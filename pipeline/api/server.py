"""
server.py
---------
API HTTP para o pipeline de geração de documentos GD.

Endpoints:
  GET  /                         → Frontend (formulário web)
  GET  /static/<file>            → Arquivos estáticos
  POST /api/gerar                → Gera .xlsx + .pdf e retorna links
  GET  /api/download/<file>      → Download do arquivo gerado
  GET  /api/equipamentos         → Lista de fabricantes/modelos (para autocomplete)
  GET  /api/status               → Health check

Uso:
  python server.py [porta]
  # Default: porta 8080
"""

import hashlib
import json
import os
import secrets
import shutil
import sys

# ── Forçar UTF-8 no stdout/stderr (essencial no Windows) ────────────
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
import time
import traceback
import urllib.parse
import uuid
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, unquote

# ─── ThreadingHTTPServer ─────────────────────────────────────────────────────
# Python >= 3.7: http.server já inclui ThreadingHTTPServer.
# Fallback manual para versões mais antigas.
try:
    from http.server import ThreadingHTTPServer
except ImportError:
    import socketserver
    class ThreadingHTTPServer(socketserver.ThreadingMixIn, HTTPServer):
        daemon_threads = True

# Adicionar diretório do pipeline ao path
PIPELINE_DIR = Path(__file__).parent.parent
if str(PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(PIPELINE_DIR))

from modelos import DadosProjeto, Painel, Inversor, UCBeneficiaria
from run_pipeline import executar_pipeline

# Diretório onde os arquivos gerados ficam disponíveis para download
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# Diretório de arquivos estáticos (frontend)
STATIC_DIR = Path(__file__).parent / "static"

# Tempo de retenção dos arquivos gerados (em segundos): 24 horas
OUTPUT_RETENTION_SECS = 24 * 3600

# Base de equipamentos (painéis e inversores)
EQUIPAMENTOS_PATH = Path(__file__).parent / "equipamentos.json"
import threading
_equip_lock = threading.Lock()

# Diretório para projetos salvos
PROJETOS_DIR = Path(__file__).parent / "projetos_salvos"
PROJETOS_DIR.mkdir(exist_ok=True)

# ── Autenticação ─────────────────────────────────────────────────────────────
USERS_PATH = Path(__file__).parent / "users.json"
SESSION_COOKIE = "memorial_session"
SESSION_DURATION = 8 * 3600  # 8 horas em segundos

# Sessões em memória: {token: {"usuario": str, "expira": float}}
_sessions: dict = {}
_sessions_lock = threading.Lock()


def _hash_senha(senha: str) -> str:
    return hashlib.sha256(senha.encode("utf-8")).hexdigest()


def _carregar_usuarios() -> dict:
    if USERS_PATH.exists():
        try:
            with open(USERS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    # Criar arquivo padrão se não existir
    padrao = {"admin": {"senha_hash": _hash_senha("memorial2026"), "nome": "Administrador"}}
    _salvar_usuarios(padrao)
    print("  [auth] users.json criado com usuário padrão: admin / memorial2026")
    return padrao


def _salvar_usuarios(users: dict):
    with open(USERS_PATH, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def _criar_sessao(usuario: str) -> str:
    token = secrets.token_hex(32)
    with _sessions_lock:
        _sessions[token] = {"usuario": usuario, "expira": time.time() + SESSION_DURATION}
    return token


def _verificar_sessao(token: str) -> str | None:
    """Retorna o nome do usuário se a sessão for válida, None caso contrário."""
    with _sessions_lock:
        sess = _sessions.get(token)
        if sess and time.time() < sess["expira"]:
            return sess["usuario"]
        if sess:
            del _sessions[token]
    return None


def _get_token_do_request(handler) -> str | None:
    """Extrai token do header Authorization (Bearer) OU do cookie (fallback)."""
    # 1. Header Authorization: Bearer <token>
    auth = handler.headers.get("Authorization", "")
    if auth.startswith("Bearer "):
        return auth[7:].strip()
    # 2. Fallback: cookie (para acesso direto não-iframe)
    cookie_header = handler.headers.get("Cookie", "")
    for parte in cookie_header.split(";"):
        parte = parte.strip()
        if parte.startswith(SESSION_COOKIE + "="):
            return parte[len(SESSION_COOKIE) + 1:]
    return None


def _rota_publica(path: str) -> bool:
    """Rotas que não requerem autenticação."""
    return path in ("/login", "/api/login", "/api/status") or path.startswith("/static/")


def _carregar_equipamentos() -> dict:
    """Carrega a base de equipamentos do JSON."""
    if EQUIPAMENTOS_PATH.exists():
        try:
            with open(EQUIPAMENTOS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {"paineis": {}, "inversores": {}}


def _salvar_equipamentos(data: dict):
    """Salva a base de equipamentos no JSON (thread-safe)."""
    with _equip_lock:
        try:
            with open(EQUIPAMENTOS_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except OSError as e:
            print(f"  [equip] Erro ao salvar equipamentos: {e}")


def _registrar_novos_equipamentos(payload: dict):
    """Após gerar documentos, salva novos fabricantes/modelos na base."""
    equip = _carregar_equipamentos()
    alterou = False

    for p in payload.get("paineis", []):
        fab = str(p.get("fabricante", "")).strip().upper()
        mod = str(p.get("modelo", "")).strip().upper()
        if fab and mod:
            if fab not in equip["paineis"]:
                equip["paineis"][fab] = []
            if mod not in equip["paineis"][fab]:
                equip["paineis"][fab].append(mod)
                equip["paineis"][fab].sort()
                alterou = True

    for inv in payload.get("inversores", []):
        fab = str(inv.get("fabricante", "")).strip().upper()
        mod = str(inv.get("modelo", "")).strip().upper()
        if fab and mod:
            if fab not in equip["inversores"]:
                equip["inversores"][fab] = []
            if mod not in equip["inversores"][fab]:
                equip["inversores"][fab].append(mod)
                equip["inversores"][fab].sort()
                alterou = True

    if alterou:
        # Ordenar fabricantes
        equip["paineis"] = dict(sorted(equip["paineis"].items()))
        equip["inversores"] = dict(sorted(equip["inversores"].items()))
        _salvar_equipamentos(equip)
        print("  [equip] Novos equipamentos registrados na base.")


def _limpar_output_antigo():
    """Remove subpastas de execuções com mais de OUTPUT_RETENTION_SECS segundos."""
    agora = time.time()
    try:
        for subdir in OUTPUT_DIR.iterdir():
            if not subdir.is_dir():
                continue
            idade = agora - subdir.stat().st_mtime
            if idade > OUTPUT_RETENTION_SECS:
                shutil.rmtree(subdir, ignore_errors=True)
    except Exception:
        pass  # Limpeza é best-effort — não deve quebrar a requisição


class APIHandler(SimpleHTTPRequestHandler):
    """Handler HTTP com rotas para API e frontend."""

    def do_GET(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)

        # ── Verificar autenticação ────────────────────────────────────
        if not _rota_publica(path):
            token = _get_token_do_request(self)
            usuario = _verificar_sessao(token) if token else None
            if not usuario:
                # API → 401; página → redireciona para /login
                if path.startswith("/api/"):
                    self._json_response(401, {"error": "Não autenticado"})
                else:
                    self.send_response(302)
                    self.send_header("Location", "/login")
                    self.end_headers()
                return

        if path == "/login":
            self._serve_file(STATIC_DIR / "login.html", "text/html")

        elif path == "/" or path == "":
            self._serve_file(STATIC_DIR / "index.html", "text/html")

        elif path.startswith("/static/"):
            filename = path[len("/static/"):]
            filepath = STATIC_DIR / filename
            content_type = self._guess_type(filename)
            self._serve_file(filepath, content_type)

        elif path.startswith("/api/download/"):
            # Path: /api/download/<exec_id>/<filename>
            subpath = path[len("/api/download/"):]
            filepath = OUTPUT_DIR / subpath
            # Segurança: não permitir path traversal
            try:
                filepath.resolve().relative_to(OUTPUT_DIR.resolve())
            except ValueError:
                self._json_response(403, {"error": "Acesso negado"})
                return
            if ".." in subpath:
                self._json_response(403, {"error": "Acesso negado"})
                return
            if not filepath.exists():
                self._json_response(404, {"error": "Arquivo não encontrado"})
                return
            content_type = (
                "application/pdf"
                if filepath.name.endswith(".pdf")
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            self._serve_file(filepath, content_type, download=True)

        elif path == "/api/equipamentos":
            equip = _carregar_equipamentos()
            self._json_response(200, equip)

        elif path == "/api/status":
            self._json_response(200, {"status": "ok", "pipeline": "ready"})

        elif path == "/api/admin/usuarios":
            self._handle_listar_usuarios()

        elif path == "/api/debug":
            self._handle_debug()

        elif path == "/api/projetos":
            # Listar projetos salvos
            projetos = []
            for filepath in PROJETOS_DIR.glob("*.json"):
                projetos.append(filepath.stem)
            projetos.sort(key=str.lower)
            self._json_response(200, {"projetos": projetos})

        elif path.startswith("/api/projetos/"):
            # Obter um projeto específico
            nome = path[len("/api/projetos/"):]
            filepath = PROJETOS_DIR / f"{nome}.json"
            if not filepath.exists():
                self._json_response(404, {"error": "Projeto não encontrado"})
                return
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                self._json_response(200, dados)
            except Exception as e:
                self._json_response(500, {"error": f"Erro ao carregar projeto: {e}"})

        elif path.startswith("/api/temperatura"):
            self._handle_temperatura(parsed)

        else:
            self._json_response(404, {"error": "Rota não encontrada"})

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/")

        # ── Verificar autenticação (exceto /api/login) ────────────────
        if not _rota_publica(path):
            token = _get_token_do_request(self)
            usuario = _verificar_sessao(token) if token else None
            if not usuario:
                self._json_response(401, {"error": "Não autenticado"})
                return

        if path == "/api/login":
            self._handle_login()
        elif path == "/api/logout":
            self._handle_logout()
        elif path == "/api/gerar":
            self._handle_gerar()
        elif path == "/api/projetos":
            self._handle_salvar_projeto()
        elif path == "/api/admin/usuarios":
            self._handle_admin_usuarios()
        elif path == "/api/admin/usuarios/remover":
            self._handle_remover_usuario()
        else:
            self._json_response(404, {"error": f"Rota POST não encontrada: {path}"})

    def _handle_login(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            payload = json.loads(body.decode("utf-8"))
            usuario = payload.get("usuario", "").strip()
            senha = payload.get("senha", "")

            users = _carregar_usuarios()
            user_data = users.get(usuario)
            if not user_data or user_data.get("senha_hash") != _hash_senha(senha):
                self._json_response(401, {"error": "Usuário ou senha inválidos"})
                return

            token = _criar_sessao(usuario)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            # Cookie como fallback (pode ser bloqueado em iframe)
            self.send_header(
                "Set-Cookie",
                f"{SESSION_COOKIE}={token}; Path=/; HttpOnly; SameSite=None; Secure; Max-Age={SESSION_DURATION}"
            )
            self._cors_headers()
            # Token também no JSON para o cliente salvar no localStorage
            body_resp = json.dumps({"ok": True, "usuario": usuario, "token": token}).encode("utf-8")
            self.send_header("Content-Length", str(len(body_resp)))
            self.end_headers()
            self.wfile.write(body_resp)
            print(f"  [auth] Login: {usuario}")
        except Exception as e:
            self._json_response(500, {"error": str(e)})

    def _handle_logout(self):
        token = _get_token_do_request(self)
        if token:
            with _sessions_lock:
                _sessions.pop(token, None)
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header(
            "Set-Cookie",
            f"{SESSION_COOKIE}=; Path=/; HttpOnly; Max-Age=0"
        )
        self._cors_headers()
        body_resp = json.dumps({"ok": True}).encode("utf-8")
        self.send_header("Content-Length", str(len(body_resp)))
        self.end_headers()
        self.wfile.write(body_resp)

    def _handle_listar_usuarios(self):
        users = _carregar_usuarios()
        lista = [{"usuario": u, "nome": d.get("nome", u)} for u, d in users.items()]
        self._json_response(200, {"usuarios": lista})

    def _handle_admin_usuarios(self):
        """POST /api/admin/usuarios — adicionar novo usuário."""
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            payload = json.loads(body.decode("utf-8"))
            usuario = payload.get("usuario", "").strip()
            senha = payload.get("senha", "")
            nome = payload.get("nome", usuario).strip()

            if not usuario or not senha:
                self._json_response(400, {"error": "Usuário e senha são obrigatórios"})
                return
            if len(senha) < 6:
                self._json_response(400, {"error": "Senha deve ter pelo menos 6 caracteres"})
                return

            users = _carregar_usuarios()
            users[usuario] = {"senha_hash": _hash_senha(senha), "nome": nome}
            _salvar_usuarios(users)
            print(f"  [auth] Usuário criado: {usuario}")
            self._json_response(200, {"ok": True, "usuario": usuario})
        except Exception as e:
            self._json_response(500, {"error": str(e)})

    def _handle_remover_usuario(self):
        """POST /api/admin/usuarios/remover — remover usuário."""
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            payload = json.loads(body.decode("utf-8"))
            usuario = payload.get("usuario", "").strip()

            if not usuario:
                self._json_response(400, {"error": "Usuário é obrigatório"})
                return

            users = _carregar_usuarios()
            if usuario not in users:
                self._json_response(404, {"error": "Usuário não encontrado"})
                return
            if len(users) == 1:
                self._json_response(400, {"error": "Não é possível remover o único usuário"})
                return

            del users[usuario]
            _salvar_usuarios(users)
            # Invalidar sessões desse usuário
            with _sessions_lock:
                to_del = [t for t, s in _sessions.items() if s["usuario"] == usuario]
                for t in to_del:
                    del _sessions[t]
            print(f"  [auth] Usuário removido: {usuario}")
            self._json_response(200, {"ok": True})
        except Exception as e:
            self._json_response(500, {"error": str(e)})

    def _handle_gerar(self):
        """Processa o formulário e gera os documentos."""
        # Limpar arquivos antigos antes de gerar novos (best-effort)
        _limpar_output_antigo()

        try:
            # Ler body JSON
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            payload = json.loads(body.decode("utf-8"))

            # Converter JSON → DadosProjeto
            dados = _json_para_dados(payload)

            # Gerar um ID único para esta execução
            exec_id = str(uuid.uuid4())[:8]
            pasta_saida = str(OUTPUT_DIR / exec_id)

            # Executar pipeline
            resultado = executar_pipeline(dados, pasta_saida=pasta_saida)

            if not resultado["ok"]:
                self._json_response(500, {
                    "error": "Pipeline concluído com avisos",
                    "relatorio": _sanitize_relatorio(resultado["relatorio"]),
                })
                return

            # Registrar novos equipamentos na base (best-effort)
            try:
                _registrar_novos_equipamentos(payload)
            except Exception:
                pass

            # Montar URLs de download
            xlsx_name = Path(resultado["xlsx"]).name
            pdf_name = Path(resultado["pdf"]).name

            response = {
                "ok": True,
                "xlsx_url": f"/api/download/{exec_id}/{urllib.parse.quote(xlsx_name)}",
                "pdf_url": f"/api/download/{exec_id}/{urllib.parse.quote(pdf_name)}",
                "xlsx_nome": xlsx_name,
                "pdf_nome": pdf_name,
            }
            if resultado.get("procuracao") and Path(resultado["procuracao"]).exists():
                proc_name = Path(resultado["procuracao"]).name
                response["procuracao_url"] = f"/api/download/{exec_id}/{urllib.parse.quote(proc_name)}"
                response["procuracao_nome"] = proc_name

            self._json_response(200, response)

        except json.JSONDecodeError as e:
            self._json_response(400, {"error": f"JSON inválido: {e}"})
        except UnicodeEncodeError as e:
            # Emoji nos prints do pipeline pode causar isso no Windows
            self._json_response(500, {"error": f"Erro de encoding (tente PYTHONIOENCODING=utf-8): {e}"})
        except (ValueError, KeyError, TypeError) as e:
            self._json_response(400, {"error": f"Dados inválidos: {e}"})
        except Exception as e:
            tb = traceback.format_exc()
            traceback.print_exc()
            self._json_response(500, {"error": f"Erro interno: {e}", "traceback": tb})

    def _handle_debug(self):
        """Diagnóstico: verifica LibreOffice, caminhos e dependências."""
        import platform, subprocess, shutil
        info = {"python": sys.version, "plataforma": platform.platform()}

        # Verificar LibreOffice
        lo_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice",
        ]
        lo_found = None
        for p in lo_paths:
            if os.path.exists(p):
                lo_found = p
                break
        info["libreoffice_path"] = lo_found or "(not found via os.path.exists)"
        info["libreoffice_in_PATH"] = bool(shutil.which("soffice") or shutil.which("soffice.exe"))

        # Tentar rodar soffice --version
        soffice_cmd = lo_found or "soffice"
        try:
            rv = subprocess.run(
                [soffice_cmd, "--version"],
                capture_output=True, text=True, timeout=15
            )
            info["libreoffice_version"] = rv.stdout.strip() or rv.stderr.strip()
        except FileNotFoundError:
            info["libreoffice_version"] = "ERRO: soffice nao encontrado"
        except Exception as e:
            info["libreoffice_version"] = f"ERRO: {e}"

        # Verificar pdftotext
        info["pdftotext"] = bool(shutil.which("pdftotext"))

        # Verificar recalc.py
        from pathlib import Path as P
        recalc_candidates = [
            P(__file__).parent.parent.parent / "lo_scripts" / "recalc.py",
            P("/app/lo_scripts/recalc.py"),
        ]
        found_recalc = next((str(p) for p in recalc_candidates if p.exists()), None)
        info["recalc_py"] = found_recalc or "(not found)"

        # Variáveis de ambiente relevantes
        info["APPDATA"] = os.environ.get("APPDATA", "(nao definido)")
        info["PATH_primeiros_500"] = os.environ.get("PATH", "")[:500]

        self._json_response(200, info)

    def _handle_salvar_projeto(self):
        """Salva o payload recebido como um projeto na pasta projetos_salvos."""
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            payload = json.loads(body.decode("utf-8"))

            nome_projeto = payload.get("nome_projeto", "").strip()
            if not nome_projeto:
                self._json_response(400, {"error": "Nome do projeto é obrigatório"})
                return

            # Limpar nome para evitar problemas de path traversal ou caracteres inválidos
            safe_name = "".join(c for c in nome_projeto if c.isalnum() or c in (" ", "-", "_")).strip()
            if not safe_name:
                self._json_response(400, {"error": "Nome do projeto inválido"})
                return

            filepath = PROJETOS_DIR / f"{safe_name}.json"
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)

            self._json_response(200, {"ok": True, "nome": safe_name})

        except Exception as e:
            self._json_response(500, {"error": str(e)})

    # Mapeamento UF → nome completo do estado (para filtrar resultados de geocoding)
    _UF_PARA_ESTADO = {
        "AC": "Acre", "AL": "Alagoas", "AP": "Amapá", "AM": "Amazonas",
        "BA": "Bahia", "CE": "Ceará", "DF": "Distrito Federal",
        "ES": "Espírito Santo", "GO": "Goiás", "MA": "Maranhão",
        "MT": "Mato Grosso", "MS": "Mato Grosso do Sul", "MG": "Minas Gerais",
        "PA": "Pará", "PB": "Paraíba", "PR": "Paraná", "PE": "Pernambuco",
        "PI": "Piauí", "RJ": "Rio de Janeiro", "RN": "Rio Grande do Norte",
        "RS": "Rio Grande do Sul", "RO": "Rondônia", "RR": "Roraima",
        "SC": "Santa Catarina", "SP": "São Paulo", "SE": "Sergipe",
        "TO": "Tocantins",
    }

    def _handle_temperatura(self, parsed):
        """GET /api/temperatura?cidade=Campinas&uf=SP
        Retorna T_min histórica (percentil 1%) via Open-Meteo.
        Requer apenas stdlib: urllib.request, urllib.parse, json.
        """
        import urllib.request
        import urllib.error

        params = urllib.parse.parse_qs(parsed.query)
        cidade = params.get("cidade", [""])[0].strip()
        uf = params.get("uf", [""])[0].strip().upper()

        if not cidade or not uf:
            self._json_response(400, {"error": "Parâmetros 'cidade' e 'uf' são obrigatórios"})
            return

        estado_esperado = self._UF_PARA_ESTADO.get(uf)
        if not estado_esperado:
            self._json_response(400, {"error": f"UF inválida: {uf}"})
            return

        # ── Passo 1: Geocoding ───────────────────────────────────────
        geo_url = (
            "https://geocoding-api.open-meteo.com/v1/search"
            f"?name={urllib.parse.quote(cidade)}&count=5&language=pt&countryCode=BR"
        )
        try:
            with urllib.request.urlopen(geo_url, timeout=10) as resp:
                geo_data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.URLError:
            self._json_response(500, {"error": "Erro de rede ao buscar dados climáticos"})
            return

        resultados = geo_data.get("results", [])
        # Filtrar por estado correspondente à UF informada
        match = next(
            (r for r in resultados if r.get("admin1", "") == estado_esperado),
            None
        )
        if match is None:
            self._json_response(404, {"error": f"Município não encontrado: {cidade}/{uf}"})
            return

        lat = match["latitude"]
        lon = match["longitude"]
        cidade_resolvida = match["name"]
        estado_resolvido = match.get("admin1", estado_esperado)

        # ── Passo 2: Dados climáticos históricos (1990–2020) ─────────
        clima_url = (
            "https://climate-api.open-meteo.com/v1/climate"
            f"?latitude={lat}&longitude={lon}"
            "&start_date=1990-01-01&end_date=2020-12-31"
            "&models=ERA5&daily=temperature_2m_min"
        )
        try:
            with urllib.request.urlopen(clima_url, timeout=10) as resp:
                clima_data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.URLError:
            self._json_response(500, {"error": "Erro de rede ao buscar dados climáticos"})
            return

        temps = clima_data.get("daily", {}).get("temperature_2m_min", [])
        # Remover possíveis valores None (dias sem dado)
        temps = [t for t in temps if t is not None]
        if not temps:
            self._json_response(500, {"error": f"Dados climáticos insuficientes para {cidade}"})
            return

        # Percentil 1%: índice floor(n * 0.01) dos valores ordenados
        temps_sorted = sorted(temps)
        idx = int(len(temps_sorted) * 0.01)
        t_min_p1 = round(temps_sorted[idx], 1)

        self._json_response(200, {
            "cidade_resolvida": cidade_resolvida,
            "estado": estado_resolvido,
            "latitude": lat,
            "longitude": lon,
            "t_min_p1": t_min_p1,
        })

    def _json_response(self, status: int, data: dict):
        """Envia resposta JSON."""
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self._cors_headers()
        self.end_headers()
        self.wfile.write(body)

    def _serve_file(self, filepath: Path, content_type: str, download: bool = False):
        """Serve um arquivo estático."""
        if not filepath.exists():
            self._json_response(404, {"error": "Arquivo não encontrado"})
            return
        data = filepath.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self._cors_headers()
        if download:
            # RFC 5987: suporte a nomes com caracteres não-ASCII (acentos, espaços)
            encoded_name = urllib.parse.quote(filepath.name, safe="")
            self.send_header(
                "Content-Disposition",
                f"attachment; filename*=UTF-8''{encoded_name}",
            )
        self.end_headers()
        self.wfile.write(data)

    def _cors_headers(self):
        """Adiciona headers CORS a qualquer resposta."""
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _guess_type(self, filename: str) -> str:
        ext = Path(filename).suffix.lower()
        return {
            ".html": "text/html",
            ".css": "text/css",
            ".js": "application/javascript",
            ".json": "application/json",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".svg": "image/svg+xml",
            ".ico": "image/x-icon",
        }.get(ext, "application/octet-stream")

    def do_OPTIONS(self):
        """CORS preflight."""
        self.send_response(200)
        self._cors_headers()
        self.end_headers()

    def log_message(self, format, *args):
        """Log mais limpo."""
        print(f"  [{self.log_date_time_string()}] {args[0]}")


def _json_para_dados(payload: dict) -> DadosProjeto:
    """Converte o payload JSON do frontend para DadosProjeto."""

    # Converter painéis
    paineis = []
    for p in payload.get("paineis", []):
        paineis.append(Painel(
            quantidade=int(p["quantidade"]),
            fabricante=str(p["fabricante"]),
            modelo=str(p["modelo"]),
            area_m2=float(p.get("area_m2", 0)),
            potencia_kw=float(p.get("potencia_kw", 0)),
        ))

    # Converter inversores
    inversores = []
    for inv in payload.get("inversores", []):
        inversores.append(Inversor(
            quantidade=int(inv["quantidade"]),
            fabricante=str(inv["fabricante"]),
            modelo=str(inv["modelo"]),
            potencia_kw=float(inv.get("potencia_kw", 0)),
            tensao_nominal_v=float(inv.get("tensao_nominal_v", 0)),
        ))

    # Converter UCs beneficiárias
    ucs = []
    for uc in payload.get("ucs_beneficiarias", []):
        ucs.append(UCBeneficiaria(
            codigo_uc=str(uc["codigo_uc"]),
            titular=str(uc["titular"]),
            cpf_cnpj=str(uc["cpf_cnpj"]),
            endereco=str(uc.get("endereco", "")),
            percentual=float(uc.get("percentual", 0)),
        ))

    # Converter carga instalada
    carga = []
    for c in payload.get("carga_instalada", []):
        carga.append((
            int(c["quantidade"]),
            str(c["equipamento"]),
            float(c["potencia_w"]),
            float(c.get("fator_demanda", 1.0)),
        ))

    return DadosProjeto(
        # Identificação
        codigo_uc=str(payload["codigo_uc"]),
        titular=str(payload["titular"]),
        classe=str(payload["classe"]),
        cpf_cnpj=str(payload["cpf_cnpj"]),
        logradouro=str(payload["logradouro"]),
        numero=str(payload["numero"]),
        bairro=str(payload["bairro"]),
        cidade=str(payload["cidade"]),
        uf=str(payload["uf"]),
        cep=str(payload["cep"]),
        email=str(payload.get("email", "")),
        telefone=str(payload.get("telefone", "")),
        celular=str(payload.get("celular", "")),

        # Padrão elétrico
        potencia_instalada_kw=float(payload["potencia_instalada_kw"]),
        tensao_atendimento_v=str(payload["tensao_atendimento_v"]),
        tipo_conexao=str(payload["tipo_conexao"]),
        tipo_ramal=str(payload["tipo_ramal"]),

        # Geração
        tipo_fonte=str(payload.get("tipo_fonte", "SOLAR FOTOVOLTAICA")),
        tipo_geracao=str(payload.get("tipo_geracao", "Empregando conversor eletrônico/inversor")),
        modalidade=str(payload["modalidade"]),
        potencia_geracao_kwp=float(payload["potencia_geracao_kwp"]),

        # Detalhes técnicos
        tipo_padrao=str(payload.get("tipo_padrao", payload["tipo_conexao"])),
        nivel_tensao_v=str(payload.get("nivel_tensao_v", payload["tensao_atendimento_v"])),
        potencia_max_disponivel_kw=float(payload.get("potencia_max_disponivel_kw", payload["potencia_instalada_kw"])),
        disjuntor_geral_a=int(payload["disjuntor_geral_a"]),
        fator_potencia=float(payload.get("fator_potencia", 0.92)),
        demanda_contratada_kw=float(payload.get("demanda_contratada_kw", 1.0)),
        dps_ca_ka=float(payload.get("dps_ca_ka", 0)),
        disjuntor_ca_a=float(payload.get("disjuntor_ca_a", 0)),
        tem_stringbox=bool(payload.get("tem_stringbox", False)),
        dps_cc_ka=float(payload.get("dps_cc_ka", 0)),
        disjuntor_cc_a=float(payload.get("disjuntor_cc_a", 0)),
        num_fases=int(payload.get("num_fases", 2)),
        cabos_por_fase=int(payload.get("cabos_por_fase", 1)),
        bitola_fase_mm2=float(payload.get("bitola_fase_mm2", 10.0)),
        bitola_neutro_mm2=float(payload.get("bitola_neutro_mm2", 10.0)),
        bitola_terra_mm2=float(payload.get("bitola_terra_mm2", 10.0)),

        # Coordenadas
        fuso=str(payload.get("fuso", "")),
        coord_x_long=float(payload.get("coord_x_long", 0)),
        coord_y_lat=float(payload.get("coord_y_lat", 0)),

        # Equipamentos
        paineis=paineis,
        inversores=inversores,
        ucs_beneficiarias=ucs,
        carga_instalada=carga,

        # Responsável técnico
        resp_nome=str(payload.get("resp_nome", "")),
        resp_cpf=str(payload.get("resp_cpf", "")),
        resp_telefone=str(payload.get("resp_telefone", "")),
        resp_email=str(payload.get("resp_email", "")),
        resp_endereco=str(payload.get("resp_endereco", "")),

        # Tipo de formulário
        tipo_fsa=str(payload.get("tipo_fsa", "SOLICITACAO")),

        # Previsão
        previsao_mes=str(payload.get("previsao_mes", "JANEIRO")),
        previsao_ano=int(payload.get("previsao_ano", 2026)),

        # Trafo
        potencia_trafo_kw=float(payload.get("potencia_trafo_kw", 0)),

        # Formulário de orçamento (dict de item_key → "X"/"SIM"/etc.)
        formulario_items=payload.get("formulario_items") or {},
    )


def _sanitize_relatorio(relatorio: dict) -> dict:
    """Remove paths internos do relatório para não expor no JSON."""
    safe = {}
    for k, v in relatorio.items():
        if isinstance(v, dict):
            safe[k] = {kk: str(vv) if "caminho" in kk else vv for kk, vv in v.items()}
        else:
            safe[k] = v
    return safe


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
    server = ThreadingHTTPServer(("0.0.0.0", port), APIHandler)
    print(f"\n{'='*50}")
    print(f"  Memorial GD — API Server  (threading)")
    print(f"  http://localhost:{port}")
    print(f"{'='*50}\n")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor encerrado.")
        server.server_close()


if __name__ == "__main__":
    main()
