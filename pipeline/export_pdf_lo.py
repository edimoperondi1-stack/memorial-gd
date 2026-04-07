"""
export_pdf_lo.py
----------------
Exporta abas específicas de um Excel para PDF usando LibreOffice.

Abordagem:
  1. Copia o xlsx para pasta temporária
  2. Instala macro LibreOffice que deleta abas indesejadas e exporta PDF
  3. Executa a macro via soffice --headless
  4. O xlsx temporário pode ser modificado, mas o original permanece intacto

Preserva shapes, drawings, imagens e toda a formatação.
"""

import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# Diretório dos scripts auxiliares do LibreOffice (soffice helper, recalc)
# Procura primeiro em lo_scripts/ (Docker), depois no skills dir (Cowork)
_POSSIBLE_SKILLS_DIRS = [
    str(Path(__file__).parent.parent / "lo_scripts"),
    "/app/lo_scripts",
    "/sessions/intelligent-adoring-hamilton/mnt/.claude/skills/xlsx/scripts",
]
SKILLS_DIR = next((d for d in _POSSIBLE_SKILLS_DIRS if Path(d).exists()), _POSSIBLE_SKILLS_DIRS[0])
if SKILLS_DIR not in sys.path:
    sys.path.insert(0, SKILLS_DIR)

from office.soffice import get_soffice_env

MACRO_DIR_LINUX = "~/.config/libreoffice/4/user/basic/Standard"
MACRO_DIR_MACOS = "~/Library/Application Support/LibreOffice/4/user/basic/Standard"
MACRO_FILENAME = "Module1.xba"


def _get_macro_dir() -> str:
    """Retorna o diretório de macros correto para o SO atual."""
    system = platform.system()
    if system == "Darwin":
        return os.path.expanduser(MACRO_DIR_MACOS)
    elif system == "Windows":
        appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
        return os.path.join(appdata, "LibreOffice", "4", "user", "basic", "Standard")
    else:
        return os.path.expanduser(MACRO_DIR_LINUX)


def _get_soffice_cmd() -> str:
    """Retorna o comando soffice. get_soffice_env() adiciona o dir ao PATH."""
    if platform.system() == "Windows":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for c in candidates:
            if os.path.exists(c):
                return c
        return "soffice.exe"
    return "soffice"


def _install_macro(abas_visiveis: list, caminho_pdf: str) -> bool:
    """Instala a macro de exportação no LibreOffice."""
    macro_dir = _get_macro_dir()
    macro_file = os.path.join(macro_dir, MACRO_FILENAME)

    # Inicializar LibreOffice se diretório não existe
    if not os.path.exists(macro_dir):
        soffice = _get_soffice_cmd()
        try:
            subprocess.run(
                [soffice, "--headless", "--terminate_after_init"],
                capture_output=True, timeout=30, env=get_soffice_env(),
            )
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
        os.makedirs(macro_dir, exist_ok=True)

    try:
        from lo_export_pdf import get_macro_content
        content = get_macro_content(abas_visiveis, caminho_pdf)
        Path(macro_file).write_text(content, encoding="utf-8")
        return True
    except Exception as e:
        print(f"  [export_pdf] Erro ao instalar macro: {e}", file=sys.stderr)
        return False


def export_pdf(caminho_xlsx: str, caminho_pdf: str, abas: list, timeout: int = 120) -> dict:
    """
    Exporta abas específicas de um xlsx para PDF via LibreOffice.

    O arquivo original NÃO é modificado — uma cópia é usada.

    Args:
        caminho_xlsx: caminho do arquivo Excel.
        caminho_pdf: caminho de destino do PDF.
        abas: lista de nomes de abas a incluir no PDF.
        timeout: timeout em segundos.

    Returns:
        dict com {"ok": bool, "caminho_pdf": str, "tamanho_bytes": int, "error": str|None}
    """
    abs_xlsx = str(Path(caminho_xlsx).absolute())
    abs_pdf = str(Path(caminho_pdf).absolute())

    # Criar diretório de saída
    Path(abs_pdf).parent.mkdir(parents=True, exist_ok=True)

    # Trabalhar numa cópia (a macro vai modificar o arquivo ao deletar abas)
    tmp_dir = tempfile.mkdtemp(prefix="export_pdf_")
    tmp_xlsx = os.path.join(tmp_dir, Path(abs_xlsx).name)
    shutil.copy2(abs_xlsx, tmp_xlsx)

    # Instalar macro
    if not _install_macro(abas, abs_pdf):
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return {"ok": False, "error": "Falha ao instalar macro LibreOffice"}

    # Executar macro
    soffice = _get_soffice_cmd()
    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "vnd.sun.star.script:Standard.Module1.ExportPDF?language=Basic&location=application",
        tmp_xlsx,
    ]

    run_timeout = None
    if platform.system() == "Linux":
        cmd = ["timeout", str(timeout)] + cmd
    elif platform.system() == "Windows":
        run_timeout = timeout + 30

    env = get_soffice_env()
    result = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=run_timeout)

    # Limpar temp
    shutil.rmtree(tmp_dir, ignore_errors=True)

    if not Path(abs_pdf).exists():
        return {
            "ok": False,
            "error": f"PDF não criado. rc={result.returncode} stdout={result.stdout[:300]} stderr={result.stderr[:300]}",
        }

    tamanho = Path(abs_pdf).stat().st_size
    return {
        "ok": True,
        "caminho_pdf": abs_pdf,
        "tamanho_bytes": tamanho,
    }


def main():
    if len(sys.argv) < 4:
        print('Uso: python export_pdf_lo.py <arquivo.xlsx> <saida.pdf> "ABA1,ABA2,ABA3" [timeout]')
        sys.exit(1)

    caminho_xlsx = sys.argv[1]
    caminho_pdf = sys.argv[2]
    abas = [a.strip() for a in sys.argv[3].split(",")]
    timeout = int(sys.argv[4]) if len(sys.argv) > 4 else 120

    resultado = export_pdf(caminho_xlsx, caminho_pdf, abas, timeout)
    print(json.dumps(resultado, indent=2))


if __name__ == "__main__":
    main()
