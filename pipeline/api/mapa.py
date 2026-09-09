"""
mapa.py
-------
Captura da imagem de satélite da localização do projeto, com a coordenada
UTM e a bússola desenhadas na própria imagem.

A imagem é produzida no servidor, não capturada da tela: assim o
enquadramento, a fonte e a posição da legenda saem idênticos em todo
projeto, e a coordenada vem calculada em vez de digitada.

O mapa no navegador serve só de visor — o usuário enquadra, e o servidor
pede exatamente aquele centro/zoom à API.

Requer a variável de ambiente GOOGLE_MAPS_API_KEY (Secret do Space).
"""

import io
import math
import os
import sys
import urllib.parse
import urllib.request
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, str(Path(__file__).parent.parent))
from utm import latlon_para_utm  # noqa: E402

STATIC_API = "https://maps.googleapis.com/maps/api/staticmap"

# Teto da Static API: 640x640 por requisição; scale=2 dobra a densidade,
# entregando 1280x1280 de pixel real.
MAX_LADO = 640
TIPOS_VALIDOS = ("hybrid", "satellite", "roadmap", "terrain")

# Fonte: no container o Dockerfile já instala fonts-dejavu-core.
# Os demais caminhos cobrem o desenvolvimento local.
_FONTES = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "C:/Windows/Fonts/segoeuib.ttf",
    "C:/Windows/Fonts/arialbd.ttf",
)


def _mundo_px(lat: float, lon: float, zoom: int):
    """
    Projeta lat/lon em pixels de mundo do Web Mercator, no zoom dado.

    O tile do Google tem 256 px, e o mundo tem 256 * 2**zoom px de lado.
    É a mesma projeção que a Static API usa, então o offset calculado aqui
    cai exatamente onde o ponto aparece na imagem devolvida.
    """
    n = 256 * (2 ** zoom)
    x = (lon + 180.0) / 360.0 * n
    seno = math.sin(math.radians(lat))
    # Clamp: a projeção diverge nos polos.
    seno = max(-0.9999, min(0.9999, seno))
    y = (0.5 - math.log((1 + seno) / (1 - seno)) / (4 * math.pi)) * n
    return x, y


def posicao_do_ponto(lat_centro: float, lon_centro: float,
                     lat_ponto: float, lon_ponto: float,
                     zoom: int, lado_px: int, escala: int):
    """
    Onde o ponto cai dentro da imagem capturada, em pixels.

    Devolve (x, y, dentro): coordenadas na imagem final e se o ponto
    está dentro do quadro. Fora do quadro o marcador não é desenhado —
    melhor não desenhar do que desenhar na borda errada.
    """
    cx, cy = _mundo_px(lat_centro, lon_centro, zoom)
    px, py = _mundo_px(lat_ponto, lon_ponto, zoom)
    meio = lado_px * escala / 2.0
    x = meio + (px - cx) * escala
    y = meio + (py - cy) * escala
    dentro = 0 <= x < lado_px * escala and 0 <= y < lado_px * escala
    return x, y, dentro


def _fonte(tamanho: int):
    for caminho in _FONTES:
        if Path(caminho).exists():
            try:
                return ImageFont.truetype(caminho, tamanho)
            except OSError:
                continue
    return ImageFont.load_default()


def _desenhar_bussola(draw: ImageDraw.ImageDraw, cx: int, cy: int, r: int):
    """
    Seta de norte no estilo de prancha: losango partido, metade vazada e
    metade cheia, com o N acima.

    Mapa de tiles em web-mercator tem o norte sempre para cima, então a
    seta é fixa — não há azimute a calcular.
    """
    preto = (0, 0, 0, 255)
    branco = (255, 255, 255, 255)

    # Fundo translucido: no hibrido o canto pode cair sobre rotulo ou pino
    # de POI do proprio Google, e a seta some no meio da poluicao.
    pad = r * 0.55
    draw.rounded_rectangle(
        [cx - r - pad, cy - r * 2.05 - pad, cx + r + pad, cy + r * 0.62 + pad],
        radius=int(r * 0.35), fill=(0, 0, 0, 130))

    topo = (cx, cy - r)
    base = (cx, cy + r * 0.62)
    esq = (cx - r * 0.52, cy + r * 0.40)
    dir_ = (cx + r * 0.52, cy + r * 0.40)

    # Contorno branco grosso primeiro: garante que a seta sobreviva tanto
    # sobre telhado claro quanto sobre vegetação escura.
    draw.polygon([topo, esq, base, dir_], fill=branco, outline=branco,
                 width=max(3, r // 6))
    draw.polygon([topo, esq, base], fill=branco, outline=preto, width=max(2, r // 14))
    draw.polygon([topo, dir_, base], fill=preto, outline=preto, width=max(2, r // 14))

    f = _fonte(int(r * 0.78))
    caixa = draw.textbbox((0, 0), "N", font=f)
    lg = caixa[2] - caixa[0]
    tx = cx - lg / 2
    ty = cy - r - (caixa[3] - caixa[1]) - r * 0.52 - caixa[1]
    for dx in (-3, -2, -1, 0, 1, 2, 3):
        for dy in (-3, -2, -1, 0, 1, 2, 3):
            draw.text((tx + dx, ty + dy), "N", font=f, fill=branco)
    draw.text((tx, ty), "N", font=f, fill=preto)


def _desenhar_marcador(draw: ImageDraw.ImageDraw, cx: int, cy: int, r: int):
    """Cruz com círculo no ponto clicado — discreta, não tapa o telhado."""
    branco = (255, 255, 255, 255)
    vermelho = (220, 30, 30, 255)
    for cor, w in ((branco, 5), (vermelho, 2)):
        draw.line([(cx - r, cy), (cx - r * 0.35, cy)], fill=cor, width=w)
        draw.line([(cx + r * 0.35, cy), (cx + r, cy)], fill=cor, width=w)
        draw.line([(cx, cy - r), (cx, cy - r * 0.35)], fill=cor, width=w)
        draw.line([(cx, cy + r * 0.35), (cx, cy + r)], fill=cor, width=w)
        draw.ellipse([cx - r * 0.30, cy - r * 0.30, cx + r * 0.30, cy + r * 0.30],
                     outline=cor, width=w)


def desenhar_overlay(img: Image.Image, lat: float, lon: float,
                     xy_marcador=None) -> Image.Image:
    """
    Aplica marcador, bussola e legenda UTM sobre a imagem capturada.

    `xy_marcador` e a posicao do ponto na imagem, em pixels. None quando o
    ponto ficou fora do quadro: nesse caso a legenda ainda sai (a coordenada
    continua valendo), mas nao se desenha marcador em lugar nenhum.
    """
    img = img.convert("RGBA")
    L, A = img.size
    camada = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(camada)

    escala = min(L, A) / 640.0

    # ── Marcador no ponto cravado ─────────────────────────────────────
    if xy_marcador is not None:
        _desenhar_marcador(draw, int(xy_marcador[0]), int(xy_marcador[1]),
                           int(22 * escala))

    # ── Bussola, canto superior direito ───────────────────────────────
    # Posicao folgada o bastante para o fundo da bussola caber inteiro:
    # com r=38, o bloco sobe ~2,6r acima do centro por causa do "N".
    r_bus = int(38 * escala)
    _desenhar_bussola(draw, L - int(80 * escala), int(122 * escala), r_bus)

    # ── Legenda UTM ───────────────────────────────────────────────────
    # Nao encosta no rodape: os termos do Google exigem que o logo (canto
    # inferior esquerdo) e a atribuicao de imagery (direita) fiquem visiveis.
    u = latlon_para_utm(lat, lon)
    texto = (f"E: {u['easting']:,.0f} m   ·   "
             f"N: {u['northing']:,.0f} m   ·   "
             f"Fuso {u['fuso']}").replace(",", ".")

    f = _fonte(int(21 * escala))
    cx = draw.textbbox((0, 0), texto, font=f)
    tw, th = cx[2] - cx[0], cx[3] - cx[1]
    pad_x, pad_y = int(20 * escala), int(11 * escala)
    margem_rodape = int(34 * escala)

    bw, bh = tw + pad_x * 2, th + pad_y * 2
    bx = (L - bw) / 2
    by = A - margem_rodape - bh
    draw.rounded_rectangle([bx, by, bx + bw, by + bh],
                           radius=int(7 * escala), fill=(0, 0, 0, 200))
    draw.text((bx + pad_x, by + pad_y - cx[1]), texto, font=f,
              fill=(255, 255, 255, 255))

    return Image.alpha_composite(img, camada).convert("RGB")


def capturar(lat: float, lon: float, zoom: int = 19,
             tipo: str = "hybrid", lado: int = MAX_LADO,
             escala: int = 2,
             lat_centro: float = None, lon_centro: float = None) -> bytes:
    """
    Busca o quadro na Maps Static API e devolve o PNG ja com overlay.

    `lat`/`lon` sao do PONTO cravado (marcador e legenda). `lat_centro`/
    `lon_centro` definem o enquadramento; quando omitidos, o ponto e o
    proprio centro — que era o comportamento anterior.

    Levanta RuntimeError se a chave nao estiver configurada, e URLError
    se a API nao responder.
    """
    chave = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    if not chave:
        raise RuntimeError(
            "GOOGLE_MAPS_API_KEY nao configurada. "
            "Defina o Secret no Space (Settings -> Variables and secrets)."
        )

    if tipo not in TIPOS_VALIDOS:
        raise ValueError(f"Tipo de mapa invalido: {tipo}")
    lado = max(200, min(int(lado), MAX_LADO))
    zoom = max(1, min(int(zoom), 21))
    escala = 2 if escala >= 2 else 1

    if lat_centro is None or lon_centro is None:
        lat_centro, lon_centro = lat, lon

    params = {
        "center": f"{lat_centro},{lon_centro}",
        "zoom": zoom,
        "size": f"{lado}x{lado}",
        "scale": escala,
        "maptype": tipo,
        "format": "png",
        "key": chave,
    }
    url = f"{STATIC_API}?{urllib.parse.urlencode(params)}"

    with urllib.request.urlopen(url, timeout=20) as resp:
        bruto = resp.read()

    x, y, dentro = posicao_do_ponto(lat_centro, lon_centro, lat, lon,
                                    zoom, lado, escala)

    img = Image.open(io.BytesIO(bruto))
    img = desenhar_overlay(img, lat, lon, (x, y) if dentro else None)

    saida = io.BytesIO()
    img.save(saida, format="PNG", optimize=True)
    return saida.getvalue()
