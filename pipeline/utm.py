"""
utm.py
------
Conversão WGS84 (lat/lon) → UTM, sem dependências externas.

Usa as fórmulas de Krüger com o elipsoide WGS84. Precisão da ordem de
milímetros dentro do fuso, muito além do necessário para localizar um
telhado numa prancha.

Existe para evitar acrescentar pyproj só por causa de uma conversão: o
resto do pipeline roda com stdlib + openpyxl/Pillow, e uma dependência
com binário compilado encarece a imagem Docker sem necessidade.

A banda de latitude (a letra) é o ponto que mais gera erro no memorial:
em Sinop-MT (~11,8 S) a banda correta é 21L, não 21K. Calculando a partir
da latitude, o rótulo sai certo sem conferência manual.
"""

import math

# Elipsoide WGS84
_A = 6378137.0                # semieixo maior (m)
_F = 1 / 298.257223563        # achatamento
_E2 = _F * (2 - _F)           # excentricidade ao quadrado
_K0 = 0.9996                  # fator de escala do UTM

_FALSE_EASTING = 500000.0
_FALSE_NORTHING = 10000000.0  # aplicado apenas no hemisfério sul

# Bandas de latitude UTM/MGRS, de 8 em 8 graus a partir de -80.
# I e O não existem (confundem com 1 e 0).
_BANDAS = "CDEFGHJKLMNPQRSTUVWX"


def zona(lon: float) -> int:
    """Número do fuso UTM (1-60) para uma longitude em graus."""
    return int((lon + 180) / 6) + 1


def banda(lat: float) -> str:
    """Letra da banda de latitude MGRS. Fora de -80..84, retorna ''."""
    if lat < -80 or lat > 84:
        return ""
    if lat > 84:
        return "X"
    idx = int((lat + 80) / 8)
    idx = min(idx, len(_BANDAS) - 1)
    return _BANDAS[idx]


def latlon_para_utm(lat: float, lon: float) -> dict:
    """
    Converte lat/lon (graus decimais, WGS84) para UTM.

    Retorna dict com: easting, northing, zona, banda, fuso, hemisferio.
    'fuso' é o rótulo pronto para a legenda (ex.: '21L').
    """
    z = zona(lon)
    b = banda(lat)

    # Meridiano central do fuso
    lon0 = math.radians(-180 + (z - 1) * 6 + 3)

    phi = math.radians(lat)
    dlam = math.radians(lon) - lon0

    n = _F / (2 - _F)
    n2, n3, n4 = n * n, n ** 3, n ** 4

    # Raio meridional retificador
    A_ret = _A / (1 + n) * (1 + n2 / 4 + n4 / 64)

    # Coeficientes de Krüger (série até 4a ordem)
    alfas = (
        n / 2 - 2 * n2 / 3 + 5 * n3 / 16,
        13 * n2 / 48 - 3 * n3 / 5,
        61 * n3 / 240,
        49561 * n4 / 161280,
    )

    # Latitude conforme, via formulação de Karney (numericamente estável)
    e = math.sqrt(_E2)
    tau = math.tan(phi)
    sigma = math.sinh(e * math.atanh(e * math.sin(phi)))
    tau_l = tau * math.sqrt(1 + sigma * sigma) - sigma * math.sqrt(1 + tau * tau)

    # atan2 com cos(dlam) no denominador: omitir esse termo custa centenas de
    # metros no northing longe do meridiano central.
    xi_ = math.atan2(tau_l, math.cos(dlam))
    eta_ = math.asinh(math.sin(dlam) / math.hypot(tau_l, math.cos(dlam)))

    xi, eta = xi_, eta_
    for j, aj in enumerate(alfas, start=1):
        xi += aj * math.sin(2 * j * xi_) * math.cosh(2 * j * eta_)
        eta += aj * math.cos(2 * j * xi_) * math.sinh(2 * j * eta_)

    easting = _K0 * A_ret * eta + _FALSE_EASTING
    northing = _K0 * A_ret * xi
    if lat < 0:
        northing += _FALSE_NORTHING

    return {
        "easting": easting,
        "northing": northing,
        "zona": z,
        "banda": b,
        "fuso": f"{z}{b}",
        "hemisferio": "S" if lat < 0 else "N",
    }


def legenda(lat: float, lon: float) -> str:
    """Linha pronta para queimar na imagem, no formato definido com o usuário."""
    u = latlon_para_utm(lat, lon)
    return f"E: {u['easting']:.0f} m  ·  N: {u['northing']:.0f} m  ·  Fuso {u['fuso']}"
