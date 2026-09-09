"""
test_utm.py
-----------
Verificação da conversão WGS84 → UTM de utm.py.

Não existe valor "oficial" fácil de citar de cor para UTM — tentar validar
por número decorado gera falso alarme. A estratégia aqui é cruzar com uma
segunda série, de derivação independente (Redfearn), e conferir o arco
meridional contra a série clássica em e². Se as três concordam, a
implementação está certa.

Rodar:  python pipeline/test_utm.py
"""

import math
import os
import sys
from pathlib import Path

# Console do Windows usa cp1252 e quebra nos acentos/box-drawing.
# Mesma abordagem já usada em pipeline/api/server.py.
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).parent))

from utm import latlon_para_utm, banda, _A, _E2, _K0


def _arco_meridional(phi: float) -> float:
    """Série clássica em e² — referência independente para o northing."""
    e2 = _E2
    e4 = e2 * e2
    e6 = e4 * e2
    return _A * (
        (1 - e2 / 4 - 3 * e4 / 64 - 5 * e6 / 256) * phi
        - (3 * e2 / 8 + 3 * e4 / 32 + 45 * e6 / 1024) * math.sin(2 * phi)
        + (15 * e4 / 256 + 45 * e6 / 1024) * math.sin(4 * phi)
        - (35 * e6 / 3072) * math.sin(6 * phi)
    )


def _redfearn(lat: float, lon: float):
    """Série de Redfearn — derivação independente da de Krüger/Karney."""
    z = int((lon + 180) / 6) + 1
    lon0 = -180 + (z - 1) * 6 + 3
    phi = math.radians(lat)
    w = math.radians(lon - lon0)
    s, c, t = math.sin(phi), math.cos(phi), math.tan(phi)
    nu = _A / math.sqrt(1 - _E2 * s * s)
    rho = _A * (1 - _E2) / (1 - _E2 * s * s) ** 1.5
    psi = nu / rho

    e = _K0 * nu * (
        w * c
        + w ** 3 / 6 * c ** 3 * (psi - t * t)
        + w ** 5 / 120 * c ** 5 * (4 * psi ** 3 * (1 - 6 * t * t)
                                   + psi * psi * (1 + 8 * t * t)
                                   - psi * 2 * t * t + t ** 4)
    ) + 500000.0

    n = _K0 * (
        _arco_meridional(phi)
        + w * w / 2 * nu * s * c
        + w ** 4 / 24 * nu * s * c ** 3 * (4 * psi * psi + psi - t * t)
        + w ** 6 / 720 * nu * s * c ** 5 * (8 * psi ** 4 * (11 - 24 * t * t)
                                            - 28 * psi ** 3 * (1 - 6 * t * t)
                                            + psi * psi * (1 - 32 * t * t)
                                            - psi * 2 * t * t + t ** 4)
    )
    if lat < 0:
        n += 1e7
    return e, n


# Pontos de teste: região de atuação, borda de fuso, equador e hemisfério norte
PONTOS = [
    ("Sinop-MT",          -11.8600, -55.5000),
    ("Sinop borda fuso",  -11.8600, -54.0500),
    ("Cuiabá-MT",         -15.6014, -56.0979),
    ("Brasília-DF",       -15.7939, -47.8828),
    ("Equador",             0.2000, -58.9000),
    ("Nova York",          40.7128, -74.0060),
]

# A banda de latitude é o campo que mais gera erro no memorial.
BANDAS = [(-11.86, "L"), (-15.79, "L"), (-17.5, "K"), (-8.5, "L"), (-7.5, "M"), (40.71, "T")]

TOL_M = 0.01  # 1 cm


def main() -> int:
    falhas = []

    print("── Krüger × Redfearn (séries independentes) ──")
    for nome, lat, lon in PONTOS:
        u = latlon_para_utm(lat, lon)
        rE, rN = _redfearn(lat, lon)
        dE, dN = u["easting"] - rE, u["northing"] - rN
        ok = abs(dE) < TOL_M and abs(dN) < TOL_M
        print(f"  {'OK ' if ok else 'FALHA '}{nome:18} fuso={u['fuso']:4} "
              f"dE={dE:+9.5f} m  dN={dN:+9.5f} m")
        if not ok:
            falhas.append(f"{nome}: dE={dE:.4f} dN={dN:.4f}")

    print("\n── Northing no meridiano central × arco meridional clássico ──")
    for nome, lat, _ in PONTOS:
        # Sobre o meridiano central o northing precisa ser exatamente k0 * M(phi).
        u = latlon_para_utm(lat, -57.0)
        n_calc = u["northing"] - (1e7 if lat < 0 else 0)
        n_ref = _K0 * _arco_meridional(math.radians(lat))
        d = n_calc - n_ref
        ok = abs(d) < TOL_M
        print(f"  {'OK ' if ok else 'FALHA '}{nome:18} lat={lat:9.4f}  d={d:+9.5f} m")
        if not ok:
            falhas.append(f"arco {nome}: d={d:.4f}")

    print("\n── Banda de latitude (o erro 21K × 21L do memorial) ──")
    for lat, esperada in BANDAS:
        obtida = banda(lat)
        ok = obtida == esperada
        print(f"  {'OK ' if ok else 'FALHA '}lat={lat:8.2f} → {obtida}  (esperado {esperada})")
        if not ok:
            falhas.append(f"banda lat={lat}: {obtida} != {esperada}")

    print()
    if falhas:
        print(f"FALHOU — {len(falhas)} problema(s):")
        for f in falhas:
            print(f"  · {f}")
        return 1
    print("Todos os testes passaram.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
