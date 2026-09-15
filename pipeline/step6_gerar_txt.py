"""
step6_gerar_txt.py
------------------
Gera um arquivo TXT com os dados resumidos do projeto, no formato usado
para auxiliar o preenchimento do restante da documentação.

Formato de saída:

    LOCALIZAÇÃO:
    11°50'28.5"S 55°30'59.0"W
    -11.841238, -55.516384
    X: 661615.64, Y: 8690572.42
    FUSO: 21L

    ENDEREÇO:<logradouro>,
    Nº <numero>, BAIRRO: <bairro>, CIDADE: <cidade>, CEP <cep>


    UC: <codigo_uc>


    PROPRIETARIO: <titular>
    CPF: <cpf_cnpj>




    TELEFONE: <telefone>
    CELULAR: <celular>

    EMAIL: <email>


    MODELO MODULOS: <qtd>x <modelo> <potencia_w>W
    FABRICANTE MODULOS: <fabricante>

    MODELO INVERSOR: <qtd>x <modelo>
    FABRICANTE INVERSOR: <fabricante>

Em ampliação (geração existente informada), sai ainda um bloco
"GERAÇÃO EXISTENTE (AMPLIAÇÃO)" com os módulos e inversores já instalados
e as potências existente / nova / total em kWp.

A localização sai nos três formatos porque cada documento pede um: a ART e
o memorial usam UTM, o Google Maps e o cadastro da Energisa usam graus
decimais, e a prancha costuma trazer graus-minutos-segundos. O formulário
só recebe UTM + fuso; lat/lon são calculados pela inversa em utm.py.
"""

from pathlib import Path

from modelos import sanitize_filename_part
from utm import formatar_dms, interpretar_fuso, utm_para_latlon


def _linhas_localizacao(dados) -> list:
    """Bloco de localização: DMS, decimal e UTM (quando o fuso permite)."""
    x = float(getattr(dados, "coord_x_long", 0) or 0)
    y = float(getattr(dados, "coord_y_lat", 0) or 0)
    fuso = str(getattr(dados, "fuso", "") or "").strip()

    linhas = ["LOCALIZAÇÃO:"]
    zona = interpretar_fuso(fuso) if (x and y) else None
    if zona:
        lat, lon = utm_para_latlon(x, y, zona[0], zona[1])
        linhas.append(formatar_dms(lat, lon))
        linhas.append(f"{lat:.6f}, {lon:.6f}")
    else:
        # Sem fuso (ou sem coordenada) não há como voltar para lat/lon —
        # melhor deixar explícito do que chutar a zona.
        linhas.append("[lat/lon não calculados: informe o fuso UTM, ex.: 21L]")
    linhas.append(f"X: {x:.2f}, Y: {y:.2f}")
    linhas.append(f"FUSO: {fuso or '[não informado]'}")
    return linhas


def _linhas_modulos(paineis: list, sufixo: str = "") -> list:
    """Uma dupla MODELO/FABRICANTE por lote de módulos."""
    linhas = []
    if paineis:
        for p in paineis:
            pot_w = int(round(float(p.potencia_kw) * 1000))
            qtd = getattr(p, "quantidade", None) or 1
            linhas.append(f"MODELO MODULOS{sufixo}: {qtd}x {p.modelo} {pot_w}W ")
            linhas.append(f"FABRICANTE MODULOS{sufixo}: {p.fabricante}  ")
            linhas.append("")
        if linhas and linhas[-1] == "":
            linhas.pop()
    else:
        linhas = [f"MODELO MODULOS{sufixo}: ", f"FABRICANTE MODULOS{sufixo}:  "]
    return linhas


def _linhas_inversores(inversores: list, sufixo: str = "") -> list:
    """Uma dupla MODELO/FABRICANTE por lote de inversores."""
    linhas = []
    if inversores:
        for inv in inversores:
            qtd = getattr(inv, "quantidade", None) or 1
            linhas.append(f"MODELO INVERSOR{sufixo}: {qtd}x {inv.modelo}   ")
            linhas.append(f"FABRICANTE INVERSOR{sufixo}:  {inv.fabricante}   ")
            linhas.append("")
        if linhas and linhas[-1] == "":
            linhas.pop()
    else:
        linhas = [f"MODELO INVERSOR{sufixo}:    ", f"FABRICANTE INVERSOR{sufixo}:     "]
    return linhas


def _kwp(paineis: list) -> float:
    """Soma quantidade × potência dos lotes de módulos, em kWp."""
    total = 0.0
    for p in paineis or []:
        qtd = getattr(p, "quantidade", None) or 1
        total += float(qtd) * float(p.potencia_kw or 0)
    return total


def _fmt_kwp(valor: float) -> str:
    """kWp com vírgula decimal, como no nome das pastas (5,40 · 15,32)."""
    return f"{valor:.2f}".replace(".", ",")


def _linhas_geracao_existente(dados) -> list:
    """
    Bloco da ampliação. Vazio quando não há geração existente — o TXT de um
    projeto novo continua igual ao de sempre.
    """
    pe = getattr(dados, "paineis_existentes", None) or []
    ie = getattr(dados, "inversores_existentes", None) or []
    if not pe and not ie:
        return []

    kwp_exist = _kwp(pe)
    kwp_nova = _kwp(getattr(dados, "paineis", None) or [])
    return [
        "",
        "",
        "GERAÇÃO EXISTENTE (AMPLIAÇÃO):",
        "",
        *_linhas_modulos(pe, " EXISTENTES"),
        "",
        *_linhas_inversores(ie, " EXISTENTE"),
        "",
        f"POTENCIA EXISTENTE: {_fmt_kwp(kwp_exist)} kWp",
        f"POTENCIA NOVA: {_fmt_kwp(kwp_nova)} kWp",
        f"POTENCIA TOTAL: {_fmt_kwp(kwp_exist + kwp_nova)} kWp",
        f"AMPLIAÇÃO DE {_fmt_kwp(kwp_exist)}KWp para {_fmt_kwp(kwp_exist + kwp_nova)}KWp",
    ]


def gerar_txt_dados(dados, pasta_saida: str) -> str:
    """Gera o arquivo TXT resumo e retorna o caminho absoluto."""
    pasta = Path(pasta_saida)
    pasta.mkdir(parents=True, exist_ok=True)

    nome_arquivo = f"{sanitize_filename_part((dados.titular or '').upper())}_UC_{sanitize_filename_part(dados.codigo_uc)}_DADOS.txt"
    caminho_txt = pasta / nome_arquivo

    linhas = [
        "",
        *_linhas_localizacao(dados),
        "",
        f"ENDEREÇO:{dados.logradouro}, ",
        f"Nº {dados.numero}, BAIRRO: {dados.bairro}, CIDADE: {dados.cidade}, CEP {dados.cep}",
        "",
        "",
        f"UC: {dados.codigo_uc}",
        "",
        "",
        f"PROPRIETARIO: {dados.titular}",
        f"CPF: {dados.cpf_cnpj}",
        "",
        "",
        "",
        "",
        f"TELEFONE: {dados.telefone}",
        f"CELULAR: {dados.celular}",
        "",
        f"EMAIL: {dados.email}",
        "",
        "",
        *_linhas_modulos(dados.paineis),
        "",
        *_linhas_inversores(dados.inversores),
        *_linhas_geracao_existente(dados),
        "",
    ]

    with open(caminho_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))

    print(f"  [step6] OK — TXT de dados gerado: {nome_arquivo}")
    return str(caminho_txt)
