"""
step6_gerar_txt.py
------------------
Gera um arquivo TXT com os dados resumidos do projeto, no formato usado
para auxiliar o preenchimento do restante da documentação.

Formato de saída:

    LOCALIZAÇÃO UTM X: <x>, Y: <y>

    ENDEREÇO:<logradouro>,
    Nº <numero>, BAIRRO: <bairro>, CIDADE: <cidade>, CEP <cep>


    UC: <codigo_uc>


    PROPRIETARIO: <titular>
    CPF: <cpf_cnpj>




    TELEFONE: <telefone>
    CELULAR: <celular>

    EMAIL: <email>


    MODELO MODULOS: <modelo> <potencia_w>W
    FABRICANTE MODULOS: <fabricante>

    MODELO INVERSOR: <modelo>
    FABRICANTE INVERSOR: <fabricante>
"""

from pathlib import Path


def gerar_txt_dados(dados, pasta_saida: str) -> str:
    """Gera o arquivo TXT resumo e retorna o caminho absoluto."""
    pasta = Path(pasta_saida)
    pasta.mkdir(parents=True, exist_ok=True)

    nome_arquivo = f"{(dados.titular or '').upper().strip()}_UC_{dados.codigo_uc}_DADOS.txt"
    caminho_txt = pasta / nome_arquivo

    # Coordenadas (podem vir como UTM nos campos coord_x_long/coord_y_lat)
    x = getattr(dados, "coord_x_long", 0) or 0
    y = getattr(dados, "coord_y_lat", 0) or 0

    # Primeiro módulo e inversor (se houver)
    painel = dados.paineis[0] if dados.paineis else None
    inv = dados.inversores[0] if dados.inversores else None

    if painel:
        pot_w = int(round(float(painel.potencia_kw) * 1000))
        modelo_mod = f"{painel.modelo} {pot_w}W"
        fab_mod = painel.fabricante
    else:
        modelo_mod = ""
        fab_mod = ""

    if inv:
        modelo_inv = inv.modelo
        fab_inv = inv.fabricante
    else:
        modelo_inv = ""
        fab_inv = ""

    linhas = [
        "",
        f"LOCALIZAÇÃO UTM X: {x}, Y: {y}",
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
        f"MODELO MODULOS: {modelo_mod} ",
        f"FABRICANTE MODULOS: {fab_mod}  ",
        "",
        f"MODELO INVERSOR: {modelo_inv}   ",
        f"FABRICANTE INVERSOR:  {fab_inv}   ",
        "",
    ]

    with open(caminho_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))

    print(f"  [step6] OK — TXT de dados gerado: {nome_arquivo}")
    return str(caminho_txt)
