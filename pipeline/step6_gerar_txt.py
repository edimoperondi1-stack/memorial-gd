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

    # Módulos — lista todos
    linhas_modulos = []
    if dados.paineis:
        for p in dados.paineis:
            pot_w = int(round(float(p.potencia_kw) * 1000))
            qtd = getattr(p, "quantidade", None) or 1
            linhas_modulos.append(f"MODELO MODULOS: {qtd}x {p.modelo} {pot_w}W ")
            linhas_modulos.append(f"FABRICANTE MODULOS: {p.fabricante}  ")
            linhas_modulos.append("")
        # remove a linha em branco final (será adicionada depois)
        if linhas_modulos and linhas_modulos[-1] == "":
            linhas_modulos.pop()
    else:
        linhas_modulos = ["MODELO MODULOS: ", "FABRICANTE MODULOS:  "]

    # Inversores — lista todos
    linhas_inversores = []
    if dados.inversores:
        for inv in dados.inversores:
            qtd = getattr(inv, "quantidade", None) or 1
            linhas_inversores.append(f"MODELO INVERSOR: {qtd}x {inv.modelo}   ")
            linhas_inversores.append(f"FABRICANTE INVERSOR:  {inv.fabricante}   ")
            linhas_inversores.append("")
        if linhas_inversores and linhas_inversores[-1] == "":
            linhas_inversores.pop()
    else:
        linhas_inversores = ["MODELO INVERSOR:    ", "FABRICANTE INVERSOR:     "]

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
        *linhas_modulos,
        "",
        *linhas_inversores,
        "",
    ]

    with open(caminho_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))

    print(f"  [step6] OK — TXT de dados gerado: {nome_arquivo}")
    return str(caminho_txt)
