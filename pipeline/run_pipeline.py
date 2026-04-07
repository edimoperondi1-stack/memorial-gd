"""
run_pipeline.py
---------------
Orquestrador do pipeline completo de geração de documentos GD.

Recebe um DadosProjeto e executa os 4 passos em sequência:
  1. Preencher o template Excel com os dados do projeto
  2. Recalcular todas as fórmulas via LibreOffice
  3. Gerar o Excel de saída (aba SAIDA, valores apenas)
  4. Gerar o PDF de saída (abas corretas conforme tipo_fsa)

Uso básico:
    from pipeline.run_pipeline import executar_pipeline
    from pipeline.modelos import DadosProjeto, Painel, Inversor

    dados = DadosProjeto(
        codigo_uc="3941140",
        titular="JOÃO DA SILVA",
        ...
    )
    resultado = executar_pipeline(dados, pasta_saida="/caminho/para/saida")
    print(resultado["xlsx"])   # caminho do Excel gerado
    print(resultado["pdf"])    # caminho do PDF gerado
"""

import os
import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
import tempfile
from pathlib import Path

# Garante que o diretório do pipeline está no path
_PIPELINE_DIR = Path(__file__).parent
if str(_PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(_PIPELINE_DIR))

from modelos import DadosProjeto
from step1_preencher import preencher_template
from step2_recalcular import recalcular, verificar_campos_criticos
from step3_gerar_xlsx import gerar_xlsx, validar_xlsx_saida
from step4_gerar_pdf import gerar_pdf, validar_pdf
from step5_gerar_procuracao import gerar_procuracao_pdf


def executar_pipeline(
    dados: DadosProjeto,
    pasta_saida: str,
    pasta_temp: str = None,
    validar: bool = True,
    timeout_recalc: int = 90,
) -> dict:
    """
    Executa o pipeline completo de geração de documentos.

    Args:
        dados: DadosProjeto com todos os campos preenchidos.
        pasta_saida: pasta onde os arquivos finais serão salvos.
        pasta_temp: pasta para arquivos temporários (padrão: /tmp).
        validar: se True, executa validações após cada step.
        timeout_recalc: segundos máximos para o recálculo do LibreOffice.

    Returns:
        dict com:
            "xlsx": caminho do arquivo .xlsx gerado
            "pdf": caminho do arquivo .pdf gerado
            "relatorio": dict com resultados de cada step
            "ok": True se tudo correu bem

    Raises:
        ValueError: se os dados de entrada forem inválidos.
        FileNotFoundError: se o template não for encontrado.
        RuntimeError: se algum step crítico falhar.
    """
    print(f"\n{'='*60}")
    print(f"  PIPELINE GD — {dados.titular} / UC {dados.codigo_uc}")
    print(f"{'='*60}")

    relatorio = {
        "step1": None,
        "step2": None,
        "step3": None,
        "step4": None,
        "step5": None,
    }

    # ── Pasta temporária ──────────────────────────────────────────
    if pasta_temp is None:
        pasta_temp = tempfile.mkdtemp(prefix="pipeline_gd_")
    else:
        Path(pasta_temp).mkdir(parents=True, exist_ok=True)

    # ── STEP 1: Preencher template ────────────────────────────────
    print(f"\n[STEP 1] Preenchendo template...")
    caminho_preenchido = preencher_template(dados, pasta_saida=pasta_temp)
    relatorio["step1"] = {"caminho": caminho_preenchido, "ok": True}
    print(f"[STEP 1] OK")

    # ── STEP 2: Recalcular fórmulas ───────────────────────────────
    print(f"\n[STEP 2] Recalculando fórmulas...")
    relatorio_recalc = recalcular(caminho_preenchido, timeout=timeout_recalc)

    if relatorio_recalc["status"] == "erros_criticos":
        raise RuntimeError(
            f"Step 2 — Erros críticos no recálculo:\n{relatorio_recalc['erros_criticos']}"
        )

    campos_criticos = verificar_campos_criticos(caminho_preenchido)
    relatorio["step2"] = {
        "recalc": relatorio_recalc,
        "campos_criticos": campos_criticos,
        "ok": len(campos_criticos.get("_problemas", [])) == 0,
    }

    if relatorio["step2"]["ok"]:
        print(f"[STEP 2] OK")
    else:
        print(f"[STEP 2] AVISO — Campos com atenção: {campos_criticos['_problemas']}")

    # ── STEP 3: Gerar Excel de saída ──────────────────────────────
    print(f"\n[STEP 3] Gerando Excel de saída...")
    caminho_xlsx = gerar_xlsx(
        caminho_preenchido=caminho_preenchido,
        pasta_saida=pasta_saida,
        nome_titular=dados.titular,
        codigo_uc=dados.codigo_uc,
    )

    validacao_xlsx = None
    if validar:
        validacao_xlsx = validar_xlsx_saida(caminho_xlsx)
        if not validacao_xlsx["ok"]:
            print(f"[STEP 3] AVISO — Problemas no Excel: {validacao_xlsx['problemas']}")
        else:
            print(f"[STEP 3] OK (validação passou)")

    relatorio["step3"] = {
        "caminho": caminho_xlsx,
        "validacao": validacao_xlsx,
        "ok": (not validar) or (validacao_xlsx and validacao_xlsx["ok"]),
    }

    # ── STEP 4: Gerar PDF ─────────────────────────────────────────
    print(f"\n[STEP 4] Gerando PDF...")
    caminho_pdf = gerar_pdf(
        caminho_preenchido=caminho_preenchido,
        pasta_saida=pasta_saida,
        nome_titular=dados.titular,
        codigo_uc=dados.codigo_uc,
        tipo_fsa=dados.tipo_fsa,
    )

    validacao_pdf = None
    if validar:
        validacao_pdf = validar_pdf(caminho_pdf, campos_esperados=[
            dados.titular.upper(),
            dados.codigo_uc,
        ])
        if not validacao_pdf["ok"]:
            print(f"[STEP 4] AVISO — Campos não encontrados no PDF: {validacao_pdf['faltando']}")
        else:
            print(f"[STEP 4] OK ({validacao_pdf['num_paginas']} páginas)")

    relatorio["step4"] = {
        "caminho": caminho_pdf,
        "validacao": validacao_pdf,
        "ok": (not validar) or (validacao_pdf and validacao_pdf["ok"]),
    }

    # ── STEP 5: Gerar Procuração ─────────────────────────────────────────
    print(f"\n[STEP 5] Gerando PDF da Procuração...")
    try:
        caminho_procuracao = gerar_procuracao_pdf(dados, pasta_saida=pasta_saida)
        relatorio["step5"] = {
            "caminho": caminho_procuracao,
            "ok": True,
        }
    except Exception as e:
        print(f"[STEP 5] AVISO — Falha ao gerar procuração: {e}")
        caminho_procuracao = None
        relatorio["step5"] = {
            "error": str(e),
            "ok": False,
        }

    # ── Resultado final ───────────────────────────────────────────
    tudo_ok = all(v["ok"] for k, v in relatorio.items() if v is not None and k != "step5")

    print(f"\n{'='*60}")
    if tudo_ok:
        print(f"  PIPELINE CONCLUIDO COM SUCESSO")
    else:
        print(f"  PIPELINE CONCLUIDO COM AVISOS")
    print(f"  Excel: {Path(caminho_xlsx).name}")
    print(f"  PDF:   {Path(caminho_pdf).name}")
    if caminho_procuracao:
        print(f"  Procuracao: {Path(caminho_procuracao).name}")
    print(f"{'='*60}\n")

    return {
        "xlsx": caminho_xlsx,
        "pdf": caminho_pdf,
        "procuracao": caminho_procuracao,
        "relatorio": relatorio,
        "ok": tudo_ok,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Bloco de teste (executa quando chamado diretamente com: python run_pipeline.py)
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    from modelos import Painel, Inversor

    # Dados do caso de teste real (Aparecido Menezes dos Santos)
    dados_teste = DadosProjeto(
        # Identificação
        codigo_uc="3941140",
        titular="APARECIDO MENEZES DOS SANTOS",
        classe="RESIDENCIAL",
        cpf_cnpj="36209562191",
        logradouro="RUA CORONEL JOÃO PESSOA",
        numero="0",
        bairro="SENHOR DOS PASSOS",
        cidade="SINOP",
        uf="MT",
        cep="78550000",
        email="",
        telefone="66992056543",
        celular="66992056543",

        # Padrão elétrico
        potencia_instalada_kw=6.0,
        tensao_atendimento_v="220",
        tipo_conexao="BIFÁSICO",
        tipo_ramal="AÉREO",

        # Geração
        tipo_fonte="SOLAR FOTOVOLTAICA",
        tipo_geracao="Empregando conversor eletrônico/inversor",
        modalidade="Compensação local",
        potencia_geracao_kwp=6.6,

        # Detalhes técnicos MD-SOLAR
        tipo_padrao="BIFÁSICO",
        nivel_tensao_v="220",
        potencia_max_disponivel_kw=6.0,
        disjuntor_geral_a=40,
        fator_potencia=0.92,
        demanda_contratada_kw=1.0,
        num_fases=2,
        cabos_por_fase=1,
        bitola_fase_mm2=10.0,
        bitola_neutro_mm2=10.0,
        bitola_terra_mm2=10.0,

        # Equipamentos
        paineis=[
            Painel(quantidade=12, fabricante="CANADIAN SOLAR", modelo="CS3N-550MS", area_m2=2.78, potencia_kw=0.55),
        ],
        inversores=[
            Inversor(quantidade=1, fabricante="GROWATT", modelo="MIC 6000TL-X", potencia_kw=6.0, tensao_nominal_v=220),
        ],

        # Relação de carga
        carga_instalada=[
            (1, "Geladeira", 120, 1.0),
            (6, "Lâmpada LED", 9, 1.0),
            (1, "Televisão", 100, 1.0),
            (1, "Chuveiro elétrico", 5500, 0.25),
            (1, "Máquina de lavar", 500, 0.25),
        ],

        # Responsável técnico
        resp_nome="ENGENHEIRO RESPONSÁVEL",
        resp_telefone="66999999999",
        resp_email="engenheiro@empresa.com",

        # Tipo de formulário
        tipo_fsa="SOLICITACAO",
    )

    PASTA_SAIDA = "/tmp/pipeline_gd_saida"
    resultado = executar_pipeline(dados_teste, pasta_saida=PASTA_SAIDA)

    print(f"Arquivos gerados em: {PASTA_SAIDA}")
    print(f"  Excel: {resultado['xlsx']}")
    print(f"  PDF:   {resultado['pdf']}")
    print(f"  Procuracao: {resultado.get('procuracao')}")
    print(f"  OK:    {resultado['ok']}")
