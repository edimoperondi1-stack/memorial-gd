"""
modelos.py
----------
Define os dados de entrada do pipeline como dataclasses Python.
Cada campo corresponde a uma célula da planilha Excel.
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Painel:
    """Um lote de painéis solares (linha da tabela em MD-SOLAR)."""
    quantidade: int
    fabricante: str
    modelo: str
    area_m2: float = 0.0
    potencia_kw: float = 0.0


@dataclass
class Inversor:
    """Um lote de inversores (linha da tabela em MD-SOLAR)."""
    quantidade: int
    fabricante: str
    modelo: str
    potencia_kw: float = 0.0
    tensao_nominal_v: float = 0.0


@dataclass
class UCBeneficiaria:
    """Uma UC beneficiária do sistema de compensação."""
    codigo_uc: str
    titular: str
    cpf_cnpj: str
    endereco: str
    percentual: float


@dataclass
class DadosProjeto:
    """
    Modelo completo de entrada do pipeline.
    Agrupa todos os campos necessários para preencher a planilha
    e gerar o Excel e PDF de saída.
    """

    # ── 1. IDENTIFICAÇÃO DA UC ──────────────────────────────────────────────
    codigo_uc: str                     # ex: "3941140"
    titular: str                       # ex: "APARECIDO MENEZES DOS SANTOS"
    classe: str                        # RESIDENCIAL | COMERCIAL | INDUSTRIAL | RURAL | PODER PÚBLICO | SERVIÇO PÚBLICO
    cpf_cnpj: str                      # apenas dígitos, ex: "36209562191"
    logradouro: str
    numero: str                        # ex: "0" ou "123"
    bairro: str
    cidade: str
    uf: str                            # ex: "MT"
    cep: str                           # apenas dígitos, ex: "78550000"
    email: str
    telefone: str                      # apenas dígitos
    celular: str                       # apenas dígitos

    # ── 2. DADOS DA UC (padrão elétrico) ────────────────────────────────────
    potencia_instalada_kw: float       # kW instalados na UC
    tensao_atendimento_v: str          # ex: "220", "127/220", "220/380"
    tipo_conexao: str                  # MONOFÁSICO | BIFÁSICO | TRIFÁSICO
    tipo_ramal: str                    # AÉREO | SUBTERRÂNEO

    # ── 3. DADOS DA GERAÇÃO ─────────────────────────────────────────────────
    tipo_fonte: str                    # ex: "SOLAR FOTOVOLTAICA"
    tipo_geracao: str                  # ex: "Empregando conversor eletrônico/inversor"
    modalidade: str                    # Compensação local | Autoconsumo remoto | Múltiplas UCs | Geração compartilhada
    potencia_geracao_kwp: float        # kWp do sistema

    # ── 4. DETALHES TÉCNICOS (MD-SOLAR) ─────────────────────────────────────
    tipo_padrao: str                   # ex: "BIFÁSICO" (mesmo que tipo_conexao na maioria dos casos)
    nivel_tensao_v: str                # ex: "220"
    potencia_max_disponivel_kw: float  # kW disponibilizados pela distribuidora
    disjuntor_geral_a: int
    fator_potencia: float              # ex: 0.92
    demanda_contratada_kw: float       # ex: 1.0
    dps_ca_ka: float = 0.0
    disjuntor_ca_a: float = 0.0
    tem_stringbox: bool = False       # Se True, habilita DPS CC e Disjuntor CC
    dps_cc_ka: float = 0.0
    disjuntor_cc_a: float = 0.0

    nivel_tensao_tipo: str = "BAIXA"   # BAIXA | ALTA
    num_fases: int = 2
    cabos_por_fase: int = 1
    bitola_fase_mm2: float = 10.0
    bitola_neutro_mm2: float = 10.0
    bitola_terra_mm2: float = 10.0
    zona: str = "URBANO"               # URBANO | RURAL

    # ── 5. COORDENADAS UTM ──────────────────────────────────────────────────
    fuso: str = ""                     # ex: "21K"
    coord_x_long: float = 0.0
    coord_y_lat: float = 0.0

    # ── 6. TRAFO / CONFIGURAÇÃO EXTRA ───────────────────────────────────────
    potencia_trafo_kw: float = 0.0
    num_hastes: int = 3
    trafo_acoplamento: str = "NÃO"     # SIM | NÃO
    trafo_exclusivo: str = "NÃO"      # SIM | NÃO
    gd_ja_instalado: str = "NÃO"      # SIM | NÃO
    previsao_mes: str = "JANEIRO"      # JANEIRO … DEZEMBRO
    previsao_ano: int = 2026

    # ── 7. EQUIPAMENTOS ─────────────────────────────────────────────────────
    paineis: list = field(default_factory=list)      # lista de Painel
    inversores: list = field(default_factory=list)   # lista de Inversor

    # ── 8. UCs BENEFICIÁRIAS (opcional) ─────────────────────────────────────
    ucs_beneficiarias: list = field(default_factory=list)  # lista de UCBeneficiaria

    # ── 9. RESPONSÁVEL TÉCNICO ───────────────────────────────────────────────
    resp_nome: str = ""
    resp_telefone: str = ""
    resp_email: str = ""

    # ── 10. RELAÇÃO DE CARGA ─────────────────────────────────────────────────
    # Lista de tuplas: (quantidade, equipamento, pot_unitaria_w, fator_demanda)
    carga_instalada: list = field(default_factory=list)

    # ── 11. TIPO DE FORMULÁRIO ───────────────────────────────────────────────
    # Define qual aba de FSA será usada e quais abas vão para o PDF
    # "SOLICITACAO" = minigeração | "FSA MICRO <=10" | "FSA MICRO >10"
    tipo_fsa: str = "SOLICITACAO"

    # ── 12. FORMULÁRIO DE ORÇAMENTO ──────────────────────────────────────────
    # Dict de item_key → valor ("X" ou "SIM"/"NÃO")
    # Se vazio/None, usa os defaults de step1_preencher.FORMULARIO_DEFAULTS
    formulario_items: dict = field(default_factory=dict)

    # ── OBSERVAÇÕES ──────────────────────────────────────────────────────────
    observacoes: str = ""
