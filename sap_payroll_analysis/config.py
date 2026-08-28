"""Configuração central do diagnóstico Payroll -> FI.

Todos os valores monetários usam `Decimal`. Nunca usar `float` para
cálculos financeiros.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal


# ---------------------------------------------------------------------------
# Parâmetros do caso em investigação (valores por omissão).
# ---------------------------------------------------------------------------

# Empresa "real" confirmada no SAP (a empresa "2010" indicada era rótulo;
# a conta 23120000 / 06-2026 move-se na empresa 1010).
EMPRESA = "1010"
ANO = 2026
MES = 6
CONTA = "23120000"
MOEDA = "EUR"

# Execuções de posting do Payroll identificadas na PCP0 para 06/2026.
POSTING_RUNS: list[str] = [
    "0000001296",
    "0000001298",
    "0000001299",
    "0000001301",
    "0000001302",
    "0000001304",
]

# Fase 2: run analisado prioritariamente (empresa 1010, conta 23120000,
# 727.258,35 EUR). O run 1299 tem o mesmo valor — NÃO somar 1298 + 1299.
PRIMARY_RUN = "0000001298"

# Chave de operação (transaction key) da linha de posting HR para contas do
# Razão. Vista em PPDIT.KTOSL; usada na determinação de contas (T030).
HR_POSTING_KEY = "HRF"

# Rubricas de referência que compõem (parcialmente) o valor de RH.
WAGE_TYPES_REFERENCIA: list[str] = ["/558", "/559"]

# Valores informados pelo utilizador para reconciliação.
VALOR_RH_REFERENCIA = Decimal("724046.64")
VALOR_FI_REFERENCIA = Decimal("727258.35")
DIFERENCA_REFERENCIA = Decimal("3211.71")

# Tolerância de arredondamento nas comparações de totais.
TOLERANCIA = Decimal("0.01")


# ---------------------------------------------------------------------------
# Segurança: whitelists explícitas.
# ---------------------------------------------------------------------------

#: Só estas tabelas podem ser lidas. Qualquer outra exige inclusão manual aqui.
READ_ONLY_TABLE_WHITELIST: frozenset[str] = frozenset(
    {
        # Payroll posting
        "PEVST",
        "PPDHD",
        "PPDIT",
        "PPOIX",
        "PPDIX",
        "PPOPX",
        "PPDST",   # Fase 4.2 — split de custeio da linha PPDIT (tem WRBTR)
        "PPDSH",   # Fase 4.2 — cabeçalho/estado do split PPDST (sem valor)
        "T52POSTRUN",
        # Determinação de contas do Payroll (customizing, só leitura)
        "T52EK",   # atribuição conta simbólica -> conta do Razão
        "T52EL",   # atribuição rubrica -> conta simbólica
        "T52EZ",   # substituição de conta por empresa
        "T52E5",   # variante: conta simbólica -> conta (por regra)
        "T030",    # determinação de contas automática (FI)
        "T52DZ",   # texto/atributos da conta simbólica
        "T52OKT",  # texto da conta simbólica
        "T52OKK",  # conta simbólica (definição)
        # Fase 5 — programa de pagamentos (Payment Medium / REGU*). Só leitura.
        "REGUH",   # dados de liquidação por beneficiário/run de pagamento
        "REGUP",   # itens processados pelo programa de pagamentos
        "REGUV",   # controlo/parâmetros do run de pagamento (F110)
        "REGUHM",  # liquidação cross-company
        "REGUT",   # ficheiros DME/TemSe por run
        "LFA1",    # mestre de fornecedores (para relação PERNR<->LIFNR, se existir)
        "LFB1",    # fornecedor por empresa (idem)
        "PA0002",  # dados pessoais (só para cruzamento de identidade, se preciso)
        # FI
        "BKPF",
        "BSEG",
        "BSIS",  # partidas individuais de contas do Razão (partidas em aberto)
        "BSAS",  # partidas individuais de contas do Razão (compensadas)
        "ACDOCA",
        "FAGLFLEXA",
        # DDIC (somente leitura de metadados)
        "DD02L",
        "DD02T",
        "DD03L",
        "DD03T",
        "DD04T",
        # textos de conta do Razão (contexto do relatório)
        "SKAT",
        "SKB1",
        # Fase 3 — cluster/diretório de Payroll (tudo transparente, só leitura)
        "HRPY_RGDIR",       # diretório de resultados de Payroll (cópia transparente do RGDIR)
        "HRPY_RGDIR_TEMP",
        "HRPY_WPBP",        # split WPBP por PERNR/SEQNR (empresa/período)
        "HRPY_GROUPING",
        "PA0001",           # atribuição organizacional (ABKRS, BUKRS)
        "T549A",            # área de contabilização -> modificador de período (PERMO)
        "T549Q",            # definição de períodos de contabilização
        "T500L",            # país/agrupamento -> MOLGA, RELID do cluster
        "T52RELID",         # RELIDs do PCL2
        "T100",             # textos de mensagens (para documentar exceções)
        # Fase 3 — tabelas transparentes de resultados de Payroll ("Payroll
        # Results Tables"). Descobertas por DDIC; neste sistema estão VAZIAS
        # (framework inactivo) mas são lidas para o catálogo automático.
        "P2RX_RT", "P2RX_RT_PERSON", "P2RX_CRT", "P2RX_BT", "P2RX_WPBP",
        "P2RX_VERSC", "P2RX_ARRRS", "P2RX_DDNTK", "P2RX_GRT",
        "HRPADNLP_P2RX_RT", "HRPADNLP_P2RX_BT",
    }
)

#: Só estas funções RFC podem ser invocadas. Todas de leitura.
ALLOWED_RFC_FUNCTIONS: frozenset[str] = frozenset(
    {
        "RFC_PING",
        "RFC_READ_TABLE",
        "RFC_GET_FUNCTION_INTERFACE",
        "DDIF_FIELDINFO_GET",
        # BAPI de LEITURA de saldos por período de conta do Razão.
        # É um "getter": não altera nada, não faz COMMIT.
        "BAPI_GL_GETGLACCPERIODBALANCES",
        # Fase 3 — leitura do resultado de Payroll (cluster PCL2). FM de
        # IMPORT/leitura: não grava, não faz COMMIT. Interface verificada por
        # RFC_GET_FUNCTION_INTERFACE (só IMPORT/CHANGING/EXPORT, sem update).
        # NOTA: neste sistema devolve DA300 ("No active nametab") em contexto
        # RFC stateless — o pacote tenta e degrada graciosamente, registando a
        # limitação. Não há alternativa de escrita (nem é usada).
        # (CU_READ_RGDIR foi investigada — mesmo DA300 — e NÃO é usada: o
        #  diretório lê-se da tabela transparente HRPY_RGDIR.)
        "PYXX_READ_PAYROLL_RESULT",
    }
)

# Fase 3 — período de Payroll do caso (ABKRS Z2, PERMO 01, mensal).
PAYROLL_PERIOD_YYYYMM = "202606"

#: Prefixos de variáveis de ambiente tentados, por ordem, para a ligação RFC.
#: O primeiro totalmente preenchido é usado. `SAP_R3_*` é o oficial deste
#: diagnóstico; `SAP_DEV_*` é aceite como fallback (mesmo host 10.1.1.101).
ENV_PREFIXES: tuple[str, ...] = ("SAP_R3_", "SAP_DEV_")

REQUIRED_ENV_SUFFIXES: tuple[str, ...] = ("USER", "PASSWD", "ASHOST", "SYSNR", "CLIENT")


@dataclass(frozen=True)
class AnalysisParams:
    """Parâmetros efectivos de uma execução do diagnóstico."""

    empresa: str = EMPRESA
    ano: int = ANO
    mes: int = MES
    conta: str = CONTA
    moeda: str = MOEDA
    posting_runs: tuple[str, ...] = tuple(POSTING_RUNS)
    primary_run: str = PRIMARY_RUN
    wage_types_referencia: tuple[str, ...] = tuple(WAGE_TYPES_REFERENCIA)
    valor_rh_referencia: Decimal = VALOR_RH_REFERENCIA
    valor_fi_referencia: Decimal = VALOR_FI_REFERENCIA
    diferenca_referencia: Decimal = DIFERENCA_REFERENCIA
    hr_posting_key: str = HR_POSTING_KEY
    tolerancia: Decimal = TOLERANCIA
    page_size: int = 5000

    @property
    def primary_run_10(self) -> str:
        return pad_run(self.primary_run)

    @property
    def conta_10(self) -> str:
        """Conta do Razão em formato interno SAP (CHAR10, zeros à esquerda)."""
        return pad_account(self.conta)

    @property
    def periodo_label(self) -> str:
        return f"{self.mes:02d}/{self.ano}"

    @property
    def poper(self) -> str:
        """Período contabilístico em formato NUMC3 (ex.: '006')."""
        return f"{self.mes:03d}"

    @property
    def gjahr(self) -> str:
        return f"{self.ano:04d}"


DEFAULTS = AnalysisParams()


def pad_account(account: str) -> str:
    """Normaliza um número de conta do Razão para CHAR10 com zeros à esquerda.

    Aceita já com zeros ou sem. Não numérico é devolvido em maiúsculas sem padding.
    """
    raw = str(account or "").strip().upper()
    if raw.isdigit():
        return raw.zfill(10)
    return raw


def pad_run(run: str) -> str:
    """Normaliza um nº de posting run para 10 dígitos com zeros à esquerda."""
    raw = str(run or "").strip()
    return raw.zfill(10) if raw.isdigit() else raw
