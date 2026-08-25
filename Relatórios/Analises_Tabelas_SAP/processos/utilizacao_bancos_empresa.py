# -*- coding: utf-8 -*-
"""
Análise histórica de utilização de Bancos Empresa no SAP.

Objetivo
--------
1. Ler todos os House Banks configurados em T012.
2. Ler execuções produtivas do programa de pagamentos em REGUV.
3. Ler utilização dos bancos em REGUH.
4. Considerar apenas o método de pagamento configurado, por padrão "S".
5. NÃO contar pagamentos individuais.
6. Contar somente uma utilização por:

       LAUFD + LAUFI + ZBUKR + HBKID

   Ou seja:
   uma execução F110 com 1 ou 500 pagamentos pelo mesmo banco
   conta apenas 1 utilização.

7. Gerar:
   - relatório resumido por Empresa + Banco;
   - relatório por Empresa + Banco + Ano;
   - JSON completo da análise.

Este processo é SOMENTE DE LEITURA.
"""

from __future__ import annotations

import csv
import json
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any


# =============================================================================
# CAMINHOS
# =============================================================================

BASE_DIR = Path(__file__).resolve().parent.parent

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))


# =============================================================================
# PARÂMETROS
# =============================================================================

METODO = "RFC"
SAP_KEY = "S4PCLNT100"
TRANSACTION = "SE16H"

ABRIR_NOVO_MODO = True
FECHAR_MODO_NO_FIM = False

GERAR_JSON = True
GERAR_CSV = False


# -----------------------------------------------------------------------------
# Método de pagamento a analisar
# -----------------------------------------------------------------------------
#
# Neste momento:
#
# S = SEPA-Fornecedor
#
# Se no futuro quiser analisar TODOS os métodos:
#
# METODO_PAGAMENTO = ""
#
# -----------------------------------------------------------------------------

METODO_PAGAMENTO = "S"


# -----------------------------------------------------------------------------
# Limites técnicos
# -----------------------------------------------------------------------------

MAX_ROWS_T001 = 5_000
MAX_ROWS_T012 = 20_000

# REGUV tem uma linha de controlo por execução.
MAX_ROWS_REGUV = 500_000

# REGUH pode ter muitas linhas porque existe uma linha por pagamento.
# Vamos desduplicá-las posteriormente por cabeçalho de execução.
MAX_ROWS_REGUH = 1_000_000


# =============================================================================
# PROCESSO
# =============================================================================

PROCESSO_ID = "utilizacao_bancos_empresa"


def filtros_reguh() -> list[dict[str, str]]:
    filtros: list[dict[str, str]] = []

    if METODO_PAGAMENTO:
        filtros.append(
            {
                "campo": "RZAWE",
                "valor": METODO_PAGAMENTO,
                "opcao": "EQ",
            }
        )

    return filtros


CONSULTAS = [

    # =========================================================================
    # 01 - EMPRESAS
    # =========================================================================
    {
        "nome": "Empresas SAP",
        "tabela": "T001",
        "filtros": [],
        "campos_saida": [
            "BUKRS",
            "BUTXT",
            "LAND1",
            "WAERS",
        ],
        "max_rows": MAX_ROWS_T001,
    },


    # =========================================================================
    # 02 - TODOS OS BANCOS EMPRESA
    # =========================================================================
    #
    # T012 é a nossa população-base.
    #
    # Mesmo que um banco nunca apareça na REGUH, ele deverá aparecer
    # no resultado final com utilização = 0.
    #
    # =========================================================================
    {
        "nome": "Todos os Bancos Empresa configurados",
        "tabela": "T012",
        "filtros": [],
        "campos_saida": [
            "BUKRS",
            "HBKID",
            "BANKS",
            "BANKL",
        ],
        "max_rows": MAX_ROWS_T012,
    },


    # =========================================================================
    # 03 - EXECUÇÕES PRODUTIVAS DO PROGRAMA DE PAGAMENTOS
    # =========================================================================
    #
    # XECHT = X:
    # execução produtiva realizada.
    #
    # LAUFD + LAUFI:
    # identificam a execução.
    #
    # =========================================================================
    {
        "nome": "Execuções produtivas do programa de pagamentos",
        "tabela": "REGUV",
        "filtros": [
            {
                "campo": "XECHT",
                "valor": "X",
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "LAUFD",
            "LAUFI",
            "XECHT",
        ],
        "max_rows": MAX_ROWS_REGUV,
    },


    # =========================================================================
    # 04 - UTILIZAÇÃO DE HOUSE BANK EM REGUH
    # =========================================================================
    #
    # REGUH possui uma linha por pagamento.
    #
    # NÃO vamos contar essas linhas.
    #
    # Posteriormente será criado um SET com:
    #
    # LAUFD
    # LAUFI
    # ZBUKR
    # HBKID
    #
    # Portanto, centenas de pagamentos da mesma execução/banco
    # contam uma única vez.
    #
    # =========================================================================
    {
        "nome": (
            f"Utilização dos Bancos Empresa - método "
            f"{METODO_PAGAMENTO or 'TODOS'}"
        ),
        "tabela": "REGUH",
        "filtros": filtros_reguh(),
        "campos_saida": [
            "LAUFD",
            "LAUFI",
            "ZBUKR",
            "HBKID",
            "HKTID",
            "RZAWE",
            "XVORL",
        ],
        "max_rows": MAX_ROWS_REGUH,
    },
]


PROCESSO = {
    "id": PROCESSO_ID,
    "titulo": "Utilização histórica dos Bancos Empresa",
    "metodo": METODO,
    "sap_key": SAP_KEY,
    "transaction": TRANSACTION,
    "abrir_novo_modo": ABRIR_NOVO_MODO,
    "fechar_modo_no_fim": FECHAR_MODO_NO_FIM,
    "max_rows": MAX_ROWS_REGUH,
    "gerar_json": GERAR_JSON,
    "gerar_csv": GERAR_CSV,
    "consultas": CONSULTAS,
}


# =============================================================================
# UTILITÁRIOS
# =============================================================================

def texto(valor: Any) -> str:
    return str(valor or "").strip()


def ano_laufd(laufd: str) -> int | None:
    """
    LAUFD normalmente vem no formato YYYYMMDD.
    """

    valor = texto(laufd)

    if len(valor) < 4:
        return None

    try:
        ano = int(valor[:4])
    except ValueError:
        return None

    if ano < 1900 or ano > 2200:
        return None

    return ano


def data_formatada(laufd: str) -> str:
    valor = texto(laufd)

    if len(valor) != 8:
        return valor

    try:
        return datetime.strptime(
            valor,
            "%Y%m%d",
        ).strftime("%d/%m/%Y")
    except ValueError:
        return valor


def chave_execucao(
    laufd: str,
    laufi: str,
) -> tuple[str, str]:
    return (
        texto(laufd),
        texto(laufi),
    )


def chave_utilizacao_banco(
    row: dict[str, str],
) -> tuple[str, str, str, str]:
    """
    Esta é a regra central da análise.

    Uma execução conta uma única vez por:

        DATA EXECUÇÃO
        ID EXECUÇÃO
        EMPRESA PAGADORA
        BANCO EMPRESA

    Não interessa quantos fornecedores/pagamentos existam na REGUH.
    """

    return (
        texto(row.get("LAUFD")),
        texto(row.get("LAUFI")),
        texto(row.get("ZBUKR")),
        texto(row.get("HBKID")),
    )


# =============================================================================
# LOCALIZAR JSON GERADO PELO ENGINE
# =============================================================================

def pasta_output() -> Path:
    # BASE_DIR:
    # Relatórios/Analises_Tabelas_SAP
    #
    # Projeto:
    # ROOT/cache/analises_tabelas_sap/...
    root_dir = BASE_DIR.parents[1]

    return (
        root_dir
        / "cache"
        / "analises_tabelas_sap"
        / PROCESSO_ID
    )


def encontrar_json_mais_recente(
    criado_depois_de: float,
) -> Path:
    pasta = pasta_output()

    candidatos = []

    if pasta.exists():
        for path in pasta.glob(
            f"{PROCESSO_ID}_*.json"
        ):
            try:
                if path.stat().st_mtime >= criado_depois_de:
                    candidatos.append(path)
            except OSError:
                continue

    if not candidatos:
        raise RuntimeError(
            "Não encontrei o JSON gerado pela execução atual."
        )

    return max(
        candidatos,
        key=lambda p: p.stat().st_mtime,
    )


# =============================================================================
# EXTRAÇÃO DOS RESULTADOS
# =============================================================================

def resultado_tabela(
    payload: dict[str, Any],
    tabela: str,
) -> list[dict[str, str]]:

    tabela = tabela.upper()

    for resultado in payload.get(
        "results",
        [],
    ):
        if (
            texto(resultado.get("table")).upper()
            == tabela
        ):
            if resultado.get("error"):
                raise RuntimeError(
                    f"Erro ao ler {tabela}: "
                    f"{resultado['error']}"
                )

            return list(
                resultado.get("rows") or []
            )

    raise RuntimeError(
        f"Tabela {tabela} não encontrada "
        "no resultado do processo."
    )


# =============================================================================
# ANÁLISE
# =============================================================================

def analisar(
    payload: dict[str, Any],
) -> dict[str, Any]:

    empresas = resultado_tabela(
        payload,
        "T001",
    )

    bancos = resultado_tabela(
        payload,
        "T012",
    )

    reguv = resultado_tabela(
        payload,
        "REGUV",
    )

    reguh = resultado_tabela(
        payload,
        "REGUH",
    )


    # -------------------------------------------------------------------------
    # Validar possível truncamento
    # -------------------------------------------------------------------------

    avisos: list[str] = []

    if len(reguv) >= MAX_ROWS_REGUV:
        avisos.append(
            "REGUV atingiu o limite configurado "
            f"de {MAX_ROWS_REGUV:,} linhas."
        )

    if len(reguh) >= MAX_ROWS_REGUH:
        avisos.append(
            "REGUH atingiu o limite configurado "
            f"de {MAX_ROWS_REGUH:,} linhas. "
            "A análise histórica pode estar truncada."
        )


    # -------------------------------------------------------------------------
    # Empresas
    # -------------------------------------------------------------------------

    empresas_por_codigo: dict[
        str,
        dict[str, str],
    ] = {}

    for row in empresas:
        bukrs = texto(
            row.get("BUKRS")
        )

        if not bukrs:
            continue

        empresas_por_codigo[bukrs] = row


    # -------------------------------------------------------------------------
    # Execuções produtivas
    # -------------------------------------------------------------------------
    #
    # REGUV-XECHT já foi filtrado no SAP.
    #
    # Guardamos:
    #
    # (LAUFD, LAUFI)
    #
    # -------------------------------------------------------------------------

    execucoes_produtivas: set[
        tuple[str, str]
    ] = set()

    for row in reguv:

        if texto(
            row.get("XECHT")
        ).upper() != "X":
            continue

        chave = chave_execucao(
            row.get("LAUFD", ""),
            row.get("LAUFI", ""),
        )

        if all(chave):
            execucoes_produtivas.add(
                chave
            )


    # -------------------------------------------------------------------------
    # Desduplicação da REGUH
    # -------------------------------------------------------------------------
    #
    # IMPORTANTE:
    #
    # Não contamos linhas de pagamento.
    #
    # Exemplo:
    #
    # LAUFD = 20260819
    # LAUFI = F11001
    # ZBUKR = 2010
    # HBKID = BPI01
    #
    # 200 pagamentos diferentes na REGUH
    #
    # resultado:
    #
    # 1 utilização
    #
    # -------------------------------------------------------------------------

    utilizacoes_unicas: dict[
        tuple[str, str, str, str],
        dict[str, str],
    ] = {}


    for row in reguh:

        laufd = texto(
            row.get("LAUFD")
        )

        laufi = texto(
            row.get("LAUFI")
        )

        bukrs = texto(
            row.get("ZBUKR")
        )

        hbkid = texto(
            row.get("HBKID")
        )


        # Sem banco empresa não serve para esta análise.
        if not hbkid:
            continue


        # Método alvo.
        if METODO_PAGAMENTO:

            metodo = texto(
                row.get("RZAWE")
            ).upper()

            if (
                metodo
                != METODO_PAGAMENTO.upper()
            ):
                continue


        # ---------------------------------------------------------------------
        # Execução produtiva
        # ---------------------------------------------------------------------

        chave_run = (
            laufd,
            laufi,
        )

        if execucoes_produtivas:

            if (
                chave_run
                not in execucoes_produtivas
            ):
                continue

        else:
            # Fallback:
            #
            # se REGUV não puder ser utilizado,
            # eliminamos registros identificados
            # explicitamente como proposta.
            if texto(
                row.get("XVORL")
            ).upper() == "X":
                continue


        chave = chave_utilizacao_banco(
            row
        )

        if not all(chave):
            continue

        utilizacoes_unicas[
            chave
        ] = row


    # -------------------------------------------------------------------------
    # Contagens
    # -------------------------------------------------------------------------

    contagem_ano: dict[
        tuple[str, str, int],
        int,
    ] = defaultdict(int)

    datas_por_banco: dict[
        tuple[str, str],
        list[str],
    ] = defaultdict(list)


    for (
        laufd,
        laufi,
        bukrs,
        hbkid,
    ) in utilizacoes_unicas:

        ano = ano_laufd(
            laufd
        )

        if ano is None:
            continue

        contagem_ano[
            (
                bukrs,
                hbkid,
                ano,
            )
        ] += 1

        datas_por_banco[
            (
                bukrs,
                hbkid,
            )
        ].append(
            laufd
        )


    # -------------------------------------------------------------------------
    # Todos os anos existentes no histórico
    # -------------------------------------------------------------------------

    anos = sorted(
        {
            ano
            for (
                _bukrs,
                _hbkid,
                ano,
            )
            in contagem_ano
        }
    )


    # -------------------------------------------------------------------------
    # Resultado LONGO:
    #
    # Empresa / Banco / Ano / Quantidade
    # -------------------------------------------------------------------------

    por_ano: list[
        dict[str, Any]
    ] = []


    # -------------------------------------------------------------------------
    # Resultado RESUMIDO:
    #
    # Uma linha por banco empresa,
    # inclusive os nunca utilizados.
    # -------------------------------------------------------------------------

    resumo: list[
        dict[str, Any]
    ] = []


    # Evita eventual duplicação em T012.
    bancos_unicos: dict[
        tuple[str, str],
        dict[str, str],
    ] = {}


    for row in bancos:

        bukrs = texto(
            row.get("BUKRS")
        )

        hbkid = texto(
            row.get("HBKID")
        )

        if not bukrs or not hbkid:
            continue

        bancos_unicos[
            (
                bukrs,
                hbkid,
            )
        ] = row


    for (
        bukrs,
        hbkid,
    ), banco in sorted(
        bancos_unicos.items()
    ):

        empresa = empresas_por_codigo.get(
            bukrs,
            {},
        )

        datas = sorted(
            datas_por_banco.get(
                (
                    bukrs,
                    hbkid,
                ),
                [],
            )
        )

        total = sum(
            contagem_ano.get(
                (
                    bukrs,
                    hbkid,
                    ano,
                ),
                0,
            )
            for ano in anos
        )


        row_resumo: dict[
            str,
            Any,
        ] = {
            "EMPRESA": bukrs,
            "DESCRICAO_EMPRESA": texto(
                empresa.get("BUTXT")
            ),
            "PAIS_EMPRESA": texto(
                empresa.get("LAND1")
            ),
            "BANCO_EMPRESA": hbkid,
            "PAIS_BANCO": texto(
                banco.get("BANKS")
            ),
            "CHAVE_BANCO": texto(
                banco.get("BANKL")
            ),
            "METODO_ANALISADO":
                METODO_PAGAMENTO
                or "TODOS",
            "TOTAL_EXECUCOES":
                total,
            "PRIMEIRA_UTILIZACAO":
                data_formatada(
                    datas[0]
                )
                if datas
                else "",
            "ULTIMA_UTILIZACAO":
                data_formatada(
                    datas[-1]
                )
                if datas
                else "",
        }


        # Uma coluna para cada ano.
        for ano in anos:

            quantidade = (
                contagem_ano.get(
                    (
                        bukrs,
                        hbkid,
                        ano,
                    ),
                    0,
                )
            )

            row_resumo[
                str(ano)
            ] = quantidade


            por_ano.append(
                {
                    "EMPRESA":
                        bukrs,
                    "DESCRICAO_EMPRESA":
                        texto(
                            empresa.get(
                                "BUTXT"
                            )
                        ),
                    "BANCO_EMPRESA":
                        hbkid,
                    "ANO":
                        ano,
                    "QTD_EXECUCOES":
                        quantidade,
                    "METODO":
                        METODO_PAGAMENTO
                        or "TODOS",
                }
            )


        resumo.append(
            row_resumo
        )


    return {
        "meta": {
            "generated_at":
                datetime.now().isoformat(
                    timespec="seconds"
                ),
            "metodo_pagamento":
                METODO_PAGAMENTO
                or "TODOS",
            "regra_contagem":
                (
                    "1 utilização por "
                    "LAUFD+LAUFI+ZBUKR+HBKID"
                ),
            "quantidade_bancos_t012":
                len(
                    bancos_unicos
                ),
            "quantidade_execucoes_produtivas":
                len(
                    execucoes_produtivas
                ),
            "quantidade_linhas_reguh_lidas":
                len(
                    reguh
                ),
            "quantidade_utilizacoes_unicas":
                len(
                    utilizacoes_unicas
                ),
            "anos_encontrados":
                anos,
            "avisos":
                avisos,
        },
        "resumo":
            resumo,
        "por_ano":
            por_ano,
    }


# =============================================================================
# GRAVAÇÃO DOS RESULTADOS
# =============================================================================

def gravar_csv(
    path: Path,
    rows: list[dict[str, Any]],
) -> None:

    if not rows:
        return

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with path.open(
        "w",
        newline="",
        encoding="utf-8-sig",
    ) as file_obj:

        writer = csv.DictWriter(
            file_obj,
            fieldnames=list(
                rows[0].keys()
            ),
            delimiter=";",
        )

        writer.writeheader()

        writer.writerows(
            rows
        )


def gravar_resultados(
    analise: dict[str, Any],
) -> list[Path]:

    output_dir = pasta_output()

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )


    json_path = (
        output_dir
        / (
            "utilizacao_bancos_empresa_"
            f"analise_{timestamp}.json"
        )
    )

    json_path.write_text(
        json.dumps(
            analise,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


    resumo_path = (
        output_dir
        / (
            "utilizacao_bancos_empresa_"
            f"resumo_{timestamp}.csv"
        )
    )

    gravar_csv(
        resumo_path,
        analise["resumo"],
    )


    ano_path = (
        output_dir
        / (
            "utilizacao_bancos_empresa_"
            f"por_ano_{timestamp}.csv"
        )
    )

    gravar_csv(
        ano_path,
        analise["por_ano"],
    )


    return [
        json_path,
        resumo_path,
        ano_path,
    ]


# =============================================================================
# TERMINAL
# =============================================================================

def imprimir_resumo(
    analise: dict[str, Any],
) -> None:

    meta = analise[
        "meta"
    ]

    print(
        "\n"
        + "=" * 100
    )

    print(
        "📊 UTILIZAÇÃO HISTÓRICA "
        "DOS BANCOS EMPRESA"
    )

    print(
        "=" * 100
    )

    print(
        "Método analisado            : "
        f"{meta['metodo_pagamento']}"
    )

    print(
        "Bancos configurados T012    : "
        f"{meta['quantidade_bancos_t012']}"
    )

    print(
        "Execuções produtivas REGUV  : "
        f"{meta['quantidade_execucoes_produtivas']}"
    )

    print(
        "Linhas REGUH lidas          : "
        f"{meta['quantidade_linhas_reguh_lidas']}"
    )

    print(
        "Utilizações únicas           : "
        f"{meta['quantidade_utilizacoes_unicas']}"
    )

    print(
        "Regra                       : "
        f"{meta['regra_contagem']}"
    )

    print(
        "Anos                        : "
        + ", ".join(
            str(x)
            for x
            in meta[
                "anos_encontrados"
            ]
        )
    )


    if meta["avisos"]:

        print(
            "\n⚠️ AVISOS"
        )

        for aviso in meta[
            "avisos"
        ]:
            print(
                f" - {aviso}"
            )


    print(
        "\n"
        + "-" * 100
    )

    print(
        "EMPRESA | BANCO | TOTAL | "
        "PRIMEIRA | ÚLTIMA"
    )

    print(
        "-" * 100
    )


    for row in analise[
        "resumo"
    ]:

        print(
            f"{row['EMPRESA']:>6} | "
            f"{row['BANCO_EMPRESA']:<10} | "
            f"{row['TOTAL_EXECUCOES']:>5} | "
            f"{row['PRIMEIRA_UTILIZACAO']:<10} | "
            f"{row['ULTIMA_UTILIZACAO']:<10}"
        )


# =============================================================================
# MAIN
# =============================================================================

def main() -> int:

    import time

    import engine


    # -------------------------------------------------------------------------
    # Evita imprimir centenas de milhares de linhas da REGUH
    # caso o engine atual exponha _print_rows.
    # -------------------------------------------------------------------------

    if hasattr(
        engine,
        "_print_rows",
    ):

        def print_rows_resumido(
            rows,
        ):
            print(
                f"✅ {len(rows)} "
                "registo(s) devolvido(s)."
            )

        engine._print_rows = (
            print_rows_resumido
        )


    inicio = time.time()


    # -------------------------------------------------------------------------
    # Executar leitura SAP
    # -------------------------------------------------------------------------

    rc = engine.executar_processo(
        PROCESSO
    )

    if rc != 0:
        print(
            "\n❌ A extração SAP "
            "terminou com erro."
        )

        return rc


    # -------------------------------------------------------------------------
    # Ler JSON da execução atual
    # -------------------------------------------------------------------------

    raw_json = (
        encontrar_json_mais_recente(
            inicio,
        )
    )

    print(
        "\n📥 A analisar:"
    )

    print(
        raw_json
    )


    payload = json.loads(
        raw_json.read_text(
            encoding="utf-8",
        )
    )


    # -------------------------------------------------------------------------
    # Análise local
    # -------------------------------------------------------------------------

    analise = analisar(
        payload
    )


    imprimir_resumo(
        analise
    )


    outputs = gravar_resultados(
        analise
    )


    print(
        "\n💾 RESULTADOS"
    )

    for path in outputs:
        print(
            f" - {path}"
        )


    print(
        "\n✅ Análise concluída."
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(
        main()
    )