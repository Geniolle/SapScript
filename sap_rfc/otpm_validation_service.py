"""
otpm_validation_service.py - Serviço RFC de leitura e validação de Cadeias de Pesquisa (OTPM).
Tabelas consultadas: TPAMA, TPAMAT, T028P.
Garante consultas estritamente READ-ONLY via RFC_READ_TABLE com credenciais unificadas.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from sap_rfc._rfc_common import (
    build_connection_params_for,
    classify_rfc_error,
    find_project_root,
    load_project_env,
    make_read_only_guard,
    read_table,
    SafetyGuard,
)


class TpamaStatus(str, Enum):
    EXISTE = "EXISTE"
    NAO_EXISTE = "NÃO EXISTE"
    ERRO_RFC = "ERRO RFC"


class T028pStatus(str, Enum):
    EXISTE = "EXISTE"
    PARCIAL = "PARCIAL"
    DIVERGENTE = "DIVERGENTE"
    NAO_EXISTE = "NÃO EXISTE"
    ERRO_RFC = "ERRO RFC"


def sanitize_panam(nome_cadeia: str) -> str:
    """Normaliza o nome da cadeia de pesquisa respeitando o limite técnico de 20 caracteres da TPAMA-PANAM.

    Remove acentos e caracteres especiais indesejados sem truncamento silencioso descontrolado.
    """
    if not nome_cadeia:
        return ""
    n_cad = "".join(
        c for c in unicodedata.normalize("NFKD", str(nome_cadeia)) if not unicodedata.combining(c)
    )
    n_cad = re.sub(r"[^A-Za-z0-9_\-\.\/ ]", "", n_cad)
    return n_cad.strip().upper()[:20]


@dataclass
class TpamaRecord:
    appli: str
    panam: str
    regex: str
    mapex: str
    note: str = ""


@dataclass
class T028pRecord:
    bukrs: str
    hbkid: str
    hktid: str
    intag: str
    panam: str
    targfi: str
    vgext: str
    vozpm: str
    enabled: str
    partn: str = ""
    prefix: str = ""
    idflag: str = ""

    def key_tuple(self) -> Tuple[str, str, str, str, str, str, str, str]:
        """Tuplo identificador do critério chave de atribuição (compatível com os scripts existentes)."""
        return (
            self.bukrs.strip().upper(),
            self.hbkid.strip().upper(),
            self.hktid.strip().upper(),
            self.vgext.strip().upper(),
            self.vozpm.strip().upper(),
            self.panam.strip().upper(),
            self.targfi.strip().upper(),
            self.prefix.strip().upper(),
        )


@dataclass
class TpamaValidationResult:
    panam: str
    status: TpamaStatus
    record: Optional[TpamaRecord] = None
    message: str = ""


@dataclass
class T028pValidationResult:
    panam: str
    status: T028pStatus
    total_esperado: int = 0
    total_encontrado: int = 0
    records_encontrados: List[T028pRecord] = field(default_factory=list)
    chaves_em_falta: List[Tuple[str, ...]] = field(default_factory=list)
    divergencias: List[str] = field(default_factory=list)
    message: str = ""


class OtpmValidationService:
    """Serviço RFC para auditoria e validação de cadeias de pesquisa OTPM."""

    ALLOWED_TABLES = ("TPAMA", "TPAMAT", "T028P")

    def __init__(self, environment: str = "DEV", connection: Any = None):
        self.environment = str(environment or "DEV").strip().upper()
        self._external_conn = connection
        self._conn = connection
        self._guard = make_read_only_guard(allowed_tables=self.ALLOWED_TABLES)

    def _get_connection(self):
        if self._conn is not None:
            return self._conn

        root = find_project_root()
        load_project_env(root)
        params = build_connection_params_for(self.environment)
        from pyrfc import Connection

        self._conn = Connection(**params)
        return self._conn

    def close(self):
        if self._external_conn is None and self._conn is not None:
            try:
                self._conn.close()
            except Exception:
                pass
            self._conn = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def check_tpama(self, nome_cadeia: str) -> TpamaValidationResult:
        """Verifica a existência exata da cadeia na TPAMA (APPLI='EBST') via RFC."""
        panam = sanitize_panam(nome_cadeia)
        if not panam:
            return TpamaValidationResult(
                panam="",
                status=TpamaStatus.NAO_EXISTE,
                message="Nome de cadeia vazio para validação.",
            )

        try:
            conn = self._get_connection()
            rows_tpama = read_table(
                conn,
                self._guard,
                table_name="TPAMA",
                fields=["APPLI", "PANAM", "REGEX", "MAPEX"],
                options=[
                    {"TEXT": f"APPLI = 'EBST' AND PANAM = '{panam}'"}
                ],
                rowcount=1,
            )

            if not rows_tpama:
                return TpamaValidationResult(
                    panam=panam,
                    status=TpamaStatus.NAO_EXISTE,
                    message=f"Cadeia '{panam}' não encontrada na TPAMA.",
                )

            r = rows_tpama[0]
            appli, p_name, regex, mapex = (
                r[0].strip(),
                r[1].strip(),
                r[2].strip() if len(r) > 2 else "",
                r[3].strip() if len(r) > 3 else "",
            )

            # Obter texto de TPAMAT (opcional, melhor esforço)
            note = ""
            try:
                rows_desc = read_table(
                    conn,
                    self._guard,
                    table_name="TPAMAT",
                    fields=["NOTE"],
                    options=[
                        {"TEXT": f"APPLI = 'EBST' AND PANAM = '{panam}'"}
                    ],
                    rowcount=1,
                )
                if rows_desc:
                    note = rows_desc[0][0].strip()
            except Exception:
                pass

            record = TpamaRecord(
                appli=appli,
                panam=p_name,
                regex=regex,
                mapex=mapex,
                note=note,
            )
            return TpamaValidationResult(
                panam=panam,
                status=TpamaStatus.EXISTE,
                record=record,
                message=f"Cadeia '{panam}' confirmada na TPAMA.",
            )

        except Exception as exc:
            err_code, err_msg = classify_rfc_error(exc)
            return TpamaValidationResult(
                panam=panam,
                status=TpamaStatus.ERRO_RFC,
                message=f"Erro RFC ao consultar TPAMA ({err_code}): {err_msg}",
            )

    def fetch_t028p_by_panam(self, nome_cadeia: str) -> Tuple[Optional[List[T028pRecord]], str]:
        """Lê todas as entradas da T028P para uma dada cadeia via RFC.

        Retorna (lista_records, mensagem). Em caso de erro RFC, lista_records é None.
        """
        panam = sanitize_panam(nome_cadeia)
        if not panam:
            return [], "Nome de cadeia vazio."

        try:
            conn = self._get_connection()
            fields = [
                "BUKRS",
                "HBKID",
                "HKTID",
                "INTAG",
                "PANAM",
                "TARGFI",
                "VGEXT",
                "VOZPM",
                "ENABLED",
                "PARTN",
                "PREFIX",
                "IDFLAG",
            ]
            rows = read_table(
                conn,
                self._guard,
                table_name="T028P",
                fields=fields,
                options=[{"TEXT": f"PANAM = '{panam}'"}],
                rowcount=0,
            )

            records: List[T028pRecord] = []
            for r in rows:
                rec = T028pRecord(
                    bukrs=r[0].strip(),
                    hbkid=r[1].strip(),
                    hktid=r[2].strip() if len(r) > 2 else "",
                    intag=r[3].strip() if len(r) > 3 else "",
                    panam=r[4].strip() if len(r) > 4 else "",
                    targfi=r[5].strip() if len(r) > 5 else "",
                    vgext=r[6].strip() if len(r) > 6 else "",
                    vozpm=r[7].strip() if len(r) > 7 else "",
                    enabled=r[8].strip() if len(r) > 8 else "",
                    partn=r[9].strip() if len(r) > 9 else "",
                    prefix=r[10].strip() if len(r) > 10 else "",
                    idflag=r[11].strip() if len(r) > 11 else "",
                )
                records.append(rec)

            return records, f"{len(records)} atribuição(ões) lida(s) na T028P para PANAM='{panam}'."

        except Exception as exc:
            err_code, err_msg = classify_rfc_error(exc)
            return None, f"Erro RFC ao consultar T028P ({err_code}): {err_msg}"

    def validate_t028p_assignments(
        self,
        nome_cadeia: str,
        expected_records: List[Dict[str, Any]],
    ) -> T028pValidationResult:
        """Valida se o conjunto esperado de atribuições para uma cadeia existe na T028P.

        Classificação:
        - EXISTE: Todos os registos esperados existem, valores equivalentes e ENABLED='X'.
        - PARCIAL: Algumas atribuições existem, mas faltam outras.
        - DIVERGENTE: As chaves existem mas atributos críticos (ex.: ENABLED, INTAG) divergem.
        - NAO_EXISTE: Nenhuma atribuição encontrada.
        - ERRO_RFC: Falha de comunicação/consulta RFC.
        """
        panam = sanitize_panam(nome_cadeia)
        total_expected = len(expected_records)

        records_found, msg = self.fetch_t028p_by_panam(panam)
        if records_found is None:
            return T028pValidationResult(
                panam=panam,
                status=T028pStatus.ERRO_RFC,
                total_esperado=total_expected,
                message=msg,
            )

        if not records_found:
            return T028pValidationResult(
                panam=panam,
                status=T028pStatus.NAO_EXISTE,
                total_esperado=total_expected,
                total_encontrado=0,
                message=f"Nenhuma atribuição encontrada na T028P para PANAM='{panam}'.",
            )

        # Mapa das chaves encontradas: key_tuple -> T028pRecord
        found_map: Dict[Tuple[str, ...], T028pRecord] = {
            rec.key_tuple(): rec for rec in records_found
        }

        missing_keys: List[Tuple[str, ...]] = []
        divergences: List[str] = []

        for exp in expected_records:
            exp_key = (
                str(exp.get("BUKRS", "")).strip().upper(),
                str(exp.get("HBKID", "")).strip().upper(),
                str(exp.get("HKTID", "")).strip().upper(),
                str(exp.get("VGEXT", "")).strip().upper(),
                str(exp.get("VOZPM", "")).strip().upper(),
                panam,
                str(exp.get("TARGFI", "")).strip().upper(),
                str(exp.get("PREFIX", "")).strip().upper(),
            )

            if exp_key not in found_map:
                missing_keys.append(exp_key)
            else:
                actual = found_map[exp_key]
                # Verificar divergências em atributos secundários ou estado ENABLED, INTAG, PARTN, IDFLAG
                exp_enabled = str(exp.get("ENABLED", "X")).strip().upper() or "X"
                if actual.enabled.strip().upper() != exp_enabled:
                    divergences.append(
                        f"Atribuição {exp_key[:4]}: ENABLED esperado='{exp_enabled}', encontrado='{actual.enabled}'"
                    )
                exp_intag = str(exp.get("INTAG", "")).strip()
                if exp_intag and actual.intag.strip() != exp_intag:
                    divergences.append(
                        f"Atribuição {exp_key[:4]}: INTAG esperado='{exp_intag}', encontrado='{actual.intag}'"
                    )
                if "PARTN" in exp:
                    exp_partn = str(exp.get("PARTN", "") or "").strip().upper()
                    if actual.partn.strip().upper() != exp_partn:
                        divergences.append(
                            f"Atribuição {exp_key[:4]}: PARTN esperado='{exp_partn}', encontrado='{actual.partn}'"
                        )
                if "IDFLAG" in exp:
                    exp_idflag = str(exp.get("IDFLAG", "") or "").strip().upper()
                    if actual.idflag.strip().upper() != exp_idflag:
                        divergences.append(
                            f"Atribuição {exp_key[:4]}: IDFLAG esperado='{exp_idflag}', encontrado='{actual.idflag}'"
                        )

        found_count = len(records_found)
        matched_expected_count = total_expected - len(missing_keys)

        if matched_expected_count == total_expected and not divergences:
            status = T028pStatus.EXISTE
            msg_res = f"Todas as {total_expected} atribuição(ões) confirmadas na T028P."
        elif matched_expected_count > 0 and len(missing_keys) > 0:
            status = T028pStatus.PARCIAL
            msg_res = f"Configuração parcial: {matched_expected_count}/{total_expected} encontradas. Faltam {len(missing_keys)} atribuição(ões)."
        elif matched_expected_count == total_expected and divergences:
            status = T028pStatus.DIVERGENTE
            msg_res = f"Atribuições existem com divergências: {'; '.join(divergences[:2])}"
        elif matched_expected_count == 0:
            status = T028pStatus.NAO_EXISTE
            msg_res = f"Nenhuma das {total_expected} atribuições esperadas foi encontrada na T028P."
        else:
            status = T028pStatus.PARCIAL
            msg_res = f"Configuração incompleta ou divergente."

        return T028pValidationResult(
            panam=panam,
            status=status,
            total_esperado=total_expected,
            total_encontrado=found_count,
            records_encontrados=records_found,
            chaves_em_falta=missing_keys,
            divergencias=divergences,
            message=msg_res,
        )
