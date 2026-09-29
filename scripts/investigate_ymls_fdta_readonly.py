from __future__ import annotations

import json
import os
from datetime import date, datetime
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from pyrfc import Connection


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "ymls_fdta_readonly_evidence.json"
ALLOWED_RFC = {"RFC_PING", "RFC_SYSTEM_INFO", "DDIF_FIELDINFO_GET", "RFC_READ_TABLE", "RPY_PROGRAM_READ"}


def normalise(value: Any) -> Any:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    if isinstance(value, dict):
        return {key: normalise(item) for key, item in value.items()}
    if isinstance(value, list):
        return [normalise(item) for item in value]
    return value


class ReadOnlySap:
    def __init__(self) -> None:
        load_dotenv(ROOT / ".env", override=False)
        self.conn = Connection(
            user=os.environ["SAP_PRD_USER"],
            passwd=os.environ["SAP_PRD_PASSWD"],
            ashost=os.environ["SAP_PRD_ASHOST"],
            sysnr=os.environ["SAP_PRD_SYSNR"],
            client=os.environ["SAP_PRD_CLIENT"],
            lang=os.getenv("SAP_PRD_LANG", "PT"),
        )

    def call(self, name: str, **kwargs: Any) -> dict[str, Any]:
        if name not in ALLOWED_RFC:
            raise RuntimeError(f"RFC não autorizado pelo modo read-only: {name}")
        return self.conn.call(name, **kwargs)

    def fields(self, table: str) -> list[dict[str, Any]]:
        result = self.call("DDIF_FIELDINFO_GET", TABNAME=table, LANGU="P", ALL_TYPES="X")
        return result.get("DFIES_TAB", [])

    def read(self, table: str, fields: list[str] | None = None, options: list[str] | None = None,
             rowcount: int = 0) -> dict[str, Any]:
        try:
            metadata = self.fields(table)
        except Exception as exc:
            return {"table": table, "options": options or [], "error": f"{type(exc).__name__}: {exc}"}
        available = {str(item.get("FIELDNAME", "")).strip() for item in metadata}
        selected = fields or [str(item.get("FIELDNAME", "")).strip() for item in metadata]
        selected = [field for field in selected if field in available]
        if not selected:
            return {"table": table, "error": "Sem campos válidos", "available_fields": sorted(available)}
        try:
            result = self.call(
                "RFC_READ_TABLE", QUERY_TABLE=table, DELIMITER="\t",
                FIELDS=[{"FIELDNAME": field} for field in selected],
                OPTIONS=[{"TEXT": value} for value in (options or [])], ROWCOUNT=rowcount,
            )
        except Exception as exc:
            return {"table": table, "fields": selected, "options": options or [], "error": f"{type(exc).__name__}: {exc}"}
        rows = []
        for raw in result.get("DATA", []):
            values = str(raw.get("WA", "")).split("\t")
            rows.append(dict(zip(selected, values)))
        return {"table": table, "fields": selected, "options": options or [], "rows": rows}

    def program(self, name: str) -> dict[str, Any]:
        try:
            result = self.call("RPY_PROGRAM_READ", PROGRAM_NAME=name)
            return normalise(result)
        except Exception as exc:
            return {"program": name, "error": f"{type(exc).__name__}: {exc}"}


def main() -> None:
    sap = ReadOnlySap()
    evidence: dict[str, Any] = {"mode": "SAP PRD READ-ONLY", "allowed_rfc": sorted(ALLOWED_RFC)}
    try:
        sap.call("RFC_PING")
        evidence["system"] = sap.call("RFC_SYSTEM_INFO").get("RFCSI_EXPORT", {})

        evidence["transaction"] = {
            "TSTC": sap.read("TSTC", ["TCODE", "PGMNA", "DYPNO", "CINFO"], ["TCODE = 'YMLS'"], 10),
            "TSTCT": sap.read("TSTCT", ["SPRSL", "TCODE", "TTEXT"], ["TCODE = 'YMLS'"], 20),
            "TSTCP": sap.read("TSTCP", None, ["TCODE = 'YMLS'"], 20),
            "TADIR": sap.read("TADIR", ["PGMID", "OBJECT", "OBJ_NAME", "DEVCLASS", "AUTHOR", "MASTERLANG"], ["OBJ_NAME = 'YMLS'"], 50),
            "YMLS_MENU_TSTC": sap.read("TSTC", ["TCODE", "PGMNA", "DYPNO", "CINFO"], ["TCODE = 'YMLS_MENU'"], 10),
            "YMLS_MENU_TSTCT": sap.read("TSTCT", ["SPRSL", "TCODE", "TTEXT"], ["TCODE = 'YMLS_MENU'"], 20),
            "YMLS_MENU_TSTCP": sap.read("TSTCP", None, ["TCODE = 'YMLS_MENU'"], 20),
            "YMLS_MENU_TADIR": sap.read("TADIR", ["PGMID", "OBJECT", "OBJ_NAME", "DEVCLASS", "AUTHOR", "MASTERLANG"], ["OBJ_NAME = 'YMLS_MENU'"], 50),
        }

        evidence["menu"] = {}
        evidence["menu"]["TMENU01"] = sap.read("TMENU01", None, ["TREE_ID = '00001'"], 2000)
        evidence["menu"]["TMENU01T"] = sap.read("TMENU01T", None, ["TREE_ID = '00001'"], 5000)
        evidence["menu"]["TMENU01R_ALL"] = sap.read("TMENU01R", None, None, 20000)

        evidence["ymls_package_objects"] = sap.read(
            "TADIR", ["PGMID", "OBJECT", "OBJ_NAME", "DEVCLASS", "AUTHOR", "MASTERLANG"],
            ["DEVCLASS LIKE 'YMLS%'"], 5000
        )

        evidence["repository_discovery"] = {
            "TFPM_tables": sap.read("DD02L", ["TABNAME", "TABCLASS", "SQLTAB", "CONTFLAG"], ["TABNAME LIKE 'TFPM%'"], 500),
            "filename_fields": sap.read(
                "DD03L", ["TABNAME", "FIELDNAME", "ROLLNAME", "POSITION"],
                ["FIELDNAME = 'DWNAM' OR FIELDNAME = 'FILENAME' OR FIELDNAME = 'FILE_NAME'"], 1000
            ),
            "formi_fields": sap.read(
                "DD03L", ["TABNAME", "FIELDNAME", "ROLLNAME", "POSITION"], ["FIELDNAME = 'FORMI'"], 1000
            ),
        }

        for table in ["REGUT", "FPAYH", "FPAYHX", "FPAYP", "PAYR", "DTA_ADMIN", "DTA_FILES"]:
            try:
                evidence.setdefault("ddic", {})[table] = [
                    {key: item.get(key) for key in ("FIELDNAME", "KEYFLAG", "ROLLNAME", "DATATYPE", "LENG", "FIELDTEXT")}
                    for item in sap.fields(table)
                ]
            except Exception as exc:
                evidence.setdefault("ddic", {})[table] = {"error": f"{type(exc).__name__}: {exc}"}

        evidence["runs"] = {}
        run_specs = {
            "old_2010_20250318_00002B": ("20250318", "00002B"),
            "new_2010_20260916_00004B": ("20260916", "00004B"),
        }
        for label, (laufd, laufi) in run_specs.items():
            evidence["runs"][label] = {}
            for table in ["REGUT", "FPAYH", "FPAYHX", "FPAYP", "PAYR"]:
                evidence["runs"][label][table] = sap.read(
                    table, None, [f"LAUFD = '{laufd}' AND LAUFI = '{laufi}'"], 200
                )
            evidence["runs"][label]["REGUT_BY_DATE"] = sap.read("REGUT", None, [f"LAUFD = '{laufd}'"], 1000)
            evidence["runs"][label]["REGUT_BY_ID"] = sap.read("REGUT", None, [f"LAUFI = '{laufi}'"], 1000)

        evidence["filename_search"] = sap.read("REGUT", None, ["DWNAM LIKE 'ISPRPTPL%'"], 2000)
        regut_fields = ["MANDT", "ZBUKR", "BANKS", "LAUFD", "LAUFI", "XVORL", "DTKEY", "LFDNR", "DTFOR", "TSNAM", "DWNAM", "DWDAT", "DWTIM", "DWUSR", "REPORT", "DTTYP", "STATUS"]
        evidence["regut_narrow"] = {
            "old_exact": sap.read("REGUT", regut_fields, ["LAUFD = '20250318' AND LAUFI = '00002B'"], 500),
            "old_date": sap.read("REGUT", regut_fields, ["LAUFD = '20250318'"], 2000),
            "old_filename": sap.read("REGUT", regut_fields, ["DWNAM LIKE 'ISPRPTPL%'"], 2000),
            "new_exact": sap.read("REGUT", regut_fields, ["LAUFD = '20260916' AND LAUFI = '00004B'"], 500),
            "new_date": sap.read("REGUT", regut_fields, ["LAUFD = '20260916'"], 2000),
        }

        evidence["format_search"] = {}
        for table in ["TFPM042F", "TFPM042FG", "TFPM042FB", "TFPM042FA", "TFPM042FE", "TFPM042FC"]:
            evidence["format_search"][table] = {
                "old": sap.read(table, None, ["FORMI = 'Z_SEPA_AP'"], 500),
                "new": sap.read(table, None, ["FORMI = 'Z_PT_CGI_XML_CT_V9'"], 500),
            }
        evidence["all_tfpm_format_search"] = {}
        for row in evidence["repository_discovery"]["TFPM_tables"].get("rows", []):
            table = str(row.get("TABNAME", "")).strip()
            if not table:
                continue
            metadata = sap.fields(table)
            if any(str(field.get("FIELDNAME", "")).strip() == "FORMI" for field in metadata):
                evidence["all_tfpm_format_search"][table] = {
                    "old": sap.read(table, None, ["FORMI = 'Z_SEPA_AP'"], 1000),
                    "new": sap.read(table, None, ["FORMI = 'Z_PT_CGI_XML_CT_V9'"], 1000),
                }

        evidence["function_modules"] = {}
        for function in ["Z_FI_PAYMEDIUM_21", "Y_MLS_BCM_FI_PAYMEDIUM_21", "Z_FI_PAYMEDIUM_MT101_41", "FI_PAYMEDIUM_DMEE_CGI_05"]:
            evidence["function_modules"][function] = {
                "TFDIR": sap.read("TFDIR", None, [f"FUNCNAME = '{function}'"], 20),
                "ENLFDIR": sap.read("ENLFDIR", None, [f"FUNCNAME = '{function}'"], 20),
            }

        program_names = {
            "YMLS_R_SHOW_MENU", "LZ_FI_PAYMEDIUM_21UXX", "LZ_FI_PAYMEDIUM_21U01",
            "LZ_FI_PAYMEDIUM_MT101_41UXX", "LZ_FI_PAYMEDIUM_MT101_41U01",
            "LDMEE_CGIUXX", "LDMEE_CGIU02",
        }
        for function_data in evidence["function_modules"].values():
            for row in function_data["TFDIR"].get("rows", []):
                for key in ("PNAME", "INCLUDE"):
                    value = str(row.get(key, "")).strip()
                    if value:
                        program_names.add(value)
                pname = str(row.get("PNAME", "")).strip()
                include_number = str(row.get("INCLUDE", "")).strip().zfill(2)
                if pname.startswith("SAPL") and include_number:
                    program_names.add(f"L{pname[4:]}U{include_number}")
        evidence["program_sources"] = {name: sap.program(name) for name in sorted(program_names)}

        evidence["ymls_payment_config"] = {}
        for table in ["YMLS_T_BCM_FNAME", "YMLS_T_BCM_FNAMH", "YMLS_T_BCM_FNAMT", "YMLS_T_BCM_PAYM", "YMLS_T_BCM_DIR", "YMLS_T_COMPF2"]:
            metadata = sap.fields(table)
            selected: list[str] = []
            length = 0
            for field in metadata:
                name = str(field.get("FIELDNAME", "")).strip()
                size = int(field.get("LENG", 0) or 0)
                if name and size <= 180 and length + size + 1 <= 450:
                    selected.append(name)
                    length += size + 1
            evidence["ymls_payment_config"][table] = sap.read(table, selected, None, 5000)

        evidence["cross_reference_search"] = {}
        for token in ["Z_SEPA_AP", "Z_PT_CGI_XML_CT_V9", "ISPRPTPL", "C_FILENAME", "FILENAME", "FILE_NAME", "DWNAM", "FORMI"]:
            evidence["cross_reference_search"][token] = sap.read(
                "WBCROSSGT", ["OTYPE", "NAME", "INCLUDE", "DIRECT", "INDIRECT", "COMPONENT", "UDATE", "UNAME"],
                [f"NAME = '{token}'"], 5000
            )

        evidence["obpm_standard_modules"] = {
            "TFPM042FB_DMEE": sap.read(
                "TFPM042FB", None, ["FNAME LIKE 'FI_PAYMEDIUM_DMEE_%'"], 5000
            ),
            "TFPM042FBC_DMEE": sap.read(
                "TFPM042FBC", None, ["FNAME LIKE 'FI_PAYMEDIUM_DMEE_%'"], 5000
            ),
        }
        evidence["obpm_variant_tables"] = {}
        for table in ["TFPM042FPB", "TFPM042FSB", "TFPM042FD", "TFPM042FM", "TFPM042FL", "TFPM042FZ"]:
            metadata = sap.fields(table)
            field_names = [str(item.get("FIELDNAME", "")).strip() for item in metadata]
            selected = [name for name in field_names if name in {"MANDT", "FORMI", "EVENT", "FNAME", "PROGRAM", "REPID", "VARIANT", "VARI", "PARAM", "PARNO", "PARID", "VALUE", "XFILESYSTEM", "FILENAME", "FILE_NAME", "PATH", "LOGICAL_FILENAME", "TREE_ID"}]
            if not selected:
                selected = field_names[:12]
            evidence["obpm_variant_tables"][table] = {
                "fields": field_names,
                "old": sap.read(table, selected, ["FORMI = 'Z_SEPA_AP'"], 5000) if "FORMI" in field_names else sap.read(table, selected, None, 5000),
                "new": sap.read(table, selected, ["FORMI = 'Z_PT_CGI_XML_CT_V9'"], 5000) if "FORMI" in field_names else {"rows": []},
            }

        evidence["cross_reference_ddic"] = {
            table: [{key: item.get(key) for key in ("FIELDNAME", "ROLLNAME", "DATATYPE", "LENG", "FIELDTEXT")} for item in sap.fields(table)]
            for table in ["WBCROSSGT", "D010TAB", "REPOSRC"]
        }
    finally:
        sap.conn.close()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(normalise(evidence), ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(OUT))


if __name__ == "__main__":
    main()
