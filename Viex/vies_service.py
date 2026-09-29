"""
Módulo de serviço VIES - lógica de validação de VAT.
Reutiliza implementação testada do script original.
Com retry automático para erros temporários.
"""

from dataclasses import dataclass
from pathlib import Path
import re
import requests
import xml.etree.ElementTree as ET
from typing import Optional, Callable
import time
import logging
from datetime import datetime


VIES_URL = "https://ec.europa.eu/taxation_customs/vies/services/checkVatService"

# Configuração
VIES_TIMEOUT = 15
VIES_MAX_RETRIES = 4
VIES_BACKOFF_TIMES = [3, 6, 12]  # segundos entre tentativas
VIES_REQUEST_INTERVAL = 1.0  # segundo entre consultas sequenciais

# Criar diretório de logs
LOGS_DIR = Path(__file__).parent / "logs"
LOGS_DIR.mkdir(exist_ok=True)

# Logging
logger = logging.getLogger("validador_vies")
handler = logging.FileHandler(LOGS_DIR / "validador_vies.log")
handler.setFormatter(logging.Formatter(
    "%(asctime)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
))
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)


# Erros temporários do VIES
TEMPORARY_ERRORS = {
    "MS_MAX_CONCURRENT_REQ",
    "SERVER_BUSY",
    "SERVICE_UNAVAILABLE",
    "MS_UNAVAILABLE",
    "TIMEOUT",
    "TEMPORARILY_UNAVAILABLE",
}


@dataclass
class ViesResult:
    vat: str
    country_code: str
    vat_number: str
    valid: bool
    name: Optional[str] = None
    address: Optional[str] = None
    request_date: Optional[str] = None


@dataclass
class ViesResultStatus:
    """Resultado da validação com status de retry."""
    vat: str
    result: Optional[ViesResult] = None
    is_success: bool = False
    error_message: Optional[str] = None
    is_temporary_error: bool = False
    attempt: int = 1
    total_attempts: int = 1


def eh_erro_temporario(erro_msg: str) -> bool:
    """Classifica se um erro é temporário ou permanente."""
    erro_upper = str(erro_msg).upper()

    for temp_error in TEMPORARY_ERRORS:
        if temp_error in erro_upper:
            return True

    # Verificar padrões comuns (case-insensitive)
    patterns = [
        "TIMEOUT",
        "TEMPORARILY",
        "BUSY",
        "CONNECTION",
        "UNAVAILABLE",
        "UNREACHABLE",
        "REFUSED",
    ]

    for pattern in patterns:
        if pattern in erro_upper:
            return True

    return False


def normalizar_vat(vat: str) -> str:
    if not vat:
        raise ValueError("O VAT não pode estar vazio.")

    vat = str(vat).strip().upper()
    vat = re.sub(r"[^A-Z0-9]", "", vat)

    if len(vat) < 3:
        raise ValueError(
            "VAT inválido. Informe o código do país seguido do número."
        )

    return vat


def separar_vat(vat: str) -> tuple[str, str]:
    vat = normalizar_vat(vat)

    country_code = vat[:2]
    vat_number = vat[2:]

    if not country_code.isalpha():
        raise ValueError(
            "O VAT deve começar pelo código do país de 2 letras."
        )

    if not vat_number:
        raise ValueError("Número VAT não informado.")

    return country_code, vat_number


def validar_vies(vat: str, timeout: int = 15) -> ViesResult:
    """
    Valida um VAT através do serviço VIES oficial.
    Retorna ViesResult com dados da consulta.
    Lança exceções em caso de erro técnico ou de validação.
    """

    country_code, vat_number = separar_vat(vat)

    soap_body = f"""<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope
    xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:urn="urn:ec.europa.eu:taxud:vies:services:checkVat:types">

    <soapenv:Header/>

    <soapenv:Body>
        <urn:checkVat>
            <urn:countryCode>{country_code}</urn:countryCode>
            <urn:vatNumber>{vat_number}</urn:vatNumber>
        </urn:checkVat>
    </soapenv:Body>

</soapenv:Envelope>
"""

    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": "",
    }

    try:
        response = requests.post(
            VIES_URL,
            data=soap_body.encode("utf-8"),
            headers=headers,
            timeout=timeout,
        )

        response.raise_for_status()

    except requests.Timeout as exc:
        raise RuntimeError(
            "O serviço VIES excedeu o tempo limite de resposta."
        ) from exc

    except requests.RequestException as exc:
        raise RuntimeError(
            f"Erro de comunicação com o serviço VIES: {exc}"
        ) from exc

    try:
        root = ET.fromstring(response.content)

    except ET.ParseError as exc:
        raise RuntimeError(
            "O VIES devolveu uma resposta XML inválida."
        ) from exc

    namespaces = {
        "soap": "http://schemas.xmlsoap.org/soap/envelope/",
        "ns": "urn:ec.europa.eu:taxud:vies:services:checkVat:types",
    }

    fault = root.find(".//soap:Fault", namespaces)

    if fault is not None:
        fault_code = fault.findtext("faultcode")
        fault_string = fault.findtext("faultstring")

        mensagem = (
            fault_string
            or fault_code
            or "Erro desconhecido."
        )

        raise RuntimeError(
            f"O serviço VIES devolveu um erro: {mensagem}"
        )

    result = root.find(
        ".//ns:checkVatResponse",
        namespaces,
    )

    if result is None:
        raise RuntimeError(
            "Resposta inesperada do serviço VIES."
        )

    def get_text(field: str) -> Optional[str]:
        element = result.find(
            f"ns:{field}",
            namespaces,
        )

        if element is None or element.text is None:
            return None

        value = element.text.strip()

        if value in {"", "---"}:
            return None

        return value

    valid_text = get_text("valid")

    return ViesResult(
        vat=f"{country_code}{vat_number}",
        country_code=get_text("countryCode") or country_code,
        vat_number=get_text("vatNumber") or vat_number,
        valid=(valid_text or "").lower() == "true",
        name=get_text("name"),
        address=get_text("address"),
        request_date=get_text("requestDate"),
    )


def validar_vies_com_retry(
    vat: str,
    on_retry_status: Optional[Callable[[ViesResultStatus], None]] = None,
) -> ViesResultStatus:
    """
    Valida um VAT com retry automático para erros temporários.

    Args:
        vat: VAT a validar
        on_retry_status: Callback para atualizar status de retry na GUI

    Returns:
        ViesResultStatus com resultado e status de retry
    """

    for tentativa in range(1, VIES_MAX_RETRIES + 1):
        try:
            resultado = validar_vies(vat, timeout=VIES_TIMEOUT)

            # Sucesso - resultado funcional
            status = ViesResultStatus(
                vat=vat,
                result=resultado,
                is_success=True,
                error_message=None,
                is_temporary_error=False,
                attempt=tentativa,
                total_attempts=tentativa,
            )

            logger.debug(
                f"{vat} | tentativa {tentativa}/{VIES_MAX_RETRIES} | "
                f"valid={resultado.valid}"
            )

            if on_retry_status:
                on_retry_status(status)

            return status

        except Exception as exc:
            erro_msg = str(exc)
            eh_temporario = eh_erro_temporario(erro_msg)

            logger.debug(
                f"{vat} | tentativa {tentativa}/{VIES_MAX_RETRIES} | "
                f"{erro_msg}"
            )

            if not eh_temporario:
                # Erro permanente - não fazer retry
                status = ViesResultStatus(
                    vat=vat,
                    result=None,
                    is_success=False,
                    error_message=erro_msg,
                    is_temporary_error=False,
                    attempt=tentativa,
                    total_attempts=tentativa,
                )

                if on_retry_status:
                    on_retry_status(status)

                raise

            # Erro temporário
            if tentativa >= VIES_MAX_RETRIES:
                # Esgotadas tentativas
                status = ViesResultStatus(
                    vat=vat,
                    result=None,
                    is_success=False,
                    error_message=f"Erro técnico após {VIES_MAX_RETRIES} tentativas",
                    is_temporary_error=True,
                    attempt=tentativa,
                    total_attempts=VIES_MAX_RETRIES,
                )

                if on_retry_status:
                    on_retry_status(status)

                logger.debug(
                    f"{vat} | esgotadas as {VIES_MAX_RETRIES} tentativas"
                )

                return status

            # Tentar novamente
            tempo_espera = VIES_BACKOFF_TIMES[tentativa - 1]

            status = ViesResultStatus(
                vat=vat,
                result=None,
                is_success=False,
                error_message=f"Erro temporário. Nova tentativa em {tempo_espera}s",
                is_temporary_error=True,
                attempt=tentativa,
                total_attempts=VIES_MAX_RETRIES,
            )

            if on_retry_status:
                on_retry_status(status)

            time.sleep(tempo_espera)
