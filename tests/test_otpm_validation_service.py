"""
tests/test_otpm_validation_service.py
Testes unitários e de integração para o serviço OtpmValidationService (TPAMA, TPAMAT, T028P).
Inclui cenários simulados (mock) e cenários com limites técnicos (20 caracteres).
"""
import unittest
from unittest.mock import MagicMock, patch

from sap_rfc.otpm_validation_service import (
    OtpmValidationService,
    TpamaStatus,
    T028pStatus,
    sanitize_panam,
    T028pRecord,
)


class TestOtpmValidationService(unittest.TestCase):

    def test_sanitize_panam_truncation_and_accents(self):
        # Limite técnico de 20 caracteres sem caracteres inválidos
        assert sanitize_panam("FIDELIDADE COMPANHIA DE SEGUROS") == "FIDELIDADE COMPANHIA"
        assert len(sanitize_panam("FIDELIDADE COMPANHIA DE SEGUROS")) <= 20
        # Acentos e caracteres especiais
        assert sanitize_panam("MUNICÍPIO DE BARCELOS") == "MUNICIPIO DE BARCELO"
        assert sanitize_panam("") == ""
        assert sanitize_panam(None) == ""

    def test_check_tpama_existente(self):
        mock_conn = MagicMock()
        service = OtpmValidationService(environment="DEV", connection=mock_conn)

        # Simular retorno do read_table para TPAMA
        with patch("sap_rfc.otpm_validation_service.read_table") as mock_read:
            mock_read.side_effect = [
                # 1ª chamada: TPAMA
                [["EBST", "FIDELIDADE", "FIDELIDADE", ""]],
                # 2ª chamada: TPAMAT
                [["FIDELIDADE COMPANHIA"]],
            ]

            result = service.check_tpama("FIDELIDADE")
            assert result.status == TpamaStatus.EXISTE
            assert result.panam == "FIDELIDADE"
            assert result.record is not None
            assert result.record.regex == "FIDELIDADE"
            assert result.record.note == "FIDELIDADE COMPANHIA"

    def test_check_tpama_nao_existente(self):
        mock_conn = MagicMock()
        service = OtpmValidationService(environment="DEV", connection=mock_conn)

        with patch("sap_rfc.otpm_validation_service.read_table") as mock_read:
            mock_read.return_value = []

            result = service.check_tpama("INEXISTENTE")
            assert result.status == TpamaStatus.NAO_EXISTE
            assert "não encontrada" in result.message

    def test_check_tpama_erro_rfc(self):
        mock_conn = MagicMock()
        service = OtpmValidationService(environment="DEV", connection=mock_conn)

        with patch("sap_rfc.otpm_validation_service.read_table") as mock_read:
            mock_read.side_effect = Exception("RFC_COMMUNICATION_ERROR: Host unreachable")

            result = service.check_tpama("FIDELIDADE")
            assert result.status == TpamaStatus.ERRO_RFC
            assert "Erro RFC" in result.message

    def test_validate_t028p_completo(self):
        mock_conn = MagicMock()
        service = OtpmValidationService(environment="DEV", connection=mock_conn)

        # Registo simulado existente em DEV
        rec = T028pRecord(
            bukrs="2100",
            hbkid="BST01",
            hktid="",
            intag="000",
            panam="FIDELIDADE",
            targfi="EBVGINT",
            vgext="NDDT",
            vozpm="-",
            enabled="X",
            prefix="ZDD4",
        )

        with patch.object(service, "fetch_t028p_by_panam") as mock_fetch:
            mock_fetch.return_value = ([rec], "1 registo lido")

            expected = [
                {
                    "BUKRS": "2100",
                    "HBKID": "BST01",
                    "HKTID": "",
                    "VGEXT": "NDDT",
                    "VOZPM": "-",
                    "TARGFI": "EBVGINT",
                    "PREFIX": "ZDD4",
                    "ENABLED": "X",
                }
            ]

            val = service.validate_t028p_assignments("FIDELIDADE", expected)
            assert val.status == T028pStatus.EXISTE
            assert val.total_esperado == 1
            assert len(val.chaves_em_falta) == 0

    def test_validate_t028p_parcial(self):
        mock_conn = MagicMock()
        service = OtpmValidationService(environment="DEV", connection=mock_conn)

        rec = T028pRecord(
            bukrs="2100",
            hbkid="BST01",
            hktid="",
            intag="000",
            panam="LEASYS MOBILITY",
            targfi="EBVGINT",
            vgext="NDDT",
            vozpm="-",
            enabled="X",
            prefix="ZDD4",
        )

        with patch.object(service, "fetch_t028p_by_panam") as mock_fetch:
            mock_fetch.return_value = ([rec], "1 registo lido")

            # Esperamos BST01 e BCP01
            expected = [
                {
                    "BUKRS": "2100",
                    "HBKID": "BST01",
                    "HKTID": "",
                    "VGEXT": "NDDT",
                    "VOZPM": "-",
                    "TARGFI": "EBVGINT",
                    "PREFIX": "ZDD4",
                    "ENABLED": "X",
                },
                {
                    "BUKRS": "2100",
                    "HBKID": "BCP01",
                    "HKTID": "",
                    "VGEXT": "NDDT",
                    "VOZPM": "-",
                    "TARGFI": "EBVGINT",
                    "PREFIX": "ZDD4",
                    "ENABLED": "X",
                },
            ]

            val = service.validate_t028p_assignments("LEASYS MOBILITY", expected)
            assert val.status == T028pStatus.PARCIAL
            assert val.total_esperado == 2
            assert len(val.chaves_em_falta) == 1
            assert val.chaves_em_falta[0][1] == "BCP01"

    def test_validate_t028p_divergente(self):
        mock_conn = MagicMock()
        service = OtpmValidationService(environment="DEV", connection=mock_conn)

        rec = T028pRecord(
            bukrs="2100",
            hbkid="BST01",
            hktid="",
            intag="000",
            panam="FIDELIDADE",
            targfi="EBVGINT",
            vgext="NDDT",
            vozpm="-",
            enabled="",  # Desativado
            prefix="ZDD4",
        )

        with patch.object(service, "fetch_t028p_by_panam") as mock_fetch:
            mock_fetch.return_value = ([rec], "1 registo lido")

            expected = [
                {
                    "BUKRS": "2100",
                    "HBKID": "BST01",
                    "HKTID": "",
                    "VGEXT": "NDDT",
                    "VOZPM": "-",
                    "TARGFI": "EBVGINT",
                    "PREFIX": "ZDD4",
                    "ENABLED": "X",  # Esperado ativo
                }
            ]

            val = service.validate_t028p_assignments("FIDELIDADE", expected)
            assert val.status == T028pStatus.DIVERGENTE
            assert len(val.divergencias) > 0

    def test_validate_t028p_nao_existe(self):
        mock_conn = MagicMock()
        service = OtpmValidationService(environment="DEV", connection=mock_conn)

        with patch.object(service, "fetch_t028p_by_panam") as mock_fetch:
            mock_fetch.return_value = ([], "0 registos")

            expected = [
                {
                    "BUKRS": "2100",
                    "HBKID": "BST01",
                    "HKTID": "",
                    "VGEXT": "NDDT",
                    "VOZPM": "-",
                    "TARGFI": "EBVGINT",
                    "PREFIX": "ZDD4",
                }
            ]

            val = service.validate_t028p_assignments("VOUGAINVEST", expected)
            assert val.status == T028pStatus.NAO_EXISTE

    def test_validate_t028p_erro_rfc(self):
        mock_conn = MagicMock()
        service = OtpmValidationService(environment="DEV", connection=mock_conn)

        with patch.object(service, "fetch_t028p_by_panam") as mock_fetch:
            mock_fetch.return_value = (None, "Falha de comunicação RFC")

            val = service.validate_t028p_assignments("FIDELIDADE", [])
            assert val.status == T028pStatus.ERRO_RFC
            assert "Falha de comunicação RFC" in val.message

    def test_validate_t028p_partn_idflag_divergence(self):
        mock_conn = MagicMock()
        service = OtpmValidationService(environment="DEV", connection=mock_conn)

        rec = T028pRecord(
            bukrs="2100",
            hbkid="BST01",
            hktid="",
            intag="000",
            panam="FIDELIDADE",
            targfi="EBAVKON",
            vgext="NDDT",
            vozpm="-",
            enabled="X",
            prefix="10000305",
            partn="BP_EXISTENTE",
            idflag="X",
        )

        with patch.object(service, "fetch_t028p_by_panam") as mock_fetch:
            mock_fetch.return_value = ([rec], "1 registo lido")

            expected = [
                {
                    "BUKRS": "2100",
                    "HBKID": "BST01",
                    "HKTID": "",
                    "VGEXT": "NDDT",
                    "VOZPM": "-",
                    "TARGFI": "EBAVKON",
                    "PREFIX": "10000305",
                    "ENABLED": "X",
                    "PARTN": "BP_DIFERENTE",
                    "IDFLAG": "",
                }
            ]

            val = service.validate_t028p_assignments("FIDELIDADE", expected)
            assert val.status == T028pStatus.DIVERGENTE
            assert any("PARTN" in d for d in val.divergencias)
            assert any("IDFLAG" in d for d in val.divergencias)

    def test_figueirimo_and_renda_foz_plaza_figu_distinction(self):
        mock_conn = MagicMock()
        service = OtpmValidationService(environment="DEV", connection=mock_conn)

        with patch("sap_rfc.otpm_validation_service.read_table") as mock_read:
            def read_side_effect(conn, guard, table_name, fields, options, rowcount):
                opt = options[0]["TEXT"] if options else ""
                if "PANAM = 'FIGUEIRIMO'" in opt:
                    return []
                if "PANAM = 'RENDA FOZ PLAZA FIGU'" in opt:
                    return [["EBST", "RENDA FOZ PLAZA FIGU", "RENDA FOZ PLAZA FIGUEIRIMO", ""]]
                return []

            mock_read.side_effect = read_side_effect

            res_fig = service.check_tpama("FIGUEIRIMO")
            assert res_fig.status == TpamaStatus.NAO_EXISTE

            res_renda = service.check_tpama("RENDA FOZ PLAZA FIGU")
            assert res_renda.status == TpamaStatus.EXISTE
            assert res_renda.record.regex == "RENDA FOZ PLAZA FIGUEIRIMO"
