from __future__ import annotations

from decimal import Decimal
from unittest import mock
from unittest import TestCase

from sap_rfc.fi_document_service import (
    _build_customer_payload,
    _build_gl_payload,
    _build_vendor_payload,
)


class FiDocumentServiceTaxLinkTests(TestCase):
    def test_customer_payload_links_tax_item_to_gl_line(self) -> None:
        payload = _build_customer_payload(
            "QAD",
            {
                "company_code": "2010",
                "posting_date": "2026-09-01",
                "document_date": "2026-09-01",
                "currency": "EUR",
                "header_text": "Teste",
                "reference": "REF-1",
                "username": "CSILVA",
                "customer_account": "0010002949",
                "revenue_gl_account": "700000",
                "amount": "100.00",
                "tax_code": "IVA",
                "tax_amount": "23.00",
                "tax_rate": "23",
                "tax_gl_account": "243000",
                "item_text": "Documento FI",
            },
        )

        self.assertEqual(payload["ACCOUNTGL"][0]["ITEMNO_TAX"], "000003")
        self.assertEqual(payload["ACCOUNTGL"][0]["TAX_CODE"], "IVA")
        self.assertEqual(payload["ACCOUNTTAX"][0]["ITEMNO_TAX"], "000002")
        self.assertEqual(payload["ACCOUNTTAX"][0]["TAX_CODE"], "IVA")
        self.assertEqual(payload["CURRENCYAMOUNT"][2]["CURR_TYPE"], "00")
        self.assertEqual(payload["CURRENCYAMOUNT"][2]["AMT_BASE"], "100.00")
        self.assertEqual(payload["CURRENCYAMOUNT"][2]["TAX_AMT"], "-23.00")

    def test_vendor_payload_links_tax_item_to_gl_line(self) -> None:
        payload = _build_vendor_payload(
            "QAD",
            {
                "company_code": "2010",
                "posting_date": "2026-09-01",
                "document_date": "2026-09-01",
                "currency": "EUR",
                "header_text": "Teste",
                "reference": "REF-2",
                "username": "CSILVA",
                "vendor_account": "0010000040",
                "expense_gl_account": "600000",
                "amount": "100.00",
                "tax_code": "IVA",
                "tax_amount": "23.00",
                "tax_rate": "23",
                "tax_gl_account": "243000",
                "item_text": "Documento FI",
            },
        )

        self.assertEqual(payload["ACCOUNTGL"][0]["ITEMNO_TAX"], "000003")
        self.assertEqual(payload["ACCOUNTGL"][0]["TAX_CODE"], "IVA")
        self.assertEqual(payload["ACCOUNTTAX"][0]["ITEMNO_TAX"], "000002")
        self.assertEqual(payload["ACCOUNTTAX"][0]["TAX_CODE"], "IVA")
        self.assertEqual(payload["CURRENCYAMOUNT"][2]["CURR_TYPE"], "00")
        self.assertEqual(payload["CURRENCYAMOUNT"][2]["AMT_BASE"], "100.00")
        self.assertEqual(payload["CURRENCYAMOUNT"][2]["TAX_AMT"], "23.00")

    @mock.patch.dict(
        "os.environ",
        {
            "SAP_FI_FORM_PAGTO": "T",
            "SAP_FI_FORM_PAGTO_FORNECEDOR": "K",
            "SAP_FI_FORM_PAGTO_CLIENTE": "C",
        },
        clear=False,
    )
    def test_vendor_payload_uses_payment_method_from_env(self) -> None:
        payload = _build_vendor_payload(
            "QAD",
            {
                "company_code": "2010",
                "posting_date": "2026-09-01",
                "document_date": "2026-09-01",
                "currency": "EUR",
                "header_text": "Teste",
                "reference": "REF-PMT",
                "username": "CSILVA",
                "vendor_account": "0010000040",
                "expense_gl_account": "600000",
                "amount": "100.00",
                "tax_code": "",
                "tax_amount": "0.00",
                "item_text": "Documento FI",
            },
        )

        self.assertEqual(payload["ACCOUNTPAYABLE"][0]["PYMT_METH"], "K")

    @mock.patch.dict(
        "os.environ",
        {
            "SAP_FI_FORM_PAGTO": "T",
            "SAP_FI_FORM_PAGTO_FORNECEDOR": "K",
            "SAP_FI_FORM_PAGTO_CLIENTE": "C",
        },
        clear=False,
    )
    def test_customer_payload_uses_payment_method_from_env(self) -> None:
        payload = _build_customer_payload(
            "QAD",
            {
                "company_code": "2010",
                "posting_date": "2026-09-01",
                "document_date": "2026-09-01",
                "currency": "EUR",
                "header_text": "Teste",
                "reference": "REF-PMT-CLI",
                "username": "CSILVA",
                "customer_account": "0010002949",
                "revenue_gl_account": "700000",
                "amount": "100.00",
                "tax_code": "",
                "tax_amount": "0.00",
                "item_text": "Documento FI",
            },
        )

        self.assertEqual(payload["ACCOUNTRECEIVABLE"][0]["PYMT_METH"], "C")

    def test_gl_payload_links_tax_item_to_selected_gl_line(self) -> None:
        payload_credit = _build_gl_payload(
            "QAD",
            {
                "company_code": "2010",
                "posting_date": "2026-09-01",
                "document_date": "2026-09-01",
                "currency": "EUR",
                "header_text": "Teste",
                "reference": "REF-3",
                "username": "CSILVA",
                "debit_gl_account": "100000",
                "credit_gl_account": "700000",
                "amount": "100.00",
                "tax_code": "IVA",
                "tax_amount": "23.00",
                "tax_direction": "credit",
                "tax_rate": "23",
                "tax_gl_account": "243000",
                "item_text": "Documento FI",
            },
        )

        self.assertEqual(payload_credit["ACCOUNTGL"][1]["ITEMNO_TAX"], "000003")
        self.assertEqual(payload_credit["ACCOUNTGL"][1]["TAX_CODE"], "IVA")
        self.assertEqual(payload_credit["ACCOUNTTAX"][0]["ITEMNO_TAX"], "000002")
        self.assertEqual(payload_credit["ACCOUNTTAX"][0]["TAX_CODE"], "IVA")
        self.assertEqual(payload_credit["CURRENCYAMOUNT"][2]["CURR_TYPE"], "00")
        self.assertEqual(payload_credit["CURRENCYAMOUNT"][2]["AMT_BASE"], "100.00")
        self.assertEqual(payload_credit["CURRENCYAMOUNT"][2]["TAX_AMT"], "-23.00")

        payload_debit = _build_gl_payload(
            "QAD",
            {
                "company_code": "2010",
                "posting_date": "2026-09-01",
                "document_date": "2026-09-01",
                "currency": "EUR",
                "header_text": "Teste",
                "reference": "REF-4",
                "username": "CSILVA",
                "debit_gl_account": "100000",
                "credit_gl_account": "700000",
                "amount": "100.00",
                "tax_code": "IVA",
                "tax_amount": "23.00",
                "tax_direction": "debit",
                "tax_rate": "23",
                "tax_gl_account": "243000",
                "item_text": "Documento FI",
            },
        )

        self.assertEqual(payload_debit["ACCOUNTGL"][0]["ITEMNO_TAX"], "000003")
        self.assertEqual(payload_debit["ACCOUNTGL"][0]["TAX_CODE"], "IVA")
        self.assertEqual(payload_debit["ACCOUNTTAX"][0]["ITEMNO_TAX"], "000001")
        self.assertEqual(payload_debit["ACCOUNTTAX"][0]["TAX_CODE"], "IVA")
        self.assertEqual(payload_debit["CURRENCYAMOUNT"][2]["CURR_TYPE"], "00")
        self.assertEqual(payload_debit["CURRENCYAMOUNT"][2]["AMT_BASE"], "100.00")
        self.assertEqual(payload_debit["CURRENCYAMOUNT"][2]["TAX_AMT"], "23.00")


class FiDocumentServiceWithholdingTaxTests(TestCase):
    def test_vendor_payload_includes_withholding_tax_table_when_configured(self) -> None:
        payload = _build_vendor_payload(
            "QAD",
            {
                "company_code": "2010",
                "posting_date": "2026-09-01",
                "document_date": "2026-09-01",
                "currency": "EUR",
                "header_text": "Teste",
                "reference": "REF-WHT",
                "username": "CSILVA",
                "vendor_account": "0010000040",
                "expense_gl_account": "600000",
                "amount": "100.00",
                "tax_code": "IVA",
                "tax_amount": "23.00",
                "tax_rate": "23",
                "tax_gl_account": "243000",
                "withholding_tax_type": "IR",
                "withholding_tax_code": "01",
                "withholding_tax_base_amount": "100.00",
                "withholding_tax_amount": "15.00",
                "item_text": "Documento FI",
            },
        )

        self.assertIn("ACCOUNTWT", payload)
        self.assertEqual(payload["ACCOUNTWT"][0]["ITEMNO_ACC"], "1")
        self.assertEqual(payload["ACCOUNTWT"][0]["WT_TYPE"], "IR")
        self.assertEqual(payload["ACCOUNTWT"][0]["WT_CODE"], "01")
        self.assertEqual(payload["ACCOUNTWT"][0]["BAS_AMT_LC"], "100.00")
        self.assertEqual(payload["ACCOUNTWT"][0]["BAS_AMT_TC"], "100.00")
        self.assertEqual(payload["ACCOUNTWT"][0]["MAN_AMT_LC"], "15.00")
        self.assertEqual(payload["ACCOUNTWT"][0]["MAN_AMT_TC"], "15.00")
        self.assertEqual(payload["ACCOUNTWT"][0]["BAS_AMT_IND"], "X")
        self.assertEqual(payload["ACCOUNTWT"][0]["MAN_AMT_IND"], "X")


class FiDocumentServiceBalanceTests(TestCase):
    def test_vendor_payload_without_tax_remains_balanced(self) -> None:
        payload = _build_vendor_payload(
            "QAD",
            {
                "company_code": "2010",
                "posting_date": "2026-09-01",
                "document_date": "2026-09-01",
                "currency": "EUR",
                "header_text": "Teste",
                "reference": "REF-NO-TAX",
                "username": "CSILVA",
                "vendor_account": "0010000040",
                "expense_gl_account": "600000",
                "amount": "100.00",
                "tax_code": "",
                "tax_amount": "23.00",
                "item_text": "Documento FI",
            },
        )

        self.assertEqual(payload["ACCOUNTTAX"], [])
        self.assertEqual(payload["CURRENCYAMOUNT"][0]["AMT_DOCCUR"], "-100.00")
        self.assertEqual(payload["CURRENCYAMOUNT"][1]["AMT_DOCCUR"], "100.00")

    def test_gl_payload_with_tax_balances_to_zero(self) -> None:
        payload = _build_gl_payload(
            "QAD",
            {
                "company_code": "2010",
                "posting_date": "2026-09-01",
                "document_date": "2026-09-01",
                "currency": "EUR",
                "header_text": "Teste",
                "reference": "REF-GL-TAX",
                "username": "CSILVA",
                "debit_gl_account": "100000",
                "credit_gl_account": "700000",
                "amount": "100.00",
                "tax_code": "IVA",
                "tax_amount": "23.00",
                "tax_direction": "credit",
                "tax_rate": "23",
                "tax_gl_account": "243000",
                "item_text": "Documento FI",
            },
        )

        total = sum(Decimal(row["AMT_DOCCUR"]) for row in payload["CURRENCYAMOUNT"])
        self.assertEqual(total, Decimal("0.00"))
