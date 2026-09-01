from __future__ import annotations

import datetime as dt
import os
import tempfile
from pathlib import Path
from unittest import TestCase, mock

import sap_rfc.f110_service as f110_service


class F110ServiceTests(TestCase):
    def _store_path(self) -> Path:
        path = Path("tests") / "_tmp_f110_sequence.sqlite3"
        path.unlink(missing_ok=True)
        return path

    def test_resolve_f110_dates_uses_today_and_tomorrow(self) -> None:
        fake_today = dt.date(2026, 8, 31)

        with mock.patch.object(f110_service, "date") as date_mock:
            date_mock.today.return_value = fake_today
            posting_date, next_due_date, run_date = f110_service._resolve_f110_dates("")

        self.assertEqual(posting_date, "20260831")
        self.assertEqual(next_due_date, "20260901")
        self.assertEqual(run_date, "20260831")

    @mock.patch("sap_rfc.f110_service.read_table")
    def test_resolve_f110_laufi_uses_env_seed_when_no_previous_rows(self, read_table_mock: mock.Mock) -> None:
        read_table_mock.return_value = []
        store_path = self._store_path()
        with mock.patch.dict(os.environ, {"SAP_F110_LAUFI": "T0001"}, clear=False), mock.patch.object(
            f110_service, "_f110_laufi_store_path", return_value=store_path
        ):
            value = f110_service._resolve_f110_laufi(
                mock.Mock(),
                mock.Mock(),
                operation_type="pagamento",
                run_date="20260831",
            )

        self.assertEqual(value, "T0001")
        self.assertEqual(read_table_mock.call_count, 3)

    @mock.patch("sap_rfc.f110_service.read_table")
    def test_resolve_f110_laufi_increments_existing_same_day(self, read_table_mock: mock.Mock) -> None:
        read_table_mock.side_effect = [
            [("20260831", "T0001"), ("20260831", "T0002")],
            [],
            [],
        ]
        store_path = self._store_path()
        with mock.patch.dict(os.environ, {"SAP_F110_LAUFI": "T0001"}, clear=False), mock.patch.object(
            f110_service, "_f110_laufi_store_path", return_value=store_path
        ):
            value = f110_service._resolve_f110_laufi(
                mock.Mock(),
                mock.Mock(),
                operation_type="pagamento",
                run_date="20260831",
            )

        self.assertEqual(value, "T0003")
        self.assertEqual(read_table_mock.call_count, 3)

    def test_build_f110_selection_params_for_supplier_includes_text_list_and_flags(self) -> None:
        params = f110_service._build_f110_selection_params(
            operation_type="pagamento",
            run_id="T0001",
            posting_date_sap="20260831",
            next_due_date_sap="20260901",
            payment_method="S",
            company_code="2010",
            account_number="0010000040",
            document_number="6050000047",
        )

        by_name = {entry["SELNAME"]: entry for entry in params}
        self.assertEqual(by_name["PAR_TEX1"]["LOW"], "BKPF-BELNR")
        self.assertEqual(by_name["PAR_LIS1"]["LOW"], "6050000047")
        self.assertEqual(by_name["PAR_XFA"]["LOW"], "X")
        self.assertEqual(by_name["PAR_XZW"]["LOW"], "X")
        self.assertEqual(by_name["PAR_XBL"]["LOW"], "X")
        self.assertIn("SEL_KRED", by_name)
        self.assertNotIn("SEL_DEBI", by_name)

    def test_build_f110_selection_params_for_customer_uses_debi(self) -> None:
        params = f110_service._build_f110_selection_params(
            operation_type="cobranca",
            run_id="C0001",
            posting_date_sap="20260831",
            next_due_date_sap="20260901",
            payment_method="Q",
            company_code="2010",
            account_number="0010002949",
            document_number="720000015620102026",
        )

        by_name = {entry["SELNAME"]: entry for entry in params}
        self.assertIn("SEL_DEBI", by_name)
        self.assertNotIn("SEL_KRED", by_name)
        self.assertEqual(by_name["PAR_TEX1"]["LOW"], "BKPF-BELNR")
        self.assertEqual(by_name["PAR_LIS1"]["LOW"], "7200000156")

    @mock.patch("sap_rfc.f110_service.read_table")
    def test_resolve_f110_laufi_advances_when_local_store_has_previous_value(self, read_table_mock: mock.Mock) -> None:
        read_table_mock.return_value = []
        store_path = self._store_path()
        with mock.patch.dict(os.environ, {"SAP_F110_LAUFI": "T0001"}, clear=False), mock.patch.object(
            f110_service, "_f110_laufi_store_path", return_value=store_path
        ):
            first = f110_service._resolve_f110_laufi(
                mock.Mock(),
                mock.Mock(),
                operation_type="pagamento",
                run_date="20260831",
            )
            second = f110_service._resolve_f110_laufi(
                mock.Mock(),
                mock.Mock(),
                operation_type="pagamento",
                run_date="20260831",
            )

        self.assertEqual(first, "T0001")
        self.assertEqual(second, "T0002")
        store_path.unlink(missing_ok=True)
