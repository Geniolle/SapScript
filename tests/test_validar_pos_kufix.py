import unittest
from unittest.mock import patch

from validar_pos_kufix import POValidatorApp, process_kufix_updates


def ekko_row(ebeln, kufix="", loekz="", wkurs="1.23450"):
    return {"WA": f"{ebeln}|1000|USD|{wkurs}|{kufix}|{loekz}|20260929"}


class FakeConnection:
    def __init__(self, states, bapi_returns=None):
        self.states = {po: list(values) for po, values in states.items()}
        self.bapi_returns = bapi_returns or {}
        self.calls = []
        self.closed = False

    def call(self, function_name, **kwargs):
        self.calls.append((function_name, kwargs))
        if function_name == "RFC_READ_TABLE":
            po = kwargs["OPTIONS"][0]["TEXT"].split("'")[1]
            states = self.states.get(po, [])
            state = states.pop(0) if states else None
            return {"DATA": [] if state is None else [ekko_row(po, **state)]}
        if function_name == "BAPI_PO_CHANGE":
            return {"RETURN": self.bapi_returns.get(kwargs["PURCHASEORDER"], [])}
        return {}

    def close(self):
        self.closed = True


def run(fake, pos):
    return process_kufix_updates(
        pos,
        connection_factory=lambda: fake,
        metadata_validator=lambda _connection: None,
    )


class ProcessKufixUpdatesTests(unittest.TestCase):
    def test_empty_kufix_success_commit_and_reread_x(self):
        fake = FakeConnection({"4500000001": [{"kufix": ""}, {"kufix": "X"}]})
        result = run(fake, ["4500000001"])["4500000001"]
        self.assertEqual(result["result"], "Marcado com sucesso")
        self.assertEqual(result["kufix_after"], "X")
        self.assertIn(("BAPI_TRANSACTION_COMMIT", {"WAIT": "X"}), fake.calls)

    def test_already_x_does_not_call_bapi(self):
        fake = FakeConnection({"4500000002": [{"kufix": "X"}]})
        result = run(fake, ["4500000002"])["4500000002"]
        self.assertEqual(result["result"], "Já estava marcado")
        self.assertNotIn("BAPI_PO_CHANGE", [name for name, _ in fake.calls])

    def test_bapi_error_rolls_back(self):
        error = [{"TYPE": "E", "ID": "ME", "NUMBER": "001", "MESSAGE": "Erro de teste"}]
        fake = FakeConnection({"4500000003": [{"kufix": ""}]}, {"4500000003": error})
        result = run(fake, ["4500000003"])["4500000003"]
        self.assertEqual(result["result"], "Erro SAP")
        self.assertIn("Erro de teste", result["sap_message"])
        self.assertIn(("BAPI_TRANSACTION_ROLLBACK", {}), fake.calls)
        self.assertNotIn("BAPI_TRANSACTION_COMMIT", [name for name, _ in fake.calls])

    def test_success_but_reread_empty_is_post_commit_error(self):
        fake = FakeConnection({"4500000004": [{"kufix": ""}, {"kufix": ""}]})
        result = run(fake, ["4500000004"])["4500000004"]
        self.assertEqual(result["result"], "Erro de validação pós-commit")

    def test_po_not_found(self):
        fake = FakeConnection({"4500000005": [None]})
        result = run(fake, ["4500000005"])["4500000005"]
        self.assertEqual(result["result"], "Não encontrado")
        self.assertNotIn("BAPI_PO_CHANGE", [name for name, _ in fake.calls])

    def test_one_error_does_not_interrupt_next_po(self):
        error = [{"TYPE": "A", "MESSAGE": "Falhou"}]
        fake = FakeConnection(
            {
                "4500000006": [{"kufix": ""}],
                "4500000007": [{"kufix": ""}, {"kufix": "X"}],
            },
            {"4500000006": error},
        )
        results = run(fake, ["4500000006", "4500000007"])
        self.assertEqual(results["4500000006"]["result"], "Erro SAP")
        self.assertEqual(results["4500000007"]["result"], "Marcado com sucesso")

    def test_cancelled_confirmation_makes_zero_change_calls(self):
        app = POValidatorApp.__new__(POValidatorApp)
        app._eligible_pos = lambda: ["4500000008"]
        app._confirm_prd_change = lambda _count: False
        with patch("validar_pos_kufix.process_kufix_updates") as process:
            app.confirm_and_mark_kufix()
        process.assert_not_called()

    def test_payload_changes_only_ex_rate_fx_and_not_exchange_rate(self):
        fake = FakeConnection({"4500000009": [{"kufix": "", "wkurs": "9.87650"}, {"kufix": "X", "wkurs": "9.87650"}]})
        result = run(fake, ["4500000009"])["4500000009"]
        bapi_call = next(kwargs for name, kwargs in fake.calls if name == "BAPI_PO_CHANGE")
        self.assertEqual(bapi_call["POHEADER"], {"EX_RATE_FX": "X"})
        self.assertEqual(bapi_call["POHEADERX"], {"EX_RATE_FX": "X"})
        self.assertNotIn("EXCH_RATE", bapi_call["POHEADER"])
        self.assertEqual(result["wkurs"], "9.87650")


if __name__ == "__main__":
    unittest.main()
