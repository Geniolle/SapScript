from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
SCRIPT_PATH = ROOT_DIR / "scripts" / "run_uat_f110_browser.py"


def load_module():
    spec = importlib.util.spec_from_file_location("run_uat_f110_browser", SCRIPT_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules.pop(spec.name, None)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def test_parser_defaults():
    mod = load_module()
    parser = mod.build_parser()
    args = parser.parse_args([])
    assert args.base_url == "http://127.0.0.1:8010"
    assert args.processo == "UAT Simulação"
    assert args.subprocesso == "Executar F110"
    assert args.modo == "massivo"
    assert args.request_option == "4"


def test_parse_param_pairs():
    mod = load_module()
    params = mod.parse_param_pairs(["company_code=2010", "vendor=10000040"])
    assert params == {"company_code": "2010", "vendor": "10000040"}
