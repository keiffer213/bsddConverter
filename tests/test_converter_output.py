import json
import subprocess
import os
import sys
from pathlib import Path
import pytest
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))
from bsddconverter.mapper import run_excel2bsdd_conversion


def test_compare_gui_vs_original(tmp_path):
    test_excel = "tests/data/test_excel_dd.xlsx"
    test_template = "tests/data/bsdd-import-model.json"
    expected_result = tmp_path / "expected_result.json"
    gui_output = tmp_path / "gui_output.json"

    # Run original converter via CLI
    test_dir = Path(__file__).resolve().parent
    converter_path = test_dir / "data" / "Excel2bSDD_converter.py"
    subprocess.run([
        "python",
        str(converter_path),
        test_excel,
        test_template,
        str(expected_result),
        "True"
    ], check=True)

    # Run GUI-backed converter function
    run_excel2bsdd_conversion(
        excel_path=test_excel,
        template_path=test_template,
        output_path=str(gui_output),
        remove_nulls=True
    )

    # Load and compare outputs
    with open(expected_result, encoding="utf-8") as f1, open(gui_output, encoding="utf-8") as f2:
        original = json.load(f1)
        gui = json.load(f2)

    assert gui == original, "Mismatch between GUI output and original script output"
