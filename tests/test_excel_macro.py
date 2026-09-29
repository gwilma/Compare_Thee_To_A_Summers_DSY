"""Runs the Excel folder-import macro inside LibreOffice (slow; opt in with RUN_EXCEL_TESTS=1)."""

import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_TESTS") != "1" or shutil.which("soffice") is None,
    reason="set RUN_EXCEL_TESTS=1 and install LibreOffice Calc + python3-uno to run",
)


def test_folder_import_macro():
    subprocess.run([sys.executable, str(ROOT / "excel/add_macros.py")], check=True, timeout=600)
    result = subprocess.run([sys.executable, str(ROOT / "excel/test_folder_import.py")], capture_output=True, text=True,
                            timeout=3000)
    assert "RESULT: PASS" in result.stdout, result.stdout[-3000:] + result.stderr[-2000:]
