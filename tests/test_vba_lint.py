"""Excel-only VBA rules (LibreOffice, used for the macro tests, does not enforce them)."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "excel"))

from lint_vba import lint  # noqa: E402


def test_folder_import_module_passes_excel_rules():
    assert lint(ROOT / "excel/FolderImport.bas") == []


def test_lint_catches_late_constants_and_undeclared_variables(tmp_path):
    bad = tmp_path / "bad.bas"
    bad.write_text('Attribute VB_Name = "Bad"\nOption Explicit\nPrivate Const A As Long = 1\n\n'
                   "Public Sub One()\n    Dim i As Long\n    For i = 1 To A\n    Next i\n    x = 2\nEnd Sub\n\n"
                   "Private Const MAX_YEARS As Long = 80\n")
    problems = lint(bad)
    assert any("MAX_YEARS" in p and "after the first procedure" in p for p in problems)
    assert any("'x' is not declared" in p for p in problems)
