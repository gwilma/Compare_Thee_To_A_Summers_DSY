"""Turn the built workbook into a macro-enabled .xlsm containing FolderImport.bas.

    python excel/add_macros.py [in.xlsx] [out.xlsm]

LibreOffice (headless, via Python-UNO) compiles the module into a VBA project
(vbaProject.bin) with document modules matching the workbook's code names. That
binary is then added to a copy of the openpyxl-built workbook, so every format,
chart and formula stays exactly as built.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
BAS = HERE / "FolderImport.bas"

VBA_CT = '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
VBA_REL = ('<Relationship Id="rIdVBA1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
           'Target="vbaProject.bin"/>')


def module_source() -> tuple[str, str]:
    text = BAS.read_text()
    name = re.search(r'Attribute VB_Name = "([^"]+)"', text).group(1)
    body = "\n".join(line for line in text.splitlines() if not line.startswith("Attribute VB_"))
    return name, "Option VBASupport 1\n" + body + "\n"


class LibreOffice:
    """A headless soffice driven over a UNO socket."""

    def __init__(self, port: int = 2002):
        import uno  # noqa: F401  (python3-uno)

        self.port = port
        self.profile = tempfile.mkdtemp(prefix="lo-profile-")
        self.proc = subprocess.Popen(
            ["soffice", "--headless", "--norestore", "--nologo", "--nodefault",
             f"-env:UserInstallation=file://{self.profile}", f"--accept=socket,host=localhost,port={port};urp;"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        import uno

        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(120):
            try:
                self.ctx = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        else:
            raise RuntimeError("LibreOffice did not start")
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

    @staticmethod
    def pv(name, value):
        import uno

        p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        p.Name, p.Value = name, value
        return p

    def load(self, path: Path, macros: bool = False):
        import uno

        props = [self.pv("Hidden", True)]
        if macros:
            props.append(self.pv("MacroExecutionMode", 4))  # ALWAYS_EXECUTE_NO_WARN
        return self.desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(path)), "_blank", 0, tuple(props))

    def store(self, doc, path: Path, filter_name: str):
        import uno

        doc.storeToURL(uno.systemPathToFileUrl(str(path)), (self.pv("FilterName", filter_name),))

    def close(self):
        try:
            self.desktop.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=60)
        except subprocess.TimeoutExpired:
            self.proc.kill()
        shutil.rmtree(self.profile, ignore_errors=True)


def build_vba_project(xlsx: Path, lo: LibreOffice) -> bytes:
    """Compile FolderImport.bas into a vbaProject.bin whose document modules match the workbook's code names."""
    import uno
    from com.sun.star.script.ModuleType import DOCUMENT, NORMAL

    doc = lo.load(xlsx)
    try:
        libs = doc.BasicLibraries
        libs.VBACompatibilityMode = True
        libs.ProjectName = "VBAProject"
        if not libs.hasByName("VBAProject"):
            libs.createLibrary("VBAProject")
        lib = libs.getByName("VBAProject")

        name, source = module_source()
        info = uno.createUnoStruct("com.sun.star.script.ModuleInfo")
        info.ModuleType = NORMAL
        lib.insertModuleInfo(name, info)
        lib.insertByName(name, source)

        sheets = doc.Sheets
        for code_name in ["ThisWorkbook"] + [sheets.getByIndex(i).CodeName for i in range(sheets.Count)]:
            if lib.hasByName(code_name):
                continue
            dinfo = uno.createUnoStruct("com.sun.star.script.ModuleInfo")
            dinfo.ModuleType = DOCUMENT
            lib.insertModuleInfo(code_name, dinfo)
            lib.insertByName(code_name, "Option VBASupport 1\n")

        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "vba.xlsm"
            lo.store(doc, out, "Calc MS Excel 2007 VBA XML")
            with zipfile.ZipFile(out) as z:
                return z.read("xl/vbaProject.bin")
    finally:
        doc.close(True)


def inject(xlsx: Path, xlsm: Path, vba: bytes) -> None:
    """Copy the workbook into a macro-enabled package with the given VBA project."""
    with zipfile.ZipFile(xlsx) as zin, zipfile.ZipFile(xlsm, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                text = data.decode()
                text = text.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
                if 'Extension="bin"' not in text:
                    text = text.replace("<Default ", VBA_CT + "<Default ", 1)
                data = text.encode()
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = data.decode().replace("</Relationships>", VBA_REL + "</Relationships>").encode()
            zout.writestr(item, data)
        zout.writestr("xl/vbaProject.bin", vba)


def main(src: str | None = None, dst: str | None = None) -> Path:
    xlsx = Path(src) if src else HERE / "summers_dsy.xlsx"
    xlsm = Path(dst) if dst else xlsx.with_suffix(".xlsm")
    lo = LibreOffice(port=int(os.environ.get("LO_PORT", 2002)))
    try:
        vba = build_vba_project(xlsx, lo)
    finally:
        lo.close()
    inject(xlsx, xlsm, vba)
    return xlsm


if __name__ == "__main__":
    print(main(*sys.argv[1:3]))
