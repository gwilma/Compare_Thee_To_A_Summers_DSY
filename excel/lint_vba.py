"""Static checks for VBA rules that Excel enforces but LibreOffice does not.

    python excel/lint_vba.py [excel/FolderImport.bas]

1. Module-level declarations (Const, Dim, Private/Public variables) must come before the first
   procedure. Excel ignores later ones, so with Option Explicit their names are "not defined".
2. With Option Explicit, every variable that is assigned, used as a loop counter or ReDim'd must be
   declared (as a parameter, a Dim/Static/Const in the procedure, or at module level).
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

PROC_START = re.compile(r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*\((.*)\)", re.I)
PROC_END = re.compile(r"^End\s+(Sub|Function|Property)\b", re.I)
MODULE_DECL = re.compile(r"^(?:(?:Public|Private|Global|Dim)\s+(?:Const\s+)?|Const\s+)(\w+)", re.I)
KEYWORDS = {"if", "elseif", "else", "end", "for", "next", "do", "loop", "while", "wend", "select", "case", "with", "exit",
            "on", "resume", "goto", "call", "set", "let", "dim", "redim", "const", "static", "option", "open", "close",
            "line", "input", "print", "get", "put", "kill", "mkdir", "erase", "debug", "msgbox", "application", "me",
            "rem", "attribute", "type", "private", "public", "sub", "function", "return", "not", "and", "or"}


def logical_lines(text: str):
    """(line number, code) with comments and string contents removed and continuations joined."""
    buf, start = "", None
    for no, raw in enumerate(text.splitlines(), start=1):
        code = strip(raw)
        if start is None:
            start = no
        if code.rstrip().endswith(" _"):
            buf += code.rstrip()[:-1] + " "
            continue
        yield start, (buf + code).strip()
        buf, start = "", None


def strip(line: str) -> str:
    out, in_str = [], False
    for i, ch in enumerate(line):
        if ch == '"':
            in_str = not in_str
            out.append('"')
            continue
        if not in_str and ch == "'":
            break
        out.append("" if in_str else ch)
    return "".join(out)


def names_in_decl(rest: str) -> list[str]:
    names = []
    for part in split_top(rest):
        m = re.match(r"\s*(?:ByVal\s+|ByRef\s+|Optional\s+|ParamArray\s+)*(\w+)", part, re.I)
        if m:
            names.append(m.group(1).lower())
    return names


def split_top(s: str) -> list[str]:
    parts, depth, cur = [], 0, ""
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    parts.append(cur)
    return parts


def lint(path: Path) -> list[str]:
    lines = list(logical_lines(path.read_text()))
    problems: list[str] = []
    explicit = any(re.match(r"^Option\s+Explicit\b", c, re.I) for _, c in lines)
    module_names, proc_names, procs = set(), set(), []
    seen_proc, current = False, None
    for no, code in lines:
        if not code:
            continue
        m = PROC_START.match(code)
        if m:
            seen_proc = True
            current = {"name": m.group(2).lower(), "params": names_in_decl(m.group(3)), "lines": [], "start": no}
            proc_names.add(m.group(2).lower())
            continue
        if PROC_END.match(code):
            procs.append(current)
            current = None
            continue
        if current is not None:
            current["lines"].append((no, code))
            continue
        d = MODULE_DECL.match(code)
        if d and not re.match(r"^(Option|Attribute)\b", code, re.I):
            if seen_proc:
                problems.append(f"line {no}: module-level declaration after the first procedure "
                                f"(Excel ignores it): {code[:70]}")
            rest = re.sub(r"^(?:(?:Public|Private|Global|Dim)\s+(?:Const\s+)?|Const\s+)", "", code, flags=re.I)
            module_names.update(names_in_decl(rest))
    if not explicit:
        return problems
    for proc in procs:
        declared = set(proc["params"]) | module_names | proc_names | {proc["name"]}
        for _, code in proc["lines"]:
            for stmt in re.split(r":\s+", code):
                m = re.match(r"^(?:Dim|Static|Const)\s+(.*)$", stmt, re.I)
                if m:
                    declared.update(names_in_decl(re.sub(r"=.*$", "", m.group(1)) if stmt.lower().startswith("const")
                                                  else m.group(1)))
        for no, code in proc["lines"]:
            for stmt in re.split(r":\s+", code):
                stmt = re.sub(r"^(?:Else)?If\s+.*?\s+Then\s+", "", stmt, flags=re.I)  # single-line If body
                targets = []
                m = re.match(r"^(?:Set\s+|Let\s+)?(\w+)\s*(?:\([^=]*\))?\s*=(?!=)", stmt, re.I)
                if m:
                    targets.append(m.group(1))
                m = re.match(r"^For\s+(?:Each\s+)?(\w+)", stmt, re.I)
                if m:
                    targets.append(m.group(1))
                m = re.match(r"^ReDim\s+(?:Preserve\s+)?(\w+)", stmt, re.I)
                if m:
                    targets.append(m.group(1))
                for t in targets:
                    if t.lower() not in declared and t.lower() not in KEYWORDS:
                        problems.append(f"line {no}: '{t}' is not declared in {proc['name']} (Option Explicit)")
    return problems


def main(argv: list[str]) -> int:
    path = Path(argv[0]) if argv else Path(__file__).resolve().parent / "FolderImport.bas"
    problems = lint(path)
    for p in problems:
        print(p)
    print(f"{path.name}: {'OK' if not problems else f'{len(problems)} problem(s)'}")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
