"""Compile check for the Basic sources: builds a test book, opens it in a GUI session (Xvfb) and calls one function per
module; a compile error of any module disables the whole library and shows a dialog, whose text is reported.

    python3 tools/basic_compile_check.py [--display :99]
"""
import argparse
import os
import sys
import tempfile
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wmslo import Office  # noqa: E402
from build_ods import build  # noqa: E402


# words of LibreOffice Basic that break a module (or LibreOffice itself) when used as the name of a variable or a parameter:
# «like» is an operator (a variable «like» crashes LibreOffice while it compiles the library), «base» and «name» are keywords
RESERVED = {"like", "base", "name", "any", "type", "mod", "is", "step", "output", "input", "line", "print", "error", "and", "or", "not",
            "xor", "eqv", "imp", "to", "then", "else", "next", "loop", "wend", "select", "case", "exit", "resume", "set", "let", "get",
            "put", "open", "close", "access", "lock", "binary", "random", "append", "shared", "global", "public", "private", "static", "dim",
            "redim", "preserve", "const", "function", "sub", "property", "option", "explicit", "compare", "new", "call", "goto", "gosub",
            "return", "on", "while", "until", "for", "each", "in", "do", "if", "with", "end", "true", "false", "nothing", "null", "empty"}


# 1.0: VBA functions LibreOffice Basic knows only with «Option VBASupport 1» (the books of WMS never set it)
VBA_ONLY = ("Round", "InStrRev", "FormatNumber", "FormatCurrency", "FormatPercent", "MonthName", "WeekdayName", "StrReverse")


def reserved_names(src_dir):
    """the variables and parameters of the Basic sources named by a reserved word: ["Module.bas:line: name", ...]"""
    import re
    bad = []
    decl = re.compile(r"^\s*(?:Dim|Private|Public|Global|Static|ReDim(?:\s+Preserve)?)\s+(.*)$", re.I)
    proc = re.compile(r"^\s*(?:Private\s+|Public\s+)?(?:Function|Sub)\s+\w+\s*\((.*)\)", re.I)
    for fn in sorted(os.listdir(src_dir)):
        if not fn.endswith(".bas"):
            continue
        for i, line in enumerate(open(os.path.join(src_dir, fn), encoding="utf-8"), 1):
            code = line.split("'")[0]
            names = []
            m = proc.match(code)
            if m:
                names = [re.sub(r"^(?:Optional\s+)?(?:ByRef\s+|ByVal\s+)?", "", x.strip(), flags=re.I).split()[0] for x in m.group(1).split(",") if x.strip()]
            else:
                m = decl.match(code)
                if m and not re.match(r"^\s*(?:Private|Public|Global)?\s*(?:Const|Function|Sub)\b", m.group(1), re.I):
                    names = [x.strip().split()[0].split("(")[0] for x in m.group(1).split(",") if x.strip()]
            for n in names:
                if n.lower() in RESERVED:
                    bad.append(f"{fn}:{i}: {n}")
            # after a one-line «If … Then ReDim …» LibreOffice Basic takes the next line into the If as well
            if re.search(r"\bThen\s+ReDim\b", code, re.I):
                bad.append(f"{fn}:{i}: однострочный If … Then ReDim (следующая строка попадёт в If)")
            # functions of the VBA mode only: compiled fine, «Sub-procedure or function procedure not defined» when run
            for f in re.findall(r"(?<![\w.])(" + "|".join(VBA_ONLY) + r")\s*\(", re.sub(r'"[^"]*"', '""', code), re.I):
                bad.append(f"{fn}:{i}: {f}() есть только в режиме VBA (Option VBASupport) — при выполнении «процедура не определена»")
    return bad


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    # the WMS book and (1.0) the books of the tools, of the transfer and of the upgrade wizard
    bad = [x for d in ("basic", "toolbox", "transfer", "upgrade") if os.path.isdir(os.path.join(root, "src", d))
           for x in reserved_names(os.path.join(root, "src", d))]
    if bad:
        print("COMPILE FAILED: ловушки LibreOffice Basic (зарезервированное слово в имени, однострочный If … ReDim): " + "; ".join(bad))
        sys.exit(1)
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    tmp = tempfile.mkdtemp(prefix="wms_cc_")
    book = os.path.join(tmp, "cc.ods")
    ob = Office("wmscc_b%d" % os.getpid(), tmp)
    try:
        build(ob, book, test=True)
    finally:
        ob.terminate()
    o = Office("wmscc_g%d" % os.getpid(), tmp, headless=False, display=a.display)
    ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
    res, dialogs = {}, []

    def call():
        try:
            doc = o.load(book, macros=0)
            res["doc"] = doc
            res["r"] = o.basic(doc, "WmsTestOps", "TestEscRoundTrip", "a;b|c=d~e%f")
            # PRIME 0.8: the second library (the pages and the additions) compiles on its first call too
            if res["r"] is True:
                res["r"] = "WmsPlus"
                res["r"] = o.basic(doc, "WmsPack", "PackDec", 4.0) == 2
        except Exception as e:
            res["e"] = str(e)[:300]

    th = threading.Thread(target=call)
    th.start()
    for _ in range(120):
        time.sleep(0.25)
        if not th.is_alive():
            break
        try:
            w = ut.getTopFocusWindow()
            texts = []
            for cid in w.getChildren():
                try:
                    for pv in w.getChild(cid).getState():
                        if pv.Name == "Text" and pv.Value:
                            texts.append(f"{cid}: {pv.Value}")
                except Exception:
                    pass
            if texts:
                dialogs.append(texts)
                for cid in ("ok", "yes", "close"):
                    if cid in list(w.getChildren()):
                        w.getChild(cid).executeAction("CLICK", ())
                        break
        except Exception:
            pass
    th.join(20)
    o.terminate()
    ok = res.get("r") is True
    __import__("shutil").rmtree(tmp, ignore_errors=True)
    print("COMPILE", "OK" if ok else "FAILED", res.get("r"), res.get("e", ""))
    for d in dialogs:
        print("DIALOG:", " | ".join(d)[:600])
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
