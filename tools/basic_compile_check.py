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


def precedence_traps(src_dir):
    """1.2.0: in LibreOffice Basic And, Or, Xor, Eqv, Imp are one level of precedence, evaluated from the left («a Or b And c» is «(a Or b) And c»): an expression
    that mixes And and Or outside the brackets does not mean what it seems to. The lines of the sources with such an expression: ["Module.bas:line: text", ...]"""
    import re
    bad = []
    for fn in sorted(os.listdir(src_dir)):
        if not fn.endswith(".bas"):
            continue
        lines = open(os.path.join(src_dir, fn), encoding="utf-8").read().split("\n")
        i = 0
        while i < len(lines):
            full, j = lines[i], i
            while full.rstrip().endswith(" _") and j + 1 < len(lines):
                j += 1
                full = full.rstrip()[:-2] + " " + lines[j].strip()
            t = re.sub(r'"[^"]*"', '""', full).split("'")[0]
            prev = None
            while prev != t:
                prev = t
                t = re.sub(r'\([^()]*\)', '', t)
            if re.search(r'\bAnd\b', t) and re.search(r'\bOr\b', t):
                bad.append(f"{fn}:{i + 1}: And и Or без скобок: {full.strip()[:80]}")
            i = j + 1
    return bad


# 1.2.1: the traps of the missing short circuit. In LibreOffice Basic «And» and «Or» evaluate BOTH sides and «IIf» BOTH branches, so a check that stands before a call does not stop the call:
# «If q = 1 Or Not X(Mid(s, q - 1, 1))» and «IIf(i > 1, Mid(s, i - 1, 1), " ")» both run «Mid(s, 0, 1)» — «Invalid procedure call», and in 1.1.0–1.2.0 one name that began with a number stopped the whole
# list of «Метизы» and of «Инструмент ЧПУ». What is allowed: the block «If … Then … End If» (nothing of it runs when the check fails), the function that answers "" outside the text (WmsTool.TlCh).
_RISKY = None


def _risky():
    import re
    global _RISKY
    if _RISKY is None:
        _RISKY = re.compile(r"(?<![\w.])(Mid|Left|Right|String)\s*\(", re.I)
    return _RISKY


def _split_args(t):
    out, depth, cur = [], 0, ""
    for ch in t:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return out


def _calls(t, name):
    """the text inside the brackets of every call name(...) of the line"""
    import re
    res = []
    for m in re.finditer(r"(?<![\w.])" + name + r"\s*\(", t, re.I):
        d, k = 1, m.end()
        while k < len(t) and d:
            d += (t[k] == "(") - (t[k] == ")")
            k += 1
        res.append((m, t[m.end():k - 1]))
    return res


def _pos_args(fname, inner):
    """the arguments of a call that must not fall below the allowed value: Mid — the start; Left, Right — the length; String — the count"""
    a = _split_args(inner)
    return a[0:1] if fname.lower() == "string" else a[1:2]


def _without_risky(op):
    """the operand with its Mid/Left/Right/String calls cut out: what is left is the check itself"""
    while True:
        m = _risky().search(op)
        if not m:
            return op
        d, k = 1, m.end()
        while k < len(op) and d:
            d += (op[k] == "(") - (op[k] == ")")
            k += 1
        op = op[:m.start()] + " " + op[k:]


def _operands(cond):
    """the operands of a condition split by And / Or outside the brackets"""
    import re
    res, depth, cur, i = [], 0, "", 0
    while i < len(cond):
        ch = cond[i]
        depth += (ch == "(") - (ch == ")")
        if depth == 0:
            m = re.match(r"\s(And|Or)\s", cond[i:], re.I)
            if m:
                res.append(cur)
                cur = ""
                i += m.end()
                continue
        cur += ch
        i += 1
    res.append(cur)
    return res


def _groups(cond):
    """the condition and the text inside every pair of brackets of it (a check may hide in them: «(q = 0 Or …) And …»)"""
    res, stack = [cond], []
    for i, ch in enumerate(cond):
        if ch == "(":
            stack.append(i)
        elif ch == ")" and stack:
            res.append(cond[stack.pop() + 1:i])
    return res


def eager_traps_in(text):
    """the lines of a Basic source (text) with a short-circuit trap: [(line number, why, the line)]"""
    import re
    out = []
    lines = text.split("\n")
    cond_re = re.compile(r"^\s*(?:ElseIf|If|Do\s+While|Do\s+Until|Loop\s+While|Loop\s+Until|While)\s+(.*?)(?:\s+Then\b.*)?$", re.I)
    sub = re.compile(r"[A-Za-z_]\w*\s*-\s*\w")
    i = 0
    while i < len(lines):
        full, j = lines[i], i
        while full.rstrip().endswith(" _") and j + 1 < len(lines):
            j += 1
            full = full.rstrip()[:-2] + " " + lines[j].strip()
        t = re.sub(r'"(?:[^"]|"")*"', '""', full).split("'")[0]
        why = []
        for _, args in _calls(t, "IIf"):
            for br in _split_args(args)[1:]:
                for m in _risky().finditer(br):
                    d, k = 1, m.end()
                    while k < len(br) and d:
                        d += (br[k] == "(") - (br[k] == ")")
                        k += 1
                    if any(sub.search(x) for x in _pos_args(m.group(1), br[m.end():k - 1])):
                        why.append(f"IIf считает обе ветки: {m.group(1)}(…-…) выполнится и когда проверка не прошла")
        m = cond_re.match(t)
        if m:
            for g in _groups(m.group(1)):
                ops = _operands(g)
                for k in range(1, len(ops)):
                    for rm in _risky().finditer(ops[k]):
                        d, q = 1, rm.end()
                        while q < len(ops[k]) and d:
                            d += (ops[k][q] == "(") - (ops[k][q] == ")")
                            q += 1
                        for x in _pos_args(rm.group(1), ops[k][rm.end():q - 1]):
                            for mm in re.finditer(r"\b([A-Za-z_]\w*)\s*-\s*\w", x):
                                if any(re.search(r"\b" + mm.group(1) + r"\b", _without_risky(ops[n])) for n in range(k)):
                                    why.append(f"And/Or не обрывают проверку: {rm.group(1)}(… {mm.group(1)} - …) выполнится и когда {mm.group(1)} не прошла")
                            mm = re.fullmatch(r"\s*([A-Za-z_]\w*)\s*", x)
                            if mm and any(re.search(r"\b" + mm.group(1) + r"\s*(=|<|<=|>|>=|<>)\s*[01]\b", _without_risky(ops[n])) for n in range(k)):
                                why.append(f"And/Or не обрывают проверку: {rm.group(1)}(…, {mm.group(1)}, …) выполнится и при {mm.group(1)} = 0")
        for w in dict.fromkeys(why):
            out.append((i + 1, w, full.strip()))
        i = j + 1
    return out


def eager_traps(src_dir):
    """the traps of the missing short circuit in the Basic sources of a folder: ["Module.bas:line: why: text", ...]"""
    bad = []
    for fn in sorted(os.listdir(src_dir)):
        if fn.endswith(".bas"):
            for ln, why, text in eager_traps_in(open(os.path.join(src_dir, fn), encoding="utf-8").read()):
                bad.append(f"{fn}:{ln}: {why}: {text[:90]}")
    return bad


# 1.2.1: the text of an error is gone after «On Error Resume Next» (in the handler itself or in any procedure it calls — WmsView.ProgressEnd has one): «Error$», «Err», «Erl» are read first.
# 1.2.0 told «список инструмента не построен:  (строка 0)» for that reason.
def lost_error_in(text):
    """the handlers of a Basic source (text) that wipe the error before reading it: [(line number of the label, why)]"""
    import re
    out = []
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if not re.match(r"^[A-Za-z_]\w*:\s*$", ln):
            continue
        blk = []
        for j in range(i + 1, min(i + 40, len(lines))):
            t = re.sub(r'"(?:[^"]|"")*"', '""', lines[j]).split("'")[0]
            if re.match(r"^\s*(End\s+(Function|Sub)|Exit\s+(Function|Sub)|Resume\b)", t, re.I) or re.match(r"^[A-Za-z_]\w*:\s*$", t):
                break
            blk.append((j + 1, t))
        wipe = next((k for k, (_, t) in enumerate(blk) if re.search(r"\bOn\s+Error\b|ProgressEnd", t, re.I)), None)
        use = next((k for k, (_, t) in enumerate(blk) if re.search(r"Error\$|\bErl\b|\bErr\b", t)), None)
        if wipe is not None and use is not None and wipe <= use:
            out.append((i + 1, f"обработчик стирает ошибку (строка {blk[wipe][0]}: On Error / ProgressEnd) раньше, чем читает её (строка {blk[use][0]})"))
    return out


def lost_error_traps(src_dir):
    """the handlers that wipe the error before reading it, in a folder of Basic sources: ["Module.bas:line: why", ...]"""
    bad = []
    for fn in sorted(os.listdir(src_dir)):
        if fn.endswith(".bas"):
            for ln, why in lost_error_in(open(os.path.join(src_dir, fn), encoding="utf-8").read()):
                bad.append(f"{fn}:{ln}: {why}")
    return bad


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    # the WMS book and (1.0) the books of the tools, of the transfer and of the upgrade wizard
    bad = [x for d in ("basic", "toolbox", "transfer", "upgrade") if os.path.isdir(os.path.join(root, "src", d))
           for x in reserved_names(os.path.join(root, "src", d)) + precedence_traps(os.path.join(root, "src", d)) + eager_traps(os.path.join(root, "src", d)) + lost_error_traps(os.path.join(root, "src", d))]
    if bad:
        print("COMPILE FAILED: ловушки LibreOffice Basic (зарезервированное слово в имени, однострочный If … ReDim, And и Or без скобок, IIf и And/Or вычисляют всё — Mid(…, 0, …), обработчик стирает ошибку до чтения): " + "; ".join(bad))
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
