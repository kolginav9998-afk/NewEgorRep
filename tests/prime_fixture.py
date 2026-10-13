"""PRIME 0.8 — the warehouse of the user as it is: a working book of release 0.7.2 after the transfer of the old table and a
working day in it, made by the sources of release 0.7.2 themselves (git archive of its commit, its build_ods.py, its
legacy_transfer.py, the macros inside its book). tests/run_prime.py upgrades copies of it to this release.

    warehouse_072(folder) -> (book, facts)   # built once per folder; facts: what the working day posted
"""
import datetime
import os
import shutil
import subprocess
import sys

import uno  # noqa: F401  (LibreOffice Python-UNO)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import Session, git_sources  # noqa: E402

COMMIT_072 = "7e05447"          # release 0.7.2 (HOTFIX 0.7.2)
ENV = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}


def run_old(src, tool, *args, timeout=3600):
    r = subprocess.run([sys.executable, os.path.join(src, "tools", tool)] + list(args), capture_output=True, text=True, timeout=timeout, env=ENV)
    return r.returncode, (r.stdout + r.stderr).strip()


def put(sh, col, r, v):
    c = sh.getCellByPosition(col, r)
    if isinstance(v, float):
        c.setValue(v)
    else:
        c.setString(v)


def warehouse_072(folder):
    """the book of 0.7.2 transferred from the main dataset of the transfer tests and worked in for a day; (book, facts) or
    raises with the reason (e.g. sources without the history of git)"""
    book = os.path.join(folder, "WMS_WORK", "WMS_PROD.ods")
    done = os.path.join(folder, "fixture_done.txt")
    if os.path.isfile(done) and os.path.isfile(book):
        return book, dict(line.split("=", 1) for line in open(done, encoding="utf-8").read().splitlines() if "=" in line)
    shutil.rmtree(folder, ignore_errors=True)
    os.makedirs(folder)
    src = os.path.join(folder, "src072")
    why = git_sources(COMMIT_072, src)
    if why:
        raise RuntimeError(why)
    cand = os.path.join(folder, "WMS_PROD_CANDIDATE.ods")
    rc, out = run_old(src, "build_ods.py", cand)
    if rc != 0:
        raise RuntimeError("книга-кандидат 0.7.2: " + out[-800:])
    import run_legacy_transfer as LT
    st = LT.staging(os.path.join(folder, "staging.ods"), LT.full_rows())
    work = os.path.join(folder, "LEGACY_WORK")
    for step in ("check", "build", "verify", "promote"):
        args = [step, "--work", work, "--staging", st]
        if step in ("check", "build"):
            args += ["--candidate", cand]
        if step == "promote":
            args += ["--to", os.path.join(folder, "WMS_WORK")]
        rc, out = run_old(src, "legacy_transfer.py", *args)
        if rc != 0:
            raise RuntimeError(f"перенос 0.7.2, шаг {step}: {out[-1200:]}")
    # a working day in 0.7.2: the engines of the macros inside the book, cells written as a storekeeper types them
    s = Session(book)
    facts = {}
    try:
        rep = s.B("StartupReport") or ""
        if "ядро 0.7.2" not in rep or "работа разрешена" not in rep:
            raise RuntimeError("книга 0.7.2 не открылась в рабочем состоянии: " + rep[:300])
        iss = s.doc.Sheets.getByName("Выдачи")
        today = float((datetime.date.today() - datetime.date(1899, 12, 30)).days)
        for r, (e, q, who) in enumerate((("ЕИ-00000427", "5", "Петров"), ("ЕИ-00000500", "20", "Сидоров"), ("ЕИ-00001001", "3", "Петров")), start=1):
            put(iss, 11, r, e)
            put(iss, 4, r, q)
            put(iss, 7, r, today)
            put(iss, 8, r, who)
            facts[f"issue{r}"] = s.I("IssuePostRow", r)
        ret = s.doc.Sheets.getByName("Возврат")
        put(ret, 1, 1, "2")
        put(ret, 5, 1, "5")
        put(ret, 7, 1, today)
        put(ret, 8, 1, "Сидоров")
        facts["return"] = s.B("ReturnPostRow", 1, module="WmsReturn")
        adj = s.doc.Sheets.getByName("Корректировки")
        put(adj, 1, 1, "Инвентаризация")
        put(adj, 2, 1, "ЕИ-00000428")
        put(adj, 7, 1, "102")
        put(adj, 8, 1, "100")
        put(adj, 10, 1, today)
        put(adj, 11, 1, "пересчёт после переноса")
        facts["inventory"] = s.B("AdjustPostRow", 1, module="WmsAdjust")
        osh = s.doc.Sheets.getByName("Заказы")
        r1001 = [r for r in range(1, 40) if osh.getCellByPosition(21, r).getString() == "ЕИ-00001001"][0]
        facts["fix"] = s.Rc("ReceiptFixRow", r1001, "12", "", "УПД-401", osh.getCellByPosition(14, r1001).getString(),
                            osh.getCellByPosition(13, r1001).getString(), "G-01", "")
        facts["fix_row"] = r1001
        s.doc.store()
    finally:
        s.close(save=True)
    bad = {k: v for k, v in facts.items() if k != "fix_row" and not str(v).startswith("OK")}
    if bad:
        raise RuntimeError(f"рабочий день в 0.7.2: {bad}")
    with open(done, "w", encoding="utf-8") as f:
        f.write("\n".join(f"{k}={v}" for k, v in facts.items()) + "\n")
    return book, facts
