"""Benchmark of the issue scenario (Core Phase 2) on large synthetic books.

    WMS_TEST_OUT=<dir> python3 tests/bench_phase2.py [N ...]          default: 100000 250000

For each N a TEST book gets N posted issues in «Выдачи», the matching journal (N ISSUE lines in the real format, with
their W records, spread over 12 monthly files) and a registry «Наличие» of N/5 EIs. Then, with WMS running: open and
startup, 200 postings, the change handler (typing an EI), 50 corrections (quantity and EI), 50 storno, the copy check
(COUNTIF over column A), the unposted count of «Главная», an autofilter saved while active and the reopen after it,
self-check, backup, save, reopen, the start after a save without WMS (full key check of «Выдачи») and the independent
oracle over the whole book.
Times are measured from Python around one UNO call (the call itself costs ~1–3 ms), so they are upper bounds.
Writes bench_phase2.json and BENCH_phase2.md into the output folder.
"""
import datetime
import json
import os
import statistics
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Session, new_wms, journal_files  # noqa: E402
from build_ods import PWD, ISSUE_LOCKS_OPEN  # noqa: E402
import issue_oracle  # noqa: E402
from com.sun.star.sheet import TableFilterField  # noqa: E402
from com.sun.star.sheet.FilterOperator import EQUAL  # noqa: E402
from com.sun.star.util import CellProtection  # noqa: E402

POSTED = "101111111111000111"                  # WmsConfig.ISSUE_LOCKS_POSTED
WHO = ["Ермолин Егор Павлович", "Офис", "Производство", "Иванов Иван Андреевич", "Петров Пётр Сергеевич"]
UNITS = ["шт", "м", "кг", "л", "упак"]
CATS = ["Крепёж", "Кабель", "Электрика", "СИЗ", "Химия", "Инструмент"]
EPOCH = datetime.date(1899, 12, 30)
TODAY = datetime.date.today()
RESULTS = {}


def num(x):
    """a number as WMS writes it into the journal (Str: no trailing zeros, a dot)"""
    return str(int(x)) if x == int(x) else f"{x:.3f}".rstrip("0").rstrip(".")


def qtext(x):
    return num(x).replace(".", ",")


def registry(nei):
    return [(f"ЕИ-{i:08d}", f"Материал {i}", f"M-{i:06d}", UNITS[i % 5], float(1000 + i % 100), f"R-{i % 40:02d}-{i % 7}",
             CATS[i % 6], "", "бенчмарк") for i in range(1, nei + 1)]


def history(n, reg, inst):
    """n posted issues: sheet rows (A:R), the journal lines of their ISSUE operations grouped by month file, and the
    balances after them"""
    bal = {i + 1: row[4] for i, row in enumerate(reg)}
    nei = len(reg)
    rows, files = [], {}
    first = TODAY - datetime.timedelta(days=365)
    for k in range(1, n + 1):
        e = 1 + (k * 7919) % nei
        canon, name, art, unit, _, place, cat, _, _ = reg[e - 1]
        q = 1 + 0.5 * (k % 7)
        before = bal[e]
        after = round(before - q, 3)
        bal[e] = after
        d = first + datetime.timedelta(days=(k - 1) * 365 // n)
        serial = (d - EPOCH).days
        who = WHO[k % 5]
        doc = f"АКТ-{k // 40}" if k % 3 == 0 else ""
        rows.append((float(k), doc, name, art, q, "", unit, float(serial), who, place, cat, canon, "", "", "", before, after,
                     "Проведено"))
        f = [f"NO={k}", f"EI={canon}", f"QTY={num(q)}", f"UNIT={unit}", f"WHO={who}", f"DATE={d.isoformat()}", f"DOC={doc}",
             f"ROW={k + 1}", f"BAL_BEFORE={num(before)}", f"BAL_AFTER={num(after)}",
             f"W=V|Выдачи|{k}|0|E|N{k}|0", f"W=V|Выдачи|{k}|2|S{name}|S{name}|1", f"W=V|Выдачи|{k}|3|S{art}|S{art}|1",
             f"W=V|Выдачи|{k}|4|S{qtext(q)}|N{num(q)}|1", f"W=V|Выдачи|{k}|6|S{unit}|S{unit}|1",
             f"W=V|Выдачи|{k}|7|N{serial}|N{serial}|1", f"W=V|Выдачи|{k}|8|S{who}|S{who}|1",
             f"W=V|Выдачи|{k}|9|S{place}|S{place}|1", f"W=V|Выдачи|{k}|10|S{cat}|S{cat}|1",
             f"W=V|Выдачи|{k}|11|S{canon}|S{canon}|1", f"W=V|Выдачи|{k}|15|N{num(before)}|N{num(before)}|1",
             f"W=V|Выдачи|{k}|16|E|N{num(after)}|1", f"W=V|Выдачи|{k}|17|E|SПроведено|1",
             f"W=L|Выдачи|{k}|0:17|{ISSUE_LOCKS_OPEN}|{POSTED}|0", f"W=V|Наличие|{e}|4|N{num(before)}|N{num(after)}|0",
             f"W=V|_SYS|6|1|N{k}|N{k + 1}|0"]
        body = f"J1;{k};{d.isoformat()}T10:00:00;{inst};ISSUE;" + ";".join(f)
        files.setdefault(f"WMS_journal_{d.year:04d}-{d.month:02d}.csv", []).append(f"{body};END;{len(body)}")
    return rows, files, bal


def protection(locked):
    p = CellProtection()
    p.IsLocked = locked
    return p


def make_book(n):
    """a registered TEST book with n posted issues and their journal (written with macros disabled)"""
    nei = max(n // 5, 1000)
    t0 = time.time()
    p = new_wms(f"bench{n}")
    s = Session(p)                               # the first start registers the book (instance, path, journal folder)
    inst = s.sysv("INSTANCE_ID")
    s.close(save=True)
    reg = registry(nei)
    rows, files, bal = history(n, reg, inst)
    current = [row[:4] + (bal[i + 1],) + row[5:] for i, row in enumerate(reg)]      # «Наличие» after the history
    jdir = os.path.join(os.path.dirname(p), "WMS_Journal")
    for f in journal_files(jdir):
        os.remove(f)
    jbytes, last_file, last_size = 0, "", 0
    for fn in sorted(files):
        data = ("\n".join(files[fn]) + "\n").encode("utf-8")
        open(os.path.join(jdir, fn), "wb").write(data)
        jbytes += len(data)
        last_file, last_size = fn, len(data)
    s = Session(p, macros=0)
    iss, stock = s.doc.Sheets.getByName("Выдачи"), s.doc.Sheets.getByName("Наличие")
    iss.unprotect(PWD)
    stock.unprotect(PWD)
    for r0 in range(0, nei, 20000):
        chunk = current[r0:r0 + 20000]
        stock.getCellRangeByPosition(0, 1 + r0, 8, r0 + len(chunk)).setDataArray(tuple(chunk))
    for r0 in range(0, n, 10000):
        chunk = rows[r0:r0 + 10000]
        iss.getCellRangeByPosition(0, 1 + r0, 17, r0 + len(chunk)).setDataArray(tuple(chunk))
    for c, bit in enumerate(POSTED):
        if bit != ISSUE_LOCKS_OPEN[c]:
            iss.getCellRangeByPosition(c, 1, c, n).CellProtection = protection(bit == "1")
    iss.protect(PWD)
    stock.protect(PWD)
    s.set_sys("LAST_SEQ", n)
    s.set_sys("TX_SEQ", n)
    s.set_sys("TX_STATE", "COMMITTED")
    s.set_sys("NEXT_NO", n + 1)
    s.set_sys("NEXT_EI", nei + 1)
    s.set_sys("JOURNAL_POS", f"{last_file}|{last_size}|{n}")
    s.set_sys("SAVE_STAMP", s.doc.getDocumentProperties().EditingCycles + 1)
    s.doc.store()
    s.close()
    initial = {row[0]: row[4] for row in reg}
    info = dict(rows=n, eis=nei, journal_files=len(files), journal_mb=round(jbytes / 1e6, 1),
                book_mb=round(os.path.getsize(p) / 1e6, 1), build_s=round(time.time() - t0))
    return p, initial, info


def stats(ms):
    ms = sorted(ms)
    return dict(n=len(ms), avg=round(statistics.mean(ms), 1), p95=round(ms[max(0, int(len(ms) * 0.95) - 1)], 1),
                max=round(ms[-1], 1))


def timed(fn, *a, **kw):
    t0 = time.perf_counter()
    r = fn(*a, **kw)
    return r, (time.perf_counter() - t0) * 1000


def startup_ms(rep):
    return int(rep.split("запуск WMS ")[1].split(" мс")[0]) if "запуск WMS " in rep else -1


def set_filter(doc, who):
    """AutoFilter «Кому выдано» = who on «Выдачи»; None removes the filter (an empty filter shows every row — a refresh
    of a range without conditions would leave the hidden rows hidden)"""
    db = doc.DatabaseRanges.getByName("WMS_ISSUES")
    if who is None:
        a = db.getDataArea()
        rng = doc.Sheets.getByIndex(a.Sheet).getCellRangeByPosition(a.StartColumn, a.StartRow, a.EndColumn, a.EndRow)
        fd = rng.createFilterDescriptor(True)
        fd.ContainsHeader = True
        rng.filter(fd)
        return
    fd = db.getFilterDescriptor()
    f = TableFilterField()
    f.Field, f.Operator, f.IsNumeric, f.StringValue = 8, EQUAL, False, who
    fd.setFilterFields((f,))
    db.refresh()


def hidden_rows(doc, n):
    """hidden rows among the first n data rows of «Выдачи»"""
    rows = doc.Sheets.getByName("Выдачи").getRows()
    return sum(1 for r in range(1, n + 1) if not rows.getByIndex(r).IsVisible)


def bench(n):
    r = {}
    p, initial, info = make_book(n)
    r["book"] = info
    print(f"[{n}] книга: {info}", flush=True)
    s = Session(p)
    try:
        rep = s.report()
        r["open_s"] = round(s.open_s, 2)
        r["startup_ms"] = startup_ms(rep)
        r["startup_state"] = rep.split(" | ")[0]
        sh = s.doc.Sheets.getByName("Выдачи")
        nei = info["eis"]
        # 200 postings at the end of the sheet (input through the API: the posting itself is measured)
        base = n + 1
        eis = [1 + (i * 97) % nei for i in range(200)]
        sh.getCellRangeByPosition(4, base, 4, base + 199).setDataArray(tuple((qtext(1 + 0.5 * (i % 3)),) for i in range(200)))
        sh.getCellRangeByPosition(7, base, 7, base + 199).setDataArray(tuple(((TODAY - EPOCH).days,) for _ in range(200)))
        sh.getCellRangeByPosition(8, base, 8, base + 199).setDataArray(tuple((WHO[i % 5],) for i in range(200)))
        sh.getCellRangeByPosition(11, base, 11, base + 199).setDataArray(tuple((str(e),) for e in eis))
        res, ms = [], []
        for i in range(200):
            x, t = timed(s.post, base + i)
            res.append(x)
            ms.append(t)
        r["post"] = stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")))
        print(f"[{n}] проведение: {r['post']}", flush=True)
        # the change handler: an EI typed into L (UI path) against a comment typed into O (handler leaves at once)
        hrow = base + 200
        ms_h, ms_o = [], []
        for i in range(50):
            s.goto("Выдачи", f"$L${hrow + i + 1}")
            _, t = timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName=str(1 + (i * 131) % nei))
            ms_h.append(t)
            s.goto("Выдачи", f"$O${hrow + i + 1}")
            _, t = timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName=f"комментарий {i}")
            ms_o.append(t)
        prev_ok = all(s.iss(hrow + i)[2] == f"Материал {1 + (i * 131) % nei}" for i in range(50))
        r["handler_ei"] = stats(ms_h) | dict(preview_ok=prev_ok)
        r["handler_other_col"] = stats(ms_o)
        # 50 corrections of old rows spread over the sheet: even — quantity (same EI), odd — another EI
        fix_rows = sorted({1 + (i * 4999) % n for i in range(50)})
        res, ms = [], []
        for i, row in enumerate(fix_rows):
            v = s.iss(row)
            e = int(v[11][3:])
            new_ei = str(e) if i % 2 == 0 else str(1 + e % nei)
            x, t = timed(s.I, "IssueFixRow", row, new_ei, qtext(v[4] + 0.5), v[8], TODAY.strftime("%d.%m.%Y"))
            res.append(x)
            ms.append(t)
        r["fix"] = stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0])
        print(f"[{n}] исправление: {r['fix']}", flush=True)
        del_rows = sorted({2 + (i * 3001) % (n - 1) for i in range(50)} - set(fix_rows))
        res, ms = [], []
        for row in del_rows:
            x, t = timed(s.I, "IssueDeleteRow", row)
            res.append(x)
            ms.append(t)
        r["storno"] = stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")))
        print(f"[{n}] сторно: {r['storno']}", flush=True)
        # copy check of a posted row (COUNTIF of its № over column A)
        ms = [timed(s.I, "PreviewRow", 1 + (i * 7777) % n, True, False)[1] for i in range(20)]
        r["copy_check"] = stats(ms)
        cnt, t = timed(s.I, "UnpostedCount")
        r["unposted_count"] = dict(ms=round(t, 1), value=cnt, expected=50)
        _, t = timed(s.U, "UiRefresh", "")
        r["main_refresh_ms"] = round(t, 1)
        r["main_unposted_text"] = s.main_status()[3]
        # autofilter «Кому = Офис» (every fifth row shown), then save with the filter on and reopen: WMS unhides the
        # rows for the save (LibreOffice would open a file with them hidden in quadratic time), keeps the filter
        # conditions and applies the filter again after the save and at the start (D-041)
        _, t1 = timed(set_filter, s.doc, "Офис")
        h1 = hidden_rows(s.doc, 1000)
        _, t2 = timed(s.doc.store)
        h2 = hidden_rows(s.doc, 1000)
        s.close()
        s = Session(p)
        rep = s.report()
        h3 = hidden_rows(s.doc, 1000)
        r["filter_save_reopen"] = dict(apply_ms=round(t1), save_s=round(t2 / 1000, 1), reopen_open_s=round(s.open_s, 2),
                                       reopen_startup_ms=startup_ms(rep), hidden_of_1000=[h1, h2, h3],
                                       modified_after_open=s.doc.isModified(), state=rep.split(" | ")[0])
        _, t3 = timed(set_filter, s.doc, None)
        r["filter_remove_ms"] = round(t3)
        r["hidden_after_remove"] = hidden_rows(s.doc, 1000)
        print(f"[{n}] фильтр: {r['filter_save_reopen']}, снятие {r['filter_remove_ms']} мс", flush=True)
        sc, t = timed(s.B, "ActionSelfCheck")
        r["selfcheck"] = dict(ms=round(t), result=sc.splitlines()[0])
        print(f"[{n}] самопроверка: {r['selfcheck']}", flush=True)
        bk, t = timed(s.B, "ActionBackupNow")
        r["backup"] = dict(ms=round(t), result=bk[:90])
        _, t = timed(s.doc.store)
        r["save_ms"] = round(t)
        s.close()
        # reopen: the saved position is confirmed → only the rest of the journal is read
        s = Session(p)
        rep = s.report()
        r["reopen"] = dict(open_s=round(s.open_s, 2), startup_ms=startup_ms(rep), state=rep.split(" | ")[0])
        s.close()
        # saved without WMS (macros off) → the next start runs the full key check of «Выдачи»
        s = Session(p, macros=0)
        s.doc.Sheets.getByName("Выдачи").getCellByPosition(14, 1).setString("правка без WMS")
        s.doc.store()
        s.close()
        s = Session(p)
        rep = s.report()
        note = [x for x in rep.split(" | ") if x.startswith("Выдачи: строк")]
        r["start_after_macros_off"] = dict(open_s=round(s.open_s, 2), startup_ms=startup_ms(rep), state=rep.split(" | ")[0],
                                           key_check=note[0][:160] if note else "")
        # the oracle: sheet, registry and journal agree over the whole history
        t0 = time.time()
        P, inf = issue_oracle.check(s.doc, s.jdir, initial)
        r["oracle"] = dict(problems=P[:5], n_problems=len(P), info=inf, s=round(time.time() - t0))
        print(f"[{n}] оракул: {len(P)} расхождений {P[:3]}", flush=True)
    finally:
        s.close()
    return r


def main():
    sizes = [int(a) for a in sys.argv[1:]] or [100000, 250000]
    for n in sizes:
        RESULTS[n] = bench(n)
        with open(os.path.join(OUT, "bench_phase2.json"), "w", encoding="utf-8") as f:
            json.dump(RESULTS, f, ensure_ascii=False, indent=1)
    with open(os.path.join(OUT, "BENCH_phase2.md"), "w", encoding="utf-8") as f:
        f.write("| Показатель | " + " | ".join(f"{n} строк" for n in sizes) + " |\n|---|" + "---|" * len(sizes) + "\n")
        keys = ["book", "open_s", "startup_ms", "startup_state", "post", "handler_ei", "handler_other_col", "fix", "storno",
                "copy_check", "unposted_count", "main_refresh_ms", "filter_save_reopen", "filter_remove_ms", "hidden_after_remove",
                "selfcheck", "backup", "save_ms", "reopen", "start_after_macros_off", "oracle"]
        for k in keys:
            f.write(f"| {k} | " + " | ".join(str(RESULTS[n].get(k, "")).replace("|", "¦") for n in sizes) + " |\n")
    print(json.dumps(RESULTS, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
