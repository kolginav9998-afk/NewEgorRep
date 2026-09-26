"""Benchmark of returns (Core Phase 4) on large synthetic books.

    WMS_TEST_OUT=<dir> python3 tests/bench_phase4.py [N ...]          default: 100000 250000

For each N a TEST book gets N movements: 80 % as in tests/bench_phase3.py (receipts on «Заказы» and posted issues of
their EIs on «Выдачи», plus expected and cancelled order positions) and 20 % posted returns on «Возврат» (each by its
own issue: three of four partial, one of four full), with «Наличие», «Заказы».X, _RET, _ISS and the protection of every
row written exactly as WMS writes them (the history itself is not journaled: the journal of the book starts with the
measured operations). Then, with WMS running: open and startup, 100 ordinary returns (issues spread over the sheet),
50 second returns of partly returned issues, 50 corrections, 50 storno, the search of an issue («Найти выдачу»: the
list of an EI, the whole button with its window preset), issue corrections and a refused storno of issues with
returns, the change handler, rows inserted above the returns, self-check, save, reopen, the start after a save without
WMS (full key check) and the independent oracle (trusted snapshot + journal replay).
Times are measured from Python around one UNO call (the call itself costs ~1–3 ms), so they are upper bounds.
Writes bench_phase4.json and BENCH_phase4.md into the output folder.
"""
import datetime
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Session, new_wms, synthetic_registry  # noqa: E402
from build_ods import PWD, ISSUE_LOCKS_OPEN, RETURN_LOCKS_OPEN  # noqa: E402
import bench_phase3 as b3  # noqa: E402
import receipt_oracle  # noqa: E402
import return_oracle  # noqa: E402

RETURN_POSTED = "111111111111110"
EPOCH, TODAY, TSER = b3.EPOCH, b3.TODAY, b3.TSER
RESULTS = {}


def generate(n):
    n_base = int(n * 0.8)
    g = b3.generate(n_base)
    iss, reg, orders = g["iss"], g["reg"], g["orders"]
    n_ret = n - n_base
    order_row = {row[21]: i for i, row in enumerate(orders) if row[21]}
    bal = {row[0]: row[4] for row in reg}
    rets, rt, sums = [], [], {}
    m = len(iss)
    for j in range(n_ret):
        k = 1 + (j * 7919) % m                          # a permutation of the issues (7919 is prime): one return per issue
        row = iss[k - 1]
        e = row[11]
        q = row[4] if j % 4 == 3 else round(row[4] / 2, 3)
        before = bal[e]
        after = round(before + q, 3)
        bal[e] = after
        d = row[7] + 3
        rr = reg[int(e[3:]) - 1]
        n_ = j + 1
        rets.append([float(n_), float(k), e, rr[1], rr[2], q, rr[3], d, row[8], rr[5], rr[6], before, after, "Проведено", ""])
        rt.append((float(n_), float(k), e, q, "LIVE", float(n_)))
        s_, c_ = sums.get(k, (0.0, 0))
        sums[k] = (round(s_ + q, 3), c_ + 1)
    for e, b in bal.items():
        reg[int(e[3:]) - 1][4] = b
        if e in order_row:
            orders[order_row[e]][23] = b
    iss_rows = [((float(k), float(k), sums[k][0], float(sums[k][1])) if k in sums else ("", "", "", "")) for k in range(1, m + 1)]
    g.update(rets=rets, rt=rt, iss_rows=iss_rows, n_ret=n_ret, sums=sums)
    return g


def make_book(n):
    t0 = time.time()
    g = generate(n)
    p = new_wms(f"bench4_{n}")
    s = Session(p)                               # the first start registers the book
    s.close(save=True)
    s = Session(p, macros=0)
    doc = s.doc
    names = ("Заказы", "Выдачи", "Наличие", "_RCV", "_ORD", "Возврат", "_RET", "_ISS")
    for nm in names:
        doc.Sheets.getByName(nm).unprotect(PWD)
    o, iss, st, rc, od, rv, rt, isx = (doc.Sheets.getByName(nm) for nm in names)
    b3.write_rows(st, g["reg"], 9)
    b3.write_rows(rc, g["rcv"], 9, first=201)
    b3.write_rows(od, [list(x) for x in g["ordp"]], 9)
    od.getCellByPosition(11, 0).setValue(g["next_ol"])
    b3.write_rows(o, g["orders"], 28, chunk=5000)
    b3.write_rows(iss, g["iss"], 18)
    b3.write_rows(rv, g["rets"], 15)
    b3.write_rows(rt, g["rt"], 6)
    b3.write_rows(isx, g["iss_rows"], 4)
    no = len(g["orders"])
    for c, bit in enumerate(b3.ORDER_POSTED):
        if bit != b3.ORDER_OPEN[c]:
            o.getCellRangeByPosition(c, 1, c, no).CellProtection = b3.protection(bit == "1")
    from com.sun.star.table import CellRangeAddress
    opened = doc.createInstance("com.sun.star.sheet.SheetCellRanges")
    for i, kind in enumerate(g["kinds"], start=1):
        if kind == "O":
            a = CellRangeAddress()
            a.Sheet, a.StartColumn, a.StartRow, a.EndColumn, a.EndRow = o.getRangeAddress().Sheet, 0, i, 20, i
            opened.addRangeAddress(a, False)
    if opened.getCount():
        opened.CellProtection = b3.protection(False)
    ni = len(g["iss"])
    for c, bit in enumerate(b3.ISSUE_POSTED):
        if bit != ISSUE_LOCKS_OPEN[c]:
            iss.getCellRangeByPosition(c, 1, c, ni).CellProtection = b3.protection(bit == "1")
    nr = g["n_ret"]
    for c, bit in enumerate(RETURN_POSTED):
        if bit != RETURN_LOCKS_OPEN[c]:
            rv.getCellRangeByPosition(c, 1, c, nr).CellProtection = b3.protection(bit == "1")
    for nm in names:
        doc.Sheets.getByName(nm).protect(PWD)
    s.set_sys("NEXT_EI", g["next_ei"])
    s.set_sys("NEXT_NO", ni + 1)
    s.set_sys("NEXT_RET", nr + 1)
    s.set_sys("SAVE_STAMP", doc.getDocumentProperties().EditingCycles + 1)
    doc.store()
    s.close()
    info = dict(movements=n, receipts=g["receipts"], issues=ni, returns=nr, order_rows=no, eis=len(g["reg"]),
                book_mb=round(os.path.getsize(p) / 1e6, 1), build_s=round(time.time() - t0))
    return p, g, info


def bench(n):
    r = {}
    p, g, info = make_book(n)
    r["book"] = info
    print(f"[{n}] книга: {info}", flush=True)
    s = Session(p)
    try:
        s.U("TestUiAuto", 1)
        rep = s.report()
        r["open_s"] = round(s.open_s, 2)
        r["startup_ms"] = b3.startup_ms(rep)
        r["startup_state"] = rep.split(" | ")[0]
        print(f"[{n}] открытие {r['open_s']} с, запуск {r['startup_ms']} мс; {r['startup_state']}", flush=True)
        base = receipt_oracle.snapshot_base(s.doc, {row[0]: row[4] for row in synthetic_registry()})
        rv = s.doc.Sheets.getByName("Возврат")
        m = len(g["iss"])
        nr = g["n_ret"]
        today = TODAY.strftime("%d.%m.%Y")
        # 1. 100 ordinary returns by issues without returns, spread over «Выдачи» (input through the API: the posting is measured)
        free = [k for k in (1 + (j * 104729 + 17) % m for j in range(400)) if k not in g["sums"]][:100]
        top = nr + 1
        rv.getCellRangeByPosition(1, top, 1, top + 99).setDataArray(tuple((float(k),) for k in free))
        rv.getCellRangeByPosition(5, top, 5, top + 99).setDataArray(tuple(("0,25",) for _ in free))
        rv.getCellRangeByPosition(7, top, 7, top + 99).setDataArray(tuple((float(TSER),) for _ in free))
        res, ms = [], []
        for i in range(100):
            x, t = b3.timed(s.Rt, "ReturnPostRow", top + i)
            res.append(x)
            ms.append(t)
        r["return"] = b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:60])
        print(f"[{n}] обычный возврат: {r['return']}", flush=True)
        # 2. 50 second returns of issues returned in part (the rest of each)
        part = [k for k, (q, c) in g["sums"].items() if q < 0.5 - 1e-9][:50]
        top2 = top + 100
        rv.getCellRangeByPosition(1, top2, 1, top2 + len(part) - 1).setDataArray(tuple((float(k),) for k in part))
        rv.getCellRangeByPosition(5, top2, 5, top2 + len(part) - 1).setDataArray(tuple(("0,25",) for _ in part))
        rv.getCellRangeByPosition(7, top2, 7, top2 + len(part) - 1).setDataArray(tuple((float(TSER),) for _ in part))
        res, ms = [], []
        for i in range(len(part)):
            x, t = b3.timed(s.Rt, "ReturnPostRow", top2 + i)
            res.append(x)
            ms.append(t)
        r["second_return"] = b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:60])
        print(f"[{n}] повторный возврат: {r['second_return']}", flush=True)
        # 3. 50 corrections of returns spread over the sheet (the partial ones: 0,25 → 0,2)
        partial_rows = [j + 1 for j in range(nr) if g["rets"][j][5] < 0.5 - 1e-9 and g["rets"][j][1] not in part]
        fx = [partial_rows[(i * 7919 + 3) % len(partial_rows)] for i in range(50)]
        fx = list(dict.fromkeys(fx))
        res, ms = [], []
        for row in fx:
            x, t = b3.timed(s.Rt, "ReturnFixRow", row, "0,2", (EPOCH + datetime.timedelta(days=int(g["rets"][row - 1][7]))).strftime("%d.%m.%Y"),
                            g["rets"][row - 1][9])
            res.append(x)
            ms.append(t)
        r["fix"] = b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:80])
        print(f"[{n}] исправление: {r['fix']}", flush=True)
        # 4. 50 storno of returns spread over the sheet
        dl = [j + 1 for j in range(nr) if (j + 1) not in fx and g["rets"][j][1] not in part]
        dl = [dl[(i * 104729 + 5) % len(dl)] for i in range(50)]
        dl = list(dict.fromkeys(dl))
        res, ms = [], []
        for row in dl:
            x, t = b3.timed(s.Rt, "ReturnDeleteRow", row)
            res.append(x)
            ms.append(t)
        r["storno"] = b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:80])
        print(f"[{n}] сторно: {r['storno']}", flush=True)
        # 5. search of an issue: the list of an EI (one array formula + the key check of each candidate), the row of an issue № by
        #    its key, the whole «Найти выдачу» button (window values preset) on a new row with the EI in C
        eis = [g["iss"][(i * 7919) % m][11] for i in range(20)]
        ms_l, found = [], []
        for e in eis:
            x, t = b3.timed(s.Rt, "IssuesOfEI", e, "", True)
            ms_l.append(t)
            found.append(len(x.split(";")) if x else 0)
        r["find_list"] = b3.stats(ms_l) | dict(issues_per_ei=found[:8])
        ms_f = []
        for i in range(20):
            k = 1 + (i * 3571) % m
            _, t = b3.timed(s.Rt, "FindIssueRow", k, False)
            ms_f.append(t)
        r["find_issue_row"] = b3.stats(ms_f)
        top3 = top2 + len(part)
        ms_b, res = [], []
        for i in range(10):
            row = top3 + i
            rv.getCellByPosition(2, row).setString(eis[i])
            s.goto("Возврат", f"$C${row + 1}")
            s.RU("TestRetInput", "0,1", today, "<как предложено>", -1)
            _, t = b3.timed(s.RU, "BtnRetFind")
            ms_b.append(t)
            res.append(s.U("TestUiLastMessage")[:50])
        r["find_button"] = b3.stats(ms_b) | dict(first=res[0])
        print(f"[{n}] поиск выдачи: список {r['find_list']}; строка {r['find_issue_row']}; кнопка {r['find_button']}", flush=True)
        # 6. issues with a partial return (untouched above): the quantity corrected down to 0,3, not below the returned 0,25
        #    (allowed), and a refused storno (one lookup in _ISS each)
        with_ret = [int(g["rets"][j][1]) for j in range(0, nr, max(1, nr // 200)) if (j + 1) not in fx and (j + 1) not in dl
                    and g["rets"][j][5] < 0.5 - 1e-9 and int(g["rets"][j][1]) not in part][:20]
        ms_if, ms_id, res_if, res_id = [], [], [], []
        for k in with_ret:
            row = g["iss"][k - 1]
            x, t = b3.timed(s.I, "IssueFixRow", k, row[11], "0,3", row[8], (EPOCH + datetime.timedelta(days=int(row[7]))).strftime("%d.%m.%Y"))
            ms_if.append(t)
            res_if.append(x)
            x, t = b3.timed(s.I, "IssueDeleteRow", k)
            ms_id.append(t)
            res_id.append(x)
        r["issue_fix_with_returns"] = b3.stats(ms_if) | dict(ok=sum(1 for x in res_if if x.startswith("OK")), first=res_if[0][:60])
        r["issue_del_refused"] = b3.stats(ms_id) | dict(refused=sum(1 for x in res_id if x.startswith("ERR")), first=res_id[0][:80])
        print(f"[{n}] ISSUE_FIX {r['issue_fix_with_returns']}; ISSUE_DEL {r['issue_del_refused']}", flush=True)
        # 6a. D-069: the storno of a receipt checks the live issues and returns of its EI (one formula of the Calc engine over
        #     the used rows): 20 receipts whose EI has issues — refused; 20 new receipts without issues — posted, then cancelled
        with_iss = {row[11] for row in g["iss"]}
        src_rows = [i + 1 for i, k in enumerate(g["kinds"]) if k == "S" and g["orders"][i][21] in with_iss]
        att = [src_rows[(i * 7919 + 1) % len(src_rows)] for i in range(20)]
        ms_rd, res_rd = [], []
        for row in att:
            x, t = b3.timed(s.Rc, "ReceiptDeleteRow", row)
            ms_rd.append(t)
            res_rd.append(x)
        r["receipt_del_refused"] = b3.stats(ms_rd) | dict(refused=sum(1 for x in res_rd if x.startswith("ERR:По этому ЕИ")), first=res_rd[0][:90])
        o = s.doc.Sheets.getByName("Заказы")
        top_o = len(g["orders"]) + 1
        new = [(f"Б4-{i}", f"Новый товар {i}", "", f"СЧ-Б4-{i}", f"N4-{i:04d}", "5", "5", "10", "шт", "", "", b3.SUPPLIERS[i % 5], "",
                float(TSER), float(TSER - 1), float(TSER - 10), float(TSER + 10), "", "Крепёж", "", "Z-4") for i in range(20)]
        o.getCellRangeByPosition(0, top_o, 20, top_o + 19).setDataArray(tuple(new))
        posted = [s.Rc("ReceiptPostRow", top_o + i) for i in range(20)]
        ms_ok, res_ok = [], []
        for i in range(20):
            x, t = b3.timed(s.Rc, "ReceiptDeleteRow", top_o + i)
            ms_ok.append(t)
            res_ok.append(x)
        r["receipt_del_ok"] = b3.stats(ms_ok) | dict(ok=sum(1 for x in res_ok if x.startswith("OK")), posted=sum(1 for x in posted if x.startswith("OK")))
        print(f"[{n}] RECEIPT_DEL: отказ (есть выдачи) {r['receipt_del_refused']}; разрешён {r['receipt_del_ok']}", flush=True)
        # 7. the change handler: the issue № typed into a new row (preview), a comment O (the handler leaves at once)
        ms_b2, ms_o = [], []
        for i in range(20):
            row = top3 + 20 + i
            s.goto("Возврат", f"$B${row + 1}")
            _, t = b3.timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName=str(1 + (i * 3571 + 11) % m))
            ms_b2.append(t)
            s.goto("Возврат", f"$O${row + 1}")
            _, t = b3.timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName=f"комментарий {i}")
            ms_o.append(t)
        r["handler_issue_no"] = b3.stats(ms_b2)
        r["handler_comment"] = b3.stats(ms_o)
        print(f"[{n}] обработчик: № выдачи {r['handler_issue_no']}; комментарий {r['handler_comment']}", flush=True)
        # 8. rows inserted above the returns: the next operation on a far row finds it by its key (one MATCH) and fixes the hint
        s.goto("Возврат", "$A$2")
        s.ui(".uno:SelectRow")
        _, t_ins = b3.timed(s.ui, ".uno:InsertRowsBefore")
        far = [j + 1 for j in range(nr - 1, 0, -1) if (j + 1) not in fx and (j + 1) not in dl and g["rets"][j][5] < 0.5 - 1e-9][0]
        x, t = b3.timed(s.Rt, "ReturnFixRow", far + 1, "0,15", (EPOCH + datetime.timedelta(days=int(g["rets"][far - 1][7]))).strftime("%d.%m.%Y"),
                        g["rets"][far - 1][9])
        x2, t2 = b3.timed(s.Rt, "ReturnFixRow", far + 1, "0,1", (EPOCH + datetime.timedelta(days=int(g["rets"][far - 1][7]))).strftime("%d.%m.%Y"),
                          g["rets"][far - 1][9])
        r["after_insert"] = dict(insert_ms=round(t_ins), first_op_ms=round(t), second_op_ms=round(t2), results=[x[:40], x2[:40]])
        print(f"[{n}] после вставки строки: {r['after_insert']}", flush=True)
        sc, t = b3.timed(s.B, "ActionSelfCheck")
        r["selfcheck"] = dict(ms=round(t), result=sc.splitlines()[0], returns=[x[:220] for x in sc.splitlines() if "возвраты" in x][:1],
                              keys=[x[:400] for x in sc.splitlines() if x.startswith(("OK ключи", "WARN ключи", "FAIL ключи"))][:1])
        print(f"[{n}] самопроверка: {r['selfcheck']}", flush=True)
        _, t = b3.timed(s.doc.store)
        r["save_ms"] = round(t)
        t0 = time.time()
        P, inf, hist = return_oracle.check(s.doc, s.jdir, {}, base=base)
        r["oracle"] = dict(problems=P[:5], n_problems=len(P), info={k: inf.get(k) for k in ("ops", "return_ops", "returns", "live_returns")},
                           recipient_pairs=len(hist), s=round(time.time() - t0))
        print(f"[{n}] оракул: {len(P)} расхождений {P[:3]}", flush=True)
        s.close()
        s = Session(p)
        rep = s.report()
        r["reopen"] = dict(open_s=round(s.open_s, 2), startup_ms=b3.startup_ms(rep), state=rep.split(" | ")[0])
        s.close()
        s = Session(p, macros=0)
        s.doc.Sheets.getByName("Возврат").getCellByPosition(14, 1).setString("правка без WMS")
        s.doc.store()
        s.close()
        s = Session(p)
        rep = s.report()
        r["start_after_macros_off"] = dict(open_s=round(s.open_s, 2), startup_ms=b3.startup_ms(rep), state=rep.split(" | ")[0],
                                           key_check=[x[:160] for x in rep.split("; ") if "Возврат: строк" in x][:1])
        print(f"[{n}] повторное открытие {r['reopen']}; после сохранения без макросов {r['start_after_macros_off']}", flush=True)
    finally:
        s.close()
    return r


def main():
    sizes = [int(a) for a in sys.argv[1:]] or [100000, 250000]
    for n in sizes:
        RESULTS[n] = bench(n)
        with open(os.path.join(OUT, "bench_phase4.json"), "w", encoding="utf-8") as f:
            json.dump(RESULTS, f, ensure_ascii=False, indent=1)
    with open(os.path.join(OUT, "BENCH_phase4.md"), "w", encoding="utf-8") as f:
        f.write("| Показатель | " + " | ".join(f"{n} движений" for n in sizes) + " |\n|---|" + "---|" * len(sizes) + "\n")
        for k in RESULTS[sizes[0]]:
            f.write(f"| {k} | " + " | ".join(str(RESULTS[n].get(k, "")).replace("|", "¦") for n in sizes) + " |\n")
    print(json.dumps(RESULTS, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
