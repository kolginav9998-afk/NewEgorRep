"""Benchmark of orders and receipts (Core Phase 3) on large synthetic books.

    WMS_TEST_OUT=<dir> python3 tests/bench_phase3.py [N ...]          default: 100000 250000

For each N a TEST book gets N movements: 40 % receipts on «Заказы» (3 of 4 on the source row of their position, 1 of 4
an additional delivery in its own row), 60 % posted issues of those EIs on «Выдачи»; plus order positions still expected
(every 50th row, half of them overdue) and cancelled ones (every 200th row); every 10th source row has an expected date Q
20 days ahead (its partly received positions become «Частично получено / просрочено» in the refresh 45 days later). «Наличие», _RCV, _ORD, the antidubl index
and the protection of every row are written exactly as WMS writes them (the history itself is not journaled: the
journal of the book starts with the measured operations). Then, with WMS running: open and startup (with the refresh of
the statuses), 100 receipts (each creates an EI), 50 second deliveries of positions spread over the sheet, 50 corrections,
50 storno, 50 issues of received EIs (the X mirror), cancellations, the status refresh (as it is and 30 days later), the
antidubl and row lookups, the change handler, rows inserted above the data, self-check, save, reopen, the start after a
save without WMS (full key check) and the independent oracle (trusted snapshot + journal replay).
Times are measured from Python around one UNO call (the call itself costs ~1–3 ms), so they are upper bounds.
Writes bench_phase3.json and BENCH_phase3.md into the output folder.
"""
import datetime
import json
import os
import statistics
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Session, new_wms, synthetic_registry  # noqa: E402
from build_ods import PWD, ISSUE_LOCKS_OPEN  # noqa: E402
import receipt_oracle  # noqa: E402
from com.sun.star.util import CellProtection  # noqa: E402
from com.sun.star.table import CellRangeAddress  # noqa: E402

ISSUE_POSTED = "101111111111000111"
ORDER_POSTED = "1111111111010111001011111001"
ORDER_OPEN = "0000000000000000000001111001"
WHO = ["Ермолин Егор Павлович", "Офис", "Производство", "Иванов Иван Андреевич", "Петров Пётр Сергеевич"]
UNITS = ["шт", "м", "кг", "л", "упак"]
CATS = ["Крепёж", "Кабель", "Электрика", "СИЗ", "Химия", "Инструмент"]
SUPPLIERS = ["Петрович", "ВсеИнструменты", "Озон", "ЭТМ", "Леруа"]
EPOCH = datetime.date(1899, 12, 30)
TODAY = datetime.date.today()
TSER = (TODAY - EPOCH).days
RESULTS = {}


def num(x):
    return str(int(x)) if x == int(x) else f"{x:.3f}".rstrip("0").rstrip(".")


def qtext(x):
    return num(x).replace(".", ",")


# ---------------------------------------------------------------- the same normalisation and hash as WmsOrders
def norm_part(s, max_len):
    out, gap = "", False
    for ch in s.strip().lower():
        code = ord(ch)
        if code == 1105:
            ch, code = "е", 1077
        if 48 <= code <= 57 or 97 <= code <= 122 or 1072 <= code <= 1103:
            if gap and out:
                out += "_"
            out += ch
            gap = False
            if len(out) >= max_len:
                break
        else:
            gap = True
    return out


def dup_key(supplier, doc, art, name, fact, dt):
    p3 = "a" + norm_part(art, 40) if art.strip() else "n" + norm_part(name, 60)
    return f"{norm_part(supplier, 40)}-{norm_part(doc, 40)}-{p3}-{num(fact).replace('.', 'd')}-{int(dt)}"


def dup_hash(s):
    p, h = 34359738337, 7
    for ch in s:
        h = (h * 131071 + ord(ch)) % p
    return float(h + 1)


def fp_s(v):
    return "E" if v == "" else ("N" + num(v) if isinstance(v, float) else "S" + v)


def status(ordq, rcv, nodoc, cancel, q):
    return receipt_oracle.position_status(ordq, rcv, nodoc, cancel, q, TSER)


def protection(locked):
    p = CellProtection()
    p.IsLocked = locked
    return p


def generate(n):
    """rows of «Заказы», «Выдачи», «Наличие», _RCV, _ORD for n movements"""
    n_rcpt = int(n * 0.4)
    n_iss = n - n_rcpt
    synth = synthetic_registry()
    reg = [list(r) for r in synth]                     # EIs 1..200 of the test book stay
    rcv_rows, ord_rows, orders, row_kind = [], [], [], []
    receipts = []                                      # (ei, row index, qty)
    first = TODAY - datetime.timedelta(days=400)
    ei, olid, k = 201, 1, 0
    open_rows = cancelled = 0
    src_of = {}
    while k < n_rcpt:
        r = len(orders) + 1
        if r % 50 == 0 or r % 200 == 7:
            # an expected position (half overdue) or a cancelled one
            a, name, art = f"З-{r // 3}", f"Позиция заказа {r}", f"P-{r:06d}"
            unit, sup = UNITS[r % 5], SUPPLIERS[r % 5]
            pdate = float((first + datetime.timedelta(days=r % 300) - EPOCH).days)
            q = float(TSER - 5 if r % 100 == 0 else TSER + 30)
            if r % 200 == 7:
                h = 10.0
                fp = "|".join([fp_s(a), fp_s(f"СЧ-{r // 3}"), fp_s(name), fp_s(art), fp_s(h), fp_s(unit), fp_s(sup), fp_s(pdate)])
                orders.append([a, name, "", f"СЧ-{r // 3}", art, "", "", h, unit, "", "", sup, "", "", "", pdate, q, "", "", "", "", "",
                               "Отменено", "", "Заказ отменён; ЕИ не создавался", "", "", ""])
                ord_rows.append((float(olid), float(r), "", fp, h, 0.0, 0.0, 0.0, "ORDER"))
                olid += 1
                row_kind.append("P")
                cancelled += 1
            else:
                orders.append([a, name, "", f"СЧ-{r // 3}", art, "", "", "10", unit, "", "", sup, "", "", "", pdate, q, "", "", "", "", "",
                               "Просрочено" if q < TSER else "Ожидается", "", "", "", "", ""])
                row_kind.append("O")
                open_rows += 1
            continue
        if k % 4 == 3 and src_of:
            # an additional delivery of an earlier position (its source row data are copied)
            s_ol = 1 + (k * 7) % (olid - 1)
            if s_ol not in src_of:
                s_ol = next(iter(src_of))
            srow, s_ord = src_of[s_ol]
            base = orders[srow - 1]
            qty = float(1 + k % 5)
            d = float((first + datetime.timedelta(days=k * 380 // n_rcpt) - EPOCH).days)
            doc = f"УПД-{k}"
            row = [base[0], base[1], doc, base[3], base[4], qty, qty, "", base[8], base[9], "", base[11], "", d, d, base[15], "", "",
                   base[18], "", f"R-{k % 40:02d}", f"ЕИ-{ei:08d}", "Дополнительное поступление", qty, "Норма", "", "", ""]
            orders.append(row)
            row_kind.append("A")
            key = dup_key(base[11], doc, base[4], base[1], qty, d)
            rcv_rows.append((f"ЕИ-{ei:08d}", float(s_ol), float(r), dup_hash(key), "LIVE", "ADD", qty, 0.0, key))
            s_ord[5] += qty
            s_ord[6] += 1
            reg.append([f"ЕИ-{ei:08d}", base[1], base[4], base[8], qty, f"R-{k % 40:02d}", base[18], "Активен", f"Заказ {base[0]}"])
            receipts.append((ei, r - 1 + 1, qty))
            ei += 1
            k += 1
            continue
        a, name, art = f"З-{r // 3}", f"Материал {r}", f"M-{r:06d}"
        unit, sup, cat = UNITS[r % 5], SUPPLIERS[r % 5], CATS[r % 6]
        h = float(10 + r % 20)
        qty = h if k % 3 else float(max(1, int(h) // 2))
        d = float((first + datetime.timedelta(days=k * 380 // n_rcpt) - EPOCH).days)
        pdate = d - 7
        doc = f"УПД-{k}"
        price = float(10 + r % 90)
        q_src = float(TSER + 20) if r % 10 == 1 else ""
        row = [a, name, doc, f"СЧ-{r // 3}", art, qty, qty, h, unit, price, "", sup, "", d, d - 1, pdate, q_src, "", cat, "", f"S-{r % 40:02d}",
               f"ЕИ-{ei:08d}", "", qty, "", "", "", ""]
        orders.append(row)
        row_kind.append("S")
        key = dup_key(sup, doc, art, name, qty, d - 1)
        rcv_rows.append((f"ЕИ-{ei:08d}", float(olid), float(r), dup_hash(key), "LIVE", "SRC", qty, 0.0, key))
        fp = "|".join([fp_s(a), fp_s(f"СЧ-{r // 3}"), fp_s(name), fp_s(art), fp_s(h), fp_s(unit), fp_s(sup), fp_s(pdate)])
        pos = [float(olid), float(r), f"ЕИ-{ei:08d}", fp, h, qty, 1.0, 0.0, ""]
        ord_rows.append(pos)
        src_of[olid] = (r, pos)
        reg.append([f"ЕИ-{ei:08d}", name, art, unit, qty, f"S-{r % 40:02d}", cat, "Активен", f"Заказ {a}"])
        receipts.append((ei, r, qty))
        ei += 1
        olid += 1
        k += 1
    # statuses and summaries of the source rows
    for pos in ord_rows:
        if pos[2]:
            r = int(pos[1])
            row = orders[r - 1]
            row[22] = status(pos[4], pos[5], 0, "", row[16] if row[16] != "" else None)
            row[24] = f"Норма; позиция: получено {qtext(pos[5])} из {qtext(pos[4])} {row[8]}"
            if pos[5] < pos[4] - 1e-9:
                row[24] += f", осталось {qtext(round(pos[4] - pos[5], 3))}"
            elif pos[5] > pos[4] + 1e-9:
                row[24] += f" (больше заказа на {qtext(round(pos[5] - pos[4], 3))})"
    # issues of received EIs (balances go down; about a quarter of every EI is issued)
    bal = {e: q for e, _, q in receipts}
    rowof = {e: r for e, r, _ in receipts}
    iss = []
    rlist = [e for e, _, _ in receipts]
    for j in range(1, n_iss + 1):
        e = rlist[(j * 7919) % len(rlist)]
        q = 0.5 if bal[e] >= 0.5 else 0.0
        if q == 0.0:
            continue
        before = bal[e]
        after = round(before - q, 3)
        bal[e] = after
        rr = reg[e - 1]
        dd = float((first + datetime.timedelta(days=30 + j * 360 // n_iss) - EPOCH).days)
        iss.append((float(len(iss) + 1), "", rr[1], rr[2], q, "", rr[3], dd, WHO[j % 5], rr[5], rr[6], rr[0], "", "", "", before, after, "Проведено"))
    for e, b in bal.items():
        reg[e - 1][4] = b
        orders[rowof[e] - 1][23] = b
    partial = [int(pos[1]) for pos in ord_rows if pos[2] and pos[5] < pos[4] - 1e-9]
    return dict(orders=orders, kinds=row_kind, reg=reg, rcv=rcv_rows, ordp=ord_rows, iss=iss, next_ei=ei, next_ol=olid,
                receipts=len(receipts), open_rows=open_rows, cancelled=cancelled, partial_rows=partial)


def write_rows(sh, rows, ncols, chunk=10000, first=1):
    for r0 in range(0, len(rows), chunk):
        part = rows[r0:r0 + chunk]
        sh.getCellRangeByPosition(0, first + r0, ncols - 1, first + r0 + len(part) - 1).setDataArray(tuple(tuple(x) for x in part))


def make_book(n):
    t0 = time.time()
    g = generate(n)
    p = new_wms(f"bench3_{n}")
    s = Session(p)                               # the first start registers the book
    s.close(save=True)
    s = Session(p, macros=0)
    doc = s.doc
    names = ("Заказы", "Выдачи", "Наличие", "_RCV", "_ORD")
    for nm in names:
        doc.Sheets.getByName(nm).unprotect(PWD)
    o, iss, st, rc, od = (doc.Sheets.getByName(nm) for nm in names)
    write_rows(st, g["reg"], 9)
    # _RCV and _ORD are dense: row = EI / OLID (the receipts start at EI 201, after the test registry)
    write_rows(rc, g["rcv"], 9, first=201)
    write_rows(od, [list(x) for x in g["ordp"]], 9)
    od.getCellByPosition(11, 0).setValue(g["next_ol"])
    write_rows(o, g["orders"], 28, chunk=5000)
    write_rows(iss, g["iss"], 18)
    # protection: every order row as WMS leaves it (posted / cancelled rows closed, expected rows open)
    no = len(g["orders"])
    for c, bit in enumerate(ORDER_POSTED):
        if bit != ORDER_OPEN[c]:
            o.getCellRangeByPosition(c, 1, c, no).CellProtection = protection(bit == "1")
    opened = doc.createInstance("com.sun.star.sheet.SheetCellRanges")
    for i, kind in enumerate(g["kinds"], start=1):
        if kind == "O":
            a = CellRangeAddress()
            a.Sheet, a.StartColumn, a.StartRow, a.EndColumn, a.EndRow = o.getRangeAddress().Sheet, 0, i, 20, i
            opened.addRangeAddress(a, False)
    if opened.getCount():
        opened.CellProtection = protection(False)
    ni = len(g["iss"])
    for c, bit in enumerate(ISSUE_POSTED):
        if bit != ISSUE_LOCKS_OPEN[c]:
            iss.getCellRangeByPosition(c, 1, c, ni).CellProtection = protection(bit == "1")
    for nm in names:
        doc.Sheets.getByName(nm).protect(PWD)
    s.set_sys("NEXT_EI", g["next_ei"])
    s.set_sys("NEXT_NO", ni + 1)
    s.set_sys("SAVE_STAMP", doc.getDocumentProperties().EditingCycles + 1)
    doc.store()
    s.close()
    info = dict(movements=n, receipts=g["receipts"], issues=ni, order_rows=no, expected_positions=g["open_rows"], cancelled=g["cancelled"],
                eis=len(g["reg"]), book_mb=round(os.path.getsize(p) / 1e6, 1), build_s=round(time.time() - t0))
    return p, g, info


def stats(ms):
    ms = sorted(ms)
    return dict(n=len(ms), avg=round(statistics.mean(ms), 1), p95=round(ms[max(0, int(len(ms) * 0.95) - 1)], 1), max=round(ms[-1], 1))


def timed(fn, *a, **kw):
    t0 = time.perf_counter()
    r = fn(*a, **kw)
    return r, (time.perf_counter() - t0) * 1000


def startup_ms(rep):
    return int(rep.split("запуск WMS ")[1].split(" мс")[0]) if "запуск WMS " in rep else -1


def note(rep, prefix):
    x = [p for p in rep.split(" | ") if p.startswith(prefix)]
    return x[0][:200] if x else ""


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
        r["startup_ms"] = startup_ms(rep)
        r["startup_refresh"] = note(rep, "статусы заказов")
        r["startup_state"] = rep.split(" | ")[0]
        print(f"[{n}] открытие {r['open_s']} с, запуск {r['startup_ms']} мс; {r['startup_refresh']}", flush=True)
        base = receipt_oracle.snapshot_base(s.doc, {row[0]: row[4] for row in synthetic_registry()})
        o = s.doc.Sheets.getByName("Заказы")
        kinds = g["kinds"]
        nrows = len(kinds)
        # 1. 100 receipts on new order rows at the end of the sheet (input through the API: the posting is measured)
        top = nrows + 1
        new = [(f"Б-{i}", f"Новый товар {i}", "", f"СЧ-Б{i}", f"N-{i:04d}", "5", "5", "10", "шт", "", "", SUPPLIERS[i % 5], "",
                float(TSER), float(TSER - 1), float(TSER - 10), float(TSER + 10), "", "Крепёж", "", "Z-1") for i in range(100)]
        o.getCellRangeByPosition(0, top, 20, top + 99).setDataArray(tuple(new))
        res, ms = [], []
        for i in range(100):
            x, t = timed(s.Rc, "ReceiptPostRow", top + i)
            res.append(x)
            ms.append(t)
        r["receipt"] = stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0])
        print(f"[{n}] приход (с новым ЕИ): {r['receipt']}", flush=True)
        # 2. 50 second deliveries of positions spread over the sheet
        src_rows = [i + 1 for i, k in enumerate(kinds) if k == "S"]
        pick = [src_rows[(i * 7919) % len(src_rows)] for i in range(50)]
        res, ms = [], []
        for i, row in enumerate(pick):
            x, t = timed(s.Rc, "ReceiptAddRow", row, "2", "2", f"УПД-Б{i}", TODAY.strftime("%d.%m.%Y"), TODAY.strftime("%d.%m.%Y"), "Q-1", "")
            res.append(x)
            ms.append(t)
        r["second_delivery"] = stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:60])
        print(f"[{n}] второе поступление: {r['second_delivery']}", flush=True)
        # 3. 50 corrections (the quantity up by 1) of receipts spread over the sheet
        fx_rows = [src_rows[(i * 104729 + 13) % len(src_rows)] for i in range(50)]
        res, ms = [], []
        for i, row in enumerate(fx_rows):
            v = s.ords(row)
            x, t = timed(s.Rc, "ReceiptFixRow", row, qtext(v[5] + 1), qtext(v[6]), v[2],
                         (EPOCH + datetime.timedelta(days=int(v[14]))).strftime("%d.%m.%Y"),
                         (EPOCH + datetime.timedelta(days=int(v[13]))).strftime("%d.%m.%Y"), v[20], qtext(v[9]) if v[9] != "" else "")
            res.append(x)
            ms.append(t)
        r["fix"] = stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:80])
        print(f"[{n}] исправление: {r['fix']}", flush=True)
        # 4. 50 storno of receipts without issues (the new ones of step 1)
        res, ms = [], []
        for i in range(50):
            x, t = timed(s.Rc, "ReceiptDeleteRow", top + 50 + i)
            res.append(x)
            ms.append(t)
        r["storno"] = stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")))
        print(f"[{n}] сторно: {r['storno']}", flush=True)
        # 5. 50 issues of received EIs (the issue also writes X of the receipt row)
        ish = s.doc.Sheets.getByName("Выдачи")
        ibase = len(g["iss"]) + 1
        eis = [201 + (i * 331) % g["receipts"] for i in range(50)]
        ish.getCellRangeByPosition(4, ibase, 4, ibase + 49).setDataArray(tuple(("0,5",) for _ in range(50)))
        ish.getCellRangeByPosition(7, ibase, 7, ibase + 49).setDataArray(tuple((float(TSER),) for _ in range(50)))
        ish.getCellRangeByPosition(8, ibase, 8, ibase + 49).setDataArray(tuple((WHO[i % 5],) for i in range(50)))
        ish.getCellRangeByPosition(11, ibase, 11, ibase + 49).setDataArray(tuple((str(e),) for e in eis))
        res, ms = [], []
        for i in range(50):
            x, t = timed(s.I, "IssuePostRow", ibase + i)
            res.append(x)
            ms.append(t)
        r["issue_with_mirror"] = stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:60])
        print(f"[{n}] выдача (с X): {r['issue_with_mirror']}", flush=True)
        # 6. cancellations: 20 expected positions, 20 rests of partial positions
        open_rows = [i + 1 for i, k in enumerate(kinds) if k == "O"]
        res, ms = [], []
        for i in range(20):
            x, t = timed(s.Rc, "OrderCancelRow", open_rows[(i * 37) % len(open_rows)])
            res.append(x)
            ms.append(t)
        r["cancel_order"] = stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")))
        partial = [row for row in g["partial_rows"] if row not in pick][:20]
        res, ms = [], []
        for row in partial:
            x, t = timed(s.Rc, "OrderCancelRestRow", row)
            res.append(x)
            ms.append(t)
        r["cancel_rest"] = stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")))
        print(f"[{n}] отмены: {r['cancel_order']} / {r['cancel_rest']}", flush=True)
        # 7. status refresh: as it is, then 30 days later (many expected positions become overdue), then back
        x, t = timed(s.Od, "RefreshStatuses")
        r["refresh"] = dict(ms=round(t), result=x)
        s.Od("TestSetToday", (TODAY + datetime.timedelta(days=45)).strftime("%d.%m.%Y"))
        x, t = timed(s.Od, "RefreshStatuses")
        r["refresh_later"] = dict(ms=round(t), result=x)
        s.Od("TestSetToday", "")
        x, t = timed(s.Od, "RefreshStatuses")
        r["refresh_back"] = dict(ms=round(t), result=x)
        print(f"[{n}] обновление статусов: {r['refresh']} / {r['refresh_later']} / {r['refresh_back']}", flush=True)
        # 8. antidubl lookups (an existing key, a missing key) and row lookups by EI
        keys = [g["rcv"][(i * 977) % len(g["rcv"])][8] for i in range(20)]
        ms_hit = [timed(s.Od, "FindDuplicate", k, 0)[1] for k in keys]
        ms_miss = [timed(s.Od, "FindDuplicate", k + "x", 0)[1] for k in keys]
        hit_ok = sum(1 for k in keys if s.Od("FindDuplicate", k, 0).startswith("ЕИ-"))
        r["antidubl_hit"] = stats(ms_hit) | dict(found=hit_ok)
        r["antidubl_miss"] = stats(ms_miss)
        ms_l = [timed(s.Od, "FindEIRow", f"ЕИ-{201 + (i * 1237) % g['receipts']:08d}", 0)[1] for i in range(20)]
        r["row_lookup"] = stats(ms_l)
        ms_c = [timed(s.Od, "CountEI", f"ЕИ-{201 + (i * 1237) % g['receipts']:08d}")[1] for i in range(20)]
        r["count_lookup"] = stats(ms_c)
        print(f"[{n}] антидубль {r['antidubl_hit']} / {r['antidubl_miss']}; поиск строки {r['row_lookup']}", flush=True)
        # 9. the change handler: Q typed into an expected row (preview), Z typed (the handler leaves at once)
        ms_q, ms_z = [], []
        for i in range(20):
            row = open_rows[(i * 53 + 1) % len(open_rows)]
            s.goto("Заказы", f"$Q${row + 1}")
            _, t = timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName=(TODAY + datetime.timedelta(days=5)).strftime("%d.%m.%Y"))
            ms_q.append(t)
            s.goto("Заказы", f"$Z${row + 1}")
            _, t = timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName=f"комментарий {i}")
            ms_z.append(t)
        r["handler_q"] = stats(ms_q)
        r["handler_other_col"] = stats(ms_z)
        # 10. rows inserted above the data: the next operation on a far row finds it by its key (one MATCH) and fixes the hint
        s.goto("Заказы", "$A$2")
        s.ui(".uno:SelectRow")
        _, t_ins = timed(s.ui, ".uno:InsertRowsBefore")
        far = src_rows[-7] + 1
        x, t = timed(s.Rc, "ReceiptAddRow", far, "1", "", "", "", TODAY.strftime("%d.%m.%Y"), "Q-2", "")
        x2, t2 = timed(s.Rc, "ReceiptAddRow", far, "1", "", "", "", TODAY.strftime("%d.%m.%Y"), "Q-2", "")
        r["after_insert"] = dict(insert_ms=round(t_ins), first_op_ms=round(t), second_op_ms=round(t2), results=[x[:40], x2[:40]])
        print(f"[{n}] после вставки строки: {r['after_insert']}", flush=True)
        sc, t = timed(s.B, "ActionSelfCheck")
        r["selfcheck"] = dict(ms=round(t), result=sc.splitlines()[0], receipts=[x for x in sc.splitlines() if "приходы" in x][:1])
        print(f"[{n}] самопроверка: {r['selfcheck']}", flush=True)
        _, t = timed(s.doc.store)
        r["save_ms"] = round(t)
        t0 = time.time()
        P, inf = receipt_oracle.check(s.doc, s.jdir, {}, base=base)
        r["oracle"] = dict(problems=P[:5], n_problems=len(P), info={k: inf[k] for k in ("ops", "receipt_ops", "receipts", "positions")},
                           s=round(time.time() - t0))
        print(f"[{n}] оракул: {len(P)} расхождений {P[:3]}", flush=True)
        s.close()
        s = Session(p)
        rep = s.report()
        r["reopen"] = dict(open_s=round(s.open_s, 2), startup_ms=startup_ms(rep), state=rep.split(" | ")[0], refresh=note(rep, "статусы заказов"))
        s.close()
        s = Session(p, macros=0)
        s.doc.Sheets.getByName("Заказы").getCellByPosition(25, 1).setString("правка без WMS")
        s.doc.store()
        s.close()
        s = Session(p)
        rep = s.report()
        r["start_after_macros_off"] = dict(open_s=round(s.open_s, 2), startup_ms=startup_ms(rep), state=rep.split(" | ")[0],
                                           key_check=note(rep, "Выдачи: строк") + " / " + [x for x in rep.split("; ") if x.startswith("Заказы: строк")][0][:120]
                                           if "Заказы: строк" in rep else note(rep, "Выдачи: строк"))
        print(f"[{n}] повторное открытие {r['reopen']}; после сохранения без макросов {r['start_after_macros_off']}", flush=True)
    finally:
        s.close()
    return r


def main():
    sizes = [int(a) for a in sys.argv[1:]] or [100000, 250000]
    for n in sizes:
        RESULTS[n] = bench(n)
        with open(os.path.join(OUT, "bench_phase3.json"), "w", encoding="utf-8") as f:
            json.dump(RESULTS, f, ensure_ascii=False, indent=1)
    with open(os.path.join(OUT, "BENCH_phase3.md"), "w", encoding="utf-8") as f:
        f.write("| Показатель | " + " | ".join(f"{n} движений" for n in sizes) + " |\n|---|" + "---|" * len(sizes) + "\n")
        for k in RESULTS[sizes[0]]:
            f.write(f"| {k} | " + " | ".join(str(RESULTS[n].get(k, "")).replace("|", "¦") for n in sizes) + " |\n")
    print(json.dumps(RESULTS, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
