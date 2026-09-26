"""Benchmark of the special receipts (Core Phase 5) on large synthetic books.

    WMS_TEST_OUT=<dir> python3 tests/bench_phase5.py [N ...]          default: 100000 250000

For each N a TEST book gets N movements, a mixed profile: 80 % as in tests/bench_phase4.py (receipts on «Заказы», issues
on «Выдачи», returns on «Возврат») and 20 % special receipts on «Иной приход»: Офис (events of 2 lines), Производство
(events of 5 lines), Детали (a pool of articles: the first line of an article creates its EI, the next ones add to it),
Старый склад, Иной. «Наличие» (with «Тип источника»), _SPR, _ART, the counters and the protection of every row are written
exactly as WMS writes them (the history itself is not journaled: the journal starts with the measured operations). Then,
with WMS running: open and startup; new EIs of Офис, Производство, Детали; a repeated part → its existing EI; Старый склад;
Иной; corrections; storno (a part's addition under the balance rule; the last line of an EI under D-069); the lookup of
an article in the index (one MATCH, the entry checked against «Наличие»), the check of a part's row (the lookup + the
card), issues and returns of special EIs, the change handler, self-check, save, reopen, the start after a save without WMS
(full key check) and the independent oracle (trusted snapshot + journal replay).
Times are measured from Python around one UNO call (the call itself costs ~1–3 ms), so they are upper bounds.
Writes bench_phase5.json and BENCH_phase5.md into the output folder.
"""
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Session, new_wms, synthetic_registry  # noqa: E402
from build_ods import PWD, ISSUE_LOCKS_OPEN, RETURN_LOCKS_OPEN, SPECIAL_LOCKS_OPEN  # noqa: E402
import bench_phase3 as b3  # noqa: E402
import bench_phase4 as b4  # noqa: E402
import receipt_oracle  # noqa: E402
import special_oracle  # noqa: E402

SPECIAL_POSTED = "1111111111111111110"
TSER = b3.TSER
RESULTS = {}
CODES = {"Офис": "OFF", "Производство": "PROD", "Детали": "DET", "Старый склад": "OLD", "Иной": "OTH"}
STYPE = {"Офис": "Офис", "Производство": "Производство", "Детали": "Детали", "Старый склад": "Старый склад", "Иной": "Иной приход"}
GROUP = {"OFF": 2, "PROD": 5, "DET": 1, "OLD": 1, "OTH": 1}
KIND = ["Офис"] * 4 + ["Производство"] * 5 + ["Детали"] * 7 + ["Старый склад"] * 2 + ["Иной"] * 2       # 20 %, 25 %, 35 %, 10 %, 10 %


def spr_key(code, artkey, name, q, d, doc):
    """the antidubl key of a special line, as WmsSpecial.DupKeyOf builds it (whole quantities here)"""
    p = "a" + artkey if artkey else "n" + receipt_oracle.norm_text(name)
    return f"{code.lower()}|{p}|{int(q)}|{int(d)}|{receipt_oracle.norm_text(doc)}"


def generate(n):
    n_base = int(n * 0.8)
    g = b4.generate(n_base)
    n_sp = n - n_base
    ei = g["next_ei"]
    rows, spr, art, stock = [], [], [], []
    ev_num = {c: 0 for c in GROUP}
    ev_fill = {c: 0 for c in GROUP}
    n_det = sum(1 for j in range(n_sp) if KIND[j % 20] == "Детали")
    pool = max(1, n_det // 4)
    part_ei, bal = {}, {}
    det_i = 0
    for j in range(n_sp):
        typ = KIND[j % 20]
        code = CODES[typ]
        if ev_fill[code] == 0:
            ev_num[code] += 1
        ev_fill[code] = (ev_fill[code] + 1) % GROUP[code]
        ev = f"{code}-{ev_num[code]:08d}"
        nL = j + 1
        d = float(TSER - 300 + j * 280 // n_sp)
        place, cat = f"X-{j % 30:02d}", b3.CATS[j % 6]
        who = "Цех 2" if code == "PROD" else ("Офис" if code == "OFF" else "")
        mark = f"инв. {j}" if code == "OLD" else ""
        artkey, art_in = "", ""
        if code == "DET":
            a = (det_i * 7919) % pool
            det_i += 1
            art_in = artkey = f"DT-{a:05d}"
            name = f"Деталь {art_in}"
        else:
            name = f"{typ} позиция {j}"
        if code == "DET" and artkey in part_ei:
            e = part_ei[artkey]
            q = 5.0
            before = bal[e]
            bal[e] = before + q
            place = stock[e - g["next_ei"]][5]
            cat = stock[e - g["next_ei"]][6]
            name = stock[e - g["next_ei"]][1]
            mode, status, ctl = "ADD", "Проведено: пополнение ЕИ", f"Пополнение ЕИ-{e:08d}: остаток {b3.num(before)} → {b3.num(bal[e])} шт · {ev}"
            row_before = before
        else:
            e = ei
            ei += 1
            q = float(10 + j % 20)
            bal[e] = q
            mode = "NEW"
            status = "Проведено: требует разбора" if code == "OTH" else "Проведено: новый ЕИ"
            ctl = f"Новый ЕИ ЕИ-{e:08d} · {ev}"
            row_before = ""
            stock.append([f"ЕИ-{e:08d}", name, art_in, "шт", q, place, cat, "Требует разбора" if code == "OTH" else "Активен",
                          f"{STYPE[typ]} {ev}", STYPE[typ]])
            if code == "DET":
                part_ei[artkey] = e
                art.append((artkey, f"ЕИ-{e:08d}", art_in))
        key = spr_key(code, artkey, name, q, d, "")
        rows.append([float(nL), typ, ev, name, art_in, q, "шт", d, place, cat, who, "", mark, f"ЕИ-{e:08d}", row_before, bal[e], status, ctl, ""])
        spr.append([float(nL), ev, code, f"ЕИ-{e:08d}", mode, q, "LIVE", float(nL), d, b3.dup_hash(key), key, artkey, "", ""])
    for e, b in bal.items():
        stock[e - g["next_ei"]][4] = b
    g.update(sp_rows=rows, spr=spr, art=art, sp_stock=stock, sp_first_ei=g["next_ei"], next_ei=ei, n_sp=n_sp, ev_num=ev_num,
             part_ei=part_ei, pool=pool)
    return g


def make_book(n):
    t0 = time.time()
    g = generate(n)
    p = new_wms(f"bench5_{n}")
    s = Session(p)                               # the first start registers the book
    s.close(save=True)
    s = Session(p, macros=0)
    doc = s.doc
    names = ("Заказы", "Выдачи", "Наличие", "_RCV", "_ORD", "Возврат", "_RET", "_ISS", "Иной приход", "_SPR", "_ART")
    for nm in names:
        doc.Sheets.getByName(nm).unprotect(PWD)
    o, iss, st, rc, od, rv, rt, isx, sp, spr, art = (doc.Sheets.getByName(nm) for nm in names)
    b3.write_rows(st, g["reg"], 9)
    # «Тип источника» of the ordinary receipts (as the receipt of Core Phase 5 writes it), then the special EIs
    st.getCellRangeByPosition(9, 201, 9, g["sp_first_ei"] - 1).setDataArray(tuple(("Поставщик",) for _ in range(201, g["sp_first_ei"])))
    b3.write_rows(st, g["sp_stock"], 10, first=g["sp_first_ei"])
    b3.write_rows(rc, g["rcv"], 9, first=201)
    b3.write_rows(od, [list(x) for x in g["ordp"]], 9)
    od.getCellByPosition(11, 0).setValue(g["next_ol"])
    b3.write_rows(o, g["orders"], 28, chunk=5000)
    b3.write_rows(iss, g["iss"], 18)
    b3.write_rows(rv, g["rets"], 15)
    b3.write_rows(rt, g["rt"], 6)
    b3.write_rows(isx, g["iss_rows"], 4)
    b3.write_rows(sp, g["sp_rows"], 19, chunk=5000)
    b3.write_rows(spr, g["spr"], 14)
    if g["art"]:
        b3.write_rows(art, g["art"], 3)
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
    for c, bit in enumerate(b4.RETURN_POSTED):
        if bit != RETURN_LOCKS_OPEN[c]:
            rv.getCellRangeByPosition(c, 1, c, nr).CellProtection = b3.protection(bit == "1")
    nsp = g["n_sp"]
    for c, bit in enumerate(SPECIAL_POSTED):
        if bit != SPECIAL_LOCKS_OPEN[c]:
            sp.getCellRangeByPosition(c, 1, c, nsp).CellProtection = b3.protection(bit == "1")
    for nm in names:
        doc.Sheets.getByName(nm).protect(PWD)
    s.set_sys("NEXT_EI", g["next_ei"])
    s.set_sys("NEXT_NO", ni + 1)
    s.set_sys("NEXT_RET", nr + 1)
    s.set_sys("NEXT_SPL", nsp + 1)
    for code, key in (("OFF", "NEXT_OFF"), ("PROD", "NEXT_PROD"), ("DET", "NEXT_DET"), ("OLD", "NEXT_OLD"), ("OTH", "NEXT_OTH")):
        s.set_sys(key, g["ev_num"][code] + 1)
    s.set_sys("SAVE_STAMP", doc.getDocumentProperties().EditingCycles + 1)
    doc.store()
    s.close()
    info = dict(movements=n, receipts=g["receipts"], issues=ni, returns=nr, special_lines=nsp, special_eis=len(g["sp_stock"]),
                parts=len(g["part_ei"]), order_rows=no, eis=g["next_ei"] - 1, book_mb=round(os.path.getsize(p) / 1e6, 1),
                build_s=round(time.time() - t0))
    return p, g, info


def put_rows(sh, top, rows):
    """input rows B..M of «Иной приход» through the API (the posting is measured, not the typing)"""
    sh.getCellRangeByPosition(1, top, 12, top + len(rows) - 1).setDataArray(tuple(tuple(r) for r in rows))


def post_many(s, top, k, key):
    res, ms = [], []
    for i in range(k):
        x, t = b3.timed(s.X, "SpecialPostRow", top + i, "", False)
        res.append(x)
        ms.append(t)
    return b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:70])


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
        sh = s.doc.Sheets.getByName("Иной приход")
        top = g["n_sp"] + 1
        d = float(TSER - 1)
        blank = ["", "", "", "", "", "", "", "", "", "", ""]

        def row(typ, name, q, place, art="", who="", mark=""):
            return [typ, "", name, art, q, "шт", d, place, "Разное", who, "", mark]
        # 1. new EIs of Офис, Производство, Детали (new articles), a part repeated (existing EI), Старый склад, Иной
        plan = [("office_new", [row("Офис", f"Бенч офис {i}", "2", "B-1") for i in range(50)]),
                ("production_new", [row("Производство", f"Бенч изделие {i}", "3", "B-2", who="Цех 2") for i in range(50)]),
                ("details_new", [row("Детали", f"Бенч деталь {i}", "10", "B-3", art=f"BN-{i:04d}") for i in range(50)]),
                ("details_repeat", [row("Детали", "", "5", "", art=f"DT-{(i * 104729) % g['pool']:05d}") for i in range(50)]),
                ("old_new", [row("Старый склад", f"Бенч старый {i}", "1", "B-4", mark=f"инв. Б{i}") for i in range(30)]),
                ("other_new", [row("Иной", f"Бенч неизвестный {i}", "1", "B-5") for i in range(30)])]
        rowsof = {}
        for key, rws in plan:
            put_rows(sh, top, rws)
            r[key] = post_many(s, top, len(rws), key)
            rowsof[key] = list(range(top, top + len(rws)))
            top += len(rws)
            print(f"[{n}] {key}: {r[key]}", flush=True)
        # 2. the lookup of an article in the index: one MATCH of the Calc engine (+ a second one for a repeated key); the check
        #    of a part's row = lookup + the entry against «Наличие» + the card
        ms_a = []
        for i in range(30):
            _, t = b3.timed(s.Od, "FindArtRow", f"DT-{(i * 3571) % g['pool']:05d}", 0)
            ms_a.append(t)
        r["article_lookup"] = b3.stats(ms_a)
        put_rows(sh, top, [row("Детали", "", "1", "", art=f"DT-{(i * 7) % g['pool']:05d}") for i in range(20)])
        ms_c, res_c = [], []
        for i in range(20):
            x, t = b3.timed(s.X, "SpecialCheckRow", top + i)
            ms_c.append(t)
            res_c.append(x)
        r["part_row_check"] = b3.stats(ms_c) | dict(first=res_c[0][:90])
        top += 20
        print(f"[{n}] поиск артикула {r['article_lookup']}; проверка строки детали {r['part_row_check']}", flush=True)
        # 3. corrections: lines of the generated history spread over the sheet (quantity +1) and the new lines (Офис: name, place)
        nsp = g["n_sp"]
        fx = list(dict.fromkeys(1 + (i * 7919 + 13) % nsp for i in range(50)))
        res, ms = [], []
        for rr in fx:
            v = g["sp_rows"][rr - 1]
            args = (b3.num(v[5] + 1), v[7], v[8], v[9], "", v[10], v[12], v[3], v[4], "шт")
            x, t = b3.timed(s.X, "SpecialFixRow", rr, *args)
            res.append(x)
            ms.append(t)
        r["fix"] = b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:80])
        print(f"[{n}] исправление: {r['fix']}", flush=True)
        # 4. storno: additions of parts spread over the sheet (the balance rule) and new EIs of Офис (the last line: D-069)
        adds = [i + 1 for i, v in enumerate(g["sp_rows"]) if v[16] == "Проведено: пополнение ЕИ" and (i + 1) not in fx]
        dl = list(dict.fromkeys(adds[(i * 104729 + 5) % len(adds)] for i in range(30)))
        res, ms = [], []
        for rr in dl:
            x, t = b3.timed(s.X, "SpecialDeleteRow", rr)
            res.append(x)
            ms.append(t)
        r["storno_addition"] = b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:80])
        res, ms = [], []
        for rr in rowsof["office_new"][:30]:
            x, t = b3.timed(s.X, "SpecialDeleteRow", rr)
            res.append(x)
            ms.append(t)
        r["storno_last_line"] = b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:80])
        print(f"[{n}] сторно пополнения {r['storno_addition']}; сторно последней строки ЕИ {r['storno_last_line']}", flush=True)
        # 5. issues and returns of special EIs (Производство, Детали, Иной: new and generated)
        iss = s.doc.Sheets.getByName("Выдачи")
        top_i = len(g["iss"]) + 1
        eis = [s.spc(rr)[13] for rr in rowsof["production_new"][:10] + rowsof["details_new"][:10] + rowsof["other_new"][:10]]
        eis += [g["sp_rows"][(i * 7919) % nsp][13] for i in range(20)]
        vals = eis
        last_i = top_i + len(eis) - 1
        iss.getCellRangeByPosition(11, top_i, 11, last_i).setDataArray(tuple((e,) for e in eis))          # L the EI
        iss.getCellRangeByPosition(4, top_i, 4, last_i).setDataArray(tuple(("0,5",) for _ in eis))       # E the quantity
        iss.getCellRangeByPosition(7, top_i, 8, last_i).setDataArray(tuple((float(TSER), "Иванов Иван Андреевич") for _ in eis))
        res, ms = [], []
        for i in range(len(vals)):
            x, t = b3.timed(s.I, "IssuePostRow", top_i + i)
            res.append(x)
            ms.append(t)
        r["issue_special_ei"] = b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:60])
        rv = s.doc.Sheets.getByName("Возврат")
        top_r = g["n_ret"] + 1
        nos = [int(iss.getCellByPosition(0, top_i + i).getValue()) for i in range(len(vals))]
        rv.getCellRangeByPosition(1, top_r, 1, top_r + len(nos) - 1).setDataArray(tuple((float(k),) for k in nos))
        rv.getCellRangeByPosition(5, top_r, 5, top_r + len(nos) - 1).setDataArray(tuple(("0,25",) for _ in nos))
        rv.getCellRangeByPosition(7, top_r, 7, top_r + len(nos) - 1).setDataArray(tuple((float(TSER),) for _ in nos))
        res, ms = [], []
        for i in range(len(nos)):
            x, t = b3.timed(s.Rt, "ReturnPostRow", top_r + i)
            res.append(x)
            ms.append(t)
        r["return_special_ei"] = b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:60])
        print(f"[{n}] выдача {r['issue_special_ei']}; возврат {r['return_special_ei']}", flush=True)
        # 6. the change handler: the type and the article of a part typed into a new row (preview with the lookup), a comment S
        ms_t, ms_a2, ms_s = [], [], []
        for i in range(10):
            rr = top + i
            s.goto("Иной приход", f"$B${rr + 1}")
            _, t = b3.timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName="Детали")
            ms_t.append(t)
            s.goto("Иной приход", f"$E${rr + 1}")
            _, t = b3.timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName=f"DT-{(i * 131) % g['pool']:05d}")
            ms_a2.append(t)
            s.goto("Иной приход", f"$S${rr + 1}")
            _, t = b3.timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName=f"комментарий {i}")
            ms_s.append(t)
        r["handler_type"] = b3.stats(ms_t)
        r["handler_article"] = b3.stats(ms_a2) | dict(preview=s.spc(top)[17][:90])
        r["handler_comment"] = b3.stats(ms_s)
        print(f"[{n}] обработчик: тип {r['handler_type']}; артикул {r['handler_article']}; комментарий {r['handler_comment']}", flush=True)
        ops = [k for k in ("office_new", "production_new", "details_new", "details_repeat", "old_new", "other_new", "article_lookup", "part_row_check",
                           "fix", "storno_addition", "storno_last_line", "issue_special_ei", "return_special_ei", "handler_type", "handler_article",
                           "handler_comment")]
        r["max_operation_ms"] = max(r[k]["max"] for k in ops)
        sc, t = b3.timed(s.B, "ActionSelfCheck")
        r["selfcheck"] = dict(ms=round(t), result=sc.splitlines()[0], special=[x[:260] for x in sc.splitlines() if "специальные приходы" in x][:1],
                              keys=[x[:400] for x in sc.splitlines() if x.startswith(("OK ключи", "WARN ключи", "FAIL ключи"))][:1])
        print(f"[{n}] самопроверка: {r['selfcheck']}", flush=True)
        _, t = b3.timed(s.doc.store)
        r["save_ms"] = round(t)
        t0 = time.time()
        P, inf, hist = special_oracle.check(s.doc, s.jdir, {}, base=base)
        r["oracle"] = dict(problems=P[:5], n_problems=len(P), info={k: inf.get(k) for k in ("ops", "special_ops", "special_lines", "special_eis", "parts",
                                                                                                 "index", "return_ops")}, s=round(time.time() - t0))
        print(f"[{n}] оракул: {len(P)} расхождений {P[:3]}", flush=True)
        s.close()
        s = Session(p)
        rep = s.report()
        r["reopen"] = dict(open_s=round(s.open_s, 2), startup_ms=b3.startup_ms(rep), state=rep.split(" | ")[0])
        s.close()
        s = Session(p, macros=0)
        s.doc.Sheets.getByName("Иной приход").getCellByPosition(18, 1).setString("правка без WMS")
        s.doc.store()
        s.close()
        s = Session(p)
        rep = s.report()
        r["start_after_macros_off"] = dict(open_s=round(s.open_s, 2), startup_ms=b3.startup_ms(rep), state=rep.split(" | ")[0],
                                           key_check=[x[:160] for x in rep.split("; ") if "Иной приход: строк" in x][:1])
        print(f"[{n}] повторное открытие {r['reopen']}; после сохранения без макросов {r['start_after_macros_off']}", flush=True)
    finally:
        s.close()
    return r


def main():
    sizes = [int(a) for a in sys.argv[1:]] or [100000, 250000]
    for n in sizes:
        RESULTS[n] = bench(n)
        with open(os.path.join(OUT, "bench_phase5.json"), "w", encoding="utf-8") as f:
            json.dump(RESULTS, f, ensure_ascii=False, indent=1)
    with open(os.path.join(OUT, "BENCH_phase5.md"), "w", encoding="utf-8") as f:
        f.write("| Показатель | " + " | ".join(f"{n} движений" for n in sizes) + " |\n|---|" + "---|" * len(sizes) + "\n")
        for k in RESULTS[sizes[0]]:
            f.write(f"| {k} | " + " | ".join(str(RESULTS[n].get(k, "")).replace("|", "¦") for n in sizes) + " |\n")
    print(json.dumps(RESULTS, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
