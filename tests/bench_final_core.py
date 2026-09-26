"""Benchmark of the Final Core (moves, write-offs, inventory corrections, the snapshot for the tools) on large books.

    WMS_TEST_OUT=<dir> python3 tests/bench_final_core.py [N ...]          default: 100000 250000

The book of tests/bench_phase5.py (N movements: receipts, issues, returns, special receipts; every row and service table
written exactly as WMS writes them) gets, with WMS running: moves, write-offs and inventory corrections of EIs spread over
the registry, their corrections and storno, the change handler of «Корректировки», the snapshot for the tools (time,
size), the self-check, the save, the reopen and the independent oracle (trusted snapshot + journal replay).
Times are measured from Python around one UNO call (the call itself costs ~1–3 ms), so they are upper bounds.
Writes bench_final_core.json and BENCH_final_core.md into the output folder.
"""
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Session, synthetic_registry  # noqa: E402
import bench_phase3 as b3  # noqa: E402
import bench_phase5 as b5  # noqa: E402
import receipt_oracle  # noqa: E402
import adjust_oracle  # noqa: E402

RESULTS = {}


def put(sh, top, rows):
    """input rows of «Корректировки» through the API: B C, G..L (the posting is measured, not the typing)"""
    for i, (k, e, g, h, j, dt, why) in enumerate(rows):
        r = top + i
        sh.getCellRangeByPosition(1, r, 2, r).setDataArray(((k, e),))
        sh.getCellRangeByPosition(6, r, 7, r).setDataArray(((g, h),))
        sh.getCellRangeByPosition(9, r, 11, r).setDataArray(((j, dt, why),))


def run(s, fn, rows, *extra):
    res, ms = [], []
    for r in rows:
        x, t = b3.timed(s.A, fn, r, *extra)
        res.append(x)
        ms.append(t)
    return b3.stats(ms) | dict(ok=sum(1 for x in res if x.startswith("OK")), first=res[0][:80])


def bench(n):
    r = {}
    p, g, info = b5.make_book(n)
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
        stock = s.doc.Sheets.getByName("Наличие")
        last = g["next_ei"] - 1
        bal = [x[0] for x in stock.getCellRangeByPosition(4, 1, 4, last).getDataArray()]
        cand = [i + 1 for i, b in enumerate(bal) if isinstance(b, float) and b >= 3]
        pick = [cand[(k * 7919 + 17) % len(cand)] for k in range(400)]
        pick = list(dict.fromkeys(pick))
        sh = s.doc.Sheets.getByName("Корректировки")
        d = float(b3.TSER)
        top = 1
        mv, wo, inv = pick[:60], pick[60:120], pick[120:180]
        put(sh, top, [("Перемещение", f"ЕИ-{e:08d}", "", "", f"BM-{k % 30}", d, "бенч") for k, e in enumerate(mv)])
        r["move"] = run(s, "AdjustPostRow", range(top, top + len(mv)))
        rows_mv = list(range(top, top + len(mv)))
        top += len(mv)
        put(sh, top, [("Списание", f"ЕИ-{e:08d}", "1", "", "", d, "бенч") for e in wo])
        r["write_off"] = run(s, "AdjustPostRow", range(top, top + len(wo)))
        rows_wo = list(range(top, top + len(wo)))
        top += len(wo)
        put(sh, top, [("Инвентаризация", f"ЕИ-{e:08d}", "", str(int(bal[e - 1]) + 1), "", d, "бенч") for e in inv])
        r["inventory"] = run(s, "AdjustPostRow", range(top, top + len(inv)))
        rows_inv = list(range(top, top + len(inv)))
        top += len(inv)
        print(f"[{n}] перемещение {r['move']}; списание {r['write_off']}; инвентаризация {r['inventory']}", flush=True)
        r["fix_move"] = run(s, "AdjustFixRow", rows_mv[:20], "BM-X", d, "бенч (исправлено)")
        r["fix_write_off"] = run(s, "AdjustFixRow", rows_wo[:20], "2", d, "бенч (исправлено)")
        r["fix_inventory"] = run(s, "AdjustFixRow", rows_inv[:20], "1", d, "бенч (исправлено)")
        r["storno_move"] = run(s, "AdjustDeleteRow", rows_mv[20:40])
        r["storno_write_off"] = run(s, "AdjustDeleteRow", rows_wo[20:40])
        r["storno_inventory"] = run(s, "AdjustDeleteRow", rows_inv[20:40])
        print(f"[{n}] исправление {r['fix_move']['avg']}/{r['fix_write_off']['avg']}/{r['fix_inventory']['avg']} мс; сторно "
              f"{r['storno_move']['avg']}/{r['storno_write_off']['avg']}/{r['storno_inventory']['avg']} мс", flush=True)
        ms_k, ms_e = [], []
        for i in range(10):
            rr = top + i
            s.goto("Корректировки", f"$B${rr + 1}")
            _, t = b3.timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName="Списание")
            ms_k.append(t)
            s.goto("Корректировки", f"$C${rr + 1}")
            _, t = b3.timed(s.o.dispatch, s.doc, ".uno:EnterString", StringName=str(pick[200 + i]))
            ms_e.append(t)
        r["handler_kind"] = b3.stats(ms_k)
        r["handler_ei"] = b3.stats(ms_e) | dict(preview=str(s.adj(top)[16])[:90])
        ops = ("move", "write_off", "inventory", "fix_move", "fix_write_off", "fix_inventory", "storno_move", "storno_write_off", "storno_inventory",
               "handler_kind", "handler_ei")
        r["max_operation_ms"] = max(r[k]["max"] for k in ops)
        print(f"[{n}] обработчик {r['handler_kind']} / {r['handler_ei']}; максимум операции {r['max_operation_ms']} мс", flush=True)
        x, t = b3.timed(s.EX, "ExportSnapshot")
        snap = s.EX("TestLastSnapshot")
        size = sum(os.path.getsize(os.path.join(dp, f)) for dp, _, fs in os.walk(snap) for f in fs)
        r["export"] = dict(ms=round(t), mb=round(size / 1e6, 1), result=x[:60], tables=x.split(" — ")[-1][:200] if " — " in x else "")
        print(f"[{n}] экспорт: {r['export']}", flush=True)
        sc, t = b3.timed(s.B, "ActionSelfCheck")
        r["selfcheck"] = dict(ms=round(t), result=sc.splitlines()[0], adjust=[x[:260] for x in sc.splitlines() if "корректировки" in x][:1])
        print(f"[{n}] самопроверка: {r['selfcheck']}", flush=True)
        _, t = b3.timed(s.doc.store)
        r["save_ms"] = round(t)
        t0 = time.time()
        P, inf, _ = adjust_oracle.check(s.doc, s.jdir, {}, base=base)
        r["oracle"] = dict(problems=P[:5], n_problems=len(P), info={k: inf.get(k) for k in ("ops", "adjust_ops", "adjustments", "adjust_live")},
                           s=round(time.time() - t0))
        print(f"[{n}] оракул: {len(P)} расхождений {P[:3]}", flush=True)
        s.close()
        s = Session(p)
        rep = s.report()
        r["reopen"] = dict(open_s=round(s.open_s, 2), startup_ms=b3.startup_ms(rep), state=rep.split(" | ")[0])
        print(f"[{n}] повторное открытие {r['reopen']}", flush=True)
    finally:
        s.close()
    return r


def main():
    sizes = [int(a) for a in sys.argv[1:]] or [100000, 250000]
    for n in sizes:
        RESULTS[n] = bench(n)
        with open(os.path.join(OUT, "bench_final_core.json"), "w", encoding="utf-8") as f:
            json.dump(RESULTS, f, ensure_ascii=False, indent=1)
    with open(os.path.join(OUT, "BENCH_final_core.md"), "w", encoding="utf-8") as f:
        f.write("| Показатель | " + " | ".join(f"{n} движений" for n in sizes) + " |\n|---|" + "---|" * len(sizes) + "\n")
        for k in RESULTS[sizes[0]]:
            f.write(f"| {k} | " + " | ".join(str(RESULTS[n].get(k, "")).replace("|", "¦") for n in sizes) + " |\n")
    print(json.dumps(RESULTS, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
