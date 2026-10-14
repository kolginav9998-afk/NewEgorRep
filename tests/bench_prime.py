"""Benchmark of the pages of PRIME 0.8 on large books: «Дашборд», «Поиск» (a word, an EI card), «Карта склада».

    WMS_TEST_OUT=<dir> python3 tests/bench_prime.py [N ...]          default: 100000 250000

The book of tests/bench_phase5.py (N movements: receipts, issues, returns, special receipts — every row and service table
written exactly as WMS writes them) gets a topology of 11 rows × 20 sections × 4 levels (880 places) and a part of its EIs
moved there; then, with WMS running: the start, the dashboard (counted, then shown again without a new operation), the
search of a word and of an EI card, the map (built, then switched between its layers). The pages never run inside an
operation: the time of a posting is the one of tests/bench_final_core.py.
Writes bench_prime.json and BENCH_prime.md into the output folder.
"""
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Session  # noqa: E402
import bench_phase3 as b3  # noqa: E402
import bench_phase5 as b5  # noqa: E402

RESULTS = {}


def timed(fn, *a, **kw):
    t0 = time.perf_counter()
    r = fn(*a, **kw)
    return r, round((time.perf_counter() - t0) * 1000)


def topology(s, rows=11, secs=20, levels=4):
    t = s.doc.Sheets.getByName("Топология")
    t.getCellByPosition(0, 0).setString("")
    r = 1
    grid = [[""] * (secs + 1) for _ in range(rows * (levels + 2) + 2)]
    for R in range(1, rows + 1):
        grid[r][0] = str(R)
        for lv in range(levels, 0, -1):
            for sec in range(1, secs + 1):
                grid[r][sec] = f"{R}-{sec}-{lv}"
            r += 1
        r += 2
    t.getCellRangeByPosition(0, 0, secs, len(grid) - 1).setDataArray(tuple(tuple(x) for x in grid))
    return rows * secs * levels


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
        places = topology(s)
        # a part of the EIs on the places of the map (the places of «Наличие» written as the storekeeper would type them)
        st = s.doc.Sheets.getByName("Наличие")
        last = g["next_ei"] - 1
        st.unprotect("wms")
        col = st.getCellRangeByPosition(5, 1, 5, last)
        vals = [[f"{(i % 11) + 1}-{(i // 11) % 20 + 1}-{(i // 220) % 4 + 1}" if i % 5 else "склад"] for i in range(last)]
        col.setDataArray(tuple(tuple(v) for v in vals))
        st.protect("wms")
        s.doc.store()
        res, r["dash_ms"] = timed(s.B, "DashRefresh", module="WmsDash")
        r["dash"] = res
        _, r["dash_again_ms"] = timed(s.B, "OnDashFocus", module="WmsDash")
        res, r["find_word_ms"] = timed(s.B, "FindRun", "Товар 1", module="WmsFind")
        r["find_word"] = res
        res, r["find_ei_ms"] = timed(s.B, "FindRun", "ЕИ-00000001", module="WmsFind")
        r["card"] = res
        res, r["map_ms"] = timed(s.B, "MapBuild", "OCC", module="WmsMap")
        r["map"] = res
        layers = {}
        for layer in ("ISS", "STALE", "ATT", "FREE"):
            _, layers[layer] = timed(s.B, "MapBuild", layer, module="WmsMap")
        r["map_layers_ms"] = layers
        r["places"] = places
        print(f"[{n}] запуск {r['startup_ms']} мс; дашборд {r['dash_ms']} мс (повторный показ {r['dash_again_ms']} мс); поиск слова "
              f"{r['find_word_ms']} мс; карточка {r['find_ei_ms']} мс; карта {r['map_ms']} мс, слои {layers}", flush=True)
        ok = all(x.startswith("OK") for x in (r["dash"], r["find_word"], r["card"], r["map"]))
        r["ok"] = ok
    finally:
        s.close()
    RESULTS[str(n)] = r
    return r


def main():
    sizes = [int(x) for x in sys.argv[1:]] or [100000, 250000]
    for n in sizes:
        bench(n)
    with open(os.path.join(OUT, "bench_prime.json"), "w", encoding="utf-8") as f:
        json.dump(RESULTS, f, ensure_ascii=False, indent=1)
    lines = ["# PRIME 0.8 — страницы на больших книгах", "", "| строк | запуск, мс | дашборд, мс | повторный показ, мс | поиск слова, мс | карточка ЕИ, мс | карта, мс | слой, мс |",
             "|---|---|---|---|---|---|---|---|"]
    for k, r in RESULTS.items():
        lines.append(f"| {k} | {r['startup_ms']} | {r['dash_ms']} | {r['dash_again_ms']} | {r['find_word_ms']} | {r['find_ei_ms']} | {r['map_ms']} | "
                     f"{max(r['map_layers_ms'].values())} |")
    open(os.path.join(OUT, "BENCH_prime.md"), "w", encoding="utf-8").write("\n".join(lines) + "\n")
    print("\n".join(lines))
    return 0 if all(r.get("ok") for r in RESULTS.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
