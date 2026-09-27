"""The fixed snapshot of the toolbox tests (FINAL WMS MARATHON §8: «tool tests use fixed snapshot fixtures»).

    python3 tests/make_fixture_snapshot.py [OUT_DIR]        default: tests/fixtures/

Run once; the result tests/fixtures/WMS_SNAPSHOT_FIXTURE/ is committed and every toolbox test reads it (copied into a
temporary WMS_Export folder), so the expected values of the tests never change with the core. A test book (synthetic EIs
1–200, five recipients — no real data) gets a deterministic series through WMS itself: ordinary orders and receipts (a
partial delivery, an overdue position), special receipts of every source (a part and its refill, an «Иной» to be
identified), issues to three recipients in July–September 2026, returns, a move, a write-off, inventory corrections (a
surplus and a shortage), a cancelled issue; then «Экспорт для инструментов». The folder is copied as WMS wrote it; only
the absolute path of the book in manifest.csv is replaced by «WMS.ods» (manifest.csv itself is not hashed) and the
registered path in service/_SYS.csv by «file:///WMS/WMS.ods» (its size and SHA-256 in the manifest recomputed), so no path of
the machine that made it is committed.
"""
import hashlib
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from harness import Session, new_wms  # noqa: E402

NAME = "WMS_SNAPSHOT_FIXTURE"


def ok(x, what):
    if not str(x).startswith("OK"):
        raise RuntimeError(f"{what}: {x}")
    return x


def series(s):
    s.U("TestUiAuto", 1)
    # ordinary orders: two received in full, one partially (a second delivery later), one overdue and never received
    s.order_input(1, A="З-101", B="Болт М12х40", E="DIN933-M12", H="200", I="шт", L="Метиз-Опт", N="10.07.2026", U="A-02-1", C="УПД-501",
                  O="10.07.2026", S="Крепёж", F="200")
    ok(s.click_ord("BtnRcvPost", 1), "приход 1")
    s.order_input(2, A="З-102", B="Кабель КГ 3х2,5", E="KG-3x2.5", H="150", I="м", L="Кабель-Сервис", N="15.08.2026", U="B-01", C="УПД-640",
                  O="15.08.2026", S="Кабель", F="100")
    ok(s.click_ord("BtnRcvPost", 2), "приход 2 (частично)")
    s.order_input(3, A="З-103", B="Перчатки нитриловые", E="GL-NIT-M", H="50", I="пар", L="СИЗ-Центр", N="02.09.2026", U="C-05", C="УПД-702",
                  O="02.09.2026", S="СИЗ", F="50")
    ok(s.click_ord("BtnRcvPost", 3), "приход 3")
    s.order_input(4, A="З-104", B="Герметик силиконовый", E="SIL-300", H="24", I="шт", L="Химпром", P="01.08.2026", Q="20.08.2026",
                  S="Химия", U="D-01")
    # special receipts of every source
    rows = [dict(B="Офис", D="Стол офисный", F="1", G="шт", H="05.07.2026", I="O-1", J="Мебель", K="Офис"),
            dict(B="Производство", D="Кронштейн сварной", F="12", G="шт", H="20.07.2026", I="P-3", J="Металл", K="Цех 2"),
            dict(B="Детали", D="Шестерня Z-40", E="GR-40", F="6", G="шт", H="03.08.2026", I="D-9", J="Детали"),
            dict(B="Старый склад", D="Дрель ударная", F="1", G="шт", H="12.08.2026", I="C-3", J="Инструмент", M="инв. 17"),
            dict(B="Иной", D="Коробка без маркировки", F="3", G="шт", H="01.09.2026", I="Z-1")]
    for i, cols in enumerate(rows, start=1):
        s.special_input(i, **cols)
        ok(s.click_spc("BtnSpcPost", i), f"иной приход {i}")
    s.special_input(6, B="Детали", E="gr-40", F="4", G="шт", H="10.09.2026")
    ok(s.click_spc("BtnSpcPost", 6), "пополнение детали")
    # issues over three months, three recipients (EIs: 1–200 synthetic, 201… the receipts above)
    iss = [("201", "20", "15.07.2026", "ива"), ("201", "30", "20.08.2026", "пет"), ("202", "25", "21.08.2026", "ива"),
           ("205", "2", "25.07.2026", "пр"), ("207", "1", "12.08.2026", "егр"), ("1", "5", "01.09.2026", "ива"),
           ("2", "10,5", "14.09.2026", "пет"), ("203", "4", "18.09.2026", "ива"), ("201", "1", "19.09.2026", "ива")]
    for i, (e, q, d, w) in enumerate(iss, start=1):
        s.issue_input(i, ei=e, qty=q, date=d, who=w)
        ok(s.post(i), f"выдача {i}")
    s.click("BtnDelete", 9)
    ok(s.U("TestUiLastMessage"), "сторно выдачи 9")
    # returns to issues 1 and 3
    s.return_input(1, issue="1", qty="5", date="30.07.2026")
    ok(s.click_ret("BtnRetPost", 1), "возврат 1")
    s.return_input(2, issue="3", qty="10", date="28.08.2026")
    ok(s.click_ret("BtnRetPost", 2), "возврат 2")
    # corrections: a move, a write-off, an inventory surplus and shortage
    adj = [dict(B="Перемещение", C="202", J="B-02", K="16.09.2026", L="освобождение стеллажа B-01"),
           dict(B="Списание", C="206", G="2", K="17.09.2026", L="брак"),
           dict(B="Инвентаризация", C="3", H="40", K="20.09.2026", L="инвентаризация сентябрь"),
           dict(B="Инвентаризация", C="4", H="10", K="20.09.2026", L="инвентаризация сентябрь")]
    for i, cols in enumerate(adj, start=1):
        s.adjust_input(i, **cols)
        ok(s.click_adj("BtnAdjPost", i), f"корректировка {i}")
    s.doc.store()


def main():
    out = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "fixtures"))
    s = Session(new_wms("fixture"))
    try:
        series(s)
        ok(s.EX("ExportSnapshot"), "экспорт")
        snap = s.EX("TestLastSnapshot")
    finally:
        s.close()
    dst = os.path.join(out, NAME)
    shutil.rmtree(dst, ignore_errors=True)
    shutil.copytree(snap, dst)
    sysf = os.path.join(dst, "service", "_SYS.csv")
    raw = open(sysf, "rb").read().decode("utf-8")
    raw = "\n".join("REGISTERED_URL;file:///WMS/WMS.ods" if ln.startswith("REGISTERED_URL;") else ln for ln in raw.split("\n"))
    open(sysf, "wb").write(raw.encode("utf-8"))
    h = hashlib.sha256(raw.encode("utf-8")).hexdigest()
    man = os.path.join(dst, "manifest.csv")
    lines = open(man, encoding="utf-8-sig").read().splitlines()
    lines = ["book;WMS.ods" if ln.startswith("book;") else ln for ln in lines]
    lines = [f"file;service/_SYS.csv;{ln.split(';')[2]};{len(raw.encode('utf-8'))};{h}" if ln.startswith("file;service/_SYS.csv;") else ln
             for ln in lines]
    open(man, "w", encoding="utf-8", newline="\n").write("\n".join(lines) + "\n")
    print("fixture:", dst, sorted(os.listdir(dst)))


if __name__ == "__main__":
    main()
