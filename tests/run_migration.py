"""Migration importer (FINAL WMS MARATHON §4) — automated scenarios of tools/migrate.py: DRY RUN → MIGRATE COPY → VERIFY →
PROMOTE on production books (never opened before, as tools/build_ods.py makes them) and, for the crash of a transfer, on
test books. The tool runs as the user runs it (a separate process, its exit code and reports); the books are then opened
with WMS and checked by the independent oracle (tests/adjust_oracle.py: the whole journal replayed, MIGRATE included).
Synthetic tables only — no real warehouse data.

    WMS_TEST_OUT=/tmp/wms_mig python3 tests/run_migration.py [case ...]

Results: WMS_TEST_OUT/results_migration.json and WMS_TEST_OUT/TEST_REPORT_migration.md.
"""
import csv
import hashlib
import os
import random
import shutil
import subprocess
import sys
import time

import uno  # noqa: F401  (LibreOffice Python-UNO)

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from harness import OUT, PROFILES, INITIAL_STOCK, Session, Results, new_wms, template  # noqa: E402
from wmslo import Office, props  # noqa: E402

R = Results()
CASES = []
TOOL = os.path.join(os.path.dirname(HERE), "tools", "migrate.py")
PWD = "wms"
HEADER = ["ЕИ", "Наименование", "Артикул", "Единица", "Количество", "Место", "Категория", "Тип источника", "Старая маркировка", "Комментарий"]
BASE_ROWS = [
    ["ЕИ-00000500", "Старая дрель", "", "шт", "3", "M-1", "Инструмент", "Старый склад", "инв. 17", ""],
    ["350", "Шестерня Z-40", "gr-40", "шт", "12", "D-9", "Детали", "Детали", "", ""],
    ["", "Коробка без маркировки", "", "шт", "0", "M-2", "", "Иной приход", "", "нашли при переезде"],
    ["77", "Кабель ВВГ 3х2,5", "", "м", "250,5", "B-02", "Кабель", "Поставщик", "К-77", ""],
]
PLACES = ["M-1", "M-2", "D-9", "B-02"]


def case(fn):
    CASES.append(fn)
    return fn


def ei(n):
    return f"ЕИ-{n:08d}"


def sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


def cdir(name):
    d = os.path.join(OUT, "cases", name)
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(d)
    return d


def table(d, rows, name="table.csv", enc="utf-8-sig", header=HEADER):
    p = os.path.join(d, name)
    with open(p, "w", encoding=enc, newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(header)
        w.writerows(rows)
    return p


def places(d, lst=PLACES):
    p = os.path.join(d, "places.txt")
    open(p, "w", encoding="utf-8").write("\n".join(lst) + "\n")
    return p


def tool(*args, timeout=1800):
    t0 = time.time()
    r = subprocess.run([sys.executable, TOOL] + [str(a) for a in args], capture_output=True, text=True, timeout=timeout)
    return r.returncode, (r.stdout + r.stderr), time.time() - t0


def prod_book(case_name, name="WMS_PROD_CANDIDATE.ods"):
    return new_wms(case_name, kind="PROD", book_name=name)


def read(p):
    return open(p, encoding="utf-8").read() if os.path.exists(p) else ""


def oracle(c, s, name, initial=None):
    P, inf, _ = s.adjcheck(0, initial={} if initial is None else initial)
    R.add(c, "оракул: " + name, not P, f"{inf}; {P[:4]}")
    return P


def stock_row(s, n):
    return tuple(s.doc.Sheets.getByName("Наличие").getCellRangeByPosition(0, n, 9, n).getDataArray()[0])


# ================================================================ MG01–MG02 DRY RUN

@case
def mg01_dry_run_clean():
    c = "MG01"
    d = cdir("mg01")
    book = prod_book("mg01b")
    h0 = sha(book)
    t = table(d, BASE_ROWS)
    rep = os.path.join(d, "dry.md")
    rc, out, _ = tool("dry-run", t, "--book", book, "--places", places(d), "--report", rep)
    md = read(rep)
    R.add(c, "dry-run чистой таблицы: код 0, план — номера сохранены (500, 350, 77), строка без ЕИ получает новый номер выше всех (501), "
             "NEXT_EI после переноса 502; книга не изменилась",
          rc == 0 and "ЕИ-00000500 |" in md and "ЕИ-00000350 |" in md and "ЕИ-00000077 |" in md and "ЕИ-00000501 (новый)" in md
          and "NEXT_EI после переноса: 502" in md and sha(book) == h0 and not os.path.exists(os.path.join(os.path.dirname(book), "WMS_Journal")),
          f"код {rc}; {out[-300:]}")


@case
def mg02_dry_run_stops():
    c = "MG02"
    d = cdir("mg02")
    book = prod_book("mg02b")
    h0 = sha(book)
    pl = places(d, PLACES + ["Z-1"])
    bad = {
        "дубль ЕИ": [BASE_ROWS[0], ["500", "Другая дрель", "", "шт", "1", "M-1", "", "Старый склад", "", ""]],
        "артикул → несколько ЕИ": [BASE_ROWS[1], ["351", "Шестерня Z-40 (вторая)", "GR–40", "шт", "1", "D-9", "", "Детали", "", ""]],
        "неизвестная единица": [["600", "Лента", "", "штук", "1", "M-1", "", "Офис", "", ""]],
        "неизвестное место": [["601", "Лента", "", "шт", "1", "Z-99", "", "Офис", "", ""]],
        "тип источника": [["602", "Лента", "", "шт", "1", "M-1", "", "", "", ""]],
        "количество": [["603", "Лента", "", "шт", "-2", "M-1", "", "Офис", "", ""]],
        "ЕИ": [["ЕИ-12x", "Лента", "", "шт", "1", "M-1", "", "Офис", "", ""]],
        "нет ключевых данных": [["604", "", "", "шт", "1", "M-1", "", "Офис", "", ""]],
    }
    res = {}
    for k, rows in bad.items():
        t = table(d, rows, name=f"bad_{len(res)}.csv")
        rep = os.path.join(d, f"bad_{len(res)}.md")
        rc, out, _ = tool("dry-run", t, "--book", book, "--places", pl, "--report", rep)
        md = read(rep)
        res[k] = (rc, f"| БЛОК | {k}" in md, "План переноса" not in md)
    R.add(c, "dry-run останавливает перенос (код 1, плана нет): дубль ЕИ, артикул детали у двух ЕИ, неизвестная единица, неизвестное место, "
             "пустой тип источника (D-087), отрицательное количество, не номер ЕИ, нет наименования; книга не изменилась",
          all(v == (1, True, True) for v in res.values()) and sha(book) == h0, f"{res}")
    # an empty type with an explicit rule of the transfer; places accepted from the table explicitly
    t = table(d, [["602", "Лента", "", "шт", "1", "M-7", "", "", "", ""]], name="empty_type.csv")
    rep = os.path.join(d, "empty_type.md")
    rc1, _, _ = tool("dry-run", t, "--book", book, "--report", rep)
    md1 = read(rep)
    rc2, _, _ = tool("dry-run", t, "--book", book, "--default-source", "Старый склад", "--accept-table-places", "--report", rep)
    md2 = read(rep)
    rc3, out3, _ = tool("dry-run", t, "--book", book, "--default-source", "Склад", "--accept-table-places")
    R.add(c, "пустой тип и место без списка: без явных правил — стоп (тип, неизвестное место); с --default-source «Старый склад» и "
             "--accept-table-places — перенос разрешён, правила записаны в отчёт; --default-source не тип — ошибка 2",
          rc1 == 1 and "| БЛОК | тип источника" in md1 and "| БЛОК | неизвестное место" in md1 and rc2 == 0 and "Старый склад" in md2
          and "приняты из таблицы явно" in md2 and "M-7" in md2 and rc3 == 2, f"{rc1} {rc2} {rc3}: {out3[-120:]}")


# ================================================================ MG03–MG06 COPY → VERIFY → PROMOTE

def migrate_base(name):
    d = cdir(name)
    book = prod_book(name + "b")
    t = table(d, BASE_ROWS)
    pl = places(d)
    out = os.path.join(d, "copy")
    rc, txt, secs = tool("copy", t, "--book", book, "--places", pl, "--out", out)
    return d, book, t, pl, out, rc, txt, secs


@case
def mg03_copy_verify_promote():
    c = "MG03"
    d, book, t, pl, out, rc, txt, secs = migrate_base("mg03")
    h0 = sha(book)
    files = sorted(os.listdir(out)) if os.path.isdir(out) else []
    rep = read(os.path.join(out, "MIGRATION_REPORT.md"))
    bk = [f for f in os.listdir(os.path.join(out, "backup")) if f.endswith(".ods")] if os.path.isdir(os.path.join(out, "backup")) else []
    R.add(c, "MIGRATE COPY: код 0; в папке переноса — копия книги, её журнал, резервная копия исходной книги (SHA-256 совпадает) и отчёт; "
             "исходная книга не изменилась и не открывалась в WMS (рядом нет журнала)",
          rc == 0 and "WMS_PROD_CANDIDATE.ods" in files and "WMS_Journal" in files and "MIGRATION_REPORT.md" in files and len(bk) == 1
          and sha(os.path.join(out, "backup", bk[0])) == h0 and sha(book) == h0 and "Перенесено ЕИ: **4**" in rep
          and not os.path.exists(os.path.join(os.path.dirname(book), "WMS_Journal")), f"код {rc}; {files}; {txt[-300:]}")
    s = Session(os.path.join(out, "WMS_PROD_CANDIDATE.ods"))
    try:
        rows = {n: stock_row(s, n) for n in (77, 350, 500, 501)}
        ents = [e for e in s.journal()[0] if e["type"] == "MIGRATE"]
        ts = sha(t)
        R.add(c, "копия: ЕИ с прежними номерами (500, 350, 77) и новый 501 для строки без ЕИ; карточки как в таблице (тип источника, "
                 "источник «Перенос (маркировка)», «Иной приход» — «Требует разбора», количество 250,5 и 0); NEXT_EI 502; деталь GR-40 в индексе; "
                 "журнал — 4 операции MIGRATE с происхождением «таблица | SHA-256 | строка»",
              rows[500][:8] == (ei(500), "Старая дрель", "", "шт", 3.0, "M-1", "Инструмент", "Активен") and rows[500][8:] == ("Перенос (инв. 17)", "Старый склад")
              and rows[350][2] == "gr-40" and rows[350][9] == "Детали" and rows[77][4] == 250.5 and rows[77][8] == "Перенос (К-77)"
              and rows[501][:8] == (ei(501), "Коробка без маркировки", "", "шт", 0.0, "M-2", "", "Требует разбора") and rows[501][8:] == ("Перенос", "Иной приход")
              and s.sysv("NEXT_EI") == 502.0 and ("GR-40", ei(350), "gr-40") in s.art_index() and len(ents) == 4
              and sorted(e["fields"]["ORIGIN"] for e in ents) == sorted(f"table.csv|{ts}|{r}" for r in (2, 3, 4, 5)),
              f"{rows}; NEXT_EI {s.sysv('NEXT_EI')}; {[e['fields'].get('ORIGIN') for e in ents]}")
        oracle(c, s, "копия после переноса")
    finally:
        s.close()
    rc, txt, _ = tool("verify", t, "--dir", out)
    vr = read(os.path.join(out, "VERIFY_REPORT.md"))
    R.add(c, "VERIFY: код 0, расхождений 0 (каждая строка таблицы — одна операция MIGRATE и одна карточка, NEXT_EI, самопроверка, оракул); "
             "записан SHA-256 проверенной книги",
          rc == 0 and "расхождений: **0**" in vr and sha(os.path.join(out, "WMS_PROD_CANDIDATE.ods")) in vr, f"код {rc}; {txt[-400:]}")
    target = os.path.join(d, "prod")
    rc, txt, _ = tool("promote", "--dir", out, "--to", target)
    prod = os.path.join(target, "WMS_PROD.ods")
    R.add(c, "PROMOTE: код 0; рабочая папка — WMS_PROD.ods и журнал; отчёт PROMOTE_REPORT.md",
          rc == 0 and os.path.exists(prod) and os.path.isdir(os.path.join(target, "WMS_Journal")) and "работа разрешена" in read(os.path.join(target, "PROMOTE_REPORT.md")),
          f"код {rc}; {txt[-300:]}")
    rc2, txt2, _ = tool("promote", "--dir", out, "--to", target)
    R.add(c, "повторный PROMOTE в существующую папку — отказ (рабочая папка создаётся заново)", rc2 == 2, txt2[-150:])
    s = Session(prod)
    try:
        s.U("TestUiAuto", 1)
        st = s.state()
        s.doc.Sheets.getByName("Получатели").getCellRangeByPosition(0, 1, 1, 1).setDataArray((("ива", "Иванов Иван Андреевич"),))
        s.issue_input(1, ei="500", qty="1", date="25.09.2026", who="ива")
        post = s.post(1)
        s.special_input(1, B="Детали", E="GR-40", F="5", G="шт", H="25.09.2026")
        sp = s.click_spc("BtnSpcPost", 1)
        R.add(c, "после PROMOTE WMS работает: состояние «работа разрешена», выдача перенесённого ЕИ-00000500, пополнение перенесённой детали "
                 "GR-40 через «Иной приход» в тот же ЕИ-00000350 (12 → 17)",
              st.get("STATE") == "CLEAN" and post.startswith("OK") and s.stock(500) == 2.0 and sp.startswith("OK") and s.spc(1)[13] == ei(350)
              and s.stock(350) == 17.0, f"{st}; {post}; {sp}; {s.stock(350)}")
        oracle(c, s, "рабочая WMS после переноса и первых операций")
    finally:
        s.close()


@case
def mg04_verify_detects_changes():
    c = "MG04"
    d, book, t, pl, out, rc, txt, _ = migrate_base("mg04")
    # a copy of the transfer changed after it: one quantity in «Наличие» (the book opened without WMS)
    bad = os.path.join(d, "copy_changed")
    shutil.copytree(out, bad)
    s = Session(os.path.join(bad, "WMS_PROD_CANDIDATE.ods"), macros=0)
    try:
        sh = s.doc.Sheets.getByName("Наличие")
        sh.unprotect(PWD)
        sh.getCellByPosition(4, 500).setValue(30)
        s.doc.store()
    finally:
        s.close()
    rc1, txt1, _ = tool("verify", t, "--dir", bad)
    vr = read(os.path.join(bad, "VERIFY_REPORT.md"))
    rc2, txt2, _ = tool("promote", "--dir", bad, "--to", os.path.join(d, "prod_bad"))
    R.add(c, "VERIFY находит изменённую копию (остаток ЕИ-00000500 30 вместо 3: сверка с таблицей и оракул), код 1; PROMOTE такой копии — отказ",
          rc1 == 1 and "строка 2:" in vr and "оракул" in vr and rc2 == 1 and not os.path.exists(os.path.join(d, "prod_bad")), f"{rc1} {rc2}; {vr[-500:]}")
    # the verified copy changed after VERIFY
    rc3, _, _ = tool("verify", t, "--dir", out)
    s = Session(os.path.join(out, "WMS_PROD_CANDIDATE.ods"), macros=0)
    try:
        sh = s.doc.Sheets.getByName("Наличие")
        sh.unprotect(PWD)
        sh.getCellByPosition(5, 77).setString("B-03")
        s.doc.store()
    finally:
        s.close()
    rc4, txt4, _ = tool("promote", "--dir", out, "--to", os.path.join(d, "prod_late"))
    R.add(c, "копия изменена после успешной VERIFY — PROMOTE отказывает («книга изменилась после проверки»)",
          rc3 == 0 and rc4 == 1 and "изменилась после проверки" in txt4 and not os.path.exists(os.path.join(d, "prod_late")), f"{rc3} {rc4}; {txt4[-150:]}")
    # the table of the transfer changed after COPY: its SHA-256 is not the one of the journal
    t2 = table(d, BASE_ROWS[:3] + [["77", "Кабель ВВГ 3х2,5", "", "м", "250", "B-02", "Кабель", "Поставщик", "К-77", ""]], name="table.csv")
    rc5, _, _ = tool("verify", t2, "--dir", bad)
    R.add(c, "VERIFY с другой таблицей (изменена после переноса) — расхождение происхождения (SHA-256), код 1", rc5 == 1 and "не эта таблица" in read(os.path.join(bad, "VERIFY_REPORT.md")), f"{rc5}")


@case
def mg05_all_or_nothing():
    c = "MG05"
    d = cdir("mg05")
    book = new_wms("mg05b")                       # a test book: EIs 1–200 exist, the fault seam is there
    h0 = sha(book)
    rows = [[str(300 + i), f"Позиция {i}", "", "шт", str(i + 1), "A-01-1", "", "Старый склад", "", ""] for i in range(6)]
    t = table(d, rows)
    out = os.path.join(d, "copy")
    rc, txt, _ = tool("copy", t, "--book", book, "--accept-table-places", "--out", out, "--fault-at", 3)     # TestSetFault(2, 1)
    failed = read(out + "_MIGRATION_FAILED.md")
    R.add(c, "сбой третьего переноса (ошибка среди записей операции): перенос остановлен, папка копии удалена целиком (проведено 2 из 6 — "
             "не остаются), отчёт MIGRATION_FAILED; исходная книга не изменилась",
          rc == 1 and not os.path.exists(out) and "Причина остановки" in failed and "2 из 6" in failed and sha(book) == h0
          and not os.path.exists(os.path.join(os.path.dirname(book), "WMS_Journal")), f"код {rc}; {txt[-300:]}")
    rc2, txt2, _ = tool("copy", t, "--book", book, "--accept-table-places", "--out", out)
    s = Session(os.path.join(out, "WMS.ods")) if rc2 == 0 else None
    try:
        ok = s is not None and all(stock_row(s, 300 + i)[4] == float(i + 1) for i in range(6)) and s.sysv("NEXT_EI") == 306.0
        R.add(c, "повтор после исправления причины: перенос всех 6 строк целиком (ЕИ 300–305 рядом с существующими 1–200, NEXT_EI 306)", rc2 == 0 and ok,
              f"{rc2}; {txt2[-200:]}")
        if s is not None:
            oracle(c, s, "тестовая книга с переносом", initial=INITIAL_STOCK)
    finally:
        if s is not None:
            s.close()
    # the book is open in WMS: no copy
    s = Session(book)
    try:
        rc3, txt3, _ = tool("copy", t, "--book", book, "--accept-table-places", "--out", os.path.join(d, "copy_open"))
    finally:
        s.close()
    R.add(c, "книга открыта в WMS (есть wms.lock) — перенос не начинается", rc3 == 1 and "открыта в WMS" in txt3 and not os.path.exists(os.path.join(d, "copy_open")),
          txt3[-150:])


@case
def mg06_formats():
    c = "MG06"
    d = cdir("mg06")
    book = prod_book("mg06b")
    pl = places(d)
    plans = []
    t1 = table(d, BASE_ROWS, name="t_utf8.csv")
    t2 = table(d, BASE_ROWS, name="t_1251.csv", enc="cp1251")
    for t in (t1, t2):
        rep = os.path.join(d, os.path.basename(t) + ".md")
        rc, _, _ = tool("dry-run", t, "--book", book, "--places", pl, "--report", rep)
        md = read(rep)
        plans.append((rc, md[md.find("## План переноса"):]))
    # the same table as ODS (typed numbers: an EI and a quantity as numbers, an article as text)
    o = Office(f"mg06calc{os.getpid()}", PROFILES)
    try:
        doc = o.desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, props(Hidden=True))
        sh = doc.Sheets.getByIndex(0)
        data = [HEADER] + [[float(r[0][-3:]) if r[0] else "", r[1], r[2], r[3], float(r[4].replace(",", ".")), r[5], r[6], r[7], r[8], r[9]] for r in BASE_ROWS]
        sh.getCellRangeByPosition(0, 0, 9, len(data) - 1).setDataArray(tuple(tuple(x) for x in data))
        t3 = os.path.join(d, "t.ods")
        doc.storeToURL(uno.systemPathToFileUrl(t3), props(FilterName="calc8"))
        t4 = os.path.join(d, "t.xlsx")
        doc.storeToURL(uno.systemPathToFileUrl(t4), props(FilterName="Calc MS Excel 2007 XML"))
        doc.close(True)
    finally:
        o.terminate()
    for t in (t3, t4):
        rep = os.path.join(d, os.path.basename(t) + ".md")
        rc, _, _ = tool("dry-run", t, "--book", book, "--places", pl, "--report", rep)
        md = read(rep)
        plans.append((rc, md[md.find("## План переноса"):]))
    R.add(c, "одна таблица в CSV UTF-8, CSV Windows-1251, ODS и XLSX (числа как числа) — одинаковый план переноса",
          all(p[0] == 0 for p in plans) and len({p[1] for p in plans}) == 1 and "ЕИ-00000350" in plans[3][1], f"{[p[0] for p in plans]}")


@case
def mg07_scale_600():
    c = "MG07"
    d = cdir("mg07")
    book = prod_book("mg07b")
    rnd = random.Random(7)
    nums = rnd.sample(range(1, 9000), 560)
    units = ["шт", "м", "кг", "упак"]
    pls = [f"R-{i:02d}" for i in range(1, 31)]
    rows = []
    for i in range(600):
        e = str(nums[i]) if i < 560 else ""
        st = ["Старый склад", "Поставщик", "Офис", "Производство", "Детали", "Иной приход"][i % 6]
        art = f"P-{i:04d}" if st == "Детали" else (f"A{i}" if i % 5 == 0 else "")
        rows.append([e, f"Товар {i}", art, units[i % 4], f"{(i % 37) + (i % 3) * 0.25:g}".replace(".", ","), pls[i % 30], "Разное", st, f"M{i}" if i % 7 == 0 else "", ""])
    t = table(d, rows)
    out = os.path.join(d, "copy")
    rc, txt, secs = tool("copy", t, "--book", book, "--accept-table-places", "--out", out, timeout=3600)
    rc2, txt2, secs2 = tool("verify", t, "--dir", out, timeout=3600)
    vr = read(os.path.join(out, "VERIFY_REPORT.md"))
    R.add(c, f"600 строк (560 со своими номерами до 9000, 40 без номера, 100 деталей): перенос и проверка без расхождений; время переноса "
             f"{secs:.0f} с ({secs / 600 * 1000:.0f} мс на строку, с запуском LibreOffice), проверки — {secs2:.0f} с",
          rc == 0 and rc2 == 0 and "расхождений: **0**" in vr, f"{rc} {rc2}; {txt[-200:]} {vr[-300:]}")


def main():
    only = set(a.lower() for a in sys.argv[1:])
    t0 = time.time()
    template("PROD")
    template("TEST")
    for fn in CASES:
        if only and not any(fn.__name__.startswith(o) for o in only):
            continue
        try:
            fn()
        except Exception as e:
            R.error(fn.__name__, e)
        subprocess.run(["pkill", "-9", "-f", "profile_.*" + str(os.getpid())], capture_output=True)
    R.save(os.path.join(OUT, "results_migration.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_migration.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
