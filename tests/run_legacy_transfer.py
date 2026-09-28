"""M7 PRIME — the transfer of the old table «Заказы» into WMS 0.7: WMS_LEGACY_TRANSFER.ods and tools/legacy_transfer.py.

    WMS_TEST_OUT=/tmp/wms_out python3 tests/run_legacy_transfer.py [case ...]

The release folder is built once by tools/build_release.py (WMS_PROD_CANDIDATE.ods, WMS_LEGACY_TRANSFER.ods, tools/,
WMS_TOOLBOX/): every case uses the engine of that folder (without the repository on its path) and the clean candidate of the
release. A staging book (the four paste sheets of the frontend) is written through the API — the rows as a user pastes them
from the old table; the frontend itself is driven through the macros its buttons are bound to (test mode: no windows). The
built WMS is read like the tests of the core read a book (tests/harness.Session) and checked by the independent oracle
(tests/adjust_oracle.py — the whole journal replayed with the rules of the task, LEGACY_RECEIPT included) and by an
independent model of the old table written here (orders from column A, statuses, balances).
L01–L32 are the scenarios of the task (§19) in its order; L33–L39 the frontend, the version check, the rules, Doctor, Search
and Analytics on transferred EIs; L40–L42 the hotfix 0.7.2: old statuses and dates as people wrote them, the real
statuses of the user's table, no status from the date (D-089).
Results: WMS_TEST_OUT/results_legacy.json and WMS_TEST_OUT/TEST_REPORT_legacy.md.
"""
import datetime
import glob
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time

import uno  # noqa: F401  (LibreOffice Python-UNO)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
from harness import OUT, PROFILES, ROOT, Session, Results, git_sources, unique, wms_versions  # noqa: E402
from wmslo import Office, props  # noqa: E402

R = Results()
CASES = []
T = datetime.date.today()
TODAY = T.strftime("%d.%m.%Y")
LET = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
       "AA", "AB"]
HDR = ["№", "Полное наименование товара", "Номер документа", "Номер счёта", "Артикул", "Фактическое количество", "Количество по документу",
       "Заказанное количество", "Единица измерения", "Цена", "Сумма", "Площадка / Поставщик", "Продавец", "Дата поступления",
       "Дата документа", "Дата заказа", "Ожидаемая дата поступления", "Покупатель", "Категория", "Кому назначено", "Место хранения",
       "Внутренний код", "Статус", "Наличие", "Контроль", "Комментарий", "Срок поставки, дней", "Возможный дубль"]
ENV = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
REL = os.path.join(OUT, "lt_release")
_mk = {}


def case(fn):
    CASES.append(fn)
    return fn


def serial(d):
    return float((d - datetime.date(1899, 12, 30)).days)


def d(k):
    """the date k days from today as Calc keeps it"""
    return serial(T + datetime.timedelta(days=k))


def ei(n):
    return f"ЕИ-{n:08d}"


def row(**kw):
    r = [""] * 28
    for k, v in kw.items():
        r[LET.index(k)] = v
    return r


def sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


# ================================================================ the release, the staging, the engine

def release():
    if not os.path.isfile(os.path.join(REL, "WMS_LEGACY_TRANSFER.ods")):
        shutil.rmtree(REL, ignore_errors=True)
        r = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "build_release.py"), REL], capture_output=True, text=True, timeout=2400,
                           env=ENV, cwd=ROOT)
        if r.returncode != 0:
            raise RuntimeError("build_release: " + (r.stdout + r.stderr)[-2000:])
    return REL


def candidate():
    return os.path.join(release(), "WMS_PROD_CANDIDATE.ods")


def maker():
    if "o" not in _mk:
        _mk["o"] = Office(unique("ltmk"), PROFILES)
    return _mk["o"]


def staging(path, orders, stock=None, issues=None, returns=None, header=True, formulas=None):
    """a staging book: the paste sheets of the frontend with the given rows (orders: rows A..AB; the others: rows with
    their own header row first)"""
    o = maker()
    doc = o.new_calc()
    try:
        sh = doc.Sheets
        sh.getByIndex(0).Name = "1_Вставить_Заказы"
        for i, n in enumerate(["1B_Вставить_Наличие", "1C_Вставить_Выдачи", "1D_Вставить_Возвраты"], 1):
            sh.insertNewByName(n, i)
        data = ([HDR] if header else []) + [list(r) for r in orders]
        if data:
            sh.getByName("1_Вставить_Заказы").getCellRangeByPosition(0, 0, 27, len(data) - 1).setDataArray(tuple(tuple(r) for r in data))
        for name, rows in (("1B_Вставить_Наличие", stock), ("1C_Вставить_Выдачи", issues), ("1D_Вставить_Возвраты", returns)):
            if rows:
                w = max(len(r) for r in rows)
                rows = [list(r) + [""] * (w - len(r)) for r in rows]
                sh.getByName(name).getCellRangeByPosition(0, 0, w - 1, len(rows) - 1).setDataArray(tuple(tuple(r) for r in rows))
        for (sname, col, r), f in (formulas or {}).items():
            sh.getByName(sname).getCellByPosition(LET.index(col), r).setFormula(f)
        doc.storeToURL(uno.systemPathToFileUrl(path), props(FilterName="calc8"))
    finally:
        doc.close(True)
    return path


def cdir(name):
    p = os.path.join(OUT, "cases", name)
    shutil.rmtree(p, ignore_errors=True)
    os.makedirs(p)
    return p


def engine(mode, work, st=None, cand=None, extra=(), **rules):
    args = [sys.executable, os.path.join(release(), "tools", "legacy_transfer.py"), mode, "--work", work]
    if st:
        args += ["--staging", st]
    if cand:
        args += ["--candidate", cand]
    for k, v in rules.items():
        args += ["--" + k.replace("_", "-"), str(v)]
    args += list(extra)
    t0 = time.time()
    r = subprocess.run(args, capture_output=True, text=True, timeout=7200, env=ENV)
    return r.returncode, (r.stdout + r.stderr).strip(), time.time() - t0


def transfer(name, orders, stock=None, issues=None, returns=None, steps=("check", "build", "verify"), **rules):
    """a case folder, its staging and the steps of the engine; (folder, staging, work, {step: (rc, out, secs)})"""
    c = cdir(name)
    st = staging(os.path.join(c, "staging.ods"), orders, stock, issues, returns)
    work = os.path.join(c, "LEGACY_WORK")
    res = {}
    for m in steps:
        if m == "promote":
            res[m] = engine(m, work, st, extra=("--to", os.path.join(c, "WMS_WORK")), **rules)
        else:
            res[m] = engine(m, work, st, candidate() if m in ("check", "build") else None, **rules)
        if res[m][0] != 0:
            break
    return c, st, work, res


def summary(work):
    """«2_Проверка» as the engine wrote it: key → value"""
    out = {}
    for ln in open(os.path.join(work, "check", "summary.csv"), encoding="utf-8").read().splitlines()[1:]:
        p = ln.split(";")
        out[p[-1]] = p[1]
    return out


def findings(work):
    rows = []
    for ln in open(os.path.join(work, "check", "errors.csv"), encoding="utf-8").read().splitlines()[1:]:
        p = ln.split(";", 5)
        rows.append(dict(sheet=p[1], row=p[2], level=p[3], kind=p[4], text=p[5]))
    return rows


def row_results(work):
    out = {}
    for ln in open(os.path.join(work, "check", "rows.csv"), encoding="utf-8").read().splitlines()[1:]:
        p = ln.split(";", 2)
        out[int(p[0])] = (p[1], p[2])
    return out


def state(work):
    return json.load(open(os.path.join(work, "state.json"), encoding="utf-8"))


def test_book(work):
    return os.path.join(work, "test", "WMS_LEGACY_TEST.ods")


def orders_of(s, n=None):
    """«Заказы» rows 1.. (A..AC) of an open book"""
    sh = s.doc.Sheets.getByName("Заказы")
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    last = n or cur.getRangeAddress().EndRow
    return [list(r) for r in sh.getCellRangeByPosition(0, 1, 28, max(last, 1)).getDataArray()]


def ready(s):
    s.U("TestUiAuto", 1)
    return s


def oracle(c, s, name):
    P, inf, _ = s.adjcheck(0, initial={})
    R.add(c, "оракул (журнал воспроизведён независимо): " + name, not P, f"{inf.get('ops')} операций; {P[:4]}")
    return P


def blocks_of(rows):
    """the orders as WMS draws them: a new block where TRIM(A)=1 on a row that is not a delivery row (AC empty)"""
    n = 0
    for r in rows:
        if str(r[0]).strip() == "1" and not r[28]:
            n += 1
    return n


# ================================================================ the old table of the scenarios (an independent model below)

def full_rows():
    """the old «Заказы» of the main scenarios: eight orders with the old statuses received, waiting, «Просрочено», partly received,
    «Частично получено / просрочено» (in WMS 0.7.2 both are open positions, D-089), without documents, cancelled, rest cancelled;
    delivery rows; a position whose first receipt is a child row; parts; a special receipt; EIs in every form; a large EI; a wrong
    old status"""
    return [
        row(A=1.0, B="Болт М8х30", C="УПД-101", D="СЧ-7", E="DIN933-M8", F=100.0, G=100.0, H=100.0, I="шт", J=5.5, K=550.0, L="Ромашка",
            M="Иванов", N=d(-30), O=d(-31), P=d(-40), Q=d(-32), R="Цех 1", S="Крепёж", T="Петров", U="A-01", V="ЕИ-00000427", W="Получено",
            X=60.0, Y="Норма", Z="коммент", AA=10.0, AB="Возможный дубль: УПД-101"),                                          # 2
        row(A=2.0, B="Гайка М8", E="DIN934-M8", H=200.0, I="шт", L="Ромашка", P=d(-40), Q=d(10), W="Ожидается"),                   # 3
        row(A=3.0, B="Шайба М8", C="УПД-102", F=100.0, H=300.0, I="шт", L="Ромашка", N=d(-20), O=d(-20), Q=d(-5), U="A-02", V="428",
            W="Частично получено", X=100.0),                                                                                   # 4
        row(A=1.0, B="Кабель ВВГ 3х2,5", C="УПД-201", E="VVG-3x2.5", F=200.0, H=500.0, I="м", L="Лютик", N=d(-15), O=d(-15), Q=d(20),
            U="B-01", V="EI-500", W="Частично получено", X=150.0),                                                              # 5
        row(A=1.0, B="Кабель ВВГ 3х2,5", C="УПД-202", F=100.0, N=d(-5), O=d(-5), U="B-02", V="00000501", X=100.0),                 # 6
        row(A=2.0, B="Лампа LED E27", H=50.0, I="шт", L="Лютик", W="Отменено"),                                                  # 7
        row(A=3.0, B="Перчатки нитриловые", F=40.0, H=100.0, I="пар", L="Лютик", N=d(-10), U="C-01", W="Частично получено / остаток отменён",
            X=40.0),                                                                                                          # 8
        row(A=1.0, B="Краска белая", F=10.0, H=10.0, I="шт", L="Колор", N=d(-3), U="D-01", V="ЕИ-00000600", W="Получено без документов",
            X=0.0),                                                                                                           # 9
        row(A=2.0, B="Смазка литиевая", H=5.0, I="кг", L="Колор", Q=d(-2), W="Просрочено"),                                      # 10
        row(A=3.0, B="Деталь насоса", E="PN-100", F=2.0, I="шт", L="Детали", N=d(-50), U="E-01", V="700", X=1.0),                  # 11
        row(A=4.0, B="Деталь насоса", E="PN-100", F=3.0, I="шт", L="Детали", N=d(-20), U="E-01", X=1.0),                           # 12
        row(A=1.0, B="Стеллаж металлический", F=1.0, I="шт", L="Офис", N=d(-100), U="Офис", V="ЕИ 800", X=1.0),                     # 13
        row(A=1.0, B="Кабель UTP", H=300.0, I="м", L="Сеть", Q=d(30), W="Частично получено"),                                     # 14
        row(C="УПД-301", F=100.0, N=d(-1), O=d(-1), U="F-01", V="ЕИ-00000900", X=80.0),                                             # 15
        row(C="УПД-302", F=50.0, N=d(0), O=d(0), U="F-01", V="ЕИ-00000901", X=50.0),                                                # 16
        row(A=1.0, B="Винт М4", C="УПД-401", F=10.0, H=10.0, I="шт", L="Метиз", N=d(-8), O=d(-8), U="G-01", V="еи-1001", W="Получено",
            X=10.0),                                                                                                          # 17
        row(A=2.0, B="Винт М5", C="УПД-401", F=10.0, H=10.0, I="шт", L="Метиз", N=d(-8), O=d(-8), U="G-01", V=1002.0, W="Получено", X=5.0),  # 18
        row(A=1.0, B="Хомут", C="УПД-501", F=5.0, H=5.0, I="шт", L="Метиз", N=d(-4), O=d(-4), U="G-02", V="ЕИ-00005000", W="Получено",
            X=5.0),                                                                                                           # 19
        row(A=1.0, B="Кран шаровый", C="УПД-601", F=4.0, H=4.0, I="шт", L="Вода", N=d(-6), O=d(-6), U="H-01", V="1100", W="Ожидается",
            X=4.0),                                                                                                           # 20
    ]


# an independent model of the rows above (written from the task, not from the engine): row → (status, EI, balance)
FULL_EXPECT = {
    2: ("Получено", ei(427), 60.0), 3: ("Ожидается", "", None), 4: ("Частично получено", ei(428), 100.0),
    5: ("Частично получено", ei(500), 150.0), 6: ("Дополнительное поступление", ei(501), 100.0), 7: ("Отменено", "", None),
    8: ("Частично получено / остаток отменён", ei(5001), 40.0), 9: ("Получено без документов", ei(600), 0.0), 10: ("Ожидается", "", None),
    14: ("Частично получено", ei(900), 80.0), 16: ("Дополнительное поступление", ei(901), 50.0), 17: ("Получено", ei(1001), 10.0),
    18: ("Получено", ei(1002), 5.0), 19: ("Получено", ei(5000), 5.0), 20: ("Получено", ei(1100), 4.0)}


def full_case():
    """the main dataset transferred once (check, build, verify, promote); later cases read its results"""
    if "full" not in _mk:
        c, st, work, res = transfer("full", full_rows(), steps=("check", "build", "verify", "promote"))
        _mk["full"] = (c, st, work, res)
    return _mk["full"]


def full_new_rows(s):
    """the new «Заказы» of the main dataset by the staging row (the plan's map: staging row → new row)"""
    c, st, work, res = full_case()
    plan = json.load(open(os.path.join(work, "check", "plan.json"), encoding="utf-8"))
    rows = orders_of(s)
    return {m["row"]: rows[m["t"] - 1] for m in plan["mapping"] if m["t"] is not None and m["fate"] == "row"}, plan


# ================================================================ L01–L02 orders from column A

@case
def l01_one_order():
    c = "L01"
    rows = [row(A=1.0, B="Позиция 1", H=10.0, I="шт", L="Поставщик-1", Q=d(5)), row(A=2.0, B="Позиция 2", H=20.0, I="шт", L="Поставщик-1"),
            row(A=3.0, B="Позиция 3", H=30.0, I="м", L="Поставщик-1")]
    cd, st, work, res = transfer("l01", rows, steps=("check", "build"))
    m = summary(work) if os.path.exists(os.path.join(work, "check", "summary.csv")) else {}
    R.add(c, "один заказ 1,2,3: проверка — 1 заказ, 3 позиции, все ожидаются, блокирующих замечаний нет",
          res["check"][0] == 0 and m.get("orders") == "1" and m.get("positions") == "3" and m.get("waiting") == "3" and m.get("blockers") == "0",
          f"{res['check'][1][-200:]}; {m.get('orders')}/{m.get('positions')}")
    s = ready(Session(test_book(work)))
    try:
        rows_ = orders_of(s, 4)
        R.add(c, "тестовая WMS: три строки заказа с № 1, 2, 3 — один блок, статусы «Ожидается», ЕИ нет (приходов не было)",
              res["build"][0] == 0 and [r[0] for r in rows_[:3]] == ["1", "2", "3"] and blocks_of(rows_[:3]) == 1
              and [r[22] for r in rows_[:3]] == ["Ожидается"] * 3 and all(r[21] == "" for r in rows_[:3]) and not any(rows_[3][:28]),
              f"{res['build'][1][-200:]}; {[r[:1] + r[21:23] for r in rows_[:4]]}")
    finally:
        s.close()


@case
def l02_several_orders():
    c = "L02"
    seq = [1, 2, 3, 1, 2, 1, 2, 3, 4]
    rows = [row(A=float(a), B=f"Товар {i + 1}", H=float(10 + i), I="шт", L="Поставщик", Q=d(3)) for i, a in enumerate(seq)]
    cd, st, work, res = transfer("l02", rows, steps=("check", "build"))
    m = summary(work)
    R.add(c, "заказы 1,2,3,1,2,1,2,3,4 — проверка: 3 заказа, 9 позиций",
          res["check"][0] == 0 and m.get("orders") == "3" and m.get("positions") == "9", f"{m.get('orders')}/{m.get('positions')}")
    s = ready(Session(test_book(work)))
    try:
        rows_ = orders_of(s, 10)
        starts = [i + 2 for i, r in enumerate(rows_[:9]) if r[0] == "1"]
        R.add(c, "тестовая WMS: те же № позиций в том же порядке; блоки заказов начинаются в строках 2, 5, 7 (линия-разделитель WMS: A = 1)",
              res["build"][0] == 0 and [r[0] for r in rows_[:9]] == [str(a) for a in seq] and starts == [2, 5, 7] and blocks_of(rows_[:9]) == 3,
              f"{[r[0] for r in rows_[:9]]}; {starts}")
    finally:
        s.close()


# ================================================================ L03–L20 on the main dataset

@case
def l03_to_l18():
    c, st, work, res = full_case()
    ok = all(res[k][0] == 0 for k in ("check", "build", "verify", "promote"))
    R.add("L03", "основной набор (8 заказов со старыми статусами: получено, ожидается, «Просрочено», частично, «Частично получено / просрочено», без документов, отменено, "
                 "остаток отменён, поступления, детали, специальный приход, ЕИ во всех формах): проверка, тестовая WMS, сверка, рабочая WMS — "
                 "все шаги выполнены",
          ok, "; ".join(f"{k}: {v[0]} {v[1][-160:]} ({v[2]:.0f} с)" for k, v in res.items()))
    m = summary(work)
    s = ready(Session(os.path.join(c, "WMS_WORK", "WMS_PROD.ods")))
    try:
        by_row, plan = full_new_rows(s)
        got = {r: (v[22], v[21], v[23] if v[23] != "" else None) for r, v in by_row.items()}
        want = {r: FULL_EXPECT[r] for r in FULL_EXPECT if r in by_row}
        bad = {r: (got.get(r), w) for r, w in want.items() if got.get(r) != w}
        R.add("L03", "полностью полученный заказ: «Получено», ЕИ-00000427, приход 100 = заказано; позиция _ORD получено 100 из 100",
              got.get(2) == FULL_EXPECT[2] and s.opos(1)[4:7] == (100.0, 100.0, 1.0), f"{got.get(2)}; {s.opos(1)}")
        R.add("L04", "ожидаемый: строка 3 «Ожидается» (ожидаемая дата впереди), ЕИ нет, позиции в _ORD нет — обычная строка заказа",
              got.get(3) == FULL_EXPECT[3], str(got.get(3)))
        R.add("L05", "просроченный в старой таблице («Просрочено», ожидаемая дата прошла, ничего не получено): в новой WMS «Ожидается» — заказ "
                     "ждём до поставки или отмены, просрочки как статуса нет (D-089)", got.get(10) == FULL_EXPECT[10], str(got.get(10)))
        R.add("L06", "частично полученный: «Частично получено» — приход 200 и поступление 100 из 500; ожидается 200",
              got.get(5) == FULL_EXPECT[5] and got.get(6) == FULL_EXPECT[6] and "осталось 200" in by_row[5][24], f"{got.get(5)}; {by_row[5][24]}")
        R.add("L07", "частично полученный с прошедшей ожидаемой датой: «Частично получено» (не «Частично получено / просрочено» — "
                     "статус от даты не зависит, D-089); «2_Проверка»: с прошедшей ожидаемой датой — информация",
              got.get(4) == FULL_EXPECT[4] and int(m.get("late_expected", "0")) >= 2, f"{got.get(4)}; late {m.get('late_expected')}")
        R.add("L08", "без документов: «Получено без документов» (у прихода нет номера и даты документа)", got.get(9) == FULL_EXPECT[9],
              str(got.get(9)))
        R.add("L09", "отменённый: «Отменено» (ORDER_CANCEL, ЕИ не создавался); «Отменено» при частичном приходе → «остаток отменён»",
              got.get(7) == FULL_EXPECT[7] and got.get(8) == FULL_EXPECT[8] and "ORDER" in [s.opos(k)[8] for k in range(1, 12)]
              and "REST" in [s.opos(k)[8] for k in range(1, 12)], f"{got.get(7)}; {got.get(8)}")
        R.add("L10", "старый ЕИ сохраняется: ЕИ-00000427 остаётся ЕИ-00000427 (строка 427 «Наличие», приход в _RCV строка 427)",
              s.stock_row(427)[0] == ei(427) and s.rcv(427)[0] == ei(427) and s.rcv(427)[4:6] == ("LIVE", "SRC"), f"{s.stock_row(427)}; {s.rcv(427)}")
        forms = {17: ("еи-1001", 1001), 18: (1002.0, 1002), 4: ("428", 428), 5: ("EI-500", 500), 6: ("00000501", 501), 13: ("ЕИ 800", 800)}
        normed = all(s.stock_row(n)[0] == ei(n) for _, n in forms.values())
        rr = row_results(work)
        R.add("L11", "разные формы ЕИ (427, 00000427, ЕИ-427, EI-427, еи 1001, число 1002) → канонический ЕИ-0000xxxx; в проверке — "
                     "«нормализовано»", normed and all("ЕИ «" in rr[r][1] or "→ ЕИ-" in rr[r][1] for r in (4, 5, 6, 13, 17, 18)),
              f"{[s.stock_row(n)[0] for _, n in forms.values()]}; {[rr[r][1][:50] for r in (4, 5, 17, 18)]}")
        part = s.stock_row(700)
        art = s.art_index()
        R.add("L13", "деталь повторно (две строки «Детали», артикул PN-100): один ЕИ-00000700, тип «Детали», индекс артикулов — одна запись; "
                     "повторная строка — пополнение того же ЕИ", part[0] == ei(700) and s.doc.Sheets.getByName("Наличие").getCellByPosition(9, 700).getString()
              == "Детали" and len([a for a in art if a[0] == "PN-100"]) == 1 and "пополнение того же ЕИ" in rr[12][1], f"{part}; {art}; {rr[12]}")
        R.add("L15", "текущий остаток меньше прихода: пришло 100, наличие 60 — остаток ЕИ 60 (израсходованное до переноса — как выдано)",
              s.stock(427) == 60.0 and by_row[2][23] == 60.0, f"{s.stock(427)}")
        R.add("L16", "нулевой остаток: ЕИ-00000600 перенесён с остатком 0 (ЕИ сохранён)", s.stock_row(600)[0] == ei(600) and s.stock(600) == 0.0,
              str(s.stock_row(600)))
        R.add("L17", "пустой ЕИ → новый номер выше всех старых: строка 8 получила ЕИ-00005001", by_row[8][21] == ei(5001) and s.stock(5001) == 40.0,
              str(by_row[8][21:24]))
        R.add("L18", "максимальный старый ЕИ 5000 → NEXT_EI 5002 (после нового 5001); «2_Проверка»: максимальный ЕИ и NEXT_EI",
              s.sysv("NEXT_EI") == 5002.0 and m.get("max_ei") == ei(5000) and m.get("next_ei") == "5002", f"{s.sysv('NEXT_EI')}; {m.get('max_ei')}")
        R.add("L03", "все строки по независимой модели старой таблицы (статус, ЕИ, наличие)", not bad, str(bad)[:600])
        # the old values are kept: the journal of every transferred receipt (W Y X AB of its rows) and the map of all rows of the
        # old table in the working folder, the waiting positions too (old → new status, EI, old → new balance)
        import csv
        import journal_oracle
        ents, _ = journal_oracle.read_journal(os.path.join(c, "WMS_WORK", "WMS_Journal"))
        j427 = next((e["fields"] for e in ents if e["type"] == "LEGACY_RECEIPT" and e["fields"].get("EI") == ei(427)), {})
        j501 = next((e["fields"] for e in ents if e["type"] == "LEGACY_RECEIPT_ADD" and e["fields"].get("EI") == ei(501)), {})
        with open(os.path.join(c, "WMS_WORK", "LEGACY_TRANSFER", "LEGACY_ROWS.csv"), encoding="utf-8") as f:
            lr = {int(r[0]): r for r in list(csv.reader(f, delimiter=";"))[1:]}
        R.add("L03", "старые W, X, Y, AB сохранены: журнал прихода (LEGACY_STATUS «Получено», LEGACY_CTL «Норма», LEGACY_STOCK 60, "
                     "LEGACY_DUP) и карта строк LEGACY_ROWS.csv рабочей папки — у каждой строки, и у ожидаемой: ЕИ, старый → новый статус, "
                     "старое → новое наличие",
              j427.get("LEGACY_STATUS") == "Получено" and j427.get("LEGACY_CTL") == "Норма" and j427.get("LEGACY_STOCK") == "60"
              and j427.get("LEGACY_DUP") == "Возможный дубль: УПД-101" and "LEGACY_DUP" in j501
              and lr[2][5:12] == [ei(427), "Получено", "Получено", "60", "60", "Норма", "Возможный дубль: УПД-101"]
              and lr[3][5:8] == ["", "Ожидается", "Ожидается"] and lr[4][6:8] == ["Частично получено", "Частично получено"]
              and lr[6][5] == ei(501) and lr[6][7] == "Дополнительное поступление",
              f"{ {k: j427.get(k) for k in ('LEGACY_STATUS', 'LEGACY_CTL', 'LEGACY_STOCK', 'LEGACY_DUP')} }; {lr.get(2)}; {lr.get(3)}; {lr.get(6)}")
        oracle("L03", s, "рабочая WMS после переноса")
    finally:
        s.close()


@case
def l12_duplicate_ei():
    c = "L12"
    rows = [row(A=1.0, B="Болт", F=5.0, H=5.0, I="шт", L="П", N=d(-3), U="A", V="ЕИ-00000077", X=5.0),
            row(A=2.0, B="Гайка", F=3.0, H=3.0, I="шт", L="П", N=d(-3), U="A", V="77", X=3.0)]
    cd, st, work, res = transfer("l12", rows, steps=("check", "build"))
    F = findings(work)
    dup = [f for f in F if f["kind"] == "дубль ЕИ"]
    R.add(c, "один ЕИ у двух разных товаров — БЛОКЕР «дубль ЕИ» у обеих строк; «2_Проверка»: дубли ЕИ 1; тестовая WMS не создаётся",
          res["check"][0] == 1 and len(dup) == 2 and all(f["level"] == "БЛОКЕР" and "разные товары" in f["text"] for f in dup)
          and summary(work).get("dup_ei") == "1" and "build" not in res and not os.path.exists(os.path.join(work, "test")),
          f"{res['check'][1][-160:]}; {dup}")
    rc, out, _ = engine("build", work, st, candidate())
    R.add(c, "«СОЗДАТЬ ТЕСТОВУЮ WMS» при блокере — отказ, копии нет", rc == 1 and not os.path.exists(os.path.join(work, "test"))
          and not os.path.exists(os.path.join(work, "test.part")), out[-200:])


@case
def l14_article_two_eis():
    c = "L14"
    rows = [row(A=1.0, B="Фильтр", E="F-10", F=2.0, I="шт", L="Детали", N=d(-9), U="E", V="310", X=2.0),
            row(A=2.0, B="Фильтр", E="f 10", F=1.0, I="шт", L="Детали", N=d(-8), U="E", V="311", X=1.0),
            row(A=3.0, B="Фильтр", E="F‑10", F=1.0, I="шт", L="Детали", N=d(-7), U="E", V="312", X=1.0)]
    cd, st, work, res = transfer("l14", rows, steps=("check",))
    F = findings(work)
    part = [f for f in F if f["kind"] == "деталь" and f["level"] == "БЛОКЕР"]
    R.add(c, "один артикул детали у двух ЕИ (F-10 с разными дефисами — один ключ артикула) — БЛОКЕР «деталь»; «f 10» — другой артикул",
          res["check"][0] == 1 and len(part) == 1 and "ЕИ-00000310" in part[0]["text"] and "ЕИ-00000312" in part[0]["text"]
          and summary(work).get("part_conflicts") == "1", str(part))


@case
def l19_formulas():
    c = "L19"
    cd = cdir("l19")
    old = os.path.join(cd, "old.ods")
    o = maker()
    doc = o.new_calc()
    doc.Sheets.getByIndex(0).Name = "Заказы"
    doc.Sheets.insertNewByName("Выдачи", 1)
    doc.Sheets.getByName("Выдачи").getCellByPosition(1, 1).setValue(35.0)
    doc.storeToURL(uno.systemPathToFileUrl(old), props(FilterName="calc8"))
    doc.close(True)
    rows = [row(A=1.0, B="Труба 20", C="УПД-9", F=50.0, H=50.0, I="м", J=12.0, L="Сталь", N=d(-5), O=d(-5), U="T-1", V="ЕИ-00000990")]
    f = {("1_Вставить_Заказы", "X", 1): f"='{uno.systemPathToFileUrl(old)}'#$Выдачи.B2",
         ("1_Вставить_Заказы", "W", 1): '=IF(F2>=H2;"Получено";"Ожидается")', ("1_Вставить_Заказы", "K", 1): "=F2*J2"}
    fe = frontend(cd)
    try:
        fe.paste(rows, formulas=f)
        before = fe.formula_count()
        fe.press("BtnLtCheck")
        after = fe.formula_count()
        x = fe.cell("1_Вставить_Заказы", "X", 1)
        fe.press("BtnLtBuild")
        msg = fe.msg()
        fe.B("LtTestConfirm", 1)
        fe.press("BtnLtClear")
        after_clear = fe.main()
    finally:
        fe.close()
    R.add(c, "«Очистить staging» до «СДЕЛАТЬ РАБОЧЕЙ»: шаги на главной начинаются заново, созданная тестовая WMS остаётся",
          after_clear[1].startswith("—") and after_clear[2].startswith("—") and os.path.isfile(test_book(os.path.join(cd, "LEGACY_WORK"))),
          str(after_clear[1:5]))
    s = ready(Session(test_book(os.path.join(cd, "LEGACY_WORK"))))
    try:
        sh = s.doc.Sheets.getByName("Заказы")
        cells = sh.queryContentCells(16).getRangeAddresses()
        links = s.doc.ExternalDocLinks.getElementNames() if hasattr(s.doc, "ExternalDocLinks") else ()
        r1 = orders_of(s, 2)[0]
        R.add(c, "формулы во вставленных данных (ссылка на старую книгу, вычисленный статус, сумма) → при «ПРОВЕРИТЬ» заменены значениями "
                 "(ссылка на старую книгу — её значение 35); тестовая WMS: в «Заказы» нет формул и внешних ссылок, наличие 35, сумма 600",
              before == 3 and after == 0 and x == ("VALUE", 35.0) and msg.startswith("OK:") and not cells and not links and r1[23] == 35.0
              and r1[10] == 600.0 and r1[22] == "Получено", f"{before}/{after}; {x}; {msg[:120]}; {cells}; {links}; {r1[10]}, {r1[22:24]}")
    finally:
        s.close()


@case
def l20_status_error():
    c = "L20"
    c_, st, work, res = full_case()
    F = findings(work)
    st20 = [f for f in F if f["row"] == "20" and f["kind"] == "статус"]
    R.add(c, "старая статусная ошибка (строка 20: «Ожидается», а приход 4 из 4 и ЕИ есть) — новый движок: «Получено», проверка сообщает "
             "(жёлтый), отчёт — «Пересчитано новым движком»",
          FULL_EXPECT[20][0] == "Получено" and st20 and st20[0]["level"] == "ПРОВЕРИТЬ" and "«Ожидается» → «Получено»" in st20[0]["text"]
          and "строка 20" in open(os.path.join(work, "LEGACY_TRANSFER_REPORT.md"), encoding="utf-8").read(), str(st20))
    rows = [row(A=1.0, B="Кран", H=4.0, I="шт", L="Вода", W="Получено"), row(A=2.0, B="Вентиль", H=2.0, I="шт", L="Вода", F=2.0, N=d(-2),
                                                                                U="V", V="1200", X=2.0, W="Отменено")]
    cd, st2, work2, res2 = transfer("l20", rows, steps=("check",))
    F2 = [f for f in findings(work2) if f["kind"] == "статус" and f["level"] == "БЛОКЕР"]
    R.add(c, "критические противоречия — БЛОКЕР: «Получено» без количества F; «Отменено», а позиция получена полностью",
          res2["check"][0] == 1 and len(F2) == 2 and summary(work2).get("status_critical") == "2", str(F2))


# ================================================================ L21–L27 new operations on the transferred WMS

def prod_session():
    c, st, work, res = full_case()
    return ready(Session(os.path.join(c, "WMS_WORK", "WMS_PROD.ods")))


@case
def l21_to_l27():
    s = prod_session()
    p = s.path
    try:
        s.doc.Sheets.getByName("Получатели").getCellRangeByPosition(0, 1, 1, 1).setDataArray((("ива", "Иванов Иван"),))
        by_row, plan = full_new_rows(s)
        t5 = [m["t"] for m in plan["mapping"] if m["row"] == 5][0]
        r = s.Rc("ReceiptAddRow", t5, "50", "", "УПД-203", TODAY, TODAY, "B-03", "")
        rows_ = orders_of(s)
        pos = [op for op in plan["positions"] if op["t"] == t5][0]
        R.add("L21", "«Ещё поступление» по перенесённой частичной позиции (Кабель ВВГ, получено 300 из 500): новый ЕИ-00005002 от NEXT_EI, "
                     "строка под поступлениями позиции, позиция получено 350, статус «Частично получено»",
              r.startswith("OK:") and "ЕИ-00005002" in r and s.opos(pos["olid"])[5] == 350.0 and rows_[t5 - 1][22] == "Частично получено"
              and rows_[t5 + 1][21] == ei(5002) and rows_[t5 + 1][28] == f"OL{pos['olid']}", f"{r}; {s.opos(pos['olid'])}; {rows_[t5 + 1][21:23]}")
        s.issue_input(1, ei="427", qty="10", date=TODAY, who="ива")
        r1 = s.post(1)
        s.issue_input(2, ei="600", qty="1", date=TODAY, who="ива")
        r2 = s.post(2)
        s.issue_input(3, ei="427", qty="51", date=TODAY, who="ива")
        r3 = s.post(3)
        R.add("L22", "выдача перенесённого ЕИ: ЕИ-00000427 выдано 10 из 60 → остаток 50 (и в «Заказы».X); ЕИ с нулевым остатком и выдача "
                     "больше остатка — отказ",
              r1.startswith("OK") and s.stock(427) == 50.0 and not r2.startswith("OK") and not r3.startswith("OK")
              and orders_of(s)[by_row_t(plan, 2) - 1][23] == 50.0, f"{r1}; {r2}; {r3}")
        s.return_input(1, issue="1", qty="5", date=TODAY)
        r4 = s.click_ret("BtnRetPost", 1)
        R.add("L23", "возврат по выдаче перенесённого ЕИ: 5 → остаток 55", r4.startswith("OK") and s.stock(427) == 55.0, r4)
        s.adjust_input(1, B="Перемещение", C="427", J="Z-9", K=TODAY, L="переезд")
        r5 = s.click_adj("BtnAdjPost", 1)
        R.add("L24", "MOVE перенесённого ЕИ: место A-01 → Z-9", r5.startswith("OK") and s.stock_row(427)[5] == "Z-9", r5)
        s.adjust_input(2, B="Списание", C="427", G="1", K=TODAY, L="брак")
        r6 = s.click_adj("BtnAdjPost", 2)
        R.add("L25", "WRITE_OFF перенесённого ЕИ: списано 1 → остаток 54", r6.startswith("OK") and s.stock(427) == 54.0, r6)
        s.adjust_input(3, B="Инвентаризация", C="427", H="50", I="54", K=TODAY, L="пересчёт")
        r7 = s.click_adj("BtnAdjPost", 3)
        R.add("L26", "INV_ADJ перенесённого ЕИ: учётный 54, фактический 50 → остаток 50", r7.startswith("OK") and s.stock(427) == 50.0, r7)
        oracle("L26", s, "перенос + «Ещё поступление», выдача, возврат, перемещение, списание, инвентаризация")
        s.close(save=True)
        s = ready(Session(p))
        st_ = s.state()
        sc = s.B("ActionSelfCheck")
        R.add("L27", "сохранение, закрытие, повторное открытие рабочей WMS: «работа разрешена», самопроверка без ошибок, остатки на месте",
              st_.get("STATE") == "CLEAN" and sc.startswith("САМОПРОВЕРКА WMS: ошибок 0") and s.stock(427) == 50.0 and s.stock(5002) == 50.0,
              f"{st_.get('STATE')}; {sc.splitlines()[0]}")
        oracle("L27", s, "после повторного открытия")
    finally:
        s.close()


def by_row_t(plan, r):
    return [m["t"] for m in plan["mapping"] if m["row"] == r][0]


# ================================================================ L28 no repeated import

@case
def l28_repeat_blocked():
    c = "L28"
    c_, st, work, res = full_case()
    prod = os.path.join(c_, "WMS_WORK", "WMS_PROD.ods")
    w2 = os.path.join(cdir("l28"), "LEGACY_WORK")
    rc, out, _ = engine("check", w2, st, prod)
    F = [f for f in findings(w2) if f["kind"] == "книга-кандидат"]
    R.add(c, "повторный перенос в уже работающую WMS (она указана как книга-кандидат) — отказ: «книга-кандидат уже использовалась»",
          rc == 1 and F and any("уже использовалась" in f["text"] for f in F), f"{out[-160:]}; {F[:2]}")
    rc2, out2, _ = engine("promote", work, st, extra=("--to", os.path.join(OUT, "cases", "l28", "WMS_WORK2")))
    R.add(c, "повторное «СДЕЛАТЬ РАБОЧЕЙ» — отказ: перенос уже сделан рабочим", rc2 == 1 and "уже сделан рабочим" in out2, out2[-200:])
    s = ready(Session(prod))
    try:
        by_row, plan = full_new_rows(s)
        t = by_row_t(plan, 2)
        r1 = s.Rc("ReceiptLegacyRow", t, "427", "60", "", "повтор", "", "", "", "", "X", "OLD")
        sh = s.doc.Sheets.getByName("Заказы")
        last = len(orders_of(s)) + 3
        sh.getCellRangeByPosition(0, last, 20, last).setDataArray((("1", "Болт М8х30", "УПД-101", "", "", 1.0, "", 1.0, "шт", "", "", "П", "",
                                                                     d(-1), "", "", "", "", "", "", "A-01"),))
        r2 = s.Rc("ReceiptLegacyRow", last, "427", "1", "", "повтор", "", "", "", "", "X", "OLD")
        R.add(c, "ядро: строка, уже проведённая переносом, повторно не проводится (SKIP); тот же ЕИ в другой строке — отказ «уже есть»",
              r1.startswith("SKIP:") and r2.startswith("ERR:") and "уже есть" in r2, f"{r1}; {r2}")
        s.close()
        s = None
    finally:
        if s is not None:
            s.close()
    cd, st3, w3, res3 = transfer("l28b", full_rows()[:4], steps=("check", "build", "verify"))
    busy = os.path.join(cd, "BUSY")
    os.makedirs(busy)
    open(os.path.join(busy, "заметки.txt"), "w").write("x")
    rc3, out3, _ = engine("promote", w3, st3, extra=("--to", busy))
    exist = os.path.join(cd, "EXIST")
    os.makedirs(exist)
    shutil.copy2(candidate(), os.path.join(exist, "WMS_PROD.ods"))
    h = sha(os.path.join(exist, "WMS_PROD.ods"))
    rc4, out4, _ = engine("promote", w3, st3, extra=("--to", exist))
    rc5, out5, _ = engine("promote", w3, st3, extra=("--to", os.path.join(cd, "NEW")))
    R.add(c, "рабочая WMS не пишется в непустую папку и поверх существующей WMS_PROD.ods (файл не тронут); в новую папку — создана",
          rc3 == 1 and "не пуста" in out3 and os.listdir(busy) == ["заметки.txt"] and rc4 == 1 and "уже существует" in out4
          and sha(os.path.join(exist, "WMS_PROD.ods")) == h and rc5 == 0 and os.path.isfile(os.path.join(cd, "NEW", "WMS_PROD.ods")),
          f"{out3[-120:]}; {out4[-120:]}; {out5[-120:]}")


# ================================================================ L29 volume: 5 000 order rows, 50 000 rows of staging

@case
def l29_volume():
    c = "L29"
    orders, n_ei, issues = [], 0, [["ЕИ", "Количество", "Дата", "Кому"]]
    for k in range(1000):
        for p in range(1, 6):
            i = k * 5 + p
            if p == 1 and k < 500:
                n_ei += 1
                n = 10000 + k
                orders.append(row(A=float(p), B=f"Товар {i}", C=f"УПД-{k}", E=f"ART-{i}", F=10.0, G=10.0, H=10.0, I="шт", J=1.5, L=f"Поставщик {k % 37}",
                                  N=d(-60 + k % 50), O=d(-60 + k % 50), P=d(-70), Q=d(-30), U=f"R-{k % 90}", V=ei(n), W="Получено", X=float(k % 11)))
            else:
                orders.append(row(A=float(p), B=f"Товар {i}", E=f"ART-{i}", H=float(p), I="шт", L=f"Поставщик {k % 37}", P=d(-10),
                                  Q=d((i % 40) - 20), W="Ожидается"))
    for j in range(45000):
        issues.append([ei(20000 + j % 3000), 1.0, d(-100), "история"])
    cd, st, work, res = transfer("l29", orders, issues=issues, steps=("check", "build", "verify"))
    m = summary(work)
    tc, tb, tv = res["check"][2], res.get("build", (0, "", 0))[2], res.get("verify", (0, "", 0))[2]
    R.add(c, "5 000 строк заказов (1 000 заказов, 500 ЕИ) и 45 000 строк выдач (1C) — 50 000 строк staging: проверка, тестовая WMS, сверка",
          all(v[0] == 0 for v in res.values()) and m.get("rows_total") == "5000" and m.get("orders") == "1000" and m.get("eis") == "500",
          f"проверка {tc:.0f} с, создание {tb:.0f} с, сверка {tv:.0f} с; {[v[1][-120:] for v in res.values()]}", timing=dict(check=tc, build=tb, verify=tv))
    R.add(c, "время: проверка ≤ 120 с, создание тестовой WMS ≤ 20 мин, сверка ≤ 5 мин", tc <= 120 and tb <= 1200 and tv <= 300,
          f"{tc:.0f} / {tb:.0f} / {tv:.0f} с")


# ================================================================ L30 crash / rollback

@case
def l30_crash():
    c = "L30"
    rows = full_rows()
    cd = cdir("l30")
    st = staging(os.path.join(cd, "staging.ods"), rows)
    work = os.path.join(cd, "LEGACY_WORK")
    h0 = sha(candidate())
    rc0, out0, _ = engine("build", work, st, candidate())
    th = state(work)["build"]["test_hash"] if rc0 == 0 else ""
    rc, out, _ = engine("build", work, st, candidate(), extra=("--kill-at", "6"))
    s_ = state(work)
    R.add(c, "LibreOffice погибает посреди создания тестовой WMS (6-я операция): создание остановлено, незаконченная копия удалена целиком, "
             "книга-кандидат не изменена, прежняя тестовая WMS не тронута; «СДЕЛАТЬ РАБОЧЕЙ» недоступно",
          rc0 == 0 and rc != 0 and not os.path.exists(os.path.join(work, "test.part")) and sha(candidate()) == h0
          and os.path.isfile(test_book(work)) and sha(test_book(work)) == th and not s_["build"]["ok"]
          and "promote_ready=0" in open(os.path.join(work, "state.txt"), encoding="utf-8").read(), f"{rc0}; {rc}: {out[-300:]}")
    rc2, out2, _ = engine("verify", work, st)
    R.add(c, "после сбоя сверка отказывает — тестовую WMS нужно создать заново", rc2 == 1 and "СОЗДАТЬ ТЕСТОВУЮ WMS" in out2, out2[-200:])
    # an operation fails among its writes (the fault seam of a test candidate): the rollback of WMS, then the whole copy goes
    tcand = os.path.join(cd, "TEST_CANDIDATE.ods")
    o = Office(unique("ltbt"), PROFILES)
    try:
        from build_ods import build
        build(o, tcand, test=True)
    finally:
        o.terminate()
    rows2 = [row(A=1.0, B="Болт", C="УПД", F=5.0, H=5.0, I="шт", L="П", N=d(-3), O=d(-3), U="A", V=ei(1000 + i), X=5.0) for i in range(5)]
    st2 = staging(os.path.join(cd, "staging2.ods"), rows2)
    w2 = os.path.join(cd, "WORK2")
    rc3, out3, _ = engine("build", w2, st2, tcand, extra=("--allow-test-candidate", "--fault-at", "3"))
    R.add(c, "операция переноса откатывается внутри WMS (сбой записи, 3-я операция) — создание остановлено с причиной, копия удалена, "
             "в рабочей папке нет ни test, ни test.part", rc3 == 1 and "ERR-RB" in out3 and not os.path.exists(os.path.join(w2, "test"))
          and not os.path.exists(os.path.join(w2, "test.part")), out3[-300:])
    rc4, out4, _ = engine("build", w2, st2, tcand, extra=("--allow-test-candidate",))
    R.add(c, "повтор без сбоя — тестовая WMS создана (самопроверка, оракул, сверка)", rc4 == 0, out4[-200:])


# ================================================================ L31 the old file is never changed

@case
def l31_source_hash():
    c = "L31"
    cd = cdir("l31")
    old = os.path.join(cd, "СТАРАЯ_ТАБЛИЦА.ods")
    o = maker()
    doc = o.new_calc()
    doc.Sheets.getByIndex(0).Name = "Заказы"
    data = [HDR] + full_rows()[:10]
    doc.Sheets.getByIndex(0).getCellRangeByPosition(0, 0, 27, len(data) - 1).setDataArray(tuple(tuple(r) for r in data))
    doc.storeToURL(uno.systemPathToFileUrl(old), props(FilterName="calc8"))
    doc.close(True)
    h0 = sha(old)
    fe = frontend(cd)
    try:
        fe.copy_from(old)
        fe.setting(4, old)
        for b in ("BtnLtCheck", "BtnLtBuild", "BtnLtVerify"):
            fe.press(b)
        fe.B("LtTestConfirm", 1)
        fe.press("BtnLtPromote")
        msg = fe.msg()
    finally:
        fe.close()
    cut = open(os.path.join(cd, "WMS_WORK", "CUTOVER_README.txt"), encoding="utf-8").read() if os.path.exists(
        os.path.join(cd, "WMS_WORK", "CUTOVER_README.txt")) else ""
    R.add(c, "старая таблица (лист «Заказы» скопирован в книгу переноса) после всех пяти шагов не изменилась: SHA-256 тот же; "
             "CUTOVER_README называет файл и этот SHA-256, правило перехода и время",
          msg.startswith("OK:") and sha(old) == h0 and h0 in cut and "все новые движения записываются только в новую WMS" in cut
          and "read-only архивом" in cut, f"{msg[:120]}; {h0[:12]}")


@case
def l31b_source_changed():
    c = "L31"
    cd = cdir("l31b")
    old = os.path.join(cd, "old.ods")
    o = maker()
    doc = o.new_calc()
    doc.Sheets.getByIndex(0).Name = "Заказы"
    data = [HDR] + full_rows()[:4]
    doc.Sheets.getByIndex(0).getCellRangeByPosition(0, 0, 27, len(data) - 1).setDataArray(tuple(tuple(r) for r in data))
    doc.storeToURL(uno.systemPathToFileUrl(old), props(FilterName="calc8"))
    doc.close(True)
    fe = frontend(cd)
    try:
        fe.copy_from(old)
        fe.setting(4, old)
        for b in ("BtnLtCheck", "BtnLtBuild", "BtnLtVerify"):
            fe.press(b)
        # the storekeeper records one more movement in the old table after it was pasted
        doc = o.load(old, macros=0, hidden=True)
        doc.Sheets.getByName("Заказы").getCellByPosition(23, 1).setValue(59.0)
        doc.store()
        doc.close(True)
        fe.B("LtTestConfirm", 1)
        fe.press("BtnLtPromote")
        msg = fe.msg()
    finally:
        fe.close()
    R.add(c, "старая таблица изменилась после проверки (в ней записали ещё движение) — «СДЕЛАТЬ РАБОЧЕЙ» отказывает: данные нужно вставить заново",
          msg.startswith("ERR:") and "изменился после проверки" in msg and not os.path.exists(os.path.join(cd, "WMS_WORK", "WMS_PROD.ods")), msg[:300])


# ================================================================ L32 the final reconciliation

@case
def l32_reconciliation():
    c = "L32"
    stock = [["ЕИ", "Наименование", "Остаток", "Место"], ["ЕИ-00000330", "Муфта", 7.0, "K-2"]]
    issues = [["ЕИ", "Количество"], ["331", 4.0], ["331", 1.0]]
    returns = [["ЕИ", "Количество"], ["331", 2.0]]
    rows = full_rows() + [row(A=1.0, B="Муфта", C="УПД-701", F=10.0, H=10.0, I="шт", L="Трубы", N=d(-9), O=d(-9), U="K-1", V="330", W="Получено"),
                          row(A=2.0, B="Отвод", C="УПД-701", F=10.0, H=10.0, I="шт", L="Трубы", N=d(-9), O=d(-9), U="K-1", V="331", W="Получено")]
    cd, st, work, res = transfer("l32", rows, stock, issues, returns, steps=("check", "build", "verify"))
    vs = {ln.split(";")[0]: ln.split(";", 1)[1] for ln in open(os.path.join(work, "verify_summary.csv"), encoding="utf-8").read().splitlines()[1:]}
    rep = open(os.path.join(work, "LEGACY_TRANSFER_REPORT.html"), encoding="utf-8").read()
    R.add(c, "итоговая сверка: критических расхождений 0, потерянных строк 0; отчёт .html и .md с категориями «Полностью совпало», "
             "«Нормализовано», «Пересчитано новым движком», «Требует внимания», «Ошибка»",
          all(v[0] == 0 for v in res.values()) and vs.get("Критических расхождений") == "0" and vs.get("Потерянных строк") == "0"
          and all(k in rep for k in ("Полностью совпало", "Нормализовано", "Пересчитано новым движком", "Требует внимания", "Ошибка"))
          and os.path.exists(os.path.join(work, "LEGACY_TRANSFER_REPORT.md")), f"{vs}; {[v[1][-100:] for v in res.values()]}")
    s = ready(Session(test_book(work)))
    try:
        R.add(c, "остаток из листа «Наличие» (1B): ЕИ-00000330 — 7 (X пуст); из выдач и возвратов (1C, 1D): ЕИ-00000331 — 10 − 5 + 2 = 7",
              s.stock(330) == 7.0 and s.stock(331) == 7.0 and s.stock_row(330)[5] == "K-2", f"{s.stock_row(330)}; {s.stock_row(331)}")
    finally:
        s.close()
    # the reconciliation itself sees a difference: a value of the book changed in memory (never stored)
    sys.path.insert(0, os.path.join(release(), "tools"))
    import legacy_transfer as lt
    plan = json.load(open(os.path.join(work, "test", "plan.json"), encoding="utf-8"))
    o = Office(unique("ltrc"), PROFILES)
    try:
        doc = o.load(test_book(work), macros=0, hidden=True)
        clean = lt.reconcile(doc, plan)["critical"]
        for n in ("Наличие", "Заказы"):
            doc.Sheets.getByName(n).unprotect("wms")
        doc.Sheets.getByName("Наличие").getCellByPosition(4, 427).setValue(61.0)
        doc.Sheets.getByName("Заказы").getCellByPosition(1, 3).setString("Подменённое имя")
        bad = lt.reconcile(doc, plan)["critical"]
        doc.setModified(False)
        doc.close(True)
    finally:
        o.terminate()
    R.add(c, "контроль сверки: изменённый в памяти остаток ЕИ-00000427 и наименование строки найдены как критические расхождения",
          not clean and any("ЕИ-00000427" in x for x in bad) and any("Подменённое имя" in x for x in bad), f"{clean[:2]}; {bad[:3]}")


# ================================================================ the frontend book (headless: the macros of its buttons)

class Frontend:
    """WMS_LEGACY_TRANSFER.ods of the release copied into a case folder (with tools/ next to it), open with its macros"""

    def __init__(self, folder):
        rel = release()
        self.folder = folder
        for f in ("WMS_LEGACY_TRANSFER.ods", "WMS_PROD_CANDIDATE.ods", "RELEASE_MANIFEST.csv"):
            shutil.copy2(os.path.join(rel, f), folder)
        shutil.copytree(os.path.join(rel, "tools"), os.path.join(folder, "tools"))
        shutil.copytree(os.path.join(rel, "WMS_TOOLBOX"), os.path.join(folder, "WMS_TOOLBOX"))
        self.o = Office(unique("ltfe"), PROFILES)
        self.doc = self.o.load(os.path.join(folder, "WMS_LEGACY_TRANSFER.ods"), macros=4)
        self.B("LtTestAuto", True)

    def B(self, f, *a):
        return self.o.basic(self.doc, "LtMain", f, *a)

    def press(self, macro):
        t0 = time.time()
        self.B(macro)
        return time.time() - t0

    def msg(self):
        return str(self.B("LtTestLastMessage") or "")

    def paste(self, rows, header=True, formulas=None):
        data = ([HDR] if header else []) + rows
        sh = self.doc.Sheets.getByName("1_Вставить_Заказы")
        sh.getCellRangeByPosition(0, 0, 27, len(data) - 1).setDataArray(tuple(tuple(r) for r in data))
        for (sname, col, r), f in (formulas or {}).items():
            self.doc.Sheets.getByName(sname).getCellByPosition(LET.index(col), r).setFormula(f)

    def copy_from(self, path):
        """the old «Заказы» A:AB as values, as a copy of the whole sheet pastes them"""
        o = self.o
        old = o.load(path, macros=0, hidden=True, ReadOnly=True)
        sh = old.Sheets.getByName("Заказы")
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        data = sh.getCellRangeByPosition(0, 0, 27, cur.getRangeAddress().EndRow).getDataArray()
        old.close(True)
        self.doc.Sheets.getByName("1_Вставить_Заказы").getCellRangeByPosition(0, 0, 27, len(data) - 1).setDataArray(data)

    def setting(self, r, v):
        self.doc.Sheets.getByName("Настройки").getCellByPosition(1, r).setString(v)

    def formula_count(self):
        n = 0
        for name in ("1_Вставить_Заказы", "1B_Вставить_Наличие", "1C_Вставить_Выдачи", "1D_Вставить_Возвраты"):
            for a in self.doc.Sheets.getByName(name).queryContentCells(16).getRangeAddresses():
                n += (a.EndRow - a.StartRow + 1) * (a.EndColumn - a.StartColumn + 1)
        return n

    def cell(self, sheet, col, r):
        cl = self.doc.Sheets.getByName(sheet).getCellByPosition(LET.index(col), r)
        return cl.getType().value, cl.getValue() if cl.getType().value == "VALUE" else cl.getString()

    def main(self):
        m = self.doc.Sheets.getByName("LEGACY TRANSFER")
        return [m.getCellByPosition(2, r).getString() for r in range(3, 10)]

    def promote_enabled(self):
        return self.doc.Sheets.getByName("LEGACY TRANSFER").getDrawPage().getForms().getByIndex(0).getByName("btnPromote").Enabled

    def close(self):
        try:
            self.doc.setModified(False)
            self.doc.close(True)
        except Exception:
            pass
        self.o.terminate()


def frontend(folder):
    return Frontend(folder)


@case
def l33_frontend():
    c = "L33"
    cd = cdir("l33")
    fe = frontend(cd)
    try:
        bad_rows = full_rows() + [row(A=1.0, B="Без количеств", I="шт", L="П")]
        fe.paste(bad_rows)
        fe.press("BtnLtCheck")
        m1 = fe.msg()
        ad = fe.doc.Sheets.getByName("1_Вставить_Заказы").getCellRangeByPosition(29, 0, 30, len(bad_rows) + 1).getDataArray()
        chk = {r[0]: (r[1], r[2]) for r in fe.doc.Sheets.getByName("2_Проверка").getCellRangeByPosition(0, 2, 2, 80).getDataArray() if r[0]}
        errs = [r for r in fe.doc.Sheets.getByName("3_Ошибки").getCellRangeByPosition(0, 1, 5, 40).getDataArray() if r[0]]
        en1 = fe.promote_enabled()
        R.add(c, "«ПРОВЕРИТЬ» с красной строкой: сообщение о блокирующих замечаниях, строка отмечена «БЛОКЕР» в AD с причиной в AE, "
                 "«2_Проверка» — итог «ПЕРЕНОС НЕВОЗМОЖЕН», «3_Ошибки» — замечание; «СДЕЛАТЬ РАБОЧЕЙ» недоступна",
              m1.startswith("ERR:") and ad[len(bad_rows)][0] == "БЛОКЕР" and "ни фактического" in ad[len(bad_rows)][1]
              and chk.get("ИТОГ", ("",))[0].startswith("ПЕРЕНОС НЕВОЗМОЖЕН") and chk["ИТОГ"][1] == "БЛОКЕР" and errs and errs[0][3] == "БЛОКЕР"
              and not en1, f"{m1[:100]}; {ad[len(bad_rows)]}; {chk.get('ИТОГ')}; {errs[:1]}")
        fe.press("BtnLtBuild")
        m2 = fe.msg()
        R.add(c, "«СОЗДАТЬ ТЕСТОВУЮ WMS» при красной строке — отказ, ничего не создано", m2.startswith("ERR:")
              and not os.path.exists(os.path.join(cd, "LEGACY_WORK", "test")), m2[:200])
        # the user fixes the row in place (deletes it) and goes on
        fe.doc.Sheets.getByName("1_Вставить_Заказы").getRows().removeByIndex(len(bad_rows), 1)
        times = [fe.press(b) for b in ("BtnLtCheck", "BtnLtBuild")]
        m3 = fe.msg()
        en2 = fe.promote_enabled()
        fe.press("BtnLtVerify")
        m4 = fe.msg()
        en3 = fe.promote_enabled()
        main = fe.main()
        ver = {r[0]: r[1] for r in fe.doc.Sheets.getByName("4_Сверка").getCellRangeByPosition(0, 2, 1, 40).getDataArray() if r[0]}
        R.add(c, "исправлено на месте → «ПРОВЕРИТЬ», «СОЗДАТЬ ТЕСТОВУЮ WMS» (кнопка ещё недоступна), «СВЕРИТЬ» — «СДЕЛАТЬ РАБОЧЕЙ» стала "
                 "доступна; главная страница: шаги 2–4 ✓, лист «4_Сверка» — «МОЖНО СДЕЛАТЬ РАБОЧЕЙ»",
              m3.startswith("OK:") and not en2 and m4.startswith("OK:") and en3 and all(x.startswith("✓") for x in main[1:4])
              and ver.get("Итог") == "МОЖНО СДЕЛАТЬ РАБОЧЕЙ", f"{m3[:80]}; {en2}/{en3}; {main[1:5]}; {ver.get('Итог')}; {times}")
        fe.B("LtTestConfirm", 0)
        fe.press("BtnLtPromote")
        m5 = fe.msg()
        R.add(c, "«СДЕЛАТЬ РАБОЧЕЙ» — вопрос «Нет»: ничего не изменено", m5 == "Ничего не изменено." and not os.path.exists(os.path.join(cd, "WMS_WORK")),
              m5)
        fe.B("LtTestConfirm", 1)
        fe.press("BtnLtPromote")
        m6 = fe.msg()
        en4 = fe.promote_enabled()
        wd = os.path.join(cd, "WMS_WORK")
        R.add(c, "«СДЕЛАТЬ РАБОЧЕЙ» — «Да»: WMS_WORK/WMS_PROD.ods с журналом, WMS_TOOLBOX, отчёт переноса, резервная копия выпуска, "
                 "CUTOVER_README.txt; кнопка снова недоступна (перенос сделан)",
              m6.startswith("OK:") and all(os.path.exists(os.path.join(wd, x)) for x in ("WMS_PROD.ods", "WMS_Journal", "WMS_TOOLBOX",
                                                                                         "LEGACY_TRANSFER/LEGACY_TRANSFER_REPORT.html",
                                                                                         "backup/release/WMS_PROD_CANDIDATE.ods", "CUTOVER_README.txt"))
              and not en4, f"{m6[:120]}; {sorted(os.listdir(wd)) if os.path.isdir(wd) else ''}")
        fe.press("BtnLtExport")
        m7 = fe.msg()
        fe.press("BtnLtBackup")
        m8 = fe.msg()
        fe.B("LtTestConfirm", 1)
        fe.press("BtnLtClear")
        left = fe.doc.Sheets.getByName("1_Вставить_Заказы").getCellRangeByPosition(0, 1, 27, 30).getDataArray()
        hdr = fe.doc.Sheets.getByName("1_Вставить_Заказы").getCellByPosition(0, 0).getString()
        main_c = fe.main()
        R.add(c, "«Экспорт отчёта» — отчёты в новой папке; «Создать резервную копию» — копия книги; «Очистить staging» — вставленное "
                 "удалено, заголовок на месте, созданные WMS не тронуты",
              m7.startswith("OK:") and glob.glob(os.path.join(cd, "LEGACY_WORK", "export_*", "LEGACY_TRANSFER_REPORT.html"))
              and m8.startswith("OK:") and glob.glob(os.path.join(cd, "LEGACY_WORK", "backups", "WMS_LEGACY_TRANSFER_*.ods"))
              and not any(any(v for v in r) for r in left) and hdr == "№ позиции" and os.path.isfile(os.path.join(wd, "WMS_PROD.ods"))
              and main_c[4].startswith("✓") and not fe.promote_enabled(), f"{m7[:80]}; {m8[:80]}; {main_c[4][:60]}")
    finally:
        fe.close()


# ================================================================ L34 the version check

@case
def l34_version():
    c = "L34"
    cd = cdir("l34")
    st = staging(os.path.join(cd, "staging.ods"), full_rows()[:3])
    fake = os.path.join(cd, "WMS_SYS5.ods")
    shutil.copy2(candidate(), fake)
    o = maker()
    doc = o.load(fake, macros=0, hidden=True)
    doc.Sheets.getByName("_SYS").unprotect("wms")
    doc.Sheets.getByName("_SYS").getCellByPosition(1, 0).setString("WMS-SYS-5")
    doc.Sheets.getByName("_SYS").protect("wms")
    doc.store()
    doc.close(True)
    rc, out, _ = engine("check", os.path.join(cd, "W1"), st, fake)
    F = [f for f in findings(os.path.join(cd, "W1")) if f["kind"] == "книга-кандидат"]
    R.add(c, "книга-кандидат другой схемы (WMS-SYS-5) — понятный отказ: инструмент работает только со схемой WMS-SYS-4",
          rc == 1 and F and "WMS-SYS-5" in F[0]["text"] and "WMS-SYS-4" in F[0]["text"], str(F[:1]))
    rc2, out2, _ = engine("build", os.path.join(cd, "W1"), st, fake)
    R.add(c, "и создание тестовой WMS с такой книгой — отказ, копии нет", rc2 == 1 and not os.path.exists(os.path.join(cd, "W1", "test")), out2[-160:])
    src = os.path.join(cd, "src070")
    why = git_sources("e65f606", src, parts=("tools", "src", "docs"))
    if why:
        R.add(c, "книга-кандидат выпуска 0.7.0 (без операций переноса) — отказ", "SKIP", why)
        return
    b = subprocess.run([sys.executable, os.path.join(src, "tools", "build_ods.py"), os.path.join(cd, "C070.ods")], capture_output=True, text=True,
                       timeout=900, env=ENV, cwd=src)
    rc3, out3, _ = engine("check", os.path.join(cd, "W2"), st, os.path.join(cd, "C070.ods"))
    F3 = [f for f in findings(os.path.join(cd, "W2")) if f["kind"] == "книга-кандидат"] if rc3 in (0, 1) else []
    R.add(c, "книга-кандидат выпуска 0.7.0 (собрана из исходников e65f606: схема та же, операций переноса нет) — отказ «не умеет переносить»",
          b.returncode == 0 and rc3 == 1 and F3 and "0.7.0" in F3[0]["text"] and "не умеет переносить" in F3[0]["text"], f"{b.stderr[-200:]}; {F3[:1]}")


# ================================================================ L35 the balance and the rules

@case
def l35_rules():
    c = "L35"
    base = dict(C="УПД", I="шт", L="П", N=d(-5), O=d(-5), U="Z")
    rows = [row(A=1.0, B="Нет наличия", F=10.0, H=10.0, V="2001", **base),
            row(A=2.0, B="Наличие больше прихода", F=5.0, H=5.0, V="2002", X=6.0, **base),
            row(A=3.0, B="Наличие текстом", F=5.0, H=5.0, V="2003", X="3,5", **base)]
    cd, st, work, res = transfer("l35", rows, steps=("check",))
    F = findings(work)
    R.add(c, "остаток не определён (нет X, 1B, 1C/1D) — по умолчанию БЛОКЕР «Требуется проверить остаток»; наличие больше прихода — "
             "БЛОКЕР «подозрительный остаток»; «3,5» текстом — принято",
          res["check"][0] == 1 and any(f["row"] == "2" and "Требуется проверить остаток" in f["text"] and f["level"] == "БЛОКЕР" for f in F)
          and any(f["row"] == "3" and "подозрительный" in f["text"] and f["level"] == "БЛОКЕР" for f in F)
          and not any(f["row"] == "4" and f["level"] == "БЛОКЕР" for f in F), str([(f["row"], f["text"][:60]) for f in F])[:500])
    cd2, st2, work2, res2 = transfer("l35b", rows[:1] + rows[2:], steps=("check", "build"), unknown_balance="zero")
    s = ready(Session(test_book(work2)))
    try:
        r1 = orders_of(s, 2)[0]
        R.add(c, "правило «Остаток, если не определён = 0»: ЕИ перенесён с остатком 0, жёлтое замечание, у строки пометка "
                 "«[перенос: проверить остаток]»; ЕИ с наличием «3,5» — остаток 3,5",
              res2["build"][0] == 0 and s.stock(2001) == 0.0 and "[перенос: проверить остаток]" in r1[25] and s.stock(2003) == 3.5
              and any(f["level"] == "ПРОВЕРИТЬ" and "Требуется проверить остаток" in f["text"] for f in findings(work2)), f"{r1[25]}; {s.stock(2003)}")
    finally:
        s.close()
    rows3 = [row(A=1.0, B="Без единицы и места", F=2.0, H=2.0, L="П", N=d(-2), V="2101", X=2.0),
             row(A=2.0, B="Без даты прихода", C="УПД-5", F=2.0, H=2.0, I="Шт.", L="П", O=d(-3), U="Y", V="2102", X=2.0, J=12.3456),
             row(A=3.0, B="Поступление без позиции", F=1.0, I="шт", N=d(-1), U="Y", V="2103", X=1.0),
             row(A=4.0, B="Пустая строка без количеств", I="шт", L="П")]
    cd3, st3, work3, res3 = transfer("l35c", rows3, steps=("check",))
    F3 = findings(work3)
    kinds = sorted({f["kind"] for f in F3 if f["level"] == "БЛОКЕР"})
    R.add(c, "без правил — БЛОКЕРЫ: нет единицы, нет места, нет даты поступления, цена с 4 знаками, поступление без позиции, строка без "
             "количеств", res3["check"][0] == 1 and sum(1 for f in F3 if f["level"] == "БЛОКЕР") >= 6, str(kinds))
    cd4, st4, work4, res4 = transfer("l35d", rows3, steps=("check", "build"), default_unit="шт", default_place="Склад", no_receipt_date="doc",
                                     price_round="round", no_order_qty="fact", empty_rows="skip")
    s = ready(Session(test_book(work4))) if res4.get("build", (1,))[0] == 0 else None
    try:
        rr = orders_of(s, 4) if s else []
        R.add(c, "с явными правилами (единица «шт», место «Склад», дата поступления = дата документа, округление цены, строка с F без H — "
                 "позиция H = F, строки без количеств пропускаются): перенос выполнен, «Шт.» → «шт», каждая замена — в проверке как нормализация",
              res4.get("build", (1,))[0] == 0 and rr[0][8] == "шт" and rr[0][20] == "Склад" and rr[1][13] == d(-3) and rr[1][9] == 12.346
              and rr[1][8] == "шт" and rr[2][7] == 1.0 and rr[2][21] == ei(2103) and len([r for r in rr if r[1]]) == 3
              and "(правило)" in open(os.path.join(work4, "check", "rows.csv"), encoding="utf-8").read(),
              f"{res4.get('build', res4.get('check'))[1][-200:]}; {[r[:10] for r in rr[:3]]}")
    finally:
        if s:
            s.close()


# ================================================================ L36 the structure of the paste; the stock sheet only

@case
def l36_structure():
    c = "L36"
    rows = full_rows()[:4]
    shifted = [[""] + r[:27] for r in [HDR] + rows]
    cd = cdir("l36")
    st = staging(os.path.join(cd, "staging.ods"), shifted, header=False)
    rc, out, _ = engine("check", os.path.join(cd, "W"), st, candidate())
    F = findings(os.path.join(cd, "W"))
    R.add(c, "вставка со сдвигом на одну колонку — БЛОКЕР с понятной причиной (заголовок не на своём месте), ничего не переносится",
          rc == 1 and any("сдвигом" in f["text"] for f in F), str([f["text"][:80] for f in F][:3]))
    stock = [["Внутренний код", "Наименование", "Единица", "Остаток", "Место"], ["ЕИ-00000950", "Ведро", "шт", 3.0, "S-1"],
             ["ЕИ-00000427", "Болт М8х30", "шт", 60.0, "A-01"]]
    cd2, st2, work2, res2 = transfer("l36b", full_rows()[:2], stock, steps=("check",))
    F2 = [f for f in findings(work2) if f["sheet"] == "1B_Вставить_Наличие"]
    R.add(c, "ЕИ с остатком только в листе «Наличие» (1B) — по умолчанию БЛОКЕР (остаток потерялся бы)",
          res2["check"][0] == 1 and any("ЕИ-00000950" in f["text"] and f["level"] == "БЛОКЕР" for f in F2), str(F2[:2]))
    cd3, st3, work3, res3 = transfer("l36c", full_rows()[:2], stock, steps=("check", "build"), stock_only="old")
    s = ready(Session(test_book(work3))) if res3.get("build", (1,))[0] == 0 else None
    try:
        R.add(c, "правило «СТАРЫЙ СКЛАД»: ЕИ-00000950 перенесён с остатком 3 и типом «Старый склад»",
              s is not None and s.stock(950) == 3.0 and s.doc.Sheets.getByName("Наличие").getCellByPosition(9, 950).getString() == "Старый склад",
              str(res3.get("build", res3["check"])[1][-200:]))
    finally:
        if s:
            s.close()
    # a position without F whose first receipt is the next row: the two rows become one receipt — two different EIs on them
    # are a blocker (neither is dropped silently); the same EI written two ways is one EI; two balances — the receipt's, with a
    # warning
    two = [row(A=1.0, B="Позиция с ЕИ", E="ART-M1", H=10.0, I="шт", L="П", V="3201"),
           row(A=1.0, B="Позиция с ЕИ", E="ART-M1", F=10.0, I="шт", L="П", N=d(-5), U="M-1", V="3202", X=10.0)]
    cd4, st4, work4, res4 = transfer("l36d", two, steps=("check",))
    F4 = findings(work4)
    R.add(c, "у позиции без F свой ЕИ, у её первого поступления другой — БЛОКЕР с обоими номерами (ни один ЕИ не теряется молча)",
          res4["check"][0] == 1 and any(f["level"] == "БЛОКЕР" and "3201" in f["text"] and "3202" in f["text"] for f in F4),
          str([(f["level"], f["text"][:100]) for f in F4][:3]))
    same = [row(A=1.0, B="Тот же ЕИ иначе", E="ART-M2", H=4.0, I="шт", L="П", V="ЕИ-3203", X=9.0),
            row(A=1.0, B="Тот же ЕИ иначе", E="ART-M2", F=4.0, I="шт", L="П", N=d(-5), U="M-2", V="3203", X=4.0)]
    cd5, st5, work5, res5 = transfer("l36e", same, steps=("check", "build"))
    F5 = findings(work5)
    s = ready(Session(test_book(work5))) if res5.get("build", (1,))[0] == 0 else None
    try:
        R.add(c, "тот же ЕИ у позиции и её первого поступления в разной записи — один ЕИ-00003203; разное наличие X — берётся наличие "
                 "поступления (4) и жёлтое предупреждение",
              s is not None and s.stock(3203) == 4.0 and orders_of(s, 1)[0][21] == ei(3203)
              and any(f["level"] == "ПРОВЕРИТЬ" and "наличие «9»" in f["text"] for f in F5),
              f"{res5.get('build', res5['check'])[1][-160:]}; {[(f['level'], f['text'][:90]) for f in F5][:3]}")
    finally:
        if s:
            s.close()
    # the table pasted twice (its header again, the same rows below): nothing is transferred twice silently
    dup = [row(A=1.0, B="Приход без ЕИ", C="УПД-7", F=3.0, H=3.0, I="шт", L="П", N=d(-4), O=d(-4), U="D-1", X=3.0),
           row(A=2.0, B="Ожидается", H=2.0, I="шт", L="П", Q=d(5))]
    cd6, st6, work6, res6 = transfer("l36f", dup + [HDR] + dup, steps=("check",))
    F6 = [f for f in findings(work6) if f["kind"] == "дубль строки"]
    R.add(c, "таблица вставлена дважды: повтор заголовка и строк найден — повтор прихода (F) БЛОКЕР (остаток удвоился бы), повтор "
             "ожидаемой позиции и заголовка — жёлтые; «2_Проверка»: повторяющиеся строки 3",
          res6["check"][0] == 1 and len(F6) == 3 and summary(work6).get("dup_rows") == "3"
          and any(f["row"] == "4" and f["level"] == "ПРОВЕРИТЬ" and "заголовка" in f["text"] for f in F6)
          and any(f["row"] == "5" and f["level"] == "БЛОКЕР" and "строкой 2" in f["text"] for f in F6)
          and any(f["row"] == "6" and f["level"] == "ПРОВЕРИТЬ" and "строкой 3" in f["text"] for f in F6),
          str([(f["row"], f["level"], f["text"][:70]) for f in F6]))
    # rule «do not transfer» for an EI of the stock sheet only: its number stays behind above NEXT_EI — said, not hidden
    cd7, st7, work7, res7 = transfer("l36g", full_rows()[:2], stock, steps=("check",), stock_only="skip")
    F7 = findings(work7)
    R.add(c, "правило «не переносить» ЕИ только из «Наличие»: ЕИ-00000950 не переносится (жёлтое), NEXT_EI 428 ниже его номера — "
             "предупреждение, что новый ЕИ может получить номер старого",
          res7["check"][0] == 0 and summary(work7).get("next_ei") == "428"
          and any(f["level"] == "ПРОВЕРИТЬ" and "ЕИ-00000950" in f["text"] and "NEXT_EI будет 428" in f["text"] for f in F7),
          str([(f["level"], f["text"][:110]) for f in F7][:4]))


# ================================================================ L37 Doctor, Search, Analytics on transferred EIs

@case
def l37_toolbox():
    c = "L37"
    c_, st, work, res = full_case()
    wd = os.path.join(c_, "WMS_WORK")
    s = ready(Session(os.path.join(wd, "WMS_PROD.ods")))
    try:
        ex = s.B("ExportSnapshot", module="WmsExport")
    finally:
        s.close()
    from run_toolbox import Tool
    res_d, res_s, res_a, found = "", "", "", []
    for name in ("WMS_DOCTOR", "WMS_SEARCH", "WMS_ANALYTICS"):
        tool = Tool(wd, name)
        try:
            if name == "WMS_DOCTOR":
                res_d = str(tool.B("DoctorRun", module="TbDoctor"))
            elif name == "WMS_SEARCH":
                tool.B("SearchRefresh", module="TbSearch")
                res_s = str(tool.B("SearchRun", "ЕИ-00000427", module="TbSearch"))
                found = [r for r in tool.rows("Поиск", 3, 9) if r[0]]
            else:
                res_a = str(tool.B("AnRefresh", float(serial(T)), module="TbAnalytics"))
        finally:
            tool.close()
    R.add(c, "рабочая WMS после переноса: «Экспорт для инструментов», WMS_DOCTOR — ошибок 0, WMS_SEARCH находит ЕИ-00000427, "
             "WMS_ANALYTICS строит сводку",
          ex.startswith("OK:") and res_d.startswith("OK:ошибок 0") and res_s.startswith("OK") and any(r[1] == ei(427) for r in found)
          and res_a.startswith("OK"), f"{ex[:80]}; {res_d[:160]}; {res_s[:80]}; {res_a[:80]}")


# ================================================================ L38 special receipts; the rule of the parts after the transfer

@case
def l38_special():
    c = "L38"
    rows = [row(A=1.0, B="Стол", F=1.0, I="шт", L="Производство", N=d(-20), U="P-1", V="3001", X=1.0),
            row(A=2.0, B="Шкаф", F=2.0, I="шт", L="Старый склад", N=d(-20), U="P-2", V="3002", X=2.0),
            row(A=3.0, B="Неизвестный ящик", F=1.0, I="шт", L="Иной", N=d(-20), U="P-3", V="3003", X=1.0),
            row(A=4.0, B="Прокладка без артикула", F=5.0, I="шт", L="Детали", N=d(-20), U="P-4", V="3004", X=4.0)]
    bad = [row(A=1.0, B="Прокладка", E="PN-7", F=2.0, I="шт", L="Детали", N=d(-9), U="P-5", V="3301", X=2.0),
           row(A=2.0, B="Прокладка", E="PN-7", F=1.0, I="шт", L="Детали", N=d(-8), U="P-5", V="abc", X=2.0)]
    cdb, stb, workb, resb = transfer("l38b", bad, steps=("check",))
    R.add(c, "два прихода одной детали (один ЕИ): неверный номер ЕИ во второй строке — БЛОКЕР этой строки (не пропадает молча)",
          resb["check"][0] == 1 and any(f["row"] == "3" and f["level"] == "БЛОКЕР" and "«abc»" in f["text"] for f in findings(workb)),
          str([(f["row"], f["level"], f["text"][:80]) for f in findings(workb)][:4]))
    cd, st, work, res = transfer("l38", rows, steps=("check", "build"))
    s = ready(Session(test_book(work))) if res.get("build", (1,))[0] == 0 else None
    try:
        card = {n: s.doc.Sheets.getByName("Наличие").getCellRangeByPosition(0, n, 9, n).getDataArray()[0] for n in (3001, 3002, 3003, 3004)} if s else {}
        R.add(c, "специальные приходы старой таблицы (Производство, Старый склад, Иной, деталь без артикула) — ЕИ с остатком и своим типом "
                 "источника; «Иной» и деталь без артикула — «Требует разбора»; строк в «Заказы» не появляется",
              s is not None and card[3001][9] == "Производство" and card[3002][9] == "Старый склад" and card[3003][9] == "Иной приход"
              and card[3003][7] == "Требует разбора" and card[3004][7] == "Требует разбора" and card[3004][4] == 4.0
              and not any(orders_of(s, 3)[0][:21]) and any("деталь без артикула" in f["text"] for f in findings(work)),
              f"{[(v[0], v[4], v[7], v[9]) for v in card.values()]}; {res.get('build', res['check'])[1][-160:]}")
    finally:
        if s:
            s.close()
    s = prod_session()
    try:
        b0 = s.stock(700)
        s.special_input(1, B="Детали", D="Деталь насоса", E="pn-100", F="2", G="шт", H=TODAY, I="E-01", K="сервис")
        r = s.click_spc("BtnSpcPost", 1)
        R.add(c, "после переноса новая WMS сама применяет правило деталей: приход детали с артикулом «pn-100» пополняет перенесённый "
                 "ЕИ-00000700 (не создаёт второй ЕИ)", r.startswith("OK") and s.stock(700) == b0 + 2.0 and s.spc(1)[13] == ei(700),
              f"{r}; {b0} → {s.stock(700)}; {s.spc(1)[13]}")
        oracle(c, s, "перенесённая деталь пополнена обычным приходом")
    finally:
        s.close()


# ================================================================ L39 how people fill a table: placeholders, units, words instead of dates

@case
def l39_real_life():
    c = "L39"
    rows = [row(A=1.0, B="Ожидаемая с прочерками", H=5.0, I="шт", L="П", F="-", V="—", X="нет", Q="на след. неделе", P="?"),
            row(A=2.0, B="Полученная с единицами", C="УПД-1", F="10 шт", H="10 шт.", I="шт", L="П", N=d(-3), O=d(-3), U="W-1", V="3101",
                X="7 шт"),
            row(A=3.0, B="Полученная, место прочерком в ЕИ", C="УПД-2", F=4.0, H=4.0, I="кг", L="П", N=d(-3), O=d(-3), U="W-2", V="3102",
                X="4 кг", AA="-")]
    cd, st, work, res = transfer("l39", rows, steps=("check", "build"))
    F = findings(work)
    rr = row_results(work)
    s = ready(Session(test_book(work))) if res.get("build", (1,))[0] == 0 else None
    try:
        new = orders_of(s, 4) if s else []
        R.add(c, "прочерки («-», «—», «нет») в числах, датах и ЕИ — пусто; количества с единицей строки («10 шт», «7 шт») — числа; "
                 "не даты в дате заказа и ожидаемой дате — в комментарий (жёлтое), перенос не останавливают",
              res["check"][0] == 0 and s is not None and new[0][5] == "" and new[0][21] == "" and new[0][15] == "" and new[0][16] == ""
              and "на след. неделе" in new[0][25] and new[1][5] == 10.0 and new[1][7] == 10.0 and s.stock(3101) == 7.0 and s.stock(3102) == 4.0
              and new[2][26] == "" and any(f["level"] == "ПРОВЕРИТЬ" and "не дата" in f["text"] for f in F) and "считается пустым" in rr[2][1],
              f"{res['check'][1][-160:]}; {[r[:8] + r[15:17] + r[21:26] for r in new[:3]]}; {rr.get(2)}")
    finally:
        if s:
            s.close()



# ================================================================ L40 old statuses as people wrote them (0.7.2)

@case
def l40_statuses():
    c = "L40"
    rows = [row(A=1.0, B="Размещен, срок впереди", H=5.0, I="шт", L="П", Q=d(5), W="Размещен"),                                   # 2
            row(A=2.0, B="Размещён, срок прошёл", H=5.0, I="шт", L="П", Q=d(-3), W="Размещён"),                                  # 3
            row(A=3.0, B="размещено, получено частично", C="УПД-40", F=2.0, H=5.0, I="шт", L="П", N=d(-4), O=d(-4), Q=d(9), U="S-1",
                V="4001", X=2.0, W="размещено"),                                                                                # 4
            row(A=4.0, B="ЗАКАЗ РАЗМЕЩЁН, получено", C="УПД-41", F=3.0, H=3.0, I="шт", L="П", N=d(-4), O=d(-4), U="S-1", V="4002", X=3.0,
                W="ЗАКАЗ РАЗМЕЩЁН"),                                                                                            # 5
            row(A=5.0, B="Неизвестный статус", H=3.0, I="шт", L="П", Q=d(10), W="Согласуется с бухгалтерией"),                   # 6
            row(A=6.0, B="Отменён", H=2.0, I="шт", L="П", W="Отменён"),                                                          # 7
            row(A=1.0, B="Оприходовано", C="УПД-42", F=4.0, H=4.0, I="шт", L="П", N=d(-2), O=d(-2), U="S-2", V="4003", X=4.0,
                W="Оприходовано"),                                                                                              # 8
            row(A=2.0, B="В пути", H=1.0, I="шт", L="П", Q=d(4), W="  в пути  "),                                                # 9
            row(A=3.0, B="Частично получен", C="УПД-43", F=1.0, H=2.0, I="шт", L="П", N=d(-1), O=d(-1), Q=d(3), U="S-2", V="4004",
                X=1.0, W="Частично получен")]                                                                                  # 10
    cd, st, work, res = transfer("l40", rows, steps=("check", "build", "verify"))
    F = findings(work)
    m = summary(work)
    s = ready(Session(test_book(work))) if res.get("build", (1,))[0] == 0 else None
    try:
        new = [r[22] for r in orders_of(s, 10)[:9]] if s else []
        want = ["Ожидается", "Ожидается", "Частично получено", "Получено", "Ожидается", "Отменено", "Получено", "Ожидается",
                "Частично получено"]
        R.add(c, "«Размещен», «Размещён», «размещено», «ЗАКАЗ РАЗМЕЩЁН», «в пути», «Оприходовано», «Отменён», «Частично получен» "
                 "(регистр, ё/е, пробелы, «заказ …») понятны и не блокируют; статус новой WMS — по данным: ожидается (и с прошедшей ожидаемой "
                 "датой — просрочки как статуса нет), частично, получено, отменено; тестовая WMS создана и сверена",
              all(v[0] == 0 for v in res.values()) and m.get("blockers") == "0" and new == want, f"{new}; {[v[1][-120:] for v in res.values()]}")
        unk = [f for f in F if "неизвестный старый статус" in f["text"]]
        R.add(c, "неизвестный старый статус («Согласуется с бухгалтерией») — жёлтое «неизвестный старый статус — состояние рассчитано по "
                 "данным», перенос не останавливается; «2_Проверка»: неизвестных старых статусов 1",
              len(unk) == 1 and unk[0]["level"] == "ПРОВЕРИТЬ" and unk[0]["row"] == "6" and "состояние рассчитано по данным" in unk[0]["text"]
              and m.get("status_unknown") == "1" and m.get("status_unclear") == "0", str(unk)[:300])
        import csv
        import journal_oracle
        ents, _ = journal_oracle.read_journal(os.path.join(work, "test", "WMS_Journal"))
        leg = {e["fields"].get("EI"): e["fields"].get("LEGACY_STATUS") for e in ents if e["type"] == "LEGACY_RECEIPT"}
        with open(os.path.join(work, "LEGACY_ROWS.csv"), encoding="utf-8") as f:
            lr = {int(r[0]): r for r in list(csv.reader(f, delimiter=";"))[1:]}
        R.add(c, "старый текст статуса сохраняется дословно (LegacyStatus) — в журнале прихода и в карте строк, рядом — рассчитанный статус; "
                 "понятый синоним отмечен как нормализация",
              leg.get(ei(4001)) == "размещено" and leg.get(ei(4002)) == "ЗАКАЗ РАЗМЕЩЁН" and lr[2][6:8] == ["Размещен", "Ожидается"]
              and lr[3][6:8] == ["Размещён", "Ожидается"] and lr[6][6:8] == ["Согласуется с бухгалтерией", "Ожидается"]
              and "понят как «Ожидается»" in lr[2][13],
              f"{leg}; {[lr[k][6:8] for k in (2, 3, 6)]}; {lr[2][13][:80]}")
    finally:
        if s:
            s.close()
    # a status that looks like a cancellation but is not in the table: its meaning changes the order — asked, not guessed; own
    # synonyms of «Настройки» give it
    odd = [row(A=1.0, B="Отменить?", H=2.0, I="шт", L="П", W="Отменить?"),
           row(A=2.0, B="Снят поставщиком", H=2.0, I="шт", L="П", W="Снят поставщиком")]
    cd2, st2, work2, res2 = transfer("l40b", odd, steps=("check",))
    F2 = [f for f in findings(work2) if f["kind"] == "статус"]
    cd3, st3, work3, res3 = transfer("l40c", odd, steps=("check",), status_map="Отменить? = Отменено; снят поставщиком=отменено")
    plan3 = json.load(open(os.path.join(work3, "check", "plan.json"), encoding="utf-8")) if res3["check"][0] == 0 else {}
    R.add(c, "статус, похожий на отмену, но не из словаря («Отменить?», «Снят поставщиком») — БЛОКЕР с подсказкой (смысл отмены не "
             "угадывается); «Свои синонимы статусов» в «Настройках» задают смысл — позиции отменены, блокеров нет",
          res2["check"][0] == 1 and len([f for f in F2 if f["level"] == "БЛОКЕР" and "похоже на отмену" in f["text"]]) == 2
          and summary(work2).get("status_unclear") == "2" and res3["check"][0] == 0
          and [p["new"] for p in plan3.get("positions", [])] == ["Отменено", "Отменено"],
          f"{[(f['row'], f['text'][:70]) for f in F2]}; {res3['check'][1][-150:]}; {[p['new'] for p in plan3.get('positions', [])]}")


# ================================================================ L41 dates as people wrote them (0.7.2)

@case
def l41_dates():
    c = "L41"
    import importlib.util
    spec = importlib.util.spec_from_file_location("lt_release_engine", os.path.join(release(), "tools", "legacy_transfer.py"))
    lt = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(lt)
    sep1 = serial(datetime.date(2026, 9, 1))
    forms = ["1.9.26", "01.09.26", "1.09.2026", "01.09.2026", "1/9/26", "01/09/2026", "1-9-26", "01-09-2026", " 01.09.2026 ",
             "01.09.2026 г.", "01.09.2026 10:30", "2026-09-01", sep1, sep1 + 0.4]
    got = {repr(f): lt.parse_legacy_date(f)[0] for f in forms}
    bad = {repr(f): lt.parse_legacy_date(f)[1] for f in ("31.02.2026", "1.9", "когда-то", "46266", 5.0, "1.9-26", "")}
    R.add(c, "parse_legacy_date: 1.9.26, 01.09.26, 1.09.2026, 01.09.2026, 1/9/26, 01/09/2026, 1-9-26, 01-09-2026, с пробелами, с «г.», "
             "со временем, 2026-09-01, число Calc и дата-время Calc → 01.09.2026 (год 26 — 2026); «31.02.2026», «1.9», слово, «46266», "
             "число 5, смешанные разделители — не дата; дата не принимается за количество или цену",
          all(v == sep1 for v in got.values()) and all(bad.values()) and lt.qty_value("1.9.26")[0] is None
          and lt.qty_value("01.09.2026")[0] is None and lt.price_value("1/9/26", False)[0] is None, f"{got}; {bad}")
    rows = [row(A=1.0, B="Короткие даты", C="УПД-50", F=2.0, H=2.0, I="шт", L="П", N="1.9.26", O="01.09.26", P="1/9/26", Q="1-9-26",
                U="D-5", V="4101", X=2.0),                                                                                     # 2
            row(A=2.0, B="Полные даты", C="УПД-51", F=1.0, H=1.0, I="шт", L="П", N="01.09.2026", O="01/09/2026", P="01-09-2026",
                Q=" 01.09.2026 ", U="D-5", V="4102", X=1.0),                                                                   # 3
            row(A=3.0, B="Даты Calc", C="УПД-52", F=1.0, H=1.0, I="шт", L="П", N=d(-3) + 0.5, O=d(-3), P=d(-10), Q=d(20), U="D-5",
                V="4103", X=1.0),                                                                                              # 4
            row(A=4.0, B="Пустые даты", H=2.0, I="шт", L="П"),                                                                  # 5
            row(A=5.0, B="Непонятные даты заказа", H=2.0, I="шт", L="П", P="когда-то", Q="31.02.2026"),                         # 6
            row(A=6.0, B="Непонятная дата документа", C="УПД-53", F=1.0, H=1.0, I="шт", L="П", N="01.09.2026", O="вчера", U="D-5",
                V="4104", X=1.0)]                                                                                              # 7
    cd, st, work, res = transfer("l41", rows, steps=("check", "build"))
    F = [f for f in findings(work) if f["kind"] == "дата"]
    s = ready(Session(test_book(work))) if res.get("build", (1,))[0] == 0 else None
    try:
        new = orders_of(s, 7) if s else [[""] * 29] * 6
        R.add(c, "в тестовой WMS: 1.9.26, 01.09.26, 1/9/26, 1-9-26, 01.09.2026, 01/09/2026, 01-09-2026, « 01.09.2026 » — настоящие даты "
                 "01.09.2026 в N, O, P, Q; дата и дата-время Calc — дата; пустые даты — пусто; ни одного замечания о датах у этих строк",
              res["check"][0] == 0 and s is not None and new[0][13:17] == [sep1] * 4 and new[1][13:17] == [sep1] * 4
              and new[2][13:17] == [d(-3), d(-3), d(-10), d(20)] and new[3][13:17] == ["", "", "", ""]
              and not any(f["row"] in ("2", "3", "4", "5") for f in F),
              f"{[r[13:17] for r in new[:4]]}; {[(f['row'], f['text'][:60]) for f in F]}; {res['check'][1][-150:]}")
        R.add(c, "непонятная дата не останавливает перенос: «когда-то» (P), «31.02.2026» (Q), «вчера» (O прихода) — жёлтые, дата пустая, "
                 "исходный текст — в комментарии Z; приход с непонятной датой документа — «Получено без документов»",
              s is not None and sorted((f["row"], f["level"]) for f in F) == [("6", "ПРОВЕРИТЬ"), ("6", "ПРОВЕРИТЬ"), ("7", "ПРОВЕРИТЬ")]
              and new[4][15:17] == ["", ""] and "когда-то" in new[4][25] and "31.02.2026" in new[4][25] and new[5][14] == ""
              and "вчера" in new[5][25] and new[5][22] == "Получено без документов",
              f"{[(f['row'], f['level'], f['text'][:70]) for f in F]}; {new[4][15:17]} {new[4][25]!r}; {new[5][14]!r} {new[5][22]} {new[5][25]!r}")
    finally:
        if s:
            s.close()
    # the one date WMS needs: the receipt date N of a receipt
    badn = [row(A=1.0, B="Непонятная дата поступления", C="УПД-54", F=1.0, H=1.0, I="шт", L="П", N="позавчера", O="01.09.2026", U="D-5",
                V="4105", X=1.0)]
    cd2, st2, work2, res2 = transfer("l41b", badn, steps=("check",))
    F2 = findings(work2)
    cd3, st3, work3, res3 = transfer("l41c", badn, steps=("check", "build"), no_receipt_date="doc")
    s = ready(Session(test_book(work3))) if res3.get("build", (1,))[0] == 0 else None
    try:
        r3 = orders_of(s, 2)[0] if s else [""] * 29
        R.add(c, "непонятная дата поступления (N) у прихода — единственный блокер по датам (без неё WMS не проведёт приход), с подсказкой; "
                 "правило «дата поступления = дата документа» — жёлтое, N = дата документа, исходное значение — в Z",
              res2["check"][0] == 1 and any(f["level"] == "БЛОКЕР" and f["row"] == "2" and "позавчера" in f["text"] and "обязательна" in f["text"]
                                            for f in F2)
              and s is not None and r3[13] == sep1 and "позавчера" in r3[25]
              and any(f["level"] == "ПРОВЕРИТЬ" and "взята дата документа" in f["text"] for f in findings(work3)),
              f"{[(f['level'], f['text'][:90]) for f in F2 if f['kind'] == 'дата']}; {r3[13]} {r3[25]!r}; {res3.get('build', res3['check'])[1][-120:]}")
    finally:
        if s:
            s.close()


# ================================================================ L42 the real statuses of the old table, no overdue (0.7.2)

@case
def l42_real_statuses():
    c = "L42"
    rows = [row(A=1.0, B="Ожидаем, срок прошёл", H=5.0, I="шт", L="П", Q=d(-12), W="Ожидаем"),                                   # 2
            row(A=2.0, B="ОЖИДАЕМ с пробелами", H=5.0, I="шт", L="П", Q=d(10), W="  ожидаем "),                                  # 3
            row(A=3.0, B="Отменён", H=5.0, I="шт", L="П", Q=d(-40), W="Отменён"),                                                # 4
            row(A=4.0, B="Отменен", H=5.0, I="шт", L="П", W="Отменен"),                                                          # 5
            row(A=5.0, B="Получен без документов", F=5.0, H=5.0, I="шт", L="П", N=d(-3), U="R-1", V="4201", W="Получен без документов",
                X=5.0),                                                                                                        # 6
            row(A=6.0, B="Ожидает размещения", H=3.0, I="шт", L="П", Z="срочно", W="Ожидает размещения"),                         # 7
            row(A=1.0, B="ожидает размещения, срок прошёл", H=3.0, I="шт", L="П", Q=d(-30), W="ожидает  размещения"),             # 8
            row(A=2.0, B="Размещён, срок прошёл", H=3.0, I="шт", L="П", Q=d(-5), W="Размещён"),                                  # 9
            row(A=3.0, B="Размещен, частично получен, срок прошёл", C="УПД-42", O=d(-2), F=1.0, H=3.0, I="шт", L="П", N=d(-2), Q=d(-5),
                U="R-1", V="4202", W="Размещен", X=1.0)]                                                                        # 10
    cd, st, work, res = transfer("l42", rows, steps=("check", "build", "verify"))
    F = findings(work)
    m = summary(work)
    s = ready(Session(test_book(work))) if res.get("build", (1,))[0] == 0 else None
    try:
        new = orders_of(s, 10)[:9] if s else []
        ws = [r[22] for r in new]
        want = ["Ожидается", "Ожидается", "Отменено", "Отменено", "Получено без документов", "Ожидается", "Ожидается", "Ожидается",
                "Частично получено"]
        R.add(c, "статусы старой таблицы пользователя: «Ожидаем», «Отменён», «Получен без документов», «Ожидает размещения», «Размещён» и их "
                 "варианты (регистр, е/ё, пробелы: «Отменен», «Размещен», «  ожидаем ») — понятны, блокеров и неизвестных статусов нет; "
                 "тестовая WMS создана и сверена; статусы — по данным",
              all(v[0] == 0 for v in res.values()) and m.get("blockers") == "0" and m.get("status_unknown") == "0" and ws == want,
              f"{ws}; {[v[1][-100:] for v in res.values()]}")
        R.add(c, "просрочки нет (D-089): ожидаемая дата прошла у 4 открытых позиций — они «Ожидается» / «Частично получено», ни одной строки "
                 "«Просрочено» в тестовой WMS; «2_Проверка»: «С прошедшей ожидаемой датой» 4 (информация, не статус)",
              s is not None and m.get("late_expected") == "4" and not any("просрочено" in str(r[22]).lower() for r in orders_of(s, 20))
              and m.get("overdue") is None, f"late {m.get('late_expected')}; {ws}")
        R.add(c, "«Ожидает размещения» — открытый заказ, ещё не размещён: статус WMS «Ожидается», смысл не теряется — в комментарии Z «Статус "
                 "до переноса: Ожидает размещения» (к прежнему комментарию); у «Размещён» и «Ожидаем» комментарий не меняется",
              s is not None and new[5][25] == "срочно; Статус до переноса: Ожидает размещения" and new[6][25] == "Статус до переноса: ожидает размещения"
              and new[7][25] == "" and new[0][25] == "", f"{[r[25] for r in new]}")
    finally:
        if s:
            s.close()


def main():
    only = set(a.lower() for a in sys.argv[1:])
    t0 = time.time()
    try:
        release()
    except Exception as e:
        R.error("setup", e)
    for fn in CASES:
        if only and not any(fn.__name__.startswith(o) for o in only):
            continue
        try:
            fn()
        except Exception as e:
            R.error(fn.__name__, e)
        subprocess.run(["pkill", "-9", "-f", "profile_(lt|s).*" + str(os.getpid())], capture_output=True)
        _mk.pop("o", None)
    R.save(os.path.join(OUT, "results_legacy.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_legacy.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
