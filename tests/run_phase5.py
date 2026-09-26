"""Core Phase 5 «Все специальные приходы» — automated scenarios (CLAUDE_TASK.md, section «ОБЯЗАТЕЛЬНЫЕ ТЕСТЫ»).

    WMS_TEST_OUT=/tmp/wms_out python3 tests/run_phase5.py [case ...]

Every case copies a freshly built test book (synthetic EIs 1–200, NEXT_EI 201, five recipients) into
WMS_TEST_OUT/cases/<case>/ and drives real LibreOffice processes like a user: typing through the UI (the change handler
of «Иной приход» runs), the macros the buttons of «Иной приход», «Выдачи», «Возврат» and «Заказы» are bound to for the
row under the cursor (questions answered in the test mode, window values preset), window close, kill -9. The book and
the journal are checked by the independent oracle tests/special_oracle.py (it runs tests/return_oracle.py and
tests/receipt_oracle.py: every ordinary receipt, issue, return and special receipt replayed from the journal and compared
with «Наличие», «Заказы», «Выдачи», «Возврат», «Иной приход», _SPR, _ART and the counters).
Results: WMS_TEST_OUT/results_phase5.json and WMS_TEST_OUT/TEST_REPORT_phase5.md.
"""
import datetime
import os
import random
import subprocess
import sys
import threading
import time
import zipfile

import uno  # noqa: F401  (LibreOffice Python-UNO)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Session, Results, new_wms, template  # noqa: E402
from com.sun.star.sheet import TableFilterField  # noqa: E402
from com.sun.star.sheet.FilterOperator import EQUAL  # noqa: E402

R = Results()
CASES = []
POSTED = "1111111111111111110"      # WmsConfig.SPECIAL_LOCKS_POSTED
OPEN = "1000000000000111110"        # WmsConfig.SPECIAL_LOCKS_OPEN
D1, D2, D3 = "21.09.2026", "22.09.2026", "23.09.2026"     # receipts, issues, returns
KEEP = "<как предложено>"            # WmsSpecialUi test seam: the window field keeps the value it proposed
PWD = "wms"
IVA, PET, EGR = "Иванов Иван Андреевич", "Петров Пётр Сергеевич", "Ермолин Егор Павлович"
FIX_FIELDS = ("qty", "date", "place", "cat", "doc", "who", "mark", "name", "art", "unit")
COUNTERS = ("LAST_SEQ", "NEXT_EI", "NEXT_SPL", "NEXT_OFF", "NEXT_PROD", "NEXT_DET", "NEXT_OLD", "NEXT_OTH")


def case(fn):
    CASES.append(fn)
    return fn


def st(s):
    d = s.state()
    return d.get("STATE"), d.get("BLOCK")


def ei(n):
    return f"ЕИ-{n:08d}"


def serial(dmy):
    d, m, y = (int(x) for x in dmy.split("."))
    return float((datetime.date(y, m, d) - datetime.date(1899, 12, 30)).days)


def ready(s):
    s.U("TestUiAuto", 1)
    return s


def njournal(s):
    return len(s.journal()[0])


def jtypes(s):
    return [e["type"] for e in s.journal()[0]]


def oracle(c, s, name, expect_tail=0):
    P, inf, hist = s.spcheck(expect_tail)
    R.add(c, "оракул: " + name, not P, f"{inf}; {P[:4]}")
    return hist


def line(s, r, typ, name=None, qty=None, unit="шт", date=D1, place="A-1", art=None, cat=None, who=None, doc=None, mark=None, ev=None, note=None):
    """one row of «Иной приход» typed like a user: the type first, the article next (a known part fills its card), then the rest"""
    s.special_input(r, B=typ, E=art, D=name, F=qty, G=unit, H=date, I=place, J=cat, K=who, L=doc, M=mark, C=ev, S=note)


def ctl(s, r):
    return str(s.spc(r)[17])


def post(s, r):
    return s.click_spc("BtnSpcPost", r)


def post_block(s, r0, r1, answer=1):
    """«Провести» for rows r0..r1 selected; the question «одно поступление?» answered «Да» (1) or «Нет» (2)"""
    s.U("TestUiAuto", answer)
    try:
        return s.click_spc("BtnSpcPost", r0, r1=r1)
    finally:
        s.U("TestUiAuto", 1)


def fix(s, r, **kw):
    s.XU("TestSpcInput", "\t".join(str(kw.get(k, KEEP)) for k in FIX_FIELDS))
    return s.click_spc("BtnSpcFix", r)


def fix_args(s, r, **kw):
    """the arguments of WmsSpecial.SpecialFixRow: the current values of the row, some of them replaced"""
    row = s.spc(r)
    q = row[5]
    d = dict(qty=str(int(q)) if q == int(q) else str(q).replace(".", ","), date=row[7], place=row[8], cat=row[9], doc=str(row[11]), who=str(row[10]),
             mark=str(row[12]), name=row[3], art=str(row[4]), unit=row[6])
    d.update(kw)
    return tuple(d[k] for k in FIX_FIELDS)


def delete(s, r):
    return s.click_spc("BtnSpcDelete", r)


def ident(s, r, typ, name, art="", cat=""):
    s.XU("TestSpcInput", "\t".join((typ, name, art, cat)))
    return s.click_spc("BtnSpcIdentify", r)


def confirms(s):
    return int(s.U("TestUiConfirmCount"))


def issue(s, r, e, q, who="ива", date=D2):
    s.issue_input(r, ei=e, qty=q, date=date, who=who)
    return s.post(r)


def ret(s, r, k, q, date=D3):
    s.return_input(r, issue=k, qty=q, date=date)
    return s.click_ret("BtnRetPost", r)


def order(s, r, a="З-1", b="Болт М10", h="100", i="шт", l="Петрович", u="A-1", f=None, n=D1):
    s.order_input(r, A=a, B=b, H=h, I=i, L=l, F=f, N=n, U=u, C="УПД-1", O=D1)


def counters(s):
    return [s.sysv(k) for k in COUNTERS]


def snap(s, rows=8, eis=(1, 201, 202, 203, 204)):
    vals = s.doc.Sheets.getByName("Иной приход").getCellRangeByPosition(0, 1, 18, rows).getDataArray()
    return (vals, [s.spc_locks(r) for r in range(1, rows + 1)], [s.card(n) for n in eis], [s.sprrec(k) for k in range(1, 6)], s.art_index(),
            counters(s), s.sysv("TX_STATE"))


def ctrl_z_harmless(s, n=5):
    a = snap(s)
    for _ in range(n):
        s.ui(".uno:Undo")
    return a == snap(s)


# ================================================================ X01–X06 Офис

@case
def x01_office_single():
    c = "X01"
    s = ready(Session(new_wms("x01")))
    try:
        line(s, 1, "Офис", "Стол письменный", "2", place="A-1")
        pre = ctl(s, 1)
        res = post(s, 1)
        row, card = s.spc(1), s.card(201)
        R.add(c, "одна позиция офиса без документа и без артикула: до проведения «Можно провести: Новый ЕИ (Офис)»; «Провести» — строка № 1, OFF-00000001, "
                 "новый ЕИ-00000201, остаток 2, статус «Проведено: новый ЕИ», строка защищена",
              pre == "Можно провести: Новый ЕИ (Офис)" and res.startswith("OK") and row[0] == 1.0 and row[1] == "Офис" and row[2] == "OFF-00000001"
              and row[13] == ei(201) and row[14] == "" and row[15] == 2.0 and row[16] == "Проведено: новый ЕИ"
              and str(row[17]).startswith("Новый ЕИ ЕИ-00000201 · OFF-00000001") and s.spc_locks(1) == POSTED, f"{pre}; {res}; {row}")
        R.add(c, "«Наличие»: одна строка ЕИ — наименование, единица, остаток, место, «Активен», источник «Офис OFF-00000001», тип источника «Офис»",
              card == (ei(201), "Стол письменный", "", "шт", 2.0, "A-1", "", "Активен", "Офис OFF-00000001", "Офис"), str(card))
        R.add(c, "счётчики: NEXT_EI 202, NEXT_SPL 2, NEXT_OFF 2, остальные источники не тронуты; в журнале одна запись SP_RECEIPT",
              counters(s)[1:] == [202, 2, 2, 1, 1, 1, 1] and jtypes(s) == ["SP_RECEIPT"], f"{counters(s)}; {jtypes(s)}")
        oracle(c, s, "одна позиция офиса")
    finally:
        s.close()


@case
def x02_office_one_event_many_lines():
    c = "X02"
    s = ready(Session(new_wms("x02")))
    try:
        for r, nm in ((1, "Стул"), (2, "Шкаф"), (3, "Тумба")):
            line(s, r, "Офис", nm, "1", who="Петрова")
        n0 = confirms(s)
        res = post_block(s, 1, 3, answer=1)
        msg = s.XU("TestSpcBlockMessage")
        evs = [s.spc(r)[2] for r in (1, 2, 3)]
        R.add(c, "три строки выделены → «Провести» → «Провести как одно поступление? — Да»: один № OFF-00000001 на все строки, у каждой строки свой ЕИ",
              res.split("\n")[0] == "OK=3;ERR=0;SKIP=0;STOP=0" and confirms(s) == n0 + 1 and evs == ["OFF-00000001"] * 3
              and [s.spc(r)[13] for r in (1, 2, 3)] == [ei(201), ei(202), ei(203)] and "Одно поступление: OFF-00000001" in msg and s.sysv("NEXT_OFF") == 2.0,
              f"{res.splitlines()[0]}; {evs}; {msg[-60:]}")
        for r, nm in ((4, "Лампа"), (5, "Полка")):
            line(s, r, "Офис", nm, "1")
        res2 = post_block(s, 4, 5, answer=2)
        evs2 = [s.spc(r)[2] for r in (4, 5)]
        R.add(c, "«Нет» — у каждой строки свой № поступления", res2.startswith("OK=2") and evs2 == ["OFF-00000002", "OFF-00000003"], f"{res2[:20]}; {evs2}")
        line(s, 6, "Офис", "Вешалка", "1", ev="off-1")
        line(s, 7, "Офис", "Зеркало", "1", ev="OFF-00000009")
        line(s, 8, "Офис", "Ковёр", "1", ev="PROD-1")
        line(s, 9, "Офис", "Урна", "1", ev="акт 5")
        out = [post(s, r) for r in (6, 7, 8, 9)]
        R.add(c, "№ поступления в C: «off-1» — строка присоединяется к выданному OFF-00000001; ещё не выданный №, № другого типа, не номер — отказ",
              out[0].startswith("OK") and s.spc(6)[2] == "OFF-00000001" and "ещё не проводилось" in out[1] and "относится к типу «Производство»" in out[2]
              and "не номер поступления" in out[3] and s.sysv("NEXT_OFF") == 4.0 and [s.spc(r)[0] for r in (7, 8, 9)] == ["", "", ""],
              str([x[:90] for x in out]))
        line(s, 10, "Офис", "Сейф", "1")
        line(s, 11, "Производство", "Стеллаж", "2")
        line(s, 12, "Офис", "Диван", "1")
        line(s, 13, "Офис", "Кулер", "1", unit=None)
        line(s, 14, "Офис", "Принтер", "1")
        res3 = post_block(s, 10, 14, answer=1)
        evs3 = [s.spc(r)[2] for r in (10, 11, 12, 14)]
        R.add(c, "блок из строк двух типов («Да»): у строк офиса один OFF-00000004, у производства свой PROD-00000001; строка без единицы отклонена, "
                 "остальные проведены (D-059)",
              res3.split("\n")[0] == "OK=4;ERR=1;SKIP=0;STOP=0" and evs3 == ["OFF-00000004", "PROD-00000001", "OFF-00000004", "OFF-00000004"]
              and s.spc(13)[0] == "" and "не указана единица измерения (G)" in ctl(s, 13), f"{res3.splitlines()[0]}; {evs3}; {ctl(s, 13)[:80]}")
        oracle(c, s, "одно поступление — несколько строк")
    finally:
        s.close()


@case
def x03_office_same_goods_twice():
    c = "X03"
    s = ready(Session(new_wms("x03")))
    try:
        line(s, 1, "Офис", "Кресло", "1")
        line(s, 2, "Офис", "Кресло", "1")
        line(s, 3, "Офис", "Кресло", "2")
        line(s, 4, "Офис", "кресло", "1", doc="Накл-5")
        out = [post(s, r) for r in (1, 2, 3, 4)]
        R.add(c, "одинаковый товар двумя отдельными приходами → разные ЕИ; антидубль: та же строка (тип, название, количество, дата, документ) — "
                 "предупреждение «Возможный дубль», другое количество или документ — без предупреждения",
              all(x.startswith("OK") for x in out) and [s.spc(r)[13] for r in (1, 2, 3, 4)] == [ei(201), ei(202), ei(203), ei(204)]
              and "Возможный дубль: строка № 1 (OFF-00000001, ЕИ-00000201)" in ctl(s, 2) and "дубль" not in ctl(s, 3) and "дубль" not in ctl(s, 4),
              f"{out}; {ctl(s, 2)[-70:]}")
        again = post(s, 1)
        again2 = s.X("SpecialPostRow", 2, "", False)
        R.add(c, "повторное «Провести» проведённой строки блокируется («уже проведено»): второй ЕИ, второй № и вторая запись журнала не созданы",
              again.startswith("SKIP:уже проведено") and again2.startswith("SKIP") and s.sysv("NEXT_EI") == 205.0 and njournal(s) == 4, f"{again}; {again2}")
        oracle(c, s, "одинаковый товар двумя приходами")
    finally:
        s.close()


@case
def x04_office_fix():
    c = "X04"
    s = ready(Session(new_wms("x04")))
    try:
        line(s, 1, "Офис", "Монитор", "3", place="A-1")
        post(s, 1)
        issue(s, 1, "201", "1")
        f1 = fix(s, 1, qty="5")
        row = s.spc(1)
        R.add(c, "«Исправить» количество 3 → 5 (выдано 1): остаток 2 → 4, статус «Проведено (исправлено): новый ЕИ», ЕИ и № поступления те же",
              f1.startswith("OK") and s.stock(201) == 4.0 and row[5] == 5.0 and row[15] == 4.0 and row[16] == "Проведено (исправлено): новый ЕИ"
              and row[13] == ei(201) and row[2] == "OFF-00000001", f"{f1}; {row[13:18]}")
        f2 = fix(s, 1, qty="0,5")
        R.add(c, "уменьшение ниже выданного (0,5 при выданном 1) — отказ, остаток не отрицательный", f2.startswith("ERR") and "отрицательным" in f2
              and s.stock(201) == 4.0, f2)
        f3 = fix(s, 1, name="Монитор 24", cat="Техника", place="B-2", doc="Акт 1", who="Петрова", date=D2)
        card, row = s.card(201), s.spc(1)
        R.add(c, "исправление наименования, категории, места, документа, «кто передал», даты: карточка ЕИ обновлена той же операцией",
              f3.startswith("OK") and card[1] == "Монитор 24" and card[6] == "Техника" and card[5] == "B-2" and row[11] == "Акт 1" and row[10] == "Петрова"
              and row[7] == serial(D2), f"{f3}; {card}")
        f4 = fix(s, 1, unit="упак")
        f5 = fix(s, 1)
        R.add(c, "единицу нельзя изменить при действующей выдаче; «ничего не изменилось» — без операции", f4.startswith("ERR") and "нельзя изменить" in f4
              and f5.startswith("SKIP") and s.card(201)[3] == "шт", f"{f4[:100]}; {f5}")
        n = njournal(s)
        fx = fix(s, 2, qty="1")
        R.add(c, "«Исправить» на пустой строке — отказ без операции", fx.startswith("ERR") and njournal(s) == n, fx)
        oracle(c, s, "исправление офиса")
    finally:
        s.close()


@case
def x05_office_storno():
    c = "X05"
    s = ready(Session(new_wms("x05")))
    try:
        line(s, 1, "Офис", "Принтер", "2")
        post(s, 1)
        issue(s, 1, "201", "1")
        n0, j0 = confirms(s), njournal(s)
        d1 = delete(s, 1)
        R.add(c, "«Удалить» при действующей выдаче ЕИ: сразу отказ (D-069) без окна подтверждения, ничего не записано",
              d1.startswith("ERR:По этому ЕИ (ЕИ-00000201) есть связанные выдачи/возвраты: действующих выдач 1, возвратов 0") and confirms(s) == n0
              and njournal(s) == j0 and s.spc(1)[16] == "Проведено: новый ЕИ", d1)
        s.I("IssueDeleteRow", 1)
        n1 = confirms(s)
        d2 = delete(s, 1)
        row, card = s.spc(1), s.card(201)
        R.add(c, "после сторно выдачи — подтверждение и сторно: остаток 0, ЕИ «Приход удалён (сторно)»; строка хранит историю (№, ЕИ, «Удалено (сторно)»), "
                 "защищена", d2.startswith("OK") and confirms(s) == n1 + 1 and row[0] == 1.0 and row[13] == ei(201) and row[16] == "Удалено (сторно)"
              and card[4] == 0.0 and card[7] == "Приход удалён (сторно)" and s.spc_locks(1) == POSTED, f"{d2}; {row[13:18]}; {card}")
        d3 = delete(s, 1)
        p3 = post(s, 1)
        line(s, 2, "Офис", "Принтер", "2")
        p4 = post(s, 2)
        R.add(c, "повторное «Удалить» — «уже удалена»; «Провести» удалённой — отказ; новый приход получает новые №: строка № 2, OFF-00000002, ЕИ-00000202",
              d3.startswith("SKIP") and p3.startswith("SKIP") and p4.startswith("OK") and s.spc(2)[:3] == (2.0, "Офис", "OFF-00000002")
              and s.spc(2)[13] == ei(202), f"{d3}; {p3}; {p4}")
        oracle(c, s, "сторно офиса")
    finally:
        s.close()


@case
def x06_office_issue_return():
    c = "X06"
    s = ready(Session(new_wms("x06")))
    try:
        line(s, 1, "Офис", "Ноутбук", "5", place="A-3")
        post(s, 1)
        out = [issue(s, r, e, "1") for r, e in ((1, "201"), (2, "ЕИ-201"), (3, "EI-00000201"), (4, "00000201"))]
        R.add(c, "выдача ЕИ офиса без отдельного кода: ЕИ введён как «201», «ЕИ-201», «EI-00000201», «00000201» — нормализован, 4 выдачи, остаток 1",
              all(x.startswith("OK") for x in out) and [s.iss(r)[11] for r in (1, 2, 3, 4)] == [ei(201)] * 4 and s.stock(201) == 1.0, str(out))
        r1 = ret(s, 1, "1", "1")
        r2 = ret(s, 2, "3", "1")
        R.add(c, "возврат по выдачам того же ЕИ (этап 4 без отдельного кода): остаток 1 → 3, тот же ЕИ", r1.startswith("OK") and r2.startswith("OK")
              and s.stock(201) == 3.0 and s.ret(1)[2] == ei(201), f"{r1}; {r2}")
        hist = oracle(c, s, "выдача и возврат офиса")
        h = hist.get(ei(201), {})
        R.add(c, "история ЕИ по журналу: откуда (Офис OFF-00000001), пришло 5, выдано 4, возвращено 2, остаток 3, место A-3",
              (h.get("origin"), h.get("received"), h.get("issued"), h.get("returned"), h.get("balance"), h.get("place"))
              == ("Офис OFF-00000001", 5.0, 4.0, 2.0, 3.0, "A-3"), str(h))
    finally:
        s.close()


# ================================================================ X07 Производство

@case
def x07_production():
    c = "X07"
    s = ready(Session(new_wms("x07")))
    try:
        line(s, 1, "Производство", "Корпус", "3", place="P-1", who="Цех 2")
        line(s, 2, "Производство", "Корпус", "4", place="P-1", art="K-10")
        line(s, 3, "Производство", "Корпус", "3", place="P-1", who="Цех 2")
        out = [post(s, r) for r in (1, 2, 3)]
        R.add(c, "приход без артикула, с артикулом, повтор того же названия: у каждого новый ЕИ (артикул производства не ключ), PROD-00000001…3; "
                 "повтор той же строки — «Возможный дубль»; «Кто передал» — цех",
              all(x.startswith("OK") for x in out) and [s.spc(r)[13] for r in (1, 2, 3)] == [ei(201), ei(202), ei(203)]
              and [s.spc(r)[2] for r in (1, 2, 3)] == ["PROD-00000001", "PROD-00000002", "PROD-00000003"] and s.card(202)[2] == "K-10"
              and s.card(202)[9] == "Производство" and "Возможный дубль: строка № 1" in ctl(s, 3) and s.art_index() == [] and s.spc(1)[10] == "Цех 2",
              f"{out}; {ctl(s, 3)[-60:]}")
        for r, (nm, q) in zip(range(4, 11), (("Вал", "5"), ("Втулка", "6"), ("Крышка", "7"), ("Шкив", "1"), ("Рама", "2"), ("Опора", "3"), ("Ось", "4"))):
            line(s, r, "Производство", nm, q, place="P-2", who="Цех 2", doc="Акт 125")
        res = post_block(s, 4, 10, answer=1)
        evs = {s.spc(r)[2] for r in range(4, 11)}
        R.add(c, "одна передача из производства из 7 позиций: один PROD-00000004 на все 7 строк, у каждой свой ЕИ (204…210)",
              res.startswith("OK=7;ERR=0") and evs == {"PROD-00000004"} and [s.spc(r)[13] for r in range(4, 11)] == [ei(n) for n in range(204, 211)],
              f"{res.splitlines()[0]}; {evs}")
        f = fix(s, 4, qty="8")
        d = delete(s, 5)
        i1 = issue(s, 1, "206", "3")
        r1 = ret(s, 1, "1", "2")
        R.add(c, "исправление, сторно, выдача и возврат ЕИ производства", f.startswith("OK") and d.startswith("OK") and i1.startswith("OK")
              and r1.startswith("OK") and s.stock(204) == 8.0 and s.stock(205) == 0.0 and s.card(205)[7] == "Приход удалён (сторно)" and s.stock(206) == 6.0,
              f"{f}; {d}; {i1}; {r1}")
        oracle(c, s, "производство")
    finally:
        s.close()


# ================================================================ X08–X14 Детали

@case
def x08_details_one_article_one_ei():
    c = "X08"
    s = ready(Session(new_wms("x08")))
    try:
        line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1", cat="Запчасти")
        p1 = post(s, 1)
        R.add(c, "новый уникальный артикул ABC-100 → новый ЕИ-00000201 (DET-00000001); индекс артикулов «ABC-100 → ЕИ-00000201»; тип источника «Детали»",
              p1.startswith("OK") and s.spc(1)[13] == ei(201) and s.spc(1)[16] == "Проведено: новый ЕИ" and s.art_index() == [("ABC-100", ei(201), "ABC-100")]
              and s.card(201)[8:] == ("Детали DET-00000001", "Детали"), f"{p1}; {s.art_index()}; {s.card(201)}")
        line(s, 2, "Детали", None, "50", art="abc‑100", unit=None, place=None)
        pre = s.spc(2)
        p2 = post(s, 2)
        row = s.spc(2)
        R.add(c, "второй приход того же артикула («abc‑100»: регистр, неразрывный дефис): до проведения в строке ЕИ-00000201, остаток 100 → 150, "
                 "наименование, единица, место и категория из карточки; после — тот же ЕИ, «Проведено: пополнение ЕИ», DET-00000002, новый ЕИ не создан",
              pre[3] == "Шестерня" and pre[6] == "шт" and pre[8] == "D-1" and pre[9] == "Запчасти" and pre[13] == ei(201) and (pre[14], pre[15]) == (100.0, 150.0)
              and p2.startswith("OK") and row[13] == ei(201) and (row[14], row[15]) == (100.0, 150.0) and row[16] == "Проведено: пополнение ЕИ"
              and row[2] == "DET-00000002" and s.stock(201) == 150.0 and s.sysv("NEXT_EI") == 202.0
              and "Пополнение ЕИ-00000201: остаток 100 → 150 шт" in str(row[17]) and jtypes(s) == ["SP_RECEIPT", "SP_REFILL"], f"{pre[:17]}; {row[13:18]}")
        line(s, 3, "Детали", "Шестерёнка", "25", art=" АВС-100 ", unit=None, place=None)
        p3 = post(s, 3)
        j3 = s.journal()[0][-1]["fields"]
        R.add(c, "третий приход: артикул кириллицей с пробелами («АВС-100») — тот же ЕИ-00000201 (175); другое введённое название — предупреждение "
                 "«Артикул уже существует: ЕИ-00000201, название «Шестерня». Будет использовано существующее наименование», в строке и карточке — «Шестерня»",
              p3.startswith("OK") and s.spc(3)[13] == ei(201) and s.stock(201) == 175.0 and s.spc(3)[3] == "Шестерня" and s.card(201)[1] == "Шестерня"
              and "Артикул уже существует: ЕИ-00000201, название «Шестерня». Будет использовано существующее наименование" in ctl(s, 3)
              and j3.get("NAME_INPUT") == "Шестерёнка", f"{p3}; {ctl(s, 3)[-120:]}")
        line(s, 4, "Детали", "Шестерня", "10", art="ABC100", place="D-2")
        line(s, 5, "Детали", "Шестерня", "10", art="XYZ-1", place="D-3")
        line(s, 6, "Детали", None, "5", art="abc-100", unit=None, place=None)
        out = [post(s, r) for r in (4, 5, 6)]
        R.add(c, "«ABC100» без дефиса — другой артикул → новый ЕИ-00000202; то же название с артикулом XYZ-1 → новый ЕИ-00000203; четвёртый приход ABC-100 → "
                 "снова ЕИ-00000201 (180)", all(x.startswith("OK") for x in out) and [s.spc(r)[13] for r in (4, 5, 6)] == [ei(202), ei(203), ei(201)]
              and s.stock(201) == 180.0, f"{out}")
        R.add(c, "индекс: три артикула — три ЕИ, у каждой детали ровно одна запись",
              sorted(s.art_index()) == [("ABC-100", ei(201), "ABC-100"), ("ABC100", ei(202), "ABC100"), ("XYZ-1", ei(203), "XYZ-1")], str(s.art_index()))
        line(s, 7, "Детали", None, "5", art="ABC-100", unit=None, place=None)
        p7 = post(s, 7)
        again = post(s, 7)
        R.add(c, "повторный реальный приход того же артикула с тем же количеством и датой — проводится (пополнение 185) с предупреждением «Возможный дубль: "
                 "строка № 6»; повторное проведение той же строки — блокируется («уже проведено»)",
              p7.startswith("OK") and "Возможный дубль: строка № 6 (DET-00000006, ЕИ-00000201)" in ctl(s, 7) and s.stock(201) == 185.0
              and again.startswith("SKIP:уже проведено") and s.stock(201) == 185.0, f"{p7}; {ctl(s, 7)[-70:]}; {again}")
        hist = oracle(c, s, "детали: один артикул — один ЕИ")
        h = hist.get(ei(201), {})
        R.add(c, "история ЕИ детали по журналу: создан «Детали DET-00000001», пополнялся 4 раза, всего пришло 185, остаток 185",
              (h.get("origin"), h.get("refills"), h.get("received"), h.get("balance")) == ("Детали DET-00000001", 4, 185.0, 185.0), str(h))
    finally:
        s.close()


@case
def x09_details_refusals():
    c = "X09"
    s = ready(Session(new_wms("x09")))
    try:
        line(s, 1, "Детали", "Шестерня", "10", place="D-1")
        p1 = post(s, 1)
        line(s, 2, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
        p2 = post(s, 2)
        line(s, 3, "Детали", None, "5", art="ABC-100", unit="упак", place=None)
        pre3 = ctl(s, 3)
        p3 = post(s, 3)
        line(s, 4, "Деталь", None, "5", art="ABC-100", unit="ШТ", place=None)
        p4 = post(s, 4)
        line(s, 5, "Склад", "Что-то", "1")
        p5 = post(s, 5)
        R.add(c, "пустой артикул детали — отказ; единица «упак» при ЕИ в «шт» — отказ «Для ЕИ-00000201 используется единица «шт». Пересчёт единиц не "
                 "настроен» (видно уже до проведения); «ШТ» — та же единица, «Деталь» — тот же тип; неизвестный тип — отказ",
              "артикул (E) обязателен" in p1 and p2.startswith("OK") and "Для ЕИ-00000201 используется единица «шт». Пересчёт единиц не настроен" in p3
              and "Для ЕИ-00000201 используется единица «шт»" in pre3 and p4.startswith("OK") and s.spc(4)[6] == "шт" and "неизвестен" in p5
              and s.stock(201) == 105.0 and jtypes(s) == ["SP_RECEIPT", "SP_REFILL"] and [s.spc(r)[0] for r in (1, 3, 5)] == ["", "", ""],
              str([x[:80] for x in (p1, p3, p4, p5)]))
        oracle(c, s, "отказы деталей")
    finally:
        s.close()


def set_art_row(s, r, row):
    art = s.doc.Sheets.getByName("_ART")
    art.unprotect(PWD)
    art.getCellRangeByPosition(0, r, 2, r).setDataArray((row,))
    art.protect(PWD)


@case
def x10_details_index_conflict():
    c = "X10"
    s = ready(Session(new_wms("x10")))
    try:
        line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
        line(s, 2, "Детали", "Болт", "50", art="QWE-7", place="D-2")
        post(s, 1)
        post(s, 2)
        oracle(c, s, "до повреждения индекса")
        set_art_row(s, 3, ("ABC-100", ei(202), "ABC-100"))
        j0, c0 = njournal(s), counters(s)
        line(s, 3, "Детали", None, "10", art="ABC-100", unit="шт", place="D-1")
        pre = ctl(s, 3)
        p3 = post(s, 3)
        R.add(c, "один артикул найден у двух ЕИ (запись индекса в обход WMS): FAIL-CLOSED «Для артикула «ABC-100» найдено несколько ЕИ (ЕИ-00000201, "
                 "ЕИ-00000202). Требуется разбор» — уже в предпросмотре и при проведении; ни один ЕИ не выбран, ничего не записано",
              p3.startswith("ERR:") and "Для артикула «ABC-100» найдено несколько ЕИ (ЕИ-00000201, ЕИ-00000202). Требуется разбор" in p3
              and "найдено несколько ЕИ" in pre and njournal(s) == j0 and counters(s) == c0 and s.stock(201) == 100.0 and s.stock(202) == 50.0
              and s.spc(3)[0] == "", f"{p3}; {pre[:80]}")
        sc = s.B("ActionSelfCheck")
        rb = s.X("ArticleIndexRebuild")
        p3b = post(s, 3)
        R.add(c, "самопроверка показывает конфликт; перестройка индекса из «Наличие»: записей 2, конфликтов 0; тот же приход пополняет ЕИ-00000201",
              "найдено несколько ЕИ" in sc and "записей 2" in rb and "конфликтов «один артикул — несколько ЕИ» 0" in rb and p3b.startswith("OK")
              and s.stock(201) == 110.0, f"{rb}; {p3b}")
        # the index points at an EI whose card has another article: a problem of WMS itself — the block stops (D-059)
        set_art_row(s, 3, ("ZZZ-1", ei(202), "ZZZ-1"))
        line(s, 4, "Офис", "Стол", "1")
        line(s, 5, "Детали", "Прокладка", "5", art="ZZZ-1", place="D-5")
        line(s, 6, "Офис", "Стул", "1")
        res = post_block(s, 4, 6, answer=2)
        head, _, sysl = res.split("\n", 2)
        R.add(c, "индекс указывает на ЕИ с другим артикулом: системная ошибка WMS (ERR-SYS «индекс артикулов не согласован»), блок остановлен на этой строке",
              head == "OK=1;ERR=0;SKIP=0;STOP=6" and "ERR-SYS:" in sysl and "не согласован" in sysl and s.spc(5)[0] == "" and s.spc(6)[0] == "",
              f"{head}; {sysl[:140]}")
        s.X("ArticleIndexRebuild")
        res2 = post_block(s, 5, 6, answer=2)
        R.add(c, "после перестройки индекса строки проводятся: ZZZ-1 — новый ЕИ детали", res2.startswith("OK=2") and s.spc(5)[13] == ei(204), res2[:40])
        oracle(c, s, "после перестройки индекса")
        # an old (migrated) conflict: two part EIs with one article in «Наличие» — the rebuild lists it, the article fails closed
        stk = s.doc.Sheets.getByName("Наличие")
        stk.unprotect(PWD)
        stk.getCellByPosition(2, 202).setString("abc-100")
        stk.protect(PWD)
        rb2 = s.X("ArticleIndexRebuild")
        line(s, 7, "Детали", None, "1", art="ABC-100", unit="шт", place="D-1")
        p7 = post(s, 7)
        R.add(c, "старый конфликт (два ЕИ деталей с одним артикулом после миграции) не исправляется автоматически: перестройка показывает его, приход "
                 "артикула — отказ до разбора", "конфликтов «один артикул — несколько ЕИ» 1" in rb2 and "ABC-100: ЕИ-00000201, ЕИ-00000202" in rb2
              and "найдено несколько ЕИ" in p7 and s.spc(7)[0] == "", f"{rb2}; {p7[:100]}")
        stk.unprotect(PWD)
        stk.getCellByPosition(2, 202).setString("QWE-7")
        stk.protect(PWD)
        s.X("ArticleIndexRebuild")
        oracle(c, s, "конфликт разобран")
    finally:
        s.close()


@case
def x11_details_place_change():
    c = "X11"
    s = ready(Session(new_wms("x11")))
    try:
        line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
        post(s, 1)
        line(s, 2, "Детали", None, "50", art="ABC-100", unit=None, place=None)
        s.type_spc("I", 2, "Z-9")
        pre = ctl(s, 2)
        n0, j0 = confirms(s), njournal(s)
        s.U("TestUiAuto", 2)
        p_no = post(s, 2)
        s.U("TestUiAuto", 1)
        R.add(c, "другое место у повторного прихода: предупреждение в «Контроль» до проведения; «Провести» спрашивает о перемещении; «Нет» — не проведено, "
                 "ничего не записано", "место ЕИ-00000201 изменится: «D-1» → «Z-9»" in pre and p_no.startswith("SKIP:перемещение ЕИ не подтверждено")
              and confirms(s) == n0 + 1 and njournal(s) == j0 and s.spc(2)[0] == "" and s.card(201)[5] == "D-1", f"{pre[-60:]}; {p_no}")
        p_yes = post(s, 2)
        j = s.journal()[0][-1]["fields"]
        R.add(c, "«Да» — приход проведён, место ЕИ стало Z-9 той же операцией (в журнале PLACE_BEFORE D-1)", p_yes.startswith("OK") and s.card(201)[5] == "Z-9"
              and s.stock(201) == 150.0 and j.get("PLACE_BEFORE") == "D-1" and j.get("PLACE") == "Z-9", f"{p_yes}; {s.card(201)}")
        line(s, 3, "Детали", None, "5", art="ABC-100", unit=None, place=None)
        s.type_spc("I", 3, "Y-1")
        line(s, 4, "Детали", "Новая деталь", "1", art="NEW-1", place="N-1")
        res = post_block(s, 3, 4, answer=2)
        R.add(c, "в выделенном блоке перемещение не подтверждается: строка с другим местом отклонена («проведите строку отдельно»), остальные проведены",
              res.split("\n")[0] == "OK=1;ERR=1;SKIP=0;STOP=0" and "перемещение ЕИ нужно подтвердить" in ctl(s, 3) and s.spc(4)[0] != ""
              and s.card(201)[5] == "Z-9", f"{res.splitlines()[0]}; {ctl(s, 3)[:90]}")
        line(s, 5, "Детали", None, "5", art="ABC-100", unit=None, place=None, cat="Крепёж")
        n1 = confirms(s)
        p5 = post(s, 5)
        R.add(c, "место не указано — подставлено текущее Z-9, без вопроса; другая категория — предупреждение, категория карточки сохранена (второй ЕИ не создан)",
              p5.startswith("OK") and confirms(s) == n1 and s.spc(5)[8] == "Z-9" and "категория карточки ЕИ-00000201" in ctl(s, 5) and s.card(201)[6] == ""
              and s.spc(5)[13] == ei(201) and s.spc(5)[9] == "", f"{p5}; {ctl(s, 5)[-90:]}")
        oracle(c, s, "место и категория детали")
    finally:
        s.close()


@case
def x12_details_fix_storno():
    c = "X12"
    s = ready(Session(new_wms("x12")))
    try:
        line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
        line(s, 2, "Детали", None, "50", art="ABC-100", unit=None, place=None)
        post(s, 1)
        post(s, 2)
        issue(s, 1, "201", "120")
        n0 = confirms(s)
        d2 = delete(s, 2)
        d1 = delete(s, 1)
        R.add(c, "пример задания: приход 100 + пополнение 50, выдано 120, остаток 30 — сторно пополнения 50 запрещено (остаток стал бы −20), сторно первого "
                 "прихода тоже; окно подтверждения не открывалось",
              "сторно сделало бы остаток отрицательным (20 не покрыто)" in d2 and "(70 не покрыто)" in d1 and confirms(s) == n0 and s.stock(201) == 30.0,
              f"{d2[:120]}; {d1[:60]}")
        fx = s.I("IssueFixRow", 1, ei(201), "80", IVA, D2)
        d2b = delete(s, 2)
        R.add(c, "выдано 80 (исправление выдачи), остаток 70: сторно пополнения 50 допустимо — остаток 20, ЕИ остаётся «Активен» (действует приход строки 1)",
              fx.startswith("OK") and d2b.startswith("OK") and confirms(s) == n0 + 1 and s.stock(201) == 20.0 and s.card(201)[7] == "Активен"
              and s.spc(2)[16] == "Удалено (сторно)", f"{fx}; {d2b}")
        line(s, 3, "Детали", None, "10", art="ABC-100", unit=None, place=None)
        post(s, 3)
        f3 = fix(s, 3, qty="40")
        f3n = fix(s, 3, name="Другое")
        f3u = fix(s, 3, unit="упак")
        f3p = fix(s, 3, place="P-5")
        R.add(c, "исправление пополнения: количество 10 → 40 (остаток 30 → 60); наименование и единицу пополнения не исправить — отказ; место — перемещение ЕИ",
              f3.startswith("OK") and s.spc(3)[16] == "Проведено (исправлено): пополнение ЕИ" and "от карточки" in f3n and "от карточки" in f3u
              and f3p.startswith("OK") and s.card(201)[5] == "P-5" and s.stock(201) == 60.0, f"{f3}; {f3n[:60]}; {f3p}")
        f1 = fix(s, 1, qty="50")
        f1b = fix(s, 1, qty="5")
        R.add(c, "исправление первого прихода детали 100 → 50 (остаток 10); до 5 — отказ (остаток стал бы отрицательным)", f1.startswith("OK")
              and s.stock(201) == 10.0 and f1b.startswith("ERR") and "отрицательным" in f1b, f"{f1}; {f1b[:80]}")
        f1a = fix(s, 1, art="ABC-101")
        line(s, 4, "Детали", "Шестерня", "7", art="ABC-100", place="D-9")
        p4 = post(s, 4)
        line(s, 5, "Детали", None, "3", art="abc-101", unit=None, place=None)
        p5 = post(s, 5)
        f1c = fix(s, 1, art="ABC-100")
        f1u = fix(s, 1, unit="кг")
        R.add(c, "артикул первой строки ABC-100 → ABC-101 (свободный): индекс изменён той же операцией; ABC-100 теперь новый ЕИ-00000202, ABC-101 пополняет "
                 "ЕИ-00000201; сделать ABC-100 снова (он у ЕИ-00000202) — отказ; единицу — отказ (есть другие приходы и выдача)",
              f1a.startswith("OK") and p4.startswith("OK") and s.spc(4)[13] == ei(202) and p5.startswith("OK") and s.spc(5)[13] == ei(201)
              and "уже принадлежит ЕИ-00000202" in f1c and "нельзя изменить" in f1u
              and sorted(s.art_index()) == [("ABC-100", ei(202), "ABC-100"), ("ABC-101", ei(201), "ABC-101")], f"{f1a}; {p4}; {p5}; {f1c[:80]}; {f1u[:60]}")
        oracle(c, s, "исправление и сторно деталей")
    finally:
        s.close()


@case
def x13_details_issues_returns():
    c = "X13"
    s = ready(Session(new_wms("x13")))
    try:
        line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
        line(s, 2, "Детали", None, "50", art="ABC-100", unit=None, place=None)
        post(s, 1)
        post(s, 2)
        out = [issue(s, 1, "201", "10", who="ива"), issue(s, 2, "201", "20", who="пет"), issue(s, 3, "201", "30", who="ива")]
        rr = [ret(s, 1, "2", "5"), ret(s, 2, "3", "30")]
        R.add(c, "несколько выдач одного ЕИ детали и возвраты по ним: остаток 150 − 60 + 35 = 125", all(x.startswith("OK") for x in out + rr)
              and s.stock(201) == 125.0, str(out + rr))
        d1 = delete(s, 1)
        R.add(c, "сторно первого прихода (100) при другом действующем приходе ЕИ: правило остатка — 125 ≥ 100, допустимо, остаток 25",
              d1.startswith("OK") and s.stock(201) == 25.0 and s.card(201)[7] == "Активен", d1)
        n0 = confirms(s)
        d2 = delete(s, 2)
        R.add(c, "сторно последнего действующего прихода ЕИ при действующих выдачах и возвратах: отказ (D-069) без подтверждения",
              d2.startswith("ERR:По этому ЕИ (ЕИ-00000201) есть связанные выдачи/возвраты: действующих выдач 3, возвратов 2") and confirms(s) == n0
              and s.stock(201) == 25.0, d2)
        oracle(c, s, "выдачи и возвраты детали")
    finally:
        s.close()


@case
def x14_double_click_copies():
    c = "X14"
    p = new_wms("x14")
    s = ready(Session(p))
    line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
    post(s, 1)
    line(s, 2, "Детали", None, "50", art="ABC-100", unit=None, place=None)
    s.goto("Иной приход", "$B$3")
    res = []

    def click():
        s.XU("BtnSpcPost")
        res.append(s.U("TestUiLastMessage"))
    th = [threading.Thread(target=click) for _ in range(2)]
    for t in th:
        t.start()
    for t in th:
        t.join(60)
    refills = [e for e in s.journal()[0] if e["type"] == "SP_REFILL"]
    R.add(c, "двойной клик «Провести» пополнения детали: одно пополнение (150, не 200), одна запись журнала, один DET-№", len(refills) == 1
          and s.stock(201) == 150.0 and s.sysv("NEXT_DET") == 3.0 and s.sysv("NEXT_SPL") == 3.0, f"ответы {res}; SP_REFILL {len(refills)}")
    line(s, 3, "Детали", None, "5", art="ABC-100", unit=None, place=None)
    ans = []
    th = [threading.Thread(target=lambda: ans.append(s.X("SpecialPostRow", 3, "", False))) for _ in range(2)]
    for t in th:
        t.start()
    for t in th:
        t.join(60)
    R.add(c, "два одновременных вызова проведения одной строки: «OK» и «уже проведено»", sorted(a.split(":")[0] for a in ans) == ["OK", "SKIP"]
          and s.stock(201) == 155.0, str(ans))
    s.goto("Иной приход", "$A$2:$S$2")
    s.ui(".uno:Copy")
    s.goto("Иной приход", "$A$9")
    s.ui(".uno:Paste")
    R.add(c, "копировать → вставить проведённую строку на защищённом листе: вставка отклонена", s.spc(8)[0] == "" and s.spc(8)[5] == "", str(s.spc(8)[:6]))
    sh = s.doc.Sheets.getByName("Иной приход")
    sh.unprotect(PWD)
    s.goto("Иной приход", "$A$2:$S$2")
    s.ui(".uno:Copy")
    s.goto("Иной приход", "$A$10")
    s.ui(".uno:Paste")
    sh.protect(PWD)
    live = ctl(s, 9)
    pr = post(s, 9)
    fx = s.X("SpecialFixRow", 9, *fix_args(s, 9, qty="1"))
    dl = s.X("SpecialDeleteRow", 9)
    R.add(c, "защита снята паролем, проведённая строка вставлена: сразу КОПИЯ, «Провести», «Исправить», «Удалить» — отказ; остаток не изменился",
          live.startswith("КОПИЯ") and "повторяет строку 2" in live and pr.startswith("SKIP") and fx.startswith("ERR") and dl.startswith("ERR")
          and s.stock(201) == 155.0, f"{live[:90]}; {pr}; {fx}; {dl}")
    cl = s.click_spc("BtnSpcClear", 9)
    R.add(c, "«Очистить» убирает копию, строка открыта для ввода", cl.startswith("OK") and s.spc(9)[0] == "" and s.spc_locks(9) == OPEN, cl)
    s.close(save=True)
    s = Session(p, macros=0)
    sh = s.doc.Sheets.getByName("Иной приход")
    sh.unprotect(PWD)
    row = list(sh.getCellRangeByPosition(0, 1, 18, 1).getDataArray()[0])
    sh.getCellRangeByPosition(0, 4, 18, 4).setDataArray((tuple(row),))
    row[0] = 99.0
    sh.getCellRangeByPosition(0, 5, 18, 5).setDataArray((tuple(row),))
    row[0] = 7.0
    sh.getCellRangeByPosition(0, 6, 18, 6).setDataArray((tuple(row),))
    sh.protect(PWD)
    sy = s.doc.Sheets.getByName("_SYS")
    sy.unprotect(PWD)
    sy.getCellByPosition(1, 19).setValue(10)                 # NEXT_SPL 10: № 7 is below it, but WMS never registered it
    sy.protect(PWD)
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        rep = s.report()
        r_ = [ctl(s, r) for r in (4, 5, 6)]
        pr = [post(s, r) for r in (4, 5, 6)]
        cl = [s.click_spc("BtnSpcClear", r) for r in (4, 5, 6)]
        orig = s.spc(1)
        R.add(c, "сохранено без макросов: копия, № который WMS не выдавала, и № без регистрации в _SPR помечены КОПИЯ при запуске; не проводятся; "
                 "«Очистить» убирает; оригинал цел",
              "без WMS" in rep and r_[0].startswith("КОПИЯ") and "повторяет" in r_[0] and r_[1].startswith("КОПИЯ") and "не выдавался" in r_[1]
              and r_[2].startswith("КОПИЯ") and "не зарегистрирована" in r_[2] and all(x.startswith("SKIP") for x in pr)
              and all(x.startswith("OK") for x in cl) and all(s.spc(r)[0] == "" for r in (4, 5, 6)) and orig[0] == 1.0
              and orig[16] == "Проведено: новый ЕИ", f"{[x[:70] for x in r_]}; {pr}; {cl}; {rep[:120]}")
        line(s, 4, "Офис", "Стол", "1")
        p4 = post(s, 4)
        R.add(c, "следующий приход получает № строки 10 (NEXT_SPL)", p4.startswith("OK") and s.spc(4)[0] == 10.0, p4)
        oracle(c, s, "копии не считаются приходами")
    finally:
        s.close()


# ================================================================ X15 Старый склад, X16 Иной приход

@case
def x15_old_warehouse():
    c = "X15"
    s = ready(Session(new_wms("x15")))
    try:
        line(s, 1, "Старый склад", "Дрель ударная", "1", place="С-3", cat="Инструмент", who="Иванов", mark="инв. 00123", note="найдена при разборе стеллажа")
        p1 = post(s, 1)
        row, card = s.spc(1), s.card(201)
        R.add(c, "найденная старая позиция без документов: новый ЕИ-00000201, OLD-00000001, обычный остаток; источник «Старый склад», старая маркировка "
                 "и комментарий сохранены",
              p1.startswith("OK") and row[2] == "OLD-00000001" and row[12] == "инв. 00123" and row[18] == "найдена при разборе стеллажа"
              and card == (ei(201), "Дрель ударная", "", "шт", 1.0, "С-3", "Инструмент", "Активен", "Старый склад OLD-00000001", "Старый склад"), f"{p1}; {card}")
        line(s, 2, "Старый склад", "Уровень строительный", "2", place="С-4", art="УР-60")
        p2 = post(s, 2)
        i1 = issue(s, 1, "201", "1")
        r1 = ret(s, 1, "1", "1")
        f2 = fix(s, 2, qty="3", mark="инв. 7")
        d2 = delete(s, 2)
        R.add(c, "второе событие OLD-00000002; выдача и возврат; исправление (количество, маркировка); сторно", p2.startswith("OK") and s.spc(2)[2] == "OLD-00000002"
              and i1.startswith("OK") and r1.startswith("OK") and s.stock(201) == 1.0 and f2.startswith("OK") and s.spc(2)[12] == "инв. 7"
              and d2.startswith("OK") and s.card(202)[7] == "Приход удалён (сторно)", f"{p2}; {i1}; {r1}; {f2}; {d2}")
        oracle(c, s, "старый склад")
    finally:
        s.close()


@case
def x16_other_receipt():
    c = "X16"
    s = ready(Session(new_wms("x16")))
    try:
        line(s, 1, "Иной", "Коробка без маркировки", "1", place="Q-1", note="фото — у кладовщика")
        p1 = post(s, 1)
        row, card = s.spc(1), s.card(201)
        R.add(c, "неопознанный товар сразу получает ЕИ-00000201 (OTH-00000001), считается остатком, «Требует разбора»; статус «Проведено: требует разбора»",
              p1.startswith("OK") and row[2] == "OTH-00000001" and row[16] == "Проведено: требует разбора" and "позднее — «Разобрать»" in str(row[17])
              and card == (ei(201), "Коробка без маркировки", "", "шт", 1.0, "Q-1", "", "Требует разбора", "Иной приход OTH-00000001", "Иной приход"),
              f"{p1}; {card}")
        i1 = issue(s, 1, "201", "1")
        r1 = ret(s, 1, "1", "1")
        R.add(c, "ЕИ иного прихода выдаётся и возвращается до разбора", i1.startswith("OK") and r1.startswith("OK") and s.stock(201) == 1.0, f"{i1}; {r1}")
        g1 = ident(s, 1, "Офис", "Коробка с кабелями", "", "Кабель")
        R.add(c, "«Разобрать» → Офис, «Коробка с кабелями», категория «Кабель»: тот же ЕИ-00000201, остаток сохранён, «Активен», тип источника «Офис»; "
                 "новый ЕИ не создан", g1.startswith("OK") and s.card(201) == (ei(201), "Коробка с кабелями", "", "шт", 1.0, "Q-1", "Кабель", "Активен",
                                                                              "Иной приход OTH-00000001 → Офис", "Офис")
              and s.spc(1)[16] == "Проведено: разобрано" and s.spc(1)[13] == ei(201) and s.sysv("NEXT_EI") == 202.0, f"{g1}; {s.card(201)}")
        g1b = ident(s, 1, "Офис", "Другое")
        R.add(c, "повторный «Разобрать» — отказ (строка уже разобрана)", g1b.startswith("ERR") and "требует разбора" in g1b, g1b)
        line(s, 2, "Иной", "Непонятная деталь", "10", place="Q-2")
        post(s, 2)
        g2 = ident(s, 2, "Детали", "Шестерня малая", "dt-5", "Запчасти")
        line(s, 3, "Детали", None, "5", art="DT-5", unit=None, place=None)
        p3 = post(s, 3)
        R.add(c, "разбор как «Детали» со свободным артикулом: артикул записан в индекс; следующий приход детали DT-5 пополняет тот же ЕИ-00000202",
              g2.startswith("OK") and ("DT-5", ei(202), "dt-5") in s.art_index() and s.card(202)[9] == "Детали" and p3.startswith("OK")
              and s.spc(3)[13] == ei(202) and s.stock(202) == 15.0, f"{g2}; {p3}; {s.art_index()}")
        line(s, 4, "Иной", "Ещё деталь", "3", place="Q-3")
        post(s, 4)
        g4 = ident(s, 4, "Детали", "Шестерня малая", "DT-5")
        g4b = ident(s, 4, "Детали", "Шестерня малая", "")
        g3 = ident(s, 3, "Офис", "x")
        R.add(c, "разбор как «Детали» с артикулом, который уже у ЕИ-00000202: отказ «та же деталь: сторно и приход детали»; без артикула — отказ; "
                 "«Разобрать» строки не иного прихода — отказ",
              "уже принадлежит ЕИ-00000202 — это та же деталь" in g4 and "артикул обязателен" in g4b and g3.startswith("ERR")
              and s.card(203)[7] == "Требует разбора", f"{g4[:100]}; {g4b[:60]}; {g3[:60]}")
        line(s, 5, "Иной", "Кабель без бирки", "20", unit="м", place="Q-5")
        post(s, 5)
        g5 = ident(s, 5, "Поставщик", "Кабель ВВГ 3×2,5", "", "Кабель")
        d4 = delete(s, 4)
        f1 = fix(s, 1, qty="2")
        R.add(c, "разбор как «Поставщик» (обычный товар без заказа); сторно иного прихода — «Приход удалён (сторно)»; исправление разобранной строки — "
                 "«Проведено (исправлено): разобрано»", g5.startswith("OK") and s.card(204)[9] == "Поставщик" and d4.startswith("OK")
              and s.card(203)[7] == "Приход удалён (сторно)" and f1.startswith("OK") and s.spc(1)[16] == "Проведено (исправлено): разобрано"
              and s.stock(201) == 2.0, f"{g5}; {d4}; {f1}")
        s.XU("TestSpcInput", "\t".join((KEEP, KEEP, KEEP, KEEP)))
        line(s, 6, "Иной", "Ящик", "1", place="Q-6")
        post(s, 6)
        g0 = s.click_spc("BtnSpcIdentify", 6)
        shown0 = s.XU("TestSpcDialog").split("\n")[0].split("\t")
        R.add(c, "окно «Разобрать» не выбирает тип за пользователя (D-081): поле типа пустое, «OK» без выбора — отказ, строка и ЕИ не изменены",
              shown0[0] == "" and g0.startswith("ERR") and "укажите, что это за приход" in g0 and s.spc(6)[16] == "Проведено: требует разбора"
              and s.card(int(s.spc(6)[13][3:]))[7] == "Требует разбора", f"{shown0}; {g0[:90]}")
        oracle(c, s, "иной приход и разбор")
    finally:
        s.close()


# ================================================================ X17–X23 crash / recovery

def reopen(p, s):
    s.kill()
    s = ready(Session(p))
    return s, s.B("ActionForceUnlock")


def spec_state(s, r, eis):
    return (s.spc(r)[:18], s.spc_locks(r), [s.card(n) for n in eis], [s.sprrec(k) for k in range(1, 5)], s.art_index(), counters(s), njournal(s))


def diff(a, b):
    names = ("строка", "защита", "«Наличие»", "_SPR", "_ART", "счётчики", "журнал")
    return [n for n, x, y in zip(names, a, b) if x != y]


def crash_series(c, tag, typ, prep=None, refill=False, eis=(201,), **row):
    """a posting of one row of type typ interrupted at every point before the journal (each time: rolled back by the
    snapshot at the next start), then after the journal (the next start sees the tail, «Восстановить» applies it once);
    then «Исправить» and «Удалить» of the same line interrupted in the middle of the writes and after the journal"""
    p = new_wms(f"x_crash_{tag}")
    s = ready(Session(p))
    try:
        if prep:
            prep(s)
        r = 2 if refill else 1
        line(s, r, typ, **row)
        s.doc.store()
        base = spec_state(s, r, eis)
        points = [("k", 0, "после резервирования № поступления"), ("k", 1, "после резервирования № строки")]
        if not refill:
            points.append(("k", 2, "после NEXT_EI"))
        points += [("sheet", "Наличие", "после изменения «Наличие»"), ("sheet", "_SPR", "после служебной записи события _SPR")]
        if typ == "Детали" and not refill:
            points.append(("sheet", "_ART", "после записи индекса артикулов"))
        points.append(("sheet", "Иной приход", "после пользовательской строки (все записи сделаны, журнала нет)"))
        fails = []
        for kind, arg, label in points:
            if kind == "k":
                s.B("TestSetFaultWrite", arg, 2)
            else:
                s.B("TestSetFaultSheet", arg, 2)
            crash = s.X("SpecialPostRow", r, "", True)
            s, unl = reopen(p, s)
            now = spec_state(s, r, eis)
            if not (crash == "CRASH-SIM" and "отменена по снимку" in unl and st(s)[0] == "CLEAN" and now == base):
                fails.append(f"{label}: {crash}; {unl[:60]}; различия {diff(base, now)}")
        what = typ + (" (пополнение ЕИ детали)" if refill else "")
        R.add(c, f"{what}: сбой в {len(points)} точках до записи в журнал ({', '.join(x[2] for x in points)}) — при каждом запуске откат по снимку: "
                 "нет ни ЕИ, ни № строки, ни № поступления, ни остатка, ввод на месте", not fails, "; ".join(fails[:3]) or f"точек {len(points)}")
        s.B("TestSetFault", 3, 2)
        crash = s.X("SpecialPostRow", r, "", True)
        s, unl = reopen(p, s)
        _, block = st(s)
        rec = s.B("ActionRecover")
        after = s.spc(r)
        again = s.X("SpecialPostRow", r, "", True)
        sp_ops = [e for e in s.journal()[0] if e["type"] in ("SP_RECEIPT", "SP_REFILL")]
        R.add(c, f"{what}: сбой после записи в журнал — при запуске «хвост», «Восстановить» доводит приход ровно один раз; повторное «Провести» — "
                 "«уже проведено»; один ЕИ, один № строки и поступления",
              crash == "CRASH-SIM" and block == "TAIL" and "восстановлено операций: 1" in rec and st(s)[0] == "CLEAN" and after[0] == float(r)
              and after[16].startswith("Проведено") and s.spc_locks(r) == POSTED and again.startswith("SKIP") and len(sp_ops) == r,
              f"{crash}; {block}; {rec[:80]}; {after[:3]}; {again}")
        # «Исправить»: in the middle of the writes, then after the journal
        base2 = spec_state(s, r, eis)
        q2 = str(int(after[5]) + 1)
        s.B("TestSetFault", 2, 2)
        crash = s.X("SpecialFixRow", r, *fix_args(s, r, qty=q2))
        s, unl = reopen(p, s)
        ok1 = crash == "CRASH-SIM" and spec_state(s, r, eis) == base2
        s.B("TestSetFault", 3, 2)
        crash2 = s.X("SpecialFixRow", r, *fix_args(s, r, qty=q2))
        s, unl = reopen(p, s)
        rec2 = s.B("ActionRecover")
        R.add(c, f"{what}: сбой посреди «Исправить» — откат; после журнала — «Восстановить» доводит исправление целиком",
              ok1 and crash2 == "CRASH-SIM" and "восстановлено операций: 1" in rec2 and s.spc(r)[5] == float(q2)
              and s.spc(r)[16].startswith("Проведено (исправлено)"), f"{crash}; {crash2}; {rec2[:80]}; {s.spc(r)[5]}")
        # «Удалить»
        base3 = spec_state(s, r, eis)
        s.B("TestSetFault", 2, 2)
        crash = s.X("SpecialDeleteRow", r)
        s, unl = reopen(p, s)
        ok1 = crash == "CRASH-SIM" and spec_state(s, r, eis) == base3
        s.B("TestSetFault", 3, 2)
        crash2 = s.X("SpecialDeleteRow", r)
        s, unl = reopen(p, s)
        rec3 = s.B("ActionRecover")
        R.add(c, f"{what}: сбой посреди «Удалить» — откат; после журнала — «Восстановить» доводит сторно целиком", ok1 and crash2 == "CRASH-SIM"
              and "восстановлено операций: 1" in rec3 and s.spc(r)[16] == "Удалено (сторно)", f"{crash}; {crash2}; {rec3[:80]}; {s.spc(r)[16]}")
        oracle(c, s, f"сбои: {what}")
    finally:
        s.close()


@case
def x17_crash_office():
    crash_series("X17", "off", "Офис", name="Стол", qty="2", place="A-1")


@case
def x18_crash_production():
    crash_series("X18", "prod", "Производство", name="Корпус", qty="3", place="P-1", who="Цех 2")


@case
def x19_crash_details_new():
    crash_series("X19", "det", "Детали", name="Шестерня", qty="100", place="D-1", art="ABC-100")


def prep_part(s):
    line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
    post(s, 1)


@case
def x20_crash_details_refill():
    crash_series("X20", "detadd", "Детали", prep=prep_part, refill=True, qty="50", art="ABC-100", unit=None, place=None)


@case
def x21_crash_old():
    crash_series("X21", "old", "Старый склад", name="Дрель", qty="1", place="С-3", mark="инв. 5")


@case
def x22_crash_other():
    crash_series("X22", "oth", "Иной", name="Коробка", qty="1", place="Q-1")


@case
def x23_abandon_keeps_numbers():
    c = "X23"
    p = new_wms("x23")
    s = ready(Session(p))
    try:
        line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
        s.doc.store()
        s.B("TestSetFault", 3, 2)
        crash = s.X("SpecialPostRow", 1, "", False)
        s, unl = reopen(p, s)
        ab = s.B("ActionAbandonTail")
        a = [e for e in s.journal()[0] if e["type"] == "ABANDON"]
        f = a[0]["fields"] if a else {}
        c1 = counters(s)
        again = post(s, 1)
        R.add(c, "«Отложить хвост» с приходом детали: № строки 1, DET-00000001 и ЕИ-00000201 больше не выдаются (ABANDON: NEXT_SPL 2, NEXT_DET 2, NEXT_EI 202); "
                 "тот же ввод проводится с новыми №: строка № 2, DET-00000002, ЕИ-00000202, индекс ABC-100 → ЕИ-00000202",
              crash == "CRASH-SIM" and "отложено операций: 1" in ab and (f.get("NEXT_SPL"), f.get("NEXT_DET"), f.get("NEXT_EI")) == ("2", "2", "202")
              and c1[1:] == [202, 2, 1, 1, 2, 1, 1] and again.startswith("OK") and s.spc(1)[0] == 2.0 and s.spc(1)[2] == "DET-00000002"
              and s.spc(1)[13] == ei(202) and s.art_index() == [("ABC-100", ei(202), "ABC-100")], f"{ab[:100]}; {f}; {c1}; {again}; {s.spc(1)[:3]}")
        oracle(c, s, "отложенный хвост")
    finally:
        s.close()


# ================================================================ X24–X29 sheet behaviour, protection, panel, old book, self-check

def set_filter(s, col, value):
    dbr = s.doc.DatabaseRanges.getByName("WMS_SPECIAL")
    fd = dbr.getFilterDescriptor()
    f = TableFilterField()
    f.Field, f.Operator, f.IsNumeric, f.StringValue = col, EQUAL, False, value
    fd.setFilterFields((f,))
    dbr.refresh()


def visible(s, r):
    return s.doc.Sheets.getByName("Иной приход").getRows().getByIndex(r).IsVisible


@case
def x24_inserted_rows_ctrl_z():
    c = "X24"
    s = ready(Session(new_wms("x24")))
    try:
        line(s, 1, "Офис", "Стол", "2")
        line(s, 2, "Детали", "Шестерня", "10", art="ABC-1", place="D-1")
        post(s, 1)
        post(s, 2)
        s.goto("Иной приход", "$A$2")
        s.ui(".uno:SelectRow")
        s.ui(".uno:InsertRowsBefore")
        s.ui(".uno:InsertRowsBefore")
        moved = (s.spc(3)[0], s.spc(4)[0])
        k = s.X("SpecialRowKind", 3)
        fx = fix(s, 3, qty="3")
        hint = s.sprrec(1)[7]
        dl = delete(s, 4)
        R.add(c, "вставка строк над приходами: строки сдвинулись, строка найдена по № (подсказка устарела), исправление и сторно из сдвинутых строк проведены, "
                 "подсказка исправлена той же операцией", moved == (1.0, 2.0) and k == "POSTED" and fx.startswith("OK") and hint == 3.0
              and dl.startswith("OK") and s.sprrec(2)[7] == 4.0 and s.stock(201) == 3.0 and s.stock(202) == 0.0, f"{moved}; {k}; {fx}; {hint}; {dl}")
        cl = s.click_spc("BtnSpcClear", 1)
        line(s, 1, "Офис", "Лампа", "1")
        p1 = post(s, 1)
        R.add(c, "вставленная строка: «Очистить» открывает её, в ней проводится новый приход (№ строки 3)", cl.startswith("OK") and p1.startswith("OK")
              and s.spc(1)[0] == 3.0, f"{cl}; {p1}")
        z1 = ctrl_z_harmless(s)
        fix(s, 1, qty="2")
        z2 = ctrl_z_harmless(s)
        delete(s, 1)
        z3 = ctrl_z_harmless(s)
        s.type_spc("S", 3, "комментарий")
        typed = s.spc(3)[18]
        s.ui(".uno:Undo")
        R.add(c, "Ctrl+Z ×5 после прихода, исправления и сторно ничего не откатывает; обычный ввод комментария S отменяется",
              z1 and z2 and z3 and typed == "комментарий" and s.spc(3)[18] == "" and s.state()["UNDO_LOCKED"] == "False", f"{z1} {z2} {z3}; {typed!r}")
        oracle(c, s, "вставка строк и Ctrl+Z")
    finally:
        s.close()


@case
def x25_filter_save_reopen():
    c = "X25"
    p = new_wms("x25")
    s = ready(Session(p))
    try:
        for r, (t, nm) in enumerate((("Офис", "Стол"), ("Производство", "Рама"), ("Офис", "Стул"), ("Производство", "Ось"), ("Офис", "Шкаф")), start=1):
            line(s, r, t, nm, "1")
        set_filter(s, 1, "Офис")
        vis = [visible(s, r) for r in range(1, 6)]
        res = post_block(s, 1, 5, answer=1)
        R.add(c, "автофильтр «Тип прихода» = Офис: «Провести» блока — только видимые строки (один OFF-00000001), скрытые не тронуты",
              vis == [True, False, True, False, True] and res.startswith("OK=3;ERR=0") and [s.spc(r)[0] for r in (2, 4)] == ["", ""]
              and {s.spc(r)[2] for r in (1, 3, 5)} == {"OFF-00000001"}, f"{vis}; {res.splitlines()[0]}")
        s.doc.store()
        vis1 = [visible(s, r) for r in range(1, 6)]
        with zipfile.ZipFile(p) as z:
            xml = z.read("content.xml").decode("utf-8")
        nfilt = xml.count('table:visibility="filter"')
        s.close()
        s = ready(Session(p))
        vis2 = [visible(s, r) for r in range(1, 6)]
        checks = {"CLEAN": st(s)[0] == "CLEAN", "не «изменена» после открытия": not s.doc.isModified(), "строки при фильтре": nfilt == 0 and vis1 == vis
                  and vis2 == vis, "статусы": [s.spc(r)[16] for r in (1, 3, 5)] == ["Проведено: новый ЕИ"] * 3,
                  "защита": [s.spc_locks(r) for r in (1, 2)] == [POSTED, OPEN], "вставка строк": xml.count('loext:insert-rows="true"') == 4,
                  "автофильтр": s.doc.DatabaseRanges.hasByName("WMS_SPECIAL"), "счётчики": counters(s)[1:] == [204, 4, 2, 1, 1, 1, 1],
                  "непроведённые на «Главной»": "иной приход: 2 строк" in s.main_status()[3]}
        set_filter(s, 1, "Производство")
        nxt = post_block(s, 1, 5, answer=1)
        checks["следующее проведение"] = nxt.startswith("OK=2") and s.spc(2)[0] == 4.0 and s.spc(4)[2] == "PROD-00000001"
        bad = [k for k, v in checks.items() if not v]
        R.add(c, "сохранение при фильтре (D-041) → закрыть → открыть: строки сохранены показанными, фильтр восстановлен; статусы, защита, опция «вставка строк», "
                 "автофильтр, счётчики целы; «Главная» показывает непроведённые строки; нумерация продолжается", not bad,
              f"не выполнено: {bad}; {s.main_status()[3][-80:]}")
        oracle(c, s, "фильтр, сохранение, повторное открытие")
    finally:
        s.close()


@case
def x26_protection_buttons_panel():
    c = "X26"
    s = ready(Session(new_wms("x26")))
    try:
        form = s.doc.Sheets.getByName("Иной приход").getDrawPage().getForms().getByIndex(0)
        names = {}
        for i in range(form.getCount()):
            ev = form.getScriptEvents(i)
            code = ev[0].ScriptCode if ev else ""
            names[form.getByIndex(i).Label] = code.split("Standard.", 1)[1].split("?")[0] if "Standard." in code else code
        want = {"Проверить": "WmsSpecialUi.BtnSpcCheck", "Провести": "WmsSpecialUi.BtnSpcPost", "Исправить": "WmsSpecialUi.BtnSpcFix",
                "Удалить": "WmsSpecialUi.BtnSpcDelete", "Разобрать": "WmsSpecialUi.BtnSpcIdentify", "Очистить": "WmsSpecialUi.BtnSpcClear"}
        R.add(c, "6 кнопок листа «Иной приход» привязаны к макросам", names == want, str(names))
        line(s, 1, "Офис", "Стол", "2")
        post(s, 1)
        before = s.spc(1)
        changed = []
        for col in "ABCDEFGHIJKLMNOPQR":
            s.type_spc(col, 1, "999")
            if s.spc(1) != before:
                changed.append(col)
        s.type_spc("S", 1, "примечание")
        open_ok = s.spc(1)[18] == "примечание"
        s.type_spc("B", 2, "Офис")
        before2 = s.spc(2)
        open_locked = []
        for col in "ANOPQR":
            s.type_spc(col, 2, "777")
            if s.spc(2) != before2:
                open_locked.append(col)
        s.goto("Иной приход", "$A$2")
        s.ui(".uno:SelectRow")
        s.ui(".uno:DeleteRows")
        s.goto("Иной приход", "$A$1:$S$3")
        s.ui(".uno:SortDescending")
        R.add(c, "защита: в проведённой строке закрыты A–R, открыт комментарий S; в непроведённой закрыты A и N–R; удаление строк и сортировка отклонены",
              not changed and open_ok and not open_locked and s.spc(1)[0] == 1.0 and s.spc_locks(1) == POSTED,
              f"изменились закрытые {changed}; S {open_ok}; в непроведённой {open_locked}")
        chk = s.click_spc("BtnSpcCheck", 1)
        line(s, 2, None, "Шкаф", "1")
        chk2 = s.click_spc("BtnSpcCheck", 2)
        R.add(c, "«Проверить»: проведённая строка — сведения о строке и ЕИ; непроведённая — результат проверки в «Контроль»",
              chk.startswith("OK:строка № 1 (действует): OFF-00000001, ЕИ-00000201 «Стол»") and "остаток ЕИ — 2" in chk and chk2.startswith("OK:можно провести")
              and ctl(s, 2).startswith("Можно провести: Новый ЕИ (Офис)"), f"{chk}; {chk2}")
        s.order_input(1, A="З-1", B="Стол", H="1", I="шт", L="Офис", F="1", N=D1, U="A-1")
        res = s.click_ord("BtnRcvPost", 1)
        R.add(c, "строка «Заказы» с поставщиком «Офис» не проводится обычным приходом — подсказка: лист «Иной приход»", res.startswith("ERR")
              and "проводится на листе «Иной приход»" in res and s.ord(1, "V") == "", res)
        order(s, 2, f="10")
        res2 = s.click_ord("BtnRcvPost", 2)
        R.add(c, "обычный приход «Заказы» не изменился: новый ЕИ, тип источника «Поставщик»", res2.startswith("OK") and s.card(202)[9] == "Поставщик"
              and s.card(202)[8] == "Заказ З-1", f"{res2}; {s.card(202)}")
        oracle(c, s, "защита, кнопки, «Заказы»")
    finally:
        s.close()


@case
def x27_dates_quantities():
    c = "X27"
    for loc in ("ru-RU", "en-US"):
        s = ready(Session(new_wms(f"x27_{loc}"), locale=loc))
        try:
            vals = ["1,5", "1.5", "0,125", "2,1250", "-1", "1e3", "1 000", "1,500", "01.02.2026", "0", "abc", "1,2345"]
            out = []
            for i, v in enumerate(vals, start=1):
                line(s, i, "Офис", f"Товар {i}", v)
                res = post(s, i)
                out.append((v, res[:2], s.spc(i)[5]))
            want_ok = {"1,5": 1.5, "1.5": 1.5, "0,125": 0.125, "2,1250": 2.125}
            ok = all((x[1] == "OK") == (x[0] in want_ok) and (x[0] not in want_ok or x[2] == want_ok[x[0]]) for x in out)
            R.add(c, f"локаль {loc}: F «1,5», «1.5», «0,125», «2,1250» — числа; «-1», «1e3», «1 000», «1,500», дата, 0, текст, 4 знака — отказ", ok, str(out))
            line(s, 13, "Офис", "Без даты", "1", date=None)
            d0 = post(s, 13)
            res_d = []
            for d in ("31.02.2026", "25.09.1999", "21-09-2026x"):
                s.type_spc("H", 13, d)
                res_d.append(post(s, 13))
            s.type_spc("H", 13, D1)
            d_ok = post(s, 13)
            R.add(c, f"локаль {loc}: дата поступления обязательна; 31.02, вне 2000–2099, не дд.мм.гггг — отказ; дд.мм.гггг принята и стала датой",
                  "не указана дата поступления (H)" in d0 and all(x.startswith("ERR") for x in res_d) and d_ok.startswith("OK")
                  and s.spc(13)[7] == serial(D1), f"{d0[:60]}; {[x[:50] for x in res_d]}; {d_ok}")
            oracle(c, s, f"количество и даты, локаль {loc}")
        finally:
            s.close()


@case
def x28_old_books_refused():
    c = "X28"
    p = new_wms("x28a")
    s = Session(p, macros=0)
    sy = s.doc.Sheets.getByName("_SYS")
    sy.unprotect(PWD)
    sy.getCellByPosition(1, 0).setString("WMS-SYS-1")
    sy.protect(PWD)
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        state, block = st(s)
        rep = s.report()
        line(s, 1, "Офис", "Стол", "1")
        res = s.X("SpecialPostRow", 1, "", False)
        R.add(c, "книга со схемой _SYS WMS-SYS-1 (ядро этапа 4 и раньше): запуск заблокирован с понятной причиной, проведение запрещено",
              state == "BLOCKED" and "WMS-SYS-1" in rep and "этапа 4" in rep and res.startswith("BLOCKED:") and s.spc(1)[0] == "",
              f"{state} {block}; {rep[:200]}; {res[:80]}")
    finally:
        s.close()
    p = new_wms("x28b")
    s = Session(p, macros=0)
    s.doc.unprotect(PWD)
    for name in ("Иной приход", "_SPR", "_ART"):
        s.doc.Sheets.removeByName(name)
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        state, block = st(s)
        rep = s.report()
        R.add(c, "книга без листов «Иной приход», _SPR, _ART: запуск заблокирован «нет листа «Иной приход»»", state == "BLOCKED"
              and "нет листа «Иной приход»" in rep, f"{state} {block}; {rep[:160]}")
    finally:
        s.close()


@case
def x29_selfcheck_block_system_error():
    c = "X29"
    s = ready(Session(new_wms("x29")))
    try:
        line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
        line(s, 2, "Офис", "Стол", "1")
        post(s, 1)
        post(s, 2)
        sc0 = s.B("ActionSelfCheck").splitlines()
        R.add(c, "самопроверка: строка «специальные приходы: … расхождений нет», ошибок 0", sc0[0].startswith("САМОПРОВЕРКА WMS: ошибок 0")
              and any("специальные приходы:" in x and "расхождений нет" in x for x in sc0), [x for x in sc0 if "специальные" in x][:1])
        stk = s.doc.Sheets.getByName("Наличие")
        stk.unprotect(PWD)
        stk.getCellByPosition(4, 202).setValue(5)
        sc1 = s.B("ActionSelfCheck")
        stk.getCellByPosition(4, 202).setValue(1)
        stk.protect(PWD)
        art = s.doc.Sheets.getByName("_ART")
        art.unprotect(PWD)
        art.getCellRangeByPosition(0, 1, 2, 1).clearContents(1 + 2 + 4 + 16)
        art.protect(PWD)
        sc2 = s.B("ActionSelfCheck")
        rb = s.X("ArticleIndexRebuild")
        sc3 = s.B("ActionSelfCheck")
        R.add(c, "самопроверка находит повреждения: остаток ЕИ выше суммы действующих приходов; деталь не записана в индексе; перестройка индекса лечит",
              "больше суммы действующих приходов" in sc1 and "деталь не записана в индексе артикулов" in sc2 and "записей 1" in rb
              and sc3.startswith("САМОПРОВЕРКА WMS: ошибок 0"), f"{[x for x in sc1.splitlines() if 'специальные' in x][:1]}; {rb}")
        for r in (3, 4, 5):
            line(s, r, "Офис", f"Товар {r}", "1")
        sh = s.doc.Sheets.getByName("Иной приход")
        sh.unprotect(PWD)
        sh.getCellByPosition(0, 12).setValue(4)          # the № the counter will give next but one is already on the sheet
        sh.protect(PWD)
        nj = njournal(s)
        head, _, sysl = post_block(s, 3, 5, answer=2).split("\n", 2)
        msg = s.XU("TestSpcBlockMessage")
        R.add(c, "системная ошибка WMS посреди блока (№ строки, который выдал бы NEXT_SPL, уже на листе): строка 4 проведена, на строке 5 обработка "
                 "остановлена, строка 6 не обрабатывалась", head == "OK=1;ERR=0;SKIP=0;STOP=5" and "ERR-SYS:" in sysl and s.spc(3)[0] == 3.0
              and s.spc(4)[0] == "" and s.spc(5)[0] == "" and njournal(s) == nj + 1 and "ОБРАБОТКА ОСТАНОВЛЕНА" in msg, f"{head}; {sysl[:140]}")
        cl = s.click_spc("BtnSpcClear", 12)
        head3 = post_block(s, 3, 5, answer=2).split("\n")[0]
        R.add(c, "причина устранена (строка с чужим № очищена): тот же блок — проведённая строка пропущена, две проведены",
              cl.startswith("OK") and head3 == "OK=2;ERR=0;SKIP=1;STOP=0" and [s.spc(r)[0] for r in (3, 4, 5)] == [3.0, 4.0, 5.0], f"{cl}; {head3}")
        oracle(c, s, "самопроверка и системная ошибка блока")
    finally:
        s.close()


# ================================================================ X30 mixed sequence of 100+ operations

@case
def x30_mixed_sequence():
    c = "X30"
    rnd = random.Random(20260927)
    s = ready(Session(new_wms("x30")))
    try:
        m = dict(bal={}, unit={}, parts={}, lines={}, issues={}, returns={}, review=[])
        errs, kinds = [], {}
        srow, irow, rrow, orow = 1, 1, 1, 1
        next_ei = 201
        who = ["ива", "пет", "егр"]
        full = {"ива": IVA, "пет": PET, "егр": EGR}
        arts = ["ABC-100", "ABC-200", "QWE-1", "QWE-2", "ZX-9", "DT-77"]
        names = ["Стол", "Стул", "Шкаф", "Корпус", "Вал", "Дрель", "Коробка", "Кабель"]
        series, n_ops = 0, 0

        def ok(kind, res):
            kinds[kind] = kinds.get(kind, 0) + (1 if res.startswith("OK") else 0)
            return res.startswith("OK")

        def new_line(n, e, q, code, mode):
            m["lines"][n] = dict(ei=e, qty=q, live=True, code=code, mode=mode, row=srow)

        for step in range(150):
            roll = rnd.random()
            live_iss = [k for k, v in m["issues"].items() if not v["del"]]
            live_ret = [k for k, v in m["returns"].items() if v["live"]]
            live_lines = [n for n, v in m["lines"].items() if v["live"]]
            if roll < 0.06:
                # an ordinary receipt of an order («Заказы»)
                q = rnd.randint(5, 30)
                order(s, orow, a=f"З-{orow}", b=f"Заказной товар {orow}", h=str(q), f=str(q))
                res = s.click_ord("BtnRcvPost", orow)
                if ok("RECEIPT", res):
                    m["bal"][ei(next_ei)] = float(q)
                    m["unit"][ei(next_ei)] = "шт"
                    next_ei += 1
                else:
                    errs.append(f"шаг {step}: приход заказа отклонён: {res[:80]}")
                orow += 1
            elif roll < 0.36:
                typ = rnd.choice(["Офис", "Производство", "Детали", "Детали", "Старый склад", "Иной"])
                code = {"Офис": "OFF", "Производство": "PROD", "Детали": "DET", "Старый склад": "OLD", "Иной": "OTH"}[typ]
                q = rnd.randint(1, 40)
                if typ == "Детали":
                    a = rnd.choice(arts)
                    known = a in m["parts"]
                    # a known part: the place stays empty — the preview fills the current place of its EI (no move)
                    line(s, srow, typ, rnd.choice(names), str(q), art=a, place=None if known else "D-" + a[-1], unit="шт")
                    res = post(s, srow)
                    if ok("SP_DET", res):
                        n = int(s.spc(srow)[0])
                        if known:
                            e = m["parts"][a]
                            m["bal"][e] += q
                            new_line(n, e, float(q), code, "ADD")
                        else:
                            e = ei(next_ei)
                            next_ei += 1
                            m["parts"][a] = e
                            m["bal"][e] = float(q)
                            m["unit"][e] = "шт"
                            new_line(n, e, float(q), code, "NEW")
                        if s.spc(srow)[13] != e:
                            errs.append(f"шаг {step}: деталь {a} проведена в {s.spc(srow)[13]}, ожидался {e}")
                    else:
                        errs.append(f"шаг {step}: приход детали {a} отклонён: {res[:100]}")
                    srow += 1
                elif rnd.random() < 0.3:
                    # one acceptance of two or three lines
                    k = rnd.randint(2, 3)
                    for j in range(k):
                        line(s, srow + j, typ, f"{rnd.choice(names)} {step}-{j}", str(rnd.randint(1, 20)), place="P-" + str(j))
                    res = post_block(s, srow, srow + k - 1, answer=1)
                    head = res.split("\n")[0]
                    if head.startswith(f"OK={k};ERR=0"):
                        evs = {s.spc(srow + j)[2] for j in range(k)}
                        if len(evs) != 1:
                            errs.append(f"шаг {step}: одно поступление получило несколько №: {evs}")
                        for j in range(k):
                            row = s.spc(srow + j)
                            e = row[13]
                            m["bal"][e] = row[5]
                            m["unit"][e] = "шт"
                            new_line(int(row[0]), e, row[5], code, "NEW")
                            m["lines"][int(row[0])]["row"] = srow + j
                            if code == "OTH":
                                m["review"].append(int(row[0]))
                            next_ei += 1
                        kinds["SP_GROUP"] = kinds.get("SP_GROUP", 0) + k
                    else:
                        errs.append(f"шаг {step}: блок {typ} {head}")
                    srow += k
                else:
                    line(s, srow, typ, f"{rnd.choice(names)} {step}", str(q), place="S-" + str(step % 7), mark="инв. " + str(step) if code == "OLD" else None)
                    res = post(s, srow)
                    if ok("SP_" + code, res):
                        e = ei(next_ei)
                        next_ei += 1
                        m["bal"][e] = float(q)
                        m["unit"][e] = "шт"
                        n = int(s.spc(srow)[0])
                        new_line(n, e, float(q), code, "NEW")
                        if code == "OTH":
                            m["review"].append(n)
                        if s.spc(srow)[13] != e:
                            errs.append(f"шаг {step}: {typ} получил {s.spc(srow)[13]}, ожидался {e}")
                    else:
                        errs.append(f"шаг {step}: {typ} отклонён: {res[:100]}")
                    srow += 1
            elif roll < 0.56:
                cand = [e for e, b in m["bal"].items() if b >= 1]
                if not cand:
                    continue
                e = rnd.choice(cand)
                q = rnd.randint(1, min(8, int(m["bal"][e])))
                w = rnd.choice(who)
                res = issue(s, irow, e[3:].lstrip("0"), str(q), who=w)
                if ok("ISSUE", res):
                    k = len(m["issues"]) + 1
                    m["issues"][k] = {"ei": e, "qty": float(q), "ret": 0.0, "row": irow, "who": full[w], "del": False}
                    m["bal"][e] -= q
                else:
                    errs.append(f"шаг {step}: выдача {q} из {e} (остаток {m['bal'][e]}) отклонена: {res[:80]}")
                irow += 1
            elif roll < 0.68 and live_iss:
                k = rnd.choice(live_iss)
                iss = m["issues"][k]
                avail = iss["qty"] - iss["ret"]
                if avail < 1:
                    continue
                q = rnd.randint(1, int(avail))
                res = ret(s, rrow, str(k), str(q))
                if ok("RETURN", res):
                    n = len(m["returns"]) + 1
                    m["returns"][n] = dict(issue=k, qty=float(q), live=True, row=rrow)
                    iss["ret"] += q
                    m["bal"][iss["ei"]] += q
                else:
                    errs.append(f"шаг {step}: возврат {q} по №{k} (можно {avail}) отклонён: {res[:80]}")
                rrow += 1
            elif roll < 0.78 and live_lines:
                n = rnd.choice(live_lines)
                L = m["lines"][n]
                q2 = max(1, int(L["qty"]) + rnd.randint(-5, 5))
                res = fix(s, L["row"], qty=str(q2))
                allowed = m["bal"][L["ei"]] - L["qty"] + q2 >= 0
                if res.startswith("OK"):
                    kinds["SP_FIX"] = kinds.get("SP_FIX", 0) + 1
                    m["bal"][L["ei"]] += q2 - L["qty"]
                    L["qty"] = float(q2)
                    if not allowed:
                        errs.append(f"шаг {step}: исправление строки № {n} до {q2} проведено, хотя остаток стал бы отрицательным")
                elif allowed and not res.startswith("SKIP"):
                    errs.append(f"шаг {step}: исправление строки № {n} до {q2} отклонено: {res[:100]}")
            elif roll < 0.86 and live_lines:
                n = rnd.choice(live_lines)
                L = m["lines"][n]
                e = L["ei"]
                others = [x for x, v in m["lines"].items() if v["live"] and v["ei"] == e and x != n]
                deps = [k for k, v in m["issues"].items() if not v["del"] and v["ei"] == e] + \
                       [k for k, v in m["returns"].items() if v["live"] and m["issues"][v["issue"]]["ei"] == e]
                allowed = m["bal"][e] >= L["qty"] - 1e-9 and (others or not deps)
                res = delete(s, L["row"])
                if res.startswith("OK"):
                    kinds["SP_DEL"] = kinds.get("SP_DEL", 0) + 1
                    L["live"] = False
                    m["bal"][e] -= L["qty"]
                    if not allowed:
                        errs.append(f"шаг {step}: сторно строки № {n} проведено, хотя должно быть запрещено")
                elif allowed:
                    errs.append(f"шаг {step}: сторно строки № {n} отклонено: {res[:100]}")
            elif roll < 0.90 and m["review"]:
                n = m["review"].pop(0)
                L = m["lines"][n]
                if not L["live"]:
                    continue
                typ = rnd.choice(["Офис", "Поставщик", "Производство", "Старый склад", "Детали"])
                a = f"ID-{n}" if typ == "Детали" else ""
                res = ident(s, L["row"], typ, f"Разобранный товар {n}", a, "Разное")
                if ok("SP_IDENTIFY", res):
                    if typ == "Детали":
                        m["parts"][a] = L["ei"]
                        arts.append(a)
                else:
                    errs.append(f"шаг {step}: разбор строки № {n} отклонён: {res[:100]}")
            elif roll < 0.95 and live_ret:
                n = rnd.choice(live_ret)
                rv = m["returns"][n]
                iss = m["issues"][rv["issue"]]
                res = s.click_ret("BtnRetDelete", rv["row"])
                if res.startswith("OK"):
                    kinds["RETURN_DEL"] = kinds.get("RETURN_DEL", 0) + 1
                    rv["live"] = False
                    iss["ret"] -= rv["qty"]
                    m["bal"][iss["ei"]] -= rv["qty"]
                elif m["bal"][iss["ei"]] >= rv["qty"]:
                    errs.append(f"шаг {step}: сторно возврата № {n} отклонено: {res[:80]}")
            elif live_iss:
                k = rnd.choice(live_iss)
                iss = m["issues"][k]
                if iss["ret"] > 0:
                    continue
                res = s.I("IssueDeleteRow", iss["row"])
                if ok("ISSUE_DEL", res):
                    iss["del"] = True
                    m["bal"][iss["ei"]] += iss["qty"]
                else:
                    errs.append(f"шаг {step}: сторно выдачи № {k} отклонено: {res[:80]}")
            if (step + 1) % 30 == 0:
                series += 1
                P, inf, _ = s.spcheck()
                R.add(c, f"оракул после серии {series} (шаги {step - 28}…{step + 1})", not P, f"{P[:3]}; операций {inf.get('ops')}")
        n_ops = sum(kinds.values())
        bal_ok = [e for e, b in m["bal"].items() if abs(s.stock(int(e[3:])) - b) > 1e-6]
        sc = s.B("ActionSelfCheck")
        R.add(c, f"{n_ops} проведённых операций подряд (обычные приходы, офис, производство, детали, старый склад, иной приход, выдачи, возвраты, "
                 "исправления, сторно, разбор): модель теста и книга совпадают, неожиданных отказов нет, полная самопроверка без ошибок",
              n_ops >= 100 and not errs and not bal_ok and sc.splitlines()[0].startswith("САМОПРОВЕРКА WMS: ошибок 0"),
              f"операций {n_ops}: {kinds}; отказы {errs[:3]}; расхождения {bal_ok[:3]}; {sc.splitlines()[0]}")
        oracle(c, s, "смешанная последовательность")
    finally:
        s.close()


# ================================================================ X31 migration dry-run (validators, no writes)

@case
def x31_migration_dryrun():
    import hashlib
    import json
    c = "X31"
    p = new_wms("x31")
    s = ready(Session(p))
    line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
    line(s, 2, "Офис", "Стол", "1", place="A-1")
    post(s, 1)
    post(s, 2)
    s.close(save=True)
    d = os.path.dirname(p)
    rows = ["ЕИ;Наименование;Артикул;Единица;Количество;Место;Категория;Тип источника;Старая маркировка;Комментарий",
            "ЕИ-00000501;Дрель ударная;;шт;1;A-1;Инструмент;Старый склад;инв. 12;",
            "EI-502;Шестерня малая;QWE-5;шт;40;D-1;Запчасти;Детали;;",
            "503;Шестерня малая;qwe‑5;шт;5;D-1;Запчасти;Детали;;",
            "00000501;Болт;;шт;10;A-1;Крепёж;;;",
            "ЕИ-00000201;Шестерня;ABC-100;шт;3;D-1;;Детали;;",
            "600;Шестерня;abc-100;шт;3;D-1;;Детали;;",
            ";Клей;;тюбик;2;Z-99;Химия;;;",
            ";Краска;;шт;-1;A-1;;;;",
            ";Лак;;шт;1,500;A-1;;;;",
            ";;;шт;1;A-1;;;;",
            ";Гайка;;;5;A-1;;;;",
            ";Подшипник;;шт;2;A-1;;Детали;;",
            "x-77;Пружина;;шт;2;A-1;;Склад;;"]
    src = os.path.join(d, "migration_input.csv")
    open(src, "w", encoding="utf-8").write("\n".join(rows) + "\n")
    h0 = hashlib.sha256(open(p, "rb").read()).hexdigest()
    tool = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools", "migration_dryrun.py")
    rj = os.path.join(d, "dryrun.json")
    run = subprocess.run([sys.executable, tool, src, "--book", p, "--json", rj, "--report", os.path.join(d, "dryrun.md")], capture_output=True,
                         text=True, timeout=300)
    h1 = hashlib.sha256(open(p, "rb").read()).hexdigest()
    out = json.load(open(rj, encoding="utf-8")) if os.path.exists(rj) else {}
    kinds = {}
    for f in out.get("findings", []):
        kinds.setdefault(f["kind"], []).append(f["row"])
    sm = out.get("summary", {})
    # D-087: an empty «Тип источника» is a blocking finding (rows 5, 8–12), an unknown one too (row 14)
    want = {"дубль ЕИ": [5], "ЕИ уже в книге": [6], "неизвестная единица": [8], "неизвестное место": [8], "количество": [9, 10],
            "тип источника": [5, 8, 9, 10, 11, 12, 14], "ЕИ": [14]}
    got = {k: sorted(v) for k, v in kinds.items() if k in want}
    art = [f["text"] for f in out.get("findings", []) if f["kind"] == "артикул → несколько ЕИ"]
    keydata = sorted(kinds.get("нет ключевых данных", []))
    R.add(c, "dry-run переноса: дубли ЕИ, ЕИ уже в книге, артикул детали у нескольких ЕИ (в таблице и против детали книги ABC-100 → ЕИ-00000201), "
             "неизвестные единица и место, отрицательное и неоднозначное количество, строки без наименования / единицы / артикула детали, "
             "неизвестный тип и неверный ЕИ; максимальный ЕИ и нужный NEXT_EI; код выхода 1",
          run.returncode == 1 and got == want and keydata == [11, 12, 13] and any("«QWE-5»" in x for x in art)
          and any("«ABC-100»" in x and "ЕИ-00000201" in x for x in art) and sm.get("max_ei_input") == "ЕИ-00000600" and sm.get("next_ei_needed") == 601,
          f"код {run.returncode}; {got}; ключевые {keydata}; артикулы {[x[:70] for x in art]}; {sm}; {run.stderr[-200:]}")
    R.add(c, "dry-run не меняет книгу: открыта только для чтения, без макросов, файл до и после совпадает", h0 == h1 and out.get("book_unchanged") is True,
          f"{h0[:12]} {h1[:12]}")
    rj2 = os.path.join(d, "dryrun_default.json")
    run3 = subprocess.run([sys.executable, tool, src, "--default-source", "Старый склад", "--json", rj2], capture_output=True, text=True, timeout=300)
    out2 = json.load(open(rj2, encoding="utf-8")) if os.path.exists(rj2) else {}
    types2 = sorted(f["row"] for f in out2.get("findings", []) if f["kind"] == "тип источника")
    run4 = subprocess.run([sys.executable, tool, src, "--default-source", "Склад"], capture_output=True, text=True, timeout=300)
    R.add(c, "пустой тип источника не считается «Старым складом» молча (D-087): только явное --default-source «Старый склад» заполняет пустые строки; "
             "неизвестное значение параметра — отказ (код 2)", types2 == [14] and out2.get("summary", {}).get("default_source") == "Старый склад"
          and run4.returncode == 2 and "не тип источника" in run4.stderr, f"{types2}; {out2.get('summary', {}).get('default_source')}; {run4.returncode}")
    clean = os.path.join(d, "migration_clean.csv")
    open(clean, "w", encoding="utf-8").write("\n".join(rows[:2] + [";Кабель;;м;12,5;A-1;Кабель;Офис;;"]) + "\n")
    run2 = subprocess.run([sys.executable, tool, clean, "--book", p], capture_output=True, text=True, timeout=300)
    R.add(c, "чистая таблица: без блокирующих замечаний — код выхода 0", run2.returncode == 0 and "блокирующих замечаний | 0" in run2.stdout,
          run2.stdout[-300:])
    s = ready(Session(p))
    try:
        oracle(c, s, "книга после dry-run")
    finally:
        s.close()


@case
def x32_details_all_storno_then_again():
    c = "X32"
    s = ready(Session(new_wms("x32")))
    try:
        line(s, 1, "Детали", "Шестерня", "100", art="ABC-100", place="D-1")
        line(s, 2, "Детали", None, "50", art="ABC-100", unit=None, place=None)
        post(s, 1)
        post(s, 2)
        d2, d1 = delete(s, 2), delete(s, 1)
        card = s.card(201)
        R.add(c, "все приходы детали удалены (сторно): остаток 0, ЕИ «Приход удалён (сторно)», запись индекса артикула сохраняется",
              d2.startswith("OK") and d1.startswith("OK") and card[4] == 0.0 and card[7] == "Приход удалён (сторно)"
              and s.art_index() == [("ABC-100", ei(201), "ABC-100")], f"{d2}; {d1}; {card}")
        line(s, 3, "Детали", None, "30", art="abc-100", unit=None, place=None)
        pre = ctl(s, 3)
        p3 = post(s, 3)
        R.add(c, "новый приход того же артикула — снова тот же ЕИ-00000201 (новый ЕИ не создан): остаток 30, ЕИ снова «Активен»",
              "снова станет активным" in pre and p3.startswith("OK") and s.spc(3)[13] == ei(201) and s.card(201)[4] == 30.0
              and s.card(201)[7] == "Активен" and s.sysv("NEXT_EI") == 202.0, f"{pre[-80:]}; {p3}; {s.card(201)}")
        oracle(c, s, "повторный приход после сторно всех приходов детали")
    finally:
        s.close()


# ================================================================ runner

def main():
    only = set(a.lower() for a in sys.argv[1:])
    t0 = time.time()
    template("TEST")
    for fn in CASES:
        if only and not any(fn.__name__.startswith(o) for o in only):
            continue
        try:
            fn()
        except Exception as e:
            R.error(fn.__name__, e)
        subprocess.run(["pkill", "-9", "-f", "profile_.*" + str(os.getpid())], capture_output=True)
    R.save(os.path.join(OUT, "results_phase5.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_phase5.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
