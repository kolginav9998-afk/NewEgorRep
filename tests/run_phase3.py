"""Core Phase 3 «Приход и заказы» — automated scenarios (CLAUDE_TASK.md, section «ТЕСТЫ»).

    WMS_TEST_OUT=/tmp/wms_out python3 tests/run_phase3.py [case ...]

Every case copies a freshly built test book (synthetic EIs 1–200, NEXT_EI 201) into WMS_TEST_OUT/cases/<case>/ and drives
real LibreOffice processes like a user: typing through the UI (the sheet's change handler runs), the macros the buttons of
«Заказы» are bound to for the row under the cursor (confirmations answered «Да» in the test mode, dialog values preset),
window close, kill -9. The book and the journal are checked by the independent oracle tests/receipt_oracle.py.
Results: WMS_TEST_OUT/results_phase3.json and WMS_TEST_OUT/TEST_REPORT_phase3.md.
"""
import datetime
import os
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
POSTED = "1111111111010111001011111001"
OPEN = "0000000000000000000001111001"
TODAY = datetime.date.today()
PAST = (TODAY - datetime.timedelta(days=10)).strftime("%d.%m.%Y")
FUTURE = (TODAY + datetime.timedelta(days=20)).strftime("%d.%m.%Y")
D0 = "25.09.2026"
KEEP = "<как предложено>"          # WmsOrdersUi test seam: the dialog field keeps the value the window proposed
PWD = "wms"


def case(fn):
    CASES.append(fn)
    return fn


def st(s):
    d = s.state()
    return d.get("STATE"), d.get("BLOCK")


def ei(n):
    return f"ЕИ-{n:08d}"


def ready(s):
    s.U("TestUiAuto", 1)
    return s


def njournal(s):
    return len(s.journal()[0])


def oracle(c, s, name, expect_tail=0, today=None):
    P, inf = s.rcheck(expect_tail, today=today)
    R.add(c, "оракул: " + name, not P, f"{inf}; {P[:4]}")
    return not P


def order(s, r, a="З-1", b="Болт М10", h="100", i="шт", l="Петрович", d="СЧ-1", e="ART-10", p="20.09.2026", q=None, cat=None):
    s.order_input(r, A=a, B=b, D=d, E=e, H=h, I=i, L=l, P=p, Q=q, S=cat)


def fact(s, r, f, n=D0, u="A-1", c="УПД-1", g=None, o="24.09.2026", j=None):
    s.order_input(r, F=f, G=g, C=c, O=o, N=n, U=u, J=j)


def post(s, r):
    return s.click_ord("BtnRcvPost", r)


def more(s, r, f, c="", g="", o="", n=D0, u="A-2", j=""):
    s.OU("TestRcvInput", f, g, c, o, n, u, j)
    return s.click_ord("BtnRcvMore", r)


def fix(s, r, f, g="", c="", o="", n=D0, u="A-1", j=""):
    s.OU("TestRcvInput", f, g, c, o, n, u, j)
    return s.click_ord("BtnRcvFix", r)


def W(s, r):
    return s.ord(r, "W")


def snap(s, rows=10, eis=(201, 202, 203, 204)):
    sh = s.doc.Sheets.getByName("Заказы")
    vals = sh.getCellRangeByPosition(0, 1, 27, rows).getDataArray()
    locks = [s.ord_locks(r) for r in range(1, rows + 1)]
    stock = [s.stock_row(n) for n in eis]
    rcv = [s.rcv(n) for n in eis]
    ordp = [s.opos(k) for k in range(1, 5)]
    sysv = [s.sysv(k) for k in ("LAST_SEQ", "NEXT_EI", "NEXT_NO", "TX_STATE", "TX_SEQ", "TX_BEFORE_IMAGE")]
    return vals, locks, stock, rcv, ordp, sysv, s.next_ol()


def ctrl_z_harmless(s, n=5):
    a = snap(s)
    for _ in range(n):
        s.ui(".uno:Undo")
    return a == snap(s)


def issue(s, r, e, q, who="егр"):
    s.issue_input(r, ei=e, qty=q, date=D0, who=who)
    return s.post(r)


# ================================================================ R01–R08 receipts

@case
def r01_full_receipt():
    c = "R01"
    s = ready(Session(new_wms("r01")))
    try:
        order(s, 1, q=FUTURE)
        pv = (W(s, 1), s.ord(1, "Y"), s.ord(1, "V"))
        fact(s, 1, "100", g="100")
        res = post(s, 1)
        row, locks, stt = s.ords(1), s.ord_locks(1), s.state()
        ents = s.journal()[0]
        e = ents[-1] if ents else {}
        f = e.get("fields", {})
        ok = (pv == ("Ожидается", "", "") and res == "OK:1" and row[21] == ei(201) and row[22] == "Получено" and row[23] == 100.0
              and row[24] == "Норма; позиция: получено 100 из 100 шт" and row[5] == 100.0 and row[6] == 100.0 and row[7] == 100.0
              and locks == POSTED and s.sysv("NEXT_EI") == 202 and stt["LAST_SEQ"] == "1" and stt["UNDO_CAN"] == "False"
              and e.get("type") == "RECEIPT" and f.get("EI") == ei(201) and f.get("QTY") == "100" and f.get("ORD_QTY") == "100"
              and f.get("DOC") == "УПД-1" and f.get("DATE") == "2026-09-25" and f.get("DOC_DATE") == "2026-09-24")
        R.add(c, "полный обычный приход: новый ЕИ-00000201 в V, W «Получено», X 100, Y «Норма», H не изменилось, строка защищена, "
                 "NEXT_EI 202, журнал RECEIPT, без окон", ok, f"предпросмотр {pv}; {res}; {row}; защита {locks}; {e.get('raw', '')[:200]}")
        reg = s.stock_row(201)
        R.add(c, "в той же операции — строка «Наличие» ЕИ-00000201: название, артикул, единица, остаток 100, место, состояние, источник «Заказ З-1»",
              reg == (ei(201), "Болт М10", "ART-10", "шт", 100.0, "A-1", "", "Активен", "Заказ З-1"), str(reg))
        oracle(c, s, "полный приход")
    finally:
        s.close()


@case
def r02_partial_receipt():
    c = "R02"
    s = ready(Session(new_wms("r02")))
    try:
        order(s, 1, q=FUTURE)
        fact(s, 1, "40", g="40")
        res = post(s, 1)
        row = s.ords(1)
        R.add(c, "частичный приход 40 из 100: «Частично получено», X 40 (остаток этого ЕИ), осталось 60, H = 100",
              res == "OK:1" and row[22] == "Частично получено" and row[23] == 40.0 and "осталось 60" in row[24] and row[7] == 100.0
              and s.opos(1)[4:8] == (100.0, 40.0, 1.0, 0.0), f"{res}; {row[20:25]}; позиция {s.opos(1)}")
        oracle(c, s, "частичный приход")
    finally:
        s.close()


@case
def r03_second_third_upd():
    c = "R03"
    s = ready(Session(new_wms("r03")))
    try:
        order(s, 1, q=FUTURE)
        fact(s, 1, "40", g="40")
        post(s, 1)
        m2 = more(s, 1, "30", c="УПД-2", g="30", o="26.09.2026", n="26.09.2026", u="A-2")
        src2, add2 = s.ords(1), s.ords(2)
        ok2 = (m2.startswith("OK:2|строка 3|" + ei(202)) and add2[0] == "З-1" and add2[3] == "СЧ-1" and add2[1] == "Болт М10" and add2[4] == "ART-10"
               and add2[8] == "шт" and add2[11] == "Петрович" and add2[2] == "УПД-2" and add2[5] == 30.0 and add2[6] == 30.0 and add2[7] == ""
               and add2[13] == 46291.0 and add2[21] == ei(202) and add2[22] == "Дополнительное поступление" and add2[23] == 30.0
               and s.ord_locks(2) == POSTED and src2[22] == "Частично получено" and "получено 70 из 100" in src2[24] and src2[23] == 40.0)
        R.add(c, "второй УПД («Ещё поступление»): новая строка — тот же № заказа и счёта, B/E/I/L скопированы, H пустая, новый ЕИ, "
                 "свой статус; исходная строка: «Частично получено», 70 из 100", ok2, f"{m2}; доп. строка {add2}; исходная W/X/Y {src2[22:25]}")
        m3 = more(s, 2, "30", c="УПД-3", g="30", o="27.09.2026", n="27.09.2026", u="A-3")
        src3 = s.ords(1)
        hs = [s.ord(r, "H") for r in range(1, 4)]
        R.add(c, "третий УПД из дополнительной строки: позиция «Получено» (40+30+30 = 100), H только в исходной строке (Σ H = 100), "
                 "три разных ЕИ",
              m3.startswith("OK:3|строка 4|" + ei(203)) and src3[22] == "Получено" and hs == [100.0, "", ""] and s.opos(1)[4:8] == (100.0, 100.0, 3.0, 0.0)
              and [s.ord(r, "V") for r in range(1, 4)] == [ei(201), ei(202), ei(203)], f"{m3}; {src3[22:25]}; H {hs}; позиция {s.opos(1)}")
        oracle(c, s, "три УПД")
    finally:
        s.close()


@case
def r04_positions_of_one_order():
    c = "R04"
    s = ready(Session(new_wms("r04")))
    try:
        order(s, 1, a="З-5", b="Кабель", e="K-1", h="50", i="м", q=FUTURE)
        order(s, 2, a="З-5", b="Розетка", e="R-2", h="20", i="шт", q=FUTURE)
        order(s, 3, a="З-5", b="Выключатель", e="V-3", h="10", i="шт", q=FUTURE)
        fact(s, 1, "50")
        fact(s, 2, "5")
        a, b = post(s, 1), post(s, 2)
        m = more(s, 2, "10", c="УПД-7", o=D0)
        # M6 §18: the delivery row goes right under its position (row 4 of the sheet), the third position moves down to row 5
        ws = [W(s, r) for r in (1, 2, 4)]
        ok = (a == "OK:1" and b == "OK:2" and m.startswith("OK:3|строка 4") and ws == ["Получено", "Частично получено", "Ожидается"]
              and s.opos(1)[4:6] == (50.0, 50.0) and s.opos(2)[4:6] == (20.0, 15.0) and s.next_ol() == 3.0 and s.ord(3, "B") == "Розетка"
              and s.ord(3, "A") == "З-5" and W(s, 3) == "Дополнительное поступление" and s.ord(4, "B") == "Выключатель")
        R.add(c, "один № заказа — три разные позиции: статусы независимы (Получено / Частично получено / Ожидается), «Ещё поступление» "
                 "относится только к своей позиции", ok, f"{a}; {b}; {m}; W {ws}; позиции {s.opos(1)} {s.opos(2)}")
        oracle(c, s, "позиции одного заказа")
    finally:
        s.close()


@case
def r05_same_name_and_article():
    c = "R05"
    s = ready(Session(new_wms("r05")))
    try:
        order(s, 1, a="З-1", b="Перчатки", e="PN-9", h="10", i="пар")
        order(s, 2, a="З-2", b="Перчатки", e="PN-9", h="10", i="пар")
        fact(s, 1, "10", c="УПД-1")
        fact(s, 2, "10", c="УПД-2")
        a, b = post(s, 1), post(s, 2)
        r1, r2 = s.stock_row(201), s.stock_row(202)
        R.add(c, "одинаковые название и артикул в разных партиях: два разных ЕИ, записи реестра не объединяются, поиск «такой товар уже есть» не выполняется",
              a == "OK:1" and b == "OK:2" and s.ord(1, "V") == ei(201) and s.ord(2, "V") == ei(202) and r1[1:4] == r2[1:4] == ("Перчатки", "PN-9", "пар")
              and r1[4] == r2[4] == 10.0 and s.stock_row(4)[4] == 12.0, f"{a} {b}; {r1}; {r2}; синтетический ЕИ-4 {s.stock_row(4)}")
        oracle(c, s, "одинаковые товары")
    finally:
        s.close()


@case
def r06_empty_article():
    c = "R06"
    s = ready(Session(new_wms("r06")))
    try:
        s.order_input(1, A="З-9", B="Скотч малярный", H="5", I="шт", L="Озон")
        fact(s, 1, "5", c="", o=None)
        res = post(s, 1)
        R.add(c, "пустой артикул: приход проводится, ЕИ создан, артикул в реестре пуст",
              res == "OK:1" and s.stock_row(201)[1:3] == ("Скотч малярный", "") and s.ord(1, "V") == ei(201), f"{res}; {s.stock_row(201)}")
        oracle(c, s, "пустой артикул")
    finally:
        s.close()


@case
def r07_no_documents_then_added():
    c = "R07"
    s = ready(Session(new_wms("r07")))
    try:
        order(s, 1, h="10")
        s.order_input(1, F="10", N=D0, U="B-1")
        res = post(s, 1)
        row = s.ords(1)
        ok1 = (res == "OK:1" and row[22] == "Получено без документов" and "нет документа (C, O)" in row[24] and s.stock(201) == 10.0
               and s.opos(1)[7] == 1.0)
        R.add(c, "получено без документов (C и O пусты): ЕИ создан, остаток доступен, статус показывает отсутствие документа", ok1,
              f"{res}; {row[20:25]}")
        iss = issue(s, 1, "201", "2")
        R.add(c, "остаток прихода без документов доступен для выдачи", iss == "OK:2" and s.stock(201) == 8.0 and s.ord(1, "X") == 8.0,
              f"{iss}; остаток {s.stock(201)}; X {s.ord(1, 'X')}")
        nx = s.sysv("NEXT_EI")
        fx = fix(s, 1, "10", g="10", c="УПД-77", o="24.09.2026", u="B-1")
        row = s.ords(1)
        R.add(c, "документ добавлен позднее через «Исправить»: тот же ЕИ, без нового прихода, NEXT_EI не изменился, статус «Получено»",
              fx.startswith("OK:") and row[21] == ei(201) and row[22] == "Получено" and row[2] == "УПД-77" and s.sysv("NEXT_EI") == nx
              and s.stock(201) == 8.0 and s.opos(1)[7] == 0.0 and s.journal()[0][-1]["type"] == "RECEIPT_FIX", f"{fx}; {row[20:25]}")
        oracle(c, s, "документ позднее")
    finally:
        s.close()


@case
def r08_shortage_surplus():
    c = "R08"
    s = ready(Session(new_wms("r08")))
    try:
        order(s, 1, a="З-1", h="10")
        order(s, 2, a="З-2", h="10")
        order(s, 3, a="З-3", h="10")
        fact(s, 1, "8", g="10", c="УПД-1")
        fact(s, 2, "12", g="10", c="УПД-2")
        fact(s, 3, "10", g="10", c="УПД-3")
        res = [post(s, r) for r in (1, 2, 3)]
        y = [s.ord(r, "Y") for r in (1, 2, 3)]
        w = [W(s, r) for r in (1, 2, 3)]
        h = [s.ord(r, "H") for r in (1, 2, 3)]
        ok = (res == ["OK:1", "OK:2", "OK:3"] and y[0].startswith("Недостача: 2 шт") and y[1].startswith("Излишек: 2 шт") and y[2].startswith("Норма")
              and w == ["Частично получено", "Получено", "Получено"] and h == [10.0, 10.0, 10.0] and "больше заказа на 2" in y[1])
        R.add(c, "G > F — недостача, G < F — излишек, G = F — норма; H не меняется; F > H — «Получено» с пометкой «больше заказа»",
              ok, f"{res}; Y {y}; W {w}; H {h}")
        oracle(c, s, "недостача и излишек")
    finally:
        s.close()


# ================================================================ R09–R11 cancellation, repeat, double click

@case
def r09_cancel_order():
    c = "R09"
    s = ready(Session(new_wms("r09")))
    try:
        order(s, 1, h="100")
        nx, nj = s.sysv("NEXT_EI"), njournal(s)
        res = s.click_ord("BtnOrdCancel", 1)
        row = s.ords(1)
        ents = s.journal()[0]
        ok = (res == "OK:1" and row[22] == "Отменено" and row[21] == "" and row[7] == 100.0 and s.ord_locks(1) == POSTED and s.sysv("NEXT_EI") == nx
              and njournal(s) == nj + 1 and ents[-1]["type"] == "ORDER_CANCEL" and s.opos(1)[8] == "ORDER" and s.stock_row(201)[0] == "")
        R.add(c, "отмена неполученного заказа: «Отменено», ЕИ не создан, NEXT_EI не израсходован, история (строка, H) осталась, журнал ORDER_CANCEL",
              ok, f"{res}; {row}; NEXT_EI {s.sysv('NEXT_EI')}")
        again = s.click_ord("BtnOrdCancel", 1)
        p2 = post(s, 1)
        m2 = more(s, 1, "5")
        cl = s.click_ord("BtnRcvClear", 1)
        R.add(c, "повторная отмена — «уже отменена»; «Провести приход», «Ещё поступление», «Очистить» по отменённой позиции — отказ",
              again.startswith("SKIP") and p2.startswith("SKIP") and m2.startswith("ERR") and cl.startswith("ERR") and njournal(s) == nj + 1,
              f"{again}; {p2}; {m2}; {cl}")
        order(s, 2, a="З-2", h="5")
        s.order_input(2, F="5")
        refused = s.Rc("OrderCancelRow", 2)
        R.add(c, "отмена строки с указанным фактом — отказ (пришедший товар проводится приходом)", refused.startswith("ERR") and W(s, 2) != "Отменено", refused)
        oracle(c, s, "отмена заказа")
    finally:
        s.close()


@case
def r10_cancel_rest():
    c = "R10"
    s = ready(Session(new_wms("r10")))
    try:
        order(s, 1, h="100", q=PAST)
        fact(s, 1, "40", g="40")
        post(s, 1)
        res = s.click_ord("BtnOrdCancelRest", 1)
        row = s.ords(1)
        ok = (res == "OK:2" and row[22] == "Частично получено / остаток отменён" and row[7] == 100.0 and row[5] == 40.0
              and "остаток 60 отменён" in row[24] and s.opos(1)[8] == "REST" and s.journal()[0][-1]["type"] == "ORDER_CANCEL_REST")
        R.add(c, "отмена остатка 60 из 100 после прихода 40: «Частично получено / остаток отменён», H = 100 и факт 40 не изменились",
              ok, f"{res}; {row[20:25]}; H {row[7]}")
        s.U("BtnSelfCheck")
        s.OU("BtnOrdRefresh")
        again = s.click_ord("BtnOrdCancelRest", 1)
        R.add(c, "отменённый остаток не считается ожидаемым: «Обновить статусы» при прошедшей Q статус не меняет; повторная отмена — SKIP",
              W(s, 1) == "Частично получено / остаток отменён" and again.startswith("SKIP"), f"{W(s, 1)}; {again}")
        order(s, 3, a="З-3", h="10")
        fact(s, 3, "10")
        post(s, 3)
        full = s.Rc("OrderCancelRestRow", 3)
        order(s, 4, a="З-4", h="10")
        open_ = s.Rc("OrderCancelRestRow", 4)
        R.add(c, "отмена остатка полностью полученной позиции и позиции без прихода — отказ с объяснением",
              full.startswith("ERR") and "полностью" in full and open_.startswith("ERR") and "«Отменить заказ»" in open_, f"{full}; {open_}")
        oracle(c, s, "отмена остатка")
    finally:
        s.close()


@case
def r11_repeat_and_double_click():
    c = "R11"
    s = ready(Session(new_wms("r11")))
    try:
        order(s, 1, h="10")
        fact(s, 1, "10")
        r1 = post(s, 1)
        before = (njournal(s), s.sysv("NEXT_EI"), s.stock(201), s.stock_row(202))
        r2 = post(s, 1)
        r3 = s.Rc("ReceiptPostRow", 1)
        after = (njournal(s), s.sysv("NEXT_EI"), s.stock(201), s.stock_row(202))
        R.add(c, "повторный клик «Провести приход»: «уже проведено», второй ЕИ не создан, остаток не удвоен, NEXT_EI и журнал не изменились",
              r1 == "OK:1" and r2.startswith("SKIP:уже проведено") and r3.startswith("SKIP") and before == after, f"{r1}; {r2}; {before} → {after}")
        order(s, 2, a="З-2", h="7")
        fact(s, 2, "7")
        s.goto("Заказы", "$B$3")
        res = []

        def click():
            s.OU("BtnRcvPost")
            res.append(s.U("TestUiLastMessage"))
        th = [threading.Thread(target=click) for _ in range(2)]
        for t in th:
            t.start()
        for t in th:
            t.join(60)
        rc = [e for e in s.journal()[0] if e["type"] == "RECEIPT"]
        R.add(c, "двойной клик (два нажатия одновременно): один ЕИ, один приход, одна запись журнала",
              len(rc) == 2 and s.sysv("NEXT_EI") == 203 and s.stock(202) == 7.0 and s.stock_row(203)[0] == "", f"ответы {res}; записей RECEIPT {len(rc)}")
        order(s, 3, a="З-3", h="3")
        fact(s, 3, "3")
        ans = []
        th = [threading.Thread(target=lambda: ans.append(s.Rc("ReceiptPostRow", 3))) for _ in range(2)]
        for t in th:
            t.start()
        for t in th:
            t.join(60)
        R.add(c, "два одновременных вызова проведения одной строки: «OK» и «уже проведено», один ЕИ",
              sorted(a[:4] for a in ans) == ["OK:3", "SKIP"] and s.sysv("NEXT_EI") == 204, str(ans))
        s.OU("TestRcvInput", "1", "", "", "", D0, "A-1", "")
        m1 = s.click_ord("BtnRcvMore", 1)
        m2 = s.click_ord("BtnRcvMore", 1)
        R.add(c, "повторный клик «Ещё поступление» без нового ввода в окне: второе поступление не создаётся",
              m1.startswith("OK") and m2.startswith("SKIP") and s.sysv("NEXT_EI") == 205, f"{m1}; {m2}")
        oracle(c, s, "повтор и двойной клик")
    finally:
        s.close()


# ================================================================ R12–R14 copies, inserted rows, filter

@case
def r12_copies():
    c = "R12"
    p = new_wms("r12")
    s = ready(Session(p))
    order(s, 1, h="10")
    fact(s, 1, "10")
    post(s, 1)
    s.goto("Заказы", "$A$2:$AB$2")
    s.ui(".uno:Copy")
    s.goto("Заказы", "$A$8")
    s.ui(".uno:Paste")
    R.add(c, "копировать → вставить проведённую строку на защищённом листе: вставка отклонена", s.ords(7)[21] == "" and s.ords(7)[0] == "",
          str(s.ords(7)[:3]))
    sh = s.doc.Sheets.getByName("Заказы")
    sh.unprotect("wms")
    s.goto("Заказы", "$A$2:$AB$2")
    s.ui(".uno:Copy")
    s.goto("Заказы", "$A$9")
    s.ui(".uno:Paste")
    sh.protect("wms")
    live = s.ord(8, "Y")
    res = post(s, 8)
    fx = s.Rc("ReceiptFixRow", 8, "5", "", "", "", D0, "A-1", "")
    dl = s.Rc("ReceiptDeleteRow", 8)
    R.add(c, "защита снята паролем, проведённая строка вставлена: сразу КОПИЯ, «Провести приход», «Исправить», «Удалить» — отказ",
          live.startswith("КОПИЯ") and "повторяет строку 2" in live and res.startswith("SKIP") and fx.startswith("ERR") and dl.startswith("ERR"),
          f"{live[:90]}; {res}; {fx}; {dl}")
    s.close(save=True)
    s = Session(p, macros=0)
    sh = s.doc.Sheets.getByName("Заказы")
    sh.unprotect("wms")
    row = list(sh.getCellRangeByPosition(0, 1, 27, 1).getDataArray()[0])
    sh.getCellRangeByPosition(0, 4, 27, 4).setDataArray((tuple(row),))
    row[21] = ei(999)
    sh.getCellRangeByPosition(0, 5, 27, 5).setDataArray((tuple(row),))
    row[21] = ei(5)
    sh.getCellRangeByPosition(0, 6, 27, 6).setDataArray((tuple(row),))
    sh.protect("wms")
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        rep = s.report()
        y = [s.ord(r, "Y") for r in (4, 5, 6)]
        pr = [post(s, r) for r in (4, 5, 6)]
        cl = [s.click_ord("BtnRcvClear", r) for r in (4, 5, 6, 8)]
        orig = s.ords(1)
        R.add(c, "сохранено без макросов: копия, ЕИ, который WMS не выдавала, и ЕИ не из прихода помечены КОПИЯ при запуске; не проводятся; «Очистить» убирает; оригинал цел",
              "без WMS" in rep and y[0].startswith("КОПИЯ") and "повторяет" in y[0] and y[1].startswith("КОПИЯ") and "не выдавался" in y[1]
              and y[2].startswith("КОПИЯ") and "не создан приходом" in y[2] and all(x.startswith("SKIP") for x in pr)
              and all(x.startswith("OK") for x in cl) and all(s.ords(r)[21] == "" for r in (4, 5, 6, 8)) and s.ord_locks(4) == OPEN
              and orig[21] == ei(201) and orig[24] == "позиция: получено 10 из 10 шт", f"{y}; {pr}; {cl}; оригинал {orig[20:25]}")
        cl0 = s.click_ord("BtnRcvClear", 1)
        R.add(c, "«Очистить» проведённой строки прихода — отказ", cl0.startswith("ERR") and s.ords(1)[21] == ei(201), cl0)
        oracle(c, s, "копии не считаются приходами")
    finally:
        s.close()


@case
def r13_inserted_rows():
    c = "R13"
    s = ready(Session(new_wms("r13")))
    try:
        order(s, 1, h="100")
        fact(s, 1, "40")
        post(s, 1)
        more(s, 1, "10", c="УПД-2", o=D0)
        order(s, 3, a="З-2", b="Гайка", e="G-1", h="5")
        s.goto("Заказы", "$A$2")
        s.ui(".uno:SelectRow")
        s.ui(".uno:InsertRowsBefore")
        s.ui(".uno:InsertRowsBefore")
        moved = (s.ord(3, "V"), s.ord(4, "V"), s.ord(5, "B"))
        m = more(s, 4, "20", c="УПД-3", o=D0)
        hint_src = s.opos(1)[1]
        iss = issue(s, 1, "202", "3")
        x = s.ord(4, "X")
        hint_add = s.rcv(202)[2]
        # M6 §18: the new delivery row (ЕИ-00000203) is right under the earlier one (row 6 of the sheet), «Гайка» moved down
        dl = s.Rc("ReceiptDeleteRow", 5)
        R.add(c, "вставка строк над приходами: «Ещё поступление» из сдвинутой строки, выдача по ЕИ сдвинутой строки (X обновлён), сторно — "
                 "строки найдены по ключу, подсказки строк исправлены той же операцией",
              moved == (ei(201), ei(202), "Гайка") and m.startswith("OK:3|строка 6|") and hint_src == 3.0 and iss == "OK:4" and x == 7.0 and hint_add == 4.0
              and dl.startswith("OK") and W(s, 3) == "Частично получено" and "получено 50 из 100" in s.ord(3, "Y") and s.ord(6, "B") == "Гайка",
              f"сдвиг {moved}; {m}; подсказка исходной {hint_src}; {iss}; X {x}; подсказка доп. {hint_add}; {dl}; {s.ord(3, 'Y')}")
        locked = s.ord_locks(1)
        s.type_ord("B", 1, "вставлено")
        typed = s.ord(1, "B")
        R.add(c, "вставленная строка между проведёнными закрыта как соседняя; «Очистить» открывает её для ввода заказа",
              s.click_ord("BtnRcvClear", 1).startswith("OK") and s.ord_locks(1) == OPEN, f"защита {locked}; ввод {typed!r}")
        order(s, 1, a="З-9", h="3")
        fact(s, 1, "3")
        R.add(c, "в очищенной вставленной строке проводится новый приход", post(s, 1).startswith("OK") and s.ord(1, "V") == ei(204),
              str(s.ords(1)[20:24]))
        oracle(c, s, "вставка строк")
    finally:
        s.close()


def set_filter(s, col, value):
    dbr = s.doc.DatabaseRanges.getByName("WMS_ORDERS")
    fd = dbr.getFilterDescriptor()
    f = TableFilterField()
    f.Field, f.Operator, f.IsNumeric, f.StringValue = col, EQUAL, False, value
    fd.setFilterFields((f,))
    dbr.refresh()


def visible(s, r):
    return s.doc.Sheets.getByName("Заказы").getRows().getByIndex(r).IsVisible


@case
def r14_filter():
    c = "R14"
    p = new_wms("r14")
    s = ready(Session(p))
    try:
        for r, (a, sup) in enumerate([("З-1", "Озон"), ("З-2", "Петрович"), ("З-3", "Озон"), ("З-4", "Петрович"), ("З-5", "Озон")], start=1):
            order(s, r, a=a, l=sup, h="10")
        fact(s, 1, "4")
        post(s, 1)
        fact(s, 3, "10")
        fact(s, 4, "10")
        fact(s, 5, "10")
        set_filter(s, 11, "Озон")
        vis = [visible(s, r) for r in range(1, 6)]
        m = more(s, 1, "6", c="УПД-2", o=D0)
        # M6 §18: the delivery row of З-1 is inserted right under it (row 3 of the sheet, supplier «Озон» — shown); the rows below move
        f3 = post(s, 4)
        s.goto("Заказы", "$A$2:$AB$7")
        s.OU("BtnRcvPost")
        batch = s.U("TestUiLastMessage")
        vis = [visible(s, r) for r in range(1, 7)]
        ok = (vis == [True, True, False, True, False, True] and m.startswith("OK:2|строка 3|") and f3.startswith("OK") and s.ord(5, "V") == ""
              and s.ord(6, "V") != "" and W(s, 1) == "Получено" and W(s, 3) == "Ожидается")
        R.add(c, "при активном автофильтре (Поставщик = Озон): приход, «Ещё поступление», проведение выделенного блока — только видимые строки, "
                 "скрытая строка с фактом не проведена", ok, f"видимость {vis}; {m}; {f3}; блок {batch}; V5 {s.ord(5, 'V')!r} V6 {s.ord(6, 'V')!r}")
        s.doc.store()
        vis1 = [visible(s, r) for r in range(1, 7)]
        with zipfile.ZipFile(p) as z:
            nfilt = z.read("content.xml").decode("utf-8").count('table:visibility="filter"')
        s.close()
        s = ready(Session(p))
        vis2 = [visible(s, r) for r in range(1, 7)]
        R.add(c, "сохранение при фильтре на «Заказы» (D-041): строки сохранены показанными, фильтр восстановлен после сохранения и при открытии",
              nfilt == 0 and vis1 == vis and vis2 == vis and not s.doc.isModified() and st(s)[0] == "CLEAN", f"{vis1} {vis2}; скрытых в файле {nfilt}")
        oracle(c, s, "фильтр")
    finally:
        s.close()


# ================================================================ R15–R18 correction and storno

@case
def r15_fix_up_down():
    c = "R15"
    s = ready(Session(new_wms("r15")))
    try:
        order(s, 1, h="50")
        fact(s, 1, "40", g="40", c="УПД-1")
        post(s, 1)
        up = fix(s, 1, "50", g="50", c="УПД-1", o="24.09.2026")
        a = (s.ord(1, "F"), s.stock(201), s.ord(1, "X"), W(s, 1), s.opos(1)[5])
        down = fix(s, 1, "30", g="50", c="УПД-1", o="24.09.2026")
        b = (s.ord(1, "F"), s.stock(201), s.ord(1, "X"), W(s, 1), s.opos(1)[5], s.ord(1, "Y"))
        same = fix(s, 1, "30", g="50", c="УПД-1", o="24.09.2026")
        R.add(c, "исправление F вверх (40→50: «Получено») и вниз (50→30: «Частично получено», недостача по документу 20): тот же ЕИ, остаток и X "
                 "пересчитаны; без изменений — «ничего не изменилось»",
              up.startswith("OK") and a == (50.0, 50.0, 50.0, "Получено", 50.0) and down.startswith("OK") and b[:5] == (30.0, 30.0, 30.0, "Частично получено", 30.0)
              and "Недостача: 20" in b[5] and same.startswith("SKIP") and s.sysv("NEXT_EI") == 202, f"{up} {a}; {down} {b}; {same}")
        pl = fix(s, 1, "30", g="50", c="УПД-1", o="24.09.2026", u="C-9")
        R.add(c, "исправление места хранения: у ЕИ одно текущее место — в «Наличие» и в U", pl.startswith("OK") and s.stock_row(201)[5] == "C-9"
              and s.ord(1, "U") == "C-9", f"{pl}; {s.stock_row(201)}")
        oracle(c, s, "исправления")
    finally:
        s.close()


@case
def r16_fix_below_issued():
    c = "R16"
    s = ready(Session(new_wms("r16")))
    try:
        order(s, 1, h="10")
        fact(s, 1, "10")
        post(s, 1)
        iss = issue(s, 1, "ЕИ-201", "7")
        x = s.ord(1, "X")
        snap0 = snap(s)
        low = fix(s, 1, "5")
        R.add(c, "принято 10, выдано 7: исправить приход на 5 — отказ «уже выдано 7», ничего не изменилось; X показывает остаток 3",
              iss == "OK:2" and x == 3.0 and low.startswith("ERR") and "уже выдано 7" in low and snap(s) == snap0, f"{iss}; X {x}; {low}")
        ok7 = fix(s, 1, "7")
        z = (s.stock(201), s.ord(1, "X"))
        ok12 = fix(s, 1, "12")
        R.add(c, "исправить на 7 (= выдано) — остаток 0; на 12 — остаток 5", ok7.startswith("OK") and z == (0.0, 0.0) and ok12.startswith("OK")
              and s.stock(201) == 5.0 and s.ord(1, "X") == 5.0, f"{ok7} {z}; {ok12} {s.stock(201)}")
        oracle(c, s, "исправление ниже выданного")
    finally:
        s.close()


@case
def r17_storno_without_issues():
    c = "R17"
    s = ready(Session(new_wms("r17")))
    try:
        order(s, 1, h="10", q=FUTURE)
        fact(s, 1, "10")
        post(s, 1)
        k0 = s.U("TestUiConfirmCount")
        dl = s.click_ord("BtnRcvDelete", 1)
        k1 = s.U("TestUiConfirmCount")
        row = s.ords(1)
        reg = s.stock_row(201)
        ok = (dl.startswith("OK") and k1 == k0 + 1 and row[21] == ei(201) and row[22] == "Ожидается" and row[23] == 0.0 and row[24].startswith("Приход удалён (сторно)")
              and reg[4] == 0.0 and reg[7] == "Приход удалён (сторно)" and s.rcv(201)[4] == "STORNO" and s.journal()[0][-1]["type"] == "RECEIPT_DEL"
              and s.ord_locks(1) == POSTED)
        R.add(c, "«Удалить» = сторно прихода без выдач (после подтверждения): остаток 0, запись ЕИ осталась в реестре со статусом «Приход удалён "
                 "(сторно)», строка — история, позиция снова ожидается", ok, f"{dl}; подтверждений {k1 - k0}; {row[20:25]}; {reg}")
        again = s.click_ord("BtnRcvDelete", 1)
        rp = post(s, 1)
        m = more(s, 1, "10", c="УПД-2", o=D0)
        R.add(c, "повторное «Удалить» — SKIP; «Провести приход» в строке сторно — отказ; новая поставка — «Ещё поступление» с новым ЕИ (ЕИ не переиспользуется)",
              again.startswith("SKIP") and rp.startswith("SKIP") and m.startswith("OK") and s.ord(2, "V") == ei(202) and W(s, 1) == "Получено",
              f"{again}; {rp}; {m}")
        s.issue_input(1, ei="201", qty="1", date=D0, who="егр")
        iz = s.post(1)
        R.add(c, "выдача по сторнированному ЕИ невозможна (остаток 0)", iz.startswith("ERR") and "равен 0" in iz, iz)
        oracle(c, s, "сторно без выдач")
    finally:
        s.close()


@case
def r18_storno_after_issue():
    c = "R18"
    s = ready(Session(new_wms("r18")))
    try:
        order(s, 1, h="10")
        fact(s, 1, "10")
        post(s, 1)
        issue(s, 1, "201", "3")
        snap0 = snap(s)
        k0 = s.U("TestUiConfirmCount")
        dl = s.click_ord("BtnRcvDelete", 1)
        k1 = s.U("TestUiConfirmCount")
        R.add(c, "сторно после выдачи 3 из 10 — отказ сразу, без окна подтверждения (действующая выдача, D-069; уже выдано 3), ничего не изменилось",
              dl.startswith("ERR") and "уже выдано 3" in s.U("TestUiLastMessage") and k1 == k0 and snap(s) == snap0, f"{dl}; подтверждений {k1 - k0}")
        di = s.I("IssueDeleteRow", 1)
        x = s.ord(1, "X")
        dl2 = s.Rc("ReceiptDeleteRow", 1)
        R.add(c, "после сторно выдачи (остаток снова 10, X = 10) сторно прихода проходит, остаток 0",
              di.startswith("OK") and x == 10.0 and dl2.startswith("OK") and s.stock(201) == 0.0 and s.ord(1, "X") == 0.0, f"{di}; X {x}; {dl2}")
        oracle(c, s, "сторно после выдачи")
    finally:
        s.close()


# ================================================================ R19–R25 crashes, recovery, Ctrl+Z

def crash_receipt(c, name, sheet=None, point=None, expect_rollback=True):
    """a receipt interrupted by a simulated crash (the book stored mid-operation, the process killed), then restarted"""
    p = new_wms(name)
    s = ready(Session(p))
    order(s, 1, h="100", q=FUTURE)
    fact(s, 1, "40", g="40")
    s.close(save=True)
    s = ready(Session(p))
    if sheet:
        s.B("TestSetFaultSheet", sheet, 2)
    else:
        s.B("TestSetFault", point, 2)
    crash = s.Rc("ReceiptPostRow", 1)
    s.kill()
    s = ready(Session(p))
    unl = s.B("ActionForceUnlock")
    return p, s, crash, unl


@case
def r19_crash_after_ei_reserved():
    c = "R19"
    p, s, crash, unl = crash_receipt(c, "r19", sheet="_SYS")
    try:
        row = s.ords(1)
        ok = (crash == "CRASH-SIM" and "отменена по снимку" in unl and st(s)[0] == "CLEAN" and s.sysv("NEXT_EI") == 201 and row[21] == ""
              and row[5] == "40" and s.stock_row(201)[0] == "" and s.rcv(201)[0] == "" and s.next_ol() == 1.0 and njournal(s) == 0)
        R.add(c, "сбой сразу после резервирования номера ЕИ (NEXT_EI записан, книга сохранена, kill): при запуске откат — NEXT_EI 201, ЕИ нигде нет, ввод на месте",
              ok, f"{crash}; {unl[:160]}; строка {row[:8]}; NEXT_EI {s.sysv('NEXT_EI')}")
        again = s.Rc("ReceiptPostRow", 1)
        R.add(c, "повторное проведение создаёт ЕИ-00000201 ровно один раз", again == "OK:1" and s.ord(1, "V") == ei(201) and s.sysv("NEXT_EI") == 202, again)
        oracle(c, s, "сбой после резервирования ЕИ")
    finally:
        s.close()


@case
def r20_crash_after_stock_row():
    c = "R20"
    p, s, crash, unl = crash_receipt(c, "r20", sheet="Наличие")
    try:
        ok = (crash == "CRASH-SIM" and "отменена по снимку" in unl and s.stock_row(201)[0] == "" and s.sysv("NEXT_EI") == 201 and s.ord(1, "V") == "")
        R.add(c, "сбой после записи строки «Наличие»: откат — строки реестра нет, NEXT_EI не израсходован", ok, f"{crash}; {unl[:120]}; {s.stock_row(201)}")
        again = s.Rc("ReceiptPostRow", 1)
        R.add(c, "повторное проведение — один ЕИ", again == "OK:1" and s.stock_row(201)[4] == 40.0, again)
        oracle(c, s, "сбой после «Наличие»")
    finally:
        s.close()


@case
def r21_crash_after_order_row():
    c = "R21"
    p, s, crash, unl = crash_receipt(c, "r21", sheet="Заказы")
    try:
        ok = (crash == "CRASH-SIM" and "отменена по снимку" in unl and s.ord(1, "V") == "" and s.ord_locks(1) == OPEN and s.stock_row(201)[0] == ""
              and s.opos(1)[0] == "" and s.sysv("NEXT_EI") == 201 and W(s, 1) == "Ожидается")
        R.add(c, "сбой после записи строки заказа (все записи сделаны, журнала нет): откат — строка снова открыта, ЕИ и позиции нет",
              ok, f"{crash}; {unl[:120]}; {s.ords(1)[20:25]}; защита {s.ord_locks(1)}")
        again = s.Rc("ReceiptPostRow", 1)
        R.add(c, "повторное проведение — один ЕИ", again == "OK:1" and s.sysv("NEXT_EI") == 202, again)
        oracle(c, s, "сбой после строки заказа")
    finally:
        s.close()


@case
def r22_crash_after_journal():
    c = "R22"
    p, s, crash, unl = crash_receipt(c, "r22", point=3)
    try:
        _, block = st(s)
        rec = s.B("ActionRecover")
        row = s.ords(1)
        ok = (crash == "CRASH-SIM" and block == "TAIL" and "восстановлено операций: 1" in rec and st(s)[0] == "CLEAN" and row[21] == ei(201)
              and row[22] == "Частично получено" and row[23] == 40.0 and s.stock(201) == 40.0 and s.sysv("NEXT_EI") == 202 and s.ord_locks(1) == POSTED)
        R.add(c, "сбой после записи в журнал: при запуске «хвост», «Восстановить» доводит приход ровно один раз (ЕИ, «Наличие», позиция, строка, статус)",
              ok, f"{crash}; {block}; {rec[:120]}; {row[20:25]}")
        again = s.Rc("ReceiptPostRow", 1)
        R.add(c, "после восстановления повторное проведение — «уже проведено»", again.startswith("SKIP"), again)
        oracle(c, s, "сбой после журнала")
    finally:
        s.close()


@case
def r23_crash_during_add():
    c = "R23"
    p = new_wms("r23")
    s = ready(Session(p))
    order(s, 1, h="100")
    fact(s, 1, "40")
    post(s, 1)
    s.close(save=True)
    s = ready(Session(p))
    s.B("TestSetFault", 2, 2)
    s.OU("TestRcvInput", "30", "", "УПД-2", D0, D0, "A-2", "")
    crash = s.click_ord("BtnRcvMore", 1)
    s.kill()
    s = ready(Session(p))
    s.B("ActionForceUnlock")
    ok1 = (crash == "CRASH-SIM" and s.ords(2)[0] == "" and s.sysv("NEXT_EI") == 202 and s.opos(1)[5] == 40.0 and s.stock_row(202)[0] == "")
    R.add(c, "сбой посреди «Ещё поступление» (до журнала): откат — новой строки, ЕИ и изменения позиции нет", ok1, f"{crash}; {s.ords(2)[:6]}; {s.opos(1)}")
    s.B("TestSetFault", 3, 2)
    crash2 = s.Rc("ReceiptAddRow", 1, "30", "", "УПД-2", D0, D0, "A-2", "")
    s.kill()
    s = ready(Session(p))
    try:
        s.B("ActionForceUnlock")
        rec = s.B("ActionRecover")
        ok2 = (crash2 == "CRASH-SIM" and "восстановлено операций: 1" in rec and s.ord(2, "V") == ei(202) and s.opos(1)[5] == 70.0
               and "получено 70 из 100" in s.ord(1, "Y") and s.sysv("NEXT_EI") == 203)
        R.add(c, "сбой после журнала при «Ещё поступление»: «Восстановить» создаёт дополнительную строку ровно один раз", ok2, f"{crash2}; {rec[:100]}")
        oracle(c, s, "сбой при доп. поступлении")
    finally:
        s.close()


@case
def r24_crash_during_fix():
    c = "R24"
    p = new_wms("r24")
    s = ready(Session(p))
    order(s, 1, h="100")
    fact(s, 1, "40")
    post(s, 1)
    s.close(save=True)
    s = ready(Session(p))
    s.B("TestSetFault", 2, 2)
    crash = s.Rc("ReceiptFixRow", 1, "60", "", "УПД-1", "24.09.2026", D0, "A-1", "")
    s.kill()
    s = ready(Session(p))
    s.B("ActionForceUnlock")
    ok1 = crash == "CRASH-SIM" and s.ord(1, "F") == 40.0 and s.stock(201) == 40.0 and s.opos(1)[5] == 40.0
    R.add(c, "сбой посреди исправления (до журнала): старый приход цел, промежуточного состояния нет", ok1, f"{crash}; F {s.ord(1, 'F')}; {s.stock(201)}")
    s.B("TestSetFault", 3, 2)
    crash2 = s.Rc("ReceiptFixRow", 1, "60", "", "УПД-1", "24.09.2026", D0, "A-1", "")
    s.kill()
    s = ready(Session(p))
    try:
        s.B("ActionForceUnlock")
        rec = s.B("ActionRecover")
        ok2 = (crash2 == "CRASH-SIM" and "восстановлено операций: 1" in rec and s.ord(1, "F") == 60.0 and s.stock(201) == 60.0 and s.opos(1)[5] == 60.0
               and s.ord(1, "V") == ei(201))
        R.add(c, "сбой после журнала при исправлении: «Восстановить» доводит RECEIPT_FIX целиком, ЕИ тот же", ok2, f"{crash2}; {rec[:100]}")
        oracle(c, s, "сбой при исправлении")
    finally:
        s.close()


@case
def r25_crash_during_storno():
    c = "R25"
    p = new_wms("r25")
    s = ready(Session(p))
    order(s, 1, h="100")
    fact(s, 1, "40")
    post(s, 1)
    s.close(save=True)
    s = ready(Session(p))
    s.B("TestSetFault", 2, 2)
    crash = s.Rc("ReceiptDeleteRow", 1)
    s.kill()
    s = ready(Session(p))
    s.B("ActionForceUnlock")
    ok1 = crash == "CRASH-SIM" and s.stock(201) == 40.0 and s.rcv(201)[4] == "LIVE" and W(s, 1) == "Частично получено"
    R.add(c, "сбой посреди сторно (до журнала): приход действует, остаток 40", ok1, f"{crash}; {s.rcv(201)}")
    s.B("TestSetFault", 3, 2)
    crash2 = s.Rc("ReceiptDeleteRow", 1)
    s.kill()
    s = ready(Session(p))
    try:
        s.B("ActionForceUnlock")
        rec = s.B("ActionRecover")
        ok2 = (crash2 == "CRASH-SIM" and "восстановлено операций: 1" in rec and s.stock(201) == 0.0 and s.rcv(201)[4] == "STORNO"
               and s.stock_row(201)[7] == "Приход удалён (сторно)" and s.opos(1)[5] == 0.0)
        R.add(c, "сбой после журнала при сторно: «Восстановить» доводит сторно; ЕИ остаётся в реестре со статусом сторно", ok2, f"{crash2}; {rec[:100]}")
        oracle(c, s, "сбой при сторно")
    finally:
        s.close()


@case
def r26_recover_unsaved_mix():
    c = "R26"
    p = new_wms("r26")
    s = ready(Session(p))
    order(s, 1, a="З-1", h="100")
    fact(s, 1, "10")
    post(s, 1)
    s.close(save=True)
    s = ready(Session(p))
    order(s, 2, a="З-2", b="Гайка", e="G-7", h="50")
    fact(s, 2, "20", c="УПД-5")
    post(s, 2)
    more(s, 2, "15", c="УПД-6", o=D0)
    fix(s, 1, "12", c="УПД-1", o="24.09.2026")
    order(s, 5, a="З-3", h="9")
    s.Rc("OrderCancelRow", 5)
    s.Rc("OrderCancelRestRow", 2)
    issue(s, 1, "202", "4")
    s.close(save=False)
    s = ready(Session(p))
    try:
        _, block = st(s)
        rec = s.B("ActionRecover")
        ok = (block == "TAIL" and "восстановлено операций: 6" in rec and st(s)[0] == "CLEAN" and s.ord(2, "V") == ei(202) and s.ord(3, "V") == ei(203)
              and s.ord(1, "F") == 12.0 and W(s, 5) == "Отменено" and W(s, 2) == "Частично получено / остаток отменён" and s.ord(2, "X") == 16.0
              and s.ord(2, "B") == "Гайка" and s.ord(5, "A") == "З-3")
        R.add(c, "«Не сохранять» после прихода, доп. поступления, исправления, отмены заказа, отмены остатка и выдачи → «Восстановить» "
                 "возвращает всё, включая введённые строки заказов", ok, f"{block}; {rec[:120]}; W {[W(s, r) for r in (1, 2, 3, 5)]}")
        oracle(c, s, "восстановление несохранённого")
    finally:
        s.close()


@case
def r27_ctrl_z():
    c = "R27"
    s = ready(Session(new_wms("r27")))
    try:
        order(s, 1, h="100")
        fact(s, 1, "40")
        post(s, 1)
        z1 = ctrl_z_harmless(s)
        more(s, 1, "5", c="УПД-2", o=D0)
        z2 = ctrl_z_harmless(s)
        fix(s, 1, "41")
        z3 = ctrl_z_harmless(s)
        s.Rc("ReceiptDeleteRow", 2)
        z4 = ctrl_z_harmless(s)
        s.Rc("OrderCancelRestRow", 1)
        z5 = ctrl_z_harmless(s)
        s.type_ord("Z", 1, "комментарий")
        typed = s.ord(1, "Z")
        s.ui(".uno:Undo")
        R.add(c, "Ctrl+Z ×5 после прихода, доп. поступления, исправления, сторно и отмены остатка ничего не откатывает; обычный ввод (комментарий Z) отменяется",
              z1 and z2 and z3 and z4 and z5 and typed == "комментарий" and s.ord(1, "Z") == "" and s.state()["UNDO_LOCKED"] == "False",
              f"{z1} {z2} {z3} {z4} {z5}; {typed!r} → {s.ord(1, 'Z')!r}")
        oracle(c, s, "Ctrl+Z")
    finally:
        s.close()


@case
def r28_save_reopen():
    c = "R28"
    p = new_wms("r28")
    s = ready(Session(p))
    order(s, 1, h="100", q=FUTURE)
    fact(s, 1, "40")
    post(s, 1)
    more(s, 1, "10", c="УПД-2", o=D0)
    order(s, 3, a="З-2", h="5", q=FUTURE)
    s.Rc("OrderCancelRow", 3)
    order(s, 4, a="З-3", b="Шуруп", h="7", q=FUTURE)
    s.close(save=True)
    with zipfile.ZipFile(p) as z:
        xml = z.read("content.xml").decode("utf-8")
    s = ready(Session(p))
    try:
        checks = {"CLEAN": st(s)[0] == "CLEAN", "не «изменена» после открытия": not s.doc.isModified(),
                  "W": [W(s, r) for r in (1, 2, 3, 4)] == ["Частично получено", "Дополнительное поступление", "Отменено", "Ожидается"],
                  "защита": [s.ord_locks(r) for r in (1, 2, 3, 4)] == [POSTED, POSTED, POSTED, OPEN],
                  # «Выдачи», «Заказы», since Core Phase 4 «Возврат» and since Core Phase 5 «Иной приход» allow inserting rows
                  "вставка строк": xml.count('loext:insert-rows="true"') == 5, "автофильтр": s.doc.DatabaseRanges.hasByName("WMS_ORDERS"),
                  "NEXT_EI": s.sysv("NEXT_EI") == 203, "NEXT_OL": s.next_ol() == 3.0}
        fact(s, 4, "7")
        nxt = post(s, 4)
        checks["следующий приход"] = nxt == "OK:4" and s.ord(4, "V") == ei(203)
        bad = [k for k, v in checks.items() if not v]
        R.add(c, "сохранить → закрыть → открыть: статусы, защита, опция «вставка строк», автофильтр, счётчики целы; книга не «изменена»; нумерация продолжается",
              not bad, f"не выполнено: {bad}")
        oracle(c, s, "после повторного открытия")
    finally:
        s.close()


# ================================================================ R29–R33 unknown EI, locale, date, antidubl, several orders

@case
def r29_unknown_forged_ei():
    c = "R29"
    s = ready(Session(new_wms("r29")))
    try:
        order(s, 1, h="10")
        s.type_ord("V", 1, "ЕИ-00000201")
        typed = s.ord(1, "V")
        fact(s, 1, "10")
        post(s, 1)
        bad = issue(s, 2, "ЕИ-00000250", "1")
        good = issue(s, 3, "ЕИ-201", "1")
        sh = s.doc.Sheets.getByName("Заказы")
        sh.unprotect("wms")
        sh.getCellByPosition(0, 6).setString("З-9")
        sh.getCellByPosition(1, 6).setString("Подделка")
        sh.getCellByPosition(21, 6).setString("ЕИ-00000150")
        sh.protect("wms")
        s.goto("Заказы", "$V$7")
        s.ui(".uno:Copy")
        s.ui(".uno:Paste")
        k = s.Od("OrderRowKind", 6)
        s.Od("CopyCheckOrder", 6)
        y = s.ord(6, "Y")
        fx = s.Rc("ReceiptFixRow", 6, "1", "", "", "", D0, "A-1", "")
        R.add(c, "V нельзя ввести вручную; неизвестный ЕИ в выдаче — отказ; ЕИ прихода выдаётся; поддельный V (ЕИ не из прихода) — КОПИЯ, операции отказывают",
              typed == "" and bad.startswith("ERR") and good == "OK:2" and k == "FOREIGN" and y.startswith("КОПИЯ") and "не создан приходом" in y
              and fx.startswith("ERR"), f"ввод V {typed!r}; {bad}; {good}; {k}; {y[:80]}; {fx}")
        oracle(c, s, "поддельный ЕИ")
    finally:
        s.close()


def locale_case(c, name, locale):
    s = ready(Session(new_wms(name), locale=locale))
    out = []
    try:
        vals = ["1,5", "1.5", "0,125", "2,1250", "1500", "1,500", "-1", "1e3", "1 000", "01.02.2026", "0", "abc"]
        for i, v in enumerate(vals, start=1):
            order(s, i, a=f"З-{i}", h="5000")
            s.order_input(i, F=v, N=D0, U="A-1")
            res = s.Rc("ReceiptPostRow", i)
            out.append((v, res[:2], s.ord(i, "F")))
        want_ok = {"1,5": 1.5, "1.5": 1.5, "0,125": 0.125, "2,1250": 2.125, "1500": 1500.0}
        ok = all((x[1] == "OK") == (x[0] in want_ok) and (x[0] not in want_ok or x[2] == want_ok[x[0]]) for x in out)
        return s, ok, out
    except Exception:
        s.close()
        raise


@case
def r30_quantity_locale():
    c = "R30"
    for loc in ("ru-RU", "en-US"):
        s, ok, out = locale_case(c, f"r30_{loc}", loc)
        try:
            R.add(c, f"локаль {loc}: F «1,5», «1.5», «0,125», «2,1250», «1500» принимаются как числа; «1,500», «-1», «1e3», «1 000», дата, 0, текст — отказ",
                  ok, str(out))
            s.order_input(20, A="З-20", B="Тест", H="1,500", I="шт", F="1", N=D0, U="A")
            h = s.Rc("ReceiptPostRow", 20)
            s.order_input(21, A="З-21", B="Тест", H="10", I="шт", F="1", G="1.500", N=D0, U="A")
            g = s.Rc("ReceiptPostRow", 21)
            s.order_input(22, A="З-22", B="Тест", H="10", I="шт", F="1", N=D0, U="A", J="12,50")
            j = s.Rc("ReceiptPostRow", 22)
            s.order_input(23, A="З-23", B="Тест", H="10", I="шт", F="1", N=D0, U="A", J="-5")
            jn = s.Rc("ReceiptPostRow", 23)
            R.add(c, f"локаль {loc}: неоднозначные H «1,500» и G «1.500» — отказ; цена J «12,50» — число 12,5; отрицательная цена — отказ",
                  h.startswith("ERR") and g.startswith("ERR") and j.startswith("OK") and s.ord(22, "J") == 12.5 and jn.startswith("ERR"), f"{h}; {g}; {j}; {jn}")
            oracle(c, s, f"локаль {loc}")
        finally:
            s.close()


@case
def r31_dates():
    c = "R31"
    s = ready(Session(new_wms("r31"), locale="en-US"))
    try:
        order(s, 1, h="5")
        s.order_input(1, F="5", U="A-1")
        no_n = s.Rc("ReceiptPostRow", 1)
        s.order_input(1, N="31.02.2026")
        bad_n = s.Rc("ReceiptPostRow", 1)
        s.order_input(1, N="25.09.1999")
        old_n = s.Rc("ReceiptPostRow", 1)
        s.order_input(1, N="25.09.2026", O="99.99.2026")
        bad_o = s.Rc("ReceiptPostRow", 1)
        s.order_input(1, O="24.09.2026", Q="30.02.2026")
        bad_q = s.Rc("ReceiptPostRow", 1)
        s.order_input(1, Q="30.09.2026")
        ok = s.Rc("ReceiptPostRow", 1)
        R.add(c, "дата: без N — отказ; 31.02 — отказ; вне 2000–2099 — отказ; неверные O и Q — отказ; текст дд.мм.гггг в локали en-US принимается и становится датой",
              no_n.startswith("ERR") and "дата поступления" in no_n and bad_n.startswith("ERR") and old_n.startswith("ERR") and bad_o.startswith("ERR")
              and bad_q.startswith("ERR") and ok == "OK:1" and s.ord(1, "N") == 46290.0 and s.ord(1, "O") == 46289.0,
              f"{no_n}; {bad_n}; {old_n}; {bad_o}; {bad_q}; {ok}; N {s.ord(1, 'N')!r}")
        oracle(c, s, "даты")
    finally:
        s.close()


@case
def r32_antidubl():
    c = "R32"
    s = ready(Session(new_wms("r32")))
    try:
        order(s, 1, a="З-1", h="10", l="Петрович", e="ART-1")
        fact(s, 1, "10", c="УПД-100", o="24.09.2026")
        post(s, 1)
        order(s, 2, a="З-2", h="10", l="Петрович", e="ART-1")
        fact(s, 2, "10", c="упд 100", o="24.09.2026")
        chk = s.click_ord("BtnRcvCheck", 2)
        prev = s.ord(2, "AB")
        res = post(s, 2)
        ab = s.ord(2, "AB")
        order(s, 3, a="З-3", h="10", l="Петрович", e="ART-1")
        fact(s, 3, "9", c="УПД-100", o="24.09.2026")
        res3 = post(s, 3)
        ok = (chk.startswith("OK") and "ЕИ-00000201" in prev and res == "OK:2" and ab.startswith("Возможный дубль: ЕИ-00000201 (строка 2)")
              and res3 == "OK:3" and s.ord(3, "AB") == "")
        R.add(c, "AB: тот же поставщик + документ (с другим написанием) + артикул + количество + дата — предупреждение «Возможный дубль», приход всё равно проводится; "
                 "другое количество — без предупреждения", ok, f"{chk}; предпросмотр {prev!r}; {res}; {ab!r}; {res3}")
        s.Rc("ReceiptDeleteRow", 1)
        s.Rc("ReceiptDeleteRow", 2)
        order(s, 4, a="З-4", h="10", l="Петрович", e="ART-1")
        fact(s, 4, "10", c="УПД-100", o="24.09.2026")
        res4 = post(s, 4)
        R.add(c, "после сторно прежних приходов их ключи не участвуют — нового предупреждения нет", res4 == "OK:6" and s.ord(4, "AB") == "", f"{res4}; {s.ord(4, 'AB')!r}")
        order(s, 5, a="З-5", b="Лента", e="", h="3", l="Озон")
        fact(s, 5, "3", c="", o=None)
        order(s, 6, a="З-6", b="лента", e="", h="3", l="озон")
        fact(s, 6, "3", c="", o=None)
        post(s, 5)
        post(s, 6)
        R.add(c, "без артикула ключ строится по нормализованному названию", s.ord(6, "AB").startswith("Возможный дубль: ЕИ-00000205"), s.ord(6, "AB"))
        oracle(c, s, "антидубль")
    finally:
        s.close()


@case
def r33_several_orders():
    c = "R33"
    s = ready(Session(new_wms("r33")))
    try:
        rows = [("З-10", "Кабель", "50"), ("З-10", "Клемма", "30"), ("З-11", "Кабель", "20"), ("З-11", "Лампа", "10"), ("З-12", "Кабель", "5"),
                ("З-12", "Реле", "8")]
        for r, (a, b, h) in enumerate(rows, start=1):
            order(s, r, a=a, b=b, e="", h=h, q=FUTURE)
        for r, f in ((1, "20"), (2, "30"), (3, "20"), (4, ""), (5, "abc"), (6, "8")):
            if f:
                fact(s, r, f, c=f"УПД-{r}")
        s.goto("Заказы", "$A$2:$AB$7")
        s.OU("BtnRcvPost")
        batch = s.U("TestUiLastMessage")
        m1 = more(s, 1, "30", c="УПД-20", o=D0)
        # M6 §18: each delivery row goes right under its position — З-10 Кабель's under row 2, then З-11 Кабель (now row 5) gets its own
        m5 = s.Rc("ReceiptAddRow", 4, "1", "", "", "", D0, "A-1", "")
        ws = [W(s, r) for r in (1, 3, 4, 6, 7, 8)]
        ok = (batch.startswith("OK=4;ERR=1;SKIP=1") and ws == ["Получено", "Получено", "Получено без документов", "Ожидается", "Ожидается", "Получено"]
              and m1.startswith("OK:5|строка 3|") and m5.startswith("OK:6|строка 6|") and W(s, 2) == W(s, 5) == "Дополнительное поступление"
              and s.opos(1)[5] == 50.0 and s.opos(3)[5] == 21.0 and s.ord(7, "Y").startswith("Не проведено"))
        R.add(c, "несколько заказов одновременно: блок из 6 строк 3 заказов — 4 прихода проведены, ошибка строки 6 в «Контроль», строка без F пропущена; "
                 "дальнейшие поступления — каждое к своей позиции (поступление без документов меняет статус своей позиции)", ok, f"{batch}; {m1}; {m5}; W {ws}")
        oracle(c, s, "несколько заказов")
    finally:
        s.close()


# ================================================================ R34–R39 statuses, special receipts, buttons, EI safety, abandon

@case
def r34_status_not_by_date():
    c = "R34"
    p = new_wms("r34")
    s = ready(Session(p))
    order(s, 1, a="З-1", h="10", q=PAST)
    order(s, 2, a="З-2", h="10", q=FUTURE)
    order(s, 3, a="З-3", h="10")
    order(s, 4, a="З-4", h="10", q=PAST)
    fact(s, 4, "5")
    post(s, 4)
    w0 = [W(s, r) for r in (1, 2, 3, 4)]
    s.close(save=True)
    # the same book as WMS before 0.7.2 left it (date statuses in W) and a row entered with macros off (no status yet)
    s = Session(p, macros=0)
    sh = s.doc.Sheets.getByName("Заказы")
    sh.unprotect("wms")
    sh.getCellByPosition(22, 1).setString("Просрочено")
    sh.getCellByPosition(22, 4).setString("Частично получено / просрочено")
    sh.protect("wms")
    sh.getCellByPosition(0, 5).setString("З-6")
    sh.getCellByPosition(1, 5).setString("Без статуса")
    sh.getCellByPosition(7, 5).setString("3")
    sh.getCellByPosition(8, 5).setString("шт")
    sh.getCellByPosition(16, 5).setValue(46000)
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        rep = s.report()
        w1 = [W(s, r) for r in (1, 2, 3, 4, 5)]
        R.add(c, "статус не зависит от даты (D-089): при вводе прошедшая ожидаемая дата Q не делает позицию «Просрочено», частично полученную — "
                 "«Частично получено / просрочено»; при открытии строка без статуса (ввод без макросов) получает «Ожидается», статусы по дате "
                 "прежней версии («Просрочено», «Частично получено / просрочено») становятся статусами по данным; в отчёте запуска — позиции с "
                 "прошедшей ожидаемой датой",
              w0 == ["Ожидается", "Ожидается", "Ожидается", "Частично получено"]
              and w1 == ["Ожидается", "Ожидается", "Ожидается", "Частично получено", "Ожидается"]
              and "статусов по дате прежней версии пересчитано 2" in rep and "обновлено статусов 3" in rep
              and "ожидаемых позиций 5, из них с прошедшей ожидаемой датой 3" in rep, f"{w0}; {w1}; {rep[-320:]}")
        q = s.doc.Sheets.getByName("Заказы").getCellByPosition(16, 1).ConditionalFormat
        cf = {q.getByIndex(i).getStyleName(): q.getByIndex(i).getFormula1() for i in range(q.getCount())}
        later = (TODAY + datetime.timedelta(days=40)).strftime("%d.%m.%Y")
        s.Od("TestSetToday", later)
        s.OU("BtnOrdRefresh")
        msg = s.U("TestUiLastMessage")
        w2 = [W(s, r) for r in (1, 2, 3, 4, 5)]
        R.add(c, "«Обновить статусы» с датой через 40 дней: ни один статус не меняется — заказ ждём до поставки или отмены; в сообщении — "
                 "позиций с прошедшей ожидаемой датой 4 (у З-2 срок тоже прошёл); на Q — условное оформление «ожидаемая дата прошла» "
                 "(подсветка по TODAY(), статус и проведение не затрагивает)",
              w2 == w1 and "из них с прошедшей ожидаемой датой 4" in msg and "обновлено статусов 0" in msg and "WMS_LateDate" in cf
              and "TODAY()" in cf.get("WMS_LateDate", "") and "$W2" in cf.get("WMS_LateDate", ""), f"{w2}; {msg}; {cf}")
        oracle(c, s, "статусы без даты", today=later)
        s.Od("TestSetToday", "")
    finally:
        s.close()


@case
def r35_special_receipts_refused():
    c = "R35"
    s = ready(Session(new_wms("r35")))
    try:
        out = []
        for r, sup in enumerate(["Офис", "Производство", "детали", "Старый склад"], start=1):
            order(s, r, a=f"З-{r}", l=sup, h="1")
            fact(s, r, "1")
            out.append((sup, post(s, r)[:4], s.ord(r, "Y")[:60]))
        R.add(c, "специальные приходы (Офис, Производство, Детали, Старый склад) обычным приходом не проводятся — понятный отказ, ЕИ не создан",
              all(x[1] == "ERR:" and "специальный приход" in x[2] for x in out) and s.sysv("NEXT_EI") == 201 and njournal(s) == 0, str(out))
    finally:
        s.close()


@case
def r36_buttons_bound():
    c = "R36"
    s = Session(new_wms("r36"))
    try:
        form = s.doc.Sheets.getByName("Заказы").getDrawPage().getForms().getByIndex(0)
        names = {}
        for i in range(form.getCount()):
            ev = form.getScriptEvents(i)
            code = ev[0].ScriptCode if ev else ""
            names[form.getByIndex(i).Label] = code.split("Standard.", 1)[1].split("?")[0] if "Standard." in code else code
        want = {"Проверить": "WmsOrdersUi.BtnRcvCheck", "Провести приход": "WmsOrdersUi.BtnRcvPost", "Ещё поступление": "WmsOrdersUi.BtnRcvMore",
                "Исправить": "WmsOrdersUi.BtnRcvFix", "Удалить": "WmsOrdersUi.BtnRcvDelete", "Отменить заказ": "WmsOrdersUi.BtnOrdCancel",
                "Отменить остаток": "WmsOrdersUi.BtnOrdCancelRest", "Обновить статусы": "WmsOrdersUi.BtnOrdRefresh", "Очистить": "WmsOrdersUi.BtnRcvClear"}
        R.add(c, "9 кнопок листа «Заказы» привязаны к макросам (запуск через меню не нужен)", all(names.get(k) == v for k, v in want.items()), str(names))
    finally:
        s.close()


@case
def r37_new_ei_safety():
    c = "R37"
    p = new_wms("r37")
    s = Session(p, macros=0)
    st_sh = s.doc.Sheets.getByName("Наличие")
    st_sh.unprotect("wms")
    st_sh.getCellRangeByPosition(0, 201, 8, 201).setDataArray((("ЕИ-00000201", "Старый товар", "", "шт", 3.0, "X", "", "", "миграция"),))
    st_sh.protect("wms")
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        order(s, 1, h="5")
        fact(s, 1, "5")
        res = post(s, 1)
        R.add(c, "в «Наличие» уже есть ЕИ с номером NEXT_EI (мигрированный код): новый ЕИ не выдаётся, приход отказан с понятной причиной, ничего не записано",
              res.startswith("ERR") and "отстаёт" in s.ord(1, "Y") and s.ord(1, "V") == "" and njournal(s) == 0 and s.sysv("NEXT_EI") == 201,
              f"{res}; {s.ord(1, 'Y')[:120]}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка сообщает, что NEXT_EI не выше максимального ЕИ книги", "FAIL приходы и заказы" in sc and "NEXT_EI" in sc,
              [x for x in sc.splitlines() if "приходы" in x][0][:200])
    finally:
        s.close()


@case
def r38_abandon_keeps_ei():
    c = "R38"
    p, s, crash, unl = crash_receipt(c, "r38", point=3)
    try:
        ab = s.B("ActionAbandonTail")
        nx = s.sysv("NEXT_EI")
        again = s.Rc("ReceiptPostRow", 1)
        R.add(c, "«Отложить хвост» с приходом ЕИ-00000201: номер этого ЕИ больше не выдаётся (NEXT_EI поднят до 202), новый приход получает ЕИ-00000202",
              "отложено операций: 1" in ab and nx == 202 and again.startswith("OK") and s.ord(1, "V") == ei(202), f"{ab[:120]}; NEXT_EI {nx}; {again}")
        oracle(c, s, "отложенный хвост")
    finally:
        s.close()


@case
def r39_many_receipts_and_selfcheck():
    c = "R39"
    s = ready(Session(new_wms("r39")))
    try:
        n = 40
        for i in range(n):
            order(s, 1 + i, a=f"З-{100 + i // 4}", b=f"Товар заказа {i}", e=f"A-{i}", h="10", q=FUTURE)
            fact(s, 1 + i, str(1 + i % 10), c=f"УПД-{i}")
        t0 = time.time()
        res = [s.Rc("ReceiptPostRow", 1 + i) for i in range(n)]
        dt = (time.time() - t0) / n * 1000
        t0 = time.time()
        adds = [s.Rc("ReceiptAddRow", 1 + i, "1", "", f"УПД-Д{i}", D0, D0, "B-1", "") for i in range(10)]
        dt2 = (time.time() - t0) / 10 * 1000
        sc = s.B("ActionSelfCheck")
        R.add(c, f"{n} приходов подряд и 10 доп. поступлений; самопроверка без ошибок",
              all(r.startswith("OK") for r in res + adds) and sc.splitlines()[0].startswith("САМОПРОВЕРКА WMS: ошибок 0"),
              f"{dt:.1f} мс/приход, {dt2:.1f} мс/доп. поступление; {sc.splitlines()[0]}", timing=dict(receipt_ms=round(dt, 1), add_ms=round(dt2, 1)))
        oracle(c, s, f"{n} приходов")
    finally:
        s.close()


@case
def r40_protection():
    c = "R40"
    s = ready(Session(new_wms("r40")))
    try:
        order(s, 1, h="10")
        fact(s, 1, "10", g="10")
        post(s, 1)
        before = s.ords(1)
        locked_changed, open_ok = [], []
        for col in "A B C D E F G H I J L N O P S U V W X Y AB".split():
            s.type_ord(col, 1, "999")
            if s.ords(1) != before:
                locked_changed.append(col)
        for col in "K M Q R T Z AA".split():
            s.type_ord(col, 1, "12.10.2026" if col == "Q" else "текст")
            open_ok.append(s.ord(1, col) != "")
        order(s, 2, a="З-2", h="5")
        before2 = s.ords(2)
        open_locked = []
        for col in "V W X Y AB".split():
            s.type_ord(col, 2, "ЕИ-00000300")
            if s.ords(2) != before2:
                open_locked.append(col)
        s.goto("Заказы", "$A$2")
        s.ui(".uno:SelectRow")
        s.ui(".uno:DeleteRows")
        s.goto("Заказы", "$A$3")
        s.ui(".uno:SelectRow")
        s.ui(".uno:DeleteRows")
        s.goto("Заказы", "$A$1:$AB$3")
        s.ui(".uno:SortDescending")
        after = (s.ord(1, "V"), s.ord(2, "A"))
        R.add(c, "защита: в проведённой строке закрыты A B C D E F G H I J L N O P S U V W X Y AB, открыты K M Q R T Z AA; в непроведённой закрыты V W X Y AB; "
                 "удаление строк и сортировка отклонены",
              not locked_changed and all(open_ok) and not open_locked and after == (ei(201), "З-2") and s.ord_locks(1) == POSTED,
              f"изменились закрытые: {locked_changed}; открытые записаны {open_ok}; в непроведённой изменились {open_locked}; после удаления/сортировки {after}")
        chk = s.click_ord("BtnRcvCheck", 1)
        R.add(c, "«Проверить» проведённой строки: приход сверен с «Наличие» и позицией — сообщение с остатком и статусом", chk.startswith("OK:ЕИ-00000201: остаток 10"),
              chk)
        oracle(c, s, "защита")
    finally:
        s.close()


@case
def r41_main_panel_unposted():
    c = "R41"
    p = new_wms("r41")
    s = ready(Session(p))
    order(s, 1, h="10")
    fact(s, 1, "10")
    order(s, 2, a="З-2", h="10")
    fact(s, 2, "3")
    post(s, 2)
    s.issue_input(1, ei="1", qty="1", date=D0, who="егр")
    s.close(save=True)
    s = ready(Session(p))
    try:
        un = s.main_status()[3]
        R.add(c, "«Главная» после открытия: непроведённые строки — выдача с ЕИ без № и строка заказа с фактом без ЕИ",
              "выдачи: 1 строк" in un and "приходы: 1 строк" in un, un)
    finally:
        s.close()


@case
def r42_crash_during_cancellations():
    c = "R42"
    p = new_wms("r42")
    s = ready(Session(p))
    order(s, 1, a="З-1", h="10")
    order(s, 2, a="З-2", h="100")
    fact(s, 2, "40")
    post(s, 2)
    s.close(save=True)
    s = ready(Session(p))
    s.B("TestSetFault", 2, 2)
    crash = s.Rc("OrderCancelRow", 1)
    s.kill()
    s = ready(Session(p))
    s.B("ActionForceUnlock")
    ok1 = crash == "CRASH-SIM" and W(s, 1) == "Ожидается" and s.ord_locks(1) == OPEN and s.next_ol() == 2.0 and s.opos(2)[0] == ""
    R.add(c, "сбой посреди «Отменить заказ» (до журнала): позиция снова ожидается, OLID не израсходован", ok1, f"{crash}; {W(s, 1)}; NEXT_OL {s.next_ol()}")
    s.B("TestSetFault", 3, 2)
    crash2 = s.Rc("OrderCancelRow", 1)
    s.kill()
    s = ready(Session(p))
    s.B("ActionForceUnlock")
    rec = s.B("ActionRecover")
    ok2 = (crash2 == "CRASH-SIM" and "восстановлено операций: 1" in rec and W(s, 1) == "Отменено" and s.next_ol() == 3.0 and s.opos(2)[8] == "ORDER"
           and s.sysv("NEXT_EI") == 202)
    R.add(c, "сбой после журнала при «Отменить заказ»: «Восстановить» отменяет позицию ровно один раз, ЕИ не создан", ok2, f"{crash2}; {rec[:100]}")
    s.B("TestSetFault", 3, 2)
    crash3 = s.Rc("OrderCancelRestRow", 2)
    s.kill()
    s = ready(Session(p))
    try:
        s.B("ActionForceUnlock")
        rec3 = s.B("ActionRecover")
        ok3 = crash3 == "CRASH-SIM" and "восстановлено операций: 1" in rec3 and W(s, 2) == "Частично получено / остаток отменён" and s.opos(1)[8] == "REST"
        R.add(c, "сбой после журнала при «Отменить остаток»: «Восстановить» доводит отмену остатка", ok3, f"{crash3}; {rec3[:100]}; {W(s, 2)}")
        oracle(c, s, "сбои при отменах")
    finally:
        s.close()


@case
def r43_autocalc_off():
    c = "R43"
    s = ready(Session(new_wms("r43")))
    try:
        order(s, 1, a="З-1", h="10", l="Петрович", e="ART-1")
        fact(s, 1, "10", c="УПД-5", o="24.09.2026")
        post(s, 1)
        s.doc.enableAutomaticCalculation(False)
        order(s, 2, a="З-2", h="10", l="Петрович", e="ART-1")
        fact(s, 2, "10", c="УПД-5", o="24.09.2026")
        res = post(s, 2)
        ab = s.ord(2, "AB")
        look = s.Od("TestLookupEI", "ЕИ-00000202")
        R.add(c, "автопересчёт Calc выключен пользователем: поиск формулами _IDX всё равно верен (метка-эхо → пересчёт) — дубль найден, "
                 "строка ЕИ найдена", res == "OK:2" and ab.startswith("Возможный дубль: ЕИ-00000201") and look == "2|1"
              and not s.doc.isAutomaticCalculationEnabled(), f"{res}; {ab!r}; поиск {look}")
        order(s, 3, a="З-3", h="10", l="Петрович", e="ART-3", q=FUTURE)
        fact(s, 3, "5", c="УПД-6", o="24.09.2026")
        p3 = post(s, 3)
        s.Od("TestSetToday", (TODAY + datetime.timedelta(days=40)).strftime("%d.%m.%Y"))
        s.OU("BtnOrdRefresh")
        w3, m3 = W(s, 3), s.U("TestUiLastMessage")
        s.Od("TestSetToday", "")
        s.OU("BtnOrdRefresh")
        m0 = s.U("TestUiLastMessage")
        R.add(c, "автопересчёт выключен: «Обновить статусы» считает позиции с прошедшей ожидаемой датой формулами Calc (свежесть результата "
                 "подтверждена меткой) — через 40 дней 1, сегодня 0; статус частично полученной позиции от даты не меняется («Частично получено»)",
              p3 == "OK:3" and w3 == "Частично получено" and W(s, 3) == "Частично получено" and "из них с прошедшей ожидаемой датой 1" in m3
              and "из них с прошедшей ожидаемой датой 0" in m0 and not s.doc.isAutomaticCalculationEnabled(), f"{p3}; {w3}; {m3}; {m0}")
        oracle(c, s, "без автопересчёта")
    finally:
        s.close()


# ================================================================ R44–R47 decisions of the acceptance (D-047, D-049, D-059)

def serial(dmy):
    d, m, y = (int(x) for x in dmy.split("."))
    return float((datetime.date(y, m, d) - datetime.date(1899, 12, 30)).days)


def add_row_of(res):
    """0-based row of the new receipt row from the result of «Ещё поступление» (OK:<seq>|строка <n>|<EI>)"""
    return int(res.split("|")[1].split()[1]) - 1


def dialog(s):
    """the fields of the last receipt window: (proposed, returned) — F G C O N U J each"""
    shown, _, back = s.OU("TestRcvDialog").partition("\n")
    return shown.split("\t"), (back.split("\t") if back else [])


@case
def r44_receipt_date():
    c = "R44"
    s = ready(Session(new_wms("r44")))
    try:
        order(s, 1, h="10")
        s.order_input(1, F="4", C="УПД-1", O="24.09.2026", U="A-1")
        before = (njournal(s), s.sysv("NEXT_EI"), s.next_ol())
        res = post(s, 1)
        row = s.ords(1)
        R.add(c, "первый приход из исходной строки с пустой датой поступления N (D-049): проведение заблокировано с объяснением, дата не подставлена, "
                 "ЕИ, позиции и записи журнала нет",
              res.startswith("ERR:") and "дата поступления" in res and row[13] == "" and row[21] == "" and row[24].startswith("Не проведено")
              and (njournal(s), s.sysv("NEXT_EI"), s.next_ol()) == before and s.ord_locks(1) == OPEN,
              f"{res}; N {row[13]!r}; V {row[21]!r}; Y {row[24][:90]}")
        s.order_input(1, N=D0)
        first = post(s, 1)
        today = "03.11.2026"
        s.Od("TestSetToday", today)
        s.OU("TestRcvInput", "3", "", "УПД-2", "", KEEP, "A-2", "")
        m1 = s.click_ord("BtnRcvMore", 1)
        shown1, back1 = dialog(s)
        r1 = add_row_of(m1) if m1.startswith("OK") else None
        n1 = s.ord(r1, "N") if r1 else None
        R.add(c, "«Ещё поступление»: окно предлагает в N сегодняшнюю дату (количество и документ пустые); подтверждено как есть — поступление с этой датой",
              first == "OK:1" and m1.startswith("OK:2") and shown1[4] == today and shown1[0] == "" and shown1[2] == "" and back1[4] == today
              and n1 == serial(today) and s.journal()[0][-1]["type"] == "RECEIPT_ADD", f"{first}; {m1}; предложено {shown1}; возвращено {back1}; N {n1!r}")
        s.OU("TestRcvInput", "2", "", "УПД-3", "", "01.11.2026", "A-2", "")
        m2 = s.click_ord("BtnRcvMore", 1)
        shown2, back2 = dialog(s)
        r2 = add_row_of(m2) if m2.startswith("OK") else None
        R.add(c, "пользователь изменил предложенную дату перед подтверждением: поступление проведено с его датой",
              m2.startswith("OK:3") and shown2[4] == today and back2[4:5] == ["01.11.2026"] and r2 is not None and s.ord(r2, "N") == serial("01.11.2026"),
              f"{m2}; предложено {shown2[4]}; возвращено {back2[4:5]}")
        nx = (njournal(s), s.sysv("NEXT_EI"))
        s.OU("TestRcvInput", "1", "", "УПД-4", "", "", "A-2", "")
        m3 = s.click_ord("BtnRcvMore", 1)
        R.add(c, "дата стёрта в окне: N обязательна — поступление не проведено, новой строки и ЕИ нет",
              m3.startswith("ERR:") and "дата поступления" in m3 and (njournal(s), s.sysv("NEXT_EI")) == nx and r2 is not None and s.ord(r2 + 1, "A") == "",
              m3)
        s.OU("TestRcvInput", KEEP, KEEP, KEEP, KEEP, KEEP, KEEP, KEEP)
        f1 = s.click_ord("BtnRcvFix", 1)
        shown3, back3 = dialog(s)
        R.add(c, "«Исправить» предлагает проведённые значения: дата поступления — проведённая, а не сегодняшняя; без правок — «ничего не изменилось»",
              shown3[4] == D0 and back3 == shown3 and f1.startswith("SKIP:ничего не изменилось"), f"{f1}; {shown3}")
        nx = (njournal(s), s.sysv("NEXT_EI"))
        miss = {}
        for k, col in enumerate("ABHIU"):
            vals = dict(A=f"З-{10 + k}", B="Товар", H="5", I="шт", F="5", N=D0, U="A-1")
            vals.pop(col)
            s.order_input(10 + k, **vals)
            miss[col] = post(s, 10 + k)
        R.add(c, "обязательные поля прихода (D-048): без № заказа, наименования, заказанного количества, единицы или места — отказ с названием поля, "
                 "ЕИ не создан", all(v.startswith("ERR:") and f"({col})" in v for col, v in miss.items()) and (njournal(s), s.sysv("NEXT_EI")) == nx,
              str(miss))
        oracle(c, s, "дата поступления", today=today)
        s.Od("TestSetToday", "")
    finally:
        s.close()


@case
def r45_status_without_date():
    c = "R45"
    p = new_wms("r45")
    s = ready(Session(p))
    try:
        order(s, 1, a="З-1", h="100", q=PAST)
        fact(s, 1, "40", g="40")
        order(s, 2, a="З-2", h="100", q=FUTURE)
        fact(s, 2, "40", g="40", c="УПД-2")
        order(s, 3, a="З-3", h="10", q=PAST)
        order(s, 4, a="З-4", h="10", q=PAST)
        fact(s, 4, "4", c="УПД-4")
        order(s, 5, a="З-5", h="10", q=PAST)
        fact(s, 5, "10", c="УПД-5")
        order(s, 6, a="З-6", h="10", q=PAST)
        s.order_input(6, F="10", N=D0, U="A-1")
        res = [post(s, r) for r in (1, 2, 4, 5, 6)]
        cr = s.click_ord("BtnOrdCancelRest", 4)
        ws = [W(s, r) for r in range(1, 7)]
        want = ["Частично получено", "Частично получено", "Ожидается", "Частично получено / остаток отменён", "Получено", "Получено без документов"]
        R.add(c, "статусы без даты (D-089 вместо D-047): частично получено — «Частично получено» и при прошедшей Q; ничего не получено и Q "
                 "прошла — «Ожидается»; остаток отменён — «Частично получено / остаток отменён»; получено полностью — «Получено» / «Получено без "
                 "документов»; Y исходной строки показывает «осталось 60»",
              res == ["OK:1", "OK:2", "OK:3", "OK:4", "OK:5"] and cr.startswith("OK") and ws == want and "осталось 60" in s.ord(1, "Y")
              and "осталось 60" in s.ord(2, "Y"), f"{res}; {cr}; {ws}; Y1 {s.ord(1, 'Y')}")
        later = (TODAY + datetime.timedelta(days=40)).strftime("%d.%m.%Y")
        s.Od("TestSetToday", later)
        s.OU("BtnOrdRefresh")
        msg = s.U("TestUiLastMessage")
        ws2 = [W(s, r) for r in range(1, 7)]
        R.add(c, "через 40 дней («Обновить статусы»): ни один статус не изменился — WMS продолжает ждать поставку; Y по-прежнему показывает, "
                 "сколько осталось; в сообщении — позиций с прошедшей ожидаемой датой 3",
              ws2 == want and "осталось 60" in s.ord(2, "Y") and "ожидаемых позиций 3, из них с прошедшей ожидаемой датой 3" in msg
              and "обновлено статусов 0" in msg, f"{ws2}; {msg}")
        s.order_input(1, Q=(TODAY + datetime.timedelta(days=60)).strftime("%d.%m.%Y"))
        w_later = W(s, 1)
        s.order_input(1, Q=PAST)
        w_back = W(s, 1)
        R.add(c, "Q исходной строки после прихода открыта: новая ожидаемая дата (позже и раньше текущей) статус не меняет — «Частично получено»",
              w_later == "Частично получено" and w_back == "Частично получено" and s.ord_locks(1) == POSTED, f"{w_later}; {w_back}")
        m = more(s, 1, "60", c="УПД-9", o=D0)
        w_full = W(s, 1)
        nr = add_row_of(m) if m.startswith("OK") else None
        d = s.click_ord("BtnRcvDelete", nr) if nr else ""
        w_storno = W(s, 1)
        R.add(c, "остаток пришёл («Ещё поступление» 60): «Получено»; сторно этого поступления — снова «Частично получено», осталось 60",
              m.startswith("OK") and w_full == "Получено" and d.startswith("OK") and w_storno == "Частично получено" and "осталось 60" in s.ord(1, "Y"),
              f"{m}; {w_full}; {d}; {w_storno}; Y1 {s.ord(1, 'Y')}")
        oracle(c, s, "статусы без даты", today=later)
        s.Od("TestSetToday", "")
        # a book of an earlier core: its date statuses become the statuses of the data (M6 §18: the cancelled delivery row of
        # З-1 stays right under it — row 3 of the sheet; the positions follow)
        src_rows = (1, 3, 4, 5, 6, 7)
        sh = s.doc.Sheets.getByName("Заказы")
        sh.unprotect("wms")
        sh.getCellByPosition(22, 1).setString("Частично получено / просрочено")
        sh.getCellByPosition(22, 4).setString("Просрочено")
        sh.protect("wms")
        s.OU("BtnOrdRefresh")
        msg2 = s.U("TestUiLastMessage")
        wa = [W(s, r) for r in src_rows]
        R.add(c, "книга прежней версии (в W «Частично получено / просрочено», «Просрочено»): «Обновить статусы» делает их статусами по данным — "
                 "«Частично получено», «Ожидается»; остальные строки не тронуты",
              wa == want and W(s, 2) == "Поступление удалено (сторно)" and "статусов по дате прежней версии пересчитано 2" in msg2
              and "обновлено статусов 2" in msg2, f"{wa}; {msg2}")
        sh.unprotect("wms")
        sh.getCellByPosition(22, 4).setString("Просрочено")
        sh.protect("wms")
        s.close(save=True)
        s = ready(Session(p))
        rep = s.report()
        ws3 = [W(s, r) for r in src_rows]
        R.add(c, "повторное открытие (настоящая дата): статус по дате прежней версии пересчитан при запуске («Просрочено» → «Ожидается»), "
                 "остальные — как были; в отчёте запуска — позиций с прошедшей ожидаемой датой 2",
              ws3 == want and "статусов по дате прежней версии пересчитано 1" in rep and "из них с прошедшей ожидаемой датой 2" in rep,
              f"{ws3}; {rep[-320:]}")
        oracle(c, s, "статусы после открытия")
    finally:
        s.close()


@case
def r46_block_posting():
    c = "R46"
    s = ready(Session(new_wms("r46")))
    try:
        for r in range(1, 7):
            order(s, r, a=f"З-{r}", h="10")
        fact(s, 1, "10")
        s.order_input(2, F="10", U="A-1")                       # no receipt date N
        fact(s, 4, "5", c="УПД-4")
        s.order_input(5, F="abc", N=D0, U="A-1")               # not a quantity
        fact(s, 6, "10", c="УПД-6")
        head, first, sysl = s.click_ord("BtnRcvPost", 1, r1=6).split("\n", 2)
        msg = s.OU("TestRcvBlockMessage")
        rc = [e for e in s.journal()[0] if e["type"] == "RECEIPT"]
        R.add(c, "блок из 6 строк (D-059): каждая строка с F — отдельная операция; ошибки строк (нет даты N, неверное F) не останавливают остальные; "
                 "строка без F пропущена; окно: проведено 3, пропущено 1, отклонено с ошибкой 2",
              head == "OK=3;ERR=2;SKIP=1;STOP=0" and sysl == "" and first.startswith("строка 3:") and "дата поступления" in first
              and [s.ord(r, "V") for r in range(1, 7)] == [ei(201), "", "", ei(202), "", ei(203)] and s.ord(2, "Y").startswith("Не проведено")
              and s.ord(5, "Y").startswith("Не проведено") and len(rc) == 3 and len({e["seq"] for e in rc}) == 3
              and "проведено 3, пропущено 1" in msg and "отклонено с ошибкой 2" in msg, f"{head}; {first}; {msg[:220]}")
        for r in range(7, 11):
            order(s, r, a=f"З-{r}", h="5")
            fact(s, r, "5", c=f"УПД-{r}")
        sh = s.doc.Sheets.getByName("Заказы")
        sh.unprotect(PWD)
        sh.getCellByPosition(21, 12).setString(ei(205))       # the EI the counter will give next but one is already written on the sheet
        sh.protect(PWD)
        nj = njournal(s)
        head2, _, sys2 = s.click_ord("BtnRcvPost", 7, r1=10).split("\n", 2)
        msg2 = s.OU("TestRcvBlockMessage")
        R.add(c, "системная ошибка WMS посреди блока (номер, который выдал бы счётчик ЕИ, уже записан на листе): строка 8 проведена, на строке 9 обработка "
                 "остановлена, строки 10–11 не обрабатывались; окно показывает итог и причину остановки",
              head2 == "OK=1;ERR=0;SKIP=0;STOP=9" and "ERR-SYS:" in sys2 and ei(205) in sys2 and s.ord(7, "V") == ei(204)
              and [s.ord(r, "V") for r in (8, 9, 10)] == ["", "", ""] and s.ord(9, "Y") == "" and s.ord(10, "Y") == "" and njournal(s) == nj + 1
              and s.sysv("NEXT_EI") == 205 and "ОБРАБОТКА ОСТАНОВЛЕНА" in msg2 and "проведено 1" in msg2, f"{head2}; {sys2[:170]}; {msg2[-220:]}")
        cl = s.click_ord("BtnRcvClear", 12)
        head2b = s.click_ord("BtnRcvPost", 7, r1=10).split("\n")[0]
        R.add(c, "причина устранена (копия очищена кнопкой «Очистить»): тот же блок — проведённая строка пропущена, остальные три проведены по одной операции",
              cl.startswith("OK") and head2b == "OK=3;ERR=0;SKIP=1;STOP=0" and [s.ord(r, "V") for r in (8, 9, 10)] == [ei(205), ei(206), ei(207)],
              f"{cl}; {head2b}")
        for r in (14, 15, 16):
            order(s, r, a=f"З-{r}", h="3")
            fact(s, r, "3", c=f"УПД-{r}")
        nj, nx = njournal(s), s.sysv("NEXT_EI")
        s.B("TestSetFault", 4, 1)                              # the next journal line is written only in part, then the write fails
        head3, _, sys3 = s.click_ord("BtnRcvPost", 14, r1=16).split("\n", 2)
        R.add(c, "ошибка журнала на первой строке блока (запись оборвалась): операция откатана, обработка блока сразу остановлена, остальные строки не "
                 "обрабатывались", head3 == "OK=0;ERR=0;SKIP=0;STOP=15" and "ERR-RB:" in sys3 and [s.ord(r, "V") for r in (14, 15, 16)] == ["", "", ""]
              and njournal(s) == nj and s.sysv("NEXT_EI") == nx and st(s)[0] == "CLEAN", f"{head3}; {sys3[:170]}")
        s.B("TestSetFault", 2, 1)                              # a write inside ApplyOperation fails
        head4, _, sys4 = s.click_ord("BtnRcvPost", 14, r1=16).split("\n", 2)
        R.add(c, "ошибка ApplyOperation на первой строке блока: откат, обработка блока остановлена",
              head4 == "OK=0;ERR=0;SKIP=0;STOP=15" and "ERR-RB:" in sys4 and [s.ord(r, "V") for r in (14, 15, 16)] == ["", "", ""] and njournal(s) == nj,
              f"{head4}; {sys4[:170]}")
        head5 = s.click_ord("BtnRcvPost", 14, r1=16).split("\n")[0]
        R.add(c, "после сбоя тот же блок проводится полностью: три строки — три операции, номера ЕИ подряд",
              head5 == "OK=3;ERR=0;SKIP=0;STOP=0" and [s.ord(r, "V") for r in (14, 15, 16)] == [ei(208), ei(209), ei(210)] and njournal(s) == nj + 3,
              head5)
        oracle(c, s, "проведение блоков")
    finally:
        s.close()


@case
def r47_block_posting_blocked():
    c = "R47"
    p = new_wms("r47")
    s = ready(Session(p))
    for r in range(1, 5):
        order(s, r, a=f"З-{r}", h="4")
        fact(s, r, "4", c=f"УПД-{r}")
    s.close(save=True)
    s = ready(Session(p))
    try:
        s.B("TestSetFault", 7, 3)                              # the completion of the operation fails and keeps failing: WMS blocks itself
        head, _, sysl = s.click_ord("BtnRcvPost", 1, r1=4).split("\n", 2)
        state1 = st(s)
        head2, _, sys2 = s.click_ord("BtnRcvPost", 1, r1=4).split("\n", 2)
        R.add(c, "критическая ошибка на первой строке блока (операция в журнале, завершение не удалось): WMS заблокирована, обработка остановлена, строки "
                 "3–5 не обрабатывались; повторное нажатие — блокировка останавливает блок на первой же непроведённой строке",
              head == "OK=0;ERR=0;SKIP=0;STOP=2" and "ERR-CRITICAL:" in sysl and state1[0] == "BLOCKED" and head2 == "OK=0;ERR=0;SKIP=1;STOP=3"
              and "BLOCKED:" in sys2 and [s.ord(r, "V") for r in (2, 3, 4)] == ["", "", ""], f"{head}; {sysl[:150]}; {state1}; {head2}; {sys2[:120]}")
    finally:
        s.close()
    s = ready(Session(p))
    try:
        _, block = st(s)
        rec = s.B("ActionRecover")
        head3 = s.click_ord("BtnRcvPost", 1, r1=4).split("\n")[0]
        R.add(c, "перезапуск: журнал впереди книги, «Восстановить» доводит приход строки 2 ровно один раз; затем блок проводит строки 3–5, строка 2 пропущена",
              block == "TAIL" and "восстановлено операций: 1" in rec and head3 == "OK=3;ERR=0;SKIP=1;STOP=0"
              and [s.ord(r, "V") for r in (1, 2, 3, 4)] == [ei(201), ei(202), ei(203), ei(204)], f"{block}; {rec[:120]}; {head3}")
        oracle(c, s, "блок после блокировки и восстановления")
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
    R.save(os.path.join(OUT, "results_phase3.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_phase3.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
