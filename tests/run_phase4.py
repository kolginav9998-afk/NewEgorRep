"""Core Phase 4 «Возвраты» — automated scenarios (CLAUDE_TASK.md, section «ОБЯЗАТЕЛЬНЫЕ ТЕСТЫ»).

    WMS_TEST_OUT=/tmp/wms_out python3 tests/run_phase4.py [case ...]

Every case copies a freshly built test book (synthetic EIs 1–200, NEXT_EI 201, five recipients) into
WMS_TEST_OUT/cases/<case>/ and drives real LibreOffice processes like a user: typing through the UI (the sheet's change
handler runs), the macros the buttons of «Возврат», «Выдачи» and «Заказы» are bound to for the row under the cursor
(confirmations answered «Да» in the test mode, window values preset), window close, kill -9. The book and the journal
are checked by the independent oracle tests/return_oracle.py (it runs tests/receipt_oracle.py: receipts, issues,
returns, «Наличие», «Заказы».X, «Выдачи», «Возврат», _RET, _ISS).
Results: WMS_TEST_OUT/results_phase4.json and WMS_TEST_OUT/TEST_REPORT_phase4.md.
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
from harness import OUT, Session, Results, new_wms, template, INITIAL_STOCK  # noqa: E402
from com.sun.star.sheet import TableFilterField  # noqa: E402
from com.sun.star.sheet.FilterOperator import EQUAL  # noqa: E402

R = Results()
CASES = []
POSTED = "111111111111110"          # WmsConfig.RETURN_LOCKS_POSTED
OPEN = "100110100011110"            # WmsConfig.RETURN_LOCKS_OPEN
TODAY = datetime.date.today()
DI = "20.09.2026"                   # the date of the issues
DR = "22.09.2026"                   # the date of the returns
KEEP = "<как предложено>"           # WmsReturnUi test seam: the window field keeps the value it proposed
PWD = "wms"
IVA, PET, EGR = "Иванов Иван Андреевич", "Петров Пётр Сергеевич", "Ермолин Егор Павлович"


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


def oracle(c, s, name, expect_tail=0, legacy_mn_empty=True):
    P, inf, hist = s.retcheck(expect_tail, legacy_mn_empty=legacy_mn_empty)
    R.add(c, "оракул: " + name, not P, f"{inf}; {P[:4]}")
    return hist


def issue(s, r, e, q, who="ива", date=DI):
    s.issue_input(r, ei=e, qty=q, date=date, who=who)
    return s.post(r)


def ret_row(s, r, k=None, q=None, date=DR, e=None, who=None, place=None):
    s.return_input(r, issue=k, ei=e, qty=q, date=date, who=who, place=place)


def post_ret(s, r):
    return s.click_ret("BtnRetPost", r)


def ret(s, r, k, q, date=DR, **kw):
    ret_row(s, r, k, q, date, **kw)
    return post_ret(s, r)


def fix_ret(s, r, q=KEEP, date=KEEP, place=KEEP):
    s.RU("TestRetInput", q, date, place, -1)
    return s.click_ret("BtnRetFix", r)


def del_ret(s, r):
    return s.click_ret("BtnRetDelete", r)


def N(s, r):
    return s.ret(r)[13]


def dialog(s):
    """the last return window: (proposed, returned, list) — quantity, date, place, line"""
    shown, back, lst = (s.RU("TestRetDialog").split("\n", 2) + ["", ""])[:3]
    return shown.split("\t"), (back.split("\t") if back else []), [x for x in lst.split("\n") if x]


def order(s, r, a="З-1", b="Болт М10", h="100", i="шт", l="Петрович", u="A-1", f=None, n="24.09.2026"):
    s.order_input(r, A=a, B=b, H=h, I=i, L=l, F=f, N=n, U=u, C="УПД-1", O="24.09.2026")


def snap(s, rows=8, eis=(1, 7, 10, 12, 201)):
    vals = s.doc.Sheets.getByName("Возврат").getCellRangeByPosition(0, 1, 14, rows).getDataArray()
    locks = [s.ret_locks(r) for r in range(1, rows + 1)]
    stock = [s.stock_row(n) for n in eis]
    rt = [s.retrec(n) for n in range(1, 5)]
    isr = [s.issrec(k) for k in range(1, 5)]
    iss = [s.iss(r) for r in range(1, 5)]
    sysv = [s.sysv(k) for k in ("LAST_SEQ", "NEXT_EI", "NEXT_NO", "NEXT_RET", "TX_STATE", "TX_SEQ", "TX_BEFORE_IMAGE")]
    return vals, locks, stock, rt, isr, iss, sysv


def ctrl_z_harmless(s, n=5):
    a = snap(s)
    for _ in range(n):
        s.ui(".uno:Undo")
    return a == snap(s)


# ================================================================ W01–W09 returns: full, partial, several, rules

@case
def w01_full_return():
    c = "W01"
    s = ready(Session(new_wms("w01")))
    try:
        r1 = issue(s, 1, "1", "10")
        s.type_ret("B", 1, "1")
        pv = s.ret(1)
        want_pv = ("", 1.0, ei(1), "Болт М8х30 оцинкованный", "DIN933-M8x30", "", "шт", "", IVA, "A-01-1", "Крепёж", 90.0, "")
        R.add(c, "№ выдачи в B → подставлены ЕИ, наименование, артикул, единица, получатель, место, категория, остаток; «Контроль» показывает "
                 "выдано / возвращено / можно вернуть", r1 == "OK:1" and pv[:13] == want_pv
              and pv[13] == f"Выдача № 1 от {DI}, «{IVA}»: выдано 10 шт, возвращено 0, можно вернуть 10", str(pv))
        s.return_input(1, qty="10", date=DR)
        ready_txt = N(s, 1)
        res = post_ret(s, 1)
        row = s.ret(1)
        ok = (res == "OK:2" and row[:13] == (1.0, 1.0, ei(1), "Болт М8х30 оцинкованный", "DIN933-M8x30", 10.0, "шт", serial(DR), IVA, "A-01-1",
                                               "Крепёж", 90.0, 100.0)
              and row[13] == "Проведено" and s.ret_locks(1) == POSTED and s.stock(1) == 100.0 and s.sysv("NEXT_RET") == 2.0
              and s.retrec(1) == (1.0, 1.0, ei(1), 10.0, "LIVE", 1.0) and s.issrec(1) == (1.0, 1.0, 10.0, 1.0))
        R.add(c, "полный возврат: № возврата 1, остаток того же ЕИ 90 → 100, строка защищена, _RET и _ISS записаны одной операцией",
              ready_txt.startswith("Можно провести:") and ok, f"{ready_txt[:60]}; {res}; {row}; _RET {s.retrec(1)}; _ISS {s.issrec(1)}")
        iss_row = s.iss(1)
        R.add(c, "строка выдачи не изменилась: количество 10, статус «Проведено», legacy M/N пусты; новых ЕИ нет",
              iss_row[4] == 10.0 and iss_row[17] == "Проведено" and iss_row[12] == "" and iss_row[13] == "" and s.sysv("NEXT_EI") == 201.0, str(iss_row))
        s.type_ret("B", 2, "1")
        pv2 = N(s, 2)
        s.return_input(2, qty="1", date=DR)
        r2 = post_ret(s, 2)
        R.add(c, "после полного возврата выдача закрыта для возвратов: ещё 1 — отказ («всё уже возвращено»), ничего не записано",
              "возвращено 10, можно вернуть 0" in pv2 and r2.startswith("ERR") and "всё уже возвращено" in r2 and s.stock(1) == 100.0
              and s.sysv("NEXT_RET") == 2.0 and N(s, 2).startswith("Не проведено"), f"{pv2}; {r2}")
        oracle(c, s, "полный возврат")
    finally:
        s.close()


@case
def w02_partial_return_receipt_ei():
    c = "W02"
    s = ready(Session(new_wms("w02")))
    try:
        order(s, 1, h="100", f="40")
        rc = s.click_ord("BtnRcvPost", 1)
        i1 = issue(s, 1, "201", "10", who="пет")
        x1 = s.ord(1, "X")
        r1 = ret(s, 1, "1", "3")
        x2, bal = s.ord(1, "X"), s.stock(201)
        R.add(c, "частичный возврат 3 из 10 по ЕИ из прихода: остаток ЕИ 30 → 33, «Заказы».X этого ЕИ обновлён той же операцией (33)",
              rc == "OK:1" and i1 == "OK:2" and x1 == 30.0 and r1 == "OK:3" and x2 == 33.0 and bal == 33.0 and s.ret(1)[13] == "Проведено",
              f"{rc}; {i1}; X {x1} → {x2}; остаток {bal}; {r1}")
        s.type_ret("B", 2, "1")
        R.add(c, "по той же выдаче можно вернуть ещё 7", "возвращено 3, можно вернуть 7" in N(s, 2), N(s, 2))
        oracle(c, s, "частичный возврат")
    finally:
        s.close()


@case
def w03_second_third_return():
    c = "W03"
    s = ready(Session(new_wms("w03")))
    try:
        issue(s, 1, "12", "10")
        a = ret(s, 1, "1", "3")
        b = ret(s, 2, "1", "2", date="23.09.2026")
        c6 = ret(s, 3, "1", "6", date="24.09.2026")
        why6 = N(s, 3)
        s.return_input(3, qty="5")
        d = post_ret(s, 3)
        R.add(c, "выдали 10 → вернули 3 → ещё 2 → попытка 6 блокируется («можно вернуть 5») → оставшиеся 5 проведены; три строки, три № возврата, "
                 "тот же ЕИ, разные даты",
              a == "OK:2" and b == "OK:3" and c6.startswith("ERR") and "можно вернуть 5" in why6 and d == "OK:4"
              and [s.ret(r)[0] for r in (1, 2, 3)] == [1.0, 2.0, 3.0] and [s.ret(r)[2] for r in (1, 2, 3)] == [ei(12)] * 3
              and [s.ret(r)[7] for r in (1, 2, 3)] == [serial(DR), serial("23.09.2026"), serial("24.09.2026")]
              and s.stock(12) == 35.0 and s.issrec(1) == (1.0, 1.0, 10.0, 3.0), f"{a}; {b}; {c6} ({why6[:80]}); {d}; _ISS {s.issrec(1)}")
        e = ret(s, 4, "1", "0,001")
        R.add(c, "всё возвращено: ещё 0,001 — отказ", e.startswith("ERR") and s.stock(12) == 35.0, e)
        oracle(c, s, "три возврата одной выдачи")
    finally:
        s.close()


@case
def w04_two_issues_one_ei():
    c = "W04"
    s = ready(Session(new_wms("w04")))
    try:
        i1 = issue(s, 1, "10", "5", who="ива")
        i2 = issue(s, 2, "10", "3", who="пет")
        r1 = ret(s, 1, "1", "5")
        s.type_ret("B", 2, "2")
        pv2 = N(s, 2)
        extra = ret(s, 3, "1", "1")
        r2 = ret(s, 2, "2", "3")
        R.add(c, "ЕИ-10: выдача №1 Иванову 5, №2 Петрову 3; возврат 5 по №1 не уменьшает обязательство №2 (по ней можно вернуть 3); "
                 "ещё 1 по №1 — отказ; 3 по №2 — проведено",
              (i1, i2, r1, r2) == ("OK:1", "OK:2", "OK:3", "OK:4") and "можно вернуть 3" in pv2 and extra.startswith("ERR") and s.stock(10) == 21.0
              and s.issrec(1) == (1.0, 1.0, 5.0, 1.0) and s.issrec(2) == (2.0, 2.0, 3.0, 1.0) and s.ret(2)[8] == PET and s.ret(1)[8] == IVA,
              f"{i1} {i2} {r1} {r2}; {pv2[:90]}; {extra}")
        hist = oracle(c, s, "две выдачи одного ЕИ")
        R.add(c, "история получателя (оракул): Иванов/ЕИ-10 выдано 5, возвращено 5, чистое 0; Петров/ЕИ-10 3, 3, 0",
              hist.get((IVA, ei(10))) == dict(issued=5.0, returned=5.0, net=0.0) and hist.get((PET, ei(10))) == dict(issued=3.0, returned=3.0, net=0.0),
              str(hist))
    finally:
        s.close()


@case
def w05_find_issue_dialog():
    c = "W05"
    s = ready(Session(new_wms("w05")))
    try:
        issue(s, 1, "7", "4", who="ива")
        issue(s, 2, "7", "6", who="пет")
        issue(s, 3, "7", "2", who="ива", date="23.09.2026")
        s.type_ret("C", 1, "7")
        s.RU("TestRetInput", "2", KEEP, KEEP, 2)
        f1 = s.click_ret("BtnRetFind", 1)
        shown, back, lst = dialog(s)
        today = TODAY.strftime("%d.%m.%Y")
        row = s.ret(1)
        R.add(c, "«Найти выдачу» по ЕИ в C: окно со списком всех выдач ЕИ-7, которые можно вернуть; дата возврата предложена сегодняшняя, место — "
                 "текущее место ЕИ; выбрана третья строка → в строку записаны № выдачи 3, количество, дата; «Контроль» — можно провести",
              len(lst) == 3 and lst[0].startswith(f"№ 1 от {DI} — «{IVA}»: выдано 4, возвращено 0, можно вернуть 4") and "«" + PET in lst[1]
              and shown[1] == today and shown[2] == "S-07" and back[3] == "2" and row[1] == 3.0 and row[5] == "2" and row[7] == serial(today)
              and row[8] == IVA and row[9] == "S-07" and row[13].startswith("Можно провести") and f1.startswith("OK"),
              f"{f1}; предложено {shown}; возвращено {back}; список {lst}; строка {row[:10]}")
        p1 = post_ret(s, 1)
        s.return_input(2, ei="7", who="ива")
        s.RU("TestRetInput", "1", "22.09.2026", KEEP, -1)
        f2 = s.click_ret("BtnRetFind", 2)
        _, _, lst2 = dialog(s)
        p2 = post_ret(s, 2)
        R.add(c, "получатель в I сужает список: только выдачи Иванова, которые ещё можно вернуть (№3 полностью возвращена — её нет); дата изменена в окне",
              p1 == "OK:4" and len(lst2) == 1 and lst2[0].startswith("№ 1 ") and f2.startswith("OK") and s.ret(2)[1] == 1.0 and s.ret(2)[7] == serial(DR)
              and p2 == "OK:5", f"{p1}; {lst2}; {f2}; {p2}")
        s.type_ret("B", 3, "2")
        s.RU("TestRetInput", "6", KEEP, "Б-1", -1)
        f3 = s.click_ret("BtnRetFind", 3)
        _, _, lst3 = dialog(s)
        p3 = post_ret(s, 3)
        R.add(c, "№ выдачи в B: окно этой выдачи (одна строка), место изменено на «Б-1» — оно становится текущим местом ЕИ той же операцией",
              len(lst3) == 1 and lst3[0].startswith("№ 2 ") and f3.startswith("OK") and p3 == "OK:6" and s.stock_row(7)[5] == "Б-1", f"{lst3}; {p3}")
        s.RU("TestRetKey", "")
        f4 = s.click_ret("BtnRetFind", 4)
        s.RU("TestRetKey", "ЕИ-00000008")
        f5 = s.click_ret("BtnRetFind", 4)
        s.RU("TestRetKey", "ЕИ-00000999")
        f6 = s.click_ret("BtnRetFind", 4)
        s.RU("TestRetKey", "7")
        s.RU("TestRetInput", "1", KEEP, KEEP, -1)
        f7 = s.click_ret("BtnRetFind", 4)
        _, _, lst7 = dialog(s)
        R.add(c, "пустая строка: вопрос «ЕИ» — отмена ничего не делает; ЕИ без выдач и неизвестный ЕИ — понятный отказ; ЕИ-7 — список из одной "
                 "оставшейся выдачи (№1: можно вернуть 3)",
              f4.startswith("SKIP") and f5.startswith("ERR") and "нет проведённых выдач" in f5 and f6.startswith("ERR") and "не найден" in f6
              and len(lst7) == 1 and "можно вернуть 3" in lst7[0] and s.ret(4)[1] == 1.0, f"{f4}; {f5[:80]}; {f6[:80]}; {lst7}")
        s.U("TestUiAuto", 1)
        f8 = s.click_ret("BtnRetFind", 1)
        R.add(c, "«Найти выдачу» на проведённой строке — отказ", f8.startswith("ERR") and "проведённый возврат" in f8, f8)
        oracle(c, s, "«Найти выдачу»")
    finally:
        s.close()


@case
def w06_return_at_zero_balance():
    c = "W06"
    s = ready(Session(new_wms("w06")))
    try:
        i1 = issue(s, 1, "5", "5")
        z = s.stock(5)
        r1 = ret(s, 1, "1", "2")
        R.add(c, "возврат при текущем остатке ЕИ = 0: разрешён, остаток 0 → 2, тот же ЕИ", i1 == "OK:1" and z == 0.0 and r1 == "OK:2" and s.stock(5) == 2.0
              and s.ret(1)[11] == 0.0 and s.ret(1)[12] == 2.0 and s.sysv("NEXT_EI") == 201.0, f"{i1}; {z}; {r1}; {s.ret(1)[11:13]}")
        oracle(c, s, "возврат при нулевом остатке")
    finally:
        s.close()


@case
def w07_more_than_issued():
    c = "W07"
    s = ready(Session(new_wms("w07")))
    try:
        issue(s, 1, "13", "4")
        a = ret(s, 1, "1", "5")
        why = N(s, 1)
        s.return_input(1, qty="4,001")
        b = post_ret(s, 1)
        s.return_input(1, qty="4")
        d = post_ret(s, 1)
        R.add(c, "возврат больше выданного (5 из 4, 4,001) — отказ с причиной в «Контроль»; ровно 4 — проведено",
              a.startswith("ERR") and "можно вернуть 4" in why and b.startswith("ERR") and d == "OK:2" and s.stock(13) == 42.0, f"{a}; {b}; {d}")
        oracle(c, s, "больше выданного")
    finally:
        s.close()


@case
def w08_wrong_recipient():
    c = "W08"
    s = ready(Session(new_wms("w08")))
    try:
        issue(s, 1, "14", "3", who="ива")
        s.return_input(1, who="пет")
        s.return_input(1, issue="1", qty="1", date=DR)
        why = N(s, 1)
        a = post_ret(s, 1)
        s.return_input(1, who="Сидоров")
        b = post_ret(s, 1)
        s.return_input(1, who="иванов иван андреевич")
        d = post_ret(s, 1)
        R.add(c, "от кого возвращено: другой получатель (Петров, Сидоров) — отказ (будущий сценарий передачи); тот же получатель в другом регистре — "
                 "проведено", a.startswith("ERR") and "не поддерживается" in why and "Петров" in why and b.startswith("ERR") and d == "OK:2"
              and s.ret(1)[8] == IVA, f"{why[:120]}; {a[:60]}; {b[:60]}; {d}; I {s.ret(1)[8]!r}")
        s.type_ret("B", 2, "1")
        R.add(c, "пустое I заполняется получателем выдачи автоматически", s.ret(2)[8] == IVA, s.ret(2)[8])
        oracle(c, s, "получатель")
    finally:
        s.close()


@case
def w09_unknown_ei_and_issue():
    c = "W09"
    s = ready(Session(new_wms("w09")))
    try:
        issue(s, 1, "16", "3")
        out = []
        for r, (b, e) in enumerate([("999", None), ("abc", None), ("0", None), ("1", "ЕИ-00000999"), ("1", "17"), (None, "ЕИ-00000999")], start=1):
            ret_row(s, r, b, "1", e=e)
            out.append((b, e, post_ret(s, r)[:4], N(s, r)[:110]))
        ok = (all(x[2] == "ERR:" for x in out) and "выдачи № 999 нет" in out[0][3] and "не номер" in out[1][3] and "не номер" in out[2][3]
              and "не совпадает с ЕИ выдачи" in out[3][3] and "не совпадает с ЕИ выдачи" in out[4][3] and "не указан № выдачи" in out[5][3])
        R.add(c, "неизвестный № выдачи (999, текст, 0), ЕИ, которого нет, другой ЕИ, ЕИ без № выдачи — отказ с понятной причиной, ничего не записано",
              ok and njournal(s) == 1 and s.sysv("NEXT_RET") == 1.0, str(out))
        s.type_ret("C", 7, "ЕИ-00000999")
        pv = N(s, 7)
        R.add(c, "предпросмотр ЕИ, которого нет в реестре", "не найден" in pv, pv)
        oracle(c, s, "неизвестные ЕИ и выдачи")
    finally:
        s.close()


# ================================================================ W10–W15 correction, storno, ISSUE_FIX / ISSUE_DEL

@case
def w10_fix_up_down():
    c = "W10"
    s = ready(Session(new_wms("w10")))
    try:
        issue(s, 1, "17", "10")
        ret(s, 1, "1", "3")
        up = fix_ret(s, 1, q="4")
        row_up = s.ret(1)
        down = fix_ret(s, 1, q="2")
        bal_down = s.stock(17)
        over = fix_ret(s, 1, q="11")
        R.add(c, "исправление возврата 3 → 4 (остаток +1), 4 → 2 (−2), 2 → 11 (выше доступного) — отказ; статус «Проведено (исправлено)»",
              up == "OK:3" and row_up[5] == 4.0 and row_up[11] == 10.0 and row_up[12] == 14.0 and row_up[13] == "Проведено (исправлено)"
              and down == "OK:4" and bal_down == 12.0 and over.startswith("ERR") and "можно не больше 10" in over and s.stock(17) == 12.0 and s.ret(1)[5] == 2.0,
              f"{up}; {row_up[5:14]}; {down}; {bal_down}; {over}")
        r2 = ret(s, 2, "1", "5")
        f_over = fix_ret(s, 1, q="6")
        f_ok = fix_ret(s, 1, q="5")
        R.add(c, "после второго возврата 5: исправление первого до 6 — отказ (другими возвратами возвращено 5, можно не больше 5); до 5 — проведено; "
                 "сумма возвратов выдачи = выданному", r2 == "OK:5" and f_over.startswith("ERR") and "можно не больше 5" in f_over and f_ok == "OK:6"
              and s.issrec(1) == (1.0, 1.0, 10.0, 2.0) and s.stock(17) == 20.0, f"{r2}; {f_over}; {f_ok}; {s.issrec(1)}")
        same = fix_ret(s, 1)
        early = fix_ret(s, 1, date="19.09.2026")
        mv = fix_ret(s, 1, place="К-5")
        R.add(c, "без изменений — «ничего не изменилось»; дата раньше выдачи — отказ; новое место — «Наличие».F = К-5 той же операцией",
              same.startswith("SKIP") and early.startswith("ERR") and "раньше даты выдачи" in early and mv == "OK:7" and s.stock_row(17)[5] == "К-5"
              and s.ret(1)[9] == "К-5", f"{same}; {early}; {mv}; {s.stock_row(17)[5]}")
        oracle(c, s, "исправление возврата")
    finally:
        s.close()


@case
def w11_storno_return():
    c = "W11"
    s = ready(Session(new_wms("w11")))
    try:
        issue(s, 1, "18", "10")
        ret(s, 1, "1", "4")
        d = del_ret(s, 1)
        row = s.ret(1)
        ok = (d == "OK:3" and row[0] == 1.0 and row[5] == 4.0 and row[13] == "Удалено (сторно)" and s.ret_locks(1) == POSTED and s.stock(18) == 17.0
              and s.retrec(1)[4] == "STORNO" and s.issrec(1) == (1.0, 1.0, 0.0, 0.0))
        R.add(c, "сторно возврата: остаток уменьшен на 4, строка осталась историей «Удалено (сторно)», _RET STORNO, итоги выдачи 0", ok,
              f"{d}; {row}; {s.retrec(1)}; {s.issrec(1)}")
        d2 = del_ret(s, 1)
        fx = fix_ret(s, 1, q="1")
        cl = s.click_ret("BtnRetClear", 1)
        R.add(c, "повторное «Удалить» — «уже удалён»; «Исправить» и «Очистить» удалённого — отказ", d2.startswith("SKIP") and fx.startswith("ERR")
              and cl.startswith("ERR") and s.ret(1)[0] == 1.0, f"{d2}; {fx}; {cl}")
        r2 = ret(s, 2, "1", "10")
        R.add(c, "№ удалённого возврата не переиспользуется: новый возврат — № 2; по выдаче снова можно вернуть всё (10)",
              r2 == "OK:4" and s.ret(2)[0] == 2.0 and s.stock(18) == 27.0, r2)
        oracle(c, s, "сторно возврата")
    finally:
        s.close()


@case
def w12_storno_after_reissue():
    c = "W12"
    s = ready(Session(new_wms("w12")))
    try:
        issue(s, 1, "15", "6", who="ива")
        r1 = ret(s, 1, "1", "5")
        i2 = issue(s, 2, "15", "5", who="пет")
        d = del_ret(s, 1)
        fdown = fix_ret(s, 1, q="2")
        R.add(c, "остаток 0 → вернули 5 → эти 5 выданы Петрову: сторно возврата и уменьшение возврата — отказ (остаток стал бы отрицательным)",
              r1 == "OK:2" and i2 == "OK:3" and s.stock(15) == 0.0 and d.startswith("ERR") and "отрицательным" in d and fdown.startswith("ERR")
              and s.ret(1)[13] == "Проведено", f"{r1}; {i2}; {d}; {fdown}")
        up = fix_ret(s, 1, q="6")
        dl2 = s.click("BtnDelete", 2)
        d2 = del_ret(s, 1)
        R.add(c, "увеличение возврата до 6 — можно (остаток 1); после сторно выдачи Петрову остаток 6 — сторно возврата проведено, остаток 0",
              up == "OK:4" and dl2 is None and s.iss(2)[17] == "Удалено (сторно)" and d2 == "OK:6" and s.stock(15) == 0.0, f"{up}; {s.iss(2)[17]}; {d2}")
        oracle(c, s, "сторно после повторной выдачи")
    finally:
        s.close()


@case
def w13_issue_fix_with_returns():
    c = "W13"
    s = ready(Session(new_wms("w13")))
    try:
        issue(s, 1, "19", "10")
        ret(s, 1, "1", "7")
        f5 = s.I("IssueFixRow", 1, "19", "5", "ива", DI)
        f7 = s.I("IssueFixRow", 1, "19", "7", "ива", DI)
        b7 = s.stock(19)
        f12 = s.I("IssueFixRow", 1, "19", "12", "ива", DI)
        R.add(c, "выдали 10, вернули 7: исправить выдачу на 5 — отказ; на 7 — можно; на 12 — можно при наличии остатка",
              f5.startswith("ERR") and "не может быть меньше возвращённого" in f5 and f7 == "OK:3" and b7 == 34.0 and f12 == "OK:4" and s.stock(19) == 29.0
              and s.iss(1)[4] == 12.0 and s.issrec(1) == (1.0, 1.0, 7.0, 1.0), f"{f5}; {f7} ({b7}); {f12} ({s.stock(19)})")
        s.U("TestFixInput", "20", "12", "ива", DI)
        fe = s.click("BtnFix", 1)
        msg_e = s.U("TestUiLastMessage")
        fw = s.I("IssueFixRow", 1, "19", "12", "пет", DI)
        fd = s.I("IssueFixRow", 1, "19", "12", "ива", "21.09.2026")
        big = s.I("IssueFixRow", 1, "19", "100", "ива", DI)
        R.add(c, "при действующих возвратах ЕИ, получатель и дата выдачи не меняются (кнопка «Исправить» и прямой вызов — отказ); больше остатка — отказ",
              fe is None and "ЕИ выдачи менять нельзя" in msg_e and fw.startswith("ERR") and "получателя" in fw and fd.startswith("ERR")
              and "дату" in fd and big.startswith("ERR") and s.iss(1)[11] == ei(19), f"{msg_e[:100]}; {fw[:80]}; {fd[:80]}; {big[:80]}")
        del_ret(s, 1)
        fe2 = s.I("IssueFixRow", 1, "20", "12", "ива", DI)
        R.add(c, "после сторно возвратов ЕИ выдачи снова можно исправить", fe2.startswith("OK") and s.iss(1)[11] == ei(20), fe2)
        oracle(c, s, "ISSUE_FIX и возвраты")
    finally:
        s.close()


@case
def w14_issue_del_with_returns():
    c = "W14"
    s = ready(Session(new_wms("w14")))
    try:
        issue(s, 1, "20", "5")
        ret(s, 1, "1", "2")
        s.click("BtnDelete", 1)
        m = s.U("TestUiLastMessage")
        d = s.I("IssueDeleteRow", 1)
        R.add(c, "сторно выдачи при действующем возврате — отказ (кнопка и прямой вызов), обратный возврат автоматически не создаётся",
              "сначала удалите (сторно)" in m and d.startswith("ERR") and s.iss(1)[17] == "Проведено" and s.stock(20) == 38.0
              and s.sysv("NEXT_RET") == 2.0, f"{m[:120]}; {d[:60]}")
        del_ret(s, 1)
        d2 = s.I("IssueDeleteRow", 1)
        R.add(c, "после сторно возврата сторно выдачи проведено: остаток 41", d2.startswith("OK") and s.iss(1)[17] == "Удалено (сторно)" and s.stock(20) == 41.0, d2)
        oracle(c, s, "ISSUE_DEL и возвраты")
    finally:
        s.close()


@case
def w15_return_after_issue_del():
    c = "W15"
    s = ready(Session(new_wms("w15")))
    try:
        issue(s, 1, "11", "3")
        s.I("IssueDeleteRow", 1)
        a = ret(s, 1, "1", "1")
        why = N(s, 1)
        s.type_ret("C", 2, "11")
        s.RU("TestRetInput", "1", KEEP, KEEP, -1)
        f = s.click_ret("BtnRetFind", 2)
        R.add(c, "возврат по удалённой (сторно) выдаче — отказ; «Найти выдачу» её не предлагает", a.startswith("ERR") and "удалена (сторно)" in why
              and f.startswith("ERR") and "нет проведённых выдач" in f and s.sysv("NEXT_RET") == 1.0, f"{a}; {why[:80]}; {f[:80]}")
        oracle(c, s, "возврат после ISSUE_DEL")
    finally:
        s.close()


# ================================================================ W16–W23 repeat, Ctrl+Z, inserted rows, copies, filter, date, quantity, reopen

@case
def w16_repeat_and_double_click():
    c = "W16"
    s = ready(Session(new_wms("w16")))
    try:
        issue(s, 1, "1", "10")
        r1 = ret(s, 1, "1", "2")
        before = (njournal(s), s.sysv("NEXT_RET"), s.stock(1), s.issrec(1))
        r2 = post_ret(s, 1)
        r3 = s.Rt("ReturnPostRow", 1)
        R.add(c, "повторный клик «Провести»: «уже проведено», второй возврат не создан", r1 == "OK:2" and r2.startswith("SKIP:уже проведено")
              and r3.startswith("SKIP") and before == (njournal(s), s.sysv("NEXT_RET"), s.stock(1), s.issrec(1)), f"{r1}; {r2}; {before}")
        ret_row(s, 2, "1", "3")
        s.goto("Возврат", "$B$3")
        res = []

        def click():
            s.RU("BtnRetPost")
            res.append(s.U("TestUiLastMessage"))
        th = [threading.Thread(target=click) for _ in range(2)]
        for t in th:
            t.start()
        for t in th:
            t.join(60)
        rets = [e for e in s.journal()[0] if e["type"] == "RETURN"]
        R.add(c, "двойной клик (два нажатия одновременно): один возврат, одна запись журнала", len(rets) == 2 and s.stock(1) == 95.0
              and s.sysv("NEXT_RET") == 3.0, f"ответы {res}; записей RETURN {len(rets)}")
        ret_row(s, 3, "1", "1")
        ans = []
        th = [threading.Thread(target=lambda: ans.append(s.Rt("ReturnPostRow", 3))) for _ in range(2)]
        for t in th:
            t.start()
        for t in th:
            t.join(60)
        R.add(c, "два одновременных вызова проведения одной строки: «OK» и «уже проведено»", sorted(a[:4] for a in ans) == ["OK:4", "SKIP"]
              and s.stock(1) == 96.0, str(ans))
        oracle(c, s, "повтор и двойной клик")
    finally:
        s.close()


@case
def w17_ctrl_z():
    c = "W17"
    s = ready(Session(new_wms("w17")))
    try:
        issue(s, 1, "12", "10")
        ret(s, 1, "1", "4")
        z1 = ctrl_z_harmless(s)
        fix_ret(s, 1, q="5")
        z2 = ctrl_z_harmless(s)
        del_ret(s, 1)
        z3 = ctrl_z_harmless(s)
        s.type_ret("O", 1, "комментарий")
        typed = s.ret(1)[14]
        s.ui(".uno:Undo")
        R.add(c, "Ctrl+Z ×5 после возврата, исправления и сторно ничего не откатывает; обычный ввод комментария O отменяется",
              z1 and z2 and z3 and typed == "комментарий" and s.ret(1)[14] == "" and s.state()["UNDO_LOCKED"] == "False", f"{z1} {z2} {z3}; {typed!r}")
        oracle(c, s, "Ctrl+Z")
    finally:
        s.close()


@case
def w18_inserted_rows():
    c = "W18"
    s = ready(Session(new_wms("w18")))
    try:
        issue(s, 1, "7", "10")
        issue(s, 2, "10", "5")
        ret(s, 1, "1", "2")
        ret(s, 2, "2", "1")
        s.goto("Возврат", "$A$2")
        s.ui(".uno:SelectRow")
        s.ui(".uno:InsertRowsBefore")
        s.ui(".uno:InsertRowsBefore")
        moved = (s.ret(3)[0], s.ret(4)[0])
        k1 = s.Rt("ReturnRowKind", 3)
        fx = fix_ret(s, 3, q="3")
        hint = s.retrec(1)[5]
        dl = del_ret(s, 4)
        R.add(c, "вставка строк над возвратами: строки сдвинулись, возврат найден по № (подсказка устарела), исправление и сторно из сдвинутых строк "
                 "проведены, подсказка строки исправлена той же операцией",
              moved == (1.0, 2.0) and k1 == "POSTED" and fx.startswith("OK") and hint == 3.0 and dl.startswith("OK") and s.retrec(2)[5] == 4.0
              and s.stock(7) == 43.0 and s.stock(10) == 16.0, f"{moved}; {k1}; {fx}; подсказка {hint}; {dl}")
        s.goto("Выдачи", "$A$2")
        s.ui(".uno:SelectRow")
        s.ui(".uno:InsertRowsBefore")
        r3 = ret(s, 5, "1", "1")
        R.add(c, "вставка строк над выдачей: возврат по № выдачи находит её в новой строке, подсказка _ISS исправлена",
              r3.startswith("OK") and s.iss(2)[0] == 1.0 and s.issrec(1)[1] == 2.0 and s.stock(7) == 44.0, f"{r3}; {s.issrec(1)}")
        locked = s.ret_locks(1)
        cl = s.click_ret("BtnRetClear", 1)
        r4 = ret(s, 1, "2", "1")
        R.add(c, "вставленная строка между проведёнными закрыта как соседняя; «Очистить» открывает её; в ней проводится новый возврат",
              cl.startswith("OK") and r4.startswith("OK") and s.ret(1)[0] == 4.0, f"защита {locked}; {cl}; {r4}")
        oracle(c, s, "вставка строк")
    finally:
        s.close()


@case
def w19_copies():
    c = "W19"
    p = new_wms("w19")
    s = ready(Session(p))
    issue(s, 1, "1", "10")
    ret(s, 1, "1", "2")
    s.goto("Возврат", "$A$2:$O$2")
    s.ui(".uno:Copy")
    s.goto("Возврат", "$A$8")
    s.ui(".uno:Paste")
    R.add(c, "копировать → вставить проведённую строку на защищённом листе: вставка отклонена", s.ret(7)[0] == "" and s.ret(7)[5] == "", str(s.ret(7)[:6]))
    sh = s.doc.Sheets.getByName("Возврат")
    sh.unprotect(PWD)
    s.goto("Возврат", "$A$2:$O$2")
    s.ui(".uno:Copy")
    s.goto("Возврат", "$A$9")
    s.ui(".uno:Paste")
    sh.protect(PWD)
    live = N(s, 8)
    res = post_ret(s, 8)
    fx = s.Rt("ReturnFixRow", 8, "1", DR, "")
    dl = s.Rt("ReturnDeleteRow", 8)
    R.add(c, "защита снята паролем, проведённая строка вставлена: сразу КОПИЯ, «Провести», «Исправить», «Удалить» — отказ",
          live.startswith("КОПИЯ") and "повторяет строку 2" in live and res.startswith("SKIP") and fx.startswith("ERR") and dl.startswith("ERR")
          and s.stock(1) == 92.0, f"{live[:90]}; {res}; {fx}; {dl}")
    s.close(save=True)
    s = Session(p, macros=0)
    sh = s.doc.Sheets.getByName("Возврат")
    sh.unprotect(PWD)
    row = list(sh.getCellRangeByPosition(0, 1, 14, 1).getDataArray()[0])
    sh.getCellRangeByPosition(0, 4, 14, 4).setDataArray((tuple(row),))
    row[0] = 99.0
    sh.getCellRangeByPosition(0, 5, 14, 5).setDataArray((tuple(row),))
    row[0] = 5.0
    sh.getCellRangeByPosition(0, 6, 14, 6).setDataArray((tuple(row),))
    sh.protect(PWD)
    sy = s.doc.Sheets.getByName("_SYS")
    sy.unprotect(PWD)
    sy.getCellByPosition(1, 7).setValue(10)                  # NEXT_RET 10: № 5 is below it, but WMS never registered it
    sy.protect(PWD)
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        rep = s.report()
        n_ = [N(s, r) for r in (4, 5, 6)]
        pr = [post_ret(s, r) for r in (4, 5, 6)]
        cl = [s.click_ret("BtnRetClear", r) for r in (4, 5, 6, 8)]
        orig = s.ret(1)
        R.add(c, "сохранено без макросов: копия, № который WMS не выдавала, и № без регистрации в _RET помечены КОПИЯ при запуске; не проводятся; "
                 "«Очистить» убирает; оригинал цел",
              "без WMS" in rep and n_[0].startswith("КОПИЯ") and "повторяет" in n_[0] and n_[1].startswith("КОПИЯ") and "не выдавался" in n_[1]
              and n_[2].startswith("КОПИЯ") and "не зарегистрирован" in n_[2] and all(x.startswith("SKIP") for x in pr)
              and all(x.startswith("OK") for x in cl) and all(s.ret(r)[0] == "" for r in (4, 5, 6, 8)) and s.ret_locks(4) == OPEN
              and orig[0] == 1.0 and orig[13] == "Проведено", f"{n_}; {pr}; {cl}; оригинал {orig[:6]}")
        cl0 = s.click_ret("BtnRetClear", 1)
        r2 = ret(s, 2, "1", "1")
        R.add(c, "«Очистить» проведённой строки — отказ; следующий возврат получает № 10 (NEXT_RET)", cl0.startswith("ERR") and s.ret(1)[0] == 1.0
              and r2.startswith("OK") and s.ret(2)[0] == 10.0, f"{cl0}; {r2}")
        oracle(c, s, "копии не считаются возвратами")
    finally:
        s.close()


def set_filter(s, col, value):
    dbr = s.doc.DatabaseRanges.getByName("WMS_RETURNS")
    fd = dbr.getFilterDescriptor()
    f = TableFilterField()
    f.Field, f.Operator, f.IsNumeric, f.StringValue = col, EQUAL, False, value
    fd.setFilterFields((f,))
    dbr.refresh()


def visible(s, r):
    return s.doc.Sheets.getByName("Возврат").getRows().getByIndex(r).IsVisible


@case
def w20_filter():
    c = "W20"
    p = new_wms("w20")
    s = ready(Session(p))
    try:
        issue(s, 1, "7", "10", who="ива")
        issue(s, 2, "10", "10", who="пет")
        for r, (k, q) in enumerate([("1", "1"), ("2", "1"), ("1", "2"), ("2", "2"), ("1", "3")], start=1):
            ret_row(s, r, k, q)
        set_filter(s, 8, IVA)
        vis = [visible(s, r) for r in range(1, 6)]
        f1 = post_ret(s, 1)
        s.goto("Возврат", "$A$2:$O$6")
        s.RU("BtnRetPost")
        blk = s.U("TestUiLastMessage")
        ok = (vis == [True, False, True, False, True] and f1 == "OK:3" and blk.startswith("OK=2;ERR=0;SKIP=1") and s.ret(2)[0] == "" and s.ret(4)[0] == ""
              and s.ret(3)[0] == 2.0 and s.ret(5)[0] == 3.0)
        R.add(c, "при активном автофильтре (получатель = Иванов): проведение строки и выделенного блока — только видимые строки, скрытые не тронуты",
              ok, f"видимость {vis}; {f1}; блок {blk.splitlines()[0]}")
        s.doc.store()
        vis1 = [visible(s, r) for r in range(1, 6)]
        with zipfile.ZipFile(p) as z:
            nfilt = z.read("content.xml").decode("utf-8").count('table:visibility="filter"')
        s.close()
        s = ready(Session(p))
        vis2 = [visible(s, r) for r in range(1, 6)]
        R.add(c, "сохранение при фильтре на «Возврат» (D-041): строки сохранены показанными, фильтр восстановлен после сохранения и при открытии",
              nfilt == 0 and vis1 == vis and vis2 == vis and not s.doc.isModified() and st(s)[0] == "CLEAN", f"{vis1} {vis2}; скрытых в файле {nfilt}")
        oracle(c, s, "фильтр")
    finally:
        s.close()


@case
def w21_dates():
    c = "W21"
    s = ready(Session(new_wms("w21"), locale="en-US"))
    try:
        issue(s, 1, "12", "10")
        ret_row(s, 1, "1", "1", date=None)
        no_h = post_ret(s, 1)
        out = [no_h]
        for d in ("31.02.2026", "19.09.2026", "25.09.1999", "22-09-2026x"):
            s.return_input(1, date=d)
            out.append(post_ret(s, 1))
        s.return_input(1, date=DR)
        ok = post_ret(s, 1)
        R.add(c, "дата возврата обязательна: без H — отказ; 31.02, раньше даты выдачи, вне 2000–2099, не дд.мм.гггг — отказ; текст дд.мм.гггг в локали "
                 "en-US принят и стал датой",
              all(x.startswith("ERR") for x in out) and "не указана дата возврата" in out[0] and "раньше даты выдачи" in out[2] and ok == "OK:2"
              and s.ret(1)[7] == serial(DR), f"{[x[:70] for x in out]}; {ok}; H {s.ret(1)[7]!r}")
        s.type_ret("B", 2, "1")
        s.RU("TestRetInput", "1", KEEP, KEEP, -1)
        s.click_ret("BtnRetFind", 2)
        shown, back, _ = dialog(s)
        h_today = s.ret(2)[7]
        s.type_ret("B", 3, "1")
        s.RU("TestRetInput", "1", "23.09.2026", KEEP, -1)
        s.click_ret("BtnRetFind", 3)
        R.add(c, "окно «Найти выдачу» предлагает сегодняшнюю дату (видна в строке до проведения); пользователь может её изменить",
              shown[1] == TODAY.strftime("%d.%m.%Y") and h_today == serial(TODAY.strftime("%d.%m.%Y")) and s.ret(3)[7] == serial("23.09.2026")
              and s.ret(2)[0] == "" and s.ret(3)[0] == "", f"{shown}; H2 {h_today}; H3 {s.ret(3)[7]}")
        oracle(c, s, "даты")
    finally:
        s.close()


def qty_case(c, name, locale):
    s = ready(Session(new_wms(name), locale=locale))
    out = []
    try:
        issue(s, 1, "7", "50")
        vals = ["1,5", "1.5", "0,125", "2,1250", "-1", "1e3", "1 000", "1,500", "01.02.2026", "0", "abc", "1,2345"]
        for i, v in enumerate(vals, start=1):
            ret_row(s, i, "1", v)
            res = post_ret(s, i)
            out.append((v, res[:2], s.ret(i)[5]))
        want_ok = {"1,5": 1.5, "1.5": 1.5, "0,125": 0.125, "2,1250": 2.125}
        ok = all((x[1] == "OK") == (x[0] in want_ok) and (x[0] not in want_ok or x[2] == want_ok[x[0]]) for x in out)
        return s, ok, out
    except Exception:
        s.close()
        raise


@case
def w22_quantity_locale():
    c = "W22"
    for loc in ("ru-RU", "en-US"):
        s, ok, out = qty_case(c, f"w22_{loc}", loc)
        try:
            R.add(c, f"локаль {loc}: F «1,5», «1.5», «0,125», «2,1250» — числа; «-1», «1e3», «1 000», «1,500», дата, 0, текст, 4 знака — отказ",
                  ok and s.stock(7) == 5.25, str(out))
            oracle(c, s, f"количество, локаль {loc}")
        finally:
            s.close()


@case
def w23_save_reopen():
    c = "W23"
    p = new_wms("w23")
    s = ready(Session(p))
    issue(s, 1, "7", "10")
    issue(s, 2, "10", "5")
    ret(s, 1, "1", "2")
    ret(s, 2, "1", "3")
    fix_ret(s, 2, q="4")
    ret(s, 3, "2", "1")
    del_ret(s, 3)
    ret_row(s, 4, "2", "2")
    s.close(save=True)
    with zipfile.ZipFile(p) as z:
        xml = z.read("content.xml").decode("utf-8")
    s = ready(Session(p))
    try:
        checks = {"CLEAN": st(s)[0] == "CLEAN", "не «изменена» после открытия": not s.doc.isModified(),
                  "N": [N(s, r)[:25] for r in (1, 2, 3)] == ["Проведено", "Проведено (исправлено)", "Удалено (сторно)"],
                  "защита": [s.ret_locks(r) for r in (1, 2, 3, 4)] == [POSTED, POSTED, POSTED, OPEN],
                  # «Выдачи», «Заказы», «Возврат» and, since Core Phase 5, «Иной приход» allow inserting rows
                  "вставка строк": xml.count('loext:insert-rows="true"') == 4, "автофильтр": s.doc.DatabaseRanges.hasByName("WMS_RETURNS"),
                  "NEXT_RET": s.sysv("NEXT_RET") == 4.0, "остатки": (s.stock(7), s.stock(10)) == (46.0, 16.0),
                  "непроведённые на «Главной»": "возвраты: 1 строк" in s.main_status()[3]}
        nxt = post_ret(s, 4)
        checks["следующий возврат"] = nxt.startswith("OK") and s.ret(4)[0] == 4.0
        bad = [k for k, v in checks.items() if not v]
        R.add(c, "сохранить → закрыть → открыть: статусы, защита, опция «вставка строк», автофильтр, NEXT_RET, остатки целы; книга не «изменена»; "
                 "нумерация продолжается", not bad, f"не выполнено: {bad}; {s.main_status()[3]}")
        oracle(c, s, "после повторного открытия")
    finally:
        s.close()


# ================================================================ W24–W31 crash / recovery

def crash_return(name, sheet=None, point=None):
    """a return interrupted by a simulated crash (the book stored mid-operation, the process killed), then restarted"""
    p = new_wms(name)
    s = ready(Session(p))
    order(s, 1, h="100", f="40")
    s.click_ord("BtnRcvPost", 1)
    issue(s, 1, "201", "10")
    ret_row(s, 1, "1", "4")
    s.close(save=True)
    s = ready(Session(p))
    if sheet:
        s.B("TestSetFaultSheet", sheet, 2)
    else:
        s.B("TestSetFault", point, 2)
    crash = s.Rt("ReturnPostRow", 1)
    s.kill()
    s = ready(Session(p))
    unl = s.B("ActionForceUnlock")
    return p, s, crash, unl


def rolled_back(s):
    row = s.ret(1)
    return (row[0] == "" and row[5] == "4" and row[13] != "Проведено" and s.ret_locks(1) == OPEN and s.sysv("NEXT_RET") == 1.0 and s.stock(201) == 30.0
            and s.ord(1, "X") == 30.0 and s.retrec(1)[0] == "" and s.issrec(1)[0] == "" and njournal(s) == 2)


def again_once(c, s):
    again = s.Rt("ReturnPostRow", 1)
    R.add(c, "повторное проведение: возврат № 1 ровно один раз, остаток 34, «Заказы».X 34", again == "OK:3" and s.ret(1)[0] == 1.0 and s.stock(201) == 34.0
          and s.ord(1, "X") == 34.0 and s.sysv("NEXT_RET") == 2.0, again)


@case
def w24_crash_after_number_reserved():
    c = "W24"
    p, s, crash, unl = crash_return("w24", sheet="_SYS")
    try:
        R.add(c, "сбой сразу после резервирования № возврата (NEXT_RET записан, книга сохранена, kill): при запуске откат — NEXT_RET 1, возврата нигде нет, "
                 "ввод на месте", crash == "CRASH-SIM" and "отменена по снимку" in unl and st(s)[0] == "CLEAN" and rolled_back(s),
              f"{crash}; {unl[:150]}; {s.ret(1)[:6]}; NEXT_RET {s.sysv('NEXT_RET')}")
        again_once(c, s)
        oracle(c, s, "сбой после резервирования №")
    finally:
        s.close()


@case
def w25_crash_after_stock():
    c = "W25"
    p, s, crash, unl = crash_return("w25", sheet="Наличие")
    try:
        R.add(c, "сбой после изменения остатка «Наличие»: откат — остаток 30, № не израсходован", crash == "CRASH-SIM" and "отменена по снимку" in unl
              and rolled_back(s), f"{crash}; {unl[:120]}; остаток {s.stock(201)}")
        again_once(c, s)
        oracle(c, s, "сбой после «Наличие»")
    finally:
        s.close()


@case
def w26_crash_after_service_record():
    c = "W26"
    for sheet in ("_RET", "_ISS"):
        p, s, crash, unl = crash_return(f"w26{sheet}", sheet=sheet)
        try:
            R.add(c, f"сбой после записи служебной таблицы {sheet}: откат — записи возврата и итогов выдачи нет", crash == "CRASH-SIM"
                  and "отменена по снимку" in unl and rolled_back(s), f"{crash}; {unl[:120]}; _RET {s.retrec(1)}; _ISS {s.issrec(1)}")
            again_once(c, s)
            oracle(c, s, f"сбой после {sheet}")
        finally:
            s.close()


@case
def w27_crash_after_user_row():
    c = "W27"
    p, s, crash, unl = crash_return("w27", sheet="Возврат")
    try:
        R.add(c, "сбой после записи строки «Возврат» (все записи сделаны, журнала нет): откат — строка снова открыта, ввод на месте", crash == "CRASH-SIM"
              and "отменена по снимку" in unl and rolled_back(s), f"{crash}; {unl[:120]}; {s.ret(1)}; защита {s.ret_locks(1)}")
        again_once(c, s)
        oracle(c, s, "сбой после строки «Возврат»")
    finally:
        s.close()


@case
def w28_crash_after_journal():
    c = "W28"
    p, s, crash, unl = crash_return("w28", point=3)
    try:
        _, block = st(s)
        rec = s.B("ActionRecover")
        row = s.ret(1)
        ok = (crash == "CRASH-SIM" and block == "TAIL" and "восстановлено операций: 1" in rec and st(s)[0] == "CLEAN" and row[0] == 1.0 and row[5] == 4.0
              and row[13] == "Проведено" and s.stock(201) == 34.0 and s.ord(1, "X") == 34.0 and s.sysv("NEXT_RET") == 2.0 and s.ret_locks(1) == POSTED
              and s.issrec(1) == (1.0, 1.0, 4.0, 1.0))
        R.add(c, "сбой после записи в журнал: при запуске «хвост», «Восстановить» доводит возврат ровно один раз (остаток, X, _RET, _ISS, строка)",
              ok, f"{crash}; {block}; {rec[:120]}; {row}")
        again = s.Rt("ReturnPostRow", 1)
        R.add(c, "после восстановления повторное проведение — «уже проведено»", again.startswith("SKIP"), again)
        oracle(c, s, "сбой после журнала")
    finally:
        s.close()


def crash_two_step(c, name, action, args, check_old, check_new, what):
    p = new_wms(name)
    s = ready(Session(p))
    issue(s, 1, "12", "10")
    issue(s, 2, "12", "5", who="пет")
    ret(s, 1, "1", "4")
    s.close(save=True)
    s = ready(Session(p))
    s.B("TestSetFault", 2, 2)
    crash = s.Rt(action, *args)
    s.kill()
    s = ready(Session(p))
    s.B("ActionForceUnlock")
    ok1 = crash == "CRASH-SIM" and check_old(s)
    R.add(c, f"сбой посреди {what} (до журнала): возврат в прежнем виде, промежуточного состояния нет", ok1, f"{crash}; {s.ret(1)}; {s.stock(12)}")
    s.B("TestSetFault", 3, 2)
    crash2 = s.Rt(action, *args)
    s.kill()
    s = ready(Session(p))
    try:
        s.B("ActionForceUnlock")
        rec = s.B("ActionRecover")
        ok2 = crash2 == "CRASH-SIM" and "восстановлено операций: 1" in rec and check_new(s)
        R.add(c, f"сбой после журнала при {what}: «Восстановить» доводит операцию целиком ровно один раз", ok2, f"{crash2}; {rec[:100]}; {s.ret(1)}")
        oracle(c, s, f"сбой при {what}")
    finally:
        s.close()


@case
def w29_crash_during_fix():
    crash_two_step("W29", "w29", "ReturnFixRow", (1, "7", DR, "Ф-1"),
                   lambda s: s.ret(1)[5] == 4.0 and s.stock(12) == 24.0 and s.issrec(1)[2] == 4.0 and s.stock_row(12)[5] == "S-12",
                   lambda s: s.ret(1)[5] == 7.0 and s.ret(1)[13] == "Проведено (исправлено)" and s.stock(12) == 27.0 and s.issrec(1)[2] == 7.0
                   and s.retrec(1)[3] == 7.0 and s.stock_row(12)[5] == "Ф-1", "исправлении возврата (RETURN_FIX)")


@case
def w30_crash_during_storno():
    crash_two_step("W30", "w30", "ReturnDeleteRow", (1,),
                   lambda s: s.ret(1)[13] == "Проведено" and s.stock(12) == 24.0 and s.retrec(1)[4] == "LIVE" and s.issrec(1)[3] == 1.0,
                   lambda s: s.ret(1)[13] == "Удалено (сторно)" and s.stock(12) == 20.0 and s.retrec(1)[4] == "STORNO" and s.issrec(1) == (1.0, 1.0, 0.0, 0.0),
                   "сторно возврата (RETURN_DEL)")


@case
def w31_abandon_keeps_number():
    c = "W31"
    p, s, crash, unl = crash_return("w31", point=3)
    try:
        ab = s.B("ActionAbandonTail")
        nx = s.sysv("NEXT_RET")
        abandon = [e for e in s.journal()[0] if e["type"] == "ABANDON"]
        again = s.Rt("ReturnPostRow", 1)
        R.add(c, "«Отложить хвост» с возвратом № 1: этот № больше не выдаётся (NEXT_RET поднят до 2, записан в ABANDON), новый возврат получает № 2",
              "отложено операций: 1" in ab and nx == 2.0 and abandon and abandon[0]["fields"].get("NEXT_RET") == "2" and again == "OK:5"
              and s.ret(1)[0] == 2.0 and s.stock(201) == 34.0, f"{ab[:120]}; NEXT_RET {nx}; {again}")
        oracle(c, s, "отложенный хвост")
    finally:
        s.close()


# ================================================================ W32 mixed sequence

@case
def w32_mixed_sequence():
    c = "W32"
    rnd = random.Random(20260926)
    s = ready(Session(new_wms("w32")))
    try:
        model = dict(bal={}, issues={}, returns={})
        for i in range(4):
            order(s, 1 + i, a=f"З-{i}", b=f"Товар заказа {i}", h="30", f=str(20 + i * 5))
            assert s.click_ord("BtnRcvPost", 1 + i).startswith("OK")
            model["bal"][ei(201 + i)] = 20.0 + i * 5
        eis = [ei(201 + i) for i in range(4)] + [ei(7), ei(10), ei(12)]
        for e in eis[4:]:
            model["bal"][e] = INITIAL_STOCK[e]
        who = ["ива", "пет", "егр"]
        full = {"ива": IVA, "пет": PET, "егр": EGR}
        irow, rrow, n_ops, errs, kinds = 1, 1, 4, [], {}
        ret_rows = {}
        for step in range(70):
            live_iss = [k for k, v in model["issues"].items() if not v["del"]]
            live_ret = [n for n, v in model["returns"].items() if v["live"]]
            roll = rnd.random()
            if roll < 0.35 or not live_iss:
                e = rnd.choice([x for x in eis if model["bal"][x] >= 1])
                q = rnd.randint(1, min(6, int(model["bal"][e])))
                w = rnd.choice(who)
                res = issue(s, irow, e[3:].lstrip("0"), str(q), who=w)
                if res.startswith("OK"):
                    k = len(model["issues"]) + 1
                    model["issues"][k] = {"ei": e, "qty": float(q), "ret": 0.0, "row": irow, "who": full[w], "del": False}
                    model["bal"][e] -= q
                irow += 1
                kind = "ISSUE"
            elif roll < 0.70:
                k = rnd.choice(live_iss)
                iss = model["issues"][k]
                avail = iss["qty"] - iss["ret"]
                if avail < 0.5:
                    q = 1
                else:
                    q = rnd.choice([1, avail, max(1, int(avail // 2))])
                res = ret(s, rrow, str(k), str(int(q)) if q == int(q) else str(q).replace(".", ","))
                if res.startswith("OK"):
                    n = len(model["returns"]) + 1
                    model["returns"][n] = dict(issue=k, qty=float(q), live=True)
                    ret_rows[n] = rrow
                    iss["ret"] += q
                    model["bal"][iss["ei"]] += q
                elif avail >= q - 1e-9:
                    errs.append(f"шаг {step}: возврат {q} по №{k} (можно {avail}) отказан: {res}")
                rrow += 1
                kind = "RETURN"
            elif roll < 0.80 and live_ret:
                n = rnd.choice(live_ret)
                rv = model["returns"][n]
                iss = model["issues"][rv["issue"]]
                q = rnd.randint(1, int(iss["qty"] - iss["ret"] + rv["qty"]))
                res = fix_ret(s, ret_rows[n], q=str(q))
                okq = model["bal"][iss["ei"]] - rv["qty"] + q >= 0
                if res.startswith("OK"):
                    iss["ret"] += q - rv["qty"]
                    model["bal"][iss["ei"]] += q - rv["qty"]
                    rv["qty"] = float(q)
                elif okq and not res.startswith("SKIP"):
                    errs.append(f"шаг {step}: исправление №{n} до {q} отказано: {res}")
                kind = "RETURN_FIX"
            elif roll < 0.90 and live_ret:
                n = rnd.choice(live_ret)
                rv = model["returns"][n]
                iss = model["issues"][rv["issue"]]
                res = del_ret(s, ret_rows[n])
                if res.startswith("OK"):
                    rv["live"] = False
                    iss["ret"] -= rv["qty"]
                    model["bal"][iss["ei"]] -= rv["qty"]
                elif model["bal"][iss["ei"]] >= rv["qty"]:
                    errs.append(f"шаг {step}: сторно №{n} отказано: {res}")
                kind = "RETURN_DEL"
            else:
                k = rnd.choice(live_iss)
                iss = model["issues"][k]
                if iss["ret"] > 0 or rnd.random() < 0.5:
                    q = max(1, int(iss["ret"]) + rnd.randint(0, 2))
                    res = s.I("IssueFixRow", iss["row"], iss["ei"], str(q), iss["who"], DI)
                    if res.startswith("OK"):
                        model["bal"][iss["ei"]] += iss["qty"] - q
                        iss["qty"] = float(q)
                    elif not res.startswith("SKIP") and q >= iss["ret"] and model["bal"][iss["ei"]] + iss["qty"] - q >= 0:
                        errs.append(f"шаг {step}: ISSUE_FIX №{k} до {q} отказано: {res}")
                    kind = "ISSUE_FIX"
                else:
                    res = s.I("IssueDeleteRow", iss["row"])
                    if res.startswith("OK"):
                        iss["del"] = True
                        model["bal"][iss["ei"]] += iss["qty"]
                    else:
                        errs.append(f"шаг {step}: ISSUE_DEL №{k} отказано: {res}")
                    kind = "ISSUE_DEL"
            kinds[kind] = kinds.get(kind, 0) + (1 if res.startswith("OK") else 0)
            n_ops += 1 if res.startswith("OK") else 0
        bal_ok = all(abs(s.stock(int(e[3:])) - b) < 1e-6 for e, b in model["bal"].items())
        sc = s.B("ActionSelfCheck")
        R.add(c, f"{n_ops} проведённых операций (приходы, выдачи, возвраты, исправления, сторно) подряд: модель теста и книга совпадают, "
                 "неожиданных отказов нет, самопроверка без ошибок",
              n_ops >= 50 and not errs and bal_ok and sc.splitlines()[0].startswith("САМОПРОВЕРКА WMS: ошибок 0"),
              f"операций {n_ops}: {kinds}; отказы {errs[:3]}; {sc.splitlines()[0]}")
        oracle(c, s, "смешанная последовательность")
    finally:
        s.close()


# ================================================================ W33–W42 unit, buttons, protection, panel, place, history, block, self-check, M/N, check

@case
def w33_unit_mismatch():
    c = "W33"
    p = new_wms("w33")
    s = ready(Session(p))
    issue(s, 1, "9", "4")
    s.close(save=True)
    s = Session(p, macros=0)
    sh = s.doc.Sheets.getByName("Наличие")
    sh.unprotect(PWD)
    sh.getCellByPosition(3, 9).setString("шт")
    sh.protect(PWD)
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        res = ret(s, 1, "1", "1")
        R.add(c, "единица ЕИ в «Наличие» (шт) не совпадает с единицей выдачи (упак): возврат отказан, пересчёт единиц не выполняется",
              res.startswith("ERR") and "пересчёт единиц не выполняется" in N(s, 1) and s.stock(9) == 10.0, f"{res}; {N(s, 1)[:120]}")
    finally:
        s.close()


@case
def w34_buttons_bound():
    c = "W34"
    s = Session(new_wms("w34"))
    try:
        form = s.doc.Sheets.getByName("Возврат").getDrawPage().getForms().getByIndex(0)
        names = {}
        for i in range(form.getCount()):
            ev = form.getScriptEvents(i)
            code = ev[0].ScriptCode if ev else ""
            names[form.getByIndex(i).Label] = code.split("Standard.", 1)[1].split("?")[0] if "Standard." in code else code
        want = {"Найти выдачу": "WmsReturnUi.BtnRetFind", "Проверить": "WmsReturnUi.BtnRetCheck", "Провести": "WmsReturnUi.BtnRetPost",
                "Исправить": "WmsReturnUi.BtnRetFix", "Удалить": "WmsReturnUi.BtnRetDelete", "Очистить": "WmsReturnUi.BtnRetClear"}
        R.add(c, "6 кнопок листа «Возврат» привязаны к макросам (запуск через меню не нужен)", names == want, str(names))
    finally:
        s.close()


@case
def w35_protection():
    c = "W35"
    s = ready(Session(new_wms("w35")))
    try:
        issue(s, 1, "7", "10")
        ret(s, 1, "1", "2")
        before = s.ret(1)
        changed = []
        for col in "ABCDEFGHIJKLMN":
            s.type_ret(col, 1, "999")
            if s.ret(1) != before:
                changed.append(col)
        s.type_ret("O", 1, "примечание")
        open_ok = s.ret(1)[14] == "примечание"
        s.type_ret("B", 2, "1")
        before2 = s.ret(2)
        open_locked = []
        for col in "ADEGKLMN":
            s.type_ret(col, 2, "777")
            if s.ret(2) != before2:
                open_locked.append(col)
        s.goto("Возврат", "$A$2")
        s.ui(".uno:SelectRow")
        s.ui(".uno:DeleteRows")
        s.goto("Возврат", "$A$1:$O$3")
        s.ui(".uno:SortDescending")
        R.add(c, "защита: в проведённой строке закрыты A–N, открыт комментарий O; в непроведённой закрыты A D E G K L M N; удаление строк и сортировка "
                 "отклонены", not changed and open_ok and not open_locked and s.ret(1)[0] == 1.0 and s.ret_locks(1) == POSTED,
              f"изменились закрытые {changed}; O {open_ok}; в непроведённой {open_locked}; A2 {s.ret(1)[0]}")
        chk = s.click_ret("BtnRetCheck", 1)
        R.add(c, "«Проверить» проведённой строки: возврат сверен с выдачей и остатком — сообщение", chk.startswith("OK:возврат № 1 (действует)")
              and "можно вернуть 8" in chk and "остаток" in chk, chk)
        oracle(c, s, "защита")
    finally:
        s.close()


@case
def w36_main_panel_unposted():
    c = "W36"
    p = new_wms("w36")
    s = ready(Session(p))
    issue(s, 1, "7", "10")
    ret_row(s, 1, "1", "2")
    ret_row(s, 2, "1", "1")
    s.close(save=True)
    s = ready(Session(p))
    try:
        un = s.main_status()[3]
        R.add(c, "«Главная» после открытия: непроведённые строки «Возврат» (количество без №) показаны", "возвраты: 2 строк" in un, un)
    finally:
        s.close()


@case
def w37_place_change():
    c = "W37"
    s = ready(Session(new_wms("w37")))
    try:
        issue(s, 1, "8", "5")
        ret_row(s, 1, "1", "2", place="Б-7")
        note = N(s, 1)
        a = post_ret(s, 1)
        e = [x for x in s.journal()[0] if x["type"] == "RETURN"][0]
        R.add(c, "место при возврате изменено: до проведения «Контроль» предупреждает «место ЕИ изменится: S-08 → Б-7»; «Наличие».F = Б-7 в той же "
                 "операции (одна строка журнала, поле PLACE), отдельного перемещения нет",
              "место ЕИ изменится: «S-08» → «Б-7»" in note and a == "OK:2" and s.stock_row(8)[5] == "Б-7" and e["fields"].get("PLACE") == "Б-7" and njournal(s) == 2
              and any(w["sheet"] == "Наличие" and w["after"].endswith("Б-7") for w in e["writes"]),
              f"{a}; {s.stock_row(8)}; {e['fields'].get('PLACE')}; {[(w['sheet'], w['col'], w['after']) for w in e['writes'] if w['sheet'] == 'Наличие']}")
        b = ret(s, 2, "1", "1")
        R.add(c, "пустое J — предлагается текущее место ЕИ (Б-7), место не меняется", b == "OK:3" and s.ret(2)[9] == "Б-7" and s.stock_row(8)[5] == "Б-7", b)
        d = del_ret(s, 1)
        R.add(c, "сторно возврата не возвращает прежнее место (место — текущее место ЕИ)", d == "OK:4" and s.stock_row(8)[5] == "Б-7", d)
        oracle(c, s, "место хранения")
    finally:
        s.close()


@case
def w38_recipient_history():
    c = "W38"
    s = ready(Session(new_wms("w38")))
    try:
        issue(s, 1, "7", "10", who="ива")
        issue(s, 2, "7", "4", who="пет")
        issue(s, 3, "10", "3", who="ива")
        issue(s, 4, "7", "2", who="ива")
        ret(s, 1, "1", "3")
        ret(s, 2, "4", "2")
        ret(s, 3, "2", "1")
        ret(s, 4, "1", "1")
        del_ret(s, 4)
        s.I("IssueDeleteRow", 3)
        hist = oracle(c, s, "история получателя")
        want = {(IVA, ei(7)): dict(issued=12.0, returned=5.0, net=7.0), (PET, ei(7)): dict(issued=4.0, returned=1.0, net=3.0)}
        R.add(c, "оракул строит историю пары «получатель + ЕИ» по зарегистрированным движениям: Иванов/ЕИ-7 выдано 12, возвращено 5, чистое 7; "
                 "Петров/ЕИ-7 4, 1, 3; удалённые выдача и возврат не учитываются", hist == want, str(hist))
    finally:
        s.close()


@case
def w39_block_posting():
    c = "W39"
    s = ready(Session(new_wms("w39")))
    try:
        issue(s, 1, "7", "20")
        ret_row(s, 1, "1", "1")
        ret_row(s, 2, "1", None)
        ret_row(s, 3, "1", "abc")
        ret_row(s, 4, "1", "2", date=None)
        ret_row(s, 5, "1", "3")
        head, first, sysl = s.click_ret("BtnRetPost", 1, r1=5).split("\n", 2)
        msg = s.RU("TestRetBlockMessage")
        R.add(c, "блок из 5 строк (D-059): каждая строка с F — отдельная операция; ошибки строк не останавливают остальные; строка без F пропущена",
              head == "OK=2;ERR=2;SKIP=1;STOP=0" and sysl == "" and first.startswith("строка 4:") and [s.ret(r)[0] for r in range(1, 6)] == [1.0, "", "", "", 2.0]
              and "проведено 2, пропущено 1" in msg and "отклонено с ошибкой 2" in msg, f"{head}; {first}; {msg[:200]}")
        for r in (6, 7, 8):
            ret_row(s, r, "1", "1")
        sh = s.doc.Sheets.getByName("Возврат")
        sh.unprotect(PWD)
        sh.getCellByPosition(0, 12).setValue(4)          # the № the counter will give next but one is already on the sheet
        sh.protect(PWD)
        nj = njournal(s)
        head2, _, sys2 = s.click_ret("BtnRetPost", 6, r1=8).split("\n", 2)
        msg2 = s.RU("TestRetBlockMessage")
        R.add(c, "системная ошибка WMS посреди блока (номер, который выдал бы NEXT_RET, уже на листе): строка 7 проведена, на строке 8 обработка "
                 "остановлена, строка 9 не обрабатывалась",
              head2 == "OK=1;ERR=0;SKIP=0;STOP=8" and "ERR-SYS:" in sys2 and s.ret(6)[0] == 3.0 and s.ret(7)[0] == "" and s.ret(8)[0] == ""
              and njournal(s) == nj + 1 and "ОБРАБОТКА ОСТАНОВЛЕНА" in msg2, f"{head2}; {sys2[:160]}")
        cl = s.click_ret("BtnRetClear", 12)
        head3 = s.click_ret("BtnRetPost", 6, r1=8).split("\n")[0]
        R.add(c, "причина устранена (строка с чужим № очищена кнопкой «Очистить»): тот же блок — проведённая строка пропущена, две проведены",
              cl.startswith("OK") and head3 == "OK=2;ERR=0;SKIP=1;STOP=0" and [s.ret(r)[0] for r in (6, 7, 8)] == [3.0, 4.0, 5.0], f"{cl}; {head3}")
        oracle(c, s, "проведение блока")
    finally:
        s.close()


@case
def w40_selfcheck():
    c = "W40"
    p = new_wms("w40")
    s = ready(Session(p))
    issue(s, 1, "7", "10")
    ret(s, 1, "1", "2")
    ret(s, 2, "1", "3")
    del_ret(s, 2)
    sc = s.B("ActionSelfCheck")
    line = [x for x in sc.splitlines() if "возвраты" in x]
    R.add(c, "самопроверка: строка «возвраты» — возвратов 2 (сторно 1), расхождений нет", sc.splitlines()[0].startswith("САМОПРОВЕРКА WMS: ошибок 0")
          and line and line[0].startswith("OK возвраты: возвратов 2 (из них сторно 1)"), str(line))
    s.close(save=True)
    s = Session(p, macros=0)
    sh = s.doc.Sheets.getByName("_ISS")
    sh.unprotect(PWD)
    sh.getCellByPosition(2, 1).setValue(5)
    sh.protect(PWD)
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        sc2 = s.B("ActionSelfCheck")
        line2 = [x for x in sc2.splitlines() if "возвраты" in x]
        R.add(c, "итоги выдачи в _ISS испорчены вручную: самопроверка — FAIL «возвраты» с расхождением", line2 and line2[0].startswith("FAIL возвраты")
              and "выдача № 1" in line2[0], str(line2))
    finally:
        s.close()


@case
def w41_legacy_mn():
    c = "W41"
    s = ready(Session(new_wms("w41")))
    try:
        issue(s, 1, "7", "10")
        s.type_iss("M", 1, "9")
        s.type_iss("N", 1, "90%")
        mn = s.iss(1)[12:14]
        s.type_ret("B", 1, "1")
        pv = N(s, 1)
        r = ret(s, 1, "1", "10")
        R.add(c, "legacy «Вернул» (M) и «Вернул в %» (N) выдачи не используются для расчёта и не меняются возвратом",
              mn[0] != "" and "возвращено 0, можно вернуть 10" in pv and r == "OK:2" and s.iss(1)[12:14] == mn, f"{pv[:90]}; {r}; M/N {mn} → {s.iss(1)[12:14]}")
        oracle(c, s, "legacy M/N", legacy_mn_empty=False)
    finally:
        s.close()


@case
def w42_check_button():
    c = "W42"
    s = ready(Session(new_wms("w42")))
    try:
        issue(s, 1, "7", "10")
        ret_row(s, 1, "1", "3")
        a = s.click_ret("BtnRetCheck", 1)
        ret_row(s, 2, "1", "30")
        b = s.click_ret("BtnRetCheck", 2)
        e = s.click_ret("BtnRetCheck", 5)
        R.add(c, "«Проверить» непроведённой строки: «Можно провести» или ошибка в «Контроль»; пустая строка — подсказка; ничего не проводится",
              a.startswith("OK:можно провести") and N(s, 1).startswith("Можно провести") and b.startswith("ERR") and N(s, 2).startswith("Ошибка:")
              and e.startswith("ERR") and njournal(s) == 1, f"{a[:80]}; {b[:80]}; {e[:60]}")
    finally:
        s.close()


@case
def w43_old_book_refused():
    c = "W43"
    p = new_wms("w43")
    s = Session(p, macros=0)
    s.doc.unprotect(PWD)
    for name in ("Возврат", "_RET", "_ISS"):
        s.doc.Sheets.removeByName(name)
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        state, block = st(s)
        rep = s.report()
        R.add(c, "книга без листов «Возврат», _RET, _ISS (книга ядра этапа 3): запуск заблокирован с понятной причиной «нет листа «Возврат»», "
                 "проведение запрещено", state == "BLOCKED" and block == "SYS_CORRUPT" and "нет листа «Возврат»" in rep, f"{state} {block}; {rep[:160]}")
    finally:
        s.close()


# ================================================================ W44–W46 acceptance (D-069): no receipt storno while its EI has live issues / returns

@case
def w44_receipt_storno_chain():
    c = "W44"
    s = ready(Session(new_wms("w44")))
    try:
        order(s, 1, h="10", f="10")
        rc = s.click_ord("BtnRcvPost", 1)
        i1 = issue(s, 1, "201", "4")
        r1 = ret(s, 1, "1", "4")
        oracle(c, s, "приход → выдача → полный возврат")
        nj = njournal(s)
        d1 = s.click_ord("BtnRcvDelete", 1)
        R.add(c, "приход 10 → выдача 4 → полный возврат 4 (остаток снова равен приходу) → «Удалить» приход — отказ: «По этому ЕИ есть связанные "
                 "выдачи/возвраты… Сначала выполните сторно зависимых операций»; ничего не записано",
              rc == "OK:1" and i1 == "OK:2" and r1 == "OK:3" and s.stock(201) == 10.0 and d1.startswith("ERR:По этому ЕИ")
              and "есть связанные выдачи/возвраты" in d1 and "выдач 1, возвратов 1" in d1 and "Сначала выполните сторно зависимых операций" in d1
              and njournal(s) == nj and s.rcv(201)[4] == "LIVE" and s.stock_row(201)[7] == "Активен", d1)
        oracle(c, s, "отказ сторно прихода")
        dr = del_ret(s, 1)
        oracle(c, s, "сторно возврата")
        d2 = s.Rc("ReceiptDeleteRow", 1)
        R.add(c, "после сторно возврата — снова отказ: действующая выдача осталась (уже выдано 4)", dr == "OK:4" and d2.startswith("ERR")
              and "выдач 1, возвратов 0" in d2 and "уже выдано 4" in d2 and s.rcv(201)[4] == "LIVE" and njournal(s) == nj + 1, f"{dr}; {d2}")
        oracle(c, s, "отказ после сторно возврата")
        di = s.I("IssueDeleteRow", 1)
        oracle(c, s, "сторно выдачи")
        d3 = s.click_ord("BtnRcvDelete", 1)
        R.add(c, "после сторно выдачи зависимых действующих движений нет — сторно прихода проведено: остаток 0, ЕИ «Приход удалён (сторно)»; "
                 "строки выдачи и возврата остались историей «Удалено (сторно)»",
              di == "OK:5" and d3 == "OK:6" and s.stock(201) == 0.0 and s.stock_row(201)[7] == "Приход удалён (сторно)" and s.rcv(201)[4] == "STORNO"
              and s.iss(1)[17] == "Удалено (сторно)" and s.ret(1)[13] == "Удалено (сторно)" and s.ret(1)[0] == 1.0, f"{di}; {d3}")
        oracle(c, s, "сторно прихода")
        i2 = issue(s, 2, "7", "1")
        r2 = ret(s, 2, "2", "1")
        R.add(c, "№ сторнированных выдачи и возврата не переиспользуются: следующие — выдача № 2, возврат № 2",
              i2 == "OK:7" and s.iss(2)[0] == 2.0 and r2 == "OK:8" and s.ret(2)[0] == 2.0, f"{i2}; {r2}")
        oracle(c, s, "нумерация после цепочки сторно")
    finally:
        s.close()


@case
def w45_receipt_storno_many():
    c = "W45"
    s = ready(Session(new_wms("w45")))
    try:
        order(s, 1, h="20", f="20")
        s.click_ord("BtnRcvPost", 1)
        res = [issue(s, 1, "201", "5", who="ива"), issue(s, 2, "201", "3", who="пет"), issue(s, 3, "201", "2", who="ива"),
               ret(s, 1, "1", "2"), ret(s, 2, "1", "3"), ret(s, 3, "2", "3")]
        oracle(c, s, "три выдачи одного ЕИ, три возврата")
        nj = njournal(s)
        out = []

        def attempt(label):
            d = s.Rc("ReceiptDeleteRow", 1)
            out.append((label, d[:95]))
            oracle(c, s, label)
            return d
        a0 = attempt("отказ: 3 выдачи, 3 возврата")
        d1 = del_ret(s, 1)
        a1 = attempt("отказ после сторно 1-го возврата")
        d2 = del_ret(s, 2)
        a2 = attempt("отказ после сторно 2-го возврата")
        d3 = del_ret(s, 3)
        a3 = attempt("отказ после сторно всех возвратов")
        ok_ret = (all(x.startswith("OK") for x in res) and all(d.startswith("OK") for d in (d1, d2, d3))
                  and "выдач 3, возвратов 3" in a0 and "уже выдано 2" in a0 and "выдач 3, возвратов 2" in a1 and "выдач 3, возвратов 1" in a2
                  and "выдач 3, возвратов 0" in a3 and "возвратов этого ЕИ" in a0 and "возвратов этого ЕИ" not in a3)
        R.add(c, "три выдачи одного ЕИ и три возврата (две выдачи возвращены полностью): сторно прихода отказано на каждом шаге, пока есть "
                 "действующие возвраты; счётчики в отказе уменьшаются", ok_ret and njournal(s) == nj + 3, str(out[:4]))
        e1 = s.I("IssueDeleteRow", 1)
        a4 = attempt("отказ после сторно выдачи № 1")
        f2 = s.I("IssueFixRow", 2, "7", "3", "пет", DI)
        a5 = attempt("отказ после переноса выдачи № 2 на другой ЕИ")
        e3 = s.I("IssueDeleteRow", 3)
        a6 = attempt("сторно прихода после сторно всех зависимых выдач")
        R.add(c, "сторно выдач по одной (одна выдача исправлением перенесена на другой ЕИ — от этого ЕИ больше не зависит): отказ, пока есть "
                 "действующие выдачи; после последней — сторно прихода проведено, остаток 0",
              e1.startswith("OK") and "выдач 2, возвратов 0" in a4 and f2.startswith("OK") and "выдач 1, возвратов 0" in a5 and e3.startswith("OK")
              and a6.startswith("OK") and s.stock(201) == 0.0 and s.rcv(201)[4] == "STORNO" and s.stock(7) == 47.0, str(out[4:]))
    finally:
        s.close()


@case
def w46_storno_chain_crash():
    c = "W46"
    p = new_wms("w46")
    s = ready(Session(p))
    order(s, 1, h="10", f="10")
    s.click_ord("BtnRcvPost", 1)
    issue(s, 1, "201", "4")
    ret(s, 1, "1", "4")
    s.close(save=True)
    steps = (("RETURN_DEL", "Rt", "ReturnDeleteRow", lambda s: s.retrec(1)[4] == "LIVE" and s.stock(201) == 10.0,
              lambda s: s.retrec(1)[4] == "STORNO" and s.stock(201) == 6.0 and s.ret(1)[13] == "Удалено (сторно)"),
             ("ISSUE_DEL", "I", "IssueDeleteRow", lambda s: s.iss(1)[17] == "Проведено" and s.stock(201) == 6.0,
              lambda s: s.iss(1)[17] == "Удалено (сторно)" and s.stock(201) == 10.0),
             ("RECEIPT_DEL", "Rc", "ReceiptDeleteRow", lambda s: s.rcv(201)[4] == "LIVE" and s.stock(201) == 10.0 and s.stock_row(201)[7] == "Активен",
              lambda s: s.rcv(201)[4] == "STORNO" and s.stock(201) == 0.0 and s.stock_row(201)[7] == "Приход удалён (сторно)"))
    for typ, mod, fn, before, after in steps:
        s = ready(Session(p))
        s.B("TestSetFault", 2, 2)
        crash = getattr(s, mod)(fn, 1)
        s.kill()
        s = ready(Session(p))
        try:
            s.B("ActionForceUnlock")
            R.add(c, f"цепочка сторно, {typ}: сбой посреди записей (до журнала) — откат, операция не проведена", crash == "CRASH-SIM" and before(s),
                  f"{crash}; {s.stock(201)}")
            oracle(c, s, f"{typ}: после отката")
            s.B("TestSetFault", 3, 2)
            crash2 = getattr(s, mod)(fn, 1)
            s.kill()
            s = ready(Session(p))
            s.B("ActionForceUnlock")
            rec = s.B("ActionRecover")
            R.add(c, f"цепочка сторно, {typ}: сбой после журнала — «Восстановить» доводит операцию ровно один раз",
                  crash2 == "CRASH-SIM" and "восстановлено операций: 1" in rec and after(s), f"{crash2}; {rec[:90]}; {s.stock(201)}")
            oracle(c, s, f"{typ}: после восстановления")
        finally:
            s.close(save=True)
    s = ready(Session(p))
    try:
        again = s.Rc("ReceiptDeleteRow", 1)
        R.add(c, "после цепочки: повторное сторно прихода — «уже удалён», выдача и возврат — история", again.startswith("SKIP")
              and s.iss(1)[17] == "Удалено (сторно)" and s.ret(1)[13] == "Удалено (сторно)", again)
        oracle(c, s, "цепочка сторно после сбоев")
    finally:
        s.close()


@case
def w47_receipt_storno_no_confirmation():
    c = "W47"
    s = ready(Session(new_wms("w47")))
    try:
        order(s, 1, h="10", f="10")
        rc = s.click_ord("BtnRcvPost", 1)
        i1 = issue(s, 1, "201", "4")
        r1 = ret(s, 1, "1", "4")
        nj = njournal(s)

        def delete(answer):
            """«Удалить» on the receipt row; answer — what a confirmation would get (1 «Да», 2 «Нет»); returns the recorded
            result and how many confirmations were asked"""
            s.U("TestUiAuto", answer)
            k0 = s.U("TestUiConfirmCount")
            d = s.click_ord("BtnRcvDelete", 1)
            k = s.U("TestUiConfirmCount") - k0
            s.U("TestUiAuto", 1)
            return d, k
        d1, k1 = delete(1)
        R.add(c, "«Удалить» приход, по ЕИ которого есть действующие выдача и возврат: отказ сразу, окно подтверждения «Удалить приход?» не "
                 "открывалось (даже с ответом «Да» наготове); ничего не записано",
              rc == "OK:1" and i1 == "OK:2" and r1 == "OK:3" and d1.startswith("ERR:По этому ЕИ") and "выдач 1, возвратов 1" in d1
              and "Сначала выполните сторно зависимых операций" in d1 and k1 == 0 and njournal(s) == nj and s.rcv(201)[4] == "LIVE",
              f"{d1[:110]}; подтверждений {k1}")
        oracle(c, s, "отказ без подтверждения")
        dr = del_ret(s, 1)
        d2, k2 = delete(1)
        R.add(c, "после сторно возврата выдача ещё действует — снова отказ сразу, без подтверждения",
              dr == "OK:4" and d2.startswith("ERR:По этому ЕИ") and "выдач 1, возвратов 0" in d2 and k2 == 0 and s.rcv(201)[4] == "LIVE",
              f"{dr}; {d2[:110]}; подтверждений {k2}")
        di = s.I("IssueDeleteRow", 1)
        d3, k3 = delete(2)
        R.add(c, "после сторно выдачи сторно прихода допустимо: только теперь спрошено подтверждение «Удалить приход …?»; ответ «Нет» — "
                 "ничего не удалено",
              di == "OK:5" and k3 == 1 and d3.startswith("Удалить приход ЕИ-00000201?") and s.rcv(201)[4] == "LIVE" and s.stock(201) == 10.0
              and njournal(s) == nj + 2, f"{di}; {d3[:60]}; подтверждений {k3}")
        d4, k4 = delete(1)
        R.add(c, "ответ «Да» — сторно прихода проведено: остаток 0, «Приход удалён (сторно)»",
              d4 == "OK:6" and k4 == 1 and s.rcv(201)[4] == "STORNO" and s.stock(201) == 0.0 and s.stock_row(201)[7] == "Приход удалён (сторно)",
              f"{d4}; подтверждений {k4}")
        d5, k5 = delete(1)
        R.add(c, "повторное «Удалить» — «уже удалён» без подтверждения", d5.startswith("SKIP") and k5 == 0, f"{d5}; подтверждений {k5}")
        oracle(c, s, "сторно прихода после подтверждения")
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
    R.save(os.path.join(OUT, "results_phase4.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_phase4.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
