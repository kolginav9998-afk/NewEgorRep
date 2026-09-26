"""Core Phase 2 «Выдачи» — automated scenarios (CLAUDE_TASK.md «ОБЯЗАТЕЛЬНЫЕ ТЕСТЫ PHASE 2»).

    WMS_TEST_OUT=/tmp/wms_out python3 tests/run_phase2.py [case ...]

Every case copies a freshly built test book (synthetic EIs, recipients) into WMS_TEST_OUT/cases/<case>/, drives real
LibreOffice processes like a user (typing through the UI so that the sheet's change handler runs, the buttons' macros
for the row under the cursor, window close, kill -9) and checks the book and the journal with tests/issue_oracle.py.
Results: WMS_TEST_OUT/results_phase2.json and WMS_TEST_OUT/TEST_REPORT_phase2.md.
"""
import os
import re
import subprocess
import sys
import threading
import time
import zipfile

import uno  # noqa: E402  (LibreOffice Python-UNO)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Session, Results, new_wms, template, INITIAL_STOCK, synthetic_registry  # noqa: E402
from com.sun.star.sheet import TableFilterField  # noqa: E402
from com.sun.star.sheet.FilterOperator import EQUAL  # noqa: E402
from com.sun.star.sheet import TableFilterField3, FilterFieldValue  # noqa: E402

R = Results()
CASES = []
POSTED_LOCKS = "101111111111000111"
OPEN_LOCKS = "101100100110000111"
COLS = "ABCDEFGHIJKLMNOPQR"


def case(fn):
    CASES.append(fn)
    return fn


def st(s):
    d = s.state()
    return d.get("STATE"), d.get("BLOCK")


def ei(n):
    return f"ЕИ-{n:08d}"


def ready(s):
    """a session in which button messages are recorded instead of shown (TEST books only)"""
    s.U("TestUiAuto", 1)
    return s


def njournal(s):
    return len(s.journal()[0])


def snapshot(s, rows=12, eis=(1, 2, 4, 5, 6, 7, 123)):
    sh = s.doc.Sheets.getByName("Выдачи")
    vals = sh.getCellRangeByPosition(0, 1, 17, rows).getDataArray()
    locks = [s.iss_locks(r) for r in range(1, rows + 1)]
    stock = [s.stock(n) for n in eis]
    sysv = [s.sysv(k) for k in ("LAST_SEQ", "NEXT_NO", "TX_STATE", "TX_SEQ", "TX_BEFORE_IMAGE", "JOURNAL_POS")]
    return vals, locks, stock, sysv


def ctrl_z_harmless(s, n=5):
    a = snapshot(s)
    for _ in range(n):
        s.ui(".uno:Undo")
    return a == snapshot(s)


def oracle(c, s, name, expect_tail=0):
    P, inf = s.icheck(expect_tail)
    R.add(c, "оракул: " + name, not P, f"{inf}; {P[:4]}")
    return not P


def fill(s, r, e, q, d="25.09.2026", w="егр"):
    s.issue_input(r, ei=e, qty=q, date=d, who=w)


# ================================================================ V01–V09 ordinary issue and the checks before it

@case
def v01_normal_issue():
    c = "V01"
    s = ready(Session(new_wms("v01")))
    try:
        state, _ = st(s)
        fill(s, 1, "ЕИ-1", "2,5", "25.09.2026", "егр")
        prev = s.iss(1)
        ok_prev = (prev[2] == "Болт М8х30 оцинкованный" and prev[3] == "DIN933-M8x30" and prev[6] == "шт" and prev[9] == "A-01-1"
                   and prev[10] == "Крепёж" and prev[15] == 100.0 and prev[11] == "ЕИ-00000001" and prev[8] == "Ермолин Егор Павлович"
                   and prev[0] == "" and prev[17] == "")
        R.add(c, "ввод ЕИ → предпросмотр C D G J K P, L канонический, сокращение → ФИО, без окон", state == "CLEAN" and ok_prev, str(prev))
        s.click("BtnPost", 1)
        res = s.U("TestUiLastMessage")
        row, locks, stt = s.iss(1), s.iss_locks(1), s.state()
        ents, _ = s.journal()
        e = ents[-1] if ents else {}
        f = e.get("fields", {})
        ok = (res == "OK:1" and row[0] == 1.0 and row[4] == 2.5 and row[15] == 100.0 and row[16] == 97.5 and row[17] == "Проведено"
              and locks == POSTED_LOCKS and s.stock(1) == 97.5 and stt["NEXT_NO"] == "2" and stt["LAST_SEQ"] == "1"
              and stt["UNDO_LOCKED"] == "False" and stt["UNDO_CAN"] == "False" and e.get("type") == "ISSUE"
              and f.get("NO") == "1" and f.get("EI") == "ЕИ-00000001" and f.get("QTY") == "2.5" and f.get("WHO") == "Ермолин Егор Павлович"
              and f.get("DATE") == "2026-09-25" and f.get("BAL_BEFORE") == "100" and f.get("BAL_AFTER") == "97.5" and len(e["raw"]) < 4096)
        R.add(c, "«Провести» для строки под курсором: № 1, P 100, Q 97,5, «Проведено», защита, остаток −2,5, журнал ISSUE",
              ok, f"{res}; {row}; защита {locks}; {stt}; журнал: {e.get('raw', '')[:220]} ({len(e.get('raw', ''))} байт)")
        variants = ["123", "00000123", "ЕИ-123", "ЕИ-00000123", "EI-123", "ei-123", "еи 123", "EИ-123", "  ЕИ-123 "]
        out = []
        for i, v in enumerate(variants, start=2):
            s.type_iss("L", i, v)
            row = s.iss(i)
            out.append((v, row[11], row[2]))
        R.add(c, "варианты ввода ЕИ приводятся к ЕИ-00000123 и подставляют данные товара",
              all(x[1] == "ЕИ-00000123" and x[2] == "Товар 123" for x in out), str(out))
        oracle(c, s, "после обычной выдачи")
    finally:
        s.close()


def refused(s, r, before_stock_n, before_stock, nj, next_no):
    return (s.stock(before_stock_n) == before_stock and njournal(s) == nj and s.sysv("NEXT_NO") == next_no
            and s.iss(r)[0] == "")


@case
def v02_unknown_ei():
    c = "V02"
    s = ready(Session(new_wms("v02")))
    try:
        out = []
        for r, v in enumerate(["ЕИ-999", "abc", "ЕИ-12x", "0", "ЕИ-", "123456789"], start=1):
            fill(s, r, v, "1")
            pv = s.iss(r)[17]
            res = s.post(r)
            out.append((v, pv[:70], res[:70], s.iss(r)[17][:60], s.iss(r)[2]))
        ok = (all(x[2].startswith("ERR") and x[3].startswith("Не проведено") and x[4] == "" for x in out)
              and "не найден в «Наличие»" in out[0][1] and njournal(s) == 0 and s.sysv("NEXT_NO") == 1)
        R.add(c, "неизвестный или неверный ЕИ: R — понятная ошибка, проведение запрещено, ничего не записано", ok, str(out))
    finally:
        s.close()


@case
def v03_zero_stock():
    c = "V03"
    s = ready(Session(new_wms("v03")))
    try:
        fill(s, 1, "3", "1")
        pv = s.iss(1)[17]
        res = s.post(1)
        R.add(c, "нулевой остаток: блок с понятной причиной", res.startswith("ERR") and "равен 0" in pv and "равен 0" in s.iss(1)[17]
              and refused(s, 1, 3, 0.0, 0, 1), f"{pv}; {res}")
    finally:
        s.close()


@case
def v04_insufficient_stock():
    c = "V04"
    s = ready(Session(new_wms("v04")))
    try:
        fill(s, 1, "5", "6")
        pv = s.iss(1)[17]
        res = s.post(1)
        ok1 = res.startswith("ERR") and "недостаточно" in pv and refused(s, 1, 5, 5.0, 0, 1)
        s.type_iss("E", 1, "5")
        res2 = s.post(1)
        fill(s, 2, "5", "1")
        res3 = s.post(2)
        R.add(c, "количество больше остатка — блок; ровно остаток — проводится до 0; после этого — блок",
              ok1 and res2 == "OK:1" and s.stock(5) == 0.0 and res3.startswith("ERR") and "равен 0" in s.iss(2)[17],
              f"6 из 5: {res} / {pv}; 5 из 5: {res2}; ещё 1: {res3}")
        oracle(c, s, "после граничных выдач")
    finally:
        s.close()


@case
def v05_qty_comma_dot():
    c = "V05"
    s = ready(Session(new_wms("v05")))
    try:
        res = []
        for r, q in enumerate(["1,5", "1.5", "0,125", "2"], start=1):
            fill(s, r, "2", q)
            res.append(s.post(r))
        vals = [s.iss(r)[4] for r in range(1, 5)]
        R.add(c, "количество с запятой и точкой: 1,5 / 1.5 / 0,125 / 2 → числа, остаток 250,5 − 5,125",
              res == ["OK:1", "OK:2", "OK:3", "OK:4"] and vals == [1.5, 1.5, 0.125, 2.0] and abs(s.stock(2) - 245.375) < 1e-9,
              f"{res}; {vals}; остаток {s.stock(2)}")
        oracle(c, s, "дробные количества")
    finally:
        s.close()


@case
def v06_bad_quantities():
    c = "V06"
    s = ready(Session(new_wms("v06")))
    try:
        bad = ["0", "-1", "1e3", "1,500", "01.05.2026", "1 000", "abc", "1,2345", "½", "100001", "1,5,0"]
        out = []
        for r, q in enumerate(bad, start=1):
            fill(s, r, "1", q)
            res = s.post(r)
            out.append((q, res[:60], s.iss(r)[17][:70]))
        R.add(c, f"некорректные количества ({len(bad)}): все отвергнуты с понятной причиной, ничего не записано",
              all(x[1].startswith("ERR") and x[2].startswith("Не проведено") for x in out) and refused(s, 1, 1, 100.0, 0, 1), str(out))
    finally:
        s.close()


@case
def v07_date_required():
    c = "V07"
    s = ready(Session(new_wms("v07")))
    try:
        out = []
        bad = [None, "31.02.2026", "29.02.2027", "31.04.2026", "abc", "25.09.1999", "01.01.2100"]
        for r, d in enumerate(bad, start=1):
            s.issue_input(r, ei="1", qty="1", date=d, who="егр")
            pv = s.iss(r)[17]
            res = s.post(r)
            out.append((d, pv[:60], res[:70], s.iss(r)[17][:40]))
        fill(s, 10, "1", "1", "29.02.2028")
        ok_leap = s.post(10)
        fill(s, 11, "1", "1", "25.09.2026")
        ok_date = s.post(11)
        R.add(c, "дата выдачи обязательна: пусто, 31.02, 29.02 невисокосного года, 31.04, текст, вне 2000–2099 — блок с причиной "
                 "(в предпросмотре и при проведении, без ошибки макроса); 29.02.2028 и обычная дата — проводятся",
              all(x[2].startswith("ERR") and x[3].startswith("Не проведено:") for x in out) and "дата" in out[0][2]
              and all(x[1].startswith("Ошибка:") for x in out[1:]) and ok_leap == "OK:1" and ok_date == "OK:2"
              and s.stock(1) == 98.0 and njournal(s) == 2,
              f"{out}; {ok_leap}; {ok_date}")
        oracle(c, s, "даты")
    finally:
        s.close()


@case
def v08_recipient_required():
    c = "V08"
    s = ready(Session(new_wms("v08")))
    try:
        s.issue_input(1, ei="1", qty="1", date="25.09.2026")
        res = s.post(1)
        R.add(c, "пустой получатель — блок", res.startswith("ERR") and "получатель" in res and refused(s, 1, 1, 100.0, 0, 1), res)
    finally:
        s.close()


@case
def v09_recipient_abbreviation():
    c = "V09"
    s = ready(Session(new_wms("v09")))
    try:
        cases = [("егр", "Ермолин Егор Павлович"), ("ЕГР", "Ермолин Егор Павлович"), ("пр", "Производство"),
                 ("Офис", "Офис"), ("Иванов Иван Андреевич", "Иванов Иван Андреевич")]
        out = []
        for r, (typed, want) in enumerate(cases, start=1):
            fill(s, r, "1", "1", "25.09.2026", typed)
            out.append((typed, s.iss(r)[8], s.iss(r)[17], s.post(r)))
        ok1 = all(x[1] == w and x[2] == "" and x[3].startswith("OK") for x, (_, w) in zip(out, cases))
        R.add(c, "сокращения (без учёта регистра) → полное имя; полное имя и «Офис»/«Производство» — без предупреждения", ok1, str(out))
        fill(s, 6, "1", "1", "25.09.2026", "Сидоров")
        warn = s.iss(6)[17]
        res = s.post(6)
        R.add(c, "новый получатель вне справочника: предупреждение без окна, проведение разрешено",
              warn.startswith("Внимание") and res.startswith("OK") and s.iss(6)[8] == "Сидоров", f"{warn}; {res}")
        # a repeated abbreviation typed into «Получатели» (UI, so the directory is re-read) becomes ambiguous
        s.goto("Получатели", "$A$7")
        s.o.dispatch(s.doc, ".uno:EnterString", StringName="егр")
        s.goto("Получатели", "$B$7")
        s.o.dispatch(s.doc, ".uno:EnterString", StringName="Егоров Роман")
        fill(s, 7, "1", "1", "25.09.2026", "егр")
        amb = s.iss(7)[17]
        res2 = s.post(7)
        sc = s.B("ActionSelfCheck")
        R.add(c, "повторное сокращение в справочнике: не подставляется, проведение запрещено, самопроверка предупреждает",
              "несколько раз" in amb and res2.startswith("ERR") and "повторяются сокращения «егр»" in sc,
              f"{amb[:90]}; {res2[:60]}; {[x for x in sc.splitlines() if 'справочник' in x]}")
        oracle(c, s, "получатели")
    finally:
        s.close()


# ================================================================ V10–V16 repeated posting, copies, fill, protection, filter

@case
def v10_repeat_post():
    c = "V10"
    s = ready(Session(new_wms("v10")))
    try:
        fill(s, 1, "1", "2")
        r1 = s.post(1)
        nj, st1 = njournal(s), s.stock(1)
        r2 = s.post(1)
        s.click("BtnPost", 1)
        r3 = s.U("TestUiLastMessage")
        R.add(c, "повторное «Провести» проведённой строки: «уже проведено», ничего не списано, журнал не растёт",
              r1 == "OK:1" and r2.startswith("SKIP:уже проведено") and r3.startswith("SKIP:уже проведено") and njournal(s) == nj
              and s.stock(1) == st1 == 98.0 and s.iss(1)[17] == "Проведено", f"{r1}; {r2}; {r3}; журнал {nj}→{njournal(s)}")
    finally:
        s.close()


@case
def v11_double_click():
    c = "V11"
    s = ready(Session(new_wms("v11")))
    try:
        fill(s, 1, "4", "3")
        s.goto("Выдачи", "$L$2")
        res = []

        def click():
            s.U("BtnPost")
            res.append(s.U("TestUiLastMessage"))
        th = [threading.Thread(target=click) for _ in range(2)]
        for t in th:
            t.start()
        for t in th:
            t.join(60)
        ents = [e for e in s.journal()[0] if e["type"] == "ISSUE"]
        # the order of the two answers is not fixed: LibreOffice Basic yields to the event loop while a macro runs, so the
        # second press may be answered («уже проведено» / «операция уже выполняется») before the first one's «OK» is
        # recorded; the recorded last message is then «OK:1». What must hold: one posting, one decrement, and a press on the
        # posted row answers «уже проведено» — checked with one more press after both have finished
        last = s.U("TestUiLastMessage")
        s.U("BtnPost")
        again = s.U("TestUiLastMessage")
        R.add(c, "двойной клик кнопки «Провести» (два нажатия одновременно): одна выдача, одно списание, повторное нажатие — «уже проведено»",
              len(ents) == 1 and s.stock(4) == 9.0 and s.iss(1)[0] == 1.0 and s.sysv("NEXT_NO") == 2
              and (last.startswith("SKIP:уже проведено") or last == "OK:1") and again.startswith("SKIP:уже проведено")
              and len([e for e in s.journal()[0] if e["type"] == "ISSUE"]) == 1,
              f"ответы нажатий {res}; последнее {last!r}; ещё нажатие {again!r}; записей ISSUE {len(ents)}; остаток {s.stock(4)}")
        # the same through the posting function, each call's own answer recorded
        fill(s, 2, "5", "1")
        ans = []
        th = [threading.Thread(target=lambda: ans.append(s.post(2))) for _ in range(2)]
        for t in th:
            t.start()
        for t in th:
            t.join(60)
        ents = [e for e in s.journal()[0] if e["type"] == "ISSUE"]
        # the second call is answered «уже проведено» — or «операция уже выполняется» when it runs inside the first one
        R.add(c, "два одновременных вызова проведения одной строки: ответы «OK» и «уже проведено» (или «уже выполняется»), одно списание",
              len(ans) == 2 and ans.count("OK:2") == 1
              and [a for a in ans if a != "OK:2"][0] in ("SKIP:уже проведено (№ 2)", "BUSY:операция уже выполняется")
              and len(ents) == 2 and s.stock(5) == 4.0,
              f"ответы {ans}; записей ISSUE {len(ents)}; остаток {s.stock(5)}")
        oracle(c, s, "двойной клик")
    finally:
        s.close()


@case
def v12_row_copies():
    c = "V12"
    p = new_wms("v12")
    s = ready(Session(p))
    fill(s, 1, "1", "2")
    s.post(1)
    # (a) copy a posted row and paste it onto an empty row of the protected sheet
    s.goto("Выдачи", "$A$2:$R$2")
    s.ui(".uno:Copy")
    s.goto("Выдачи", "$A$10")
    s.ui(".uno:Paste")
    pasted = s.iss(9)
    R.add(c, "копировать → вставить проведённую строку на защищённом листе: вставка отклонена, № не появился",
          pasted[0] == "" and pasted[17] == "", str(pasted[:3]))
    # (b) the sheet unprotected with the password while WMS runs, a posted row pasted: the handler marks the copy
    sh = s.doc.Sheets.getByName("Выдачи")
    sh.unprotect("wms")
    s.goto("Выдачи", "$A$2:$R$2")
    s.ui(".uno:Copy")
    s.goto("Выдачи", "$A$11")
    s.ui(".uno:Paste")
    sh.protect("wms")
    live = s.iss(10)[17]
    res_live = s.post(10)
    R.add(c, "защита снята паролем, проведённая строка вставлена: обработчик сразу помечает КОПИЯ, «Провести» отказывает",
          live.startswith("КОПИЯ") and "повторяется" in live and res_live.startswith("SKIP"), f"{live[:90]}; {res_live}")
    s.close(save=True)
    # (c) macros disabled: a copy of the posted row and a key WMS never issued are written and saved
    s = Session(p, macros=0)
    sh = s.doc.Sheets.getByName("Выдачи")
    sh.unprotect("wms")
    row = list(sh.getCellRangeByPosition(0, 1, 17, 1).getDataArray()[0])
    sh.getCellRangeByPosition(0, 4, 17, 4).setDataArray((tuple(row),))
    row[0] = 99.0
    sh.getCellRangeByPosition(0, 5, 17, 5).setDataArray((tuple(row),))
    sh.protect("wms")
    s.doc.store()
    s.close()
    s = ready(Session(p))
    try:
        rep = s.report()
        c1, c2 = s.iss(4)[17], s.iss(5)[17]
        res = s.post(4)
        clr = s.I("IssueClearRow", 4)
        clr2 = s.I("IssueClearRow", 5)
        clr3 = s.I("IssueClearRow", 10)
        after = (s.iss(4), s.iss_locks(4))
        R.add(c, "сохранено без макросов: копия и ключ, который WMS не выдавала, помечены КОПИЯ при запуске; не проводятся; «Очистить» убирает",
              "без WMS" in rep and c1.startswith("КОПИЯ") and "повторяет" in c1 and c2.startswith("КОПИЯ") and "не выдавался" in c2
              and res.startswith("SKIP") and clr.startswith("OK") and clr2.startswith("OK") and clr3.startswith("OK")
              and after[0][0] == "" and after[1] == OPEN_LOCKS and s.stock(1) == 98.0,
              f"{c1[:70]}; {c2[:70]}; {res}; {clr}; строка после очистки {after}")
        oracle(c, s, "копии не считаются проведёнными")
    finally:
        s.close()


@case
def v13_fill_series():
    c = "V13"
    s = ready(Session(new_wms("v13")))
    try:
        fill(s, 1, "1", "2")
        s.post(1)
        s.goto("Выдачи", "$A$2:$A$6")
        s.ui(".uno:FillDown")
        keys = [s.iss(r)[0] for r in range(1, 6)]
        R.add(c, "«Заполнить вниз» по № от проведённой строки: отказ, поддельных № нет", keys == [1.0, "", "", "", ""], str(keys))
        s.goto("Выдачи", "$A$2:$R$6")
        s.ui(".uno:FillDown")
        rows = [s.iss(r)[0] for r in range(1, 6)]
        R.add(c, "«Заполнить вниз» целой проведённой строкой: отказ", rows == [1.0, "", "", "", ""], str(rows))
        s.type_iss("L", 10, "ЕИ-20")
        s.type_iss("E", 10, "1")
        s.goto("Выдачи", "$L$11:$L$15")
        s.ui(".uno:FillSeries", FillDir="B", FillCmd="L", FillStep="1", FillDateCmd="D", FillStart="", FillMax="")
        s.goto("Выдачи", "$E$11:$E$15")
        s.ui(".uno:FillSeries", FillDir="B", FillCmd="L", FillStep="1", FillDateCmd="D", FillStart="", FillMax="")
        got = [(s.iss(r)[11], s.iss(r)[2], s.iss(r)[4], s.iss(r)[0], s.iss(r)[17]) for r in range(10, 15)]
        # Calc reads «ЕИ-00000010» as text + the number −10, so the series runs 10, 9, 8, 7, 6: whatever the direction,
        # every filled row must show the data of its own EI
        names = {row[0]: row[1] for row in synthetic_registry()}
        R.add(c, "«Заполнить ряд» ЕИ и количества в непроведённых строках: предпросмотр обновлён во всех строках (данные своего ЕИ), "
                 "проведённых строк нет",
              len({g[0] for g in got}) == 5 and all(g[0] in names and g[1] == names[g[0]] for g in got)
              and [g[2] for g in got] == ["1", "2", "3", "4", "5"] and all(g[3] == "" and g[4] == "" for g in got) and njournal(s) == 1,
              str(got))
        oracle(c, s, "после заполнения")
    finally:
        s.close()


@case
def v14_posted_fields_protected():
    c = "V14"
    s = ready(Session(new_wms("v14")))
    try:
        fill(s, 1, "1", "2")
        s.post(1)
        before = s.iss(1)
        changed = []
        for col in "ACDEFGHIJKLPQR":
            s.type_iss(col, 1, "999")
            if s.iss(1) != before:
                changed.append(col)
        s.goto("Выдачи", "$A$2")
        s.ui(".uno:SelectRow")
        s.ui(".uno:DeleteRows")
        after_del = s.iss(1)
        s.goto("Выдачи", "$A$1:$R$3")
        s.ui(".uno:SortDescending")
        after_sort = s.iss(1)
        R.add(c, "проведённая строка: ручной ввод в A C D E F G H I J K L P Q R, удаление строки и сортировка — отклонены",
              not changed and after_del == before and after_sort == before and s.iss_locks(1) == POSTED_LOCKS,
              f"изменились колонки: {changed}; после удаления {after_del[:2]}; после сортировки {after_sort[:2]}")
        oracle(c, s, "защита")
    finally:
        s.close()


@case
def v15_comment_editable():
    c = "V15"
    s = ready(Session(new_wms("v15")))
    try:
        fill(s, 1, "1", "2")
        s.post(1)
        nj = njournal(s)
        s.type_iss("O", 1, "выдано по заявке 17")
        s.type_iss("B", 1, "АКТ-5")
        row = s.iss(1)
        R.add(c, "в проведённой строке «Комментарий» (O) и «№ документа» (B) редактируются; учёт и журнал не меняются",
              row[14] == "выдано по заявке 17" and row[1] == "АКТ-5" and row[17] == "Проведено" and njournal(s) == nj and s.stock(1) == 98.0,
              str(row))
        oracle(c, s, "комментарий")
    finally:
        s.close()


def set_filter(s, col, value):
    dbr = s.doc.DatabaseRanges.getByName("WMS_ISSUES")
    fd = dbr.getFilterDescriptor()
    f = TableFilterField()
    f.Field, f.Operator, f.IsNumeric, f.StringValue = col, EQUAL, False, value
    fd.setFilterFields((f,))
    dbr.refresh()


def visible(s, r):
    return s.doc.Sheets.getByName("Выдачи").getRows().getByIndex(r).IsVisible


@case
def v16_active_autofilter():
    c = "V16"
    s = ready(Session(new_wms("v16")))
    try:
        for r, (e, w) in enumerate([("1", "Офис"), ("2", "пет"), ("4", "Офис"), ("5", "пет")], start=1):
            fill(s, r, e, "1", "25.09.2026", w)
            s.post(r)
        fill(s, 5, "6", "0,5", "25.09.2026", "Офис")          # entered, not posted yet
        set_filter(s, 8, "Офис")
        vis = [visible(s, r) for r in range(1, 6)]
        p5 = s.click("BtnPost", 5) or s.U("TestUiLastMessage")
        s.U("TestFixInput", "ЕИ-00000001", "3", "Офис", "25.09.2026")
        s.click("BtnFix", 1)
        fx = s.U("TestUiLastMessage")
        s.click("BtnDelete", 3)
        dl = s.U("TestUiLastMessage")
        vis2 = [visible(s, r) for r in range(1, 6)]
        hidden_rows = (s.iss(2), s.iss(4))
        ok = (vis == [True, False, True, False, True] and p5 == "OK:5" and fx.startswith("OK") and dl.startswith("OK") and vis2 == vis
              and hidden_rows[0][17] == "Проведено" and hidden_rows[1][17] == "Проведено"
              and s.stock(6) == 3.25 and s.stock(1) == 97.0 and s.stock(4) == 12.0)
        R.add(c, "при активном автофильтре (Кому = Офис): «Провести», «Исправить», «Удалить» меняют видимые строки под курсором, скрытые не тронуты, фильтр остаётся",
              ok, f"видимость до {vis} после {vis2}; {p5}; {fx}; {dl}; скрытые: {hidden_rows[0][17]}, {hidden_rows[1][17]}")
        oracle(c, s, "автофильтр")
    finally:
        s.close()


# ================================================================ V17–V19 correction and storno

@case
def v17_fix_quantity():
    c = "V17"
    s = ready(Session(new_wms("v17")))
    try:
        fill(s, 1, "4", "5", "25.09.2026", "ива")
        s.post(1)
        up = s.I("IssueFixRow", 1, "ЕИ-00000004", "8", "Иванов Иван Андреевич", "25.09.2026")
        a = (s.iss(1)[4], s.iss(1)[15], s.iss(1)[16], s.stock(4), s.iss(1)[17])
        down = s.I("IssueFixRow", 1, "4", "2,5", "ива", "25.09.2026")
        b = (s.iss(1)[4], s.iss(1)[15], s.iss(1)[16], s.stock(4))
        snap = snapshot(s)
        nj = njournal(s)
        too = s.I("IssueFixRow", 1, "4", "13", "ива", "25.09.2026")
        R.add(c, "исправление количества вверх (5→8) и вниз (8→2,5): остаток и P/Q пересчитаны; больше доступного (13 из 12) — отказ без изменений",
              up.startswith("OK") and a == (8.0, 12.0, 4.0, 4.0, "Проведено (исправлено)") and down.startswith("OK") and b == (2.5, 12.0, 9.5, 9.5)
              and too.startswith("ERR") and "недостаточно" in too and snapshot(s) == snap and njournal(s) == nj,
              f"вверх {up} {a}; вниз {down} {b}; 13: {too}")
        same = s.I("IssueFixRow", 1, "4", "2,5", "ива", "25.09.2026")
        who = s.I("IssueFixRow", 1, "4", "2,5", "пет", "26.09.2026")
        R.add(c, "исправление без изменений — «ничего не изменилось»; смена получателя и даты — проводится",
              same.startswith("SKIP") and who.startswith("OK") and s.iss(1)[8] == "Петров Пётр Сергеевич" and s.stock(4) == 9.5, f"{same}; {who}")
        oracle(c, s, "исправления количества")
    finally:
        s.close()


@case
def v18_fix_ei():
    c = "V18"
    s = ready(Session(new_wms("v18")))
    try:
        fill(s, 1, "1", "10")
        s.post(1)
        res = s.I("IssueFixRow", 1, "ЕИ-7", "5", "егр", "25.09.2026")
        row = s.iss(1)
        ok = (res.startswith("OK") and row[11] == "ЕИ-00000007" and row[2] == "Товар 7" and row[4] == 5.0 and s.stock(1) == 100.0
              and s.stock(7) == (7 * 7) % 50 + 1 - 5 and row[15] == float((7 * 7) % 50 + 1) and row[0] == 1.0)
        R.add(c, "исправление ЕИ (1 → 7): старый ЕИ вернул 10, новый уменьшился на 5, строка — данные нового ЕИ, № тот же", ok,
              f"{res}; {row}; остатки 1={s.stock(1)} 7={s.stock(7)}")
        snap = snapshot(s)
        bad = s.I("IssueFixRow", 1, "3", "1", "егр", "25.09.2026")
        bad2 = s.I("IssueFixRow", 1, "999", "1", "егр", "25.09.2026")
        R.add(c, "исправление на ЕИ с нулевым остатком или неизвестный ЕИ — отказ, прежнее состояние цело",
              bad.startswith("ERR") and bad2.startswith("ERR") and snapshot(s) == snap, f"{bad}; {bad2}")
        oracle(c, s, "исправление ЕИ")
    finally:
        s.close()


@case
def v19_delete_storno():
    c = "V19"
    s = ready(Session(new_wms("v19")))
    try:
        fill(s, 1, "2", "10,5")
        fill(s, 2, "2", "1")
        s.post(1)
        s.post(2)
        d = s.I("IssueDeleteRow", 1)
        row = s.iss(1)
        ents = s.journal()[0]
        fill(s, 3, "2", "1")
        nxt = s.post(3)
        again = s.I("IssueDeleteRow", 1)
        clr = s.I("IssueClearRow", 1)
        fix = s.I("IssueFixRow", 1, "2", "1", "егр", "25.09.2026")
        ok = (d.startswith("OK") and row[17] == "Удалено (сторно)" and row[0] == 1.0 and row[4] == 10.5 and row[11] == "ЕИ-00000002"
              and s.iss_locks(1) == POSTED_LOCKS and ents[-1]["type"] == "ISSUE_DEL" and ents[-1]["fields"]["NO"] == "1"
              and nxt == "OK:4" and s.iss(3)[0] == 3.0 and abs(s.stock(2) - 248.5) < 1e-9 and again.startswith("SKIP")
              and clr.startswith("ERR") and fix.startswith("ERR") and all(visible(s, r) for r in range(1, 4)))
        R.add(c, "«Удалить» = сторно: остаток вернулся, журнал ISSUE_DEL, строка осталась «Удалено (сторно)» с № и данными, № не переиспользован, строка не скрыта",
              ok, f"{d}; {row}; следующая {nxt} №{s.iss(3)[0]}; повторно {again}; очистить {clr}; исправить {fix}; остаток {s.stock(2)}")
        oracle(c, s, "сторно")
    finally:
        s.close()


# ================================================================ V20–V25 crashes, recovery, Ctrl+Z, save/reopen

@case
def v20_crash_before_journal():
    c = "V20"
    p = new_wms("v20")
    s = ready(Session(p))
    fill(s, 1, "1", "2")
    s.post(1)
    fill(s, 2, "4", "3", "25.09.2026", "пет")
    s.B("TestSetFault", 1, 1)
    rb = s.post(2)
    row_rb = s.iss(2)
    ok_rb = rb.startswith("ERR-RB") and row_rb[0] == "" and row_rb[4] == "3" and row_rb[17].startswith("Не проведено") and s.stock(4) == 12.0
    R.add(c, "ошибка до записи журнала → полный откат: ввод на месте, № и остаток не тронуты, причина в R", ok_rb, f"{rb}; {row_rb}")
    s.close(save=True)
    s = ready(Session(p))
    s.B("TestSetFault", 2, 2)
    crash = s.post(2)
    s.kill()
    s = ready(Session(p))
    try:
        unl = s.B("ActionForceUnlock")
        state, _ = st(s)
        row = s.iss(2)
        ok = (crash == "CRASH-SIM" and "отменена по снимку" in unl and state == "CLEAN" and row[0] == "" and s.stock(4) == 12.0
              and njournal(s) == 1 and s.sysv("NEXT_NO") == 2)
        R.add(c, "книга сохранена посреди выдачи (до журнала) + kill → при запуске откат по маркеру, выдачи нет нигде", ok,
              f"{crash}; {unl[:200]}; строка {row}")
        again = s.post(2)
        R.add(c, "повторное «Провести» проводит один раз", again == "OK:2" and s.stock(4) == 9.0, again)
        oracle(c, s, "сбой до журнала")
    finally:
        s.close()


@case
def v21_crash_after_journal():
    c = "V21"
    p = new_wms("v21")
    s = ready(Session(p))
    fill(s, 1, "1", "2")
    s.post(1)
    s.close(save=True)
    s = ready(Session(p))
    fill(s, 2, "6", "1,25", "25.09.2026", "ива")
    s.B("TestSetFault", 3, 2)
    crash = s.post(2)
    s.kill()
    s = ready(Session(p))
    try:
        unl = s.B("ActionForceUnlock")
        _, block = st(s)
        rec = s.B("ActionRecover")
        state, _ = st(s)
        row = s.iss(2)
        ok = (crash == "CRASH-SIM" and block == "TAIL" and "восстановлено операций: 1" in rec and state == "CLEAN" and row[0] == 2.0
              and row[17] == "Проведено" and s.stock(6) == 2.5)
        R.add(c, "книга сохранена после записи журнала + kill → откат по маркеру, «Восстановить» доводит выдачу из журнала ровно один раз", ok,
              f"{crash}; {unl[:160]}; {rec[:120]}; строка {row}")
        oracle(c, s, "сбой после журнала")
    finally:
        s.close()


@case
def v22_recover_unsaved():
    c = "V22"
    p = new_wms("v22")
    s = ready(Session(p))
    fill(s, 1, "1", "1")
    s.post(1)
    s.close(save=True)
    s = ready(Session(p))
    for r, (e, q, w) in enumerate([("2", "3,5", "егр"), ("4", "2", "Офис"), ("5", "1", "пр")], start=2):
        fill(s, r, e, q, "26.09.2026", w)
        s.post(r)
    s.I("IssueFixRow", 3, "4", "4", "Офис", "26.09.2026")
    s.I("IssueDeleteRow", 4)
    s.I("IssueDeleteRow", 1)
    s.close(save=False)
    s = ready(Session(p))
    try:
        rep = s.report()
        _, block = st(s)
        before = s.iss(2)
        rec = s.B("ActionRecover")
        state, _ = st(s)
        rows = [s.iss(r) for r in range(1, 5)]
        ok = (block == "TAIL" and before[0] == "" and "восстановлено операций: 6" in rec and state == "CLEAN"
              and [r[17] for r in rows] == ["Удалено (сторно)", "Проведено", "Проведено (исправлено)", "Удалено (сторно)"]
              and rows[1][11] == "ЕИ-00000002" and rows[1][4] == 3.5 and rows[1][8] == "Ермолин Егор Павлович" and rows[2][4] == 4.0)
        R.add(c, "«Не сохранять» после 3 выдач, исправления и 2 сторно → запуск: блок «хвост»; «Восстановить» возвращает всё, включая ввод",
              ok, f"{rep[:150]}; {rec[:120]}; {[r[17] for r in rows]}")
        oracle(c, s, "восстановление несохранённого")
        rec2 = s.B("ActionRecover")
        R.add(c, "повторное «Восстановить» — «не требуется», ничего не меняет", "не требуется" in rec2 and not s.icheck()[0], rec2[:100])
    finally:
        s.close(save=True)
    s = Session(p)
    try:
        R.add(c, "после сохранения и перезапуска — чисто", st(s)[0] == "CLEAN" and not s.icheck()[0], s.report()[:120])
    finally:
        s.close()


@case
def v23_repeat_recovery_after_crash():
    c = "V23"
    p = new_wms("v23")
    s = ready(Session(p))
    fill(s, 1, "1", "1")
    s.post(1)
    s.close(save=True)
    s = ready(Session(p))
    fill(s, 2, "2", "2")
    s.post(2)
    fill(s, 3, "4", "1")
    s.post(3)
    s.kill()
    s = ready(Session(p))
    s.B("ActionForceUnlock")
    rec1 = s.B("ActionRecover")
    s.kill()                                                   # crash again right after the recovery, nothing saved
    s = ready(Session(p))
    try:
        s.B("ActionForceUnlock")
        _, block = st(s)
        rec2 = s.B("ActionRecover")
        rec3 = s.B("ActionRecover")
        R.add(c, "восстановление, затем снова авария без сохранения → повторное «Восстановить» доводит те же 2 выдачи ровно один раз",
              "восстановлено операций: 2" in rec1 and block == "TAIL" and "восстановлено операций: 2" in rec2 and "не требуется" in rec3
              and s.stock(2) == 248.5 and s.stock(4) == 11.0, f"{rec1[:60]}; {block}; {rec2[:60]}; {rec3[:60]}")
        oracle(c, s, "повторное восстановление")
    finally:
        s.close()


@case
def v24_ctrl_z():
    c = "V24"
    s = ready(Session(new_wms("v24")))
    try:
        fill(s, 1, "1", "2")
        s.post(1)
        z1 = ctrl_z_harmless(s)
        s.I("IssueFixRow", 1, "1", "3", "егр", "25.09.2026")
        z2 = ctrl_z_harmless(s)
        s.I("IssueDeleteRow", 1)
        z3 = ctrl_z_harmless(s)
        stt = s.state()
        s.type_iss("O", 5, "ручной ввод")
        typed = s.iss(5)[14]
        s.ui(".uno:Undo")
        R.add(c, "Ctrl+Z ×5 после проведения, исправления и удаления ничего не откатывает; обычный ввод отменяется",
              z1 and z2 and z3 and stt["UNDO_LOCKED"] == "False" and typed == "ручной ввод" and s.iss(5)[14] == "",
              f"после проведения {z1}, исправления {z2}, удаления {z3}; {stt}")
        oracle(c, s, "Ctrl+Z")
    finally:
        s.close()


@case
def v25_save_reopen():
    c = "V25"
    p = new_wms("v25")
    s = ready(Session(p))
    fill(s, 1, "1", "2")
    s.post(1)
    fill(s, 2, "2", "1")
    s.post(2)
    s.I("IssueDeleteRow", 2)
    fill(s, 3, "4", "1")                  # entered, not posted
    s.close(save=True)
    with zipfile.ZipFile(p) as z:
        xml = z.read("content.xml").decode("utf-8")
    s = ready(Session(p))
    try:
        state, _ = st(s)
        rows = [s.iss(r) for r in range(1, 4)]
        un = s.main_status()[3]
        checks = {"CLEAN": state == "CLEAN", "не «изменена» после открытия": not s.doc.isModified(),
                  "лист Выдачи": s.active_sheet() == "Выдачи", "R1": rows[0][17] == "Проведено",
                  "R2": rows[1][17] == "Удалено (сторно)", "ввод цел": rows[2][0] == "" and rows[2][2] == "Перчатки нитриловые",
                  "защита 1": s.iss_locks(1) == POSTED_LOCKS, "защита 3": s.iss_locks(3) == OPEN_LOCKS,
                  "вставка строк": 'loext:insert-rows="true"' in xml, "автофильтр": s.doc.DatabaseRanges.hasByName("WMS_ISSUES"),
                  "Главная": "1 строк" in un}
        nxt = s.post(3)
        # OK:<seq>: post, post, storno were seq 1..3; the new issue gets seq 4 and № 3 (№ 2 of the storno is not reused)
        checks["проведение"] = nxt == "OK:4" and s.iss(3)[0] == 3.0
        bad = [k for k, v in checks.items() if not v]
        R.add(c, "сохранить → закрыть → открыть: строки, статусы, защита, опция «вставка строк», автофильтр целы; непроведённая строка видна на «Главной»; "
                 "открытие без изменений не делает книгу «изменённой»",
              not bad, f"не выполнено: {bad}; {state}; {[r[17] for r in rows]}; «Главная»: {un}; {nxt}; активный лист {s.active_sheet()}")
        oracle(c, s, "после повторного открытия")
    finally:
        s.close()


# ================================================================ V26–V30 the rest of the contract

@case
def v26_percent_column():
    c = "V26"
    s = ready(Session(new_wms("v26")))
    try:
        fill(s, 1, "1", "2")
        s.type_iss("F", 1, "50")
        pv = s.iss(1)[17]
        res = s.post(1)
        R.add(c, "«Количество в %» (F) не используется для проведения: заполненная F — понятный отказ, ничего не списано",
              res.startswith("ERR") and "%" in pv and s.stock(1) == 100.0 and njournal(s) == 0, f"{pv}; {res}")
    finally:
        s.close()


@case
def v27_ui_blocked_and_buttons():
    c = "V27"
    p = new_wms("v27")
    s = Session(p)
    names = {}
    for sheet in ("Выдачи", "Главная"):
        form = s.doc.Sheets.getByName(sheet).getDrawPage().getForms().getByIndex(0)
        for i in range(form.getCount()):
            ev = form.getScriptEvents(i)
            code = ev[0].ScriptCode if ev else ""
            ok_loc = code.startswith("vnd.sun.star.script:Standard.") and code.endswith("?language=Basic&location=document")
            names[form.getByIndex(i).Label] = (sheet, code.split("Standard.", 1)[1].split("?")[0] if ok_loc else code)
    want = {"Провести": "WmsUi.BtnPost", "Исправить": "WmsUi.BtnFix", "Удалить": "WmsUi.BtnDelete", "Очистить": "WmsUi.BtnClear",
            "Восстановить": "WmsUi.BtnRecover", "Отложить хвост журнала": "WmsUi.BtnAbandon", "Снять блокировку WMS": "WmsUi.BtnUnlock",
            "Сделать рабочим файлом": "WmsUi.BtnRegister", "Самопроверка": "WmsUi.BtnSelfCheck", "Резервная копия": "WmsUi.BtnBackup"}
    R.add(c, "кнопки: «Провести», «Исправить», «Удалить», «Очистить» на «Выдачи» и 6 действий восстановления на «Главной» привязаны к макросам",
          all(names.get(k, (None, None))[1] == v for k, v in want.items()), str(names))
    s.kill()
    s = ready(Session(p))
    try:
        active = s.active_sheet()
        ms = s.main_status()
        fill(s, 1, "1", "1")
        res = s.post(1)
        r_status = s.iss(1)[17]
        ok1 = (active == "Главная" and ms[0].startswith("ЗАБЛОКИРОВАНО [LOCKED]") and "WMS уже открыта" in ms[1]
               and "Снять блокировку WMS" in ms[2] and res.startswith("BLOCKED") and "WMS заблокирована" in r_status and "Главная" in r_status)
        R.add(c, "блокировка при запуске: книга открывается на «Главной» с причиной и подсказкой; «Провести» отказывает и пишет причину в R",
              ok1, f"активный лист {active}; {ms[:3]}; R: {r_status[:120]}")
        s.U("BtnUnlock")
        ms2 = s.main_status()
        res2 = s.post(1)
        R.add(c, "кнопка «Снять блокировку WMS» (с подтверждением): состояние «Работа разрешена», последнее действие показано, проведение работает",
              ms2[0] == "Работа разрешена" and "Снять блокировку" in ms2[5] and res2 == "OK:1", f"{ms2}; {res2}")
    finally:
        s.close(save=False)
    s = ready(Session(p))
    try:
        ms = s.main_status()
        s.U("BtnRecover")
        ms2 = s.main_status()
        R.add(c, "несохранённая выдача: «Главная» показывает хвост журнала; кнопка «Восстановить» доводит его",
              ms[0].startswith("ЗАБЛОКИРОВАНО [TAIL]") and "Восстановить" in ms[2] and ms2[0] == "Работа разрешена"
              and "восстановлено операций: 1" in ms2[5] and s.iss(1)[17] == "Проведено", f"{ms[:3]}; {ms2[5][:80]}")
        s.U("BtnSelfCheck")
        sc = s.main_status()[5]
        R.add(c, "кнопка «Самопроверка»: отчёт на «Главной», реестр и справочник проверены, ошибок нет",
              sc.startswith("САМОПРОВЕРКА WMS: ошибок 0") and "реестр «Наличие»" in sc and "справочник «Получатели»" in sc, sc[:300])
        oracle(c, s, "интерфейс")
    finally:
        s.close()


@case
def v28_crash_during_fix():
    c = "V28"
    p = new_wms("v28")
    s = ready(Session(p))
    fill(s, 1, "1", "10")
    s.post(1)
    s.close(save=True)
    s = ready(Session(p))
    s.B("TestSetFault", 2, 2)
    crash = s.I("IssueFixRow", 1, "7", "5", "пр", "25.09.2026")
    s.kill()
    s = ready(Session(p))
    s.B("ActionForceUnlock")
    state1, _ = st(s)
    row1 = s.iss(1)
    ok1 = (crash == "CRASH-SIM" and state1 == "CLEAN" and row1[11] == "ЕИ-00000001" and row1[4] == 10.0 and s.stock(1) == 90.0
           and s.stock(7) == (7 * 7) % 50 + 1)
    R.add(c, "авария посреди исправления (до журнала): при запуске старое движение цело, частичного состояния нет", ok1,
          f"{crash}; {row1}; остатки 1={s.stock(1)} 7={s.stock(7)}")
    s.B("TestSetFault", 3, 2)
    crash2 = s.I("IssueFixRow", 1, "7", "5", "пр", "25.09.2026")
    s.kill()
    s = ready(Session(p))
    try:
        s.B("ActionForceUnlock")
        rec = s.B("ActionRecover")
        row2 = s.iss(1)
        ok2 = (crash2 == "CRASH-SIM" and "восстановлено операций: 1" in rec and row2[11] == "ЕИ-00000007" and row2[4] == 5.0
               and row2[17] == "Проведено (исправлено)" and s.stock(1) == 100.0 and s.stock(7) == (7 * 7) % 50 + 1 - 5)
        R.add(c, "авария после записи исправления в журнал: «Восстановить» доводит составную операцию целиком (сторно + новое движение)",
              ok2, f"{crash2}; {rec[:100]}; {row2}")
        oracle(c, s, "исправление при сбоях")
    finally:
        s.close()


@case
def v29_inserted_row():
    c = "V29"
    s = ready(Session(new_wms("v29")))
    try:
        fill(s, 1, "1", "1")
        s.post(1)
        fill(s, 2, "2", "1")
        s.post(2)
        s.goto("Выдачи", "$A$3")
        s.ui(".uno:SelectRow")
        s.ui(".uno:InsertRowsBefore")
        moved = s.iss(3)[0] == 2.0 and s.iss(2)[0] == ""
        locked = s.iss_locks(2)
        s.type_iss("L", 2, "4")
        typed_before = s.iss(2)[11]
        clr = s.I("IssueClearRow", 2)
        fill(s, 2, "4", "2")
        res = s.post(2)
        R.add(c, "вставка строки между проведёнными: разрешена; вставленная строка закрыта как соседняя, «Очистить» открывает её для ввода, выдача проводится",
              moved and locked == POSTED_LOCKS and typed_before == "" and clr.startswith("OK") and res == "OK:3" and s.iss(3)[0] == 2.0,
              f"сдвиг {moved}; защита вставленной {locked}; ввод до «Очистить» {typed_before!r}; {clr}; {res}")
        oracle(c, s, "вставленная строка")
    finally:
        s.close()


@case
def v30_many_issues_and_selfcheck():
    c = "V30"
    s = ready(Session(new_wms("v30")))
    try:
        n = 60
        for i in range(n):
            e = 8 + (i % 40)
            s.issue_input(1 + i, ei=str(e), qty="1", date="25.09.2026", who=["егр", "Офис", "пр"][i % 3])
        t0 = time.time()
        res = [s.post(1 + i) for i in range(n)]
        dt = (time.time() - t0) / n * 1000
        sc = s.B("ActionSelfCheck")
        R.add(c, f"{n} выдач подряд по разным ЕИ; самопроверка без ошибок", all(r.startswith("OK") for r in res)
              and sc.splitlines()[0].startswith("САМОПРОВЕРКА WMS: ошибок 0"), f"{dt:.1f} мс/выдача; {sc.splitlines()[0]}",
              timing=round(dt, 1))
        oracle(c, s, f"{n} выдач")
    finally:
        s.close()


def set_filter_values(s, col, values):
    """an AutoFilter with several ticked values (what the dropdown list produces): TableFilterField3"""
    dbr = s.doc.DatabaseRanges.getByName("WMS_ISSUES")
    fd = dbr.getFilterDescriptor()
    f = TableFilterField3()
    f.Field, f.Operator = col, uno.getConstantByName("com.sun.star.sheet.FilterOperator2.EQUAL")
    vals = []
    for v in values:
        x = FilterFieldValue()
        x.IsNumeric, x.StringValue = False, v
        x.FilterType = uno.getConstantByName("com.sun.star.sheet.FilterFieldType.STRING")     # the default is NUMERIC
        vals.append(x)
    f.Values = tuple(vals)
    fd.setFilterFields3((f,))
    dbr.refresh()


@case
def v31_save_with_filter():
    c = "V31"
    p = new_wms("v31")
    s = ready(Session(p))
    try:
        for i in range(9):
            fill(s, 1 + i, str(8 + i), "1", "25.09.2026", ["Офис", "пр", "егр"][i % 3])
            s.post(1 + i)
        fill(s, 10, "20", "1", "25.09.2026", "пет")                  # entered, not posted
        set_filter_values(s, 8, ["Офис", "Производство"])
        vis0 = [visible(s, r) for r in range(1, 11)]
        t0 = time.time()
        s.doc.store()
        save_s = time.time() - t0
        vis1 = [visible(s, r) for r in range(1, 11)]
        fields = s.doc.DatabaseRanges.getByName("WMS_ISSUES").getFilterDescriptor().getFilterFields3()
        modified = s.doc.isModified()
        harmless = ctrl_z_harmless(s)
        with zipfile.ZipFile(p) as z:
            xml = z.read("content.xml").decode("utf-8")
        nfilt = xml.count('table:visibility="filter"')
        ok1 = (vis0 == [True, True, False] * 3 + [False] and vis1 == vis0 and len(fields) == 1 and len(fields[0].Values) == 2
               and not modified and nfilt == 0 and harmless)
        R.add(c, "сохранение при активном автофильтре (два значения): на время сохранения скрытие строк снято, условия фильтра "
                 "не удалены, все строки сохранены показанными (файл со скрытыми строками LibreOffice открывал бы минутами); "
                 "после сохранения фильтр восстановлен, книга не «изменена», Ctrl+Z безвреден",
              ok1, f"видимость {vis0} → {vis1}; условий {len(fields)}; строк, сохранённых скрытыми, {nfilt}; изменена {modified}; "
                   f"сохранение {save_s:.2f} с")
        s.close()
        s = ready(Session(p))
        state, _ = st(s)
        vis2 = [visible(s, r) for r in range(1, 11)]
        modified2 = s.doc.isModified()
        s.goto("Выдачи", "$L$11")
        nxt = s.click("BtnPost") or s.U("TestUiLastMessage")
        R.add(c, "после повторного открытия: WMS снова применила тот же фильтр (файл открылся без скрытых строк), книга не «изменена», "
                 "работа разрешена; «Провести» работает при фильтре",
              state == "CLEAN" and vis2 == vis0 and not modified2 and nxt == "OK:10", f"{state}; {vis2}; изменена {modified2}; {nxt}")
        oracle(c, s, "сохранение с фильтром")
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
    R.save(os.path.join(OUT, "results_phase2.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_phase2.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
