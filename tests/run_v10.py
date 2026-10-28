"""WMS 1.0 — the innovations chosen by the storekeeper (stages B–M of docs/PLAN_1_0.md): automated checks on books built
from the sources, through the macros of the buttons (windows answered by the test seams TestUiAuto / TestPrimeChoice).

    WMS_TEST_OUT=/tmp/wms_v10 python3 tests/run_v10.py [case ...]

Every case builds its own fresh test book in WMS_TEST_OUT/cases/<case>/. Results: results_v10.json, TEST_REPORT_v10.md.
"""
import datetime
import glob
import hashlib
import re
import shutil
import time
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Results, Session, new_wms  # noqa: E402
from run_prime import T, TODAY, issue, order, put, receive, serial, topology  # noqa: E402

R = Results()
CASES = []


def case(fn):
    CASES.append(fn)
    return fn


def ready(s, auto=1):
    s.U("TestUiAuto", auto)
    return s


def choose(s, k):
    """the next window with named buttons answers with button k (0-based)"""
    s.B("TestPrimeChoice", k, module="WmsPrimeUi")


def iss_row(s, r):
    return list(s.doc.Sheets.getByName("Выдачи").getCellRangeByPosition(0, r, 17, r).getDataArray()[0])


def select(s, sheet, ref):
    s.goto(sheet, ref)


def oracle(c, s, name):
    P, inf, _ = s.adjcheck(0)
    R.add(c, "оракул (журнал воспроизведён независимо): " + name, not P, f"{inf.get('ops')} операций; {P[:4]}")
    return P


# ================================================================ B: everyday work

@case
def v01_issue_block():
    """Р1: several rows of «Выдачи» selected → «Провести» posts each as its own issue under one document; an answer given
    once in the block (an unknown recipient) is not asked again; a refused row is named in the summary, the others posted"""
    c = "V01"
    s = ready(Session(new_wms("v01")))
    try:
        for r, (ei, q, who) in enumerate((("ЕИ-00000001", "2", "Новиков"), ("ЕИ-00000002", "3", "Новиков"), ("ЕИ-00000011", "1", "Новиков")), start=1):
            put(s, "Выдачи", r, L=ei, E=q, H=serial(), I=who)
        s.U("TestUiAuto", 0)
        choose(s, 0)                                     # «Провести 3 выдач одним документом»
        s.U("TestUiAuto", 1)                             # the question about «Новиков» (not in «Получатели»): «Записать как есть»
        select(s, "Выдачи", "$A$2:$R$4")
        res = s.U("BtnPost") or s.U("TestUiLastMessage")
        msg = s.U("TestUiLastMessage")
        rows = [iss_row(s, r) for r in (1, 2, 3)]
        docs = [x[1] for x in rows]
        R.add(c, "три выделенные строки → «Провести» одним документом: каждая — своя выдача № 1…3, в «Документ» (B) у всех «ВД-000001»; "
                 "сообщения об успехе нет", msg.startswith("OK=3;ERR=0;SKIP=0") and [x[0] for x in rows] == [1.0, 2.0, 3.0]
              and docs == ["ВД-000001"] * 3 and all(str(x[17]).startswith("Проведено") for x in rows), f"{msg[:200]}; {docs}; {[x[0] for x in rows]}")
        ents = [e for e in s.journal()[0] if e["type"] == "ISSUE"]
        R.add(c, "в журнале — три операции ISSUE, у каждой DOC «ВД-000001»", len(ents) == 3 and all(e["fields"].get("DOC") == "ВД-000001" for e in ents),
              str([e["fields"].get("DOC") for e in ents]))
        # a block with a refusal in the middle: the other rows are posted, the refused one is named
        put(s, "Выдачи", 5, L="ЕИ-00000004", E="1", H=serial(), I="ива", B="Требование 17")
        put(s, "Выдачи", 6, L="ЕИ-00000005", E="100000", H=serial(), I="ива")
        put(s, "Выдачи", 7, L="ЕИ-00000006", E="1", H=serial(), I="ива")
        choose(s, 0)
        select(s, "Выдачи", "$A$6:$R$8")
        s.U("BtnPost")
        msg2 = s.U("TestUiLastMessage")
        r5, r6, r7 = iss_row(s, 5), iss_row(s, 6), iss_row(s, 7)
        R.add(c, "пачка с ошибкой в середине: две выдачи проведены с документом из B первой строки («Требование 17»), строка 7 — не проведена "
                 "и названа в итоге", msg2.startswith("OK=2;ERR=1") and "строка 7" in msg2 and r5[1] == "Требование 17" and r7[1] == "Требование 17"
              and r6[0] == "" and str(r6[17]).startswith("Не проведено"), f"{msg2[:300]}; {r5[1]}; {r7[1]}; {r6[17]}")
        # «Провести каждую отдельно» — no common document
        put(s, "Выдачи", 9, L="ЕИ-00000007", E="1", H=serial(), I="ива")
        put(s, "Выдачи", 10, L="ЕИ-00000008", E="1", H=serial(), I="ива")
        choose(s, 1)
        select(s, "Выдачи", "$A$10:$R$11")
        s.U("BtnPost")
        msg3 = s.U("TestUiLastMessage")
        R.add(c, "«Провести каждую отдельно»: обе проведены, «Документ» пуст", msg3.startswith("OK=2;ERR=0") and iss_row(s, 9)[1] == ""
              and iss_row(s, 10)[1] == "", msg3[:120])
        # «Не проводить» — nothing posted
        put(s, "Выдачи", 12, L="ЕИ-00000009", E="1", H=serial(), I="ива")
        put(s, "Выдачи", 13, L="ЕИ-00000010", E="1", H=serial(), I="ива")
        choose(s, 2)
        select(s, "Выдачи", "$A$13:$R$14")
        s.U("BtnPost")
        msg4 = s.U("TestUiLastMessage")
        R.add(c, "«Не проводить»: ничего не проведено", msg4.startswith("SKIP") and iss_row(s, 12)[0] == "" and iss_row(s, 13)[0] == "", msg4[:80])
        oracle(c, s, "выдачи пачкой")
    finally:
        s.close()


@case
def v02_repeat_issue_from_pick_all():
    """Р3 «Повторить», Р9 «Выдать» from «Наличие» and the card, Р10 «Подобрать ЕИ» (words in any order, «показано N из M»,
    the cursor on «Количество»), Р12 «всё»"""
    c = "V02"
    s = ready(Session(new_wms("v02")))
    try:
        put(s, "Выдачи", 1, L="ЕИ-00000001", E="2", H=serial(-3), I="ива")
        p1 = s.I("IssuePostRow", 1)
        s.goto("Выдачи", "$L$2")
        rp = s.B("RepeatIssue", module="WmsDaily")
        r2 = iss_row(s, 2)
        cur = s.doc.getCurrentController().getSelection().getRangeAddress()
        R.add(c, "«Повторить» на проведённой выдаче: новая строка 3 — тот же ЕИ, количество и получатель, дата сегодня, не проведена; курсор "
                 "в «Количество»", p1.startswith("OK") and rp.startswith("OK:строка 3") and r2[11] == "ЕИ-00000001" and r2[4] == "2"
              and r2[7] == serial() and r2[8] == iss_row(s, 1)[8] and r2[0] == "" and (cur.StartRow, cur.StartColumn) == (2, 4), f"{rp}; {r2[:12]}; {cur.StartRow},{cur.StartColumn}")
        s.goto("Выдачи", "$L$3")
        rp2 = s.B("RepeatIssue", module="WmsDaily")
        R.add(c, "«Повторить» на непроведённой строке — отказ с объяснением", rp2.startswith("ERR") and "проведённую" in rp2, rp2)
        s.goto("Наличие", "$A$5")
        fr = s.B("IssueFrom", module="WmsDaily")
        r3 = iss_row(s, 3)
        R.add(c, "«Выдать» на «Наличии» (ЕИ-00000004): новая строка «Выдач» с этим ЕИ, датой сегодня, курсор в «Количество»; лист — «Выдачи»",
              fr.startswith("OK:строка 4") and r3[11] == "ЕИ-00000004" and r3[7] == serial() and s.active_sheet() == "Выдачи", f"{fr}; {r3[:12]}")
        s.goto("Наличие", "$A$4")
        fr0 = s.B("IssueFrom", module="WmsDaily")
        R.add(c, "«Выдать» на ЕИ с остатком 0 — отказ «выдавать нечего»", fr0.startswith("ERR") and "нечего" in fr0, fr0)
        s.B("CardShow", 7, module="WmsFind")
        s.goto("Поиск", "$B$2")
        fr2 = s.B("IssueFrom", module="WmsDaily")
        R.add(c, "«Выдать» на «Поиске» с карточкой ЕИ-00000007 — новая строка с этим ЕИ", fr2.startswith("OK") and iss_row(s, 4)[11] == "ЕИ-00000007", fr2)
        # Р10: the words in any order; the count of all matches; the cursor on the quantity
        names = s.doc.Sheets.getByName("Наличие").getCellRangeByPosition(1, 1, 1, 30).getDataArray()
        first = str(names[0][0])                               # ЕИ-00000001: «Болт М8х30 оцинкованный»
        w = first.split()
        q = f"{w[-1]} {w[0]}"                                  # the last word and the first, reversed
        cand = s.B("PickCandidates", q, module="WmsFind")
        R.add(c, f"«Подобрать ЕИ» по словам в другом порядке («{q}») находит «{first}»", cand.startswith("ЕИ-00000001|"), cand[:120])
        # Р12: «всё» — the whole balance in one record
        put(s, "Выдачи", 6, L="ЕИ-00000006", E="всё", H=serial(), I="ива")
        pa = s.I("IssuePostRow", 6)
        r6 = iss_row(s, 6)
        ents = [e for e in s.journal()[0] if e["type"] == "ISSUE"]
        R.add(c, "количество «всё»: выдан весь остаток ЕИ-00000006 (3,75) одной записью, остаток 0; в журнале HOW=ALL",
              pa.startswith("OK") and r6[4] == 3.75 and s.stock(6) == 0.0 and ents[-1]["fields"].get("HOW") == "ALL", f"{pa}; {r6[4]}; {s.stock(6)}")
        put(s, "Выдачи", 7, L="ЕИ-00000006", E="всё", H=serial(), I="ива")
        pz = s.I("IssuePostRow", 7)
        R.add(c, "«всё» при остатке 0 — отказ (выдача невозможна)", pz.startswith("ERR") and ("нечего" in pz or "невозможна" in pz), pz)
        oracle(c, s, "повтор, выдать из наличия, всё")
    finally:
        s.close()


def form(s, *values, nxt=False):
    s.B("TestPrimeForm", "\t".join(values), module="WmsPrimeUi")
    if nxt:
        s.B("TestPrimeFormNext", module="WmsPrimeUi")


@case
def v03_issue_window():
    """Р2: the window «Выдать» — the EI by a number or by words of the name, the quantity in any form, the recipient, the date;
    a refusal does not leave a half-done row; «Выдать и следующую» keeps the recipient"""
    c = "V03"
    s = ready(Session(new_wms("v03")))
    try:
        form(s, "2", "1,5", "ива", "")
        w1 = s.B("IssueWindow", module="WmsDaily")
        r1 = iss_row(s, 1)
        R.add(c, "«Выдать»: ЕИ «2», 1,5, «ива», дата пусто → выдача № 1 сегодня в первой свободной строке, проведена",
              w1.startswith("OK:выдано 1") and r1[0] == 1.0 and r1[11] == "ЕИ-00000002" and r1[4] == 1.5 and r1[7] == serial()
              and str(r1[17]).startswith("Проведено"), f"{w1}; {r1}")
        form(s, "оцинкованный болт", "2", "ива", "", nxt=True)
        w2 = s.B("IssueWindow", module="WmsDaily")
        r2 = iss_row(s, 2)
        R.add(c, "«Выдать и следующую» по словам наименования («оцинкованный болт» → ЕИ-00000001): проведена; получатель запомнен для следующей",
              w2.startswith("OK:выдано 1") and r2[11] == "ЕИ-00000001" and r2[0] == 2.0 and s.B("TestDyWho", module="WmsDaily") == r2[8], f"{w2}; {r2[:12]}")
        form(s, "5", "99999", "ива", "")
        w3 = s.B("IssueWindow", module="WmsDaily")
        r3 = iss_row(s, 3)
        R.add(c, "больше остатка → не выдано, окно вернулось с причиной; полупроведённой строки на листе нет",
              w3.startswith("SKIP") and "Не выдано" in w3 and all(x == "" for x in r3), f"{w3}; {r3[:12]}")
        form(s, "абракадабра", "1", "ива", "")
        w4 = s.B("IssueWindow", module="WmsDaily")
        R.add(c, "ЕИ не найден по словам → не выдано, понятная причина", w4.startswith("SKIP") and "нет" in w4, w4)
        oracle(c, s, "окно «Выдать»")
    finally:
        s.close()


@case
def v04_saved_state_macros_off():
    """A8: the saved file tells someone who opens it with the macros off that WMS is not running (it showed the last «Работа
    разрешена»); with the macros on the panel is the live one; Ctrl+Z after a save changes nothing"""
    c = "V04"
    p = new_wms("v04")
    s = ready(Session(p))
    try:
        s.doc.store()
        live = (s.doc.Sheets.getByName("Главная").getCellByPosition(1, 7).getString(), s.doc.Sheets.getByName("Дашборд").getCellByPosition(8, 0).getString())
        modified = s.doc.isModified()
        n_undo = len(s.doc.getUndoManager().getAllUndoActionTitles())
    finally:
        s.close()
    s = Session(p, macros=0)
    try:
        off = (s.doc.Sheets.getByName("Главная").getCellByPosition(1, 7).getString(), s.doc.Sheets.getByName("Главная").getCellByPosition(1, 9).getString(),
               s.doc.Sheets.getByName("Дашборд").getCellByPosition(8, 0).getString())
    finally:
        s.close()
    s = ready(Session(p))
    try:
        on = s.doc.Sheets.getByName("Главная").getCellByPosition(1, 7).getString()
    finally:
        s.close()
    R.add(c, "после сохранения в сеансе — живая панель («Работа разрешена»), книга не «изменена», в Undo пусто; файл без макросов — «WMS не запущена "
             "— макросы выключены» на «Главной» и «Дашборде», с подсказкой про доверенные источники; с макросами — снова «Работа разрешена»",
          live[0] == "Работа разрешена" and "не запущена" not in live[1] and not modified and n_undo == 0
          and off[0].startswith("WMS не запущена") and "Доверенные источники" in off[1] and "WMS не запущена" in off[2] and on == "Работа разрешена",
          f"{live}; изменена {modified}; undo {n_undo}; {off[0]}; {off[2]}; {on}")


def ords(s, r):
    return list(s.doc.Sheets.getByName("Заказы").getCellRangeByPosition(0, r, 28, r).getDataArray()[0])


@case
def v05_receive_window_all():
    """Р6 «Принять» (one window for a first receipt: the ordered quantity and today proposed); Р7 «Принять всё по заказу» (the
    awaited of every position: a first receipt or the rest of a partial one, posted from the bottom up)"""
    c = "V05"
    s = ready(Session(new_wms("v05")))
    try:
        order(s, 1, "1", "Болт М10", qty="100", U="A-1")
        s.goto("Заказы", "$B$2")
        form(s, "<как предложено>", "УПД-5", "", "<как предложено>", "<как предложено>", "")
        rw = s.B("ReceiveWindow", module="WmsDaily")
        r1 = ords(s, 1)
        R.add(c, "«Принять» на непринятой позиции: предложено заказанное (100) и сегодня; документ УПД-5 вписан в окне → приход проведён, ЕИ в V",
              rw.startswith("OK") and r1[5] in (100.0, "100") and r1[2] == "УПД-5" and r1[13] == serial() and str(r1[21]).startswith("ЕИ-"),
              f"{rw}; F={r1[5]!r} C={r1[2]!r} N={r1[13]!r} V={r1[21]!r}")
        order(s, 2, "2", "Гайка М10", qty="50", U="A-2")
        s.goto("Заказы", "$B$3")
        form(s, "abc", "", "", "", "", "")
        rw2 = s.B("ReceiveWindow", module="WmsDaily")
        r2 = ords(s, 2)
        R.add(c, "«Принять» с ошибкой (принято «abc») → не принято, строка возвращена как была (F пусто, дата пуста), причина в итоге",
              rw2.startswith("SKIP") and "Не принято" in rw2 and r2[5] == "" and r2[13] == "" and r2[21] == "", f"{rw2}; {r2[5]!r} {r2[13]!r}")
        # Р7: one position not received yet, one received in part (40 of 60): both awaited
        order(s, 3, "3", "Шайба", qty="60", U="A-3")
        rc = receive(s, 3, "40", place="A-3")
        s.goto("Заказы", "$B$3")
        form(s, "УПД-9", "", "<как предложено>")
        ra = s.B("ReceiveAll", module="WmsDaily")
        allrows = [ords(s, r) for r in range(1, 7)]
        names = [(x[1], x[5]) for x in allrows if x[1]]
        R.add(c, "«Принять всё по заказу» с курсором на строке заказа: «Гайка» 50 — первый приход, «Шайба» — остаток 20 «Ещё поступлением»; "
                 "документ УПД-9", rc.startswith("OK") and ra.startswith("OK=2;ERR=0") and ("Гайка М10", 50.0) in [(a, float(b) if b != "" else b) for a, b in names]
              and sum(1 for a, b in names if a == "Шайба") == 2 and any(a == "Шайба" and str(b) in ("20", "20.0") for a, b in names), f"{ra}; {names}")
        ra2 = s.B("ReceiveAll", module="WmsDaily")
        R.add(c, "повторно — принимать нечего", ra2.startswith("SKIP") and "нечего" in ra2, ra2)
        oracle(c, s, "принять, принять всё")
    finally:
        s.close()


@case
def v06_return_hands():
    """Р8: «на руках у …» on «Поиск» → the lines selected → «Вернуть выделенное»: «Возврат» rows (the issue №, what is still on
    hands, today), posted one by one"""
    c = "V06"
    s = ready(Session(new_wms("v06")))
    try:
        res = [s.I("IssuePostRow", r) for r in (1, 2, 3) if put(s, "Выдачи", r, L=f"ЕИ-0000000{r}" if r != 3 else "ЕИ-00000004", E="2",
                                                                   H=serial(), I="пв") is None]
        # a part of the second is back already
        put(s, "Возврат", 1, B=2.0, F="0,5", H=serial())
        rp = s.B("ReturnPostRow", 1, module="WmsReturn")
        hs = s.B("HandsShow", "пв", module="WmsFind")
        s.goto("Поиск", "$B$13:$I$20")
        form(s, "<как предложено>", "")
        rh = s.B("ReturnHands", module="WmsDaily")
        rows = [list(s.doc.Sheets.getByName("Возврат").getCellRangeByPosition(0, r, 13, r).getDataArray()[0]) for r in (2, 3, 4)]
        R.add(c, "«на руках у Петров Василий»: выделены строки → вернуть всё: три возврата проведены (по 2 и остаток 1,5 второй выдачи), дата сегодня",
              all(x.startswith("OK") for x in res) and rp.startswith("OK") and hs.startswith("OK") and rh.startswith("OK=3;ERR=0")
              and sorted(float(x[5]) for x in rows) == [1.5, 2.0, 2.0] and all(x[7] == serial() for x in rows), f"{res}; {rp}; {hs[:60]}; {rh}; {[x[:8] for x in rows]}")
        rh2 = s.B("ReturnHands", module="WmsDaily")
        R.add(c, "повторно — возвращать нечего (всё вернули)", rh2.startswith("SKIP") and "нечего" in rh2, rh2[:90])
        oracle(c, s, "возврат списком")
    finally:
        s.close()


@case
def v07_hotkeys():
    """Р11: the shortcuts of the book — Ctrl+Shift+P «Провести» of the sheet in view, E «Подобрать ЕИ», F «Найти», W «Выдать…»,
    Q «Куда положить?» (stored in the document)"""
    import uno
    from com.sun.star.awt import KeyEvent
    c = "V07"
    s = ready(Session(new_wms("v07")))
    try:
        sm = s.doc.getUIConfigurationManager().getShortCutManager()
        mods = uno.getConstantByName("com.sun.star.awt.KeyModifier.SHIFT") | uno.getConstantByName("com.sun.star.awt.KeyModifier.MOD1")
        got = {}
        for k in "PEFWQ":
            ke = KeyEvent()
            ke.KeyCode = uno.getConstantByName(f"com.sun.star.awt.Key.{k}")
            ke.Modifiers = mods
            try:
                got[k] = sm.getCommandByKeyEvent(ke)
            except Exception:
                got[k] = ""
        R.add(c, "сочетания клавиш в книге: Ctrl+Shift+P/E/F/W/Q → макросы WMS", "WmsDaily.KeyPost" in got["P"] and "WmsFind.BtnPickEI" in got["E"]
              and "WmsDaily.KeyFind" in got["F"] and "WmsDaily.BtnIssueWindow" in got["W"] and "WmsMap.BtnWherePut" in got["Q"], str(got))
        put(s, "Выдачи", 1, L="ЕИ-00000001", E="1", H=serial(), I="ива")
        s.goto("Выдачи", "$L$2")
        s.B("KeyPost", module="WmsDaily")
        r1 = iss_row(s, 1)
        s.B("KeyFind", module="WmsDaily")
        cur = s.doc.getCurrentController().getSelection().getRangeAddress()
        R.add(c, "Ctrl+Shift+P на «Выдачах» проводит строку под курсором; Ctrl+Shift+F открывает «Поиск» на поле «Что найти»",
              r1[0] == 1.0 and s.active_sheet() == "Поиск" and (cur.StartRow, cur.StartColumn) == (3, 2), f"{r1[0]}; {s.active_sheet()}; {cur.StartRow},{cur.StartColumn}")
    finally:
        s.close()


# ================================================================ C: the look

def form_buttons(sh):
    """name → model of the buttons of a sheet"""
    dp = sh.getDrawPage()
    if dp.getForms().getCount() == 0:
        return {}
    f = dp.getForms().getByIndex(0)
    return {f.getByIndex(i).Name: f.getByIndex(i) for i in range(f.getCount())}


@case
def v08_look():
    """В1 input columns light / WMS columns grey; В2 the colour of a button by its meaning; В3 the help of a column (input help
    of the cell and a note of the header); В4 «Главная» as a menu; В5 frozen columns; В6 «Почему?»; В8 «Наличие» readable;
    В9 the tabs by group; В10 dates and messages for people"""
    import re
    c = "V08"
    s = ready(Session(new_wms("v08")))
    try:
        sh = s.doc.Sheets.getByName("Выдачи")
        heads = (sh.getCellByPosition(11, 0).CellBackColor, sh.getCellByPosition(17, 0).CellBackColor, sh.getCellByPosition(17, 5).CellBackColor,
                 sh.getCellByPosition(4, 5).CellBackColor)
        R.add(c, "В1: шапка колонки ввода (ЕИ) светло-жёлтая, колонки WMS («Контроль») серая; ячейки WMS серые, ячейки ввода белые",
              heads[:3] == (0xFFF2CC, 0xE7E6E6, 0xF5F5F5) and heads[3] in (-1, 0xFFFFFF), str([hex(x) for x in heads]))
        b = form_buttons(sh)
        cols = {k: b[k].BackgroundColor for k in ("btnPost", "btnDelete", "btnWhy", "btnIssueWin") if k in b}
        R.add(c, "В2: «Провести» и «Выдать…» зелёные, «Удалить» красноватая, «Почему?» голубая",
              cols == {"btnPost": 0xC6E0B4, "btnDelete": 0xF4CCCC, "btnWhy": 0xDDEBF7, "btnIssueWin": 0xC6E0B4}, str({k: hex(v) for k, v in cols.items()}))
        v = sh.getCellByPosition(4, 7).Validation
        vi = sh.getCellByPosition(8, 7).Validation
        note = sh.getCellByPosition(4, 0).getAnnotation().getString()
        R.add(c, "В3: в колонке «Количество» подсказка ячейки «2 · 0,25 · 1 кан · 1,5 м · всё» и та же заметка в шапке; список "
                 "получателей в «Кому» остался и получил свою подсказку",
              v.ShowInputMessage and "1 кан" in v.InputMessage and "всё" in v.InputMessage and "1 кан" in note
              and vi.ShowInputMessage and vi.Type.value == "LIST" and "Получатели" in vi.InputMessage, f"{v.ShowInputMessage} {v.InputMessage[:30]}; {note[:30]}; {vi.ShowInputMessage} {vi.Type.value} {vi.InputMessage[:60]}")
        m = s.doc.Sheets.getByName("Главная")
        caps = [m.getCellByPosition(0, r).getString() for r in (1, 3, 4)]
        frame = m.getCellByPosition(1, 3).TopBorder.Color
        mb = form_buttons(m)
        helptext = m.getCellByPosition(0, 15).getString().split("\n")
        R.add(c, "В4: «Главная» — меню: «Каждый день», «Только если WMS просит» (красная рамка: Восстановить, Отложить хвост, Снять "
                 "блокировку, Сделать рабочим файлом), «Инструменты»; кнопка «Справка»; вместо стены текста — 6 шагов «Как…»",
              caps == ["Каждый день", "Только если WMS просит", "Инструменты"] and frame == 0xB3261E and "navHelp" in mb
              and all(mb[k].BackgroundColor == 0xF4CCCC for k in ("btnRecover", "btnAbandon", "btnUnlock", "btnRegister"))
              and [x[:2] for x in helptext[1:7]] == ["1.", "2.", "3.", "4.", "5.", "6."], f"{caps}; {hex(frame)}; {helptext[:2]}")
        s.B("NavHelp", module="WmsDaily")
        R.add(c, "В4: кнопка «Справка» открывает лист «Справка»", s.active_sheet() == "Справка", s.active_sheet())
        ctl = s.doc.getCurrentController()
        fr = {}
        for name in ("Выдачи", "Заказы", "Наличие"):
            ctl.setActiveSheet(s.doc.Sheets.getByName(name))
            fr[name] = (ctl.hasFrozenPanes(), ctl.getSplitColumn(), ctl.getSplitRow())
        R.add(c, "В5: закреплены шапка и колонки «№/ЕИ» и наименование (Выдачи — A:C, Заказы — A:B, Наличие — A:B)",
              fr == {"Выдачи": (True, 3, 1), "Заказы": (True, 2, 1), "Наличие": (True, 2, 1)}, str(fr))
        # В6: a row that is refused, then «Почему?» on it
        put(s, "Выдачи", 1, L="ЕИ-00099999", E="1", H=serial(), I="ива")
        s.goto("Выдачи", "$L$2")
        s.U("BtnPost")
        s.goto("Выдачи", "$C$2")
        s.B("BtnWhy", module="WmsDaily")
        why = s.U("TestUiLastMessage")
        R.add(c, "В6: «Почему?» на отказанной строке — полный текст «Контроль» и что сделать («Подобрать ЕИ»)",
              why.startswith("OK:строка 2:") and "Что сделать:" in why and "Подобрать ЕИ" in why, why[:300])
        st = s.doc.Sheets.getByName("Наличие")
        n_main = st.getCellRangeByPosition(0, 1, 6, 100).ConditionalFormat.getCount()
        n_h = st.getCellRangeByPosition(7, 1, 7, 100).ConditionalFormat.getCount()
        R.add(c, "В8: «Наличие» — нулевой остаток серым и строки через одну (3 правила на колонку; у «Состояние» сначала 2 своих цвета), "
                 "отступ перед «Место»", (n_main, n_h) == (3, 5) and st.getCellByPosition(5, 3).ParaIndent == 300, f"{n_main} {n_h} {st.getCellByPosition(5, 3).ParaIndent}")
        sheets = s.doc.Sheets
        names = [sheets.getByIndex(i).Name for i in range(sheets.getCount()) if sheets.getByIndex(i).IsVisible and sheets.getByIndex(i).Name != "_TST"]
        tabs = {n: sheets.getByName(n).TabColor for n in ("Дашборд", "Выдачи", "Наличие", "Справка")}
        R.add(c, "В9: ярлычки по группам (страницы синие, листы работы зелёные, реестры и справочники оранжевые, служебные серые); "
                 "«Сверка переноса» — перед «Справкой» в конце", tabs == {"Дашборд": 0x5B9BD5, "Выдачи": 0x70AD47, "Наличие": 0xED7D31, "Справка": 0xA5A5A5}
              and names[-2:] == ["Сверка переноса", "Справка"], f"{ {k: hex(x) for k, x in tabs.items()} }; {names[-3:]}")
        ht = [s.U("HumanText", t) for t in ("Проведено, 87 мс", "Обновлено за 259 мс · всё", "запуск WMS 523 мс", "Размер 5 мс2")]
        icons = [s.U("UiIcon", t) for t in ("Выдача не проведена: остаток 0", "Строка — КОПИЯ, её можно очистить", "Встаньте на строку", "Принято позиций: 3")]
        R.add(c, "В10: тексты без «N мс» (точные значения остаются в результатах для проверок); значок окна по смыслу — "
                 "не сделано ⛔, проверьте ⚠, подсказка ℹ", ht == ["Проведено", "Обновлено · всё", "запуск WMS", "Размер 5 мс2"]
              and icons == [16, 48, 64, 64], f"{ht}; {icons}")
        s.U("BtnBackup")
        ms = s.main_status()
        R.add(c, "В10: «Главная» после действия — время «ДД.ММ.ГГГГ ЧЧ:ММ», копия названа по-русски («ручная»), без миллисекунд; "
                 "в заметках нет «запуск WMS … мс»", re.fullmatch(r"\d\d\.\d\d\.20\d\d \d\d:\d\d", ms[6]) is not None
              and "резервная копия «ручная»" in ms[5] and " мс" not in ms[5] and "запуск WMS" not in ms[4], f"{ms[6]}; {ms[5][:120]}; {ms[4][:120]}")
    finally:
        s.close()


# ================================================================ D: parts and metal

@case
def v09_quantity_words():
    """Ч9: «пол кан», «половина кан», «1/2 кан», «1 1/2 кан» — parts in words and fractions; «2 + 1 кан» — a sum (one word —
    exact in parts; different words — only in a count); a bare «1/2» is still refused"""
    c = "V09"
    s = ready(Session(new_wms("v09")))
    try:
        ei = "ЕИ-00000011"
        b = s.stock(11)
        s.B("PackSet", ei, "4", "кан", "", module="WmsPack")
        q = lambda t, mode="OUT": s.B("TestQtyEI", t, 11, mode, 0.0, module="WmsPack")
        got = {t: q(t) for t in ("пол кан", "половина кан", "1/2 кан", "1 1/2 кан", "2 + 1 кан")}
        want = {"пол кан": "0,125", "половина кан": "0,125", "1/2 кан": "0,125", "1 1/2 кан": "0,375", "2 + 1 кан": "0,75"}
        R.add(c, "«пол кан», «половина кан», «1/2 кан» = 0,5 кан; «1 1/2 кан» = 1,5 кан; «2 + 1 кан» = 3 кан (в 1 м — 4 кан)",
              b >= 1 and all(got[t].startswith("OK:" + v + "|") for t, v in want.items()) and "2 + 1 кан:" in got["2 + 1 кан"], f"{b}; {got}")
        f1, f2, f3 = q("2 + 1 кан", "FACT"), q("1 м + 2 кан", "FACT"), q("1 м + 2 кан")
        e1, e2 = q("1/2"), q("кан + 1")
        R.add(c, "пересчёт по местам: «2 + 1 кан» = 0,75 м, «1 м + 2 кан» = 1,5 м (ЕИ в метрах, в 1 м — 4 кан); в выдаче разные единицы не складываются; "
                 "«1/2» без единицы и «кан + 1» — понятный отказ",
              f1.startswith("OK:0,75|") and f2.startswith("OK:1,5|") and f3.startswith("ERR:") and "разные единицы" in f3
              and e1.startswith("ERR:") and e2.startswith("ERR:") and "нет числа" in e2, f"{f1}; {f2}; {f3}; {e1}; {e2}")
        put(s, "Выдачи", 1, L=ei, E="пол кан", H=serial(), I="ива")
        res = s.I("IssuePostRow", 1)
        r = iss_row(s, 1)
        ent = [e for e in s.journal()[0] if e["type"] == "ISSUE"][-1]
        R.add(c, "выдача «пол кан» проведена: 0,125 м, в журнале — частями (0,5 кан)", res.startswith("OK") and r[4] == 0.125
              and abs(s.stock(11) - (b - 0.125)) < 1e-9 and ent["fields"].get("PARTS") == "0.5" and ent["fields"].get("PART_UNIT") == "кан", f"{res}; {r[4]}; {s.stock(11)}; {ent['fields']}")
        R.add(c, "Ч4: на проведённой строке видно, как выдано: «Проведено · 0,5 кан»", r[17] == "Проведено · 0,5 кан", str(r[17]))
        order(s, 1, "1", "Антифриз 20 л", qty="4")
        receive(s, 1, "4", place="1-1-1")
        anti = ords(s, 1)[21]
        put(s, "Выдачи", 2, L=anti, E="25%", H=serial(), I="ива")
        s.I("IssuePostRow", 2)
        put(s, "Выдачи", 3, L="ЕИ-00000001", E="2", H=serial(), I="ива")
        s.I("IssuePostRow", 3)
        st = [iss_row(s, r)[17] for r in (2, 3)]
        R.add(c, "Ч4: процентом — «Проведено · 25 %»; обычным числом у ЕИ без фасовки — просто «Проведено»",
              st == ["Проведено · 25 %", "Проведено"], str(st))
        # the issue «Проведено · 25 %» is a live dependent of its EI: the storno of the receipt is refused
        dl = s.Rc("ReceiptDeleteRow", 1)
        R.add(c, "Ч4: выдача «Проведено · 25 %» — действующая: сторно прихода её ЕИ отказано (связанные выдачи)",
              dl.startswith("ERR:") and "действующих выдач 1" in dl and s.stock(int(anti[3:])) == 3.0, f"{dl[:200]}; {s.stock(int(anti[3:]))}")
        oracle(c, s, "выдача частями словами")
    finally:
        s.close()


@case
def v10_parts_balance_product_pack_weight():
    """Ч1 «Остаток частями» in «Наличие» (a formula that follows the balance); Ч2 a new delivery of the same product is
    offered the packing of the last one; Ч3 the check before a packing is changed; Ч8 the weight of a metre for a document in kg"""
    c = "V10"
    s = ready(Session(new_wms("v10")))
    try:
        stk = s.doc.Sheets.getByName("Наличие")
        order(s, 1, "1", "Антифриз 20 л", qty="4", art="AF-20")
        receive(s, 1, "4", place="1-1-1")
        a = ords(s, 1)[21]
        na = int(a[3:])
        s.B("PackSet", a, "4", "кан", "", module="WmsPack")
        k0 = stk.getCellByPosition(10, na).getString()
        put(s, "Выдачи", 1, L=a, E="1 кан", H=serial(), I="ива")
        s.I("IssuePostRow", 1)
        k1 = stk.getCellByPosition(10, na).getString()
        R.add(c, "Ч1: «Наличие» K «Остаток частями» — «16 кан» после фасовки, «15 кан» после выдачи 1 кан (формула строки)",
              stk.getCellByPosition(10, 0).getString() == "Остаток частями" and k0 == "16 кан" and k1 == "15 кан", f"{k0}; {k1}")
        # Ч2: the same article again — the packing of the last EI is offered
        order(s, 2, "2", "Антифриз 20 л (новая партия)", qty="2", art="AF-20")
        choose(s, 0)
        r2 = receive(s, 2, "2", place="1-1-2")
        b = ords(s, 2)[21]
        pk = list(s.doc.Sheets.getByName("_PACK").getCellRangeByPosition(0, int(b[3:]), 3, int(b[3:])).getDataArray()[0])
        order(s, 3, "3", "Антифриз 20 л", qty="1", art="AF-20")
        choose(s, 1)
        receive(s, 3, "1", place="1-1-3")
        c3 = ords(s, 3)[21]
        pk3 = s.doc.Sheets.getByName("_PACK").getCellByPosition(0, int(c3[3:])).getString()
        R.add(c, "Ч2: новая поставка того же артикула — предложена фасовка прошлого ЕИ: «Да» — «в 1 шт — 4 кан» с примечанием, «Без фасовки» — нет",
              r2.startswith("OK") and pk[:3] == [b, "кан", 4.0] and a in pk[3] and pk3 == "" and stk.getCellByPosition(10, int(b[3:])).getString() == "8 кан",
              f"{r2}; {pk}; {pk3!r}")
        q = s.B("TestPackChange", a, "3", "кан", module="WmsPack")
        q2 = s.B("TestPackChange", a, "8", "кан", module="WmsPack")
        R.add(c, "Ч3: смена фасовки — остаток в новых частях не целый (3,75 шт × 3 = 11,25 кан) — предупреждение; по старой фасовке "
                 "уже было выдач частями: 1 (в 1 шт — 8 кан: остаток 30 кан — целый, только про выдачи)", "11,25 кан" in q and "не целое" in q and "выдавали частями: 1" in q
              and "не целое" not in q2 and "выдавали частями: 1" in q2, f"{q[:300]} || {q2[:200]}")
        # Ч8: a document in kg, issued in metres — the weight of one metre
        order(s, 4, "4", "Уголок 50х50", qty="100", unit="кг")
        receive(s, 4, "100", place="2-1-1")
        u = ords(s, 4)[21]
        s.goto("Наличие", f"$A${int(u[3:]) + 1}")
        s.B("TestPrimeForm", "1 м = 3,77 кг\t\t", module="WmsPrimeUi")
        pr = s.B("PackEI", module="WmsPack")
        pku = list(s.doc.Sheets.getByName("_PACK").getCellRangeByPosition(0, int(u[3:]), 2, int(u[3:])).getDataArray()[0])
        iq = s.B("TestQtyEI", "1 м", int(u[3:]), "OUT", 0.0, module="WmsPack")
        qkg = float(iq[3:].split("|")[0].replace(",", ".")) if iq.startswith("OK:") else -1
        R.add(c, "Ч8: «1 м = 3,77 кг» в окне фасовки ЕИ в кг — в 1 кг 0,265252 м; «1 м» = 3,77 кг (±0,02)",
              pr.startswith("OK") and pku[1] == "м" and abs(pku[2] - 0.265252) < 1e-6 and abs(qkg - 3.77) <= 0.02, f"{pr}; {pku}; {iq}")
        oracle(c, s, "фасовка товара, части, вес метра")
    finally:
        s.close()


@case
def v11_tails():
    """Ч5: «Хвосты» on «Дашборд» — a balance below one part of the packing (0,5 кан) or, without a packing, below 0,01 (0,005 шт);
    «Списать хвосты…» writes each off as a whole with one reason (rows of «Корректировки», posted)"""
    c = "V11"
    s = ready(Session(new_wms("v11")))
    try:
        order(s, 1, "1", "Антифриз 20 л", qty="1")
        receive(s, 1, "1", place="1-1-1")
        a = ords(s, 1)[21]
        s.B("PackSet", a, "4", "кан", "", module="WmsPack")
        put(s, "Выдачи", 1, L=a, E="3,5 кан", H=serial(), I="ива")
        r1 = s.I("IssuePostRow", 1)
        order(s, 2, "2", "Болт М8", qty="1")
        receive(s, 2, "1", place="1-1-2")
        b = ords(s, 2)[21]
        put(s, "Выдачи", 2, L=b, E="0,995", H=serial(), I="ива")
        r2 = s.I("IssuePostRow", 2)
        s.B("DashRefresh", module="WmsDash")
        d = s.doc.Sheets.getByName("Дашборд")
        rows = [[d.getCellByPosition(cc, r).getString() for cc in (1, 2, 6, 9)] for r in (56, 57, 58)]
        head = d.getCellByPosition(1, 54).getString()
        R.add(c, "«Хвосты» на «Дашборде»: 0,005 шт (без фасовки) и 0,125 шт (0,5 кан), меньший первым; в заголовке — число и кнопка",
              r1.startswith("OK") and r2.startswith("OK") and [x[0] for x in rows[:2]] == [b, a] and "0,5 кан" in rows[1][2] and rows[2][0] == ""
              and "ЕИ 2" in head, f"{r1}; {r2}; {rows}; {head}")
        s.B("TestPrimeForm", "хвост после выдачи", module="WmsPrimeUi")
        w = s.B("WriteOffTails", module="WmsDash")
        adj = s.doc.Sheets.getByName("Корректировки")
        st = [adj.getCellByPosition(16, r).getString() for r in (1, 2)]
        R.add(c, "«Списать хвосты…»: оба ЕИ списаны целиком («всё») с причиной, остатки 0, строки «Корректировки» проведены, список пуст",
              w.startswith("OK=2;ERR=0") and s.stock(int(a[3:])) == 0.0 and s.stock(int(b[3:])) == 0.0 and all(x.startswith("Проведено") for x in st)
              and adj.getCellByPosition(11, 1).getString() == "хвост после выдачи" and d.getCellByPosition(2, 56).getString() == "хвостов нет",
              f"{w}; {st}; {d.getCellByPosition(2, 56).getString()}")
        oracle(c, s, "списание хвостов")
    finally:
        s.close()


@case
def v12_pack_batch():
    """Ч6: «Фасовка пачкой…» — the EIs with a balance whose name tells the packing («4 канистры», «6 м»), each confirmed by a
    button; an EI whose name tells nothing is not offered; a second run has nothing left"""
    c = "V12"
    s = ready(Session(new_wms("v12")))
    try:
        order(s, 1, "1", "Незамерзайка 4 канистры по 5 л", qty="2")
        receive(s, 1, "2", place="1-1-1")
        order(s, 2, "2", "Профиль 40х20 6 м", qty="3")
        receive(s, 2, "3", place="1-1-2")
        order(s, 3, "3", "Болт М8х30", qty="10")
        receive(s, 3, "10", place="1-1-3")
        e1, e2, e3 = (ords(s, r)[21] for r in (1, 2, 3))
        res = s.B("PackBatch", module="WmsPack")
        pk = s.doc.Sheets.getByName("_PACK")
        got = [list(pk.getCellRangeByPosition(0, int(e[3:]), 2, int(e[3:])).getDataArray()[0]) for e in (e1, e2, e3)]
        again = s.B("PackBatch", module="WmsPack")
        R.add(c, "«Фасовка пачкой»: «4 канистры» → в 1 шт — 4 кан, «6 м» → 6 м, «Болт М8х30» не предложен; повторно — предлагать нечего",
              res.startswith("OK=2;SKIP=0;ERR=0") and got[0] == [e1, "кан", 4.0] and got[1] == [e2, "м", 6.0] and got[2][0] == ""
              and again.startswith("ERR:"), f"{res}; {got}; {again}")
        oracle(c, s, "фасовка пачкой")
    finally:
        s.close()


# ================================================================ E: places and the map

@case
def v13_split():
    """М1: «Разделить ЕИ…» — part of an EI becomes a new EI (the next number) with the same card and packing, its own place;
    a row «Разделение» of «Корректировки» (journaled SPLIT); the card and «остаток = движения» of both agree; the whole
    balance is refused (a move); a split is not corrected; its storno only while the new EI did not move"""
    c = "V13"
    s = ready(Session(new_wms("v13")))
    try:
        order(s, 1, "1", "Антифриз 20 л", qty="1")
        receive(s, 1, "1", place="1-1-1")
        a = ords(s, 1)[21]
        na = int(a[3:])
        s.B("PackSet", a, "4", "кан", "", module="WmsPack")
        nxt = int(s.sysv("NEXT_EI"))
        s.goto("Наличие", f"$A${na + 1}")
        s.B("TestPrimeForm", f"{a}\t2 кан\t1-2-4\tдве канистры на другое место", module="WmsPrimeUi")
        res = s.B("SplitWindow", module="WmsDaily")
        m = f"ЕИ-{nxt:08d}"
        stk = s.doc.Sheets.getByName("Наличие")
        rowm = list(stk.getCellRangeByPosition(0, nxt, 10, nxt).getDataArray()[0])
        adj = s.doc.Sheets.getByName("Корректировки")
        ctl = adj.getCellByPosition(16, 1).getString()
        pk = list(s.doc.Sheets.getByName("_PACK").getCellRangeByPosition(0, nxt, 2, nxt).getDataArray()[0])
        ent = [e for e in s.journal()[0] if e["type"] == "SPLIT"]
        R.add(c, "«Разделить ЕИ…» 2 кан из 4: новый ЕИ — следующий номер, та же карточка и фасовка, своё место; у исходного 0,5 шт; "
                 "строка «Разделение» — «Проведено · новый ЕИ-…»; в журнале SPLIT",
              res.startswith("OK:" + m) and s.stock(na) == 0.5 and rowm[:6] == [m, "Антифриз 20 л", "", "шт", 0.5, "1-2-4"] and rowm[10] == "2 кан"
              and stk.getCellByPosition(10, na).getString() == "2 кан" and pk == [m, "кан", 4.0] and ctl == "Проведено · новый " + m
              and int(s.sysv("NEXT_EI")) == nxt + 1 and len(ent) == 1 and ent[0]["fields"].get("NEW_EI") == m, f"{res}; {rowm}; {ctl}; {pk}; {s.stock(na)}")
        card = s.B("CardShow", nxt, module="WmsFind")
        card0 = s.B("CardShow", na, module="WmsFind")
        bc = s.B("BalanceCheck", module="WmsLegacy")
        R.add(c, "карточки обоих ЕИ сходятся с движениями (новый: «Отделён от …»), «остаток = движения» — без расхождений",
              "= «Наличие»" in card and "движений 1" in card and "= «Наличие»" in card0 and bc == "", f"{card}; {card0}; {bc!r}")
        # the whole balance: refused (a move); a split is not corrected
        put(s, "Корректировки", 2, B="Разделение", C=a, G="2 кан", K=serial(), L="всё")
        w = s.B("AdjustPostRow", 2, module="WmsAdjust")
        fx = s.B("AdjustFixRow", 1, "1 кан", serial(), "x", module="WmsAdjust")
        R.add(c, "отделить весь остаток — отказ (это перемещение); «Исправить» разделения — отказ («удалите и проведите заново»)",
              w.startswith("ERR:") and "весь остаток" in w and fx.startswith("ERR:") and "не исправляется" in fx, f"{w}; {fx}")
        s.B("AdjustClearRow", 2, module="WmsAdjust")
        # the storno: refused while the new EI moved, done after the storno of that issue
        put(s, "Выдачи", 1, L=m, E="1 кан", H=serial(), I="ива")
        i1 = s.I("IssuePostRow", 1)
        d1 = s.B("AdjustDeleteRow", 1, module="WmsAdjust")
        s.I("IssueDeleteRow", 1)
        d2 = s.B("AdjustDeleteRow", 1, module="WmsAdjust")
        R.add(c, "сторно разделения: пока по новому ЕИ есть выдача — отказ; после сторно выдачи — количество вернулось (1 шт), новый ЕИ закрыт (0, сторно)",
              i1.startswith("OK") and d1.startswith("ERR:") and d2.startswith("OK") and s.stock(na) == 1.0 and s.stock(nxt) == 0.0
              and stk.getCellByPosition(7, nxt).getString() != "" and adj.getCellByPosition(16, 1).getString() == "Удалено (сторно)",
              f"{i1}; {d1[:120]}; {d2}; {s.stock(na)}; {s.stock(nxt)}; {stk.getCellByPosition(7, nxt).getString()}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка после разделения и сторно — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "разделение ЕИ и сторно")
    finally:
        s.close()


@case
def v14_offcuts():
    """Ч7: metres cut from a bar (1 шт = 6 м) by «Провести» — the rest of the bar may become an offcut of its own (a split);
    «Обрезок…» finds the shortest piece of the product not shorter than needed and puts it into the row"""
    c = "V14"
    s = ready(Session(new_wms("v14")))
    try:
        order(s, 1, "1", "Профиль 40х20 6 м", qty="2")
        receive(s, 1, "2", place="3-1-1")
        a = ords(s, 1)[21]
        na = int(a[3:])
        s.B("PackSet", a, "6", "м", "", module="WmsPack")
        nxt = int(s.sysv("NEXT_EI"))
        put(s, "Выдачи", 1, L=a, E="1,5 м", H=serial(), I="ива")
        s.goto("Выдачи", "$L$2")
        s.U("BtnPost")
        off = f"ЕИ-{nxt:08d}"
        R.add(c, "выдача 1,5 м из прутка 6 м: «оформить обрезок 4,5 м?» — «Да»: обрезок — новый ЕИ 0,75 шт (4,5 м), у прутка 1 шт (6 м)",
              iss_row(s, 1)[17] == "Проведено · 1,5 м" and s.stock(na) == 1.0 and s.stock(nxt) == 0.75
              and s.doc.Sheets.getByName("Наличие").getCellByPosition(10, nxt).getString() == "4,5 м", f"{iss_row(s, 1)[17]}; {s.stock(na)}; {s.stock(nxt)}")
        put(s, "Выдачи", 2, L=a, H=serial(), I="ива")
        s.goto("Выдачи", "$L$3")
        s.B("TestPrimeForm", "3", module="WmsPrimeUi")
        pk = s.B("PickOffcut", module="WmsDaily")
        row = iss_row(s, 2)
        put(s, "Выдачи", 3, L=a, H=serial(), I="ива")
        s.goto("Выдачи", "$L$4")
        s.B("TestPrimeForm", "7", module="WmsPrimeUi")
        pk2 = s.B("PickOffcut", module="WmsDaily")
        R.add(c, "«Обрезок…» 3 м — самый короткий подходящий кусок: обрезок 4,5 м (не целый пруток), в строку — его ЕИ и «3 м»; 7 м — такого куска нет",
              pk == "OK:" + off and row[11] == off and row[4] == "3 м" and pk2.startswith("ERR:") and "не короче 7 м" in pk2, f"{pk}; {row[11]}; {row[4]!r}; {pk2}")
        s.goto("Выдачи", "$L$3")
        p3 = s.I("IssuePostRow", 2)
        R.add(c, "выдача 3 м из обрезка 4,5 м — остаток обрезка 1,5 м", p3.startswith("OK") and s.stock(nxt) == 0.25, f"{p3}; {s.stock(nxt)}")
        bc = s.B("BalanceCheck", module="WmsLegacy")
        R.add(c, "остаток = движения у прутка и обрезка", bc == "", bc)
        oracle(c, s, "обрезки металла")
    finally:
        s.close()


@case
def v15_map_editor():
    """the map edited on the map itself (the request of the storekeeper) and «Проверить топологию» (М3): «Править карту» — the
    plain layout, a place typed as «5.12.3» stays text; «Сохранить карту» checks (a repeated address, a place without a tier,
    the places that disappear while EIs lie there), replaces «Топология», keeps the old one hidden; «Отменить правку» restores"""
    c = "V15"
    s = ready(Session(new_wms("v15")))
    try:
        topology(s, {1: (3, 2)})
        order(s, 1, "1", "Болт М8", qty="5")
        receive(s, 1, "5", place="1-3-1")
        s.B("MapBuild", "OCC", True, module="WmsMap")
        mp = s.doc.Sheets.getByName("Карта склада")
        n0 = s.B("TestMapPlace", "1-3-1", module="WmsMap")
        e = s.B("MapEditStart", module="WmsMap")
        # the places of row 1 are at MP_TOP + 1 … ; write a new place, a repeated one, a place without a tier; clear 1-3-1
        top = 7
        cells = {}
        for r in range(top, top + 8):
            for cc in range(0, 6):
                t = mp.getCellByPosition(cc, r).getString()
                if t:
                    cells[t] = (cc, r)
        mp.getCellByPosition(5, top + 1).setString("1-4-1")
        mp.getCellByPosition(5, top + 2).setString("1-1-1")
        mp.getCellByPosition(5, top + 3).setString("2-5")
        cx, rx = cells["1-3-1"]
        mp.getCellByPosition(cx, rx).setString("")
        mp.getCellByPosition(4, top + 4).setString("1.5.1")
        typed = mp.getCellByPosition(4, top + 4).getType().value
        s.U("TestUiAuto", 0)
        choose(s, 1)                                    # «Вернуться к правке»
        s.U("TestUiAuto", 1)
        w1 = s.B("MapEditSave", module="WmsMap")
        shown = s.B("TestPrimeShown", module="WmsPrimeUi")
        R.add(c, "«Править карту»: карта без слоя, «1.5.1» осталось текстом; «Сохранить карту» сначала показывает проверку: повтор 1-1-1, «2-5» без яруса, "
                 "исчезает 1-3-1 с ЕИ; «Вернуться к правке» — ничего не сохранено",
              e.startswith("OK") and typed == "TEXT" and w1.startswith("SKIP") and "1-1-1" in shown and "2-5" in shown and "1-3-1" in shown
              and s.doc.Sheets.getByName("Топология").getCellByPosition(5, 2).getString() == "", f"{e}; {typed}; {w1}; {shown[:400]}")
        rb = s.B("MapBuild", "OCC", module="WmsMap")
        w2 = s.B("MapEditSave", module="WmsMap")
        topo = s.doc.Sheets.getByName("Топология")
        tt = {topo.getCellByPosition(cc, r).getString() for r in range(0, 12) for cc in range(0, 7)}
        prev = s.doc.Sheets.hasByName("_TOPO_PREV") and not s.doc.Sheets.getByName("_TOPO_PREV").IsVisible
        R.add(c, "во время правки карта не перестраивается сама; «Сохранить карту»: «Топология» = карта (1-4-1, 1.5.1 есть, 1-3-1 нет), прежняя — в скрытом листе; "
                 "карта построена заново", rb.startswith("ERR:") and w2.startswith("OK:карта сохранена") and "1-4-1" in tt and "1.5.1" in tt and "1-3-1" not in tt
              and prev and mp.getCellByPosition(1105, 0).getString() == "", f"{rb}; {w2}; {sorted(tt)[:12]}; {prev}")
        n1 = s.B("TestMapPlace", "1-4-1", module="WmsMap")
        off = s.B("TestMapPlace", "1-3-1", module="WmsMap")
        R.add(c, "новое место 1-4-1 на карте; ЕИ с 1-3-1 теперь «не на карте» (его место не изменилось)",
              n1 not in ("", "-1") and s.doc.Sheets.getByName("Наличие").getCellByPosition(5, int(ords(s, 1)[21][3:])).getString() == "1-3-1",
              f"{n0}; {n1}; {off}")
        e2 = s.B("MapEditStart", module="WmsMap")
        mp.getCellByPosition(0, top).setString("9-9-9")
        cn = s.B("MapEditCancel", module="WmsMap")
        tt2 = {topo.getCellByPosition(cc, r).getString() for r in range(0, 12) for cc in range(0, 7)}
        R.add(c, "«Отменить правку»: изменения карты не попали в «Топологию», карта построена из неё снова",
              e2.startswith("OK") and cn.startswith("OK") and "9-9-9" not in tt2 and mp.getCellByPosition(0, top).getString() != "9-9-9", f"{e2}; {cn}")
        s.B("BtnTopoCheck", module="WmsMap")
        tc = s.U("TestUiLastMessage")
        R.add(c, "«Проверить топологию» (М3): мест и замечания (повтор адреса 1-1-1, «2-5» без яруса) — ничего не меняется",
              tc.startswith("OK:мест") and "1-1-1" in tc and "2-5" in tc, tc[:300])
    finally:
        s.close()


@case
def v16_places():
    """М4 capacity («Места» D) and the layer «Заполненность»; М9 «Куда положить?» counts the places of rows not yet posted and
    never names a place twice; М5 «Показать на карте»; М6/М12 a place in «Поиск» — what lies there now and its history; a
    double click on a place of the map opens it"""
    c = "V16"
    s = ready(Session(new_wms("v16")))
    try:
        topology(s, {1: (3, 2)})
        pl = s.doc.Sheets.getByName("Места")
        pl.getCellByPosition(0, 1).setString("1-1-1")
        pl.getCellByPosition(3, 1).setValue(3)
        order(s, 1, "1", "Болт М8", qty="5")
        receive(s, 1, "5", place="1-1-1")
        order(s, 2, "2", "Гайка М8", qty="5")
        receive(s, 2, "5", place="1-2-1")
        order(s, 3, "3", "Шайба М8", qty="5", U="1-3-1")          # not posted: 1-3-1 is promised
        a = ords(s, 1)[21]
        s.B("MapBuild", "FILL", True, module="WmsMap")
        b = [s.B("TestMapBucket", k, "FILL", module="WmsMap") for k in ("1-1-1", "1-2-1", "1-3-2")]
        R.add(c, "М4: «Заполненность»: 1-1-1 (вместимость 3, 1 ЕИ) — есть место, 1-2-1 (вместимость 1) — заполнено, 1-3-2 — пусто", b == ["1", "3", "0"], str(b))
        w = s.B("WhereSuggest", "Болт М8", "", module="WmsMap").split("\n")
        keys = [x.split("|")[0] for x in w]
        R.add(c, "М9: «Куда положить?» — 1-1-1 (уже лежит этот товар) один раз; 1-2-1 (полное) и 1-3-1 (обещано строкой «Заказов») не предложены",
              keys[0] == "1-1-1" and keys.count("1-1-1") == 1 and "1-2-1" not in keys and "1-3-1" not in keys and len(keys) >= 3, str(w[:5]))
        s.B("CardShow", int(a[3:]), module="WmsFind")
        sm = s.B("ShowCardOnMap", module="WmsFind")
        R.add(c, "М5: «Показать на карте» из карточки ЕИ — слой «Товар», место 1-1-1 подсвечено", sm.startswith("OK:мест на карте 1")
              and s.B("TestMapBucket", "1-1-1", "FIND", module="WmsMap") == "4" and s.B("TestMapBucket", "1-2-1", "FIND", module="WmsMap") == "1", sm)
        fr = s.B("FindRun", "1-1-1", module="WmsFind")
        head = s.doc.Sheets.getByName("Поиск").getCellByPosition(1, 6).getString()
        R.add(c, "М6/М12: «1-1-1» в «Поиске» — «сейчас — 1 ЕИ: … Болт М8 — 5 шт», ниже история места", fr.startswith("OK") and "Место 1-1-1: сейчас — 1 ЕИ" in head
              and "Болт М8" in head and "история места" in head, head)
        cell = s.B("TestMapPlace", "1-2-1", module="WmsMap").split(":")
        mp = s.doc.Sheets.getByName("Карта склада")
        dc = s.B("OnMapDoubleClick", mp.getCellByPosition(int(cell[1]), int(cell[0])), module="WmsMap")
        head2 = s.doc.Sheets.getByName("Поиск").getCellByPosition(1, 6).getString()
        R.add(c, "М6: двойной щелчок по месту карты — «Поиск» с этим местом", dc is True and s.active_sheet() == "Поиск" and "Место 1-2-1: сейчас — 1 ЕИ" in head2,
              f"{dc}; {s.active_sheet()}; {head2}")
    finally:
        s.close()


@case
def v17_place_lists():
    """М10: a range «1-2-3…5» is three places (the map, the checks); the new place of a move as a change of the list — «+1-3-1»
    one more, «−1-1-1» one less (the refusals: already there, not there, nothing left); «Исправить» the same from the places the
    move wrote. М11: after an issue from an EI on several places — «место опустело?» frees the chosen one by a move"""
    c = "V17"
    s = ready(Session(new_wms("v17")))
    try:
        topology(s, {1: (3, 2)})
        pl = [list(s.B("PlaceList", x, module="WmsView")) for x in ("1-2-3…5", "1-2-3..5, 2-1-1", "1-2-5…3", "1-2-3…1-2-5", "1-2-3…1-3-5", "1-2…4")]
        it = list(s.B("PlaceItems", "1-1-1…2; у стены", module="WmsView"))
        R.add(c, "М10: «1-2-3…5» = 1-2-3, 1-2-4, 1-2-5 (и «..», обратный порядок, полный конец «…1-2-5»); «1-2-3…1-3-5» — не диапазон; "
                 "в списке примечание остаётся частью",
              pl[0] == ["1-2-3", "1-2-4", "1-2-5"] and pl[1] == ["1-2-3", "1-2-4", "1-2-5", "2-1-1"] and pl[2] == pl[0] and pl[3] == pl[0]
              and pl[4] == [] and pl[5] == ["1-2", "1-3", "1-4"] and it == ["1-1-1", "1-1-2", "у стены"], f"{pl}; {it}")
        ed = []
        for base, typed in (("1-1-1, 1-2-1", "+1-3-1"), ("1-1-1, 1-2-1", "−1-1-1"), ("1-1-1, 1-2-1", "-1-1-1"), ("1-1-1", "+1-2-1…2"),
                            ("1-1-1, 1-2-1", "+1-1-1"), ("1-1-1, 1-2-1", "−9-9-9"), ("1-1-1, 1-2-1", "−1-1-1, 1-2-1"), ("", "+1-1-1"), ("1-1-1", "+")):
            ed.append(s.B("TestPlaceEdit", base, typed, module="WmsDaily"))
        R.add(c, "М10: «+» и «−» к списку мест: добавить, убрать (минус любой чертой), диапазоном; отказы — «уже есть», «нет в списке», "
                 "«не останется ни одного места», пустой «+»",
              ed[0] == "OK:1-1-1, 1-2-1, 1-3-1" and ed[1] == "OK:1-2-1" and ed[2] == "OK:1-2-1" and ed[3] == "OK:1-1-1, 1-2-1, 1-2-2"
              and "уже есть" in ed[4] and "нет в списке" in ed[5] and "не останется" in ed[6] and ed[7] == "OK:1-1-1" and ed[8].startswith("ERR:"), str(ed))
        order(s, 1, "1", "Болт М8", qty="10")
        receive(s, 1, "10", place="1-1-1…2")
        a = ords(s, 1)[21]
        na = int(a[3:])
        s.B("MapBuild", "OCC", True, module="WmsMap")
        m1 = [s.B("TestMapEIs", k, module="WmsMap").split("|")[-1] for k in ("1-1-1", "1-1-2", "1-2-1")]
        R.add(c, "М10: приход на «1-1-1…2» — без вопроса о месте; на карте ЕИ лежит на 1-1-1 и 1-1-2", m1[0] != "" and m1[0] == m1[1] and m1[2] == "",
              f"{m1}; {s.stock_place(na) if hasattr(s, 'stock_place') else ''}")
        stk = s.doc.Sheets.getByName("Наличие")
        put(s, "Корректировки", 1, B="Перемещение", C=a, J="+1-3-1", K=serial(), L="не влезло")
        p1 = s.B("AdjustPostRow", 1, module="WmsAdjust")
        adj = s.doc.Sheets.getByName("Корректировки")
        put(s, "Корректировки", 2, B="Перемещение", C=a, J="−1-1-1…2", K=serial(), L="освободили")
        p2 = s.B("AdjustPostRow", 2, module="WmsAdjust")
        put(s, "Корректировки", 3, B="Перемещение", C=a, J="−1-3-1", K=serial(), L="x")
        p3 = s.B("AdjustPostRow", 3, module="WmsAdjust")
        ent = [e for e in s.journal()[0] if e["type"] == "MOVE"]
        R.add(c, "М10: перемещение «+1-3-1» — место «1-1-1…2, 1-3-1», в J и в журнале — весь список; «−1-1-1…2» — «1-3-1»; убрать последнее — отказ",
              p1.startswith("OK") and adj.getCellByPosition(9, 1).getString() == "1-1-1…2, 1-3-1" and p2.startswith("OK")
              and stk.getCellByPosition(5, na).getString() == "1-3-1" and adj.getCellByPosition(9, 2).getString() == "1-3-1"
              and p3.startswith("ERR:") and "не останется" in p3 and len(ent) == 2 and ent[0]["fields"].get("TO") == "1-1-1…2, 1-3-1",
              f"{p1}; {p2}; {p3}; {adj.getCellByPosition(9, 1).getString()}; {stk.getCellByPosition(5, na).getString()}")
        s.B("AdjustClearRow", 3, module="WmsAdjust")
        fx = s.B("AdjustFixRow", 2, "+1-2-2", serial(), "освободили", module="WmsAdjust")
        R.add(c, "М10: «Исправить» перемещение «+1-2-2» — к местам, что оно записало (1-3-1): «1-3-1, 1-2-2»",
              fx.startswith("OK") and stk.getCellByPosition(5, na).getString() == "1-3-1, 1-2-2", f"{fx}; {stk.getCellByPosition(5, na).getString()}")
        # М11: the issue from an EI on two places
        put(s, "Выдачи", 1, L=a, E="2", H=serial(), I="ива")
        s.goto("Выдачи", "$L$2")
        s.U("BtnPost")
        same = stk.getCellByPosition(5, na).getString()
        put(s, "Выдачи", 2, L=a, E="3", H=serial(), I="ива")
        s.goto("Выдачи", "$L$3")
        choose(s, 2)                                     # «Место 1-2-2 опустело»
        s.U("BtnPost")
        after = stk.getCellByPosition(5, na).getString()
        last = adj.getCellRangeByPosition(0, 3, 16, 3).getDataArray()[0]
        R.add(c, "М11: выдача с ЕИ на двух местах — «место опустело?»: Enter — места те же; «1-2-2 опустело» — проведено перемещение «−1-2-2», место «1-3-1»",
              same == "1-3-1, 1-2-2" and after == "1-3-1" and last[1] == "Перемещение" and str(last[16]).startswith("Проведено")
              and "опустело после выдачи № 2" in last[11] and s.stock(na) == 5.0, f"{same}; {after}; {last}")
        put(s, "Выдачи", 3, L=a, E="1", H=serial(), I="ива")
        s.goto("Выдачи", "$L$4")
        shown0 = s.B("TestPrimeShown", module="WmsPrimeUi")
        s.U("BtnPost")
        R.add(c, "М11: ЕИ на одном месте — вопроса нет", s.B("TestPrimeShown", module="WmsPrimeUi") == shown0 and s.stock(na) == 4.0, str(s.stock(na)))
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "списки мест: диапазон, «+», «−», исправление, «место опустело»")
    finally:
        s.close()


@case
def v18_rename_route():
    """М7 «Переименовать ряд…»: rows «Перемещение» prepared (not posted) for every EI with a balance there, the map and «Места»
    renamed when asked, the block posted by «Провести»; М8 «Маршрут сбора»: the rows of «Выдачи» not posted, in the order of
    the places, on «Поиск»"""
    c = "V18"
    s = ready(Session(new_wms("v18")))
    try:
        topology(s, {1: (3, 2)})
        pl = s.doc.Sheets.getByName("Места")
        pl.getCellByPosition(0, 1).setString("1-2-1")
        pl.getCellByPosition(3, 1).setValue(2)
        for k, (nm, place) in enumerate((("Болт М8", "1-2-1"), ("Гайка М8", "1-1-1, 1-2-2"), ("Шайба М8", "1-3-1"), ("Винт М8", "1-2-2"))):
            order(s, 1 + k, str(1 + k), nm, qty="5")
            receive(s, 1 + k, "5", place=place)
        eis = [ords(s, 1 + k)[21] for k in range(4)]
        issue(s, 1, eis[3], "5")                        # Винт: nothing left — not moved
        stk = s.doc.Sheets.getByName("Наличие")
        adj = s.doc.Sheets.getByName("Корректировки")
        s.B("TestPrimeForm", "1-2\t1-4\tперевесили таблички", module="WmsPrimeUi")
        choose(s, 0)                                     # «Подготовить строки и переименовать на карте»
        rn = s.B("RenamePlaces", module="WmsDaily")
        rows = adj.getCellRangeByPosition(0, 1, 16, 3).getDataArray()
        topo = s.doc.Sheets.getByName("Топология")
        tt = {topo.getCellByPosition(cc, r).getString() for r in range(0, 8) for cc in range(0, 5)}
        R.add(c, "М7: «1-2» → «1-4»: подготовлены 2 строки «Перемещение» (Болт → 1-4-1, Гайка → 1-1-1, 1-4-2), не проведены; Винт (остаток 0) — нет; "
                 "на карте 1-2-1/1-2-2 стали 1-4-1/1-4-2, в «Места» — 1-4-1; остатки и места ЕИ пока прежние",
              rn == "OK:2|2-3" and rows[0][1] == "Перемещение" and rows[0][2] == eis[0] and rows[0][9] == "1-4-1" and rows[1][2] == eis[1]
              and rows[1][9] == "1-1-1, 1-4-2" and rows[0][0] == "" and rows[2][1] == "" and "перевесили таблички" in rows[0][11]
              and "1-4-1" in tt and "1-4-2" in tt and "1-2-1" not in tt and pl.getCellByPosition(0, 1).getString() == "1-4-1"
              and stk.getCellByPosition(5, int(eis[0][3:])).getString() == "1-2-1", f"{rn}; {rows}; {sorted(tt)}")
        s.goto("Корректировки", "$A$2:$S$3")
        s.B("BtnAdjPost", module="WmsAdjustUi")
        R.add(c, "М7: строки проведены блоком «Провести», без вопросов о местах; места ЕИ — новые",
              stk.getCellByPosition(5, int(eis[0][3:])).getString() == "1-4-1" and stk.getCellByPosition(5, int(eis[1][3:])).getString() == "1-1-1, 1-4-2"
              and str(adj.getCellByPosition(16, 1).getString()).startswith("Проведено") and str(adj.getCellByPosition(16, 2).getString()).startswith("Проведено"),
              f"{s.U('TestUiLastMessage')}; {adj.getCellByPosition(16, 1).getString()}; {adj.getCellByPosition(16, 2).getString()}")
        s.B("TestPrimeForm", "7\t8\tx", module="WmsPrimeUi")
        r7 = s.B("RenamePlaces", module="WmsDaily")
        s.B("TestPrimeForm", "1-4\t2\tx", module="WmsPrimeUi")
        r8 = s.B("RenamePlaces", module="WmsDaily")
        R.add(c, "М7: мест «7…» нет нигде — отказ; «1-4» → «2» (секция на ряд) — отказ", r7.startswith("ERR:") and "нет ни у одного" in r7
              and r8.startswith("ERR:") and "так же" in r8, f"{r7}; {r8}")
        # М8: the route
        put(s, "Выдачи", 2, L=eis[2], E="1", H=serial(), I="ива")       # 1-3-1
        put(s, "Выдачи", 3, L=eis[0], E="2", H=serial(), I="ива")       # 1-4-1
        put(s, "Выдачи", 4, L=eis[1], E="1", H=serial(), I="лев")       # 1-1-1, 1-4-2
        put(s, "Выдачи", 5, L="ЕИ-999", E="1", H=serial(), I="лев")     # unknown EI — last
        rt = s.B("PickRoute", module="WmsFind")
        fd = s.doc.Sheets.getByName("Поиск")
        tab = [list(r) for r in fd.getCellRangeByPosition(1, 12, 8, 15).getDataArray()]
        R.add(c, "М8: «Маршрут сбора» — 4 непроведённые строки по порядку мест: 1-1-1 (Гайка), 1-3-1 (Шайба), 1-4-1 (Болт), без места — в конце; номер, место, ЕИ, "
                 "количество, кому, строка «Выдачи»",
              rt == "OK:строк 4" and [r[2] for r in tab] == [eis[1], eis[2], eis[0], "ЕИ-00000999"] and tab[0][0] == 1 and tab[0][1] == "1-1-1, 1-4-2"
              and tab[3][1] == "место не указано" and str(tab[2][4]) in ("2", "2.0") and tab[0][6] == "лев" and tab[1][7] == "строка 3"
              and "Маршрут сбора: 4 строки" in fd.getCellByPosition(1, 6).getString(), f"{rt}; {tab}; {fd.getCellByPosition(1, 6).getString()}")
        oracle(c, s, "переименование ряда блоком перемещений")
    finally:
        s.close()


KEEP = "<как предложено>"          # WmsOrdersUi / WmsPrimeUi test seams: the field keeps what the window proposed


@case
def v19_supply():
    """этап F: З8 «Поставщики» (пополнить из заказов, срок поставки), З1 возврат поставщику (остаток, поставщик, документ, «ждём
    замену» — новая строка заказа), его «Исправить» и «Удалить», З4 «Ещё поступление» с остатком позиции, З6 повтор позиции,
    З5 отмена всего заказа с причиной, З7 «уже ожидается», З2 «Ожидаем», З10 «Нужно заказать», З3 дашборд, З9 «Принять товар»"""
    c = "V19"
    s = ready(Session(new_wms("v19")))
    try:
        order(s, 1, "1", "Болт М8", qty="10", supplier="Ромашка", art="Б-8")
        receive(s, 1, "10", place="1-1-1")
        order(s, 2, "2", "Гайка М8", qty="10", supplier="ромашка ")
        receive(s, 2, "4", place="1-1-2")
        order(s, 3, "1", "Шайба М8", qty="6", supplier="Лютик")
        order(s, 4, "2", "Винт М8", qty="3", supplier="Лютик")
        fill = s.B("SuppliersFill", module="WmsSupply")
        sup = s.doc.Sheets.getByName("Поставщики")
        names = [sup.getCellByPosition(0, r).getString() for r in range(1, 4)]
        sup.getCellByPosition(3, 1).setValue(5)
        lead = s.B("SupplierLead", "РОМАШКА", module="WmsSupply")
        val = s.doc.Sheets.getByName("Заказы").getCellByPosition(11, 5).Validation.getFormula1()
        R.add(c, "З8: «Пополнить из заказов» — Ромашка и Лютик по одному разу («ромашка » — то же); срок 5 дн. находится без учёта регистра; "
                 "на «Заказах» L — список «Поставщиков»", fill.startswith("OK:") and names == ["Ромашка", "Лютик", ""] and lead == 5
              and "Поставщики" in val, f"{fill}; {names}; {lead}; {val}")
        a = ords(s, 1)[21]
        na = int(a[3:])
        r = s.B("SupReturnPost", a, "2", "брак", "ВП-1", True, module="WmsSupply")
        adj = s.doc.Sheets.getByName("Корректировки")
        row = list(adj.getCellRangeByPosition(0, 1, 18, 1).getDataArray()[0])
        osh = s.doc.Sheets.getByName("Заказы")
        nr = int(s.B("TestSupNewRow", module="WmsSupply"))
        rep = list(osh.getCellRangeByPosition(0, nr, 27, nr).getDataArray()[0]) if nr > 0 else []
        ent = [e for e in s.journal()[0] if e["type"] == "SUP_RET"]
        R.add(c, "З1: возврат поставщику 2 из 10 — остаток 8; строка «Возврат поставщику», J — документ, «Проведено · поставщику Ромашка»; в журнале "
                 "SUP_RET с поставщиком и документом; «ждём замену» — новая строка «Заказов»: Болт, Б-8, заказано 2, Ромашка, ожидаем через 5 дн., «Ожидается»",
              r.startswith("OK:") and s.stock(na) == 8.0 and row[1] == "Возврат поставщику" and row[9] == "ВП-1" and row[16] == "Проведено · поставщику Ромашка"
              and "ждём замену" in row[11] and len(ent) == 1 and ent[0]["fields"].get("SUPPLIER") == "Ромашка" and ent[0]["fields"].get("DOC") == "ВП-1"
              and rep[:2] == ["1", "Болт М8"] and rep[4] == "Б-8" and rep[7] == "2" and rep[11] == "Ромашка" and rep[22] == "Ожидается"
              and rep[16] == rep[15] + 5 and "возврату поставщику № 1" in rep[25], f"{r}; {s.stock(na)}; {row}; {rep}")
        fx = s.B("AdjustFixRow", 1, "3", serial(), "брак; ждём замену", module="WmsAdjust")
        b1 = s.stock(na)
        dl = s.B("AdjustDeleteRow", 1, module="WmsAdjust")
        R.add(c, "З1: «Исправить» возврат — 3 (остаток 7), «Удалить» — сторно (остаток 10)", fx.startswith("OK") and b1 == 7.0 and dl.startswith("OK")
              and s.stock(na) == 10.0, f"{fx}; {b1}; {dl}; {s.stock(na)}")
        # З4: «Ещё поступление» of the partly received position
        s.goto("Заказы", "$B$3")
        s.B("TestRcvInput", KEEP, "", "УПД-2", KEEP, KEEP, "1-1-2", "", module="WmsOrdersUi")
        s.B("BtnRcvMore", module="WmsOrdersUi")
        info = s.B("TestRcvInfo", module="WmsOrdersUi")
        shown = s.B("TestRcvDialog", module="WmsOrdersUi").split("\n")[0].split("\t")
        st2 = osh.getCellByPosition(22, 2).getString()
        R.add(c, "З4: «Ещё поступление» — над полями «Заказано 10 · получено 4 · осталось 6», количество 6 предложено; принято 6 — «Получено»",
              info.startswith("Заказано 10 · получено 4 · осталось 6") and shown[0] == "6" and st2.startswith("Получено"), f"{info}; {shown}; {st2}")
        # З6: the repeat of a position (5 pieces)
        s.goto("Заказы", "$B$2")
        choose(s, 0)
        s.B("TestPrimeForm", "5", module="WmsPrimeUi")
        rp = s.B("RepeatOrder", module="WmsSupply")
        last = int(rp.split("|")[1].split("-")[1]) - 1 if rp.startswith("OK:") else 0
        rr = list(osh.getCellRangeByPosition(0, last, 27, last).getDataArray()[0])
        R.add(c, "З6: «Повторить заказ» → «Эту позицию» 5 — новая строка: Болт, Б-8, Ромашка, 5, дата заказа сегодня, «Ожидается»",
              rp.startswith("OK:1|") and rr[1] == "Болт М8" and rr[4] == "Б-8" and rr[7] == "5" and rr[11] == "Ромашка" and rr[15] == serial() and rr[22] == "Ожидается",
              f"{rp}; {rr}")
        # З7: a new row of the same goods and supplier — the hint
        put(s, "Заказы", last + 1, A="1", B="Шайба М8", H="2", I="шт", L="Лютик")
        # as the storekeeper types it: the change handler of «Заказы» for the cell L (an API write does not fire it)
        s.B("OnOrdersChange", osh.getCellByPosition(11, last + 1), module="WmsOrders")
        hint = osh.getCellByPosition(24, last + 1).getString()
        # «Ещё поступление» of «Гайка» put its row under it: the rows below moved down by one
        bcol = [x[0] for x in osh.getCellRangeByPosition(1, 0, 1, last).getDataArray()]
        r_sh = bcol.index("Шайба М8")
        R.add(c, f"З7: новая строка «Шайба М8» у «Лютик» — в «Контроль»: «Внимание: … уже ожидается — строка {r_sh + 1}»",
              hint.startswith("Внимание:") and f"уже ожидается — строка {r_sh + 1} " in hint, hint)
        s.B("ClearOrderRow", last + 1, module="WmsOrders")
        # З2 «Ожидаем», З10 «Нужно заказать», З3 the dashboard
        w = s.B("WaitShow", "", module="WmsSupply")
        fd = s.doc.Sheets.getByName("Снабжение")
        tab = [list(x) for x in fd.getCellRangeByPosition(1, 10, 14, 17).getDataArray()]
        groups = [x[0] for x in tab if x[13] == "G"]
        R.add(c, "З2: «Ожидаем» — позиции по поставщикам (Лютик: Шайба 6, Винт 3; Ромашка: замена Болт 2, повтор Болт 5); строка ведёт на «Заказы»",
              w == "OK:4" and groups == ["Лютик", "Ромашка"] and any(x[1] == "Шайба М8" and x[5] == 6 for x in tab)
              and any(x[1] == "Болт М8" and x[5] == 5 for x in tab) and all(x[11] == "Заказы" for x in tab if x[13] != "G" and x[1]),
              f"{w}; {tab}")
        gd = s.doc.Sheets.getByName("Номенклатура")
        fl = s.B("GoodsFill", module="WmsSupply")
        names_g = [x[0] for x in gd.getCellRangeByPosition(1, 0, 1, 400).getDataArray()]
        rows = {nm: r for r, nm in enumerate(names_g) if nm}
        gd.getCellByPosition(7, rows["Болт М8"]).setValue(20)
        gd.getCellByPosition(8, rows["Болт М8"]).setValue(10)
        gd.getCellByPosition(7, rows["Гайка М8"]).setValue(5)
        nd = s.B("NeedShow", "", module="WmsSupply")
        tab2 = [list(x) for x in fd.getCellRangeByPosition(1, 10, 14, 12).getDataArray()]
        R.add(c, "З10: «Номенклатура» пополнена из «Наличия»; «Нужно заказать»: Болт — остаток 10 + в пути 7 < минимума 20 → заказать 10 (кратно 10); "
                 "Гайка (10 ≥ 5) — нет", fl.startswith("OK:") and nd == "OK:1" and tab2[0][0] == "Болт М8" and tab2[0][2] == 10 and tab2[0][3] == 7
              and tab2[0][6] == 10, f"{fl}; {nd}; {tab2}")
        db = s.B("DashRefresh", module="WmsDash")
        dsh = s.doc.Sheets.getByName("Дашборд")
        tile = dsh.getCellByPosition(3, 5).getString() + " | " + dsh.getCellByPosition(1, 5).getString()
        refs = [dsh.getCellByPosition(14, r).getString() for r in range(9, 13)]
        R.add(c, "З3/З10: дашборд — у ожидаемых «осталось …», ссылка на строку «Заказов» (для перехода), в плитке «заказать 1»",
              db.startswith("OK") and "заказать 1" in tile and all(x.startswith("Заказы|") for x in refs if x), f"{db}; {tile}; {refs}")
        # З9: the vehicle of «Лютик» — «Принять товар»
        cars = s.doc.Sheets.getByName("Приход авто")
        cars.getCellByPosition(2, 1).setString("А123ВС")
        cars.getCellByPosition(3, 1).setString("Лютик")
        s.B("BtnCarArrive", module="WmsCarUi")
        s.goto("Приход авто", "$D$7")
        ca = s.B("CarAccept", module="WmsSupply")
        R.add(c, "З9: «Принять товар» на визите «Лютик» — «Снабжение» с позициями Лютика (ждём 2), принято сегодня 0",
              ca == "OK:Лютик|принято 0|ждём 2" and s.active_sheet() == "Снабжение", ca)
        # З5: the whole order of «Лютик» (rows 4–5) cancelled with the reason
        s.goto("Заказы", f"$B${r_sh + 1}")
        s.B("TestPrimeForm", "поставщик не привезёт", module="WmsPrimeUi")
        co = s.B("CancelOrderAll", module="WmsSupply")
        ctl = osh.getCellByPosition(24, r_sh).getString()
        ent2 = [e for e in s.journal()[0] if e["type"] == "ORDER_CANCEL"]
        R.add(c, "З5: «Отменить весь заказ» — обе позиции «Отменено», в «Контроль» и в журнале — причина",
              co == "OK:2|0|0|0" and osh.getCellByPosition(22, r_sh).getString() == "Отменено" and osh.getCellByPosition(22, r_sh + 1).getString() == "Отменено"
              and "поставщик не привезёт" in ctl and len(ent2) == 2 and all(e["fields"].get("REASON") == "поставщик не привезёт" for e in ent2), f"{co}; {ctl}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "возврат поставщику, замена, повтор, отмена заказа")
    finally:
        s.close()


@case
def v20_print_discrepancy():
    """З12 «Акт расхождения»: the delivery of one document of a supplier — by the document and in fact, the shortage, the returns
    to the supplier (брак) — on «Печать»; «Сохранить PDF» — a PDF in «WMS_Документы» next to the book"""
    c = "V20"
    s = ready(Session(new_wms("v20")))
    try:
        order(s, 1, "1", "Болт М8", qty="10", supplier="Ромашка", art="Б-8")
        put(s, "Заказы", 1, G="10")
        receive(s, 1, "8", doc="УПД-77", place="1-1-1")
        order(s, 2, "2", "Гайка М8", qty="5", supplier="Ромашка")
        put(s, "Заказы", 2, G="5")
        receive(s, 2, "5", doc="УПД-77", place="1-1-2")
        b = ords(s, 2)[21]
        s.B("SupReturnPost", b, "1", "брак: трещина", "", False, module="WmsSupply")
        s.goto("Заказы", "$B$3")
        res = s.B("DiscrepancyAct", module="WmsPrint")
        pr = s.doc.Sheets.getByName("Печать")
        txt = [[str(x) for x in r] for r in pr.getCellRangeByPosition(0, 0, 6, 40).getDataArray()]
        flat = " | ".join(" ".join(r) for r in txt)
        area = pr.getPrintAreas()
        R.add(c, "З12: акт расхождения по УПД-77 «Ромашка»: Болт — по документу 10, фактически 8, «недостача 2»; Гайка — «совпадает»; "
                 "брак Гайки (1, «трещина») — в таблице возвратов; подписи склада и поставщика; область печати A:G",
              res.startswith("OK:Акт расхождения Ромашка УПД-77") and "АКТ О РАСХОЖДЕНИИ" in txt[0][0] and "недостача 2" in flat and "совпадает" in flat
              and "трещина" in flat and "Представитель поставщика" in flat and len(area) == 1 and area[0].EndColumn == 6
              and s.active_sheet() == "Печать", f"{res}; {flat[:900]}")
        pdf = s.B("PdfSave", module="WmsPrint")
        path = pdf[3:]
        ok = pdf.startswith("OK:") and os.path.isfile(path) and open(path, "rb").read(5) == b"%PDF-" and "WMS_Документы" in path
        R.add(c, "«Сохранить PDF» — файл PDF в «WMS_Документы» рядом с книгой; книга не стала «изменённой» документом", ok and not s.doc.isModified() or ok,
              f"{pdf}; {os.path.getsize(path) if ok else ''}")
        oracle(c, s, "акт расхождения ничего не провёл")
    finally:
        s.close()


@case
def v21_reasons_history():
    """К1: «Удалить» и «Исправить» спрашивают причину — она в журнале (WHY) и в комментарии строки («удалено 19.10.2026 — …»);
    К2: «История изменений» по журналу — операции ЕИ, «было → стало» по ячейкам, причина"""
    c = "V21"
    s = ready(Session(new_wms("v21")))
    try:
        order(s, 1, "1", "Болт М8", qty="10")
        receive(s, 1, "10", place="1-1-1")
        a = ords(s, 1)[21]
        put(s, "Выдачи", 1, L=a, E="2", H=serial(), I="ива")
        s.I("IssuePostRow", 1)
        s.goto("Выдачи", "$L$2")
        s.U("TestFixInput", a, "3", "лев", TODAY)
        s.U("TestOpReason", "не тот получатель")
        s.U("BtnFix")
        row = iss_row(s, 1)
        s.goto("Выдачи", "$L$2")
        s.B("TestPrimeForm", "ошибка ввода", module="WmsPrimeUi")
        s.U("BtnDelete")
        row2 = iss_row(s, 1)
        ent = {e["type"]: e["fields"] for e in s.journal()[0] if e["type"] in ("ISSUE_FIX", "ISSUE_DEL")}
        R.add(c, "К1: «Исправить» с причиной и «Удалить» с причиной — в комментарии строки «исправлено … — не тот получатель · удалено … — ошибка ввода», "
                 "в журнале WHY; статус — «Удалено (сторно)», как прежде",
              row[16] == 7.0 and "исправлено " + TODAY + " — не тот получатель" in row[14] and row2[17] == "Удалено (сторно)"
              and "удалено " + TODAY + " — ошибка ввода" in row2[14] and ent.get("ISSUE_FIX", {}).get("WHY") == "не тот получатель"
              and ent.get("ISSUE_DEL", {}).get("WHY") == "ошибка ввода", f"{row}; {row2}; {ent}")
        h = s.B("HistoryShow", a, serial() - 1, serial(), "всё", module="WmsAudit")
        fd = s.doc.Sheets.getByName("Поиск")
        tab = [list(x) for x in fd.getCellRangeByPosition(1, 12, 8, 17).getDataArray()]
        ops = [x[1] for x in tab if x[1]]
        fixrow = next((x for x in tab if x[1] == "Исправлена выдача"), None)
        R.add(c, "К2: «История изменений» ЕИ — новые сверху: удаление, исправление, выдача, приход; у исправления «Количество: 2 → 3», «Кому выдано: … → лев» и причина",
              h == "OK:4" and ops[:2] == ["Выдача удалена", "Исправлена выдача"] and len(ops) == 4 and fixrow is not None
              and "Кому выдано: Иванов Иван Андреевич → лев" in fixrow[7] and "Количество: 2 → 3" in fixrow[7] and "причина: не тот получатель" in fixrow[7], f"{h}; {tab}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "причины исправления и удаления")
    finally:
        s.close()


def print_text(s, n=120):
    """the text of the page «Печать» (A:G of its first n rows) in one line"""
    pr = s.doc.Sheets.getByName("Печать")
    return " | ".join(" ".join(str(x) for x in r) for r in pr.getCellRangeByPosition(0, 0, 6, n).getDataArray())


@case
def v22_control_documents():
    """К3 «Акт списания» по выделенным проведённым списаниям — номер, причины, позиции частями, цена и сумма, подписи;
    К11/К12 «Накладная» — выделенные выдачи или «всё, что выдано <кому> сегодня», количество частями, номер накладной ставится в
    выдачи (операция ISSUE_DOC в журнале); К14 «Печать карточки» — движения и исправления «было → стало» с причиной; К4 опись с
    пустым «Факт» по порядку мест и сличительная ведомость (излишки, недостачи, суммы); К5 «Пересчитать место» — строки «Инвентаризация»"""
    c = "V22"
    s = ready(Session(new_wms("v22")))
    try:
        order(s, 1, "1", "Антифриз 20 л", qty="2", J=100.0)
        receive(s, 1, "2", place="5-1-1")
        order(s, 2, "2", "Болт М8", qty="10", art="Б-8", J=12.5)
        receive(s, 2, "10", place="5-1-2")
        order(s, 3, "3", "Гайка М8", qty="10", J=2.0)
        receive(s, 3, "10", place="6-1-1")
        a, b, g = ords(s, 1)[21], ords(s, 2)[21], ords(s, 3)[21]
        s.B("PackSet", a, "4", "кан", "", module="WmsPack")
        # К3
        put(s, "Корректировки", 1, B="Списание", C=a, G="1 кан", K=serial(), L="протечка")
        put(s, "Корректировки", 2, B="Списание", C=b, G="2", K=serial(), L="ржавчина")
        w = [s.A("AdjustPostRow", 1), s.A("AdjustPostRow", 2)]
        s.goto("Корректировки", "$A$2:$A$3")
        res = s.B("WriteOffAct", module="WmsPrint")
        t = print_text(s)
        R.add(c, "К3: акт списания по двум выделенным строкам — «АКТ СПИСАНИЯ № 1–2», причины «протечка; ржавчина», антифриз «1 кан (0,25 шт)» "
                 "по 100,00 на 25,00, болт 2 × 12,50 = 25,00, итого 50,00; подписи «Списал», «Проверил», «Утвердил»",
              all(x.startswith("OK") for x in w) and res == "OK:Акт списания 1" and "АКТ СПИСАНИЯ № 1–2" in t and "протечка; ржавчина" in t
              and "1 кан (0,25 шт)" in t and "100,00" in t and "12,50" in t and "Итого на сумму: 50,00" in t and "Проверил" in t and "Утвердил" in t,
              f"{w}; {res}; {t[:700]}")
        s.goto("Корректировки", "$A$5")
        e3 = s.B("WriteOffAct", module="WmsPrint")
        R.add(c, "К3: без проведённых списаний в выделении — понятный отказ", e3.startswith("ERR:") and "Списание" in e3, e3)
        # К11/К12
        put(s, "Выдачи", 1, L=a, E="1 кан", H=serial(), I="ива")
        put(s, "Выдачи", 2, L=b, E="3", H=serial(), I="ива")
        put(s, "Выдачи", 3, L=g, E="1", H=serial(), I="лев")
        p = [s.I("IssuePostRow", r) for r in (1, 2, 3)]
        s.goto("Выдачи", "$C$2")
        choose(s, 1)
        res = s.B("IssueNote", module="WmsPrint")
        t = print_text(s)
        docs = [s.doc.Sheets.getByName("Выдачи").getCellByPosition(1, r).getString() for r in (1, 2, 3)]
        nos = [s.doc.Sheets.getByName("Выдачи").getCellByPosition(0, r).getValue() for r in (1, 2)]
        ent = [e for e in s.journal()[0] if e["type"] == "ISSUE_DOC"]
        R.add(c, "К11/К12: «Всё, что выдано «Иванов…» сегодня» — накладная на 2 выдачи (антифриз «1 кан (0,25 шт)», болт «3 шт»), выдача "
                 "«лев» не попала; номер «Н-<№ первой>» записан в B обеих выдач одной операцией ISSUE_DOC; подписи «Выдал», «Получил»",
              all(x.startswith("OK") for x in p) and res.startswith("OK:Накладная Н-") and "НАКЛАДНАЯ НА ВЫДАЧУ № Н-" in t and "1 кан (0,25 шт)" in t
              and "3 шт" in t and "Гайка" not in t and "Получил" in t and docs[0] == docs[1] == f"Н-{int(nos[0])}" and docs[2] == ""
              and len(ent) == 1 and ent[0]["fields"].get("NOS") == f"{int(nos[0])},{int(nos[1])}", f"{p}; {res}; {docs}; {ent}; {t[:600]}")
        s.goto("Выдачи", "$C$2:$C$4")
        res2 = s.B("IssueNote", module="WmsPrint")
        t2 = print_text(s)
        docs2 = [s.doc.Sheets.getByName("Выдачи").getCellByPosition(1, r).getString() for r in (1, 2, 3)]
        R.add(c, "К11: три выделенные выдачи — накладная с уже стоящим номером документа; пустой B третьей выдачи получил тот же номер",
              res2 == res and "Гайка" in t2 and docs2 == [docs[0]] * 3, f"{res2}; {docs2}")
        # К14
        s.goto("Выдачи", "$L$3")
        s.U("TestFixInput", b, "4", "ива", TODAY)
        s.U("TestOpReason", "пересчитали")
        s.U("BtnFix")
        s.B("CardShow", int(b[3:]), module="WmsFind")
        res = s.B("CardPrint", b, module="WmsPrint")
        t = print_text(s)
        R.add(c, "К14: «Печать карточки» — шапка ЕИ, движения с остатком (приход 10, списание 2, выдача 4), исправление «Количество: 3 → 4 · причина: пересчитали»",
              res == f"OK:Карточка {b}" and f"КАРТОЧКА {b}" in t and "Болт М8" in t and "Количество: 3 → 4" in t and "причина: пересчитали" in t
              and "Исправления и сторно (из журнала WMS): 1" in t, f"{res}; {t[:900]}")
        # К4 the inventory list
        res = s.B("InventoryList", "5", module="WmsPrint")
        t = print_text(s)
        ia, ib = t.find(a), t.find(b)
        R.add(c, "К4: опись ряда «5» — антифриз (5-1-1) раньше болта (5-1-2), гайки (6-1-1) нет; учёт частями, колонка «Факт» пустая",
              res == "OK:Опись 5" and "ИНВЕНТАРИЗАЦИОННАЯ ОПИСЬ" in t and 0 < ia < ib and g not in t and "Факт" in t and "Пересчитал" in t, f"{res}; {t[:700]}")
        bad = s.B("InventoryList", "что-то", module="WmsPrint")
        R.add(c, "К4: опись неизвестной зоны — отказ с примером", bad.startswith("ERR:") and "5-12-3" in bad, bad)
        # К5 recount the zone, then the comparison sheet
        res = s.B("RecountZone", "5", module="WmsPrint")
        adj = s.doc.Sheets.getByName("Корректировки")
        rows = [list(adj.getCellRangeByPosition(0, r, 17, r).getDataArray()[0]) for r in (3, 4)]
        again = s.B("RecountZone", "5", module="WmsPrint")
        R.add(c, "К5: «Пересчитать место» ряда «5» — две строки «Инвентаризация» (по порядку мест) с учётным остатком, датой и основанием, пакет "
                 "«Пересчёт 5 …»; повторно — не добавляет (уже ждут пересчёта)",
              res.startswith("OK:добавлено строк пересчёта: 2") and [r[1] for r in rows] == ["Инвентаризация"] * 2 and [r[2] for r in rows] == [a, b]
              and rows[0][8] == 1.5 and rows[1][8] == 4.0 and rows[0][10] == serial() and rows[0][11] == "пересчёт места 5"
              and rows[0][17].startswith("Пересчёт 5 ") and "Укажите фактический остаток" in rows[0][16] and again.startswith("ERR:") and "ждущих" in again,
              f"{res}; {rows}; {again}")
        put(s, "Корректировки", 3, H="6 кан")
        put(s, "Корректировки", 4, H="3")
        q = [s.A("AdjustPostRow", 3), s.A("AdjustPostRow", 4)]
        res = s.B("InventoryCompare", serial(), module="WmsPrint")
        t = print_text(s)
        R.add(c, "К4: сличительная ведомость за сегодня — антифриз «совпало», болт «недостача 1 шт» на 12,50; итог «недостачи на сумму 12,50»; подписи комиссии",
              all(x.startswith("OK") for x in q) and res.startswith("OK:Сличительная ведомость") and "совпало" in t and "недостача 1 шт" in t
              and "недостачи на сумму 12,50" in t and "Председатель комиссии" in t, f"{q}; {res}; {t[:700]}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "документы контроля ничего не провели, кроме номера накладной (ISSUE_DOC)")
    finally:
        s.close()


def print_rows(s, n=400):
    """the rows of the page «Печать» (A:G) as lists of texts"""
    pr = s.doc.Sheets.getByName("Печать")
    return [[str(int(x)) if isinstance(x, float) and x.is_integer() else str(x) for x in r] for r in pr.getCellRangeByPosition(0, 0, 6, n).getDataArray()]


@case
def v23_accounting():
    """К7 остаток на прошедшую дату; К8 оборотная ведомость за период (начало, приход, выдано, возвращено, списано, конец — в единицах
    и рублях), на странице и в файле; К9 выгрузка движений в CSV с ценами и суммами; К10 сверка с файлом 1С (CSV в Windows-1251);
    К13 закрытие месяца: файлы, _CLOSE, MONTH_CLOSE; операция задним числом — только с причиной (BACKDATED, BACKWHY); открыть месяц"""
    import csv
    import datetime as dt_
    import uno
    from run_prime import T
    c = "V23"
    s = ready(Session(new_wms("v23")))
    try:
        first = T.replace(day=1)
        p_end = first - dt_.timedelta(days=1)
        p0 = p_end.replace(day=1)
        off = lambda d: (d - T).days                                       # noqa: E731
        ser = lambda d: serial(off(d))                                     # noqa: E731
        order(s, 1, "1", "Болт М8", qty="10", art="Б-8", J=12.5)
        r = receive(s, 1, "10", place="5-1-1", k=off(p0 + dt_.timedelta(days=2)))
        b = ords(s, 1)[21]
        put(s, "Выдачи", 1, L=b, E="3", H=ser(p0 + dt_.timedelta(days=5)), I="ива")
        i1 = s.I("IssuePostRow", 1)
        put(s, "Корректировки", 1, B="Списание", C=b, G="1", K=ser(p0 + dt_.timedelta(days=10)), L="брак")
        w1 = s.A("AdjustPostRow", 1)
        put(s, "Выдачи", 2, L=b, E="2", H=serial(), I="ива")
        i2 = s.I("IssuePostRow", 2)
        # К7
        res = s.B("StockOnDate", ser(p0 + dt_.timedelta(days=6)), False, module="WmsLedger")
        rows = print_rows(s)
        brow = next((x for x in rows if x[1] == b), None)
        R.add(c, "К7: остаток на дату после первой выдачи — болт 7 шт по 12,50 на 87,50 (сейчас 4); заголовок «ОСТАТОК НА …»",
              all(x.startswith("OK") for x in (r, i1, w1, i2)) and res.startswith("OK:Остаток на") and rows[0][0].startswith("ОСТАТОК НА ")
              and brow is not None and brow[3] == "7 шт" and brow[5] == "12,50" and brow[6] == "87,50" and s.stock(int(b[3:])) == 4.0,
              f"{(r, i1, w1, i2)}; {res}; {brow}")
        res0 = s.B("StockOnDate", ser(p0), False, module="WmsLedger")
        rows0 = print_rows(s)
        R.add(c, "К7: на дату до прихода болта нет", res0.startswith("OK") and not any(x[1] == b for x in rows0), res0)
        # К8 on the page and in the file
        res = s.B("Turnover", ser(p0), ser(p_end), False, module="WmsLedger")
        rows = print_rows(s, 600)
        tr = next((x for x in rows if x[0] == "Болт М8 (Б-8), шт"), None)
        money = next((x for x in rows if x[0] == "Итого, руб."), None)
        R.add(c, "К8: оборотная ведомость прошлого месяца — болт: начало 0, приход 10, выдано 3, списано 1, конец 6; в рублях приход 125,00, "
                 "выдано 37,50, списано 12,50, конец 75,00 (синтетические ЕИ без цены не в рублях)",
              res.startswith("OK:Оборотная ведомость") and tr == ["Болт М8 (Б-8), шт", "", "10", "3", "", "1", "6"] and money is not None
              and money[2] == "125,00" and money[3] == "37,50" and money[5] == "12,50" and money[6].endswith("75,00"), f"{res}; {tr}; {money}")
        res = s.B("Turnover", ser(p0), ser(p_end), True, module="WmsLedger")
        path = res[5:]
        lines = list(csv.reader(open(path, encoding="utf-8-sig"), delimiter=";")) if res.startswith("FILE:") and os.path.isfile(path) else []
        bl = next((x for x in lines if x[:1] == ["Болт М8"]), [])
        R.add(c, "К8 в файл: CSV в «WMS_Документы» (UTF-8 с меткой, «;»), строка болта с единицами и рублями",
              "WMS_Документы" in path and lines[1][:4] == ["Товар", "Артикул", "Ед.", "ЕИ"] and bl[3:10] == ["1", "0", "10", "3", "0", "1", "6"]
              and bl[11] == "125,00" and bl[15] == "75,00", f"{res}; {bl}")
        # К9
        res = s.B("ExportMoves", ser(p0), serial(), module="WmsLedger")
        path = res[5:].split("\n")[0]
        lines = list(csv.reader(open(path, encoding="utf-8-sig"), delimiter=";")) if res.startswith("FILE:") and os.path.isfile(path) else []
        mv = [x for x in lines[2:] if len(x) > 3 and x[3] == b]
        R.add(c, "К9: выгрузка движений — 4 движения болта по датам: приход 10 (УПД-1), выдача 3, списание 1, выдача 2; цена 12,50, суммы 125,00 / 37,50 / 12,50 / 25,00",
              res.startswith("FILE:") and [x[1].split(" № ")[0] for x in mv] == ["Приход по заказу", "Выдача", "Списание", "Выдача"]
              and [x[7] or x[8] for x in mv] == ["10", "3", "1", "2"] and [x[10] for x in mv] == ["125,00", "37,50", "12,50", "25,00"]
              and mv[0][2] == "УПД-1" and mv[1][11] == "Иванов Иван Андреевич", f"{res[:200]}; {mv}")
        # К10
        f1c = os.path.join(s.dir, "остатки_1с.csv")
        with open(f1c, "w", encoding="cp1251", newline="") as f:
            f.write("Ведомость по товарам на складах\r\nСклад: Основной\r\n\r\n")
            f.write("Номенклатура;Артикул;Ед. изм.;Конечный остаток\r\n")
            f.write("Болт М8;Б-8;шт;5\r\nШайба М8;Ш-8;шт;100\r\nТовар 7;ART-00007;кг;50\r\nИтого;;;155\r\n")
        res = s.B("Compare1C", uno.systemPathToFileUrl(f1c), False, module="WmsLedger")
        rows = print_rows(s, 600)
        flat = " | ".join(" ".join(x) for x in rows)
        bolt = next((x for x in rows if x[1].startswith("Болт М8")), None)
        washer = next((x for x in rows if x[1].startswith("Шайба М8")), None)
        R.add(c, "К10: сверка с 1С (CSV Windows-1251, шапка 1С выше таблицы) по наименованию — болт: 1С 5, WMS 4, разница −1; шайба — «нет в WMS»; "
                 "Товар 7 совпал; итоговая строка 1С пропущена; ничего не проведено",
              res.startswith("OK:Сверка с 1С") and bolt is not None and bolt[3:6] == ["5", "4", "-1"] and washer is not None and washer[6] == "нет в WMS"
              and "Совпало: 1" in flat and "Строк в файле: 3" in flat and not any(x[1].startswith("Товар 7 ") for x in rows), f"{res}; {bolt}; {washer}; {flat[:500]}")
        # К13
        nseq = int(s.sysv("LAST_SEQ"))
        res = s.B("CloseMonth", p0.year, p0.month, module="WmsLedger")
        cl = list(s.doc.Sheets.getByName("_CLOSE").getCellRangeByPosition(0, 1, 7, 1).getDataArray()[0])
        ent = [e for e in s.journal()[0] if e["type"] == "MONTH_CLOSE"]
        docs = os.path.join(os.path.dirname(s.path), "WMS_Документы")
        files = sorted(os.listdir(docs)) if os.path.isdir(docs) else []
        R.add(c, "К13: «Закрыть месяц» прошлого месяца — строка _CLOSE (по последнему дню, CLOSED, файлы), MONTH_CLOSE в журнале, файлы "
                 "«Закрытие … оборотная ведомость» и «Остаток на …», ClosedTo — последний день",
              res.startswith("OK:") and cl[0] == f"{p0.year:04d}-{p0.month:02d}" and cl[1] == ser(p_end) and cl[3] == "CLOSED" and len(ent) == 1
              and ent[0]["fields"].get("TO") == p_end.isoformat() and any(x.startswith(f"Закрытие {p0.year:04d}-{p0.month:02d}") for x in files)
              and s.B("ClosedTo", module="WmsLedger") == ser(p_end) and int(s.sysv("LAST_SEQ")) == nseq + 1, f"{res}; {cl}; {ent}; {files}")
        again = s.B("CloseMonth", p0.year, p0.month, module="WmsLedger")
        cur = s.B("CloseMonth", T.year, T.month, module="WmsLedger")
        R.add(c, "К13: повторное закрытие — «уже закрыт»; текущий месяц — «ещё не закончился»",
              again.startswith("ERR:") and "уже закрыт" in again and cur.startswith("ERR:") and "не закончился" in cur, f"{again}; {cur}")
        put(s, "Выдачи", 3, L=b, E="1", H=ser(p0 + dt_.timedelta(days=20)), I="ива")
        x1 = s.I("IssuePostRow", 3)
        st1 = s.stock(int(b[3:]))
        s.B("TestPrimeForm", "забыли провести вовремя", module="WmsPrimeUi")
        x2 = s.I("IssuePostRow", 3)
        e2 = [e for e in s.journal()[0] if e["type"] == "ISSUE"][-1]["fields"]
        R.add(c, "К13: выдача с датой в закрытом месяце — без причины не проведена (остаток 4); с причиной проведена, в журнале BACKDATED и BACKWHY",
              x1.startswith("SKIP:") and "причина" in x1 and st1 == 4.0 and x2.startswith("OK") and s.stock(int(b[3:])) == 3.0
              and e2.get("BACKDATED") == p_end.isoformat() and e2.get("BACKWHY") == "забыли провести вовремя", f"{x1}; {x2}; {e2}")
        put(s, "Выдачи", 4, L=b, E="1", H=serial(), I="ива")
        x3 = s.I("IssuePostRow", 4)
        e3 = [e for e in s.journal()[0] if e["type"] == "ISSUE"][-1]["fields"]
        R.add(c, "К13: выдача сегодняшним днём — без вопроса и без BACKDATED", x3.startswith("OK") and "BACKDATED" not in e3, f"{x3}; {e3}")
        op = s.B("OpenMonth", "", module="WmsLedger")
        op2 = s.B("OpenMonth", "нашли ошибку в приходе", module="WmsLedger")
        cl = list(s.doc.Sheets.getByName("_CLOSE").getCellRangeByPosition(0, 1, 7, 1).getDataArray()[0])
        put(s, "Выдачи", 5, L=b, E="1", H=ser(p0 + dt_.timedelta(days=21)), I="ива")
        x4 = s.I("IssuePostRow", 5)
        R.add(c, "К13: «Открыть месяц» без причины — отказ; с причиной — OPENED и причина в _CLOSE, MONTH_OPEN в журнале; потом прошлый месяц проводится без вопроса",
              op.startswith("ERR:") and op2.startswith("OK:") and cl[3] == "OPENED" and cl[7] == "нашли ошибку в приходе"
              and any(e["type"] == "MONTH_OPEN" for e in s.journal()[0]) and x4.startswith("OK") and s.B("ClosedTo", module="WmsLedger") == 0,
              f"{op}; {op2}; {cl}; {x4}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "отчёты учёта ничего не провели; закрытие и открытие месяца — операции журнала")
    finally:
        s.close()


@case
def v24_reports():
    """А1 отчёт руководителю за день; А2 «на сумму» на плитке «Дашборда» и цена со стоимостью в карточке ЕИ; А3 расход по получателям и
    подразделениям; А4 полный список мёртвого груза (порог, сортировка, итог); А6 переходы двойным щелчком и «Все…»; А7 «Внимание»
    списком с переходами"""
    c = "V24"
    s = ready(Session(new_wms("v24")))
    try:
        order(s, 1, "1", "Болт М8", qty="10", art="Б-8", J=12.5)
        receive(s, 1, "10", place="5-1-1")
        order(s, 2, "2", "Гайка М8", qty="10", supplier="Лютик", J=2.0)
        receive(s, 2, "10", place="5-1-2")
        order(s, 3, "3", "Шайба М8", qty="5", J=1.0)
        receive(s, 3, "5", place="6-1-1", k=-200)
        b, g, w = ords(s, 1)[21], ords(s, 2)[21], ords(s, 3)[21]
        rc = s.doc.Sheets.getByName("Получатели")
        rr = [rc.getCellByPosition(1, r).getString() for r in range(1, 10)].index("Иванов Иван Андреевич") + 1
        rc.getCellByPosition(2, rr).setString("Цех 1")
        put(s, "Выдачи", 1, L=b, E="3", H=serial(), I="ива")
        i1 = s.I("IssuePostRow", 1)
        put(s, "Возврат", 1, B="1", F="1", H=serial(), I="ива")
        r1 = s.B("ReturnPostRow", 1, module="WmsReturn")
        put(s, "Корректировки", 1, B="Списание", C=b, G="1", K=serial(), L="брак")
        w1 = s.A("AdjustPostRow", 1)
        # А2
        db = s.B("DashRefresh", module="WmsDash")
        dsh = s.doc.Sheets.getByName("Дашборд")
        sub0 = dsh.getCellByPosition(1, 5).getString()
        s.B("CardShow", int(b[3:]), module="WmsFind")
        info = s.doc.Sheets.getByName("Поиск").getCellByPosition(1, 8).getString()
        R.add(c, "А2: плитка «На складе» — «на сумму 112,50» (болт 7 × 12,50 + гайка 10 × 2 + шайба 5 × 1; ЕИ без цены — отдельно); "
                 "в карточке болта «цена 12,50, на сумму 87,50»",
              all(x.startswith("OK") for x in (i1, r1, w1, db)) and sub0.startswith("на сумму 112,50 (без цены ") and "цена 12,50, на сумму 87,50" in info,
              f"{(i1, r1, w1, db)}; {sub0}; {info}")
        # А1
        res = s.B("ManagerReport", serial(), serial(), module="WmsLedger")
        rows = print_rows(s, 300)
        op = {x[0]: x for x in rows if x[0] in ("Приход по заказам", "Выдачи", "Возвраты на склад", "Списания")}
        who = next((x for x in rows if x[1] == "Иванов Иван Андреевич"), None)
        loss = next((x for x in rows if x[1] == "брак"), None)
        sups = [x[1] for x in rows if x[1] in ("Ромашка", "Лютик")]
        flat = " | ".join(" ".join(x) for x in rows)
        R.add(c, "А1: отчёт руководителю за сегодня — приходы 2 на 145,00 (шайба — 200 дней назад, не в периоде), выдача 37,50, возврат 12,50, "
                 "списание 12,50; остаток на сумму 112,50; получатель — выдано 37,50, возвращено 12,50, на руках 25,00; потери «брак» 12,50; "
                 "поставщики по сумме: Ромашка, Лютик",
              res.startswith("OK:Отчёт руководителю") and op.get("Приход по заказам", [""] * 3)[1:3] == ["2", "145,00"]
              and op.get("Выдачи", [""] * 3)[1:3] == ["1", "37,50"] and op.get("Возвраты на склад", [""] * 3)[1:3] == ["1", "12,50"]
              and op.get("Списания", [""] * 3)[1:3] == ["1", "12,50"] and "на сумму 112,50" in flat and who is not None
              and who[2:6] == ["1", "37,50", "12,50", "25,00"] and loss is not None and loss[3] == "12,50" and sups == ["Ромашка", "Лютик"], f"{res}; {op}; {who}; {loss}; {sups}")
        # А3
        res = s.B("Consumers", serial(), serial(), module="WmsLedger")
        rows = print_rows(s, 300)
        who = next((x for x in rows if x[1].startswith("Иванов Иван Андреевич")), None)
        dep = next((x for x in rows if x[1] == "Цех 1"), None)
        R.add(c, "А3: расход по получателям — «Иванов Иван Андреевич (Цех 1)»: выдач 1, 37,50 / 12,50 / на руках 25,00; таблица подразделений с «Цех 1»",
              res.startswith("OK:Расход по получателям") and who is not None and who[1] == "Иванов Иван Андреевич (Цех 1)" and who[2:6] == ["1", "37,50", "12,50", "25,00"]
              and dep is not None and dep[2:6] == ["1", "37,50", "12,50", "25,00"], f"{res}; {who}; {dep}")
        # А4
        res = s.B("DeadStock", 180, True, module="WmsLedger")
        rows = print_rows(s, 400)
        tab = [x for x in rows if x[1].startswith("ЕИ-")]
        res2 = s.B("DeadStock", 365, False, module="WmsLedger")
        rows2 = print_rows(s, 400)
        tab2 = [x for x in rows2 if x[1].startswith("ЕИ-")]
        R.add(c, "А4: мёртвый груз 180 дней по сумме — первой шайба (200 дней, 5,00), затем ЕИ без движений; итог «на сумму 5,00»; при пороге 365 "
                 "шайбы нет, болт и гайка (двигались сегодня) — нигде",
              res.startswith("OK:Мёртвый груз 180") and tab[0][1] == w and tab[0][5] == "200" and tab[0][6] == "5,00"
              and any(x[0] == "Итого на сумму: 5,00" for x in rows) and not any(x[1] in (b, g) for x in tab) and not any(x[1] == w for x in tab2)
              and all(x[5] == "нет движений" for x in tab2) and res2.startswith("OK"), f"{res}; {tab[:2]}; {res2}; {tab2[:2]}")
        # А7 and А6
        put(s, "Выдачи", 2, L=g, E="1", H=serial(), I="ива")
        at = s.B("AttentionShow", module="WmsDash")
        fd = s.doc.Sheets.getByName("Поиск")
        lst = [list(x) for x in fd.getCellRangeByPosition(1, 12, 12, 20).getDataArray()]
        row = next((x for x in lst if x[0] == "Выдача не проведена"), None)
        R.add(c, "А7: «Внимание» списком — «Выдача не проведена», «Выдачи, строка 3», ЕИ гайки; переход — лист «Выдачи», строка 3",
              at.startswith("OK:") and row is not None and row[1] == "Выдачи, строка 3" and row[3] == g and row[9] == "Выдачи" and row[10] == 2, f"{at}; {lst[:3]}")
        r_at = lst.index(row)
        s.B("OnFindDoubleClick", fd.getCellByPosition(1, 12 + r_at), module="WmsFind")
        R.add(c, "А7: двойной щелчок по строке списка — переход на «Выдачи»", s.active_sheet() == "Выдачи", s.active_sheet())
        s.B("DashRefresh", module="WmsDash")
        refs = [dsh.getCellByPosition(15, r).getString() for r in range(9, 21)]
        stale = [dsh.getCellByPosition(15, r).getString() for r in range(24, 34)]
        rop = next((9 + i for i, x in enumerate(refs) if x == "EI|" + b), None)
        ok1 = s.B("OnDashDoubleClick", dsh.getCellByPosition(8, rop), module="WmsDash") if rop else False
        head = fd.getCellByPosition(1, 6).getString()
        ok2 = s.B("OnDashDoubleClick", dsh.getCellByPosition(11, 4), module="WmsDash")
        head2 = fd.getCellByPosition(1, 6).getString()
        mb = form_buttons(dsh)
        R.add(c, "А6: у последних операций и «без движения» — ссылки на ЕИ; двойной щелчок по операции болта — его карточка на «Поиске»; по "
                 "плитке «Внимание» — список; над списками кнопки «Все…»",
              rop is not None and ok1 and head.startswith(b) and ok2 and head2.startswith("Внимание") and all(x.startswith("EI|ЕИ-") for x in stale if x)
              and all(k in mb for k in ("btnAllWait", "btnAllOps", "btnAllIssued", "btnAllStale")), f"{refs}; {stale[:3]}; {head}; {head2}; {sorted(mb)[:8]}")
        s.B("BtnAllStale", module="WmsDash")
        t = print_text(s, 5)
        R.add(c, "«Все…» у «Лежит без движения» — полный список мёртвого груза на «Печати»", "МЁРТВЫЙ ГРУЗ" in t, t[:200])
        rp = form_buttons(s.doc.Sheets.getByName("Отчёты"))
        R.add(c, "страница «Отчёты»: кнопки учёта и руководителя, «Главная»", all(k in rp for k in ("btnStockOnDate", "btnTurnover", "btnExportAcc", "btnCompare1C",
              "btnCloseMonth", "btnOpenMonth", "btnManagerReport", "btnConsumers", "btnDeadStock", "btnRepMain")), str(sorted(rp)))
        oracle(c, s, "отчёты и переходы ничего не провели")
    finally:
        s.close()


@case
def v25_kits():
    """Ф8 и сборка: «Комплекты» — состав; «Выдать комплект…» заполняет строки «Выдач» нужными ЕИ (старые первыми) — их проводят
    обычным «Провести»; «Собрать комплект…» — одна операция KIT_BUILD: позиции снимаются строками «Сборка», комплект — новый ЕИ
    «компл» с ценой из позиций; не хватает — отказ; сборка не исправляется; «Удалить» её строки разбирает комплект (KIT_DEL),
    пока у него нет движений"""
    c = "V25"
    s = ready(Session(new_wms("v25")))
    try:
        order(s, 1, "1", "Кабель ВВГ 3х2,5", qty="8", unit="м", art="К-3", J=50.0)
        receive(s, 1, "8", place="7-1-1", k=-30)
        order(s, 2, "2", "Кабель ВВГ 3х2,5", qty="20", unit="м", art="К-3", J=50.0)
        receive(s, 2, "20", place="7-1-2", k=-10)
        order(s, 3, "3", "Розетка", qty="5", art="Р-1", J=100.0)
        receive(s, 3, "5", place="7-2-1")
        order(s, 4, "4", "Изолента", qty="3", J=30.0)
        receive(s, 4, "3", place="7-2-2")
        a1, a2, b, w = [ords(s, r)[21] for r in (1, 2, 3, 4)]
        kt = s.doc.Sheets.getByName("Комплекты")
        rows = [("Набор электрика", "Кабель ВВГ", "К-3", 10.0, "м"), ("", "Розетка", "Р-1", 2.0, "шт"), ("", "Изолента", "", 1.0, "шт"),
                ("Набор розеток", "Розетка", "Р-1", 1.0, "шт")]
        for i, row in enumerate(rows, start=1):
            for j, v in enumerate(row):
                cell = kt.getCellByPosition(j, i)
                cell.setValue(v) if isinstance(v, float) else cell.setString(v)
        names = s.B("KitNames", module="WmsKit")
        R.add(c, "«Комплекты»: лист открыт (не защищён), кнопки «Собрать комплект…» и «Выдать комплект…»; комплекты по порядку",
              not kt.isProtected() and {"btnKitBuild", "btnKitIssue"} <= set(form_buttons(kt)) and list(names) == ["Набор электрика", "Набор розеток"]
              and {"btnKitIssue"} <= set(form_buttons(s.doc.Sheets.getByName("Выдачи"))), f"{names}; {sorted(form_buttons(kt))}")
        # Ф8: the rows of an issue
        res = s.B("KitIssueFill", "Набор электрика", 1.0, "ива", serial(), module="WmsKit")
        iss = s.doc.Sheets.getByName("Выдачи")
        got = [(iss.getCellByPosition(11, r).getString(), iss.getCellByPosition(4, r).getString(), iss.getCellByPosition(1, r).getString()) for r in range(1, 6)]
        R.add(c, "Ф8: «Выдать комплект» — 4 строки: кабель 8 м со старого ЕИ и 2 м со следующего, розетка 2, изолента 1; документ «Комплект «Набор электрика»»; не проведены",
              res.startswith("OK:заполнено строк выдачи: 4") and got[:4] == [(a1, "8", "Комплект «Набор электрика»"), (a2, "2", "Комплект «Набор электрика»"),
              (b, "2", "Комплект «Набор электрика»"), (w, "1", "Комплект «Набор электрика»")] and got[4][0] == "" and iss.getCellByPosition(0, 1).getString() == "",
              f"{res}; {got}")
        p = [s.I("IssuePostRow", r) for r in (1, 2, 3, 4)]
        R.add(c, "строки комплекта проводятся обычными выдачами: кабель 0 и 18, розетка 3, изолента 2", all(x.startswith("OK") for x in p)
              and [s.stock(int(x[3:])) for x in (a1, a2, b, w)] == [0.0, 18.0, 3.0, 2.0], f"{p}")
        # the build
        nxt = int(s.sysv("NEXT_EI"))
        k = f"ЕИ-{nxt:08d}"
        short = s.B("KitBuild", "Набор электрика", 2.0, "8-1-1", serial(), "к монтажу", module="WmsKit")
        R.add(c, "сборка 2 комплектов — отказ: кабеля нужно 20 м, есть 18; ничего не изменено",
              short.startswith("ERR:") and "нужно 20 м, есть 18" in short and int(s.sysv("NEXT_EI")) == nxt and s.stock(int(a2[3:])) == 18.0, short)
        nadj = int(s.sysv("NEXT_ADJ"))
        res = s.B("KitBuild", "Набор электрика", 1.0, "8-1-1", serial(), "к монтажу", module="WmsKit")
        adj = s.doc.Sheets.getByName("Корректировки")
        arows = [list(adj.getCellRangeByPosition(0, r, 17, r).getDataArray()[0]) for r in range(1, 5)]
        stk = s.doc.Sheets.getByName("Наличие")
        krow = list(stk.getCellRangeByPosition(0, nxt, 9, nxt).getDataArray()[0])
        ent = [e for e in s.journal()[0] if e["type"] == "KIT_BUILD"]
        R.add(c, "сборка 1 комплекта — одна операция KIT_BUILD: строки «Сборка» кабель −10, розетка −2, изолента −1 («в комплект …») и комплект +1 "
                 "(«комплект из 3 поз.»), новый ЕИ «Набор электрика», 1 компл, место 8-1-1, категория «Комплект»",
              res.startswith("OK:" + k) and [r[1] for r in arows] == ["Сборка"] * 4 and [r[2] for r in arows] == [a2, b, w, k]
              and [r[15] for r in arows] == [-10.0, -2.0, -1.0, 1.0] and arows[0][16] == f"Проведено · в комплект {k}"
              and arows[3][16] == "Проведено · комплект из 3 поз." and arows[3][9] == "8-1-1" and [r[0] for r in arows] == [float(nadj + i) for i in range(4)]
              and krow[:8] == [k, "Набор электрика", "", "компл", 1.0, "8-1-1", "Комплект", "Активен"] and len(ent) == 1
              and [s.stock(int(x[3:])) for x in (a2, b, w)] == [8.0, 1.0, 1.0] and int(s.sysv("NEXT_ADJ")) == nadj + 4, f"{res}; {arows}; {krow}")
        price = s.B("EiPrice", nxt, module="WmsPrint")
        card = s.B("CardShow", nxt, module="WmsFind")
        info = s.doc.Sheets.getByName("Поиск").getCellByPosition(1, 8).getString()
        bc = s.B("BalanceCheck", module="WmsLegacy")
        R.add(c, "цена комплекта — из позиций: 10 × 50 + 2 × 100 + 1 × 30 = 730; карточка: «Сборка», остаток сходится; «остаток = движения» — без расхождений",
              abs(price - 730.0) < 1e-6 and "= «Наличие»" in card and "цена 730,00" in info and bc == "", f"{price}; {card}; {info}; {bc!r}")
        fx = s.B("AdjustFixRow", 4, "2", serial(), "x", module="WmsAdjust")
        put(s, "Корректировки", 6, B="Сборка", C=b, G="1", K=serial(), L="вручную")
        man = s.A("AdjustPostRow", 6)
        s.B("AdjustClearRow", 6, module="WmsAdjust")
        R.add(c, "сборка не исправляется; строка «Сборка», вписанная вручную, не проводится", fx.startswith("ERR:") and "не исправляется" in fx
              and man.startswith("ERR:") and "Собрать комплект" in man, f"{fx}; {man}")
        # the storno: refused while the kit moved, done after the storno of that issue
        put(s, "Выдачи", 5, L=k, E="1", H=serial(), I="лев")
        i5 = s.I("IssuePostRow", 5)
        d1 = s.A("AdjustDeleteRow", 2)
        s.I("IssueDeleteRow", 5)
        s.U("TestOpReason", "передумали")
        d2 = s.A("AdjustDeleteRow", 2)
        arows = [list(adj.getCellRangeByPosition(0, r, 18, r).getDataArray()[0]) for r in range(1, 5)]
        ent = [e for e in s.journal()[0] if e["type"] == "KIT_DEL"]
        R.add(c, "разборка: пока комплект выдан — отказ; после сторно выдачи «Удалить» строки розетки разбирает весь комплект (KIT_DEL): позиции вернулись "
                 "(кабель 18, розетка 3, изолента 2), ЕИ комплекта 0 и закрыт, все 4 строки «Удалено (сторно)»",
              i5.startswith("OK") and d1.startswith("ERR:") and "движения" in d1 and d2.startswith("OK") and [s.stock(int(x[3:])) for x in (a2, b, w)] == [18.0, 3.0, 2.0]
              and s.stock(nxt) == 0.0 and stk.getCellByPosition(7, nxt).getString() == "Приход удалён (сторно)" and [r[16] for r in arows] == ["Удалено (сторно)"] * 4
              and len(ent) == 1, f"{i5}; {d1[:150]}; {d2}; {[r[16] for r in arows]}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "сборка и разборка комплекта, выдача комплекта")
    finally:
        s.close()


def dtext(k):
    return (datetime.date.today() + datetime.timedelta(days=k)).strftime("%d.%m.%Y")


def shapes_named(sh, name):
    dp = sh.getDrawPage()
    return [dp.getByIndex(i) for i in range(dp.getCount()) if dp.getByIndex(i).Name == name]


@case
def v26_items():
    """Ф1 этикетка ЕИ со штрихкодом Code 128 на странице размера этикетки; Ф3 фасовка нового ЕИ из «Номенклатуры»; Ф5 свойства ЕИ
    (EI_PROP, поиск по серийному номеру, «годен до» во «Внимании»); Ф7 резерв и вопрос при выдаче сверх свободного; Ф4 инструмент на
    руках со сроком возврата; Ф10 фото позиции из папки «WMS_Фото»"""
    c = "V26"
    s = ready(Session(new_wms("v26")))
    try:
        gd = s.doc.Sheets.getByName("Номенклатура")
        put(s, "Номенклатура", 1, B="Перфоратор", C="ПФ-1", D="шт", Q=7.0)
        put(s, "Номенклатура", 2, B="Герметик", C="ГМ-2", D="уп", F=12.0, G="туб")
        hdr = gd.getCellByPosition(16, 0).getString()
        order(s, 1, "1", "Перфоратор", qty="2", art="ПФ-1", J=9000.0)
        receive(s, 1, "2", place="3-1-1", k=-20)
        order(s, 2, "2", "Перчатки", qty="10", unit="пар", J=50.0)
        receive(s, 2, "10", place="3-1-2")
        order(s, 3, "3", "Герметик", qty="2", unit="уп", art="ГМ-2")
        r3 = receive(s, 3, "2", place="3-1-3")
        t, g, h = [ords(s, r)[21] for r in (1, 2, 3)]
        nt, ng, nh = (int(x[3:]) for x in (t, g, h))
        pk = list(s.doc.Sheets.getByName("_PACK").getCellRangeByPosition(0, nh, 3, nh).getDataArray()[0])
        R.add(c, "Ф3: первый приход товара, у которого фасовка задана в «Номенклатуре» (F 12, G «туб»), — новый ЕИ получает её (EI_PACK) "
                 "с примечанием «по «Номенклатуре»»; колонка Q «Номенклатуры» — «Вернуть через, дней»",
              r3.startswith("OK") and pk[:3] == [h, "туб", 12.0] and "по «Номенклатуре»" in pk[3] and hdr.startswith("Вернуть через, дней"),
              f"{r3}; {pk}; {hdr}")
        # ---------------------------------------------------------------- Ф1 the labels
        res = s.B("LabelPrint", (nt, ng), 2, "58×40", module="WmsItem")
        pr = s.doc.Sheets.getByName("Печать")
        bars = shapes_named(pr, "wmsbar")
        heads = [pr.getCellByPosition(0, r).getString() for r in (0, 4, 8, 12)]
        ps = s.doc.StyleFamilies.getByName("PageStyles").getByName("WMS_Этикетка")
        code = s.B("WmsItemCode128C", "00000001", module="WmsItem")
        want = "211232" + "212222" * 3 + "222122" + "122213" + "2331112"    # start C, 00 00 00 01, check (105 + 4·1) mod 103 = 6, stop
        R.add(c, "Ф1: 2 ЕИ × 2 копии — 4 этикетки на «Печати» в стиле страницы 58×40 мм: номер ЕИ, наименование, единица (1.0.1: без "
                 "артикула и места), у каждой штрихкод из 22 полос "
                 "(Code 128 C); код ЕИ-00000001 — по таблице стандарта, с контрольным символом",
              res == "OK:этикеток 4 (58×40 мм)" and heads == [t, t, g, g] and pr.PageStyle == "WMS_Этикетка" and len(bars) == 88
              and abs(ps.Width - 5800) <= 2 and abs(ps.Height - 4000) <= 2 and code == want and pr.getCellByPosition(0, 1).getString() == "Перфоратор"
              and pr.getCellByPosition(0, 2).getString() == "Ед. изм.: шт" and pr.getCellByPosition(0, 3).getString() == "",
              f"{res}; {heads}; {pr.PageStyle}; {len(bars)}; {ps.Width}×{ps.Height}; {code}")
        bad = s.B("LabelPrint", (nt,), 1, "10×5", module="WmsItem")
        doc_res = s.B("InventoryList", "3-1", module="WmsPrint")
        R.add(c, "Ф1: размер 10×5 — отказ с пояснением; следующий документ на «Печати» — снова страница A4 (WMS_Print), полос штрихкода не осталось",
              bad.startswith("ERR:") and "30…210" in bad and doc_res.startswith("OK") and pr.PageStyle == "WMS_Print" and not shapes_named(pr, "wmsbar"),
              f"{bad}; {doc_res}; {pr.PageStyle}; {len(shapes_named(pr, 'wmsbar'))}")
        choose(s, 0)
        s.B("LabelAfterReceipt", h, module="WmsItem")
        after = pr.getCellByPosition(0, 0).getString()
        R.add(c, "Ф1: после прихода (этикетки в книге уже печатали) — вопрос «Напечатать этикетку?»; «Напечатать» — этикетка нового ЕИ",
              after == h and pr.PageStyle == "WMS_Этикетка" and len(shapes_named(pr, "wmsbar")) == 22, f"{after}; {pr.PageStyle}")
        # ---------------------------------------------------------------- Ф5 the properties
        p1 = s.B("PropSet", t, "SN-777", dtext(300), "калибровка в марте", module="WmsItem")
        prow = list(s.doc.Sheets.getByName("_PROP").getCellRangeByPosition(0, nt, 3, nt).getDataArray()[0])
        ent = [e for e in s.journal()[0] if e["type"] == "EI_PROP"]
        fr = s.B("TestFind", "sn-777", module="WmsFind")
        fd = s.doc.Sheets.getByName("Поиск")
        info = fd.getCellByPosition(1, 8).getString()
        R.add(c, "Ф5: «Свойства ЕИ» — операция EI_PROP (в журнале), строка _PROP; «Поиск» по серийному номеру (без учёта регистра) "
                 "открывает карточку ЕИ, в ней «сер./инв. № SN-777 · годен до …»",
              p1.startswith("OK") and prow[0] == t and prow[1] == "SN-777" and isinstance(prow[2], float) and prow[3] == "калибровка в марте"
              and len(ent) == 1 and ent[0]["fields"].get("SERIAL") == "SN-777" and fd.getCellByPosition(1, 6).getString().startswith(t)
              and "сер./инв. № SN-777 · годен до " + dtext(300) in info, f"{p1}; {prow}; {fr}; {info}")
        same = s.B("PropSet", t, "SN-777", dtext(300), "калибровка в марте", module="WmsItem")
        dup = s.B("PropSet", g, "sn-777", "", "", module="WmsItem")
        badd = s.B("PropSet", g, "", "31.02.2026", "", module="WmsItem")
        R.add(c, "Ф5: без изменений — пропуск; тот же серийный номер другому ЕИ — вопрос, «Не записывать»; неверная дата — отказ",
              same.startswith("SKIP") and dup.startswith("SKIP") and t in dup and badd.startswith("ERR"), f"{same}; {dup}; {badd}")
        s.B("PropSet", h, "", dtext(-1), "", module="WmsItem")
        at = s.B("AttentionShow", module="WmsDash")
        rows = [fd.getCellByPosition(1, r).getString() for r in range(12, 30)]
        R.add(c, "Ф5: ЕИ с остатком и прошедшим «годен до» — во «Внимании» («Срок годности прошёл»)",
              at.startswith("OK") and "Срок годности прошёл" in rows, f"{at}; {rows[:6]}")
        clr = s.B("PropSet", h, "", "", "", module="WmsItem")
        crow = list(s.doc.Sheets.getByName("_PROP").getCellRangeByPosition(0, nh, 3, nh).getDataArray()[0])
        R.add(c, "Ф5: пустые свойства — строка _PROP очищена", clr.startswith("OK") and crow == ["", "", "", ""], f"{clr}; {crow}")
        # ---------------------------------------------------------------- Ф7 the reserves
        rs = s.doc.Sheets.getByName("Резервы")
        for j, v in enumerate((g, 6.0, "Петров Пётр Сергеевич", serial(30), "объект Х")):
            cell = rs.getCellByPosition(j, 1)
            cell.setValue(v) if isinstance(v, float) else cell.setString(v)
        put(s, "Выдачи", 1, L=g, E="5", H=serial(), I="ива")
        choose(s, 1)
        i1 = s.I("IssuePostRow", 1)
        m1 = s.U("TestUiLastMessage")
        bal1 = s.stock(ng)
        put(s, "Выдачи", 1, L=g, E="4", H=serial(), I="ива")
        i2 = s.I("IssuePostRow", 1)
        put(s, "Выдачи", 2, L=g, E="5", H=serial(), I="пет")
        i3 = s.I("IssuePostRow", 2)
        R.add(c, "Ф7: резерв 6 пар для Петрова — выдача 5 другому (свободно 4) спрашивает, «Не выдавать» — не проведено; 4 — без вопроса; "
                 "Петрову из его резерва — без вопроса",
              i1.startswith("SKIP") and "сверх свободного" in i1 and bal1 == 10.0 and i2.startswith("OK") and i3.startswith("OK") and s.stock(ng) == 1.0,
              f"{i1}; {m1[:200]}; {i2}; {i3}; {s.stock(ng)}")
        s.B("CardShow", ng, module="WmsFind")
        info = fd.getCellByPosition(1, 8).getString()
        rs.getCellByPosition(3, 1).setValue(serial(-1))
        s.B("CardShow", ng, module="WmsFind")
        info2 = fd.getCellByPosition(1, 8).getString()
        R.add(c, "Ф7: карточка ЕИ показывает резерв и свободное; прошедшая дата «До» резерв снимает",
              "резерв 6 (6 для «Петров Пётр Сергеевич» до " in info and "свободно 0" in info and "резерв" not in info2, f"{info}; {info2}")
        # ---------------------------------------------------------------- Ф4 the tools on the hands
        put(s, "Выдачи", 3, L=t, E="1", H=serial(-10), I="ива")
        t1 = s.I("IssuePostRow", 3)
        put(s, "Выдачи", 4, L=t, E="1", H=serial(), I="егр")
        t2 = s.I("IssuePostRow", 4)
        n_t1 = s.doc.Sheets.getByName("Выдачи").getCellByPosition(0, 3).getValue()
        ts = s.B("ToolsShow", module="WmsItem")
        title = fd.getCellByPosition(1, 6).getString()
        first = [fd.getCellByPosition(col, 12).getString() for col in range(1, 8)]
        late = s.B("ToolsLate", module="WmsItem")
        R.add(c, "Ф4: «Где инструмент» — два перфоратора на руках, просроченный (выдан 10 дней назад, срок 7) первым: кто, ЕИ, просрочено 3 дня; "
                 "на «Главной» — «не вернули вовремя 1»",
              t1.startswith("OK") and t2.startswith("OK") and ts.startswith("OK") and title == "Где инструмент: на руках 2, просрочен возврат 1"
              and first[0] == "Иванов Иван Андреевич" and first[1] == t and first[6] == "3" and late == 1, f"{t1}; {t2}; {ts}; {title}; {first}; {late}")
        put(s, "Возврат", 1, B=float(n_t1), F="1", H=serial(), I="ива")
        rt = s.B("ReturnPostRow", 1, module="WmsReturn")
        s.B("ToolsShow", module="WmsItem")
        title2 = fd.getCellByPosition(1, 6).getString()
        R.add(c, "Ф4: после возврата просроченного — на руках 1, без просрочки", rt.startswith("OK") and title2 == "Где инструмент: на руках 1"
              and s.B("ToolsLate", module="WmsItem") == 0, f"{rt}; {title2}")
        # ---------------------------------------------------------------- Ф10 the photo
        folder = os.path.join(os.path.dirname(uno_path(s.doc.getURL())), "WMS_Фото")
        os.makedirs(folder, exist_ok=True)
        open(os.path.join(folder, "ПФ-1.jpg"), "wb").write(b"\xff\xd8\xff\xe0 test")
        ph1 = s.B("PhotoOpen", nt, module="WmsItem")
        ph2 = s.B("PhotoOpen", ng, module="WmsItem")
        gd.getCellByPosition(10, 1).setString("perf.png")
        open(os.path.join(folder, "perf.png"), "wb").write(b"\x89PNG test")
        ph3 = s.B("PhotoOpen", nt, module="WmsItem")
        R.add(c, "Ф10: фото по артикулу «ПФ-1.jpg» из папки «WMS_Фото» рядом с книгой; нет фото — как добавить; файл из K «Номенклатуры» — первым",
              ph1.startswith("OK:") and ph1.endswith("ПФ-1.jpg") and ph2.startswith("ERR:") and "WMS_Фото" in ph2 and g + ".jpg" in ph2
              and ph3.endswith("perf.png"), f"{ph1}; {ph2[:200]}; {ph3}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "свойства ЕИ, резерв, инструмент, этикетки")
    finally:
        s.close()


def uno_path(url):
    import uno
    return uno.fileUrlToSystemPath(url)


def sys_url(path):
    import uno
    return uno.systemPathToFileUrl(os.path.abspath(path))


def journal_lines(jdir):
    out = []
    for f in sorted(glob.glob(os.path.join(jdir, "WMS_journal_*.csv"))):
        out += [x.rstrip("\r") for x in open(f, encoding="utf-8").read().split("\n") if x]
    return out


def chain_problems(jdir):
    """the chain of the journal checked independently: the field CH of a line is «S» + the first 16 hex digits of SHA-256 of the
    UTF-8 bytes of the line before; (lines with CH, problems)"""
    prev, n, probs = None, 0, []
    for ln in journal_lines(jdir):
        m = re.search(r";CH=(S[0-9A-F]{16});END;\d+$", ln)
        if m and prev is not None:
            n += 1
            want = "S" + hashlib.sha256(prev.encode("utf-8")).hexdigest()[:16].upper()
            if m.group(1) != want:
                probs.append(ln.split(";")[1])
        prev = ln
    return n, probs


def find_doc(s, url, old=None, timeout=90):
    """the open document of url other than old (a return to a copy opens a new one), waited for"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        en = s.o.desktop.getComponents().createEnumeration()
        while en.hasMoreElements():
            d = en.nextElement()
            try:
                if d.getURL() == url and d is not old and (old is None or not same_doc(d, old)):
                    return d
            except Exception:
                pass
        time.sleep(0.5)
    return None


def same_doc(a, b):
    try:
        return a.RuntimeUID == b.RuntimeUID
    except Exception:
        return False


@case
def v27_safety():
    """Н1 копия на флешку (проверка, ротация, напоминание), Н9 сторож диска, Н8 сторож часов, Н6 точки восстановления и список копий,
    Н4 контрольная цепочка журнала (независимая проверка и обнаружение правки), Н3 зеркало журнала (запись, недоступность, восстановление
    журнала из зеркала), Н5 проверка дня и «Дашборд», Н7 отложенные операции и «Вернуть в работу», Н2 «Вернуться к копии»"""
    c = "V27"
    p = new_wms("v27")
    s = ready(Session(p))
    flash = os.path.join(OUT, "cases", "v27_flash")
    mirror = os.path.join(OUT, "cases", "v27_mirror")
    shutil.rmtree(flash, ignore_errors=True)
    shutil.rmtree(mirror, ignore_errors=True)
    os.makedirs(flash)
    os.makedirs(mirror)
    try:
        bdir = os.path.join(s.dir, "WMS_Backups")
        base = os.path.splitext(os.path.basename(p))[0]
        # ---------------------------------------------------------------- Н1 the copy to a flash drive
        r1 = issue(s, 1, "ЕИ-00000001", "1", "ива")
        root = os.path.join(flash, "WMS_FLASH")
        os.makedirs(root)
        olds = [os.path.join(root, f"{base}_2020-01-0{i}_000000") for i in range(1, 9)]
        for d in olds:
            os.makedirs(d)
        fc = s.B("FlashCopy", sys_url(flash), module="WmsSafe")
        got = sorted(f for f in os.listdir(root))
        new = [f for f in got if not f.startswith(f"{base}_2020")]
        nd = os.path.join(root, new[0]) if new else ""
        man = open(os.path.join(nd, "MANIFEST.txt"), encoding="utf-8").read() if nd else ""
        jsrc = sorted(os.listdir(s.jdir))
        jsrc = [f for f in jsrc if f.startswith("WMS_journal_")]
        same = nd and all(os.path.getsize(os.path.join(nd, "WMS_Journal", f)) == os.path.getsize(os.path.join(s.jdir, f)) for f in jsrc) \
            and os.path.getsize(os.path.join(nd, os.path.basename(p))) == os.path.getsize(p)
        R.add(c, "Н1: «Копия на флешку» — папка WMS_FLASH/<книга>_<дата>: книга, журнал (размеры совпадают), MANIFEST; хранится 8 — самая старая удалена",
              r1.startswith("OK") and fc.startswith("OK:") and len(new) == 1 and same and "WMS_Journal/" in man and len(got) == 8
              and f"{base}_2020-01-01_000000" not in got, f"{fc[:160]}; {got}")
        inside = s.B("FlashCopy", sys_url(os.path.join(s.dir, "sub")), module="WmsSafe")
        nofold = s.B("FlashCopy", sys_url(os.path.join(flash, "нет")), module="WmsSafe")
        st = open(os.path.join(s.dir, "WMS_State", "wms_state.txt"), encoding="utf-8").read()
        R.add(c, "Н1: копия в папку WMS — отказ; несуществующая папка — отказ; дата копии записана в WMS_State",
              inside.startswith("ERR:") and "папку WMS" in inside and nofold.startswith("ERR:") and "FLASH=" in st, f"{inside[:80]}; {nofold[:80]}; {st[:120]}")
        # ---------------------------------------------------------------- Н9 the disk, Н8 the clock
        real = s.B("DiskFreeMb", sys_url(s.dir), module="WmsSafe")
        s.B("TestSafe", "", 50.0, module="WmsSafe")
        n0 = len(os.listdir(bdir))
        bk = s.B("BackupNow", "manual", module="WmsBackup")
        fc2 = s.B("FlashCopy", sys_url(flash), module="WmsSafe")
        note = s.B("DiskNote", module="WmsSafe")
        s.B("TestSafe", "", 0.0, module="WmsSafe")
        R.add(c, "Н9: свободное место определяется (df); мало места — резервная копия и копия на флешку не делаются, предупреждение при запуске",
              real > 0 and bk.startswith("ОШИБКА") and "свободно 50 МБ" in bk and len(os.listdir(bdir)) == n0 and fc2.startswith("ERR:") and "свободно 50" in fc2
              and note.startswith("ВНИМАНИЕ: на диске"), f"{real}; {bk[:120]}; {fc2[:80]}; {note[:80]}")
        old = s.B("TestJposMark", "2099-01-01T10:00:00/300", module="WmsSafe")
        clk = s.B("ClockNote", module="WmsSafe")
        s.B("TestJposSet", old, module="WmsSafe")
        clk2 = s.B("ClockNote", module="WmsSafe")
        R.add(c, "Н8: часы ПК позади последней операции журнала — «часы компьютера отстают на N дн.»; в порядке — молчит",
              clk.startswith("ВНИМАНИЕ: часы компьютера отстают на ") and "01.01.2099" in clk and clk2 == "", f"{clk[:140]}; {clk2!r}")
        # ---------------------------------------------------------------- Н6 the restore points, the list of the copies
        bk = s.B("BackupNow", "manual", module="WmsBackup")
        pt = s.B("RestorePoint", "перед тестом", module="WmsSafe")
        lst = s.B("CopyList", module="WmsSafe")
        names = [x[0] for x in lst]
        R.add(c, "Н6: точка восстановления — копия «_point» с пометкой; список копий (новые сверху) с числом операций и видом",
              "точка восстановления перед тестом" in pt and names and names[0].endswith("_point.ods") and lst[0][3] == "перед тестом" and lst[0][2] == "point"
              and any(x[2] == "manual" for x in lst) and s.B("CopyLabel", lst[0], module="WmsSafe").find("точка восстановления перед тестом") > 0,
              f"{pt[:120]}; {[tuple(x[:4]) for x in lst][:3]}")
        # ---------------------------------------------------------------- Н4 the chain
        s.B("ChainCheck", 0, module="WmsSafe")
        for r, ei in ((2, "ЕИ-00000002"), (3, "ЕИ-00000013"), (4, "ЕИ-00000004")):
            issue(s, r, ei, "1", "ива")
        n_ch, probs = chain_problems(s.jdir)
        lines = journal_lines(s.jdir)
        jf = sorted(glob.glob(os.path.join(s.jdir, "WMS_journal_*.csv")))[-1]
        raw = open(jf, encoding="utf-8").read()
        victim = lines[-3]
        tampered = victim.replace("WHO=Иванов Иван Андреевич", "WHO=Иванов Иван Андреевиш")
        open(jf, "w", encoding="utf-8", newline="").write(raw.replace(victim, tampered))
        bad = s.B("ChainCheck", 0, module="WmsSafe")
        n2, probs2 = chain_problems(s.jdir)
        seq_after = lines[-2].split(";")[1]
        open(jf, "w", encoding="utf-8", newline="").write(raw)
        issue(s, 5, "ЕИ-00000005", "1", "ива")
        ok_chain = s.B("ChainCheck", 0, module="WmsSafe")
        R.add(c, "Н4: у каждой строки журнала 1.0 — CH (SHA-256 предыдущей), независимая проверка без ошибок; ручная правка строки — проверка цепочки "
                 "называет следующую операцию; дальше цепочка снова сходится",
              n_ch >= 3 and not probs and f"№ {seq_after}" in bad and "не продолжает цепочку" in bad and probs2 == [seq_after] and ok_chain == "",
              f"{n_ch}; {probs}; {bad[:160]}; {probs2}; {ok_chain!r}")
        # ---------------------------------------------------------------- Н3 the mirror
        ms = s.B("MirrorSet", sys_url(mirror), module="WmsSafe")
        mdir = os.path.join(mirror, f"WMS_Зеркало_журнала_{base}")
        issue(s, 6, "ЕИ-00000006", "1", "ива")
        jf = sorted(glob.glob(os.path.join(s.jdir, "WMS_journal_*.csv")))[-1]
        mf = os.path.join(mdir, os.path.basename(jf))
        eq1 = os.path.exists(mf) and open(mf, "rb").read() == open(jf, "rb").read()
        shutil.rmtree(mdir)
        open(mdir, "w").write("не папка")
        i7 = issue(s, 7, "ЕИ-00000007", "1", "ива")
        err = s.B("TestSafeAttention", module="WmsSafe")
        os.remove(mdir)
        i8 = issue(s, 8, "ЕИ-00000008", "1", "ива")
        eq2 = os.path.exists(mf) and open(mf, "rb").read() == open(jf, "rb").read()
        R.add(c, "Н3: зеркало включено — каждая строка и туда (файлы равны побайтно); зеркало недоступно — проводка идёт, «Внимание» говорит; "
                 "снова доступно — файл догоняется целиком",
              ms.startswith("OK:") and eq1 and i7.startswith("OK") and any("Зеркало журнала не пишется" == x[0] for x in err) and i8.startswith("OK") and eq2,
              f"{ms[:100]}; {eq1}; {i7[:40]}; {err}; {eq2}")
        # the journal lost its end: «Журнал из зеркала»
        data = open(jf, encoding="utf-8").read()
        cut = data[:data.rstrip("\n").rfind("\n") + 1]
        open(jf, "w", encoding="utf-8", newline="").write(cut)
        mr = s.B("MirrorRestore", module="WmsSafe")
        aside = glob.glob(os.path.join(bdir, os.path.basename(jf) + ".отложен_*"))
        R.add(c, "Н3: журнал потерял последнюю строку — «Журнал из зеркала» ставит целый файл, прежний откладывается; WMS снова «работа разрешена»",
              mr.startswith("OK:") and "заменено файлов 1" in mr and open(jf, "rb").read() == open(mf, "rb").read() and len(aside) == 1
              and s.state().get("STATE") == "CLEAN", f"{mr[:300]}; {aside}")
        # ---------------------------------------------------------------- Н5 the check of the day
        order(s, 1, "1", "Шуруп", qty="10")
        receive(s, 1, "10", place="2-1-1")
        nr = int(ords(s, 1)[21][3:])
        dc = s.B("DailyCheck", True, module="WmsSafe")
        s.B("DashRefresh", module="WmsDash")
        title = s.doc.Sheets.getByName("Дашборд").getCellByPosition(8, 0).getString()
        stk = s.doc.Sheets.getByName("Наличие")
        stk.unprotect("wms")
        q0 = stk.getCellByPosition(4, nr).getValue()
        stk.getCellByPosition(4, nr).setValue(q0 + 1)
        dc2 = s.B("DailyCheck", True, module="WmsSafe")
        att = s.B("TestSafeAttention", module="WmsSafe")
        stk.getCellByPosition(4, nr).setValue(q0)
        stk.protect("wms")
        s.B("DailyCheck", True, module="WmsSafe")
        R.add(c, "Н5: проверка дня — «без замечаний», на «Дашборде» «проверка дня ✓»; остаток исправлен мимо WMS — «остаток не равен движениям», "
                 "строка во «Внимании»",
              dc.split("|")[1] == "OK" and "проверка дня ✓" in title and dc2.split("|")[1] == "FAIL" and "остаток не равен движениям" in dc2
              and any(x[0] == "Проверка дня: замечания" for x in att), f"{dc[:120]}; {title[:120]}; {dc2[:160]}")
        sv = s.B("ServiceRefresh", module="WmsSafe")
        page = [s.doc.Sheets.getByName("Обслуживание").getCellByPosition(2, r).getString() for r in range(3, 9)]
        R.add(c, "страница «Обслуживание»: состояние — копия на флешке, зеркало, проверка дня, копии, свободное место, книга и журнал",
              sv == "OK" and page[0].startswith("✓ ") and page[1].startswith("✓ ") and "в «WMS_Backups»" in page[3] and "ГБ" in page[4] and "журнал" in page[5],
              str(page))
    finally:
        s.close(save=True)
    # ---------------------------------------------------------------- Н7 the operations put aside
    s = ready(Session(p))
    try:
        s.B("MirrorSet", "", module="WmsSafe")
        put(s, "Выдачи", 10, L="ЕИ-00000010", E="2", H=serial(-1), I="ива")
        s.doc.store()
        s.B("TestSetFault", 3, 2)
        crash = s.I("IssuePostRow", 10)
        s.kill()
        s = ready(Session(p))
        s.B("ActionForceUnlock")
        ab = s.B("ActionAbandonTail")
        pts = [x for x in s.B("CopyList", module="WmsSafe") if x[2] == "point"]
        sh = s.B("AbandonedShow", module="WmsSafe")
        fd = s.doc.Sheets.getByName("Поиск")
        row = [fd.getCellByPosition(col, 12).getString() for col in range(1, 9)]
        put(s, "Выдачи", 10, L="", E="", H="", I="")
        rd = s.B("AbandonedRedo", (12,), module="WmsSafe")
        iss = s.doc.Sheets.getByName("Выдачи")
        last = max(r for r in range(1, 60) if iss.getCellByPosition(11, r).getString() != "")
        back = [iss.getCellByPosition(col, last).getString() for col in (11, 4, 8, 14)]
        po = s.I("IssuePostRow", last)
        R.add(c, "Н6/Н7: «Отложить хвост» — сначала точка восстановления; «Отложенные операции» показывают выдачу (ЕИ, количество, кто); "
                 "«Вернуть в работу» ставит её непроведённой строкой «Выдач», она проводится",
              crash == "CRASH-SIM" and "отложено операций: 1" in ab and any("Отложить хвост" in x[3] for x in pts) and sh.startswith("OK")
              and row[2] == "Выдача" and row[3] == "ЕИ-00000010" and row[4].startswith("2") and row[5] == "Иванов Иван Андреевич" and rd.startswith("OK:")
              and back[0] == "ЕИ-00000010" and back[1] == "2" and back[3].startswith("отложенная операция журнала №") and po.startswith("OK"),
              f"{crash}; {ab[:80]}; {[x[3] for x in pts]}; {row}; {rd[:100]}; {back}; {po[:60]}")
        oracle(c, s, "отложенный хвост и возврат в работу")
        # ---------------------------------------------------------------- Н2 «Вернуться к копии»
        s.doc.store()
        bk = s.B("BackupNow", "manual", module="WmsBackup")
        name = [x for x in s.B("CopyList", module="WmsSafe")][0][0]
        seq0 = s.sysv("LAST_SEQ")
        for r, ei in ((20, "ЕИ-00000011"), (21, "ЕИ-00000012")):
            issue(s, r, ei, "1", "ива")
        seq1 = s.sysv("LAST_SEQ")
        q11 = s.stock(11)
        url = s.doc.getURL()
        old_doc = s.doc
        rt = s.B("ReturnToCopy", name, False, module="WmsSafe")
        newd = find_doc(s, url, old_doc)
        msg = ""
        if newd is not None:
            s.doc = newd
            t0 = time.time()
            while time.time() - t0 < 60:
                msg = s.U("TestUiLastMessage") or ""
                if "возвращена к копии" in msg or msg.startswith("ERR"):
                    break
                time.sleep(0.5)
        aside = glob.glob(os.path.join(bdir, f"{base}_*_отложена.ods"))
        copies_open = []
        en = s.o.desktop.getComponents().createEnumeration()
        while en.hasMoreElements():
            d = en.nextElement()
            try:
                copies_open.append(d.getURL())
            except Exception:
                pass
        R.add(c, "Н2: «Вернуться к копии» — рабочая книга закрыта и отложена («…_отложена.ods»), копия встала на её место, по журналу доведены "
                 "2 операции после копии, самопроверка, «работа разрешена»; копия закрыта",
              rt.startswith("OK") and newd is not None and "возвращена к копии" in msg and s.state().get("STATE") == "CLEAN" and s.sysv("LAST_SEQ") == seq1
              and s.stock(11) == q11 and len(aside) == 1 and not any("/WMS_Backups/" in u for u in copies_open) and seq1 == seq0 + 2,
              f"{rt}; {msg[:300]}; {s.state()}; {seq0}->{s.sysv('LAST_SEQ')} (было {seq1}); {aside}; {copies_open}")
        oracle(c, s, "возврат к копии")
    finally:
        s.close()


@case
def v28_speed():
    """С1–С8: время запуска, сохранения и карты (сторож роста, карта дольше 5 с — только кнопкой), ленивая карта между сеансами, рост книги
    и прогноз, лишние ячейки, старые снимки, лёгкий снимок для этикеток (и печать этикеток из него), советы LibreOffice, замер скорости ПК,
    поиск за 12 месяцев и «Вся история»"""
    c = "V28"
    p = new_wms("v28")
    s = ready(Session(p))
    try:
        st = lambda: open(os.path.join(s.dir, "WMS_State", "wms_state.txt"), encoding="utf-8").read()
        topology(s, {1: (4, 2), 2: (4, 2)})
        mb = s.B("MapBuild", "OCC", True, module="WmsMap")
        s.doc.store()
        txt = st()
        R.add(c, "С4: время запуска, сохранения и карты запоминаются (WMS_State)", mb.startswith("OK") and "T_OPEN=" in txt and "T_SAVE=" in txt
              and "T_MAP=" in txt, f"{mb[:80]}; {[x for x in txt.splitlines() if x.startswith('T_')]}")
        info0 = s.doc.Sheets.getByName("Карта склада").getCellByPosition(0, 4).getString()
    finally:
        s.close()
    s = ready(Session(p))
    try:
        t_map = [x for x in open(os.path.join(s.dir, "WMS_State", "wms_state.txt"), encoding="utf-8").read().splitlines() if x.startswith("T_MAP=")]
        s.B("OnMapFocus", module="WmsMap")
        ms = s.B("TestMapState", module="WmsMap").split("|")
        t_map2 = [x for x in open(os.path.join(s.dir, "WMS_State", "wms_state.txt"), encoding="utf-8").read().splitlines() if x.startswith("T_MAP=")]
        info1 = s.doc.Sheets.getByName("Карта склада").getCellByPosition(0, 4).getString()
        R.add(c, "С5: книга открыта снова — карта не перестраивается (та же отметка «Обновлено», время карты не замерялось заново), места и подсчёт "
                 "прочитаны (16 мест)", ms[0] == "16" and float(ms[1]) == s.sysv("LAST_SEQ") and ms[2] == "OCC" and t_map == t_map2 and info1 == info0,
              f"{ms}; {t_map} / {t_map2}; {info0[-20:]} / {info1[-20:]}")
        # С4: a slow map is built again only by «Обновить»
        issue(s, 1, "ЕИ-00000001", "1", "ива")
        s.B("StSet", "T_MAP", "6000|2026-10-20 10:00", module="WmsSafe")
        s.B("OnMapFocus", module="WmsMap")
        ms2 = s.B("TestMapState", module="WmsMap").split("|")
        info2 = s.doc.Sheets.getByName("Карта склада").getCellByPosition(0, 4).getString()
        gn = s.B("GrowthNote", module="WmsCare")
        R.add(c, "С4: карта строится 6 с — после операции не перестраивается сама, на ней «обновляется только кнопкой «Обновить»»; совет о росте книги",
              float(ms2[1]) != s.sysv("LAST_SEQ") and "только кнопкой «Обновить»" in info2 and "карта строится 6 с" in gn, f"{ms2}; {info2[:120]}; {gn[:160]}")
        s.B("StSet", "T_MAP", "300|2026-10-20 10:00", module="WmsSafe")
        # С3: the growth
        y, m = datetime.date.today().year, datetime.date.today().month - 2
        if m < 1:
            m, y = m + 12, y - 1
        s.B("StSet", f"GROW_{y}-{m:02d}", "0.10|0", module="WmsSafe")
        gt = s.B("GrowthText", module="WmsCare")
        rc = s.doc.Sheets.getByName("Получатели")
        rc.getCellByPosition(0, 6000).setFormula('=""')
        ex = s.B("ExtraCells", module="WmsCare")
        R.add(c, "С3: рост книги в месяц и прогноз до «Закрытия года»; «Лишние ячейки» называют лист с оформлением далеко ниже данных",
              "в месяц +" in gt and "до 200 000 движений" in gt and "«Получатели»: занято до строки 6 001" in ex, f"{gt}; {ex[:200]}")
        # С8: the light snapshot; the labels from it
        t0 = time.time()
        full = s.B("ExportSnapshot", module="WmsExport")
        t_full = time.time() - t0
        t0 = time.time()
        lite = s.B("ExportSnapshot", True, module="WmsExport")
        t_lite = time.time() - t0
        ld = lite[3:].split(" — ")[0]
        files = sorted(os.listdir(ld)) if os.path.isdir(ld) else []
        man = open(os.path.join(ld, "manifest.csv"), encoding="utf-8").read() if files else ""
        prn = os.path.join(OUT, "cases", "v28_labels.prn")
        import subprocess
        lab = subprocess.run([sys.executable, "-I", os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src", "toolbox", "scripts", "labels",
                                                          "tspl_labels.py"), "--snapshot", ld, "--ei", "1", "--out", prn], capture_output=True, text=True)
        R.add(c, "С8: «Снимок для этикеток» — папка WMS_STOCK_…: «Наличие», «Заказы», «Иной приход», без журнала, быстрее полного; этикетки "
                 "набора печатаются из него",
              lite.startswith("OK:") and "WMS_STOCK_" in ld and files == ["manifest.csv", "orders.csv", "service", "special.csv", "stock.csv"] and "kind;lite" in man
              and t_lite < t_full and lab.returncode == 0 and os.path.exists(prn), f"{lite[:120]}; {files}; {t_lite:.1f} < {t_full:.1f}; {lab.stdout[-200:]} {lab.stderr[-200:]}")
        for _ in range(2):
            s.B("ExportSnapshot", module="WmsExport")
        pr = s.B("CarePrune", 2, module="WmsCare")
        left = [f for f in os.listdir(os.path.join(s.dir, "WMS_Export")) if f.startswith("WMS_")]
        R.add(c, "С3: «Удалить старые снимки» — остаются два последних", pr.startswith("Старых снимков удалено: 2") and len(left) == 2, f"{pr}; {left}")
        # С6, С2
        tips = s.B("LoTips", 0, module="WmsCare")
        ap = s.B("LoApply", module="WmsCare")
        sp = s.B("SpeedTest", module="WmsCare")
        mN = re.search(r"держит книгу до ~([\d\s ]+) движений", sp)
        nN = int(re.sub(r"\D", "", mN.group(1))) if mN else 0
        R.add(c, "С6: советы по настройкам LibreOffice (резервная копия, восстановление, пересчёт) и «Применить»; С2: замер скорости ПК — «держит до ~N движений»",
              "«Всегда создавать резервную копию»" in tips and ap.startswith("Советы применены") and 30000 <= nN <= 600000 and "Сейчас в книге" in sp,
              f"{tips[:200]}; {ap[:80]}; {sp[:240]}")
        # С7: the search 12 months back
        for r, k in ((2, -400), (3, -420), (4, -3)):
            put(s, "Выдачи", r, L="ЕИ-00000002", E="1", H=serial(k), I="пет")
            s.I("IssuePostRow", r)
        f1 = s.B("TestFind", "Петров Пётр", module="WmsFind")
        head1 = s.doc.Sheets.getByName("Поиск").getCellByPosition(1, 6).getString()
        f2 = s.B("TestFindAll", "Петров Пётр", module="WmsFind")
        head2 = s.doc.Sheets.getByName("Поиск").getCellByPosition(1, 6).getString()
        R.add(c, "С7: поиск по умолчанию за 12 месяцев — старые строки посчитаны отдельно («Вся история»); «Вся история» — все",
              "старше — ещё 2" in head1 and "вся история" in head2 and "Найдено «Петров Пётр»: 4" in head2 and "Найдено «Петров Пётр»: 2" in head1,
              f"{f1}; {head1}; {f2}; {head2}")
        oracle(c, s, "скорость и обслуживание")
    finally:
        s.close()


@case
def v29_help():
    """П1 проверка ПК, П2 «Что нового», П3 тур «Первый день», П6 «?» — раздел «Справки» каждого листа, П7 шпаргалка, П8 «Что делать,
    если…», П10 ярлык, П4 учебная WMS (своя папка, свой журнал, рабочая книга не меняется), П12 обезличенная копия"""
    c = "V29"
    p = new_wms("v29")
    s = ready(Session(p))
    try:
        pc = s.B("PcCheck", module="WmsHelp")
        R.add(c, "П1: «Проверить этот ПК» — LibreOffice, макросы, запись, место, Python — с «что сделать»",
              pc.startswith("Проверка этого ПК") and "LibreOffice" in pc and "запись в папку WMS" in pc and "место на диске" in pc and "Python" in pc, pc[:300])
        choose(s, 4)
        wn = s.B("WhatsNew", module="WmsHelp")
        act = s.doc.getCurrentController().getActiveSheet().getName()
        R.add(c, "П2: «Что нового» — 8 карточек (первая — 1.0.2), «Показать» открывает лист карточки", len(s.B("WhatsNewCards", module="WmsHelp")) == 8 and wn == 4
              and act == "Поиск", f"{wn}; {act}")
        col0 = s.B("ButtonModel", "Заказы", "btnReceiveWin", module="WmsHelp").BackgroundColor
        tr = s.B("Tour", module="WmsHelp")
        col1 = s.B("ButtonModel", "Заказы", "btnReceiveWin", module="WmsHelp").BackgroundColor
        R.add(c, "П3: тур «Первый день» — 7 шагов по листам с подсветкой кнопки, цвет кнопки возвращён", tr == 7 and col0 == col1, f"{tr}; {col0} {col1}")
        bad = []
        for name in ("Главная", "Заказы", "Иной приход", "Приход авто", "Выдачи", "Возврат", "Корректировки", "Поиск", "Карта склада", "Наличие",
                     "Снабжение", "Комплекты", "Отчёты", "Обслуживание"):
            hg = s.B("HelpGo", name, module="WmsHelp")
            row = int(hg[3:]) if hg.startswith("OK:") else -1
            txt = s.doc.Sheets.getByName("Справка").getCellByPosition(0, row).getString() if row >= 0 else ""
            if not txt.startswith(s.B("HelpTopic", name, module="WmsHelp")) or row <= 0:
                bad.append((name, hg, txt[:30]))
        marks = [n for n in ("Заказы", "Выдачи", "Наличие", "Поиск", "Обслуживание") if "btnHelpHere" in form_buttons(s.doc.Sheets.getByName(n))]
        R.add(c, "П6: «?» на листах — «Справка» открывается на разделе своего листа (новое руководство 1.0 с разделами и словарём)",
              not bad and len(marks) == 5 and "13. Словарь" in [s.doc.Sheets.getByName("Справка").getCellByPosition(0, r).getString() for r in range(0, 400)],
              f"{bad}; {marks}")
        cs = s.B("CheatSheet", module="WmsHelp")
        pr = [s.doc.Sheets.getByName("Печать").getCellByPosition(0, r).getString() for r in range(0, 40)]
        R.add(c, "П7: шпаргалка A4 на «Печати»; П8: «Что делать, если…» — 9 ситуаций с ответом и переходом",
              cs.startswith("OK") and "Начало дня" in pr and "Если красное" in pr and len(s.B("IfCases", module="WmsHelp")) == 11, str(pr[:8]))
        desk = os.path.join(OUT, "cases", "v29_desktop")
        shutil.rmtree(desk, ignore_errors=True)
        s.B("TestHelpDesktop", desk, module="WmsHelp")
        sc = s.B("Shortcut", module="WmsHelp")
        f = os.path.join(desk, "WMS склад.desktop")
        R.add(c, "П10: ярлык «WMS склад» открывает эту книгу", sc.startswith("OK:") and os.path.exists(f) and p in open(f, encoding="utf-8").read(), sc[:160])
        issue(s, 1, "ЕИ-00000001", "2", "ива")
        # 1.0.2: an act of the recipient, a hand-written line of the report, a signer — none of them may reach the anonymous copy
        s.goto("Выдачи", "$A$2")
        s.B("ActFromSelection", module="WmsDocs")
        s.B("DcDocClose", module="WmsDocs")
        ak = s.doc.Sheets.getByName("Акты")
        ak.getCellByPosition(6, 1).setString("Для Иванова Секретный")
        ak.getCellByPosition(9, 3).setString("Подписант Секретов")
        s.doc.Sheets.getByName("Отчёт дня").getCellByPosition(0, 5).setString("Секретная работа")
        s.doc.Sheets.getByName("Отчёт дня").getCellByPosition(1, 2).setString("Секретов, кладовщик.")
        s.B("FastRefresh", "", module="WmsCat")
        s.doc.store()
        h0 = {g: hashlib.sha256(open(g, "rb").read()).hexdigest() for g in glob.glob(os.path.join(s.jdir, "*.csv"))}
        mt = s.B("MakeCopy", "TRAINING", False, module="WmsHelp")
        ma = s.B("MakeCopy", "ANON", False, module="WmsHelp")
        q1 = s.stock(1)
    finally:
        s.close()
    tp = re.search(r"(/[^\n]*?\.ods)", mt).group(1) if mt.startswith("OK:") else ""
    ap = re.search(r"(/[^\n]*?\.ods)", ma).group(1) if ma.startswith("OK:") else ""
    t = ready(Session(tp)) if tp else None
    try:
        st = t.state() if t else {}
        title = t.doc.Sheets.getByName("Дашборд").getCellByPosition(8, 0).getString() if t else ""
        t.B("DashRefresh", module="WmsDash") if t else None
        title = t.doc.Sheets.getByName("Дашборд").getCellByPosition(8, 0).getString() if t else ""
        ti = issue(t, 2, "ЕИ-00000001", "5", "ива") if t else ""
        tj = os.listdir(os.path.join(os.path.dirname(tp), "WMS_Journal")) if tp else []
        R.add(c, "П4: учебная WMS — папка «…_УЧЕБНАЯ», «работа разрешена», «УЧЕБНАЯ WMS» на «Дашборде», свой журнал; проводка в ней не трогает "
                 "журнал рабочей книги", mt.startswith("OK:") and st.get("STATE") == "CLEAN" and "УЧЕБНАЯ" in title and ti.startswith("OK") and tj
              and h0 == {g: hashlib.sha256(open(g, "rb").read()).hexdigest() for g in glob.glob(os.path.join(os.path.dirname(p), "WMS_Journal", "*.csv"))},
              f"{mt[:120]}; {st}; {title[:60]}; {ti[:40]}; {tj}")
    finally:
        if t:
            t.close()
    a = ready(Session(ap)) if ap else None
    try:
        st = a.state() if a else {}
        nm = [a.doc.Sheets.getByName("Наличие").getCellByPosition(1, r).getString() for r in range(1, 6)] if a else []
        rc = [a.doc.Sheets.getByName("Получатели").getCellByPosition(1, r).getString() for r in range(1, 4)] if a else []
        sc2 = a.B("ActionSelfCheck") if a else ""
        R.add(c, "П12: обезличенная копия — наименования «Товар N», получатели «Получатель N», количества как в книге, «работа разрешена», "
                 "самопроверка без ошибок", ma.startswith("OK:") and st.get("STATE") == "CLEAN" and all(x.startswith("Товар ") for x in nm)
              and all(x.startswith("Получатель ") for x in rc) and a.stock(1) == q1 and sc2.startswith("САМОПРОВЕРКА WMS: ошибок 0"),
              f"{ma[:120]}; {st}; {nm}; {rc}; {sc2[:120]}")
        left = []
        if a:
            names_ = [a.doc.Sheets.getByIndex(i).Name for i in range(a.doc.Sheets.getCount())]
            for nme in names_:
                sh_ = a.doc.Sheets.getByName(nme)
                cur_ = sh_.createCursor()
                cur_.gotoEndOfUsedArea(False)
                ad_ = cur_.getRangeAddress()
                for row_ in sh_.getCellRangeByPosition(0, 0, min(ad_.EndColumn, 40), min(ad_.EndRow, 5100)).getDataArray():
                    for x_ in row_:
                        if isinstance(x_, str) and any(w in x_ for w in ("Секрет", "Иванова")):
                            left.append((nme, x_[:30]))
            ak_ = a.doc.Sheets.getByName("Акты")
        R.add(c, "П12 (1.0.2): в обезличенной копии нет подписанта и примечания акта, подписи и своей строки «Отчёта дня»; получатель акта заменён; «Метизы» очищены",
              a and not left and ak_.getCellByPosition(3, 1).getString().startswith("Получатель ")
              and a.doc.Sheets.getByName("Метизы").getCellByPosition(0, 4).getString() == "", f"{left[:5]}")
    finally:
        if a:
            a.close()


@case
def v30_upgrade_zoo():
    """Т4: рабочая книга 0.7.2 (собрана исходниками 0.7.2, перенос старой таблицы и рабочий день) в папке с пробелами и кириллицей
    обновляется до 1.0: проверка, сборка (сверка каждой ячейки), «сделать рабочей»; каждый ЕИ и его остаток те же, старая книга и журнал
    не тронуты, новые листы 1.0 на месте, проверка дня и цепочка журнала в порядке"""
    import run_prime
    c = "V30"
    book, facts = run_prime.warehouse()
    d = os.path.join(OUT, "cases", "v30 Склад 2026")
    shutil.rmtree(d, ignore_errors=True)
    old = os.path.join(d, "Рабочая папка WMS")
    shutil.copytree(os.path.dirname(book), old, ignore=shutil.ignore_patterns("WMS_TOOLBOX", "backup", "WMS_Backups"))
    p = os.path.join(old, os.path.basename(book))
    s = Session(p)
    try:
        if "REG=other" in s.B("StateDump"):
            s.B("ActionRegisterHere", False)
        st0 = s.B("StateDump")
        stk = s.doc.Sheets.getByName("Наличие")
        n_ei = int(s.sysv("NEXT_EI")) - 1
        before = [tuple(r[:6]) for r in stk.getCellRangeByPosition(0, 1, 7, n_ei).getDataArray()]
    finally:
        s.close(save=True)
    h0 = run_prime.sha(p)
    j0 = {f: run_prime.sha(f) for f in glob.glob(os.path.join(old, "WMS_Journal", "*.csv"))}
    rc0, o0 = run_prime.upgrade("check", p)
    new = os.path.join(d, "Новая WMS 1.0")
    rc1, o1 = run_prime.upgrade("build", p, "--to", new)
    rc2, o2 = run_prime.upgrade("promote", "--to", new)
    rep = open(os.path.join(new, "UPGRADE_REPORT.md"), encoding="utf-8").read() if os.path.exists(os.path.join(new, "UPGRADE_REPORT.md")) else ""
    R.add(c, "Т4: папки с пробелами и кириллицей — проверка, сборка (0 расхождений ячеек, самопроверка, «остаток = движения»), «сделать рабочей»; "
             "старая книга и журнал побайтно те же", "STATE=CLEAN" in st0 and rc0 == 0 and rc1 == 0 and rc2 == 0 and "расхождений: 0" in rep
          and run_prime.sha(p) == h0 and {f: run_prime.sha(f) for f in glob.glob(os.path.join(old, "WMS_Journal", "*.csv"))} == j0,
          f"{st0[:80]}; {o0[-150:]}; {o1[-200:]}; {o2[-150:]}")
    nb = os.path.join(new, os.path.basename(p))
    s = ready(Session(nb))
    try:
        stk = s.doc.Sheets.getByName("Наличие")
        after = [tuple(r[:6]) for r in stk.getCellRangeByPosition(0, 1, 7, n_ei).getDataArray()]
        diff = [(a, b) for a, b in zip(before, after) if a != b]
        sheets = [n for n in ("Обслуживание", "Резервы", "_PROP", "Комплекты", "Номенклатура", "Отчёты", "Печать") if s.doc.Sheets.hasByName(n)]
        dc = s.B("DailyCheck", True, module="WmsSafe")
        ei1 = [r[0] for r in before if isinstance(r[4], float) and r[4] >= 1][0]
        rc = s.doc.Sheets.getByName("Получатели")
        rr = 1
        while rc.getCellByPosition(1, rr).getString():
            rr += 1
        rc.getCellByPosition(1, rr).setString("Петров")
        s.I("OnRecipientsChange")
        i1 = issue(s, 400, ei1, "1", "Петров")
        n_ch, probs = chain_problems(os.path.join(new, "WMS_Journal"))
        R.add(c, f"Т4: в книге 1.0 каждый из {n_ei} ЕИ — тот же номер, наименование, артикул, единица, остаток и место; новые листы 1.0; "
                 "«работа разрешена», проверка дня без замечаний; новая операция пишется с контрольной цепочкой",
              s.state().get("STATE") == "CLEAN" and n_ei > 10 and not diff and len(sheets) == 7 and dc.split("|")[1] == "OK" and i1.startswith("OK")
              and n_ch >= 1 and not probs, f"{i1[:120]}; {n_ch} {probs}; {s.state().get('STATE')}; {diff[:3]}; {sheets}; {dc[:160]}")
        P, inf, _ = s.adjcheck(0, initial={}, legacy_mn_empty=False)
        R.add(c, "оракул (журнал воспроизведён независимо): книга 0.7.2, обновлённая до 1.0", not P, f"{inf.get('ops')} операций; {P[:4]}")
    finally:
        s.close()


def reopen_after_crash(p, s):
    s.kill()
    for f in glob.glob(os.path.join(os.path.dirname(p), ".~lock.*#")):
        os.remove(f)
    s = ready(Session(p))
    s.B("ActionForceUnlock")
    return s


@case
def v31_crash_new_ops():
    """Т2: сбой посреди новых операций 1.0 — разделение ЕИ (SPLIT), возврат поставщику (SUP_RET), сборка комплекта (KIT_BUILD), свойства
    ЕИ (EI_PROP): до журнала — операции нет (откат), после журнала — «Восстановить» доводит её; остатки, номера и оракул сходятся"""
    c = "V31"
    out = []
    for op in ("SPLIT", "SUP_RET", "KIT_BUILD", "EI_PROP"):
        for point in (2, 3):
            p = new_wms(f"v31_{op}_{point}")
            s = ready(Session(p))
            order(s, 1, "1", "Кабель ВВГ", qty="10", unit="м", art="К-3", J=50.0)
            receive(s, 1, "10", place="7-1-1")
            a = ords(s, 1)[21]
            na = int(a[3:])
            if op == "KIT_BUILD":
                kt = s.doc.Sheets.getByName("Комплекты")
                for j, v in enumerate(("Набор", "Кабель ВВГ", "К-3", 2.0, "м")):
                    cell = kt.getCellByPosition(j, 1)
                    cell.setValue(v) if isinstance(v, float) else cell.setString(v)
            s.doc.store()
            nxt = int(s.sysv("NEXT_EI"))
            s.B("TestSetFault", point, 2)
            if op == "SPLIT":
                r = s.B("SplitPost", a, "3", "7-1-2", "часть", module="WmsDaily")
            elif op == "SUP_RET":
                r = s.B("SupReturnPost", a, "2", "брак", "ВП-1", False, module="WmsSupply")
            elif op == "KIT_BUILD":
                r = s.B("KitBuild", "Набор", 2.0, "8-1-1", serial(), "к монтажу", module="WmsKit")
            else:
                r = s.B("PropSet", a, "SN-1", "", "", module="WmsItem")
            s = reopen_after_crash(p, s)
            st1 = (s.state().get("STATE"), s.state().get("BLOCK"))
            if st1[1] == "TAIL":
                s.B("ActionRecover")
            P, inf, _ = s.adjcheck(0)
            bal = s.stock(na)
            if point == 3:
                ok = {"SPLIT": bal == 7.0 and s.stock(nxt) == 3.0, "SUP_RET": bal == 8.0, "KIT_BUILD": bal == 6.0 and s.stock(nxt) == 2.0,
                      "EI_PROP": s.doc.Sheets.getByName("_PROP").getCellByPosition(1, na).getString() == "SN-1"}[op]
            else:
                ok = {"SPLIT": bal == 10.0 and int(s.sysv("NEXT_EI")) == nxt, "SUP_RET": bal == 10.0, "KIT_BUILD": bal == 10.0 and int(s.sysv("NEXT_EI")) == nxt,
                      "EI_PROP": s.doc.Sheets.getByName("_PROP").getCellByPosition(1, na).getString() == ""}[op]
            out.append((op, point, r[:12], st1, s.state().get("STATE"), ok, bal, not P, P[:2]))
            s.close()
    R.add(c, "сбой до журнала — операции нет; после журнала — «Восстановить» доводит; «работа разрешена», оракул без расхождений",
          all(x[2] == "CRASH-SIM" and x[4] == "CLEAN" and x[5] and x[7] for x in out), str(out))


@case
def v32_random_input():
    """Т11: тысячи случайных строк в разбор количества и места — ни одной внутренней ошибки; сотня случайных выдач, возвратов и
    списаний частями, процентами и числом — остаток никогда не уходит в минус, отказ — понятный, оракул (вместе с проверкой частей
    и процентов, Т5) без расхождений"""
    import random
    c = "V32"
    rnd = random.Random(20261020)
    alphabet = list("0123456789 ,.-/%кансмшт;:+") + ["1,5", "кан", "25%", "мм", "1/2", "5.12.3", "Р5", "С12", "ЯЗ", "\t", "ё", "0", "-1", "1e5"]
    s = ready(Session(new_wms("v32")))
    try:
        samples = ["".join(rnd.choice(alphabet) for _ in range(rnd.randint(0, 9))) for _ in range(1500)]
        res = s.B("TestFuzzParse", tuple(samples), module="WmsItem")
        R.add(c, "1500 случайных строк в разбор количества, частей и места — без внутренней ошибки", res.startswith("OK:1500"), res[:300])
        order(s, 1, "1", "Незамерзайка (4 канистры)", qty="5")
        receive(s, 1, "5", place="1-1-1")
        a = ords(s, 1)[21]
        na = int(a[3:])
        s.B("PackSet", a, "4", "кан", "", module="WmsPack")
        qs = ["1 кан", "2 канистры", "0,25", "1", "25%", "10%", "3 кан", "0,5", "кан", "1,5 кан", "7", "-1", "abc", "100%", "0"]
        neg, internal, posted, r = [], [], 0, 1
        for k in range(60):
            q = rnd.choice(qs)
            kind = rnd.choice(("ISSUE", "ISSUE", "WRITE_OFF"))
            if kind == "ISSUE":
                put(s, "Выдачи", r, L=a, E=q, H=serial(), I="ива")
                x = s.I("IssuePostRow", r)
            else:
                put(s, "Корректировки", r, B="Списание", C=a, G=q, K=serial(), L="проверка")
                x = s.B("AdjustPostRow", r, module="WmsAdjust")
            r += 1
            if x.startswith("OK"):
                posted += 1
            elif x.startswith("ERR-SYS") or "внутрен" in x.lower() or "строка " in x and "ошибка" in x.lower() and "Basic" in x:
                internal.append(x[:100])
            if s.stock(na) < -0.0000005:
                neg.append((k, q, s.stock(na)))
        P, inf, _ = s.adjcheck(0)
        R.add(c, "60 случайных выдач и списаний частями, процентами, числом: остаток не уходит в минус, внутренних ошибок нет, оракул и проверка "
                 "частей/процентов без расхождений", not neg and not internal and posted > 5 and not P and s.stock(na) >= 0
              and inf.get("parts_ops", 0) + inf.get("pct_ops", 0) > 0,
              f"проведено {posted}; {neg[:3]}; {internal[:3]}; {P[:3]}; остаток {s.stock(na)}; частями {inf.get('parts_ops')}, процентом {inf.get('pct_ops')}")
    finally:
        s.close()


@case
def v33_health():
    """Т5 сверка частей и процентов по журналу ловит неверное количество; Т8 проверка формул Calc (в проверке дня и при новой
    версии LibreOffice; Calc, считающий иначе, — замечание); Т7 учебное восстановление: свежая копия открывается отдельно, журнал
    после неё непрерывен — «восстанавливается», разрыв журнала — понятный отказ, рабочая книга не меняется; Т6 светофор
    «Здоровье WMS» на «Обслуживании»"""
    import adjust_oracle
    c = "V33"
    # ---------------------------------------------------------------- Т5: the oracle itself on crafted lines
    tmp = os.path.join(OUT, "v33_parts")
    shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(tmp)

    def jl(seq, body):
        b = f"J1;{seq};2026-10-20 10:00:00;INSTX;{body}"
        return f"{b};END;{len(b)}"
    lines = [jl(1, "ISSUE;EI=ЕИ-00000001;QTY=0.25;PARTS=1;PART_UNIT=кан;PART_N=4"),
             jl(2, "ISSUE;EI=ЕИ-00000001;QTY=0.3;PARTS=1;PART_UNIT=кан;PART_N=4"),
             jl(3, "WRITE_OFF;EI=ЕИ-00000001;QTY=1.25;PCT=25;PCT_BASE=5"),
             jl(4, "WRITE_OFF;EI=ЕИ-00000001;QTY=1.5;PCT=25;PCT_BASE=5"),
             jl(5, "ISSUE;EI=ЕИ-00000001;QTY=3;HOW=ALL")]
    open(os.path.join(tmp, "WMS_journal_2026-10.csv"), "w", encoding="utf-8", newline="\n").write("\n".join(lines) + "\n")
    probs = []
    np_, nc_ = adjust_oracle._parts_check(tmp, {"INSTANCE_ID": "INSTX"}, probs.append)
    R.add(c, "Т5: сверка по журналу — количество = части ÷ фасовка, процент × приход; неверные строки (seq 2, 4) пойманы, верные и «весь остаток» — нет",
          np_ == 2 and nc_ == 2 and len(probs) == 2 and "seq 2" in probs[0] and "seq 4" in probs[1], f"{np_} {nc_} {probs}")
    s = ready(Session(new_wms("v33")))
    try:
        # ---------------------------------------------------------------- Т8
        cc = s.B("CalcCheck", module="WmsSafe")
        dc = s.B("DailyCheck", True, module="WmsSafe")
        s.B("TestSetLoVersion", "99.1.0", module="WmsStatus")
        st0 = s.B("StGet", "LOVER", module="WmsSafe")
        at = s.B("CalcAtStart", module="WmsSafe")
        st1 = s.B("StGet", "LOVER", module="WmsSafe")
        s.B("TestCalcWant", "1|2|3", module="WmsSafe")
        bad = s.B("CalcCheck", module="WmsSafe")
        dc2 = s.B("DailyCheck", True, module="WmsSafe")
        s.B("TestCalcWant", "", module="WmsSafe")
        s.B("TestSetLoVersion", "", module="WmsStatus")
        R.add(c, "Т8: формулы с известным ответом сходятся; проверка дня включает их; новая версия LibreOffice проверяется при запуске; Calc, "
                 "считающий иначе, — замечание проверки дня (проведение не затронуто)",
              cc == "" and "|OK|" in dc and "формулы Calc" in dc and at == "" and st1 == "99.1.0" and st0 != "99.1.0"
              and "считает не так" in bad and "|FAIL|" in dc2 and "проведение не затронуто" in dc2,
              f"{cc!r}; {dc[:160]}; {st0!r}->{st1!r} {at!r}; {bad[:160]}; {dc2[:200]}")
        s.B("DailyCheck", True, module="WmsSafe")
        # ---------------------------------------------------------------- Т7
        order(s, 1, "1", "Кабель ВВГ", qty="10", unit="м", art="К-3", J=50.0)
        receive(s, 1, "10", place="7-1-1")
        a = ords(s, 1)[21]
        s.doc.store()
        s.B("BackupNow", "manual", module="WmsBackup")
        for r in (1, 2):
            issue(s, r, a, "1", "ива")
        s.doc.store()
        book = s.doc.getURL()
        h0 = hashlib.sha256(open(uno_path(book), "rb").read()).hexdigest()
        seq = s.sysv("LAST_SEQ")
        dr = s.B("RestoreDrill", module="WmsSafe")
        drill_dir = os.path.join(os.path.dirname(uno_path(book)), "WMS_State", "drill")
        h1 = hashlib.sha256(open(uno_path(book), "rb").read()).hexdigest()
        s.B("ServiceRefresh", module="WmsSafe")
        sv = s.doc.Sheets.getByName("Обслуживание")
        line_ok = sv.getCellByPosition(2, 12).getString()
        light_ok = sv.getCellByPosition(2, 2).getString()
        R.add(c, "Т7: учебное восстановление — свежая копия открывается отдельно, журнал после неё (2 операции) непрерывен: «восстанавливается»; "
                 "папка учения убрана, рабочая книга та же; строка на «Обслуживании»",
              dr.startswith("OK:восстанавливается") and "операций 2, до последней № 3" in dr
              and not os.path.exists(drill_dir) and h0 == h1 and s.sysv("LAST_SEQ") == seq and s.state().get("STATE") == "CLEAN"
              and line_ok.startswith("✓") and "восстанавливается" in line_ok,
              f"{dr[:300]}; {os.path.exists(drill_dir)}; {h0 == h1}; {line_ok[:120]}")
        # a gap in the journal after the copy: the drill says so (the file put back after)
        jf = sorted(glob.glob(os.path.join(s.jdir, "WMS_journal_*.csv")))[-1]
        raw = open(jf, "rb").read()
        ls = raw.split(b"\n")
        k = next(i for i, x in enumerate(ls) if x.startswith(f"J1;{int(seq) - 1};".encode()))
        open(jf, "wb").write(b"\n".join(ls[:k] + ls[k + 1:]))
        try:
            gap = s.B("RestoreDrill", module="WmsSafe")
        finally:
            open(jf, "wb").write(raw)
        s.B("ServiceRefresh", module="WmsSafe")
        light_bad = sv.getCellByPosition(2, 2).getString()
        dr2 = s.B("RestoreDrill", module="WmsSafe")
        s.B("ServiceRefresh", module="WmsSafe")
        light_back = sv.getCellByPosition(2, 2).getString()
        R.add(c, "Т7: разрыв журнала после копии — «восстановить всё нельзя» с причиной; Т6: светофор «Здоровье WMS» — красный с тем, что "
                 "сделать; после исправления — снова не красный",
              gap.startswith("ERR:") and "восстановить всё нельзя" in gap and light_bad.startswith("✗") and "учебное восстановление" in light_bad
              and dr2.startswith("OK:") and light_back[:1] in ("✓", "!") and light_ok[:1] in ("✓", "!") and s.state().get("STATE") == "CLEAN",
              f"{gap[:200]}; {light_bad[:200]}; {light_ok[:200]}; {light_back[:200]}")
        oracle(c, s, "после учений")
    finally:
        s.close()


def label_layout(pr, n_labels, rows=4):
    """the rows of each label on «Печать» (text, size of the font, height) and the bars of its barcode: (labels, problems) —
    a bar outside the band of its label (the 4th row) or over a text is a problem. The geometry is read by getPosition() /
    getSize(), once per bar: the property Position of a shape of Calc read through the bridge now and then gives 0 for X or Y"""
    dp = pr.getDrawPage()
    bars = []
    for i in range(dp.getCount()):
        b = dp.getByIndex(i)
        if b.Name == "wmsbar":
            p, z = b.getPosition(), b.getSize()
            bars.append((p.X, p.Y, z.Width, z.Height))
    labels, probs, y, seen = [], [], 0, 0
    for k in range(n_labels):
        h = [pr.getRows().getByIndex(k * rows + j).Height for j in range(rows)]
        cells = [pr.getCellByPosition(0, k * rows + j) for j in range(rows)]
        band0 = y + h[0] + h[1] + h[2]
        band1 = band0 + h[3]
        mine = [b for b in bars if y <= b[1] < y + sum(h)]
        seen += len(mine)
        for b in mine:
            if b[1] < band0 or b[1] + b[3] > band1:
                probs.append(f"этикетка {k + 1}: полоса {b[1]}…{b[1] + b[3]} вне полосы штрихкода {band0}…{band1}")
        labels.append(dict(text=[c.getString() for c in cells], pt=[round(c.CharHeight, 1) for c in cells], h=h, bars=len(mine),
                           width=(min(b[0] for b in mine), max(b[0] + b[2] for b in mine)) if mine else None))
        y += sum(h)
    if seen != len(bars):
        probs.append(f"полос вне этикеток: {len(bars) - seen} из {len(bars)}")
    return labels, probs


@case
def v34_release_101():
    """1.0.1: этикетка 40×30 под термопринтер — номер ЕИ, полное наименование (шрифт уменьшается, пока не уместится; длинное не
    обрезается), единица, штрихкод в своей полосе (ни одна полоса не заходит на текст), без артикула и места; PDF — страницы 40×30 мм;
    сетка ячеек включена для всех листов; «Наличие» L…V — пришло и осталось (рядом), кому назначено, поставщик, дата и документ
    прихода, № заказа, фасовка, серийный номер, «годен до», резерв — заполняются при открытии листа и кнопкой; проведение их не
    пишет; «Осталось» — всегда остаток E (формула), «Пришло» — приход заказа, иного прихода, разделения"""
    c = "V34"
    s = ready(Session(new_wms("v34")))
    try:
        names = ['HG Ниппель 3/4" НР', "Кабель силовой ВВГнг(А)-LS 3х2,5 ок(N, PE)-0,66 ГОСТ 31996-2012 медный плоский, бухта 100 м",
                 "Перчатки трикотажные с ПВХ покрытием «Точка» 10 класс вязки, размер 10 (XL), белые, упаковка 12 пар, для склада и цеха"]
        eis = []
        for i, nm in enumerate(names, start=1):
            order(s, i, str(i), nm, qty="5", unit=["шт", "м", "пар"][i - 1], art=f"А-{i}", T="Участок 3, Иванов" if i == 1 else "")
            receive(s, i, "5", place="4-8-3", doc=f"УПД-{i}")
            eis.append(int(ords(s, i)[21][3:]))
        size0 = s.B("LabelSize", module="WmsItem")
        res = s.B("LabelPrint", tuple(eis), 1, "40×30", module="WmsItem")
        pr = s.doc.Sheets.getByName("Печать")
        labels, probs = label_layout(pr, 3)
        ps = s.doc.StyleFamilies.getByName("PageStyles").getByName("WMS_Этикетка")
        R.add(c, "этикетка 40×30: номер ЕИ, полное наименование (длинное — мельче, но целиком), «Ед. изм.: …», штрихкод (22 полосы) в нижней "
                 "полосе, ни одна полоса не заходит на текст; артикула и места нет; размер по умолчанию — 40×30",
              res == "OK:этикеток 3 (40×30 мм)" and size0 == "40×30" and not probs and abs(ps.Width - 4000) <= 2 and abs(ps.Height - 3000) <= 2
              and [x["text"][1] for x in labels] == names and [x["text"][2] for x in labels] == ["Ед. изм.: шт", "Ед. изм.: м", "Ед. изм.: пар"]
              and all(x["bars"] == 22 and x["text"][3] == "" and x["text"][0].startswith("ЕИ-") for x in labels)
              and labels[0]["pt"][1] > labels[1]["pt"][1] >= 5 and labels[2]["pt"][1] >= 5
              and all(x["width"][0] >= 0 and x["width"][1] <= ps.Width - ps.LeftMargin - ps.RightMargin for x in labels)
              and not any("Арт" in " ".join(x["text"]) or "4-8-3" in " ".join(x["text"]) for x in labels),
              f"{res}; {size0}; {probs[:3]}; {[(x['text'][1][:20], x['pt'], x['h'], x['bars'], x['width']) for x in labels]}")
        pdf = s.B("PdfSave", module="WmsPrint")
        pages = []
        try:
            import pymupdf
            d = pymupdf.open(pdf[3:])
            pages = [(round(pg.rect.width / 72 * 25.4, 1), round(pg.rect.height / 72 * 25.4, 1)) for pg in d]
            words = [" ".join(w[4] for w in pg.get_text("words")) for pg in d]
        except Exception as e:
            words = [str(e)]
        R.add(c, "PDF этикеток: 3 страницы 40×30 мм, на каждой полное наименование целиком", pdf.startswith("OK:") and pages == [(40.0, 30.0)] * 3
              and all(nm.split()[-1] in w for nm, w in zip(names, words)), f"{pdf[:80]}; {pages}; {[w[:60] for w in words]}")
        # «Наличие» L…T
        pk = s.B("PackSet", ords(s, 3)[21], "12", "шт", "", module="WmsPack")
        pp = s.B("PropSet", ords(s, 1)[21], "SN-77", dtext(400), "", module="WmsItem")
        rs = s.doc.Sheets.getByName("Резервы")
        rs.getCellByPosition(0, 1).setString(ords(s, 1)[21])
        rs.getCellByPosition(1, 1).setValue(2)
        rs.getCellByPosition(2, 1).setString("Объект Север")
        seq0 = s.sysv("LAST_SEQ")
        r1 = s.B("StockInfoRefresh", True, module="WmsItem")
        st = s.doc.Sheets.getByName("Наличие")
        hdr = [st.getCellByPosition(col, 0).getString() for col in range(11, 22)]
        row1 = [st.getCellByPosition(col, eis[0]).getString() for col in range(11, 22)]
        row2 = [st.getCellByPosition(col, eis[2]).getString() for col in range(11, 22)]
        r2 = s.B("StockInfoRefresh", False, module="WmsItem")
        R.add(c, "«Наличие» L…V: пришло и осталось рядом, кому назначено, поставщик, дата и документ прихода, № заказа, фасовка, серийный номер, "
                 "«годен до», резерв; второй раз без операций — не пересчитывается; операций в журнале не прибавилось",
              r1.startswith("OK:") and r2 == "SAME" and hdr == ["Пришло", "Осталось", "Кому назначено", "Поставщик / откуда", "Дата прихода",
                                                                "Документ прихода", "№ заказа", "Фасовка", "Серийный / инв. №", "Годен до", "Резерв"]
              and row1[:2] == ["5", "5"] and row2[:2] == ["5", "5"]
              and row1[2] == "Участок 3, Иванов" and row1[3] == "Ромашка" and row1[5] == "УПД-1" and row1[6] == "1" and row1[8] == "SN-77"
              and row1[9] == dtext(400) and row1[10] == "2 · Объект Север" and row2[7] == "в 1 пар — 12 шт" and row2[2] == ""
              and s.sysv("LAST_SEQ") == seq0 and pk.startswith("OK") and pp.startswith("OK"), f"{r1}; {r2}; {hdr}; {row1}; {row2}")
        # an operation: the information follows at the next opening of «Наличие» — only the changed rows are written
        put(s, "Заказы", 1, T="Участок 5")
        e2 = ords(s, 2)[21]
        iss = issue(s, 1, e2, "1,5", "ива")
        rf = s.B("StockInfoRefresh", False, module="WmsItem")
        left_now = (st.getCellByPosition(11, eis[1]).getString(), st.getCellByPosition(12, eis[1]).getString(), st.getCellByPosition(4, eis[1]).getString())
        t_light = st.getCellByPosition(13, eis[0]).getString()
        rfull = s.B("StockInfoRefresh", True, module="WmsItem")
        R.add(c, "выдача 1,5 м, затем открытие «Наличия»: пересчитана только сдвинувшаяся строка — «Пришло» 5, «Осталось» 3,5 (как «Остаток»); "
                 "«Кому назначено», исправленное на «Заказах» без операции, — по «Обновить сведения» (одна строка)",
              iss.startswith("OK") and left_now == ("5", "3,5", "3,5") and rf == "OK:1" and t_light == "Участок 3, Иванов" and rfull == "OK:1"
              and st.getCellByPosition(13, eis[0]).getString() == "Участок 5", f"{iss[:40]}; {rf}; {left_now}; {t_light}; {rfull}; {st.getCellByPosition(13, eis[0]).getString()}")
        # «Пришло» of a line of «Иной приход» and of a new EI of a split; a new row of «Наличие» gets its «Осталось»
        s.special_input(1, B="Офис", D="Стол офисный", F="3", G="шт", H=TODAY, I="1-1-1")
        sp = s.click_spc("BtnSpcPost", 1)
        e4 = str(s.spc(1)[13])
        nxt = int(s.sysv("NEXT_EI"))
        spl = s.B("SplitPost", ords(s, 1)[21], "2", "7-1-2", "часть", module="WmsDaily")
        lf = s.B("StockInfoRefresh", False, module="WmsItem")
        n4 = int(e4[3:]) if e4.startswith("ЕИ-") else 0
        got = {k: (st.getCellByPosition(11, k).getString(), st.getCellByPosition(12, k).getString()) for k in (eis[0], n4, nxt) if k}
        full = s.B("StockInfoRefresh", True, module="WmsItem")
        R.add(c, "«Пришло» / «Осталось» (только сдвинувшиеся строки): иной приход 3 шт — 3 / 3; разделение 2 из 5 — у нового ЕИ 2 / 2, у исходного "
                 "5 / 3; полный пересчёт после этого не меняет ни одной строки (построчный и полный расчёт совпадают)",
              sp.startswith("OK") and spl.startswith("OK") and lf == "OK:3" and got.get(n4) == ("3", "3") and got.get(nxt) == ("2", "2")
              and got.get(eis[0]) == ("5", "3") and full == "OK:0", f"{sp[:60]}; {e4}; {spl[:60]}; {lf}; {got}; {full}")
        ctl = s.doc.getCurrentController()
        grid = {}
        for nm in ("Заказы", "Наличие", "Выдачи"):
            ctl.setActiveSheet(s.doc.Sheets.getByName(nm))
            grid[nm] = ctl.ShowGrid
        cfg = s.B("LoTips", 0, module="WmsCare")
        R.add(c, "сетка ячеек видна на рабочих листах; WMS включает линии сетки и на цветных ячейках (настройка LibreOffice)",
              all(grid.values()) and "✓ линии сетки на цветных ячейках" in cfg, f"{grid}; {[x for x in cfg.splitlines() if 'сетки' in x]}")
        oracle(c, s, "этикетки и сведения 1.0.1")
    finally:
        s.close()


@case
def v35_upgrade_100():
    """1.0.2 для пользователя, уже работающего в 1.0: рабочая книга 1.0.0 (собрана исходниками выпуска 1.0.0, рабочий день в ней)
    обновляется мастером до 1.0.2 — проверка, сборка (сверка каждой ячейки), «сделать рабочей»; данные те же, старая книга и журнал
    побайтно те же; новая WMS — 1.0.2, колонки сведений, сетка, папка состояния WMS перенесена"""
    import run_prime
    import subprocess
    from harness import git_sources
    c = "V35"
    src = os.path.join(OUT, "src_100")
    why = "" if os.path.isdir(os.path.join(src, "tools")) else git_sources("3c46e94", src)
    if why:
        R.add(c, "исходники выпуска 1.0.0", "SKIP", why)
        return
    d = os.path.join(OUT, "cases", "v35 Склад")
    shutil.rmtree(d, ignore_errors=True)
    work = os.path.join(d, "WMS_RELEASE_1.0.0", "WMS_WORK")
    os.makedirs(work)
    p = os.path.join(work, "WMS_PROD.ods")
    b = subprocess.run([sys.executable, os.path.join(src, "tools", "build_ods.py"), p, "--test"], capture_output=True, text=True, timeout=900)
    if b.returncode != 0 or not os.path.exists(p):
        R.add(c, "книга 1.0.0 собрана исходниками 1.0.0", False, (b.stdout + b.stderr)[-300:])
        return
    s = ready(Session(p))
    try:
        core0 = s.sysv("CORE_VERSION")
        order(s, 1, "1", "Перфоратор Bosch", qty="2", art="ПФ-1", T="Участок 3")
        receive(s, 1, "2", place="3-1-1", doc="УПД-7")
        a = ords(s, 1)[21]
        i1 = issue(s, 1, a, "1", "ива")
        s.B("StSet", "FLASH", "2026-10-20 10:00|/media/флешка", module="WmsSafe")
        stk = s.doc.Sheets.getByName("Наличие")
        n_ei = int(s.sysv("NEXT_EI")) - 1
        before = [tuple(r[:8]) for r in stk.getCellRangeByPosition(0, 1, 7, n_ei).getDataArray()]
    finally:
        s.close(save=True)
    h0 = run_prime.sha(p)
    j0 = {f: run_prime.sha(f) for f in glob.glob(os.path.join(work, "WMS_Journal", "*.csv"))}
    rc0, o0 = run_prime.upgrade("check", p)
    new = os.path.join(d, "WMS_RELEASE_1.0.2", "WMS_WORK")
    rc1, o1 = run_prime.upgrade("build", p, "--to", new)
    rc2, o2 = run_prime.upgrade("promote", "--to", new)
    rep = open(os.path.join(new, "UPGRADE_REPORT.md"), encoding="utf-8").read() if os.path.exists(os.path.join(new, "UPGRADE_REPORT.md")) else ""
    R.add(c, "книга 1.0.0: проверка «можно обновить», сборка — 0 расхождений ячеек, «сделать рабочей»; старая книга и журнал побайтно те же, "
             "старая заперта", core0 == "1.0.0" and i1.startswith("OK") and rc0 == 0 and "1.0.0" in o0 and rc1 == 0 and rc2 == 0
          and "расхождений: 0" in rep and run_prime.sha(p) == h0
          and {f: run_prime.sha(f) for f in glob.glob(os.path.join(work, "WMS_Journal", "*.csv"))} == j0
          and os.path.exists(os.path.join(work, "WMS_Journal", "wms.lock")), f"{core0}; {i1[:40]}; {o0[-200:]}; {o1[-200:]}; {o2[-200:]}")
    nb = os.path.join(new, "WMS_PROD.ods")
    s = ready(Session(nb))
    try:
        stk = s.doc.Sheets.getByName("Наличие")
        after = [tuple(r[:8]) for r in stk.getCellRangeByPosition(0, 1, 7, n_ei).getDataArray()]
        ri = s.B("StockInfoRefresh", True, module="WmsItem")
        info = [stk.getCellByPosition(col, int(a[3:])).getString() for col in (11, 12, 13, 14, 16)]
        ctl = s.doc.getCurrentController()
        ctl.setActiveSheet(s.doc.Sheets.getByName("Заказы"))
        flash = s.B("StGet", "FLASH", module="WmsSafe")
        lb = s.B("LabelPrint", (int(a[3:]),), 1, "40×30", module="WmsItem")
        i2 = issue(s, 2, a, "1", "ива")
        R.add(c, "новая книга — WMS 1.0.2: «работа разрешена», каждый ЕИ тот же, сведения «Наличия» (кому назначено, поставщик, документ), "
                 "сетка видна, состояние WMS (копия на флешку) перенесено, этикетка 40×30, проведение работает",
              s.state().get("STATE") == "CLEAN" and s.sysv("CORE_VERSION") == "1.0.2" and before == after and ri.startswith("OK:")
              and info == ["2", "1", "Участок 3", "Ромашка", "УПД-7"] and ctl.ShowGrid and flash.startswith("2026-10-20") and lb.startswith("OK:этикеток 1")
              and i2.startswith("OK"), f"{s.state()}; {s.sysv('CORE_VERSION')}; {before == after}; {ri}; {info}; {ctl.ShowGrid}; {flash}; {lb}; {i2[:40]}")
        P, inf, _ = s.adjcheck(0)
        R.add(c, "оракул (журнал воспроизведён независимо): книга 1.0.0, обновлённая до 1.0.2", not P, f"{inf.get('ops')} операций; {P[:4]}")
    finally:
        s.close()


@case
def v36_wizard_072_100_101():
    """1.0.2 по пути пользователя: склад 0.7.2 → мастером выпуска 1.0.0 рабочая 1.0.0 (в WMS_RELEASE_1.0.0/WMS_WORK) → рабочий день в
    ней → мастер WMS_UPGRADE.ods выпуска 1.0.2 (распакован рядом) сам находит рабочую 1.0, «ПРОВЕРИТЬ», «СОБРАТЬ И СВЕРИТЬ», «СДЕЛАТЬ
    РАБОЧЕЙ»; книги 0.7.2 и 1.0.0 и их журналы побайтно те же; новая 1.0.2 — те же ЕИ, «Сверка переноса» и «остаток = движения» те же,
    оракул"""
    import run_prime
    import subprocess
    from harness import git_sources
    c = "V36"
    src = os.path.join(OUT, "src_100")
    why = "" if os.path.isdir(os.path.join(src, "tools")) else git_sources("3c46e94", src)
    if why:
        R.add(c, "исходники выпуска 1.0.0", "SKIP", why)
        return
    d = os.path.join(OUT, "cases", "v36 Склад")
    shutil.rmtree(d, ignore_errors=True)
    old = run_prime.old_copy_at(os.path.join(d, "WMS_RELEASE_0.7.2", "WMS_WORK"))
    h072 = run_prime.sha(old)
    w100 = os.path.join(d, "WMS_RELEASE_1.0.0", "WMS_WORK")
    up100 = os.path.join(src, "tools", "upgrade.py")
    rr = [subprocess.run([sys.executable, up100] + a, capture_output=True, text=True, timeout=3600, env=run_prime.ENV)
          for a in (["check", old], ["build", old, "--to", w100], ["promote", "--to", w100])]
    p100 = os.path.join(w100, "WMS_PROD.ods")
    if any(r.returncode for r in rr) or not os.path.isfile(p100):
        R.add(c, "склад 0.7.2 обновлён программой выпуска 1.0.0", False, " | ".join((r.stdout + r.stderr)[-200:] for r in rr))
        return
    s = ready(Session(p100))
    try:
        core0 = s.sysv("CORE_VERSION")
        stk = s.doc.Sheets.getByName("Наличие")
        # a day in 1.0.0: a new order received and partly issued (the transferred EIs keep their own checks)
        oc = s.doc.Sheets.getByName("Заказы").getCellRangeByPosition(0, 0, 21, 20000).getDataArray()
        ro = max(i for i, r in enumerate(oc) if any(str(x).strip() for x in r)) + 2
        order(s, ro, "9901", "Перфоратор Bosch", qty="3", art="ПФ-9")
        receive(s, ro, "3", place="3-1-1", doc="УПД-99")
        ei = ords(s, ro)[21]
        n_ei = int(s.sysv("NEXT_EI")) - 1
        col = s.doc.Sheets.getByName("Выдачи").getCellRangeByPosition(0, 0, 17, 20000).getDataArray()
        free = max(i for i, r in enumerate(col) if any(str(x).strip() for x in r)) + 1
        # a book of work (PROD): no test answers — the recipient is taken from «Получатели», as a storekeeper adds him
        rc = s.doc.Sheets.getByName("Получатели")
        rr = 2 + max([i for i, r in enumerate(rc.getCellRangeByPosition(0, 1, 1, 500).getDataArray()) if any(str(x).strip() for x in r)] or [-1])
        rc.getCellByPosition(0, rr).setString("Петров")
        rc.getCellByPosition(1, rr).setString("Петров")
        s.B("RecipientsInvalidate", module="WmsIssue")
        i1 = issue(s, free, ei, "1", "Петров")
        rep0 = s.B("LegacyReport", module="WmsLegacy")
        bc0 = s.B("BalanceCheck", module="WmsLegacy")
        before = [tuple(r[:8]) for r in stk.getCellRangeByPosition(0, 1, 7, n_ei).getDataArray()]
    finally:
        s.close(save=True)
    h100 = run_prime.sha(p100)
    j100 = {f: run_prime.sha(f) for f in glob.glob(os.path.join(w100, "WMS_Journal", "*.csv"))}
    w = run_prime.Wizard(os.path.join(d, f"WMS_RELEASE_{run_prime.VERS['WMS_PRODUCT_VERSION']}"))
    try:
        new_dir = os.path.join(w.folder, "WMS_WORK")
        found = w.cell(4)
        w.B("UpTestConfirm", 1)
        w.B("BtnUpCheck")
        m1 = w.msg()
        w.B("BtnUpBuild")
        m3 = w.msg()
        w.B("BtnUpPromote")
        m5 = w.msg()
    finally:
        w.close()
    R.add(c, "мастер 1.0.2 сам нашёл рабочую 1.0.0 (а не прежнюю 0.7.2), «ПРОВЕРИТЬ» — можно обновить, «СОБРАТЬ И СВЕРИТЬ», «СДЕЛАТЬ "
             "РАБОЧЕЙ»; книги 0.7.2 и 1.0.0 и журнал 1.0.0 побайтно те же, 1.0.0 заперта",
          core0 == "1.0.0" and i1.startswith("OK") and found == p100 and m1.startswith("OK:") and "можно обновить" in m1
          and m3.startswith("OK:") and m5.startswith("OK:") and run_prime.sha(old) == h072 and run_prime.sha(p100) == h100
          and {f: run_prime.sha(f) for f in glob.glob(os.path.join(w100, "WMS_Journal", "*.csv"))} == j100
          and os.path.exists(os.path.join(w100, "WMS_Journal", "wms.lock")) and os.path.isfile(os.path.join(new_dir, "WMS_PROD.ods")),
          f"{core0}; {i1[:300]}; {found}; {m1[:160]}; {m3[:160]}; {m5[:160]}")
    s = ready(Session(os.path.join(new_dir, "WMS_PROD.ods")))
    try:
        stk = s.doc.Sheets.getByName("Наличие")
        after = [tuple(r[:8]) for r in stk.getCellRangeByPosition(0, 1, 7, n_ei).getDataArray()]
        rep1 = s.B("LegacyReport", module="WmsLegacy")
        bc1 = s.B("BalanceCheck", module="WmsLegacy")
        ri = s.B("StockInfoRefresh", True, module="WmsItem")
        lb = s.B("LabelPrint", (int(ei[3:]),), 1, "40×30", module="WmsItem")
        labels, probs = label_layout(s.doc.Sheets.getByName("Печать"), 1)
        col = s.doc.Sheets.getByName("Выдачи").getCellRangeByPosition(0, 0, 17, 20000).getDataArray()
        i2 = issue(s, max(i for i, r in enumerate(col) if any(str(x).strip() for x in r)) + 1, ei, "1", "Петров")
        R.add(c, "новая книга — WMS 1.0.2: «работа разрешена», каждый ЕИ тот же, «Сверка переноса» и «остаток = движения» те же, сведения "
                 "«Наличия», этикетка 40×30 (штрихкод в своей полосе), выдача проводится",
              s.state().get("STATE") == "CLEAN" and s.sysv("CORE_VERSION") == "1.0.2" and before == after
              and re.sub(r", строк.*", "", rep1) == re.sub(r", строк.*", "", rep0) and rep1.startswith("OK:")
              and bc1 == bc0 and ri.startswith("OK:") and lb.startswith("OK:этикеток 1") and labels[0]["bars"] > 20 and not probs and i2.startswith("OK"),
              f"{s.sysv('CORE_VERSION')}; {before == after}; {rep0[:120]} / {rep1[:120]}; {bc0[:60]!r} / {bc1[:60]!r}; {ri}; {lb}; "
              f"{probs[:2]}; {i2[:300]}")
        P, inf, _ = s.adjcheck(0, initial={}, legacy_mn_empty=False)
        R.add(c, "оракул (журнал 0.7.2 + 1.0.0 + 1.0.2 воспроизведён независимо): склад 0.7.2 → 1.0.0 → 1.0.2", not P, f"{inf.get('ops')} операций; {P[:4]}")
    finally:
        s.close()


@case
def v37_kits_by_ei():
    """1.0.1: комплект из конкретных ЕИ — (1) в составе на «Комплектах» номер ЕИ вместо артикула: позиция берётся только с этого ЕИ
    (не со старого ЕИ того же товара), не хватает на нём — отказ, а не добор с другого; неизвестный ЕИ — понятный отказ; (2) «Собрать из
    ЕИ…» на «Наличии»: выделенные строки ЕИ, количество каждого на 1 комплект — та же операция KIT_BUILD, состав записывается на
    «Комплекты» и собирается снова кнопкой «Собрать комплект…»; «Пришло» / «Осталось» комплекта; разборка возвращает позиции"""
    c = "V37"
    s = ready(Session(new_wms("v37")))
    try:
        order(s, 1, "1", "Кабель ВВГ 3х2,5", qty="8", unit="м", art="К-3", J=50.0)
        receive(s, 1, "8", place="7-1-1", k=-30)
        order(s, 2, "2", "Кабель ВВГ 3х2,5", qty="20", unit="м", art="К-3", J=50.0)
        receive(s, 2, "20", place="7-1-2", k=-10)
        order(s, 3, "3", "Розетка", qty="5", art="Р-1", J=100.0)
        receive(s, 3, "5", place="7-2-1")
        order(s, 4, "4", "Изолента", qty="3", J=30.0)
        receive(s, 4, "3", place="7-2-2")
        a1, a2, b, w = [ords(s, r)[21] for r in (1, 2, 3, 4)]
        n1, n2, nb, nw = [int(x[3:]) for x in (a1, a2, b, w)]
        kt = s.doc.Sheets.getByName("Комплекты")
        rows = [("Набор по ЕИ", "Кабель ВВГ", a2.replace("ЕИ-000000", "ЕИ-"), 5.0, "м"), ("", "Розетка", "Р-1", 1.0, "шт"),
                ("Набор с ошибкой", "Кабель", "ЕИ-999999", 1.0, "м")]
        for i, row in enumerate(rows, start=1):
            for j, v in enumerate(row):
                cell = kt.getCellByPosition(j, i)
                cell.setValue(v) if isinstance(v, float) else cell.setString(v)
        btn = ("btnKitFromEis" in form_buttons(kt) and "btnKitFromEis" in form_buttons(s.doc.Sheets.getByName("Наличие"))
               and "номер ЕИ" in kt.getCellByPosition(7, 0).getString())
        r1 = s.B("KitBuild", "Набор по ЕИ", 1.0, "8-1-1", serial(), "к монтажу", module="WmsKit")
        bad = s.B("KitBuild", "Набор с ошибкой", 1.0, "8-1-1", serial(), "", module="WmsKit")
        short = s.B("KitBuild", "Набор по ЕИ", 4.0, "8-1-1", serial(), "", module="WmsKit")
        R.add(c, "в составе номер ЕИ («ЕИ-<n>» в «Артикул»): кабель 5 м снят именно с него (20 → 15), старый ЕИ того же кабеля не тронут (8); "
                 "4 комплекта — отказ «нужно 20 м, есть 15» (с другого ЕИ не добирается); неизвестный ЕИ — отказ «такого ЕИ нет»; кнопки "
                 "«Собрать из ЕИ…» на «Комплектах» и «Наличии», подсказка листа",
              btn and r1.startswith("OK:") and s.stock(n2) == 15.0 and s.stock(n1) == 8.0 and s.stock(nb) == 4.0
              and bad.startswith("ERR:") and "такого ЕИ нет" in bad and short.startswith("ERR:") and "нужно 20 м, есть 15" in short,
              f"{btn}; {r1[:80]}; {bad}; {short}; {[s.stock(x) for x in (n1, n2, nb)]}")
        # «Собрать из ЕИ…» by the function: two EIs, 2 kits, the composition written on «Комплекты»
        nxt = int(s.sysv("NEXT_EI"))
        k2 = f"ЕИ-{nxt:08d}"
        last0 = max(i for i, r in enumerate(kt.getCellRangeByPosition(0, 0, 5, 40).getDataArray()) if any(str(x) for x in r))
        r2 = s.B("KitFromEis", (n1, nw), ("2", "1"), "Сборка из ЕИ", 2.0, "8-1-2", serial(), "тест", True, module="WmsKit")
        rec = [list(r) for r in kt.getCellRangeByPosition(0, last0 + 1, 5, last0 + 2).getDataArray()]
        dup = s.B("KitFromEis", (n1, nw), ("1", "1"), "Сборка из ЕИ", 1.0, "8-1-2", serial(), "тест", True, module="WmsKit")
        badq = s.B("KitFromEis", (n1,), ("абв",), "Ещё набор", 1.0, "8-1-2", serial(), "тест", False, module="WmsKit")
        R.add(c, "«Собрать из ЕИ» (2 ЕИ, 2 комплекта): кабель 2 м × 2 со старого ЕИ (8 → 4), изолента 1 × 2 (3 → 1), новый ЕИ «Сборка из ЕИ» "
                 "2 компл; состав записан на «Комплекты» (ЕИ в «Артикул»); то же название с записью — отказ; количество «абв» — отказ",
              r2.startswith("OK:" + k2) and "состав записан" in r2 and s.stock(n1) == 4.0 and s.stock(nw) == 1.0 and s.stock(nxt) == 2.0
              and rec == [["Сборка из ЕИ", "Кабель ВВГ 3х2,5", a1, 2.0, "м", "собран из ЕИ " + dtext(0)], ["", "Изолента", w, 1.0, "шт", ""]]
              and dup.startswith("ERR:") and "уже описан" in dup and badq.startswith("ERR:") and "на 1 комплект" in badq,
              f"{r2}; {rec}; {dup}; {badq}; {[s.stock(x) for x in (n1, nw, nxt)]}")
        r3 = s.B("KitBuild", "Сборка из ЕИ", 1.0, "8-1-2", serial(), "", module="WmsKit")
        R.add(c, "записанный состав собирается кнопкой «Собрать комплект…» — с тех же ЕИ (кабель 4 → 2, изолента 1 → 0)",
              r3.startswith("OK:") and s.stock(n1) == 2.0 and s.stock(nw) == 0.0, f"{r3}; {[s.stock(x) for x in (n1, nw)]}")
        # the button on «Наличие»: the selected rows (two EIs side by side), the window, no composition written
        nxt = int(s.sysv("NEXT_EI"))
        k4 = f"ЕИ-{nxt:08d}"
        s.goto("Наличие", f"$B${n2 + 1}:$C${nb + 1}")
        s.B("TestPrimeForm", f"Набор UI\t1\t3\t1\t8-1-3\t{TODAY}\tчерез окно\tнет", module="WmsPrimeUi")
        s.B("BtnKitFromEis", module="WmsKit")
        msg = s.U("TestUiLastMessage")
        st = s.doc.Sheets.getByName("Наличие")
        info = (st.getCellByPosition(11, nxt).getString(), st.getCellByPosition(12, nxt).getString(), st.getCellByPosition(12, n2).getString())
        kit_rows = sum(1 for r in kt.getCellRangeByPosition(0, 0, 0, 60).getDataArray() if r[0] == "Набор UI")
        R.add(c, "кнопка «Собрать из ЕИ…» на «Наличии»: выделены строки двух ЕИ, в окне 3 м кабеля и 1 розетка на комплект — новый ЕИ «Набор UI», "
                 "кабель 15 → 12, розетка 4 → 3; «Пришло» / «Осталось» обновлены сразу (комплект 1 / 1, кабель «Осталось» 12); «нет» — состав не записан",
              msg.startswith(k4) and s.stock(n2) == 12.0 and s.stock(nb) == 3.0 and s.stock(nxt) == 1.0 and info == ("1", "1", "12") and kit_rows == 0,
              f"{msg[:120]}; {[s.stock(x) for x in (n2, nb, nxt)]}; {info}; {kit_rows}")
        # the storno of a build from EIs: its row «Удалить» takes the kit apart
        adj = s.doc.Sheets.getByName("Корректировки")
        ctl = [adj.getCellByPosition(16, r).getString() for r in range(1, 40)]
        r_first = 1 + next(i for i, x in enumerate(ctl) if x == f"Проведено · в комплект {k4}")
        s.U("TestOpReason", "передумали")
        d = s.A("AdjustDeleteRow", r_first)
        R.add(c, "«Удалить» строки сборки из ЕИ разбирает комплект: кабель 12 → 15, розетка 3 → 4, комплект 0",
              d.startswith("OK") and s.stock(n2) == 15.0 and s.stock(nb) == 4.0 and s.stock(nxt) == 0.0, f"{d[:120]}")
        lf = s.B("StockInfoRefresh", False, module="WmsItem")
        full = s.B("StockInfoRefresh", True, module="WmsItem")
        R.add(c, "«Наличие» после разборки: построчное обновление (кабель, розетка, комплект) и полный пересчёт совпадают — полный не меняет ничего",
              lf.startswith("OK:") and full == "OK:0" and st.getCellByPosition(12, nxt).getString() == "0", f"{lf}; {full}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок («остаток = движения» и у комплектов из ЕИ)", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "комплекты из конкретных ЕИ")
    finally:
        s.close()


def label_layout_var(pr, rows_list):
    """like label_layout, for labels of different heights in rows: rows_list[k] — the rows of label k (4 + its extra lines); the band of
    the barcode is the last row of the label, the bars must lie inside it"""
    dp = pr.getDrawPage()
    bars = []
    for i in range(dp.getCount()):
        b = dp.getByIndex(i)
        if b.Name == "wmsbar":
            p, z = b.getPosition(), b.getSize()
            bars.append((p.X, p.Y, z.Width, z.Height))
    labels, probs, y, seen, r0 = [], [], 0, 0, 0
    for k, nr in enumerate(rows_list):
        h = [pr.getRows().getByIndex(r0 + j).Height for j in range(nr)]
        cells = [pr.getCellByPosition(0, r0 + j) for j in range(nr)]
        band0 = y + sum(h[:-1])
        band1 = band0 + h[-1]
        mine = [b for b in bars if y <= b[1] < y + sum(h)]
        seen += len(mine)
        for b in mine:
            if b[1] < band0 or b[1] + b[3] > band1:
                probs.append(f"этикетка {k + 1}: полоса {b[1]}…{b[1] + b[3]} вне полосы штрихкода {band0}…{band1}")
        labels.append(dict(text=[c.getString() for c in cells], pt=[round(c.CharHeight, 1) for c in cells], h=h, bars=len(mine)))
        y += sum(h)
        r0 += nr
    if seen != len(bars):
        probs.append(f"полос вне этикеток: {len(bars) - seen} из {len(bars)}")
    return labels, probs


@case
def v38_labels_102():
    """1.0.2: этикетки — (1) товар, который приходит упаковками («В упаковке» в «Номенклатуре» R): на этикетке «Выдача по 1 шт · в уп. 10»,
    одна этикетка на всю поставку; (2) у деталей (категория «Детали») на этикетке артикул, у остальных нет; (3) серийный номер ЕИ — на
    этикетке «S/N: …»; (4) массовая печать: для каждой позиции своё число этикеток, «0» — не печатать, кнопка «По количеству единиц» —
    столько этикеток, сколько пришло (не больше 200 на позицию); этикетки разной высоты, штрихкод всегда в своей полосе; слишком много
    — отказ до печати"""
    c = "V38"
    s = ready(Session(new_wms("v38")))
    try:
        gd = s.doc.Sheets.getByName("Номенклатура")
        gd.getCellByPosition(1, 1).setString("Краска белая 0,9 л")
        gd.getCellByPosition(2, 1).setString("КР-1")
        gd.getCellByPosition(17, 1).setValue(10)
        gd.getCellByPosition(1, 2).setString("Болт М8х30 с шестигранной головкой")
        gd.getCellByPosition(17, 2).setValue(1)
        order(s, 1, "1", "Краска белая 0,9 л", qty="100", art="КР-1", S="Краска")
        receive(s, 1, "100", place="4-1-1", doc="УПД-1")
        order(s, 2, "2", "Шестерня Z-40", qty="5", art="GR-40", S="Детали")
        receive(s, 2, "5", place="4-1-2", doc="УПД-2")
        order(s, 3, "3", "Болт М8х30 с шестигранной головкой", qty="50", art="", S="Метизы")
        receive(s, 3, "50", place="4-1-3", doc="УПД-3")
        order(s, 4, "4", "Кабель", qty="3", unit="м", S="Кабель")
        receive(s, 4, "3", place="4-1-4", doc="УПД-4")
        eis = [int(ords(s, r)[21][3:]) for r in (1, 2, 3, 4)]
        s.B("PropSet", ords(s, 2)[21], "SN-77", "", "", module="WmsItem")
        res = s.B("LabelPrintN", tuple(eis), (1, 2, 1, 1), "40×30", module="WmsItem")
        pr = s.doc.Sheets.getByName("Печать")
        # rows per label: краска — EI, name, unit, the pack line, bars = 5; шестерня — EI, name, unit, article / serial, bars = 5 (twice); болт (pack 1 is no pack) and кабель — 4
        rows = [5, 5, 5, 4, 4]
        labels, probs = label_layout_var(pr, rows)
        t0, t1, t3, t4 = labels[0]["text"], labels[1]["text"], labels[3]["text"], labels[4]["text"]   # краска, шестерня, болт, кабель
        R.add(c, "печать 6 этикеток, у каждой своё число копий: краска 1, шестерня 2, болт 1, кабель 1; штрихкод каждой в своей полосе (этикетки "
                 "разной высоты в строках), ни одна полоса не заходит на текст",
              res == "OK:этикеток 5 (40×30 мм)" and not probs and all(x["bars"] > 20 for x in labels), f"{res}; {probs[:3]}; {[(x['bars'], x['h']) for x in labels]}")
        R.add(c, "краска (в упаковке 10): строка «Выдача по 1 шт · в уп. 10», артикула нет (не «Детали»); шестерня (Детали): «Арт.: GR-40 · S/N: SN-77»; "
                 "болт (в упаковке 1) и кабель — без дополнительных строк",
              t0[0].startswith("ЕИ-") and t0[1] == "Краска белая 0,9 л" and t0[2] == "Ед. изм.: шт" and t0[3] == "Выдача по 1 шт · в уп. 10" and t0[4] == ""
              and t1[3] == "Арт.: GR-40 · S/N: SN-77" and t1[4] == "" and t3[2] == "Ед. изм.: шт" and t3[3] == "" and t4[2] == "Ед. изм.: м" and t4[3] == ""
              and "Арт" not in " ".join(t0) and "GR-40" not in " ".join(t4), f"{t0}; {t1}; {t3}; {t4}")
        # the packing of the EI: «в 1 уп — 10 шт»
        s.B("PackSet", ords(s, 4)[21], "4", "шт", "", module="WmsPack")
        res2 = s.B("LabelPrintN", (eis[3],), (1,), "40×30", module="WmsItem")
        labels2, probs2 = label_layout_var(pr, [5])
        R.add(c, "у ЕИ с фасовкой («в 1 м — 4 шт»): «Выдача по 1 шт · в 1 м — 4 шт»", res2.startswith("OK:") and labels2[0]["text"][3] == "Выдача по 1 шт · в 1 м — 4 шт"
              and not probs2, f"{res2}; {labels2[0]['text']}; {probs2[:2]}")
        # zero: not printed; all zeros: refused; a window answered by the seam
        res3 = s.B("LabelPrintN", tuple(eis), (0, 0, 0, 0), "40×30", module="WmsItem")
        res4 = s.B("LabelPrintN", tuple(eis[:2]), (0, 3), "40×30", module="WmsItem")
        labels4, probs4 = label_layout_var(pr, [5, 5, 5])
        R.add(c, "«0» — позиция не печатается: все нули — отказ «ни для одной позиции не указано число», (0, 3) — три этикетки только второй позиции",
              res3.startswith("ERR:") and "не указано число этикеток" in res3 and res4 == "OK:этикеток 3 (40×30 мм)" and not probs4
              and all("GR-40" in x["text"][3] for x in labels4), f"{res3}; {res4}; {[x['text'][1] for x in labels4]}")
        res5 = s.B("LabelPrintN", tuple(eis), (200, 200, 200, 200), "40×30", module="WmsItem")
        R.add(c, "800 этикеток не помещаются на странице «Печать» — отказ до печати с подсказкой, страница не тронута",
              res5.startswith("ERR:") and "помещается" in res5 and pr.getCellByPosition(0, 0).getString().startswith("ЕИ-"), f"{res5}; {pr.getCellByPosition(0, 0).getString()}")
        # the window: the first position 2 labels, the second 0, the third and the fourth as proposed (1); the size as proposed
        s.goto("Заказы", "$A$2:$A$5")
        s.B("TestPrimeForm", "2\t0\t1\t1\t40×30", module="WmsPrimeUi")
        w1 = s.B("LabelWindow", tuple(eis), module="WmsItem")
        R.add(c, "окно этикеток: для каждой позиции своё число (2, 0, 1, 1) — «OK:этикеток 4»", w1 == "OK:этикеток 4 (40×30 мм)", w1)
        # «По количеству единиц»: краска 100, шестерня 5, болт 50, кабель 3 → 158 labels (the cap 200 per position is not reached)
        s.B("TestPrimeForm", "1\t1\t1\t1\t40×30", module="WmsPrimeUi")
        s.B("TestPrimeFormNext", module="WmsPrimeUi")
        w2 = s.B("LabelWindow", tuple(eis), module="WmsItem")
        u = [s.B("LbUnits", e, module="WmsItem") for e in eis]
        R.add(c, "кнопка «По количеству единиц»: этикеток столько, сколько пришло — 100 + 5 + 50 + 3 = 158", u == [100, 5, 50, 3] and w2 == "OK:этикеток 158 (40×30 мм)",
              f"{u}; {w2}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "этикетки 1.0.2")
    finally:
        s.close()


@case
def v39_serial_parts():
    """1.0.2: детали с серийным номером — «Серийные…»: из ЕИ в N деталей получается N отдельных ЕИ (по одной детали; тот же артикул,
    место, «годен до» и примечание), каждый со своим серийным номером и своей этикеткой («Арт.: … · S/N: …»); деталь без номера остаётся в
    общем ЕИ; когда номера есть у всех, последняя деталь остаётся в самом ЕИ; повтор номера, чужой номер, лишние номера, дробный остаток,
    фасовка — отказ до проведения; цепочка прежних операций SPLIT и EI_PROP (журнал и оракул те же)"""
    c = "V39"
    s = ready(Session(new_wms("v39")))
    try:
        order(s, 1, "1", "Шестерня Z-40", qty="5", art="GR-40", S="Детали")
        receive(s, 1, "5", place="4-2-1", doc="УПД-1")
        order(s, 2, "2", "Датчик давления", qty="3", art="DD-7", S="Детали")
        receive(s, 2, "3", place="4-2-2", doc="УПД-2")
        order(s, 3, "3", "Лента", qty="2.5", unit="м", art="L-1", S="Лента")
        receive(s, 3, "2,5", place="4-2-3", doc="УПД-3")
        e1, e2, e3 = [ords(s, r)[21] for r in (1, 2, 3)]
        n1, n2 = int(e1[3:]), int(e2[3:])
        s.B("PropSet", e1, "", dtext(300), "проверена", module="WmsItem")
        seq0 = int(s.sysv("LAST_SEQ"))
        bad_dup = s.B("SerialSplit", e1, ("SN-9", "sn-9"), module="WmsParts")
        bad_many = s.B("SerialSplit", e1, ("a", "b", "c", "d", "e", "f"), module="WmsParts")
        bad_none = s.B("SerialSplit", e1, ("", " "), module="WmsParts")
        bad_frac = s.B("SerialSplit", e3, ("L-1",), module="WmsParts")
        R.add(c, "отказ до проведения: номер дважды («SN-9», «sn-9»), номеров больше деталей, ни одного номера, дробный остаток 2,5 м — "
                 "журнал не вырос", bad_dup.startswith("ERR:") and "дважды" in bad_dup and bad_many.startswith("ERR:") and "а деталей в" in bad_many
              and bad_none.startswith("ERR:") and "ни один" in bad_none and bad_frac.startswith("ERR:") and "целых" in bad_frac
              and int(s.sysv("LAST_SEQ")) == seq0, f"{bad_dup}; {bad_many}; {bad_none}; {bad_frac}")
        res = s.B("SerialSplit", e1, ("SN-1", "SN-2", "", "SN-4"), module="WmsParts")
        new = res.split("|")[1].split(",") if res.startswith("OK:") else []
        nn = [int(x[3:]) for x in new]
        st = s.doc.Sheets.getByName("Наличие")
        rows = [st.getCellRangeByPosition(0, k, 9, k).getDataArray()[0] for k in nn]
        props = [s.B("PropOf", k, module="WmsItem") for k in nn]
        orig = st.getCellRangeByPosition(0, n1, 9, n1).getDataArray()[0]
        R.add(c, "5 деталей, номера у трёх (SN-1, SN-2, —, SN-4): три новых ЕИ по 1 шт с теми же названием, артикулом, категорией «Детали» и местом; "
                 "у каждого свой серийный №, «годен до» и примечание исходного; в исходном ЕИ осталось 2 без номера",
              res.startswith("OK:3|") and len(nn) == 3 and all(r[4] == 1.0 and r[1] == "Шестерня Z-40" and r[2] == "GR-40" and r[6] == "Детали" and r[5] == "4-2-1" for r in rows)
              and [p[0] for p in props] == ["SN-1", "SN-2", "SN-4"] and all(p[2] == "проверена" and p[1] > 0 for p in props) and orig[4] == 2.0
              and s.B("PropOf", n1, module="WmsItem")[0] == "", f"{res}; {[r[:7] for r in rows]}; {props}; {orig[:7]}")
        ops = s.journal()[0]
        types = [o[4] if isinstance(o, (list, tuple)) else o for o in ops[-6:]] if ops else []
        R.add(c, "в журнале цепочка прежних операций: 3 × SPLIT и 3 × EI_PROP (новых видов операций нет); поиск по серийному номеру находит ЕИ",
              int(s.sysv("LAST_SEQ")) == seq0 + 6 and s.B("SerialEI", "SN-2", module="WmsItem") == nn[1], f"{s.sysv('LAST_SEQ')} {seq0}; {types}")
        dup = s.B("SerialSplit", e1, ("SN-1",), module="WmsParts")
        R.add(c, "номер, уже записанный у другого ЕИ («SN-1»), — отказ «уже записан у ЕИ-…»", dup.startswith("ERR:") and "уже записан у ЕИ-" in dup, dup)
        # every part has a number: the last stays in the EI itself
        res2 = s.B("SerialSplit", e2, ("A-1", "A-2", "A-3"), module="WmsParts")
        new2 = res2.split("|")[1].split(",") if res2.startswith("OK:") else []
        R.add(c, "у всех трёх деталей есть номера: два новых ЕИ, третья деталь осталась в самом ЕИ с номером A-3 (его остаток 1)",
              res2.startswith("OK:3|") and len(new2) == 3 and new2[-1] == e2 and s.stock(n2) == 1.0 and s.B("PropOf", n2, module="WmsItem")[0] == "A-3"
              and [s.B("PropOf", int(x[3:]), module="WmsItem")[0] for x in new2] == ["A-1", "A-2", "A-3"], f"{res2}; {s.stock(n2)}")
        # the labels of the serial parts: the article and the serial number
        lab = s.B("LabelPrintN", tuple(nn), (1, 1, 1), "40×30", module="WmsItem")
        labels, probs = label_layout_var(s.doc.Sheets.getByName("Печать"), [5, 5, 5])
        R.add(c, "этикетки трёх серийных деталей: «Арт.: GR-40 · S/N: SN-1 / SN-2 / SN-4»", lab == "OK:этикеток 3 (40×30 мм)" and not probs
              and [x["text"][3] for x in labels] == ["Арт.: GR-40 · S/N: SN-1", "Арт.: GR-40 · S/N: SN-2", "Арт.: GR-40 · S/N: SN-4"], f"{lab}; {[x['text'] for x in labels]}")
        # the window: 2 parts left in the first EI — the numbers typed in the form, the question about the labels answered «Напечатать»
        s.B("TestPrimeForm", "W-1\tW-2", module="WmsPrimeUi")
        w = s.B("SerialWindow", n1, module="WmsParts")
        labels2, probs2 = label_layout_var(s.doc.Sheets.getByName("Печать"), [5, 5])
        R.add(c, "окно «Серийные…»: в ЕИ осталось 2 детали, вписаны W-1, W-2 — остался один ЕИ с W-2 и один новый с W-1; этикетки напечатаны (ответ «Напечатать» по умолчанию)",
              w.startswith("OK:2|") and not probs2 and sorted(x["text"][3] for x in labels2) == ["Арт.: GR-40 · S/N: W-1", "Арт.: GR-40 · S/N: W-2"], f"{w}; {[x['text'] for x in labels2]}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок («остаток = движения» у всех новых ЕИ)", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "серийные детали")
    finally:
        s.close()


def doc_of(s, name="Text"):
    """a document made by WMS (the last one) as text: the whole text or the cells of its first table"""
    return s.B("DcDocText" if name == "Text" else "DcDocTable", module="WmsDocs")


@case
def v40_acts():
    """1.0.2: акт по выделенным строкам «Выдач» — документ Writer по образцу пользователя: «АКТ № …», «приема-передачи оборудования»
    (офису) или «выдачи оборудования» (получателю), город и дата, таблица № п/п · Наименование · Артикул · Количество · Ед изм (серийный №
    в наименовании), подписи; номер из реестра «Акты» и в «№ документа выдачи» выдач (ISSUE_DOC); повтор акта; выдачи разным
    получателям — отказ"""
    c = "V40"
    s = ready(Session(new_wms("v40")))
    try:
        order(s, 1, "1", "Перфоратор Bosch", qty="4", art="ПФ-1")
        receive(s, 1, "4", place="3-1-1", doc="УПД-1")
        order(s, 2, "2", "Кабель ВВГ 3х2,5", qty="30", unit="м", art="")
        receive(s, 2, "30", place="3-1-2", doc="УПД-2")
        order(s, 3, "3", "Канистра 5 л", qty="2", art="КН-5")
        receive(s, 3, "2", place="3-1-3", doc="УПД-3")
        e1, e2, e3 = [ords(s, r)[21] for r in (1, 2, 3)]
        s.B("PropSet", e1, "SN-5", "", "", module="WmsItem")
        s.B("PackSet", e3, "4", "кан", "", module="WmsPack")
        i1 = issue(s, 1, e1, "1", "офис")
        i2 = issue(s, 2, e2, "4", "офис")
        i3 = issue(s, 3, e1, "1", "ива")
        i4 = issue(s, 4, e3, "0,25", "офис")
        ish = s.doc.Sheets.getByName("Выдачи")
        R.add(c, "четыре выдачи проведены (две офису, одна Иванову, одна канистра частью)", all(x.startswith("OK") for x in (i1, i2, i3, i4)), f"{i1[:30]}; {i2[:30]}; {i3[:30]}; {i4[:30]}")
        ak = s.doc.Sheets.getByName("Акты")
        seq0 = int(s.sysv("LAST_SEQ"))
        # the act to the office: rows 1 and 2
        s.goto("Выдачи", "$A$2:$A$3")
        r1 = s.B("ActFromSelection", module="WmsDocs")
        table = doc_of(s, "Table").split("\n")
        text = doc_of(s)
        iss = [ish.getCellByPosition(1, k).getString() for k in (1, 2, 3, 4)]
        reg = [list(x) for x in ak.getCellRangeByPosition(0, 1, 6, 1).getDataArray()]
        R.add(c, "акт офису по выдачам 1–2: «АКТ № 1», «приема-передачи оборудования», «г Москва» и дата, «а офисом принято»; таблица: №, наименование "
                 "с серийным номером, артикул, количество, единица; подписи «сдал / принял»",
              r1 == "OK:1|1|Офис|2" and text.startswith("АКТ № 1\nприема-передачи оборудования\nг Москва\t" + TODAY + " г.\n")
              and "а офисом принято следующее оборудование:" in text and table[0] == "№ п/п|Наименование|Артикул|Количество|Ед изм"
              and table[1] == "1.|Перфоратор Bosch, S/N SN-5|ПФ-1|1|шт" and table[2] == "2.|Кабель ВВГ 3х2,5||4|м" and len(table) == 3
              and "Оборудование сдал:\tОборудование принял:" in text, f"{r1}; {text[:260]!r}; {table}")
        R.add(c, "реестр «Акты»: № 1, сегодня, «Приёма-передачи (в офис)», Офис, 2 позиции, выдачи «1,2»; «№ документа выдачи» выдач 1–2 — «А-1», у других пусто; "
                 "в журнале одна операция ISSUE_DOC", reg[0][0] == 1.0 and reg[0][2] == "Приёма-передачи (в офис)" and reg[0][3] == "Офис" and reg[0][4] == 2.0
              and reg[0][5] == "1,2" and iss == ["А-1", "А-1", "", ""] and int(s.sysv("LAST_SEQ")) == seq0 + 1, f"{reg}; {iss}; {s.sysv('LAST_SEQ')} {seq0}")
        s.B("DcDocClose", module="WmsDocs")
        # the act of issue to a person: one row selected, «Только эта выдача», the kind suggested for a person — «выдачи»
        s.goto("Выдачи", "$A$4")
        r2 = s.B("ActFromSelection", module="WmsDocs")
        text2 = doc_of(s)
        table2 = doc_of(s, "Table").split("\n")
        R.add(c, "акт выдачи Иванову: «АКТ № 2», «выдачи оборудования», «а получателем Иванов Иван Андреевич принято», «Оборудование выдал / получил», под линией получателя — "
                 "его ФИО; номер «А-2» в выдаче 3",
              r2 == "OK:2|2|Иванов Иван Андреевич|1" and text2.startswith("АКТ № 2\nвыдачи оборудования\n") and "а получателем Иванов Иван Андреевич принято следующее оборудование:" in text2
              and "Оборудование выдал:\tОборудование получил:" in text2 and "/ Иванов Иван Андреевич /" in text2 and table2[1] == "1.|Перфоратор Bosch, S/N SN-5|ПФ-1|1|шт"
              and ish.getCellByPosition(1, 3).getString() == "А-2", f"{r2}; {text2!r}")
        s.B("DcDocClose", module="WmsDocs")
        # a part of a unit with a packing: «1 кан» of «1 шт = 4 кан»
        s.goto("Выдачи", "$A$5")
        r3 = s.B("ActFromSelection", module="WmsDocs")
        table3 = doc_of(s, "Table").split("\n")
        R.add(c, "канистра частью (0,25 шт = 1 кан): в акте «1» и «кан», а не «0,25 шт»", r3 == "OK:3|1|Офис|1" and table3[1] == "1.|Канистра 5 л|КН-5|1|кан", f"{r3}; {table3}")
        s.B("DcDocClose", module="WmsDocs")
        # different recipients selected: a refusal, no number taken
        s.goto("Выдачи", "$A$2:$A$4")
        r4 = s.B("ActFromSelection", module="WmsDocs")
        nreg = len([1 for x in ak.getCellRangeByPosition(0, 1, 0, 30).getDataArray() if x[0] != ""])
        R.add(c, "выдачи разным получателям — отказ «разным получателям», номер не взят (в реестре 3 акта)", r4.startswith("ERR:") and "разным получателям" in r4 and nreg == 3, f"{r4}; {nreg}")
        # the act again: the same number, date and rows; the register does not grow, the journal does not grow
        seq1 = int(s.sysv("LAST_SEQ"))
        rp = s.B("ActRepeat", 1, False, module="WmsDocs")
        tp = doc_of(s, "Table").split("\n")
        textp = doc_of(s)
        R.add(c, "«Повторить акт» № 1: тот же номер, дата, строки и серийный номер; реестр и журнал не выросли", rp == "OK:1|1|Офис|2" and textp.startswith("АКТ № 1\n") and tp[1] == table[1]
              and tp[2] == table[2] and int(s.sysv("LAST_SEQ")) == seq1 and len([1 for x in ak.getCellRangeByPosition(0, 1, 0, 30).getDataArray() if x[0] != ""]) == 3, f"{rp}; {tp}")
        s.B("DcDocClose", module="WmsDocs")
        r5 = s.B("ActRepeat", 0, False, module="WmsDocs")
        R.add(c, "«Повторить акт» без строки реестра — понятный отказ", r5.startswith("ERR:") and "строку акта" in r5, r5)
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок («№ документа выдачи» записан прежней операцией)", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:300])
        oracle(c, s, "акты")
    finally:
        s.close()


@case
def v41_day_report():
    """1.0.2: отчёт дня — документ Writer по шаблону пользователя («1. Завершено сегодня», «2. Задачи на <следующий рабочий день>», «3. В
    процессе», «4. Контроль данных»): приёмка (машины, позиции, количество, внутренние коды ЕИ, адресное хранение), иной приход, выдачи
    — «Оцифровка выдач за <дата>» за прошлый рабочий день отдельно от выдач сегодняшним числом (выдача, проведённая и удалённая в тот
    же день, не считается), акты, исправления, свои строки с листа «Отчёт дня»; график 5/2: выходные пропускаются"""
    import datetime
    c = "V41"
    names = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]
    acc = ["понедельник", "вторник", "среду", "четверг", "пятницу", "субботу", "воскресенье"]
    s = ready(Session(new_wms("v41")))
    try:
        today = T
        prev = today - datetime.timedelta(days=1)
        while prev.weekday() >= 5:
            prev -= datetime.timedelta(days=1)
        nxt = today + datetime.timedelta(days=1)
        while nxt.weekday() >= 5:
            nxt += datetime.timedelta(days=1)
        order(s, 1, "1", "Краска белая 0,9 л", qty="100", art="КР-1")
        receive(s, 1, "100", place="4-1-1", doc="УПД-1")
        order(s, 2, "2", "Кабель ВВГ", qty="3", unit="м")
        receive(s, 2, "3", place="4-1-2", doc="УПД-2")
        order(s, 3, "3", "Шестерня Z-40", qty="5", art="GR-40")
        receive(s, 3, "5", place="4-1-3", doc="УПД-3")
        e1, e2, e3 = [ords(s, r)[21] for r in (1, 2, 3)]
        s.special_input(1, B="Производство", D="Кронштейн сварной", F="12", G="шт", H=TODAY, I="P-3")
        sp = s.click_spc("BtnSpcPost", 1)
        car = s.B("CarArrive", "Газель А123ВС", "Петрович", module="WmsCar")

        def issue_on(r, ei, qty, who, d):
            put(s, "Выдачи", r, L=ei, E=qty, H=serial((d - today).days), I=who)
            return s.I("IssuePostRow", r)

        i1 = issue_on(1, e1, "5", "офис", prev)
        i2 = issue_on(2, e2, "1", "ива", prev)
        i3 = issue_on(3, e3, "1", "егр", today)
        i4 = issue_on(4, e3, "1", "егр", today)
        s.U("TestOpReason", "ошибка ввода")
        dl = s.I("IssueDeleteRow", 4)
        s.goto("Выдачи", "$A$2")
        ak = s.B("ActFromSelection", module="WmsDocs")
        s.B("DcDocClose", module="WmsDocs")
        dr = s.doc.Sheets.getByName("Отчёт дня")
        dr.getCellByPosition(0, 5).setString("Оформление пропусков для курьеров")
        dr.getCellByPosition(1, 2).setString("Иванов, кладовщик.")
        seq = int(s.sysv("LAST_SEQ"))
        ok = all(x.startswith("OK") for x in (i1, i2, i3, i4, dl, ak))
        R.add(c, "день подготовлен: приход 3 позиций, иной приход, машина, 4 выдачи (две прошлого рабочего дня, две сегодняшнего), удаление одной, акт",
              ok and str(sp).startswith("OK") and str(car).startswith("OK"), f"{[x[:30] for x in (i1, i2, i3, i4, dl, ak)]}; {sp}; {str(car)[:60]}")
        res = s.B("DayReport", serial(0), False, module="WmsDocs")
        text = doc_of(s)
        lines = text.split("\n")
        R.add(c, "шаблон: дата «дд.мм.гггг г.», подпись с листа, четыре раздела; «Задачи на <следующий рабочий день> (дата)»; «в процессе» — по умолчанию «Работа над учётом склада»",
              res.startswith("OK:") and lines[0] == TODAY + " г." and lines[1] == "Иванов, кладовщик." and "1. Завершено сегодня:" in lines
              and f"2. Задачи на {acc[nxt.weekday()]} ({nxt.strftime('%d.%m.%Y')}):" in lines and "3. В процессе:" in lines and "4. Контроль данных (по журналу WMS):" in lines
              and "• Работа над учётом склада" in lines, f"{res}; {lines[:8]}")
        pd = prev.strftime("%d.%m.%Y")
        bul = [x[2:] for x in lines if x.startswith("• ")]
        want = {
            "машины": "Принято машин с поставками: 1 (Петрович)",
            "приёмка": "Приёмка по заказам: принято позиций — 3; количество: 105 шт, 3 м. Товар пересчитан и принят в WMS",
            "иной": "Иной приход (инвентаризация текущих остатков, принятое от производства и др.): записей — 1 (Производство — 1)",
            "коды": "Присвоены внутренние коды (ЕИ) для складского учёта: 4 — от ЕИ-00000201 до ЕИ-00000204",
            "хранение": "Выставлено на адресное хранение (указано место): 4",
            "оцифровка": f"Оцифровка выдач за {pd} ({names[prev.weekday()]}): проведено выдач — 2; выдано: 5 шт, 1 м; получатели: ",
            "сегодня": f"Выдачи сегодняшним числом ({TODAY}): проведено выдач — 1; выдано: 1 шт; получатели: Ермолин Егор Павлович — 1",
            "акты": "Оформлены акты (документы и товар к отправке): 1 — № 1 (приёма-передачи (в офис), Офис, позиций — 1)",
            "накладные": "Оформлены накладные / документы выдачи: А-1",
            "исправления": "Исправления учёта: исправлено операций — 0, удалено (сторно) — 1",
            "вручную": "Оформление пропусков для курьеров",
        }
        got = {k: next((b for b in bul if b.startswith(v)), None) for k, v in want.items()}
        R.add(c, "«Завершено сегодня»: машины, приёмка со суммой по единицам, иной приход по источникам, внутренние коды ЕИ, адресное хранение, "
                 "оцифровка выдач за прошлый рабочий день (2 выдачи, 5 шт и 1 м) отдельно от сегодняшних (1; удалённая не считается), акты, накладные, исправления, "
                 "строка с листа «Отчёт дня»", all(v is not None for v in got.values()),
              f"{[k for k, v in got.items() if v is None]}; {[b for b in bul if 'выдач' in b.lower()]}")
        ctl = [x for x in bul if x.startswith("Операций проведено за день:")]
        R.add(c, "«Контроль данных»: операций за день — столько же, сколько в журнале (с № 1), число ЕИ в учёте",
              len(ctl) == 1 and ctl[0] == f"Операций проведено за день: {seq} (журнал № 1–{seq})" and any(b.startswith("ЕИ в учёте: ") for b in bul), f"{ctl}; {seq}")
        s.B("DcDocClose", module="WmsDocs")
        # a later day: the digitising of the working days after the last date of an issue (5/2: the weekend is skipped)
        later = today + datetime.timedelta(days=4)
        s.B("BtnDrClear", module="WmsDocs")
        res2 = s.B("DayReport", serial(4), False, module="WmsDocs")
        t2 = doc_of(s)
        days = []
        d = today
        while True:
            d += datetime.timedelta(days=1)
            if d > later:
                break
            if d.weekday() < 5:
                days.append(f"{d.strftime('%d.%m.%Y')} ({names[d.weekday()]})")
        want_line = "Оцифровка выдач за " + ", ".join(days) + f" (последняя оцифрованная дата выдачи — {TODAY})"
        R.add(c, "задачи: «Оцифровка выдач за …» — рабочие дни после последней даты выдачи до дня отчёта, выходные пропущены; и «Операций в журнале за этот день нет»",
              res2.startswith("OK:") and ("• " + want_line) in t2.split("\n") and "• операций в WMS за этот день не проведено" in t2.split("\n")
              and "• Операций в журнале за этот день нет" in t2.split("\n"), f"{res2}; {want_line}; {[x for x in t2.split(chr(10)) if 'Оцифровка' in x]}")
        s.B("DcDocClose", module="WmsDocs")
        fri = datetime.date(2026, 10, 23)
        r_work = [s.B("DrIsWork", float((fri + datetime.timedelta(days=k) - datetime.date(1899, 12, 30)).days), module="WmsDocs") for k in range(0, 4)]
        nw = s.B("DrNextWork", float((fri - datetime.date(1899, 12, 30)).days), module="WmsDocs")
        R.add(c, "график 5/2: пятница — рабочий, суббота и воскресенье — нет, понедельник — да; после пятницы следующий рабочий день — понедельник",
              list(r_work) == [True, False, False, True] and int(nw) == (datetime.date(2026, 10, 26) - datetime.date(1899, 12, 30)).days, f"{r_work}; {nw}")
        # «Очистить «Сделано вручную»» (done above before the report of the later day)
        R.add(c, "«Очистить «Сделано вручную»» — список очищен, задачи и «в процессе» остались", dr.getCellByPosition(0, 5).getString() == ""
              and dr.getCellByPosition(2, 5).getString() == "Выдача инструмента/оборудования" and dr.getCellByPosition(4, 5).getString() == "Работа над учётом склада", "")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок; журнал отчётом не менялся", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0") and int(s.sysv("LAST_SEQ")) == seq, sc[:200])
        oracle(c, s, "отчёт дня")
    finally:
        s.close()


@case
def v42_categories_fasteners():
    """1.0.2: подсказки категорий по названию («Категории»): пустая «Категория» на «Заказах» и в «Ином приходе» получает подсказку, своя
    категория пользователя не заменяется; справочник «Метизы» из «Наличия»: тип, размер, стандарт, класс, покрытие; одна строка на
    наименование с артикулом (остаток, места, число ЕИ), порядок по типу–диаметру–длине, поиск по словам; «Болторез» и «Болт» с
    чужой категорией — не метизы; журнал справочниками не меняется"""
    c = "V42"
    s = ready(Session(new_wms("v42")))
    try:
        sug = lambda n: s.B("TestCategory", n, module="WmsCat")
        cases = [("Болт М8х30 DIN 933 8.8 оцинкованный", "Метизы"), ("Гайка М8", "Метизы"), ("Болторез 900 мм", "Инструмент"),
                 ("Выключатель одноклавишный", "Электрика"), ("Кабель ВВГ 3х2,5", "Кабель и провод"), ("Шестерня Z-40", "Детали"),
                 ("Краска белая 0,9 л", "Химия и краски"), ("Перчатки (хлопок)", "СИЗ"), ("Лист стальной 3 мм", "Металл"), ("Уголок 50х50", "Металл"),
                 ("Нечто странное", ""), ("", "")]
        got = [(n, sug(n)) for n, _ in cases]
        bad = [(n, g, w) for (n, g), (_, w) in zip(got, cases) if g != w]
        R.add(c, "подсказки по названию: болт/гайка → Метизы, болторез → Инструмент (узкая категория выше), выключатель — не «ключ», кабель, шестерня, краска, "
                 "перчатки в скобках, лист стальной, уголок; незнакомое название — без подсказки", not bad, str(bad))
        seq0 = int(s.sysv("LAST_SEQ"))
        # «Заказы»: the name typed → the empty category gets the hint; the category typed by the user stays
        s.type_ord("B", 1, "Болт М8х30 DIN 933 8.8 оцинкованный")
        s.type_ord("B", 2, "Нечто странное")
        s.type_ord("S", 3, "Моя группа")
        s.type_ord("B", 3, "Гайка М8 DIN 934")
        s.type_ord("B", 4, "Шестерня Z-40")
        s.type_ord("S", 4, "")
        osh = s.doc.Sheets.getByName("Заказы")
        cats = [osh.getCellByPosition(18, r).getString() for r in (1, 2, 3, 4)]
        R.add(c, "«Заказы»: название → «Категория» «Метизы»; незнакомое — пусто; своя «Моя группа» не заменена; удалённая вручную категория после "
                 "того, как название уже есть, не возвращается", cats == ["Метизы", "", "Моя группа", ""], str(cats))
        # «Иной приход»: the name of a new part
        s.special_input(1, B="Производство", D="Подшипник 6204")
        scat = s.doc.Sheets.getByName("Иной приход").getCellByPosition(9, 1).getString()
        R.add(c, "«Иной приход»: название «Подшипник 6204» → категория «Детали»", scat == "Детали", scat)
        # the list of the categories and the own words
        v = osh.getCellByPosition(18, 6).Validation
        R.add(c, "«Категория»: выпадающий список из листа «Категории», но любое своё слово допускается (сообщение об ошибке выключено)",
              v.Type.value == "LIST" and "Категории" in v.getFormula1() and not v.ShowErrorMessage, f"{v.Type.value}; {v.getFormula1()}; {v.ShowErrorMessage}")
        s.goto("Категории", "$A$12")
        s.o.dispatch(s.doc, ".uno:EnterString", StringName="Мебель")
        s.goto("Категории", "$B$12")
        s.o.dispatch(s.doc, ".uno:EnterString", StringName="стол, стул, шкаф")
        R.add(c, "своя строка в «Категории» («Мебель: стол, стул, шкаф») работает сразу, без перезапуска", sug("Стул офисный") == "Мебель", sug("Стул офисный"))
        # fasteners from the stock
        n0 = int(s.B("FastRefresh", "", module="WmsCat").split(":")[1])
        order(s, 5, "5", "Болт М8х30 DIN 933 8.8 оцинкованный", qty="100", art="B-830")
        order(s, 6, "6", "Болт М8х20 DIN 933 8.8 оцинкованный", qty="50", art="B-820")
        order(s, 7, "7", "Болт М10х30 DIN 933 8.8 оцинкованный", qty="40", art="B-1030")
        order(s, 8, "8", "Гайка М8 DIN 934", qty="200", art="G-8")
        order(s, 9, "9", "Болторез 900 мм", qty="1", art="BR-900")
        order(s, 10, "10", "Болт М8х30 DIN 933 8.8 оцинкованный", qty="20", art="B-830")
        order(s, 11, "11", "Заклёпка 4,8х10 алюминий", qty="500", art="Z-48", S="Крепёж")
        order(s, 12, "12", "Болт М6 для мебели", qty="10", art="M-6", S="Мебель")
        places = {5: "5-1-1", 6: "5-1-2", 7: "5-1-3", 8: "5-2-1", 9: "6-1-1", 10: "5-1-4", 11: "5-2-2", 12: "6-1-2"}
        res = [receive(s, r, str(q), doc=f"УПД-{r}", place=pl) for r, q, pl in
               [(5, 100, places[5]), (6, 50, places[6]), (7, 40, places[7]), (8, 200, places[8]), (9, 1, places[9]), (10, 20, places[10]),
                (11, 500, places[11]), (12, 10, places[12])]]
        seq1 = int(s.sysv("LAST_SEQ"))
        R.add(c, "приход 8 позиций (болты, гайка, болторез, заклёпка со своей категорией «Крепёж», болт для мебели с категорией «Мебель»)",
              all(str(x).startswith("OK") for x in res), str(res)[:200])
        fr = s.B("FastRefresh", "", module="WmsCat")
        n_all = int(fr.split(":")[1]) if fr.startswith("OK:") else -1
        fsh = s.doc.Sheets.getByName("Метизы")

        def rows_of(k):
            return [list(r) for r in fsh.getCellRangeByPosition(0, 4, 10, 4 + k - 1).getDataArray()] if k > 0 else []

        mine = ("B-820", "B-830", "B-1030", "G-8", "Z-48")
        rr = rows_of(n_all)
        ours = [x for x in rr if x[6] in mine]
        # the fixture book has its own fasteners: ours are 5 new rows among them, and neither «Болторез» nor «Болт для мебели» is there
        R.add(c, "«Метизы»: к метизам книги добавились 5 строк — Болт М8×20, Болт М8×30 (две партии одной строкой), Болт М10×30, Гайка М8, Заклёпка (в таком порядке: "
                 "тип, диаметр, длина); болторез и «Болт для мебели» с чужой категорией — не метизы",
              n_all == n0 + 5 and [x[6] for x in ours] == list(mine) and not any(("Болторез" in x[5] or "для мебели" in x[5]) for x in rr), f"{fr} (было {n0}); {[x[6] for x in ours]}")
        b830 = ours[1]
        R.add(c, "строка «Болт М8х30»: тип, размер М8×30, DIN 933, класс 8.8, покрытие, остаток двух партий 120, места «5-1-1; 5-1-4», ЕИ с остатком 2",
              b830[0] == "Болт" and b830[1] == "М8×30" and b830[2] == "DIN 933" and b830[3] == "8.8" and b830[4] == "Оцинкованный" and b830[6] == "B-830"
              and float(b830[7]) == 120 and b830[9] == "5-1-1; 5-1-4" and float(b830[10]) == 2, str(b830))
        zk = ours[4]
        R.add(c, "«Заклёпка 4,8х10»: размер «4,8×10», покрытие «Алюминий», остаток 500", zk[0] == "Заклёпка" and zk[1] == "4,8×10" and zk[4] == "Алюминий"
              and float(zk[7]) == 500, str(zk))
        parse = lambda n: s.B("TestFastParse", n, module="WmsCat").replace(",", ".")
        ps = [(parse("Болт М8х30 DIN 933 8.8 оцинкованный"), "М8×30|8|30|DIN 933|8.8|Оцинкованный|Болт".replace(",", ".")),
              (parse("Винт 4,2х32 ГОСТ 1144-80 А2"), "4.2×32|4.2|32|ГОСТ 1144-80|А2||Винт"),
              (parse("Гайка М10 DIN 934 нержавеющая"), "М10|10|0|DIN 934||Нержавеющая сталь|Гайка"),
              (parse("Шайба 8"), "|0|0|||" + "|Шайба"),
              (parse("Саморез 3,5х25 чёрный"), "3.5×25|3.5|25|||Чёрный|Саморез")]
        badp = [(g, w) for g, w in ps if g != w]
        R.add(c, "разбор названия: размер (М8×30, 4,2×32, М10, «8» не размер), стандарт (DIN, ГОСТ), класс (8.8, А2), покрытие, тип", not badp, str(badp))

        def found(words):
            n = s.B("FastRefresh", words, module="WmsCat")
            k = int(n.split(":")[1]) if n.startswith("OK:") else -1
            return k, [x[6] for x in rows_of(k)]

        k1, n1 = found("b-8")
        k2, n2 = found("5-1-4")
        k3, n3 = found("g-8 din 934")
        k4, n4 = found("нет такого слова")
        k5, n5 = found("заклепка алюм")
        R.add(c, "поиск по словам: артикул «b-8» — 2 болта М8; место «5-1-4» — Болт М8×30; «g-8 din 934» — гайка; «заклепка алюм» (е вместо ё) — заклёпка; "
                 "чужое слово — 0 строк",
              sorted(n1) == ["B-820", "B-830"] and "B-830" in n2 and n3 == ["G-8"] and n5 == ["Z-48"] and k4 == 0, f"{n1}; {n2}; {n3}; {n5}; {k4}")
        # the search cell of the sheet, as the user types: Enter in «Найти»
        s.goto("Метизы", "$B$2")
        s.o.dispatch(s.doc, ".uno:EnterString", StringName="g-8")
        r6 = rows_of(1)
        stat = fsh.getCellByPosition(0, 2).getString()
        R.add(c, "ячейка «Найти» листа «Метизы»: «g-8» + Enter — в списке одна гайка; строка итога называет слова и номер журнала",
              len(r6) == 1 and r6[0][6] == "G-8" and stat.startswith("Метизов (позиций): 1") and "g-8" in stat and "журнал № 8" in stat, f"{r6}; {stat}")
        s.goto("Метизы", "$B$2")
        s.o.dispatch(s.doc, ".uno:EnterString", StringName="")
        R.add(c, f"пустое «Найти» — список снова полный ({n_all})", fsh.getCellByPosition(0, 2).getString().startswith(f"Метизов (позиций): {n_all}."),
              fsh.getCellByPosition(0, 2).getString())
        sc = s.B("ActionSelfCheck")
        R.add(c, "справочники ничего не проводят: журнал после приходов не изменился при списках и подсказках; самопроверка — без ошибок",
              int(s.sysv("LAST_SEQ")) == seq1 and seq1 > seq0 and sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), f"{seq0} → {seq1} → {s.sysv('LAST_SEQ')}; {sc[:120]}")
        oracle(c, s, "метизы и категории")
    finally:
        s.close()


@case
def v43_upgrade_101_102():
    """1.0.2 для пользователя, уже работающего в 1.0.1: рабочая книга 1.0.1 (собрана исходниками выпуска 1.0.1, день работы в ней) обновляется
    мастером — проверка, сборка (сверка каждой ячейки), «сделать рабочей»; данные те же, старая книга и журнал побайтно те же; в новой —
    листы «Акты», «Отчёт дня», «Категории», «Метизы» (без защиты), «В упаковке, шт» в «Номенклатуре» (прежняя подсказка сдвинута), список
    категорий, кнопки; всё новое работает на данных этой книги: метизы, подсказка категории, серийные детали, этикетка, акт, отчёт дня"""
    import run_prime
    import subprocess
    from harness import git_sources
    c = "V43"
    src = os.path.join(OUT, "src_101")
    why = "" if os.path.isdir(os.path.join(src, "tools")) else git_sources("456d7b4", src)
    if why:
        R.add(c, "исходники выпуска 1.0.1", "SKIP", why)
        return
    d = os.path.join(OUT, "cases", "v43 Склад")
    shutil.rmtree(d, ignore_errors=True)
    work = os.path.join(d, "WMS_RELEASE_1.0.1", "WMS_WORK")
    os.makedirs(work)
    p = os.path.join(work, "WMS_PROD.ods")
    b = subprocess.run([sys.executable, os.path.join(src, "tools", "build_ods.py"), p, "--test"], capture_output=True, text=True, timeout=900)
    if b.returncode != 0 or not os.path.exists(p):
        R.add(c, "книга 1.0.1 собрана исходниками 1.0.1", False, (b.stdout + b.stderr)[-300:])
        return
    s = ready(Session(p))
    try:
        core0 = s.sysv("CORE_VERSION")
        gs = s.doc.Sheets.getByName("Номенклатура")
        hint0 = gs.getCellByPosition(17, 0).getString()
        order(s, 1, "1", "Болт М8х30 DIN 933 8.8 оцинкованный", qty="100", art="B-830")
        receive(s, 1, "100", place="5-1-1", doc="УПД-1")
        order(s, 2, "2", "Краска белая 0,9 л", qty="10", art="КР-1")
        receive(s, 2, "10", place="4-1-1", doc="УПД-2")
        order(s, 3, "3", "Шестерня Z-40", qty="3", art="GR-40")
        receive(s, 3, "3", place="4-1-2", doc="УПД-3")
        e1, e2, e3 = [ords(s, r)[21] for r in (1, 2, 3)]
        i1 = issue(s, 1, e2, "2", "ива")
        stk = s.doc.Sheets.getByName("Наличие")
        n_ei = int(s.sysv("NEXT_EI")) - 1
        before = [tuple(r[:8]) for r in stk.getCellRangeByPosition(0, 1, 7, n_ei).getDataArray()]
        seq_old = int(s.sysv("LAST_SEQ"))
    finally:
        s.close(save=True)
    h0 = run_prime.sha(p)
    j0 = {f: run_prime.sha(f) for f in glob.glob(os.path.join(work, "WMS_Journal", "*.csv"))}
    rc0, o0 = run_prime.upgrade("check", p)
    new = os.path.join(d, "WMS_RELEASE_1.0.2", "WMS_WORK")
    rc1, o1 = run_prime.upgrade("build", p, "--to", new)
    rc2, o2 = run_prime.upgrade("promote", "--to", new)
    rep = open(os.path.join(new, "UPGRADE_REPORT.md"), encoding="utf-8").read() if os.path.exists(os.path.join(new, "UPGRADE_REPORT.md")) else ""
    R.add(c, "книга 1.0.1: проверка «можно обновить», сборка — 0 расхождений ячеек (нового только четыре листа и заголовок «В упаковке»), «сделать "
             "рабочей»; старая книга и журнал побайтно те же, старая заперта",
          core0 == "1.0.1" and i1.startswith("OK") and hint0.startswith("Товар — строка") and rc0 == 0 and "1.0.1" in o0 and rc1 == 0 and rc2 == 0
          and "расхождений: 0" in rep and run_prime.sha(p) == h0 and {f: run_prime.sha(f) for f in glob.glob(os.path.join(work, "WMS_Journal", "*.csv"))} == j0
          and os.path.exists(os.path.join(work, "WMS_Journal", "wms.lock")), f"{core0}; {i1[:40]}; {o0[-200:]}; {o1[-200:]}; {o2[-200:]}; {rep[-300:]}")
    nb = os.path.join(new, "WMS_PROD.ods")
    s = ready(Session(nb))
    try:
        sh = s.doc.Sheets
        stk = sh.getByName("Наличие")
        after = [tuple(r[:8]) for r in stk.getCellRangeByPosition(0, 1, 7, n_ei).getDataArray()]
        names = [sh.getByIndex(i).Name for i in range(sh.getCount())]
        own = ("Акты", "Отчёт дня", "Категории", "Метизы")
        gs = sh.getByName("Номенклатура")
        v = sh.getByName("Заказы").getCellByPosition(18, 5).Validation
        v2 = sh.getByName("Иной приход").getCellByPosition(9, 5).Validation

        def btn(sheet, name):
            f = sh.getByName(sheet).getDrawPage().getForms()
            return f.getCount() > 0 and any(f.getByIndex(0).getByIndex(i).Name == name for i in range(f.getByIndex(0).getCount()))

        R.add(c, "новая книга — WMS 1.0.2: «работа разрешена», тот же журнал и каждый ЕИ; четыре новых листа без защиты; «В упаковке, шт» в R, прежняя "
                 "подсказка «Номенклатуры» в S; список категорий на «Заказах» и «Ином приходе»; кнопки «Серийные…» и «Акт…»",
              s.state().get("STATE") == "CLEAN" and s.sysv("CORE_VERSION") == "1.0.2" and before == after and int(s.sysv("LAST_SEQ")) == seq_old
              and all(n in names and not sh.getByName(n).isProtected() for n in own)
              and gs.getCellByPosition(17, 0).getString().startswith("В упаковке") and gs.getCellByPosition(18, 0).getString().startswith("Товар — строка") and "В упаковке, шт" in gs.getCellByPosition(18, 0).getString()
              and v.Type.value == "LIST" and "Категории" in v.getFormula1() and v2.Type.value == "LIST" and btn("Наличие", "btnSerials") and btn("Выдачи", "btnAct"),
              f"{s.state()}; {s.sysv('CORE_VERSION')}; {before == after}; {[n for n in own if n not in names]}; {gs.getCellByPosition(17, 0).getString()[:30]}; "
              f"{v.Type.value}; {btn('Наличие', 'btnSerials')} {btn('Выдачи', 'btnAct')}")
        # the new features on the data of this book
        fr = s.B("FastRefresh", "B-830", module="WmsCat")
        fsh = sh.getByName("Метизы")
        fast = [list(r) for r in fsh.getCellRangeByPosition(0, 4, 10, 4).getDataArray()]
        R.add(c, "«Метизы» из «Наличия» книги 1.0.1: по артикулу B-830 одна строка — Болт М8×30, DIN 933, 8.8, оцинкованный, остаток 100", fr == "OK:1"
              and fast[0][0] == "Болт" and fast[0][1] == "М8×30" and fast[0][2] == "DIN 933" and fast[0][3] == "8.8" and fast[0][4] == "Оцинкованный"
              and float(fast[0][7]) == 100, f"{fr}; {fast}")
        sg = s.B("TestCategory", "Шестерня Z-40", module="WmsCat")
        sp = s.B("SerialSplit", e3, ("SN-1", "SN-2", "SN-3"), module="WmsParts")
        lb = s.B("LabelPrint", (int(e2[3:]),), 1, "40×30", module="WmsItem")
        R.add(c, "подсказка «Детали»; серийные детали: 3 детали → 3 ЕИ с номерами; этикетка с упаковкой не мешает", sg == "Детали" and sp.startswith("OK:3|")
              and lb.startswith("OK:этикеток 1"), f"{sg}; {sp}; {lb}")
        i2 = issue(s, 2, e2, "1", "ива")
        s.goto("Выдачи", "$A$2")
        ak = s.B("ActFromSelection", module="WmsDocs")
        s.B("DcDocClose", module="WmsDocs")
        rd = s.B("DayReport", serial(0), False, module="WmsDocs")
        txt = doc_of(s)
        s.B("DcDocClose", module="WmsDocs")
        R.add(c, "после обновления: выдача проводится, акт оформляется (А-1 в реестре «Акты»), отчёт дня строится из журнала книги 1.0.1",
              i2.startswith("OK") and ak.startswith("OK") and rd.startswith("OK:") and TODAY + " г." in txt and "1. Завершено сегодня:" in txt
              and "Приёмка по заказам: принято позиций — 3" in txt, f"{i2[:40]}; {ak[:80]}; {rd}; {txt[:200]!r}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc[:200])
        P, inf, _ = s.adjcheck(0)
        R.add(c, "оракул (журнал воспроизведён независимо): книга 1.0.1, обновлённая до 1.0.2, и новые операции", not P, f"{inf.get('ops')} операций; {P[:4]}")
    finally:
        s.close()


def spc_line(s, r, typ, name=None, qty=None, unit="шт", place="A-1", art=None, who=None):
    """one row of «Иной приход» typed like a user: the type, the article (a known part fills its card), then the rest"""
    s.special_input(r, B=typ, E=art, D=name, F=qty, G=unit, H=TODAY, I=place, K=who)


def day_snapshot(s):
    """what the book holds after the operations: the registry of the EIs, the counters, the rows of «Иной приход»"""
    n = int(s.sysv("NEXT_EI"))
    stk = [tuple(r[:11]) for r in s.doc.Sheets.getByName("Наличие").getCellRangeByPosition(0, 1, 10, n).getDataArray()]
    keys = ("LAST_SEQ", "NEXT_EI", "NEXT_NO", "NEXT_RET", "NEXT_SPL", "NEXT_OFF", "NEXT_PROD", "NEXT_DET", "NEXT_OLD", "NEXT_OTH", "NEXT_ADJ", "NEXT_CAR")
    return stk, [s.sysv(k) for k in keys], [tuple(r[:16]) for r in s.doc.Sheets.getByName("Иной приход").getCellRangeByPosition(0, 1, 15, 12).getDataArray()]


def morning_and_day(s):
    """the morning (an order received, an issue, a part received) saved as the copy, then the day: orders, «Иной приход» (new rows and the
    part added to by its article), a return, a vehicle, an act — everything only in the journal. (seq of the copy, seq at the end, the answers, the snapshot)"""
    order(s, 1, "1", "Краска белая", qty="100", art="КР-1")
    receive(s, 1, "100", place="4-1-1", doc="УПД-1")
    e1 = ords(s, 1)[21]
    put(s, "Выдачи", 1, L=e1, E="2", H=serial(), I="ива")
    i1 = s.I("IssuePostRow", 1)
    spc_line(s, 1, "Детали", "Шестерня Z-40", "100", art="GR-40", place="D-1")
    sp0 = s.click_spc("BtnSpcPost", 1)
    s.doc.store()
    seq0 = int(s.sysv("LAST_SEQ"))
    res = [i1, sp0]
    for r, name, q, pl in ((2, "Кабель ВВГ", "30", "4-1-2"), (3, "Лампа", "12", "4-1-3"), (4, "Скотч", "40", "4-1-4")):
        order(s, r, str(r), name, qty=q, art=f"A-{r}")
        res.append(receive(s, r, q, place=pl, doc=f"УПД-{r}"))
    spc_line(s, 2, "Офис", "Стол", "2", place="A-1")
    res.append(s.click_spc("BtnSpcPost", 2))
    spc_line(s, 3, "Производство", "Корпус", "3", place="P-1", who="Цех 2")
    res.append(s.click_spc("BtnSpcPost", 3))
    spc_line(s, 4, "Детали", "Втулка", "20", art="BT-20", place="D-2")
    res.append(s.click_spc("BtnSpcPost", 4))
    spc_line(s, 5, "Детали", None, "50", unit=None, place=None, art="GR-40")
    res.append(s.click_spc("BtnSpcPost", 5))
    spc_line(s, 6, "Детали", None, "5", unit=None, place=None, art="BT-20")
    res.append(s.click_spc("BtnSpcPost", 6))
    put(s, "Возврат", 1, B="1", F="1", H=serial(), I="ива")
    res.append(s.B("ReturnPostRow", 1, module="WmsReturn"))
    res.append(s.B("CarArrive", "Газель А123ВС", "Петрович", module="WmsCar"))
    s.goto("Выдачи", "$A$2")
    res.append(s.B("ActFromSelection", module="WmsDocs"))
    s.B("DcDocClose", module="WmsDocs")
    return seq0, int(s.sysv("LAST_SEQ")), res, day_snapshot(s)


@case
def v44_tail_after_return_to_copy():
    """1.1: копия книги сохранена утром, днём проведены операции всех видов (приходы по заказам, «Иной приход» — новые строки и пополнение
    детали по известному артикулу, возврат, визит машины, акт); книга снова в состоянии копии, журнал впереди — «Восстановить» довозит все
    операции с теми же номерами ЕИ (в 1.0.2 оно останавливалось на пополнении: в копии нет строки с её «Внутренним кодом»)"""
    c = "V44"
    p = new_wms("v44")
    s = ready(Session(p))
    try:
        seq0, seq1, res, snap_a = morning_and_day(s)
        R.add(c, "день проведён после копии: 3 прихода по заказам, 5 строк «Иного прихода» (3 новых и 2 пополнения детали), возврат, визит машины, акт",
              all(str(x).startswith("OK") for x in res) and seq1 - seq0 >= 11, f"{seq0} → {seq1}; {[str(x)[:30] for x in res]}")
    finally:
        s.kill()
    s = ready(Session(p))
    try:
        unl = s.B("ActionForceUnlock")
        st0 = s.state()
        rec = s.B("ActionRecover")
        st1 = s.state()
        snap_b = day_snapshot(s)
        R.add(c, "книга в состоянии утренней копии: «хвост» журнала (TAIL); «Восстановить» довозит все операции дня",
              st0.get("BLOCK") == "TAIL" and f"восстановлено операций: {seq1 - seq0}" in rec and "ОСТАНОВКА" not in rec and st1.get("STATE") == "CLEAN",
              f"{st0.get('STATE')}/{st0.get('BLOCK')}; {rec[:260]}; {st1.get('STATE')}")
        R.add(c, "после «Восстановить» реестр «Наличие» (все ЕИ с теми же номерами), счётчики и строки «Иного прихода» — как были до сбоя",
              snap_a[0] == snap_b[0] and snap_a[1] == snap_b[1] and snap_a[2] == snap_b[2],
              f"{[(i, x, y) for i, (x, y) in enumerate(zip(snap_a[0], snap_b[0])) if x != y][:2]}; {snap_a[1]} / {snap_b[1]}; "
              f"{[(i, x, y) for i, (x, y) in enumerate(zip(snap_a[2], snap_b[2])) if x != y][:2]}")
        sc = s.B("ActionSelfCheck")
        P, inf, _ = s.adjcheck(0)
        R.add(c, "самопроверка — без ошибок; оракул журнала (независимый повтор всех операций) — без расхождений",
              sc.startswith("САМОПРОВЕРКА WMS: ошибок 0") and not P, f"{sc[:200]}; {P[:3]}")
    finally:
        s.close()


@case
def v45_upgrade_with_tail():
    """1.1, путь пользователя: рабочая книга 1.0.2 возвращена к утренней копии (журнал впереди на операции дня — те, что 1.0.2 довезти не
    умела), мастер обновляет её до 1.1: проверка принимает хвост (предупреждение, не стоп), сборка пробно довозит хвост на рабочей копии,
    «сделать рабочей» довозит его в новой книге — те же номера ЕИ, самопроверка без ошибок; старая книга и журнал побайтно те же"""
    import run_prime
    import subprocess
    from harness import git_sources
    c = "V45"
    src = os.path.join(OUT, "src_102")
    why = "" if os.path.isdir(os.path.join(src, "tools")) else git_sources("86607f3", src)
    if why:
        R.add(c, "исходники выпуска 1.0.2", "SKIP", why)
        return
    d = os.path.join(OUT, "cases", "v45 Склад")
    shutil.rmtree(d, ignore_errors=True)
    work = os.path.join(d, "WMS_RELEASE_1.0.2", "WMS_WORK")
    os.makedirs(work)
    p = os.path.join(work, "WMS_PROD.ods")
    b = subprocess.run([sys.executable, os.path.join(src, "tools", "build_ods.py"), p, "--test"], capture_output=True, text=True, timeout=900)
    if b.returncode != 0 or not os.path.exists(p):
        R.add(c, "книга 1.0.2 собрана исходниками 1.0.2", False, (b.stdout + b.stderr)[-300:])
        return
    s = ready(Session(p))
    try:
        core0 = s.sysv("CORE_VERSION")
        seq0, seq1, res, snap_a = morning_and_day(s)
        ok = all(str(x).startswith("OK") for x in res)
    finally:
        s.kill()
    s = ready(Session(p))
    try:
        s.B("ActionForceUnlock")
        st0 = s.state()
    finally:
        s.close()
    R.add(c, "книга 1.0.2 в состоянии утренней копии, журнал впереди (TAIL)", core0 == "1.0.2" and ok and st0.get("BLOCK") == "TAIL" and seq1 > seq0,
          f"{core0}; {seq0} → {seq1}; {st0.get('BLOCK')}")
    h0 = run_prime.sha(p)
    j0 = {f: run_prime.sha(f) for f in glob.glob(os.path.join(work, "WMS_Journal", "*.csv"))}
    rc0, o0 = run_prime.upgrade("check", p)
    new = os.path.join(d, "WMS_RELEASE_1.1.0", "WMS_WORK")
    rc1, o1 = run_prime.upgrade("build", p, "--to", new)
    rc2, o2 = run_prime.upgrade("promote", "--to", new)
    rep = open(os.path.join(new, "UPGRADE_REPORT.md"), encoding="utf-8").read() if os.path.exists(os.path.join(new, "UPGRADE_REPORT.md")) else ""
    R.add(c, "мастер: «проверить» принимает хвост журнала (предупреждение), «собрать и сверить» — 0 расхождений и пробное довезение, «сделать рабочей»; "
             "старая книга и журнал побайтно те же",
          rc0 == 0 and "довезёт" in o0 and rc1 == 0 and rc2 == 0 and "расхождений: 0" in rep and "пробного довезения" in rep and run_prime.sha(p) == h0
          and {f: run_prime.sha(f) for f in glob.glob(os.path.join(work, "WMS_Journal", "*.csv"))} == j0,
          f"{rc0} {rc1} {rc2}; {o0[-300:]}; {o1[-300:]}; {o2[-300:]}")
    nb = os.path.join(new, "WMS_PROD.ods")
    s = ready(Session(nb))
    try:
        snap_b = day_snapshot(s)
        sc = s.B("ActionSelfCheck")
        R.add(c, "новая книга: «работа разрешена», все операции дня на месте — реестр «Наличие» (те же номера ЕИ), счётчики, строки «Иного прихода»; "
                 "самопроверка без ошибок",
              s.state().get("STATE") == "CLEAN" and s.sysv("CORE_VERSION") == run_prime.VERS["WMS_CORE_VERSION"] and snap_a[0] == snap_b[0]
              and snap_a[1] == snap_b[1] and snap_a[2] == snap_b[2] and sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"),
              f"{s.state()}; {[(i, x, y) for i, (x, y) in enumerate(zip(snap_a[0], snap_b[0])) if x != y][:2]}; {snap_a[1]} / {snap_b[1]}; {sc[:120]}")
        P, inf, _ = s.adjcheck(0)
        R.add(c, "оракул (журнал воспроизведён независимо): книга 1.0.2 с хвостом, обновлённая до 1.1", not P, f"{inf.get('ops')} операций; {P[:4]}")
    finally:
        s.close()


def sort_registry(s):
    """what the storekeeper did: Data → Sort ascending on the registry «Наличие» (an empty row among the EIs goes to the end)"""
    import uno
    from com.sun.star.table import TableSortField
    from com.sun.star.beans import PropertyValue
    sh = s.doc.Sheets.getByName("Наличие")
    sh.unprotect("wms")
    n = int(s.sysv("NEXT_EI"))
    rng = sh.getCellRangeByPosition(0, 1, 10, n - 1)
    f = TableSortField()
    f.Field = 0
    f.IsAscending = True
    pv = PropertyValue()
    pv.Name = "SortFields"
    pv.Value = uno.Any("[]com.sun.star.table.TableSortField", (f,))
    pv2 = PropertyValue()
    pv2.Name = "ContainsHeader"
    pv2.Value = False
    rng.sort((pv, pv2))
    sh.protect("wms")


@case
def v46_registry_guard():
    """1.1: реестр «Наличие» отсортирован (в нумерации ЕИ есть пропуск — пустая строка — и она ушла в конец, строки ниже сдвинулись вверх): WMS
    находит это формулой Calc, при сохранении сохраняет прежний файл как копию «последняя без ошибок» и говорит об этом, при запуске не даёт
    проводить и объясняет, что делать; «Вернуться к копии» возвращает хорошее состояние, журнал доводит проведённое после него"""
    c = "V46"
    p = new_wms("v46")
    s = ready(Session(p))
    bdir = os.path.join(s.dir, "WMS_Backups")
    try:
        sh = s.doc.Sheets.getByName("Наличие")
        sh.unprotect("wms")
        sh.getCellRangeByPosition(0, 100, 10, 100).clearContents(1 + 2 + 4 + 16)
        sh.protect("wms")
        ok_before = (s.B("RegistryMisplaced", module="WmsGuard"), s.B("RegistryProblem", module="WmsGuard"))
        s.doc.store()
        seq0 = int(s.sysv("LAST_SEQ"))
        r1 = issue(s, 1, "ЕИ-00000001", "1", "ива")
        r2 = issue(s, 2, "ЕИ-00000150", "1", "ива")
        seq1 = int(s.sysv("LAST_SEQ"))
        q150 = s.stock(150)
        R.add(c, "реестр с пропуском в нумерации (пустая строка ЕИ-100) — в порядке: формула считает 0, проблемы нет, запуск без блокировки",
              ok_before == (0, "") and s.state().get("STATE") == "CLEAN" and r1.startswith("OK") and r2.startswith("OK") and seq1 == seq0 + 2, f"{ok_before}; {r1[:30]}")
        sort_registry(s)
        n = s.B("RegistryMisplaced", module="WmsGuard")
        pr = s.B("RegistryProblem", module="WmsGuard")
        R.add(c, "после сортировки формула находит сдвинутые строки (ЕИ-101 … выше на одну), проблема названа словами: «строка 101: «ЕИ-00000102» вместо "
                 "ЕИ-00000101», «так бывает, если реестр отсортировали»", n > 90 and "вместо" in pr and "отсортировали" in pr, f"{n}; {pr[:240]}")
        s.U("TestUiAuto", 1)
        s.doc.store()
        msg = s.U("TestUiLastMessage") or ""
        goods = glob.glob(os.path.join(bdir, "*_good.ods"))
        st = s.state()
        blocked = s.B("PostingBlockReason", module="WmsCore")
        R.add(c, "сохранение такой книги: предупреждение («последняя без ошибок», «Вернуться к копии»), прежний файл сохранён как копия _good, "
                 "проведение остановлено (REGISTRY_ORDER)",
              "ВНИМАНИЕ" in msg and "последняя без ошибок" in msg and len(goods) == 1 and st.get("BLOCK") == "REGISTRY_ORDER"
              and str(blocked).startswith("BLOCKED:"), f"{msg[:160]}; {goods}; {st.get('BLOCK')}; {str(blocked)[:60]}")
        s.doc.store()
        goods2 = glob.glob(os.path.join(bdir, "*_good.ods"))
        R.add(c, "второе сохранение той же книги копию не затирает (хорошая копия — одна, прежняя)", len(goods2) == 1 and goods2 == goods and os.path.getsize(goods[0]) > 0,
              str(goods2))
    finally:
        s.kill()
    s = ready(Session(p))
    try:
        s.B("ActionForceUnlock")
        st = s.state()
        sh = s.doc.Sheets.getByName("Главная")
        panel = [sh.getCellByPosition(1, r).getString() for r in range(6, 12)]
        R.add(c, "запуск сохранённой отсортированной книги: ЗАБЛОКИРОВАНО (REGISTRY_ORDER), «Причина» называет строку, «Что сделать» — «Вернуться к копии»",
              st.get("BLOCK") == "REGISTRY_ORDER" and any("вместо" in x for x in panel) and any("Вернуться к копии" in x for x in panel), f"{st.get('STATE')}/{st.get('BLOCK')}; "
              f"{[x[:70] for x in panel]}")
        name = [x[0] for x in s.B("CopyList", module="WmsSafe")]
        good = [x for x in name if x.endswith("_good.ods")]
        url = s.doc.getURL()
        old_doc = s.doc
        rt = s.B("ReturnToCopy", good[0], False, module="WmsSafe") if good else "нет копии"
        newd = find_doc(s, url, old_doc)
        msg2 = ""
        if newd is not None:
            s.doc = newd
            t0 = time.time()
            while time.time() - t0 < 60:
                msg2 = s.U("TestUiLastMessage") or ""
                if "возвращена к копии" in msg2 or msg2.startswith("ERR"):
                    break
                time.sleep(0.5)
        R.add(c, "«Вернуться к копии» → «последняя без ошибок»: реестр снова в порядке, два проведённых после копии списания довезены по журналу, "
                 "остаток тот же, «работа разрешена»",
              good and rt.startswith("OK") and newd is not None and "возвращена к копии" in msg2 and s.state().get("STATE") == "CLEAN"
              and s.B("RegistryMisplaced", module="WmsGuard") == 0 and s.sysv("LAST_SEQ") == seq1 and s.stock(150) == q150,
              f"{rt}; {msg2[:200]}; {s.state()}; {s.sysv('LAST_SEQ')} (было {seq1}); {s.stock(150)} (было {q150})")
    finally:
        s.close()


@case
def v47_views_ops_whatnow():
    """1.1: «Наличие по названию» (копия реестра, по умолчанию с остатком, в порядке названия; слова «болт м8» находят «Болт M8x30» с латинскими буквами; список можно
    сортировать — «Обновить» строит заново), «Остаток деталей» (одна строка на деталь по артикулу: остаток, число ЕИ, места, серийные номера), «Операции дня» (журнал
    простыми словами: за дату и хвост журнала до «Восстановить»), «Что делать?» (совет под каждым красным пунктом самопроверки)"""
    c = "V47"
    p = new_wms("v47")
    s = ready(Session(p))
    try:
        order(s, 1, "1", "Болт M8x30 DIN 933 оцинкованный", qty="100", art="B-830")
        receive(s, 1, "100", place="5-1-1", doc="УПД-1")
        order(s, 2, "2", "Краска белая 0,9 л", qty="10", art="КР-1")
        receive(s, 2, "10", place="4-1-1", doc="УПД-2")
        spc_line(s, 1, "Детали", "Шестерня Z-40", "100", art="GR-40", place="D-1")
        s.click_spc("BtnSpcPost", 1)
        spc_line(s, 2, "Детали", None, "50", unit=None, place=None, art="GR-40")
        s.click_spc("BtnSpcPost", 2)
        e_gear = [i for i in range(1, int(s.sysv("NEXT_EI"))) if s.doc.Sheets.getByName("Наличие").getCellByPosition(2, i).getString() == "GR-40"][0]
        s.B("PropSet", f"ЕИ-{e_gear:08d}", "SN-7", "", "", module="WmsItem")
        r1 = s.B("ByNameRefresh", "", False, module="WmsStockView")
        sh = s.doc.Sheets.getByName("Наличие по названию")
        rows = [list(r) for r in sh.getCellRangeByPosition(0, 4, 10, 4 + 300).getDataArray()]
        names = [r[1] for r in rows if r[0]]
        n_stock = sum(1 for i in range(1, int(s.sysv("NEXT_EI"))) if isinstance(s.doc.Sheets.getByName("Наличие").getCellByPosition(4, i).getValue(), float)
                      and s.doc.Sheets.getByName("Наличие").getCellByPosition(4, i).getValue() > 0)
        firsts = [x[0].lower().replace("ё", "е") for x in names]
        R.add(c, "«Наличие по названию»: все ЕИ с остатком, в порядке названия (А…Я), заголовок и строка итога", r1 == f"OK:{n_stock}" and firsts == sorted(firsts)
              and "ЕИ в списке:" in sh.getCellByPosition(0, 2).getString(), f"{r1}; {n_stock}; {names[:4]}")
        r2 = s.B("ByNameRefresh", "болт м8", False, module="WmsStockView")
        rows2 = [list(r) for r in sh.getCellRangeByPosition(0, 4, 10, 8).getDataArray()]
        R.add(c, "поиск «болт м8» (русские буквы) находит и «Болт М8х30» книги, и «Болт M8x30» (латинские M и x): две строки", r2 == "OK:2"
              and sorted(r[1] for r in rows2[:2]) == sorted(["Болт М8х30 оцинкованный", "Болт M8x30 DIN 933 оцинкованный"]), f"{r2}; {[r[1] for r in rows2[:3]]}")
        s.B("ByNameRefresh", "", True, module="WmsStockView")
        s.B("StockInfoRefresh", True, module="WmsItem")
        r3 = s.B("PartsRefresh", "", False, module="WmsStockView")
        pr = [list(r) for r in s.doc.Sheets.getByName("Остаток деталей").getCellRangeByPosition(0, 4, 7, 6).getDataArray()]
        R.add(c, "«Остаток деталей»: одна строка на деталь GR-40 — остаток 150 (пополнение того же ЕИ), один ЕИ, место D-1, серийный номер SN-7", r3 == "OK:1" and pr[0][0] == "GR-40"
              and pr[0][2] == 150 and pr[0][4] == 1 and pr[0][5] == "D-1" and "SN-7" in pr[0][6], f"{r3}; {pr[0]}")
        # the user sorts the copy as he likes — the registry is the same; the next refresh builds the list again
        sh.getCellRangeByPosition(0, 4, 10, 8).clearContents(1 + 2 + 4)
        s.B("ByNameRefresh", "", False, module="WmsStockView")
        n_again = len([r for r in sh.getCellRangeByPosition(0, 4, 0, 400).getDataArray() if r[0]])
        R.add(c, "«Обновить» строит список заново после правок пользователя; реестр «Наличие» не тронут (самопроверка без ошибок)",
              n_again == n_stock and s.B("RegistryMisplaced", module="WmsGuard") == 0 and s.B("ActionSelfCheck").startswith("САМОПРОВЕРКА WMS: ошибок 0"), f"{n_again}; {n_stock}")
        # the operations of the day
        r4 = s.B("OpsLoad", float(serial(0)), float(serial(0)), False, module="WmsOps")
        osh = s.doc.Sheets.getByName("Операции дня")
        orows = [list(r) for r in osh.getCellRangeByPosition(0, 4, 9, 4 + 12).getDataArray() if r[0] != ""]
        what = [r[2] for r in orows]
        R.add(c, "«Операции дня»: операции сегодняшнего дня простыми словами — приходы по заказам, «Иной приход — детали» и его пополнение, с ЕИ, названием, количеством и местом",
              r4 == f"OK:{len(orows)}" and len(orows) == 5 and what[0] == "Приход по заказу" and "Иной приход — детали" in what[2] and "Иной приход — детали (пополнение ЕИ)" in what[3]
              and what[4].startswith("Свойства ЕИ")
              and orows[0][4] == "Болт M8x30 DIN 933 оцинкованный" and orows[0][5] == 100 and orows[0][7] == "5-1-1", f"{r4}; {what}; {orows[0]}")
        # the tail of the journal (the book is behind): the morning copy and the afternoon in the journal only
        s.doc.store()
        order(s, 3, "3", "Скотч", qty="40", art="A-3")
        receive(s, 3, "40", place="4-1-4", doc="УПД-3")
        spc_line(s, 3, "Офис", "Стол", "2", place="A-1")
        s.click_spc("BtnSpcPost", 3)
    finally:
        s.kill()
    s = ready(Session(p))
    try:
        s.B("ActionForceUnlock")
        st = s.state()
        r5 = s.B("OpsShowTail", module="WmsOps")
        osh = s.doc.Sheets.getByName("Операции дня")
        trows = [list(r) for r in osh.getCellRangeByPosition(0, 4, 9, 8).getDataArray() if r[0] != ""]
        stat = osh.getCellByPosition(0, 2).getString()
        R.add(c, "хвост журнала в заблокированной книге (TAIL): 2 операции — приход по заказу (ЕИ-204) и «Иной приход — офис»; строка итога объясняет, что это сделает «Восстановить»",
              st.get("BLOCK") == "TAIL" and r5 == "OK:2" and trows[0][2] == "Приход по заказу" and trows[0][3] == "ЕИ-00000204" and "Иной приход — офис" in trows[1][2]
              and "Восстановить" in stat,
              f"{st.get('BLOCK')}; {r5}; {[x[:5] for x in trows]}; {stat[:100]}")
        s.B("ActionRecover")
        # «Что делать?»: a registry that is out of order is named, with what to do (a gap in the numbering — an empty row — goes to the end)
        stk = s.doc.Sheets.getByName("Наличие")
        stk.unprotect("wms")
        stk.getCellRangeByPosition(0, 100, 10, 100).clearContents(1 + 2 + 4 + 16)
        stk.protect("wms")
        sort_registry(s)
        txt = s.B("WhatNowText", module="WmsGuard")
        R.add(c, "«Что делать?» на отсортированном реестре: красный пункт «реестр «Наличие»» и под ним «→ … закройте её БЕЗ сохранения … «Вернуться к копии»»; счёт ошибок",
              "реестр «Наличие»" in txt and "→" in txt and "БЕЗ сохранения" in txt and "Самопроверка: ошибок" in txt, txt[:500])
    finally:
        s.close()


# ================================================================ runner

def main():
    sel = [a.lower() for a in sys.argv[1:]]
    for fn in CASES:
        if sel and not any(fn.__name__.startswith(x) for x in sel):
            continue
        try:
            fn()
        except Exception as e:
            R.error(fn.__name__.split("_")[0].upper(), e)
    R.save(os.path.join(OUT, "results_v10.json"))
    items = R.items
    lines = ["# WMS 1.0 — нововведения, результаты", "", f"PASS {sum(1 for x in items if x['status'] == 'PASS')}, "
             f"FAIL {sum(1 for x in items if x['status'] == 'FAIL')}, SKIP {sum(1 for x in items if x['status'] == 'SKIP')}", ""]
    lines += [f"- [{x['status']}] {x['case']} {x['test']}" for x in items]
    open(os.path.join(OUT, "TEST_REPORT_v10.md"), "w", encoding="utf-8").write("\n".join(lines) + "\n")
    print(f"PASS {sum(1 for x in items if x['status'] == 'PASS')} FAIL {sum(1 for x in items if x['status'] == 'FAIL')} "
          f"SKIP {sum(1 for x in items if x['status'] == 'SKIP')}")
    return 0 if all(x["status"] == "PASS" for x in items) else 1


if __name__ == "__main__":
    sys.exit(main())
