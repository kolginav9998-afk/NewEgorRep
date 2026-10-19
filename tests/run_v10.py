"""WMS 1.0 — the innovations chosen by the storekeeper (stages B–M of docs/PLAN_1_0.md): automated checks on books built
from the sources, through the macros of the buttons (windows answered by the test seams TestUiAuto / TestPrimeChoice).

    WMS_TEST_OUT=/tmp/wms_v10 python3 tests/run_v10.py [case ...]

Every case builds its own fresh test book in WMS_TEST_OUT/cases/<case>/. Results: results_v10.json, TEST_REPORT_v10.md.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Results, Session, new_wms  # noqa: E402
from run_prime import TODAY, issue, order, put, receive, serial, topology  # noqa: E402

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
