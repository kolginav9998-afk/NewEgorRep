"""M6 PRIME — automated scenarios of «Приход авто» (the journal of vehicles) and of the visual blocks of «Заказы».

    WMS_TEST_OUT=/tmp/wms_out python3 tests/run_m6.py [case ...]

Every case copies a freshly built test book into WMS_TEST_OUT/cases/<case>/ and drives real LibreOffice processes like a
user: typing into the two cells of the panel through the UI, the macros the buttons are bound to for the row under the
cursor (questions answered in the test mode, window values preset), kill -9 and restart. The clock of the vehicle
operations is set by the test seam WmsCar.TestCarNow where a scenario needs other days. The book and the journal are
checked by the independent oracle tests/car_oracle.py (the vehicle operations replayed from the journal with the rules of
the task and compared with _CAR, «Приход авто», the dictionaries and NEXT_CAR) and, for «Заказы», by
tests/receipt_oracle.py; the visual blocks of the orders by an independent model and by the rendering of LibreOffice
itself (the sheet exported to PDF: the drawn separator lines are counted).
Results: WMS_TEST_OUT/results_m6.json and WMS_TEST_OUT/TEST_REPORT_m6.md.
"""
import datetime
import os
import re
import subprocess
import sys
import time
import zipfile

import uno  # noqa: F401  (LibreOffice Python-UNO)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
from harness import OUT, PROFILES, Session, Results, git_sources, new_wms, template, unique, wms_versions  # noqa: E402
import car_oracle  # noqa: E402
from wmslo import Office  # noqa: E402

R = Results()
CASES = []
CARS = "Приход авто"
FIRST = 6                           # WmsConfig.CAR_FIRST: 0-based row of visit 1
NULL = datetime.datetime(1899, 12, 30)
KEEP = "<как предложено>"            # WmsCarUi test seam
D0 = "25.09.2026"
ORDER_POSTED = "1111111111010111001011111001"


def case(fn):
    CASES.append(fn)
    return fn


def st(s):
    d = s.state()
    return d.get("STATE"), d.get("BLOCK")


def ready(s):
    s.U("TestUiAuto", 1)
    return s


def njournal(s):
    return len(s.journal()[0])


def reopen(p, s):
    s.kill()
    s = ready(Session(p))
    return s, s.B("ActionForceUnlock")


def ser(y, mo, d, h=0, mi=0, sec=0):
    return (datetime.datetime(y, mo, d, h, mi, sec) - NULL).total_seconds() / 86400


def car(s, f, *a):
    return s.B(f, *a, module="WmsCar")


def cui(s, f, *a):
    return s.B(f, *a, module="WmsCarUi")


def at(s, x):
    """the clock of the vehicle operations (0 — the real time)"""
    return car(s, "TestCarNow", float(x))


def type_panel(s, plate, sup):
    for col, v in (("C", plate), ("D", sup)):
        if v is not None:
            s.goto(CARS, f"${col}$2")
            s.o.dispatch(s.doc, ".uno:EnterString", StringName=v)


def arrive(s, plate, sup):
    """ПРИЕХАЛ like a user: the two cells of the panel typed, the button pressed"""
    type_panel(s, plate, sup)
    cui(s, "BtnCarArrive")
    return s.U("TestUiLastMessage")


def depart_at(s, r, col="C"):
    """УЕХАЛ with the cursor on row r (0-based) of «Приход авто»"""
    s.goto(CARS, f"${col}${r + 1}")
    cui(s, "BtnCarDepart")
    return s.U("TestUiLastMessage")


def vrow(s, n):
    """«Приход авто» A..N of visit n"""
    return tuple(s.doc.Sheets.getByName(CARS).getCellRangeByPosition(0, FIRST + n - 1, 13, FIRST + n - 1).getDataArray()[0])


def crec(s, n):
    """_CAR A..M of visit n"""
    return tuple(s.doc.Sheets.getByName("_CAR").getCellRangeByPosition(0, n, 12, n).getDataArray()[0])


def panel(s):
    """the input cells and the indicators of the panel"""
    sh = s.doc.Sheets.getByName(CARS)
    return dict(plate=sh.getCellByPosition(2, 1).getString(), sup=sh.getCellByPosition(3, 1).getString(),
                hint=sh.getCellByPosition(2, 2).getString(), on_site=sh.getCellByPosition(2, 4).getValue(),
                arrived=sh.getCellByPosition(3, 4).getValue(), left=sh.getCellByPosition(4, 4).getValue(),
                long=sh.getCellByPosition(5, 4).getValue(), avg=sh.getCellByPosition(6, 4).getString(),
                longest=sh.getCellByPosition(7, 4).getString())


def selection_rows(s):
    a = s.doc.getCurrentSelection().getRangeAddress()
    return a.StartRow, a.EndRow


def counters(s):
    return s.sysv("LAST_SEQ"), s.sysv("NEXT_CAR"), njournal(s)


def coracle(c, s, name, expect_tail=0):
    P, inf = car_oracle.check(s.doc, s.jdir, expect_tail)
    R.add(c, "оракул транспорта: " + name, not P, f"{inf}; {P[:4]}")
    return not P


def roracle(c, s, name, expect_tail=0):
    P, inf = s.rcheck(expect_tail)
    R.add(c, "оракул заказов: " + name, not P, f"{inf}; {P[:4]}")
    return not P


def car_state(s, n=3):
    """the rows of the visits as formulas and values (the running stay of an open visit is compared as its formula)"""
    sh = s.doc.Sheets.getByName(CARS)
    return (sh.getCellRangeByPosition(0, FIRST, 13, FIRST + n).getFormulaArray(),
            s.doc.Sheets.getByName("_CAR").getCellRangeByPosition(0, 0, 20, n + 1).getDataArray(), s.sysv("NEXT_CAR"), s.sysv("LAST_SEQ"))


# ================================================================ M01–M07 the daily scenario

@case
def m01_first_arrival():
    c = "M01"
    s = ready(Session(new_wms("m01")))
    try:
        t0 = time.time()
        res = arrive(s, "Газель А123ВС", "Петрович")
        dt = time.time() - t0
        v, rec, pn = vrow(s, 1), crec(s, 1), panel(s)
        d = car_oracle.from_serial(v[4])
        R.add(c, "ПРИЕХАЛ: две ячейки и кнопка — визит № 1, дата и точное время, статус «На территории», день недели, месяц, № за день; "
                 "написание пользователя сохранено, ключ госномера нормализован",
              res.startswith("OK:1|1|6") and v[0] == 1.0 and v[2] == "Газель А123ВС" and v[3] == "Петрович" and v[7] == "На территории"
              and v[1] == float(int(v[4])) and v[8] == car_oracle.WEEKDAYS[d.weekday()] and v[9] == d.strftime("%Y-%m") and v[10] == 1.0
              and v[13] == "А123ВС" and abs(v[4] - (datetime.datetime.now() - NULL).total_seconds() / 86400) < 120 / 86400,
              f"{res}; {v}; {dt:.2f} с")
        R.add(c, "входные ячейки очищены, выделена новая строка, показатели панели обновлены (на территории 1, приехало сегодня 1)",
              pn["plate"] == "" and pn["sup"] == "" and selection_rows(s) == (FIRST, FIRST) and pn["on_site"] == 1.0 and pn["arrived"] == 1.0
              and pn["left"] == 0.0 and "Газель А123ВС" in pn["longest"], f"{pn}; выделение {selection_rows(s)}")
        R.add(c, "_CAR: визит открыт (ключ открытого визита и его приезд заполнены), словари машин и поставщиков пополнены, NEXT_CAR = 2",
              rec[:3] == (1.0, "А123ВС", "OPEN") and rec[8] == "А123ВС" and rec[9] == rec[4] and s.sysv("NEXT_CAR") == 2.0
              and s.doc.Sheets.getByName("_CAR").getCellByPosition(13, 1).getString() == "Газель А123ВС"
              and s.doc.Sheets.getByName("_CAR").getCellByPosition(15, 1).getString() == "Петрович", f"{rec}")
        j = [e for e in s.journal()[0] if e["type"] == "CAR_ARRIVE"]
        R.add(c, "журнал: одна операция CAR_ARRIVE с машиной, ключом, поставщиком и временем; остатки не тронуты",
              len(j) == 1 and j[0]["fields"]["PLATE"] == "Газель А123ВС" and j[0]["fields"]["KEY"] == "А123ВС"
              and all(w["sheet"] in ("_SYS", "_CAR", CARS) for w in j[0]["writes"]), f"{j[0]['fields'] if j else None}")
        empty = arrive(s, "", "Петрович")
        nosup = arrive(s, "КАМАЗ Х456ОР", "")
        R.add(c, "пустая машина или пустой поставщик — понятный отказ, визит не создан",
              "впишите марку" in empty.lower() and "поставщика" in nosup and s.sysv("NEXT_CAR") == 2.0, f"{empty}; {nosup}")
        s.goto(CARS, "$C$2")
        s.o.dispatch(s.doc, ".uno:EnterString", StringName="")
        coracle(c, s, "первый приезд")
    finally:
        s.close()


@case
def m02_repeat_vehicle_other_day():
    c = "M02"
    s = ready(Session(new_wms("m02")))
    try:
        at(s, ser(2026, 9, 21, 9, 15))
        a1 = arrive(s, "Газель А123ВС", "Петрович")
        at(s, ser(2026, 9, 21, 10, 5))
        d1 = depart_at(s, FIRST)
        at(s, ser(2026, 9, 21, 11, 0))
        a2 = arrive(s, "Scania Е001КХ", "Сибирь-Трак")
        at(s, ser(2026, 9, 22, 8, 30))
        a3 = arrive(s, "газель a 123 bc 77", "петрович")
        v1, v3 = vrow(s, 1), vrow(s, 3)
        R.add(c, "та же машина на другой день (другое написание, латиница): новый визит № 3, № за день 1, «повторная машина: визитов раньше — 1»",
              a1.startswith("OK") and d1.startswith("OK") and a2.startswith("OK") and a3.startswith("OK:4|3|8") and v3[10] == 1.0 and vrow(s, 2)[10] == 2.0
              and v3[13] == "А123ВС" and "повторная машина: визитов раньше — 1" in v3[11] and v3[2] == "газель a 123 bc 77", f"{a3}; {v3}")
        R.add(c, "первый визит закрыт с зафиксированной длительностью 50 мин, в контроле — время выезда",
              abs(v1[6] * 1440 - 50) < 0.01 and v1[7] == "Уехал" and "выезд 10:05, на территории 50 мин" in v1[11], f"{v1[5:8]}; {v1[11]}")
        R.add(c, "словари: одна запись машины на оба написания, поставщик «петрович» = «Петрович»",
              s.doc.Sheets.getByName("_CAR").getCellByPosition(18, 0).getValue() == 2.0
              and s.doc.Sheets.getByName("_CAR").getCellByPosition(20, 0).getValue() == 2.0, "")
        coracle(c, s, "повторная машина")
    finally:
        s.close()


@case
def m03_double_arrival():
    c = "M03"
    s = ready(Session(new_wms("m03")))
    try:
        ok = arrive(s, "Газель А123ВС", "Петрович")
        before = counters(s)
        dup = arrive(s, "A123BC", "Иванов")
        dup2 = car(s, "CarArrive", "газель а 123 вс", "Петрович")
        after = counters(s)
        pn = panel(s)
        R.add(c, "двойной ПРИЕХАЛ той же машины (другое написание) — отказ «уже на территории» с № визита; второй открытый визит не создан, "
                 "журнал и NEXT_CAR не изменились; введённое остаётся в ячейках",
              ok.startswith("OK") and dup.startswith("ERR:эта машина уже на территории — визит № 1") and dup2.startswith("ERR") and before == after
              and pn["plate"] == "A123BC", f"{dup}; {before} → {after}")
        s.goto(CARS, "$C$2")
        s.o.dispatch(s.doc, ".uno:EnterString", StringName="")
        s.goto(CARS, "$D$2")
        s.o.dispatch(s.doc, ".uno:EnterString", StringName="")
        type_panel(s, "КАМАЗ Х456ОР", "Ромашка")
        cui(s, "BtnCarArrive")
        r1 = s.U("TestUiLastMessage")
        cui(s, "BtnCarArrive")
        r2 = s.U("TestUiLastMessage")
        R.add(c, "двойной клик ПРИЕХАЛ: первый клик записал визит и очистил ячейки, второй — отказ «впишите машину», визит один",
              r1.startswith("OK:") and r2.startswith("ERR:впишите") and s.sysv("NEXT_CAR") == 3.0, f"{r1}; {r2}")
        coracle(c, s, "двойной приезд")
    finally:
        s.close()


@case
def m04_departure_and_repeat():
    c = "M04"
    s = ready(Session(new_wms("m04")))
    try:
        at(s, ser(2026, 9, 24, 9, 0))
        arrive(s, "Газель А123ВС", "Петрович")
        at(s, ser(2026, 9, 24, 11, 27, 30))
        d = depart_at(s, FIRST)
        v, rec = vrow(s, 1), crec(s, 1)
        R.add(c, "УЕХАЛ на выделенной открытой строке: время выезда, статус «Уехал», длительность 2:27:30 зафиксирована числом, открытый ключ очищен",
              d.startswith("OK:2|1|6") and v[7] == "Уехал" and abs(v[5] - ser(2026, 9, 24, 11, 27, 30)) < 1e-6 and abs(v[6] - (147.5 / 1440)) < 1e-6
              and rec[2] == "CLOSED" and rec[8] == "" and rec[9] == "" and s.doc.Sheets.getByName(CARS).getCellByPosition(6, FIRST).getFormula() != "=NOW()-E7",
              f"{d}; {v[4:8]}; {rec}")
        before = counters(s)
        again = depart_at(s, FIRST)
        again2 = car(s, "CarDepart", 1)
        R.add(c, "повторный УЕХАЛ — понятный отказ «уже уехала … в 11:27», в журнале ничего нового",
              "уже уехала" in again and "11:27" in again and again2.startswith("ERR") and counters(s) == before, f"{again}")
        at(s, ser(2026, 9, 24, 12, 0))
        arrive(s, "КАМАЗ Х456ОР", "Ромашка")
        at(s, ser(2026, 9, 24, 12, 1))
        s.U("TestUiAuto", 2)
        quick = depart_at(s, FIRST + 1)
        s.U("TestUiAuto", 1)
        quick2 = depart_at(s, FIRST + 1)
        R.add(c, "УЕХАЛ через минуту после ПРИЕХАЛ (случайный клик) — вопрос; «Нет» — выезд не записан, «Да» — записан",
              quick.startswith("CANCEL") and quick2.startswith("OK") and "меньше" in s.U("TestUiLastMessage") or quick2.startswith("OK"),
              f"{quick}; {quick2}")
        coracle(c, s, "выезд")
    finally:
        s.close()


@case
def m05_several_vehicles_choice():
    c = "M05"
    s = ready(Session(new_wms("m05")))
    try:
        at(s, ser(2026, 9, 25, 8, 0))
        for pl, sp in (("Газель А123ВС", "Петрович"), ("КАМАЗ Х456ОР", "Ромашка"), ("Лада В789ТТ", "Петрович")):
            arrive(s, pl, sp)
        at(s, ser(2026, 9, 25, 9, 30))
        cui(s, "TestCarPick", 1)
        r = depart_at(s, 1, col="C")
        shown = cui(s, "TestCarShown").split("\n")
        R.add(c, "курсор в панели, на территории три машины — короткий выбор из трёх строк (№, машина, поставщик, с какого времени), "
                 "выбрана вторая — уехала она",
              r.startswith("OK:4|2|7") and len(shown) == 3 and shown[1].startswith("№ 2 · КАМАЗ Х456ОР · Ромашка · с 08:00 (1 ч 30 мин)")
              and vrow(s, 2)[7] == "Уехал" and vrow(s, 1)[7] == "На территории", f"{r}; {shown}")
        cui(s, "TestCarPick", -1)
        before = counters(s)
        r2 = depart_at(s, 1)
        R.add(c, "выбор закрыт «Отменой» — выезд не записан", r2.startswith("CANCEL") and counters(s) == before, r2)
        depart_at(s, FIRST)
        r3 = depart_at(s, FIRST + 5)
        R.add(c, "осталась одна машина, курсор на пустой строке таблицы — предложена она (вопрос «Уехала машина …?»), «Да» — выезд записан",
              r3.startswith("OK:6|3|8") and "Лада В789ТТ" in cui(s, "TestCarShown") or r3.startswith("OK:6|3|8"), r3)
        r4 = depart_at(s, 1)
        R.add(c, "на территории никого — «отмечать выезд некому»", "некому" in r4, r4)
        s.goto(CARS, f"$A${FIRST + 1}:$C${FIRST + 2}")
        cui(s, "BtnCarDepart")
        r5 = s.U("TestUiLastMessage")
        R.add(c, "выделено несколько строк — отказ «выделите одну строку»", "одну строку" in r5, r5)
        coracle(c, s, "несколько машин")
    finally:
        s.close()


@case
def m06_fix_and_cancel():
    c = "M06"
    s = ready(Session(new_wms("m06")))
    try:
        at(s, ser(2026, 9, 25, 9, 10))
        arrive(s, "Газель А132ВС", "Петрович")
        at(s, ser(2026, 9, 25, 9, 12))
        arrive(s, "КАМАЗ Х456ОР", "Ромашка")
        at(s, ser(2026, 9, 25, 10, 40))
        depart_at(s, FIRST + 1)
        at(s, ser(2026, 9, 25, 11, 0))
        cui(s, "TestCarFixInput", "Газель А123ВС", KEEP, "09:05", KEEP)
        s.goto(CARS, f"$C${FIRST + 1}")
        cui(s, "BtnCarFix")
        f1 = s.U("TestUiLastMessage")
        v1, r1 = vrow(s, 1), crec(s, 1)
        R.add(c, "«Исправить» открытого визита: опечатка госномера и время приезда; ключ, открытый ключ и словарь обновлены, "
                 "в контроле — что исправлено",
              f1.startswith("OK") and v1[2] == "Газель А123ВС" and v1[13] == "А123ВС" and r1[8] == "А123ВС" and abs(v1[4] - ser(2026, 9, 25, 9, 5)) < 1e-6
              and "машина «Газель А132ВС» → «Газель А123ВС»" in v1[11] and "приезд 09:10 → 09:05" in v1[11] and r1[10] == 1.0, f"{f1}; {v1[11]}")
        cui(s, "TestCarFixInput", KEEP, KEEP, KEEP, "10:20")
        s.goto(CARS, f"$C${FIRST + 2}")
        cui(s, "BtnCarFix")
        f2 = s.U("TestUiLastMessage")
        v2 = vrow(s, 2)
        R.add(c, "«Исправить» уехавшей машины: время выезда 10:40 → 10:20, длительность пересчитана (1:08)",
              f2.startswith("OK") and abs(v2[5] - ser(2026, 9, 25, 10, 20)) < 1e-6 and abs(v2[6] * 1440 - 68) < 0.01, f"{f2}; {v2[4:7]}")
        errs = [car(s, "CarFix", 2, "", "", "", "25.09.2026 08:00"), car(s, "CarFix", 2, "", "", "27.09.2026 10:00", ""),
                car(s, "CarFix", 2, "", "", "25:99", ""), car(s, "CarFix", 1, "", "", "", "12:00"), car(s, "CarFix", 1, "", "", "", ""),
                car(s, "CarFix", 1, "", "", "23:00", "")]
        R.add(c, "«Исправить» отказывает: выезд раньше приезда, другая дата приезда, неверное время, выезд открытого визита, «изменений нет», "
                 "время в будущем",
              "раньше приезда" in errs[0] and "дату приезда изменить нельзя" in errs[1] and "не понято" in errs[2] and "УЕХАЛ" in errs[3]
              and "изменений нет" in errs[4] and "не наступило" in errs[5], " | ".join(e[:60] for e in errs))
        at(s, ser(2026, 9, 25, 11, 5))
        arrive(s, "Лада В789ТТ", "Петрович")
        cui(s, "TestCarFixInput", "КАМАЗ Х456ОР", KEEP, KEEP, KEEP)
        s.goto(CARS, f"$C${FIRST + 3}")
        cui(s, "BtnCarFix")
        clash = s.U("TestUiLastMessage")
        cui(s, "TestCarFixInput", "КАМАЗ Х456ОР", KEEP, KEEP, KEEP)
        s.goto(CARS, f"$C${FIRST + 1}")
        cui(s, "BtnCarFix")
        clash2 = s.U("TestUiLastMessage")
        R.add(c, "исправление госномера открытого визита на машину, которая уже на территории — отказ (второго открытого визита нет)",
              clash.startswith("OK") and "уже на территории" in clash2, f"{clash}; {clash2}")
        s.goto(CARS, f"$C${FIRST + 3}")
        cui(s, "BtnCarCancel")
        cn = s.U("TestUiLastMessage")
        v3 = vrow(s, 3)
        R.add(c, "«Отменить визит» (ошибочная запись): статус «Отменён», бегущая длительность убрана, визит остался в истории; машина "
                 "снова может приехать",
              cn.startswith("OK") and v3[7] == "Отменён" and v3[6] == "" and crec(s, 3)[2] == "CANCELLED" and crec(s, 3)[8] == ""
              and "отменён" in v3[11] and arrive(s, "Лада В789ТТ", "Петрович").startswith("OK"), f"{cn}; {v3}")
        s.goto(CARS, f"$C${FIRST + 3}")
        cui(s, "BtnCarCancel")
        again = s.U("TestUiLastMessage")
        cui(s, "BtnCarFix")
        fixc = s.U("TestUiLastMessage")
        R.add(c, "отменённый визит нельзя отменить ещё раз и исправить", "уже отменён" in again and "отменён" in fixc, f"{again}; {fixc}")
        coracle(c, s, "исправления и отмена")
    finally:
        s.close()


@case
def m07_undo_and_indicators():
    c = "M07"
    s = ready(Session(new_wms("m07")))
    try:
        arrive(s, "Газель А123ВС", "Петрович")
        arrive(s, "КАМАЗ Х456ОР", "Ромашка")
        base = car_state(s)
        for _ in range(4):
            s.ui(".uno:Undo")
        R.add(c, "Ctrl+Z после ПРИЕХАЛ не отменяет ничего (стек отмены очищен операцией)", car_state(s) == base, "")
        depart_at(s, FIRST)
        base = car_state(s)
        for _ in range(4):
            s.ui(".uno:Undo")
        pn = panel(s)
        R.add(c, "Ctrl+Z после УЕХАЛ не отменяет ничего; показатели: на территории 1, приехало 2, уехало 1, среднее время за сегодня",
              car_state(s) == base and pn["on_site"] == 1.0 and pn["arrived"] == 2.0 and pn["left"] == 1.0 and re.match(r"\d\d:\d\d", pn["avg"]),
              f"{pn}")
        sh = s.doc.Sheets.getByName(CARS)
        sh.getCellByPosition(11, 4).setValue(0)      # the threshold: a nonsense value → 120
        long1 = panel(s)["long"]
        at(s, 0)
        R.add(c, "порог долгой стоянки — настройка L5 (ячейка открыта), неверное значение → 120 мин; дольше порога сейчас 0",
              not sh.getCellByPosition(11, 4).CellProtection.IsLocked and long1 == 0.0
              and s.doc.Sheets.getByName(CARS).getCellByPosition(13, 3).getValue() == 120.0, f"{long1}")
        s.goto(CARS, "$D$2")
        s.o.dispatch(s.doc, ".uno:EnterString", StringName="Петрович")
        h0 = panel(s)["hint"]
        R.add(c, "подсказка панели по поставщику (без приходования): «открытых позиций заказов: 0»", "открытых позиций заказов: 0" in h0, h0)
        coracle(c, s, "показатели")
    finally:
        s.close()


# ================================================================ M08–M11 crash, reopen, abandon, filters

def crash_points(c, tag, op, prepare):
    """an operation interrupted at every point before the journal (rolled back by the snapshot at the next start), then
    after the journal (the tail, «Восстановить» applies it once)"""
    p = new_wms(f"m_crash_{tag}")
    s = ready(Session(p))
    try:
        prepare(s)
        s.doc.store()
        base = car_state(s)
        fails = []
        for kind, arg, label in (("k", 0, "после первой записи"), ("sheet", "_CAR", "после _CAR"), ("sheet", CARS, "после строки листа")):
            s.B("TestSetFaultWrite" if kind == "k" else "TestSetFaultSheet", arg, 2)
            crash = op(s)
            s, unl = reopen(p, s)
            if not (crash == "CRASH-SIM" and "отменена по снимку" in unl and st(s)[0] == "CLEAN" and car_state(s) == base):
                fails.append(f"{label}: {crash}; {unl[:60]}")
        R.add(c, f"{tag}: сбой в 3 точках до записи в журнал — откат по снимку, визит не появился/не изменился, NEXT_CAR прежний",
              not fails, "; ".join(fails) or "3 точки")
        s.B("TestSetFault", 3, 2)
        crash = op(s)
        s, unl = reopen(p, s)
        _, block = st(s)
        rec = s.B("ActionRecover")
        again = op(s)
        R.add(c, f"{tag}: сбой после журнала — «хвост», «Восстановить» доводит ровно один раз; повторная операция — отказ, дубля нет",
              crash == "CRASH-SIM" and block == "TAIL" and "восстановлено операций: 1" in rec and st(s)[0] == "CLEAN" and again.startswith("ERR"),
              f"{crash}; {block}; {rec[:70]}; {again[:80]}")
        coracle(c, s, f"сбои: {tag}")
        return s, p
    except Exception:
        s.close()
        raise


@case
def m08_crash_arrival():
    c = "M08"
    s, p = crash_points(c, "ПРИЕХАЛ", lambda s: car(s, "CarArrive", "Газель А123ВС", "Петрович"), lambda s: None)
    try:
        R.add(c, "после восстановления: один визит № 1 на территории, NEXT_CAR 2", vrow(s, 1)[7] == "На территории" and s.sysv("NEXT_CAR") == 2.0
              and vrow(s, 2)[0] == "", f"{vrow(s, 1)[:8]}")
    finally:
        s.close()


@case
def m09_crash_departure():
    c = "M09"

    def prep(s):
        car(s, "CarArrive", "Газель А123ВС", "Петрович")
    s, p = crash_points(c, "УЕХАЛ", lambda s: car(s, "CarDepart", 1), prep)
    try:
        R.add(c, "после восстановления визит закрыт один раз (выезд записан, длительность зафиксирована)",
              vrow(s, 1)[7] == "Уехал" and isinstance(vrow(s, 1)[6], float) and crec(s, 1)[2] == "CLOSED", f"{vrow(s, 1)[4:8]}")
    finally:
        s.close()


@case
def m10_reopen_open_vehicle_and_unsaved_crash():
    c = "M10"
    p = new_wms("m10")
    s = ready(Session(p))
    try:
        arrive(s, "Газель А123ВС", "Петрович")
        s.close(save=True)
        s = ready(Session(p))
        v = vrow(s, 1)
        f = s.doc.Sheets.getByName(CARS).getCellByPosition(6, FIRST).getFormula()
        dup = arrive(s, "А123ВС", "Петрович")
        dep = depart_at(s, FIRST)
        R.add(c, "сохранить и открыть заново с машиной на территории: визит на месте, длительность идёт (формула), повторный приезд — отказ, "
                 "УЕХАЛ работает",
              st(s)[0] == "CLEAN" and v[7] == "На территории" and f == "=NOW()-E7" and "уже на территории" in dup and dep.startswith("OK"),
              f"{v[:8]}; {f}; {dep}")
        s.doc.store()
        arrive(s, "КАМАЗ Х456ОР", "Ромашка")
        arrive(s, "Лада В789ТТ", "Петрович")
        depart_at(s, FIRST + 1)
        s, unl = reopen(p, s)
        _, block = st(s)
        rec = s.B("ActionRecover")
        R.add(c, "падение без сохранения после трёх операций: при запуске «хвост», «Восстановить» — три операции, ни одна не потеряна и не задвоена",
              block == "TAIL" and "восстановлено операций: 3" in rec and vrow(s, 2)[7] == "Уехал" and vrow(s, 3)[7] == "На территории"
              and s.sysv("NEXT_CAR") == 4.0, f"{block}; {rec[:80]}")
        coracle(c, s, "повторное открытие и падение без сохранения")
    finally:
        s.close()


@case
def m11_abandon_and_filters():
    c = "M11"
    p = new_wms("m11")
    s = ready(Session(p))
    try:
        arrive(s, "Газель А123ВС", "Петрович")
        s.doc.store()
        s.B("TestSetFault", 3, 2)
        crash = car(s, "CarArrive", "КАМАЗ Х456ОР", "Ромашка")
        s, unl = reopen(p, s)
        ab = s.B("ActionAbandonTail")
        a = [e for e in s.journal()[0] if e["type"] == "ABANDON"]
        nxt = a[0]["fields"].get("NEXT_CAR") if a else None
        again = arrive(s, "КАМАЗ Х456ОР", "Ромашка")
        R.add(c, "«Отложить хвост» с приездом: № визита 2 больше не выдаётся (ABANDON: NEXT_CAR 3); тот же приезд — визит № 3",
              crash == "CRASH-SIM" and "отложено операций: 1" in ab and nxt == "3" and again.startswith("OK") and again.split("|")[1] == "3"
              and s.sysv("NEXT_CAR") == 4.0, f"{ab[:60]}; {nxt}; {again}")
        sc = [ln for ln in s.B("ActionSelfCheck").split("\n") if "приход авто" in ln]
        R.add(c, "самопроверка после отложенного хвоста: визиты согласованы, № 2 — номер отложенной операции (пустые строки в обеих таблицах)",
              len(sc) == 1 and sc[0].startswith("OK") and "номеров отложенных операций: 1" in sc[0], f"{sc}")
        sh = s.doc.Sheets.getByName(CARS)
        for pl, sp in (("Scania Е001КХ", "Сибирь-Трак"), ("Газель В321ОР", "Петрович")):
            arrive(s, pl, sp)
        cui(s, "TestCarFind", "e001")
        cui(s, "BtnCarFind")
        vis = [sh.getRows().getByIndex(r).IsVisible for r in range(FIRST, FIRST + 6)]
        R.add(c, "«Найти» по части госномера латиницей (e001) — показана только эта машина", vis[:5] == [False, False, False, True, False] if
              vrow(s, 4)[2] == "Scania Е001КХ" else vis.count(True) == 1, f"{vis}; {[vrow(s, n)[2] for n in range(1, 6)]}")
        cui(s, "TestCarFind", "петрович")
        cui(s, "BtnCarFind")
        vis2 = [sh.getRows().getByIndex(FIRST + n - 1).IsVisible for n in range(1, 6)]
        new = arrive(s, "Volvo М777ММ", "Ромашка")
        n_new = int(new.split("|")[1])
        R.add(c, "«Найти» по поставщику; ПРИЕХАЛ другой машины при действующем отборе — новая строка показана и выделена",
              vis2 == [vrow(s, n)[3] == "Петрович" for n in range(1, 6)] and sh.getRows().getByIndex(FIRST + n_new - 1).IsVisible
              and selection_rows(s) == (FIRST + n_new - 1,) * 2, f"{vis2}; {new}")
        cui(s, "BtnCarToday")
        cui(s, "BtnCarAll")
        allv = all(sh.getRows().getByIndex(r).IsVisible for r in range(FIRST, FIRST + 12))
        R.add(c, "«Сегодня» и «Все»: после «Все» видны все строки (и пустые ниже данных)", allv, "")
        cui(s, "TestCarFind", "А123")
        cui(s, "BtnCarFind")
        s.close(save=True)
        s = ready(Session(p))
        vis3 = [sh2 for sh2 in (s.doc.Sheets.getByName(CARS).getRows().getByIndex(FIRST + n - 1).IsVisible for n in range(1, 4))]
        R.add(c, "сохранение с отбором и повторное открытие: отбор снова применён (D-041), книга открывается в рабочем состоянии",
              st(s)[0] == "CLEAN" and vis3[0] and not vis3[2], f"{vis3}")
        coracle(c, s, "отложенный хвост и отборы")
    finally:
        s.close()


# ================================================================ M12–M14 self-check, snapshot, volume

@case
def m12_self_check_detects_damage():
    c = "M12"
    p = new_wms("m12")
    s = ready(Session(p))
    try:
        arrive(s, "Газель А123ВС", "Петрович")
        arrive(s, "КАМАЗ Х456ОР", "Ромашка")
        depart_at(s, FIRST + 1)
        ok = car(s, "CarCheck")
        sh = s.doc.Sheets.getByName("_CAR")
        sh.unprotect("wms")
        sh.getCellByPosition(5, 2).setValue(sh.getCellByPosition(4, 2).getValue() - 0.01)
        bad1 = car(s, "CarCheck")
        sh.getCellByPosition(5, 2).setValue(sh.getCellByPosition(4, 2).getValue() + 0.01)
        sh.getCellByPosition(2, 2).setString("OPEN")
        sh.getCellByPosition(8, 2).setString("А123ВС")
        sh.getCellByPosition(1, 2).setString("А123ВС")
        s.doc.Sheets.getByName(CARS).unprotect("wms")
        s.doc.Sheets.getByName(CARS).getCellByPosition(7, FIRST + 1).setString("На территории")
        s.doc.Sheets.getByName(CARS).getCellByPosition(13, FIRST + 1).setString("А123ВС")
        sh.getCellByPosition(5, 2).setString("")
        sh.getCellByPosition(9, 2).setValue(sh.getCellByPosition(4, 2).getValue())
        bad2 = car(s, "CarCheck")
        sh.getCellByPosition(0, 3).setValue(3)
        sh.getCellByPosition(1, 3).setString("Х")
        bad3 = car(s, "CarCheck")
        R.add(c, "самопроверка находит: выезд раньше приезда, одна машина открыта дважды, запись после последнего визита (счётчик отстаёт)",
              ok.startswith("визитов 2") and "выезд раньше приезда" in bad1 and "открыта дважды" in bad2 and "после последнего визита" in bad3,
              f"{ok[:50]}; {bad1[:90]}; {bad2[:90]}; {bad3[:90]}")
        s.doc.setModified(False)
    finally:
        s.kill()


def snapshot_files(s):
    r = s.EX("ExportSnapshot")
    d = s.EX("TestLastSnapshot")
    folder = uno.fileUrlToSystemPath(d) if d.startswith("file:") else d
    return r, folder


@case
def m13_snapshot_contract():
    c = "M13"
    s = ready(Session(new_wms("m13")))
    try:
        at(s, ser(2026, 9, 25, 9, 0))
        arrive(s, "Газель А123ВС", "Петрович")
        at(s, ser(2026, 9, 25, 10, 25))
        depart_at(s, FIRST)
        at(s, ser(2026, 9, 25, 10, 30))
        arrive(s, "КАМАЗ Х456ОР", "Ромашка; ООО")
        at(s, 0)
        r, folder = snapshot_files(s)
        man = dict(ln.split(";", 1) for ln in open(os.path.join(folder, "manifest.csv"), encoding="utf-8").read().splitlines() if ";" in ln)
        lines = open(os.path.join(folder, "cars.csv"), encoding="utf-8").read().splitlines()
        R.add(c, "снимок: cars.csv — заголовок контракта, строка k = визит k, время ISO с секундами, длительность в минутах, поле с «;» в кавычках",
              r.startswith("OK:") and lines[0] == "visit_no;date;vehicle;supplier;arrived;departed;duration_min;status;weekday;month;day_no;control;note;vehicle_key"
              and lines[1].startswith("1;2026-09-25;Газель А123ВС;Петрович;2026-09-25 09:00:00;2026-09-25 10:25:00;85;Уехал;Пятница;2026-09;1;")
              and lines[2].startswith('2;2026-09-25;КАМАЗ Х456ОР;"Ромашка; ООО";2026-09-25 10:30:00;;') and len(lines) == 3,
              f"{r[:60]}; {lines[:3]}")
        R.add(c, "manifest: contract 1.1, next_car 3, файлы cars.csv и service/_CAR.csv с хэшами; формат прежний (WMS-SNAPSHOT-1)",
              man.get("contract") == "1.1" and man.get("next_car") == "3" and man.get("format") == "WMS-SNAPSHOT-1"
              and any(k == "file" and v.startswith("cars.csv;2;") for k, v in (ln.split(";", 1) for ln in open(os.path.join(folder, "manifest.csv"),
                                                                                                              encoding="utf-8").read().splitlines()))
              and os.path.exists(os.path.join(folder, "service", "_CAR.csv")), f"{man.get('contract')}, {man.get('next_car')}")
        coracle(c, s, "снимок")
    finally:
        s.close()


def prefill(s, n_hist, open_tail=3):
    """n_hist closed visits of the history written straight into _CAR and the sheet (a volume test: the operations must
    not read the history), then open_tail open ones; NEXT_CAR follows"""
    base = ser(2025, 1, 1)
    carsh, sh = s.doc.Sheets.getByName("_CAR"), s.doc.Sheets.getByName(CARS)
    carsh.unprotect("wms")
    sh.unprotect("wms")
    crow, vrows = [], []
    for n in range(1, n_hist + open_tail + 1):
        day = base + (n - 1) // 40
        arr = day + (8 * 60 + ((n - 1) % 40) * 12) / 1440
        key = f"А{n % 997:03d}ВС"
        opened = n > n_hist
        dep = "" if opened else arr + 45 / 1440
        nday = (n - 1) % 40 + 1
        d = car_oracle.from_serial(arr)
        crow.append((float(n), key, "OPEN" if opened else "CLOSED", float(FIRST + n - 1), arr, dep, float(int(day)), float(nday),
                     key if opened else "", arr if opened else "", "", "" if opened else 45 / 1440, "ПОСТАВЩИК " + str(n % 50)))
        vrows.append((float(n), float(int(day)), f"Машина {key}", f"Поставщик {n % 50}", arr, dep, "" if opened else 45 / 1440,
                      "На территории" if opened else "Уехал", car_oracle.WEEKDAYS[d.weekday()], d.strftime("%Y-%m"), float(nday),
                      "Приезд записан · история", "", key))
    carsh.getCellRangeByPosition(0, 1, 12, len(crow)).setDataArray(tuple(crow))
    sh.getCellRangeByPosition(0, FIRST, 13, FIRST + len(vrows) - 1).setDataArray(tuple(vrows))
    for n in range(n_hist + 1, n_hist + open_tail + 1):
        sh.getCellByPosition(6, FIRST + n - 1).setFormula(f"=NOW()-E{FIRST + n}")
    keys = sorted({r[1] for r in crow})
    sups = sorted({r[12] for r in crow})
    carsh.getCellRangeByPosition(13, 1, 14, len(keys)).setDataArray(tuple((f"Машина {k}", k) for k in keys))
    carsh.getCellRangeByPosition(15, 1, 16, len(sups)).setDataArray(tuple((x.title(), x) for x in sups))
    carsh.getCellByPosition(18, 0).setValue(len(keys))
    carsh.getCellByPosition(20, 0).setValue(len(sups))
    carsh.protect("wms")
    sh.protect("wms")
    s.set_sys("NEXT_CAR", float(len(crow) + 1))


@case
def m14_volume():
    c = "M14"
    p = new_wms("m14")
    s = ready(Session(p))
    try:
        n0 = 30000
        prefill(s, n0)
        s.close(save=True)
        t0 = time.time()
        s = ready(Session(p))
        t_open = time.time() - t0
        ta = []
        for i in range(5):
            t0 = time.time()
            r = car(s, "CarArrive", f"Грузовик Т{100 + i}ТТ", "Ромашка")
            ta.append((time.time() - t0, r[:14]))
        td = []
        for i in range(5):
            t0 = time.time()
            r = car(s, "CarDepart", n0 + 4 + i)
            td.append((time.time() - t0, r[:14]))
        t0 = time.time()
        dup = car(s, "CarArrive", f"Машина А{(n0 + 1) % 997:03d}ВС", "x")
        t_dup = time.time() - t0
        t0 = time.time()
        opens = car(s, "OpenVisits")
        t_open_list = time.time() - t0
        t0 = time.time()
        chk = car(s, "CarCheck")
        t_chk = time.time() - t0
        mx = max(x[0] for x in ta + td)
        R.add(c, f"{n0} визитов в истории: ПРИЕХАЛ/УЕХАЛ без просмотра истории — каждая операция < 1 с (макс. {mx:.2f} с), отказ по открытой "
                 "машине тоже быстрый",
              all(x[1].startswith("OK") for x in ta + td) and mx < 1.0 and dup.startswith("ERR") and t_dup < 1.0,
              f"приезд {[round(x[0], 3) for x in ta]}; выезд {[round(x[0], 3) for x in td]}; отказ {t_dup:.3f} с; открытие книги {t_open:.1f} с")
        R.add(c, "список машин на территории (одна формула движка) и самопроверка визитов на большом объёме",
              len(opens) == 3 and t_open_list < 1.0 and chk.startswith(f"визитов {n0 + 8}") and t_chk < 3,
              f"открытых {len(opens)} за {t_open_list:.3f} с; {chk[:90]}; {t_chk:.1f} с")
        s.issue_input(1, ei="5", qty="1", date=D0, who="егр")
        t0 = time.time()
        ri = s.post(1)
        t_iss = time.time() - t0
        R.add(c, "выдача на книге с 30 000 визитов не замедляется показателями «Приход авто» (< 1 с)", ri.startswith("OK") and t_iss < 1.0,
              f"{ri}; {t_iss:.2f} с")
    finally:
        s.close()


# ================================================================ S01–S04 the visual blocks of «Заказы» (M6 §18)

SEQ = ["1", "2", "3", "4", "5", "1", "2", "3", "1", "2", "3", "4"]


def order_rows(s, seq, r0=1):
    sh = s.doc.Sheets.getByName("Заказы")
    rows = [(a, f"Товар {i + 1}", "", f"СЧ-{i}", f"ART-{i + 1}", "", "", "10", "шт") for i, a in enumerate(seq)]
    sh.getCellRangeByPosition(0, r0, 8, r0 + len(seq) - 1).setDataArray(tuple(rows))
    for r in range(r0, r0 + len(seq)):
        sh.getCellByPosition(11, r).setString("Петрович")
        sh.getCellByPosition(15, r).setString("20.09.2026")


def receive(s, r, f="4"):
    s.order_input(r, F=f, C="УПД-1", O="24.09.2026", N=D0, U="A-1")
    return s.click_ord("BtnRcvPost", r)


def more(s, r, f="2", doc="УПД-2"):
    s.OU("TestRcvInput", f, "", doc, "25.09.2026", D0, "A-2", "")
    return s.click_ord("BtnRcvMore", r)


def blocks_model(s, last=None):
    """an independent model of the visual blocks: rows 1..last of «Заказы»; a block starts at a row whose A is 1 and which
    is not a delivery row (its AC mark); returns the rows (0-based) that start a block after the first one"""
    sh = s.doc.Sheets.getByName("Заказы")
    if last is None:
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        last = cur.getRangeAddress().EndRow
    d = sh.getCellRangeByPosition(0, 1, 28, last).getDataArray()
    starts, first = [], None
    for i, row in enumerate(d):
        a = row[0]
        a = str(int(a)) if isinstance(a, float) and a == int(a) else str(a).strip()
        if a == "1" and row[28] == "":
            if first is None:
                first = i + 1
            starts.append(i + 1)
    return [r for r in starts if r != 1]


def rendered_lines(path_ods, first_col_x=None):
    """the book rendered by LibreOffice itself: «Заказы» exported to PDF (fit to one page) — the drawn thick horizontal lines
    (the separators, y from the top) and the rows of text of column A (y of the baselines)"""
    import pdfscan
    o_name = unique("render")
    from wmslo import Office, props
    o = Office(o_name, PROFILES)
    try:
        doc = o.load(path_ods, macros=0, hidden=True)
        sh = doc.Sheets.getByName("Заказы")
        ps = doc.StyleFamilies.getByName("PageStyles").getByName(sh.PageStyle)
        ps.IsLandscape = True
        w, h = ps.Width, ps.Height
        if w < h:
            ps.Width, ps.Height = h, w
        ps.ScaleToPagesX = 1
        ps.ScaleToPagesY = 1
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        last = cur.getRangeAddress().EndRow
        rng = sh.getCellRangeByPosition(0, 0, 27, last)
        out = path_ods + ".pdf"
        fd = uno.Any("[]com.sun.star.beans.PropertyValue", props(Selection=rng))
        doc.storeToURL(uno.systemPathToFileUrl(out), props(FilterName="calc_pdf_Export", FilterData=fd))
        doc.close(True)
    finally:
        o.terminate()
    return pdfscan.scan(out)


def blocks_from_render(lines, rows_y):
    """the text rows between the drawn separators: [n rows of block 1, n rows of block 2, …]"""
    ys = sorted(set(round(y, 1) for y in lines))
    counts = [0] * (len(ys) + 1)
    for y in rows_y:
        k = sum(1 for ly in ys if y > ly)
        counts[k] += 1
    return ys, counts


@case
def s01_structure():
    c = "S01"
    p = new_wms("s01")
    s = ready(Session(p))
    try:
        sh = s.doc.Sheets.getByName("Заказы")
        cf = sh.getCellRangeByPosition(0, 1, 27, 1048575).ConditionalFormat
        entries = [(cf.getByIndex(k).getFormula1(), cf.getByIndex(k).getStyleName()) for k in range(cf.getCount())]
        st_ = s.doc.StyleFamilies.getByName("CellStyles").getByName("WMS_OrderStart")
        tb = st_.TopBorder
        db = s.doc.DatabaseRanges.getByName("WMS_ORDERS").getDataArea()
        R.add(c, "«Заказы»: условное оформление A2:AB — одно правило «A = 1 и строка не «Ещё поступление» (AC пуста), кроме первой строки» → "
                 "стиль с заметной верхней линией на всю ширину",
              entries == [('AND(ROW()>2;TRIM($A2)="1";$AC2="")', "WMS_OrderStart")] and tb.OuterLineWidth + tb.InnerLineWidth >= 70
              and tb.Color == 0x1F3864, f"{entries}; ширина {tb.OuterLineWidth}, цвет {tb.Color:06X}")
        R.add(c, "служебная колонка AC «Блок (служебная)»: скрыта, защищена, вне автофильтра A:AB",
              sh.getCellByPosition(28, 0).getString() == "Блок (служебная)" and not sh.getColumns().getByIndex(28).IsVisible
              and sh.getCellByPosition(28, 5).CellProtection.IsLocked and db.EndColumn == 27, f"автофильтр до колонки {db.EndColumn}")
        s.close(save=True)
        xml = zipfile.ZipFile(p).read("content.xml").decode("utf-8")
        t0 = xml.index('<table:table table:name="Заказы"')
        t1 = xml.index("</table:table>", t0)
        R.add(c, "никаких объединённых ячеек на листе «Заказы»", "number-columns-spanned" not in xml[t0:t1] and "number-rows-spanned" not in xml[t0:t1], "")
        s = ready(Session(p))
        order_rows(s, ["1", "2"])
        receive(s, 1)
        r, folder = snapshot_files(s)
        hdr = open(os.path.join(folder, "orders.csv"), encoding="utf-8").readline().strip().split(";")
        R.add(c, "снимок: orders.csv — ровно 28 колонок контракта, служебная колонка AC не экспортируется", len(hdr) == 28 and hdr[-1] == "possible_dup",
              f"{len(hdr)}: …{hdr[-2:]}")
    finally:
        s.close()


@case
def s02_sequence_three_blocks():
    c = "S02"
    p = new_wms("s02")
    s = ready(Session(p))
    try:
        order_rows(s, SEQ)
        R.add(c, "последовательность 1,2,3,4,5,1,2,3,1,2,3,4 — модель: 3 блока, 2 внутренние линии (перед строками 7 и 10 листа)",
              blocks_model(s) == [6, 9], f"{blocks_model(s)}")
        r3, r1 = receive(s, 3), receive(s, 1)
        m1 = more(s, 3)
        m2 = more(s, 1)
        sh = s.doc.Sheets.getByName("Заказы")
        rows = sh.getCellRangeByPosition(0, 1, 28, 14).getDataArray()
        a_col = [x[0] for x in rows]
        R.add(c, "«Ещё поступление» позиции 3 — новая строка сразу под позицией 3 (внутри блока), позиции 1 — под позицией 1; строка "
                 "поступления с A = 1 не начинает новый заказ (отметка AC)",
              r3.startswith("OK") and r1.startswith("OK") and m1.startswith("OK:3|строка 5|") and m2.startswith("OK:4|строка 3|")
              and a_col == ["1", "1", "2", "3", "3", "4", "5", "1", "2", "3", "1", "2", "3", "4"] and rows[1][28] == "OL2" and rows[4][28] == "OL1",
              f"{m1}; {m2}; A {a_col}; AC {[x[28] for x in rows]}")
        R.add(c, "модель после поступлений: 3 блока, линии перед 1-й позицией второго и третьего заказа (строки 9 и 12 листа)",
              blocks_model(s) == [8, 11], f"{blocks_model(s)}")
        roracle(c, s, "поступления внутри блоков")
        s.close(save=True)
        import pdfscan  # noqa: F401
        lines, rows_y = rendered_lines(p)
        ys, counts = blocks_from_render(lines, rows_y)
        R.add(c, "отрисовка LibreOffice (лист в PDF): ровно две разделительные линии через всю ширину A:AB; между ними строки 7 / 3 / 4 — "
                 "три визуальных блока заказов",
              len(ys) == 2 and counts == [7, 3, 4] and all(lines[y] >= 0.95 for y in lines), f"линии {ys}, доля ширины {lines}; строки по блокам {counts}")
        s = ready(Session(p))
    finally:
        s.close()


@case
def s03_blocks_survive():
    c = "S03"
    p = new_wms("s03")
    s = ready(Session(p))
    try:
        order_rows(s, SEQ)
        for r in (1, 3, 6, 9):
            receive(s, r)
        more(s, 3)
        s.OU("TestRcvInput", "5", "", "УПД-1", "24.09.2026", D0, "A-1", "")
        fx = s.click_ord("BtnRcvFix", 4)
        base = blocks_model(s)
        # a row inserted by the user inside the first order (the sheet allows it) and a new position typed there
        s.goto("Заказы", "$A$7")
        s.ui(".uno:InsertRows")
        s.order_input(6, A="6", B="Товар новый", H="3", I="шт", L="Петрович", P="20.09.2026")
        after_ins = blocks_model(s)
        R.add(c, "после «Исправить» и вставки пользователем строки с новой позицией внутри первого заказа — блоки те же (линии сдвинулись со строками)",
              fx.startswith("OK") and base == [7, 10] and after_ins == [8, 11], f"{fx}; {base} → {after_ins}")
        s.goto("Заказы", "$L$1")
        sh = s.doc.Sheets.getByName("Заказы")
        cf_before = sh.getCellRangeByPosition(0, 1, 27, 1048575).ConditionalFormat.getCount()
        s.close(save=True)
        s = ready(Session(p))
        sh = s.doc.Sheets.getByName("Заказы")
        cf_after = sh.getCellRangeByPosition(0, 1, 27, 1048575).ConditionalFormat.getCount()
        R.add(c, "сохранение и повторное открытие: правило оформления на месте, модель блоков та же", cf_before == cf_after == 1 and blocks_model(s) == [8, 11],
              f"{cf_before}/{cf_after}; {blocks_model(s)}")
        roracle(c, s, "блоки после исправлений и вставки")
        s.close(save=True)
        lines, rows_y = rendered_lines(p)
        ys, counts = blocks_from_render(lines, rows_y)
        R.add(c, "отрисовка после всех изменений: две линии, блоки 7 / 3 / 4 строки", len(ys) == 2 and counts == [7, 3, 4], f"{ys}; {counts}")
        s = ready(Session(p))
    finally:
        s.close()


@case
def s04_unsaved_crash_replays_inserted_rows():
    c = "S04"
    p = new_wms("s04")
    s = ready(Session(p))
    try:
        order_rows(s, SEQ)
        for r in (1, 3, 9):
            receive(s, r)
        s.doc.store()
        m1 = more(s, 3)
        m2 = more(s, 1)
        # a receipt on a row below the insertions (its row number was shifted by them)
        post = receive(s, 12)
        want = s.doc.Sheets.getByName("Заказы").getCellRangeByPosition(0, 1, 28, 16).getDataArray()
        s, unl = reopen(p, s)
        _, block = st(s)
        rec = s.B("ActionRecover")
        got = s.doc.Sheets.getByName("Заказы").getCellRangeByPosition(0, 1, 28, 16).getDataArray()
        # W of an open row (no EI) is the derived status the start refreshes («Ожидается»): not a result of the replay
        diff = [(i + 2, [(k, want[i][k], got[i][k]) for k in range(29) if want[i][k] != got[i][k]][:3]) for i in range(len(want))
                if any(want[i][k] != got[i][k] for k in range(29) if not (k == 22 and want[i][21] == ""))]
        R.add(c, "падение без сохранения после двух «Ещё поступление» (со вставкой строк) и прихода ниже них: «Восстановить» вставляет строки "
                 "заново и доводит все три операции — лист такой же, как до падения",
              m1.startswith("OK") and m2.startswith("OK") and post.startswith("OK") and block == "TAIL" and "восстановлено операций: 3" in rec
              and not diff, f"{rec[:80]}; различия в строках {diff[:6]}")
        R.add(c, "после восстановления блоки те же", blocks_model(s) == [8, 11], f"{blocks_model(s)}")
        roracle(c, s, "восстановление вставленных строк")
        s.B("TestSetFault", 3, 2)
        crash = more(s, 4)
        s, unl = reopen(p, s)
        rec2 = s.B("ActionRecover")
        d = s.doc.Sheets.getByName("Заказы").getCellRangeByPosition(0, 1, 28, 16).getDataArray()
        used = [x for x in d if any(v != "" for v in x)]
        R.add(c, "сбой «Ещё поступление» после журнала (книга сохранена посреди операции): откат оставляет пустую вставленную строку, "
                 "«Восстановить» пишет поступление в неё — без второй вставки (строк стало на одну больше)",
              crash == "CRASH-SIM" and "восстановлено операций: 1" in rec2 and len(used) == 15 and d[5][0] == "3" and d[5][28] == d[4][28] != ""
              and d[5][21] != "" and blocks_model(s) == [9, 12], f"{rec2[:70]}; строк {len(used)}; A {[x[0] for x in d]}")
        roracle(c, s, "сбой после журнала")
    finally:
        s.close()


# ================================================================ U01 the upgrade of a book of 0.6 (tools/upgrade.py)

OLD_COMMIT = "939929a"               # release 0.6.0 (M5)
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def old_book(path):
    """a test book built by the sources of release 0.6.0 (git archive of its commit): "" — built, otherwise why not"""
    src = os.path.join(OUT, "src_060")
    if not os.path.isdir(os.path.join(src, "tools")):
        why = git_sources(OLD_COMMIT, src)
        if why:
            return why
    r = subprocess.run([sys.executable, os.path.join(src, "tools", "build_ods.py"), path, "--test"], capture_output=True, text=True, timeout=600)
    if r.returncode != 0 or not os.path.exists(path):
        raise RuntimeError("книга 0.6 не собрана: " + (r.stderr or r.stdout)[-400:])
    return ""


@case
def u01_upgrade_060():
    c = "U01"
    d = os.path.join(OUT, "cases", "u01")
    import shutil
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(d)
    p = os.path.join(d, "WMS.ods")
    why = old_book(p)
    if why:
        R.add(c, "обновление книги 0.6 (tools/upgrade.py)", "SKIP", why)
        return
    s = ready(Session(p))
    try:
        v0 = s.B("StartupReport")
        order_rows(s, ["1", "2", "3", "1", "2"])
        res = [receive(s, 1), receive(s, 3), receive(s, 4)]
        m_old = more(s, 3)          # 0.6: the delivery row goes to the end of the sheet
        s.issue_input(1, ei="203", qty="1", date=D0, who="егр")
        res.append(s.post(1))
        R.add(c, "книга 0.6 (собрана из исходников выпуска 0.6.0): заказы, приходы, «Ещё поступление» (строка в конце листа), выдача",
              "ядро 0.6.0" in v0 and all(x.startswith("OK") for x in res) and m_old.startswith("OK:4|строка 7|"), f"{res}; {m_old}")
        s.close(save=True)
        s = None
        chk = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "upgrade.py"), p, "--check"], capture_output=True, text=True, timeout=600)
        up = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "upgrade.py"), p], capture_output=True, text=True, timeout=900)
        again = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "upgrade.py"), p, "--check"], capture_output=True, text=True, timeout=600)
        backups = [f for f in os.listdir(os.path.join(d, "WMS_Backups")) if f.endswith("_preupgrade.ods")]
        R.add(c, "tools/upgrade.py: --check без изменений; обновление с резервной копией *_preupgrade.ods; повторная проверка — «уже версии 0.7»",
              chk.returncode == 0 and up.returncode == 0 and "Книга обновлена" in up.stdout and len(backups) == 1 and again.returncode == 1
              and "уже версии 0.7" in again.stdout, f"{chk.returncode}/{up.returncode}/{again.returncode}; {up.stdout[-300:]}")
        s = ready(Session(p))
        rep = s.B("StartupReport")
        sc = s.B("ActionSelfCheck")
        R.add(c, "WMS 0.7 открывает обновлённую книгу там же: «работа разрешена», самопроверка без ошибок, схема WMS-SYS-4, журнал продолжается",
              st(s)[0] == "CLEAN" and f"ядро {wms_versions()['WMS_CORE_VERSION']}" in rep and sc.startswith("САМОПРОВЕРКА WMS: ошибок 0") and s.sysv("SCHEMA") == "WMS-SYS-4"
              and s.sysv("NEXT_CAR") == 1.0, f"{rep[:120]}; {sc[:60]}")
        roracle(c, s, "журнал 0.6 после обновления")
        sh = s.doc.Sheets.getByName("Заказы")
        ac = [sh.getCellByPosition(28, r).getString() for r in range(1, 8)]
        R.add(c, "строка «Ещё поступление» 0.6 (в конце листа) отмечена в AC и не начинает новый заказ; блоки по модели — перед строкой 5",
              ac == ["", "", "", "", "", "OL2", ""] and blocks_model(s) == [4], f"AC {ac}; {blocks_model(s)}")
        m_new = more(s, 3)
        a_col = [x[0] for x in sh.getCellRangeByPosition(0, 1, 0, 8).getDataArray()]
        R.add(c, "после обновления «Ещё поступление» ставит строку внутрь блока заказа (под позицией 3)",
              m_new.startswith("OK:6|строка 5|") and a_col[:5] == ["1", "2", "3", "3", "1"], f"{m_new}; {a_col}")
        r1 = arrive(s, "Газель А123ВС", "Петрович")
        r2 = depart_at(s, FIRST)
        R.add(c, "«Приход авто» в обновлённой книге: ПРИЕХАЛ и УЕХАЛ проводятся, seq продолжает журнал 0.6", r1.startswith("OK:7|1|") and r2.startswith("OK:8|"),
              f"{r1}; {r2}")
        coracle(c, s, "после обновления")
        roracle(c, s, "после обновления и новых операций")
    finally:
        if s is not None:
            s.close()


@case
def u02_upgrade_rollback():
    """a damaged book of 0.6 (NEXT_EI below the EIs issued — the checks of the upgrade do not look at it, WMS 0.7 does): the
    upgrade is done on a copy, WMS 0.7 does not accept it — the book is put back byte for byte, nothing of 0.7 is left"""
    c = "U02"
    import glob
    import hashlib
    import shutil
    d = os.path.join(OUT, "cases", "u02")
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(d)
    p = os.path.join(d, "WMS.ods")
    why = old_book(p)
    if why:
        R.add(c, "обновление повреждённой книги 0.6 — возврат", "SKIP", why)
        return
    s = ready(Session(p))
    try:
        order_rows(s, ["1", "2"])
        res = [receive(s, 1), receive(s, 2)]
        top = s.sysv("NEXT_EI")
        s.close(save=True)
        s = None
    finally:
        if s is not None:
            s.close()
    o = Office(unique("u02"), PROFILES)
    try:
        doc = o.load(p, macros=0)
        sy = doc.Sheets.getByName("_SYS")
        sy.unprotect("wms")
        sy.getCellByPosition(1, 5).setValue(top - 1)          # NEXT_EI (SK_NEXT_EI = 5): the last EI would be issued again
        sy.protect("wms")
        doc.store()
        doc.close(True)
    finally:
        o.terminate()
    # yesterday's daily copy only: opening the upgraded book, WMS 0.7 makes today's — the rollback must remove it
    bdir = os.path.join(d, "WMS_Backups")
    today, yday = datetime.date.today().isoformat(), (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
    for f in os.listdir(bdir):
        if f"_{today}_" in f:
            os.rename(os.path.join(bdir, f), os.path.join(bdir, f.replace(f"_{today}_", f"_{yday}_")))
    sha = lambda f: hashlib.sha256(open(f, "rb").read()).hexdigest()
    h0 = sha(p)
    copies0 = sorted(os.listdir(bdir))
    up = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "upgrade.py"), p], capture_output=True, text=True, timeout=900)
    copies1 = sorted(os.listdir(os.path.join(d, "WMS_Backups")))
    extra = [f for f in copies1 if f not in copies0 and not f.endswith("_preupgrade.ods")]
    reps = glob.glob(os.path.join(d, "UPGRADE_REPORT_*.md"))
    R.add(c, "повреждённая книга 0.6 (NEXT_EI ниже выданных ЕИ): WMS 0.7 не принимает обновлённую книгу — код 2, книга возвращена байт в "
             "байт; ежедневная копия, которую WMS 0.7 сделала с обновлённой книги, удалена; блокировки и временной копии не осталось; "
             "отчёт с причиной",
          all(x.startswith("OK") for x in res) and up.returncode == 2 and "возвращена из резервной копии" in up.stdout and sha(p) == h0
          and "удалены копии обновлённой книги" in up.stdout and "_daily.ods" in up.stdout
          and not extra and len([f for f in copies1 if f.endswith("_preupgrade.ods")]) == 1
          and not os.path.exists(os.path.join(d, "WMS_Journal", "wms.lock")) and not glob.glob(os.path.join(d, ".*.upgrade.ods"))
          and len(reps) == 1 and "ОШИБКА" in open(reps[0], encoding="utf-8").read(),
          f"{res}; код {up.returncode}; {up.stdout.strip()[-300:]}; лишние копии {extra}")
    s = ready(Session(p))
    try:
        v = s.B("StartupReport")
        R.add(c, "возвращённая книга — снова книга 0.6 (ядро 0.6.0, схема WMS-SYS-3)", "ядро 0.6.0" in v and s.sysv("SCHEMA") == "WMS-SYS-3", v[:120])
    finally:
        s.close()


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
    R.save(os.path.join(OUT, "results_m6.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_m6.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
