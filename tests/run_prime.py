"""PRIME 0.8 — stage 1: the protection window of an operation (a sheet protected with a foreign password), the new
corrections (POSITION_FIX, LEGACY_FIX, ORDER_UNCANCEL, MIGRATE_FIX, the cancellation of a transferred receipt), the rule
«остаток = сумма движений» and «Сверка переноса», the windows of «Исправить», the recovery of the new operations, and the
upgrade of a working book of 0.7.2 (tools/upgrade.py: check, build, promote, every stop).

    WMS_TEST_OUT=/tmp/wms_prime python3 tests/run_prime.py [case ...]

The book of 0.7.2 is made by the sources of release 0.7.2 themselves (tests/prime_fixture.py): its transfer tool moves the
main dataset of the transfer tests into its candidate, its macros post a working day. Every book is checked by the
independent oracle (tests/adjust_oracle.py: the whole journal replayed with the rules of the task).
Results: WMS_TEST_OUT/results_prime.json and WMS_TEST_OUT/TEST_REPORT_prime.md.
"""
import datetime
import glob
import hashlib
import json
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, ROOT, Results, Session, new_wms  # noqa: E402

R = Results()
CASES = []
T = datetime.date.today()
TODAY = T.strftime("%d.%m.%Y")
ENV = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
UPGRADE = os.path.join(ROOT, "tools", "upgrade.py")
FOREIGN = "чужой-пароль"
ORDER_COLS = [chr(65 + i) for i in range(26)] + ["AA", "AB", "AC"]


def case(fn):
    CASES.append(fn)
    return fn


def serial(k=0):
    return float((T + datetime.timedelta(days=k) - datetime.date(1899, 12, 30)).days)


def sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


def ready(s):
    s.U("TestUiAuto", 1)
    return s


def state(s):
    d = s.state()
    return d.get("STATE"), d.get("BLOCK")


def put(s, sheet, r, **cols):
    """cells of a row written through the API (numbers as numbers, the rest as text)"""
    sh = s.doc.Sheets.getByName(sheet)
    for k, v in cols.items():
        c = sh.getCellByPosition(ORDER_COLS.index(k), r)
        if isinstance(v, float):
            c.setValue(v)
        else:
            c.setString(v)


def order(s, r, a, name, qty="10", unit="шт", supplier="Ромашка", art="", **more):
    put(s, "Заказы", r, A=a, B=name, H=qty, I=unit, L=supplier, **({"E": art} if art else {}), **more)


def receive(s, r, f, doc="УПД-1", place="A-01", k=0):
    put(s, "Заказы", r, F=f, C=doc, O=serial(k), N=serial(k), U=place)
    return s.Rc("ReceiptPostRow", r)


def issue(s, r, ei, qty, who="егр"):
    put(s, "Выдачи", r, L=ei, E=qty, H=serial(), I=who)
    return s.I("IssuePostRow", r)


def ords(s, r):
    return list(s.doc.Sheets.getByName("Заказы").getCellRangeByPosition(0, r, 28, r).getDataArray()[0])


def leg(s, n):
    return list(s.doc.Sheets.getByName("_LEG").getCellRangeByPosition(0, n, 10, n).getDataArray()[0])


def oracle(c, s, name, initial=None, **kw):
    P, inf, _ = s.adjcheck(0, initial=initial, **kw) if initial is not None else s.adjcheck(0, **kw)
    R.add(c, "оракул (журнал воспроизведён независимо): " + name, not P, f"{inf.get('ops')} операций; {P[:4]}")
    return P


def njournal(s):
    return len(s.journal()[0])


def protect_foreign(s, name):
    sh = s.doc.Sheets.getByName(name)
    if sh.isProtected():
        sh.unprotect("wms")
    sh.protect(FOREIGN)


def unprotect_foreign(s, name):
    sh = s.doc.Sheets.getByName(name)
    sh.unprotect(FOREIGN)
    sh.protect("wms")


def reopen(p, s, kill=True):
    if kill:
        s.kill()
        for f in glob.glob(os.path.join(os.path.dirname(p), ".~lock.*#")):
            os.remove(f)
    else:
        s.close(save=True)
    s = ready(Session(p))
    if kill and state(s)[1] == "LOCKED":
        # the lock of the killed process: what a storekeeper does after a crash — «Снять блокировку WMS»
        s.B("ActionForceUnlock")
    return s


def raw(p, fn):
    """the book opened with macros disabled (as another program would), fn(doc), stored"""
    s = Session(p, macros=0)
    try:
        fn(s.doc)
        s.doc.store()
    finally:
        s.close()


# ================================================================ PR01 the protection window: a foreign password

@case
def pr01_foreign_password():
    c = "PR01"
    p = new_wms("pr01")
    s = ready(Session(p))
    try:
        r1 = issue(s, 1, "ЕИ-00000001", "1")
        seq0, j0 = s.sysv("LAST_SEQ"), njournal(s)
        protect_foreign(s, "Выдачи")
        r2 = issue(s, 2, "ЕИ-00000001", "1")
        st = state(s)
        a2 = s.doc.Sheets.getByName("Выдачи").getCellByPosition(0, 2).getString()
        R.add(c, "лист «Выдачи» защищён чужим паролем: выдача отказана до начала операции понятным сообщением; ничего не изменено (seq, "
                 "журнал, строка, маркер); WMS не заблокирована",
              r1.startswith("OK") and r2.startswith("ERR-RB:") and "«Выдачи» защищён не паролем WMS" in r2 and "ничего не изменено" in r2
              and "Сервис → Защитить лист" in r2 and st == ("CLEAN", "") and s.sysv("LAST_SEQ") == seq0 and njournal(s) == j0 and a2 == ""
              and s.sysv("TX_STATE") in ("COMMITTED", "NONE") and s.sysv("TX_BEFORE_IMAGE") == "", f"{r2[:200]}; {st}")
        pw = s.B("PreWorkCheck", module="WmsStatus")
        R.add(c, "«Проверка перед работой»: ошибка «пароли листов» с названием листа и подсказкой",
              any(ln.startswith("FAIL") and "пароли листов" in ln and "«Выдачи»" in ln for ln in pw.split("\n")), pw[:300])
        s = reopen(p, s, kill=False)
        rep = s.report()
        R.add(c, "запуск: «работа разрешена» и предупреждение о листе с чужим паролем; другие листы работают (приход по заказу проведён)",
              state(s) == ("CLEAN", "") and "«Выдачи» защищены не паролем WMS" in rep
              and (order(s, 1, "1", "Шайба") or receive(s, 1, "5")).startswith("OK"), rep[:300])
        unprotect_foreign(s, "Выдачи")
        r3 = issue(s, 2, "ЕИ-00000001", "1")
        R.add(c, "чужая защита снята — выдача проведена, лист «Выдачи» снова защищён паролем WMS",
              r3.startswith("OK") and s.doc.Sheets.getByName("Выдачи").isProtected() and s.B("ForeignProtection", "Выдачи") == "", r3)
        oracle(c, s, "после отказа и повторной выдачи")
    finally:
        s.close()


@case
def pr02_incident_333():
    """the incident of the warehouse (seq 333 of 0.7.2): an operation stopped in the middle, the book saved, a sheet of the
    operation protected with a foreign password before the next start — 0.7.2 said MARKER_CORRUPT without the cause"""
    c = "PR02"
    p = new_wms("pr02")
    s = ready(Session(p))
    seq0 = s.sysv("LAST_SEQ")
    s.B("TestSetFault", 2, 2)
    r = issue(s, 1, "ЕИ-00000001", "1")
    R.add(c, "сбой посреди выдачи: книга сохранена с незавершённой операцией (маркер STARTED)", r == "CRASH-SIM", r)
    s.kill()
    for f in glob.glob(os.path.join(os.path.dirname(p), ".~lock.*#")):
        os.remove(f)
    raw(p, lambda d: (d.Sheets.getByName("Выдачи").isProtected() and d.Sheets.getByName("Выдачи").unprotect("wms"),
                      d.Sheets.getByName("Выдачи").protect(FOREIGN)))
    s = ready(Session(p))
    try:
        unl = s.B("ActionForceUnlock") if state(s)[1] == "LOCKED" else ""
        st = state(s)
        rep = s.report() + " " + unl
        hint = s.B("BlockHint", "PROTECTION_FOREIGN", module="WmsUi")
        R.add(c, "запуск: ЗАБЛОКИРОВАНО с настоящей причиной — лист «Выдачи» защищён не паролем WMS, незавершённую операцию нельзя откатить; "
                 "подсказка «снимите эту защиту» (в 0.7.2 было MARKER_CORRUPT без причины)",
              st == ("BLOCKED", "PROTECTION_FOREIGN") and "«Выдачи»" in rep and "нельзя откатить" in rep and "Снимите эту защиту" in hint
              and s.sysv("TX_STATE") == "STARTED", f"{st}; {rep[:250]}")
    finally:
        s.close()
    raw(p, lambda d: (d.Sheets.getByName("Выдачи").unprotect(FOREIGN)))
    s = ready(Session(p))
    try:
        rep = s.report()
        R.add(c, "чужая защита снята — запуск откатил операцию по снимку: «работа разрешена», seq тот же, строка пуста, лист снова защищён "
                 "паролем WMS (после сбоя WMS сама возвращает свою защиту)",
              state(s) == ("CLEAN", "") and "отменена по снимку" in rep and s.sysv("LAST_SEQ") == seq0
              and s.doc.Sheets.getByName("Выдачи").getCellByPosition(0, 1).getString() == ""
              and s.doc.Sheets.getByName("Выдачи").isProtected() and s.B("ForeignProtection", "Выдачи") == "", rep[:250])
        r = issue(s, 1, "ЕИ-00000001", "1")
        R.add(c, "после этого выдача проводится как обычно", r.startswith("OK"), r)
        oracle(c, s, "после сбоя, чужой защиты и отката")
    finally:
        s.close()


@case
def pr03_window_after_crash():
    """every point of a crash inside an operation: after the next start the sheets of the window are protected again"""
    c = "PR03"
    res = []
    for point in (1, 2, 3):
        p = new_wms(f"pr03_{point}")
        s = ready(Session(p))
        s.B("TestSetFault", point, 2)
        r = issue(s, 1, "ЕИ-00000002", "1")
        s = reopen(p, s)
        if state(s)[1] == "TAIL":
            s.B("ActionRecover")
        prot = all(s.doc.Sheets.getByName(n).isProtected() for n in ("Выдачи", "Наличие", "_SYS"))
        res.append((point, r, state(s), prot, s.sysv("LAST_SEQ")))
        P, _, _ = s.adjcheck(0)
        res.append(not P)
        s.close()
    R.add(c, "сбой в точках 1, 2 (до журнала: откат) и 3 (после журнала: восстановление): листы снова защищены, «работа разрешена», "
             "оракул без расхождений", all(x[2] == ("CLEAN", "") and x[3] for x in res if isinstance(x, tuple)) and all(x for x in res if isinstance(x, bool)),
          str(res))


# ================================================================ PR04 «Исправить позицию» — POSITION_FIX

@case
def pr04_position_fix():
    c = "PR04"
    p = new_wms("pr04")
    s = ready(Session(p))
    try:
        order(s, 1, "4501", "Болт М8", qty="30", art="DIN933", P=serial(-10), S="Крепёж")
        r1 = receive(s, 1, "10")
        e1 = ords(s, 1)[21]
        r2 = s.Rc("ReceiptAddRow", 1, "5", "", "УПД-2", TODAY, TODAY, "A-02", "")
        e2 = ords(s, 2)[21]
        fx = s.Rc("PositionFixRow", 2, "4502", "Болт М8х40 оцинк.", "СЧ-77", "DIN933-40", "25", "шт", "Лютик", (T - datetime.timedelta(days=12)).strftime("%d.%m.%Y"),
                  "Метизы")
        a, b = ords(s, 1), ords(s, 2)
        n1, n2 = int(e1[3:]), int(e2[3:])
        cards = [s.card(n1), s.card(n2)]
        R.add(c, "позиция с поступлением: «Исправить позицию» из строки поступления — № заказа, наименование, счёт, артикул, заказано, поставщик, "
                 "дата заказа, категория в обеих строках позиции; карточки обоих ЕИ в «Наличие»; «Источник» — новый № заказа",
              r1.startswith("OK") and r2.startswith("OK") and fx.startswith("OK")
              and [a[i] for i in (0, 1, 3, 4, 11, 18)] == ["4502", "Болт М8х40 оцинк.", "СЧ-77", "DIN933-40", "Лютик", "Метизы"]
              and [b[i] for i in (0, 1, 3, 4, 11, 18)] == ["4502", "Болт М8х40 оцинк.", "СЧ-77", "DIN933-40", "Лютик", "Метизы"]
              and a[7] == 25.0 and b[7] == "" and a[15] == serial(-12) and b[15] == serial(-12)
              and all(cd[1] == "Болт М8х40 оцинк." and cd[2] == "DIN933-40" and cd[6] == "Метизы" and cd[8] == "Заказ 4502" for cd in cards),
              f"{r1}; {r2}; {fx}; {a[:19]}; {cards}")
        R.add(c, "заказано 25 из полученных 15: статус «Частично получено», в «Контроле» — «осталось 10»",
              a[22] == "Частично получено" and "осталось 10" in a[24], f"{a[22]}; {a[24]}")
        sk = s.Rc("PositionFixRow", 1, "4502", "Болт М8х40 оцинк.", "СЧ-77", "DIN933-40", "25", "шт", "Лютик",
                  (T - datetime.timedelta(days=12)).strftime("%d.%m.%Y"), "Метизы")
        R.add(c, "те же значения ещё раз — «ничего не изменилось», операции нет", sk.startswith("SKIP:"), sk)
        iss = issue(s, 1, e1, "2")
        ru = s.Rc("PositionFixRow", 1, "4502", "Болт М8х40 оцинк.", "СЧ-77", "DIN933-40", "25", "кг", "Лютик", "", "Метизы")
        R.add(c, "единица не меняется, пока по ЕИ позиции есть выдача: отказ с объяснением и подсказкой, что остальное исправить можно",
              iss.startswith("OK") and ru.startswith("ERR:") and "единицу измерения нельзя изменить" in ru and "Остальные поля" in ru, ru)
        bad = [s.Rc("PositionFixRow", 1, "", "x", "", "", "25", "шт", "Лютик", "", ""), s.Rc("PositionFixRow", 1, "4502", "", "", "", "25", "шт", "Лютик", "", ""),
               s.Rc("PositionFixRow", 1, "4502", "x", "", "", "0", "шт", "Лютик", "", ""), s.Rc("PositionFixRow", 1, "4502", "x", "", "", "25", "", "Лютик", "", ""),
               s.Rc("PositionFixRow", 1, "4502", "x", "", "", "25", "шт", "Лютик", "32.13.2026", ""),
               s.Rc("PositionFixRow", 1, "4502", "x", "", "", "25", "шт", "Офис", "", "")]
        R.add(c, "пустой № заказа, пустое наименование, заказано 0, пустая единица, неверная дата, поставщик — специальный приход: каждая "
                 "ошибка — отказ с причиной, ничего не изменено", all(x.startswith("ERR:") for x in bad) and ords(s, 1)[1] == "Болт М8х40 оцинк.",
              " | ".join(x[:70] for x in bad))
        order(s, 3, "4503", "Гайка", qty="5")
        R.add(c, "строка без прихода — «Исправить позицию» не нужно: исправляется прямо в строке",
              s.Rc("PositionFixRow", 3, "4503", "Гайка М8", "", "", "5", "шт", "Ромашка", "", "").startswith("ERR:строка не проведена"), "")
        m3 = s.Rc("ReceiptAddRow", 1, "3", "", "УПД-3", TODAY, TODAY, "A-03", "")
        r3 = ords(s, 3)
        R.add(c, "после исправления «Ещё поступление» копирует уже исправленную позицию", m3.startswith("OK") and r3[1] == "Болт М8х40 оцинк."
              and r3[11] == "Лютик" and r3[0] == "4502", f"{m3}; {r3[:12]}")
        oracle(c, s, "исправления позиции")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc.splitlines()[0])
    finally:
        s.close()


# ================================================================ PR05 «Коррекция переноса» — LEGACY_FIX, the cancellation of a transferred receipt

def legacy_receipt(s, r, ei, f, bal, a="9001", name="Кабель", x=None):
    order(s, r, a, name, qty=f)
    put(s, "Заказы", r, F=f, C="УПД-L", O=serial(-60), N=serial(-60), U="L-01")
    return s.Rc("ReceiptLegacyRow", r, ei, bal, "", f"staging|test|{r}", "Получено", "", x if x is not None else bal, "", "X", "OLD")


@case
def pr05_legacy_fix():
    c = "PR05"
    p = new_wms("pr05")
    s = ready(Session(p))
    try:
        r0 = legacy_receipt(s, 1, "ЕИ-00000700", "10", "4")
        lg0 = leg(s, 700)
        R.add(c, "перенесённый приход (F 10, остаток при переносе 4): реестр переноса — приход 10, остаток 4, израсходовано до переноса 6, "
                 "старое X, откуда остаток, строка старой таблицы, seq",
              r0.startswith("OK") and lg0[:8] == ["ЕИ-00000700", "RECEIPT", 10.0, 4.0, 6.0, "4", "X", "staging|test|1"] and lg0[8] == float(r0[3:]),
              f"{r0}; {lg0}")
        iss = issue(s, 1, "ЕИ-00000700", "1")
        rf = s.Rc("ReceiptFixRow", 1, "6", "", "УПД-L", (T - datetime.timedelta(days=60)).strftime("%d.%m.%Y"),
                  (T - datetime.timedelta(days=60)).strftime("%d.%m.%Y"), "L-01", "")
        R.add(c, "«Исправить» прихода ниже израсходованного: отказ называет «израсходовано до переноса 6» и «после переноса выдано 1» и "
                 "ведёт к «Коррекции переноса»", iss.startswith("OK") and rf.startswith("ERR:") and "израсходовано до переноса 6" in rf
              and "после переноса выдано или списано 1" in rf and "Коррекция переноса" in rf, rf)
        bad = [s.Rc("LegacyFixRow", 1, "10", "3", "", TODAY, ""), s.Rc("LegacyFixRow", 1, "10", "3", "А-1", "31.02.2026", ""),
               s.Rc("LegacyFixRow", 1, "10", "11", "А-1", TODAY, ""), s.Rc("LegacyFixRow", 1, "10", "0,5", "А-1", TODAY, ""),
               s.Rc("LegacyFixRow", 1, "10", "", "А-1", TODAY, ""), s.Rc("LegacyFixRow", 1, "0", "0", "А-1", TODAY, "")]
        R.add(c, "без № акта, неверная дата акта, остаток при переносе больше прихода, остаток меньше уже выданного после переноса, пустой "
                 "остаток, приход 0 — отказы с причиной, ничего не изменено",
              all(x.startswith("ERR:") for x in bad) and "№ акта" in bad[0] and "больше полученного" in bad[2] and "уже выдано или списано 1" in bad[3]
              and s.stock(700) == 3.0, " | ".join(x[:80] for x in bad))
        sk = s.Rc("LegacyFixRow", 1, "10", "4", "А-1", TODAY, "")
        fx = s.Rc("LegacyFixRow", 1, "12", "7", "А-17", TODAY, "пересчёт по полке")
        lg1 = leg(s, 700)
        R.add(c, "без изменений — «ничего не изменилось»; коррекция F 10 → 12, остаток при переносе 4 → 7: остаток 3 → 6 (выдача после "
                 "переноса сохранена), приход 12 в «Заказах» и _RCV, позиция получено 12, реестр: 12 / 7 / 5, коррекций 1, акт",
              sk.startswith("SKIP:") and fx.startswith("OK") and s.stock(700) == 6.0 and ords(s, 1)[5] == 12.0 and s.rcv(700)[6] == 12.0
              and s.opos(1)[5] == 12.0 and lg1[2:5] == [12.0, 7.0, 5.0] and lg1[9] == 1.0 and "акт № А-17 от" in lg1[10] and "пересчёт по полке" in lg1[10],
              f"{sk}; {fx}; {s.stock(700)}; {lg1}")
        dc = s.Rc("ReceiptDeleteCheck", 1)
        st_ = s.I("IssueDeleteRow", 1)
        d2 = s.Rc("ReceiptDeleteRow", 1)
        R.add(c, "удаление перенесённого прихода: при выдаче после переноса — отказ «связанные выдачи»; после сторно выдачи — удалён, остаток 0 "
                 "(израсходованное до переноса уходит вместе с приходом)",
              dc.startswith("ERR:") and "выдач" in dc and st_.startswith("OK") and d2.startswith("OK") and s.stock(700) == 0.0
              and s.rcv(700)[4] == "STORNO", f"{dc[:120]}; {st_}; {d2}")
        r2 = legacy_receipt(s, 3, "ЕИ-00000701", "8", "8", a="9002", name="Провод")
        R.add(c, "«Коррекция переноса» обычного прихода (не из переноса) — отказ с подсказкой, что исправлять обычным «Исправить» и «Инвентаризацией»",
              (order(s, 5, "9003", "Шнур") or receive(s, 5, "3")).startswith("OK") and s.Rc("LegacyFixRow", 5, "3", "3", "А", TODAY, "").startswith("ERR:")
              and r2.startswith("OK"), "")
        oracle(c, s, "коррекции переноса и сторно перенесённого прихода")
        bc = s.B("BalanceCheck", module="WmsLegacy")
        R.add(c, "«остаток = сумма движений» — все ЕИ сходятся (перенесённые с учётом израсходованного до переноса)", bc == "", bc)
    finally:
        s.close()


# ================================================================ PR06 «Снять отмену» — ORDER_UNCANCEL

@case
def pr06_uncancel():
    c = "PR06"
    p = new_wms("pr06")
    s = ready(Session(p))
    try:
        order(s, 1, "5001", "Лампа", qty="5")
        oc = s.Rc("OrderCancelRow", 1)
        un = s.Rc("OrderUncancelRow", 1)
        a = ords(s, 1)
        locks = s.ord_locks(1)
        pr = receive(s, 1, "5")
        R.add(c, "заказ отменён без прихода → «Снять отмену»: строка снова заказ («Ожидается», ввод открыт), приход проводится",
              oc.startswith("OK") and un.startswith("OK") and a[22] == "Ожидается" and "Отмена снята" in a[24] and locks == "0000000000000000000001111001"
              and pr.startswith("OK") and ords(s, 1)[22] == "Получено", f"{oc}; {un}; {a[22]}; {locks}; {pr}")
        order(s, 2, "5002", "Перчатки", qty="100")
        r1 = receive(s, 2, "40")
        cr = s.Rc("OrderCancelRestRow", 2)
        st1 = ords(s, 2)[22]
        un2 = s.Rc("OrderUncancelRow", 2)
        st2 = ords(s, 2)[22]
        more = s.Rc("ReceiptAddRow", 2, "60", "", "УПД-9", TODAY, TODAY, "A-9", "")
        R.add(c, "остаток отменён → «Снять отмену»: «Частично получено», недостающее снова ждём; «Ещё поступление» на 60 — «Получено»",
              r1.startswith("OK") and cr.startswith("OK") and st1 == "Частично получено / остаток отменён" and un2.startswith("OK")
              and st2 == "Частично получено" and more.startswith("OK") and ords(s, 2)[22] == "Получено", f"{st1}; {un2}; {st2}; {more}")
        order(s, 4, "5003", "Краска", qty="10")
        r4 = receive(s, 4, "10")
        d4 = s.Rc("ReceiptDeleteRow", 4)
        c4 = s.Rc("OrderCancelRow", 4)
        u4 = s.Rc("OrderUncancelRow", 4)
        m4 = s.Rc("ReceiptAddRow", 4, "10", "", "УПД-10", TODAY, TODAY, "A-10", "")
        R.add(c, "приход удалён и заказ отменён → «Снять отмену»: «Ожидается», «Ещё поступление» снова принимается",
              all(x.startswith("OK") for x in (r4, d4, c4, u4, m4)) and ords(s, 4)[22] == "Получено", f"{r4}; {d4}; {c4}; {u4}; {m4}")
        sk = [s.Rc("OrderUncancelRow", 1), s.Rc("OrderUncancelRow", 2)]
        R.add(c, "не отменённая позиция — «снимать нечего»", all(x.startswith("SKIP:") for x in sk), str(sk))
        oracle(c, s, "снятие отмены")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc.splitlines()[0])
    finally:
        s.close()


# ================================================================ PR07 «Исправить ЕИ переноса» — MIGRATE_FIX

def migrate(s, ei, name, qty, stype, art="", unit="шт", place="M-01", cat=""):
    return s.B("MigrateEI", ei, name, art, unit, qty, place, cat, stype, "", f"table|sha|{ei}", module="WmsMigrate")


@case
def pr07_migrate_fix():
    c = "PR07"
    p = new_wms("pr07")
    s = ready(Session(p))
    try:
        m1 = migrate(s, "ЕИ-00000800", "Стеллаж", "3", "Иной приход")
        m2 = migrate(s, "ЕИ-00000801", "Насос", "2", "Детали", art="PN-1")
        lg = leg(s, 800)
        R.add(c, "перенос ЕИ без заказа: «Требует разбора»; в реестре переноса — MIGRATE, остаток 3",
              m1.startswith("OK") and m2.startswith("OK") and s.card(800)[7] == "Требует разбора" and lg[1] == "MIGRATE" and lg[3] == 3.0, f"{m1}; {lg}")
        f1 = s.B("MigrateFixEI", "ЕИ-00000800", "Стеллаж металлический", "", "шт", "Мебель", "Офис", "", "", "", "", module="WmsLegacy")
        R.add(c, "«Разобрать»: тип «Офис», наименование и категория — ЕИ «Активен», остаток тот же, акт не нужен (остаток не менялся)",
              f1.startswith("OK") and s.card(800)[1] == "Стеллаж металлический" and s.card(800)[7] == "Активен" and s.card(800)[9] == "Офис"
              and s.stock(800) == 3.0, f"{f1}; {s.card(800)}")
        f2 = s.B("MigrateFixEI", "ЕИ-00000800", "Стеллаж металлический", "", "шт", "Мебель", "Офис", "2", "", "", "", module="WmsLegacy")
        f3 = s.B("MigrateFixEI", "ЕИ-00000800", "Стеллаж металлический", "", "шт", "Мебель", "Офис", "2", "А-5", TODAY, "пересчёт", module="WmsLegacy")
        R.add(c, "остаток при переносе 3 → 2: без акта — отказ; с актом — остаток 2, реестр 2, коррекций 2",
              f2.startswith("ERR:") and "акта" in f2 and f3.startswith("OK") and s.stock(800) == 2.0 and leg(s, 800)[3] == 2.0 and leg(s, 800)[9] == 2.0,
              f"{f2}; {f3}")
        x = [s.B("MigrateFixEI", "ЕИ-00000800", "Стеллаж", "PN-1", "шт", "", "Детали", "", "", "", "", module="WmsLegacy"),
             s.B("MigrateFixEI", "ЕИ-00000801", "Насос", "PN-1", "шт", "", "Офис", "", "", "", "", module="WmsLegacy"),
             s.B("MigrateFixEI", "ЕИ-00000800", "Стеллаж", "", "шт", "", "Склад", "", "", "", "", module="WmsLegacy"),
             s.B("MigrateFixEI", "ЕИ-00000001", "x", "", "шт", "", "Офис", "", "", "", "", module="WmsLegacy")]
        R.add(c, "артикул детали уже у другого ЕИ, деталь перестаёт быть деталью, неизвестный тип, ЕИ не из переноса — отказы с причиной",
              all(v.startswith("ERR:") for v in x) and "один артикул детали" in x[0] and "«Детали»" in x[1] and "неизвестен" in x[2] and "не перенесён" in x[3],
              " | ".join(v[:70] for v in x))
        iss = issue(s, 1, "ЕИ-00000800", "1")
        fu = s.B("MigrateFixEI", "ЕИ-00000800", "Стеллаж металлический", "", "компл", "Мебель", "Офис", "", "", "", "", module="WmsLegacy")
        R.add(c, "единица не меняется после выдачи", iss.startswith("OK") and fu.startswith("ERR:") and "единицу измерения нельзя" in fu, fu)
        info = s.B("MigrateFixInfo", "ЕИ-00000800", module="WmsLegacy")
        R.add(c, "окно получает карточку ЕИ переноса (наименование, тип, остаток при переносе, текущий остаток, место)",
              info.split("|")[0] == "Стеллаж металлический" and info.split("|")[4] == "Офис" and info.split("|")[5] == "2" and info.split("|")[6] == "1", info)
        oracle(c, s, "исправления ЕИ переноса")
    finally:
        s.close()


# ================================================================ PR08 «остаток = сумма движений» and «Сверка переноса»

@case
def pr08_balance_report():
    c = "PR08"
    p = new_wms("pr08")
    s = ready(Session(p))
    try:
        legacy_receipt(s, 1, "ЕИ-00000900", "10", "4", x="5")
        order(s, 2, "9101", "Хомут")
        receive(s, 2, "3")
        m = migrate(s, "ЕИ-00000990", "Ящик", "2", "Офис")
        bc0 = s.B("BalanceCheck", module="WmsLegacy")
        st = s.doc.Sheets.getByName("Наличие")
        st.unprotect("wms")
        st.getCellByPosition(4, 900).setValue(7.0)
        st.protect("wms")
        bc1 = s.B("BalanceCheck", module="WmsLegacy")
        mod0 = s.doc.isModified()
        s.doc.setModified(False)
        rep = s.B("LegacyReport", module="WmsLegacy")
        sh = s.doc.Sheets.getByName("Сверка переноса")
        rows = sh.getCellRangeByPosition(0, 4, 16, 8).getDataArray()
        R.add(c, "всё сходится — пусто; остаток ЕИ-00000900 изменён в обход WMS — названы ЕИ, остаток по движениям 4 и в «Наличие» 7",
              bc0 == "" and "ЕИ-00000900" in bc1 and "по движениям 4" in bc1 and "«Наличие» 7" in bc1, f"{bc0!r}; {bc1}")
        R.add(c, "«Сверка переноса»: первая строка — расхождение (разница 3, «пересчитайте ЕИ на месте»), затем перенос со старым X ≠ остатку при "
                 "переносе («проверьте»), затем в порядке; ЕИ, пришедшие обычным приходом, без расхождений не показаны; лист защищён, книга не "
                 "помечена изменённой",
              rep.startswith("OK:перенесённых ЕИ 2, расхождений с движениями 1") and rows[0][0] == "ЕИ-00000900" and rows[0][12] == 3.0
              and "пересчитайте" in rows[0][16] and "X = «5»" in rows[0][16] and rows[1][0] == "ЕИ-00000990" and rows[1][16] == "в порядке"
              and rows[2][0] == "" and sh.isProtected() and not s.doc.isModified() and mod0, f"{m}; {rep}; {rows[:3]}")
        hdr = sh.getCellRangeByPosition(0, 3, 16, 3).getDataArray()[0]
        R.add(c, "шапка «Сверки переноса» — 17 понятных колонок (ЕИ … Что проверить)", hdr[0] == "ЕИ" and hdr[16] == "Что проверить" and "Израсходовано до переноса" in hdr,
              str(hdr))
    finally:
        s.close()


# ================================================================ PR09 the windows of «Исправить» (test seams)

@case
def pr09_windows():
    c = "PR09"
    p = new_wms("pr09")
    s = ready(Session(p))
    try:
        legacy_receipt(s, 1, "ЕИ-00000950", "10", "6")
        order(s, 2, "9201", "Труба", qty="20")
        receive(s, 2, "20")
        s.goto("Заказы", "$B$2")
        menu1 = s.B("FixChoices", 1, "RECEIVED", module="WmsPrimeUi")
        menu2 = s.B("FixChoices", 2, "RECEIVED", module="WmsPrimeUi")
        R.add(c, "«Исправить» на перенесённом приходе предлагает: этот приход, позицию, коррекцию переноса; на обычном — без коррекции переноса",
              menu1.count("\n") == 3 and "LEG|" in menu1 and "POS|" in menu1 and "RCV|" in menu1 and "LEG|" not in menu2, f"{menu1!r}; {menu2!r}")
        s.B("TestPrimeChoice", 2, module="WmsPrimeUi")
        s.B("TestPrimeForm", "\t".join(["10", "5", "А-3", TODAY, "пересчёт"]), module="WmsPrimeUi")
        s.goto("Заказы", "$B$2")
        s.OU("BtnRcvFix")
        last = s.U("TestUiLastMessage")
        shown = s.B("TestPrimeShown", module="WmsPrimeUi")
        R.add(c, "окно «Коррекция переноса»: показаны числа переноса и текущий остаток, проведено — остаток 6 → 5",
              last.startswith("OK") and s.stock(950) == 5.0 and "по переносу получено 10" in shown and "израсходовано до переноса 4" in shown, f"{last}; {shown[:200]}")
        s.B("TestPrimeChoice", 1, module="WmsPrimeUi")
        s.B("TestPrimeForm", "\t".join(["<как предложено>", "Труба ПНД 32"]), module="WmsPrimeUi")
        s.goto("Заказы", "$B$3")
        s.OU("BtnRcvFix")
        last2 = s.U("TestUiLastMessage")
        R.add(c, "окно «Исправить позицию»: предложены значения позиции, изменено только наименование", last2.startswith("OK") and ords(s, 2)[1] == "Труба ПНД 32",
              last2)
        s.B("TestPrimeChoice", -1, module="WmsPrimeUi")
        s.goto("Заказы", "$B$3")
        s.OU("BtnRcvFix")
        R.add(c, "окно выбора закрыто без выбора — ничего не исправлено", s.U("TestUiLastMessage").startswith("SKIP"), s.U("TestUiLastMessage"))
        migrate(s, "ЕИ-00000960", "Шкаф", "1", "Иной приход")
        s.goto("Наличие", "$B$961")
        s.B("TestPrimeForm", "\t".join(["Шкаф офисный", "", "шт", "Мебель", "Офис"]), module="WmsPrimeUi")
        s.B("BtnMigrateFix", module="WmsPrimeUi")
        R.add(c, "«Наличие» → «Исправить ЕИ переноса» для ЕИ под курсором: разобран как «Офис»", s.U("TestUiLastMessage").startswith("OK")
              and s.card(960)[7] == "Активен", s.U("TestUiLastMessage"))
        s.goto("Наличие", "$B$2")
        s.B("BtnMigrateFix", module="WmsPrimeUi")
        msg = s.U("TestUiLastMessage")
        R.add(c, "ЕИ не из переноса — понятное объяснение, где его исправлять", msg.startswith("ERR:") and "не из переноса" in msg, msg)
        s.B("BtnLegacyReport", module="WmsPrimeUi")
        R.add(c, "«Сверка переноса» — лист построен и открыт", s.U("TestUiLastMessage").startswith("OK") and s.active_sheet() == "Сверка переноса",
              s.U("TestUiLastMessage"))
        oracle(c, s, "исправления через окна")
    finally:
        s.close()


# ================================================================ PR10 the recovery of the new operations

@case
def pr10_recovery():
    c = "PR10"
    out = []
    for op in ("LEGACY_FIX", "POSITION_FIX", "ORDER_UNCANCEL", "MIGRATE_FIX"):
        for point in (2, 3):
            p = new_wms(f"pr10_{op}_{point}")
            s = ready(Session(p))
            legacy_receipt(s, 1, "ЕИ-00000980", "10", "4")
            order(s, 2, "9301", "Кран", qty="10")
            receive(s, 2, "4")
            s.Rc("OrderCancelRestRow", 2)
            migrate(s, "ЕИ-00000985", "Тумба", "2", "Иной приход")
            s.B("TestSetFault", point, 2)
            if op == "LEGACY_FIX":
                r = s.Rc("LegacyFixRow", 1, "10", "5", "А-1", TODAY, "")
            elif op == "POSITION_FIX":
                r = s.Rc("PositionFixRow", 2, "9301", "Кран шаровый", "", "", "10", "шт", "Ромашка", "", "")
            elif op == "ORDER_UNCANCEL":
                r = s.Rc("OrderUncancelRow", 2)
            else:
                r = s.B("MigrateFixEI", "ЕИ-00000985", "Тумба", "", "шт", "", "Офис", "", "", "", "", module="WmsLegacy")
            s = reopen(p, s)
            st1 = state(s)
            if st1[1] == "TAIL":
                s.B("ActionRecover")
            P, inf, _ = s.adjcheck(0)
            done = {"LEGACY_FIX": s.stock(980) == 5.0, "POSITION_FIX": ords(s, 2)[1] == "Кран шаровый",
                    "ORDER_UNCANCEL": ords(s, 2)[22] == "Частично получено", "MIGRATE_FIX": s.card(985)[9] == "Офис"}[op]
            out.append((op, point, r, st1, state(s), done == (point == 3), not P))
            s.close()
    R.add(c, "сбой каждой новой операции до журнала (откат: операции нет) и после журнала («Восстановить» доводит её из журнала): "
             "«работа разрешена», оракул без расхождений", all(x[2] == "CRASH-SIM" and x[4] == ("CLEAN", "") and x[5] and x[6] for x in out), str(out))


# ================================================================ PR11 … the upgrade of a working book of 0.7.2

def upgrade(*args, fault=None):
    env = dict(ENV)
    if fault:
        env["WMS_UPGRADE_TEST_FAULT"] = fault
    r = subprocess.run([sys.executable, UPGRADE] + list(args), capture_output=True, text=True, timeout=3600, env=env)
    return r.returncode, (r.stdout + r.stderr).strip()


def warehouse():
    import prime_fixture
    return prime_fixture.warehouse_072(os.path.join(OUT, "w072"))


def old_copy(name):
    """a copy of the warehouse of 0.7.2 in a folder of its own, registered there by its own macros (its working path)"""
    book, facts = warehouse()
    d = os.path.join(OUT, "cases", name)
    shutil.rmtree(d, ignore_errors=True)
    shutil.copytree(os.path.dirname(book), os.path.join(d, "OLD"), ignore=shutil.ignore_patterns("WMS_TOOLBOX", "backup", "WMS_Backups"))
    p = os.path.join(d, "OLD", os.path.basename(book))
    s = Session(p)
    try:
        if "REG=other" in s.B("StateDump"):
            s.B("ActionRegisterHere", False)
        st = s.B("StateDump")
        if "STATE=CLEAN" not in st:
            raise RuntimeError("копия книги 0.7.2 не открылась: " + st[:200])
    finally:
        s.close(save=True)
    return d, p


@case
def pr11_upgrade_072():
    c = "PR11"
    d, p = old_copy("pr11")
    h0 = sha(p)
    j0 = {f: sha(f) for f in glob.glob(os.path.join(d, "OLD", "WMS_Journal", "*"))}
    rc0, o0 = upgrade("check", p)
    R.add(c, "проверка книги 0.7.2: схема WMS-SYS-4, ядро 0.7.2, операций и журнал совпадают — «можно обновить»; книга не изменена",
          rc0 == 0 and "WMS-SYS-4" in o0 and "0.7.2-hotfix" in o0 and "можно обновить" in o0 and sha(p) == h0, o0[-300:])
    new = os.path.join(d, "NEW")
    rc1, o1 = upgrade("build", p, "--to", new)
    rep = open(os.path.join(new, "UPGRADE_REPORT.md"), encoding="utf-8").read() if os.path.exists(os.path.join(new, "UPGRADE_REPORT.md")) else ""
    summ = open(os.path.join(new, "UPGRADE_SUMMARY.txt"), encoding="utf-8").read() if os.path.exists(os.path.join(new, "UPGRADE_SUMMARY.txt")) else ""
    R.add(c, "сборка в новой папке: копия книги и журнала, изменения 0.8, сверка каждой ячейки — 0 расхождений; WMS 0.8 открывает рабочую "
             "копию проверки: самопроверка 0 ошибок, «остаток = движения», оракул без новых расхождений; старая книга и журнал побайтно те же",
          rc1 == 0 and "расхождений: 0" in rep and "ошибок 0" in rep and "все ЕИ сходятся" in rep and "новых 0" in rep and sha(p) == h0
          and {f: sha(f) for f in glob.glob(os.path.join(d, "OLD", "WMS_Journal", "*"))} == j0 and not os.path.exists(os.path.join(new, ".verify")),
          o1[-300:] + " || " + rep[-600:])
    names = ("Болт", "Ромашка", "Петров", "Кабель")
    R.add(c, "отчёт без данных склада (UPGRADE_SUMMARY.txt): счётчики, хэши, PASS — без наименований, поставщиков и людей; его можно прислать",
          "RESULT=BUILD_PASS" in summ and "problems=0" in summ and not any(n in summ for n in names), summ[-400:])
    nb = os.path.join(new, os.path.basename(p))
    s = Session(nb, macros=0)
    try:
        sv = {k: s.sysv(k) for k in ("SCHEMA", "UPGRADE_FROM", "LAST_SEQ", "NEXT_EI")}
        legn = sum(1 for r in s.doc.Sheets.getByName("_LEG").getCellRangeByPosition(0, 1, 0, 6000).getDataArray() if r[0])
        has = all(s.doc.Sheets.hasByName(n) for n in ("_LEG", "Сверка переноса"))
        prot = all(s.doc.Sheets.getByName(n).isProtected() for n in ("_LEG", "Сверка переноса", "Наличие", "Заказы", "Выдачи"))
        hidden = not s.doc.Sheets.getByName("_LEG").IsVisible
    finally:
        s.close()
    R.add(c, "новая книга: схема WMS-SYS-5, UPGRADE_FROM (папка, seq, хэш, дата, размер журнала), реестр переноса из журнала (14 ЕИ), лист "
             "«Сверка переноса»; листы защищены, реестр скрыт", sv["SCHEMA"] == "WMS-SYS-5" and str(sv["UPGRADE_FROM"]).count("|") == 4
          and has and prot and hidden and legn == 14, f"{sv}; leg {legn}")
    rc2, o2 = upgrade("promote", "--to", new)
    rc3, o3 = upgrade("promote", "--to", new)
    R.add(c, "«СДЕЛАТЬ РАБОЧЕЙ»: новая книга зарегистрирована в своей папке, самопроверка, сохранена и открыта снова; повторно — отказ; "
             "файл «ПЕРЕХОД_ВЫПОЛНЕН.txt» с объяснением", rc2 == 0 and rc3 == 1 and "уже рабочая" in o3
          and os.path.exists(os.path.join(new, "ПЕРЕХОД_ВЫПОЛНЕН.txt")) and sha(p) == h0, f"{o2[-200:]}; {o3[-120:]}")
    s = ready(Session(nb))
    try:
        rep = s.report()
        R.add(c, "WMS 0.8 в новой папке: «работа разрешена», ядро 0.8.0-prime, предупреждений о старой книге нет",
              state(s) == ("CLEAN", "") and "ядро 0.8.0-prime" in rep and "старой книге" not in rep, rep[:200])
        osh = s.doc.Sheets.getByName("Заказы")
        r427 = [r for r in range(1, 40) if osh.getCellByPosition(21, r).getString() == "ЕИ-00000427"][0]
        lf = s.Rc("LegacyFixRow", r427, "100", "58", "А-1", TODAY, "пересчёт")
        R.add(c, "в обновлённой книге «Коррекция переноса» ЕИ-00000427 (перенесён с остатком 60, выдано 5 в 0.7.2): остаток 55 → 53",
              lf.startswith("OK") and s.stock(427) == 53.0, f"{lf}; {s.stock(427)}")
        oracle(c, s, "журнал 0.7.2 + 0.8 в обновлённой книге", initial={}, legacy_mn_empty=False)
    finally:
        s.close(save=True)
    # someone works in the old book after the switch
    s = Session(p)
    try:
        r = issue(s, 9, "ЕИ-00000427", "1", who="Петров")
    finally:
        s.close(save=True)
    s = ready(Session(nb))
    try:
        rep = s.report()
        R.add(c, "в старой книге после перехода провели выдачу — новая WMS при запуске предупреждает: «в старой книге … проводились операции … "
                 "их нет в этой WMS»", r.startswith("OK") and "в старой книге" in rep and "их нет в этой WMS" in rep, f"{r}; {rep[:400]}")
    finally:
        s.close()


@case
def pr12_upgrade_stops():
    c = "PR12"
    d, p = old_copy("pr12")
    h0 = sha(p)
    outs = {}
    jdir = os.path.join(d, "OLD", "WMS_Journal")
    # the book open (the lock of WMS, the lock of LibreOffice)
    open(os.path.join(jdir, "wms.lock"), "w").write("x")
    outs["lock"] = upgrade("check", p)
    os.remove(os.path.join(jdir, "wms.lock"))
    lo = os.path.join(d, "OLD", ".~lock." + os.path.basename(p) + "#")
    open(lo, "w").write("x")
    outs["lolock"] = upgrade("check", p)
    os.remove(lo)
    # a sheet with a foreign password
    q = os.path.join(d, "F", os.path.basename(p))
    shutil.copytree(os.path.join(d, "OLD"), os.path.join(d, "F"))
    raw(q, lambda doc: (doc.Sheets.getByName("Иной приход").unprotect("wms"), doc.Sheets.getByName("Иной приход").protect(FOREIGN)))
    outs["foreign"] = upgrade("check", q)
    # a tail of the journal (an operation the book does not hold)
    t = os.path.join(d, "T", os.path.basename(p))
    shutil.copytree(os.path.join(d, "OLD"), os.path.join(d, "T"))
    raw(t, lambda doc: (doc.Sheets.getByName("_SYS").unprotect("wms"), doc.Sheets.getByName("_SYS").getCellByPosition(1, 4).setValue(
        doc.Sheets.getByName("_SYS").getCellByPosition(1, 4).getValue() - 1), doc.Sheets.getByName("_SYS").protect("wms")))
    outs["tail"] = upgrade("check", t)
    # an unfinished operation (the marker STARTED)
    m = os.path.join(d, "M", os.path.basename(p))
    shutil.copytree(os.path.join(d, "OLD"), os.path.join(d, "M"))
    raw(m, lambda doc: (doc.Sheets.getByName("_SYS").unprotect("wms"), doc.Sheets.getByName("_SYS").getCellByPosition(1, 10).setString("STARTED"),
                        doc.Sheets.getByName("_SYS").protect("wms")))
    outs["marker"] = upgrade("check", m)
    # lines of another WMS in the journal
    fj = os.path.join(d, "J", os.path.basename(p))
    shutil.copytree(os.path.join(d, "OLD"), os.path.join(d, "J"))
    jf = sorted(glob.glob(os.path.join(d, "J", "WMS_Journal", "WMS_journal_*.csv")))[-1]
    body = "J1;999;2026-01-01 00:00:00;FFFFFFFFFFFFFFFF;ISSUE;NO=1"
    with open(jf, "a", encoding="utf-8") as f:
        f.write(f"{body};END;{len(body)}\n")
    outs["foreign_journal"] = upgrade("check", fj)
    # not a book of WMS; no file; an upgraded book
    plain = os.path.join(d, "plain.ods")
    s = Session(new_wms("pr12_plain"), macros=0)
    try:
        s.doc.Sheets.getByName("_SYS").unprotect("wms")
        s.doc.unprotect("wms")
        s.doc.Sheets.removeByName("_SYS")
        s.doc.storeToURL(__import__("uno").systemPathToFileUrl(plain), ())
    finally:
        s.close()
    outs["plain"] = upgrade("check", plain)
    outs["nofile"] = upgrade("check", os.path.join(d, "нет.ods"))
    outs["new08"] = upgrade("check", new_wms("pr12_08"))
    want = {"lock": "WMS открыта", "lolock": "открыта в LibreOffice", "foreign": "«Иной приход» защищены не паролем WMS", "tail": "которых нет в книге",
            "marker": "незавершённая операция", "foreign_journal": "другой WMS", "plain": "не книга WMS", "nofile": "нет файла", "new08": "уже версии 0.8"}
    bad = {k: v[1][-160:] for k, v in outs.items() if v[0] != 1 or want[k] not in v[1]}
    R.add(c, "стоп-проверки с понятной причиной: WMS открыта; книга открыта в LibreOffice; лист с чужим паролем; хвост журнала; незавершённая "
             "операция; строки другой WMS в журнале; не книга WMS; нет файла; книга уже 0.8 — код 1, книга не изменена",
          not bad and sha(p) == h0, str(bad)[:900])
    # the folders
    os.makedirs(os.path.join(d, "BUSY"))
    open(os.path.join(d, "BUSY", "x.txt"), "w").write("x")
    r1 = upgrade("build", p, "--to", os.path.join(d, "BUSY"))
    r2 = upgrade("build", p, "--to", os.path.join(d, "OLD"))
    R.add(c, "новая папка не пуста или совпадает с папкой книги — отказ", r1[0] == 1 and "не пуста" in r1[1] and r2[0] == 1, f"{r1[1][-100:]}; {r2[1][-100:]}")
    # a fault in the middle: all or nothing
    res = []
    for point in ("after_copy", "after_changes", "before_wms"):
        to = os.path.join(d, "FAULT_" + point)
        rc, o = upgrade("build", p, "--to", to, fault=point)
        left = sorted(os.listdir(to))
        res.append((point, rc, left))
    R.add(c, "сбой посреди сборки (после копии, после изменений, перед проверкой WMS 0.8): код 2, новая книга и журнал удалены, остались только "
             "отчёты; старая книга та же",
          all(rc == 2 and left == sorted(["UPGRADE_REPORT.md", "UPGRADE_STATE.json", "UPGRADE_STATE.txt", "UPGRADE_SUMMARY.txt", "progress.txt"]) for _, rc, left in res)
          and sha(p) == h0, str(res))
    rc, o = upgrade("promote", "--to", os.path.join(d, "FAULT_after_copy"))
    R.add(c, "«СДЕЛАТЬ РАБОЧЕЙ» после неудачной сборки — отказ", rc == 1 and "нет успешной сборки" in o, o[-160:])
    # rehearsal; work in the old book between build and promote
    rh = os.path.join(d, "REH")
    rr = upgrade("build", p, "--to", rh, "--rehearsal")
    rp = upgrade("promote", "--to", rh)
    R.add(c, "репетиция: сборка и сверка как при обновлении, отчёт «REHEARSAL_PASS»; «СДЕЛАТЬ РАБОЧЕЙ» для репетиции — отказ",
          rr[0] == 0 and "REHEARSAL_PASS" in open(os.path.join(rh, "UPGRADE_SUMMARY.txt"), encoding="utf-8").read() and rp[0] == 1 and "репетиция" in rp[1],
          f"{rr[1][-120:]}; {rp[1][-120:]}")
    nw = os.path.join(d, "NEW")
    rb = upgrade("build", p, "--to", nw)
    s = Session(p)
    try:
        ri = issue(s, 9, "ЕИ-00000427", "1", who="Петров")
    finally:
        s.close(save=True)
    rp2 = upgrade("promote", "--to", nw)
    R.add(c, "после сборки в старой книге провели выдачу — «СДЕЛАТЬ РАБОЧЕЙ» отказывает: соберите заново, чтобы ничего не потерять",
          rb[0] == 0 and ri.startswith("OK") and rp2[0] == 1 and "работали после сборки" in rp2[1], f"{ri}; {rp2[1][-160:]}")
    rb2 = upgrade("build", p, "--to", os.path.join(d, "NEW2"))
    s = Session(os.path.join(d, "NEW2", os.path.basename(p)), macros=0)
    try:
        s.doc.Sheets.getByName("Получатели").getCellByPosition(5, 5).setString("правка")
        s.doc.store()
    finally:
        s.close()
    rp3 = upgrade("promote", "--to", os.path.join(d, "NEW2"))
    R.add(c, "новую книгу изменили после сборки — «СДЕЛАТЬ РАБОЧЕЙ» отказывает", rb2[0] == 0 and rp3[0] == 1 and "изменилась после сборки" in rp3[1],
          rp3[1][-160:])


@case
def pr13_transfer_writes_registry():
    """a transfer of the old table by this release fills the registry of the transfer itself"""
    c = "PR13"
    import run_legacy_transfer as LT
    cd, st, work, res = LT.transfer("pr13", LT.full_rows(), steps=("check", "build"))
    ok = all(res[k][0] == 0 for k in res)
    s = ready(Session(LT.test_book(work)))
    try:
        legs = [r for r in s.doc.Sheets.getByName("_LEG").getCellRangeByPosition(0, 1, 10, 6000).getDataArray() if r[0]]
        rep = s.B("LegacyReport", module="WmsLegacy")
        bc = s.B("BalanceCheck", module="WmsLegacy")
        r427 = [r for r in legs if r[0] == "ЕИ-00000427"]
        R.add(c, "перенос старой таблицы программой этого выпуска: реестр переноса заполнен операциями переноса (ЕИ-00000427: пришло 100, "
                 "остаток 60, израсходовано 40, X «60», «колонка X»); «Сверка переноса» и «остаток = движения» без расхождений",
              ok and len(legs) == 14 and r427 and r427[0][1:5] == ("RECEIPT", 100.0, 60.0, 40.0) and r427[0][6] == "X"
              and "расхождений с движениями 0" in rep and bc == "", f"{[res[k][1][-120:] for k in res if res[k][0]]}; {len(legs)}; {r427}; {rep}")
    finally:
        s.close()


# ================================================================ PR14 the wizard WMS_UPGRADE.ods (the buttons of the release)

def release_prime():
    rel = os.path.join(OUT, "rel_prime")
    if not os.path.isfile(os.path.join(rel, "WMS_UPGRADE.ods")):
        shutil.rmtree(rel, ignore_errors=True)
        r = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "build_release.py"), rel], capture_output=True, text=True, timeout=2400,
                           env=ENV, cwd=ROOT)
        if r.returncode != 0:
            raise RuntimeError("build_release: " + (r.stdout + r.stderr)[-2000:])
    return rel


class Wizard:
    """WMS_UPGRADE.ods of the release in a case folder (with tools/, src/ and WMS_TOOLBOX/ next to it), open with its macros"""

    def __init__(self, folder):
        from harness import PROFILES, unique
        from wmslo import Office
        rel = release_prime()
        os.makedirs(folder, exist_ok=True)
        shutil.copy2(os.path.join(rel, "WMS_UPGRADE.ods"), folder)
        for d in ("tools", "src", "docs", "WMS_TOOLBOX"):
            shutil.copytree(os.path.join(rel, d), os.path.join(folder, d))
        self.folder = folder
        self.o = Office(unique("upwz"), PROFILES)
        self.doc = self.o.load(os.path.join(folder, "WMS_UPGRADE.ods"), macros=4)
        self.B("UpTestAuto", True)

    def B(self, f, *a):
        return self.o.basic(self.doc, "UpMain", f, *a)

    def msg(self):
        return str(self.B("UpTestLastMessage") or "")

    def cell(self, r, sheet="ОБНОВЛЕНИЕ", c=2):
        return self.doc.Sheets.getByName(sheet).getCellByPosition(c, r).getString()

    def enabled(self, name):
        return self.doc.Sheets.getByName("ОБНОВЛЕНИЕ").getDrawPage().getForms().getByIndex(0).getByName(name).Enabled

    def close(self):
        try:
            en = self.o.desktop.getComponents().createEnumeration()
            while en.hasMoreElements():
                c = en.nextElement()
                if hasattr(c, "getURL") and c.getURL() != self.doc.getURL() and hasattr(c, "setModified"):
                    # like a user closing the window: the unload events of WMS release its lock
                    c.setModified(False)
                    try:
                        self.o.dispatch(c, ".uno:CloseDoc")
                    except Exception:
                        c.close(True)
        except Exception:
            pass
        self.doc.setModified(False)
        self.doc.close(True)
        self.o.terminate()


def old_copy_at(folder):
    """a copy of the warehouse of 0.7.2 in folder (WMS_PROD.ods and its journal), registered there by its own macros"""
    book, facts = warehouse()
    shutil.rmtree(folder, ignore_errors=True)
    shutil.copytree(os.path.dirname(book), folder, ignore=shutil.ignore_patterns("WMS_TOOLBOX", "backup", "WMS_Backups"))
    p = os.path.join(folder, os.path.basename(book))
    s = Session(p)
    try:
        if "REG=other" in s.B("StateDump"):
            s.B("ActionRegisterHere", False)
    finally:
        s.close(save=True)
    return p


@case
def pr14_wizard():
    c = "PR14"
    d = os.path.join(OUT, "cases", "pr14")
    shutil.rmtree(d, ignore_errors=True)
    old = old_copy_at(os.path.join(d, "WMS_RELEASE_0.7.2", "WMS_WORK"))
    h0 = sha(old)
    w = Wizard(os.path.join(d, "WMS_RELEASE_0.8.0"))
    try:
        new_dir = os.path.join(w.folder, "WMS_WORK")
        R.add(c, "мастер при открытии сам находит рабочую книгу 0.7 в папке прежнего выпуска и предлагает новую папку WMS_WORK своего выпуска",
              w.cell(4) == old and w.cell(5).rstrip("/") == new_dir and not w.enabled("btnPromote"), f"{w.cell(4)}; {w.cell(5)}")
        w.B("UpTestConfirm", 1)
        w.B("BtnUpCheck")
        m1 = w.msg()
        R.add(c, "«1. ПРОВЕРИТЬ»: «можно обновить», шаг отмечен ✓, книга не изменена", m1.startswith("OK:") and "можно обновить" in m1
              and w.cell(8).startswith("✓") and sha(old) == h0, f"{m1[:300]}; {w.cell(8)}")
        w.B("BtnUpRehearsal")
        m2 = w.msg()
        reh = glob.glob(os.path.join(w.folder, "UPGRADE_WORK", "репетиция_*"))
        R.add(c, "«2. РЕПЕТИЦИЯ»: прошла в своей папке (UPGRADE_WORK/репетиция_…), шаг ✓, «СДЕЛАТЬ РАБОЧЕЙ» всё ещё недоступна, книга не изменена",
              m2.startswith("OK:") and len(reh) == 1 and w.cell(9).startswith("✓") and not w.enabled("btnPromote") and sha(old) == h0
              and not os.path.exists(new_dir), f"{m2[:200]}; {reh}; {w.cell(9)}")
        w.B("BtnUpBuild")
        m3 = w.msg()
        rep = [w.cell(r, "Отчёт", 0) for r in range(0, 120)]
        R.add(c, "«3. СОБРАТЬ И СВЕРИТЬ»: новая книга в WMS_WORK выпуска 0.8, шаг ✓, кнопка «СДЕЛАТЬ РАБОЧЕЙ» доступна, отчёт на листе «Отчёт» "
                 "(заголовки, пункты, «НОВАЯ КНИГА ГОТОВА»)", m3.startswith("OK:") and w.cell(10).startswith("✓") and w.enabled("btnPromote")
              and os.path.isfile(os.path.join(new_dir, "WMS_PROD.ods")) and any("НОВАЯ КНИГА ГОТОВА" in x for x in rep)
              and any(x.startswith("•  ") for x in rep) and sha(old) == h0, f"{m3[:200]}; {w.cell(10)}; {[x for x in rep if x][:6]}")
        w.B("UpTestConfirm", 0)
        w.B("BtnUpPromote")
        m4 = w.msg()
        R.add(c, "«4. СДЕЛАТЬ РАБОЧЕЙ» — вопрос с объяснением (работать только в новой, старая — откат); «Нет» — ничего не изменено",
              m4 == "Ничего не изменено." and "Старая книга остаётся как есть" in str(w.B("UpTestLastAsk")) and w.enabled("btnPromote")
              and not os.path.exists(os.path.join(new_dir, "ПЕРЕХОД_ВЫПОЛНЕН.txt")), f"{m4}; {w.B('UpTestLastAsk')[:200]}")
        w.B("UpTestConfirm", 1)
        w.B("BtnUpPromote")
        m5 = w.msg()
        opened = []
        en = w.o.desktop.getComponents().createEnumeration()
        while en.hasMoreElements():
            x = en.nextElement()
            if hasattr(x, "getURL") and x.getURL():
                opened.append(__import__("uno").fileUrlToSystemPath(x.getURL()))
        R.add(c, "переход: рабочая WMS — новая книга, шаг ✓, кнопка снова недоступна, «ПЕРЕХОД_ВЫПОЛНЕН.txt», инструменты выпуска рядом с книгой, "
                 "новая WMS открыта по «Да»; старая книга не изменена",
              m5.startswith("OK:") and w.cell(11).startswith("✓") and not w.enabled("btnPromote")
              and os.path.exists(os.path.join(new_dir, "ПЕРЕХОД_ВЫПОЛНЕН.txt")) and os.path.isdir(os.path.join(new_dir, "WMS_TOOLBOX"))
              and os.path.join(new_dir, "WMS_PROD.ods") in opened and sha(old) == h0, f"{m5[:200]}; {w.cell(11)}; {opened}")
        w.B("BtnUpBuild")
        m6 = w.msg()
        R.add(c, "повторная сборка в папку, где уже рабочая WMS 0.8 — отказ «работайте в ней»", m6.startswith("ERR:") and "уже рабочая WMS 0.8" in m6, m6)
        shutil.move(os.path.join(w.folder, "docs"), os.path.join(w.folder, "docs.away"))
        w.doc.Sheets.getByName("ОБНОВЛЕНИЕ").getCellByPosition(2, 4).setString(old)
        w.B("BtnUpCheck")
        m9 = w.msg()
        shutil.move(os.path.join(w.folder, "docs.away"), os.path.join(w.folder, "docs"))
        R.add(c, "папка выпуска неполная (нет docs/) — проверка останавливается и говорит, чего нет", m9.startswith("ERR:") and "неполная" in m9
              and "docs/OPERATOR_GUIDE.md" in m9, m9[:300])
        w.doc.Sheets.getByName("ОБНОВЛЕНИЕ").getCellByPosition(2, 4).setString("")
        w.B("BtnUpCheck")
        m7 = w.msg()
        w.doc.Sheets.getByName("ОБНОВЛЕНИЕ").getCellByPosition(2, 4).setString(os.path.join(d, "нет.ods"))
        w.B("BtnUpCheck")
        m8 = w.msg()
        R.add(c, "книга не выбрана / нет файла — понятное сообщение, программа не запускается", m7.startswith("ERR:") and "не выбрана" in m7
              and m8.startswith("ERR:") and "нет файла" in m8, f"{m7}; {m8}")
    finally:
        w.close()
    s = ready(Session(os.path.join(new_dir, "WMS_PROD.ods")))
    try:
        R.add(c, "новая WMS после мастера: «работа разрешена», ядро 0.8, самопроверка без ошибок", state(s) == ("CLEAN", "")
              and s.B("ActionSelfCheck").startswith("САМОПРОВЕРКА WMS: ошибок 0"), s.report()[:200])
    finally:
        s.close()


def main():
    sel = [a.lower() for a in sys.argv[1:]]
    for fn in CASES:
        if sel and not any(fn.__name__.startswith(x) for x in sel):
            continue
        try:
            fn()
        except Exception as e:
            R.error(fn.__name__.split("_")[0].upper(), e)
    R.save(os.path.join(OUT, "results_prime.json"))
    items = R.items
    lines = ["# PRIME 0.8 — этап 1: результаты", "", f"PASS {sum(1 for x in items if x['status'] == 'PASS')}, "
             f"FAIL {sum(1 for x in items if x['status'] == 'FAIL')}, SKIP {sum(1 for x in items if x['status'] == 'SKIP')}", ""]
    lines += [f"- [{x['status']}] {x['case']} {x['test']}" for x in items]
    open(os.path.join(OUT, "TEST_REPORT_prime.md"), "w", encoding="utf-8").write("\n".join(lines) + "\n")
    print(f"PASS {sum(1 for x in items if x['status'] == 'PASS')} FAIL {sum(1 for x in items if x['status'] == 'FAIL')} "
          f"SKIP {sum(1 for x in items if x['status'] == 'SKIP')}")
    return 0 if all(x["status"] == "PASS" for x in items) else 1


if __name__ == "__main__":
    sys.exit(main())
