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
from harness import OUT, ROOT, Results, Session, new_wms, wms_versions  # noqa: E402
import journal_oracle as jo  # noqa: E402

VERS = wms_versions()

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
        has = all(s.doc.Sheets.hasByName(n) for n in ("_LEG", "Сверка переноса", "_PACK"))
        prot = all(s.doc.Sheets.getByName(n).isProtected() for n in ("_LEG", "Сверка переноса", "Наличие", "Заказы", "Выдачи", "_PACK"))
        hidden = not s.doc.Sheets.getByName("_LEG").IsVisible and not s.doc.Sheets.getByName("_PACK").IsVisible
    finally:
        s.close()
    R.add(c, "новая книга: схема WMS-SYS-5, UPGRADE_FROM (папка, seq, хэш, дата, размер журнала), реестр переноса из журнала (14 ЕИ), лист "
             "«Сверка переноса», лист фасовки _PACK; листы защищены, служебные скрыты", sv["SCHEMA"] == "WMS-SYS-5" and str(sv["UPGRADE_FROM"]).count("|") == 4
          and has and prot and hidden and legn == 14, f"{sv}; leg {legn}")
    rc2, o2 = upgrade("promote", "--to", new)
    rc3, o3 = upgrade("promote", "--to", new)
    R.add(c, "«СДЕЛАТЬ РАБОЧЕЙ»: новая книга зарегистрирована в своей папке, самопроверка, сохранена и открыта снова; повторно — отказ; "
             "файл «ПЕРЕХОД_ВЫПОЛНЕН.txt» с объяснением", rc2 == 0 and rc3 == 1 and "уже рабочая" in o3
          and os.path.exists(os.path.join(new, "ПЕРЕХОД_ВЫПОЛНЕН.txt")) and sha(p) == h0, f"{o2[-200:]}; {o3[-120:]}")
    s = ready(Session(nb))
    try:
        rep = s.report()
        R.add(c, f"WMS {VERS['WMS_PRODUCT_VERSION']} в новой папке: «работа разрешена», ядро {VERS['WMS_CORE_VERSION']}, предупреждений о старой книге нет",
              state(s) == ("CLEAN", "") and f"ядро {VERS['WMS_CORE_VERSION']}" in rep and "старой книге" not in rep, rep[:200])
        pw = [ln.split("\t") for ln in s.B("PreWorkCheck", module="WmsStatus").split("\n")]
        prot = [f for f in pw if len(f) >= 3 and f[1] == "защиты"]
        R.add(c, "«Проверка перед работой» обновлённой книги: защиты без замечаний («Топология» и «Получатели» — листы ввода, не защищаются; "
                 "_LEG и _PACK скрыты)", len(prot) == 1 and prot[0][0] == "OK", str(prot))
        osh = s.doc.Sheets.getByName("Заказы")
        r427 = [r for r in range(1, 40) if osh.getCellByPosition(21, r).getString() == "ЕИ-00000427"][0]
        lf = s.Rc("LegacyFixRow", r427, "100", "58", "А-1", TODAY, "пересчёт")
        R.add(c, "в обновлённой книге «Коррекция переноса» ЕИ-00000427 (перенесён с остатком 60, выдано 5 в 0.7.2): остаток 55 → 53",
              lf.startswith("OK") and s.stock(427) == 53.0, f"{lf}; {s.stock(427)}")
        names = [s.doc.Sheets.getByIndex(i).Name for i in range(s.doc.Sheets.getCount()) if s.doc.Sheets.getByIndex(i).IsVisible]
        db = s.B("DashRefresh", module="WmsDash")
        fd = s.B("FindRun", "427", module="WmsFind")
        card = page(s, "Поиск", 9, 9)[0][1]
        main_btn = s.doc.Sheets.getByName("Главная").getDrawPage().getCount()
        R.add(c, "обновлённая книга — с новыми страницами: «Дашборд» первым, «Поиск», «Карта склада», «Топология»; «Главная» 1.0 (25 кнопок над строками состояния, с «Справкой» и «?»); "
                 "дашборд считается; карточка ЕИ-00000427 — остаток сходится с движениями переноса и 0.7.2",
              names[:2] == ["Дашборд", "Главная"] and "Поиск" in names and "Карта склада" in names and "Топология" in names and main_btn == 25
              and db.startswith("OK") and fd.startswith("OK:ЕИ-00000427") and card.startswith("✓"), f"{names}; {main_btn}; {db}; {fd}; {card}")
        btns = {}
        for n in ("Заказы", "Иной приход", "Выдачи", "Корректировки", "Наличие"):
            fm = s.doc.Sheets.getByName(n).DrawPage.Forms.getByIndex(0)
            btns[n] = [fm.getByIndex(i).Name for i in range(fm.Count)]
        R.add(c, "в обновлённой книге и на старых листах — кнопки 0.8: «Куда положить?» («Заказы», «Иной приход»), «Подобрать ЕИ» и «Фасовка ЕИ» "
                 "(«Выдачи», «Корректировки»), «Фасовка ЕИ» на «Наличие», каждая один раз", all(btns[n].count("btnWherePut") == 1 for n in ("Заказы", "Иной приход"))
              and all(btns[n].count("btnPickEI") == 1 and btns[n].count("btnPackEI") == 1 for n in ("Выдачи", "Корректировки"))
              and btns["Наличие"].count("btnPackEI") == 1, str(btns))
        wcf = s.doc.Sheets.getByName("Заказы").getCellRangeByPosition(22, 1, 22, 1048575).ConditionalFormat
        wst = [wcf.getByIndex(k).getStyleName() for k in range(wcf.getCount())]
        acf = s.doc.Sheets.getByName("Заказы").getCellRangeByPosition(0, 1, 15, 1048575).ConditionalFormat
        R.add(c, "в обновлённой книге статус W «Заказов» — в цвете, и линия начала заказа на нём осталась (начало заказа со статусом — "
                 "линия и цвет, затем линия, затем цвет); A:P — линия заказа, как в 0.7.2",
              wst == ["ST_Done_Line", "ST_Wait_Line", "ST_Part_Line", "ST_Gray_Line", "WMS_OrderStart", "ST_Done", "ST_Wait", "ST_Part", "ST_Gray"]
              and acf.getCount() == 1 and acf.getByIndex(0).getStyleName() == "WMS_OrderStart", str(wst))
        pk = s.B("PackSet", "ЕИ-00000427", "4", "кан", "", module="WmsPack")
        # 1.0: a recipient not in «Получатели» is asked about (a window a test of a working book cannot answer) — added as people do
        rc = s.doc.Sheets.getByName("Получатели")
        rr = 1
        while rc.getCellByPosition(1, rr).getString():
            rr += 1
        rc.getCellByPosition(1, rr).setString("Петров")
        s.I("OnRecipientsChange")
        put(s, "Выдачи", 300, L="ЕИ-00000427", E="2 кан", H=serial(), I="Петров")
        pi = s.I("IssuePostRow", 300)
        R.add(c, "в обновлённой книге фасовка и выдача частями работают: в 1 шт — 4 кан, «2 кан» — 0,5 шт", pk.startswith("OK") and pi.startswith("OK")
              and s.iss(300)[4] == 0.5, f"{pk}; {pi}; {s.iss(300)[4]}")
        oracle(c, s, "журнал 0.7.2 + 0.8 в обновлённой книге", initial={}, legacy_mn_empty=False)
    finally:
        s.close(save=True)
    # someone opens the old book after the switch: it is locked (1.0); after «Снять блокировку WMS» an issue is posted in it
    lock = open(os.path.join(d, "OLD", "WMS_Journal", "wms.lock"), encoding="utf-8").read() if os.path.exists(
        os.path.join(d, "OLD", "WMS_Journal", "wms.lock")) else ""
    s = Session(p)
    try:
        rep_old = s.report()
        blocked = "LOCKED" in s.B("StateDump") and "ЭТА КНИГА БОЛЬШЕ НЕ РАБОЧАЯ" in rep_old
        s.B("ActionForceUnlock")
        r = issue(s, 9, "ЕИ-00000427", "1", who="Петров")
    finally:
        s.close(save=True)
    R.add(c, "старая книга после перехода заперта: открывается с блокировкой «ЭТА КНИГА БОЛЬШЕ НЕ РАБОЧАЯ: WMS перенесена в папку …»; "
             "журнал и книга не тронуты до «Снять блокировку WMS»", blocked and "session=MOVED-" in lock, f"{rep_old[:300]}; {lock[:160]}")
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
    # a book of the intermediate PRIME 0.8: the same schema, the core 0.8.0-prime (written at its first start)
    real08 = new_wms("pr12_real08")
    s = Session(real08, macros=0)
    try:
        sy = s.doc.Sheets.getByName("_SYS")
        sy.unprotect("wms")
        sy.getCellByPosition(1, 3).setString("0.8.0-prime")
        s.doc.store()
    finally:
        s.close()
    h08 = sha(real08)
    outs["real08"] = upgrade("check", real08)
    want = {"lock": "WMS открыта", "lolock": "открыта в LibreOffice", "foreign": "«Иной приход» защищены не паролем WMS", "tail": "которых нет в книге",
            "marker": "незавершённая операция", "foreign_journal": "другой WMS", "plain": "не книга WMS", "nofile": "нет файла", "new08": f"уже версии {VERS['WMS_PRODUCT_VERSION']}",
            "real08": "промежуточный выпуск PRIME 0.8"}
    bad = {k: v[1][-160:] for k, v in outs.items() if v[0] != 1 or want[k] not in v[1]}
    R.add(c, "стоп-проверки с понятной причиной: WMS открыта; книга открыта в LibreOffice; лист с чужим паролем; хвост журнала; незавершённая "
             "операция; строки другой WMS в журнале; не книга WMS; нет файла; книга уже этой версии; книга промежуточной 0.8 — «мастер переносит "
             "книгу 0.7, ничего не изменено» — код 1, книги не изменены",
          not bad and sha(p) == h0 and sha(real08) == h08 and "Ничего не изменено" in outs["real08"][1], str(bad)[:900])
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
    w = Wizard(os.path.join(d, f"WMS_RELEASE_{VERS['WMS_PRODUCT_VERSION']}"))
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
        R.add(c, "повторная сборка в папку, где уже рабочая WMS 1.0 — отказ «работайте в ней»", m6.startswith("ERR:") and "уже рабочая WMS 1.0" in m6, m6)
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
        R.add(c, f"новая WMS после мастера: «работа разрешена», ядро {VERS['WMS_CORE_VERSION']}, самопроверка без ошибок", state(s) == ("CLEAN", "")
              and s.B("ActionSelfCheck").startswith("САМОПРОВЕРКА WMS: ошибок 0"), s.report()[:200])
    finally:
        s.close()


# ================================================================ PR15 … PR18 the pages of PRIME: «Дашборд», «Поиск», «Карта склада»

def page(s, name, r0, r1, c0=0, c1=13):
    sh = s.doc.Sheets.getByName(name)
    return [[sh.getCellByPosition(c, r).getString() for c in range(c0, c1 + 1)] for r in range(r0, r1 + 1)]


def unmodified_call(s, func, *a, module):
    """a view refreshed on a book saved just before: (result, the book still not «modified», the journal untouched)"""
    s.doc.store()
    j0 = njournal(s)
    res = s.B(func, *a, module=module)
    return res, not s.doc.isModified(), njournal(s) == j0


@case
def pr15_dashboard():
    c = "PR15"
    p = new_wms("pr15")
    s = ready(Session(p))
    try:
        order(s, 1, "1", "Болт", qty="100", Q=serial(-3))
        receive(s, 1, "40")
        order(s, 2, "1", "Гайка", qty="50", Q=serial(0))
        order(s, 3, "1", "Шайба", qty="5", Q=serial(4))
        order(s, 4, "1", "Винт", qty="5")
        iss = [issue(s, 1, "ЕИ-00000001", "2", who="Петров"), issue(s, 2, "ЕИ-00000001", "1", who="Сидоров"), issue(s, 3, "ЕИ-00000002", "1", who="Петров")]
        put(s, "Выдачи", 4, L="ЕИ-00000003", E="1", I="Петров")          # typed, not posted
        r, unmod, nj = unmodified_call(s, "DashRefresh", module="WmsDash")
        v = page(s, "Дашборд", 0, 67, 0, 14)
        tiles = [v[4][1 + 2 * i] for i in range(6)]
        R.add(c, "дашборд считается по листам книги: плитки — ЕИ с остатком, ожидается позиций (частично полученная тоже), пришло и выдано сегодня, "
                 "машины, внимание (непроведённая строка выдачи); книга не помечена изменённой, журнал не тронут",
              r.startswith("OK") and all(x.startswith("OK") for x in iss) and tiles[1] == "4" and tiles[2] == "1" and tiles[3] == "3" and tiles[5] == "1"
              and "непроведённых 1" in v[5][11] and "получателей 2" in v[5][7] and unmod and nj, f"{r}; {tiles}; {v[5]}")
        lst = [row[1:7] for row in v[9:13]]
        R.add(c, "«Ожидается»: позиции по ожидаемой дате (Q), раньше — выше: просрочено 3 дн., сегодня, через 4 дн.; без даты — в счётчике плитки",
              lst[0][1].startswith("Болт · осталось") and "просрочено 3 дн." == lst[0][5] and lst[1][1].startswith("Гайка") and lst[1][5] == "сегодня" and lst[2][5] == "через 4 дн."
              and "без даты 1" in v[5][3], str(lst))
        ops = [row[7:13] for row in v[9:14]]
        R.add(c, "«Последние операции» — из журнала, новые сверху: выдача, кому, количество с единицей", ops[0][1] == "Выдача" and "ЕИ-00000002" in ops[0][3]
              and "Петров" in ops[0][3] and ops[0][5].startswith("1 "), str(ops[:3]))
        top = [row[1:7] for row in v[24:26]]
        R.add(c, "«Чаще всего выдают — 30 дней»: по числу выдач, с количеством и единицей", top[0][3] == "2" and top[0][4] == "3", str(top))
        days = v[37]
        R.add(c, "«Последние 14 дней»: сегодня первым, число выдач и полоска", days[2] == "3" and "█" in days[3] and days[8] == "1", str(days[:10]))
        R.add(c, "в шапке — состояние WMS: «Работа разрешена»", "Работа разрешена" in v[0][8] and v[0][14].startswith("CLEAN"), v[0][8])
        # the page shown again with no new operation: nothing recounted; after an operation: recounted
        foot0 = v[67][1]
        s.B("OnDashFocus", module="WmsDash")
        same = page(s, "Дашборд", 67, 67)[0][1] == foot0
        issue(s, 5, "ЕИ-00000002", "1", who="Иванов")
        s.B("OnDashFocus", module="WmsDash")
        v2 = page(s, "Дашборд", 4, 4)[0]
        R.add(c, "показ страницы без новых операций — не пересчитывается; после выдачи — пересчитана (выдано сегодня 4)", same and v2[7] == "4", f"{same}; {v2[7]}")
        oracle(c, s, "дашборд ничего не провёл и не изменил")
    finally:
        s.close()


@case
def pr16_find_card():
    c = "PR16"
    p = new_wms("pr16")
    s = ready(Session(p))
    try:
        order(s, 1, "7", "Кабель UTP", qty="100", supplier="Сеть")
        rr = receive(s, 1, "100", doc="УПД-77", place="5-12-3")
        ei = ords(s, 1)[21]
        n = int(ei[3:]) if ei.startswith("ЕИ-") else 0
        i1 = issue(s, 1, ei, "30", who="Петров")
        i2 = issue(s, 2, ei, "5", who="Сидоров")
        d2 = s.I("IssueDeleteRow", 2)
        put(s, "Возврат", 1, B="1", F="10", H=serial(), I="Петров")
        rt = s.B("ReturnPostRow", 1, module="WmsReturn")
        put(s, "Корректировки", 1, B="Списание", C=ei, G="2", K=serial(), L="брак")
        wo = s.B("AdjustPostRow", 1, module="WmsAdjust")
        put(s, "Корректировки", 2, B="Перемещение", C=ei, J="5-12-4", K=serial(), L="переложили")
        mv = s.B("AdjustPostRow", 2, module="WmsAdjust")
        r, unmod, nj = unmodified_call(s, "FindRun", str(n), module="WmsFind")
        v = page(s, "Поиск", 6, 20, 0, 12)
        moves = [row[1:9] for row in v[6:12]]
        R.add(c, "«Поиск» номера ЕИ → карточка: приход по заказу +100, выдача −30, удалённая выдача серым без влияния, возврат +10, списание −2, "
                 "перемещение без влияния; бегущий остаток 100 → 70 → 70 → 80 → 78 → 78; «✓ Остаток сходится»; книга не изменена",
              all(x.startswith("OK") for x in (rr, i1, i2, d2, rt, wo, mv, r)) and moves[0][1].startswith("Приход по заказу") and moves[0][4] == "100"
              and moves[1][5] == "30" and moves[2][5] == "" and "удалена" in moves[2][7] and moves[3][4] == "10" and moves[4][5] == "2"
              and [m[6] for m in moves] == ["100", "70", "70", "80", "78", "78"] and v[3][1].startswith("✓") and unmod and nj,
              f"{r}; {moves}; {v[3][1][:80]}")
        R.add(c, "шапка карточки: ЕИ, наименование, остаток, место (после перемещения), откуда", v[0][1].startswith(ei) and "Кабель UTP" in v[0][1]
              and "Остаток 78" in v[1][1] and "5-12-4" in v[1][1], f"{v[0][1]}; {v[1][1]}")
        for q in ("ЕИ-" + str(n), f"еи {n}", ei):
            s.B("FindRun", q, module="WmsFind")
        same = page(s, "Поиск", 6, 6)[0][1].startswith(ei)
        R.add(c, "ЕИ понимается в любой записи: «ЕИ-427», «еи 427», «ЕИ-00000427», «427»", same, "")
        rh = s.B("FindRun", "петров", module="WmsFind")
        R.add(c, "имя получателя → «На руках у …»", rh.startswith("OK:на руках у Петров"), rh)
        r2 = s.B("FindRun", "петров*", module="WmsFind")
        v2 = page(s, "Поиск", 6, 20, 0, 12)
        sheets_found = sorted({row[1] for row in v2[6:] if row[1]})
        R.add(c, "поиск слова (без учёта регистра) по всем листам: выдачи и возврат Петрова, получатель; строка — номер строки листа",
              r2.startswith("OK") and "Выдачи" in sheets_found and "Возврат" in sheets_found and v2[0][1].startswith("Найдено «петров»"), f"{r2}; {sheets_found}")
        r3 = s.B("FindRun", "5-12-4", module="WmsFind")
        v3 = page(s, "Поиск", 6, 14, 0, 12)
        R.add(c, "поиск места: ЕИ на месте и итог «в «Наличие» по найденному»", r3.startswith("OK") and any(row[3] == ei for row in v3[6:]) and "ЕИ 1" in v3[1][1],
              f"{r3}; {v3[1][1]}")
        r4 = s.B("FindRun", "нет-такого-слова", module="WmsFind")
        R.add(c, "ничего не найдено — понятная строка, таблица пуста", r4.startswith("OK:найдено 0") and page(s, "Поиск", 6, 6)[0][1].startswith("Ничего не найдено"), r4)
        # «Перейти к строке» and the double click
        s.B("FindRun", "петров*", module="WmsFind")
        s.goto("Поиск", "$B$13")
        s.B("BtnFindGo", module="WmsFind")
        R.add(c, "«Перейти к строке» открывает лист найденной строки", s.active_sheet() in ("Выдачи", "Возврат", "Получатели") and s.U("TestUiLastMessage").startswith("OK:"),
              f"{s.active_sheet()}; {s.U('TestUiLastMessage')}")
        s.B("FindRun", "кабель utp", module="WmsFind")
        cell = s.doc.Sheets.getByName("Поиск").getCellByPosition(1, 12)
        dc = s.B("OnFindDoubleClick", cell, module="WmsFind")
        R.add(c, "двойной щелчок по найденной строке с ЕИ — карточка ЕИ (ячейка не редактируется)", dc is True and page(s, "Поиск", 6, 6)[0][1].startswith(ei),
              f"{dc}; {page(s, 'Поиск', 6, 6)[0][1]}")
        r5 = s.B("FindRun", "1", module="WmsFind")
        R.add(c, "ЕИ начальных данных книги без прихода — карточка без ложной ошибки «не сходится»", "Первичного прихода в книге нет" in page(s, "Поиск", 9, 9)[0][1],
              page(s, "Поиск", 9, 9)[0][1])
        oracle(c, s, "поиск и карточка ничего не провели")
    finally:
        s.close()


def topology(s, rows):
    """a small layout on «Топология»: rows {ряд: (секций, ярусов)}; between the rows an empty line and a wall"""
    t = s.doc.Sheets.getByName("Топология")
    t.getCellByPosition(0, 0).setString("")
    r = 1
    for R_, (S, L) in rows.items():
        t.getCellByPosition(0, r).setString(str(R_))
        for lv in range(L, 0, -1):
            for sec in range(1, S + 1):
                t.getCellByPosition(sec, r).setString(f"{R_}-{sec}-{lv}")
            r += 1
        for sec in range(1, S + 1):
            t.getCellByPosition(sec, r).CellBackColor = 0x000000
        r += 2


@case
def pr17_map():
    c = "PR17"
    p = new_wms("pr17")
    s = ready(Session(p))
    try:
        e0 = s.B("MapBuild", "OCC", module="WmsMap")
        R.add(c, "топология не задана — понятная причина (вставить или загрузить), без ошибки", e0.startswith("ERR:") and "Загрузить топологию" in e0
              and "ряд-секция-ярус" in e0, e0)
        topology(s, {1: (6, 3), 2: (6, 3)})
        b0 = s.B("MapBuild", "OCC", module="WmsMap")
        off0 = int(b0.split("не на карте ")[1].split(",")[0]) if "не на карте " in b0 else -1
        places = {}
        for k, (pl_text, f) in enumerate([("1-2-3", "10"), ("01.02.03", "5"), ("Р1 С2 Я3", "4"), ("2/6/1", "3"), ("1-5", "2"), ("Офис", "1"), ("2-6-1", "8")]):
            order(s, 1 + k, "1", f"Товар {k}", qty=f)
            places[pl_text] = receive(s, 1 + k, f, place=pl_text)
        r1, unmod, nj = unmodified_call(s, "MapBuild", "OCC", module="WmsMap")
        R.add(c, "карта по топологии: 36 мест; «1-2-3», «01.02.03», «Р1 С2 Я3» — одно место (3 ЕИ), «2/6/1» и «2-6-1» — другое; «1-5» без яруса и «Офис» "
                 "— не на карте (+2 к ЕИ начальных данных книги); книга не изменена, журнал не тронут",
              all(x.startswith("OK") for x in places.values()) and r1.startswith("OK:мест 36, занято 2, свободно 34") and off0 >= 0
              and f"не на карте {off0 + 2}," in r1 and unmod and nj, f"{b0}; {r1}; {places}")
        rc = s.B("TestMapPlace", "1-2-3", module="WmsMap").split(":")
        info = s.B("TestMapClick", int(rc[0]), int(rc[1]), module="WmsMap")
        R.add(c, "щелчок по месту 1-2-3 — что на нём лежит: три ЕИ с остатками", info.startswith("OK:1-2-3:") and info.count("ЕИ-") == 3 and "10 шт" in info, info)
        rc2 = s.B("TestMapPlace", "1-1-1", module="WmsMap").split(":")
        info2 = s.B("TestMapClick", int(rc2[0]), int(rc2[1]), module="WmsMap")
        R.add(c, "пустое место — «пусто — место свободно»", "свободно" in info2, info2)
        mp = s.doc.Sheets.getByName("Карта склада")
        grid = mp.getCellByPosition(600 + int(rc[1]), int(rc[0])).getValue()
        R.add(c, "место на карте — текст топологии как есть; код цвета слоя «Занятость» для 3 ЕИ — 3", grid == 3.0
              and mp.getCellByPosition(int(rc[1]), int(rc[0])).getString() == "1-2-3", f"{grid}")
        res = {}
        for layer in ("ISS", "STALE", "ATT", "FREE", "OCC"):
            res[layer] = s.B("MapBuild", layer, module="WmsMap")
        R.add(c, "слои переключаются (выдачи 30 дней, без движения, внимание, свободные, занятость) — каждый с легендой", all(x.startswith("OK") for x in res.values())
              and "■■" in mp.getCellByPosition(0, 5).getString(), str({k: v[:30] for k, v in res.items()}))
        s.B("BtnMapOff", module="WmsMap")
        off = page(s, "Поиск", 6, 12 + off0 + 3, 0, 12)
        mine = {row[4]: row[7] for row in off[6:] if row[4] in ("Товар 4", "Товар 5")}
        R.add(c, "«Не на карте» — список на «Поиск»: «Товар 4» (место «1-5», без яруса) и «Товар 5» («Офис», такого места нет), с подсказкой, как поправить",
              off[0][1].startswith(f"ЕИ с остатком не на карте: {off0 + 2}") and mine.get("Товар 4") == "1-5" and mine.get("Товар 5") == "Офис"
              and "Перемещение" in off[1][1], f"{off[0][1]}; {mine}")
        ei = ords(s, 6)[21]
        put(s, "Корректировки", 1, B="Перемещение", C=ei, J="1-1-1", K=serial(), L="на место")
        mv = s.B("AdjustPostRow", 1, module="WmsAdjust")
        r2 = s.B("MapBuild", "OCC", module="WmsMap")
        R.add(c, "после перемещения на место карты — ЕИ на карте (занято 3, не на карте на один меньше)", mv.startswith("OK") and "занято 3" in r2
              and f"не на карте {off0 + 1}," in r2, f"{mv}; {r2}")
        keys = [s.B("PlaceKey", x, module="WmsView") for x in ("5.12.3", "05/12/03", "Р5 С12 Я3", "R5-S12-L3", "A-01-1", "B-01-1", "5-12", "Офис", "12")]
        R.add(c, "место читается как его пишут: 5.12.3, 05/12/03, «Р5 С12 Я3», R5-S12-L3 → 5-12-3; буквы ряда сохраняются (A-1-1 ≠ B-1-1; латинские A, B — как русские А, В); 5-12 — без яруса",
              keys == ["5-12-3"] * 4 + ["А-1-1", "В-1-1", "5-12", "", ""], str(keys))
        oracle(c, s, "карта ничего не провела")
    finally:
        s.close()


@case
def pr18_topology_file():
    """«Загрузить топологию…»: the sheet of a file with the most places becomes «Топология» (cells, colours, sizes, merged cells)"""
    c = "PR18"
    p = new_wms("pr18")
    f = os.path.join(os.path.dirname(p), "topo.ods")
    s0 = Session(new_wms("pr18_src"), macros=0)
    try:
        doc = s0.o.new_calc()
        sh = doc.Sheets.getByIndex(0)
        sh.Name = "Схема"
        for sec in range(1, 9):
            for lv in (1, 2):
                cell = sh.getCellByPosition(sec, 3 - lv)
                cell.setString(f"3-{sec}-{lv}")
                cell.CellBackColor = 0xDDEEFF
        sh.getCellRangeByPosition(1, 4, 8, 4).CellBackColor = 0x000000
        sh.getCellByPosition(0, 1).setString("3")
        sh.getCellRangeByPosition(0, 1, 0, 2).merge(True)
        sh.getColumns().getByIndex(3).Width = 1500
        doc.Sheets.insertNewByName("Прочее", 1)
        doc.Sheets.getByName("Прочее").getCellByPosition(0, 0).setString("1-1-1")
        doc.storeToURL(__import__("uno").systemPathToFileUrl(f), ())
        doc.close(True)
        empty = s0.o.new_calc()
        empty.Sheets.getByIndex(0).getCellByPosition(0, 0).setString("просто таблица")
        empty.storeToURL(__import__("uno").systemPathToFileUrl(os.path.join(os.path.dirname(p), "empty.ods")), ())
        empty.close(True)
    finally:
        s0.close()
    s = ready(Session(p))
    try:
        order = lambda: [s.doc.Sheets.getByIndex(i).Name for i in range(s.doc.Sheets.getCount())]
        before = order()
        r = s.B("TopoLoad", __import__("uno").systemPathToFileUrl(f), module="WmsMap")
        t = s.doc.Sheets.getByName("Топология")
        R.add(c, "из файла взят лист с местами (16 мест, лист «Схема»), он стал «Топология» на своём месте, книга защищена как была",
              r.startswith("OK:мест 16") and "«Схема»" in r and t.getCellByPosition(3, 1).getString() == "3-3-2"
              and abs(t.getColumns().getByIndex(3).Width - 1500) <= 20 and s.doc.isProtected() and order() == before,
              f"{r}; {t.getCellByPosition(3, 1).getString()}; {t.getColumns().getByIndex(3).Width}; {s.doc.isProtected()}; {order()}")
        m = s.B("MapBuild", "OCC", module="WmsMap")
        R.add(c, "карта по загруженной топологии: 16 мест", m.startswith("OK:мест 16"), m)
        bad = os.path.join(os.path.dirname(p), "empty.ods")
        r2 = s.B("TopoLoad", __import__("uno").systemPathToFileUrl(bad), module="WmsMap")
        R.add(c, "файл без мест «ряд-секция-ярус» — отказ «это не топология склада», прежняя «Топология» на месте",
              r2.startswith("ERR:") and "не топология" in r2 and s.doc.Sheets.getByName("Топология").getCellByPosition(3, 1).getString() == "3-3-2", r2)
        oracle(c, s, "загрузка топологии ничего не провела")
    finally:
        s.close()


@case
def pr19_where_put():
    """«Куда положить?»: where the goods already lie, then the free places of the map nearest to them"""
    c = "PR19"
    p = new_wms("pr19")
    s = ready(Session(p))
    try:
        topology(s, {1: (6, 3), 2: (6, 3)})
        order(s, 1, "1", "Перчатки", qty="10")
        r1 = receive(s, 1, "10", place="1-3-2")
        order(s, 2, "2", "Перчатки", qty="5")
        sug = s.B("WhereSuggest", "Перчатки", "", module="WmsMap").split("\n")
        R.add(c, "подсказка: сначала место, где уже лежит этот товар (1-3-2), затем свободные рядом — того же ряда и секции (1-3-1, 1-3-3)",
              r1.startswith("OK") and sug[0].startswith("1-3-2|") and "уже лежит" in sug[0] and sug[1].split("|")[0] in ("1-3-1", "1-3-3")
              and sug[2].split("|")[0] in ("1-3-1", "1-3-3") and len(sug) == 8, str(sug[:4]))
        s.goto("Заказы", "$B$3")
        s.B("TestPrimeChoice", 1, module="WmsPrimeUi")
        res = s.B("WherePut", module="WmsMap")
        R.add(c, "выбор второго варианта — место записано в «Место хранения» непроведённой строки (это ввод, не операция: журнал не тронут)",
              res.startswith("OK:место 1-3-") and ords(s, 2)[20] == sug[1].split("|")[0], f"{res}; {ords(s, 2)[20]}")
        rr = receive(s, 2, "5", place=ords(s, 2)[20])
        s.goto("Заказы", "$B$2")
        s.B("TestPrimeChoice", 0, module="WmsPrimeUi")
        res2 = s.B("WherePut", module="WmsMap")
        R.add(c, "для проведённой строки место не пишется — подсказка «Корректировки» → «Перемещение»", rr.startswith("OK") and "Перемещение" in res2
              and ords(s, 1)[20] == "1-3-2", res2)
        s.goto("Заказы", "$B$2")
        s.B("TestPrimeChoice", -1, module="WmsPrimeUi")
        res3 = s.B("WherePut", module="WmsMap")
        s.goto("Выдачи", "$B$2")
        res4 = s.B("WherePut", module="WmsMap")
        R.add(c, "окно закрыто — ничего не записано; на другом листе — понятный отказ", res3.startswith("SKIP") and res4.startswith("ERR:") and "«Заказы»" in res4,
              f"{res3}; {res4}")
        oracle(c, s, "«Куда положить?» ничего не провёл")
    finally:
        s.close()


@case
def pr20_hands():
    """«На руках»: the issues of a recipient minus the returns of each of them"""
    c = "PR20"
    p = new_wms("pr20")
    s = ready(Session(p))
    try:
        put(s, "Получатели", 30, A="пв", B="Петров Василий")
        a = [issue(s, 1, "ЕИ-00000001", "5", who="пв"), issue(s, 2, "ЕИ-00000002", "2", who="Петров Василий"), issue(s, 3, "ЕИ-00000004", "1", who="егр")]
        put(s, "Возврат", 1, B="1", F="5", H=serial(), I="Петров Василий")
        b = s.B("ReturnPostRow", 1, module="WmsReturn")
        r, unmod, nj = unmodified_call(s, "FindRun", "пв", module="WmsFind")
        v = page(s, "Поиск", 6, 15, 0, 12)
        R.add(c, "«пв» (сокращение) → «На руках у Петров Василий»: 1 выдача не возвращена (ЕИ-00000002, 2), возвращённая ниже; книга не изменена",
              all(x.startswith("OK") for x in a + [b]) and r.startswith("OK:на руках у Петров Василий: 1 из 2") and v[6][3] == "ЕИ-00000002"
              and v[6][7] == "2" and v[7][3] == "ЕИ-00000001" and v[7][7] == "0" and v[7][6] == "5" and unmod and nj, f"{r}; {a}; {b}; {v[6][1:9]}; {v[7][1:9]}")
        r2 = s.B("FindRun", "пв*", module="WmsFind")
        R.add(c, "«пв*» — поиск текста по листам, а не «на руках»", r2.startswith("OK:найдено") and page(s, "Поиск", 6, 6)[0][1].startswith("Найдено «пв»"), r2)
        oracle(c, s, "«на руках» ничего не провёл")
    finally:
        s.close()


@case
def pr21_pick_ei():
    """«Подобрать ЕИ»: the EIs with a balance by a name, an article or a place, the oldest first, written into the row"""
    c = "PR21"
    p = new_wms("pr21")
    s = ready(Session(p))
    try:
        order(s, 1, "1", "Перчатки нитриловые", qty="10", unit="пар")
        rc = receive(s, 1, "10", place="D-02")
        cand = s.B("PickCandidates", "перчатки", module="WmsFind").split("\n")
        new_ei = cand[-1].split("|")[0]
        R.add(c, "по части наименования — ЕИ с остатком, старый первым (ЕИ-00000004, затем новый приход); в строке — товар, остаток и место",
              rc.startswith("OK") and len(cand) == 2 and cand[0].startswith("ЕИ-00000004|ЕИ-00000004 · Перчатки нитриловые · 12 пар · D-01")
              and new_ei != "ЕИ-00000004" and cand[1].endswith("· 10 пар · D-02"), f"{rc}; {cand}")
        many = s.B("PickCandidates", "товар", module="WmsFind").split("\n")
        by_art = s.B("PickCandidates", "pn-9", module="WmsFind")
        by_place = s.B("PickCandidates", "D-02", module="WmsFind")
        R.add(c, "по артикулу (без учёта регистра) и по месту; длинный список — не больше 15 строк, старые первыми",
              by_art.startswith("ЕИ-00000004|") and "\n" not in by_art and by_place.startswith(new_ei + "|") and len(many) == 15
              and many[0].startswith("ЕИ-00000007|") and many[1].startswith("ЕИ-00000008|"), f"{by_art}; {by_place}; {many[:2]}; {len(many)}")
        zero = s.B("PickCandidates", "лампа", module="WmsFind")
        empty = s.B("PickCandidates", "", module="WmsFind")
        R.add(c, "товар без остатка и пустой запрос — понятный отказ, а не пустой список",
              zero.startswith("ERR:ЕИ с остатком по «лампа» нет") and empty.startswith("ERR:ничего не вписано"), f"{zero}; {empty}")
        j0 = njournal(s)
        s.goto("Выдачи", "$B$2")
        s.B("TestPrimeForm", "перчатки", module="WmsPrimeUi")
        s.B("TestPrimeChoice", 1, module="WmsPrimeUi")
        res = s.B("PickEI", module="WmsFind")
        cell = s.doc.Sheets.getByName("Выдачи").getCellByPosition(11, 1).getString()
        R.add(c, "«Выдачи»: выбран второй ЕИ — записан в «ЕИ» непроведённой строки 2 (это ввод, не операция: журнал не тронут)",
              res == f"OK:{new_ei} записан в строку 2" and cell == new_ei and njournal(s) == j0, f"{res}; {cell}")
        put(s, "Выдачи", 1, E="3", H=serial(), I="егр")
        iss = s.I("IssuePostRow", 1)
        s.goto("Выдачи", "$B$2")
        res2 = s.B("PickEI", module="WmsFind")
        R.add(c, "строка проводится с подобранным ЕИ; в проведённой строке ЕИ не подбирается — подсказка «Исправить»",
              iss.startswith("OK") and res2.startswith("ERR:строка 2 уже проведена") and "«Исправить»" in res2, f"{iss}; {res2}")
        s.goto("Корректировки", "$B$2")
        s.B("TestPrimeForm", "PN-9", module="WmsPrimeUi")
        s.B("TestPrimeChoice", 0, module="WmsPrimeUi")
        res3 = s.B("PickEI", module="WmsFind")
        adj = s.doc.Sheets.getByName("Корректировки").getCellByPosition(2, 1).getString()
        R.add(c, "«Корректировки»: ЕИ по артикулу записан в «ЕИ» строки 2", res3 == "OK:ЕИ-00000004 записан в строку 2" and adj == "ЕИ-00000004",
              f"{res3}; {adj}")
        s.goto("Выдачи", "$B$3")
        closed = s.B("PickEI", module="WmsFind")
        s.B("TestPrimeForm", "перчатки", module="WmsPrimeUi")
        s.B("TestPrimeChoice", -1, module="WmsPrimeUi")
        nochoice = s.B("PickEI", module="WmsFind")
        s.goto("Заказы", "$B$3")
        other = s.B("PickEI", module="WmsFind")
        R.add(c, "окно закрыто или ЕИ не выбран — строка не тронута; на другом листе — понятный отказ",
              closed.startswith("SKIP") and nochoice.startswith("SKIP") and s.doc.Sheets.getByName("Выдачи").getCellByPosition(11, 2).getString() == ""
              and other.startswith("ERR:") and "«Выдачи»" in other and "«Корректировки»" in other, f"{closed}; {nochoice}; {other}")
        forms = {n: [s.doc.Sheets.getByName(n).DrawPage.Forms.getByIndex(0).getByIndex(i).Name
                     for i in range(s.doc.Sheets.getByName(n).DrawPage.Forms.getByIndex(0).Count)] for n in ("Выдачи", "Корректировки")}
        R.add(c, "кнопка «Подобрать ЕИ» есть на «Выдачи» и «Корректировки»", all("btnPickEI" in v for v in forms.values()), str(forms))
        oracle(c, s, "«Подобрать ЕИ» ничего не провёл, кроме выдачи строки 2")
    finally:
        s.close()


def ei_row(canon):
    """the row of «Наличие» (= the number) of an EI"""
    return int(str(canon)[3:])


def journal_last(s, typ):
    """the fields of the last journal line of type typ"""
    entries, _ = jo.read_journal(s.jdir)
    for e in reversed(entries):
        if e["type"] == typ:
            return e["fields"]
    return {}


@case
def pr22_parts():
    """D-099: учёт по документу, выдача частями — фасовка ЕИ («в 1 шт — 4 кан», «1 шт = 6 м»), «1 кан», «1,5 м», 25%, 0,25;
    возврат, списание, инвентаризация и «Исправить» частями; без «хвостов»"""
    c = "PR22"
    p = new_wms("pr22")
    s = ready(Session(p))
    try:
        order(s, 1, "1", "Незамерзайка -30 (4 канистры по 5 л)", qty="1")
        rc1 = receive(s, 1, "1", place="1-1-1")
        ant = ords(s, 1)[21]
        guess = [s.B("TestPackGuess", x, module="WmsPack") for x in ("Незамерзайка -30 (4 канистры по 5 л)", "Профиль 40х20 L=6000 мм",
                                                                    "Профиль 40х20, 6 м", "Масло 3 бут.", "Болт М8")]
        R.add(c, "фасовка угадывается по наименованию: 4 канистры → 4 кан, L=6000 мм → 6 м, 6 м → 6 м, 3 бут. → 3 бут; «Болт М8» — ничего",
              guess == ["4|кан", "6|м", "6|м", "3|бут", "0|"], str(guess))
        s.goto("Наличие", f"$A${ei_row(ant) + 1}")
        s.B("TestPrimeForm", "", module="WmsPrimeUi")
        pk = s.B("PackEI", module="WmsPack")
        prow = list(s.doc.Sheets.getByName("_PACK").getCellRangeByPosition(0, ei_row(ant), 3, ei_row(ant)).getDataArray()[0])
        R.add(c, "«Фасовка ЕИ» на строке «Наличие»: предложено по наименованию, записано операцией EI_PACK (в журнале), остаток не тронут",
              rc1.startswith("OK") and pk.startswith("OK:" + ant + ": в 1 шт — 4 кан") and prow[:3] == [ant, "кан", 4.0] and s.stock(ei_row(ant)) == 1.0
              and journal_last(s, "EI_PACK").get("N") == "4", f"{pk}; {prow}")
        put(s, "Выдачи", 1, L=ant, E="1 кан", H=serial(), I="егр")
        s.I("PreviewRow", 1, False, False)
        prev = s.iss(1)[17]
        a1 = s.I("IssuePostRow", 1)
        j1 = journal_last(s, "ISSUE")
        R.add(c, "«1 кан» в «Количество»: до проведения в «Контроле» — «К выдаче: 1 кан = 0,25 шт · останется 0,75 шт (3 кан)»; проведено 0,25 шт (учёт по документу), в журнале — 1 кан",
              prev == "К выдаче: 1 кан = 0,25 шт · останется 0,75 шт (3 кан)" and a1.startswith("OK") and s.iss(1)[4] == 0.25 and s.stock(ei_row(ant)) == 0.75
              and j1.get("QTY") == "0.25" and j1.get("PARTS") == "1" and j1.get("PART_UNIT") == "кан", f"{prev}; {a1}; {s.iss(1)[4]}; {j1}")
        put(s, "Выдачи", 2, L=ant, E="2 канистры", H=serial(), I="егр")
        a2 = s.I("IssuePostRow", 2)
        put(s, "Выдачи", 3, L=ant, E="1 к", H=serial(), I="егр")
        a3 = s.I("IssuePostRow", 3)
        put(s, "Выдачи", 3, E="1 бут")
        a3b = s.I("IssuePostRow", 3)
        put(s, "Выдачи", 3, E="2 кан")
        a3c = s.I("IssuePostRow", 3)
        put(s, "Выдачи", 3, E="1 кан")
        a3d = s.I("IssuePostRow", 3)
        R.add(c, "«2 канистры» — 0,5 шт; непонятная единица («1 к», «1 бут») и больше остатка («2 кан» при 1) — понятный отказ; последняя канистра закрывает остаток ровно в 0",
              a2.startswith("OK") and a3.startswith("ERR:") and a3b.startswith("ERR:") and "«кан»" in a3b and "«шт»" in a3b
              and a3c.startswith("ERR:") and "остаток 1 кан" in a3c and a3d.startswith("OK") and s.stock(ei_row(ant)) == 0.0,
              f"{a2}; {a3}; {a3b}; {a3c}; {a3d}; {s.stock(ei_row(ant))}")
        # three parts that 0,001 cannot hold exactly: 1/3 each, no remainder after the third
        order(s, 2, "2", "Масло 3 бут.", qty="1")
        receive(s, 2, "1", place="1-1-2")
        oil = ords(s, 2)[21]
        ps = s.B("PackSet", oil, "3", "бут", "", module="WmsPack")
        bals = []
        for r in (4, 5, 6):
            put(s, "Выдачи", r, L=oil, E="1 бут", H=serial(), I="егр")
            bals.append((s.I("IssuePostRow", r)[:2], s.iss(r)[4], s.stock(ei_row(oil))))
        R.add(c, "3 бут в 1 шт: каждая бутылка — 0,333 или 0,334 шт, после третьей остаток ровно 0 (без «хвоста» 0,001)",
              ps.startswith("OK") and [b[0] for b in bals] == ["OK"] * 3 and bals[2][2] == 0.0 and abs(sum(b[1] for b in bals) - 1.0) < 1e-9, str(bals))
        # metal: a profile cut into parts
        order(s, 3, "3", "Профиль 40х20, 6 м", qty="2", unit="шт")
        receive(s, 3, "2", place="1-2-1")
        prof = ords(s, 3)[21]
        s.B("PackSet", prof, "6", "м", "профиль по 6 м", module="WmsPack")
        put(s, "Выдачи", 7, L=prof, E="1,5 м", H=serial(), I="егр")
        m1 = s.I("IssuePostRow", 7)
        put(s, "Выдачи", 8, L=prof, E="0,5", H=serial(), I="егр")
        m2 = s.I("IssuePostRow", 8)
        put(s, "Выдачи", 9, L=prof, E="25%", H=serial(), I="егр")
        s.I("PreviewRow", 9, False, False)
        prev9 = s.iss(9)[17]
        m3 = s.I("IssuePostRow", 9)
        j3 = journal_last(s, "ISSUE")
        R.add(c, "металл: «1,5 м» профиля 6 м — 0,25 шт; 0,5 шт числом; «25%» — от пришедших 2 шт = 0,5 шт (4,5 м остатка видно в подсказке)",
              m1.startswith("OK") and s.iss(7)[4] == 0.25 and m2.startswith("OK") and m3.startswith("OK") and s.iss(9)[4] == 0.5
              and prev9.startswith("К выдаче: 25% от пришедших 2 шт = 0,5 шт (3 м)") and j3.get("PCT") == "25" and s.stock(ei_row(prof)) == 0.75,
              f"{m1}; {m2}; {prev9}; {m3}; {j3}; {s.stock(ei_row(prof))}")
        put(s, "Выдачи", 10, L=prof, F=50.0, H=serial(), I="егр")
        f1 = s.I("IssuePostRow", 10)
        put(s, "Выдачи", 10, F=10.0)
        f2 = s.I("IssuePostRow", 10)
        put(s, "Выдачи", 11, L=prof, E="5 м", F=10.0, H=serial(), I="егр")
        f3 = s.I("IssuePostRow", 11)
        R.add(c, "«Количество в %» (F): 50% (1 шт) больше остатка — отказ с остатком в метрах; 10% = 0,2 шт; E и F вместе — «заполните что-то одно»",
              f1.startswith("ERR:") and "4,5 м" in f1 and f2.startswith("OK") and s.iss(10)[4] == 0.2 and f3.startswith("ERR:") and "что-то одно" in f3,
              f"{f1}; {f2}; {f3}")
        # the receipt: «по документу 1 шт, в наименовании 4 канистры, факт 4» — asked: by the document and the packing, or as typed
        order(s, 5, "5", "Незамерзайка зимняя, 4 канистры", qty="1")
        put(s, "Заказы", 5, G=1.0)
        rk1 = receive(s, 5, "4", place="1-3-1")
        w1 = ords(s, 5)[21]
        pk1 = list(s.doc.Sheets.getByName("_PACK").getCellRangeByPosition(0, ei_row(w1), 2, ei_row(w1)).getDataArray()[0])
        order(s, 6, "6", "Незамерзайка зимняя, 4 канистры", qty="1")
        put(s, "Заказы", 6, G=1.0)
        s.U("TestUiAuto", 2)
        rk2 = receive(s, 6, "4", place="1-3-1")
        s.U("TestUiAuto", 1)
        w2 = ords(s, 6)[21]
        R.add(c, "приход «по документу 1 шт, факт 4, в наименовании 4 канистры» — вопрос: «Да» — проведено 1 шт и фасовка «в 1 шт — 4 кан»; "
                 "«Нет» — 4 шт как есть, без фасовки",
              rk1.startswith("OK") and s.stock(ei_row(w1)) == 1.0 and ords(s, 5)[5] == 1.0 and pk1 == [w1, "кан", 4.0]
              and rk2.startswith("OK") and s.stock(ei_row(w2)) == 4.0 and s.B("PackLine", ei_row(w2), "шт", module="WmsPack") == "",
              f"{rk1}; {s.stock(ei_row(w1))}; {pk1}; {rk2}; {s.stock(ei_row(w2))}")
        # a whole number without a unit for an EI with a packing: asked — «1» is 1 шт (4 бан) or 1 бан
        order(s, 4, "4", "Краска (4 банки)", qty="2")
        receive(s, 4, "2", place="1-2-2")
        paint = ords(s, 4)[21]
        s.B("PackSet", paint, "4", "бан", "", module="WmsPack")
        s.U("TestUiAuto", 2)
        put(s, "Выдачи", 12, L=paint, E=1.0, H=serial(), I="егр")
        q1 = s.I("IssuePostRow", 12)
        s.U("TestUiAuto", 1)
        put(s, "Выдачи", 13, L=paint, E=1.0, H=serial(), I="егр")
        q2 = s.I("IssuePostRow", 13)
        R.add(c, "у ЕИ с фасовкой целое число без единицы — вопрос «1 шт (4 бан) или 1 бан?»: «Нет» — 1 бан (0,25 шт), «Да» — 1 шт",
              q1.startswith("OK") and s.iss(12)[4] == 0.25 and q2.startswith("OK") and s.iss(13)[4] == 1.0 and s.stock(ei_row(paint)) == 0.75,
              f"{q1}; {s.iss(12)[4]}; {q2}; {s.iss(13)[4]}; {s.stock(ei_row(paint))}")
        # a packing that is not a whole number: a pipe of 1 шт came by weight, 23,5 кг — «5 кг», then the rest closes exactly
        order(s, 7, "7", "Труба 57х3,5", qty="1")
        receive(s, 7, "1", place="1-3-2")
        pipe = ords(s, 7)[21]
        pw = s.B("PackSet", pipe, "23,5", "кг", "", module="WmsPack")
        put(s, "Выдачи", 14, L=pipe, E="5 кг", H=serial(), I="егр")
        w1 = s.I("IssuePostRow", 14)
        b1 = s.stock(ei_row(pipe))
        put(s, "Выдачи", 15, L=pipe, E="18,5 кг", H=serial(), I="егр")
        w2 = s.I("IssuePostRow", 15)
        R.add(c, "дробная фасовка «в 1 шт — 23,5 кг»: «5 кг» — 0,213 шт, остаток 18,5 кг; «18,5 кг» закрывает остаток ровно в 0",
              pw.startswith("OK") and w1.startswith("OK") and s.iss(14)[4] == 0.213 and b1 == 0.787 and w2.startswith("OK")
              and s.stock(ei_row(pipe)) == 0.0, f"{pw}; {w1}; {s.iss(14)[4]}; {b1}; {w2}; {s.stock(ei_row(pipe))}")
        # «Исправить» in parts, a return, a write-off and a count in parts
        fx = s.I("IssueFixRow", 7, prof, "2 м", "егр", serial())
        R.add(c, "«Исправить» выдачи частями: «2 м» вместо 1,5 м — 0,333 шт, остаток пересчитан", fx.startswith("OK") and s.iss(7)[4] == 0.333
              and abs(s.stock(ei_row(prof)) - 0.467) < 1e-9, f"{fx}; {s.iss(7)[4]}; {s.stock(ei_row(prof))}")
        put(s, "Возврат", 1, B="1", F="1 кан", H=serial(), I="егр")
        rt = s.B("ReturnPostRow", 1, module="WmsReturn")
        put(s, "Корректировки", 1, B="Списание", C=ant, G="1 кан", K=serial(), L="разлили")
        wo = s.B("AdjustPostRow", 1, module="WmsAdjust")
        put(s, "Корректировки", 2, B="Инвентаризация", C=oil, H="2 бут", K=serial(), L="нашли")
        inv = s.B("AdjustPostRow", 2, module="WmsAdjust")
        R.add(c, "возврат «1 кан» по выдаче 1 кан — ровно 0,25 шт (выдача закрыта); списание «1 кан»; инвентаризация «2 бут» — остаток 0,667 шт",
              rt.startswith("OK") and wo.startswith("OK") and inv.startswith("OK") and s.stock(ei_row(ant)) == 0.0 and s.stock(ei_row(oil)) == 0.667
              and s.doc.Sheets.getByName("Возврат").getCellByPosition(5, 1).getValue() == 0.25, f"{rt}; {wo}; {inv}; {s.stock(ei_row(ant))}; {s.stock(ei_row(oil))}")
        e1 = s.B("TestQtyEI", "1 кан", 1, "OUT", 0.0, module="WmsPack")
        e2 = s.B("TestQtyEI", "25%", 1, "OUT", 0.0, module="WmsPack")
        e3 = s.B("TestQtyEI", "2 шт", 1, "OUT", 0.0, module="WmsPack")
        e4 = s.B("TestQtyEI", "1500 мм", ei_row(prof), "OUT", 0.0, module="WmsPack")
        R.add(c, "у ЕИ без фасовки «1 кан» — подсказка «задайте фасовку»; процент без прихода в книге — понятный отказ; «2 шт» — единица документа; "
                 "«1500 мм» при фасовке в метрах — отказ, а не 1500 м", e1.startswith("ERR:") and "Фасовка ЕИ" in e1 and e2.startswith("ERR:")
              and "нет прихода" in e2 and e3.startswith("OK:2|") and e4.startswith("ERR:"), f"{e1}; {e2}; {e3}; {e4}")
        bad = [s.B("PackSet", prof, x, u, "", module="WmsPack") for x, u in (("2000", "мм"), ("4", "4кан"), ("4", "шт"), ("4", ""))]
        off = s.B("PackSet", prof, "", "", "", module="WmsPack")
        prow2 = list(s.doc.Sheets.getByName("_PACK").getCellRangeByPosition(0, ei_row(prof), 3, ei_row(prof)).getDataArray()[0])
        R.add(c, "фасовка с ошибкой не записывается (больше 1000, цифры в названии, та же единица, без названия); пустое число — фасовка снята",
              all(x.startswith("ERR:") for x in bad) and off.startswith("OK") and prow2 == ["", "", "", ""], f"{bad}; {off}; {prow2}")
        cd = s.B("FindRun", oil, module="WmsFind")
        v = page(s, "Поиск", 7, 20, 0, 9)
        notes = [row[8] for row in v if row[2].startswith("Выдача")]
        R.add(c, "карточка ЕИ на «Поиске»: остаток и частями, фасовка «в 1 шт — 3 бут», у каждой выдачи — сколько бутылок",
              cd.startswith("OK") and "(2 бут)" in v[0][1] and "в 1 шт — 3 бут" in v[0][1] and notes[:3] == ["1 бут"] * 3, f"{cd}; {v[0][1]}; {notes}")
        oracle(c, s, "выдачи, возвраты, списания и пересчёт частями; фасовка в _PACK — по журналу")
    finally:
        s.close()


@case
def pr26_parts_everywhere():
    """1.0 (A1): «1 шт или 1 кан?» — в списании, инвентаризации, возврате и «Исправить», «1» больше остатка (это может быть 1 кан),
    процент меньше 1 («0,25» в %), «1 кан 2», «останется …» в «Контроле», % после «Исправить», дата в будущем, получатель с опечаткой"""
    c = "PR26"
    p = new_wms("pr26")
    s = ready(Session(p))
    try:
        order(s, 1, "1", "Незамерзайка (4 канистры)", qty="2")
        receive(s, 1, "2", place="1-1-1")
        ei = ords(s, 1)[21]
        n = ei_row(ei)
        s.B("PackSet", ei, "4", "кан", "", module="WmsPack")
        put(s, "Выдачи", 1, L=ei, E="1 кан", H=serial(), I="егр")
        s.I("PreviewRow", 1, False, False)
        prev = s.iss(1)[17]
        put(s, "Выдачи", 1, E="1")
        s.I("PreviewRow", 1, False, False)
        prev1 = s.iss(1)[17]
        R.add(c, "«Контроль» до проведения: сколько останется частями; у «1» без единицы — как написать канистры",
              prev == "К выдаче: 1 кан = 0,25 шт · останется 1,75 шт (7 кан)" and "если 1 кан — так и напишите «1 кан»" in prev1, f"{prev!r}; {prev1!r}")
        put(s, "Выдачи", 1, E="1 кан 2")
        bad = s.I("IssuePostRow", 1)
        R.add(c, "«1 кан 2» — отказ (не 1 кан молча)", bad.startswith("ERR:") and "лишнее" in bad, bad)
        # write-off «1»: asked — the second button is 1 кан
        s.U("TestUiAuto", 2)
        put(s, "Корректировки", 1, B="Списание", C=ei, G="1", K=serial(), L="разлили")
        wo = s.B("AdjustPostRow", 1, module="WmsAdjust")
        s.U("TestUiAuto", 1)
        shown = s.B("TestPrimeShown", module="WmsPrimeUi")
        R.add(c, "списание «1» у ЕИ с фасовкой — вопрос с кнопками по действию; «Списать 1 кан» — 0,25 шт",
              wo.startswith("OK") and s.doc.Sheets.getByName("Корректировки").getCellByPosition(6, 1).getValue() == 0.25 and s.stock(n) == 1.75
              and "Списать 1 шт — это 4 кан" in shown and "Списать 1 кан — это 0,25 шт" in shown and "Не проводить" in shown, f"{wo}; {shown!r}")
        # a count «3»: 3 шт or 3 кан
        s.U("TestUiAuto", 2)
        put(s, "Корректировки", 2, B="Инвентаризация", C=ei, H="3", K=serial(), L="пересчёт")
        inv = s.B("AdjustPostRow", 2, module="WmsAdjust")
        s.U("TestUiAuto", 1)
        R.add(c, "инвентаризация «3» — вопрос; «Факт 3 кан» — остаток 0,75 шт (а не 3 шт с излишком)", inv.startswith("OK") and s.stock(n) == 0.75,
              f"{inv}; {s.stock(n)}")
        # «1» more than the balance (0,75 шт = 3 кан): only «1 кан» can be meant — asked with that button first
        put(s, "Выдачи", 2, L=ei, E="1", H=serial(), I="егр")
        sh1 = s.I("IssuePostRow", 2)
        shown2 = s.B("TestPrimeShown", module="WmsPrimeUi")
        R.add(c, "«1» больше остатка 0,75 шт — вопрос «Выдать 1 кан?»: выдано 1 кан = 0,25 шт", sh1.startswith("OK") and s.iss(2)[4] == 0.25
              and s.stock(n) == 0.5 and "Выдать 1 кан" in shown2 and "1 шт нельзя" in shown2, f"{sh1}; {s.iss(2)[4]}; {shown2!r}")
        # a return «1» of an issue of 1 кан: only 1 кан can come back
        put(s, "Возврат", 1, B=str(int(s.iss(2)[0])), F="1", H=serial(), I="егр")
        rt = s.B("ReturnPostRow", 1, module="WmsReturn")
        R.add(c, "возврат «1» по выдаче 1 кан — «Вернуть 1 кан»: ровно 0,25 шт", rt.startswith("OK")
              and s.doc.Sheets.getByName("Возврат").getCellByPosition(5, 1).getValue() == 0.25 and s.stock(n) == 0.75, f"{rt}; {s.stock(n)}")
        # a percent below 1: «0,25» in F — 0,25 % or 25 %
        s.U("TestUiAuto", 2)
        put(s, "Выдачи", 3, L=ei, F=0.25, H=serial(), I="егр")
        pc = s.I("IssuePostRow", 3)
        s.U("TestUiAuto", 1)
        shown3 = s.B("TestPrimeShown", module="WmsPrimeUi")
        R.add(c, "«0,25» в «%» — вопрос «0,25% или 25%?»: выбрано 25% от пришедших 2 шт = 0,5 шт", pc.startswith("OK") and s.iss(3)[4] == 0.5
              and "меньше одного процента" in shown3 and "Выдать 25% — это 0,5 шт" in shown3, f"{pc}; {s.iss(3)[4]}; {shown3!r}")
        # «Исправить» of that row: the quantity in parts, F tells the new quantity
        fx = s.I("IssueFixRow", 3, ei, "1 кан", "егр", serial())
        f_after = s.doc.Sheets.getByName("Выдачи").getCellByPosition(5, 3).getString()
        fx2 = s.I("IssueFixRow", 3, ei, "10%", "егр", serial())
        f_after2 = s.doc.Sheets.getByName("Выдачи").getCellByPosition(5, 3).getString()
        inp = s.B("QtyForInput", n, 0.25, module="WmsPack")
        R.add(c, "«Исправить» на «1 кан» — в «%» (F) не остаётся старый процент; на «10%» — в F 10%; окно «Исправить» показывает «1 кан», а не 0,25",
              fx.startswith("OK") and f_after == "" and fx2.startswith("OK") and f_after2 == "10%" and s.iss(3)[4] == 0.2 and inp == "1 кан",
              f"{fx}; {f_after!r}; {fx2}; {f_after2!r}; {inp}")
        # the date of a movement in the future: asked; «Не проводить» leaves the row as it is
        s.U("TestUiAuto", 2)
        put(s, "Выдачи", 4, L=ei, E="1 кан", H=serial(30), I="егр")
        fut = s.I("IssuePostRow", 4)
        s.U("TestUiAuto", 1)
        fut2 = s.I("IssuePostRow", 4)
        R.add(c, "дата выдачи в будущем — вопрос: «Не проводить» — строка не проведена, «Провести с датой …» — проведена",
              fut.startswith("SKIP:") and "дату" in fut and fut2.startswith("OK"), f"{fut}; {fut2}")
        # a recipient with a typo: the name of the directory offered
        put(s, "Получатели", 5, A="пи", B="Петров Иван")
        s.I("OnRecipientsChange")
        s.B("TestPrimeChoice", 0, module="WmsPrimeUi")
        put(s, "Выдачи", 5, L=ei, E="1 кан", H=serial(), I="Петров  Иван.")
        who = s.I("IssuePostRow", 5)
        R.add(c, "получатель «Петров  Иван.» не из справочника — предложен «Петров Иван» и записан он (на руках не разделится на двоих)",
              who.startswith("OK") and s.iss(5)[8] == "Петров Иван", f"{who}; {s.iss(5)[8]!r}")
        oracle(c, s, "вопросы 1.0 перед проведением")
    finally:
        s.close()


@case
def pr27_receipts():
    """1.0 (A2): приход — вопрос о возможном повторе, «Ещё поступление» с фасовкой позиции и вопросом «1 шт = 4 кан», подсказка
    фасовки с запятой, «Отменить заказ» сначала проверяет, «Исправить приход» не меняет место перемещённого ЕИ"""
    c = "PR27"
    p = new_wms("pr27")
    s = ready(Session(p))
    today = serial()
    try:
        g = [s.B("TestPackGuess", x, module="WmsPack") for x in ("Профиль 20х20, 2,5 м", "Труба 11,7 м", "Уголок 4.5 м")]
        R.add(c, "подсказка фасовки не теряет запятую: «2,5 м» → 2,5 м, «11,7 м» → 11,7 м, «4.5 м» → 4,5 м",
              g == ["2,5|м", "11,7|м", "4,5|м"], str(g))
        order(s, 1, "1", "Антифриз (4 канистры)", qty="2", supplier="Ромашка", art="AF-1")
        put(s, "Заказы", 1, G=1.0)
        r1 = receive(s, 1, "4", doc="УПД-7", place="1-1-1")
        ei1 = ords(s, 1)[21]
        # the same receipt again on another order row: asked; «Не проводить» leaves it unposted
        order(s, 2, "2", "Антифриз (4 канистры)", qty="2", supplier="Ромашка", art="AF-1")
        put(s, "Заказы", 2, G=1.0)
        put(s, "Заказы", 2, F="1", C="УПД-7", O=today, N=today, U="1-1-2")
        s.B("TestPrimeChoice", 1, module="WmsPrimeUi")
        d1 = s.Rc("ReceiptPostRow", 2)
        shown = s.B("TestPrimeShown", module="WmsPrimeUi")
        d2 = s.Rc("ReceiptPostRow", 2)
        R.add(c, "такое же поступление (поставщик, документ, артикул, количество, дата) — вопрос «возможный повтор»: «Не проводить» — не проведено; "
                 "«Провести — это другое поступление» — проведено с отметкой в AB",
              r1.startswith("OK") and d1.startswith("SKIP:") and "повтор" in d1 and "Провести — это другое поступление" in shown
              and d2.startswith("OK") and "Возможный дубль" in str(ords(s, 2)[27]), f"{r1}; {d1}; {d2}; AB {ords(s, 2)[27]!r}")
        # «Ещё поступление» of position 1: by the document 1 шт, the fact 4 — asked; the new EI gets the packing of ЕИ of the position
        dtxt = T.strftime("%d.%m.%Y")
        m1 = s.Rc("ReceiptAddRow", 1, "4", "1", "УПД-8", dtxt, dtxt, "1-1-3", "")
        ei2 = m1.split("|")[-1] if m1.startswith("OK") else ""
        pk2 = s.B("PackLine", ei_row(ei2), "шт", module="WmsPack") if ei2 else ""
        R.add(c, "«Ещё поступление»: «по документу 1 шт, факт 4» — вопрос, проведено 1 шт; у нового ЕИ фасовка позиции «в 1 шт — 4 кан»",
              m1.startswith("OK") and s.stock(ei_row(ei2)) == 1.0 and pk2 == "в 1 шт — 4 кан", f"{m1}; {pk2!r}")
        # «Отменить заказ» of a position that was received: refused before any question
        s.U("TestUiAuto", 2)
        s.goto("Заказы", "$A$2")
        s.OU("BtnOrdCancel")
        msg = s.U("TestUiLastMessage")
        s.U("TestUiAuto", 1)
        R.add(c, "«Отменить заказ» по полученной позиции — сразу понятный отказ, без вопроса «Отменить заказ?»",
              msg.startswith("ERR:") and "не отменяется" in msg, msg)
        # «Исправить приход» of an EI moved since: its current place stays
        put(s, "Корректировки", 1, B="Перемещение", C=ei1, J="2-2-2", K=today, L="переложили")
        mv = s.B("AdjustPostRow", 1, module="WmsAdjust")
        fx = s.Rc("ReceiptFixRow", 1, "4", "1", "УПД-7", dtxt, dtxt, "1-1-9", "")
        R.add(c, "«Исправить приход» с новым местом после перемещения ЕИ — текущее место ЕИ (2-2-2) не меняется, строка прихода — 1-1-9",
              mv.startswith("OK") and fx.startswith("OK") and s.stock_row(ei_row(ei1))[5] == "2-2-2" and ords(s, 1)[20] == "1-1-9",
              f"{mv}; {fx}; {s.stock_row(ei_row(ei1))}; U {ords(s, 1)[20]!r}")
        oracle(c, s, "приходы 1.0")
    finally:
        s.close()


@case
def pr28_places():
    """1.0 (A3, A4): места и карта — проверка места при вводе с подсказкой, справочник «Места» (пол, улица), повтор адреса на
    схеме, «5.12.3»-дата в схеме, латинские P/C/A, смешанные разделители, точный поиск места, «EI-…» латиницей, место возврата
    «добавить / заменить», карта пересчитывается после операций, «на руках» без разделения по точкам и пробелам"""
    c = "PR28"
    p = new_wms("pr28")
    s = ready(Session(p))
    try:
        topology(s, {1: (3, 2)})
        t = s.doc.Sheets.getByName("Топология")
        t.getCellByPosition(5, 1).setString("1-1-2")                 # the same address drawn twice (a long place)
        cdate = t.getCellByPosition(6, 1)
        cdate.NumberFormat = s.doc.NumberFormats.getStandardFormat(2, s.doc.CharLocale)
        cdate.setValue(float((datetime.date(2003, 12, 5) - datetime.date(1899, 12, 30)).days))   # «5.12.3» became a date
        s.B("OnTopoChange", module="WmsMap")
        pl = s.doc.Sheets.getByName("Места")
        pl.getCellByPosition(0, 1).setString("Улица")
        pl.getCellByPosition(1, 1).setString("Улица")
        pl.getCellByPosition(0, 2).setString("1-2-2")
        pl.getCellByPosition(4, 2).setString("да")
        s.B("OnPlacesChange", module="WmsMap")
        known = [s.B("PlaceKnown", k, module="WmsMap") for k in ("1-1-1", "5-12-3", "УЛИЦА", "1-24")]
        keys = [s.B("PlaceKey", x, module="WmsView") for x in ("P1-C1-1", "A-1-2", "А-1-2")]
        lst = list(s.B("PlaceList", "1-1-1, 1-2-1; 1-3-1", module="WmsView"))
        R.add(c, "место по схеме: 1-1-1 на карте, «5.12.3»-дата в схеме — место 5-12-3, «Улица» — из «Места», «1-24» неизвестно; "
                 "латинские P/C — ряд и секция, латинская A = русская А; «, ;» вместе — три места",
              known == [1, 1, 2, 0] and keys[0] == "1-1-1" and keys[1] == keys[2] and lst == ["1-1-1", "1-2-1", "1-3-1"], f"{known}; {keys}; {lst}")
        # a place typed with a slip: asked, the nearest place of the map offered
        order(s, 1, "1", "Болт М8", qty="5")
        s.B("TestPrimeChoice", 1, module="WmsPrimeUi")
        r1 = receive(s, 1, "5", place="1-11")
        shown = s.B("TestPrimeShown", module="WmsPrimeUi")
        ei1 = ords(s, 1)[21]
        if not ei1:
            raise RuntimeError(f"приход не проведён: {r1}; {shown!r}; контроль {ords(s, 1)[24]!r}")
        R.add(c, "приход с местом «1-11» (нет на карте) — вопрос с похожими местами; выбрано «1-1-1» — оно и записано в строку и в «Наличие»",
              r1.startswith("OK") and "нет на карте" in shown and s.stock_row(ei_row(ei1))[5] == "1-1-1" and ords(s, 1)[20] == "1-1-1",
              f"{r1}; {shown!r}; {s.stock_row(ei_row(ei1))}; U {ords(s, 1)[20]!r}")
        order(s, 2, "2", "Песок", qty="3")
        r2 = receive(s, 2, "3", place="Улица")
        order(s, 3, "3", "Гайка М8", qty="4")
        r3 = receive(s, 3, "4", place="1-1-2")
        mb = s.B("MapBuild", "OCC", True, module="WmsMap")
        occ = s.B("TestMapEIsAll", "1-1-2", module="WmsMap").split(";")
        R.add(c, "карта: ЕИ на «Улице» не «не на карте» (справочник «Места»); адрес 1-1-2 нарисован дважды — обе ячейки заняты",
              r2.startswith("OK") and r3.startswith("OK") and "вне стеллажей 1," in mb and len(occ) == 2 and all(x.split("|")[1] != "" for x in occ),
              f"{r2}; {r3}; {mb}; {occ}")
        # where to put: neither the closed 1-2-2 nor the other cell of 1-1-2
        s.goto("Заказы", "$A$5")
        order(s, 4, "4", "Шайба", qty="1")
        ws = s.B("WhereSuggest", "Шайба", "", module="WmsMap")
        R.add(c, "«Куда положить?» не предлагает закрытое место (1-2-2 «закрыто» в «Местах») и занятое 1-1-2 во второй ячейке схемы",
              "1-2-2" not in ws and "1-1-2|" not in ws, ws)
        # the map is counted again when it is shown after operations
        put(s, "Выдачи", 1, L=ei1, E="5", H=serial(), I="егр")
        i1 = s.I("IssuePostRow", 1)
        s.B("OnMapFocus", module="WmsMap")
        after = s.B("TestMapEIs", "1-1-1", module="WmsMap")
        R.add(c, "после выдачи всего остатка карта при открытии пересчитана: на 1-1-1 пусто", i1.startswith("OK") and after.endswith("|"), f"{i1}; {after!r}")
        # the exact place in «Поиск», a Latin «EI-»
        order(s, 5, "5", "Винт", qty="2")
        r5 = receive(s, 5, "2", place="1-1-1")
        order(s, 6, "6", "Шуруп", qty="2")
        r6 = receive(s, 6, "2", place="11-1-1")
        f1 = s.B("FindList", "1-1-1", module="WmsFind")
        rows = [r for r in page(s, "Поиск", 12, 40, 0, 9) if r[1] == "Наличие"]
        places = sorted(set(r[7] for r in rows))
        n5 = int(ords(s, 5)[21].split("-")[1]) if ords(s, 5)[21] else 0
        fe = s.B("FindRun", f"EI-{n5}", module="WmsFind")
        R.add(c, "«Поиск» места 1-1-1 не находит 11-1-1; «EI-…» латиницей — карточка ЕИ",
              r5.startswith("OK") and r6.startswith("OK") and f1.startswith("OK") and "11-1-1" not in places and fe.startswith("OK:ЕИ-"),
              f"{r5}; {r6}; {f1}; {places}; {fe}")
        # the place of a return: part of the EI still on the shelf — «добавить» (the first button)
        e5 = ords(s, 5)[21]
        put(s, "Выдачи", 2, L=e5, E="1", H=serial(), I="егр")
        s.I("IssuePostRow", 2)
        put(s, "Возврат", 1, B=str(int(s.iss(2)[0])), F="1", H=serial(), I="егр", J="1-3-1")
        rt = s.B("ReturnPostRow", 1, module="WmsReturn")
        R.add(c, "возврат на другое место, когда часть ЕИ лежит на прежнем — вопрос; «Добавить место» — у ЕИ «1-1-1, 1-3-1»",
              rt.startswith("OK") and s.stock_row(ei_row(e5))[5] == "1-1-1, 1-3-1", f"{rt}; {s.stock_row(ei_row(e5))}")
        # «на руках» — one person whatever the spaces and dots
        put(s, "Выдачи", 3, L=e5, E="1", H=serial(), I="Петров И.И.")
        put(s, "Выдачи", 4, L=ords(s, 6)[21], E="1", H=serial(), I="петров  и. и.")
        s.I("IssuePostRow", 3)
        s.I("IssuePostRow", 4)
        h = s.B("HandsShow", "Петров И. И.", module="WmsFind")
        R.add(c, "«на руках у Петров И. И.» — обе выдачи («Петров И.И.» и «петров  и. и.»)", h.startswith("OK") and ": 2 из 2" in h, h)
        oracle(c, s, "места 1.0")
    finally:
        s.close()


@case
def pr23_many_places():
    """D-100: одна позиция на нескольких местах («1-2-3, 1-3-3»): карта, выдачи по местам, «Куда положить?» добавляет место, перемещение списком"""
    c = "PR23"
    p = new_wms("pr23")
    s = ready(Session(p))
    try:
        topology(s, {1: (6, 3), 2: (6, 3)})
        keys = [s.B("PlaceList", x, module="WmsView") for x in ("1-2-3, 1-3-3", "1-2-3; 1-3-3", "1-2-3 + 1-3-3", "1-2-3,1-3-3", "5,12,3", "Офис", "1-2-3, у стены")]
        R.add(c, "место читается списком: «1-2-3, 1-2-4», «;», «+», без пробела — два места; «5,12,3» — одно место 5-12-3; «Офис» — не место; «1-2-3, у стены» — место 1-2-3 (пометка не мешает)",
              [list(k) for k in keys] == [["1-2-3", "1-3-3"]] * 4 + [["5-12-3"], [], ["1-2-3"]], str(keys))
        order(s, 1, "1", "Незамерзайка (4 канистры)", qty="1")
        rc = receive(s, 1, "1", place="1-2-3, 1-3-3")
        ei = ords(s, 1)[21]
        b0 = s.B("MapBuild", "OCC", module="WmsMap")
        m3, m4 = s.B("TestMapEIs", "1-2-3", module="WmsMap"), s.B("TestMapEIs", "1-3-3", module="WmsMap")
        n = str(ei_row(ei))
        R.add(c, "ЕИ на двух местах — на карте занято оба (места 1-2-3 и 1-3-3 показывают этот ЕИ), ЕИ не в «Не на карте»",
              rc.startswith("OK") and "занято 2," in b0 and n in m3.split("|")[1].split(",") and n in m4.split("|")[1].split(","), f"{b0}; {m3}; {m4}")
        put(s, "Выдачи", 1, L=ei, E="0,25", H=serial(), I="егр")
        iss = s.I("IssuePostRow", 1)
        s.B("MapBuild", "ISS", module="WmsMap", )
        mp = s.doc.Sheets.getByName("Карта склада")
        rc3 = s.B("TestMapPlace", "1-2-3", module="WmsMap").split(":")
        rc4 = s.B("TestMapPlace", "1-3-3", module="WmsMap").split(":")
        g3 = mp.getCellByPosition(600 + int(rc3[1]), int(rc3[0])).getValue()
        g4 = mp.getCellByPosition(600 + int(rc4[1]), int(rc4[0])).getValue()
        R.add(c, "слой «Выдачи 30 дней»: выдача ЕИ видна на обоих его местах", iss.startswith("OK") and g3 == 1.0 and g4 == 1.0, f"{iss}; {g3}; {g4}")
        order(s, 2, "2", "Незамерзайка (4 канистры)", qty="2", U="1-1-1")
        s.goto("Заказы", "$B$3")
        s.B("TestPrimeChoice", 0, module="WmsPrimeUi")
        wp = s.B("WherePut", module="WmsMap")
        u = ords(s, 2)[20]
        R.add(c, "«Куда положить?» при уже указанном месте — вопрос «добавить или заменить»; «Да» — товар займёт несколько мест («1-1-1, 1-2-3»)",
              wp.startswith("OK:место 1-1-1, ") and u.startswith("1-1-1, 1-2-"), f"{wp}; {u}")
        put(s, "Корректировки", 1, B="Перемещение", C=ei, J="2-1-1, 2-1-2", K=serial(), L="не влезло")
        mv = s.B("AdjustPostRow", 1, module="WmsAdjust")
        b2 = s.B("MapBuild", "OCC", module="WmsMap")
        m21 = s.B("TestMapEIs", "2-1-2", module="WmsMap")
        R.add(c, "перемещение списком мест: ЕИ теперь на 2-1-1 и 2-1-2 (1-2-3 и 1-3-3 свободны)", mv.startswith("OK") and n in m21.split("|")[1].split(",")
              and n not in s.B("TestMapEIs", "1-2-3", module="WmsMap").split("|")[1].split(","), f"{mv}; {b2}; {m21}")
        oracle(c, s, "несколько мест у ЕИ — только место, учёт не тронут")
    finally:
        s.close()


@case
def pr24_parts_recovery():
    """D-099 и сбой: фасовка EI_PACK, выдача, возврат, списание частями и процентом — сбой до журнала (откат) и после журнала
    («Восстановить» доводит операцию из журнала, вместе с частями)"""
    c = "PR24"
    out = []
    for op in ("EI_PACK", "ISSUE", "RETURN", "WRITE_OFF"):
        for point in (2, 3):
            p = new_wms(f"pr24_{op}_{point}")
            s = ready(Session(p))
            order(s, 1, "1", "Незамерзайка -30 (4 канистры)", qty="1")
            receive(s, 1, "1", place="1-1-1")
            ant = ords(s, 1)[21]
            if op != "EI_PACK":
                s.B("PackSet", ant, "4", "кан", "", module="WmsPack")
            if op == "RETURN":
                issue(s, 1, ant, "2 кан")
            s.B("TestSetFault", point, 2)
            if op == "EI_PACK":
                r = s.B("PackSet", ant, "4", "кан", "", module="WmsPack")
            elif op == "ISSUE":
                r = issue(s, 1, ant, "1 кан")
            elif op == "RETURN":
                put(s, "Возврат", 1, B="1", F="1 кан", H=serial(), I="егр")
                r = s.B("ReturnPostRow", 1, module="WmsReturn")
            else:
                put(s, "Корректировки", 1, B="Списание", C=ant, G="25%", K=serial(), L="разлили")
                r = s.B("AdjustPostRow", 1, module="WmsAdjust")
            s = reopen(p, s)
            st1 = state(s)
            if st1[1] == "TAIL":
                s.B("ActionRecover")
            P, inf, _ = s.adjcheck(0)
            pk = list(s.doc.Sheets.getByName("_PACK").getCellRangeByPosition(0, ei_row(ant), 2, ei_row(ant)).getDataArray()[0])
            bal = s.stock(ei_row(ant))
            if point == 3:
                # done from the journal: the balance in the unit of the document, the journal keeps how it was given
                j = journal_last(s, op)
                ok = {"EI_PACK": pk == [ant, "кан", 4.0] and bal == 1.0 and j.get("N") == "4",
                      "ISSUE": bal == 0.75 and j.get("PARTS") == "1" and j.get("PART_UNIT") == "кан",
                      "RETURN": bal == 0.75 and j.get("PARTS") == "1",
                      "WRITE_OFF": bal == 0.75 and j.get("PCT") == "25"}[op]
            else:
                # rolled back: nothing of the operation is left
                ok = {"EI_PACK": pk == ["", "", ""] and bal == 1.0, "ISSUE": bal == 1.0, "RETURN": bal == 0.5, "WRITE_OFF": bal == 1.0}[op]
            out.append((op, point, r, st1, state(s), ok, bal, not P, P[:2]))
            s.close()
    R.add(c, "сбой фасовки и движений частями до журнала (операции нет, остаток и фасовка как были) и после журнала («Восстановить» "
             "доводит её: остаток в единице документа, в журнале — части или процент): «работа разрешена», оракул без расхождений",
          all(x[2] == "CRASH-SIM" and x[4] == ("CLEAN", "") and x[5] and x[7] for x in out), str(out))


@case
def pr25_status_colours():
    """«Цвет статуса» (0.8): every rule of the colour evaluated by Calc itself on the real texts of the statuses — the first true
    rule gives the colour; the start of an order keeps its line"""
    c = "PR25"
    p = new_wms("pr25")
    s = ready(Session(p))
    try:
        topo = s.doc.Sheets.getByName("Топология")
        cell = topo.getCellByPosition(1000, 0)

        def first_style(sheet, col, text):
            cf = s.doc.Sheets.getByName(sheet).getCellRangeByPosition(col, 1, col, 1048575).ConditionalFormat
            letter = chr(65 + col) if col < 26 else "A" + chr(65 + col - 26)
            for k in range(cf.getCount()):
                f = cf.getByIndex(k).getFormula1()
                if f.startswith("AND(ROW()>2"):
                    continue                      # the start of an order: checked by S01 (the same test with the line)
                cell.setFormula("=" + f.replace(f"${letter}2", '"' + text.replace('"', '""') + '"'))
                if cell.getValue() == 1:
                    return cf.getByIndex(k).getStyleName()
            return ""
        exp = {("Заказы", 22): {"Получено": "ST_Done", "Получено без документов": "ST_Done", "Дополнительное поступление": "ST_Done",
                                "Ожидается": "ST_Wait", "Частично получено": "ST_Part", "Частично получено / остаток отменён": "ST_Part",
                                "Отменено": "ST_Gray", "Поступление удалено (сторно)": "ST_Gray", "Заказ": ""},
               ("Заказы", 24): {"КОПИЯ строки 5": "ST_Bad", "Не проведено: нет количества": "ST_Bad", "Проведено": "",
                                "Ошибка: не указано место": "ST_Bad", "Внимание: возможный дубль": "ST_Part"},
               ("Наличие", 7): {"Требует разбора": "ST_Part", "Приход удалён (сторно)": "ST_Gray", "Активен": ""},
               ("Выдачи", 17): {"Проведено": "ST_Done", "Проведено, исправлено": "ST_Done", "Удалено (сторно)": "ST_Gray",
                                "КОПИЯ строки 3": "ST_Bad", "Не проведено: остаток 1 кан": "ST_Bad", "К выдаче: 1 кан = 0,25 шт": "",
                                "Ошибка: остаток 0": "ST_Bad", "Внимание: получатель не найден": "ST_Part"}}
        got = {k: {t: first_style(k[0], k[1], t) for t in v} for k, v in exp.items()}
        cell.setString("")
        bad = {k: {t: (got[k][t], v[t]) for t in v if got[k][t] != v[t]} for k, v in exp.items()}
        R.add(c, "цвет статуса — первое верное правило на настоящих текстах: получено (и без документов, и доп. поступление) — зелёный, "
                 "ожидается — синий, частично (и с отменённым остатком) — жёлтый, отменено и сторно — серый; «Требует разбора» — жёлтый, "
                 "«Приход удалён (сторно)» — серый; копия и не проведено — красный; прочее — без цвета",
              all(not v for v in bad.values()), str(bad))
    finally:
        s.close()


@case
def pr29_tool_data():
    """1.0: «СДЕЛАТЬ РАБОЧЕЙ» carries the data of the old tools to the tools next to the new book: «Журнал работы», the acts
    («История», the number of the next act), the settings, the folders WMS_Labels and WMS_Docs"""
    c = "PR29"
    sys.path.insert(0, os.path.join(ROOT, "tools"))
    import upgrade
    from wmslo import Office
    d = os.path.join(OUT, "cases", "pr29")
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(d)
    r = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "build_toolbox.py"), os.path.join(d, "rel")], capture_output=True, text=True,
                       timeout=1200)
    old_dir, new_dir = os.path.join(d, "OLD"), os.path.join(d, "NEW")
    shutil.copytree(os.path.join(d, "rel", "WMS_TOOLBOX"), os.path.join(old_dir, "WMS_TOOLBOX"))
    shutil.copytree(os.path.join(d, "rel", "WMS_TOOLBOX"), os.path.join(new_dir, "WMS_TOOLBOX"))
    os.makedirs(os.path.join(old_dir, "WMS_Labels"))
    open(os.path.join(old_dir, "WMS_Labels", "history.csv"), "w", encoding="utf-8").write("job;time;ei;copies;how\n1;2026-10-01;ЕИ-00000001;2;PDF\n")
    o = Office(f"pr29{os.getpid()}", os.path.join(OUT, "profiles"))
    try:
        mg = o.load(os.path.join(old_dir, "WMS_TOOLBOX", "WMS_MANAGER.ods"), macros=0, hidden=True)
        mg.Sheets.getByName("Журнал работы").getCellRangeByPosition(0, 1, 3, 2).setDataArray(
            ((46300.0, "Разгрузка / погрузка", "фура", 2.5), (46301.0, "Работа погрузчиком", "стеллаж", 1.0)))
        mg.store()
        mg.close(True)
        dc = o.load(os.path.join(old_dir, "WMS_TOOLBOX", "WMS_DOCS.ods"), macros=0, hidden=True)
        st = dc.Sheets.getByName("Настройки")
        row = [k for k in range(30) if st.getCellByPosition(0, k).getString() == "Номер следующего акта"][0]
        st.getCellByPosition(1, row).setString("7")
        st.getCellByPosition(1, 3).setString("file:///старый/снимок/")
        dc.Sheets.getByName("История").getCellRangeByPosition(0, 1, 5, 1).setDataArray(((6.0, "01.10.2026", "акт передачи", "Петров", 3.0, "ACT-00006.odt"),))
        dc.store()
        dc.close(True)
        rep = upgrade.carry_toolbox(o, os.path.join(old_dir, "WMS_TOOLBOX"), os.path.join(new_dir, "WMS_TOOLBOX"))
        folders = upgrade.carry_tool_folders(old_dir, new_dir)
        mg = o.load(os.path.join(new_dir, "WMS_TOOLBOX", "WMS_MANAGER.ods"), macros=0, hidden=True)
        log = mg.Sheets.getByName("Журнал работы").getCellRangeByPosition(0, 1, 3, 2).getDataArray()
        mg.close(True)
        dc = o.load(os.path.join(new_dir, "WMS_TOOLBOX", "WMS_DOCS.ods"), macros=0, hidden=True)
        st = dc.Sheets.getByName("Настройки")
        nxt = st.getCellByPosition(1, row).getString()
        snap = st.getCellByPosition(1, 3).getString()
        hist = dc.Sheets.getByName("История").getCellRangeByPosition(0, 1, 5, 1).getDataArray()
        dc.close(True)
    finally:
        o.terminate()
    R.add(c, "данные старых инструментов переходят к новым: «Журнал работы» (2 строки), акты — «История» и номер следующего акта 7, "
             "настройки (кроме снимка в работе), папка WMS_Labels с историей этикеток",
          r.returncode == 0 and log[0][1] == "Разгрузка / погрузка" and log[1][3] == 1.0 and nxt == "7" and "старый" not in snap
          and hist[0][5] == "ACT-00006.odt" and os.path.isfile(os.path.join(new_dir, "WMS_Labels", "history.csv"))
          and any("WMS_MANAGER" in x and "Журнал работы" in x for x in rep) and folders == ["папка WMS_Labels скопирована"],
          f"{rep}; {folders}; {log}; {nxt}; {snap}; {hist}")


def main():
    sel = [a.lower() for a in sys.argv[1:]]
    for fn in CASES:
        if sel and not any(fn.__name__.startswith(x) for x in sel):
            continue
        try:
            fn()
        except Exception as e:
            R.error(fn.__name__.split("_")[0].upper(), e)
        # the staging office of the transfer tests (run_legacy_transfer.maker) is not left running after a case
        lt = sys.modules.get("run_legacy_transfer")
        if lt is not None and "o" in lt._mk:
            try:
                lt._mk.pop("o").terminate()
            except Exception:
                pass
    R.save(os.path.join(OUT, "results_prime.json"))
    items = R.items
    lines = ["# PRIME 0.8 — результаты", "", f"PASS {sum(1 for x in items if x['status'] == 'PASS')}, "
             f"FAIL {sum(1 for x in items if x['status'] == 'FAIL')}, SKIP {sum(1 for x in items if x['status'] == 'SKIP')}", ""]
    lines += [f"- [{x['status']}] {x['case']} {x['test']}" for x in items]
    open(os.path.join(OUT, "TEST_REPORT_prime.md"), "w", encoding="utf-8").write("\n".join(lines) + "\n")
    print(f"PASS {sum(1 for x in items if x['status'] == 'PASS')} FAIL {sum(1 for x in items if x['status'] == 'FAIL')} "
          f"SKIP {sum(1 for x in items if x['status'] == 'SKIP')}")
    return 0 if all(x["status"] == "PASS" for x in items) else 1


if __name__ == "__main__":
    sys.exit(main())
