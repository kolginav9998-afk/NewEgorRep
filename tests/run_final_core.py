"""Final Core (FINAL WMS MARATHON §2, §3) — automated scenarios: moves, write-offs, inventory corrections, the snapshot
for the tools and the batches of the tools.

    WMS_TEST_OUT=/tmp/wms_out python3 tests/run_final_core.py [case ...]

Every case copies a freshly built test book (synthetic EIs 1–200, NEXT_EI 201, five recipients) into
WMS_TEST_OUT/cases/<case>/ and drives real LibreOffice processes like a user: typing through the UI (the change handler
of «Корректировки» runs), the macros the buttons are bound to for the row under the cursor (questions answered in the
test mode, window values preset), kill -9 and restart. The book and the journal are checked by the independent oracle
tests/adjust_oracle.py (it runs the oracles of Phases 3–5: every operation replayed from the journal and compared with
«Наличие», «Заказы», «Выдачи», «Возврат», «Иной приход», «Корректировки», the service tables and the counters).
Results: WMS_TEST_OUT/results_final_core.json and WMS_TEST_OUT/TEST_REPORT_final_core.md.
"""
import csv
import glob
import hashlib
import os
import random
import subprocess
import sys
import time

import uno  # noqa: F401  (LibreOffice Python-UNO)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Session, Results, new_wms, template  # noqa: E402

R = Results()
CASES = []
POSTED = "1111111111111111110"      # WmsConfig.ADJUST_LOCKS_POSTED
OPEN = "1001110000001111100"        # WmsConfig.ADJUST_LOCKS_OPEN
D1, D2, D3, D4 = "21.09.2026", "22.09.2026", "23.09.2026", "24.09.2026"
KEEP = "<как предложено>"            # WmsAdjustUi test seam: the window field keeps the value it proposed
COUNTERS = ("LAST_SEQ", "NEXT_EI", "NEXT_NO", "NEXT_RET", "NEXT_SPL", "NEXT_ADJ")
KIND = {"m": "Перемещение", "w": "Списание", "i": "Инвентаризация"}


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


def jfields(s, t):
    return [e["fields"] for e in s.journal()[0] if e["type"] == t]


def oracle(c, s, name, expect_tail=0):
    P, inf, hist = s.adjcheck(expect_tail)
    R.add(c, "оракул: " + name, not P, f"{inf}; {P[:4]}")
    return hist


def row(s, r, k, e, qty=None, fact=None, book=None, place=None, date=D1, reason="причина", note=None):
    """one row of «Корректировки» typed like a user: the kind, the EI, the value of the kind, the date, the reason"""
    s.adjust_input(r, B=KIND.get(k, k), C=e, G=qty, H=fact, I=book, J=place, K=date, L=reason, S=note)


def post(s, r):
    return s.click_adj("BtnAdjPost", r)


def fix(s, r, main=KEEP, date=KEEP, reason=KEEP):
    s.AU("TestAdjInput", main, date, reason)
    return s.click_adj("BtnAdjFix", r)


def delete(s, r):
    return s.click_adj("BtnAdjDelete", r)


def ctl(s, r):
    return str(s.adj(r)[16])


def place(s, n):
    return s.stock_row(n)[5]


def confirms(s):
    return int(s.U("TestUiConfirmCount"))


def counters(s):
    return [s.sysv(k) for k in COUNTERS]


def issue(s, r, e, q, who="ива", date=D2):
    s.issue_input(r, ei=e, qty=q, date=date, who=who)
    return s.post(r)


def order(s, r, a="З-1", b="Болт М10", h="100", i="шт", l="Петрович", u="A-1", f=None, n=D1):
    s.order_input(r, A=a, B=b, H=h, I=i, L=l, F=f, N=n, U=u, C="УПД-1", O=D1)


def reopen(p, s):
    s.kill()
    s = ready(Session(p))
    return s, s.B("ActionForceUnlock")


def adj_state(s, r, eis):
    return (s.adj(r)[:18], s.adj_locks(r), [s.stock_row(n) for n in eis], [s.adjrec(k) for k in range(1, 5)], counters(s), njournal(s))


def diff(a, b):
    names = ("строка", "защита", "«Наличие»", "_ADJ", "счётчики", "журнал")
    return [n for n, x, y in zip(names, a, b) if x != y]


# ================================================================ Y01–Y04 posting

@case
def y01_move():
    c = "Y01"
    s = ready(Session(new_wms("y01")))
    try:
        j0 = njournal(s)
        row(s, 1, "m", "1", place="Z-9", reason="переезд стеллажа")
        prev = s.adj(1)
        res = post(s, 1)
        a = s.adj(1)
        f = (jfields(s, "MOVE") or [{}])[0]
        R.add(c, "перемещение: ЕИ → новое место → дата → причина; строка показывает старое место (M) и остаток (N = O), статус «Проведено», "
                 "№ 1 из NEXT_ADJ; в «Наличие» место «Z-9», количество не изменилось",
              res == "OK:1" and prev[12] == "A-01-1" and prev[16].startswith("Можно провести") and a[0] == 1.0 and a[1] == "Перемещение"
              and a[2] == ei(1) and a[9] == "Z-9" and a[12] == "A-01-1" and a[13] == 100.0 and a[14] == 100.0 and a[16] == "Проведено"
              and s.adj_locks(1) == POSTED and place(s, 1) == "Z-9" and s.stock(1) == 100.0 and s.sysv("NEXT_ADJ") == 2.0,
              f"{res}; {prev[12:17]}; {a}")
        R.add(c, "журнал: одна операция MOVE с номером, ЕИ, местом «откуда» и «куда», остатком, датой и причиной; _ADJ: MOVE, изменение 0, LIVE",
              njournal(s) == j0 + 1 and (f.get("ADJ"), f.get("KIND"), f.get("EI"), f.get("FROM"), f.get("TO"), f.get("BAL"), f.get("DATE"), f.get("REASON"))
              == ("1", "MOVE", ei(1), "A-01-1", "Z-9", "100", "2026-09-21", "переезд стеллажа")
              and s.adjrec(1)[:8] == (1.0, "MOVE", ei(1), 0.0, "LIVE", 1.0, "A-01-1", "Z-9"), f"{f}; {s.adjrec(1)}")
        oracle(c, s, "перемещение")
    finally:
        s.close()


@case
def y02_refusals():
    c = "Y02"
    s = ready(Session(new_wms("y02")))
    try:
        cases = [
            ("m", "1", dict(place="A-01-1"), "совпадает с текущим местом"),
            ("m", "1", dict(place=""), "не указано новое место"),
            ("m", "1", dict(place="B-1", qty="1"), "перемещение не меняет количество"),
            ("w", "1", dict(qty="101"), "нельзя: остаток"),
            ("w", "1", dict(qty="0"), "больше 0"),
            ("w", "1", dict(qty="-1"), "отрицательное"),
            ("w", "1", dict(qty="1/2"), "дробь"),
            ("w", "1", dict(qty="1", place="B-1"), "только G"),
            ("i", "4", dict(fact="12"), "расхождения нет"),
            ("i", "4", dict(fact="abc"), "недопустимый символ"),
            ("Прочее", "1", dict(qty="1"), "неизвестен"),
            ("w", "999", dict(qty="1"), "не найден"),
            ("w", "ЕИ-12", dict(qty="1", reason=""), "причина"),
            ("w", "2", dict(qty="1", date="31.02.2026"), "не является датой"),
        ]
        fails = []
        n0, j0 = s.sysv("NEXT_ADJ"), njournal(s)
        for i, (k, e, kw, want) in enumerate(cases, start=1):
            row(s, i, k, e, **kw)
            res = post(s, i)
            if not (res.startswith("ERR") and want in res and s.adj(i)[0] == "" and ("Не проведено" in ctl(s, i))):
                fails.append(f"{k}/{e}/{kw}: {res[:90]} | {ctl(s, i)[:60]}")
        R.add(c, f"отказы ({len(cases)} случаев: то же место, пустое место, количество у перемещения, списание больше остатка / 0 / минус / дробь, "
                 "лишнее поле, инвентаризация без расхождения / не число, неизвестный вид, нет ЕИ, нет причины, неверная дата) — ничего не проведено, "
                 "причина в «Контроль»", not fails and s.sysv("NEXT_ADJ") == n0 and njournal(s) == j0, "; ".join(fails[:3]))
        # an EI whose receipt was cancelled has nothing to correct
        order(s, 20, f="5")
        rc = s.click_ord("BtnRcvPost", 20)
        dl = s.click_ord("BtnRcvDelete", 20)
        row(s, 20, "w", "201", qty="1")
        res = post(s, 20)
        R.add(c, "ЕИ, приход которого удалён (сторно), не корректируется: отказ «корректировать нечего»",
              rc.startswith("OK") and dl.startswith("OK") and res.startswith("ERR") and "корректировать нечего" in res, f"{rc}; {dl}; {res}")
        oracle(c, s, "отказы")
    finally:
        s.close()


@case
def y03_write_off():
    c = "Y03"
    s = ready(Session(new_wms("y03")))
    try:
        row(s, 1, "w", "2", qty="10,5", reason="брак")
        r1 = post(s, 1)
        row(s, 2, "w", "5", qty="5", reason="истёк срок")
        r2 = post(s, 2)
        a = s.adj(1)
        R.add(c, "списание: остаток уменьшается на количество (250,5 → 240), разница P −10,5; списание всего остатка — остаток 0, ЕИ остаётся",
              r1 == "OK:1" and r2 == "OK:2" and s.stock(2) == 240.0 and a[6] == 10.5 and a[13] == 250.5 and a[14] == 240.0 and a[15] == -10.5
              and s.stock(5) == 0.0 and s.stock_row(5)[0] == ei(5), f"{r1}; {r2}; {a}; {s.stock(5)}")
        row(s, 3, "w", "5", qty="1")
        r3 = post(s, 3)
        R.add(c, "после списания всего остатка следующее списание — отказ (нельзя уйти ниже нуля)", r3.startswith("ERR") and "остаток" in r3, r3)
        # the balance mirror «Заказы».X of a receipt EI follows the write-off; it is not an issue to a person
        order(s, 1, f="10")
        rc = s.click_ord("BtnRcvPost", 1)
        row(s, 4, "w", "201", qty="3")
        r4 = post(s, 4)
        x = s.ord(1, "X")
        iss = s.doc.Sheets.getByName("Выдачи").getCellRangeByPosition(0, 1, 17, 3).getDataArray()
        R.add(c, "списание ЕИ обычного прихода: «Заказы».X показывает новый остаток (7); на листе «Выдачи» ничего не появилось (это не выдача)",
              rc.startswith("OK") and r4.startswith("OK") and x == 7.0 and all(v == "" for rr in iss for v in rr), f"{rc}; {r4}; X={x}")
        oracle(c, s, "списание")
    finally:
        s.close()


@case
def y04_inventory():
    c = "Y04"
    s = ready(Session(new_wms("y04")))
    try:
        row(s, 1, "i", "4", fact="14", reason="инвентаризация 21.09")
        prev = s.adj(1)
        r1 = post(s, 1)
        row(s, 2, "i", "6", fact="3,5", reason="инвентаризация 21.09")
        r2 = post(s, 2)
        row(s, 3, "i", "7", fact="0", reason="не найдено")
        r3 = post(s, 3)
        f = jfields(s, "INV_ADJ")
        R.add(c, "инвентаризация: учётный остаток (I, подставлен) → фактический (H) → разница (P): +2 (12 → 14), −0,25 (3,75 → 3,5), факт 0 — "
                 "остаток 0 (было 50); в журнале BOOK, FACT, DIFF — корректировка, история не переписывается",
              prev[8] == 12.0 and prev[15] == 2.0 and r1 == "OK:1" and r2 == "OK:2" and r3 == "OK:3" and s.stock(4) == 14.0 and s.stock(6) == 3.5
              and s.stock(7) == 0.0 and [(x.get("BOOK"), x.get("FACT"), x.get("DIFF")) for x in f] == [("12", "14", "2"), ("3.75", "3.5", "-0.25"), ("50", "0", "-50")],
              f"{prev[8]}; {r1}; {r2}; {r3}; {[(x.get('BOOK'), x.get('FACT'), x.get('DIFF')) for x in f]}")
        # the book balance the count was compared with: a movement after the count makes the row stale
        row(s, 4, "i", "8", fact="20", book="7", reason="пересчёт по снимку")
        is1 = issue(s, 1, "8", "1")
        r4 = post(s, 4)
        s.type_adj("I", 4, "6")
        r5 = post(s, 4)
        R.add(c, "учётный остаток пересчёта (I) не совпадает с текущим (после пересчёта была выдача) — отказ «учётный остаток изменился»; "
                 "с верным учётным — проведено, разница от него",
              is1.startswith("OK") and r4.startswith("ERR") and "учётный остаток изменился" in r4 and r5.startswith("OK") and s.stock(8) == 20.0
              and s.adj(4)[15] == 14.0, f"{is1}; {r4[:100]}; {r5}; {s.adj(4)[7:9]} {s.adj(4)[15]}")
        oracle(c, s, "инвентаризация")
    finally:
        s.close()


# ================================================================ Y05–Y09 correction and storno

@case
def y05_fix_move():
    c = "Y05"
    s = ready(Session(new_wms("y05")))
    try:
        row(s, 1, "m", "1", place="Z-9")
        post(s, 1)
        shown0 = None
        f1 = fix(s, 1, main="Z-10", reason="уточнение")
        shown0 = s.AU("TestAdjDialog").split("\n")[0].split("\t")
        R.add(c, "«Исправить» перемещения: окно предлагает проведённые место, дату, причину; новое место — ЕИ в «Z-10», статус «Проведено "
                 "(исправлено)», одна операция MOVE_FIX", f1.startswith("OK") and shown0 == ["Z-9", D1, "причина"] and place(s, 1) == "Z-10"
              and ctl(s, 1) == "Проведено (исправлено)" and s.adjrec(1)[7] == "Z-10" and len(jfields(s, "MOVE_FIX")) == 1, f"{f1}; {shown0}; {place(s, 1)}")
        f2 = fix(s, 1)
        f3 = fix(s, 1, main="A-01-1")
        row(s, 2, "m", "1", place="Q-5")
        post(s, 2)
        f4 = fix(s, 1, main="Z-11")
        R.add(c, "без изменений — «ничего не изменилось»; исправление на место «до перемещения» — отказ (это «Удалить»); после следующего "
                 "перемещения ЕИ исправить первое нельзя (ЕИ уже не там) — отказ, место не тронуто",
              f2.startswith("SKIP") and "ничего не изменилось" in f2 and f3.startswith("ERR") and "Удалить" in f3 and f4.startswith("ERR")
              and "менялось" in f4 and place(s, 1) == "Q-5", f"{f2}; {f3[:80]}; {f4[:80]}; {place(s, 1)}")
        oracle(c, s, "исправление перемещения")
    finally:
        s.close()


@case
def y06_fix_write_off():
    c = "Y06"
    s = ready(Session(new_wms("y06")))
    try:
        row(s, 1, "w", "4", qty="5")
        post(s, 1)
        f1 = fix(s, 1, main="2")
        b1 = s.stock(4)
        f2 = fix(s, 1, main="13")
        f3 = fix(s, 1, main="12", date=D2, reason="брак партии")
        a = s.adj(1)
        R.add(c, "«Исправить» списания: 5 → 2 — остаток 7 → 10; больше, чем было до списания (13 при 12) — отказ; 12 с новой датой и причиной — "
                 "остаток 0, N O пересчитаны, статус «Проведено (исправлено)»",
              f1.startswith("OK") and b1 == 10.0 and f2.startswith("ERR") and "нельзя" in f2 and f3.startswith("OK") and s.stock(4) == 0.0
              and a[6] == 12.0 and a[13] == 12.0 and a[14] == 0.0 and a[11] == "брак партии" and ctl(s, 1) == "Проведено (исправлено)",
              f"{f1}; {b1}; {f2[:70]}; {f3}; {a}")
        oracle(c, s, "исправление списания")
    finally:
        s.close()


@case
def y07_fix_inventory():
    c = "Y07"
    s = ready(Session(new_wms("y07")))
    try:
        row(s, 1, "i", "4", fact="20")
        post(s, 1)
        is1 = issue(s, 1, "4", "15")
        f1 = fix(s, 1, main="14")
        f2 = fix(s, 1, main="12")
        f3 = fix(s, 1, main="18")
        R.add(c, "«Исправить» инвентаризации: учётный остаток пересчёта (12) не меняется; факт 20 → 14 после выдачи 15 сделал бы остаток −1 — отказ; "
                 "факт = учётному — отказ (это «Удалить»); факт 18 — разница +6, остаток 5 → 3",
              is1.startswith("OK") and f1.startswith("ERR") and "отрицательным" in f1 and f2.startswith("ERR") and "Удалить" in f2
              and f3.startswith("OK") and s.stock(4) == 3.0 and s.adj(1)[15] == 6.0 and s.adj(1)[8] == 12.0, f"{is1}; {f1[:70]}; {f2[:70]}; {f3}; {s.stock(4)}")
        oracle(c, s, "исправление инвентаризации")
    finally:
        s.close()


@case
def y08_delete_move():
    c = "Y08"
    s = ready(Session(new_wms("y08")))
    try:
        row(s, 1, "m", "1", place="Z-9")
        post(s, 1)
        row(s, 2, "m", "1", place="Z-10")
        post(s, 2)
        k0 = confirms(s)
        d1 = delete(s, 1)
        k1 = confirms(s)
        d2 = delete(s, 2)
        d3 = delete(s, 1)
        R.add(c, "«Удалить» перемещения, после которого ЕИ перемещали дальше, — отказ первым окном, без подтверждения; сторно последнего "
                 "перемещения возвращает ЕИ на прежнее место («Z-10» → «Z-9»); затем и первое — «A-01-1»",
              d1.startswith("ERR") and "менялось" in d1 and k1 == k0 and d2.startswith("OK") and d3.startswith("OK") and place(s, 1) == "A-01-1"
              and ctl(s, 1) == ctl(s, 2) == "Удалено (сторно)" and s.adjrec(1)[4] == s.adjrec(2)[4] == "STORNO" and s.adj_locks(1) == POSTED,
              f"{d1[:80]}; {k0}/{k1}; {d2}; {d3}; {place(s, 1)}")
        d4 = delete(s, 1)
        f4 = fix(s, 1, main="X-1")
        R.add(c, "повторное «Удалить» — «уже удалена», «Исправить» удалённой — отказ; номера не переиспользуются",
              d4.startswith("SKIP") and f4.startswith("ERR") and s.sysv("NEXT_ADJ") == 3.0, f"{d4}; {f4[:60]}")
        oracle(c, s, "сторно перемещений")
    finally:
        s.close()


@case
def y09_delete_quantity():
    c = "Y09"
    s = ready(Session(new_wms("y09")))
    try:
        row(s, 1, "w", "4", qty="2")
        post(s, 1)
        d1 = delete(s, 1)
        row(s, 2, "i", "5", fact="9")
        post(s, 2)
        is1 = issue(s, 1, "5", "8")
        d2 = delete(s, 2)
        s.goto("Выдачи", "$L$2")
        s.U("TestUiAuto", 1)
        s.click("BtnDelete", 1)
        di = s.U("TestUiLastMessage")
        d3 = delete(s, 2)
        R.add(c, "сторно списания возвращает количество (10 → 12); сторно инвентаризации +4 после выдачи 8 из 9 сделало бы остаток −3 — отказ; "
                 "после сторно выдачи — проведено (остаток 5)",
              d1.startswith("OK") and s.stock(4) == 12.0 and is1.startswith("OK") and d2.startswith("ERR") and "отрицательным" in d2
              and di.startswith("OK") and d3.startswith("OK") and s.stock(5) == 5.0, f"{d1}; {is1}; {d2[:80]}; {di}; {d3}; {s.stock(5)}")
        oracle(c, s, "сторно списания и инвентаризации")
    finally:
        s.close()


# ================================================================ Y10 D-069 for the corrections

@case
def y10_receipt_storno_dependents():
    c = "Y10"
    s = ready(Session(new_wms("y10")))
    try:
        order(s, 1, f="10")
        s.click_ord("BtnRcvPost", 1)
        row(s, 1, "w", "201", qty="2")
        post(s, 1)
        k0 = confirms(s)
        d1 = s.click_ord("BtnRcvDelete", 1)
        k1 = confirms(s)
        dl = delete(s, 1)
        d2 = s.click_ord("BtnRcvDelete", 1)
        R.add(c, "D-069 и корректировки: сторно прихода при действующем списании его ЕИ — отказ первым окном «есть действующие корректировки… "
                 "сначала выполните их сторно», без подтверждения; после сторно списания — приход удалён",
              d1.startswith("ERR") and "корректировки" in d1 and k1 == k0 and dl.startswith("OK") and d2.startswith("OK")
              and s.stock_row(201)[7] == "Приход удалён (сторно)" and s.stock(201) == 0.0, f"{d1[:120]}; {k0}/{k1}; {dl}; {d2}")
        # a special receipt with a live move
        s.special_input(1, B="Офис", D="Стол", F="1", G="шт", H=D1, I="A-1")
        sp = s.click_spc("BtnSpcPost", 1)
        row(s, 2, "m", "202", place="B-2")
        post(s, 2)
        d3 = s.click_spc("BtnSpcDelete", 1)
        dm = delete(s, 2)
        d4 = s.click_spc("BtnSpcDelete", 1)
        R.add(c, "сторно последнего прихода «Иной приход» при действующем перемещении его ЕИ — отказ; после сторно перемещения — проведено",
              sp.startswith("OK") and d3.startswith("ERR") and "корректировки" in d3 and dm.startswith("OK") and d4.startswith("OK"),
              f"{sp}; {d3[:100]}; {dm}; {d4}")
        oracle(c, s, "D-069 и корректировки")
    finally:
        s.close()


# ================================================================ Y11–Y13 blocks, idempotency, copies, Ctrl+Z, inserted rows

@case
def y11_block_idempotency_copies():
    c = "Y11"
    s = ready(Session(new_wms("y11")))
    try:
        row(s, 1, "m", "10", place="N-1")
        row(s, 2, "w", "11", qty="1000")
        row(s, 3, "i", "12", fact="1")
        res = s.click_adj("BtnAdjPost", 1, r1=4)
        R.add(c, "«Провести» для выделенных строк: каждая строка — своя операция; ошибка одной строки не останавливает остальные (проведено 2, "
                 "ошибка 1, пустая строка пропущена)", res.startswith("OK=2;ERR=1;SKIP=1") and s.adj(1)[0] == 1.0 and s.adj(3)[0] == 2.0
              and s.adj(2)[0] == "", res.split(chr(10))[0])
        again = post(s, 1)
        blk = s.click_adj("BtnAdjPost", 1, r1=3)
        R.add(c, "повторное «Провести» проведённой строки — «уже проведено», ничего не изменилось; блок из проведённых строк — пропущены",
              again.startswith("SKIP") and "уже проведено" in again and blk.startswith("OK=0;ERR=1;SKIP=2") and s.sysv("NEXT_ADJ") == 3.0, f"{again}; {blk[:30]}")
        # a copy of a posted row (its № pasted into another row) is marked, never posted
        sh = s.doc.Sheets.getByName("Корректировки")
        was = sh.isProtected()
        sh.unprotect("wms")
        sh.getCellRangeByPosition(0, 6, 18, 6).setDataArray(sh.getCellRangeByPosition(0, 1, 18, 1).getDataArray())
        if was:
            sh.protect("wms")
        k = s.A("AdjustRowKind", 6)
        p6 = post(s, 6)
        cl = s.click_adj("BtnAdjClear", 6)
        R.add(c, "строка-копия проведённой (тот же №): не проводится, помечается КОПИЯ; «Очистить» убирает её, проведённая строка цела",
              k == "FOREIGN" and p6.startswith("SKIP") and cl.startswith("OK") and s.adj(6)[0] == "" and s.adj(1)[0] == 1.0, f"{k}; {p6}; {cl}")
        oracle(c, s, "блоки, повтор, копии")
    finally:
        s.close()


@case
def y12_ctrl_z_inserted_rows():
    c = "Y12"
    s = ready(Session(new_wms("y12")))
    try:
        row(s, 1, "m", "13", place="N-1")
        row(s, 2, "w", "14", qty="1")
        post(s, 1)
        post(s, 2)
        a = (s.adj(1), s.adj(2), s.stock_row(13), s.stock_row(14), counters(s))
        for _ in range(6):
            s.ui(".uno:Undo")
        b = (s.adj(1), s.adj(2), s.stock_row(13), s.stock_row(14), counters(s))
        R.add(c, "Ctrl+Z после проведения не отменяет корректировки (строки, «Наличие», счётчики не изменились)", a == b, "")
        s.goto("Корректировки", "$A$2")
        s.ui(".uno:SelectRow")
        s.ui(".uno:InsertRowsBefore")
        s.ui(".uno:InsertRowsBefore")
        moved = (s.adj(3)[0], s.adj(4)[0])
        fx = fix(s, 4, main="2")
        dl = delete(s, 3)
        R.add(c, "вставка строк над корректировками: строки найдены по № (подсказка устарела), исправление и сторно из сдвинутых строк "
                 "проведены, подсказки _ADJ исправлены той же операцией",
              moved == (1.0, 2.0) and fx.startswith("OK") and dl.startswith("OK") and s.adjrec(1)[5] == 3.0 and s.adjrec(2)[5] == 4.0
              and place(s, 13) == "S-13", f"{moved}; {fx}; {dl}; {s.adjrec(1)[5]} {s.adjrec(2)[5]}")
        oracle(c, s, "Ctrl+Z и вставка строк")
    finally:
        s.close()


# ================================================================ Y13–Y16 crash and recovery

def crash_series(c, tag, k, e, eis, fixv, **kw):
    """a posting interrupted at every point before the journal (each time: rolled back by the snapshot at the next
    start), then after the journal (the tail, «Восстановить» applies it once); then «Исправить» and «Удалить»
    interrupted in the middle of the writes and after the journal"""
    p = new_wms(f"y_crash_{tag}")
    s = ready(Session(p))
    try:
        r = 1
        row(s, r, k, e, **kw)
        s.doc.store()
        base = adj_state(s, r, eis)
        points = [("k", 0, "после NEXT_ADJ"), ("sheet", "Наличие", "после «Наличие»"), ("sheet", "_ADJ", "после _ADJ"),
                  ("sheet", "Корректировки", "после строки листа (журнала нет)")]
        fails = []
        for kind, arg, label in points:
            if kind == "k":
                s.B("TestSetFaultWrite", arg, 2)
            else:
                s.B("TestSetFaultSheet", arg, 2)
            crash = s.A("AdjustPostRow", r)
            s, unl = reopen(p, s)
            now = adj_state(s, r, eis)
            if not (crash == "CRASH-SIM" and "отменена по снимку" in unl and st(s)[0] == "CLEAN" and now == base):
                fails.append(f"{label}: {crash}; {unl[:60]}; различия {diff(base, now)}")
        R.add(c, f"{KIND[k]}: сбой в {len(points)} точках до записи в журнал — при каждом запуске откат по снимку: нет ни №, ни изменения "
                 "«Наличие», ввод на месте", not fails, "; ".join(fails[:3]) or f"точек {len(points)}")
        s.B("TestSetFault", 3, 2)
        crash = s.A("AdjustPostRow", r)
        s, unl = reopen(p, s)
        _, block = st(s)
        rec = s.B("ActionRecover")
        after = s.adj(r)
        again = s.A("AdjustPostRow", r)
        R.add(c, f"{KIND[k]}: сбой после записи в журнал — при запуске «хвост», «Восстановить» доводит операцию ровно один раз; повторное "
                 "«Провести» — «уже проведено»",
              crash == "CRASH-SIM" and block == "TAIL" and "восстановлено операций: 1" in rec and st(s)[0] == "CLEAN" and after[0] == 1.0
              and after[16] == "Проведено" and s.adj_locks(r) == POSTED and again.startswith("SKIP"), f"{crash}; {block}; {rec[:80]}; {after[:3]}; {again}")
        base2 = adj_state(s, r, eis)
        s.B("TestSetFault", 2, 2)
        crash = s.A("AdjustFixRow", r, fixv, s.adj(r)[10], "исправлено")
        s, unl = reopen(p, s)
        ok1 = crash == "CRASH-SIM" and adj_state(s, r, eis) == base2
        s.B("TestSetFault", 3, 2)
        crash2 = s.A("AdjustFixRow", r, fixv, s.adj(r)[10], "исправлено")
        s, unl = reopen(p, s)
        rec2 = s.B("ActionRecover")
        R.add(c, f"{KIND[k]}: сбой посреди «Исправить» — откат; после журнала — «Восстановить» доводит исправление целиком",
              ok1 and crash2 == "CRASH-SIM" and "восстановлено операций: 1" in rec2 and ctl(s, r) == "Проведено (исправлено)",
              f"{crash}; {crash2}; {rec2[:80]}; {ctl(s, r)}")
        base3 = adj_state(s, r, eis)
        s.B("TestSetFault", 2, 2)
        crash = s.A("AdjustDeleteRow", r)
        s, unl = reopen(p, s)
        ok1 = crash == "CRASH-SIM" and adj_state(s, r, eis) == base3
        s.B("TestSetFault", 3, 2)
        crash2 = s.A("AdjustDeleteRow", r)
        s, unl = reopen(p, s)
        rec3 = s.B("ActionRecover")
        R.add(c, f"{KIND[k]}: сбой посреди «Удалить» — откат; после журнала — «Восстановить» доводит сторно целиком", ok1 and crash2 == "CRASH-SIM"
              and "восстановлено операций: 1" in rec3 and ctl(s, r) == "Удалено (сторно)", f"{crash}; {crash2}; {rec3[:80]}; {ctl(s, r)}")
        oracle(c, s, f"сбои: {KIND[k]}")
    finally:
        s.close()


@case
def y13_crash_move():
    crash_series("Y13", "move", "m", "20", (20,), "M-2", place="M-1")


@case
def y14_crash_write_off():
    crash_series("Y14", "wo", "w", "21", (21,), "2", qty="1")


@case
def y15_crash_inventory():
    crash_series("Y15", "inv", "i", "22", (22,), "30", fact="40")


@case
def y16_abandon_keeps_numbers():
    c = "Y16"
    p = new_wms("y16")
    s = ready(Session(p))
    try:
        row(s, 1, "w", "23", qty="1")
        s.doc.store()
        s.B("TestSetFault", 3, 2)
        crash = s.A("AdjustPostRow", 1)
        s, unl = reopen(p, s)
        ab = s.B("ActionAbandonTail")
        a = [e for e in s.journal()[0] if e["type"] == "ABANDON"]
        f = a[0]["fields"] if a else {}
        again = post(s, 1)
        R.add(c, "«Отложить хвост» со списанием: № 1 больше не выдаётся (ABANDON: NEXT_ADJ 2); тот же ввод проводится с № 2",
              crash == "CRASH-SIM" and "отложено операций: 1" in ab and f.get("NEXT_ADJ") == "2" and again.startswith("OK") and s.adj(1)[0] == 2.0
              and s.sysv("NEXT_ADJ") == 3.0, f"{ab[:80]}; {f.get('NEXT_ADJ')}; {again}; {s.adj(1)[0]}")
        oracle(c, s, "отложенный хвост")
    finally:
        s.close()


# ================================================================ Y17–Y19 self-check, startup, main page, save

@case
def y17_self_check_and_keys():
    c = "Y17"
    s = ready(Session(new_wms("y17")))
    try:
        row(s, 1, "m", "30", place="R-1")
        row(s, 2, "w", "31", qty="1")
        post(s, 1)
        post(s, 2)
        sc = s.B("ActionSelfCheck")
        okline = [ln for ln in sc.split("\n") if "корректировки" in ln]
        s.doc.Sheets.getByName("_ADJ").unprotect("wms")
        s.doc.Sheets.getByName("_ADJ").getCellByPosition(4, 2).setString("XXX")
        s.set_sys("NEXT_ADJ", 2)
        sc2 = s.B("ActionSelfCheck")
        bad = [ln for ln in sc2.split("\n") if "корректировки" in ln]
        R.add(c, "самопроверка: раздел «корректировки» — OK (2, NEXT_ADJ выше всех №); испорченное состояние в _ADJ и NEXT_ADJ ниже № — FAIL",
              okline and okline[0].startswith("OK") and bad and bad[0].startswith("FAIL") and ("NEXT_ADJ" in bad[0] or "состояние" in bad[0]),
              f"{okline[:1]}; {bad[:1]}")
    finally:
        s.close()


@case
def y18_startup_old_books():
    c = "Y18"
    p = new_wms("y18")
    s = Session(p, macros=0)
    try:
        sy = s.doc.Sheets.getByName("_SYS")
        sy.unprotect("wms")
        sy.getCellByPosition(1, 0).setString("WMS-SYS-2")
        s.doc.store()
    finally:
        s.close()
    s = ready(Session(p))
    try:
        state, block = st(s)
        rep = s.report()
        row(s, 1, "w", "1", qty="1")
        res = s.A("AdjustPostRow", 1)
        R.add(c, "книга ядра этапа 5 (схема WMS-SYS-2): запуск заблокирован с понятной причиной («этапа 5», «Корректировки»), проведение запрещено",
              state == "BLOCKED" and "WMS-SYS-2" in rep and "этапа 5" in rep and res.startswith("BLOCKED:"), f"{state}; {rep[:160]}; {res[:60]}")
    finally:
        s.close()
    p = new_wms("y18b")
    s = Session(p, macros=0)
    try:
        s.doc.unprotect("wms")
        s.doc.Sheets.removeByName("Корректировки")
        s.doc.store()
    finally:
        s.close()
    s = ready(Session(p))
    try:
        state, block = st(s)
        rep = s.report()
        R.add(c, "книга без листа «Корректировки»: запуск заблокирован («нет листа «Корректировки»»)", state == "BLOCKED" and "Корректировки" in rep,
              f"{state}; {rep[:160]}")
    finally:
        s.close()


@case
def y19_main_page_save_reopen():
    c = "Y19"
    p = new_wms("y19")
    s = ready(Session(p))
    try:
        row(s, 1, "m", "40", place="U-1")
        post(s, 1)
        row(s, 2, "w", "41", qty="1")
        s.U("UiRefresh", "")
        main = "\n".join(s.main_status())
        s.close(save=True)
        s = ready(Session(p))
        st1 = st(s)[0]
        row(s, 3, "i", "42", fact="1")
        r3 = post(s, 3)
        R.add(c, "«Главная» показывает непроведённые корректировки; сохранение → открытие: состояние «работа разрешена», корректировка на месте, "
                 "нумерация продолжается (№ 2)", "корректировки: 1 строк" in main and st1 == "CLEAN" and s.adj(1)[0] == 1.0 and r3 == f"OK:{int(s.sysv('LAST_SEQ'))}"
              and s.adj(3)[0] == 2.0, f"{main[-120:]}; {st1}; {r3}")
        oracle(c, s, "после открытия")
    finally:
        s.close()


# ================================================================ Y20–Y21 the snapshot for the tools

def read_csv(path):
    with open(path, encoding="utf-8", newline="") as f:
        return list(csv.reader(f, delimiter=";"))


@case
def y20_snapshot():
    c = "Y20"
    s = ready(Session(new_wms("y20")))
    try:
        order(s, 1, f="10", b='Болт "М10"; оцинк.')
        s.click_ord("BtnRcvPost", 1)
        issue(s, 1, "201", "3")
        s.return_input(1, issue="1", qty="1", date=D3)
        rt = s.click_ret("BtnRetPost", 1)
        s.special_input(1, B="Старый склад", D="Дрель", F="1", G="шт", H=D1, I="C-3")
        s.click_spc("BtnSpcPost", 1)
        row(s, 1, "m", "202", place="C-4")
        post(s, 1)
        row(s, 2, "w", "201", qty="2")
        post(s, 2)
        row(s, 3, "i", "1", fact="99")
        post(s, 3)
        delete(s, 3)
        mod0, seq0, j0 = s.doc.isModified(), s.sysv("LAST_SEQ"), njournal(s)
        res = s.EX("ExportSnapshot")
        snap = s.EX("TestLastSnapshot")
        man = {r[0]: r[1:] for r in read_csv(os.path.join(snap, "manifest.csv"))[1:]}
        files = {r[0]: r for r in read_csv(os.path.join(snap, "manifest.csv"))[1:] if r[0] == "file"}
        hashes_ok, rows_ok = True, True
        for r in read_csv(os.path.join(snap, "manifest.csv"))[1:]:
            if r[0] != "file":
                continue
            fp = os.path.join(snap, r[1])
            h = hashlib.sha256(open(fp, "rb").read()).hexdigest()
            if h != r[4] or os.path.getsize(fp) != int(float(r[3])):
                hashes_ok = False
            if int(r[2]) >= 0 and "/" not in r[1] and len(read_csv(fp)) - 1 != int(r[2]):
                rows_ok = False
        R.add(c, "«Экспорт для инструментов»: папка WMS_Export/WMS_SNAPSHOT_<время>_seq<N>; manifest.csv — формат WMS-SNAPSHOT-1, версии, "
                 "схема, экземпляр, LAST_SEQ, позиция журнала, счётчики; у каждого файла число строк, размер и SHA-256 (сверены)",
              res.startswith("OK:") and man.get("format") == ["WMS-SNAPSHOT-1"] and man.get("schema") == ["WMS-SYS-3"]
              and man.get("last_seq") == [str(int(seq0))] and man.get("next_adj") == ["4"] and hashes_ok and rows_ok
              and os.path.basename(snap.rstrip("/")).startswith(f"WMS_SNAPSHOT_") and snap.rstrip("/").endswith(f"_seq{int(seq0)}"),
              f"{res[:120]}; {man.get('format')}; hashes {hashes_ok}; rows {rows_ok}")
        stock = read_csv(os.path.join(snap, "stock.csv"))
        hdr = stock[0]
        by = {r[0]: r for r in stock[1:] if r and r[0]}
        orders = read_csv(os.path.join(snap, "orders.csv"))
        adj = read_csv(os.path.join(snap, "adjustments.csv"))
        iss = read_csv(os.path.join(snap, "issues.csv"))
        sysv = {r[0]: r[1] for r in read_csv(os.path.join(snap, "service", "_SYS.csv")) if len(r) > 1}
        R.add(c, "таблицы снимка: первая строка — английские имена колонок, строка k файла = строка k листа; «Наличие» — остатки и места как "
                 "в книге (числа с точкой); «Заказы» — текст с «;» и кавычками в кавычках, даты ГГГГ-ММ-ДД; «Выдачи», «Корректировки» со "
                 "статусами (сторно — «Удалено (сторно)»); служебная _SYS — счётчики",
              hdr[:5] == ["ei", "name", "article", "unit", "qty"] and by[ei(201)][4] == "6" and by[ei(202)][5] == "C-4" and by[ei(2)][4] == "250.5"
              and orders[0][21] == "ei" and orders[1][1] == 'Болт "М10"; оцинк.' and orders[1][13] == "2026-09-21" and orders[1][21] == ei(201)
              and iss[1][4] == "3" and iss[1][7] == "2026-09-22" and [r[16] for r in adj[1:4]] == ["Проведено", "Проведено", "Удалено (сторно)"]
              and adj[1][10] == "2026-09-21" and sysv.get("NEXT_ADJ") == "4",
              f"{rt}; {hdr[:5]}; {by.get(ei(201), [])[:6]}; {orders[1][:2] if len(orders) > 1 else ''}; {[r[16] for r in adj[1:4]]}")
        svc = sorted(os.path.basename(x) for x in glob.glob(os.path.join(snap, "service", "*.csv")))
        jr = sorted(os.path.basename(x) for x in glob.glob(os.path.join(snap, "journal", "*.csv")))
        R.add(c, "снимок: служебные таблицы (service/) и копия журнала (journal/); книга не изменилась (флаг «изменена», LAST_SEQ, журнал — "
                 "как до экспорта); временной папки .part не осталось",
              "_ADJ.csv" in svc and "_SYS.csv" in svc and len(svc) == 8 and jr and s.doc.isModified() == mod0 and s.sysv("LAST_SEQ") == seq0
              and njournal(s) == j0 and not glob.glob(os.path.join(os.path.dirname(snap.rstrip("/")), ".*.part")), f"{svc}; {jr}")
        res2 = s.EX("ExportSnapshot")
        snap2 = s.EX("TestLastSnapshot")
        R.add(c, "повторный экспорт в ту же секунду — отдельная папка (снимки неизменяемы)", res2.startswith("OK:") and snap2 != snap, f"{snap2}")
        # a blocked WMS does not export
        s.set_sys("SCHEMA", "WMS-SYS-2")
        s.B("WmsStartup")
        res3 = s.EX("ExportSnapshot")
        R.add(c, "заблокированная WMS снимок не создаёт (только из рабочего состояния)", res3.startswith("ERR") and "заблокирована" in res3, res3[:100])
    finally:
        s.close()


# ================================================================ Y21–Y22 batches of the tools

def batch(path, target, bid, header, rows, extra=()):
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("#WMS-BATCH-1\n")
        if bid is not None:
            f.write(f"#id;{bid}\n")
        f.write(f"#target;{target}\n#source;WMS_INVENTORY\n")
        for x in extra:
            f.write(x + "\n")
        f.write(";".join(header) + "\n")
        for r in rows:
            f.write(";".join(r) + "\n")
    return path


def load(s, path, answer=1):
    s.EX("TestBatchFile", path)
    s.U("TestUiAuto", answer)
    try:
        s.EX("BtnLoadBatch")
        return s.U("TestUiLastMessage")
    finally:
        s.U("TestUiAuto", 1)


@case
def y21_batch_adjust():
    c = "Y21"
    s = ready(Session(new_wms("y21")))
    try:
        d = os.path.dirname(s.path)
        good = batch(os.path.join(d, "inv1.csv"), "Корректировки", "INV-20260921-1", ["kind", "ei", "fact", "book", "date", "reason"],
                     [["Инвентаризация", "ЕИ-00000050", "5", "1", D1, "инвентаризация 21.09"],
                      ["Инвентаризация", "51", "0", "8", D1, "инвентаризация 21.09"],
                      ["Инвентаризация", "52", "9", "15", D1, "инвентаризация 21.09"]])
        issue(s, 1, "52", "1")
        n0 = s.sysv("NEXT_ADJ")
        res = load(s, good)
        rows = [s.adj(r) for r in (1, 2, 3)]
        R.add(c, "пакет инвентаризации (WMS-BATCH-1): строки загружены на «Корректировки» (только колонки ввода, R «Пакет» = id) и по "
                 "подтверждению проведены обычными проверками: две проведены, у третьей учётный остаток изменился (после снимка была выдача) — "
                 "отказ в «Контроль»", res.startswith("OK=2;ERR=1")
              and [r[0] for r in rows] == [1.0, 2.0, ""] and [r[17] for r in rows] == ["INV-20260921-1"] * 3 and "учётный остаток изменился" in str(rows[2][16])
              and s.stock(50) == 5.0 and s.stock(51) == 0.0 and s.sysv("NEXT_ADJ") == n0 + 2, f"{res[:60]}; {[r[0] for r in rows]}; {rows[2][16]}")
        res2 = load(s, good)
        R.add(c, "тот же пакет повторно — отказ «уже загружен», ничего не добавлено", res2.startswith("ERR") and "уже загружен" in res2
              and s.adj(4)[0] == "" and s.adj(4)[1] == "", res2[:100])
        bads = {
            "без строки формата": ("nofmt.csv", None),
            "неизвестная колонка": ("badcol.csv", batch(os.path.join(d, "badcol.csv"), "Корректировки", "B2", ["kind", "ei", "price"], [["Списание", "1", "5"]])),
            "формула": ("formula.csv", batch(os.path.join(d, "formula.csv"), "Корректировки", "B3", ["kind", "ei", "qty"], [["Списание", "1", "=1+1"]])),
            "число полей": ("fields.csv", batch(os.path.join(d, "fields.csv"), "Корректировки", "B4", ["kind", "ei", "qty"], [["Списание", "1"]])),
            "лист": ("target.csv", batch(os.path.join(d, "target.csv"), "Наличие", "B5", ["ei"], [["1"]])),
            "нет id": ("noid.csv", batch(os.path.join(d, "noid.csv"), "Корректировки", None, ["kind", "ei", "qty"], [["Списание", "1", "1"]])),
        }
        with open(os.path.join(d, "nofmt.csv"), "w", encoding="utf-8") as f:
            f.write("#id;B1\n#target;Корректировки\nkind;ei;qty\nСписание;1;1\n")
        fails = []
        for name, (fn, _) in bads.items():
            r_ = load(s, os.path.join(d, fn))
            if not r_.startswith("ERR"):
                fails.append(f"{name}: {r_[:80]}")
        R.add(c, "неверные пакеты (нет строки формата, неизвестная колонка, формула, число полей, недопустимый лист, нет id) — отказ целиком, "
                 "ни одной строки не загружено", not fails and s.adj(4)[1] == "" and s.adj(5)[1] == "", "; ".join(fails))
        res3 = load(s, batch(os.path.join(d, "wo.csv"), "Корректировки", "WO-1", ["kind", "ei", "qty", "date", "reason"],
                             [["Списание", "60", "1", "2026-09-22", "брак"]]), answer=2)
        R.add(c, "«Нет» на вопрос «Провести сейчас?» — строки остаются непроведёнными для просмотра; дата пакета ГГГГ-ММ-ДД введена как дд.мм.гггг",
              res3.startswith("OK:загружено 1") and s.adj(4)[1] == "Списание" and s.adj(4)[0] == "" and s.adj(4)[17] == "WO-1"
              and s.adj(4)[10] == D2, f"{res3}; {s.adj(4)[:3]}; дата {s.adj(4)[10]!r}")
        oracle(c, s, "пакеты корректировок")
    finally:
        s.close()


@case
def y22_batch_special_orders():
    c = "Y22"
    s = ready(Session(new_wms("y22")))
    try:
        d = os.path.dirname(s.path)
        res = load(s, batch(os.path.join(d, "old.csv"), "Иной приход", "OLD-STOCK-1", ["type", "name", "qty", "unit", "date", "place"],
                            [["Старый склад", "Дрель", "1", "шт", D1, "C-1"], ["Старый склад", "Пила", "2", "шт", D1, "C-2"]],
                            extra=("#shared;yes",)))
        sp = [s.spc(r) for r in (1, 2)]
        R.add(c, "пакет «Иной приход» (#shared yes): две строки проведены одним поступлением (один № OLD-…), в комментарии отметка пакета",
              res.startswith("OK=2;ERR=0") and sp[0][2] == sp[1][2] == "OLD-00000001" and "[пакет OLD-STOCK-1]" in sp[0][18], f"{res[:80]}; {[x[2] for x in sp]}; {sp[0][18]}")
        res2 = load(s, batch(os.path.join(d, "ord.csv"), "Заказы", "ORD-1", ["order_no", "name", "qty_ordered", "unit", "supplier", "place", "qty_fact",
                                                                            "date_received"],
                             [["З-7", "Гайка М6", "100", "шт", "Метиз", "A-2", "100", D2], ["З-7", "Шайба", "50", "шт", "Метиз", "A-2", "", ""]]))
        R.add(c, "пакет «Заказы»: две строки заказа загружены; строка с фактом (F) проведена приходом (новый ЕИ), строка без факта — ожидаемый заказ",
              res2.startswith("OK=1;ERR=0") and str(s.ord(1, "V")).startswith("ЕИ-") and s.ord(2, "V") == "" and "[пакет ORD-1]" in str(s.ord(1, "Z")),
              f"{res2[:100]}; {s.ord(1, 'V')}; {s.ord(2, 'V')}")
        oracle(c, s, "пакеты прихода")
    finally:
        s.close()


# ================================================================ Y23 a long mixed series

@case
def y23_mixed_series():
    c = "Y23"
    s = ready(Session(new_wms("y23")))
    rnd = random.Random(20260926)
    try:
        r_adj, r_iss, n_ok, n_err, posted = 1, 1, 0, 0, []
        for step in range(1, 91):
            k = rnd.choice("mmwwiiIDFx")
            e = rnd.randint(60, 90)
            if k == "m":
                row(s, r_adj, "m", str(e), place=f"P-{rnd.randint(1, 9)}")
            elif k == "w":
                row(s, r_adj, "w", str(e), qty=str(rnd.randint(1, 4)))
            elif k == "i":
                row(s, r_adj, "i", str(e), fact=str(rnd.randint(0, 60)))
            elif k == "I":
                res = issue(s, r_iss, str(e), str(rnd.randint(1, 3)))
                r_iss += 1
                n_ok += res.startswith("OK")
                continue
            elif k in "DF" and posted:
                rr = rnd.choice(posted)
                if k == "D":
                    res = delete(s, rr)
                else:
                    kind = s.adj(rr)[1]
                    main = f"P-{rnd.randint(10, 19)}" if kind == "Перемещение" else str(rnd.randint(1, 3))
                    res = fix(s, rr, main=main)
                n_ok += res.startswith("OK")
                n_err += res.startswith("ERR")
                continue
            else:
                continue
            res = post(s, r_adj)
            if res.startswith("OK"):
                n_ok += 1
                posted.append(r_adj)
            else:
                n_err += 1
            r_adj += 1
            if step % 30 == 0:
                P, inf, _ = s.adjcheck()
                R.add(c, f"оракул после шагов 1…{step}", not P, f"{P[:3]}; операций {inf.get('adjust_ops')}")
        R.add(c, f"смешанная серия (перемещения, списания, инвентаризация, выдачи, исправления, сторно): проведено {n_ok}, отказов {n_err}; "
                 "ни одного отрицательного остатка", n_ok > 40 and all(s.stock(n) >= 0 for n in range(60, 91)), f"ok {n_ok}, err {n_err}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка после серии — без ошибок", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc.split("\n")[0])
    finally:
        s.close()



# ================================================================ Y24 MIGRATE (the transfer of the existing EIs, for tools/migrate.py)

def migrate(s, e, name="Старая дрель", art="", unit="шт", qty="1", place="M-1", cat="Инструмент", stype="Старый склад", mark="", origin="t.csv|abc|2"):
    return s.B("MigrateEI", e, name, art, unit, qty, place, cat, stype, mark, origin, module="WmsMigrate")


@case
def y24_migrate():
    c = "Y24"
    p = new_wms("y24")
    s = ready(Session(p))
    try:
        r1 = migrate(s, "ЕИ-00000500", qty="3", mark="инв. 17")
        r2 = migrate(s, "350", name="Шестерня Z-40", art="gr-40", stype="Детали", qty="12", place="D-9", cat="Детали")
        r3 = migrate(s, "ЕИ-00000900", name="Коробка", stype="Иной приход", qty="0")
        card = s.card(500)
        f = (jfields(s, "MIGRATE") or [{}])[0]
        R.add(c, "перенос ЕИ со старыми номерами: ЕИ-00000500, ЕИ-00000350 (деталь, в индексе), ЕИ-00000900 (иной приход, остаток 0, «Требует "
                 "разбора») — номера сохранены, дыры в реестре допустимы, NEXT_EI = 901; карточка: источник «Перенос (инв. 17)», тип «Старый склад»; "
                 "журнал MIGRATE хранит происхождение (файл, SHA-256, строку)",
              r1.startswith("OK") and r2.startswith("OK") and r3.startswith("OK") and card[:8] == (ei(500), "Старая дрель", "", "шт", 3.0, "M-1", "Инструмент", "Активен")
              and card[8] == "Перенос (инв. 17)" and card[9] == "Старый склад" and s.sysv("NEXT_EI") == 901.0 and s.stock_row(900)[7] == "Требует разбора"
              and ("GR-40", ei(350), "gr-40") in s.art_index() and f.get("ORIGIN") == "t.csv|abc|2" and f.get("MARK") == "инв. 17",
              f"{r1}; {r2}; {r3}; {card}; {s.sysv('NEXT_EI')}; {f.get('ORIGIN')}")
        bads = [migrate(s, "500"), migrate(s, "501", stype=""), migrate(s, "502", stype="Склад"), migrate(s, "503", qty="-1"),
                migrate(s, "504", name="Шестерня 2", art="GR-40", stype="Детали"), migrate(s, "505", name=""), migrate(s, "506", art="", stype="Детали")]
        R.add(c, "перенос отказывает (ничего не записано): ЕИ уже есть, пустой тип источника (D-087), неизвестный тип, отрицательное количество, "
                 "артикул детали уже у другого ЕИ, нет наименования, деталь без артикула",
              all(x.startswith("ERR") for x in bads) and s.sysv("NEXT_EI") == 901.0 and len(jfields(s, "MIGRATE")) == 3, "; ".join(x[:60] for x in bads))
        # the migrated part is refilled through «Иной приход» (its article → the same EI); the storno of the refill keeps the EI
        s.special_input(1, B="Детали", E="GR-40", F="5", G="шт", H=D1)
        sp = s.click_spc("BtnSpcPost", 1)
        after = s.stock(350)
        dl = s.click_spc("BtnSpcDelete", 1)
        R.add(c, "деталь, перенесённая с артикулом, пополняется через «Иной приход» в тот же ЕИ-00000350 (12 → 17); сторно пополнения возвращает "
                 "12, ЕИ остаётся «Активен» (перенесённый остаток — не приход строки)",
              sp.startswith("OK") and s.spc(1)[13] == ei(350) and after == 17.0 and dl.startswith("OK") and s.stock(350) == 12.0
              and s.stock_row(350)[7] == "Активен", f"{sp}; {s.spc(1)[13]}; {after}; {dl}; {s.stock_row(350)[7]}")
        is1 = issue(s, 1, "500", "1")
        row(s, 1, "w", "500", qty="1")
        wo = post(s, 1)
        R.add(c, "перенесённый ЕИ живёт как обычный: выдача и списание", is1.startswith("OK") and wo.startswith("OK") and s.stock(500) == 1.0, f"{is1}; {wo}")
        # a crash of a transfer: before the journal — rolled back; after it — «Восстановить» completes it once
        s.doc.store()
        s.B("TestSetFault", 2, 2)
        cr = migrate(s, "700")
        s, unl = reopen(p, s)
        rolled = s.stock_row(700)[0] == "" and s.sysv("NEXT_EI") == 901.0
        s.B("TestSetFault", 3, 2)
        cr2 = migrate(s, "700")
        s, unl = reopen(p, s)
        rec = s.B("ActionRecover")
        again = migrate(s, "700")
        R.add(c, "сбой переноса до журнала — откат по снимку; после журнала — «Восстановить» доводит перенос ровно один раз, повтор — отказ «уже есть»",
              cr == "CRASH-SIM" and rolled and cr2 == "CRASH-SIM" and "восстановлено операций: 1" in rec and s.stock_row(700)[0] == ei(700)
              and again.startswith("ERR") and "уже есть" in again, f"{cr}; {rolled}; {cr2}; {rec[:60]}; {again[:60]}")
        oracle(c, s, "перенос")
    finally:
        s.close()

@case
def y25_self_check_surplus_and_migrated():
    c = "Y25"
    s = ready(Session(new_wms("y25")))
    try:
        order(s, 1, f="10")
        r1 = s.click_ord("BtnRcvPost", 1)                      # EI 201: received 10
        s.special_input(1, B="Офис", D="Стол", F="5", G="шт", H=D1, I="A-1")
        r2 = s.click_spc("BtnSpcPost", 1)                      # EI 202: received 5
        row(s, 1, "i", "201", fact="12")
        a1 = post(s, 1)                                         # a surplus +2 over the receipt
        row(s, 2, "i", "202", fact="8")
        a2 = post(s, 2)                                         # a surplus +3
        m = migrate(s, "ЕИ-00000600", name="Шестерня", art="GR-7", stype="Детали", qty="4")
        s.special_input(2, B="Детали", E="GR-7", F="2", G="шт", H=D1)
        r3 = s.click_spc("BtnSpcPost", 2)                      # the migrated part refilled: 4 → 6
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка без ошибок, когда инвентаризация нашла излишек (остаток выше прихода: обычный 10 → 12, иной приход 5 → 8) и "
                 "перенесённая деталь пополнена «Иным приходом» (4 → 6)",
              all(x.startswith("OK") for x in (r1, r2, a1, a2, m, r3)) and s.stock(201) == 12.0 and s.stock(202) == 8.0 and s.stock(600) == 6.0
              and sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), f"{r1}; {r2}; {a1}; {a2}; {m}; {r3}; " + " / ".join(
                  ln[:160] for ln in sc.splitlines() if ln.startswith("FAIL") or ln.startswith("САМО")))
        # the balance forged above receipt + surplus is still found
        stk = s.doc.Sheets.getByName("Наличие")
        stk.unprotect("wms")
        stk.getCellByPosition(4, 201).setValue(20)
        stk.getCellByPosition(4, 202).setValue(30)
        sc2 = s.B("ActionSelfCheck")
        bad = [ln for ln in sc2.splitlines() if ln.startswith("FAIL")]
        R.add(c, "подделанный остаток выше прихода и излишка (20 > 10 + 2, 30 > 5 + 3) самопроверка находит",
              any("201" in ln and "излишков инвентаризации 2" in ln for ln in bad) and any("202" in ln and "излишков инвентаризации 3" in ln for ln in bad),
              " / ".join(ln[:200] for ln in bad))
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
    R.save(os.path.join(OUT, "results_final_core.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_final_core.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
