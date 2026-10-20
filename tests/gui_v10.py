"""GUI check of the windows of WMS 1.0 on a virtual screen (Xvfb), pressed like a user (Т3), and the input typed through
the interface of Calc in the Russian locale (Т1).

    Xvfb :99 &  WMS_TEST_OUT=/tmp/wms_gui python3 tests/gui_v10.py [--display :99]

1. Typed like a person (GoToCell + EnterString, the change handler of the sheet runs, Calc converts what it converts):
   «25%» in the quantity of «Выдачи» — a quarter of what came in; «1,5 кан» — the parts of the packing; «1/2» — not a date,
   refused with an understandable message; a full date — a date, «20.10» without the year (Calc of the Russian locale reads
   it as the number 20,1) — refused with «наберите полностью»; the place «5.12.3» in «Корректировки»
   stays the text «5.12.3» (not a date) and the EI moves there.
2. The question «штуки или части?» (an issue «1» of an EI with a packing): a real window with named buttons, pressed through
   its accessibility tree — «… 1 кан …» issues a quarter, «Не проводить» posts nothing.
3. «Первый день» (the tour): its first window, «Закончить» ends it after one step.
4. «Обслуживание» opened by its tab: the light «Здоровье WMS» is filled when the page is opened.
5. «Учебное восстановление»: the question «Проверить» pressed, the answer «восстанавливается» closed by «OK»; the working
   book does not change.
Prints one line per step; exits 0 when every step passed.
"""
import argparse
import datetime
import hashlib
import os
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
from harness import new_wms, PROFILES, unique  # noqa: E402
from wmslo import Office  # noqa: E402
from gui_phase4 import step, state, RESULTS  # noqa: E402

import uno  # noqa: E402

EI = "ЕИ-00000001"
TODAY = float((datetime.date.today() - datetime.date(1899, 12, 30)).days)
D = datetime.date.today().strftime("%d.%m.%Y")


def col_ix(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def window_buttons(o, title_start):
    """the push buttons of the open window whose title starts with title_start — {label: accessible context}"""
    role = uno.getConstantByName("com.sun.star.accessibility.AccessibleRole.PUSH_BUTTON")
    tk = o.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", o.ctx)
    for k in range(tk.getTopWindowCount()):
        ctx = tk.getTopWindow(k).getAccessibleContext()
        if ctx is None or not ctx.getAccessibleName().startswith(title_start):
            continue
        found, todo = {}, [ctx]
        while todo:
            c = todo.pop()
            if c.getAccessibleRole() == role:
                found[c.getAccessibleName()] = c
            todo += [c.getAccessibleChild(i).getAccessibleContext() for i in range(c.getAccessibleChildCount())]
        if found:
            return found
    return {}


def press(o, ut, doc, module, macro, args, title_start, pick):
    """runs module.macro(*args) in a thread, waits for the window titled title_start… and presses the button whose label
    pick(label) accepts; the WMS message that follows is closed with «OK». Returns (the labels seen, the macro's result,
    the last message of WMS)"""
    res = {}
    th = threading.Thread(target=lambda: res.setdefault("r", o.basic(doc, module, macro, *args)), daemon=True)
    th.start()
    labels = None
    for _ in range(120):
        time.sleep(0.25)
        try:
            btn = window_buttons(o, title_start)
            if btn:
                labels = sorted(btn)
                [b for lbl, b in btn.items() if pick(lbl)][0].doAccessibleAction(0)
                break
        except Exception:
            pass
    for _ in range(240):
        th.join(0.25)
        if not th.is_alive():
            break
        try:
            w = ut.getTopFocusWindow()
            if str(state(w).get("Text", "")) == "WMS" and "ok" in list(w.getChildren()):
                w.getChild("ok").executeAction("CLICK", ())
        except Exception:
            pass
    th.join(30)
    return labels or [], res.get("r"), o.basic(doc, "WmsUi", "TestUiLastMessage")


def goto(o, doc, sheet, ref):
    doc.getCurrentController().setActiveSheet(doc.Sheets.getByName(sheet))
    o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${sheet}.{ref}")


def typed(o, doc, sheet, ref, text):
    goto(o, doc, sheet, ref)
    o.dispatch(doc, ".uno:EnterString", StringName=text)
    o.idle()


def stock(doc, n):
    return doc.Sheets.getByName("Наличие").getCellByPosition(4, n).getValue()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    p = new_wms("gui_v10")
    o = Office(unique("guiv10"), PROFILES, headless=False, display=a.display)
    try:
        ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
        doc = o.load(p, macros=4)
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        o.basic(doc, "WmsPack", "PackSet", EI, "4", "кан", "")
        # an EI that came in by a receipt (5 шт, 4 кан in 1 шт): the percents are counted from what came in
        zk = doc.Sheets.getByName("Заказы")
        for col, v in (("A", "1"), ("B", "Незамерзайка (4 канистры)"), ("H", "5"), ("I", "шт"), ("L", "Ромашка"), ("F", "5"), ("C", "УПД-1"),
                       ("U", "1-1-1")):
            zk.getCellByPosition(col_ix(col), 1).setString(v)
        for col in ("N", "O"):
            zk.getCellByPosition(col_ix(col), 1).setValue(TODAY)
        rc = o.basic(doc, "WmsReceipt", "ReceiptPostRow", 1)
        ein = zk.getCellByPosition(21, 1).getString()
        n = int(ein[3:])
        o.basic(doc, "WmsPack", "PackSet", ein, "4", "кан", "")
        b0 = stock(doc, n)
        # ---------------------------------------------------------------- 1. typed like a person
        iss = doc.Sheets.getByName("Выдачи")
        for r, q, d in ((1, "25%", D), (2, "1,5 кан", D), (3, "1/2", D), (4, "1", D[:5])):
            typed(o, doc, "Выдачи", f"$L${r + 1}", ein)
            typed(o, doc, "Выдачи", f"$E${r + 1}", q)
            typed(o, doc, "Выдачи", f"$H${r + 1}", d)
            typed(o, doc, "Выдачи", f"$I${r + 1}", "ива")
        cells = [(iss.getCellByPosition(4, r).getType().value, iss.getCellByPosition(4, r).getString()) for r in (1, 2, 3)]
        h = iss.getCellByPosition(7, 1)
        res = [o.basic(doc, "WmsIssue", "IssuePostRow", r) for r in (1, 2, 3, 4)]
        b1 = stock(doc, n)
        step("Т1: набрано в «Выдачах» как человеком: «25%» — четверть прихода (1,25), «1,5 кан» — 0,375, «1/2» — не дата, отказ понятный; "
             "дата «дд.мм.гггг» — дата; «дд.мм» без года — отказ «наберите полностью»",
             rc.startswith("OK") and cells[2] == ("TEXT", "1/2") and res[0].startswith("OK") and res[1].startswith("OK") and not res[2].startswith("OK")
             and "ERR-SYS" not in res[2] and "дробь" in res[2] and abs((b0 - b1) - 1.625) < 0.002
             and h.getType().value == "VALUE" and h.getString() == D and res[3].startswith("ERR") and "без года" in res[3],
             f"{rc[:40]}; {cells}; {res}; {b0} → {b1}; дата {h.getString()!r}")
        adj = doc.Sheets.getByName("Корректировки")
        typed(o, doc, "Корректировки", "$B$2", "Перемещение")
        typed(o, doc, "Корректировки", "$C$2", "ЕИ-00000002")
        typed(o, doc, "Корректировки", "$J$2", "5.12.3")
        typed(o, doc, "Корректировки", "$K$2", D)
        typed(o, doc, "Корректировки", "$L$2", "переложили")
        pc = adj.getCellByPosition(9, 1)
        mv = o.basic(doc, "WmsAdjust", "AdjustPostRow", 1)
        place = doc.Sheets.getByName("Наличие").getCellByPosition(5, 2).getString()
        step("Т1: место «5.12.3», набранное в «Корректировки», — текст (не дата), ЕИ перемещён туда",
             (pc.getType().value, pc.getString()) == ("TEXT", "5.12.3") and mv.startswith("OK") and place == "5.12.3",
             f"{pc.getType().value} {pc.getString()!r}; {mv}; {place!r}")
        # ---------------------------------------------------------------- 2. «штуки или части?»
        for r in (5, 6):
            for col, v in (("L", EI), ("E", "1"), ("H", D), ("I", "ива")):
                typed(o, doc, "Выдачи", f"${col}${r + 1}", v)
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        b2 = stock(doc, 1)
        lab, r4, _ = press(o, ut, doc, "WmsIssue", "IssuePostRow", (5,), "WMS — штуки или части?", lambda x: " 1 кан — " in x)
        b3 = stock(doc, 1)
        lab2, r5, _ = press(o, ut, doc, "WmsIssue", "IssuePostRow", (6,), "WMS — штуки или части?", lambda x: x == "Не проводить")
        b4 = stock(doc, 1)
        step("Т3: окно «штуки или части?» — настоящие кнопки; «… 1 кан …» выдаёт 0,25, «Не проводить» — ничего",
             len(lab) == 3 and "Не проводить" in lab and str(r4).startswith("OK") and abs((b2 - b3) - 0.25) < 1e-9
             and not str(r5).startswith("OK") and b4 == b3, f"{lab}; {r4}; {r5}; {b2} → {b3} → {b4}")
        # ---------------------------------------------------------------- 3. the tour
        lab, k, _ = press(o, ut, doc, "WmsHelp", "Tour", (), "Первый день — шаг 1 из ", lambda x: x == "Закончить")
        step("Т3: «Первый день» — окно шага 1 с «Дальше» и «Закончить»; «Закончить» — конец после первого шага",
             "Дальше" in lab and "Закончить" in lab and k == 1, f"{lab}; {k}")
        # ---------------------------------------------------------------- 4. «Обслуживание» by its tab
        doc.getCurrentController().setActiveSheet(doc.Sheets.getByName("Обслуживание"))
        o.idle()
        light = ""
        for _ in range(40):
            light = doc.Sheets.getByName("Обслуживание").getCellByPosition(2, 2).getString()
            if light:
                break
            time.sleep(0.25)
        step("Т6: «Обслуживание», открытое ярлыком, показывает светофор «Здоровье WMS»", light[:1] in ("✓", "!", "✗") and len(light) > 10, light[:160])
        # ---------------------------------------------------------------- 5. the drill
        doc.store()
        h0 = hashlib.sha256(open(p, "rb").read()).hexdigest()
        lab, _, msg = press(o, ut, doc, "WmsSafe", "BtnRestoreDrill", (), "Учебное восстановление", lambda x: x == "Проверить")
        h1 = hashlib.sha256(open(p, "rb").read()).hexdigest()
        step("Т7: «Учебное восстановление» — «Проверить», ответ «восстанавливается» (закрыт «OK»); рабочая книга не изменилась",
             lab == ["Отмена", "Проверить"] and "восстанавливается" in msg and h0 == h1, f"{lab}; {msg[:200]}; {h0 == h1}")
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        doc.setModified(False)
        o.dispatch(doc, ".uno:CloseDoc")
    finally:
        o.terminate()
    ok = bool(RESULTS) and all(RESULTS)
    print(f"GUI: {sum(RESULTS)} из {len(RESULTS)} PASS", flush=True)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
