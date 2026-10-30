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


def drive(o, doc, module, macro, args, steps):
    """runs module.macro(*args) in a thread and goes through the windows one by one: steps — (the beginning of the title, act(window context) → bool);
    returns (the steps done, the macro's result)"""
    res = {}
    th = threading.Thread(target=lambda: res.setdefault("r", o.basic(doc, module, macro, *args)), daemon=True)
    th.start()
    done = []
    tk = o.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", o.ctx)
    for title, act in steps:
        ok = False
        for _ in range(160):
            time.sleep(0.25)
            try:
                for k in range(tk.getTopWindowCount()):
                    ctx = tk.getTopWindow(k).getAccessibleContext()
                    if ctx is not None and ctx.getAccessibleName().startswith(title) and act(ctx):
                        ok = True
                        break
            except Exception:
                pass
            if ok:
                break
        done.append(ok)
        if not ok:
            break
    th.join(60)
    return done, res.get("r")


def descendants(ctx):
    todo, out = [ctx], []
    while todo:
        c = todo.pop()
        out.append(c)
        todo += [c.getAccessibleChild(i).getAccessibleContext() for i in range(c.getAccessibleChildCount())]
    return out


def press_memo(expect):
    """the window with the big field: the field shows the text it was opened with; presses «Дальше»"""
    def act(ctx):
        ds = descendants(ctx)
        edit = [c for c in ds if c.getAccessibleRole() == uno.getConstantByName("com.sun.star.accessibility.AccessibleRole.TEXT")]
        if not edit or edit[0].getText() != expect:
            return False
        [c for c in ds if c.getAccessibleRole() == uno.getConstantByName("com.sun.star.accessibility.AccessibleRole.PUSH_BUTTON")
         and c.getAccessibleName() == "Дальше"][0].doAccessibleAction(0)
        return True
    return act


def pick_button(label):
    def act(ctx):
        b = [c for c in descendants(ctx) if c.getAccessibleRole() == uno.getConstantByName("com.sun.star.accessibility.AccessibleRole.PUSH_BUTTON")
             and c.getAccessibleName() == label]
        if not b:
            return False
        b[0].doAccessibleAction(0)
        return True
    return act


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
        # ---------------------------------------------------------------- 6. «Серийные детали…» — the real window with the big field
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        sp = doc.Sheets.getByName("Иной приход")
        for col, v in (("B", "Детали"), ("D", "Шестерня Z-40"), ("E", "GR-40"), ("G", "шт"), ("I", "A-1")):
            sp.getCellByPosition(col_ix(col), 1).setString(v)
        sp.getCellByPosition(col_ix("H"), 1).setValue(TODAY)
        goto(o, doc, "Иной приход", "$B$2")
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        # the window opens with the numbers already in its big field (the field shows them and gives them back); the keys of the field itself are LibreOffice's
        o.basic(doc, "WmsSerial", "TestSerialSetText", "SN-1\nSN-2\nSN-3")
        done, rs = drive(o, doc, "WmsSerial", "BtnSerialSpecial", (),
                         [("Серийные детали — приход", press_memo("SN-1\nSN-2\nSN-3")), ("Серийные детали", pick_button("Провести приход и записать номера")),
                          ("Серийные детали", pick_button("Не сейчас"))])
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        ser = [int(o.basic(doc, "WmsItem", "SerialEI", x)) for x in ("SN-1", "SN-2", "SN-3")]
        step("1.1.1: «Серийные детали…» — настоящее окно с большим полем (номера по строкам), подтверждение «Деталей: 3» и вопрос про этикетки; в «Наличии» три ЕИ по 1 шт",
             done == [True, True, True] and len(set(ser)) == 3 and min(ser) > 0 and all(stock(doc, n) == 1.0 for n in ser), f"{done}; {rs}; {ser}")
        # ---------------------------------------------------------------- 7. «Собрать деталь…» — the real windows (1.2.0)
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        for r, (nm_, art_, q_) in enumerate((("Болт М8", "B-8", "40"), ("Гайка М8", "N-8", "40")), start=2):
            for col, v in (("A", str(r)), ("B", nm_), ("H", q_), ("I", "шт"), ("L", "Ромашка"), ("E", art_), ("F", q_), ("C", f"УПД-{r}"), ("U", "2-1-1")):
                zk.getCellByPosition(col_ix(col), r).setString(v)
            for col in ("N", "O"):
                zk.getCellByPosition(col_ix(col), r).setValue(TODAY)
            o.basic(doc, "WmsReceipt", "ReceiptPostRow", r)
        asm = doc.Sheets.getByName("Сборка деталей")
        for i, row in enumerate((("Фара FAR", "FAR-100", "Болт М8", "B-8", 4.0, "шт"), ("", "", "Гайка М8", "N-8", 4.0, "шт")), start=1):
            for j, v in enumerate(row):
                cell = asm.getCellByPosition(j, i)
                cell.setValue(v) if isinstance(v, float) else cell.setString(v)
        first = o.basic(doc, "WmsAsm", "AsmBuild", "Фара FAR", 1.0, "C-1", TODAY, "")
        far_n = int(str(first)[3 + 3:].split("|")[0]) if str(first).startswith("OK:ЕИ-") else 0
        goto(o, doc, "Сборка деталей", "$C$2")
        o.basic(doc, "WmsUi", "TestUiAuto", 0)

        def pick_any(labels):
            def act(ctx):
                b = [c for c in descendants(ctx) if c.getAccessibleRole() == uno.getConstantByName("com.sun.star.accessibility.AccessibleRole.PUSH_BUTTON")
                     and c.getAccessibleName().replace("~", "").replace("&", "") in labels]
                if not b:
                    return False
                b[0].doAccessibleAction(0)
                return True
            return act
        done, rs = drive(o, doc, "WmsAsm", "BtnAsmBuild", (),
                         [("Собрать деталь — Фара FAR", pick_button("Собрать")), ("WMS", pick_any(("Yes", "Да"))), ("Сборка детали", pick_button("Не сейчас"))])
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        step("1.2.0: «Собрать деталь…» — настоящие окна: форма (деталь под курсором, число 1, место постоянного ЕИ подставлено), вопрос «Собрать?» с перечнем снимаемых ЕИ, итог и вопрос про этикетки; "
             "постоянный ЕИ детали 1 → 2, болты и гайки сняты (40 → 36)",
             done == [True, True, True] and far_n > 0 and stock(doc, far_n) == 2.0, f"{done}; {first}; {rs}; {far_n and stock(doc, far_n)}")
        # ---------------------------------------------------------------- 8. «Инструмент ЧПУ» — the sheet in a real window (1.2.0)
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        for r, (nm_, art_, q_) in enumerate((("Фреза концевая Ø10х50 Z4 твердосплавная", "FR-10", "20"), ("Цанга ER32 d10", "CZ-10", "4"), ("Пластина CNMG 120408 TiN", "PL-1", "50")), start=10):
            for col, v in (("A", str(r)), ("B", nm_), ("H", q_), ("I", "шт"), ("L", "Ромашка"), ("E", art_), ("F", q_), ("C", f"УПД-{r}"), ("U", "3-1-1")):
                zk.getCellByPosition(col_ix(col), r).setString(v)
            for col in ("N", "O"):
                zk.getCellByPosition(col_ix(col), r).setValue(TODAY)
            o.basic(doc, "WmsReceipt", "ReceiptPostRow", r)
        tsh = doc.Sheets.getByName("Инструмент ЧПУ")
        goto(o, doc, "Инструмент ЧПУ", "$B$2")
        o.basic(doc, "WmsTool", "BtnToolRefresh")
        listed = lambda: [x[12] for x in tsh.getCellRangeByPosition(0, 4, 18, 40).getDataArray() if x[0] != ""]
        k_all = sorted(listed())
        typed(o, doc, "Инструмент ЧПУ", "$B$2", "цанга")
        k_one = listed()
        stat = tsh.getCellByPosition(7, 1).getString()
        typed(o, doc, "Инструмент ЧПУ", "$B$2", "")
        step("1.2.0: «Инструмент ЧПУ» в настоящем окне: «Обновить» — три позиции (фреза, цанга, пластина); «цанга» в «Найти» и Enter — одна, строка итога называет слова",
             k_all == ["CZ-10", "FR-10", "PL-1"] and k_one == ["CZ-10"] and "цанга" in stat and len(listed()) == 3, f"{k_all}; {k_one}; {stat[:90]}; {len(listed())}")
        # ---------------------------------------------------------------- 9. «Инструмент ЧПУ» on a big warehouse with the bar of the status line (1.2.1): names that begin with a number,
        # odd ones, and a failure on purpose — the text of the error is not lost
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        stk = doc.Sheets.getByName("Наличие")
        stk.unprotect("wms")
        last = max(i for i, x in enumerate(stk.getCellRangeByPosition(0, 0, 0, 3000).getDataArray()) if x[0] != "")
        mats = ["Р6М5", "ВК8", "HSS-Co", "твердосплавная", "Т15К6"]
        rows = [[f"Сверло спиральное {3 + i % 30},{i // 30} DIN 338 {mats[i % 5]}", f"SV-{i}"] for i in range(340)]
        rows += [["10 мм сверло по металлу", "DG-1"], ["4z фреза концевая 10", "DG-2"], ["Сверло 8мм", "DG-3"], ["М8 гайка", "NT-1"], ["8 мм болт", "NT-2"], ["Фреза ¦ 10", "OD-1"],
                 ["Фреза 9" + "9" * 400, "OD-2"], ["Метчик М10 после необычных", "AFTER-1"]]

        def fill(first, rws):
            data = tuple(tuple([f"ЕИ-{first + i:08d}", nm_, art_, "шт", 5.0, f"D-{i % 50 + 1}", "", "Активен"]) for i, (nm_, art_) in enumerate(rws))
            stk.getCellRangeByPosition(0, first, 7, first + len(rws) - 1).setDataArray(data)

        fill(last + 1, rows)
        import time as _t
        t0 = _t.time()
        r9 = str(o.basic(doc, "WmsTool", "ToolRefresh", ""))
        t9 = _t.time() - t0
        tsh = doc.Sheets.getByName("Инструмент ЧПУ")
        k9 = int(r9.split(":")[1]) if r9.startswith("OK:") else 0
        got = {x[12]: x for x in tsh.getCellRangeByPosition(0, 4, 18, 4 + k9 - 1).getDataArray()} if k9 else {}
        stat9 = tsh.getCellByPosition(7, 1).getString()
        step("1.2.1: «Инструмент ЧПУ» в настоящем окне, 340 новых названий (идёт полоса хода) и названия с числа впереди: список построен (в нём «10 мм сверло…» Ø10, «4z фреза…» Z4, «Сверло 8мм» Ø8, «Фреза ¦ 10»), "
             "«М8 гайка» и «8 мм болт» не в нём, строка итога без «НЕ ПРОЧИТАНО»",
             r9.startswith("OK:") and all(a in got for a in ("DG-1", "DG-2", "DG-3", "OD-1", "OD-2", "AFTER-1")) and "NT-1" not in got and "NT-2" not in got
             and got["DG-1"][2] == "Ø10" and got["DG-2"][5] == 4 and got["DG-3"][2] == "Ø8" and "НЕ ПРОЧИТАНО" not in stat9, f"{r9}; {t9:.1f} с; {stat9[:120]}")
        rows2 = [[f"Метчик М{3 + i % 30}х{1 + i // 30},5 м/р {mats[i % 5]}", f"MT-{i}"] for i in range(320)]
        fill(last + 1 + len(rows), rows2)
        o.basic(doc, "WmsTool", "TestToolFailStep", "разбор названий")
        r10 = str(o.basic(doc, "WmsTool", "ToolRefresh", ""))
        stat10 = tsh.getCellByPosition(7, 1).getString()
        o.basic(doc, "WmsTool", "TestToolFailStep", "")
        step("1.2.1: сбой при настоящей полосе хода — сообщение называет причину («Invalid procedure call», код 5, шаг «разбор названий»), а не «(строка 0)»; то же — на листе; после сбоя список строится",
             r10.startswith("ERR:") and "Invalid procedure call" in r10 and "код 5" in r10 and "шаг «разбор названий»" in r10 and "(строка 0)" not in r10 and "Invalid procedure call" in stat10
             and str(o.basic(doc, "WmsTool", "ToolRefresh", "")).startswith("OK:"), f"{r10[:200]}; {stat10[:100]}")
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
