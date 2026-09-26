"""GUI check of «Иной приход» on a virtual screen (Xvfb), used like a user.

    Xvfb :99 &  WMS_TEST_OUT=/tmp/wms_gui python3 tests/gui_phase5.py [--display :99]

1. The AutoFilter dropdown of the protected «Иной приход» opens on «Тип прихода», a type can be unticked, OK filters the
   rows; the save keeps the filter (D-041) and the reopen shows it again.
2. The real windows (in the headless tests they are replaced by preset answers): the question «Провести выделенные строки
   как одно поступление?» for a selected block — «Отмена» posts nothing, «Нет» gives a № per row («Да» of this Basic
   message box has no id for the UI test and does not take a key press — the answer «Да» is checked by
   tests/run_phase5.py X02, X07, X25); the confirmation of a part's move to another place — «Нет» posts nothing;
   «Удалить» of a line whose EI has a live issue — the refusal is the first window (only «OK»), no confirmation; an
   allowed storno asks «Удалить строку прихода № …?» — «Нет» cancels nothing; the window «Исправить» proposes the posted
   values («OK» without changes — «ничего не изменилось», «Отмена» — nothing); the window «Разобрать» with the list of
   types — «Отмена» does nothing, the list opens with no type selected and «OK» without a choice is refused with nothing
   changed (D-081), a type chosen in the list and «OK» identify the line (the same EI). UITest names no control of a
   Basic dialog, so the list is read and chosen through the window's accessibility tree.
Prints one line per step; exits 0 when every step passed.
"""
import argparse
import os
import sys
import threading
import time
import zipfile

import uno

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
from harness import new_wms, PROFILES, unique  # noqa: E402
from wmslo import Office  # noqa: E402
from gui_phase4 import step, state, open_filter, press_ok, entries, run_dialog, run_msg, issue_row, RESULTS  # noqa: E402

SH = "Иной приход"
D1 = 46286                                   # 21.09.2026


def spc_row(sh, r, typ, name, qty, place, art=None, unit="шт"):
    for col, v in ((1, typ), (3, name), (5, qty), (6, unit), (8, place), (4, art)):
        if v is not None:
            sh.getCellByPosition(col, r).setString(v)
    sh.getCellByPosition(7, r).setValue(D1)


def run_ask3(o, ut, doc, answer):
    """«Провести» for the selection: waits for the question with three answers and answers it — «Отмена» (its id is «yes»)
    or «Нет» (id «close»); returns (the buttons seen, the macro's result)"""
    th = threading.Thread(target=lambda: o.basic(doc, "WmsSpecialUi", "BtnSpcPost"), daemon=True)
    th.start()
    kids = None
    for _ in range(80):
        time.sleep(0.25)
        try:
            w = ut.getTopFocusWindow()
            if str(state(w).get("Text", "")) == "WMS":
                kids = [k for k in w.getChildren() if k in ("ok", "yes", "no", "close", "cancel")]
                w.getChild({"Отмена": "yes", "Нет": "close"}[answer]).executeAction("CLICK", ())
                break
        except Exception:
            pass
    # the summary after the block (a message with «OK») is closed too
    for _ in range(40):
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
    return kids or [], o.basic(doc, "WmsUi", "TestUiLastMessage")


def type_list(o, title):
    """The type list of the open window whose title starts with title — found through the accessibility tree,
    the way an assistive technology sees the window (UITest names no control of a Basic dialog)"""
    role = uno.getConstantByName("com.sun.star.accessibility.AccessibleRole.LIST")
    tk = o.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", o.ctx)
    for k in range(tk.getTopWindowCount()):
        ctx = tk.getTopWindow(k).getAccessibleContext()
        if ctx is None or not ctx.getAccessibleName().startswith(title):
            continue
        todo = [ctx]
        while todo:
            c = todo.pop()
            if c.getAccessibleRole() == role and c.getAccessibleChildCount() > 0:
                return c
            todo += [c.getAccessibleChild(i).getAccessibleContext() for i in range(c.getAccessibleChildCount())]
    return None


def run_identify(o, ut, doc, choice):
    """«Разобрать»: waits for the window, counts the selected types (none may be selected), selects choice in the list
    when given and presses «OK»; a refusal message after it is closed. Returns (selected when opened, the result)"""
    th = threading.Thread(target=lambda: o.basic(doc, "WmsSpecialUi", "BtnSpcIdentify"), daemon=True)
    th.start()
    sel0 = None
    for _ in range(80):
        time.sleep(0.25)
        try:
            w = ut.getTopFocusWindow()
            if "Разобрать иной приход" in str(state(w).get("Text", "")) and "ok" in list(w.getChildren()):
                lst = type_list(o, "Разобрать иной приход")
                sel0 = lst.getSelectedAccessibleChildCount()
                if choice:
                    names = [lst.getAccessibleChild(k).getAccessibleContext().getAccessibleName()
                             for k in range(lst.getAccessibleChildCount())]
                    lst.selectAccessibleChild(names.index(choice))
                w.getChild("ok").executeAction("CLICK", ())
                break
        except Exception:
            pass
    for _ in range(40):
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
    return sel0, o.basic(doc, "WmsUi", "TestUiLastMessage")


def goto(o, doc, ref):
    o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${SH}.{ref}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    p = new_wms("gui_special")
    o = Office(unique("gui5"), PROFILES, headless=False, display=a.display)
    try:
        ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
        doc = o.load(p, macros=4)
        sh = doc.Sheets.getByName(SH)
        res = []
        for r, (t, nm, q, pl) in enumerate((("Офис", "Стол", "1", "A-1"), ("Производство", "Рама", "2", "P-1"), ("Офис", "Стул", "1", "A-2"),
                                            ("Производство", "Ось", "3", "P-2")), start=1):
            spc_row(sh, r, t, nm, q, pl)
            res.append(o.basic(doc, "WmsSpecial", "SpecialPostRow", r, "", False))
        spc_row(sh, 5, "Офис", "Шкаф", "1", "A-3")
        spc_row(sh, 6, "Офис", "Полка", "1", "A-4")
        step("4 строки иного прихода проведены, лист «Иной приход» защищён", res == [f"OK:{i}" for i in range(1, 5)] and sh.isProtected(), str(res))
        fw, lst = open_filter(o, ut, doc, sh, 1)
        ent = entries(lst)
        texts = [t for _, t in ent]
        step("окно автофильтра «Тип прихода» открылось на защищённом листе", "Офис" in texts and "Производство" in texts, str(ent))
        for cid, text in ent:
            if text == "Производство":
                lst.getChild(cid).executeAction("CLICK", ())
        press_ok(o, fw)
        vis = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 7)]
        step("снята галочка «Производство» → OK: строки производства скрыты", vis == [True, False, True, False, True, True], str(vis))
        # the question «одно поступление?» — real window
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        k0 = o.basic(doc, "WmsUi", "TestUiConfirmCount")
        goto(o, doc, "$A$6:$S$7")
        kids, msg = run_ask3(o, ut, doc, "Отмена")
        step("выделены две строки → «Провести»: вопрос «Провести как одно поступление?» (Да / Нет / Отмена); «Отмена» — ничего не проведено",
             msg.startswith("SKIP:проведение отменено") and "close" in kids and "yes" in kids and sh.getCellByPosition(0, 5).getString() == ""
             and o.basic(doc, "WmsUi", "TestUiConfirmCount") == k0 + 1, f"кнопки {kids}; {msg}")
        goto(o, doc, "$A$6:$S$7")
        kids, msg = run_ask3(o, ut, doc, "Нет")
        evs = [sh.getCellByPosition(2, r).getString() for r in (5, 6)]
        step("«Нет»: обе строки проведены, у каждой свой № поступления и свой ЕИ", msg.startswith("OK=2;ERR=0") and evs == ["OFF-00000003", "OFF-00000004"]
             and [sh.getCellByPosition(13, r).getString() for r in (5, 6)] == ["ЕИ-00000205", "ЕИ-00000206"], f"{msg.splitlines()[0]}; {evs}")
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        doc.store()
        vis_saved = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 7)]
        with zipfile.ZipFile(p) as z:
            nfilt = z.read("content.xml").decode("utf-8").count('table:visibility="filter"')
        step("сохранение при фильтре «Иной приход»: строки сохранены показанными, фильтр восстановлен, книга не «изменена»",
             vis_saved == vis and not doc.isModified() and nfilt == 0, f"{vis_saved}; строк, сохранённых скрытыми, {nfilt}")
        o.dispatch(doc, ".uno:CloseDoc")
        doc = o.load(p, macros=4)
        sh = doc.Sheets.getByName(SH)
        vis_reopen = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 7)]
        step("повторное открытие: фильтр «Иной приход» возвращён, книга не «изменена»", vis_reopen == vis and not doc.isModified(), str(vis_reopen))
        fw, lst = open_filter(o, ut, doc, sh, 1)
        for cid, text in entries(lst):
            if text == "Производство":
                lst.getChild(cid).executeAction("CLICK", ())
        press_ok(o, fw)
        step("галочка возвращена → все строки видны", all(sh.getRows().getByIndex(r).IsVisible for r in range(1, 7)), "")
        # a part: the move to another place is confirmed in a window
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        spc_row(sh, 9, "Детали", "Шестерня", "100", "D-1", art="ABC-100")
        rp = o.basic(doc, "WmsSpecial", "SpecialPostRow", 9, "", False)
        spc_row(sh, 10, "Детали", "Шестерня", "50", "Z-9", art="ABC-100")
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        goto(o, doc, "$B$11")
        kids, msg = run_msg(o, ut, doc, "BtnSpcPost", "WmsSpecialUi")
        step("пополнение детали с другим местом: окно «Место ЕИ детали изменится … Провести?» (Да / Нет); «Нет» — строка не проведена",
             rp.startswith("OK") and "close" in kids and msg.startswith("SKIP:перемещение ЕИ не подтверждено") and sh.getCellByPosition(0, 10).getString() == ""
             and doc.Sheets.getByName("Наличие").getCellByPosition(5, 207).getString() == "D-1", f"{rp}; кнопки {kids}; {msg}")
        # «Удалить»: a refusal is the first window; an allowed storno asks for the confirmation
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        iss = doc.Sheets.getByName("Выдачи")
        issue_row(iss, 1, "207", "10", "Иванов Иван Андреевич")
        ri = o.basic(doc, "WmsIssue", "IssuePostRow", 1)
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        k1 = o.basic(doc, "WmsUi", "TestUiConfirmCount")
        goto(o, doc, "$B$10")
        kids, msg = run_msg(o, ut, doc, "BtnSpcDelete", "WmsSpecialUi")
        k2 = o.basic(doc, "WmsUi", "TestUiConfirmCount")
        step("«Удалить» строки детали, по ЕИ которой есть выдача: первое окно — отказ (только «OK»), подтверждения нет; строка не удалена",
             ri.startswith("OK") and "ok" in kids and "close" not in kids and msg.startswith("ERR:По этому ЕИ") and k2 == k1
             and sh.getCellByPosition(16, 9).getString() == "Проведено: новый ЕИ", f"{ri}; кнопки {kids}; {msg[:80]}")
        goto(o, doc, "$B$2")
        kids, msg = run_msg(o, ut, doc, "BtnSpcDelete", "WmsSpecialUi")
        k3 = o.basic(doc, "WmsUi", "TestUiConfirmCount")
        step("«Удалить» строки без выдач: окно подтверждения «Удалить строку прихода № 1?» (Да / Нет); «Нет» — строка не удалена",
             "close" in kids and "ok" not in kids and msg.startswith("SKIP:удаление не подтверждено") and k3 == k2 + 1
             and sh.getCellByPosition(16, 1).getString() == "Проведено: новый ЕИ", f"кнопки {kids}; {msg[:40]}")
        # «Исправить»: the window with the posted values
        goto(o, doc, "$B$3")
        title, msg = run_dialog(o, ut, doc, "BtnSpcFix", "Исправить строку", "cancel", module="WmsSpecialUi")
        step("окно «Исправить» открывается кнопкой; «Отмена» — ничего не исправлено", title is not None and msg.startswith("SKIP:окно закрыто без исправления"),
             f"{title}; {msg}")
        goto(o, doc, "$B$3")
        title, msg = run_dialog(o, ut, doc, "BtnSpcFix", "Исправить строку", "ok", module="WmsSpecialUi")
        shown, back = (o.basic(doc, "WmsSpecialUi", "TestSpcDialog").split("\n") + [""])[:2]
        step("окно «Исправить» заполнено проведёнными значениями (количество, дата, место, …, наименование, единица): «OK» без правок — «ничего не изменилось»",
             title is not None and msg.startswith("SKIP:ничего не изменилось") and shown.split("\t") == ["2", "21.09.2026", "P-1", "", "", "", "", "Рама", "", "шт"]
             and back == shown, f"{title}; {msg}; предложено {shown.split(chr(9))}")
        # «Разобрать»: the window with the list of types
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        spc_row(sh, 11, "Иной", "Коробка", "1", "Q-1")
        rp2 = o.basic(doc, "WmsSpecial", "SpecialPostRow", 11, "", False)
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        goto(o, doc, "$B$12")
        title, msg = run_dialog(o, ut, doc, "BtnSpcIdentify", "Разобрать иной приход", "cancel", module="WmsSpecialUi")
        step("окно «Разобрать» открывается кнопкой; «Отмена» — строка не разобрана", rp2.startswith("OK") and title is not None
             and msg.startswith("SKIP:окно закрыто без разбора") and sh.getCellByPosition(16, 11).getString() == "Проведено: требует разбора", f"{title}; {msg}")
        goto(o, doc, "$B$12")
        sel0, msg = run_identify(o, ut, doc, None)
        card0 = doc.Sheets.getByName("Наличие").getCellRangeByPosition(0, 208, 9, 208).getDataArray()[0]
        step("окно «Разобрать» открывается без выбранного типа (D-081); «OK» без выбора — отказ «укажите, что это за приход», ничего не изменено",
             sel0 == 0 and msg.startswith("ERR") and "укажите, что это за приход" in msg and card0[7] == "Требует разбора"
             and sh.getCellByPosition(16, 11).getString() == "Проведено: требует разбора", f"выбрано при открытии {sel0}; {msg[:80]}; {card0[7]}")
        goto(o, doc, "$B$12")
        sel1, msg = run_identify(o, ut, doc, "Офис")
        card = doc.Sheets.getByName("Наличие").getCellRangeByPosition(0, 208, 9, 208).getDataArray()[0]
        step("выбран тип «Офис» в списке → «OK»: тот же ЕИ-00000208 — «Активен», тип источника «Офис»",
             msg.startswith("OK") and card[0] == "ЕИ-00000208" and card[7] == "Активен" and card[9] == "Офис"
             and sh.getCellByPosition(16, 11).getString() == "Проведено: разобрано", f"{msg}; {card}")
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
