"""GUI check of «Корректировки» and «Главная» on a virtual screen (Xvfb), used like a user.

    Xvfb :99 &  WMS_TEST_OUT=/tmp/wms_gui python3 tests/gui_final_core.py [--display :99]

1. The list of the kinds in B «Вид» of the protected sheet: the three kinds are offered (a validation list).
2. The real windows (in the headless tests they are replaced by preset values): «Исправить» opens with the posted values
   of the row (the new place, the date, the reason) — «Отмена» does nothing, «OK» without changes — «ничего не
   изменилось», a place typed into the window's field (through the accessibility tree of the window, the way an assistive
   technology does it: UITest names no control of a Basic dialog) — the correction is posted; «Удалить» asks «Удалить
   корректировку № …?» — «Нет» cancels nothing; «Удалить» of a move after which the EI moved on is refused in the first
   window (only «OK»), no confirmation.
3. «Экспорт для инструментов» on «Главная»: the button creates the snapshot, the message names its folder.
Prints one line per step; exits 0 when every step passed.
"""
import argparse
import glob
import os
import sys
import threading
import time

import uno

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
from harness import new_wms, PROFILES, unique  # noqa: E402
from wmslo import Office  # noqa: E402
from gui_phase4 import step, state, run_dialog, run_msg, RESULTS  # noqa: E402

SH = "Корректировки"


def adj_row(sh, r, kind, e, place=None, qty=None):
    sh.getCellByPosition(1, r).setString(kind)
    sh.getCellByPosition(2, r).setString(e)
    if place:
        sh.getCellByPosition(9, r).setString(place)
    if qty:
        sh.getCellByPosition(6, r).setString(qty)
    sh.getCellByPosition(10, r).setValue(46286)             # 21.09.2026
    sh.getCellByPosition(11, r).setString("проверка окна")


def goto(o, doc, ref):
    o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${SH}.{ref}")


def edit_field(o, title, index, text):
    """types text into the index-th edit field of the open window whose title starts with title — through the
    accessibility tree of the window (UITest names no control of a Basic dialog)"""
    role = uno.getConstantByName("com.sun.star.accessibility.AccessibleRole.TEXT")
    tk = o.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", o.ctx)
    for k in range(tk.getTopWindowCount()):
        ctx = tk.getTopWindow(k).getAccessibleContext()
        if ctx is None or not ctx.getAccessibleName().startswith(title):
            continue
        found, todo = [], [ctx]
        while todo:
            c = todo.pop(0)
            if c.getAccessibleRole() == role:
                found.append(c)
            todo += [c.getAccessibleChild(i).getAccessibleContext() for i in range(c.getAccessibleChildCount())]
        found[index].setText(text)
        return True
    return False


def run_fix_typed(o, ut, doc, text):
    """«Исправить»: waits for the window, types a new place into its first field, presses «Исправить» (OK)"""
    th = threading.Thread(target=lambda: o.basic(doc, "WmsAdjustUi", "BtnAdjFix"), daemon=True)
    th.start()
    typed = False
    for _ in range(80):
        time.sleep(0.25)
        try:
            w = ut.getTopFocusWindow()
            if "Исправить корректировку" in str(state(w).get("Text", "")) and "ok" in list(w.getChildren()):
                typed = edit_field(o, "Исправить корректировку", 0, text)
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
    return typed, o.basic(doc, "WmsUi", "TestUiLastMessage")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    p = new_wms("gui_final_core")
    o = Office(unique("guifc"), PROFILES, headless=False, display=a.display)
    try:
        ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
        doc = o.load(p, macros=4)
        sh = doc.Sheets.getByName(SH)
        v = sh.getCellByPosition(1, 1).Validation
        step("B «Вид» на защищённом листе «Корректировки» — список: Перемещение, Списание, Инвентаризация",
             sh.isProtected() and v.getFormula1() == '"Перемещение";"Списание";"Инвентаризация"', v.getFormula1())
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        adj_row(sh, 1, "Перемещение", "ЕИ-00000001", place="Z-1")
        adj_row(sh, 2, "Списание", "ЕИ-00000002", qty="1")
        res = [o.basic(doc, "WmsAdjust", "AdjustPostRow", r) for r in (1, 2)]
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        step("две корректировки проведены (перемещение, списание)", res == ["OK:1", "OK:2"], str(res))
        goto(o, doc, "$B$2")
        title, msg = run_dialog(o, ut, doc, "BtnAdjFix", "Исправить корректировку", "cancel", module="WmsAdjustUi")
        shown = (o.basic(doc, "WmsAdjustUi", "TestAdjDialog").split("\n") + [""])[0].split("\t")
        step("окно «Исправить» открывается кнопкой с проведёнными значениями (новое место, дата, причина); «Отмена» — ничего не исправлено",
             title is not None and msg.startswith("SKIP:окно закрыто без исправления") and shown == ["Z-1", "21.09.2026", "проверка окна"],
             f"{title}; {msg}; {shown}")
        goto(o, doc, "$B$2")
        title, msg = run_dialog(o, ut, doc, "BtnAdjFix", "Исправить корректировку", "ok", module="WmsAdjustUi")
        step("«OK» без правок — «ничего не изменилось»", title is not None and msg.startswith("SKIP:ничего не изменилось"), f"{title}; {msg}")
        goto(o, doc, "$B$2")
        typed, msg = run_fix_typed(o, ut, doc, "Z-2")
        place = doc.Sheets.getByName("Наличие").getCellByPosition(5, 1).getString()
        step("в поле окна введено новое место «Z-2» → «Исправить»: ЕИ в «Z-2», строка «Проведено (исправлено)»",
             typed and msg.startswith("OK") and place == "Z-2" and sh.getCellByPosition(16, 1).getString() == "Проведено (исправлено)", f"{typed}; {msg}; {place}")
        k0 = o.basic(doc, "WmsUi", "TestUiConfirmCount")
        goto(o, doc, "$B$3")
        kids, msg = run_msg(o, ut, doc, "BtnAdjDelete", "WmsAdjustUi")
        step("«Удалить» списания: окно «Удалить корректировку № 2?» (Да / Нет); «Нет» — ничего не удалено",
             msg.startswith("SKIP:удаление не подтверждено") and "close" in kids and sh.getCellByPosition(16, 2).getString() == "Проведено"
             and o.basic(doc, "WmsUi", "TestUiConfirmCount") == k0 + 1, f"кнопки {kids}; {msg}")
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        adj_row(sh, 3, "Перемещение", "ЕИ-00000001", place="Z-3")
        r3 = o.basic(doc, "WmsAdjust", "AdjustPostRow", 3)
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        k0 = o.basic(doc, "WmsUi", "TestUiConfirmCount")
        goto(o, doc, "$B$2")
        kids, msg = run_msg(o, ut, doc, "BtnAdjDelete", "WmsAdjustUi")
        step("«Удалить» перемещения, после которого ЕИ перемещали: первое окно — отказ (только «OK»), подтверждения нет",
             r3.startswith("OK") and kids == ["ok"] and msg.startswith("ERR:") and "менялось" in msg
             and o.basic(doc, "WmsUi", "TestUiConfirmCount") == k0, f"{r3}; кнопки {kids}; {msg[:90]}")
        doc.getCurrentController().setActiveSheet(doc.Sheets.getByName("Главная"))
        kids, msg = run_msg(o, ut, doc, "BtnExport", "WmsExport")
        snaps = glob.glob(os.path.join(os.path.dirname(p), "WMS_Export", "WMS_SNAPSHOT_*"))
        step("«Экспорт для инструментов» на «Главной»: снимок создан, окно называет папку (только «OK»)",
             msg.startswith("OK:") and kids == ["ok"] and len(snaps) == 1 and os.path.exists(os.path.join(snaps[0], "manifest.csv")),
             f"кнопки {kids}; {msg[:100]}")
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
