"""GUI check of «Заказы» on a virtual screen (Xvfb), used like a user.

    Xvfb :99 &  WMS_TEST_OUT=/tmp/wms_gui python3 tests/gui_phase3.py [--display :99]

1. The AutoFilter dropdown of the protected «Заказы» opens, a supplier can be unticked, OK filters the rows, «Провести
   приход» works for a visible row while filtered; the save keeps the filter (D-041) and the reopen shows it again.
2. The real input window of «Ещё поступление» and «Исправить» (in the headless tests the window is replaced by preset
   values): it opens from the button, «Отмена» does nothing, «OK» returns exactly the values WMS proposed; «Ещё
   поступление» proposes today's date as the receipt date N (D-049), «Исправить» the posted values.
Prints one line per step; exits 0 when every step passed.
"""
import argparse
import datetime
import os
import sys
import threading
import time
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
from harness import new_wms, PROFILES, unique  # noqa: E402
from wmslo import Office, props  # noqa: E402

RESULTS = []


def step(name, ok, detail=""):
    RESULTS.append(ok)
    print(f"[{'PASS' if ok else 'FAIL'}] {name}: {detail}", flush=True)


def state(obj):
    return {pv.Name: pv.Value for pv in obj.getState()}


def open_filter(o, ut, doc, sheet, col):
    doc.getCurrentController().setActiveSheet(sheet)
    o.idle()
    time.sleep(0.5)
    ut.getTopFocusWindow().getChild("grid_window").executeAction("LAUNCH", props(AUTOFILTER="", COL=str(col), ROW="0"))
    o.idle()
    time.sleep(0.3)
    fw = ut.getFloatWindow()
    return fw, fw.getChild("check_list_box")


def press_ok(o, fw):
    fw.getChild("ok").executeAction("CLICK", ())
    o.idle()
    time.sleep(0.3)


def entries(lst):
    n = int(state(lst).get("Children", "0"))
    return [(str(i), state(lst.getChild(str(i))).get("Text")) for i in range(n)]


def order_row(sh, r, a, sup, f=None):
    sh.getCellByPosition(0, r).setString(a)
    sh.getCellByPosition(1, r).setString(f"Товар {a}")
    sh.getCellByPosition(7, r).setString("10")
    sh.getCellByPosition(8, r).setString("шт")
    sh.getCellByPosition(11, r).setString(sup)
    if f:
        sh.getCellByPosition(5, r).setString(f)
        sh.getCellByPosition(13, r).setValue(46288)          # N 23.09.2026 (not today: the windows must tell the two apart)
        sh.getCellByPosition(20, r).setString("A-1")


def run_dialog(o, ut, doc, macro, title_part, button):
    """runs the button macro in a thread, waits for its window and presses «ok» or «cancel» without changing the fields
    (the fields keep the values WMS proposed); returns (window title or None, the macro's recorded result)"""
    th = threading.Thread(target=lambda: o.basic(doc, "WmsOrdersUi", macro), daemon=True)
    th.start()
    for _ in range(80):
        time.sleep(0.25)
        try:
            w = ut.getTopFocusWindow()
            title = str(state(w).get("Text", ""))
            if title_part in title and button in list(w.getChildren()):
                w.getChild(button).executeAction("CLICK", ())
                break
        except Exception:
            pass
    else:
        return None, "окно не открылось"
    # a message window after the action (a refusal) is closed like a user
    for _ in range(20):
        th.join(0.25)
        if not th.is_alive():
            break
        try:
            w = ut.getTopFocusWindow()
            if "ok" in list(w.getChildren()) and title_part not in str(state(w).get("Text", "")):
                w.getChild("ok").executeAction("CLICK", ())
        except Exception:
            pass
    th.join(30)
    return title, o.basic(doc, "WmsUi", "TestUiLastMessage")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    p = new_wms("gui_orders")
    o = Office(unique("gui3"), PROFILES, headless=False, display=a.display)
    try:
        ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
        doc = o.load(p, macros=4)
        sh = doc.Sheets.getByName("Заказы")
        res = []
        for i, sup in enumerate(["Озон", "Петрович", "Озон", "Петрович"]):
            order_row(sh, 1 + i, f"З-{i + 1}", sup, "4")
            res.append(o.basic(doc, "WmsReceipt", "ReceiptPostRow", 1 + i))
        order_row(sh, 5, "З-5", "Озон", "10")               # entered, not posted yet
        step("4 прихода проведены, лист «Заказы» защищён", res == ["OK:1", "OK:2", "OK:3", "OK:4"] and sh.isProtected(), str(res))
        fw, lst = open_filter(o, ut, doc, sh, 11)
        ent = entries(lst)
        step("окно автофильтра «Площадка / Поставщик» открылось на защищённом листе", [t for _, t in ent] == ["Озон", "Петрович"], str(ent))
        for cid, text in ent:
            if text == "Петрович":
                lst.getChild(cid).executeAction("CLICK", ())
        press_ok(o, fw)
        vis = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)]
        step("снята галочка «Петрович» → OK: строки «Петрович» скрыты", vis == [True, False, True, False, True], str(vis))
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Заказы.$B$6")
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        o.basic(doc, "WmsOrdersUi", "BtnRcvPost")
        msg = o.basic(doc, "WmsUi", "TestUiLastMessage")
        step("«Провести приход» для видимой строки при фильтре", msg == "OK:5" and sh.getCellByPosition(21, 5).getString() == "ЕИ-00000205", msg)
        doc.store()
        vis_saved = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)]
        with zipfile.ZipFile(p) as z:
            nfilt = z.read("content.xml").decode("utf-8").count('table:visibility="filter"')
        step("сохранение при фильтре «Заказы»: строки сохранены показанными, фильтр восстановлен, книга не «изменена»",
             vis_saved == vis and not doc.isModified() and nfilt == 0, f"{vis_saved}; строк, сохранённых скрытыми, {nfilt}")
        # closed like a user (the window's close command runs the WMS unload event: the lock is released)
        o.dispatch(doc, ".uno:CloseDoc")
        doc = o.load(p, macros=4)
        sh = doc.Sheets.getByName("Заказы")
        vis_reopen = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)]
        step("повторное открытие: фильтр «Заказы» возвращён, книга не «изменена»", vis_reopen == vis and not doc.isModified(), str(vis_reopen))
        fw, lst = open_filter(o, ut, doc, sh, 11)
        for cid, text in entries(lst):
            if text == "Петрович":
                lst.getChild(cid).executeAction("CLICK", ())
        press_ok(o, fw)
        step("галочка возвращена → все строки видны", all(sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)), "")
        # the real input window (interactive mode: no preset values). Typing into its fields is standard VCL behaviour; the
        # check is the round trip of the window: what WMS proposes comes back unchanged with «OK», «Отмена» does nothing
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        sysh = doc.Sheets.getByName("_SYS")
        nx = sysh.getCellByPosition(1, 5).getValue()
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Заказы.$B$2")
        title, msg = run_dialog(o, ut, doc, "BtnRcvMore", "Ещё поступление", "cancel")
        step("окно «Ещё поступление» открывается кнопкой; «Отмена» — ничего не проведено",
             title is not None and msg.startswith("SKIP:окно закрыто") and sysh.getCellByPosition(1, 5).getValue() == nx, f"{title}; {msg}")
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Заказы.$B$2")
        title, msg = run_dialog(o, ut, doc, "BtnRcvMore", "Ещё поступление", "ok")
        shown, _, back = o.basic(doc, "WmsOrdersUi", "TestRcvDialog").partition("\n")
        today = datetime.date.today().strftime("%d.%m.%Y")
        step("«OK» без ввода: окно предложило в N сегодняшнюю дату и вернуло все поля как предложены (F пусто) — отказ «не указано количество», "
             "новой строки и ЕИ нет",
             "не указано количество" in msg and sysh.getCellByPosition(1, 5).getValue() == nx and sh.getCellByPosition(21, 6).getString() == ""
             and shown.split("\t")[4] == today and back == shown, f"{title}; {msg[:120]}; предложено {shown.split(chr(9))}")
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Заказы.$B$2")
        title, msg = run_dialog(o, ut, doc, "BtnRcvFix", "Исправить приход", "ok")
        shown, _, back = o.basic(doc, "WmsOrdersUi", "TestRcvDialog").partition("\n")
        step("окно «Исправить» заполнено проведёнными значениями (дата поступления — проведённая): «OK» без правок — «ничего не изменилось», "
             "все 7 полей вернулись точно",
             title is not None and msg.startswith("SKIP:ничего не изменилось") and shown.split("\t")[4] == "23.09.2026" and back == shown,
             f"{title}; {msg}; предложено {shown.split(chr(9))}")
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
