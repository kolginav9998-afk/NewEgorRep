"""GUI check of «Выдачи» on a virtual screen (Xvfb): the AutoFilter dropdown of the protected sheet, used like a user.

    Xvfb :99 &  WMS_TEST_OUT=/tmp/wms_gui python3 tests/gui_phase2.py [--display :99]

LibreOffice 24.2 has no «use AutoFilter» sheet protection option, so the check is that the real dropdown opens on the
protected sheet, a value can be unticked, OK filters the rows, «Провести» works for a visible row while filtered, and the
dropdown clears the filter again. Prints one line per step; exits 0 when every step passed.
"""
import argparse
import os
import sys
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
    """opens the AutoFilter dropdown of column col like a click on its button; returns (popup, value list).
    LibreOffice must be idle before and after, otherwise the popup closes at once (it loses the focus)."""
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    p = new_wms("gui_filter")
    o = Office(unique("gui"), PROFILES, headless=False, display=a.display)
    try:
        ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
        doc = o.load(p, macros=4)
        sh = doc.Sheets.getByName("Выдачи")
        who = ["Офис", "Производство", "Офис", "Производство"]
        res = []
        for i, w in enumerate(who):
            r = 1 + i
            sh.getCellByPosition(11, r).setString(str(10 + i))
            sh.getCellByPosition(4, r).setString("1")
            sh.getCellByPosition(7, r).setValue(46290)
            sh.getCellByPosition(8, r).setString(w)
            res.append(o.basic(doc, "WmsIssue", "IssuePostRow", r))
        sh.getCellByPosition(11, 5).setString("20")          # an unposted row for «Офис»
        sh.getCellByPosition(4, 5).setString("1")
        sh.getCellByPosition(7, 5).setValue(46290)
        sh.getCellByPosition(8, 5).setString("Офис")
        step("4 выдачи проведены, лист защищён", res == ["OK:1", "OK:2", "OK:3", "OK:4"] and sh.isProtected(), str(res))
        fw, lst = open_filter(o, ut, doc, sh, 8)
        ent = entries(lst)
        step("окно автофильтра «Кому выдано» открылось на защищённом листе", [t for _, t in ent] == ["Офис", "Производство"], str(ent))
        for cid, text in ent:
            if text == "Производство":
                lst.getChild(cid).executeAction("CLICK", ())
        press_ok(o, fw)
        vis = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)]
        step("снята галочка «Производство» → OK: строки «Производство» скрыты, остальные видны",
             vis == [True, False, True, False, True], str(vis))
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Выдачи.$L$6")
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        o.basic(doc, "WmsUi", "BtnPost")
        msg = o.basic(doc, "WmsUi", "TestUiLastMessage")
        step("«Провести» для видимой строки при фильтре", msg == "OK:5" and sh.getCellByPosition(0, 5).getValue() == 5, msg)
        doc.store()
        vis_saved = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)]
        with zipfile.ZipFile(p) as z:
            nfilt = z.read("content.xml").decode("utf-8").count('table:visibility="filter"')
        step("сохранение при фильтре: скрытие строк на время сохранения снято (все строки сохранены показанными), "
             "условия фильтра не удалены, после сохранения фильтр восстановлен, книга не «изменена»",
             vis_saved == vis and not doc.isModified() and nfilt == 0, f"{vis_saved}; строк, сохранённых скрытыми, {nfilt}")
        doc.close(True)
        doc = o.load(p, macros=4)
        sh = doc.Sheets.getByName("Выдачи")
        vis_reopen = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)]
        step("повторное открытие: WMS вернула фильтр из окна (несколько значений), книга не «изменена»",
             vis_reopen == vis and not doc.isModified(), str(vis_reopen))
        fw, lst = open_filter(o, ut, doc, sh, 8)
        for cid, text in entries(lst):
            if text == "Производство":
                lst.getChild(cid).executeAction("CLICK", ())
        press_ok(o, fw)
        vis = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)]
        step("галочка возвращена → все строки видны", all(vis), str(vis))
        doc.setModified(False)
        doc.close(True)
    finally:
        o.terminate()
    ok = bool(RESULTS) and all(RESULTS)
    print(f"GUI: {sum(RESULTS)} из {len(RESULTS)} PASS", flush=True)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
