"""GUI check of «Возврат» on a virtual screen (Xvfb), used like a user.

    Xvfb :99 &  WMS_TEST_OUT=/tmp/wms_gui python3 tests/gui_phase4.py [--display :99]

1. The AutoFilter dropdown of the protected «Возврат» opens, a recipient can be unticked, OK filters the rows, «Провести»
   works for a visible row while filtered; the save keeps the filter (D-041) and the reopen shows it again.
2. The real windows (in the headless tests they are replaced by preset values): «Найти выдачу» opens from the button —
   «Отмена» does nothing, «Заполнить строку» returns exactly the values WMS proposed (today's date, the current place)
   and fills the row, nothing is posted; with an EI and several issues the window shows the list and the first issue is
   chosen; the question «ЕИ» of an empty row can be cancelled; «Исправить» proposes the posted values («OK» without
   changes — «ничего не изменилось»); the confirmation of «Удалить» answered «Нет» cancels nothing.
3. «Удалить» of a receipt on «Заказы» (D-069): while its EI has a live issue and return the first window is the refusal
   (only «OK») — no confirmation window; once the storno is allowed the confirmation «Удалить приход …?» is asked and
   «Нет» cancels nothing.
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
IVA, PET = "Иванов Иван Андреевич", "Петров Пётр Сергеевич"


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


def run_dialog(o, ut, doc, macro, title_part, button, module="WmsReturnUi", exact=False):
    """runs the button macro in a thread, waits for its window and presses the button without changing the fields
    (they keep the values WMS proposed); returns (window title or None, the macro's recorded result). A message box
    of Basic shows its «Нет» button as «close»; exact: the window title must be title_part itself (the document
    window's title contains «WMS» too)."""
    th = threading.Thread(target=lambda: o.basic(doc, module, macro), daemon=True)
    th.start()
    for _ in range(80):
        time.sleep(0.25)
        try:
            w = ut.getTopFocusWindow()
            title = str(state(w).get("Text", ""))
            if (title == title_part if exact else title_part in title) and button in list(w.getChildren()):
                w.getChild(button).executeAction("CLICK", ())
                break
        except Exception:
            pass
    else:
        return None, "окно не открылось"
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


def run_msg(o, ut, doc, macro, module):
    """runs the button macro in a thread and waits for the first window titled «WMS»; returns (the buttons of that window,
    the macro's recorded result). A message of Basic shows its button as «ok»; a confirmation («Да» / «Нет») shows only
    «Нет», as «close» («Да» has no id). The window is closed with «OK» or, for a confirmation, with «Нет»."""
    th = threading.Thread(target=lambda: o.basic(doc, module, macro), daemon=True)
    th.start()
    kids = None
    for _ in range(80):
        time.sleep(0.25)
        try:
            w = ut.getTopFocusWindow()
            if str(state(w).get("Text", "")) == "WMS":
                kids = [k for k in w.getChildren() if k in ("ok", "yes", "no", "close", "cancel")]
                w.getChild("ok" if "ok" in kids else "close").executeAction("CLICK", ())
                break
        except Exception:
            pass
    th.join(30)
    return kids or [], o.basic(doc, "WmsUi", "TestUiLastMessage")


def issue_row(sh, r, e, q, who):
    sh.getCellByPosition(11, r).setString(e)
    sh.getCellByPosition(4, r).setString(q)
    sh.getCellByPosition(7, r).setValue(46285)              # 20.09.2026
    sh.getCellByPosition(8, r).setString(who)


def return_row(sh, r, k, q):
    sh.getCellByPosition(1, r).setValue(k)
    if q:
        sh.getCellByPosition(5, r).setString(q)
    sh.getCellByPosition(7, r).setValue(46287)              # 22.09.2026


def dialog(o, doc):
    shown, back, lst = (o.basic(doc, "WmsReturnUi", "TestRetDialog").split("\n", 2) + ["", ""])[:3]
    return shown.split("\t"), (back.split("\t") if back else []), [x for x in lst.split("\n") if x]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    p = new_wms("gui_returns")
    o = Office(unique("gui4"), PROFILES, headless=False, display=a.display)
    try:
        ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
        doc = o.load(p, macros=4)
        iss = doc.Sheets.getByName("Выдачи")
        sh = doc.Sheets.getByName("Возврат")
        res = []
        for i, (e, who) in enumerate([("7", IVA), ("7", PET), ("10", IVA), ("10", PET)], start=1):
            issue_row(iss, i, e, "10", who)
            res.append(o.basic(doc, "WmsIssue", "IssuePostRow", i))
        for r, k in enumerate([1, 2, 3, 4], start=1):
            return_row(sh, r, k, "1")
            res.append(o.basic(doc, "WmsReturn", "ReturnPostRow", r))
        return_row(sh, 5, 1, "2")                            # entered, not posted yet (Иванов)
        o.basic(doc, "WmsReturn", "PreviewReturnRow", 5, False)
        step("4 выдачи и 4 возврата проведены, лист «Возврат» защищён", res == [f"OK:{i}" for i in range(1, 9)] and sh.isProtected(), str(res))
        fw, lst = open_filter(o, ut, doc, sh, 8)
        ent = entries(lst)
        step("окно автофильтра «От кого возвращено» открылось на защищённом листе", [t for _, t in ent] == [IVA, PET], str(ent))
        for cid, text in ent:
            if text == PET:
                lst.getChild(cid).executeAction("CLICK", ())
        press_ok(o, fw)
        vis = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)]
        step("снята галочка «Петров» → OK: строки Петрова скрыты", vis == [True, False, True, False, True], str(vis))
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Возврат.$B$6")
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        o.basic(doc, "WmsReturnUi", "BtnRetPost")
        msg = o.basic(doc, "WmsUi", "TestUiLastMessage")
        step("«Провести» для видимой строки при фильтре", msg == "OK:9" and sh.getCellByPosition(0, 5).getValue() == 5, msg)
        doc.store()
        vis_saved = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)]
        with zipfile.ZipFile(p) as z:
            nfilt = z.read("content.xml").decode("utf-8").count('table:visibility="filter"')
        step("сохранение при фильтре «Возврат»: строки сохранены показанными, фильтр восстановлен, книга не «изменена»",
             vis_saved == vis and not doc.isModified() and nfilt == 0, f"{vis_saved}; строк, сохранённых скрытыми, {nfilt}")
        o.dispatch(doc, ".uno:CloseDoc")
        doc = o.load(p, macros=4)
        sh = doc.Sheets.getByName("Возврат")
        vis_reopen = [sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)]
        step("повторное открытие: фильтр «Возврат» возвращён, книга не «изменена»", vis_reopen == vis and not doc.isModified(), str(vis_reopen))
        fw, lst = open_filter(o, ut, doc, sh, 8)
        for cid, text in entries(lst):
            if text == PET:
                lst.getChild(cid).executeAction("CLICK", ())
        press_ok(o, fw)
        step("галочка возвращена → все строки видны", all(sh.getRows().getByIndex(r).IsVisible for r in range(1, 6)), "")
        # the real windows (interactive mode: no preset values)
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        sysh = doc.Sheets.getByName("_SYS")
        nx = sysh.getCellByPosition(1, 7).getValue()
        sh.getCellByPosition(1, 6).setValue(3)                 # row 7: issue № 3, nothing else yet
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Возврат.$B$7")
        title, msg = run_dialog(o, ut, doc, "BtnRetFind", "Найти выдачу", "cancel")
        step("окно «Найти выдачу» открывается кнопкой; «Отмена» — строка не тронута", title is not None and msg.startswith("SKIP:окно закрыто")
             and sh.getCellByPosition(7, 6).getString() == "" and sysh.getCellByPosition(1, 7).getValue() == nx, f"{title}; {msg}")
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Возврат.$B$7")
        title, msg = run_dialog(o, ut, doc, "BtnRetFind", "Найти выдачу", "ok")
        shown, back, lst = dialog(o, doc)
        today = datetime.date.today()
        step("«Заполнить строку» без ввода: окно предложило сегодняшнюю дату и текущее место ЕИ и вернуло их как предложено; строка заполнена "
             "(дата видна в H), ничего не проведено",
             title is not None and msg.startswith("OK:") and shown[1] == today.strftime("%d.%m.%Y") and shown[2] == "S-10" and back == shown and len(lst) == 1
             and sh.getCellByPosition(7, 6).getValue() == (today - datetime.date(1899, 12, 30)).days and sh.getCellByPosition(0, 6).getString() == ""
             and sysh.getCellByPosition(1, 7).getValue() == nx, f"{title}; {msg[:100]}; предложено {shown}; список {lst}")
        sh.getCellByPosition(2, 7).setString("7")               # row 8: the EI only — the window lists its issues
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Возврат.$C$8")
        title, msg = run_dialog(o, ut, doc, "BtnRetFind", "Найти выдачу", "ok")
        shown, back, lst = dialog(o, doc)
        step("по ЕИ в C: окно со списком выдач ЕИ-7 (две), выбрана первая — в строку записан её №",
             title is not None and len(lst) == 2 and back[3] == "0" and sh.getCellByPosition(1, 7).getValue() == 1, f"{title}; {lst}; {back}")
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Возврат.$B$9")
        title, msg = run_dialog(o, ut, doc, "BtnRetFind", "найти выдачу", "cancel")
        step("пустая строка: вопрос «ЕИ» — «Отмена» ничего не делает", title is not None and msg.startswith("SKIP:ЕИ не указан"), f"{title}; {msg}")
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Возврат.$B$2")
        title, msg = run_dialog(o, ut, doc, "BtnRetFix", "Исправить возврат", "ok")
        shown, back, _ = dialog(o, doc)
        step("окно «Исправить» заполнено проведёнными значениями: «OK» без правок — «ничего не изменилось»",
             title is not None and msg.startswith("SKIP:ничего не изменилось") and shown[:3] == ["1", "22.09.2026", "S-07"] and back == shown,
             f"{title}; {msg}; предложено {shown}")
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Возврат.$B$2")
        title, msg = run_dialog(o, ut, doc, "BtnRetDelete", "WMS", "close", exact=True)
        step("подтверждение «Удалить» — «Нет»: возврат не удалён", title is not None and msg.startswith("SKIP:удаление не подтверждено")
             and sh.getCellByPosition(13, 1).getString() == "Проведено", f"{title}; {msg}")
        # «Удалить» of a receipt (D-069): the refusal first, the confirmation only for a storno that is allowed
        o.basic(doc, "WmsUi", "TestUiAuto", 1)
        orders, iss = doc.Sheets.getByName("Заказы"), doc.Sheets.getByName("Выдачи")
        rcv = doc.Sheets.getByName("_RCV")
        for col, v in ((0, "З-GUI-1"), (1, "Товар GUI"), (7, "10"), (8, "шт"), (5, "10"), (20, "A-1")):
            orders.getCellByPosition(col, 1).setString(v)
        orders.getCellByPosition(13, 1).setValue(46285)         # N 20.09.2026
        rr = [o.basic(doc, "WmsReceipt", "ReceiptPostRow", 1)]
        issue_row(iss, 5, "201", "2", IVA)
        rr.append(o.basic(doc, "WmsIssue", "IssuePostRow", 5))
        return_row(sh, 10, 5, "1")
        rr.append(o.basic(doc, "WmsReturn", "ReturnPostRow", 10))
        o.basic(doc, "WmsUi", "TestUiAuto", 0)
        k0 = o.basic(doc, "WmsUi", "TestUiConfirmCount")
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Заказы.$B$2")
        kids, msg = run_msg(o, ut, doc, "BtnRcvDelete", "WmsOrdersUi")
        k1 = o.basic(doc, "WmsUi", "TestUiConfirmCount")
        step("«Удалить» прихода, по ЕИ которого есть выдача и возврат: первое окно — отказ (только «OK»), окна подтверждения "
             "«Удалить приход?» нет; приход не удалён",
             all(x.startswith("OK") for x in rr) and "ok" in kids and "close" not in kids and msg.startswith("ERR:По этому ЕИ") and k1 == k0
             and rcv.getCellByPosition(4, 201).getString() == "LIVE", f"{rr}; кнопки окна {kids}; {msg[:90]}; подтверждений {k1 - k0}")
        rd = [o.basic(doc, "WmsReturn", "ReturnDeleteRow", 10), o.basic(doc, "WmsIssue", "IssueDeleteRow", 5)]
        o.dispatch(doc, ".uno:GoToCell", ToPoint="$Заказы.$B$2")
        kids2, msg2 = run_msg(o, ut, doc, "BtnRcvDelete", "WmsOrdersUi")
        k2 = o.basic(doc, "WmsUi", "TestUiConfirmCount")
        step("после сторно возврата и выдачи сторно допустимо: окно подтверждения «Удалить приход …?» («Да» / «Нет»); «Нет» — приход не удалён",
             all(x.startswith("OK") for x in rd) and "close" in kids2 and "ok" not in kids2 and msg2.startswith("Удалить приход ЕИ-00000201?")
             and k2 == k1 + 1 and rcv.getCellByPosition(4, 201).getString() == "LIVE", f"{rd}; кнопки окна {kids2}; {msg2[:40]}; подтверждений {k2 - k1}")
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
