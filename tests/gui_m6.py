"""GUI check of «Приход авто» and of the ways from «Главная» on a virtual screen (Xvfb), used like a user.

    Xvfb :99 &  WMS_TEST_OUT=/tmp/wms_gui python3 tests/gui_m6.py [--display :99]

1. The panel: the rows 0..5 (the input cells, the indicators, the header) are frozen — the table scrolls under them; the
   buttons ПРИЕХАЛ, УЕХАЛ, «Найти», «Сегодня», «Все», «Исправить», «Отменить визит» lie in the frozen rows and are bound
   to their macros; the two input cells offer the lists of the known vehicles and suppliers (validation lists, AutoInput
   stays off); the threshold cell and the comment column are open, everything else is locked.
2. The real windows (in the headless tests they are replaced by preset values): ПРИЕХАЛ with an empty cell — a message;
   УЕХАЛ with the cursor in the panel and two vehicles on the territory — the short list («Отмена» changes nothing,
   «Уехала» records the chosen one); with one vehicle — the question «Уехала машина …?» («Нет» changes nothing);
   «Исправить» opens with the values of the visit («Исправить» without changes — «изменений нет»).
3. The AutoFilter of the table (header in row 6) opens on the protected sheet and filters by the status.
4. «Главная»: the seven ways to the working sheets; «Приход авто» opens with the cursor in «Марка / госномер».
Prints one line per step; exits 0 when every step passed.
"""
import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
from harness import new_wms, PROFILES, unique  # noqa: E402
from wmslo import Office, props  # noqa: E402
from gui_phase4 import step, state, run_dialog, run_msg, press_ok, entries, RESULTS  # noqa: E402

CARS = "Приход авто"


def buttons(doc, sheet):
    """name → (label, macro, top row of the shape)"""
    dp = sheet.getDrawPage()
    out = {}
    for i in range(dp.getCount()):
        shp = dp.getByIndex(i)
        try:
            m = shp.getControl()
        except Exception:
            continue
        form = m.getParent()
        idx = [form.getByIndex(k).Name for k in range(form.getCount())].index(m.Name)
        ev = form.getScriptEvents(idx)
        macro = ev[0].ScriptCode.split("Standard.")[1].split("?")[0] if ev else ""
        y = shp.getPosition().Y
        row = 0
        while sheet.getCellByPosition(0, row + 1).Position.Y <= y:
            row += 1
        out[m.Name] = (m.Label, macro, row)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    p = new_wms("gui_m6")
    o = Office(unique("guim6"), PROFILES, headless=False, display=a.display)
    try:
        ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
        doc = o.load(p, macros=4)
        sh = doc.Sheets.getByName(CARS)
        ctl = doc.getCurrentController()
        ctl.setActiveSheet(sh)
        o.idle()
        step("лист «Приход авто»: закреплены строки 1–6 (панель и заголовок таблицы)", ctl.hasFrozenPanes() and ctl.getSplitRow() == 6,
             f"закрепление {ctl.hasFrozenPanes()}, строк {ctl.getSplitRow()}")
        b = buttons(doc, sh)
        want = {"btnCarArrive": ("ПРИЕХАЛ", "WmsCarUi.BtnCarArrive"), "btnCarDepart": ("УЕХАЛ", "WmsCarUi.BtnCarDepart"),
                "btnCarFind": ("Найти", "WmsCarUi.BtnCarFind"), "btnCarToday": ("Сегодня", "WmsCarUi.BtnCarToday"),
                "btnCarAll": ("Все", "WmsCarUi.BtnCarAll"), "btnCarFix": ("Исправить", "WmsCarUi.BtnCarFix"),
                "btnCarCancel": ("Отменить визит", "WmsCarUi.BtnCarCancel")}
        step("кнопки панели лежат в закреплённых строках и привязаны к своим макросам",
             all(b.get(k, ("", "", 99))[:2] == v and b[k][2] <= 2 for k, v in want.items()), str(b))
        c2, d2 = sh.getCellByPosition(2, 1), sh.getCellByPosition(3, 1)
        vl = [c.Validation for c in (c2, d2)]
        step("ячейки «Марка / госномер» и «Поставщик» открыты и предлагают списки известных машин и поставщиков (сортированные)",
             not c2.CellProtection.IsLocked and not d2.CellProtection.IsLocked and [v.Type.value for v in vl] == ["LIST", "LIST"]
             and vl[0].getFormula1().endswith("$N$2:$N$1048576") and vl[1].getFormula1().endswith("$P$2:$P$1048576") and vl[0].ShowList == 2,
             f"{[v.getFormula1() for v in vl]}")
        locks = [sh.getCellByPosition(c, 10).CellProtection.IsLocked for c in range(14)]
        step("строки таблицы защищены, открыт только «Комментарий» (M); порог L5 открыт", locks == [True] * 12 + [False, True]
             and not sh.getCellByPosition(11, 4).CellProtection.IsLocked and sh.isProtected(), str(locks))
        # ПРИЕХАЛ with an empty supplier → a message
        o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${CARS}.$C$2")
        o.dispatch(doc, ".uno:EnterString", StringName="Газель А123ВС")
        kids, msg = run_msg(o, ut, doc, "BtnCarArrive", "WmsCarUi")
        step("ПРИЕХАЛ без поставщика — окно с объяснением (только «OK»), визит не создан", kids == ["ok"] and "поставщика" in msg
             and sh.getCellByPosition(0, 6).getString() == "", f"{kids}; {msg}")
        o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${CARS}.$D$2")
        o.dispatch(doc, ".uno:EnterString", StringName="Петрович")
        o.basic(doc, "WmsCarUi", "BtnCarArrive")
        r1 = o.basic(doc, "WmsUi", "TestUiLastMessage")
        o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${CARS}.$C$2")
        o.dispatch(doc, ".uno:EnterString", StringName="КАМАЗ Х456ОР")
        o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${CARS}.$D$2")
        o.dispatch(doc, ".uno:EnterString", StringName="Ромашка")
        o.basic(doc, "WmsCarUi", "BtnCarArrive")
        r2 = o.basic(doc, "WmsUi", "TestUiLastMessage")
        sel = doc.getCurrentSelection().getRangeAddress()
        step("две машины приехали кнопкой: ячейки очищены, выделена новая строка", r1.startswith("OK:") and r2.startswith("OK:")
             and c2.getString() == "" and d2.getString() == "" and (sel.StartRow, sel.EndRow) == (7, 7), f"{r1}; {r2}; строка {sel.StartRow}")
        # УЕХАЛ from the panel: the list of the two vehicles
        o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${CARS}.$C$2")
        title, msg = run_dialog(o, ut, doc, "BtnCarDepart", "УЕХАЛ", "cancel", module="WmsCarUi")
        shown = o.basic(doc, "WmsCarUi", "TestCarShown").split("\n")
        step("УЕХАЛ с курсором в панели: окно выбора из двух машин; «Отмена» — ничего не записано", title is not None and len(shown) == 2
             and msg.startswith("CANCEL") and sh.getCellByPosition(7, 6).getString() == "На территории", f"{title}; {shown}; {msg}")
        o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${CARS}.$C$2")
        title, msg = run_dialog(o, ut, doc, "BtnCarDepart", "УЕХАЛ", "ok", module="WmsCarUi")
        step("окно выбора: «Уехала» — записан выезд первой в списке машины", msg.startswith("OK:") and sh.getCellByPosition(7, 6).getString() == "Уехал"
             and sh.getCellByPosition(7, 7).getString() == "На территории", f"{title}; {msg}")
        o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${CARS}.$C$2")
        kids, msg = run_msg(o, ut, doc, "BtnCarDepart", "WmsCarUi")
        step("одна машина на территории — вопрос «Уехала машина …?» («Да» / «Нет»); «Нет» — выезд не записан", "close" in kids
             and msg.startswith("CANCEL") and sh.getCellByPosition(7, 7).getString() == "На территории", f"{kids}; {msg}")
        o.dispatch(doc, ".uno:GoToCell", ToPoint=f"${CARS}.$C$7")
        title, msg = run_dialog(o, ut, doc, "BtnCarFix", "Исправить визит", "ok", module="WmsCarUi")
        shown = o.basic(doc, "WmsCarUi", "TestCarShown").split("\t")
        step("«Исправить»: окно с машиной, поставщиком и временем визита; без изменений — «изменений нет»",
             title is not None and shown[:2] == ["Газель А123ВС", "Петрович"] and "изменений нет" in msg, f"{title}; {shown}; {msg}")
        # the AutoFilter of the table: header row 6
        ctl.setActiveSheet(sh)
        o.idle()
        time.sleep(0.5)
        ut.getTopFocusWindow().getChild("grid_window").executeAction("LAUNCH", props(AUTOFILTER="", COL="7", ROW="5"))
        o.idle()
        time.sleep(0.3)
        fw = ut.getFloatWindow()
        lst = fw.getChild("check_list_box")
        ent = entries(lst)
        for cid, text in ent:
            if text == "Уехал":
                lst.getChild(cid).executeAction("CLICK", ())
        press_ok(o, fw)
        vis = [sh.getRows().getByIndex(r).IsVisible for r in (5, 6, 7)]
        step("автофильтр «Статус» на защищённом листе: снята галочка «Уехал» — уехавшая машина скрыта", {"Уехал", "На территории"} <= {t for _, t in ent}
             and vis == [True, False, True], f"{ent}; {vis}")
        o.basic(doc, "WmsCarUi", "BtnCarAll")
        # «Главная»: the ways
        mb = buttons(doc, doc.Sheets.getByName("Главная"))
        nav = {"navOrders": "Заказы", "navSpecial": "Иной приход", "navCars": "Приход авто", "navIssues": "Выдачи", "navReturns": "Возврат",
               "navAdjust": "Корректировки", "navStock": "Наличие"}
        step("«Главная»: семь заметных переходов (Заказы, Иной приход, Приход авто, Выдачи, Возврат, Корректировки, Наличие) в строке 2",
             all(mb.get(k, ("",))[0] == v and mb[k][2] == 1 for k, v in nav.items()), str({k: mb.get(k) for k in nav}))
        ctl.setActiveSheet(doc.Sheets.getByName("Главная"))
        o.basic(doc, "WmsUi", "NavCars")
        a_ = doc.getCurrentSelection().getRangeAddress()
        s1 = ctl.getActiveSheet().getName()
        o.basic(doc, "WmsUi", "NavIssues")
        a2 = doc.getCurrentSelection().getRangeAddress()
        step("переход «Приход авто» — курсор в «Марка / госномер»; «Выдачи» — первая свободная строка, колонка ЕИ",
             s1 == CARS and (a_.StartColumn, a_.StartRow) == (2, 1) and ctl.getActiveSheet().getName() == "Выдачи" and (a2.StartColumn, a2.StartRow) == (11, 1),
             f"{s1} {a_.StartColumn}:{a_.StartRow}; {ctl.getActiveSheet().getName()} {a2.StartColumn}:{a2.StartRow}")
        doc.setModified(False)
    finally:
        o.terminate()
    print(f"GUI: {sum(RESULTS)} из {len(RESULTS)} PASS", flush=True)
    sys.exit(0 if all(RESULTS) else 1)


if __name__ == "__main__":
    main()
