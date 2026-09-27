"""GUI check of the release candidate (FINAL WMS MARATHON §5) on a virtual screen (Xvfb), used like a user.

    Xvfb :99 &  WMS_TEST_OUT=/tmp/wms_gui python3 tests/gui_candidate.py [--display :99]

A production book (MODE PROD, never opened) opened with WMS in a real window:
1. «Главная»: the version next to the title; the buttons «Состояние системы» and «Проверка перед работой».
2. «Проверка перед работой»: a window «WMS» with only «OK» names the result («ошибок 0»); the section of «Главная» lists the
   checks.
3. «Состояние системы»: no window; the section shows the versions, the working file, LibreOffice, the journal, the backups.
4. «Справка»: a visible protected sheet with the guide of the storekeeper, the release notes and the backup / recovery guide.
Prints one line per step; exits 0 when every step passed.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
from harness import new_wms, PROFILES, unique, wms_versions  # noqa: E402
from wmslo import Office  # noqa: E402
from gui_phase4 import step, run_msg, RESULTS  # noqa: E402

STATUS_ROW = 24                                   # WmsStatus.STATUS_ROW


def section(doc):
    rows = doc.Sheets.getByName("Главная").getCellRangeByPosition(0, STATUS_ROW, 1, STATUS_ROW + 30).getDataArray()
    return [(a, b) for a, b in rows if a or b]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    p = new_wms("gui_candidate", kind="PROD", book_name="WMS_PROD_CANDIDATE.ods")
    o = Office(unique("guicand"), PROFILES, headless=False, display=a.display)
    try:
        ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
        doc = o.load(p, macros=4)
        main_sh = doc.Sheets.getByName("Главная")
        ver = main_sh.getCellByPosition(1, 0).getString()
        form = main_sh.getDrawPage().getForms().getByIndex(0)
        labels = [form.getByIndex(i).Label for i in range(form.getCount())]
        state = main_sh.getCellByPosition(1, 2).getString()
        step("«Главная» рабочей книги: «Работа разрешена», версия рядом с заголовком, кнопки «Состояние системы» и «Проверка перед работой»",
             state == "Работа разрешена" and ver.startswith("версия " + wms_versions()["WMS_PRODUCT_VERSION"]) and "Состояние системы" in labels and "Проверка перед работой" in labels,
             f"{state}; «{ver}»; {labels}")
        kids, msg = run_msg(o, ut, doc, "BtnPreWorkCheck", "WmsStatus")
        sec = section(doc)
        names = [x[0] for x in sec[1:]]
        step("«Проверка перед работой»: окно «WMS» (только «OK») с итогом «ошибок 0»; на «Главной» — раздел с проверками",
             kids == ["ok"] and msg.startswith("ПРОВЕРКА ПЕРЕД РАБОТОЙ: ошибок 0") and sec[0][0] == "ПРОВЕРКА ПЕРЕД РАБОТОЙ"
             and {"WMS", "LibreOffice", "журнал", "резервные копии"} <= set(names), f"кнопки {kids}; {msg[:90]}; {names}")
        o.basic(doc, "WmsStatus", "BtnSystemStatus")
        sec = dict(section(doc)[1:])
        head = section(doc)[0][0]
        step("«Состояние системы»: раздел без окна — версии, рабочий файл (зарегистрирован), LibreOffice, журнал, резервные копии",
             head == "СОСТОЯНИЕ СИСТЕМЫ" and wms_versions()["WMS_CORE_VERSION"] in sec.get("WMS", "") and "(зарегистрирован)" in sec.get("рабочий файл", "")
             and "LibreOffice" in sec and "журнал" in sec and "резервные копии" in sec, f"{head}; {list(sec)}")
        hlp = doc.Sheets.getByName("Справка")
        doc.getCurrentController().setActiveSheet(hlp)
        cur = hlp.createCursor()
        cur.gotoEndOfUsedArea(False)
        col = [r[0] for r in hlp.getCellRangeByPosition(0, 0, 0, cur.getRangeAddress().EndRow).getDataArray()]
        heads = [x for x in col if x.startswith("WMS — ")]
        step("«Справка»: видимый защищённый лист — руководство кладовщика, примечания к выпуску, резервные копии и восстановление",
             hlp.IsVisible and hlp.isProtected() and heads == ["WMS — руководство кладовщика", "WMS — примечания к выпуску",
                                                               "WMS — резервные копии и восстановление"], str(heads))
        doc.setModified(False)
        o.dispatch(doc, ".uno:CloseDoc")
    finally:
        o.terminate()
    ok = bool(RESULTS) and all(RESULTS)
    print(f"GUI: {sum(RESULTS)} из {len(RESULTS)} PASS", flush=True)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
