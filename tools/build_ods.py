"""Assemble a WMS workbook from the text sources (MASTER SPEC v0.3 §3: Basic sources are text, the ODS is a build artefact).

    python3 tools/build_ods.py OUT.ods            production book: «Главная», «Выдачи», «Наличие», «Получатели», _SYS
    python3 tools/build_ods.py OUT.ods --test     test book: + test module, the synthetic _TST sheet, synthetic EIs and
                                                  recipients, MODE=TEST

Needs LibreOffice with Python-UNO on the developer machine. Nothing is installed on the warehouse PC.
"""
import argparse
import os
import random
import sys
import tempfile
import zipfile
import uno
from com.sun.star.awt import Point, Size
from com.sun.star.lang import Locale
from com.sun.star.script import ScriptEventDescriptor
from com.sun.star.table import CellRangeAddress
from com.sun.star.util import CellProtection

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wmslo import Office, add_basic, bind_event, props  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src", "basic")
TEST_SRC = os.path.join(ROOT, "tests", "basic")
PWD = "wms"                      # WmsConfig.PROTECT_PWD
SYS_KEYS = ["SCHEMA", "INSTANCE_ID", "MODE", "CORE_VERSION", "LAST_SEQ", "NEXT_EI", "NEXT_NO", "NEXT_RET",
            "JOURNAL_POS", "REGISTERED_URL", "TX_STATE", "TX_SEQ", "TX_TYPE", "TX_TIME", "TX_BEFORE_IMAGE",
            "SAVE_STAMP", "SAVE_SEQ", "MAX_QTY", "KEY_SHEETS"]        # WmsConfig.SysKeyNames()
SK_SAVE_STAMP = SYS_KEYS.index("SAVE_STAMP")
MODULES = ["WmsConfig", "WmsCore", "WmsJournal", "WmsLock", "WmsBackup", "WmsRecovery", "WmsDiagnostics", "WmsIssue", "WmsUi"]
TEST_MODULES = ["WmsTestOps"]
TEST_SHEET = "_TST"
LAST_ROW = 1048575

# «Выдачи» A:R — the fixed user interface (MASTER SPEC v0.3 §2; WmsConfig IC_*)
ISSUES = "Выдачи"
ISSUE_HEADERS = ["№", "№ документа выдачи", "Наименование", "Артикул", "Количество", "Количество в %", "Единица измерения",
                 "Дата выдачи", "Кому выдано", "Место хранения", "Категория", "Внутренний код", "Вернул", "Вернул в %",
                 "Комментарий", "Остаток до выдачи", "Остаток после выдачи", "Контроль"]
ISSUE_WIDTHS = [1500, 2600, 6000, 3000, 2300, 2300, 2300, 2600, 4800, 2800, 3000, 3300, 2000, 2000, 4500, 2500, 2500, 7000]
ISSUE_LOCKS_OPEN = "101100100110000111"      # WmsConfig.ISSUE_LOCKS_OPEN: unposted rows, inputs B E F H I L M N O open
# «Наличие» — the permanent EI registry, row index = EI number (spec §5, v0.1 §24/§37; WmsConfig SC_*)
STOCK = "Наличие"
STOCK_HEADERS = ["ЕИ", "Наименование", "Артикул", "Единица измерения", "Остаток", "Место хранения", "Категория", "Состояние", "Источник"]
STOCK_WIDTHS = [3300, 6500, 3200, 2300, 2300, 2800, 3000, 2500, 3000]
RCPT = "Получатели"
MAIN = "Главная"
# _SYS KEY_SHEETS: sheet|keyCol|ctlCol|counterRow|inputCol — the full key check after a save without WMS (spec §15)
ISSUE_KEY_SPEC = f"{ISSUES}|0|17|{SYS_KEYS.index('NEXT_NO')}|11"
# _TST (Phase 1 test sheet): A key (issued from NEXT_NO), B text (input), C quantity (text input), D control/status
TEST_KEY_SPEC = f"{TEST_SHEET}|0|3|{SYS_KEYS.index('NEXT_NO')}|1"

TEST_RECIPIENTS = [("егр", "Ермолин Егор Павлович"), ("офис", "Офис"), ("пр", "Производство"),
                   ("ива", "Иванов Иван Андреевич"), ("пет", "Петров Пётр Сергеевич")]


def synthetic_registry(n=200):
    """Deterministic EIs of the test book (the test oracle knows their initial balances from here)."""
    units = ["шт", "м", "кг", "л", "упак"]
    cats = ["Крепёж", "Кабель", "Электрика", "СИЗ", "Химия", "Инструмент"]
    special = {
        1: ("Болт М8х30 оцинкованный", "DIN933-M8x30", "шт", 100.0, "A-01-1", "Крепёж"),
        2: ("Кабель ВВГнг 3х2,5", "ВВГ-3x2.5", "м", 250.5, "B-02", "Кабель"),
        3: ("Лампа LED E27 10W", "LED-E27-10", "шт", 0.0, "C-03", "Электрика"),
        4: ("Перчатки нитриловые", "PN-9", "пар", 12.0, "D-01", "СИЗ"),
        5: ("Краска белая 0,9 л", "KR-W-09", "шт", 5.0, "E-02", "Химия"),
        6: ("Смазка литиевая", "LIT-400", "кг", 3.75, "E-03", "Химия"),
    }
    rows = []
    for i in range(1, n + 1):
        if i in special:
            name, art, unit, qty, place, cat = special[i]
        else:
            name, art, unit = f"Товар {i}", f"ART-{i:05d}", units[i % 5]
            qty, place, cat = float((i * 7) % 50 + 1), f"S-{i % 20:02d}", cats[i % 6]
        rows.append((f"ЕИ-{i:08d}", name, art, unit, qty, place, cat, "", "тест"))
    return rows


def read_source(path):
    s = open(path, encoding="utf-8").read()
    return s.replace("\r\n", "\n")


def number_format(doc, code, lang=("en", "US")):
    nf = doc.NumberFormats
    loc = Locale(lang[0], lang[1], "")
    k = nf.queryKey(code, loc, False)
    if k == -1:
        k = nf.addNew(code, loc)
    return k


def protection(locked):
    p = CellProtection()
    p.IsLocked = locked
    return p


def header(sheet, titles, widths, height=None):
    for c, (t, w) in enumerate(zip(titles, widths)):
        cell = sheet.getCellByPosition(c, 0)
        cell.setString(t)
        sheet.getColumns().getByIndex(c).Width = w
    rng = sheet.getCellRangeByPosition(0, 0, len(titles) - 1, 0)
    rng.CharWeight = 150
    rng.IsTextWrapped = True
    rng.VertJustify = 3            # bottom: the buttons sit in the upper part of the header row
    rng.CellBackColor = 0xE7E6E6
    if height:
        sheet.getRows().getByIndex(0).Height = height


def autofilter(doc, sheet_index, name, ncols):
    a = CellRangeAddress()
    a.Sheet, a.StartColumn, a.StartRow, a.EndColumn, a.EndRow = sheet_index, 0, 0, ncols - 1, LAST_ROW
    doc.DatabaseRanges.addNewByName(name, a)
    doc.DatabaseRanges.getByName(name).AutoFilter = True


def add_button(doc, sheet, name, label, macro, x, y, w, h):
    """A form push button bound to a Basic macro (no focus on click: the cell cursor stays on the user's row)."""
    model = doc.createInstance("com.sun.star.form.component.CommandButton")
    model.Name = name
    model.Label = label
    model.FocusOnClick = False
    model.Printable = False
    shape = doc.createInstance("com.sun.star.drawing.ControlShape")
    shape.setPosition(Point(x, y))
    shape.setSize(Size(w, h))
    shape.setControl(model)
    sheet.getDrawPage().add(shape)
    form = model.getParent()
    idx = [form.getByIndex(i).Name for i in range(form.getCount())].index(name)
    ev = ScriptEventDescriptor()
    ev.ListenerType = "XActionListener"
    ev.EventMethod = "actionPerformed"
    ev.AddListenerParam = ""
    ev.ScriptType = "Script"
    ev.ScriptCode = f"vnd.sun.star.script:Standard.{macro}?language=Basic&location=document"
    form.registerScriptEvent(idx, ev)


def build_issues(doc, sh):
    header(sh, ISSUE_HEADERS, ISSUE_WIDTHS, height=1700)
    text = number_format(doc, "@", ("ru", "RU"))
    date = number_format(doc, "DD.MM.YYYY")
    whole = number_format(doc, "0")
    sh.getCellRangeByPosition(0, 1, 0, LAST_ROW).NumberFormat = whole
    for c in (4, 11):                          # E quantity, L EI: text input (spec §12, locale-independent parsing)
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
    sh.getCellRangeByPosition(7, 1, 7, LAST_ROW).NumberFormat = date
    for c, bit in enumerate(ISSUE_LOCKS_OPEN):
        if bit == "0":
            sh.getCellRangeByPosition(c, 1, c, LAST_ROW).CellProtection = protection(False)
    x = 150
    for name, label, macro in (("btnPost", "Провести", "WmsUi.BtnPost"), ("btnFix", "Исправить", "WmsUi.BtnFix"),
                               ("btnDelete", "Удалить", "WmsUi.BtnDelete"), ("btnClear", "Очистить", "WmsUi.BtnClear")):
        add_button(doc, sh, name, label, macro, x, 120, 2750, 650)
        x += 2900


def build_stock(doc, sh, rows):
    header(sh, STOCK_HEADERS, STOCK_WIDTHS)
    sh.getCellRangeByPosition(0, 1, 0, LAST_ROW).NumberFormat = number_format(doc, "@", ("ru", "RU"))
    if rows:
        sh.getCellRangeByPosition(0, 1, len(STOCK_HEADERS) - 1, len(rows)).setDataArray(rows)


def build_recipients(sh, rows):
    header(sh, ["Сокращение", "Получатель"], [3500, 8000])
    if rows:
        sh.getCellRangeByPosition(0, 1, 1, len(rows)).setDataArray(tuple(rows))


def build_main(doc, sh):
    sh.getColumns().getByIndex(0).Width = 5200
    sh.getColumns().getByIndex(1).Width = 17500
    t = sh.getCellByPosition(0, 0)
    t.setString("WMS — состояние и восстановление")
    t.CharWeight = 150
    t.CharHeight = 16
    captions = ["Состояние", "Причина", "Что сделать", "Непроведённые выдачи", "Заметки запуска", "Последнее действие", "Обновлено"]
    for i, cap in enumerate(captions):
        c = sh.getCellByPosition(0, 2 + i)
        c.setString(cap)
        c.CharWeight = 150
    sh.getCellByPosition(1, 2).setString("WMS ещё не запускалась в этой книге — откройте её с включёнными макросами")
    vals = sh.getCellRangeByPosition(0, 2, 1, 8)
    vals.IsTextWrapped = True
    vals.VertJustify = 1           # top
    y = sh.getCellByPosition(0, 10).Position.Y
    buttons = (("btnRecover", "Восстановить", "WmsUi.BtnRecover"), ("btnAbandon", "Отложить хвост журнала", "WmsUi.BtnAbandon"),
               ("btnUnlock", "Снять блокировку WMS", "WmsUi.BtnUnlock"), ("btnRegister", "Сделать рабочим файлом", "WmsUi.BtnRegister"),
               ("btnSelfCheck", "Самопроверка", "WmsUi.BtnSelfCheck"), ("btnBackup", "Резервная копия", "WmsUi.BtnBackup"))
    for i, (name, label, macro) in enumerate(buttons):
        add_button(doc, sh, name, label, macro, 200 + (i % 2) * 6700, y + (i // 2) * 1000, 6400, 800)
    help_row = 15
    sh.getCellByPosition(0, help_row).setString(
        "Выдача: на листе «Выдачи» введите ЕИ в «Внутренний код», количество, дату и получателя и нажмите «Провести». "
        "«Исправить» и «Удалить» работают для проведённой строки под курсором, «Очистить» — для непроведённой строки или копии.")
    sh.getCellRangeByPosition(0, help_row, 1, help_row).merge(True)
    sh.getCellByPosition(0, help_row).IsTextWrapped = True
    sh.getRows().getByIndex(help_row).Height = 1500


def patch_content(path, fn):
    """Rewrite content.xml of an ODS through fn(str) -> str (mimetype stays first and stored)."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "content.xml":
                data = fn(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data, compress_type=zipfile.ZIP_STORED if item.filename == "mimetype" else zipfile.ZIP_DEFLATED)
    os.replace(tmp, path)


def allow_insert_rows(xml, sheet_name):
    """Sheet protection option «insert rows» (spec §16; LibreOffice has no API for it — P1 used the same XML edit)."""
    head = f'<table:table table:name="{sheet_name}"'
    i = xml.index(head)
    j = xml.index(">", i) + 1
    old = '<loext:table-protection loext:select-protected-cells="true" loext:select-unprotected-cells="true"/>'
    if not xml.startswith(old, j):
        raise RuntimeError(f"unexpected protection XML for {sheet_name}: {xml[j:j + 160]}")
    new = old.replace('/>', ' loext:insert-rows="true"/>')
    return xml[:j] + new + xml[j + len(old):]


def build(office, out_path, test=False, instance_id=None):
    out_path = os.path.abspath(out_path)
    doc = office.new_calc()
    sheets = doc.Sheets
    main = sheets.getByIndex(0)
    main.Name = MAIN
    for i, name in enumerate((ISSUES, STOCK, RCPT, "_SYS"), start=1):
        sheets.insertNewByName(name, i)
    sys_sh = sheets.getByName("_SYS")
    values = {
        "SCHEMA": "WMS-SYS-1", "INSTANCE_ID": instance_id or "%016X" % random.getrandbits(64), "MODE": "TEST" if test else "PROD",
        "CORE_VERSION": "", "LAST_SEQ": 0, "NEXT_EI": 1, "NEXT_NO": 1, "NEXT_RET": 1, "JOURNAL_POS": "", "REGISTERED_URL": "",
        "TX_STATE": "NONE", "TX_SEQ": 0, "TX_TYPE": "", "TX_TIME": "", "TX_BEFORE_IMAGE": "", "SAVE_STAMP": 0, "SAVE_SEQ": 0,
        "MAX_QTY": 100000, "KEY_SHEETS": ISSUE_KEY_SPEC + (";" + TEST_KEY_SPEC if test else ""),
    }
    registry = synthetic_registry() if test else []
    if registry:
        values["NEXT_EI"] = len(registry) + 1
    for i, k in enumerate(SYS_KEYS):
        sys_sh.getCellByPosition(0, i).setString(k)
        v = values[k]
        if isinstance(v, (int, float)):
            sys_sh.getCellByPosition(1, i).setValue(v)
        else:
            sys_sh.getCellByPosition(1, i).setString(v)
    build_main(doc, main)
    build_issues(doc, sheets.getByName(ISSUES))
    build_stock(doc, sheets.getByName(STOCK), registry)
    build_recipients(sheets.getByName(RCPT), TEST_RECIPIENTS if test else [])
    autofilter(doc, 1, "WMS_ISSUES", len(ISSUE_HEADERS))
    autofilter(doc, 2, "WMS_STOCK", len(STOCK_HEADERS))
    if test:
        sheets.insertNewByName(TEST_SHEET, sheets.getCount())
        tst = sheets.getByName(TEST_SHEET)
        for c, h in enumerate(["№", "Текст", "Количество", "Статус"]):
            tst.getCellByPosition(c, 0).setString(h)
        tst.getCellRangeByPosition(0, 1, 3, LAST_ROW).CellProtection = protection(False)
        # quantity is entered as text and parsed by WMS (spec §12)
        tst.getCellRangeByPosition(2, 1, 2, LAST_ROW).NumberFormat = number_format(doc, "@", ("ru", "RU"))
        tst.protect(PWD)
    names = MODULES + (TEST_MODULES if test else [])
    for m in names:
        folder = TEST_SRC if m in TEST_MODULES else SRC
        add_basic(doc, m, read_source(os.path.join(folder, m + ".bas")))
    bind_event(doc.Events, "OnLoad", "WmsCore.OnDocLoad")
    bind_event(doc.Events, "OnPrepareUnload", "WmsCore.OnDocPrepareUnload")
    bind_event(doc.Events, "OnSave", "WmsCore.OnDocSave")
    bind_event(doc.Events, "OnSaveDone", "WmsCore.OnDocSaveDone")
    bind_event(doc.Events, "OnSaveFailed", "WmsCore.OnDocSaveDone")
    bind_event(sheets.getByName(ISSUES).Events, "OnChange", "WmsIssue.OnIssuesChange")
    bind_event(sheets.getByName(RCPT).Events, "OnChange", "WmsIssue.OnRecipientsChange")
    ctl = doc.getCurrentController()
    for name in (ISSUES, STOCK, RCPT):
        ctl.setActiveSheet(sheets.getByName(name))
        ctl.freezeAtPosition(0, 1)
    ctl.setActiveSheet(sheets.getByName(ISSUES))
    ctl.select(sheets.getByName(ISSUES).getCellByPosition(11, 1))
    sys_sh.IsVisible = False
    for name in (MAIN, ISSUES, STOCK, "_SYS"):
        sheets.getByName(name).protect(PWD)
    doc.protect(PWD)
    if os.path.exists(out_path):
        os.remove(out_path)
    doc.storeAsURL(uno.systemPathToFileUrl(out_path), props(FilterName="calc8"))
    doc.close(True)
    patch_content(out_path, lambda x: allow_insert_rows(x, ISSUES))
    # the saved file must carry the save stamp WMS expects (EditingCycles after this save), otherwise the first
    # start would report "saved without WMS"; fix it up with macros disabled so no event interferes
    doc = office.load(out_path, macros=0, hidden=True)
    sys_sh = doc.Sheets.getByName("_SYS")
    ec = doc.getDocumentProperties().EditingCycles
    sys_sh.getCellByPosition(1, SK_SAVE_STAMP).setValue(ec + 1)
    doc.getCurrentController().setActiveSheet(doc.Sheets.getByName(ISSUES))     # the book opens on «Выдачи»
    doc.store()
    doc.close(True)
    doc = office.load(out_path, macros=0, hidden=True)
    ec = doc.getDocumentProperties().EditingCycles
    stamp = doc.Sheets.getByName("_SYS").getCellByPosition(1, SK_SAVE_STAMP).getValue()
    doc.close(True)
    if stamp != ec:
        raise RuntimeError(f"save stamp {stamp} != EditingCycles {ec}")
    with zipfile.ZipFile(out_path) as z:
        xml = z.read("content.xml").decode("utf-8")
    if 'loext:insert-rows="true"' not in xml:
        raise RuntimeError("the «insert rows» protection option of «Выдачи» did not survive the save")
    return out_path


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out")
    ap.add_argument("--test", action="store_true", help="test book (test module, _TST sheet, synthetic EIs, MODE=TEST)")
    a = ap.parse_args()
    prof = tempfile.mkdtemp(prefix="wms_build_")
    o = Office("wmsbuild%d" % os.getpid(), prof)
    try:
        print("built", build(o, a.out, test=a.test))
    finally:
        o.terminate()


if __name__ == "__main__":
    main()
