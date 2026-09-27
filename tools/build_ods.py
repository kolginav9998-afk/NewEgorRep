"""Assemble a WMS workbook from the text sources (MASTER SPEC v0.3 §3: Basic sources are text, the ODS is a build artefact).

    python3 tools/build_ods.py OUT.ods            production book: «Главная», «Заказы», «Иной приход», «Выдачи», «Возврат»,
                                                  «Корректировки», «Наличие», «Получатели», service sheets _ORD, _RCV, _SPR,
                                                  _ART, _RET, _ISS, _ADJ, _IDX, _SYS
    python3 tools/build_ods.py OUT.ods --test     test book: + test module, the synthetic _TST sheet, synthetic EIs and
                                                  recipients, MODE=TEST

Needs LibreOffice with Python-UNO on the developer machine. Nothing is installed on the warehouse PC.
"""
import argparse
import os
import random
import re
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
            "SAVE_STAMP", "SAVE_SEQ", "MAX_QTY", "KEY_SHEETS",
            "NEXT_SPL", "NEXT_OFF", "NEXT_PROD", "NEXT_DET", "NEXT_OLD", "NEXT_OTH", "NEXT_ADJ"]   # WmsConfig.SysKeyNames() (WMS-SYS-3)
SCHEMA = "WMS-SYS-3"                                                                  # WmsConfig.WMS_SYS_SCHEMA
SK_SAVE_STAMP = SYS_KEYS.index("SAVE_STAMP")
MODULES = ["WmsConfig", "WmsCore", "WmsJournal", "WmsLock", "WmsBackup", "WmsRecovery", "WmsDiagnostics", "WmsIssue", "WmsUi",
           "WmsOrders", "WmsReceipt", "WmsOrdersUi", "WmsReturn", "WmsReturnUi", "WmsSpecial", "WmsSpecialUi", "WmsAdjust", "WmsAdjustUi",
           "WmsExport", "WmsMigrate", "WmsStatus"]
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
STOCK_HEADERS = ["ЕИ", "Наименование", "Артикул", "Единица измерения", "Остаток", "Место хранения", "Категория", "Состояние", "Источник",
                 "Тип источника"]
STOCK_WIDTHS = [3300, 6500, 3200, 2300, 2300, 2800, 3000, 2500, 3500, 3000]
RCPT = "Получатели"
HELP = "Справка"                  # the guide of the storekeeper, release notes, backup / recovery (docs/*.md), read-only
HELP_DOCS = ["OPERATOR_GUIDE.md", "RELEASE_NOTES.md", "BACKUP_RECOVERY.md"]
STATUS_ROW = 24                   # WmsStatus.STATUS_ROW
MAIN = "Главная"
# «Заказы» A:AB — the fixed user interface (MASTER SPEC v0.3 §2; WmsConfig OC_*)
ORDERS = "Заказы"
ORDER_HEADERS = ["№ заказа", "Полное наименование товара", "Номер документа", "Номер счёта", "Артикул", "Фактическое количество",
                 "Количество по документу", "Заказанное количество", "Единица измерения", "Цена", "Сумма", "Площадка / Поставщик",
                 "Продавец", "Дата поступления", "Дата документа", "Дата заказа", "Ожидаемая дата поступления", "Покупатель",
                 "Категория", "Кому назначено", "Место хранения", "Внутренний код", "Статус", "Наличие", "Контроль", "Комментарий",
                 "Срок поставки, дней", "Возможный дубль"]
ORDER_WIDTHS = [2600, 6500, 3000, 2800, 3000, 2600, 2600, 2600, 2300, 2300, 2300, 3500, 3000, 2600, 2600, 2600, 2800, 3000,
                3000, 3000, 2800, 3300, 4200, 2300, 7000, 4500, 2300, 4500]
ORDER_LOCKS_OPEN = "0000000000000000000001111001"      # WmsConfig.ORDER_LOCKS_OPEN: V W X Y AB locked
# service sheets of the orders (WmsConfig OD_*, RV_*, IX_*)
ORD = "_ORD"
ORD_HEADERS = ["OLID", "Строка (индекс)", "Ключ (ЕИ исходной строки)", "Отпечаток позиции", "Заказано", "Получено", "Поступлений",
               "Без документов", "Отмена"]
RCV = "_RCV"
RCV_HEADERS = ["ЕИ", "OLID", "Строка (индекс)", "Антидубль", "Состояние", "Вид", "Количество", "Без документов", "Ключ антидубля"]
IDX = "_IDX"
# «Возврат» A:O — returns of issued goods (Core Phase 4; WmsConfig RC_*); the service data is kept in _RET and _ISS
RETURNS = "Возврат"
RETURN_HEADERS = ["№ возврата", "№ выдачи", "Внутренний код", "Наименование", "Артикул", "Количество возврата", "Единица измерения",
                  "Дата возврата", "От кого возвращено", "Место хранения", "Категория", "Остаток до возврата", "Остаток после возврата",
                  "Контроль", "Комментарий"]
RETURN_WIDTHS = [1900, 1900, 3300, 6000, 3000, 2400, 2300, 2600, 4800, 2800, 3000, 2500, 2500, 7000, 4500]
RETURN_LOCKS_OPEN = "100110100011110"      # WmsConfig.RETURN_LOCKS_OPEN: unposted rows, inputs B C F H I J O open
RET = "_RET"
RET_HEADERS = ["№ возврата", "№ выдачи", "ЕИ", "Количество", "Состояние", "Строка (индекс)"]      # WmsConfig RT_*
ISS = "_ISS"
ISS_HEADERS = ["№ выдачи", "Строка (индекс)", "Возвращено", "Возвратов"]                           # WmsConfig IS_*
# «Иной приход» A:S — special receipts (Core Phase 5; WmsConfig XC_*): Офис, Производство, Детали, Старый склад, Иной
SPECIAL = "Иной приход"
SPECIAL_HEADERS = ["№ строки", "Тип прихода", "№ поступления", "Наименование", "Артикул", "Количество", "Единица измерения",
                   "Дата поступления", "Место хранения", "Категория", "Кто передал", "Документ", "Старая маркировка", "Внутренний код",
                   "Остаток до", "Остаток после", "Статус", "Контроль", "Комментарий"]
SPECIAL_WIDTHS = [1800, 2900, 3300, 6000, 3000, 2300, 2300, 2600, 2800, 3000, 3500, 3000, 3000, 3300, 2300, 2300, 4800, 7000, 4500]
SPECIAL_LOCKS_OPEN = "1000000000000111110"      # WmsConfig.SPECIAL_LOCKS_OPEN: unposted rows, inputs B..M and S open
SPECIAL_TYPES = ["Офис", "Производство", "Детали", "Старый склад", "Иной"]      # WmsConfig.SRC_NAMES
SPR = "_SPR"
SPR_HEADERS = ["№ строки", "Событие", "Тип", "ЕИ", "Режим", "Количество", "Состояние", "Строка (индекс)", "Дата", "Антидубль",
               "Ключ антидубля", "Артикул (норм.)", "Разобрано как", "Исправлений"]                  # WmsConfig SR_*
ART = "_ART"
ART_HEADERS = ["Артикул (норм.)", "ЕИ", "Артикул"]                                                  # WmsConfig AR_*
# «Корректировки» A:S — moves, write-offs, inventory corrections (Final Core; WmsConfig AC_*); the service data in _ADJ
ADJUST = "Корректировки"
ADJUST_HEADERS = ["№", "Вид", "Внутренний код", "Наименование", "Артикул", "Единица измерения", "Списать, количество",
                  "Фактический остаток", "Учётный остаток", "Новое место", "Дата", "Причина / основание", "Место до", "Остаток до",
                  "Остаток после", "Разница", "Контроль", "Пакет", "Комментарий"]
ADJUST_WIDTHS = [1500, 3000, 3300, 6000, 3000, 2300, 2300, 2500, 2500, 2800, 2600, 5000, 2800, 2300, 2300, 2300, 7000, 3500, 4500]
ADJUST_LOCKS_OPEN = "1001110000001111100"      # WmsConfig.ADJUST_LOCKS_OPEN: unposted rows, inputs B C G H I J K L R S open
ADJUST_KINDS = ["Перемещение", "Списание", "Инвентаризация"]      # WmsConfig.ADJ_KIND_NAMES
ADJ = "_ADJ"
ADJ_HEADERS = ["№", "Вид", "ЕИ", "Изменение остатка", "Состояние", "Строка (индекс)", "Место до", "Место после", "Учётный остаток",
               "Фактический остаток", "Дата"]                                                   # WmsConfig AJ_*
# _SYS KEY_SHEETS: sheet|keyCol|ctlCol|counterRow|inputCol[|EI|RET|SPR] — the full key check after a save without WMS (spec §15)
ISSUE_KEY_SPEC = f"{ISSUES}|0|17|{SYS_KEYS.index('NEXT_NO')}|11"
# «Заказы»: the key is the EI in V (text), issued from NEXT_EI; the input column is F (a quantity without a receipt)
ORDER_KEY_SPEC = f"{ORDERS}|21|24|{SYS_KEYS.index('NEXT_EI')}|5|EI"
# «Возврат»: the key is the return № in A, issued from NEXT_RET and registered in _RET; the input column is F (quantity)
RETURN_KEY_SPEC = f"{RETURNS}|0|13|{SYS_KEYS.index('NEXT_RET')}|5|RET"
# «Иной приход»: the key is the line № in A, issued from NEXT_SPL and registered in _SPR; the input column is F (quantity)
SPECIAL_KEY_SPEC = f"{SPECIAL}|0|17|{SYS_KEYS.index('NEXT_SPL')}|5|SPR"
# «Корректировки»: the key is the № in A, issued from NEXT_ADJ and registered in _ADJ; the input column is B (the kind)
ADJUST_KEY_SPEC = f"{ADJUST}|0|16|{SYS_KEYS.index('NEXT_ADJ')}|1|ADJ"
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
        rows.append((f"ЕИ-{i:08d}", name, art, unit, qty, place, cat, "", "тест", ""))
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


def build_returns(doc, sh):
    header(sh, RETURN_HEADERS, RETURN_WIDTHS, height=1700)
    text = number_format(doc, "@", ("ru", "RU"))
    whole = number_format(doc, "0")
    # A return №, B issue №: whole numbers; C EI and F quantity: text input (spec §12, locale-independent parsing); H date
    for c in (0, 1):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = whole
    for c in (2, 5):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
    sh.getCellRangeByPosition(7, 1, 7, LAST_ROW).NumberFormat = number_format(doc, "DD.MM.YYYY")
    for c, bit in enumerate(RETURN_LOCKS_OPEN):
        if bit == "0":
            sh.getCellRangeByPosition(c, 1, c, LAST_ROW).CellProtection = protection(False)
    x = 150
    for name, label, macro in (("btnRetFind", "Найти выдачу", "WmsReturnUi.BtnRetFind"), ("btnRetCheck", "Проверить", "WmsReturnUi.BtnRetCheck"),
                               ("btnRetPost", "Провести", "WmsReturnUi.BtnRetPost"), ("btnRetFix", "Исправить", "WmsReturnUi.BtnRetFix"),
                               ("btnRetDelete", "Удалить", "WmsReturnUi.BtnRetDelete"), ("btnRetClear", "Очистить", "WmsReturnUi.BtnRetClear")):
        add_button(doc, sh, name, label, macro, x, 120, 2750, 650)
        x += 2900


def build_special(doc, sh):
    header(sh, SPECIAL_HEADERS, SPECIAL_WIDTHS, height=1700)
    text = number_format(doc, "@", ("ru", "RU"))
    whole = number_format(doc, "0")
    # A line №: whole numbers; C event, E article, F quantity, N EI: text input (identifiers as typed, quantities parsed by
    # WMS — spec §12); H date; O P balances
    sh.getCellRangeByPosition(0, 1, 0, LAST_ROW).NumberFormat = whole
    for c in (2, 4, 5, 11, 12, 13):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
    sh.getCellRangeByPosition(7, 1, 7, LAST_ROW).NumberFormat = number_format(doc, "DD.MM.YYYY")
    for c, bit in enumerate(SPECIAL_LOCKS_OPEN):
        if bit == "0":
            sh.getCellRangeByPosition(c, 1, c, LAST_ROW).CellProtection = protection(False)
    # B «Тип прихода»: a list to choose from (WMS checks the value itself, other text is not blocked by Calc)
    rng = sh.getCellRangeByPosition(1, 1, 1, LAST_ROW)
    v = rng.Validation
    v.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
    v.ShowList = 1
    v.IgnoreBlankCells = True
    v.ShowErrorMessage = False
    v.setFormula1(";".join(f'"{t}"' for t in SPECIAL_TYPES))
    rng.Validation = v
    x = 150
    for name, label, macro in (("btnSpcCheck", "Проверить", "WmsSpecialUi.BtnSpcCheck"), ("btnSpcPost", "Провести", "WmsSpecialUi.BtnSpcPost"),
                               ("btnSpcFix", "Исправить", "WmsSpecialUi.BtnSpcFix"), ("btnSpcDelete", "Удалить", "WmsSpecialUi.BtnSpcDelete"),
                               ("btnSpcIdentify", "Разобрать", "WmsSpecialUi.BtnSpcIdentify"),
                               ("btnSpcClear", "Очистить", "WmsSpecialUi.BtnSpcClear")):
        add_button(doc, sh, name, label, macro, x, 120, 2750, 650)
        x += 2900


def build_adjust(doc, sh):
    header(sh, ADJUST_HEADERS, ADJUST_WIDTHS, height=1700)
    text = number_format(doc, "@", ("ru", "RU"))
    # A №: whole numbers; C EI and G H I quantities: text input (identifiers as typed, quantities parsed by WMS — spec §12);
    # K date; N O P balances
    sh.getCellRangeByPosition(0, 1, 0, LAST_ROW).NumberFormat = number_format(doc, "0")
    for c in (2, 6, 7, 8, 9, 17):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
    sh.getCellRangeByPosition(10, 1, 10, LAST_ROW).NumberFormat = number_format(doc, "DD.MM.YYYY")
    for c, bit in enumerate(ADJUST_LOCKS_OPEN):
        if bit == "0":
            sh.getCellRangeByPosition(c, 1, c, LAST_ROW).CellProtection = protection(False)
    # B «Вид»: a list to choose from (WMS checks the value itself)
    rng = sh.getCellRangeByPosition(1, 1, 1, LAST_ROW)
    v = rng.Validation
    v.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
    v.ShowList = 1
    v.IgnoreBlankCells = True
    v.ShowErrorMessage = False
    v.setFormula1(";".join(f'"{t}"' for t in ADJUST_KINDS))
    rng.Validation = v
    x = 150
    for name, label, macro in (("btnAdjCheck", "Проверить", "WmsAdjustUi.BtnAdjCheck"), ("btnAdjPost", "Провести", "WmsAdjustUi.BtnAdjPost"),
                               ("btnAdjFix", "Исправить", "WmsAdjustUi.BtnAdjFix"), ("btnAdjDelete", "Удалить", "WmsAdjustUi.BtnAdjDelete"),
                               ("btnAdjClear", "Очистить", "WmsAdjustUi.BtnAdjClear")):
        add_button(doc, sh, name, label, macro, x, 120, 2750, 650)
        x += 2900


def build_stock(doc, sh, rows):
    header(sh, STOCK_HEADERS, STOCK_WIDTHS)
    sh.getCellRangeByPosition(0, 1, 0, LAST_ROW).NumberFormat = number_format(doc, "@", ("ru", "RU"))
    if rows:
        sh.getCellRangeByPosition(0, 1, len(STOCK_HEADERS) - 1, len(rows)).setDataArray(rows)


def build_orders(doc, sh):
    header(sh, ORDER_HEADERS, ORDER_WIDTHS, height=2500)
    text = number_format(doc, "@", ("ru", "RU"))
    date = number_format(doc, "DD.MM.YYYY")
    # A C D E: identifiers as typed (no lost leading zeros); F G H J: quantities and price are text input parsed by WMS
    # (spec §12, locale-independent); N O P Q: dates; V: the EI (text)
    for c in (0, 2, 3, 4, 5, 6, 7, 9, 21):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
    for c in (13, 14, 15, 16):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = date
    for c, bit in enumerate(ORDER_LOCKS_OPEN):
        if bit == "0":
            sh.getCellRangeByPosition(c, 1, c, LAST_ROW).CellProtection = protection(False)
    rows = ((("btnRcvCheck", "Проверить", "WmsOrdersUi.BtnRcvCheck"), ("btnRcvPost", "Провести приход", "WmsOrdersUi.BtnRcvPost"),
             ("btnRcvMore", "Ещё поступление", "WmsOrdersUi.BtnRcvMore"), ("btnRcvFix", "Исправить", "WmsOrdersUi.BtnRcvFix"),
             ("btnRcvDelete", "Удалить", "WmsOrdersUi.BtnRcvDelete")),
            (("btnOrdCancel", "Отменить заказ", "WmsOrdersUi.BtnOrdCancel"),
             ("btnOrdCancelRest", "Отменить остаток", "WmsOrdersUi.BtnOrdCancelRest"),
             ("btnOrdRefresh", "Обновить статусы", "WmsOrdersUi.BtnOrdRefresh"), ("btnRcvClear", "Очистить", "WmsOrdersUi.BtnRcvClear")))
    for i, row in enumerate(rows):
        x = 150
        for name, label, macro in row:
            add_button(doc, sh, name, label, macro, x, 120 + i * 760, 2900, 650)
            x += 3050


def build_service(doc, sheets):
    """_ORD (order positions), _RCV (receipts), _RET (returns), _ISS (returns of each issue), _IDX (lookup formulas of the
    Calc engine)"""
    ordsh = sheets.getByName(ORD)
    for c, h in enumerate(ORD_HEADERS):
        ordsh.getCellByPosition(c, 0).setString(h)
    ordsh.getCellByPosition(10, 0).setString("NEXT_OL")
    ordsh.getCellByPosition(11, 0).setValue(1)
    ordsh.getCellRangeByPosition(3, 1, 3, LAST_ROW).NumberFormat = number_format(doc, "@", ("ru", "RU"))
    rcv = sheets.getByName(RCV)
    for c, h in enumerate(RCV_HEADERS):
        rcv.getCellByPosition(c, 0).setString(h)
    rcv.getCellRangeByPosition(8, 1, 8, LAST_ROW).NumberFormat = number_format(doc, "@", ("ru", "RU"))
    for name, heads in ((RET, RET_HEADERS), (ISS, ISS_HEADERS), (SPR, SPR_HEADERS), (ART, ART_HEADERS), (ADJ, ADJ_HEADERS)):
        sh = sheets.getByName(name)
        for c, h in enumerate(heads):
            sh.getCellByPosition(c, 0).setString(h)
    for c in (1, 10, 11):                  # _SPR: event, antidubl key, article key are texts
        sheets.getByName(SPR).getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = number_format(doc, "@", ("ru", "RU"))
    for c in (0, 2):                       # _ART: the article key and the article as typed are texts
        sheets.getByName(ART).getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = number_format(doc, "@", ("ru", "RU"))
    for c in (1, 2, 4, 6, 7):              # _ADJ: kind, EI, state and the places are texts
        sheets.getByName(ADJ).getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = number_format(doc, "@", ("ru", "RU"))
    idx = sheets.getByName(IDX)
    o = f"$'{ORDERS}'"
    r = f"$'{RCV}'"
    cells = [   # WmsConfig IX_*: row 0 key, 1 start, 2 nonce, 3 echo, 4.. formulas
        ("Искомое значение", None), ("Начало поиска (строк от первой строки данных)", 0), ("Метка поиска", 0),
        ("Метка (эхо): формулы пересчитаны", "=$B$3"),
        # whole columns: rows inserted by the user never shift or break these references; the headers never match a key
        ("Заказы.V: MATCH", f"=MATCH($B$1;INDEX({o}.$V:$V;$B$2+2):INDEX({o}.$V:$V;1048576);0)+$B$2"),
        ("Заказы.V: COUNTIF", f"=COUNTIF({o}.$V:$V;$B$1)"),
        ("_RCV.D: MATCH", f"=MATCH($B$1;INDEX({r}.$D:$D;$B$2+2):INDEX({r}.$D:$D;1048576);0)+$B$2"),
        ("_RCV.D: COUNTIF", f"=COUNTIF({r}.$D:$D;$B$1)"),
        ("Заказы.A: MATCH", f"=MATCH($B$1;INDEX({o}.$A:$A;$B$2+2):INDEX({o}.$A:$A;1048576);0)+$B$2"),
        ("Заказы.W: MATCH", f"=MATCH($B$1;INDEX({o}.$W:$W;$B$2+2):INDEX({o}.$W:$W;1048576);0)+$B$2"),
    ]
    for i, (label, v) in enumerate(cells):
        idx.getCellByPosition(0, i).setString(label)
        if isinstance(v, str):
            idx.getCellByPosition(1, i).setFormula(v)
        elif v is not None:
            idx.getCellByPosition(1, i).setValue(v)
    # WmsConfig IX_LIST: the scratch cell of one-shot array formulas (written, read and cleared at once by WMS); an array
    # formula cannot be entered into a locked cell of a protected sheet, so this one cell is unlocked (the sheet is hidden)
    idx.getCellByPosition(0, len(cells)).setString("Временная формула: строки «Заказов» для пересчёта статусов, выдачи ЕИ для возврата")
    idx.getCellByPosition(1, len(cells)).CellProtection = protection(False)
    i_ = f"$'{ISSUES}'"
    v = f"$'{RETURNS}'"
    x_ = f"$'{SPECIAL}'"
    a_ = f"$'{ART}'"
    p_ = f"$'{SPR}'"
    k_ = f"$'{ADJUST}'"
    more = [    # WmsConfig IX_I_MATCH, IX_RA_MATCH, IX_RA_COUNT (Phase 4), IX_XA_*, IX_ART_*, IX_XD_* (Phase 5), after the scratch cell
        ("Выдачи.A: MATCH", f"=MATCH($B$1;INDEX({i_}.$A:$A;$B$2+2):INDEX({i_}.$A:$A;1048576);0)+$B$2"),
        ("Возврат.A: MATCH", f"=MATCH($B$1;INDEX({v}.$A:$A;$B$2+2):INDEX({v}.$A:$A;1048576);0)+$B$2"),
        ("Возврат.A: COUNTIF", f"=COUNTIF({v}.$A:$A;$B$1)"),
        ("Иной приход.A: MATCH", f"=MATCH($B$1;INDEX({x_}.$A:$A;$B$2+2):INDEX({x_}.$A:$A;1048576);0)+$B$2"),
        ("Иной приход.A: COUNTIF", f"=COUNTIF({x_}.$A:$A;$B$1)"),
        ("_ART.A: MATCH", f"=MATCH($B$1;INDEX({a_}.$A:$A;$B$2+2):INDEX({a_}.$A:$A;1048576);0)+$B$2"),
        ("_ART.A: COUNTIF", f"=COUNTIF({a_}.$A:$A;$B$1)"),
        ("_SPR.J: MATCH", f"=MATCH($B$1;INDEX({p_}.$J:$J;$B$2+2):INDEX({p_}.$J:$J;1048576);0)+$B$2"),
        ("_SPR.J: COUNTIF", f"=COUNTIF({p_}.$J:$J;$B$1)"),
        # WmsConfig IX_AA_MATCH, IX_AA_COUNT (Final Core)
        ("Корректировки.A: MATCH", f"=MATCH($B$1;INDEX({k_}.$A:$A;$B$2+2):INDEX({k_}.$A:$A;1048576);0)+$B$2"),
        ("Корректировки.A: COUNTIF", f"=COUNTIF({k_}.$A:$A;$B$1)"),
    ]
    for i, (label, f) in enumerate(more, start=len(cells) + 1):
        idx.getCellByPosition(0, i).setString(label)
        idx.getCellByPosition(1, i).setFormula(f)
    idx.getColumns().getByIndex(0).Width = 7000
    idx.getColumns().getByIndex(1).Width = 5000


def build_recipients(sh, rows):
    header(sh, ["Сокращение", "Получатель"], [3500, 8000])
    if rows:
        sh.getCellRangeByPosition(0, 1, 1, len(rows)).setDataArray(tuple(rows))


def core_versions():
    """WMS_PRODUCT_VERSION, WMS_CORE_VERSION, WMS_SYS_SCHEMA of src/basic/WmsConfig.bas"""
    txt = read_source(os.path.join(SRC, "WmsConfig.bas"))
    return {k: re.search(rf'Public Const {k} = "([^"]+)"', txt).group(1) for k in ("WMS_PRODUCT_VERSION", "WMS_CORE_VERSION", "WMS_SYS_SCHEMA")}


def build_main(doc, sh):
    sh.getColumns().getByIndex(0).Width = 5200
    sh.getColumns().getByIndex(1).Width = 17500
    t = sh.getCellByPosition(0, 0)
    t.setString("WMS — состояние и восстановление")
    t.CharWeight = 150
    t.CharHeight = 16
    v = core_versions()
    sh.getCellByPosition(1, 0).setString(f"версия {v['WMS_PRODUCT_VERSION']} · ядро {v['WMS_CORE_VERSION']} · схема {v['WMS_SYS_SCHEMA']}")
    sh.getCellByPosition(1, 0).VertJustify = 2      # centre, next to the title
    captions = ["Состояние", "Причина", "Что сделать", "Непроведённые строки", "Заметки запуска", "Последнее действие", "Обновлено"]
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
               ("btnSelfCheck", "Самопроверка", "WmsUi.BtnSelfCheck"), ("btnBackup", "Резервная копия", "WmsUi.BtnBackup"),
               ("btnExport", "Экспорт для инструментов", "WmsExport.BtnExport"), ("btnLoadBatch", "Загрузить пакет", "WmsExport.BtnLoadBatch"),
               ("btnStatus", "Состояние системы", "WmsStatus.BtnSystemStatus"), ("btnPreWork", "Проверка перед работой", "WmsStatus.BtnPreWorkCheck"))
    for i, (name, label, macro) in enumerate(buttons):
        add_button(doc, sh, name, label, macro, 200 + (i % 2) * 6700, y + (i // 2) * 1000, 6400, 800)
    help_row = 22
    sh.getCellByPosition(0, help_row).setString(
        "Приход: на листе «Заказы» в строке заказа укажите фактическое количество (F), дату поступления (N), место хранения (U), "
        "при наличии — документ (C, G, O) и нажмите «Провести приход»: будет создан новый ЕИ. Следующая поставка той же позиции — "
        "«Ещё поступление». Выдача: на листе «Выдачи» введите ЕИ в «Внутренний код», количество, дату и получателя и нажмите «Провести». "
        "Возврат: на листе «Возврат» укажите № выдачи (B) или нажмите «Найти выдачу», затем количество (F) и дату возврата (H) и "
        "нажмите «Провести»: количество вернётся в остаток того же ЕИ. "
        "Приход не по заказу (Офис, Производство, Детали, Старый склад, Иной): лист «Иной приход» — тип прихода (B), наименование, "
        "артикул (для деталей обязателен: тот же артикул пополняет тот же ЕИ), количество, единица, дата, место и «Провести»; "
        "несколько выделенных строк можно провести одним поступлением (один № OFF/PROD/DET/OLD/OTH). "
        "Перемещение, списание, инвентаризация: лист «Корректировки» — вид (B), ЕИ (C), для списания количество (G), для "
        "инвентаризации фактический остаток (H), для перемещения новое место (J), дата (K), причина (L) и «Провести». "
        "«Экспорт для инструментов» создаёт снимок для WMS_TOOLBOX в папке WMS_Export рядом с книгой; «Загрузить пакет» — "
        "строки из файла инструмента (они проводятся обычными проверками). "
        "«Исправить» и «Удалить» работают для проведённой строки под курсором, «Очистить» — для непроведённой строки или копии. "
        "Перед началом рабочего дня — «Проверка перед работой»; версии, журнал и резервные копии — «Состояние системы»; "
        "руководство кладовщика, что нового в версии и как восстановить WMS из резервной копии — лист «Справка».")
    sh.getCellRangeByPosition(0, help_row, 1, help_row).merge(True)
    sh.getCellByPosition(0, help_row).IsTextWrapped = True
    sh.getRows().getByIndex(help_row).Height = 6300
    t = sh.getCellByPosition(0, STATUS_ROW)
    t.setString("СОСТОЯНИЕ СИСТЕМЫ")
    t.CharWeight = 150
    sh.getCellByPosition(1, STATUS_ROW).setString("нажмите «Состояние системы» или «Проверка перед работой»")
    sh.getCellRangeByPosition(0, STATUS_ROW + 1, 1, STATUS_ROW + 31).IsTextWrapped = True


def md_plain(s):
    """a line of the guides without the Markdown marks: **bold**, `code`, links"""
    s = s.replace("**", "").replace("`", "")
    return re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", s)


def build_help(sh):
    """«Справка»: docs/OPERATOR_GUIDE.md, RELEASE_NOTES.md and BACKUP_RECOVERY.md as text (headings bold, a table row per line)"""
    sh.getColumns().getByIndex(0).Width = 26000
    r = 0
    for name in HELP_DOCS:
        for raw in open(os.path.join(ROOT, "docs", name), encoding="utf-8").read().splitlines():
            line = raw.rstrip()
            if re.fullmatch(r"\|?[\s:|-]+\|?", line) and "-" in line:
                continue                                   # the separator row of a Markdown table
            head = re.match(r"(#+)\s+(.*)", line)
            if line.startswith("|"):
                line = " — ".join(md_plain(c.strip()) for c in line.strip("|").split("|"))
            elif head:
                line = md_plain(head.group(2))
            else:
                line = md_plain(line)
            c = sh.getCellByPosition(0, r)
            if line:
                c.setString(line)
            if head:
                c.CharWeight = 150
                c.CharHeight = {1: 15, 2: 13}.get(len(head.group(1)), 11)
            r += 1
        r += 2
    sh.getCellRangeByPosition(0, 0, 0, r).IsTextWrapped = True


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
    for i, name in enumerate((ORDERS, SPECIAL, ISSUES, RETURNS, ADJUST, STOCK, RCPT, HELP, ORD, RCV, SPR, ART, RET, ISS, ADJ, IDX, "_SYS"), start=1):
        sheets.insertNewByName(name, i)
    sys_sh = sheets.getByName("_SYS")
    values = {
        "SCHEMA": SCHEMA, "INSTANCE_ID": instance_id or "%016X" % random.getrandbits(64), "MODE": "TEST" if test else "PROD",
        "CORE_VERSION": "", "LAST_SEQ": 0, "NEXT_EI": 1, "NEXT_NO": 1, "NEXT_RET": 1, "JOURNAL_POS": "", "REGISTERED_URL": "",
        "TX_STATE": "NONE", "TX_SEQ": 0, "TX_TYPE": "", "TX_TIME": "", "TX_BEFORE_IMAGE": "", "SAVE_STAMP": 0, "SAVE_SEQ": 0,
        "MAX_QTY": 100000,
        "KEY_SHEETS": ISSUE_KEY_SPEC + ";" + ORDER_KEY_SPEC + ";" + RETURN_KEY_SPEC + ";" + SPECIAL_KEY_SPEC + ";" + ADJUST_KEY_SPEC
                      + (";" + TEST_KEY_SPEC if test else ""),
        "NEXT_SPL": 1, "NEXT_OFF": 1, "NEXT_PROD": 1, "NEXT_DET": 1, "NEXT_OLD": 1, "NEXT_OTH": 1, "NEXT_ADJ": 1,
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
    build_orders(doc, sheets.getByName(ORDERS))
    build_issues(doc, sheets.getByName(ISSUES))
    build_returns(doc, sheets.getByName(RETURNS))
    build_special(doc, sheets.getByName(SPECIAL))
    build_adjust(doc, sheets.getByName(ADJUST))
    build_stock(doc, sheets.getByName(STOCK), registry)
    build_recipients(sheets.getByName(RCPT), TEST_RECIPIENTS if test else [])
    build_help(sheets.getByName(HELP))
    build_service(doc, sheets)
    order = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    autofilter(doc, order.index(ORDERS), "WMS_ORDERS", len(ORDER_HEADERS))
    autofilter(doc, order.index(SPECIAL), "WMS_SPECIAL", len(SPECIAL_HEADERS))
    autofilter(doc, order.index(ISSUES), "WMS_ISSUES", len(ISSUE_HEADERS))
    autofilter(doc, order.index(RETURNS), "WMS_RETURNS", len(RETURN_HEADERS))
    autofilter(doc, order.index(ADJUST), "WMS_ADJUST", len(ADJUST_HEADERS))
    autofilter(doc, order.index(STOCK), "WMS_STOCK", len(STOCK_HEADERS))
    # lookups of the Calc engine (_IDX): whole-cell comparison, no regular expressions; wildcards are escaped by WMS
    doc.setPropertyValue("MatchWholeCell", True)
    doc.setPropertyValue("RegularExpressions", False)
    doc.setPropertyValue("Wildcards", True)
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
    bind_event(sheets.getByName(ORDERS).Events, "OnChange", "WmsOrders.OnOrdersChange")
    bind_event(sheets.getByName(RETURNS).Events, "OnChange", "WmsReturn.OnReturnsChange")
    bind_event(sheets.getByName(SPECIAL).Events, "OnChange", "WmsSpecial.OnSpecialChange")
    bind_event(sheets.getByName(ADJUST).Events, "OnChange", "WmsAdjust.OnAdjustChange")
    bind_event(sheets.getByName(RCPT).Events, "OnChange", "WmsIssue.OnRecipientsChange")
    ctl = doc.getCurrentController()
    for name in (ORDERS, SPECIAL, ISSUES, RETURNS, ADJUST, STOCK, RCPT):
        ctl.setActiveSheet(sheets.getByName(name))
        ctl.freezeAtPosition(0, 1)
    ctl.setActiveSheet(sheets.getByName(ISSUES))
    ctl.select(sheets.getByName(ISSUES).getCellByPosition(11, 1))
    for name in (ORD, RCV, SPR, ART, RET, ISS, ADJ, IDX, "_SYS"):
        sheets.getByName(name).IsVisible = False
    for name in (MAIN, ORDERS, SPECIAL, ISSUES, RETURNS, ADJUST, STOCK, HELP, ORD, RCV, SPR, ART, RET, ISS, ADJ, IDX, "_SYS"):
        sheets.getByName(name).protect(PWD)
    doc.protect(PWD)
    if os.path.exists(out_path):
        os.remove(out_path)
    doc.storeAsURL(uno.systemPathToFileUrl(out_path), props(FilterName="calc8"))
    doc.close(True)
    patch_content(out_path, lambda x: allow_insert_rows(allow_insert_rows(allow_insert_rows(allow_insert_rows(allow_insert_rows(
        x, ISSUES), ORDERS), RETURNS), SPECIAL), ADJUST))
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
    if xml.count('loext:insert-rows="true"') != 5:
        raise RuntimeError("the «insert rows» protection option of «Выдачи»/«Заказы»/«Возврат»/«Иной приход»/«Корректировки» did not survive the save")
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
