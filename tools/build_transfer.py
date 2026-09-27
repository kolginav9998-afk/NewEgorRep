"""Builds WMS_LEGACY_TRANSFER.ods — the book of the transfer of the old table «Заказы» into WMS 0.7 (M7 PRIME).

    python3 tools/build_transfer.py OUT_DIR

The book is the frontend of tools/legacy_transfer.py: it lies in the release folder next to WMS_PROD_CANDIDATE.ods and
the folder tools/. Sheets: «LEGACY TRANSFER» (the five steps as big buttons with their state, «Очистить staging»,
«Открыть ошибки», «Экспорт отчёта», «Создать резервную копию», the help), «1_Вставить_Заказы» (the old «Заказы» A:AB by
one Ctrl+V; the result of the check of every row in AD:AE), «1B_Вставить_Наличие», «1C_Вставить_Выдачи»,
«1D_Вставить_Возвраты» (optional), «2_Проверка» (the counters, red / yellow / green), «3_Ошибки», «4_Сверка»,
«Настройки» (the folders and the explicit rules of the transfer). The Basic module src/transfer/LtMain.bas.
"""
import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import uno  # noqa: E402
from wmslo import Office, add_basic, bind_event, props  # noqa: E402
from build_ods import add_button, add_cond, cell_style, number_format, read_source  # noqa: E402
from build_toolbox import validation_list  # noqa: E402

BOOK = "WMS_LEGACY_TRANSFER.ods"
MAIN = "LEGACY TRANSFER"
ORDERS = "1_Вставить_Заказы"
STOCK = "1B_Вставить_Наличие"
ISSUES = "1C_Вставить_Выдачи"
RETURNS = "1D_Вставить_Возвраты"
CHECK = "2_Проверка"
ERRORS = "3_Ошибки"
VERIFY = "4_Сверка"
SETTINGS = "Настройки"
ORDER_HEADERS = ["№ позиции", "Полное наименование товара", "Номер документа", "Номер счёта", "Артикул", "Фактическое количество",
                 "Количество по документу", "Заказанное количество", "Единица измерения", "Цена", "Сумма", "Площадка / Поставщик", "Продавец",
                 "Дата поступления", "Дата документа", "Дата заказа", "Ожидаемая дата поступления", "Покупатель", "Категория", "Кому назначено",
                 "Место хранения", "Внутренний код", "Статус", "Наличие", "Контроль", "Комментарий", "Срок поставки, дней", "Возможный дубль"]
ORDER_WIDTHS = [1800, 6500, 2800, 2600, 3000, 2400, 2400, 2400, 2000, 2000, 2200, 3500, 3000, 2600, 2600, 2600, 2800, 3000, 3000, 3000,
                2800, 3300, 4200, 2200, 5000, 4500, 2300, 3500]
SIDE = {STOCK: (["ЕИ", "Наименование", "Артикул", "Единица", "Остаток", "Место", "Категория"], [3300, 6500, 3200, 2300, 2300, 2800, 3000]),
        ISSUES: (["ЕИ", "Количество", "Дата", "Кому"], [3300, 2600, 2600, 5000]),
        RETURNS: (["ЕИ", "Количество", "Дата", "От кого"], [3300, 2600, 2600, 5000])}
STEPS = [("btnPaste", "1. ВСТАВИТЬ ДАННЫЕ", "LtMain.BtnLtPaste"), ("btnCheck", "2. ПРОВЕРИТЬ", "LtMain.BtnLtCheck"),
         ("btnBuild", "3. СОЗДАТЬ ТЕСТОВУЮ WMS", "LtMain.BtnLtBuild"), ("btnVerify", "4. СВЕРИТЬ", "LtMain.BtnLtVerify"),
         ("btnPromote", "5. СДЕЛАТЬ РАБОЧЕЙ", "LtMain.BtnLtPromote")]
EXTRA = [("btnClear", "Очистить staging", "LtMain.BtnLtClear"), ("btnErrors", "Открыть ошибки", "LtMain.BtnLtErrors"),
         ("btnExport", "Экспорт отчёта", "LtMain.BtnLtExport"), ("btnBackup", "Создать резервную копию", "LtMain.BtnLtBackup")]
UNITS = "шт|м|кг|л|упак|пар|компл|рул|м2|м3|т|г|мл|см|мм|лист|набор|пог. м|бух"
# «Настройки»: rows 1.. (LtMain LS_*)
SETTINGS_ROWS = [
    ("Настройка", "Значение", "Пояснение"),
    ("Книга-кандидат WMS", "WMS_PROD_CANDIDATE.ods", "чистая книга выпуска (путь от этой книги); она не изменяется — перенос идёт в её копию"),
    ("Рабочая папка переноса", "LEGACY_WORK", "копия данных, результаты проверки, тестовая WMS, отчёты"),
    ("Папка рабочей WMS", "WMS_WORK", "создаётся кнопкой «СДЕЛАТЬ РАБОЧЕЙ»; должна не существовать или быть пустой"),
    ("Файл старой таблицы (необязательно)", "", "только для отметки SHA-256 в отчёте и CUTOVER_README; файл не открывается"),
    ("Единица, если не указана", "", "пусто — строка без единицы красная; например «шт» — подставляется (отмечается как нормализация)"),
    ("Место хранения, если не указано", "", "пусто — приход без места красный; например «Склад» — подставляется"),
    ("Остаток, если не определён", "СТОП", "нет X «Наличие», нет листа 1B и выдач 1C/1D: СТОП — красная строка; 0 — остаток 0; F — всё полученное"),
    ("Строка с F без H (нет позиции выше)", "СТОП", "СТОП — красная строка; ПОЗИЦИЯ H=F — отдельная позиция, заказано = получено"),
    ("Нет даты поступления (N)", "СТОП", "СТОП — красная строка; ДАТА ДОКУМЕНТА — берётся дата документа (O)"),
    ("ЕИ только в листе «Наличие» (1B)", "СТОП", "ЕИ с остатком, которого нет в «Заказы»: СТОП; НЕ ПЕРЕНОСИТЬ; СТАРЫЙ СКЛАД — перенести как «Старый склад»"),
    ("Цена с более чем 3 знаками", "СТОП", "СТОП — красная строка; ОКРУГЛИТЬ — до 3 знаков"),
    ("Строки без количеств (нет H и F)", "СТОП", "СТОП — красная строка; ПРОПУСТИТЬ — строка не переносится (указывается в отчёте)"),
    ("Известные единицы", UNITS, "через «|»; другие единицы — жёлтые (перенос не останавливают)"),
    ("Известные места", "", "через «|»; пусто — принимаются места таблицы"),
    ("Версия книги переноса", "", "WMS_LEGACY_TRANSFER"),
]
CHOICES = {7: ["СТОП", "0", "F"], 8: ["СТОП", "ПОЗИЦИЯ H=F"], 9: ["СТОП", "ДАТА ДОКУМЕНТА"], 10: ["СТОП", "НЕ ПЕРЕНОСИТЬ", "СТАРЫЙ СКЛАД"],
           11: ["СТОП", "ОКРУГЛИТЬ"], 12: ["СТОП", "ПРОПУСТИТЬ"]}
HELP = (
    "Как перенести старую таблицу. 1) Откройте старую таблицу, на листе «Заказы» выделите колонки A:AB целиком и нажмите Ctrl+C; "
    "здесь на листе «1_Вставить_Заказы» выберите A1 и нажмите Ctrl+V (кнопка «1. ВСТАВИТЬ ДАННЫЕ» открывает этот лист). Колонка A — "
    "№ позиции внутри заказа: каждая строка с A = 1 начинает следующий заказ. Строка без заказанного количества (H), но с фактическим (F) — "
    "ещё одно поступление позиции выше. Необязательно: старые «Наличие», «Выдачи», «Возвраты» — на листы 1B, 1C, 1D (с заголовками). "
    "2) «ПРОВЕРИТЬ» ничего не меняет: формулы заменяются значениями, строится полный план переноса; итог — лист «2_Проверка», замечания — "
    "«3_Ошибки» и колонки AD, AE листа заказов (красный — блокирует перенос, жёлтый — перенесётся, посмотрите, зелёный — готово). "
    "Исправьте красные строки прямо здесь (или задайте правило в «Настройках») и проверьте снова. 3) «СОЗДАТЬ ТЕСТОВУЮ WMS» — копия чистой "
    "книги-кандидата выпуска, в ней весь перенос обычными операциями WMS, самопроверка, независимая проверка журнала, сохранение и "
    "повторное открытие (LEGACY_WORK/test/WMS_LEGACY_TEST.ods — её можно открыть и посмотреть). 4) «СВЕРИТЬ» — тестовая WMS против "
    "вставленных данных, отчёт LEGACY_TRANSFER_REPORT.html. 5) «СДЕЛАТЬ РАБОЧЕЙ» доступна только после чистой сверки: папка WMS_WORK с "
    "WMS_PROD.ods, CUTOVER_README.txt, отчётом и WMS_TOOLBOX. С этого момента все новые движения — только в новой WMS; старая таблица — "
    "архив. Старая таблица никогда не открывается и не изменяется этой книгой.")


def fill_header(sh, titles, widths, row=0, height=1100):
    sh.getCellRangeByPosition(0, row, len(titles) - 1, row).setDataArray((tuple(titles),))
    for c, w in enumerate(widths):
        sh.getColumns().getByIndex(c).Width = w
    rng = sh.getCellRangeByPosition(0, row, len(titles) - 1, row)
    rng.CharWeight = 150
    rng.IsTextWrapped = True
    rng.CellBackColor = 0xE7E6E6
    rng.VertJustify = 2
    sh.getRows().getByIndex(row).Height = height


def level_formats(doc, sh, rng, col, base_row, first_word=True):
    """red / yellow / green by the level word in column col (БЛОКЕР / ПРОВЕРИТЬ / ГОТОВО; the row results «БЛОКЕР» /
    «Проверить» / «Готово» too)"""
    L = chr(ord("A") + col) if col < 26 else "A" + chr(ord("A") + col - 26)
    add_cond(sh, rng, f'UPPER(${L}{base_row + 1})="БЛОКЕР"', "LT_Red", 0, base_row)
    add_cond(sh, rng, f'UPPER(${L}{base_row + 1})="ПРОВЕРИТЬ"', "LT_Yellow", 0, base_row)
    add_cond(sh, rng, f'OR(UPPER(${L}{base_row + 1})="ГОТОВО";UPPER(${L}{base_row + 1})="OK")', "LT_Green", 0, base_row)


def build(o, out_dir):
    out_dir = os.path.abspath(out_dir)
    path = os.path.join(out_dir, BOOK)
    doc = o.new_calc()
    try:
        cell_style(doc, "LT_Red", CellBackColor=0xF4C7C3)
        cell_style(doc, "LT_Yellow", CellBackColor=0xFCE8B2)
        cell_style(doc, "LT_Green", CellBackColor=0xB7E1CD)
        cell_style(doc, "LT_Title", CharHeight=18, CharWeight=150)
        sheets = doc.Sheets
        sheets.getByIndex(0).Name = MAIN
        for i, name in enumerate([ORDERS, STOCK, ISSUES, RETURNS, CHECK, ERRORS, VERIFY, SETTINGS], start=1):
            sheets.insertNewByName(name, i)
        text = number_format(doc, "@", ("ru", "RU"))
        date = number_format(doc, "DD.MM.YYYY", ("ru", "RU"))
        # ---------------------------------------------------------------- the main page
        m = sheets.getByName(MAIN)
        m.getColumns().getByIndex(0).Width = 9800
        m.getColumns().getByIndex(1).Width = 500
        m.getColumns().getByIndex(2).Width = 21000
        t = m.getCellByPosition(0, 0)
        t.setString("WMS LEGACY TRANSFER — перенос старой таблицы «Заказы» в WMS 0.7")
        t.CellStyle = "LT_Title"
        m.getCellByPosition(0, 1).setString("Старая таблица не изменяется. Всё строится в новой копии; рабочей становится только "
                                            "проверенная и сверенная копия.")
        m.getRows().getByIndex(2).Height = 400
        for i, (name, label, macro) in enumerate(STEPS):
            r = 3 + i
            m.getRows().getByIndex(r).Height = 1300
            c = m.getCellByPosition(2, r)
            c.VertJustify = 2
            c.IsTextWrapped = True
        m.getCellByPosition(2, 3).setFormula('="Вставлено строк «Заказы»: "&MAX(0;COUNTA($\'1_Вставить_Заказы\'.$B$1:$B$1048576)-1)'
                                              '&"; «Наличие»: "&MAX(0;COUNTA($\'1B_Вставить_Наличие\'.$A$1:$A$1048576)-1)'
                                              '&"; «Выдачи»: "&MAX(0;COUNTA($\'1C_Вставить_Выдачи\'.$A$1:$A$1048576)-1)'
                                              '&"; «Возвраты»: "&MAX(0;COUNTA($\'1D_Вставить_Возвраты\'.$A$1:$A$1048576)-1)')
        for r, s in ((4, "— ещё не выполнялась"), (5, "— ещё не создавалась"), (6, "— ещё не выполнялась"),
                     (7, "— недоступно: сначала шаги 2, 3 и 4 без ошибок")):
            m.getCellByPosition(2, r).setString(s)
        rng = m.getCellRangeByPosition(2, 4, 2, 7)
        add_cond(m, rng, 'LEFT($C5;1)="✓"', "LT_Green", 2, 4)
        add_cond(m, rng, 'LEFT($C5;1)="✗"', "LT_Red", 2, 4)
        add_cond(m, rng, 'LEFT($C5;1)="…"', "LT_Yellow", 2, 4)
        m.getCellByPosition(0, 9).setString("Состояние:")
        m.getCellByPosition(0, 9).CharWeight = 150
        m.getCellByPosition(2, 9).setString("Вставьте старый лист «Заказы» (шаг 1) и нажмите «2. ПРОВЕРИТЬ»")
        m.getCellByPosition(2, 9).IsTextWrapped = True
        m.getRows().getByIndex(9).Height = 1100
        m.getRows().getByIndex(11).Height = 1000
        h = m.getCellByPosition(0, 13)
        h.setString(HELP)
        h.IsTextWrapped = True
        m.getCellRangeByPosition(0, 13, 2, 13).merge(True)
        m.getRows().getByIndex(13).Height = 7600
        h.VertJustify = 1
        for i, (name, label, macro) in enumerate(STEPS):
            y = m.getCellByPosition(0, 3 + i).Position.Y
            add_button(doc, m, name, label, macro, 150, y + 100, 9500, 1100, font=13, bold=True,
                       back=(0xC6E0B4 if name == "btnPromote" else 0xDDEBF7))
        y = m.getCellByPosition(0, 11).Position.Y
        for i, (name, label, macro) in enumerate(EXTRA):
            add_button(doc, m, name, label, macro, 150 + i * 7700, y + 100, 7400, 800)
        # ---------------------------------------------------------------- the paste sheets
        o_ = sheets.getByName(ORDERS)
        fill_header(o_, ORDER_HEADERS, ORDER_WIDTHS)
        o_.getColumns().getByIndex(28).Width = 500
        o_.getCellRangeByPosition(29, 0, 30, 0).setDataArray((("Результат проверки", "Что не так / что сделано"),))
        o_.getColumns().getByIndex(29).Width = 3200
        o_.getColumns().getByIndex(30).Width = 18000
        hr = o_.getCellRangeByPosition(29, 0, 30, 0)
        hr.CharWeight = 150
        hr.CellBackColor = 0xD9D2E9
        level_formats(doc, o_, o_.getCellRangeByPosition(29, 1, 30, 200000), 29, 1)
        for name, (heads, widths) in SIDE.items():
            sh = sheets.getByName(name)
            fill_header(sh, heads, widths)
        # ---------------------------------------------------------------- results
        c_ = sheets.getByName(CHECK)
        c_.getCellByPosition(0, 0).setString("Проверка переноса (заполняется кнопкой «2. ПРОВЕРИТЬ»; красный — блокирует перенос, "
                                            "жёлтый — требуется проверка, зелёный — готово)")
        c_.getCellByPosition(0, 0).CharWeight = 150
        fill_header(c_, ["Показатель", "Значение", "Уровень", "Пояснение"], [11000, 7000, 3200, 16000], row=1, height=600)
        level_formats(doc, c_, c_.getCellRangeByPosition(0, 2, 3, 400), 2, 2)
        e_ = sheets.getByName(ERRORS)
        fill_header(e_, ["№", "Лист", "Строка", "Уровень", "Вид", "Описание"], [1500, 5200, 1800, 3200, 4200, 30000], height=600)
        level_formats(doc, e_, e_.getCellRangeByPosition(0, 1, 5, 200000), 3, 1)
        v_ = sheets.getByName(VERIFY)
        v_.getCellByPosition(0, 0).setString("Сверка тестовой WMS с данными (заполняется кнопкой «4. СВЕРИТЬ»; полный отчёт — "
                                            "LEGACY_TRANSFER_REPORT.html в рабочей папке переноса)")
        v_.getCellByPosition(0, 0).CharWeight = 150
        fill_header(v_, ["Показатель", "Значение"], [11000, 24000], row=1, height=600)
        # ---------------------------------------------------------------- settings
        s_ = sheets.getByName(SETTINGS)
        s_.getCellRangeByPosition(0, 0, 2, len(SETTINGS_ROWS) - 1).setDataArray(tuple(SETTINGS_ROWS))
        s_.getCellRangeByPosition(0, 0, 2, 0).CharWeight = 150
        for c, w in enumerate([8500, 7500, 22000]):
            s_.getColumns().getByIndex(c).Width = w
        s_.getCellRangeByPosition(1, 1, 1, len(SETTINGS_ROWS) - 1).CellBackColor = 0xFFF2CC
        s_.getCellRangeByPosition(1, 1, 1, len(SETTINGS_ROWS) - 1).NumberFormat = text
        for r, items in CHOICES.items():
            validation_list(s_.getCellByPosition(1, r), items)
        s_.getCellRangeByPosition(2, 1, 2, len(SETTINGS_ROWS) - 1).IsTextWrapped = True
        src = read_source(os.path.join(ROOT, "src", "transfer", "LtMain.bas"))
        import re
        ver = re.search(r'Public Const LT_VERSION = "([^"]+)"', src).group(1)
        s_.getCellByPosition(1, len(SETTINGS_ROWS) - 1).setString(ver)
        add_basic(doc, "LtMain", src)
        bind_event(doc.Events, "OnLoad", "LtMain.LtOnLoad")
        doc.getCurrentController().setActiveSheet(m)
        doc.storeAsURL(uno.systemPathToFileUrl(path), props(FilterName="calc8"))
    finally:
        doc.close(True)
    return path


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out")
    a = ap.parse_args()
    out = os.path.abspath(a.out)
    if os.path.exists(os.path.join(out, BOOK)):
        print(f"ОШИБКА: {os.path.join(out, BOOK)} уже существует — собирайте в новую папку", file=sys.stderr)
        return 2
    os.makedirs(out, exist_ok=True)
    import tempfile
    import shutil
    prof = tempfile.mkdtemp(prefix="wms_transfer_")
    o = Office(f"ltbuild{os.getpid()}", prof)
    try:
        p = build(o, out)
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    print(f"OK: {p}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
