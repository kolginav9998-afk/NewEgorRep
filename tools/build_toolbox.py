"""Assemble WMS_TOOLBOX/ — the separate tools of WMS (FINAL WMS MARATHON §6, §7) — from the text sources.

    python3 tools/build_toolbox.py OUT_DIR            → OUT_DIR/WMS_TOOLBOX/ (it must not exist: the books of a toolbox in use
                                                      keep the data of the people — «Журнал работы», the history of the acts)

Every tool is a LibreOffice book with Basic (src/toolbox/*.bas): it reads the snapshot «Экспорт для инструментов» of WMS
(docs/EXPORT_CONTRACT.md; by default the folder ../WMS_Export next to the toolbox folder) and writes batches for WMS
(../WMS_Batches); it never opens the working WMS book. The launcher WMS_TOOLBOX.ods opens the tools by relative paths;
WMS_PROD.ods does not need the toolbox. Scripts (python3, standard library only) go to reconcile/, backup/, labels/, insights/.
Needs LibreOffice with Python-UNO on the developer machine (like tools/build_ods.py).
"""
import argparse
import os
import re
import shutil
import sys
import tempfile

import uno

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wmslo import Office, add_basic, props  # noqa: E402
from build_ods import add_button, header, read_source  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src", "toolbox")
SCRIPTS = os.path.join(SRC, "scripts")


def tb_version():
    """TbCommon.TB_VERSION"""
    return re.search(r'Public Const TB_VERSION = "([^"]+)"', read_source(os.path.join(SRC, "TbCommon.bas"))).group(1)


SETTINGS = [("Параметр", "Значение"), ("Папка снимков WMS", "../WMS_Export"), ("Папка пакетов для WMS", "../WMS_Batches"),
            ("Снимок в работе", ""), ("Версия инструмента", "")]
ACT_KINDS = ["Акт передачи", "Передача деталей", "Передача сторонней организации"]
WORK_KINDS = ["Разгрузка / погрузка", "Работа погрузчиком", "Запрос бухгалтерии", "Закупки", "Проблемная поставка", "Прочие работы"]

# name → dict(title, help, sheets [(name, header row, headers, widths)], buttons [(label, macro)], module, extra settings, lists)
TOOLS = {
    "WMS_INVENTORY": dict(
        title="Инвентаризация",
        help="«Обновить из снимка» — строки пересчёта (ЕИ, наименование, место, учётный остаток) из последнего снимка WMS; впишите "
             "фактический остаток (колонка H); «Сформировать пакет» — файл корректировок для WMS (в WMS: «Главная» → «Загрузить пакет»). "
             "Строка, остаток которой изменился после снимка, в WMS не проведётся — её пересчитывают заново. Фильтры — лист «Настройки».",
        sheets=[("Пересчёт", 1, ["ЕИ", "Наименование", "Артикул", "Единица", "Место", "Категория", "Учётный остаток", "Фактический остаток",
                                 "Разница", "Комментарий"], [3300, 6000, 3000, 2000, 2500, 3000, 2600, 2600, 2300, 5000])],
        buttons=[("Обновить из снимка", "TbInventory.BtnInvRefresh"), ("Сформировать пакет", "TbInventory.BtnInvBatch")],
        module="TbInventory", extra=[("Фильтр: место начинается с", ""), ("Фильтр: категория", "")]),
    "WMS_ANALYTICS": dict(
        title="Аналитика",
        help="«Обновить из снимка» — лист «Сводка» (остатки, стоимость, источники, категории, места, движения по месяцам, наиболее "
             "выдаваемые позиции, заказы, оборачиваемость, получатели, проблемные позиции, транспорт), лист «Транспорт» (машины "
             "«Приход авто» по дням, неделям, месяцам, часам, поставщикам; повторяющиеся машины; на территории) и лист «Графики». "
             "«Найти закономерности» — движок закономерностей по снимку (python3): лист «Найденные закономерности» (Критично / "
             "Внимание / Наблюдение, основание, уверенность, что сделать) и лист «Расход» (дни запаса, дата исчерпания или "
             "«Недостаточно данных»). Пороги — лист «Настройки».",
        sheets=[("Сводка", -1, [], [9000, 3200, 3200, 3200, 3200, 3200, 3200, 3200]),
                ("Транспорт", -1, [], [7000, 5500, 4500, 4200, 3000, 3600, 6000]),
                ("Графики", -1, [], [4000]),
                ("Найденные закономерности", -1, [], [2600, 5200, 6500, 11000, 9000, 2800, 8000]),
                ("Расход", -1, [], [3300, 6000, 3000, 2400, 2400, 2400, 2400, 2400, 2800, 2800, 2400, 2800, 3200, 12000])],
        buttons=[("Обновить из снимка", "TbAnalytics.BtnAnRefresh"), ("Найти закономерности", "TbAnalytics.BtnAnInsights")],
        module="TbAnalytics", extra=[("Малый остаток: не больше", "2"), ("Длительный визит машины, мин", "120")]),
    "WMS_MANAGER": dict(
        title="Отчёт руководителю",
        help="«Сегодня», «Неделя», «Месяц» — отчёт по последнему снимку за этот период: складские операции, ЕИ в движении, "
             "инвентаризационные расхождения, заказы и документы, машины «Приход авто» (приехало, уехало, среднее время) и работа, "
             "не видная в движении товара, — «Журнал работы» (разгрузка и погрузка, погрузчик, бухгалтерия, закупки, проблемные "
             "поставки, прочее; ведёт кладовщик). Любой период: B2 (День / Неделя / Месяц), B3 (дата внутри, пусто — сегодня), "
             "«Отчёт за период (B2–B3)». «Сформировать отчёт руководителю» — документ (ODT и PDF) в ../WMS_Reports; «В PDF» — лист. "
             "«Контроль дня» — перед концом смены: машины на территории, непроведённые строки, неразобранный «Иной приход», "
             "проблемы диагностики, приходы без документов, прошедшие ожидаемые даты, резервная копия, снимок; и «Что изменилось сегодня».",
        sheets=[("Отчёт", -1, [], [9000, 4000, 3000, 3000]), ("Журнал работы", 0, ["Дата", "Вид работы", "Что сделано", "Часы", "Комментарий"],
                                                                 [2600, 5000, 9000, 1800, 6000]),
                ("Контроль дня", -1, [], [3000, 7000, 14000, 11000])],
        buttons=[("Обновить из снимка", "TbManager.BtnMgrRefresh"), ("Сегодня", "TbManager.BtnMgrToday"), ("Неделя", "TbManager.BtnMgrWeek"),
                 ("Месяц", "TbManager.BtnMgrMonth"), ("Отчёт за период (B2–B3)", "TbManager.BtnMgrReport"),
                 ("Сформировать отчёт руководителю", "TbManager.BtnMgrDoc"), ("В PDF", "TbManager.BtnMgrPdf"), ("Контроль дня", "TbManager.BtnMgrControl")],
        module="TbManager", modules=["TbDoctor"], extra=[]),
    "WMS_SEARCH": dict(
        title="Поиск и история ЕИ",
        help="«Обновить из снимка», затем на листе «Поиск» введите в B1 ЕИ, артикул, часть наименования, место, категорию, источник, "
             "получателя, поставщика или госномер машины и нажмите «Найти». Лист «Карточка»: ЕИ в B1 → «Показать карточку» — данные ЕИ "
             "и история (приход, пополнения, перемещения, выдачи, возвраты, списания, инвентаризация) по датам. Лист «Поставщик»: "
             "поставщик в B1 → «Заказы и машины поставщика». Лист «Машина»: госномер в B1 (в любом написании) → «История приездов».",
        sheets=[("Поиск", 2, ["Где", "ЕИ", "Наименование", "Артикул", "Место", "Категория", "Количество", "Тип / кто", "Дата", "Состояние"],
                 [3300, 3300, 6000, 3000, 2500, 3000, 2300, 5000, 2600, 4000]),
                ("Карточка", -1, [], [4000, 6000, 800, 2600, 6000, 2300, 3500, 5000, 4000]),
                ("Поставщик", -1, [], [3300, 6000, 4500, 3300, 3000, 2600, 2600, 2600, 5000]),
                ("Машина", -1, [], [2600, 2600, 5000, 4500, 3600, 3600, 1800, 3500])],
        buttons=[("Обновить из снимка", "TbSearch.BtnSearchRefresh"), ("Найти", "TbSearch.BtnSearch"), ("Показать карточку", "TbSearch.BtnCard"),
                 ("Заказы и машины поставщика", "TbSearch.BtnSupplier"), ("История приездов машины", "TbSearch.BtnVehicle")],
        module="TbSearch", extra=[]),
    "WMS_DOCTOR": dict(
        title="Диагностика",
        help="«Проверить» — аудит последнего снимка WMS и его журнала: целостность снимка, версии схемы и контракта, защиты листов и "
             "структуры книги, журнал (формат, непрерывность, экземпляр), остатки против движений, отрицательные остатки, ЕИ без "
             "карточки, связи возвратов, повторы номеров и копии, счётчики, артикулы деталей, возможные дубли, непроведённые строки, "
             "визиты машин «Приход авто», незавершённая загрузка пакетов, комплект инструментов, лишние файлы, резервная копия за "
             "сегодня. Ничего не исправляет: отчёт и план действий на листе «Отчёт».",
        sheets=[("Отчёт", 2, ["Статус", "Проверка", "Подробности", "Что делать"], [1800, 5000, 12000, 12000])],
        buttons=[("Проверить", "TbDoctor.BtnDoctor")],
        module="TbDoctor", extra=[]),
    "WMS_LABELS": dict(
        title="Этикетки",
        help="«Обновить из снимка», затем на листе «Этикетки» перечислите ЕИ (A) и число копий (B) — вручную или «Подобрать ЕИ» по "
             "фильтрам справа (место, категория, тип источника, даты получения, номера, только с остатком) — и нажмите «Построить "
             "этикетки»: лист «Макет», одна этикетка — одна страница размера этикетки (ЕИ крупно, наименование, артикул, место, штрихкод "
             "Code 128). «Предпросмотр» — одна этикетка так, как она напечатается. «Печать» — обычное окно печати (системный принтер / "
             "CUPS), «В PDF» — файл для проверки; задания записываются в «История печати», «Повторная печать» возвращает задание "
             "(номер — «Настройки», пусто — последнее). Термопринтер с командами TSPL — скрипт labels/tspl_labels.py.",
        sheets=[("Этикетки", 0, ["ЕИ", "Копий"], [4000, 2000, 600, 6200, 4200]),
                ("История печати", 0, ["№ задания", "Дата и время", "ЕИ", "Копий", "Как"], [2400, 3600, 3600, 1800, 2400])],
        buttons=[("Обновить из снимка", "TbLabels.BtnLblRefresh"), ("Подобрать ЕИ", "TbLabels.BtnLblPick"), ("Построить этикетки", "TbLabels.BtnLblBuild"),
                 ("Предпросмотр", "TbLabels.BtnLblPreview"), ("Печать", "TbLabels.BtnLblPrint"), ("В PDF", "TbLabels.BtnLblPdf"),
                 ("Повторная печать", "TbLabels.BtnLblReprint")],
        module="TbLabels", extra=[("Ширина этикетки, мм", "58"), ("Высота этикетки, мм", "40"), ("Наименование (да/нет)", "да"),
                                  ("Артикул (да/нет)", "да"), ("Место (да/нет)", "да"), ("Штрихкод (да/нет)", "да"), ("Размер шрифта ЕИ", "20"),
                                  ("Повторить задание № (пусто — последнее)", "")]),
    "WMS_DOCS": dict(
        title="Акты",
        help="«Обновить из снимка»; на листе «Акт» заполните вид (акт передачи; передача деталей — у каждой позиции артикул; передача "
             "сторонней организации — укажите организацию), дату, «Передал», «Принял» и позиции — вручную или «Подобрать по выдачам» "
             "(выдачи получателю «Принял» за период B5–B6); «Создать документ» — акт (ODT, по желанию и PDF) в папке ../WMS_Docs с "
             "номером, сторонами, таблицей позиций и подписями; лист «История» — созданные акты. Номер и история сохраняются вместе с "
             "этой книгой; акт с уже занятым номером не перезаписывается.",
        sheets=[("Акт", -1, [], [4500, 7000, 3000, 1800, 2500, 6000]), ("История", 0, ["№", "Дата", "Вид", "Принял", "Позиций", "Файл"],
                                                                        [1500, 2600, 4000, 6000, 1800, 12000])],
        buttons=[("Обновить из снимка", "TbDocs.BtnDocsRefresh"), ("Подобрать по выдачам", "TbDocs.BtnDocsPick"), ("Создать документ", "TbDocs.BtnDocsCreate")],
        module="TbDocs", extra=[("Сохранять и PDF (да/нет)", "нет"), ("Номер следующего акта", "1"), ("Папка актов", "../WMS_Docs")]),
    "WMS_ARCHIVE": dict(
        title="Архив",
        help="«Анализ» — размер рабочей книги и число строк движений по месяцам; какие месяцы старше порога «Настройки» и могут уйти "
             "в архив. «Создать архивный пакет» — папка ../WMS_Archive/ARCHIVE_<с>_<по>: строки движений этих месяцев из снимка, "
             "файлы журнала и manifest с SHA-256 (пакет проверяется сразу). «Открыть пакет» — таблицы пакета для просмотра. Рабочая "
             "книга WMS при этом не меняется: история из неё автоматически не удаляется.",
        sheets=[("Анализ", -1, [], [7000, 5000, 4000])],
        buttons=[("Анализ", "TbArchive.BtnArcAnalyze"), ("Создать архивный пакет", "TbArchive.BtnArcPackage"), ("Открыть пакет", "TbArchive.BtnArcOpen")],
        module="TbArchive", extra=[("Хранить в рабочей книге, месяцев", "12"), ("Архив: с месяца (ГГГГ-ММ)", ""), ("Архив: по месяц (ГГГГ-ММ)", ""),
                                   ("Папка архива", "../WMS_Archive")]),
    "WMS_IMPORTER": dict(
        title="Массовый приход",
        help="«Загрузить таблицу» — файл поставщика (CSV, ODS, XLSX; первая строка — заголовки) на лист «Таблица»; лист «Сопоставление» "
             "— какая колонка файла идёт в какую колонку WMS (угадывается по заголовкам; колонку C можно поправить); цель — «Иной "
             "приход» или «Заказы» («Настройки»). «Проверить» — правила WMS для каждой строки и повторы (строка файла, совпадающая с "
             "другой; поступление, которое уже есть в WMS по последнему снимку); итог в колонке «Проверка», лист «Предпросмотр» — "
             "ровно то, что уйдёт в пакет. «Сформировать пакет» — только без ошибок: файл для WMS («Главная» → «Загрузить пакет»).",
        sheets=[("Таблица", -1, [], [5000] * 12), ("Сопоставление", 0, ["Ключ пакета", "Колонка WMS", "Колонка файла", "Обязательно"],
                                                    [3500, 6500, 6500, 3000]), ("Предпросмотр", -1, [], [4500] * 16)],
        buttons=[("Загрузить таблицу", "TbImporter.BtnImpLoad"), ("Проверить", "TbImporter.BtnImpCheck"), ("Сформировать пакет", "TbImporter.BtnImpBatch")],
        module="TbImporter", extra=[("Цель пакета (Иной приход / Заказы)", "Иной приход"), ("Тип прихода по умолчанию (пусто — только из файла)", ""),
                                    ("Одно поступление на тип (да/нет)", "нет"), ("Кодировка CSV (UTF-8 / Windows-1251)", "UTF-8")]),
}

LAUNCHER = dict(
    title="Инструменты WMS",
    help="Инструменты WMS работают со снимком «Экспорт для инструментов» (в WMS: «Главная») и никогда не открывают рабочую книгу WMS. "
         "Изменения склада возвращаются в WMS только пакетами («Главная» → «Загрузить пакет»), где WMS проверяет и проводит их "
         "сама. «Обновить» — какие инструменты есть в папке, последний снимок, пакеты; кнопки открывают инструменты; «Проверка перед "
         "сменой» — скрипт backup/wms_healthcheck.py для папки WMS (Настройки, B6); «Контроль дня» — открывает «Отчёт руководителю» "
         "и сразу проверяет смену; «Сверка названий» — файл старых названий → отчёт с кандидатами-ЕИ (отметьте «да» у верного ЕИ и "
         "сохраните), «Соответствие по сверке» — отмеченный отчёт → соответствие старых названий и ЕИ, подтверждённое копится в словаре "
         "и предлагается снова (скрипты — python3). Для работы WMS эта книга не нужна.",
    buttons=[("Обновить", "TbLauncher.BtnLnRefresh"), ("Проверка перед сменой", "TbLauncher.BtnLnHealth"), ("Контроль дня", "TbLauncher.BtnLnControl"),
             ("Сверка названий", "TbLauncher.BtnLnReconcile"), ("Соответствие по сверке", "TbLauncher.BtnLnConfirm")],
    extra=[("Папка WMS (для проверки перед сменой)", "..")])


def validation_list(rng, items):
    v = rng.Validation
    v.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
    v.ShowList = 1
    v.IgnoreBlankCells = True
    v.ShowErrorMessage = False
    v.setFormula1(";".join(f'"{t}"' for t in items))
    rng.Validation = v


def main_sheet(doc, name, title, help_text, buttons, version):
    main = doc.Sheets.getByIndex(0)
    main.Name = "Главная"
    main.getColumns().getByIndex(0).Width = 26000
    t = main.getCellByPosition(0, 0)
    t.setString(f"{name} — {title}")
    t.CharWeight = 150
    t.CharHeight = 16
    main.getCellByPosition(0, 1).setString(f"WMS_TOOLBOX {version}. Читает снимок «Экспорт для инструментов» (WMS 0.6 и 0.7) и никогда не открывает "
                                           "рабочую книгу WMS.")
    # the help below the buttons (four in a row, 1 cm per row)
    hrow = 5 + 3 * ((len(buttons) + 3) // 4)
    h = main.getCellByPosition(0, hrow)
    h.setString(help_text)
    h.IsTextWrapped = True
    main.getRows().getByIndex(hrow).Height = 3200
    y = main.getCellByPosition(0, 3).Position.Y
    for i, (label, macro) in enumerate(buttons):
        add_button(doc, main, f"btn{i}", label, macro, 200 + (i % 4) * 6400, y + (i // 4) * 1000, 6200, 800)
    return main


def settings_sheet(doc, extra, version):
    doc.Sheets.insertNewByName("Настройки", doc.Sheets.getCount())
    st = doc.Sheets.getByName("Настройки")
    rows = [list(r) for r in SETTINGS + extra]
    rows[4][1] = version
    st.getCellRangeByPosition(0, 0, 1, len(rows) - 1).setDataArray(tuple(tuple(r) for r in rows))
    st.getCellRangeByPosition(0, 0, 1, 0).CharWeight = 150
    st.getColumns().getByIndex(0).Width = 9500
    st.getColumns().getByIndex(1).Width = 11000
    return st


def layout(doc, name, st):
    """the fixed cells of some work sheets"""
    sh = doc.Sheets
    if name == "WMS_MANAGER":
        rep = sh.getByName("Отчёт")
        rep.getCellRangeByPosition(0, 0, 1, 4).setDataArray((("ОТЧЁТ РУКОВОДИТЕЛЮ", ""), ("Период (День / Неделя / Месяц)", "Месяц"),
                                                             ("Дата внутри периода (пусто — сегодня)", ""), ("С", ""), ("По", "")))
        rep.getCellByPosition(0, 0).CharWeight = 150
        validation_list(rep.getCellByPosition(1, 1), ["День", "Неделя", "Месяц"])
        log = sh.getByName("Журнал работы")
        validation_list(log.getCellRangeByPosition(1, 1, 1, 10000), WORK_KINDS)
        for c in (0,):
            log.getCellRangeByPosition(c, 1, c, 10000).NumberFormat = date_format(doc)
        for c in (1, 2):
            rep.getCellByPosition(1, 2 + c).NumberFormat = date_format(doc)
        rep.getCellByPosition(1, 2).NumberFormat = date_format(doc)
    elif name == "WMS_SEARCH":
        f = sh.getByName("Поиск")
        f.getCellByPosition(0, 0).setString("Что искать:")
        f.getCellByPosition(0, 0).CharWeight = 150
        f.getCellByPosition(1, 0).CellBackColor = 0xFFF2CC
        for name, label in (("Карточка", "ЕИ:"), ("Поставщик", "Поставщик:"), ("Машина", "Госномер / машина:")):
            c = sh.getByName(name)
            c.getCellByPosition(0, 0).setString(label)
            c.getCellByPosition(0, 0).CharWeight = 150
            c.getCellByPosition(1, 0).CellBackColor = 0xFFF2CC
    elif name == "WMS_DOCS":
        a = sh.getByName("Акт")
        a.getCellRangeByPosition(0, 0, 1, 6).setDataArray((("Вид акта", "Акт передачи"), ("Дата", ""), ("Передал", ""), ("Принял", ""),
                                                           ("Период с (для подбора)", ""), ("Период по", ""), ("Организация", "")))
        for r in (1, 4, 5):
            a.getCellByPosition(1, r).NumberFormat = date_format(doc)
        a.getCellRangeByPosition(0, 8, 5, 8).setDataArray((("ЕИ", "Наименование", "Артикул", "Ед.", "Количество", "Примечание"),))
        a.getCellRangeByPosition(0, 8, 5, 8).CharWeight = 150
        a.getCellRangeByPosition(0, 0, 0, 6).CharWeight = 150
        validation_list(a.getCellByPosition(1, 0), ACT_KINDS)
    elif name == "WMS_IMPORTER":
        validation_list(st.getCellByPosition(1, 5), ["Иной приход", "Заказы"])
        validation_list(st.getCellByPosition(1, 6), ["", "Офис", "Производство", "Детали", "Старый склад", "Иной"])
        validation_list(st.getCellByPosition(1, 8), ["UTF-8", "Windows-1251"])
    elif name == "WMS_LABELS":
        lab = sh.getByName("Этикетки")
        lab.getCellRangeByPosition(0, 1, 0, 5000).NumberFormat = text_format(doc)
        lab.getCellRangeByPosition(3, 0, 4, 9).setDataArray((("ПОДБОР ЕИ ИЗ СНИМКА («Подобрать ЕИ»)", "Условие (пусто — любое)"),
                                                             ("Место начинается с", ""), ("Категория", ""), ("Тип источника", ""),
                                                             ("Получены с (дата)", ""), ("Получены по (дата)", ""), ("ЕИ с номера", ""),
                                                             ("ЕИ по номер", ""), ("Только с остатком (да/нет)", "да"), ("Копий каждой", "1")))
        lab.getCellRangeByPosition(3, 0, 4, 0).CharWeight = 150
        lab.getCellRangeByPosition(4, 1, 4, 9).CellBackColor = 0xFFF2CC
        for r in (4, 5):
            lab.getCellByPosition(4, r).NumberFormat = date_format(doc)
        validation_list(lab.getCellByPosition(4, 3), ["", "Поставщик", "Офис", "Производство", "Детали", "Старый склад", "Иной приход"])
        validation_list(lab.getCellByPosition(4, 8), ["да", "нет"])


def date_format(doc):
    lc = uno.createUnoStruct("com.sun.star.lang.Locale")
    lc.Language, lc.Country = "ru", "RU"
    nf = doc.getNumberFormats()
    k = nf.queryKey("DD.MM.YYYY", lc, False)
    return k if k != -1 else nf.addNew("DD.MM.YYYY", lc)


def text_format(doc):
    lc = uno.createUnoStruct("com.sun.star.lang.Locale")
    lc.Language, lc.Country = "ru", "RU"
    nf = doc.getNumberFormats()
    k = nf.queryKey("@", lc, False)
    return k if k != -1 else nf.addNew("@", lc)


def build_tool(o, out_dir, name, spec, version):
    doc = o.new_calc()
    # every tool book: «Все инструменты» — back to the launcher of the folder (M6 §11, the navigation between the books)
    main_sheet(doc, name, spec["title"], spec["help"], spec["buttons"] + [("Все инструменты", "TbCommon.BtnTbLauncher")], version)
    for i, (sname, hrow, heads, widths) in enumerate(spec["sheets"], start=1):
        doc.Sheets.insertNewByName(sname, i)
        sh = doc.Sheets.getByName(sname)
        for c, w in enumerate(widths):
            sh.getColumns().getByIndex(c).Width = w
        if hrow >= 0 and heads:
            header(sh, heads, widths)
            if hrow > 0:
                sh.getRows().insertByIndex(0, hrow)
                rng = sh.getCellRangeByPosition(0, 0, len(heads) - 1, hrow - 1)
                rng.CellBackColor = -1
                rng.CharWeight = 100
    st = settings_sheet(doc, spec["extra"], version)
    layout(doc, name, st)
    for m in ["TbCommon", spec["module"]] + spec.get("modules", []):
        add_basic(doc, m, read_source(os.path.join(SRC, m + ".bas")))
    doc.getCurrentController().setActiveSheet(doc.Sheets.getByIndex(0))
    p = os.path.join(out_dir, name + ".ods")
    doc.storeAsURL(uno.systemPathToFileUrl(p), props(FilterName="calc8"))
    doc.close(True)
    return p


def build_launcher(o, out_dir, version):
    doc = o.new_calc()
    tools = list(TOOLS)
    buttons = LAUNCHER["buttons"] + [(f"Открыть: {TOOLS[t]['title']}", f"TbLauncher.BtnOpen{i}") for i, t in enumerate(tools)]
    main_sheet(doc, "WMS_TOOLBOX", LAUNCHER["title"], LAUNCHER["help"], buttons, version)
    doc.Sheets.insertNewByName("Инструменты", 1)
    sh = doc.Sheets.getByName("Инструменты")
    for c, w in enumerate([5200, 5500, 16000, 3000]):
        sh.getColumns().getByIndex(c).Width = w
    sh.getCellRangeByPosition(0, 0, 0, 1).setDataArray((("Последний снимок WMS",), ("Пакеты для WMS",)))
    sh.getCellRangeByPosition(0, 2, 3, 2).setDataArray((("Файл", "Инструмент", "Назначение / команда", "Состояние"),))
    sh.getCellRangeByPosition(0, 2, 3, 2).CharWeight = 150
    sh.getCellRangeByPosition(0, 0, 0, 1).CharWeight = 150
    settings_sheet(doc, LAUNCHER["extra"], version)
    for m in ("TbCommon", "TbLauncher"):
        add_basic(doc, m, read_source(os.path.join(SRC, m + ".bas")))
    doc.getCurrentController().setActiveSheet(doc.Sheets.getByIndex(0))
    p = os.path.join(out_dir, "WMS_TOOLBOX.ods")
    doc.storeAsURL(uno.systemPathToFileUrl(p), props(FilterName="calc8"))
    doc.close(True)
    return p


def build(o, out):
    version = tb_version()
    tb = os.path.join(os.path.abspath(out), "WMS_TOOLBOX")
    if os.path.exists(tb):
        raise ValueError(f"{tb} уже существует — в книгах работающего набора данные людей (журнал работы, история актов); "
                         "соберите в новую папку и перенесите данные вручную")
    os.makedirs(tb)
    built = [build_tool(o, tb, name, spec, version) for name, spec in TOOLS.items()]
    built.append(build_launcher(o, tb, version))
    for sub in ("reconcile", "backup", "labels", "insights"):
        shutil.copytree(os.path.join(SCRIPTS, sub), os.path.join(tb, sub), ignore=shutil.ignore_patterns("__pycache__"))
    shutil.copy2(os.path.join(ROOT, "docs", "TOOLBOX.md"), os.path.join(tb, "README.md"))
    shutil.copy2(os.path.join(ROOT, "docs", "EXPORT_CONTRACT.md"), os.path.join(tb, "EXPORT_CONTRACT.md"))
    with open(os.path.join(tb, "VERSION"), "w", encoding="utf-8") as f:
        f.write(f"WMS_TOOLBOX {version}\n")
    return tb, built


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out")
    a = ap.parse_args()
    if os.path.exists(os.path.join(os.path.abspath(a.out), "WMS_TOOLBOX")):
        print(f"ОШИБКА: {os.path.join(os.path.abspath(a.out), 'WMS_TOOLBOX')} уже существует — собирайте в новую папку "
              "(в книгах работающего набора — журнал работы и история актов)", file=sys.stderr)
        return 2
    prof = tempfile.mkdtemp(prefix="wms_tb_")
    o = Office("wmstb%d" % os.getpid(), prof)
    try:
        tb, built = build(o, a.out)
        print("OK:", tb, [os.path.basename(x) for x in built])
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
