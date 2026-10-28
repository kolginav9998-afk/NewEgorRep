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
import shutil
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
from wmslo import Office, add_basic, bind_event, props, macro_url, PLUS_MODULES, PLUS_LIB  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src", "basic")
TEST_SRC = os.path.join(ROOT, "tests", "basic")
PWD = "wms"                      # WmsConfig.PROTECT_PWD
SYS_KEYS = ["SCHEMA", "INSTANCE_ID", "MODE", "CORE_VERSION", "LAST_SEQ", "NEXT_EI", "NEXT_NO", "NEXT_RET",
            "JOURNAL_POS", "REGISTERED_URL", "TX_STATE", "TX_SEQ", "TX_TYPE", "TX_TIME", "TX_BEFORE_IMAGE",
            "SAVE_STAMP", "SAVE_SEQ", "MAX_QTY", "KEY_SHEETS",
            "NEXT_SPL", "NEXT_OFF", "NEXT_PROD", "NEXT_DET", "NEXT_OLD", "NEXT_OTH", "NEXT_ADJ",
            "NEXT_CAR", "UPGRADE_FROM"]                                               # WmsConfig.SysKeyNames() (WMS-SYS-5)
SCHEMA = "WMS-SYS-5"                                                                  # WmsConfig.WMS_SYS_SCHEMA
SK_SAVE_STAMP = SYS_KEYS.index("SAVE_STAMP")
# the core of the accounting — library «Standard» (the modules of 0.7.2); the pages and the additions — library «WmsPlus»
# (wmslo.PLUS_MODULES: their names do not slow the lookups of the core down)
CORE_MODULES = ["WmsConfig", "WmsCore", "WmsJournal", "WmsLock", "WmsBackup", "WmsRecovery", "WmsDiagnostics", "WmsIssue", "WmsUi",
                "WmsOrders", "WmsReceipt", "WmsOrdersUi", "WmsReturn", "WmsReturnUi", "WmsSpecial", "WmsSpecialUi", "WmsAdjust", "WmsAdjustUi",
                "WmsCar", "WmsCarUi", "WmsExport", "WmsMigrate", "WmsStatus"]
MODULES = CORE_MODULES + PLUS_MODULES
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
STATUS_ROW = 17                   # WmsStatus.STATUS_ROW
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
ORDER_BLOCK_HEADER = "Блок (служебная)"                # WmsConfig.ORDER_BLOCK_HEADER: AC, the mark of a delivery row (M6 §18)
ORDER_BLOCK_COL = 28
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
ADJUST_KINDS = ["Перемещение", "Списание", "Инвентаризация", "Разделение", "Возврат поставщику"]      # WmsConfig.ADJ_KIND_NAMES
ADJ = "_ADJ"
ADJ_HEADERS = ["№", "Вид", "ЕИ", "Изменение остатка", "Состояние", "Строка (индекс)", "Место до", "Место после", "Учётный остаток",
               "Фактический остаток", "Дата", "Новый ЕИ"]                                                   # WmsConfig AJ_*
# «Приход авто» — the journal of vehicles (M6 PRIME §1; WmsConfig CP_*, CC_*, CR_*, CD_*): a panel of frozen rows 0..4 (two
# input cells, the hint, the indicators, the threshold) above the table (header row 5, visit n on row 5 + n); _CAR — the visits
# PRIME 0.8: the registry of the transfer (WmsConfig LG_*) and its report «Сверка переноса» (LEGREP_*)
LEG = "_LEG"
LEG_HEADERS = ["ЕИ", "Вид", "Приход при переносе", "Остаток при переносе", "Израсходовано до переноса", "Старое наличие (X)",
               "Откуда взят остаток", "Строка старой таблицы", "Операция переноса (seq)", "Коррекций переноса", "Последняя коррекция"]
# PRIME 0.8 — the packing of an EI (WmsConfig SH_PACK, PK_*): row = EI, written only by the operation EI_PACK (D-099)
PACK = "_PACK"
PACK_HEADERS = ["ЕИ", "Часть (как называть)", "Частей в 1 единице документа", "Примечание"]
# PRIME 0.8 — the view sheets (WmsConfig SH_DASH … SH_TOPO): built from the data, never journaled
DASH = "Дашборд"
FIND = "Поиск"
MAP = "Карта склада"
TOPO = "Топология"
VIEW_SHEETS = [DASH, FIND, MAP, TOPO]
# 1.0: the directory of the places outside the racks (WmsMap.SH_PLACES) — the storekeeper's own sheet, not protected
PLACES = "Места"
PLACES_HEADERS = ["Место", "Тип", "Зона", "Вместимость, ячеек", "Закрыто", "Примечание"]
PLACE_KINDS = ["Стеллаж", "Пол", "Улица", "Офис", "Другое"]
# 1.0 (З8, З10, Ф3): the directories of the suppliers and of the goods — the storekeeper's own sheets, not protected; «Снабжение»
# (WmsSupply) — the page «Ожидаем» / «Нужно заказать»
SUPPLIERS = "Поставщики"
SUPPLIER_HEADERS = ["Поставщик", "Телефон", "Контакт", "Обычный срок поставки, дней", "Примечание"]
GOODS = "Номенклатура"
GOODS_HEADERS = ["Код", "Наименование", "Артикул", "Ед.", "Категория", "Фасовка: частей в 1 ед.", "Часть (как называть)", "Минимум",
                 "Заказывать по", "Поставщик", "Фото (файл в папке «WMS_Фото»)", "Примечание", "Остаток (все ЕИ)", "ЕИ с остатком", "В пути",
                 "Обновлено"]
GOODS_WMS_COLS = 4          # the last columns M..P: written by «Обновить» (WmsSupply), not by the storekeeper
SUPPLY = "Снабжение"
# 1.0 (З12, К3, К4, К11, К12, К14, А1): the page of the documents on paper — WmsPrint fills it, «Печать…» / «Сохранить PDF»
PRINT = "Печать"
# 1.0 (этап G): the closed months (WmsLedger SH_CLOSE, CL_*): a row per closing, written only by MONTH_CLOSE / MONTH_OPEN
CLOSE = "_CLOSE"
CLOSE_HEADERS = ["Месяц", "Закрыт по", "Когда закрыт", "Состояние", "Остаток на конец", "На сумму", "Файлы", "Причина открытия"]
# 1.0 (этап G): «Отчёты» (WmsLedger) — the reports of the accounting and of the manager, made on «Печать» or into a file
REPORTS = "Отчёты"
REPORT_GROUPS = [
    ("Учёт и бухгалтерия", [
        ("btnStockOnDate", "Остаток на дату…", "WmsLedger.BtnStockOnDate",
         "Остаток всех ЕИ на любую прошедшую дату, с ценой и суммой — на печать или в файл."),
        ("btnTurnover", "Оборотная ведомость…", "WmsLedger.BtnTurnover",
         "За период по товарам: начало, приход, выдано, возвращено, списано, конец — в единицах и рублях."),
        ("btnExportAcc", "Выгрузка для бухгалтерии…", "WmsLedger.BtnExport",
         "Файл CSV всех движений за период: документы, ЕИ, цены, суммы, получатели (открывается в Excel, загружается в 1С)."),
        ("btnCompare1C", "Сверка с 1С…", "WmsLedger.BtnCompare1C",
         "Выберите файл остатков из 1С (CSV, XLSX, ODS) — список расхождений с WMS. Ничего не проводится."),
        ("btnCloseMonth", "Закрыть месяц…", "WmsLedger.BtnCloseMonth",
         "Итоги месяца сохраняются в файлы и фиксируются в журнале; операция задним числом — только с предупреждением и причиной."),
        ("btnOpenMonth", "Открыть месяц…", "WmsLedger.BtnOpenMonth",
         "Снять закрытие последнего закрытого месяца — с причиной (она остаётся в журнале)."),
    ]),
    ("Руководителю", [
        ("btnManagerReport", "Отчёт руководителю…", "WmsLedger.BtnManagerReport",
         "Сегодня, неделя, месяц или период: операции и суммы, остаток на сумму, получатели, потери, поставщики, машины."),
        ("btnConsumers", "Расход по получателям…", "WmsLedger.BtnConsumers",
         "Сколько выдано, возвращено и на руках — у каждого получателя и подразделения (колонка C «Получателей»), на какую сумму."),
        ("btnDeadStock", "Мёртвый груз…", "WmsLedger.BtnDeadStock",
         "Все ЕИ без движения дольше 90, 180 или 365 дней — по дням или по сумме, итог «на сумму …»."),
    ]),
]
PRINT_WIDTHS = [900, 2600, 6300, 2300, 1500, 2000, 2400]        # A..G = 18 cm: an A4 page with margins of 1,5 cm
# 1.0 (Ф8): «Комплекты» — the storekeeper's own directory of the kits (WmsKit SH_KITS, KT_*): not protected, not journaled
KITS = "Комплекты"
KITS_HEADERS = ["Комплект", "Позиция (наименование)", "Артикул", "Количество на 1 комплект", "Ед.", "Примечание"]
# 1.0 (Ф5): _PROP — the properties of the EIs, one row per EI (row = EI number; WmsItem SH_PROP, PR_*), written only by EI_PROP
PROP = "_PROP"
PROP_HEADERS = ["ЕИ", "Серийный / инв. №", "Годен до", "Примечание"]
# 1.0 (Ф7): «Резервы» — the storekeeper's own sheet of the reserves (WmsItem SH_RESERVES, RS_*): not protected, not journaled
RESERVES = "Резервы"
RESERVE_HEADERS = ["ЕИ", "Количество", "Для кого / объект", "До (дата)", "Примечание"]
# 1.0 (Ф4): «Номенклатура» Q — the days a returnable thing (a tool) may stay on the hands (WmsItem GD_RETDAYS)
GOODS_RET_COL = 16
GOODS_RET_HEADER = "Вернуть через, дней (инструмент)"
# 1.0.2: «Номенклатура» R — how many pieces come in one pack (WmsItem GD_PACKQ): the label of such goods says «Выдача по 1 шт · в уп. 10».
# The hint of the sheet (R1 up to 1.0.1) moved one column to the right
GOODS_PACKQ_COL = 17
GOODS_PACKQ_HEADER = "В упаковке, шт (на этикетке: выдача по 1 шт)"
# 1.0 (этапы J, K): «Обслуживание» (WmsSafe) — the safety of the data and the care of the book: the state on top (refreshed when
# the page is opened), the buttons with what they do below
SERVICE = "Обслуживание"
SERVICE_STATE = ["Копия на флешке", "Зеркало журнала", "Проверка дня", "Копии книги", "Свободно на диске", "Книга и журнал", "Рост книги",
                 "Скорость", "Этот компьютер", "Учебное восстановление"]
SERVICE_STATE_ROW = 3            # WmsSafe.SV_STATE_ROW
SERVICE_GROUPS = [
    ("Сохранность данных", [
        ("btnFlashCopy", "Копия на флешку…", "WmsSafe.BtnFlashCopy",
         "Книга и журнал — на флешку или другой диск, каждый файл проверяется. Раз в неделю: при поломке компьютера данные останутся."),
        ("btnReturnCopy", "Вернуться к копии…", "WmsSafe.BtnReturnToCopy",
         "Копии книги с датой и числом операций. Книга вернётся к выбранной, операции после неё WMS доведёт по журналу и проверит."),
        ("btnMirror", "Зеркало журнала…", "WmsSafe.BtnMirror",
         "Вторая копия журнала на другом диске или флешке: каждая проведённая операция пишется и туда."),
        ("btnMirrorRestore", "Журнал из зеркала", "WmsSafe.BtnMirrorRestore",
         "Журнал испорчен или потерял конец — целые файлы берутся из зеркала, прежние откладываются в «WMS_Backups»."),
        ("btnAbandoned", "Отложенные операции", "WmsSafe.BtnAbandoned",
         "Что ушло в «Отложить хвост журнала» и не попало в книгу; выдачи возвращаются в «Выдачи» одной кнопкой."),
        ("btnDirsBook", "Справочники из копии…", "WmsDirs.BtnDirsFromBook",
         "Справочники («Получатели», «Номенклатура», «Акты», «Категории»…) в журнал не пишутся: после возврата к копии они такие, как были в копии. Возьмите их из нужной копии или из отложенной книги."),
    ]),
    ("Проверки", [
        ("btnDailyCheck", "Проверить сейчас", "WmsSafe.BtnDailyCheck",
         "То же, что WMS делает при первом запуске дня: остаток = движения, фасовка, ключи листов, цепочка журнала, зеркало."),
        ("btnPreWork", "Проверка перед работой", "WmsStatus.BtnPreWorkCheck",
         "LibreOffice, макросы, запись в папки, копии, свободное место, защита листов."),
        ("btnDrill", "Учебное восстановление", "WmsSafe.BtnRestoreDrill",
         "Раз в месяц: самая свежая копия открывается отдельно, журнал от неё проверяется до последней операции — «восстанавливается»."),
    ]),
    ("Книга и скорость", [
        ("btnSpeedTest", "Замер скорости ПК", "WmsCare.BtnSpeedTest",
         "Несколько секунд: вычисления, формулы, диск — и сколько движений этот компьютер держит комфортно."),
        ("btnLoTips", "Советы LibreOffice", "WmsCare.BtnLoTips",
         "Настройки LibreOffice для большой книги: резервная копия при сохранении, восстановление, пересчёт при загрузке."),
        ("btnExtraCells", "Лишние ячейки", "WmsCare.BtnExtraCells",
         "Листы, у которых занятая область тянется далеко ниже данных (файл больше и медленнее). Ничего не меняет."),
        ("btnPruneSnap", "Удалить старые снимки", "WmsCare.BtnPruneSnapshots",
         "Снимки «Экспорт для инструментов», кроме двух последних, удаляются; книга и журнал не затрагиваются."),
        ("btnLiteSnap", "Снимок для этикеток", "WmsCare.BtnLiteSnapshot",
         "Лёгкий снимок для набора «Этикетки»: только «Наличие», «Заказы», «Иной приход», без журнала — секунды."),
    ]),
    ("Этот компьютер и обучение", [
        ("btnPcCheck", "Проверить этот ПК", "WmsHelp.BtnPcCheck",
         "За секунды, без интернета: LibreOffice, макросы и доверенная папка, запись в папку WMS, место на диске, Python."),
        ("btnShortcut", "Ярлык на рабочий стол", "WmsHelp.BtnShortcut",
         "Ярлык «WMS склад» открывает именно эту книгу — по привычке не откроется старая."),
        ("btnTraining", "Учебная WMS…", "WmsHelp.BtnTraining",
         "Отдельная копия с пометкой «УЧЕБНАЯ» и своим журналом: пробовать «1 кан», «Исправить», «Удалить» без риска."),
        ("btnAnonCopy", "Обезличенная копия…", "WmsHelp.BtnAnonCopy",
         "Для разработчика: наименования, люди, поставщики заменены на «Товар 1»… — ошибку можно повторить без ваших данных."),
    ]),
]
SHEETS_10 = [PLACES, SUPPLIERS, GOODS, SUPPLY, PRINT, REPORTS, CLOSE, KITS, PROP, RESERVES, SERVICE]
LEGREP = "Сверка переноса"
LEGREP_HEAD = 3                   # WmsConfig.LEGREP_HEAD
LEGREP_HEADERS = ["ЕИ", "Наименование", "Ед.", "Откуда", "Приход при переносе", "Остаток при переносе", "Израсходовано до переноса",
                  "Старое X", "Откуда взят остаток", "Движения после переноса", "Остаток по движениям", "Остаток в «Наличие»", "Разница",
                  "Коррекций", "Последняя коррекция", "Место", "Что проверить"]
LEGREP_WIDTHS = [3300, 6500, 1800, 3800, 2600, 2600, 2600, 2300, 3800, 2600, 2600, 2600, 2200, 2000, 6000, 2600, 12000]

CARS = "Приход авто"
CAR = "_CAR"
CAR_HEAD_ROW = 5                  # WmsConfig.CAR_HEAD_ROW
CAR_HEADERS = ["№ визита", "Дата", "Марка / госномер", "Поставщик", "Приезд", "Выезд", "Длительность", "Статус", "День недели",
               "Месяц / год", "№ за день", "Контроль", "Комментарий", "Ключ машины"]
CAR_WIDTHS = [1900, 2500, 6500, 6000, 2600, 2600, 2600, 3800, 3300, 2400, 2100, 9000, 5000, 3000]
CAR_SERVICE_HEADERS = ["№ визита", "Ключ машины", "Состояние", "Строка (индекс)", "Приезд", "Выезд", "День", "№ за день",
                       "Ключ открытого визита", "Приезд открытого визита", "Исправлений", "Длительность (сут.)", "Ключ поставщика",
                       "Словарь: машина", "Словарь: ключ машины", "Словарь: поставщик", "Словарь: ключ поставщика"]      # WmsConfig CR_*, CD_*
CAR_LONG_MIN = 120                # WmsConfig.CAR_LONG_MIN: the default of the threshold cell L5
CS_OPEN, CS_GONE, CS_CANCEL = "На территории", "Уехал", "Отменён"
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


def autofilter(doc, sheet_index, name, ncols, start_row=0):
    a = CellRangeAddress()
    a.Sheet, a.StartColumn, a.StartRow, a.EndColumn, a.EndRow = sheet_index, 0, start_row, ncols - 1, LAST_ROW
    doc.DatabaseRanges.addNewByName(name, a)
    doc.DatabaseRanges.getByName(name).AutoFilter = True


def add_button(doc, sheet, name, label, macro, x, y, w, h, font=None, bold=False, back=None):
    """A form push button bound to a Basic macro (no focus on click: the cell cursor stays on the user's row)."""
    model = doc.createInstance("com.sun.star.form.component.CommandButton")
    model.Name = name
    model.Label = label
    model.FocusOnClick = False
    model.Printable = False
    if font:
        model.FontHeight = font
    if bold:
        model.FontWeight = 150
    if back is not None:
        model.BackgroundColor = back
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
    ev.ScriptCode = macro_url(macro)
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


def cell_style(doc, name, **props_):
    """a cell style of the book (conditional formats refer to styles by name)"""
    fam = doc.StyleFamilies.getByName("CellStyles")
    if not fam.hasByName(name):
        fam.insertByName(name, doc.createInstance("com.sun.star.style.CellStyle"))
    st = fam.getByName(name)
    for k, v in props_.items():
        setattr(st, k, v)
    return st


def add_cond(sheet, rng, formula, style, base_col, base_row):
    """a conditional format entry «formula is true → style» (relative references are relative to base_col/base_row)"""
    from com.sun.star.table import CellAddress
    base = CellAddress()
    base.Sheet, base.Column, base.Row = sheet.RangeAddress.Sheet, base_col, base_row
    cf = rng.ConditionalFormat
    cf.addNew((props(Operator=uno.Enum("com.sun.star.sheet.ConditionOperator", "FORMULA"), Formula1=formula, StyleName=style,
                     SourcePosition=base)))
    rng.ConditionalFormat = cf


def border_line(width, color=0x000000):
    from com.sun.star.table import BorderLine2
    b = BorderLine2()
    b.Color, b.OuterLineWidth, b.LineWidth = color, width, width
    return b


def list_validation(rng, source):
    v = rng.Validation
    v.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
    v.ShowList = 2                          # sorted ascending
    v.IgnoreBlankCells = True
    v.ShowErrorMessage = False              # any text may be typed: the list only helps
    v.setFormula1(source)
    rng.Validation = v


def car_formulas():
    """the indicator formulas of the panel of «Приход авто» (row 4) and its hidden helpers (column N): {(col,row): formula}"""
    c = f"$'{CAR}'"
    o = f"$'{ORDERS}'"
    st, dep, day, oarr, dur = (f"{c}.${x}$2:${x}$1048576" for x in "CFGJL")
    thr = "$N$4"                               # the threshold in effect (the setting L5 when it is a positive number)
    mn = "$N$5"                                # the arrival of the longest open visit
    sup = 'TRIM($D$2)'
    l_, w_, q_ = f"{o}.$L$2:$L$1048576", f"{o}.$W$2:$W$1048576", f"{o}.$Q$2:$Q$1048576"
    # the open positions (with the date statuses of a book before 0.7.2 until its refresh) and those whose expected date Q
    # has passed — information for the storekeeper, never a status (D-089)
    open_st = ("Ожидается", "Частично получено", "Просрочено", "Частично получено / просрочено")
    cnt = "+".join(f'COUNTIFS({l_};"="&{sup};{w_};"{x}")' for x in open_st)
    over = "+".join(f'COUNTIFS({l_};"="&{sup};{w_};"{x}";{q_};"<"&TODAY())' for x in open_st)
    hm = f'RIGHT("0"&HOUR({mn});2)&":"&RIGHT("0"&MINUTE({mn});2)'
    return {
        (2, 2): f'=IF({sup}="";"Впишите машину и поставщика → ПРИЕХАЛ. При выезде выделите строку машины → УЕХАЛ.";'
                f'"У поставщика «"&{sup}&"» открытых позиций заказов: "&({cnt})&IF(({over})>0;", из них ожидаемая дата прошла: "&({over});""))',
        (13, 3): f"=IF(AND(ISNUMBER($L$5);$L$5>=1);$L$5;{CAR_LONG_MIN})",
        (13, 4): f'=IF(COUNT({oarr})=0;"";MIN({oarr}))',
        (2, 4): f'=COUNTIF({st};"OPEN")',
        (3, 4): f'=COUNTIFS({day};TODAY();{st};"<>CANCELLED")',
        (4, 4): f'=COUNTIFS({dep};">="&TODAY();{dep};"<"&(TODAY()+1);{st};"CLOSED")',
        (5, 4): f'=COUNTIFS({oarr};"<"&(NOW()-{thr}/1440))',
        (6, 4): f'=IFERROR(AVERAGEIFS({dur};{day};TODAY();{st};"CLOSED");"—")',
        (7, 4): f'=IF(COUNT({oarr})=0;"—";INDEX($C$7:$C$1048576;MATCH({mn};{oarr};0))&" · с "&IF(INT({mn})<TODAY();DAY({mn})&"."&'
                f'RIGHT("0"&MONTH({mn});2)&" ";"")&{hm}&" · "&INT((NOW()-{mn})*24)&" ч "&MOD(INT((NOW()-{mn})*1440);60)&" мин")',
    }


def build_cars(doc, sh):
    """«Приход авто»: the panel (frozen with the header), the table A:N of the visits, buttons, lists, conditional formats"""
    cols = sh.getColumns()
    for c, w in enumerate(CAR_WIDTHS):
        cols.getByIndex(c).Width = w
    cols.getByIndex(13).IsVisible = False                  # N: the search key (and the helpers of the panel)
    rows = sh.getRows()
    for r, hgt in enumerate((700, 1100, 650, 750, 900, 1000)):
        rows.getByIndex(r).Height = hgt
    t = sh.getCellByPosition(0, 0)
    t.setString("ПРИХОД АВТО")
    t.CharWeight, t.CharHeight = 150, 14
    for c, label in ((2, "Марка / госномер"), (3, "Поставщик")):
        x = sh.getCellByPosition(c, 0)
        x.setString(label)
        x.CharWeight, x.VertJustify = 150, 3
    text = number_format(doc, "@", ("ru", "RU"))
    inp = sh.getCellRangeByPosition(2, 1, 3, 1)
    inp.NumberFormat = text
    inp.CellProtection = protection(False)
    inp.CharHeight, inp.CharWeight, inp.CellBackColor, inp.VertJustify = 14, 150, 0xFFF2CC, 2
    for side in ("TopBorder", "BottomBorder", "LeftBorder", "RightBorder"):
        setattr(inp, side, border_line(35, 0x7F6000))
    list_validation(sh.getCellRangeByPosition(2, 1, 2, 1), f"$'{CAR}'.$N$2:$N$1048576")
    list_validation(sh.getCellRangeByPosition(3, 1, 3, 1), f"$'{CAR}'.$P$2:$P$1048576")
    hint = sh.getCellRangeByPosition(2, 2, 2, 2)
    hint.CharColor, hint.CharHeight = 0x44546A, 10
    for c, label in ((2, "На территории сейчас"), (3, "Приехало сегодня"), (4, "Уехало сегодня"), (5, "Дольше порога"),
                     (6, "Среднее время сегодня"), (7, "Самая долгая текущая машина"), (11, "Порог долгой стоянки, мин")):
        x = sh.getCellByPosition(c, 3)
        x.setString(label)
        x.IsTextWrapped, x.CharColor, x.CharHeight, x.VertJustify = True, 0x595959, 9, 3
    vals = sh.getCellRangeByPosition(2, 4, 11, 4)
    vals.CharHeight, vals.CharWeight, vals.VertJustify = 13, 150, 2
    vals.NumberFormat = number_format(doc, "0")
    sh.getCellByPosition(6, 4).NumberFormat = number_format(doc, "[HH]:MM")
    for (c, r), f in car_formulas().items():
        sh.getCellByPosition(c, r).setFormula(f)
    thr = sh.getCellByPosition(11, 4)
    thr.setValue(CAR_LONG_MIN)
    thr.CellProtection = protection(False)
    thr.CellBackColor = 0xFFF2CC
    # the header of the table (row 5) and its data rows
    for c, h in enumerate(CAR_HEADERS):
        sh.getCellByPosition(c, CAR_HEAD_ROW).setString(h)
    hdr = sh.getCellRangeByPosition(0, CAR_HEAD_ROW, len(CAR_HEADERS) - 1, CAR_HEAD_ROW)
    hdr.CharWeight, hdr.IsTextWrapped, hdr.VertJustify, hdr.CellBackColor = 150, True, 3, 0xE7E6E6
    first = CAR_HEAD_ROW + 1
    sh.getCellRangeByPosition(0, first, 0, LAST_ROW).NumberFormat = number_format(doc, "0")
    sh.getCellRangeByPosition(1, first, 1, LAST_ROW).NumberFormat = number_format(doc, "DD.MM.YYYY")
    sh.getCellRangeByPosition(4, first, 5, LAST_ROW).NumberFormat = number_format(doc, "HH:MM")
    sh.getCellRangeByPosition(6, first, 6, LAST_ROW).NumberFormat = number_format(doc, "[HH]:MM")
    sh.getCellRangeByPosition(10, first, 10, LAST_ROW).NumberFormat = number_format(doc, "0")
    sh.getCellRangeByPosition(12, first, 12, LAST_ROW).CellProtection = protection(False)      # M: the comment of the user
    # a vehicle on the territory is highlighted, a long stay (over the threshold of L5) in red, a cancelled visit greyed
    cell_style(doc, "WMS_CarOpen", CellBackColor=0xE2EFDA)
    cell_style(doc, "WMS_CarLong", CellBackColor=0xFFC7CE, CharColor=0x9C0006, CharWeight=150)
    cell_style(doc, "WMS_CarCancel", CharColor=0x808080, CharStrikeout=1)
    table = sh.getCellRangeByPosition(0, first, 12, LAST_ROW)
    add_cond(sh, table, f'AND($H{first + 1}="{CS_OPEN}";(NOW()-$E{first + 1})*1440>$N$4)', "WMS_CarLong", 0, first)
    add_cond(sh, table, f'$H{first + 1}="{CS_OPEN}"', "WMS_CarOpen", 0, first)
    add_cond(sh, table, f'$H{first + 1}="{CS_CANCEL}"', "WMS_CarCancel", 0, first)
    add_cond(sh, sh.getCellRangeByPosition(7, 4, 7, 4), 'AND(ISNUMBER($N$5);(NOW()-$N$5)*1440>$N$4)', "WMS_CarLong", 7, 4)
    add_cond(sh, sh.getCellRangeByPosition(5, 4, 5, 4), '$F$5>0', "WMS_CarLong", 5, 4)
    # the buttons: ПРИЕХАЛ over E1:F2, УЕХАЛ over G1:H2, the compact ones over I1:K2
    def x_(c):
        return sh.getCellByPosition(c, 0).Position.X
    y0, y1 = sh.getCellByPosition(0, 0).Position.Y, sh.getCellByPosition(0, 1).Position.Y
    hh = sh.getCellByPosition(0, 2).Position.Y - y0
    add_button(doc, sh, "btnCarArrive", "ПРИЕХАЛ", "WmsCarUi.BtnCarArrive", x_(4) + 60, y0 + 60, x_(6) - x_(4) - 120, hh - 120,
               font=18, bold=True, back=0xC6EFCE)
    add_button(doc, sh, "btnCarDepart", "УЕХАЛ", "WmsCarUi.BtnCarDepart", x_(6) + 60, y0 + 60, x_(8) - x_(6) - 120, hh - 120,
               font=18, bold=True, back=0xF8CBAD)
    small = ((("btnCarFind", "Найти", "WmsCarUi.BtnCarFind", 8, 9), ("btnCarToday", "Сегодня", "WmsCarUi.BtnCarToday", 9, 10),
              ("btnCarAll", "Все", "WmsCarUi.BtnCarAll", 10, 11)),
             (("btnCarFix", "Исправить", "WmsCarUi.BtnCarFix", 8, 9), ("btnCarCancel", "Отменить визит", "WmsCarUi.BtnCarCancel", 9, 11)))
    for i, row in enumerate(small):
        y = y0 if i == 0 else y1
        h = (y1 - y0) if i == 0 else (sh.getCellByPosition(0, 2).Position.Y - y1)
        for name, label, macro, c0, c1 in row:
            add_button(doc, sh, name, label, macro, x_(c0) + 40, y + 40, x_(c1) - x_(c0) - 80, h - 80)


def build_car_service(doc, sh):
    """_CAR: the visits (row = visit №) and the dictionaries of the input lists with their sizes in S1 and U1"""
    for c, h in enumerate(CAR_SERVICE_HEADERS):
        sh.getCellByPosition(c, 0).setString(h)
    sh.getCellByPosition(17, 0).setString("Машин в словаре")
    sh.getCellByPosition(18, 0).setValue(0)
    sh.getCellByPosition(19, 0).setString("Поставщиков в словаре")
    sh.getCellByPosition(20, 0).setValue(0)
    text = number_format(doc, "@", ("ru", "RU"))
    for c in (1, 2, 8, 12, 13, 14, 15, 16):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text


def idx_car_formulas():
    """_IDX rows IX_CO_MATCH..IX_CS_MATCH (WmsConfig): the open visit of a vehicle key, their count, the visits of a key, the
    dictionaries"""
    c = f"$'{CAR}'"
    return [
        ("_CAR.I: MATCH", f"=MATCH($B$1;INDEX({c}.$I:$I;$B$2+2):INDEX({c}.$I:$I;1048576);0)+$B$2"),
        ("_CAR.I: COUNTIF", f"=COUNTIF({c}.$I$2:$I$1048576;$B$1)"),
        ("_CAR.B: COUNTIF", f"=COUNTIF({c}.$B$2:$B$1048576;$B$1)"),
        ("_CAR.O: MATCH", f"=MATCH($B$1;INDEX({c}.$O:$O;$B$2+2):INDEX({c}.$O:$O;1048576);0)+$B$2"),
        ("_CAR.Q: MATCH", f"=MATCH($B$1;INDEX({c}.$Q:$Q;$B$2+2):INDEX({c}.$Q:$Q;1048576);0)+$B$2"),
    ]


def build_stock(doc, sh, rows):
    header(sh, STOCK_HEADERS, STOCK_WIDTHS)
    sh.getCellRangeByPosition(0, 1, 0, LAST_ROW).NumberFormat = number_format(doc, "@", ("ru", "RU"))
    if rows:
        sh.getCellRangeByPosition(0, 1, len(STOCK_HEADERS) - 1, len(rows)).setDataArray(rows)
    stock_buttons(doc, sh)


def stock_buttons(doc, sh):
    """PRIME 0.8: the buttons of «Наличие» in the upper part of its header row (the registry stays read-only)"""
    sh.getRows().getByIndex(0).Height = 1700
    sh.getCellRangeByPosition(0, 0, len(STOCK_HEADERS) - 1, 0).VertJustify = 3
    add_button(doc, sh, "btnMigrateFix", "Исправить ЕИ переноса", "WmsPrimeUi.BtnMigrateFix", 150, 120, 4600, 650)
    add_button(doc, sh, "btnLegacyReport", "Сверка переноса", "WmsPrimeUi.BtnLegacyReport", 4900, 120, 3600, 650)
    add_button(doc, sh, "btnPackEI", "Фасовка ЕИ", "WmsPack.BtnPackEI", 8650, 120, 3000, 650, back=0xC6E0B4)


def build_leg(doc, sh):
    """PRIME 0.8: _LEG — the registry of the transfer, one row per EI that came from the old table (row = EI number)"""
    for c, h in enumerate(LEG_HEADERS):
        sh.getCellByPosition(c, 0).setString(h)
    text = number_format(doc, "@", ("ru", "RU"))
    for c in (0, 1, 5, 6, 7, 10):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text


def build_pack(doc, sh):
    """PRIME 0.8: _PACK — the packing of an EI («в 1 шт — 4 кан», «1 шт = 6 м»), one row per EI (row = EI number)"""
    for c, h in enumerate(PACK_HEADERS):
        sh.getCellByPosition(c, 0).setString(h)
    text = number_format(doc, "@", ("ru", "RU"))
    for c in (0, 1, 3):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text


def build_legrep(doc, sh):
    """PRIME 0.8: «Сверка переноса» — the read-only report of the transfer, rebuilt by its button (WmsLegacy.LegacyReport)"""
    t = sh.getCellByPosition(0, 0)
    t.setString("Сверка переноса и остатков")
    t.CharWeight = 150
    t.CharHeight = 14
    sh.getRows().getByIndex(0).Height = 1000
    sh.getCellByPosition(0, 1).setString("Нажмите «Обновить сверку»: WMS сравнит остаток каждого перенесённого ЕИ с переносом и движениями после него")
    sh.getRows().getByIndex(1).Height = 1200
    note = sh.getCellRangeByPosition(0, 1, len(LEGREP_HEADERS) - 1, 1)
    note.merge(True)
    note.IsTextWrapped = True
    note.VertJustify = 1
    for c, (h, w) in enumerate(zip(LEGREP_HEADERS, LEGREP_WIDTHS)):
        cell = sh.getCellByPosition(c, LEGREP_HEAD)
        cell.setString(h)
        sh.getColumns().getByIndex(c).Width = w
    rng = sh.getCellRangeByPosition(0, LEGREP_HEAD, len(LEGREP_HEADERS) - 1, LEGREP_HEAD)
    rng.CharWeight = 150
    rng.IsTextWrapped = True
    rng.CellBackColor = 0xE7E6E6
    sh.getCellRangeByPosition(len(LEGREP_HEADERS) - 1, LEGREP_HEAD + 1, len(LEGREP_HEADERS) - 1, LAST_ROW).IsTextWrapped = True
    add_button(doc, sh, "btnLegacyRefresh", "Обновить сверку", "WmsPrimeUi.BtnLegacyReport", 9500, 100, 4200, 800, bold=True, back=0xDDEBF7)


# ================================================================ PRIME 0.8: «Дашборд», «Поиск», «Карта склада», «Топология»

INK = 0x1F3864          # titles
TEXT = 0x263238         # text
MUTED = 0x5B6B7A        # captions, notes
LINE = 0xD5DDE6         # thin lines
PANEL = 0xF3F6FA        # tiles


MAP_COLORS = {      # WmsMap.LayerColors: (background, text) of the codes 0 … 4 of every layer
    "OCC": [(0xF4F6F8, 0x263238), (0xC6E8C9, 0x263238), (0x8FD19E, 0x263238), (0x4CAF6A, 0xFFFFFF), (0x2E7D4F, 0xFFFFFF)],
    "ISS": [(0xE8ECF0, 0x263238), (0xFFE8A3, 0x263238), (0xFFC061, 0x263238), (0xF28C38, 0xFFFFFF), (0xD9472B, 0xFFFFFF)],
    "STALE": [(0xF4F6F8, 0x263238), (0xDCEBFA, 0x263238), (0x9CC3EB, 0x263238), (0x5B97D6, 0xFFFFFF), (0x2F6DB5, 0xFFFFFF)],
    "ATT": [(0xF4F6F8, 0x263238), (0xE8ECF0, 0x263238), (0xFFC061, 0x263238), (0xD9472B, 0xFFFFFF), (0xE8ECF0, 0x263238)],
    "FREE": [(0xE8ECF0, 0x263238), (0x8FD19E, 0x263238), (0xE8ECF0, 0x263238), (0xE8ECF0, 0x263238), (0xE8ECF0, 0x263238)],
    # 1.0: «Заполненность» (М4) and «Товар» (М5)
    "FILL": [(0xF4F6F8, 0x263238), (0xC6E8C9, 0x263238), (0xE8ECF0, 0x263238), (0xFFC061, 0x263238), (0xD9472B, 0xFFFFFF)],
    "FIND": [(0xF4F6F8, 0x263238), (0xE8ECF0, 0x263238), (0xE8ECF0, 0x263238), (0xE8ECF0, 0x263238), (0x2F6DB5, 0xFFFFFF)],
}


def view_styles(doc):
    for layer, cols in MAP_COLORS.items():
        for k, (bg, ink) in enumerate(cols):
            cell_style(doc, f"MAP_{layer}_{k}", CellBackColor=bg, CharColor=ink)
    cell_style(doc, "DB_Good", CellBackColor=0xE3F4E6, CharColor=0x1E6B34)
    cell_style(doc, "DB_Alert", CellBackColor=0xFDE2DD, CharColor=0xB3261E)
    cell_style(doc, "DB_Warn", CellBackColor=0xFFF2CC, CharColor=0x7F5F00)
    cell_style(doc, "DB_StateOk", CharColor=0x1E6B34)
    cell_style(doc, "DB_StateBad", CharColor=0xB3261E)


def box(rng, color=LINE, width=18):
    from com.sun.star.table import TableBorder2
    b = TableBorder2()
    ln = border_line(width, color)
    b.TopLine = b.BottomLine = b.LeftLine = b.RightLine = ln
    b.IsTopLineValid = b.IsBottomLineValid = b.IsLeftLineValid = b.IsRightLineValid = True
    rng.TableBorder2 = b


def underline(rng, color=LINE, width=18):
    rng.BottomBorder = border_line(width, color)


def text_cell(sh, c, r, s, size=10, bold=False, color=TEXT, wrap=False, valign=None, halign=None):
    cell = sh.getCellByPosition(c, r)
    cell.setString(s)
    cell.CharHeight = size
    cell.CharWeight = 150 if bold else 100
    cell.CharColor = color
    if wrap:
        cell.IsTextWrapped = True
    if valign is not None:
        cell.VertJustify = valign
    if halign is not None:
        cell.HoriJustify = halign
    return cell


def merge(sh, c0, r0, c1, r1):
    rng = sh.getCellRangeByPosition(c0, r0, c1, r1)
    rng.merge(True)
    return rng


VIEW_NAV = (("Заказы", "WmsUi.NavOrders"), ("Иной приход", "WmsUi.NavSpecial"), ("Приход авто", "WmsUi.NavCars"), ("Выдачи", "WmsUi.NavIssues"),
            ("Возврат", "WmsUi.NavReturns"), ("Корректировки", "WmsUi.NavAdjust"), ("Наличие", "WmsUi.NavStock"), ("Поиск", "WmsUi.NavFind"),
            ("Карта склада", "WmsUi.NavMap"), ("Дашборд", "WmsUi.NavDash"), ("Главная", "WmsUi.NavMain"))


def view_nav(doc, sh, row, prefix, skip, x0=150, total=30000, h=760):
    """a strip of the ways to the working sheets in `row` (the page itself left out)"""
    items = [x for x in VIEW_NAV if x[0] != skip]
    w = (total - 60 * (len(items) - 1)) // len(items)
    y = sh.getCellByPosition(0, row).Position.Y
    for i, (label, macro) in enumerate(items):
        add_button(doc, sh, f"{prefix}{i}", label, macro, x0 + i * (w + 60), y + 70, w, h, back=0xDDEBF7)


DB_TILES = ["НА СКЛАДЕ — ЕИ С ОСТАТКОМ", "ОЖИДАЕТСЯ ПОЗИЦИЙ", "ПРИШЛО СЕГОДНЯ", "ВЫДАНО СЕГОДНЯ", "МАШИНЫ НА ТЕРРИТОРИИ", "ВНИМАНИЕ"]


def build_dash(doc, sh):
    """«Дашборд» (WmsDash): six tiles, four lists, the last 14 days — the cells WmsDash.DashRefresh fills"""
    view_styles(doc)
    sh.getColumns().getByIndex(0).Width = 400
    for c in range(1, 13):
        sh.getColumns().getByIndex(c).Width = 2500
    sh.getColumns().getByIndex(13).Width = 400
    sh.getColumns().getByIndex(14).IsVisible = False            # WmsDash.DB_SVC_COL
    heights = {0: 1150, 1: 950, 2: 250, 3: 520, 4: 1250, 5: 820, 6: 350, 7: 650, 8: 480, 21: 350, 22: 650, 23: 480, 34: 350, 35: 650, 36: 480}
    for r, h in heights.items():
        sh.getRows().getByIndex(r).Height = h
    for r in list(range(9, 21)) + list(range(24, 34)) + list(range(37, 51)):
        sh.getRows().getByIndex(r).Height = 440
    merge(sh, 1, 0, 7, 0)
    text_cell(sh, 1, 0, "Склад сегодня", size=20, bold=True, color=INK, valign=2)
    merge(sh, 8, 0, 12, 0)
    st = text_cell(sh, 8, 0, "●  WMS запускается…", size=12, bold=True, color=MUTED, valign=2, halign=3)
    add_cond(sh, sh.getCellRangeByPosition(8, 0, 12, 0), 'LEFT($O$1;5)="CLEAN"', "DB_StateOk", 8, 0)
    add_cond(sh, sh.getCellRangeByPosition(8, 0, 12, 0), 'AND($O$1<>"";LEFT($O$1;5)<>"CLEAN")', "DB_StateBad", 8, 0)
    whole = number_format(doc, "0")
    for i, cap in enumerate(DB_TILES):
        c0 = 1 + 2 * i
        merge(sh, c0, 3, c0 + 1, 3)
        merge(sh, c0, 4, c0 + 1, 4)
        merge(sh, c0, 5, c0 + 1, 5)
        tile = sh.getCellRangeByPosition(c0, 3, c0 + 1, 5)
        tile.CellBackColor = PANEL
        box(tile)
        text_cell(sh, c0, 3, cap, size=8, bold=True, color=MUTED, valign=3, halign=2, wrap=True)
        v = text_cell(sh, c0, 4, "—", size=26, bold=True, color=INK, valign=2, halign=2)
        v.NumberFormat = whole
        text_cell(sh, c0, 5, "", size=8, color=MUTED, wrap=True, valign=1, halign=2)
    att = sh.getCellRangeByPosition(11, 3, 12, 5)
    add_cond(sh, att, "$L$5>0", "DB_Alert", 11, 3)
    add_cond(sh, att, 'AND(ISNUMBER($L$5);$L$5=0)', "DB_Good", 11, 3)
    # the four lists and the days
    def section(r, c0, c1, title):
        merge(sh, c0, r, c1, r)
        text_cell(sh, c0, r, title, size=12, bold=True, color=INK, valign=3)
        underline(sh.getCellRangeByPosition(c0, r, c1, r), INK, 26)

    def heads(r, cols, right=()):
        # 1.0: the headers of the numbers stand over them (right-aligned, as the numbers are)
        for c, s in cols:
            text_cell(sh, c, r, s, size=8, bold=True, color=MUTED, valign=3, halign=(3 if c in right else 1))

    section(7, 1, 6, "Ожидается: срок прошёл и ближайшие поставки")
    section(7, 7, 12, "Последние операции")
    heads(8, [(1, "Заказ"), (2, "Наименование"), (4, "Поставщик"), (5, "Ожидалось"), (6, "Срок"), (7, "Когда"), (8, "Операция"),
              (10, "Что"), (12, "Кол-во")], right=(12,))
    section(22, 1, 6, "Чаще всего выдают — 30 дней")
    section(22, 7, 12, "Лежит без движения 90+ дней")
    heads(23, [(1, "Наименование"), (4, "Выдач"), (5, "Количество"), (6, "Ед."), (7, "ЕИ"), (8, "Наименование"), (11, "Остаток"), (12, "Дней")],
          right=(4, 5, 11, 12))
    section(35, 1, 12, "Последние 14 дней: выдачи и приходы")
    heads(36, [(1, "День"), (2, "Выдач"), (8, "Приходов")])
    for r in range(9, 21):
        merge(sh, 2, r, 3, r)
        merge(sh, 8, r, 9, r)
        merge(sh, 10, r, 11, r)
        underline(sh.getCellRangeByPosition(1, r, 12, r), 0xEEF1F4, 10)
    for r in range(24, 34):
        merge(sh, 1, r, 3, r)
        merge(sh, 8, r, 10, r)
        underline(sh.getCellRangeByPosition(1, r, 12, r), 0xEEF1F4, 10)
    for r in range(37, 51):
        merge(sh, 3, r, 7, r)
        merge(sh, 9, r, 12, r)
    for r0, r1 in ((9, 20), (24, 33), (37, 50)):
        lists = sh.getCellRangeByPosition(1, r0, 12, r1)
        lists.CharHeight = 9
        lists.CharColor = TEXT
        lists.VertJustify = 2
    for c0, c1, r0, r1 in ((12, 12, 9, 20), (4, 5, 24, 33), (11, 12, 24, 33)):
        sh.getCellRangeByPosition(c0, r0, c1, r1).HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "RIGHT")
    sh.getCellRangeByPosition(6, 9, 6, 20).CharHeight = 8
    add_cond(sh, sh.getCellRangeByPosition(6, 9, 6, 20), 'LEFT($G10;10)="просрочено"', "DB_StateBad", 6, 9)
    sh.getCellRangeByPosition(3, 37, 7, 50).CharColor = 0x3A7BD5
    sh.getCellRangeByPosition(9, 37, 12, 50).CharColor = 0x34A853
    sh.getCellRangeByPosition(2, 37, 2, 50).HoriJustify = 3
    sh.getCellRangeByPosition(8, 37, 8, 50).HoriJustify = 3
    for c in (4, 5, 11, 12):
        sh.getCellRangeByPosition(c, 24, c, 33).HoriJustify = 3
    sh.getCellRangeByPosition(12, 9, 12, 20).HoriJustify = 3
    # 1.0 (Ч5): the tails — a balance below one part (WmsDash.DB_TAILS_HEAD … ), «Списать хвосты…»
    section(54, 1, 12, "Хвосты — остаток меньше одной части")
    heads(55, [(1, "ЕИ"), (2, "Наименование"), (6, "Остаток"), (9, "Место")])
    for r in range(56, 66):
        sh.getRows().getByIndex(r).Height = 440
        merge(sh, 2, r, 5, r)
        merge(sh, 6, r, 8, r)
        merge(sh, 9, r, 12, r)
        underline(sh.getCellRangeByPosition(1, r, 12, r), 0xEEF1F4, 10)
    tl = sh.getCellRangeByPosition(1, 56, 12, 65)
    tl.CharHeight = 9
    tl.CharColor = TEXT
    tl.VertJustify = 2
    for r, h in ((53, 350), (54, 650), (55, 480), (66, 350)):
        sh.getRows().getByIndex(r).Height = h
    y = sh.getCellByPosition(0, 54).Position.Y
    add_button(doc, sh, "btnTails", "Списать хвосты…", "WmsDash.BtnTails", sh.getCellByPosition(10, 54).Position.X, y + 40, 7400, 560, back=0xF4CCCC)
    merge(sh, 1, 67, 12, 67)
    text_cell(sh, 1, 67, "Страница считается по листам книги при открытии и после каждой операции; она ничего не изменяет.", size=8, color=MUTED)
    view_nav(doc, sh, 1, "dbNav", DASH, total=26900)
    y = sh.getCellByPosition(0, 1).Position.Y
    add_button(doc, sh, "btnDashRefresh", "Обновить", "WmsDash.BtnDashRefresh", 150 + 27000, y + 70, 2850, 760, bold=True, back=0xC6E0B4)
    sh.getCellRangeByPosition(0, 0, 14, 70).CellProtection = protection(True)


def build_find(doc, sh):
    """«Поиск» (WmsFind): the field, the results or the card of an EI"""
    view_styles(doc)
    widths = [400, 2700, 4000, 3300, 7400, 2300, 1700, 2600, 9000, 400]
    for c, w in enumerate(widths):
        sh.getColumns().getByIndex(c).Width = w
    for c in (10, 11, 12):                         # WmsFind.FD_SRC_SHEET … FD_SRC_EI
        sh.getColumns().getByIndex(c).IsVisible = False
    for r, h in {0: 1100, 1: 950, 2: 250, 3: 900, 4: 600, 5: 250, 6: 800, 7: 520, 8: 520, 9: 600, 10: 250, 11: 560}.items():
        sh.getRows().getByIndex(r).Height = h
    merge(sh, 1, 0, 8, 0)
    text_cell(sh, 1, 0, "Поиск и карточка ЕИ", size=20, bold=True, color=INK, valign=2)
    text_cell(sh, 1, 3, "Что найти", size=12, bold=True, color=INK, valign=2, halign=3)
    inp = merge(sh, 2, 3, 5, 3)
    cell = sh.getCellByPosition(2, 3)
    cell.CharHeight = 13
    cell.VertJustify = 2
    cell.NumberFormat = number_format(doc, "@", ("ru", "RU"))
    inp.CellBackColor = 0xFFF8E1
    box(inp, 0xE0B84C, 26)
    merge(sh, 1, 4, 8, 4)
    text_cell(sh, 1, 4, "ЕИ (427, ЕИ-427), наименование, артикул, № заказа, поставщик, получатель, место (5-12-3), документ, номер машины — "
                        "впишите и нажмите Enter. Ищется по всем рабочим листам.", size=9, color=MUTED, wrap=True, valign=1)
    merge(sh, 1, 6, 8, 6)
    text_cell(sh, 1, 6, "Впишите, что найти, и нажмите Enter", size=15, bold=True, color=INK, valign=2)
    for r in (7, 8, 9):
        merge(sh, 1, r, 8, r)
        text_cell(sh, 1, r, "", size=10, color=TEXT, valign=2)
    add_cond(sh, sh.getCellRangeByPosition(1, 9, 8, 9), 'LEFT($B$10;1)="✓"', "DB_Good", 1, 9)
    add_cond(sh, sh.getCellRangeByPosition(1, 9, 8, 9), 'LEFT($B$10;1)="✗"', "DB_Alert", 1, 9)
    hdr = sh.getCellRangeByPosition(1, 11, 8, 11)
    hdr.CharWeight = 150
    hdr.CharHeight = 9
    hdr.CharColor = MUTED
    hdr.CellBackColor = 0xE7ECF2
    hdr.VertJustify = 2
    body = sh.getCellRangeByPosition(1, 12, 8, 3100)
    body.CharHeight = 9
    body.CharColor = TEXT
    sh.getCellRangeByPosition(2, 12, 2, 3100).HoriJustify = 2
    view_nav(doc, sh, 1, "fdNav", FIND, total=30500)
    y = sh.getCellByPosition(0, 3).Position.Y
    x = sh.getCellByPosition(6, 3).Position.X
    add_button(doc, sh, "btnFind", "Найти", "WmsFind.BtnFind", x + 80, y + 90, 2600, 720, bold=True, back=0xC6E0B4)
    add_button(doc, sh, "btnFindCard", "Карточка ЕИ", "WmsFind.BtnFindCard", x + 2760, y + 90, 3300, 720, back=0xDDEBF7)
    add_button(doc, sh, "btnFindGo", "Перейти к строке", "WmsFind.BtnFindGo", x + 6140, y + 90, 3800, 720, back=0xDDEBF7)
    add_button(doc, sh, "btnFindPack", "Фасовка ЕИ", "WmsPack.BtnPackEI", x + 10020, y + 90, 3000, 720, back=0xDDEBF7)
    sh.getCellRangeByPosition(0, 0, 12, 3100).CellProtection = protection(True)
    sh.getCellByPosition(2, 3).CellProtection = protection(False)


MAP_LAYERS = (("Занятость", "WmsMap.BtnMapOcc"), ("Выдачи 30 дней", "WmsMap.BtnMapIssues"), ("Без движения", "WmsMap.BtnMapStale"),
              ("Внимание", "WmsMap.BtnMapAttention"), ("Свободные", "WmsMap.BtnMapFree"), ("Не на карте", "WmsMap.BtnMapOff"),
              ("Уместить", "WmsMap.BtnMapFit"), ("Обновить", "WmsMap.BtnMapRefresh"), ("Загрузить топологию…", "WmsMap.BtnTopoLoad"))


def build_map(doc, sh):
    """«Карта склада» (WmsMap): the header (title, the layers, the summary, the place under the cursor, the legend); the map
    itself is the layout of «Топология», copied below MP_TOP by WmsMap.MapBuild"""
    for r, h in {0: 1100, 1: 950, 2: 900, 3: 560, 4: 640, 5: 520, 6: 250}.items():
        sh.getRows().getByIndex(r).Height = h
    text_cell(sh, 0, 0, "Карта склада", size=20, bold=True, color=INK, valign=2)
    text_cell(sh, 0, 3, "Карта строится по листу «Топология»: вставьте туда свою схему склада (места подписаны «ряд-секция-ярус», "
                        "например 5-12-3) или нажмите «Загрузить топологию…».", size=10, color=MUTED, valign=2)
    text_cell(sh, 0, 4, "", size=11, bold=True, color=INK, valign=2)
    text_cell(sh, 0, 5, "", size=10, color=TEXT, valign=2)
    view_nav(doc, sh, 1, "mpNav", MAP, total=30000)
    y = sh.getCellByPosition(0, 2).Position.Y
    x = 150
    for i, (label, macro) in enumerate(MAP_LAYERS):
        w = (3500 if i < 5 else 2900) if i < 8 else 4300
        add_button(doc, sh, f"mpLayer{i}", label, macro, x, y + 80, w, 740, bold=(i < 5), back=(0xFCE8B2 if i < 5 else 0xDDEBF7))
        x += w + (300 if i == 4 else 60)
    sh.getCellRangeByPosition(0, 0, 200, 3000).CellProtection = protection(True)


def build_topo(sh):
    """«Топология»: the layout of the warehouse as the storekeeper drew it (not protected: it is theirs to edit)"""
    c = sh.getCellByPosition(0, 0)
    c.setString("Вставьте сюда схему склада (Ctrl+V в ячейку A1) или загрузите её файл кнопкой «Загрузить топологию…» на листе «Карта склада». "
                "Места подписываются «ряд-секция-ярус», например 5-12-3; остальное (стены, проходы, подписи) остаётся как нарисовано.")
    c.CharColor = MUTED
    c.CharPosture = uno.Enum("com.sun.star.awt.FontSlant", "ITALIC")


def status_colors(doc, sheets):
    """PRIME 0.8: the state of a row seen at a glance — the status / control cell coloured (the cell only, the row stays calm):
    received — green, expected — blue, partly — amber, cancelled or deleted — grey, not posted or a copy — red"""
    cell_style(doc, "ST_Done", CellBackColor=0xE3F4E6, CharColor=0x1E6B34)
    cell_style(doc, "ST_Wait", CellBackColor=0xE3EEFB, CharColor=0x1F4E8C)
    cell_style(doc, "ST_Part", CellBackColor=0xFFF1D6, CharColor=0x8A5A00)
    cell_style(doc, "ST_Gray", CellBackColor=0xEEF1F4, CharColor=0x6B7785)
    cell_style(doc, "ST_Bad", CellBackColor=0xFDE2DD, CharColor=0xB3261E)

    def col(sheet, c, rules, first=1):
        sh = sheets.getByName(sheet)
        letter = chr(65 + c) if c < 26 else "A" + chr(65 + c - 26)
        rng = sh.getCellRangeByPosition(c, first, c, LAST_ROW)
        if sheet == ORDERS:
            # «Заказы» already carries the line above every order (order_separators) and a conditional format applies its
            # first true entry only: the start of an order with a status gets both (the line and the colour), then the line
            # alone, then the colour alone — as Q does with its mark of a passed date
            cf = rng.ConditionalFormat
            cf.clear()
            rng.ConditionalFormat = cf
            start = 'ROW()>2;TRIM($A2)="1";$AC2=""'
            line = border_line(88, 0x1F3864)
            for test, style in rules:
                st = doc.StyleFamilies.getByName("CellStyles").getByName(style)
                cell_style(doc, style + "_Line", CellBackColor=st.CellBackColor, CharColor=st.CharColor, TopBorder=line)
                add_cond(sh, rng, f"AND({start};{test.replace('X', f'${letter}{first + 1}')})", style + "_Line", c, first)
            add_cond(sh, rng, f"AND({start})", "WMS_OrderStart", c, first)
        for test, style in rules:
            add_cond(sh, rng, test.replace("X", f"${letter}{first + 1}"), style, c, first)

    # 1.0: the preview of a row — «Ошибка: …» red, «Внимание: …» amber (it was without colour)
    posted = [('LEFT(X;9)="Проведено"', "ST_Done"), ('LEFT(X;7)="Удалено"', "ST_Gray"), ('LEFT(X;5)="КОПИЯ"', "ST_Bad"),
              ('LEFT(X;12)="Не проведено"', "ST_Bad"), ('LEFT(X;6)="Ошибка"', "ST_Bad"), ('LEFT(X;8)="Внимание"', "ST_Part")]
    col(ORDERS, 22, [('OR(LEFT(X;8)="Получено";X="Дополнительное поступление")', "ST_Done"), ('X="Ожидается"', "ST_Wait"), ('LEFT(X;8)="Частично"', "ST_Part"),
                     ('OR(X="Отменено";LEFT(X;19)="Поступление удалено")', "ST_Gray")])
    col(ORDERS, 24, [('LEFT(X;5)="КОПИЯ"', "ST_Bad"), ('LEFT(X;12)="Не проведено"', "ST_Bad"), ('LEFT(X;6)="Ошибка"', "ST_Bad"),
                     ('LEFT(X;8)="Внимание"', "ST_Part")])
    col(ISSUES, 17, posted)
    col(RETURNS, 13, posted)
    col(ADJUST, 16, posted)
    col(SPECIAL, 16, posted)
    col(SPECIAL, 17, [('LEFT(X;5)="КОПИЯ"', "ST_Bad"), ('LEFT(X;12)="Не проведено"', "ST_Bad"), ('LEFT(X;6)="Ошибка"', "ST_Bad"),
                      ('LEFT(X;8)="Внимание"', "ST_Part")])
    col(STOCK, 7, [('X="Требует разбора"', "ST_Part"), ('LEFT(X;13)="Приход удалён"', "ST_Gray")])


def fit_headers(sh, titles, height):
    """the header row at least `height` high; each column at least as wide as the longest word of its header (bold 10 pt)"""
    row = sh.getRows().getByIndex(0)
    if row.Height < height:
        row.Height = height
    for c, t in enumerate(titles):
        word = max((len(x) for x in t.replace("/", " ").split()), default=0)
        need = 255 * word + 350
        col = sh.getColumns().getByIndex(c)
        if col.Width < need:
            col.Width = need


def build_places(doc, sh):
    """«Места» (1.0, М2): the places outside the racks of the map — a floor, the yard, an office: WMS knows them (not «не на
    карте»), «Куда положить?» does not propose a closed one. The storekeeper's own sheet: not protected, not journaled"""
    header(sh, PLACES_HEADERS, [4200, 3000, 3000, 3600, 2600, 9000])
    text = number_format(doc, "@", ("ru", "RU"))
    sh.getCellRangeByPosition(0, 1, 0, LAST_ROW).NumberFormat = text
    list_validation(sh.getCellRangeByPosition(1, 1, 1, 5000), ";".join(f'"{k}"' for k in PLACE_KINDS))
    list_validation(sh.getCellRangeByPosition(4, 1, 4, 5000), '"да"')
    hint = sh.getCellByPosition(7, 0)
    hint.setString("Сюда — места вне стеллажей карты (пол, улица, офис, «у ворот»): одна строка — одно место. "
                   "«Закрыто» = да — место не предлагается в «Куда положить?». Места стеллажей задаёт лист «Топология».")
    hint.CharColor = MUTED
    hint.IsTextWrapped = True
    sh.getColumns().getByIndex(7).Width = 11000


def build_suppliers(doc, sh):
    """«Поставщики» (1.0, З8): one spelling of every supplier — the list of «Заказы» L and «Приход авто»; the usual delivery
    time gives the expected date of a new order made by WMS (a replacement, a repeat). The storekeeper's own sheet"""
    header(sh, SUPPLIER_HEADERS, [7000, 3600, 4200, 3400, 9000], height=900)
    text = number_format(doc, "@", ("ru", "RU"))
    sh.getCellRangeByPosition(0, 1, 2, LAST_ROW).NumberFormat = text
    sh.getCellRangeByPosition(3, 1, 3, LAST_ROW).NumberFormat = number_format(doc, "0")
    hint = sh.getCellByPosition(6, 0)
    hint.setString("Одна строка — один поставщик, одним написанием: тогда «Заказы», «Приход авто» и «Ожидаем» его не путают. "
                   "«Пополнить из заказов» добавит поставщиков, которых здесь ещё нет.")
    hint.CharColor = MUTED
    hint.IsTextWrapped = True
    sh.getColumns().getByIndex(6).Width = 11000


GOODS_HINT = ("Товар — строка: артикул (или наименование, если артикула нет) объединяет все его ЕИ. «Минимум» и «Заказывать по» — для "
              "«Нужно заказать» на странице «Снабжение». «В упаковке, шт» — сколько штук в упаковке: на этикетке такого товара будет "
              "«Выдача по 1 шт · в уп. 10». Серые колонки WMS заполняет кнопкой «Обновить».")


def build_goods(doc, sh):
    """«Номенклатура» (1.0, З10 / Ф3): the goods (an article, or the name when there is none, joins the EIs of one goods):
    the minimum and the batch of «Нужно заказать», the usual supplier, the default packing, the photo; M..P — written by WMS"""
    header(sh, GOODS_HEADERS, [2200, 7000, 3000, 1600, 3000, 2600, 2600, 2200, 2400, 4800, 4200, 5000, 2600, 2400, 2400, 2800], height=1300)
    text = number_format(doc, "@", ("ru", "RU"))
    for c in (0, 1, 2, 3, 4, 6, 9, 10, 11):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
    n = len(GOODS_HEADERS)
    wms = sh.getCellRangeByPosition(n - GOODS_WMS_COLS, 0, n - 1, 0)
    wms.CellBackColor = WMS_HEAD_10
    sh.getCellRangeByPosition(0, 0, n - GOODS_WMS_COLS - 1, 0).CellBackColor = INPUT_HEAD_10
    sh.getCellRangeByPosition(n - GOODS_WMS_COLS, 1, n - 1, 5000).CharColor = MUTED
    sh.getCellRangeByPosition(n - 1, 1, n - 1, LAST_ROW).NumberFormat = number_format(doc, "DD.MM.YYYY")
    hint = sh.getCellByPosition(n + 2, 0)
    hint.setString(GOODS_HINT)
    hint.CharColor = MUTED
    hint.IsTextWrapped = True
    sh.getColumns().getByIndex(n + 2).Width = 11000


def build_supply(doc, sh):
    """«Снабжение» (WmsSupply): the filter by supplier, «Ожидаем» (the open positions by supplier) and «Нужно заказать» (the goods
    below their minimum); a double click — the row of «Заказы»"""
    view_styles(doc)
    widths = [400, 5600, 7400, 2800, 2000, 2000, 2200, 1500, 2700, 2600, 2700, 400]
    for c, w in enumerate(widths):
        sh.getColumns().getByIndex(c).Width = w
    for c in (12, 13, 14):                         # WmsSupply.SP_SRC_SHEET, SP_SRC_ROW, SP_MARK
        sh.getColumns().getByIndex(c).IsVisible = False
    for r, h in {0: 1100, 1: 950, 2: 250, 3: 900, 4: 600, 5: 250, 6: 800, 7: 520, 8: 520, 9: 560}.items():
        sh.getRows().getByIndex(r).Height = h
    merge(sh, 1, 0, 10, 0)
    text_cell(sh, 1, 0, "Снабжение: ожидаем и нужно заказать", size=20, bold=True, color=INK, valign=2)
    text_cell(sh, 1, 3, "Поставщик", size=12, bold=True, color=INK, valign=2, halign=3)
    inp = merge(sh, 2, 3, 3, 3)
    cell = sh.getCellByPosition(2, 3)
    cell.CharHeight = 13
    cell.VertJustify = 2
    cell.NumberFormat = number_format(doc, "@", ("ru", "RU"))
    inp.CellBackColor = 0xFFF8E1
    box(inp, 0xE0B84C, 26)
    list_validation(sh.getCellByPosition(2, 3), f"$'{SUPPLIERS}'.$A$2:$A$1048576")
    merge(sh, 1, 4, 10, 4)
    text_cell(sh, 1, 4, "Пусто — все поставщики. «Ожидаем» — что заказано и ещё не пришло; «Нужно заказать» — товары «Номенклатуры» ниже "
                        "минимума. Двойной щелчок по строке — строка «Заказов».", size=9, color=MUTED, wrap=True, valign=1)
    merge(sh, 1, 6, 10, 6)
    text_cell(sh, 1, 6, "Нажмите «Ожидаем» или «Нужно заказать»", size=15, bold=True, color=INK, valign=2)
    for r in (7, 8):
        merge(sh, 1, r, 10, r)
        text_cell(sh, 1, r, "", size=10, color=TEXT, valign=2)
    hdr = sh.getCellRangeByPosition(1, 9, 10, 9)
    hdr.CharWeight = 150
    hdr.CharHeight = 9
    hdr.CharColor = MUTED
    hdr.CellBackColor = 0xE7ECF2
    hdr.VertJustify = 2
    hdr.IsTextWrapped = True
    body = sh.getCellRangeByPosition(1, 10, 10, 3100)
    body.CharHeight = 9
    body.CharColor = TEXT
    cell_style(doc, "SP_Group", CharWeight=150, CellBackColor=0xEEF3F8, CharColor=INK)
    cell_style(doc, "SP_Late", CharColor=0x9C0006)
    add_cond(sh, sh.getCellRangeByPosition(1, 10, 10, 3100), '$O11="G"', "SP_Group", 1, 10)
    add_cond(sh, sh.getCellRangeByPosition(1, 10, 10, 3100), '$O11="L"', "SP_Late", 1, 10)
    view_nav(doc, sh, 1, "spNav", "", total=30500)
    y = sh.getCellByPosition(0, 3).Position.Y
    x = sh.getCellByPosition(4, 3).Position.X
    add_button(doc, sh, "btnWait", "Ожидаем", "WmsSupply.BtnWait", x + 80, y + 90, 2500, 720, bold=True, back=0xC6E0B4)
    add_button(doc, sh, "btnNeed", "Нужно заказать", "WmsSupply.BtnNeed", x + 2660, y + 90, 3400, 720, bold=True, back=0xC6E0B4)
    add_button(doc, sh, "btnSupGo", "Перейти к строке", "WmsSupply.BtnSupplyGo", x + 6140, y + 90, 3600, 720, back=0xDDEBF7)
    add_button(doc, sh, "btnSupDir", "Поставщики", "WmsSupply.NavSuppliers", x + 9820, y + 90, 2900, 720, back=0xDDEBF7)
    add_button(doc, sh, "btnGoodsDir", "Номенклатура", "WmsSupply.NavGoods", x + 12800, y + 90, 3000, 720, back=0xDDEBF7)
    sh.getCellRangeByPosition(0, 0, 15, 3100).CellProtection = protection(True)
    sh.getCellByPosition(2, 3).CellProtection = protection(False)
    # protected with the other pages at the end of the build (a protected sheet stops the cell styles created later)


def build_print(doc, sh):
    """«Печать» (WmsPrint): a page of the documents — A4 portrait, one page wide; the buttons right of the page are not printed"""
    ps = doc.StyleFamilies.getByName("PageStyles")
    if not ps.hasByName("WMS_Print"):
        ps.insertByName("WMS_Print", doc.createInstance("com.sun.star.style.PageStyle"))
    st = ps.getByName("WMS_Print")
    st.IsLandscape = False
    st.Width, st.Height = 21000, 29700
    st.LeftMargin = st.RightMargin = 1500
    st.TopMargin = st.BottomMargin = 1200
    st.HeaderIsOn = False
    st.FooterIsOn = False
    st.ScaleToPagesX = 1
    st.ScaleToPagesY = 0
    st.PrintGrid = False
    sh.PageStyle = "WMS_Print"
    for c, w in enumerate(PRINT_WIDTHS):
        sh.getColumns().getByIndex(c).Width = w
    sh.getColumns().getByIndex(len(PRINT_WIDTHS)).Width = 600
    body = sh.getCellRangeByPosition(0, 0, len(PRINT_WIDTHS) - 1, 2000)
    body.CharHeight = 10
    body.CharColor = TEXT
    text_cell(sh, 0, 0, "Здесь появляется документ для печати: акт, накладная, опись, карточка ЕИ, отчёт.", size=13, bold=True, color=INK)
    text_cell(sh, 0, 2, "Его делают кнопки на рабочих листах («Акт расхождения», «Акт списания», «Накладная», «Опись», «Печать карточки», "
                        "«Отчёт»). «Печать…» — на принтер, «Сохранить PDF» — файл в папке «WMS_Документы» рядом с книгой.", size=10, color=MUTED)
    x = sum(PRINT_WIDTHS) + 900
    add_button(doc, sh, "btnPrint", "Печать…", "WmsPrint.BtnPrint", x, 200, 3200, 760, bold=True, back=0xC6E0B4)
    add_button(doc, sh, "btnPdf", "Сохранить PDF", "WmsPrint.BtnPdf", x, 1060, 3200, 760, back=0xDDEBF7)
    add_button(doc, sh, "btnPrintBack", "Назад", "WmsPrint.BtnBack", x, 1920, 3200, 760, back=0xDDEBF7)
    sh.getCellRangeByPosition(0, 0, 12, 2000).CellProtection = protection(True)
    # protected with the other pages at the end of the build


def build_reports(doc, sh):
    """«Отчёты» (WmsLedger): a menu — a button and what it makes; every report goes to «Печать» (print, PDF) or into a file"""
    for c, w in enumerate([400, 6600, 17000, 400]):
        sh.getColumns().getByIndex(c).Width = w
    sh.getRows().getByIndex(0).Height = 1100
    sh.getRows().getByIndex(1).Height = 900
    merge(sh, 1, 0, 2, 0)
    text_cell(sh, 1, 0, "Отчёты и учёт", size=20, bold=True, color=INK, valign=2)
    merge(sh, 1, 1, 2, 1)
    text_cell(sh, 1, 1, "Отчёт появляется на странице «Печать» — оттуда «Печать…» или «Сохранить PDF»; файлы CSV и PDF — в папке «WMS_Документы» "
                        "рядом с книгой. Отчёты ничего не проводят; «Закрыть месяц» и «Открыть месяц» — операции журнала.", size=9, color=MUTED,
              wrap=True, valign=1)
    r = 3
    for caption, items in REPORT_GROUPS:
        sh.getRows().getByIndex(r).Height = 700
        text_cell(sh, 1, r, caption, size=12, bold=True, color=INK, valign=3)
        r += 1
        for name, label, macro, desc in items:
            sh.getRows().getByIndex(r).Height = 900
            text_cell(sh, 2, r, desc, size=10, color=TEXT, wrap=True, valign=2)
            y = sh.getCellByPosition(1, r).Position.Y
            x = sh.getCellByPosition(1, r).Position.X
            add_button(doc, sh, name, label, macro, x, y + 80, 6400, 740, bold=name in ("btnStockOnDate", "btnTurnover", "btnManagerReport"),
                       back=0xF4CCCC if name == "btnOpenMonth" else (0xC6E0B4 if name == "btnCloseMonth" else 0xDDEBF7))
            r += 1
        r += 1
    add_button(doc, sh, "btnRepMain", "Главная", "WmsUi.NavMain", sh.getCellByPosition(2, 0).Position.X + 13200, 160, 3700, 760, back=0xDDEBF7)
    sh.getCellRangeByPosition(0, 0, 12, 200).CellProtection = protection(True)
    # protected with the other pages at the end of the build


def build_service_page(doc, sh):
    """«Обслуживание» (WmsSafe): the state of the safety of the data (filled when the page is opened) and the buttons"""
    for c, w in enumerate([400, 6600, 17000, 400]):
        sh.getColumns().getByIndex(c).Width = w
    sh.getRows().getByIndex(0).Height = 1100
    sh.getRows().getByIndex(1).Height = 700
    merge(sh, 1, 0, 2, 0)
    text_cell(sh, 1, 0, "Обслуживание", size=20, bold=True, color=INK, valign=2)
    merge(sh, 1, 1, 2, 1)
    text_cell(sh, 1, 1, "Сохранность данных и здоровье книги. Состояние обновляется при каждом открытии страницы; ничего здесь не проводит операций, "
                        "кроме возврата к копии.", size=9, color=MUTED, wrap=True, valign=1)
    # Т6: one light for the whole WMS (WmsSafe.Health): ✓ green, ! yellow (soon), ✗ red (now)
    sh.getRows().getByIndex(SERVICE_STATE_ROW - 1).Height = 1000
    text_cell(sh, 1, SERVICE_STATE_ROW - 1, "Здоровье WMS", size=12, bold=True, color=INK, valign=2)
    text_cell(sh, 2, SERVICE_STATE_ROW - 1, "", size=11, bold=True, color=TEXT, wrap=True, valign=2)
    hc = sh.getCellRangeByPosition(2, SERVICE_STATE_ROW - 1, 2, SERVICE_STATE_ROW - 1)
    add_cond(sh, hc, 'LEFT($C' + str(SERVICE_STATE_ROW) + ';1)="✓"', "DB_Good", 2, SERVICE_STATE_ROW - 1)
    add_cond(sh, hc, 'LEFT($C' + str(SERVICE_STATE_ROW) + ';1)="✗"', "DB_Alert", 2, SERVICE_STATE_ROW - 1)
    add_cond(sh, hc, 'LEFT($C' + str(SERVICE_STATE_ROW) + ';1)="!"', "DB_Warn", 2, SERVICE_STATE_ROW - 1)
    r = SERVICE_STATE_ROW
    for label in SERVICE_STATE:
        sh.getRows().getByIndex(r).Height = 560
        text_cell(sh, 1, r, label, size=10, bold=True, color=INK, valign=2)
        text_cell(sh, 2, r, "", size=10, color=TEXT, wrap=False, valign=2)
        r += 1
    add_cond(sh, sh.getCellRangeByPosition(2, SERVICE_STATE_ROW, 2, r - 1), 'LEFT($C' + str(SERVICE_STATE_ROW + 1) + ';1)="✓"', "DB_Good", 2,
             SERVICE_STATE_ROW)
    add_cond(sh, sh.getCellRangeByPosition(2, SERVICE_STATE_ROW, 2, r - 1), 'LEFT($C' + str(SERVICE_STATE_ROW + 1) + ';1)="✗"', "DB_Alert", 2,
             SERVICE_STATE_ROW)
    r += 1
    for caption, items in SERVICE_GROUPS:
        sh.getRows().getByIndex(r).Height = 700
        text_cell(sh, 1, r, caption, size=12, bold=True, color=INK, valign=3)
        r += 1
        for name, label, macro, desc in items:
            sh.getRows().getByIndex(r).Height = 900
            text_cell(sh, 2, r, desc, size=10, color=TEXT, wrap=True, valign=2)
            y = sh.getCellByPosition(1, r).Position.Y
            x = sh.getCellByPosition(1, r).Position.X
            add_button(doc, sh, name, label, macro, x, y + 80, 6400, 740, bold=name in ("btnFlashCopy", "btnDailyCheck"),
                       back=0xC6E0B4 if name in ("btnFlashCopy", "btnDailyCheck") else 0xDDEBF7)
            r += 1
        r += 1
    add_button(doc, sh, "btnSvMain", "Главная", "WmsUi.NavMain", sh.getCellByPosition(2, 0).Position.X + 13200, 160, 3700, 760, back=0xDDEBF7)
    sh.getCellRangeByPosition(0, 0, 12, 200).CellProtection = protection(True)


def service_10(doc, sheets):
    """1.0 (этап J): the page «Обслуживание» for a new book and an upgraded one; its state is refreshed when it is opened"""
    if not sheets.hasByName(SERVICE):
        names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
        sheets.insertNewByName(SERVICE, names.index(PRINT) if PRINT in names else names.index(FIND) + 1)
        build_service_page(doc, sheets.getByName(SERVICE))
    bind_event(sheets.getByName(SERVICE).Events, "OnFocus", "WmsSafe.OnServiceFocus")


def service_110(doc, sheets):
    """1.1: the button «Справочники из копии…» on «Обслуживание» of a book that was made before it (the page is not built again): after the last button of the page"""
    if not sheets.hasByName(SERVICE):
        return
    sh = sheets.getByName(SERVICE)
    if has_button(sh, "btnDirsBook"):
        return
    desc = [x for x in SERVICE_GROUPS[0][1] if x[0] == "btnDirsBook"][0]
    # the first free row after the last button of the page: a button's shape position decides
    form = sh.getDrawPage()
    maxy = 0
    for i in range(form.getCount()):
        shp = form.getByIndex(i)
        maxy = max(maxy, shp.Position.Y + shp.Size.Height)
    r = 0
    while sh.getCellByPosition(1, r).Position.Y < maxy + 100 and r < 400:
        r += 1
    sh.getRows().getByIndex(r).Height = 900
    text_cell(sh, 2, r, desc[3], size=10, color=TEXT, wrap=True, valign=2)
    y = sh.getCellByPosition(1, r).Position.Y
    x = sh.getCellByPosition(1, r).Position.X
    add_button(doc, sh, desc[0], desc[1], desc[2], x, y + 80, 6400, 740, back=0xDDEBF7)


def dash_10(doc, sheets):
    """1.0 (А6): «Все…» over the four lists of «Дашборд» and the hidden column of the references of its right lists (a double
    click on a row opens its order, its EI card or the goods on «Поиск»)"""
    sh = sheets.getByName(DASH)
    sh.getColumns().getByIndex(15).IsVisible = False
    for name, macro, c_end, r in (("btnAllWait", "WmsDash.BtnAllWait", 7, 7), ("btnAllOps", "WmsDash.BtnAllOps", 13, 7),
                                  ("btnAllIssued", "WmsDash.BtnAllIssued", 7, 22), ("btnAllStale", "WmsDash.BtnAllStale", 13, 22)):
        if not has_button(sh, name):
            x = sh.getCellByPosition(c_end, r).Position.X - 1900
            y = sh.getCellByPosition(0, r).Position.Y + 60
            add_button(doc, sh, name, "Все…", macro, x, y, 1800, 520, back=0xDDEBF7)


# 1.0.1: a component may be one EI (its number in «Артикул»); «Собрать из ЕИ…» builds from the EIs selected on «Наличие»
KITS_HINT = ("Одна строка — одна позиция комплекта: комплект, наименование, артикул — или номер ЕИ (ЕИ-427: позиция берётся только с "
             "этого ЕИ), количество на 1 комплект, единица. «Собрать комплект…» — позиции снимаются с их ЕИ, готовые комплекты "
             "становятся новым ЕИ (единица «компл»). «Собрать из ЕИ…» — выделите на «Наличии» строки нужных ЕИ и нажмите эту кнопку "
             "там: комплект соберётся именно из них, состав можно записать сюда. «Выдать комплект…» — WMS заполняет строки «Выдач» "
             "нужными ЕИ, их проводят обычным «Провести».")


def build_kits(doc, sh):
    """«Комплекты» (WmsKit): a kit — its rows (the name in A of the first row is enough, an empty A continues the kit above);
    the buttons over the header; the sheet stays open for the storekeeper"""
    header(sh, KITS_HEADERS, [5200, 6500, 3000, 3200, 1800, 6000], height=1700)
    text = number_format(doc, "@", ("ru", "RU"))
    for c in (0, 1, 2, 4, 5):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
    hint = sh.getCellByPosition(7, 0)
    hint.setString(KITS_HINT)
    hint.IsTextWrapped = True
    hint.VertJustify = 1
    hint.CharHeight = 9
    hint.CharColor = MUTED
    sh.getColumns().getByIndex(6).Width = 400
    sh.getColumns().getByIndex(7).Width = 12000
    add_button(doc, sh, "btnKitBuild", "Собрать комплект…", "WmsKit.BtnKitBuild", 150, 120, 4400, 650, bold=True, back=0xC6E0B4)
    add_button(doc, sh, "btnKitIssue", "Выдать комплект…", "WmsKit.BtnKitIssue", 4700, 120, 4400, 650, back=0xC6E0B4)


def kits_10(doc, sheets):
    """1.0 (Ф8): the directory «Комплекты» for a new book and an upgraded one"""
    if not sheets.hasByName(KITS):
        names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
        at = names.index(GOODS) + 1 if GOODS in names else (names.index(RCPT) + 1 if RCPT in names else len(names))
        sheets.insertNewByName(KITS, at)
        build_kits(doc, sheets.getByName(KITS))


def build_reserves(doc, sh):
    """«Резервы» (WmsItem, Ф7): a row — a quantity of an EI kept for a recipient or an object until a date; «Выдачи» asks before
    giving the others more than is free of it. The storekeeper's own sheet: not protected, not journaled"""
    header(sh, RESERVE_HEADERS, [3200, 2600, 7000, 2800, 7000], height=2000)
    text = number_format(doc, "@", ("ru", "RU"))
    for c in (0, 2, 4):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
    sh.getCellRangeByPosition(3, 1, 3, LAST_ROW).NumberFormat = number_format(doc, "DD.MM.YYYY")
    hint = sh.getCellByPosition(6, 0)
    hint.setString("Резерв — часть ЕИ, отложенная для получателя или объекта: ЕИ (427 или ЕИ-00000427), сколько, для кого, до какой даты "
                   "(пусто — пока не уберёте строку). Выдавая другому больше свободного, WMS спросит. Выдача тому, для кого резерв, "
                   "не спрашивает; прошедшая дата «До» резерв снимает. Выдали по резерву — уменьшите количество или удалите строку.")
    hint.IsTextWrapped = True
    hint.VertJustify = 1
    hint.CharHeight = 9
    hint.CharColor = MUTED
    sh.getColumns().getByIndex(5).Width = 400
    sh.getColumns().getByIndex(6).Width = 16000


def items_10(doc, sheets):
    """1.0 (Ф4, Ф5, Ф7): what a new book and an upgraded one both get — the hidden _PROP, the sheet «Резервы» and the column Q
    «Вернуть через, дней» of «Номенклатура»"""
    if not sheets.hasByName(PROP):
        names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
        sheets.insertNewByName(PROP, names.index("_SYS") if "_SYS" in names else sheets.getCount())
        sh = sheets.getByName(PROP)
        for c, h in enumerate(PROP_HEADERS):
            sh.getCellByPosition(c, 0).setString(h)
        text = number_format(doc, "@", ("ru", "RU"))
        for c in (0, 1, 3):
            sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
        sh.getCellRangeByPosition(2, 1, 2, LAST_ROW).NumberFormat = number_format(doc, "DD.MM.YYYY")
        sh.IsVisible = False
    if not sheets.hasByName(RESERVES):
        names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
        at = names.index(KITS) + 1 if KITS in names else (names.index(GOODS) + 1 if GOODS in names else len(names))
        sheets.insertNewByName(RESERVES, at)
        build_reserves(doc, sheets.getByName(RESERVES))
    gs = sheets.getByName(GOODS)
    h = gs.getCellByPosition(GOODS_RET_COL, 0)
    if h.getString().strip() == "":
        h.setString(GOODS_RET_HEADER)
        ref = gs.getCellByPosition(GOODS_RET_COL - 5, 0)
        h.CharWeight = 150
        h.IsTextWrapped = True
        h.VertJustify = ref.VertJustify
        h.HoriJustify = ref.HoriJustify
        h.CharHeight = ref.CharHeight
        h.CellBackColor = INPUT_HEAD_10
        gs.getColumns().getByIndex(GOODS_RET_COL).Width = 3000
        gs.getCellRangeByPosition(GOODS_RET_COL, 1, GOODS_RET_COL, LAST_ROW).CharColor = gs.getCellByPosition(GOODS_RET_COL - 5, 1).CharColor


def goods_102(doc, sheets):
    """1.0.2: the column R «В упаковке, шт» of «Номенклатура» (WmsItem GD_PACKQ) for a new book and an upgraded one. The hint of the
    sheet stood in R1 up to 1.0.1: it moves to S1 with its look; a cell R1 holding something else is not touched"""
    gs = sheets.getByName(GOODS)
    h = gs.getCellByPosition(GOODS_PACKQ_COL, 0)
    old = h.getString().strip()
    if old == GOODS_PACKQ_HEADER:
        return
    if old != "" and not old.startswith("Товар — строка"):
        return
    if old != "":
        t = gs.getCellByPosition(GOODS_PACKQ_COL + 1, 0)
        t.setString(GOODS_HINT)
        for prop in ("CharColor", "IsTextWrapped", "VertJustify", "HoriJustify", "CharHeight", "CharWeight"):
            setattr(t, prop, getattr(h, prop))
        gs.getColumns().getByIndex(GOODS_PACKQ_COL + 1).Width = 11000
    ref = gs.getCellByPosition(GOODS_RET_COL, 0)
    h.setString(GOODS_PACKQ_HEADER)
    for prop in ("CharWeight", "CellBackColor", "IsTextWrapped", "VertJustify", "HoriJustify", "CharHeight", "CharColor"):
        setattr(h, prop, getattr(ref, prop))
    gs.getColumns().getByIndex(GOODS_PACKQ_COL).Width = 3600
    gs.getCellRangeByPosition(GOODS_PACKQ_COL, 1, GOODS_PACKQ_COL, LAST_ROW).CharColor = gs.getCellByPosition(GOODS_RET_COL, 1).CharColor


# 1.0.2: «Акты» — the register of the acts (WmsDocs SH_ACTS, AK_*): the storekeeper's own sheet, not protected, not journaled; «Отчёт дня» —
# the lists of the daily report (WmsDocs SH_DAYREP, DR_*)
ACTS = "Акты"
ACTS_HEADERS = ["№ акта", "Дата", "Вид", "Получатель", "Позиций", "Выдачи (№)", "Заметка"]
ACTS_SETTINGS = [("Настройки акта", "Значение"), ("Город в шапке", "г Москва"), ("Получатель «офис» (если несколько — через запятую)", "Офис"),
                 ("Подпись кладовщика под актом (необязательно)", "")]
DAYREP = "Отчёт дня"
DAYREP_TASKS = ["Выдача инструмента/оборудования", "Взаимодействие с руководителем склада", "Оформление пропусков", "Приём поставок на склад",
                "Оформление документов и товара к отправке в офис"]
DAYREP_PROG = ["Работа над учётом склада"]
DAYREP_HINT = ("«Подготовить отчёт» открывает отчёт за выбранный день отдельным документом Writer: что сделано по данным WMS (приёмка, выдачи — "
               "и оцифровка выдач за прошлые рабочие дни, акты, исправления), задачи на следующий рабочий день, в процессе, контроль данных. "
               "Графики 5/2: выходные пропускаются. Здесь — то, чего WMS знать не может: впишите своё по спискам ниже. «Сделано вручную» "
               "очищается кнопкой в начале нового дня; задачи и «в процессе» остаются, пока вы их не измените.")
DAYREP_ROW0 = 5                      # WmsDocs DR_ROW0


def build_acts(doc, sh):
    """«Акты» (WmsDocs): the register of the acts — № акта, дата, вид, получатель, позиций, номера выдач; the settings of the act on the
    right (the city of the head, the recipient that is the office, the signature of the storekeeper); «Повторить акт» prints the act of a
    row of the register again"""
    header(sh, ACTS_HEADERS, [2200, 2600, 6600, 7000, 2000, 6000, 8000], height=1700)
    text = number_format(doc, "@", ("ru", "RU"))
    for c in (2, 3, 5, 6):
        sh.getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
    sh.getCellRangeByPosition(1, 1, 1, LAST_ROW).NumberFormat = number_format(doc, "DD.MM.YYYY")
    sh.getColumns().getByIndex(7).Width = 600
    sh.getColumns().getByIndex(8).Width = 9500
    sh.getColumns().getByIndex(9).Width = 5000
    for k, (a, v) in enumerate(ACTS_SETTINGS):
        sh.getCellByPosition(8, k).setString(a)
        sh.getCellByPosition(9, k).setString(v)
        sh.getCellByPosition(9, k).NumberFormat = text
        sh.getCellByPosition(9, k).CellBackColor = INPUT_HEAD_10 if k else 0xE7E6E6
    sh.getCellRangeByPosition(8, 0, 9, 0).CharWeight = 150
    sh.getCellRangeByPosition(8, 0, 9, 0).VertJustify = 3
    sh.getCellRangeByPosition(8, 1, 8, 3).CharColor = MUTED
    sh.getCellRangeByPosition(8, 1, 8, 3).IsTextWrapped = True
    add_button(doc, sh, "btnActRepeat", "Повторить акт", "WmsDocs.BtnActRepeat", 150, 120, 4200, 650, back=0xC6E0B4)


def build_dayrep(doc, sh):
    """«Отчёт дня» (WmsDocs): the date and the signature of the report, three lists written by hand — done by hand today (A), tasks for the
    next working day (C), in progress (E) — and the button that builds the report"""
    sh.getColumns().getByIndex(0).Width = 10500
    sh.getColumns().getByIndex(1).Width = 5500
    sh.getColumns().getByIndex(2).Width = 10500
    sh.getColumns().getByIndex(3).Width = 600
    sh.getColumns().getByIndex(4).Width = 10500
    sh.getRows().getByIndex(0).Height = 1500
    text = number_format(doc, "@", ("ru", "RU"))
    sh.getCellRangeByPosition(0, 0, 4, 30).NumberFormat = text
    sh.getCellByPosition(0, 1).setString("Дата отчёта (пусто — сегодня), дд.мм.гггг")
    sh.getCellByPosition(0, 2).setString("Подпись в отчёте")
    sh.getCellByPosition(1, 2).setString("Егор, кладовщик.")
    for r in (1, 2):
        sh.getCellByPosition(1, r).CellBackColor = INPUT_HEAD_10
        sh.getCellByPosition(0, r).CharColor = MUTED
    hint = sh.getCellByPosition(2, 1)
    hint.setString(DAYREP_HINT)
    hint.IsTextWrapped = True
    hint.VertJustify = 1
    hint.CharHeight = 9
    hint.CharColor = MUTED
    merge(sh, 2, 1, 4, 3)
    for c, t in ((0, "Сегодня сделано вручную (чего нет в WMS)"), (2, "Задачи на следующий рабочий день"), (4, "В процессе")):
        cell = sh.getCellByPosition(c, DAYREP_ROW0 - 1)
        cell.setString(t)
        cell.CharWeight = 150
        cell.CellBackColor = 0xE7E6E6
    for c in (0, 2, 4):
        rng = sh.getCellRangeByPosition(c, DAYREP_ROW0, c, DAYREP_ROW0 + 19)
        rng.CellBackColor = INPUT_HEAD_10
        rng.IsTextWrapped = True
    for k, t in enumerate(DAYREP_TASKS):
        sh.getCellByPosition(2, DAYREP_ROW0 + k).setString(t)
    for k, t in enumerate(DAYREP_PROG):
        sh.getCellByPosition(4, DAYREP_ROW0 + k).setString(t)
    add_button(doc, sh, "btnDayReport", "Подготовить отчёт", "WmsDocs.BtnDayReport", 150, 120, 5400, 760, bold=True, back=0xC6E0B4)
    add_button(doc, sh, "btnDrClear", "Очистить «Сделано вручную»", "WmsDocs.BtnDrClear", 5700, 120, 5700, 760, back=0xDDEBF7)


def docs_102(doc, sheets):
    """1.0.2: the sheets «Акты» and «Отчёт дня» for a new book and an upgraded one (after «Печать»); nothing that exists is touched"""
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    at = names.index(PRINT) + 1 if PRINT in names else len(names)
    if not sheets.hasByName(ACTS):
        sheets.insertNewByName(ACTS, at)
        build_acts(doc, sheets.getByName(ACTS))
        at += 1
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(DAYREP):
        sheets.insertNewByName(DAYREP, names.index(ACTS) + 1 if ACTS in names else len(names))
        build_dayrep(doc, sheets.getByName(DAYREP))


# 1.0.2: «Категории» — the hints of the categories by the name (WmsCat SH_CATS): a category and the words that tell it in a name; «Метизы» —
# the directory of the fasteners made from «Наличие» by the names (WmsCat SH_FAST, FS_*): the storekeeper's own sheets, not protected, not journaled
CATS = "Категории"
CATS_HEADERS = ["Категория", "Слова в названии (через запятую; части слова достаточно: «болт» найдёт «болты»; «=болт» — только целое слово)",
                "Не считать (слова, при которых категория не подходит)", "Позиций", "ЕИ с остатком"]
CATS_WIDTHS = [6000, 20000, 9000, 2400, 2800]
CATS_SEED = [
    ("Инструмент", "дрел, перфоратор, шуруповерт, отверт, молоток, ключ, болгарк, болторез, пассатиж, плоскогуб, кусачк, рулетк, уровен, дальномер, лобзик, паяльник, пила"),
    ("Метизы", "болт, гайк, шайб, винт, саморез, шуруп, шпильк, дюбел, анкер, заклепк, гвозд, хомут, шплинт"),
    ("Кабель и провод", "кабел, провод, шнур"),
    ("Электрика", "лампа, лампоч, розетк, выключател, автомат, удлинител, светильник, клемм, патрон, пускател"),
    ("СИЗ", "перчатк, каск, очки, респиратор, наушник, жилет, комбинезон, ботинк, сапог, щиток, маска"),
    ("Металл", "труб, уголок, швеллер, арматур, двутавр, полоса, пруток, профиль, лист стал, лист оцинк"),
    ("Химия и краски", "краск, эмаль, растворител, клей, герметик, смазк, масло, очистител, грунт, ацетон"),
    ("Расходники", "диск отрез, диск шлиф, электрод, сверло, лента, скотч, изолент, ветош, салфетк, пленк"),
    ("Детали", "шестерн, подшипник, втулк, кронштейн, фланец, муфт, звездочк, ремень, редуктор, датчик"),
    ("Офис", "бумаг, картридж, ручка, степлер, папк, скрепк"),
]
FAST = "Метизы"
FAST_HEADERS = ["Тип", "Размер", "Ø", "Длина", "Стандарт", "Аналоги стандарта", "Класс", "Покрытие / материал", "Головка / шлиц", "Наименование (как в учёте)", "Артикул", "Остаток",
                "Ед.", "Места", "ЕИ с остатком", "Заметка"]
FAST_WIDTHS = [2600, 2600, 1300, 1600, 2800, 4200, 1800, 4000, 3600, 9000, 3200, 2200, 1500, 5000, 2200, 7000]
FAST_HINT = ("Справочник метизов из «Наличия»: болты, гайки, шайбы, винты, саморезы … WMS узнаёт их по названию (и по категории «Метизы» / «Крепёж», в том числе из «Номенклатуры») и "
             "разбирает его: размер (М8×30), диаметр и длина, стандарт (DIN 933) и его аналоги по листу «Стандарты» (ISO 4017, ГОСТ 7805), класс прочности (8.8, А2), покрытие, головка и шлиц. "
             "Одна строка — одно наименование с артикулом: остаток всех его ЕИ и места. «Найти» понимает слова, как страница «Поиск» («болт м8 30», «din 933», «-нерж»). Строка под ним — "
             "фильтры столбцов: текст — слова, которые должны быть («оцинк»); числа — «8», «8 10 12» (любое из), «8-12», «>20», «<=40»; нажмите Enter. «Скрыть нулевые» — «да». В "
             "«Заметке» — возможные дубли (то же самое под другим названием; одно название под разными артикулами). Список обновляется при открытии листа и кнопкой «Обновить». "
             "Только просмотр: ничего не проводится.")
FAST_HINT_COL = 18
# the sheets 1.0.2 adds to a book of 1.0.0 / 1.0.1
SHEETS_102 = [ACTS, DAYREP, CATS, FAST]
# the sheets 1.1 adds to a book of 1.0.x (filled in as the release grows)
SHEETS_110 = ["Наличие по названию", "Остаток деталей", "Операции дня", "Проверка названий", "Стандарты", "Синонимы", "Без категории", "Настройки"]
FAST_FIND_ROW = 1
FAST_FILTER_ROW = 2
FAST_HEAD = 3
FAST_KEY_COL = 16


CATS_HINT = ("Подсказки категорий. Вписали наименование товара на «Заказах» — пустая «Категория» получает подсказку: сначала категорию товара из «Номенклатуры» (если она там есть), "
             "затем по словам отсюда: выигрывает категория, у которой в названии нашлось больше слов (целое слово весит больше, чем его начало); при равенстве — та, что выше. "
             "Слово «=болт» — только целое слово; слово «-болторез» (или слова в столбце «Не считать») — категория не подходит, если оно есть в названии. Можно написать свою категорию — "
             "подсказка её не трогает. Свои строки и слова — добавляйте и меняйте: это ваш справочник, учёт он не меняет. «Позиций» и «ЕИ с остатком» считаются при открытии листа: "
             "категория берётся из реестра, из «Номенклатуры» или по словам. Список слева — выпадающий список «Категории» на «Заказах» и «Ином приходе». Названия без категории — на листе «Без категории».")
CATS_HINT_COL = 6


def cats_style_new_columns(doc, sh):
    """the columns of 1.1 of «Категории»: the words to leave out (an input), the two counted columns (grey: not an input)"""
    text = number_format(doc, "@", ("ru", "RU"))
    sh.getCellRangeByPosition(2, 0, 2, 300).NumberFormat = text
    sh.getCellRangeByPosition(2, 1, 2, 300).CellBackColor = INPUT_HEAD_10
    sh.getCellRangeByPosition(2, 1, 2, 300).IsTextWrapped = True
    sh.getCellRangeByPosition(3, 1, 4, 300).CellBackColor = 0xF2F2F2
    sh.getCellRangeByPosition(3, 1, 4, 300).HoriJustify = 3
    sh.getCellRangeByPosition(3, 1, 4, 300).CharColor = MUTED
    for c, w in enumerate(CATS_WIDTHS):
        sh.getColumns().getByIndex(c).Width = w
    sh.getColumns().getByIndex(5).Width = 600
    sh.getColumns().getByIndex(CATS_HINT_COL).Width = 12000


def build_cats(doc, sh):
    header(sh, CATS_HEADERS, CATS_WIDTHS, height=900)
    text = number_format(doc, "@", ("ru", "RU"))
    sh.getCellRangeByPosition(0, 0, 1, 300).NumberFormat = text
    sh.getCellRangeByPosition(0, 1, 1, 300).CellBackColor = INPUT_HEAD_10
    for r, (a, w) in enumerate(CATS_SEED, start=1):
        sh.getCellByPosition(0, r).setString(a)
        sh.getCellByPosition(1, r).setString(w)
    sh.getCellRangeByPosition(1, 1, 1, 300).IsTextWrapped = True
    cats_style_new_columns(doc, sh)
    hint = sh.getCellByPosition(CATS_HINT_COL, 0)
    hint.setString(CATS_HINT)
    hint.IsTextWrapped = True
    hint.VertJustify = 1
    hint.CharHeight = 9
    hint.CharColor = MUTED
    bind_event(sh.Events, "OnChange", "WmsCat.OnCatsChange")
    bind_event(sh.Events, "OnFocus", "WmsCat.OnCatsFocus")


def cats_110(doc, sheets):
    """1.1: «Категории» of a book of 1.0.2 gets the columns «Не считать», «Позиций», «ЕИ с остатком» (C, D, E) and the hint moves from D1 to G1; the categories and the words (A, B) are
    not touched. A sheet in which the storekeeper has written something into C…G of his own, or a sheet that has the columns already, is left as it is"""
    if not sheets.hasByName(CATS):
        return
    sh = sheets.getByName(CATS)
    if sh.getCellByPosition(2, 0).getString().strip() == CATS_HEADERS[2]:
        return
    d1 = sh.getCellByPosition(3, 0).getString().strip()
    busy = any(str(x) != "" for row in sh.getCellRangeByPosition(2, 1, CATS_HINT_COL, 300).getDataArray() for x in row)
    if busy or any(sh.getCellByPosition(c, 0).getString().strip() != "" for c in (2, 4, 5, CATS_HINT_COL)) or (d1 != "" and not d1.startswith("Подсказки категорий")):
        return
    ref = sh.getCellByPosition(0, 0)
    for c in (2, 3, 4):
        h = sh.getCellByPosition(c, 0)
        h.setString(CATS_HEADERS[c])
        for prop in ("CharWeight", "CellBackColor", "CharColor", "HoriJustify", "CharHeight"):
            setattr(h, prop, getattr(ref, prop))
        h.IsTextWrapped = True
        h.VertJustify = 2
    cats_style_new_columns(doc, sh)
    hint = sh.getCellByPosition(CATS_HINT_COL, 0)
    hint.setString(CATS_HINT)
    hint.IsTextWrapped = True
    hint.VertJustify = 1
    hint.CharHeight = 9
    hint.CharColor = MUTED
    bind_event(sh.Events, "OnChange", "WmsCat.OnCatsChange")
    bind_event(sh.Events, "OnFocus", "WmsCat.OnCatsFocus")


def build_fast(doc, sh):
    """«Метизы» (1.1 layout): the buttons, the search (row 2) and its «Скрыть нулевые», the filters of the columns (row 3), the head of the table (row 4), the list; the key of the sort
    is hidden, the hint is on the right"""
    header(sh, FAST_HEADERS, FAST_WIDTHS, height=1500)
    n = len(FAST_HEADERS)
    sh.getRows().getByIndex(0).Height = 1500
    for c in range(n):
        sh.getCellByPosition(c, FAST_HEAD).setString(FAST_HEADERS[c])
        sh.getCellByPosition(c, 0).setString("")
    hd = sh.getCellRangeByPosition(0, FAST_HEAD, n - 1, FAST_HEAD)
    hd.CharWeight = 150
    hd.IsTextWrapped = True
    hd.VertJustify = 2
    hd.CellBackColor = 0xE7E6E6
    sh.getCellRangeByPosition(0, 0, n - 1, 0).CellBackColor = 0xFFFFFF
    text = number_format(doc, "@", ("ru", "RU"))
    for c in (0, 1, 4, 5, 6, 7, 8, 9, 10, 12, 13, 15):
        sh.getCellRangeByPosition(c, FAST_HEAD + 1, c, 6000).NumberFormat = text
    # the filters of the columns are texts too («>20», «8-12»)
    sh.getCellRangeByPosition(0, FAST_FILTER_ROW, n - 1, FAST_FILTER_ROW).NumberFormat = text
    sh.getCellRangeByPosition(0, FAST_FILTER_ROW, n - 1, FAST_FILTER_ROW).CellBackColor = INPUT_HEAD_10
    sh.getCellByPosition(0, FAST_FIND_ROW).setString("Найти:")
    sh.getCellByPosition(0, FAST_FIND_ROW).CharWeight = 150
    sh.getCellByPosition(1, FAST_FIND_ROW).CellBackColor = INPUT_HEAD_10
    sh.getCellByPosition(1, FAST_FIND_ROW).NumberFormat = text
    merge(sh, 1, FAST_FIND_ROW, 3, FAST_FIND_ROW)
    sh.getCellByPosition(4, FAST_FIND_ROW).setString("Скрыть нулевые:")
    sh.getCellByPosition(4, FAST_FIND_ROW).CharWeight = 150
    sh.getCellByPosition(4, FAST_FIND_ROW).HoriJustify = 3
    sh.getCellByPosition(5, FAST_FIND_ROW).setString("нет")
    sh.getCellByPosition(5, FAST_FIND_ROW).CellBackColor = INPUT_HEAD_10
    sh.getCellByPosition(5, FAST_FIND_ROW).NumberFormat = text
    sh.getCellByPosition(0, 2).CharColor = MUTED
    hint = sh.getCellByPosition(FAST_HINT_COL, 0)
    hint.setString(FAST_HINT)
    hint.IsTextWrapped = True
    hint.VertJustify = 1
    hint.CharHeight = 9
    hint.CharColor = MUTED
    sh.getColumns().getByIndex(FAST_KEY_COL).IsVisible = False
    sh.getColumns().getByIndex(FAST_KEY_COL + 1).Width = 600
    sh.getColumns().getByIndex(FAST_HINT_COL).Width = 12000
    sh.getCellRangeByPosition(11, FAST_HEAD + 1, 11, 6000).HoriJustify = 3
    sh.getCellRangeByPosition(2, FAST_HEAD + 1, 3, 6000).HoriJustify = 3
    if not has_button(sh, "btnFastRefresh"):
        add_button(doc, sh, "btnFastRefresh", "Обновить", "WmsCat.BtnFastRefresh", 150, 120, 3000, 650, bold=True, back=0xC6E0B4)
    bind_event(sh.Events, "OnFocus", "WmsCat.OnFastFocus")
    bind_event(sh.Events, "OnChange", "WmsCat.OnFastChange")


def fast_110(doc, sheets):
    """1.1: «Метизы» of a book of 1.0.2 is made again in the layout of 1.1 (new columns, filters). The sheet is a view made from «Наличия» — there is no data of the storekeeper in it:
    it is cleared (cells and formats) and built anew; its button stays; the list appears at the first opening"""
    if not sheets.hasByName(FAST):
        return
    sh = sheets.getByName(FAST)
    if sh.getCellByPosition(2, FAST_HEAD).getString().strip() == FAST_HEADERS[2]:
        return
    rng = sh.getCellRangeByPosition(0, 0, 40, 6200)
    rng.merge(False)
    rng.clearContents(1 + 2 + 4 + 8 + 16 + 32 + 64 + 256 + 512)
    for c in range(0, 30):
        sh.getColumns().getByIndex(c).IsVisible = True
    build_fast(doc, sh)


# 1.1: «Настройки» (WmsSettings SH_SET) — the settings of the storekeeper: a table «что настроить — значение — пояснение», the code of a setting in the hidden column D
SETTINGS = "Настройки"
SETTINGS_HEADERS = ["Что настроить", "Значение", "Пояснение", "(код)"]
SETTINGS_WIDTHS = [11000, 3200, 24000, 2400]
SETTINGS_ROWS = [
    ("Большая операция: спрашивать дважды от количества", 1000,
     "Выдача или списание от этого количества (в единицах учёта: штуки, метры, килограммы) спрашивается дважды — лишний ноль при вводе обходится дорого. 0 — не спрашивать.", "BIG_QTY"),
]


def build_settings(doc, sh):
    header(sh, SETTINGS_HEADERS, SETTINGS_WIDTHS, height=900)
    text = number_format(doc, "@", ("ru", "RU"))
    sh.getCellRangeByPosition(0, 0, 3, 60).NumberFormat = text
    sh.getCellRangeByPosition(1, 1, 1, 40).CellBackColor = INPUT_HEAD_10
    sh.getCellRangeByPosition(1, 1, 1, 40).HoriJustify = 3
    sh.getCellRangeByPosition(0, 1, 2, 40).IsTextWrapped = True
    for r, (a, v, e, code) in enumerate(SETTINGS_ROWS, start=1):
        sh.getCellByPosition(0, r).setString(a)
        sh.getCellByPosition(1, r).setValue(v)
        sh.getCellByPosition(1, r).NumberFormat = 0
        sh.getCellByPosition(2, r).setString(e)
        sh.getCellByPosition(2, r).CharColor = MUTED
        sh.getCellByPosition(3, r).setString(code)
    sh.getColumns().getByIndex(3).IsVisible = False


# 1.1: «Без категории» (WmsCat SH_NOCAT) — the names of the registry with no category in the registry nor in «Номенклатура», with the hint; the column «Категория» is filled with the hint and
# goes to «Номенклатура» by the button (an empty category of an existing goods only)
NOCAT = "Без категории"
NOCAT_HEADERS = ["Наименование", "Артикул", "ЕИ с остатком", "Остаток", "Ед.", "Подсказка по словам", "Категория (оставьте, измените или очистите)", "В «Номенклатуре»", "(строка)"]
NOCAT_WIDTHS = [9000, 3200, 2200, 2200, 1500, 4000, 5200, 4400, 1200]
NOCAT_HINT = ("Названия «Наличия», у которых нет категории ни в реестре, ни в «Номенклатуре». В столбце «Категория» уже стоит подсказка по словам листа «Категории» — исправьте её "
              "или очистите. «Записать в «Номенклатуру»» вносит категории в пустую клетку «Категория» уже существующего товара (то, что там написано, не заменяется; товаров, "
              "которых в «Номенклатуре» нет, кнопка не создаёт — добавьте товар и повторите). В учёт и журнал ничего не пишется: категория работает в «Метизах», в подсказках и в поиске "
              "«категория:…». «Найти» отбирает по словам названия.")


def build_nocat(doc, sh):
    build_view110(doc, sh, NOCAT_HEADERS, NOCAT_WIDTHS, NOCAT_HINT, "Найти:", None)
    text = number_format(doc, "@", ("ru", "RU"))
    sh.getCellRangeByPosition(0, 4, 1, 6000).NumberFormat = text
    sh.getCellRangeByPosition(5, 4, 7, 6000).NumberFormat = text
    sh.getCellRangeByPosition(6, 4, 6, 6000).CellBackColor = INPUT_HEAD_10
    sh.getCellRangeByPosition(3, 4, 3, 6000).HoriJustify = 3
    sh.getColumns().getByIndex(8).IsVisible = False
    sh.getCellByPosition(0, 2).CharColor = MUTED
    add_button(doc, sh, "btnNoCatRefresh", "Обновить", "WmsCat.BtnNoCatRefresh", 150, 120, 3000, 650, bold=True, back=0xC6E0B4)
    add_button(doc, sh, "btnNoCatApply", "Записать в «Номенклатуру»", "WmsCat.BtnNoCatApply", 3250, 120, 5400, 650, back=0xDDEBF7)
    bind_event(sh.Events, "OnFocus", "WmsCat.OnNoCatFocus")
    bind_event(sh.Events, "OnChange", "WmsCat.OnNoCatChange")


def categories_102(doc, sheets):
    """1.0.2: «Категории» and «Метизы» for a new book and an upgraded one (after «Комплекты»), the list of the categories on «Заказы» and
    «Иной приход» (a drop-down that accepts any word); nothing that exists is touched"""
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    at = names.index(KITS) + 1 if KITS in names else (names.index(GOODS) + 1 if GOODS in names else len(names))
    if not sheets.hasByName(CATS):
        sheets.insertNewByName(CATS, at)
        build_cats(doc, sheets.getByName(CATS))
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(FAST):
        sheets.insertNewByName(FAST, names.index(CATS) + 1)
        build_fast(doc, sheets.getByName(FAST))
    for sheet, col in ((ORDERS, 18), (SPECIAL, 9)):
        sh = sheets.getByName(sheet)
        rng = sh.getCellRangeByPosition(col, 1, col, LAST_ROW)
        v = rng.Validation
        v.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
        v.setFormula1("$" + CATS + ".$A$2:$A$300")
        v.ShowList = 1
        v.ShowErrorMessage = False
        v.IgnoreBlankCells = True
        rng.Validation = v


# 1.1: «Наличие по названию», «Остаток деталей» (WmsStockView) and «Операции дня» (WmsOps): sheets that only show — the storekeeper's own, not protected,
# not journaled; the data is made again by the buttons and when the sheet is opened after a posting
BYNAME = "Наличие по названию"
PARTS = "Остаток деталей"
OPSHEET = "Операции дня"
BYNAME_HEADERS = ["ЕИ", "Наименование", "Артикул", "Ед.", "Остаток", "Место", "Категория", "Состояние", "Откуда", "Серийный / инв. №", "Дата прихода"]
BYNAME_WIDTHS = [2800, 9000, 3200, 1500, 2200, 2600, 3000, 2400, 4500, 3200, 2600]
PARTS_HEADERS = ["Артикул", "Наименование", "Остаток", "Ед.", "ЕИ с остатком", "Места", "Серийные номера", "Последний приход"]
PARTS_WIDTHS = [3200, 9000, 2200, 1500, 2200, 5000, 7000, 2800]
OPS_HEADERS = ["№ журнала", "Дата и время", "Что сделано", "ЕИ", "Наименование", "Количество", "Ед.", "Место / кому", "Документ", "Подробности"]
OPS_WIDTHS = [1700, 3400, 6500, 2800, 7000, 2000, 1300, 5200, 3400, 9000]
BYNAME_HINT = ("Наличие по названию: копия «Наличия», по одной строке на ЕИ (по умолчанию только с остатком), в порядке названия. Впишите слова в «Найти» — например "
               "«болт м8 цинк» — и нажмите Enter: останутся ЕИ, у которых есть все слова (русские и латинские буквы М/M, х/x не различаются). Эту таблицу можно сортировать "
               "и фильтровать как угодно: «Обновить» (и открытие листа после проведения) строит её заново. Сам реестр «Наличие» сортировать нельзя.")
PARTS_HINT = ("Остаток деталей: одна строка на деталь (по артикулу): сколько осталось на всех её ЕИ, на скольких ЕИ, где лежит, какие серийные номера. Деталь — категория «Детали» "
              "или приход как деталей. Сортируйте и фильтруйте как удобно; «Обновить» строит таблицу заново.")
OPS_HINT = ("Операции дня: что делала WMS по журналу — простыми словами. «Показать» — операции за дату (пусто — сегодня; «по дату» — несколько дней), «Хвост журнала» — "
            "операции, которых ещё нет в книге (перед «Восстановить» или «Отложить хвост журнала» видно, что это), «Вчера» — прошлый рабочий день. Журнал только читается.")
NAMECHK = "Проверка названий"
NAMECHK_HEADERS = ["Где", "Строка", "ЕИ", "Как записано (¶ — разрыв строки)", "Как должно быть", "Что не так", "Что делать"]
NAMECHK_WIDTHS = [2800, 1500, 2800, 9000, 8000, 7000, 9000]
NAMECHK_HINT = ("Проверка названий: названия, разбитые на несколько строк («Винт» / «м3х25» / «Din 912» / «A» / «2»), с лишними пробелами, с разорванным размером (М 8), "
                "классом (A 2) и стандартом строчными (din). «Проверить» — найти их в «Наличии», «Заказах», «Ином приходе» и «Номенклатуре»; «Исправить открытые строки» — "
                "привести к одной строке то, что ещё не проведено, и справочник «Номенклатура». Названия проведённых позиций WMS сама не меняет: исправляйте их кнопкой «Исправить» "
                "на строке-источнике. Новое название, вписанное в «Заказы» или «Иной приход», приводится к одной строке сразу.")
SHEETS_110_VIEWS = [BYNAME, PARTS, OPSHEET, NAMECHK]

# 1.1: «Стандарты» (WmsSpec SH_STD) — the standards that mean the same part (DIN / ISO / ГОСТ), «Синонимы» (WmsSearch SH_SYN) — the words that mean the same:
# the directories of the search, the storekeeper's own sheets (not protected, not journaled); the seeds are examples that he edits
STANDARDS = "Стандарты"
SYNONYMS = "Синонимы"
STANDARDS_HEADERS = ["DIN", "ISO", "ГОСТ"]
STANDARDS_SEED = [
    ("DIN 931", "ISO 4014", "ГОСТ 7798"), ("DIN 933", "ISO 4017", "ГОСТ 7805"), ("DIN 934", "ISO 4032", "ГОСТ 5915"), ("DIN 439", "ISO 4035", "ГОСТ 5916"),
    ("DIN 125", "ISO 7089", "ГОСТ 11371"), ("DIN 9021", "ISO 7093", "ГОСТ 6958"), ("DIN 127", "", "ГОСТ 6402"), ("DIN 912", "ISO 4762", "ГОСТ 11738"),
    ("DIN 7991", "ISO 10642", ""), ("DIN 965", "ISO 7046", "ГОСТ 17475"), ("DIN 603", "ISO 8677", "ГОСТ 7802"), ("DIN 985", "ISO 10511", ""),
    ("DIN 982", "ISO 7040", ""), ("DIN 1587", "", "ГОСТ 11860"), ("DIN 94", "ISO 1234", "ГОСТ 397"), ("DIN 7", "ISO 2338", "ГОСТ 3128"),
    ("DIN 471", "", "ГОСТ 13942"), ("DIN 472", "", "ГОСТ 13943"), ("DIN 913", "ISO 4026", ""), ("DIN 914", "ISO 4027", ""), ("DIN 915", "ISO 4028", ""),
    ("DIN 916", "ISO 4029", ""),
]
STANDARDS_HINT = ("Какие стандарты — об одной и той же детали. Поиск «din 933» найдёт и «ISO 4017», и «ГОСТ 7805», если они в одной строке. Строка — одна деталь; пустая "
                  "клетка допустима (достаточно двух стандартов). Это ориентиры: перед закупкой сверяйтесь с документами — эквивалентность бывает не полной (класс точности, "
                  "шаг резьбы, размер под ключ). Допишите свои строки; год у ГОСТ («ГОСТ 7805-70») писать не нужно. WMS эту таблицу только читает, учёт она не меняет.")
SYNONYMS_HEADERS = ["Слова, означающие одно и то же (через запятую)"]
SYNONYMS_SEED = ["саморез, шуруп", "гровер, пружинная", "акб, аккумулятор", "перчатки, рукавицы", "болгарка, ушм", "стремянка, лестница"]
SYNONYMS_HINT = ("Слова одной строки поиск считает равными: «саморез» найдёт и «шуруп». Одно слово — одна клетка списка через запятую; по одному слову на деталь, "
                 "без пробелов внутри. Падежи и числа писать не нужно («болты» найдёт «болт»). Это примеры — меняйте под себя. Учёт эта таблица не меняет.")


def build_standards(doc, sh):
    header(sh, STANDARDS_HEADERS, [4000, 4000, 4500], height=900)
    text = number_format(doc, "@", ("ru", "RU"))
    sh.getCellRangeByPosition(0, 0, 2, 400).NumberFormat = text
    sh.getCellRangeByPosition(0, 1, 2, 400).CellBackColor = INPUT_HEAD_10
    for r, row in enumerate(STANDARDS_SEED, start=1):
        for c, v in enumerate(row):
            if v:
                sh.getCellByPosition(c, r).setString(v)
    hint = sh.getCellByPosition(4, 0)
    hint.setString(STANDARDS_HINT)
    hint.IsTextWrapped = True
    hint.VertJustify = 1
    hint.CharHeight = 9
    hint.CharColor = MUTED
    sh.getColumns().getByIndex(3).Width = 600
    sh.getColumns().getByIndex(4).Width = 12000
    bind_event(sh.Events, "OnChange", "WmsSpec.OnStdChange")


def build_synonyms(doc, sh):
    header(sh, SYNONYMS_HEADERS, [14000], height=900)
    text = number_format(doc, "@", ("ru", "RU"))
    sh.getCellRangeByPosition(0, 0, 0, 300).NumberFormat = text
    sh.getCellRangeByPosition(0, 1, 0, 300).CellBackColor = INPUT_HEAD_10
    for r, v in enumerate(SYNONYMS_SEED, start=1):
        sh.getCellByPosition(0, r).setString(v)
    hint = sh.getCellByPosition(2, 0)
    hint.setString(SYNONYMS_HINT)
    hint.IsTextWrapped = True
    hint.VertJustify = 1
    hint.CharHeight = 9
    hint.CharColor = MUTED
    sh.getColumns().getByIndex(1).Width = 600
    sh.getColumns().getByIndex(2).Width = 12000
    bind_event(sh.Events, "OnChange", "WmsSearch.OnSynChange")


def build_view110(doc, sh, headers, widths, hint, find_label, zero_label, merge_find=True):
    header(sh, headers, widths, height=1500)
    sh.getRows().getByIndex(0).Height = 1500
    n = len(headers)
    for c in range(n):
        sh.getCellByPosition(c, 3).setString(headers[c])
        sh.getCellByPosition(c, 0).setString("")
    hd = sh.getCellRangeByPosition(0, 3, n - 1, 3)
    hd.CharWeight = 150
    hd.IsTextWrapped = True
    hd.VertJustify = 2
    hd.CellBackColor = 0xE7E6E6
    sh.getCellRangeByPosition(0, 0, n - 1, 0).CellBackColor = 0xFFFFFF
    text = number_format(doc, "@", ("ru", "RU"))
    if find_label:
        sh.getCellByPosition(0, 1).setString(find_label)
        sh.getCellByPosition(0, 1).CharWeight = 150
        sh.getCellByPosition(1, 1).CellBackColor = INPUT_HEAD_10
        sh.getCellByPosition(1, 1).NumberFormat = text
        if merge_find:
            merge(sh, 1, 1, 3, 1)
    if zero_label:
        sh.getCellByPosition(4, 1).setString(zero_label)
        sh.getCellByPosition(4, 1).CharWeight = 150
        sh.getCellByPosition(4, 1).HoriJustify = 3
        sh.getCellByPosition(5, 1).setString("нет")
        sh.getCellByPosition(5, 1).CellBackColor = INPUT_HEAD_10
        sh.getCellByPosition(5, 1).NumberFormat = text
    sh.getCellByPosition(0, 2).CharColor = MUTED
    h = sh.getCellByPosition(n + 1, 0)
    h.setString(hint)
    h.IsTextWrapped = True
    h.VertJustify = 1
    h.CharHeight = 9
    h.CharColor = MUTED
    sh.getColumns().getByIndex(n).Width = 600
    sh.getColumns().getByIndex(n + 1).Width = 12000
    return sh


def build_byname(doc, sh):
    build_view110(doc, sh, BYNAME_HEADERS, BYNAME_WIDTHS, BYNAME_HINT, "Найти:", "С нулевым остатком (да/нет):")
    sh.getCellRangeByPosition(4, 4, 4, 60000).HoriJustify = 3
    sh.getCellRangeByPosition(10, 4, 10, 60000).NumberFormat = number_format(doc, "DD.MM.YYYY")
    add_button(doc, sh, "btnByNameRefresh", "Обновить", "WmsStockView.BtnByNameRefresh", 150, 120, 3000, 650, bold=True, back=0xC6E0B4)
    bind_event(sh.Events, "OnFocus", "WmsStockView.OnByNameFocus")
    bind_event(sh.Events, "OnChange", "WmsStockView.OnByNameChange")


def build_parts(doc, sh):
    build_view110(doc, sh, PARTS_HEADERS, PARTS_WIDTHS, PARTS_HINT, "Найти:", "С нулевым остатком (да/нет):")
    sh.getCellRangeByPosition(2, 4, 2, 60000).HoriJustify = 3
    sh.getCellRangeByPosition(7, 4, 7, 60000).NumberFormat = number_format(doc, "DD.MM.YYYY")
    add_button(doc, sh, "btnPartsRefresh", "Обновить", "WmsStockView.BtnPartsRefresh", 150, 120, 3000, 650, bold=True, back=0xC6E0B4)
    bind_event(sh.Events, "OnFocus", "WmsStockView.OnPartsFocus")
    bind_event(sh.Events, "OnChange", "WmsStockView.OnPartsChange")


def build_ops(doc, sh):
    build_view110(doc, sh, OPS_HEADERS, OPS_WIDTHS, OPS_HINT, "С даты:", None, merge_find=False)
    sh.getCellByPosition(2, 1).setString("по дату:")
    sh.getCellByPosition(2, 1).CharWeight = 150
    sh.getCellByPosition(2, 1).HoriJustify = 3
    sh.getCellByPosition(3, 1).CellBackColor = INPUT_HEAD_10
    sh.getCellRangeByPosition(0, 4, 0, 6000).HoriJustify = 3
    sh.getCellRangeByPosition(5, 4, 5, 6000).HoriJustify = 3
    add_button(doc, sh, "btnOpsShow", "Показать", "WmsOps.BtnOpsShow", 150, 120, 2800, 650, bold=True, back=0xC6E0B4)
    add_button(doc, sh, "btnOpsTail", "Хвост журнала", "WmsOps.BtnOpsTail", 3050, 120, 3400, 650, back=0xF4CCCC)
    add_button(doc, sh, "btnOpsYest", "Вчера (рабочий день)", "WmsOps.BtnOpsYesterday", 6550, 120, 4200, 650, back=0xDDEBF7)
    bind_event(sh.Events, "OnFocus", "WmsOps.OnOpsFocus")
    bind_event(sh.Events, "OnChange", "WmsOps.OnOpsChange")


def build_namechk(doc, sh):
    build_view110(doc, sh, NAMECHK_HEADERS, NAMECHK_WIDTHS, NAMECHK_HINT, None, None)
    sh.getCellRangeByPosition(0, 4, NAMECHK_HEADERS.__len__() - 1, 3000).IsTextWrapped = True
    add_button(doc, sh, "btnNamesCheck", "Проверить", "WmsNames.BtnNamesCheck", 150, 120, 3000, 650, bold=True, back=0xC6E0B4)
    add_button(doc, sh, "btnNamesFix", "Исправить открытые строки", "WmsNames.BtnNamesFix", 3250, 120, 5400, 650, back=0xDDEBF7)


def views_110(doc, sheets):
    """1.1: the three sheets that only show, for a new book and an upgraded one: «Наличие по названию» and «Остаток деталей» after «Наличие», «Операции дня»
    after «Отчёты»; nothing that exists is touched"""
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(BYNAME):
        sheets.insertNewByName(BYNAME, names.index(STOCK) + 1 if STOCK in names else len(names))
        build_byname(doc, sheets.getByName(BYNAME))
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(PARTS):
        sheets.insertNewByName(PARTS, names.index(BYNAME) + 1)
        build_parts(doc, sheets.getByName(PARTS))
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(OPSHEET):
        sheets.insertNewByName(OPSHEET, names.index(REPORTS) + 1 if REPORTS in names else len(names))
        build_ops(doc, sheets.getByName(OPSHEET))
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(NAMECHK):
        sheets.insertNewByName(NAMECHK, names.index(OPSHEET) + 1)
        build_namechk(doc, sheets.getByName(NAMECHK))
    # the directories of the search follow «Метизы» (or «Категории», or «Номенклатура»)
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    after = FAST if FAST in names else (CATS if CATS in names else (GOODS if GOODS in names else None))
    if not sheets.hasByName(STANDARDS):
        sheets.insertNewByName(STANDARDS, names.index(after) + 1 if after else len(names))
        build_standards(doc, sheets.getByName(STANDARDS))
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(SYNONYMS):
        sheets.insertNewByName(SYNONYMS, names.index(STANDARDS) + 1)
        build_synonyms(doc, sheets.getByName(SYNONYMS))
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(NOCAT):
        sheets.insertNewByName(NOCAT, names.index(CATS) + 1 if CATS in names else len(names))
        build_nocat(doc, sheets.getByName(NOCAT))
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(SETTINGS):
        sheets.insertNewByName(SETTINGS, names.index(NOCAT) + 1 if NOCAT in names else len(names))
        build_settings(doc, sheets.getByName(SETTINGS))
    # the sheets of 1.0.2 get the layout of 1.1 (a new book has it already)
    cats_110(doc, sheets)
    fast_110(doc, sheets)


# 1.1: the buttons of «Главная» that 1.1 adds (the free places of the third row of the system buttons)
MAIN_BUTTONS_110 = [("btnDayOps", "Операции дня", "WmsOps.BtnOpsOpen", 10), ("btnWhatNow", "Что делать?", "WmsGuard.BtnWhatNow", 11)]


def main_buttons_110(doc, sheets):
    sh = sheets.getByName(MAIN)
    for name, label, macro, i in MAIN_BUTTONS_110:
        if not has_button(sh, name):
            y = sh.getCellByPosition(0, 3 + i // 4).Position.Y
            add_button(doc, sh, name, label, macro, MAIN_LEFT + 100 + (i % 4) * 5700, y + 50, 5640, 750, back=0xDDEBF7)


# 1.0.2: the new buttons — only the missing ones are added (a new book and an upgraded one both)
BUTTONS_102 = [
    (STOCK, "btnSerials", "Серийные…", "WmsParts.BtnSerials", 23450, 880, 3400, 650, 0xDDEBF7),
    (ISSUES, "btnAct", "Акт…", "WmsDocs.BtnAct", 25100, 880, 2900, 650, 0xC6E0B4),
]


def buttons_102(doc, sheets):
    for sheet, name, label, macro, x, y, w, h, back in BUTTONS_102:
        if sheets.hasByName(sheet):
            sh = sheets.getByName(sheet)
            if not has_button(sh, name):
                add_button(doc, sh, name, label, macro, x, y, w, h, back=back)


def ledger_10(doc, sheets):
    """1.0 (этап G): what a new book and an upgraded one both get — the hidden _CLOSE (the closed months) and the page «Отчёты»"""
    if not sheets.hasByName(CLOSE):
        names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
        sheets.insertNewByName(CLOSE, names.index("_SYS") if "_SYS" in names else sheets.getCount())
        sh = sheets.getByName(CLOSE)
        for c, h in enumerate(CLOSE_HEADERS):
            sh.getCellByPosition(c, 0).setString(h)
        text = number_format(doc, "@", ("ru", "RU"))
        for c in (0, 2, 3, 6, 7):
            sh.getCellRangeByPosition(c, 1, c, 1000).NumberFormat = text
        sh.getCellRangeByPosition(1, 1, 1, 1000).NumberFormat = number_format(doc, "DD.MM.YYYY")
        sh.IsVisible = False
    if not sheets.hasByName(REPORTS):
        names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
        sheets.insertNewByName(REPORTS, names.index(PRINT) if PRINT in names else names.index(FIND) + 1)
        build_reports(doc, sheets.getByName(REPORTS))
    # А3: the department of a recipient — the storekeeper's column C of «Получатели»
    rc = sheets.getByName(RCPT)
    cur = rc.createCursor()
    cur.gotoEndOfUsedArea(False)
    last = cur.getRangeAddress().EndRow
    # only an empty column C gets the header (cells holding an empty text count as empty): a column the storekeeper already
    # uses for something else stays as it is
    if all(str(r[0]).strip() == "" for r in rc.getCellRangeByPosition(2, 0, 2, max(last, 0)).getDataArray()):
        h = rc.getCellByPosition(2, 0)
        h.setString("Подразделение")
        h.CharWeight = 150
        h.IsTextWrapped = True
        h.VertJustify = 3
        h.CellBackColor = rc.getCellByPosition(1, 0).CellBackColor
        rc.getColumns().getByIndex(2).Width = 5000
    dash_10(doc, sheets)


def print_10(doc, sheets):
    """1.0: the page «Печать» for a new book and an upgraded one"""
    if not sheets.hasByName(PRINT):
        names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
        sheets.insertNewByName(PRINT, names.index(SUPPLY) + 1 if SUPPLY in names else names.index(FIND) + 1)
        build_print(doc, sheets.getByName(PRINT))


def supply_10(doc, sheets):
    """1.0 (этап F): what a new book and an upgraded one both get — «Поставщики», «Номенклатура», «Снабжение»; the list of the
    suppliers on «Заказы» L and on the panel of «Приход авто»; the events of «Снабжение»"""
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(SUPPLIERS):
        sheets.insertNewByName(SUPPLIERS, names.index(RCPT) + 1)
        build_suppliers(doc, sheets.getByName(SUPPLIERS))
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(GOODS):
        sheets.insertNewByName(GOODS, names.index(SUPPLIERS) + 1)
        build_goods(doc, sheets.getByName(GOODS))
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    if not sheets.hasByName(SUPPLY):
        sheets.insertNewByName(SUPPLY, names.index(FIND) + 1)
        build_supply(doc, sheets.getByName(SUPPLY))
    src = f"$'{SUPPLIERS}'.$A$2:$A$1048576"
    list_validation(sheets.getByName(ORDERS).getCellRangeByPosition(11, 1, 11, LAST_ROW), src)
    list_validation(sheets.getByName(CARS).getCellByPosition(3, 1), src)          # the panel: «Поставщик» D2 (WmsConfig CP_SUP_COL, CP_IN_ROW)
    bind_event(sheets.getByName(SUPPLY).Events, "OnDoubleClick", "WmsSupply.OnSupplyDoubleClick")
    bind_event(sheets.getByName(SUPPLY).Events, "OnChange", "WmsSupply.OnSupplyChange")


def layout_10(doc, sheets):
    """1.0: what a new book and an upgraded one both get (the data are not touched, only formats): the place columns are text —
    «5.12.3» typed there stays a place, Calc does not take it for a date (Заказы U, Возврат J, Иной приход I, Корректировки J);
    the directory «Места» after «Топология»; the events of «Топология» and «Места» (the known places are read again)"""
    text = number_format(doc, "@", ("ru", "RU"))
    # the styles of the layers of the map, the new ones of 1.0 too (a protected book cannot create them at run time)
    view_styles(doc)
    for sheet, c in ((ORDERS, 20), (RETURNS, 9), (SPECIAL, 8), (ADJUST, 9)):
        sheets.getByName(sheet).getCellRangeByPosition(c, 1, c, LAST_ROW).NumberFormat = text
    # the headers of the working sheets: the buttons never hide a line of a header and no word is cut in two
    # («Количест-во», «измерени-я»): the header row is higher, a column is never narrower than the longest word of its header
    for sheet, titles, height in ((ISSUES, ISSUE_HEADERS, 2500), (RETURNS, RETURN_HEADERS, 2500), (SPECIAL, SPECIAL_HEADERS, 2500),
                                  (ADJUST, ADJUST_HEADERS, 2500), (ORDERS, ORDER_HEADERS, 3100)):
        fit_headers(sheets.getByName(sheet), titles, height)
    # the recipient of «Выдачи» (I) and «Возврат» (I): a list of «Получатели» to choose from — any other text may still be typed,
    # WMS asks about a name it does not know
    for sheet in (ISSUES, RETURNS):
        list_validation(sheets.getByName(sheet).getCellRangeByPosition(8, 1, 8, LAST_ROW), f"$'{RCPT}'.$B$2:$B$1048576")
    # 1.0 (М1): the kind «Разделение» in the list of «Корректировки» B, the column «Новый ЕИ» of _ADJ (L)
    rng = sheets.getByName(ADJUST).getCellRangeByPosition(1, 1, 1, LAST_ROW)
    v = rng.Validation
    v.setFormula1(";".join(f'"{t}"' for t in ADJUST_KINDS))
    rng.Validation = v
    adj = sheets.getByName(ADJ)
    if adj.getCellByPosition(len(ADJ_HEADERS) - 1, 0).getString() == "":
        adj.getCellByPosition(len(ADJ_HEADERS) - 1, 0).setString(ADJ_HEADERS[-1])
    # «Наличие» K: the balance in the parts of the packing («3 кан», a formula WMS puts in the row of an EI with a packing — Ч1)
    st = sheets.getByName(STOCK)
    k = st.getCellByPosition(len(STOCK_HEADERS), 0)
    if k.getString() == "":
        j = st.getCellByPosition(len(STOCK_HEADERS) - 1, 0)
        k.setString("Остаток частями")
        for prop in ("CharWeight", "CellBackColor", "IsTextWrapped", "VertJustify", "HoriJustify", "TopBorder", "BottomBorder", "LeftBorder",
                     "RightBorder", "CharHeight"):
            setattr(k, prop, getattr(j, prop))
    if st.getColumns().getByIndex(len(STOCK_HEADERS)).Width < 3000:
        st.getColumns().getByIndex(len(STOCK_HEADERS)).Width = 3000
    # the status W of «Заказы» is read whole («Частично получено / остаток отменён»)
    w = sheets.getByName(ORDERS).getColumns().getByIndex(22)
    if w.Width < 5600:
        w.Width = 5600
    if not sheets.hasByName(PLACES):
        names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
        sheets.insertNewByName(PLACES, names.index(TOPO) + 1)
        build_places(doc, sheets.getByName(PLACES))
    bind_event(sheets.getByName(TOPO).Events, "OnChange", "WmsMap.OnTopoChange")
    bind_event(sheets.getByName(PLACES).Events, "OnChange", "WmsMap.OnPlacesChange")
    # the state line of «Дашборд» is written while the book is saved («WMS не запущена» for the macros off): unlocked, so that
    # no change of the protection is needed then (WmsDash.DashSavedState)
    sheets.getByName(DASH).getCellByPosition(8, 0).CellProtection = protection(False)


# 1.0 (В5): the columns frozen with the header row (№ / EI and the name of the row)
FREEZE_COLS_10 = {ISSUES: 3, ORDERS: 2, RETURNS: 4, ADJUST: 4, SPECIAL: 4, STOCK: 2}
# 1.0 (В9): the colour of the sheet tabs by group — pages, the sheets the storekeeper works in, the registers and directories,
# the service sheets
TAB_COLORS_10 = [((DASH, MAIN, FIND, MAP, SUPPLY, REPORTS, SERVICE, PRINT), 0x5B9BD5), ((ORDERS, SPECIAL, CARS, ISSUES, RETURNS, ADJUST), 0x70AD47),
                 ((STOCK, RCPT, TOPO, PLACES, SUPPLIERS, GOODS, KITS, RESERVES), 0xED7D31), ((LEGREP, HELP), 0xA5A5A5)]
# 1.0 (В1): what the storekeeper fills in is white, what WMS writes is grey (data rows), the headers likewise
INPUT_HEAD_10 = 0xFFF2CC
WMS_HEAD_10 = 0xE7E6E6
WMS_FILL_10 = 0xF5F5F5
LOCKS_10 = {ISSUES: ISSUE_LOCKS_OPEN, ORDERS: ORDER_LOCKS_OPEN, RETURNS: RETURN_LOCKS_OPEN, SPECIAL: SPECIAL_LOCKS_OPEN, ADJUST: ADJUST_LOCKS_OPEN}
# 1.0 (В2): one meaning per colour of a button — doing (green), undoing (red), helping (blue), the rest neutral
BTN_GREEN, BTN_RED, BTN_BLUE = 0xC6E0B4, 0xF4CCCC, 0xDDEBF7


def button_color(name):
    n = name.lower()
    if any(k in n for k in ("post", "receive", "issuewin", "issuefrom", "returnhands", "find")) and "go" not in n:
        return BTN_GREEN
    if any(k in n for k in ("delete", "cancel", "abandon", "unlock", "clear")):
        return BTN_RED
    return BTN_BLUE


def colors_10(doc, sheets):
    """В1, В2, В9 (the data are not touched): the input columns and those WMS writes, the buttons by meaning, the tabs by group"""
    for sheet, locks in LOCKS_10.items():
        sh = sheets.getByName(sheet)
        for c, bit in enumerate(locks):
            sh.getCellByPosition(c, 0).CellBackColor = INPUT_HEAD_10 if bit == "0" else WMS_HEAD_10
            if bit == "1":
                sh.getCellRangeByPosition(c, 1, c, LAST_ROW).CellBackColor = WMS_FILL_10
    for i in range(sheets.getCount()):
        sh = sheets.getByIndex(i)
        dp = sh.getDrawPage()
        if dp.getForms().getCount() == 0:
            continue
        form = dp.getForms().getByIndex(0)
        for k in range(form.getCount()):
            m = form.getByIndex(k)
            if m.supportsService("com.sun.star.form.component.CommandButton") and sh.Name not in (MAIN, DASH, MAP):
                m.BackgroundColor = button_color(m.Name)
    for names, color in TAB_COLORS_10:
        for name in names:
            if sheets.hasByName(name):
                sheets.getByName(name).TabColor = color
    # «Сверка переноса» is looked at once after the transfer: it goes to the end, before «Справка»
    if sheets.hasByName(LEGREP) and sheets.hasByName(HELP):
        names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
        if names.index(LEGREP) != names.index(HELP) - 1:
            sheets.moveByName(LEGREP, names.index(HELP))


# 1.0.1: «Наличие» L…V — what is known of the goods of an EI: how much came in and how much is left (side by side), for whom
# they were ordered, where they came from, the packing, the serial number, «годен до», the reserve. View only: WmsItem.
# StockInfoRefresh fills them from the sheets when «Наличие» is opened after an operation and by «Обновить сведения» (never
# inside a posting); «Осталось» is a formula over E, so it is never behind the balance
STOCK_INFO_COL = 11                # WmsItem.SI_COL0
STOCK_INFO_HEADERS = ["Пришло", "Осталось", "Кому назначено", "Поставщик / откуда", "Дата прихода", "Документ прихода", "№ заказа",
                      "Фасовка", "Серийный / инв. №", "Годен до", "Резерв"]
STOCK_INFO_WIDTHS = [2300, 2300, 3800, 4400, 2600, 3200, 2400, 3800, 3200, 2600, 4400]
STOCK_INFO_QTY = (0, 1)               # «Пришло», «Осталось» (a formula: always the balance of E) — the number format of E
STOCK_INFO_DATES = (4, 9)             # «Дата прихода», «Годен до»


def stock_info_10(doc, sheets):
    """1.0.1: the headers, widths and formats of «Наличие» L…V, the button «Обновить сведения», the event of the sheet, the
    autofilter over them — for a new book and an upgraded one (the data are not touched)"""
    st = sheets.getByName(STOCK)
    j = st.getCellByPosition(len(STOCK_HEADERS) - 1, 0)
    date = number_format(doc, "DD.MM.YYYY", ("ru", "RU"))
    for k, (title, w) in enumerate(zip(STOCK_INFO_HEADERS, STOCK_INFO_WIDTHS)):
        c = st.getCellByPosition(STOCK_INFO_COL + k, 0)
        if c.getString() == "":
            c.setString(title)
            for prop in ("CharWeight", "CellBackColor", "IsTextWrapped", "VertJustify", "HoriJustify", "TopBorder", "BottomBorder",
                         "LeftBorder", "RightBorder", "CharHeight"):
                setattr(c, prop, getattr(j, prop))
        col = st.getColumns().getByIndex(STOCK_INFO_COL + k)
        if col.Width < w:
            col.Width = w
    for k in STOCK_INFO_DATES:
        st.getCellRangeByPosition(STOCK_INFO_COL + k, 1, STOCK_INFO_COL + k, LAST_ROW).NumberFormat = date
    qty = st.getCellByPosition(4, 1).NumberFormat
    for k in STOCK_INFO_QTY:
        st.getCellRangeByPosition(STOCK_INFO_COL + k, 1, STOCK_INFO_COL + k, LAST_ROW).NumberFormat = qty
    if not has_button(st, "btnStockInfo"):
        add_button(doc, st, "btnStockInfo", "Обновить сведения", "WmsItem.BtnStockInfo", 15100, 880, 4200, 650, back=0xDDEBF7)
    # 1.0.1: a kit from the EIs selected here; on «Комплекты» the same button and the hint that tells it
    if not has_button(st, "btnKitFromEis"):
        add_button(doc, st, "btnKitFromEis", "Собрать из ЕИ…", "WmsKit.BtnKitFromEis", 19450, 880, 3800, 650, back=0xC6E0B4)
    if sheets.hasByName(KITS):
        ks = sheets.getByName(KITS)
        if not has_button(ks, "btnKitFromEis"):
            add_button(doc, ks, "btnKitFromEis", "Собрать из ЕИ…", "WmsKit.BtnKitFromEis", 9250, 120, 4400, 650, back=0xC6E0B4)
        ks.getCellByPosition(7, 0).setString(KITS_HINT)
    bind_event(st.Events, "OnFocus", "WmsItem.OnStockFocus")
    if doc.DatabaseRanges.hasByName("WMS_STOCK"):
        dbr = doc.DatabaseRanges.getByName("WMS_STOCK")
        a = dbr.DataArea
        if a.EndColumn < STOCK_INFO_COL + len(STOCK_INFO_HEADERS) - 1:
            a.EndColumn = STOCK_INFO_COL + len(STOCK_INFO_HEADERS) - 1
            dbr.DataArea = a
            dbr.AutoFilter = True


# 1.0 (В8): «Наличие» reads easier — a zero balance grey, every second row a shade darker, a gap between the balance and the place.
# One conditional format per column range (Calc applies the first true entry of the first format of a cell): the status of H
# keeps its colours first
STOCK_ZERO, STOCK_BAND = 0x9AA3AD, 0xF3F6FA


def stock_view_10(doc, sheets):
    sh = sheets.getByName(STOCK)
    cell_style(doc, "WMS_StockZero", CharColor=STOCK_ZERO)
    cell_style(doc, "WMS_StockBand", CellBackColor=STOCK_BAND)
    cell_style(doc, "WMS_StockZeroBand", CharColor=STOCK_ZERO, CellBackColor=STOCK_BAND)
    used, zero, even = '$A2<>""', 'N($E2)=0', 'ISEVEN(ROW())'
    common = [(f"AND({used};{zero};{even})", "WMS_StockZeroBand"), (f"AND({used};{zero})", "WMS_StockZero"), (f"AND({used};{even})", "WMS_StockBand")]
    status = [('$H2="Требует разбора"', "ST_Part"), ('LEFT($H2;13)="Приход удалён"', "ST_Gray")]
    last = len(STOCK_HEADERS) + len(STOCK_INFO_HEADERS)          # with K «Остаток частями» and L…V (1.0.1)
    for c0, c1, rules in ((0, 6, common), (7, 7, status + common), (8, last, common)):
        rng = sh.getCellRangeByPosition(c0, 1, c1, LAST_ROW)
        cf = rng.ConditionalFormat
        cf.clear()
        rng.ConditionalFormat = cf
        for formula, style in rules:
            add_cond(sh, rng, formula, style, c0, 1)
    sh.getCellRangeByPosition(5, 1, 5, LAST_ROW).ParaIndent = 300


# 1.0 (В3): what may be typed in a column — the input help of Calc (shown while the cell is selected) and the same text as a
# note of the header (shown when the mouse is over it). (title, text) by column; the data and the lists of the columns stay
_DATE = ("Дата", "ДД.ММ.ГГГГ, например 14.10.2026. Будущая или очень старая дата — WMS переспросит.")
_EI = ("ЕИ", "EI-427, просто 427 или скан этикетки. Не помните номер — «Подобрать ЕИ» (Ctrl+Shift+E).")
_PLACE = ("Место", "1-2-3; несколько мест — через запятую: 1-2-3, 1-2-4. Свободное место подскажет «Куда положить?» (Ctrl+Shift+Q).")
_QTY_IN = ("Количество", "Сколько пришло: 4 · 2,5 · 10 — в единице документа.")
HINTS_10 = {
    ISSUES: {1: ("№ документа", "Необязательно. Несколько выдач одним документом: выделите строки и «Провести» — WMS предложит общий номер."),
             4: ("Количество", "2 · 0,25 · 1 кан · 1,5 м · всё (весь остаток). Целое число у ЕИ с фасовкой — WMS спросит: шт или кан."),
             5: ("Количество в %", "25 = 25 % прихода этого ЕИ (можно и «25%» в «Количество»). Меньше 1 % — WMS переспросит."),
             7: ("Дата выдачи", "Обязательно. " + _DATE[1]),
             8: ("Кому выдано", "Из списка (стрелка справа) или новое имя — WMS спросит, добавить ли его в «Получатели»."),
             11: _EI},
    ORDERS: {5: ("Фактическое количество", "Сколько пришло по факту: 4 · 2,5 · 10. Затем «Провести приход» — или всё в одном окне: «Принять…»."),
             13: ("Дата поступления", _DATE[1]), 20: _PLACE},
    RETURNS: {1: ("№ выдачи", "Номер проведённой выдачи (колонка A «Выдач»). Всё, что на руках у человека, — «Поиск»: имя и «Вернуть выделенное»."),
              5: ("Количество возврата", "2 · 0,5 · 1 кан · 1,5 м — не больше, чем выдано и ещё не возвращено."),
              7: ("Дата возврата", _DATE[1]),
              8: ("От кого возвращено", "Кто вернул — из списка (стрелка справа) или имя, как в «Получатели»."),
              9: ("Место", "Куда положили: 1-2-3. У ЕИ на нескольких местах WMS спросит — добавить место или заменить.")},
    SPECIAL: {1: ("Тип прихода", "Офис, Производство, Детали, Старый склад или Иной — из списка."), 5: _QTY_IN,
              7: ("Дата поступления", _DATE[1]), 8: _PLACE},
    ADJUST: {1: ("Вид", "Перемещение, Списание, Инвентаризация, Разделение (часть ЕИ — в новый ЕИ) или Возврат поставщику (брак, пересорт) — из списка."), 2: _EI,
             6: ("Списать / отделить / вернуть", "Списание и возврат поставщику: 2 · 0,25 · 1 кан · 1,5 м · всё. Разделение: сколько отделить в новый ЕИ — 2 кан, 4,5 м."),
             7: ("Фактический остаток", "Инвентаризация: сколько насчитали — 3 · 3 кан · 4,5 м; 0 — на месте пусто. Совпало с учётом — так и запишется."),
             9: ("Новое место", "Перемещение: куда переложили — 1-2-3, несколько мест через запятую, «1-2-3…5»; «+1-2-5» — добавить место, «−1-2-3» — убрать. "
                 "Разделение: место нового ЕИ. Возврат поставщику: номер документа возврата."),
             10: _DATE},
}


def hints_10(doc, sheets):
    from com.sun.star.table import CellAddress
    for sheet, cols in HINTS_10.items():
        sh = sheets.getByName(sheet)
        notes = sh.getAnnotations()
        for c, (title, text) in cols.items():
            rng = sh.getCellRangeByPosition(c, 1, c, LAST_ROW)
            v = rng.Validation
            v.ShowInputMessage = True
            v.InputTitle = title
            v.InputMessage = text
            rng.Validation = v
            head = sh.getCellByPosition(c, 0)
            if head.getAnnotation().getString() == "":
                a = CellAddress()
                a.Sheet, a.Column, a.Row = sh.RangeAddress.Sheet, c, 0
                notes.insertNew(a, text)


# 1.0 (Р11): the keyboard shortcuts of the book (stored in the document: they act in this book only)
HOTKEYS_10 = [("P", "WmsDaily.KeyPost", "Провести (лист в работе)"), ("E", "WmsFind.BtnPickEI", "Подобрать ЕИ"),
              ("F", "WmsDaily.KeyFind", "Найти (страница «Поиск»)"), ("W", "WmsDaily.BtnIssueWindow", "Выдать… (окно)"),
              ("Q", "WmsMap.BtnWherePut", "Куда положить?")]


def hotkeys_10(doc):
    from com.sun.star.awt import KeyEvent
    sm = doc.getUIConfigurationManager().getShortCutManager()
    mods = uno.getConstantByName("com.sun.star.awt.KeyModifier.SHIFT") | uno.getConstantByName("com.sun.star.awt.KeyModifier.MOD1")
    for key, macro, _ in HOTKEYS_10:
        ke = KeyEvent()
        ke.KeyCode = uno.getConstantByName(f"com.sun.star.awt.Key.{key}")
        ke.Modifiers = mods
        sm.setKeyEvent(ke, macro_url(macro))
    sm.store()
    doc.getUIConfigurationManager().store()


def has_button(sheet, name):
    dp = sheet.getDrawPage()
    if dp.getForms().getCount() == 0:
        return False
    form = dp.getForms().getByIndex(0)
    return any(form.getByIndex(i).Name == name for i in range(form.getCount()))


# 1.0: the buttons of the everyday work (stage B): «Выдачи» gets a second row of buttons (its header as high as that of
# «Заказы»), «Наличие» and «Поиск» get «Выдать». (sheet, name, label, macro, x, y, w, h, colour); only the missing ones are added
BUTTONS_10 = [
    (ISSUES, "btnIssueWin", "Выдать…", "WmsDaily.BtnIssueWindow", 150, 880, 2750, 650, 0xC6E0B4),
    (ISSUES, "btnRepeat", "Повторить", "WmsDaily.BtnRepeatIssue", 150 + 2900, 880, 2750, 650, 0xDDEBF7),
    (ISSUES, "btnWhy", "Почему?", "WmsDaily.BtnWhy", 150 + 2 * 2900, 880, 2750, 650, 0xDDEBF7),
    (ISSUES, "btnOffcut", "Обрезок…", "WmsDaily.BtnOffcut", 150 + 3 * 2900, 880, 2750, 650, 0xDDEBF7),
    (ISSUES, "btnRoute", "Маршрут сбора", "WmsFind.BtnPickRoute", 150 + 4 * 2900, 880, 2750, 650, 0xDDEBF7),
    (ORDERS, "btnWhy", "Почему?", "WmsDaily.BtnWhy", 150 + 6 * 3050, 120, 2900, 650, 0xDDEBF7),
    (RETURNS, "btnWhy", "Почему?", "WmsDaily.BtnWhy", 150 + 6 * 2900, 120, 2750, 650, 0xDDEBF7),
    (SPECIAL, "btnWhy", "Почему?", "WmsDaily.BtnWhy", 150 + 7 * 2900, 120, 2750, 650, 0xDDEBF7),
    (ADJUST, "btnWhy", "Почему?", "WmsDaily.BtnWhy", 150 + 7 * 2900, 120, 2750, 650, 0xDDEBF7),
    (STOCK, "btnIssueFrom", "Выдать", "WmsDaily.BtnIssueFrom", 11800, 120, 2400, 650, 0xC6E0B4),
    (STOCK, "btnPackBatch", "Фасовка пачкой…", "WmsPack.BtnPackBatch", 14350, 120, 3600, 650, 0xDDEBF7),
    (STOCK, "btnSplit", "Разделить ЕИ…", "WmsDaily.BtnSplitWindow", 18100, 120, 3400, 650, 0xDDEBF7),
    # 1.0: the map edited on the map itself (по просьбе пользователя) and the check of the layout (М3)
    (MAP, "btnMapEdit", "Править карту", "WmsMap.BtnMapEdit", 9000, 170, 3400, 760, 0xDDEBF7),
    (MAP, "btnMapSave", "Сохранить карту", "WmsMap.BtnMapSave", 12500, 170, 3600, 760, 0xC6E0B4),
    (MAP, "btnMapCancel", "Отменить правку", "WmsMap.BtnMapCancel", 16200, 170, 3600, 760, 0xF4CCCC),
    (MAP, "btnTopoCheck", "Проверить топологию", "WmsMap.BtnTopoCheck", 19900, 170, 4300, 760, 0xDDEBF7),
    (MAP, "btnMapFill", "Заполненность", "WmsMap.BtnMapFill", 24300, 170, 3400, 760, 0xFCE8B2),
    (MAP, "btnRename", "Переименовать ряд…", "WmsDaily.BtnRenamePlaces", 27800, 170, 4400, 760, 0xDDEBF7),
    (FIND, "btnShowOnMap", "Показать на карте", "WmsFind.BtnShowOnMap", None, None, 3900, 720, 0xDDEBF7),
    (ADJUST, "btnSplit", "Разделить ЕИ…", "WmsDaily.BtnSplitWindow", 150 + 8 * 2900, 120, 2750, 650, 0xDDEBF7),
    (ORDERS, "btnReceiveWin", "Принять…", "WmsDaily.BtnReceiveWindow", 150 + 5 * 3050, 120, 2900, 650, 0xC6E0B4),
    (ORDERS, "btnReceiveAll", "Принять всё по заказу", "WmsDaily.BtnReceiveAll", 150 + 5 * 3050, 120 + 760, 3900, 650, 0xC6E0B4),
    (FIND, "btnIssueFrom", "Выдать", "WmsDaily.BtnIssueFrom", None, None, 2400, 720, 0xC6E0B4),
    (FIND, "btnReturnHands", "Вернуть выделенное", "WmsDaily.BtnReturnHands", None, None, 3900, 720, 0xC6E0B4),
    # 1.0 (этап F): orders and suppliers
    (ORDERS, "btnRepeatOrder", "Повторить заказ", "WmsSupply.BtnRepeatOrder", 19400, 880, 3300, 650, 0xDDEBF7),
    (ORDERS, "btnCancelOrderAll", "Отменить весь заказ", "WmsSupply.BtnCancelOrderAll", 22800, 880, 3900, 650, 0xF4CCCC),
    (ORDERS, "btnSupply", "Снабжение", "WmsSupply.NavSupply", 150 + 7 * 3050, 120, 2900, 650, 0xDDEBF7),
    (STOCK, "btnSupReturn", "Вернуть поставщику…", "WmsSupply.BtnSupReturn", 21600, 120, 4300, 650, 0xDDEBF7),
    (ADJUST, "btnSupReturn", "Вернуть поставщику…", "WmsSupply.BtnSupReturn", 150 + 9 * 2900, 120, 4200, 650, 0xDDEBF7),
    (SUPPLIERS, "btnSuppliersFill", "Пополнить из заказов", "WmsSupply.BtnSuppliersFill", 18200, 120, 4600, 650, 0xDDEBF7),
    (GOODS, "btnGoodsRefresh", "Обновить", "WmsSupply.BtnGoodsRefresh", None, None, 2600, 650, 0xDDEBF7),
    (GOODS, "btnGoodsFill", "Пополнить из «Наличия»", "WmsSupply.BtnGoodsFill", None, None, 5000, 650, 0xDDEBF7),
    (CARS, "btnCarAccept", "Принять товар", "WmsSupply.BtnCarAccept", None, None, None, None, 0xC6E0B4),
    (ORDERS, "btnDiscAct", "Акт расхождения", "WmsPrint.BtnDiscrepancyAct", 26800, 880, 3600, 650, 0xDDEBF7),
    # 1.0 (этап G): control and documents
    (FIND, "btnHistory", "История изменений", "WmsAudit.BtnHistory", None, None, 4000, 720, 0xDDEBF7),
    (FIND, "btnCardPrint", "Печать карточки", "WmsPrint.BtnCardPrint", None, None, 3900, 720, 0xDDEBF7),
    (ISSUES, "btnIssueNote", "Накладная", "WmsPrint.BtnIssueNote", 150 + 5 * 2900, 880, 2750, 650, 0xDDEBF7),
    # «Корректировки»: the second row of buttons (its header made as high as that of «Выдачи»)
    (ADJUST, "btnRecount", "Пересчитать место…", "WmsPrint.BtnRecountZone", 150, 880, 3900, 650, 0xC6E0B4),
    (ADJUST, "btnInvList", "Опись…", "WmsPrint.BtnInventoryList", 150 + 4050, 880, 2750, 650, 0xDDEBF7),
    (ADJUST, "btnInvCompare", "Сличительная ведомость…", "WmsPrint.BtnInventoryCompare", 150 + 4050 + 2900, 880, 4900, 650, 0xDDEBF7),
    (ADJUST, "btnWriteOffAct", "Акт списания", "WmsPrint.BtnWriteOffAct", 150 + 4050 + 2900 + 5050, 880, 3000, 650, 0xDDEBF7),
    # 1.0 (этап I): the kits
    (ISSUES, "btnKitIssue", "Выдать комплект…", "WmsKit.BtnKitIssue", 150 + 6 * 2900, 880, 3600, 650, 0xC6E0B4),
    (ADJUST, "btnKitBuild", "Собрать комплект…", "WmsKit.BtnKitBuild", 150 + 4050 + 2900 + 5050 + 3150, 880, 3900, 650, 0xC6E0B4),
    # 1.0 (этап I): the label, the properties, the photo, the tools on the hands, the reserves («Наличие» and «Поиск» — a second row)
    (STOCK, "btnLabel", "Этикетка", "WmsItem.BtnLabel", 150, 880, 2600, 650, 0xC6E0B4),
    (STOCK, "btnProps", "Свойства ЕИ…", "WmsItem.BtnProps", 2900, 880, 3400, 650, 0xDDEBF7),
    (STOCK, "btnPhoto", "Фото", "WmsItem.BtnPhoto", 6450, 880, 2000, 650, 0xDDEBF7),
    (STOCK, "btnTools", "Где инструмент", "WmsItem.BtnTools", 8600, 880, 3600, 650, 0xDDEBF7),
    (STOCK, "btnReserves", "Резервы", "WmsItem.NavReserves", 12350, 880, 2600, 650, 0xDDEBF7),
    (ORDERS, "btnLabel", "Этикетка", "WmsItem.BtnLabel", 150 + 8 * 3050, 120, 2900, 650, 0xC6E0B4),
    (SPECIAL, "btnLabel", "Этикетка", "WmsItem.BtnLabel", 150 + 8 * 2900, 120, 2750, 650, 0xC6E0B4),
    (ISSUES, "btnTools", "Где инструмент", "WmsItem.BtnTools", 150 + 6 * 2900 + 3750, 880, 3600, 650, 0xDDEBF7),
    (FIND, "btnLabel", "Этикетка", "WmsItem.BtnLabel", None, None, 2400, 640, 0xC6E0B4),
    (FIND, "btnProps", "Свойства ЕИ…", "WmsItem.BtnProps", None, None, 3900, 640, 0xDDEBF7),
    (FIND, "btnPhoto", "Фото", "WmsItem.BtnPhoto", None, None, 1900, 640, 0xDDEBF7),
    (FIND, "btnTools", "Где инструмент", "WmsItem.BtnTools", None, None, 3900, 640, 0xDDEBF7),
    # 1.0 (С7): the search over the whole history (by default 12 months)
    (FIND, "btnFindAll", "Вся история", "WmsFind.BtnFindAll", None, None, 2800, 720, 0xDDEBF7),
]


def buttons_10(doc, sheets):
    fit_headers(sheets.getByName(ISSUES), ISSUE_HEADERS, 3100)
    fit_headers(sheets.getByName(ADJUST), ADJUST_HEADERS, 2700)
    fit_headers(sheets.getByName(STOCK), STOCK_HEADERS, 2700)
    for sheet, name, label, macro, x, y, w, h, back in BUTTONS_10:
        sh = sheets.getByName(sheet)
        if x is None and sheet == CARS:
            # 1.0 (З9): over L1:L2 of the panel, after «Исправить» / «Отменить визит»
            x = sh.getCellByPosition(11, 0).Position.X + 40
            y = sh.getCellByPosition(0, 0).Position.Y + 40
            w = sh.getCellByPosition(12, 0).Position.X - x - 40
            h = sh.getCellByPosition(0, 2).Position.Y - y - 40
        elif x is None and sheet == GOODS:
            # over the hint right of the table
            x = sh.getCellByPosition(len(GOODS_HEADERS) + 2, 0).Position.X + {"btnGoodsRefresh": 0, "btnGoodsFill": 2700}[name]
            y = 120
        elif x is None and name in ("btnLabel", "btnProps", "btnPhoto", "btnTools", "btnFindAll"):
            # «Поиск»: the second row, under «Выдать» … (right of the hint of row 5)
            y = sh.getCellByPosition(0, 4).Position.Y + 70
            x = sh.getCellByPosition(6, 3).Position.X + {"btnLabel": 13100, "btnProps": 15580, "btnPhoto": 19560, "btnTools": 21540, "btnFindAll": 25540}[name]
        elif x is None:
            # «Поиск»: the row of its buttons (build_find), after «Фасовка ЕИ»
            y = sh.getCellByPosition(0, 3).Position.Y + 90
            x = sh.getCellByPosition(6, 3).Position.X + {"btnIssueFrom": 13100, "btnReturnHands": 15580, "btnShowOnMap": 19560,
                                                         "btnHistory": 23540, "btnCardPrint": 27640}[name]
        if not has_button(sh, name):
            add_button(doc, sh, name, label, macro, x, y, w, h, back=back)
    find_layout_10(sheets)
    help_buttons_10(doc, sheets)


# 1.0 (П6): «?» — the section of «Справка» of the sheet, right of the first row of its buttons
HELP_MARK_SHEETS = (ORDERS, SPECIAL, CARS, ISSUES, RETURNS, ADJUST, STOCK, FIND, MAP, DASH, SUPPLY, REPORTS, SERVICE, KITS, RESERVES, GOODS,
                    SUPPLIERS, PLACES, MAIN)


def help_buttons_10(doc, sheets):
    """1.0 (П2, П3, П6, П7, П8): the buttons of «Справка» over its text and «?» on the working sheets and the pages"""
    hs = sheets.getByName(HELP)
    for r in range(HELP_TOP):
        hs.getRows().getByIndex(r).Height = 520
    x = 150
    for name, label, macro in HELP_BUTTONS:
        w = 3600 if len(label) > 11 else 2800
        if not has_button(hs, name):
            add_button(doc, hs, name, label, macro, x, 60, w, 900, bold=name == "btnWhatIf", back=0xC6E0B4 if name == "btnWhatIf" else 0xDDEBF7)
        x += w + 100
    for name in HELP_MARK_SHEETS:
        if not sheets.hasByName(name):
            continue
        sh = sheets.getByName(name)
        if has_button(sh, "btnHelpHere"):
            continue
        dp = sh.getDrawPage()
        right, top, hgt = 150, None, 650
        for i in range(dp.getCount()):
            shp = dp.getByIndex(i)
            if not shp.supportsService("com.sun.star.drawing.ControlShape"):
                continue
            p, z = shp.getPosition(), shp.getSize()
            if top is None or p.Y < top - 100:
                top, right, hgt = p.Y, p.X + z.Width, z.Height
            elif abs(p.Y - top) <= 100:
                right = max(right, p.X + z.Width)
        add_button(doc, sh, "btnHelpHere", "?", "WmsHelp.BtnHelpHere", right + 100, top if top is not None else 120, 700, hgt, bold=True, back=0xFFF2CC)


# 1.0: «Поиск» — its buttons in two rows that fit a screen of 1920 points (row 4: find and act, row 5: the card of the EI)
FIND_ROWS = (
    (3, 90, 720, (("btnFind", 2600), ("btnFindCard", 3300), ("btnFindGo", 3800), ("btnIssueFrom", 2400), ("btnReturnHands", 3900),
                  ("btnShowOnMap", 3900), ("btnFindAll", 2800))),
    (4, 80, 640, (("btnFindPack", 3000), ("btnHistory", 4000), ("btnCardPrint", 3900), ("btnLabel", 2400), ("btnProps", 3400),
                  ("btnPhoto", 1900), ("btnTools", 3600))),
)


def find_layout_10(sheets):
    """1.0: the buttons of «Поиск» placed in the two rows of FIND_ROWS (a new book and an upgraded one); the hint of row 5
    narrowed to B:F (two lines) so that the second row of buttons stands right of it"""
    sh = sheets.getByName(FIND)
    sh.getCellRangeByPosition(1, 4, 8, 4).merge(False)
    hint = sh.getCellRangeByPosition(1, 4, 5, 4)
    hint.merge(True)
    hint.IsTextWrapped = True
    hint.VertJustify = 1
    sh.getRows().getByIndex(4).Height = 900
    dp = sh.getDrawPage()
    shapes = {}
    for i in range(dp.getCount()):
        shp = dp.getByIndex(i)
        if shp.supportsService("com.sun.star.drawing.ControlShape"):
            shapes[shp.getControl().Name] = shp
    x0 = sh.getCellByPosition(6, 3).Position.X + 80
    for row, dy, h, items in FIND_ROWS:
        y = sh.getCellByPosition(0, row).Position.Y + dy
        x = x0
        for name, w in items:
            if name in shapes:
                shapes[name].setPosition(Point(x, y))
                shapes[name].setSize(Size(w, h))
            x += w + 80


def where_buttons(doc, sheets):
    """PRIME 0.8: «Куда положить?» on «Заказы» (the free place of the second row of its buttons) and on «Иной приход» (after its
    buttons); «Подобрать ЕИ» on «Выдачи» and «Корректировки» (after their buttons)"""
    add_button(doc, sheets.getByName(ORDERS), "btnWherePut", "Куда положить?", "WmsMap.BtnWherePut", 150 + 4 * 3050, 120 + 760, 2900, 650,
               back=0xC6E0B4)
    add_button(doc, sheets.getByName(SPECIAL), "btnWherePut", "Куда положить?", "WmsMap.BtnWherePut", 150 + 6 * 2900, 120, 2750, 650,
               back=0xC6E0B4)
    # «Подобрать ЕИ»: the EI by a name, an article or a place («Выдачи» after its four buttons, «Корректировки» after its own)
    add_button(doc, sheets.getByName(ISSUES), "btnPickEI", "Подобрать ЕИ", "WmsFind.BtnPickEI", 150 + 4 * 2900, 120, 2750, 650, back=0xC6E0B4)
    add_button(doc, sheets.getByName(ADJUST), "btnPickEI", "Подобрать ЕИ", "WmsFind.BtnPickEI", 150 + 5 * 2900, 120, 2750, 650, back=0xC6E0B4)
    # «Фасовка ЕИ»: how the EI of the row is handed out («в 1 шт — 4 кан», «1 шт = 6 м»)
    add_button(doc, sheets.getByName(ISSUES), "btnPackEI", "Фасовка ЕИ", "WmsPack.BtnPackEI", 150 + 5 * 2900, 120, 2750, 650, back=0xC6E0B4)
    add_button(doc, sheets.getByName(ADJUST), "btnPackEI", "Фасовка ЕИ", "WmsPack.BtnPackEI", 150 + 6 * 2900, 120, 2750, 650, back=0xC6E0B4)


def view_events(sheets):
    """the events of the view sheets (the same in a new book and in an upgraded one)"""
    bind_event(sheets.getByName(DASH).Events, "OnFocus", "WmsDash.OnDashFocus")
    bind_event(sheets.getByName(DASH).Events, "OnDoubleClick", "WmsDash.OnDashDoubleClick")
    bind_event(sheets.getByName(FIND).Events, "OnChange", "WmsFind.OnFindChange")
    bind_event(sheets.getByName(FIND).Events, "OnDoubleClick", "WmsFind.OnFindDoubleClick")
    bind_event(sheets.getByName(MAP).Events, "OnSelect", "WmsMap.OnMapSelect")
    bind_event(sheets.getByName(MAP).Events, "OnFocus", "WmsMap.OnMapFocus")
    bind_event(sheets.getByName(MAP).Events, "OnDoubleClick", "WmsMap.OnMapDoubleClick")


def build_views(doc, sheets):
    build_dash(doc, sheets.getByName(DASH))
    build_find(doc, sheets.getByName(FIND))
    build_map(doc, sheets.getByName(MAP))
    build_topo(sheets.getByName(TOPO))


def order_separators(doc, sh):
    """M6 §18: a noticeable line across A:AB above every order but the first — where the user's numbering of the positions
    (A) restarts at 1 on a row that is not a delivery row of «Ещё поступление» (AC holds its mark «OL<OrderLineID>»). A
    conditional format of the Calc engine: it follows inserted, posted, corrected, filtered rows and survives the save;
    no merged cells. The hidden column AC (locked, outside the autofilter and the snapshot) is created here too."""
    c = sh.getCellByPosition(ORDER_BLOCK_COL, 0)
    c.setString(ORDER_BLOCK_HEADER)
    c.CharWeight = 150
    sh.getCellRangeByPosition(ORDER_BLOCK_COL, 1, ORDER_BLOCK_COL, LAST_ROW).NumberFormat = number_format(doc, "@", ("ru", "RU"))
    sh.getColumns().getByIndex(ORDER_BLOCK_COL).IsVisible = False
    cell_style(doc, "WMS_OrderStart", TopBorder=border_line(88, 0x1F3864))
    add_cond(sh, sh.getCellRangeByPosition(0, 1, len(ORDER_HEADERS) - 1, LAST_ROW), 'AND(ROW()>2;TRIM($A2)="1";$AC2="")', "WMS_OrderStart", 0, 1)
    # the expected date Q of a position still expected has passed: an unobtrusive mark on Q — the status stays «Ожидается» /
    # «Частично получено», the posting and the receipts are not affected (D-089). A conditional format applies its first true
    # entry only, so Q gets its own one: the order line with the mark, the order line, the mark
    start = 'ROW()>2;TRIM($A2)="1";$AC2=""'
    late = 'ISNUMBER($Q2);$Q2<TODAY();OR($W2="Ожидается";$W2="Частично получено")'
    cell_style(doc, "WMS_LateDate", CellBackColor=0xFFEB9C, CharColor=0x9C5700)
    cell_style(doc, "WMS_OrderStartLate", TopBorder=border_line(88, 0x1F3864), CellBackColor=0xFFEB9C, CharColor=0x9C5700)
    q = sh.getCellRangeByPosition(16, 1, 16, LAST_ROW)
    cf = q.ConditionalFormat
    cf.clear()
    q.ConditionalFormat = cf
    add_cond(sh, q, f"AND({start};{late})", "WMS_OrderStartLate", 16, 1)
    add_cond(sh, q, f"AND({start})", "WMS_OrderStart", 16, 1)
    add_cond(sh, q, f"AND({late})", "WMS_LateDate", 16, 1)


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
    order_separators(doc, sh)
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
    ] + idx_car_formulas()      # WmsConfig IX_CO_MATCH .. IX_CS_MATCH (M6)
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


def add_nav(doc, sh):
    """«Главная»: the ways to the working sheets (M6 PRIME §11) — a strip of buttons in row 2"""
    sh.getRows().getByIndex(1).Height = 1000
    y = sh.getCellByPosition(0, 1).Position.Y
    nav = (("navOrders", ORDERS, "WmsUi.NavOrders"), ("navSpecial", SPECIAL, "WmsUi.NavSpecial"), ("navCars", CARS, "WmsUi.NavCars"),
           ("navIssues", ISSUES, "WmsUi.NavIssues"), ("navReturns", RETURNS, "WmsUi.NavReturns"), ("navAdjust", ADJUST, "WmsUi.NavAdjust"),
           ("navStock", STOCK, "WmsUi.NavStock"))
    w = 3200
    for i, (name, label, macro) in enumerate(nav):
        add_button(doc, sh, name, label, macro, MAIN_LEFT + 100 + i * (w + 60), y + 80, w, 840, bold=True, back=0xDDEBF7)


def main_nav_views(doc, sh):
    """PRIME 0.8: «Главная» leads to the new pages too — the second strip, under the ways to the working sheets"""
    sh.getRows().getByIndex(2).Height = 1000
    y = sh.getCellByPosition(0, 2).Position.Y
    # 1.0: «Снабжение» (этап F) and «Отчёты» (этап G) among them
    # 1.0 (этап J): «Обслуживание» — seven buttons, as wide as those of the strip above
    w = 3200
    for i, (name, label, macro) in enumerate((("navDash", "Дашборд", "WmsUi.NavDash"), ("navFind", "Поиск", "WmsUi.NavFind"),
                                              ("navMap", "Карта склада", "WmsUi.NavMap"), ("navSupply", "Снабжение", "WmsSupply.NavSupply"),
                                              ("navReports", "Отчёты", "WmsLedger.NavReports"), ("navService", "Обслуживание", "WmsSafe.NavService"),
                                              ("navHelp", "Справка", "WmsDaily.NavHelp"))):
        add_button(doc, sh, name, label, macro, MAIN_LEFT + 100 + i * (w + 60), y + 80, w, 840, bold=True, back=0xC6E0B4)


def version_text():
    v = core_versions()
    return f"версия {v['WMS_PRODUCT_VERSION']} · ядро {v['WMS_CORE_VERSION']} · схема {v['WMS_SYS_SCHEMA']}"


MAIN_HELP_ROW = 15


def main_help_text():
    """1.0 (В4): six short «Как…» instead of a wall of text; everything else is in «Справка»"""
    return "\n".join([
        "КАК РАБОТАТЬ — коротко (подробно — лист «Справка»)",
        "1. Начать день: «Проверка перед работой» — всё в порядке → работайте.",
        "2. Принять товар: «Заказы» → строка заказа → «Принять…» (или F, N, U и «Провести приход»). Не по заказу — «Иной приход».",
        "3. Выдать: «Выдачи» → «Выдать…» (Ctrl+Shift+W) или ЕИ, количество, дата, кому → «Провести» (Ctrl+Shift+P).",
        "4. Принять возврат: «Поиск» → имя человека → выделить строки → «Вернуть выделенное» (или лист «Возврат»).",
        "5. Исправить ошибку: проведённая строка → «Исправить» или «Удалить»; непроведённая — «Очистить». Не проводится — «Почему?».",
        "6. Найти: «Поиск» — номер ЕИ, получатель или любое слово (Ctrl+Shift+F); где лежит и куда положить — «Карта склада».",
        "Если вверху «ЗАБЛОКИРОВАНО» — ничего не проводится: прочтите «Что сделать» и нажмите кнопку из красной рамки."])


# 1.0 (В4): «Главная» as a menu — the group of each row of buttons (row, caption, colour of the caption)
MAIN_GROUPS = [(1, "Каждый день", 0x1F3864), (3, "Только если WMS просит", 0xB3261E), (4, "Инструменты", 0x1F3864)]
MAIN_LEFT = 5200       # the width of column A (the captions); the buttons are in column B


MAIN_ROWS = 7          # WmsUi.MR_STATE: the lines of the state start here (rows 7 … 13)


def build_main(doc, sh):
    """«Главная» (PRIME 0.8 layout): the title, the ways to the sheets and pages, the actions of the system — all above the lines
    of the state, which grow with their text (nothing below them is a button: they never cover one)"""
    sh.getColumns().getByIndex(0).Width = MAIN_LEFT
    sh.getColumns().getByIndex(1).Width = 23000
    t = sh.getCellByPosition(0, 0)
    t.setString("Состояние WMS")
    t.CharWeight = 150
    t.CharHeight = 16
    sh.getRows().getByIndex(0).Height = 1000
    sh.getCellByPosition(1, 0).setString(version_text())
    sh.getCellByPosition(1, 0).VertJustify = 2      # centre, next to the title
    sh.getCellByPosition(1, 0).HoriJustify = 3
    sh.getCellByPosition(1, 0).CharColor = MUTED
    add_nav(doc, sh)
    main_nav_views(doc, sh)
    for r in (3, 4, 5):
        sh.getRows().getByIndex(r).Height = 850
    sh.getRows().getByIndex(6).Height = 300
    buttons = (("btnRecover", "Восстановить", "WmsUi.BtnRecover"), ("btnAbandon", "Отложить хвост журнала", "WmsUi.BtnAbandon"),
               ("btnUnlock", "Снять блокировку WMS", "WmsUi.BtnUnlock"), ("btnRegister", "Сделать рабочим файлом", "WmsUi.BtnRegister"),
               ("btnPreWork", "Проверка перед работой", "WmsStatus.BtnPreWorkCheck"), ("btnSelfCheck", "Самопроверка", "WmsUi.BtnSelfCheck"),
               ("btnBackup", "Резервная копия", "WmsUi.BtnBackup"), ("btnStatus", "Состояние системы", "WmsStatus.BtnSystemStatus"),
               ("btnExport", "Экспорт для инструментов", "WmsExport.BtnExport"), ("btnLoadBatch", "Загрузить пакет", "WmsExport.BtnLoadBatch"))
    for i, (name, label, macro) in enumerate(buttons):
        y = sh.getCellByPosition(0, 3 + i // 4).Position.Y
        add_button(doc, sh, name, label, macro, MAIN_LEFT + 100 + (i % 4) * 5700, y + 50, 5640, 750,
                   back=0xF4CCCC if i < 4 else (0xC6E0B4 if name == "btnPreWork" else 0xDDEBF7))
    for r, cap, color in MAIN_GROUPS:
        c = sh.getCellByPosition(0, r)
        c.setString(cap)
        c.CharWeight = 150
        c.CharColor = color
        c.VertJustify = 2
        c.IsTextWrapped = True
    # the buttons that only WMS asks for: a red frame around their row
    red = border_line(70, 0xB3261E)
    frame = sh.getCellRangeByPosition(0, 3, 1, 3)
    frame.CellBackColor = 0xFDE2DD
    for side in ("TopBorder", "BottomBorder", "LeftBorder", "RightBorder"):
        setattr(frame, side, red)
    captions = ["Состояние", "Причина", "Что сделать", "Непроведённые строки", "Заметки запуска", "Последнее действие", "Обновлено"]
    for i, cap in enumerate(captions):
        c = sh.getCellByPosition(0, MAIN_ROWS + i)
        c.setString(cap)
        c.CharWeight = 150
    sh.getCellByPosition(1, MAIN_ROWS).setString("WMS ещё не запускалась в этой книге — откройте её с включёнными макросами")
    vals = sh.getCellRangeByPosition(0, MAIN_ROWS, 1, MAIN_ROWS + 6)
    vals.IsTextWrapped = True
    vals.VertJustify = 1           # top
    help_row = MAIN_HELP_ROW
    sh.getCellByPosition(0, help_row).setString(main_help_text())
    sh.getCellRangeByPosition(0, help_row, 1, help_row).merge(True)
    sh.getCellByPosition(0, help_row).IsTextWrapped = True
    sh.getCellByPosition(0, help_row).VertJustify = 1
    sh.getRows().getByIndex(help_row).Height = 4600
    t = sh.getCellByPosition(0, STATUS_ROW)
    t.setString("СОСТОЯНИЕ СИСТЕМЫ")
    t.CharWeight = 150
    sh.getCellByPosition(1, STATUS_ROW).setString("нажмите «Состояние системы» или «Проверка перед работой»")
    sh.getCellRangeByPosition(0, STATUS_ROW + 1, 1, STATUS_ROW + 31).IsTextWrapped = True


def rebuild_main(doc, sh):
    """an upgraded book gets «Главная» of this release (it holds no data: WMS writes its state at every start)"""
    dp = sh.getDrawPage()
    while dp.getCount():
        dp.remove(dp.getByIndex(0))
    sh.getRows().removeByIndex(0, 200)
    sh.getColumns().removeByIndex(0, 10)
    build_main(doc, sh)


def md_plain(s):
    """a line of the guides without the Markdown marks: **bold**, `code`, links"""
    s = s.replace("**", "").replace("`", "")
    return re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", s)


HELP_TOP = 2                       # 1.0 (П2, П3, П7, П8): the buttons of «Справка» stand in rows 1–2, the text begins under them
HELP_BUTTONS = [("btnWhatsNew", "Что нового", "WmsHelp.BtnWhatsNew"), ("btnTour", "Первый день", "WmsHelp.BtnTour"),
                ("btnWhatIf", "Что делать, если…", "WmsHelp.BtnWhatIf"), ("btnCheat", "Шпаргалка A4", "WmsHelp.BtnCheatSheet"),
                ("btnHelpMain", "Главная", "WmsUi.NavMain")]


def build_help(sh):
    """«Справка»: docs/OPERATOR_GUIDE.md, RELEASE_NOTES.md and BACKUP_RECOVERY.md as text (headings bold, a table row per line)"""
    sh.getColumns().getByIndex(0).Width = 26000
    r = HELP_TOP
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


def apply_views(doc):
    """the view of every working sheet (saved in settings.xml, only by a document with a view): the header rows frozen —
    «Приход авто» freezes its whole panel with the header (M6 §1: the buttons stay in view, no moving shapes); the cursor
    where the work starts; the book opens on «Дашборд» (PRIME 0.8)"""
    ctl = doc.getCurrentController()
    sheets = doc.Sheets
    # 1.0 (В5): with the header the columns that say whose row it is stay in view when scrolled right (№ / EI and the name)
    for name in (ORDERS, SPECIAL, ISSUES, RETURNS, ADJUST, STOCK, RCPT):
        ctl.setActiveSheet(sheets.getByName(name))
        ctl.freezeAtPosition(FREEZE_COLS_10.get(name, 0), 1)
    ctl.setActiveSheet(sheets.getByName(LEGREP))
    ctl.freezeAtPosition(1, LEGREP_HEAD + 1)
    ctl.setActiveSheet(sheets.getByName(CARS))
    ctl.freezeAtPosition(0, CAR_HEAD_ROW + 1)
    ctl.select(sheets.getByName(CARS).getCellByPosition(2, 1))
    ctl.setActiveSheet(sheets.getByName(ISSUES))
    ctl.select(sheets.getByName(ISSUES).getCellByPosition(11, 1))
    ctl.setActiveSheet(sheets.getByName(MAP))
    ctl.freezeAtPosition(0, 7)                      # WmsMap.MP_TOP: the header stays in view over the map
    # 1.0.1: the grid of the cells on — it is one setting of the whole view (not of a sheet): turned off for the pages of 1.0 it
    # was off on every sheet and the cells of the working sheets ran together
    ctl.ShowGrid = True
    ctl.setActiveSheet(sheets.getByName(FIND))
    ctl.select(sheets.getByName(FIND).getCellByPosition(2, 3))
    ctl.setActiveSheet(sheets.getByName(DASH))
    ctl.select(sheets.getByName(DASH).getCellByPosition(0, 0))
    ctl.ShowPageBreaks = False


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
    for i, name in enumerate((ORDERS, SPECIAL, CARS, ISSUES, RETURNS, ADJUST, STOCK, FIND, MAP, LEGREP, RCPT, TOPO, HELP, ORD, RCV, SPR, ART,
                              RET, ISS, ADJ, CAR, LEG, PACK, IDX, "_SYS"), start=1):
        sheets.insertNewByName(name, i)
    sheets.insertNewByName(DASH, 0)
    sys_sh = sheets.getByName("_SYS")
    values = {
        "SCHEMA": SCHEMA, "INSTANCE_ID": instance_id or "%016X" % random.getrandbits(64), "MODE": "TEST" if test else "PROD",
        "CORE_VERSION": "", "LAST_SEQ": 0, "NEXT_EI": 1, "NEXT_NO": 1, "NEXT_RET": 1, "JOURNAL_POS": "", "REGISTERED_URL": "",
        "TX_STATE": "NONE", "TX_SEQ": 0, "TX_TYPE": "", "TX_TIME": "", "TX_BEFORE_IMAGE": "", "SAVE_STAMP": 0, "SAVE_SEQ": 0,
        "MAX_QTY": 100000,
        "KEY_SHEETS": ISSUE_KEY_SPEC + ";" + ORDER_KEY_SPEC + ";" + RETURN_KEY_SPEC + ";" + SPECIAL_KEY_SPEC + ";" + ADJUST_KEY_SPEC
                      + (";" + TEST_KEY_SPEC if test else ""),
        "NEXT_SPL": 1, "NEXT_OFF": 1, "NEXT_PROD": 1, "NEXT_DET": 1, "NEXT_OLD": 1, "NEXT_OTH": 1, "NEXT_ADJ": 1, "NEXT_CAR": 1,
        "UPGRADE_FROM": "",
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
    build_cars(doc, sheets.getByName(CARS))
    build_car_service(doc, sheets.getByName(CAR))
    build_stock(doc, sheets.getByName(STOCK), registry)
    build_leg(doc, sheets.getByName(LEG))
    build_pack(doc, sheets.getByName(PACK))
    build_legrep(doc, sheets.getByName(LEGREP))
    build_recipients(sheets.getByName(RCPT), TEST_RECIPIENTS if test else [])
    build_help(sheets.getByName(HELP))
    build_views(doc, sheets)
    status_colors(doc, sheets)
    where_buttons(doc, sheets)
    layout_10(doc, sheets)
    supply_10(doc, sheets)
    print_10(doc, sheets)
    ledger_10(doc, sheets)
    kits_10(doc, sheets)
    items_10(doc, sheets)
    service_10(doc, sheets)
    buttons_10(doc, sheets)
    colors_10(doc, sheets)
    hints_10(doc, sheets)
    stock_info_10(doc, sheets)
    stock_view_10(doc, sheets)
    goods_102(doc, sheets)
    docs_102(doc, sheets)
    categories_102(doc, sheets)
    buttons_102(doc, sheets)
    views_110(doc, sheets)
    main_buttons_110(doc, sheets)
    hotkeys_10(doc)
    build_service(doc, sheets)
    order = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    autofilter(doc, order.index(ORDERS), "WMS_ORDERS", len(ORDER_HEADERS))
    autofilter(doc, order.index(SPECIAL), "WMS_SPECIAL", len(SPECIAL_HEADERS))
    autofilter(doc, order.index(ISSUES), "WMS_ISSUES", len(ISSUE_HEADERS))
    autofilter(doc, order.index(RETURNS), "WMS_RETURNS", len(RETURN_HEADERS))
    autofilter(doc, order.index(ADJUST), "WMS_ADJUST", len(ADJUST_HEADERS))
    autofilter(doc, order.index(CARS), "WMS_CARS", len(CAR_HEADERS), start_row=CAR_HEAD_ROW)
    autofilter(doc, order.index(STOCK), "WMS_STOCK", STOCK_INFO_COL + len(STOCK_INFO_HEADERS))
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
    view_events(sheets)
    for name in (ORD, RCV, SPR, ART, RET, ISS, ADJ, CAR, LEG, PACK, IDX, "_SYS"):
        sheets.getByName(name).IsVisible = False
    for name in (MAIN, ORDERS, SPECIAL, CARS, ISSUES, RETURNS, ADJUST, STOCK, LEGREP, HELP, ORD, RCV, SPR, ART, RET, ISS, ADJ, CAR, LEG, PACK, IDX,
                 "_SYS", DASH, FIND, MAP, SUPPLY, PRINT, CLOSE, REPORTS, PROP, SERVICE):
        sheets.getByName(name).protect(PWD)
    doc.protect(PWD)
    if os.path.exists(out_path):
        os.remove(out_path)
    doc.storeAsURL(uno.systemPathToFileUrl(out_path), props(FilterName="calc8"))
    doc.close(True)
    patch_content(out_path, lambda x: allow_insert_rows(allow_insert_rows(allow_insert_rows(allow_insert_rows(allow_insert_rows(
        x, ISSUES), ORDERS), RETURNS), SPECIAL), ADJUST))
    # the saved file must carry the save stamp WMS expects (EditingCycles after this save), otherwise the first
    # start would report "saved without WMS"; fix it up with macros disabled so no event interferes. The document is
    # loaded with a view (not hidden): only a view saves its settings — the frozen rows and the cursor of every sheet
    doc = office.load(out_path, macros=0, hidden=False)
    sys_sh = doc.Sheets.getByName("_SYS")
    ec = doc.getDocumentProperties().EditingCycles
    sys_sh.getCellByPosition(1, SK_SAVE_STAMP).setValue(ec + 1)
    apply_views(doc)
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
        shutil.rmtree(prof, ignore_errors=True)


if __name__ == "__main__":
    main()
