' WmsConfig — constants, _SYS layout, settings, LibreOffice settings management (AutoInput).
' MASTER SPEC v0.3: §3 sources as text, §10/§29 AutoInput, §16 protection, §20 journal, §22 lock, §24 chunks, §26 backups,
' §2 and §14 fixed sheets «Выдачи», «Заказы», «Возврат», «Иной приход» and their service structures.
Option Explicit

Public Const WMS_CORE_VERSION = "0.7.1-legacy"
' the product (the release of the whole WMS book and its tools; the snapshot for the tools names it)
Public Const WMS_PRODUCT_VERSION = "0.7.1"
' Phase 5 appended the counters of the special receipts (rows 19..24), Final Core the counter of the corrections (row 25),
' M6 the counter of the vehicle visits (row 26); a book of an earlier core has WMS-SYS-1 (Phase 1–4), WMS-SYS-2 (Phase 5)
' or WMS-SYS-3 (Final Core, release 0.6.0 — tools/upgrade.py brings it to this schema with its data)
Public Const WMS_SYS_SCHEMA = "WMS-SYS-4"
Public Const WMS_SYS_SCHEMA_OLD = "WMS-SYS-1"
Public Const WMS_SYS_SCHEMA_P5 = "WMS-SYS-2"
Public Const WMS_SYS_SCHEMA_FC = "WMS-SYS-3"

Public Const SYS_SHEET = "_SYS"
Public Const JOURNAL_DIR = "WMS_Journal"
Public Const BACKUP_DIR = "WMS_Backups"
Public Const LOCK_FILE = "wms.lock"
Public Const JOURNAL_PREFIX = "WMS_journal_"
Public Const JOURNAL_EXT = ".csv"
Public Const AUTOINPUT_STATE_FILE = "wms_autoinput.state"

' Sheet protection guards against accidents, not against a determined user (spec §16: the password is not used in daily work).
Public Const PROTECT_PWD = "wms"

' _SYS: column A = key, column B = value; row index (0-based) = SK_* constant. Layout is checked at every start.
Public Const SK_SCHEMA = 0
Public Const SK_INSTANCE = 1
Public Const SK_MODE = 2
Public Const SK_CORE_VERSION = 3
Public Const SK_LAST_SEQ = 4
Public Const SK_NEXT_EI = 5
Public Const SK_NEXT_NO = 6
Public Const SK_NEXT_RET = 7
Public Const SK_JPOS = 8
Public Const SK_REG_URL = 9
Public Const SK_TX_STATE = 10
Public Const SK_TX_SEQ = 11
Public Const SK_TX_TYPE = 12
Public Const SK_TX_TIME = 13
Public Const SK_TX_BI = 14
Public Const SK_SAVE_STAMP = 15
Public Const SK_SAVE_SEQ = 16
Public Const SK_MAX_QTY = 17
Public Const SK_KEY_SHEETS = 18
' Phase 5 (WMS-SYS-2): the line № of «Иной приход» and the numbers of the receipt events per source
Public Const SK_NEXT_SPL = 19
Public Const SK_NEXT_OFF = 20
Public Const SK_NEXT_PROD = 21
Public Const SK_NEXT_DET = 22
Public Const SK_NEXT_OLD = 23
Public Const SK_NEXT_OTH = 24
'' Final Core (WMS-SYS-3): the № of the corrections (MOVE / WRITE_OFF / INV_ADJ)
Public Const SK_NEXT_ADJ = 25
' M6 (WMS-SYS-4): the № of the vehicle visits of «Приход авто»
Public Const SK_NEXT_CAR = 26
Public Const SYS_ROWS = 27

' transaction marker states
Public Const TX_NONE = "NONE"
Public Const TX_STARTED = "STARTED"
Public Const TX_COMMITTED = "COMMITTED"

' settings
Public Const BACKUP_KEEP_DAILY = 14
Public Const BACKUP_KEEP_MANUAL = 10
Public Const BACKUP_KEEP_UPGRADE = 5
Public Const BACKUP_KEEP_MIGRATION = 5
Public Const CHECK_CHUNK_ROWS = 20000
Public Const JPOS_VERIFY_WINDOW = 4096
Public Const QTY_MAX_DECIMALS = 3
Public Const QTY_MAX_INT_DIGITS = 9

' ---------------------------------------------------------------- Phase 2: issues by EI (spec §2, §4, §5, §9, §10, §16)
Public Const SH_MAIN = "Главная"
Public Const SH_ISSUES = "Выдачи"
Public Const SH_STOCK = "Наличие"
Public Const SH_RCPT = "Получатели"

' «Выдачи» A:R — a fixed user interface (spec §2): 0-based column indices
Public Const IC_NO = 0
Public Const IC_DOC = 1
Public Const IC_NAME = 2
Public Const IC_ART = 3
Public Const IC_QTY = 4
Public Const IC_PCT = 5
Public Const IC_UNIT = 6
Public Const IC_DATE = 7
Public Const IC_WHO = 8
Public Const IC_PLACE = 9
Public Const IC_CAT = 10
Public Const IC_EI = 11
Public Const IC_RET = 12
Public Const IC_RETPCT = 13
Public Const IC_NOTE = 14
Public Const IC_BEFORE = 15
Public Const IC_AFTER = 16
Public Const IC_CTL = 17
Public Const IC_LAST = 17

' cell protection of a «Выдачи» row, one character per column A..R ("1" = locked). Unposted: the inputs B E F H I L and
' the informational M N O are open, everything WMS fills is locked. Posted: only B M N O stay open (D-013).
Public Const ISSUE_LOCKS_OPEN = "101100100110000111"
Public Const ISSUE_LOCKS_POSTED = "101111111111000111"

' «Контроль» (R) of a row that WMS posted, corrected or cancelled; a copy is marked КОПИЯ (spec §15)
Public Const ST_POSTED = "Проведено"
Public Const ST_FIXED = "Проведено (исправлено)"
Public Const ST_DELETED = "Удалено (сторно)"

' «Наличие» — the permanent EI registry, one row per EI, row index = EI number (spec §5, v0.1 §24/§37)
Public Const SC_EI = 0
Public Const SC_NAME = 1
Public Const SC_ART = 2
Public Const SC_UNIT = 3
Public Const SC_QTY = 4
Public Const SC_PLACE = 5
Public Const SC_CAT = 6
Public Const SC_STATE = 7
Public Const SC_SRC = 8
' Phase 5: the kind of source of the EI (Поставщик / Офис / Производство / Детали / Старый склад / Иной приход) — for
' filters and analytics; I «Источник» keeps the document of origin («Заказ <№>», «Офис OFF-00000001», …)
Public Const SC_STYPE = 9
Public Const SC_LAST = 9

Public Const EI_PREFIX = "ЕИ-"
Public Const EI_DIGITS = 8
Public Const MAX_SHEET_ROW = 1048575

' a change of more rows at once (paste, fill) gets its preview only for this many rows; the rest is cleared (spec §13)
Public Const PREVIEW_MAX_ROWS = 500
' «Главная»: at most this many separate blocks of rows without № are examined for the unposted count
Public Const UNPOSTED_MAX_BLOCKS = 20000

' ---------------------------------------------------------------- Phase 3: orders and ordinary receipts (spec §2, §4–§6, §14–§16)
Public Const SH_ORDERS = "Заказы"
Public Const SH_ORD = "_ORD"
Public Const SH_RCV = "_RCV"
Public Const SH_IDX = "_IDX"

' «Заказы» A:AB — a fixed user interface (spec §2): 0-based column indices
Public Const OC_ORDER = 0
Public Const OC_NAME = 1
Public Const OC_DOC = 2
Public Const OC_INVOICE = 3
Public Const OC_ART = 4
Public Const OC_FACT = 5
Public Const OC_DOCQTY = 6
Public Const OC_ORDQTY = 7
Public Const OC_UNIT = 8
Public Const OC_PRICE = 9
Public Const OC_SUM = 10
Public Const OC_SUPPLIER = 11
Public Const OC_SELLER = 12
Public Const OC_RDATE = 13
Public Const OC_DDATE = 14
Public Const OC_ODATE = 15
Public Const OC_EDATE = 16
Public Const OC_BUYER = 17
Public Const OC_CAT = 18
Public Const OC_ASSIGNED = 19
Public Const OC_PLACE = 20
Public Const OC_EI = 21
Public Const OC_STATUS = 22
Public Const OC_STOCK = 23
Public Const OC_CTL = 24
Public Const OC_NOTE = 25
Public Const OC_DAYS = 26
Public Const OC_DUP = 27
Public Const OC_LAST = 27
' AC — outside the user interface A:AB (hidden, locked, not in the snapshot): the mark «OL<OrderLineID>» of a row created by
' «Ещё поступление». The visual blocks of the orders (M6 §18) start where the user's numbering of the positions (A)
' restarts at 1 on a row without this mark; a delivery row stays in the block of its order.
Public Const OC_BLOCK = 28
Public Const ORDER_BLOCK_HEADER = "Блок (служебная)"

' cell protection of a «Заказы» row, one character per column A..AB ("1" = locked). An order row that has no receipt yet:
' everything the user enters is open, what WMS fills (V W X Y AB) is locked. A row with a receipt or a cancelled position:
' everything that describes the movement, the position or the status is locked; K M Q R T Z AA stay open.
Public Const ORDER_LOCKS_OPEN = "0000000000000000000001111001"
Public Const ORDER_LOCKS_POSTED = "1111111111010111001011111001"

' W «Статус»: the business status of an order position (on its source row only, spec §6, v0.1 §15)
Public Const OS_WAITING = "Ожидается"
Public Const OS_OVERDUE = "Просрочено"
Public Const OS_PARTIAL = "Частично получено"
Public Const OS_RECEIVED = "Получено"
Public Const OS_NODOCS = "Получено без документов"
Public Const OS_CANCELLED = "Отменено"
Public Const OS_REST_CANCELLED = "Частично получено / остаток отменён"
Public Const OS_PARTIAL_OVERDUE = "Частично получено / просрочено"
' W of an additional receipt row (technical status: the position status is shown on its source row)
Public Const OS_ADD = "Дополнительное поступление"
Public Const OS_ADD_STORNO = "Поступление удалено (сторно)"

' «Наличие» H «Состояние» of an EI created by a receipt
Public Const EI_ST_ACTIVE = "Активен"
Public Const EI_ST_STORNO = "Приход удалён (сторно)"
' Phase 5: an EI of «Иной приход» that is stock already but not identified yet
Public Const EI_ST_REVIEW = "Требует разбора"
' «Наличие» J «Тип источника» of an EI created by an ordinary receipt of «Заказы»
Public Const STYPE_SUPPLIER = "Поставщик"

' _ORD — order positions (OrderLineID = row index, dense), created by the first operation of a position (spec §14, D-006)
Public Const OD_ID = 0
Public Const OD_ROW = 1
Public Const OD_KEY = 2
Public Const OD_FP = 3
Public Const OD_ORD = 4
Public Const OD_RCV = 5
Public Const OD_CNT = 6
Public Const OD_NODOC = 7
Public Const OD_CANCEL = 8
Public Const OD_LAST = 8
' _ORD!L1 holds NEXT_OL, the next OrderLineID
Public Const OD_NEXT_COL = 11
Public Const OD_CANCEL_ORDER = "ORDER"
Public Const OD_CANCEL_REST = "REST"

' _RCV — receipts, one row per EI created by a receipt (row index = EI number, dense like «Наличие»)
Public Const RV_EI = 0
Public Const RV_OL = 1
Public Const RV_ROW = 2
Public Const RV_DUP = 3
Public Const RV_STATE = 4
Public Const RV_KIND = 5
Public Const RV_QTY = 6
Public Const RV_NODOC = 7
Public Const RV_DUPKEY = 8
Public Const RV_LAST = 8
Public Const RV_LIVE = "LIVE"
Public Const RV_STORNO = "STORNO"
Public Const RV_SRC = "SRC"
Public Const RV_ADD = "ADD"

' _IDX — lookup cells: the criterion is written into column B, formulas of the Calc engine answer (spec §25: MATCH/COUNTIF,
' not a Basic loop). Rows (0-based) of column B:
Public Const IX_KEY = 0
Public Const IX_START = 1
Public Const IX_NONCE = 2
Public Const IX_ECHO = 3
Public Const IX_V_MATCH = 4
Public Const IX_V_COUNT = 5
Public Const IX_DUP_MATCH = 6
Public Const IX_DUP_COUNT = 7
Public Const IX_A_MATCH = 8
Public Const IX_W_MATCH = 9
' the scratch cell (unlocked) for one-shot array formulas: rows of «Заказы» whose date-dependent status must be checked,
' issues of one EI for «Найти выдачу»
Public Const IX_LIST = 10
' Phase 4: «Выдачи».A (issue №) MATCH, «Возврат».A (return №) MATCH and COUNTIF
Public Const IX_I_MATCH = 11
Public Const IX_RA_MATCH = 12
Public Const IX_RA_COUNT = 13
' Phase 5: «Иной приход».A (line №) MATCH and COUNTIF, _ART.A (article key of the parts) MATCH and COUNTIF, _SPR.J
' (antidubl hash of the special receipts) MATCH and COUNTIF
Public Const IX_XA_MATCH = 14
Public Const IX_XA_COUNT = 15
Public Const IX_ART_MATCH = 16
Public Const IX_ART_COUNT = 17
Public Const IX_XD_MATCH = 18
Public Const IX_XD_COUNT = 19
'' Final Core: «Корректировки».A (the № of a correction)
Public Const IX_AA_MATCH = 20
Public Const IX_AA_COUNT = 21
' M6: _CAR.I (the key of the vehicle of an open visit) MATCH and COUNTIF, _CAR.B (the key of every visit) COUNTIF, the
' dictionaries of _CAR: vehicles (O) and suppliers (Q) MATCH
Public Const IX_CO_MATCH = 22
Public Const IX_CO_COUNT = 23
Public Const IX_CK_COUNT = 24
Public Const IX_CP_MATCH = 25
Public Const IX_CS_MATCH = 26
Public Const IX_LAST = 26

' ---------------------------------------------------------------- Phase 4: returns of issued goods (spec §11, §14, D-004, D-011)
Public Const SH_RETURNS = "Возврат"
Public Const SH_RET = "_RET"
Public Const SH_ISS = "_ISS"

' «Возврат» A:O (spec §11 lists the fields, the task of Phase 4 fixes the columns): 0-based column indices
Public Const RC_NO = 0
Public Const RC_ISSUE = 1
Public Const RC_EI = 2
Public Const RC_NAME = 3
Public Const RC_ART = 4
Public Const RC_QTY = 5
Public Const RC_UNIT = 6
Public Const RC_DATE = 7
Public Const RC_WHO = 8
Public Const RC_PLACE = 9
Public Const RC_CAT = 10
Public Const RC_BEFORE = 11
Public Const RC_AFTER = 12
Public Const RC_CTL = 13
Public Const RC_NOTE = 14
Public Const RC_LAST = 14

' cell protection of a «Возврат» row, one character per column A..O ("1" = locked). Unposted: the inputs B C F H I J and
' the comment O are open, everything WMS fills (A D E G K L M N) is locked. Posted or cancelled: only O stays open.
Public Const RETURN_LOCKS_OPEN = "100110100011110"
Public Const RETURN_LOCKS_POSTED = "111111111111110"

' _RET — returns, one row per return № (row index = return №, dense like «Наличие»)
Public Const RT_NO = 0
Public Const RT_ISSUE = 1
Public Const RT_EI = 2
Public Const RT_QTY = 3
Public Const RT_STATE = 4
Public Const RT_ROW = 5
Public Const RT_LAST = 5

' _ISS — the returns of one issue (row index = issue №, dense; a row appears with the first return of the issue):
' № выдачи, row hint of the issue on «Выдачи», returned (sum of the live returns), number of live returns
Public Const IS_NO = 0
Public Const IS_ROW = 1
Public Const IS_RET = 2
Public Const IS_CNT = 3
Public Const IS_LAST = 3
' «Найти выдачу»: at most this many of the latest returnable issues of one EI are listed
Public Const ISSUE_LIST_MAX = 100

' ---------------------------------------------------------------- Phase 5: special receipts (spec §4, §7, §8, v0.1 §10–§12, §23)
' One user sheet for every receipt that is not an order of «Заказы»: Офис, Производство, Детали, Старый склад, Иной.
Public Const SH_SPECIAL = "Иной приход"
Public Const SH_SPR = "_SPR"
Public Const SH_ART = "_ART"

' «Иной приход» A:S: 0-based column indices
Public Const XC_NO = 0
Public Const XC_TYPE = 1
Public Const XC_EVENT = 2
Public Const XC_NAME = 3
Public Const XC_ART = 4
Public Const XC_QTY = 5
Public Const XC_UNIT = 6
Public Const XC_DATE = 7
Public Const XC_PLACE = 8
Public Const XC_CAT = 9
Public Const XC_WHO = 10
Public Const XC_DOC = 11
Public Const XC_MARK = 12
Public Const XC_EI = 13
Public Const XC_BEFORE = 14
Public Const XC_AFTER = 15
Public Const XC_STATUS = 16
Public Const XC_CTL = 17
Public Const XC_NOTE = 18
Public Const XC_LAST = 18

' cell protection of an «Иной приход» row, one character per column A..S ("1" = locked). Unposted: the inputs B..M and
' the comment S are open, what WMS fills (A N O P Q R) is locked. Posted or cancelled: only S stays open.
Public Const SPECIAL_LOCKS_OPEN = "1000000000000111110"
Public Const SPECIAL_LOCKS_POSTED = "1111111111111111110"

' Q «Статус» of a posted line (the kind of the line after the prefix; «Проведено (исправлено): » after a correction)
Public Const XS_POSTED = "Проведено: "
Public Const XS_FIXED = "Проведено (исправлено): "
Public Const XK_NEW = "новый ЕИ"
Public Const XK_ADD = "пополнение ЕИ"
Public Const XK_REVIEW = "требует разбора"
Public Const XK_IDENT = "разобрано"

' _SPR — the lines of the special receipts, one row per line № (row index = line №, dense like «Наличие»)
Public Const SR_NO = 0
Public Const SR_EVENT = 1
Public Const SR_TYPE = 2
Public Const SR_EI = 3
Public Const SR_MODE = 4
Public Const SR_QTY = 5
Public Const SR_STATE = 6
Public Const SR_ROW = 7
Public Const SR_DATE = 8
Public Const SR_DUP = 9
Public Const SR_DUPKEY = 10
Public Const SR_ARTKEY = 11
Public Const SR_IDENT = 12
Public Const SR_FIXES = 13
Public Const SR_LAST = 13
' SR_MODE: the line created its EI (NEW) or added to an existing EI of a part (ADD)
Public Const SR_NEW = "NEW"
Public Const SR_ADD = "ADD"

' _ART — the index «article of a part → its EI» (one row per article, appended; spec §8: one article — one EI)
Public Const AR_KEY = 0
Public Const AR_EI = 1
Public Const AR_ART = 2
Public Const AR_LAST = 2

' the sources: code (_SPR, journal), prefix of the event ID, name on the sheet, «Тип источника» of the EI in «Наличие»
Public Const SRC_CODES = "OFF|PROD|DET|OLD|OTH"
Public Const SRC_NAMES = "Офис|Производство|Детали|Старый склад|Иной"
Public Const SRC_STYPES = "Офис|Производство|Детали|Старый склад|Иной приход"
' the «Тип источника» of a detail EI (one article — one EI; a later receipt of the article refills it)
Public Const SRC_STYPES_PART = "Детали"
Public Const EVENT_DIGITS = 8

' «Площадка / Поставщик» values that start a special receipt (spec §7): not an ordinary receipt of «Заказы» — it is
' posted on the sheet «Иной приход» (Core Phase 5)
Public Const SPECIAL_SUPPLIERS = "офис|производство|детали|старый склад|иной|иной приход"
' «Обновить статусы» examines at most this many separate blocks of rows without a status
Public Const REFRESH_MAX_BLOCKS = 20000

' ---------------------------------------------------------------- Final Core: corrections of the stock (FINAL WMS MARATHON §2)
' One light user sheet for the last operations of the core: MOVE (a place change of an EI), WRITE_OFF (a write-off, not an
' issue to a person), INV_ADJ (an inventory correction: the difference between the actual and the book balance).
Public Const SH_ADJUST = "Корректировки"
Public Const SH_ADJ = "_ADJ"

' «Корректировки» A:S: 0-based column indices
Public Const AC_NO = 0
Public Const AC_KIND = 1
Public Const AC_EI = 2
Public Const AC_NAME = 3
Public Const AC_ART = 4
Public Const AC_UNIT = 5
Public Const AC_QTY = 6
Public Const AC_FACT = 7
Public Const AC_BOOK = 8
Public Const AC_PLACE = 9
Public Const AC_DATE = 10
Public Const AC_REASON = 11
Public Const AC_FROM = 12
Public Const AC_BEFORE = 13
Public Const AC_AFTER = 14
Public Const AC_DIFF = 15
Public Const AC_CTL = 16
Public Const AC_BATCH = 17
Public Const AC_NOTE = 18
Public Const AC_LAST = 18

' cell protection of a «Корректировки» row, one character per column A..S ("1" = locked). Unposted: the inputs B C G H I J
' K L, the batch mark R and the comment S are open; what WMS fills (A D E F M N O P Q) is locked. Posted or cancelled:
' only S stays open.
Public Const ADJUST_LOCKS_OPEN = "1001110000001111100"
Public Const ADJUST_LOCKS_POSTED = "1111111111111111110"

' the kinds: name in column B, code in _ADJ and the journal
Public Const ADJ_KIND_NAMES = "Перемещение|Списание|Инвентаризация"
Public Const ADJ_KIND_CODES = "MOVE|WRITE_OFF|INV_ADJ"

' _ADJ — the corrections, one row per № (row index = №, dense like «Наличие»): kind, EI, change of the balance
' (WRITE_OFF < 0, INV_ADJ the difference, MOVE 0), state LIVE / STORNO, row hint, places before and after the move,
' the book and the actual balance of an inventory, the date
Public Const AJ_NO = 0
Public Const AJ_KIND = 1
Public Const AJ_EI = 2
Public Const AJ_QTY = 3
Public Const AJ_STATE = 4
Public Const AJ_ROW = 5
Public Const AJ_FROM = 6
Public Const AJ_TO = 7
Public Const AJ_BOOK = 8
Public Const AJ_FACT = 9
Public Const AJ_DATE = 10
Public Const AJ_LAST = 10

' ---------------------------------------------------------------- M6: «Приход авто» — the journal of vehicles (M6 PRIME §1)
' Not a receipt of goods: the balances never change. The user types two cells of the panel above the table («Марка /
' госномер», «Поставщик») and presses ПРИЕХАЛ; at the departure the row of the vehicle is selected and УЕХАЛ pressed.
' Everything else is written by WMS: one operation CAR_ARRIVE / CAR_DEPART / CAR_FIX / CAR_CANCEL each.
Public Const SH_CARS = "Приход авто"
Public Const SH_CAR = "_CAR"

' the panel (rows 0..4, frozen with the header row): the input cells, the hint, the indicators, the threshold setting
Public Const CP_IN_ROW = 1
Public Const CP_PLATE_COL = 2
Public Const CP_SUP_COL = 3
Public Const CP_HINT_ROW = 2
Public Const CP_VAL_ROW = 4
Public Const CP_THR_COL = 11
' the header of the table and its first data row: visit n is on row CAR_FIRST + n - 1 (dense, like «Наличие»; the rows
' are written by WMS only — locked, no rows can be inserted, sorting a protected sheet is refused)
Public Const CAR_HEAD_ROW = 5
Public Const CAR_FIRST = 6

' «Приход авто» A:N: 0-based column indices of the table
Public Const CC_NO = 0
Public Const CC_DATE = 1
Public Const CC_PLATE = 2
Public Const CC_SUP = 3
Public Const CC_ARR = 4
Public Const CC_DEP = 5
Public Const CC_DUR = 6
Public Const CC_STATUS = 7
Public Const CC_WDAY = 8
Public Const CC_MONTH = 9
Public Const CC_NDAY = 10
Public Const CC_CTL = 11
Public Const CC_NOTE = 12
' N: the search key of the vehicle (hidden column: «Найти» filters by it)
Public Const CC_KEY = 13
Public Const CC_LAST = 13

' H «Статус» of a visit
Public Const CS_OPEN = "На территории"
Public Const CS_GONE = "Уехал"
Public Const CS_CANCEL = "Отменён"

' _CAR — the visits, one row per visit № (row index = №, dense like «Наличие»): key of the vehicle, state, the row of the
' table, arrival, departure, the day and the № of the day, the key and the arrival of an open visit (only while it is open:
' the lookups of the Calc engine find an open visit of a vehicle and the longest one without a loop over the history),
' fixes, duration (days), the key of the supplier
Public Const CR_NO = 0
Public Const CR_KEY = 1
Public Const CR_STATE = 2
Public Const CR_ROW = 3
Public Const CR_ARR = 4
Public Const CR_DEP = 5
Public Const CR_DAY = 6
Public Const CR_NDAY = 7
Public Const CR_OKEY = 8
Public Const CR_OARR = 9
Public Const CR_FIXES = 10
Public Const CR_DUR = 11
Public Const CR_SKEY = 12
Public Const CR_LAST = 12
Public Const CR_OPEN = "OPEN"
Public Const CR_CLOSED = "CLOSED"
Public Const CR_CANCELLED = "CANCELLED"
' the dictionaries of _CAR (the lists of the two input cells): vehicles N:O (as typed first, key), suppliers P:Q; row k =
' entry k; their sizes in _CAR!S1 and _CAR!U1
Public Const CD_PLATE = 13
Public Const CD_PKEY = 14
Public Const CD_SUP = 15
Public Const CD_SKEY = 16
Public Const CD_NPLATE_COL = 18
Public Const CD_NSUP_COL = 20

' a stay longer than this (minutes) is marked on the sheet unless the threshold cell of the panel holds another number
Public Const CAR_LONG_MIN = 120
' УЕХАЛ within this many minutes after ПРИЕХАЛ asks first (a click on the wrong button)
Public Const CAR_QUICK_MIN = 2
' the texts of a vehicle and a supplier are limited (a stray paste into the input cell is refused)
Public Const CAR_TEXT_MAX = 120

Function SysKeyNames() As Variant
    SysKeyNames = Array("SCHEMA", "INSTANCE_ID", "MODE", "CORE_VERSION", "LAST_SEQ", "NEXT_EI", "NEXT_NO", "NEXT_RET", _
        "JOURNAL_POS", "REGISTERED_URL", "TX_STATE", "TX_SEQ", "TX_TYPE", "TX_TIME", "TX_BEFORE_IMAGE", _
        "SAVE_STAMP", "SAVE_SEQ", "MAX_QTY", "KEY_SHEETS", "NEXT_SPL", "NEXT_OFF", "NEXT_PROD", "NEXT_DET", "NEXT_OLD", "NEXT_OTH", "NEXT_ADJ", _
        "NEXT_CAR")
End Function

' keys of _SYS that must hold numbers
Function SysNumericKeys() As Variant
    SysNumericKeys = Array(SK_LAST_SEQ, SK_NEXT_EI, SK_NEXT_NO, SK_NEXT_RET, SK_TX_SEQ, SK_SAVE_STAMP, SK_SAVE_SEQ, SK_MAX_QTY, _
        SK_NEXT_SPL, SK_NEXT_OFF, SK_NEXT_PROD, SK_NEXT_DET, SK_NEXT_OLD, SK_NEXT_OTH, SK_NEXT_ADJ, SK_NEXT_CAR)
End Function

' counters that must be ≥ 1 (numbers handed out from 1)
Function SysCounterKeys() As Variant
    SysCounterKeys = Array(SK_NEXT_EI, SK_NEXT_NO, SK_NEXT_RET, SK_MAX_QTY, SK_NEXT_SPL, SK_NEXT_OFF, SK_NEXT_PROD, SK_NEXT_DET, _
        SK_NEXT_OLD, SK_NEXT_OTH, SK_NEXT_ADJ, SK_NEXT_CAR)
End Function

' ---------------------------------------------------------------- AutoInput (spec §10, §29; decision D-031)
' The setting belongs to the LibreOffice user profile and is global: it affects every Calc document of this profile
' while WMS runs. The original value is remembered in a small file inside the profile (not in the book), so that it
' survives a crash or an unsaved session and is restored at the next clean close of WMS.

Private Function CalcInputAccess(bUpdate As Boolean) As Object
    Dim cp As Object, a(0) As New com.sun.star.beans.PropertyValue
    cp = CreateUnoService("com.sun.star.configuration.ConfigurationProvider")
    a(0).Name = "nodepath"
    a(0).Value = "/org.openoffice.Office.Calc/Input"
    If bUpdate Then
        CalcInputAccess = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", a())
    Else
        CalcInputAccess = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", a())
    End If
End Function

Function AutoInputGet() As Boolean
    AutoInputGet = CalcInputAccess(False).getByName("AutoInput")
End Function

Sub AutoInputSet(bOn As Boolean)
    Dim ua As Object
    ua = CalcInputAccess(True)
    ua.replaceByName("AutoInput", bOn)
    ua.commitChanges()
End Sub

Private Function AutoInputStateUrl() As String
    AutoInputStateUrl = CreateUnoService("com.sun.star.util.PathSettings").UserConfig & "/" & AUTOINPUT_STATE_FILE
End Function

' Called at WMS start. Returns a short note for the startup report.
Function AutoInputDisable() As String
    Dim sfa As Object, url As String, wasOn As Boolean
    On Error GoTo EH
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    url = AutoInputStateUrl()
    wasOn = AutoInputGet()
    If Not sfa.exists(url) Then
        ' remember the user's own value only once: after a crash the file still holds the original
        WmsCore.WriteTextFile(url, "original=" & IIf(wasOn, "1", "0") & Chr(10))
    End If
    If wasOn Then AutoInputSet(False)
    AutoInputDisable = "автоввод Calc выключен на время работы WMS (настройка общая для LibreOffice)"
    Exit Function
EH:
    AutoInputDisable = "ВНИМАНИЕ: не удалось выключить автоввод Calc: " & Error$
End Function

' Called at clean close. Restores the remembered value and forgets it.
Function AutoInputRestore() As String
    Dim sfa As Object, url As String, s As String
    On Error GoTo EH
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    url = AutoInputStateUrl()
    If Not sfa.exists(url) Then
        AutoInputRestore = ""
        Exit Function
    End If
    s = WmsCore.ReadTextFile(url)
    If InStr(s, "original=1") > 0 Then AutoInputSet(True)
    If InStr(s, "original=0") > 0 Then AutoInputSet(False)
    sfa.kill(url)
    AutoInputRestore = "автоввод Calc восстановлен"
    Exit Function
EH:
    AutoInputRestore = "ВНИМАНИЕ: не удалось восстановить автоввод Calc: " & Error$
End Function
