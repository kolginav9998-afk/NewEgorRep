"""WMS_LEGACY_TRANSFER — перенос старой таблицы «Заказы» (A:AB) в новую WMS 0.7 (задание M7 PRIME).

    python3 tools/legacy_transfer.py check   --staging STAGING.ods --work WORK [--candidate BOOK] [--settings FILE] [rules]
    python3 tools/legacy_transfer.py build   --staging STAGING.ods --work WORK --candidate BOOK [--settings FILE] [rules]
    python3 tools/legacy_transfer.py verify  --staging STAGING.ods --work WORK [--settings FILE] [rules]
    python3 tools/legacy_transfer.py promote --work WORK --to WMS_WORK [--name WMS_PROD.ods]
    python3 tools/legacy_transfer.py status  --work WORK

The engine of the book WMS_LEGACY_TRANSFER.ods (its buttons run it); it can be run by hand as well. STAGING is a copy of
that book (the frontend stores it with the formulas already frozen to their values): the sheets «1_Вставить_Заказы» (the
old «Заказы» A:AB pasted with one Ctrl+V, the header row included or not), and optionally «1B_Вставить_Наличие» (the old
stock), «1C_Вставить_Выдачи» and «1D_Вставить_Возвраты» (the old issues and returns: EI and quantity by the header).
The old file itself is never opened or written.

CHECK    (dry run, nothing but WORK/check/ is written) — the whole transfer is planned in memory: the orders (column A is
         the № of the position inside its order: every source row with A = 1 starts the next order), the positions and
         their further deliveries (a row without the ordered quantity H but with a received quantity F, attached to the
         position above it in the same order), the EIs (427, 00000427, ЕИ-427, EI-427 → ЕИ-00000427; never renumbered;
         a row without an EI gets a new number above every old one), the current balance of every EI (X «Наличие» when it
         is reliable, else the sheet 1B, else received − issued (1C) + returned (1D), else the explicit rule — never F
         without a rule), the old status W (kept as LegacyStatus) against the status the new engine computes, the parts
         (one article of «Детали» — one EI). Every finding has a level: BLOCK (red — the transfer is refused), WARN
         (yellow — transferred, look at it), OK (green). Output: summary.csv (sheet «2_Проверка»), errors.csv («3_Ошибки»),
         rows.csv (the result of every pasted row), plan.json (the exact plan), CHECK_REPORT.md.
BUILD    only with 0 blocking findings: the clean candidate of the release (WMS_PROD_CANDIDATE.ods, schema WMS-SYS-4,
         core with the transfer operations) is copied — the candidate itself is never changed — to WORK/test.part/, WMS
         opens the copy and the plan is executed by the operations of the core: the order rows are written like a user
         types them, every old receipt is posted by LEGACY_RECEIPT / LEGACY_RECEIPT_ADD (the old EI and its current
         balance; the position in _ORD, the receipt in _RCV, the row of «Наличие», the counters), a cancelled order by
         ORDER_CANCEL / ORDER_CANCEL_REST, a special receipt (Офис, Производство, Детали, Старый склад, Иной) by MIGRATE;
         LEGACY_BEGIN / LEGACY_END mark the transfer in the journal. Then the statuses are refreshed, the index of the parts
         is rebuilt, the self-check and the independent oracle (the whole journal replayed) must find nothing, the
         reconciliation with the plan must have no critical difference; the book is saved, closed, opened again and
         checked again. Only then WORK/test.part/ becomes WORK/test/ (WMS_LEGACY_TEST.ods). Any refusal stops the build:
         WORK/test.part/ is removed as a whole — the transfer happens whole or not at all.
VERIFY   the test WMS against the staging: every pasted row is found (moved, merged or transferred as an EI — nothing
         lost), orders, positions, open rests, received quantities, old EIs, balances, dates, documents, suppliers,
         articles, units, places, categories, prices and sums, statuses, NEXT_EI, the links receipt → position, no negative
         balance; the self-check and the oracle once more. LEGACY_TRANSFER_REPORT.md / .html: «Полностью совпало»,
         «Нормализовано», «Пересчитано новым движком», «Требует внимания», «Ошибка».
PROMOTE  only after a successful VERIFY of the unchanged test WMS and unchanged staging and rules: WMS_WORK/WMS_PROD.ods
         with its journal is created (the folder must not exist or be empty — an existing PROD is never written), WMS
         registers it there as the working book; WMS_TOOLBOX of the release is put next to it; the reports, a backup of the
         release candidate and CUTOVER_README.txt (from now on every movement goes to the new WMS only; the old table stays
         a read-only archive; time, SHA-256 of the staging and of the new WMS).
STATUS   the state of the steps (WORK/state.txt for the frontend).
Exit code: 0 done, 1 a stop by the rules (nothing changed), 2 the input could not be read or a tool error.
"""
import argparse
import datetime
import hashlib
import html
import json
import os
import re
import shutil
import sys
import tempfile
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "tests"))

TOOL_VERSION = "1.0.0"
SCHEMA = "WMS-SYS-4"
MIN_CORE = (0, 7, 1)
SHEET_ORDERS = "1_Вставить_Заказы"
SHEET_STOCK = "1B_Вставить_Наличие"
SHEET_ISSUES = "1C_Вставить_Выдачи"
SHEET_RETURNS = "1D_Вставить_Возвраты"
STAGING_SHEETS = (SHEET_ORDERS, SHEET_STOCK, SHEET_ISSUES, SHEET_RETURNS)
TEST_BOOK = "WMS_LEGACY_TEST.ods"
JOURNAL_DIR = "WMS_Journal"
LOCK_FILE = "wms.lock"
SELF_CHECK_OK = "САМОПРОВЕРКА WMS: ошибок 0"
REPORT_NAME = "LEGACY_TRANSFER_REPORT"

COLS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y",
        "Z", "AA", "AB"]
COL_NAMES = ["№ позиции", "Наименование", "Номер документа", "Номер счёта", "Артикул", "Фактическое количество", "Количество по документу",
             "Заказанное количество", "Единица измерения", "Цена", "Сумма", "Площадка / Поставщик", "Продавец", "Дата поступления",
             "Дата документа", "Дата заказа", "Ожидаемая дата поступления", "Покупатель", "Категория", "Кому назначено", "Место хранения",
             "Внутренний код (ЕИ)", "Статус", "Наличие", "Контроль", "Комментарий", "Срок поставки, дней", "Возможный дубль"]
(C_A, C_B, C_C, C_D, C_E, C_F, C_G, C_H, C_I, C_J, C_K, C_L, C_M, C_N, C_O, C_P, C_Q, C_R, C_S, C_T, C_U, C_V, C_W, C_X, C_Y, C_Z,
 C_AA, C_AB) = range(28)
TEXT_COLS = (C_B, C_C, C_D, C_E, C_L, C_M, C_R, C_S, C_T, C_U, C_Z)
# the columns the tool types into «Заказы» (the inputs of an order row); V W X Y AB are written by WMS only
INPUT_COLS = tuple(range(C_A, C_U + 1)) + (C_Z, C_AA)
# the identity of a position that «Ещё поступление» copies from the source row into a delivery row (WmsReceipt)
IDENTITY_COLS = (C_A, C_B, C_D, C_E, C_I, C_L, C_M, C_P, C_R, C_S, C_T)

OS_WAITING, OS_OVERDUE, OS_PARTIAL, OS_RECEIVED = "Ожидается", "Просрочено", "Частично получено", "Получено"
OS_NODOCS, OS_CANCELLED = "Получено без документов", "Отменено"
OS_REST_CANCELLED, OS_PARTIAL_OVERDUE = "Частично получено / остаток отменён", "Частично получено / просрочено"
OS_ADD = "Дополнительное поступление"
RECEIVED_FAMILY = (OS_PARTIAL, OS_PARTIAL_OVERDUE, OS_RECEIVED, OS_NODOCS, OS_REST_CANCELLED)
EI_ACTIVE, EI_REVIEW = "Активен", "Требует разбора"
STYPE_SUPPLIER = "Поставщик"
SPECIAL = {"офис": "Офис", "производство": "Производство", "детали": "Детали", "старый склад": "Старый склад", "иной": "Иной приход",
           "иной приход": "Иной приход"}
PART = "Детали"
DEFAULT_UNITS = ["шт", "м", "кг", "л", "упак", "пар", "компл", "рул", "м2", "м3", "т", "г", "мл", "см", "мм", "лист", "набор", "пог. м", "бух"]
UNIT_SYNONYMS = {"штук": "шт", "штука": "шт", "штуки": "шт", "шт.": "шт", "метр": "м", "метров": "м", "метра": "м", "м.": "м",
                 "кг.": "кг", "килограмм": "кг", "литр": "л", "литров": "л", "л.": "л", "упаковка": "упак", "упак.": "упак",
                 "уп.": "упак", "уп": "упак", "пара": "пар", "пары": "пар", "пар.": "пар", "комплект": "компл", "компл.": "компл",
                 "к-т": "компл", "кмп": "компл", "рулон": "рул", "рул.": "рул", "лист.": "лист", "листов": "лист", "м²": "м2",
                 "кв.м": "м2", "кв. м": "м2", "м³": "м3", "куб.м": "м3", "куб. м": "м3", "гр": "г", "гр.": "г", "мл.": "мл",
                 "см.": "см", "мм.": "мм", "набор.": "набор", "бухта": "бух", "бух.": "бух", "п.м": "пог. м", "пог.м": "пог. м",
                 "пог. м.": "пог. м", "т.": "т", "тонна": "т"}

# the old statuses (W) as the users wrote them → the status of WMS (the key: lower case, ё = е, punctuation as spaces)
OLD_STATUS = {
    "ожидается": OS_WAITING, "ожидание": OS_WAITING, "ожидаем": OS_WAITING, "в пути": OS_WAITING, "заказано": OS_WAITING,
    "в заказе": OS_WAITING, "ждем": OS_WAITING, "не получено": OS_WAITING,
    "просрочено": OS_OVERDUE, "просрочен": OS_OVERDUE, "просрочка": OS_OVERDUE,
    "частично получено": OS_PARTIAL, "частично": OS_PARTIAL, "получено частично": OS_PARTIAL, "частично получен": OS_PARTIAL,
    "частично получено / просрочено": OS_PARTIAL_OVERDUE, "частично получено просрочено": OS_PARTIAL_OVERDUE,
    "частично просрочено": OS_PARTIAL_OVERDUE, "частично / просрочено": OS_PARTIAL_OVERDUE,
    "получено": OS_RECEIVED, "получен": OS_RECEIVED, "доставлено": OS_RECEIVED, "принято": OS_RECEIVED, "выполнено": OS_RECEIVED,
    "закрыто": OS_RECEIVED, "получено полностью": OS_RECEIVED,
    "получено без документов": OS_NODOCS, "без документов": OS_NODOCS, "получено б / д": OS_NODOCS, "получено без док": OS_NODOCS,
    "получено нет документов": OS_NODOCS,
    "отменено": OS_CANCELLED, "отменен": OS_CANCELLED, "отмена": OS_CANCELLED, "аннулировано": OS_CANCELLED, "отказ": OS_CANCELLED,
    "частично получено / остаток отменен": OS_REST_CANCELLED, "частично получено остаток отменен": OS_REST_CANCELLED,
    "остаток отменен": OS_REST_CANCELLED, "частично / остаток отменен": OS_REST_CANCELLED,
    "дополнительное поступление": OS_ADD, "доп поступление": OS_ADD, "поступление": OS_ADD,
}

LEVELS = ("BLOCK", "WARN", "INFO")
RESULT_BLOCK, RESULT_WARN, RESULT_OK, RESULT_SKIP = "БЛОКЕР", "Проверить", "Готово", "Пропущено"
QTY_MAX_DECIMALS, QTY_MAX_INT_DIGITS, MAX_QTY_DEFAULT = 3, 9, 100000.0
EPS = 1e-6
D2000, D2099 = 36526, 73050            # serials of 01.01.2000 and 31.12.2099 (Calc: 30.12.1899 = 0)

# the rules of the transfer (sheet «Настройки» of the frontend; «stop» — the safe default: a blocking finding)
DEFAULT_SETTINGS = dict(default_unit="", default_place="", unknown_balance="stop", no_order_qty="stop", no_receipt_date="stop",
                        stock_only="stop", price_round="stop", empty_rows="stop", units="|".join(DEFAULT_UNITS), places="",
                        source_file="")
SETTING_CHOICES = dict(unknown_balance=("stop", "zero", "fact"), no_order_qty=("stop", "fact"), no_receipt_date=("stop", "doc"),
                       stock_only=("stop", "skip", "old"), price_round=("stop", "round"), empty_rows=("stop", "skip"))
BALANCE_MARK = "[перенос: проверить остаток]"


# ================================================================ small tools

def now():
    return datetime.datetime.now()


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def sha256_text(s):
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


def write_text(path, s):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:
        f.write(s)
    os.replace(tmp, path)


def csv_field(v):
    if v is None:
        return ""
    if isinstance(v, float):
        s = f"{v:.6f}".rstrip("0").rstrip(".") if v != int(v) else str(int(v))
        return s.replace(".", ",")
    s = str(v).replace("\r", " ").replace("\n", " ")
    if any(ch in s for ch in ';"'):
        s = '"' + s.replace('"', '""') + '"'
    return s


def write_csv(path, header, rows):
    write_text(path, "\n".join(";".join(csv_field(x) for x in r) for r in [header] + list(rows)) + "\n")


def qty_text(q):
    """a quantity as WMS parses it: a comma, no trailing zeros ("0" for zero)"""
    s = f"{float(q):.3f}".rstrip("0").rstrip(".")
    return s.replace(".", ",") if s else "0"


def num_text(q):
    return qty_text(q) if q is not None else "—"


def serial_date(d):
    return float((d - datetime.date(1899, 12, 30)).days)


def date_of_serial(x):
    return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(x))


def date_text(x):
    return date_of_serial(x).strftime("%d.%m.%Y") if isinstance(x, (int, float)) and x else ""


def today_serial():
    return serial_date(datetime.date.today())


class Progress:
    """WORK/progress.txt «phase|done|total|text» — the frontend shows it while the engine runs"""

    def __init__(self, work, phase):
        self.path = os.path.join(work, "progress.txt")
        self.phase = phase
        self.last = 0.0

    def __call__(self, done, total, text, force=False):
        t = time.time()
        if not force and t - self.last < 0.4:
            return
        self.last = t
        try:
            write_text(self.path, f"{self.phase}|{done}|{total}|{text}\n")
        except OSError:
            pass


# ================================================================ the rules of WMS (written for this tool from the specification of the core)

def is_empty(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


# a placeholder written instead of an empty cell in a column of numbers, dates or EIs
BLANK_LIKE = {"-", "--", "---", "—", "–", ".", "нет", "н/д", "н.д.", "n/a", "na"}
BLANK_COLS = (C_F, C_G, C_H, C_J, C_K, C_N, C_O, C_P, C_Q, C_V, C_X, C_AA)


def blank_like(v):
    return isinstance(v, str) and v.strip().lower() in BLANK_LIKE


def strip_unit(v, unit):
    """(value, True) for a quantity written with the unit of its row («60 шт» when the unit is «шт»), else (v, False)"""
    if not isinstance(v, str) or not txt(unit):
        return v, False
    m = re.fullmatch(r"\s*(\d[\d.,]*)\s*([^\d\s.,][^\d]*?)\s*", v)
    if m and unit_value(m.group(2))[0] == unit_value(unit)[0]:
        return m.group(1), True
    return v, False


def txt(v):
    """a cell as text: a whole number without «,0», a fraction with a comma, text trimmed"""
    if v is None:
        return ""
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return (f"{v:.6f}".rstrip("0").rstrip(".")).replace(".", ",")
    return str(v).strip()


def norm_space(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def name_key(v):
    s = txt(v).lower().replace("ё", "е")
    return re.sub(r"[^0-9a-zа-я]+", " ", s).strip()


_LATIN = {"А": "A", "В": "B", "Е": "E", "Ё": "E", "К": "K", "М": "M", "Н": "H", "О": "O", "Р": "P", "С": "C", "Т": "T", "У": "Y", "Х": "X"}
_DROP = set("\t\n\r       ­")


def article_key(s):
    """the key of an article as the index of the parts keeps it (WmsSpecial.ArticleKey): upper case, no spaces, every dash
    as "-", Cyrillic lookalikes as Latin"""
    out = []
    for ch in txt(s).upper():
        if ch in _DROP:
            continue
        if "‐" <= ch <= "―" or ch == "−":
            ch = "-"
        out.append(_LATIN.get(ch, ch))
    return "".join(out)


def parse_qty_text(s, allow_zero=False):
    """(value, "") or (None, reason): the locale-independent parser of WMS (WmsCore.ParseQtyText, D-030)"""
    t = str(s)
    t = t.strip(" \t  ")
    if t == "":
        return None, "не указано количество"
    ip, fp, nsep = "", "", 0
    for ch in t:
        if "0" <= ch <= "9":
            if nsep == 0:
                ip += ch
            else:
                fp += ch
        elif ch in ",.":
            nsep += 1
            if nsep > 1:
                return None, f"«{t}»: несколько разделителей — похоже на дату или разряды"
        elif ch in "-−":
            return None, f"«{t}»: отрицательное количество"
        elif ch in "eE":
            return None, f"«{t}»: экспоненциальная запись"
        elif ch == "/" or ch == "⁄" or "¼" <= ch <= "¾" or "⅐" <= ch <= "⅞":
            return None, f"«{t}»: дробь — нужно десятичное число"
        elif ch in "   '":
            return None, f"«{t}»: пробелы и разделители разрядов не допускаются"
        else:
            return None, f"«{t}»: недопустимый символ «{ch}»"
    if ip == "":
        return None, f"«{t}»: нет целой части"
    if nsep == 1 and fp == "":
        return None, f"«{t}»: после разделителя нет цифр"
    if len(ip) > 1 and ip[0] == "0":
        return None, f"«{t}»: лишние нули в начале числа"
    if len(ip) > QTY_MAX_INT_DIGITS:
        return None, f"«{t}»: слишком большое число"
    if len(fp) == 3 and ip != "0" and len(ip) <= 3:
        return None, f"«{t}»: неоднозначно — {ip},{fp.rstrip('0') or '0'} или {ip}{fp}? Для целого запишите {ip}{fp}"
    fp = fp.rstrip("0")
    if len(fp) > QTY_MAX_DECIMALS:
        return None, f"«{t}»: не более {QTY_MAX_DECIMALS} знаков после запятой"
    v = float(ip + ("." + fp if fp else ""))
    if v <= 0 and not allow_zero:
        return None, f"«{t}»: количество должно быть больше 0"
    return v, ""


def qty_value(v, allow_zero=False, max_qty=MAX_QTY_DEFAULT):
    """(value, "") or (None, reason) for a cell value: a number (> 0, at most 3 decimals) or text by the parser of WMS"""
    if isinstance(v, float):
        if v < 0 or (v == 0 and not allow_zero):
            return None, f"{txt(v)} — количество должно быть больше 0"
        if abs(v * 1000 - round(v * 1000)) > 1e-6:
            return None, f"{txt(v)} — не более {QTY_MAX_DECIMALS} знаков после запятой"
        q = round(v, 3)
    else:
        q, why = parse_qty_text(v, allow_zero)
        if q is None:
            return None, why
    if q > max_qty:
        return None, f"{qty_text(q)} — больше порога MAX_QTY ({qty_text(max_qty)})"
    return q, ""


def price_value(v, rounding=False):
    """(value, "", note) or (None, reason, "") — J «Цена»: a number ≥ 0 with at most 3 decimals"""
    if isinstance(v, float):
        if v < 0:
            return None, "отрицательная цена", ""
        if abs(v * 1000 - round(v * 1000)) > 1e-6:
            if rounding:
                return round(v, 3), "", f"цена {txt(v)} округлена до {qty_text(round(v, 3))} (правило)"
            return None, f"цена {txt(v)} — больше 3 знаков после запятой", ""
        return round(v, 3), "", ""
    t = txt(v).replace(" ", "").replace(" ", "")
    t = re.sub(r"(руб\.?|р\.?|₽)$", "", t, flags=re.I).strip()
    q, why = parse_qty_text(t, allow_zero=True)
    if q is None:
        if rounding:
            m = re.fullmatch(r"(\d+)[.,](\d+)", t)
            if m:
                x = round(float(m.group(1) + "." + m.group(2)), 3)
                return x, "", f"цена «{txt(v)}» округлена до {qty_text(x)} (правило)"
        return None, f"цена «{txt(v)}»: {why}", ""
    return q, "", ""


def number_value(v):
    """K «Сумма», AA «Срок поставки»: a number (text with a comma or spaces accepted) or None"""
    if isinstance(v, float):
        return v
    t = txt(v).replace(" ", "").replace(" ", "")
    t = re.sub(r"(руб\.?|р\.?|₽|дн\.?|дней|день|дня)$", "", t, flags=re.I).strip()
    if re.fullmatch(r"-?\d+([.,]\d+)?", t):
        return float(t.replace(",", "."))
    return None


def date_value(v):
    """(serial, "", note) or (None, reason, ""): a Calc date number (the time part dropped) or a text date: dd.mm.yyyy
    (as WMS reads it); dd.mm.yy, yyyy-mm-dd, dd/mm/yyyy are accepted and noted as normalized; 2000–2099 only"""
    note = ""
    if isinstance(v, float):
        x = float(int(v))
        if v != x:
            note = "время отброшено"
    else:
        t = txt(v)
        m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", t)
        m2 = re.fullmatch(r"(\d{1,2})[./](\d{1,2})[./](\d{2})", t)
        m3 = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", t)
        m4 = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", t)
        try:
            if m:
                d = datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            elif m2:
                d = datetime.date(2000 + int(m2.group(3)), int(m2.group(2)), int(m2.group(1)))
                note = f"дата «{t}» прочитана как {d:%d.%m.%Y}"
            elif m3:
                d = datetime.date(int(m3.group(1)), int(m3.group(2)), int(m3.group(3)))
                note = f"дата «{t}» прочитана как {d:%d.%m.%Y}"
            elif m4:
                d = datetime.date(int(m4.group(3)), int(m4.group(2)), int(m4.group(1)))
                note = f"дата «{t}» прочитана как {d:%d.%m.%Y}"
            else:
                return None, f"«{t}» не является датой (дд.мм.гггг)", ""
        except ValueError:
            return None, f"«{t}» — такой даты нет", ""
        x = serial_date(d)
    if x < D2000 or x > D2099:
        return None, f"{date_text(x) if D2000 - 40000 < x < D2099 + 40000 else txt(v)} — вне диапазона 2000–2099", ""
    return x, "", note


def ei_value(v):
    """(n, canonical, normalized?) or (None, None, reason): 427, 00000427, ЕИ-427, EI-427, ЕИ 427 → ЕИ-00000427"""
    if isinstance(v, float):
        if v != int(v) or not 1 <= v <= 99999999:
            return None, None, f"«{txt(v)}» не похоже на номер ЕИ"
        n = int(v)
        return n, f"ЕИ-{n:08d}", True
    raw = str(v).strip()
    t = raw.upper()
    if len(t) >= 2 and t[0] in "ЕE" and t[1] in "ИI":
        t = t[2:]
        if t[:1] in ("-", " ", "–", "—", "_", " "):
            t = t[1:].strip()
    if not t or len(t) > 8 or not t.isdigit():
        return None, None, f"«{raw}» не похоже на номер ЕИ (пример: ЕИ-00000427 или 427)"
    n = int(t)
    if n < 1:
        return None, None, "номер ЕИ должен быть больше 0"
    canon = f"ЕИ-{n:08d}"
    return n, canon, raw != canon


def unit_value(v):
    """(unit, note): the unit as WMS keeps it — lower case, the usual abbreviations"""
    raw = norm_space(txt(v))
    u = raw.lower()
    u = UNIT_SYNONYMS.get(u, u)
    if u.endswith(".") and u[:-1] in DEFAULT_UNITS:
        u = u[:-1]
    return u, (f"единица «{raw}» → «{u}»" if u != raw else "")


def pos_no(v):
    """(n, "", note) or (None, reason, ""): A — the № of the position inside its order"""
    if isinstance(v, float):
        if v == int(v) and 1 <= v <= 99999:
            return int(v), "", ""
        return None, f"«{txt(v)}» — не номер позиции", ""
    t = txt(v)
    m = re.fullmatch(r"0*(\d{1,5})\.?", t)
    if not m or int(m.group(1)) < 1:
        return None, f"«{t}» — не номер позиции", ""
    n = int(m.group(1))
    return n, "", (f"№ позиции «{t}» → «{n}»" if t != str(n) else "")


def special_source(v):
    """the kind of source of a special receipt (L «Площадка / Поставщик», WmsOrders.SpecialSupplier) or \"\""""
    return SPECIAL.get(norm_space(txt(v)).lower(), "")


def old_status(v):
    """(the status of WMS or None, the text as it was)"""
    raw = norm_space(txt(v))
    if not raw:
        return None, ""
    t = raw.lower().replace("ё", "е")
    t = re.sub(r"[()\.,;:!«»\"]+", " ", t)
    t = re.sub(r"\s*/\s*", " / ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return OLD_STATUS.get(t) or OLD_STATUS.get(t.replace(" / ", " ")), raw


def position_status(ordq, rcv, nodoc, cancel, edate, today):
    """WmsOrders.PositionStatus"""
    if cancel == "ORDER":
        return OS_CANCELLED
    if rcv <= EPS:
        if cancel == "REST":
            return OS_CANCELLED
        return OS_OVERDUE if edate and edate < today else OS_WAITING
    if rcv >= ordq - EPS:
        return OS_NODOCS if nodoc > 0 else OS_RECEIVED
    if cancel == "REST":
        return OS_REST_CANCELLED
    return OS_PARTIAL_OVERDUE if edate and edate < today else OS_PARTIAL


# ================================================================ settings

def load_settings(path=None, overrides=None):
    s = dict(DEFAULT_SETTINGS)
    if path and os.path.exists(path):
        for ln in open(path, encoding="utf-8-sig").read().splitlines():
            k, sep, v = ln.partition("=")
            if sep and k.strip() in s:
                s[k.strip()] = v.strip()
    for k, v in (overrides or {}).items():
        if v is not None:
            s[k] = v
    problems = []
    for k, choices in SETTING_CHOICES.items():
        s[k] = (s[k] or "stop").strip().lower()
        if s[k] not in choices:
            problems.append(f"правило «{k}»: «{s[k]}» — допустимо {', '.join(choices)}")
            s[k] = "stop"
    return s, problems


def settings_hash(s):
    return sha256_text(json.dumps({k: s[k] for k in sorted(DEFAULT_SETTINGS)}, ensure_ascii=False, sort_keys=True))


def split_list(s):
    return [x.strip() for x in re.split(r"[|\n;]", s or "") if x.strip()]


# ================================================================ reading the staging (a copy of the frontend book)

class Staging:
    def __init__(self):
        self.sheets = {}                # name → rows (lists of values; "" empty, float, str)
        self.formulas = {}              # name → formula cells found (the frontend freezes them before storing the copy)
        self.errors = {}                # name → cells with an error result
        self.missing = []
        self.path = ""
        self.hash = ""


def _used(sheet, max_col=None):
    cur = sheet.createCursor()
    cur.gotoEndOfUsedArea(False)
    a = cur.getRangeAddress()
    last_col = a.EndColumn if max_col is None else max_col
    if a.EndRow < 0:
        return []
    return [list(r) for r in sheet.getCellRangeByPosition(0, 0, last_col, a.EndRow).getDataArray()]


def _count_cells(ranges):
    n = 0
    for a in ranges.getRangeAddresses():
        n += (a.EndRow - a.StartRow + 1) * (a.EndColumn - a.StartColumn + 1)
    return n


def read_staging(o, path):
    """every staging sheet of the book at path, read-only (macros never run); values as Calc computed them"""
    st = Staging()
    st.path = os.path.abspath(path)
    doc = o.load(st.path, macros=0, hidden=True, ReadOnly=True)
    if doc is None:
        raise ValueError(f"не удалось открыть {path}")
    try:
        for name in STAGING_SHEETS:
            if not doc.Sheets.hasByName(name):
                st.missing.append(name)
                st.sheets[name] = []
                continue
            sh = doc.Sheets.getByName(name)
            rows = _used(sh, 27 if name == SHEET_ORDERS else None)
            for r in rows:
                for i, v in enumerate(r):
                    if v is None:
                        r[i] = "#ОШИБКА"
            st.sheets[name] = rows
            try:
                st.formulas[name] = _count_cells(sh.queryContentCells(16))
                st.errors[name] = _count_cells(sh.queryFormulaCells(4))
            except Exception:
                st.formulas[name] = st.errors[name] = 0
    finally:
        try:
            doc.close(True)
        except Exception:
            pass
    canon = {n: [[(round(v, 9) if isinstance(v, float) else str(v)) for v in r] for r in rows] for n, rows in st.sheets.items()}
    # trailing empty rows and columns do not change the data
    for n in canon:
        rows = [r[:max([i + 1 for i, v in enumerate(r) if v != ""] or [0])] for r in canon[n]]
        while rows and not rows[-1]:
            rows.pop()
        canon[n] = rows
    st.hash = sha256_text(json.dumps(canon, ensure_ascii=False, sort_keys=True))
    return st


# ================================================================ CHECK: the plan of the transfer (nothing is written)

HEADER_KEYS = {C_B: ("наимен",), C_E: ("артикул",), C_F: ("фактич",), C_G: ("по документ",), C_H: ("заказан",), C_I: ("единиц",),
               C_L: ("поставщик", "площадк"), C_N: ("дата поступ",), C_U: ("место",), C_V: ("внутрен", "код", "еи"), C_W: ("статус",),
               C_X: ("наличи",)}
SIDE_KEYS = {
    "ei": ("еи", "ei", "внутренний код", "внутренний код (еи)", "код еи", "код", "№ еи", "номер еи"),
    "qty": ("количество", "кол-во", "qty", "остаток", "наличие", "выдано", "количество выдачи", "количество возврата", "возвращено",
            "сколько", "текущий остаток"),
    "name": ("наименование", "название", "товар", "полное наименование товара", "полное наименование"),
    "art": ("артикул",),
    "unit": ("единица", "единица измерения", "ед", "ед.", "ед. изм.", "ед.изм."),
    "place": ("место", "место хранения", "ячейка"),
    "cat": ("категория",),
}


class SRow:
    """one row of «1_Вставить_Заказы»"""

    def __init__(self, r, v):
        self.r = r                  # the row number as the sheet shows it
        self.v = v                  # A..AB as read (placeholders «-», «нет»… become empty in classify)
        self.orig = list(v)         # A..AB exactly as pasted: the old W X Y AB of the record
        self.kind = ""              # EMPTY HEADER POSITION DELIVERY SPECIAL NOQTY
        self.a = None
        self.block = None
        self.level = ""             # the worst finding: BLOCK / WARN / ""
        self.msgs = []
        self.norm = []              # what was normalized
        self.fate = ""              # row / merged / special / skipped
        self.t = None               # the row of «Заказы» (0-based) it becomes
        self.pos = None
        self.special = None


class Position:
    def __init__(self, row, block):
        self.row = row
        self.block = block
        self.deliveries = []
        self.merged = None          # the first delivery row merged into the source row (the position row had no F)
        self.vals = None            # the planned input values of the source row (col → value)
        self.src_of = None          # col → the SRow the value came from
        self.receipts = []
        self.ordq = 0.0
        self.rcv = 0.0
        self.nodoc = 0
        self.cancel = ""
        self.edate = 0.0
        self.old = None
        self.old_raw = ""
        self.new = ""
        self.status_note = ""
        self.olid = None
        self.t = None


class Item:
    """one EI of the transfer: a receipt of an order position (SRC / ADD) or a special receipt (MIGRATE)"""

    def __init__(self, kind):
        self.kind = kind            # SRC ADD MIGRATE
        self.rows = []              # the staging rows it comes from
        self.pos = None
        self.t = None
        self.vals = None
        self.n = None
        self.ei = ""
        self.ei_src = ""            # OLD NORM NEW
        self.ei_raw = ""
        self.fact = 0.0
        self.bal = None
        self.bal_src = ""
        self.place = ""
        self.place_now = ""
        self.name = self.art = self.unit = self.cat = ""
        self.stype = STYPE_SUPPLIER
        self.state = EI_ACTIVE
        self.nodoc = 0
        self.xraw = ""
        self.olid = None
        self.mark = ""
        self.review = False


class Analysis:
    def __init__(self, st, settings, today=None, max_qty=MAX_QTY_DEFAULT, candidate=None):
        self.st = st
        self.s = settings
        self.today = today or today_serial()
        self.max_qty = max_qty
        self.candidate = candidate or {}
        self.F = []
        self.rows = []
        self.positions = []
        self.items = []
        self.blocks = 0
        self.header = None
        self.side = {}
        self.metrics = []
        self.plan = None
        self.known_units = [u.lower() for u in split_list(settings.get("units"))] or DEFAULT_UNITS
        self.known_places = split_list(settings.get("places"))

    # ------------------------------------------------------------ findings
    def find(self, level, kind, text, srow=None, sheet=SHEET_ORDERS, row=None):
        self.F.append(dict(sheet=sheet, row=(srow.r if srow is not None else row), level=level, kind=kind, text=text))
        if srow is not None:
            srow.msgs.append((level, text))
            if level == "BLOCK" or (level == "WARN" and srow.level != "BLOCK"):
                srow.level = level

    def blockers(self):
        return [f for f in self.F if f["level"] == "BLOCK"]

    def warnings(self):
        return [f for f in self.F if f["level"] == "WARN"]

    # ------------------------------------------------------------ the whole check
    def run(self):
        for p in getattr(self, "settings_problems", []):
            self.find("BLOCK", "настройки", p, row=0, sheet="Настройки")
        for name in self.st.missing:
            if name == SHEET_ORDERS:
                self.find("BLOCK", "структура", f"в книге нет листа «{SHEET_ORDERS}» — используйте WMS_LEGACY_TRANSFER.ods этого выпуска",
                          row=0, sheet=name)
        for name, n in self.st.errors.items():
            if n:
                self.find("INFO", "формула", f"ячеек с ошибкой формулы: {n} (значение прочитано как «#ОШИБКА»)", row=0, sheet=name)
        self.read_side_sheets()
        self.read_orders()
        self.group()
        self.build_positions()
        self.collect_items()
        self.assign_eis()
        self.balances()
        self.statuses()
        self.layout()
        self.check_candidate()
        self.summarize()
        return self

    # ------------------------------------------------------------ 1B / 1C / 1D
    def side_sheet(self, name, need):
        rows = self.st.sheets.get(name) or []
        out = dict(present=False, rows=[], cols={}, skipped=0)
        if not any(any(not is_empty(v) for v in r) for r in rows):
            return out
        out["present"] = True
        hdr = None
        for i, r in enumerate(rows[:10]):
            cols = {}
            for c, v in enumerate(r):
                h = norm_space(txt(v)).lower().rstrip(":")
                for key, names in SIDE_KEYS.items():
                    if h in names and key not in cols:
                        cols[key] = c
            if all(k in cols for k in need):
                hdr, out["cols"] = i, cols
                break
        if hdr is None:
            self.find("BLOCK", "структура", f"лист «{name}»: не найден заголовок с колонками {', '.join(need)} "
                      f"(например: «ЕИ», «Количество») — вставьте лист вместе со строкой заголовка", row=1, sheet=name)
            return out
        cols = out["cols"]
        for i, r in enumerate(rows[hdr + 1:], start=hdr + 2):
            r = list(r) + [""] * 20
            if all(is_empty(v) for v in r):
                continue
            ev = r[cols["ei"]]
            if is_empty(ev):
                out["skipped"] += 1
                continue
            n, canon, why = ei_value(ev)
            if n is None:
                self.find("BLOCK", "ЕИ", f"«{txt(ev)}» в колонке ЕИ — {why}; исправьте или удалите строку", row=i, sheet=name)
                continue
            q, why = qty_value(r[cols["qty"]], allow_zero=True, max_qty=float("inf")) if not is_empty(r[cols["qty"]]) else (None, "нет количества")
            if q is None:
                self.find("BLOCK", "количество", f"{canon}: количество — {why}", row=i, sheet=name)
                continue
            rec = dict(row=i, n=n, ei=canon, qty=q)
            for k in ("name", "art", "unit", "place", "cat"):
                rec[k] = txt(r[cols[k]]) if k in cols else ""
            out["rows"].append(rec)
        if out["skipped"]:
            self.find("INFO", "структура", f"лист «{name}»: строк без ЕИ пропущено {out['skipped']} (итоги, подписи)", row=0, sheet=name)
        return out

    def read_side_sheets(self):
        self.side["stock"] = self.side_sheet(SHEET_STOCK, ("ei", "qty"))
        self.side["issues"] = self.side_sheet(SHEET_ISSUES, ("ei", "qty"))
        self.side["returns"] = self.side_sheet(SHEET_RETURNS, ("ei", "qty"))
        self.stock_by = {}
        for rec in self.side["stock"]["rows"]:
            self.stock_by.setdefault(rec["n"], []).append(rec)
        self.issued, self.returned = {}, {}
        for rec in self.side["issues"]["rows"]:
            self.issued[rec["n"]] = self.issued.get(rec["n"], 0.0) + rec["qty"]
        for rec in self.side["returns"]["rows"]:
            self.returned[rec["n"]] = self.returned.get(rec["n"], 0.0) + rec["qty"]
        self.moves = self.side["issues"]["present"] or self.side["returns"]["present"]

    # ------------------------------------------------------------ «1_Вставить_Заказы»
    @staticmethod
    def header_hits(r, off=0):
        texts = [txt(x).lower() for x in r]
        hits = 0
        for c, ks in HEADER_KEYS.items():
            j = c + off
            if 0 <= j < len(texts) and any(k in texts[j] for k in ks):
                hits += 1
        return hits

    def read_orders(self):
        rows = self.st.sheets.get(SHEET_ORDERS) or []
        start = 0
        for i, r in enumerate(rows[:15]):
            h0 = self.header_hits(r)
            best = max(((self.header_hits(r, off), off) for off in (-3, -2, -1, 1, 2, 3)), default=(0, 0))
            if best[0] >= 5 and best[0] > h0:
                off = best[1]
                self.find("BLOCK", "структура", f"данные вставлены со сдвигом на {abs(off)} колонк{'у' if abs(off) == 1 else 'и'} "
                          f"{'вправо' if off > 0 else 'влево'}: заголовок «Полное наименование» стоит в колонке {COLS[C_B + off] if 0 <= C_B + off < 28 else '?'}, "
                          "а должен быть в B. Выделите на старом листе «Заказы» колонки A:AB целиком и вставьте в ячейку A1", row=i + 1)
                self.header = i
                start = len(rows)
                break
            if h0 >= 5:
                # the first header: a header again further down is a row of the data (a second paste) and is shown there
                self.header = i
                start = i + 1
                break
        if self.header is None:
            self.find("INFO", "структура", "строка заголовка не найдена — данные читаются с первой строки", row=1)
        elif self.header > 0 and start < len(rows):
            above = sum(1 for r in rows[:self.header] if any(not is_empty(x) for x in r))
            if above:
                self.find("WARN", "структура", f"строк над заголовком не переносится: {above} (заголовок в строке {self.header + 1}) — если "
                          "это данные, вставьте лист «Заказы» в ячейку A1 целиком, с его первой строки", row=1)
        seen = {}
        for i in range(start, len(rows)):
            v = list(rows[i]) + [""] * (28 - len(rows[i]))
            row = SRow(i + 1, v[:28])
            if all(is_empty(x) for x in row.v):
                row.kind, row.fate = "EMPTY", "skipped"
            elif self.header_hits(row.v) >= 5:
                row.kind, row.fate = "HEADER", "skipped"
                row.norm.append("повтор строки заголовка — пропущена")
                self.find("WARN", "дубль строки", "повтор строки заголовка среди данных — если таблица вставлена дважды, удалите вторую копию "
                          "на этом листе", row)
            else:
                # the same row twice (a second paste of the table, a copied line): never transferred twice silently — a receipt
                # (F) twice would double its stock, so it stops; any other row is shown
                key = tuple(txt(x) for x in row.v)
                first = seen.setdefault(key, row.r)
                self.classify(row)
                if first != row.r:
                    if not is_empty(row.v[C_F]):
                        self.find("BLOCK", "дубль строки", f"строка полностью совпадает со строкой {first} — приход перенёсся бы дважды: "
                                  "удалите лишнюю копию на этом листе (старый файл не меняется); если это действительно два одинаковых "
                                  "поступления, отличите их (например, комментарием Z)", row)
                    else:
                        self.find("WARN", "дубль строки", f"строка полностью совпадает со строкой {first} — если это повторная вставка, "
                                  "удалите лишнюю копию на этом листе", row)
            self.rows.append(row)
        if not any(r.kind not in ("EMPTY", "HEADER") for r in self.rows):
            self.find("BLOCK", "структура", f"на листе «{SHEET_ORDERS}» нет данных — вставьте старый лист «Заказы» (A:AB) в ячейку A1", row=1)

    def classify(self, row):
        v = row.v
        for c in BLANK_COLS:
            if blank_like(v[c]):
                row.norm.append(f"{COLS[c]}: «{v[c].strip()}» считается пустым")
                v[c] = ""
        if not is_empty(v[C_A]):
            row.a, why, note = pos_no(v[C_A])
            if row.a is None:
                row.a_bad = why
            if note:
                row.norm.append(note)
        sp = special_source(v[C_L])
        if sp:
            row.kind = "SPECIAL"
        elif not is_empty(v[C_H]):
            row.kind = "POSITION"
        elif not is_empty(v[C_F]):
            row.kind = "DELIVERY"
        else:
            row.kind = "NOQTY"

    # ------------------------------------------------------------ orders (blocks) and deliveries
    def match_parent(self, row, positions):
        nk, ek = name_key(row.v[C_B]), article_key(row.v[C_E])
        for pos in reversed(positions):
            src = pos.row
            if row.a is not None and row.a != src.a:
                continue
            if nk or ek:
                if not ((nk and nk == name_key(src.v[C_B])) or (ek and ek == article_key(src.v[C_E]))):
                    continue
            return pos
        return None

    def start_block(self, row):
        self.blocks += 1
        self.cur = []
        self.block_first[self.blocks] = row

    def group(self):
        self.cur = []
        self.block_first = {}
        prev_a = None
        for row in self.rows:
            if row.kind in ("EMPTY", "HEADER"):
                continue
            if row.kind == "DELIVERY":
                parent = self.match_parent(row, self.cur) if self.blocks else None
                if parent is not None:
                    row.block = parent.block
                    row.pos = parent
                    parent.deliveries.append(row)
                    continue
                if self.s["no_order_qty"] == "fact":
                    row.kind = "POSITION"
                    row.h_rule = True
                    row.norm.append("строка без H стала отдельной позицией: H = F (правило)")
                else:
                    self.find("BLOCK", "структура", "поступление (F без заказанного количества H), но выше в этом заказе нет позиции с тем же "
                              "№ (A) и товаром — если это отдельная позиция, укажите H (или включите правило «строка с F без H — "
                              "позиция с H = F»)", row)
                    row.fate = "skipped"
                    continue
            if row.kind == "NOQTY":
                if self.s["empty_rows"] == "skip":
                    row.fate = "skipped"
                    row.norm.append("строка без количеств пропущена (правило)")
                    self.find("WARN", "структура", "строка без заказанного (H) и полученного (F) количества — пропущена по правилу", row)
                else:
                    self.find("BLOCK", "структура", "нет ни заказанного (H), ни фактического (F) количества — это не позиция и не "
                              "поступление: заполните H или удалите строку (или включите правило «пропускать строки без количеств»)", row)
                    row.fate = "skipped"
                continue
            # a position or a special receipt: A = 1 starts the next order
            if self.blocks == 0 or row.a == 1:
                if self.blocks == 0 and row.a not in (None, 1):
                    self.find("WARN", "структура", f"первый заказ начинается с позиции {row.a}, а не с 1", row)
                self.start_block(row)
            elif prev_a is not None and row.a is not None and row.a != prev_a + 1:
                self.find("INFO", "структура", f"нумерация позиций: после {prev_a} идёт {row.a}", row)
            prev_a = row.a if row.a is not None else prev_a
            row.block = self.blocks
            if row.kind == "POSITION":
                pos = Position(row, self.blocks)
                row.pos = pos
                self.positions.append(pos)
                self.cur.append(pos)
        for b, first in self.block_first.items():
            if first.kind == "SPECIAL" and any(p.block == b for p in self.positions):
                self.find("WARN", "структура", "заказ начинается со строки специального прихода — в новой WMS она уходит в «Наличие», "
                          "и над следующими позициями этого заказа не будет линии-разделителя", first)

    # ------------------------------------------------------------ values of one row
    def norm_values(self, raw, src_of, role, pos=None):
        """the input values of a row of «Заказы» (col → value; None — empty) by the rules of WMS; findings go to the
        staging row each value came from. role: OPEN (a position without a receipt), RECEIPT (a position with its first
        receipt), DELIVERY, SPECIAL"""
        out = {}
        s = self.s
        # the unit a quantity may be written with («60 шт»): of the row, of its position, the default one
        unit_raw = raw[C_I] or (pos.row.v[C_I] if pos is not None else "") or s["default_unit"]

        def bad(c, text, kind="значение"):
            self.find("BLOCK", kind, f"{COLS[c]} «{COL_NAMES[c]}»: {text}", src_of[c])

        def note(c, text):
            src_of[c].norm.append(f"{COLS[c]}: {text}")

        for c in TEXT_COLS:
            out[c] = txt(raw[c]) or None
        # A
        if role == "DELIVERY":
            out[C_A] = str(pos.row.a) if pos.row.a is not None else None
        elif role != "SPECIAL":
            if is_empty(raw[C_A]):
                bad(C_A, "не указан № позиции", "обязательное поле")
            elif getattr(src_of[C_A], "a_bad", "") and src_of[C_A].a is None:
                bad(C_A, src_of[C_A].a_bad)
            out[C_A] = str(src_of[C_A].a) if src_of[C_A].a is not None else None
        else:
            out[C_A] = txt(raw[C_A]) or None
        # B
        if not out[C_B] and role != "DELIVERY":
            bad(C_B, "не указано наименование", "обязательное поле")
        # H
        if role in ("OPEN", "RECEIPT"):
            if is_empty(raw[C_H]):
                if getattr(src_of[C_A], "h_rule", False):
                    out[C_H] = None           # filled below from F
                else:
                    bad(C_H, "не указано заказанное количество", "обязательное поле")
            else:
                v, cut = strip_unit(raw[C_H], unit_raw)
                if cut:
                    note(C_H, f"«{raw[C_H]}» → {v}")
                q, why = qty_value(v, max_qty=self.max_qty)
                if q is None:
                    bad(C_H, why, "количество")
                out[C_H] = q
        else:
            out[C_H] = None
        # F G (H above): a quantity written with the unit of its row is its number
        for c in (C_F, C_G):
            if is_empty(raw[c]):
                out[c] = None
            else:
                v, cut = strip_unit(raw[c], unit_raw)
                if cut:
                    note(c, f"«{raw[c]}» → {v}")
                q, why = qty_value(v, max_qty=self.max_qty)
                if q is None:
                    bad(c, why, "количество")
                out[c] = q
        if role in ("RECEIPT", "DELIVERY", "SPECIAL") and out[C_F] is None and is_empty(raw[C_F]):
            bad(C_F, "не указано фактическое количество", "обязательное поле")
        if getattr(src_of[C_A], "h_rule", False) and role in ("OPEN", "RECEIPT") and out[C_H] is None:
            out[C_H] = out[C_F]
            note(C_H, f"H = F = {num_text(out[C_F])} (правило)")
        # I (a delivery takes the unit of its position: a different one is refused in check_delivery_identity)
        if role == "DELIVERY":
            out[C_I] = None
        elif is_empty(raw[C_I]):
            if s["default_unit"]:
                u, _ = unit_value(s["default_unit"])
                out[C_I] = u
                note(C_I, f"единица не указана → «{u}» (правило)")
            else:
                out[C_I] = None
                if role in ("RECEIPT", "SPECIAL"):
                    bad(C_I, "не указана единица измерения (укажите или задайте единицу по умолчанию в «Настройках»)", "обязательное поле")
                elif role == "OPEN":
                    self.find("WARN", "единица", "I: единица измерения не указана — понадобится при приходе", src_of[C_I])
        else:
            u, n = unit_value(raw[C_I])
            out[C_I] = u
            if n:
                note(C_I, n)
            if u and u.lower() not in self.known_units:
                self.find("WARN", "единица", f"I: неизвестная единица «{u}» (нет в списке «Известные единицы»)", src_of[C_I])
        # J
        if is_empty(raw[C_J]):
            out[C_J] = None
        else:
            x, why, n = price_value(raw[C_J], s["price_round"] == "round")
            if x is None:
                bad(C_J, why + " (или включите правило округления цены)" if "знаков" in why else why, "цена")
            out[C_J] = x
            if n:
                note(C_J, n)
        # K, AA: numbers when they are numbers
        for c in (C_K, C_AA):
            if is_empty(raw[c]):
                out[c] = None
            else:
                x = number_value(raw[c])
                if x is None:
                    out[c] = txt(raw[c])
                    self.find("WARN", "значение", f"{COLS[c]} «{COL_NAMES[c]}»: «{txt(raw[c])}» не число — перенесено как текст", src_of[c])
                else:
                    out[c] = x
                    if not isinstance(raw[c], float):
                        note(c, f"«{txt(raw[c])}» → {txt(x)}")
        # dates: N O are facts of a receipt (a wrong one stops the transfer); P Q only describe the order — what is not a
        # date there is kept in the comment Z instead of being lost or stopping the transfer
        for c in (C_N, C_O, C_P, C_Q):
            if is_empty(raw[c]):
                out[c] = None
                continue
            x, why, n = date_value(raw[c])
            if x is None and c in (C_P, C_Q):
                keep = f"{COL_NAMES[c]}: {txt(raw[c])}"
                out[C_Z] = (out[C_Z] + "; " if out[C_Z] else "") + keep
                self.find("WARN", "дата", f"{COLS[c]}: «{txt(raw[c])}» не дата ({why}) — записано в комментарий (Z)", src_of[c])
                out[c] = None
                continue
            if x is None:
                bad(c, why, "дата")
            out[c] = x
            if n:
                note(c, n)
        if role in ("RECEIPT", "DELIVERY", "SPECIAL") and out[C_N] is None and is_empty(raw[C_N]):
            if s["no_receipt_date"] == "doc" and out[C_O] is not None:
                out[C_N] = out[C_O]
                note(C_N, f"дата поступления не указана → дата документа {date_text(out[C_O])} (правило)")
            else:
                bad(C_N, "не указана дата поступления (или включите правило «дата поступления = дата документа»)", "обязательное поле")
        if role in ("RECEIPT", "DELIVERY") and out[C_N] and out[C_N] > self.today:
            self.find("WARN", "дата", f"N: дата поступления {date_text(out[C_N])} в будущем", src_of[C_N])
        # U
        if not out[C_U]:
            if role in ("RECEIPT", "DELIVERY", "SPECIAL"):
                if s["default_place"]:
                    out[C_U] = s["default_place"].strip()
                    note(C_U, f"место не указано → «{out[C_U]}» (правило)")
                else:
                    bad(C_U, "не указано место хранения (укажите или задайте место по умолчанию в «Настройках»)", "обязательное поле")
        if out[C_U] and self.known_places and out[C_U] not in self.known_places:
            self.find("WARN", "место", f"U: неизвестное место «{out[C_U]}» (нет в списке «Известные места»)", src_of[C_U])
        # a special supplier in L is not an order of «Заказы»
        return out

    # ------------------------------------------------------------ positions, their deliveries, the merge of the first delivery
    def build_positions(self):
        for pos in self.positions:
            row = pos.row
            raw = list(row.v)
            src_of = [row] * 28
            first = pos.deliveries[0] if pos.deliveries else None
            if is_empty(raw[C_F]) and first is not None:
                # WMS: the first receipt of a position is made in its source row — the first delivery row becomes that
                # receipt (its values move into the source row; nothing is lost: the row is listed as merged)
                pos.merged = first
                pos.deliveries = pos.deliveries[1:]
                for c in (C_F, C_G, C_C, C_J, C_N, C_O, C_U, C_V, C_X):
                    if is_empty(first.v[c]):
                        continue
                    if not is_empty(raw[c]) and txt(raw[c]) != txt(first.v[c]):
                        if c == C_V:
                            a, b = ei_value(raw[c])[1], ei_value(first.v[c])[1]
                            if a is None or a != b:
                                # two EIs for one receipt: taking either would silently drop the other
                                self.find("BLOCK", "ЕИ", f"V: у позиции указан ЕИ «{txt(raw[c])}», а у её первого поступления (строка "
                                          f"{first.r}) — «{txt(first.v[c])}»: это один приход, неясно, какой ЕИ настоящий — оставьте один",
                                          row)
                        elif c == C_X:
                            self.find("WARN", "остаток", f"X: у позиции указано наличие «{txt(raw[c])}», а у её первого поступления "
                                      f"(строка {first.r}) — «{txt(first.v[c])}»: переносится наличие поступления (у строки позиции "
                                      "без F своего прихода нет)", row)
                        else:
                            row.norm.append(f"{COLS[c]}: значение позиции «{txt(raw[c])}» заменено значением поступления «{txt(first.v[c])}»")
                    raw[c] = first.v[c]
                    src_of[c] = first
                for c in (C_K, C_AA):
                    if is_empty(raw[c]) and not is_empty(first.v[c]):
                        raw[c], src_of[c] = first.v[c], first
                if not is_empty(first.v[C_Z]):
                    raw[C_Z] = (txt(raw[C_Z]) + "; " if not is_empty(raw[C_Z]) else "") + txt(first.v[C_Z])
                first.fate = "merged"
                first.norm.append(f"первое поступление позиции перенесено в её строку (строка {row.r}): так позиция устроена в WMS 0.7")
                self.check_delivery_identity(first, pos)
            role = "RECEIPT" if not is_empty(raw[C_F]) else "OPEN"
            pos.raw, pos.src_of = raw, src_of
            pos.vals = self.norm_values(raw, src_of, role)
            if role == "OPEN":
                if not is_empty(raw[C_V]):
                    self.find("BLOCK", "ЕИ", f"V: указан ЕИ «{txt(raw[C_V])}», но фактическое количество F пусто — у ЕИ без прихода нет "
                              "остатка: укажите F (и дату, место) или очистите V", src_of[C_V])
                if not is_empty(raw[C_X]):
                    x = number_value(raw[C_X])
                    if x is None or abs(x) > EPS:
                        self.find("WARN", "остаток", f"X: наличие «{txt(raw[C_X])}» у позиции без прихода — не переносится", src_of[C_X])
            row.fate = "row"
            for d in pos.deliveries:
                d.fate = "row"
                self.check_delivery_identity(d, pos)
                raw_d = list(d.v)
                dv = self.norm_values(raw_d, [d] * 28, "DELIVERY", pos)
                # the identity of the position comes from its source row (as «Ещё поступление» copies it)
                for c in IDENTITY_COLS:
                    dv[c] = pos.vals.get(c)
                dv[C_H] = None
                d.vals = dv

    def check_delivery_identity(self, d, pos):
        diff = []
        for c in IDENTITY_COLS:
            if c == C_A:
                continue
            a, b = txt(d.v[c]), txt(pos.row.v[c])
            if a and a != b:
                if c == C_I:
                    ua, _ = unit_value(a)
                    ub, _ = unit_value(b)
                    if ua != ub:
                        self.find("BLOCK", "единица", f"I: единица поступления «{a}» не совпадает с единицей позиции «{b}» (строка {pos.row.r}) — "
                                  "одна позиция принимается в одной единице", d)
                    continue
                diff.append(f"{COLS[c]} «{a}» → «{b}»")
        if diff:
            d.norm.append("реквизиты позиции взяты из её строки " + str(pos.row.r) + ": " + ", ".join(diff[:6]))

    # ------------------------------------------------------------ the EIs of the transfer
    def collect_items(self):
        for pos in self.positions:
            if pos.vals.get(C_F) is not None or not is_empty(pos.raw[C_F]):
                it = Item("SRC")
                it.pos, it.vals = pos, pos.vals
                it.rows = [pos.row] + ([pos.merged] if pos.merged else [])
                it.ei_raw = pos.raw[C_V]
                it.ei_row = pos.src_of[C_V]
                it.xraw = pos.raw[C_X]
                it.x_row = pos.src_of[C_X]
                pos.receipts.append(it)
                self.items.append(it)
            for d in pos.deliveries:
                it = Item("ADD")
                it.pos, it.vals = pos, d.vals
                it.rows = [d]
                it.ei_raw, it.ei_row = d.v[C_V], d
                it.xraw, it.x_row = d.v[C_X], d
                pos.receipts.append(it)
                self.items.append(it)
        for it in self.items:
            v = it.vals
            it.fact = v.get(C_F) or 0.0
            it.place = v.get(C_U) or ""
            it.name, it.art, it.unit, it.cat = it.pos.vals.get(C_B) or "", it.pos.vals.get(C_E) or "", it.pos.vals.get(C_I) or "", it.pos.vals.get(C_S) or ""
            it.nodoc = 0 if (v.get(C_C) and v.get(C_O)) else 1
        # special receipts: parts (Детали) grouped by their article — one article, one EI; the others one EI per row
        groups = {}
        for row in self.rows:
            if row.kind != "SPECIAL" or row.fate == "skipped":
                continue
            row.fate = "special"
            vals = self.norm_values(list(row.v), [row] * 28, "SPECIAL")
            stype = special_source(row.v[C_L])
            key = article_key(row.v[C_E]) if stype == PART else ""
            if stype == PART and key:
                it = groups.get(key)
                if it is None:
                    it = Item("MIGRATE")
                    groups[key] = it
                    self.items.append(it)
                    it.stype, it.key, it.vals = PART, key, vals
                it.rows.append(row)
                it.fact += vals.get(C_F) or 0.0
                row.special = it
                continue
            it = Item("MIGRATE")
            it.rows, it.vals, it.stype = [row], vals, stype
            it.fact = vals.get(C_F) or 0.0
            if stype == PART:
                it.stype, it.review = "Иной приход", True
                self.find("WARN", "деталь", "деталь без артикула (E) — переносится как ЕИ «Требует разбора» (тип «Иной приход»); "
                          "после переноса укажите артикул и разберите её", row)
            row.special = it
            self.items.append(it)
        for it in self.items:
            if it.kind != "MIGRATE":
                continue
            first = it.rows[0]
            it.name, it.art, it.unit, it.cat = it.vals.get(C_B) or "", it.vals.get(C_E) or "", it.vals.get(C_I) or "", it.vals.get(C_S) or ""
            it.place = it.vals.get(C_U) or ""
            it.state = EI_REVIEW if it.stype == "Иной приход" else EI_ACTIVE
            evs = [(r, r.v[C_V]) for r in it.rows if not is_empty(r.v[C_V])]
            it.ei_raw, it.ei_row = (evs[0][1], evs[0][0]) if evs else ("", first)
            it.xs = [(r, r.v[C_X]) for r in it.rows if not is_empty(r.v[C_X])]
            it.nodoc = 0
            if len(it.rows) > 1:
                names = {name_key(r.v[C_B]) for r in it.rows}
                units = {unit_value(r.v[C_I])[0] for r in it.rows if not is_empty(r.v[C_I])}
                if len(units) > 1:
                    self.find("BLOCK", "деталь", f"деталь «{it.art}»: строки {', '.join(str(r.r) for r in it.rows)} в разных единицах "
                              f"({', '.join(sorted(units))}) — одна деталь — один ЕИ в одной единице", first)
                if len(names) > 1:
                    self.find("WARN", "деталь", f"деталь «{it.art}»: в строках {', '.join(str(r.r) for r in it.rows)} разные наименования — "
                              f"взято «{it.name}»", first)
                for r in it.rows[1:]:
                    r.norm.append(f"повторный приход детали «{it.art}» — пополнение того же ЕИ (строка {first.r})")
                distinct = {}
                for i, (r, ev) in enumerate(evs):
                    n, canon, why = ei_value(ev)
                    if n is not None:
                        distinct.setdefault(canon, []).append(r.r)
                    elif i:
                        # the EI of the group is the first one (checked with the others); a wrong number further down must not vanish
                        self.find("BLOCK", "ЕИ", f"V: {why}", r)
                if len(distinct) > 1:
                    self.find("BLOCK", "деталь", f"артикул детали «{it.art}» у разных ЕИ: " + "; ".join(f"{k} (строки {', '.join(map(str, v))})"
                              for k, v in distinct.items()) + " — в новой WMS одна деталь — один ЕИ; решите, какой ЕИ настоящий", first)
                    it.ei_conflict = True

    def assign_eis(self):
        claims = {}
        self.max_seen = 0
        for it in self.items:
            if is_empty(it.ei_raw):
                continue
            n, canon, why = ei_value(it.ei_raw)
            if n is None:
                self.find("BLOCK", "ЕИ", f"V: {why}", it.ei_row)
                continue
            it.n, it.ei = n, canon
            it.ei_src = "NORM" if why else "OLD"
            if why:
                it.ei_row.norm.append(f"V: ЕИ «{txt(it.ei_raw)}» → {canon}")
            claims.setdefault(n, []).append(it)
        for row in self.rows:
            if row.kind in ("EMPTY", "HEADER") or is_empty(row.v[C_V]):
                continue
            n, canon, why = ei_value(row.v[C_V])
            if n is not None:
                self.max_seen = max(self.max_seen, n)
        for side in ("stock", "issues", "returns"):
            for rec in self.side[side]["rows"]:
                self.max_seen = max(self.max_seen, rec["n"])
        self.dup_eis = 0
        for n, lst in claims.items():
            if len(lst) < 2:
                continue
            self.dup_eis += 1
            rows = sorted({r.r for it in lst for r in it.rows})
            keys = {(name_key(it.name), article_key(it.art), it.unit) for it in lst}
            why = ("разные товары: " + "; ".join(sorted({f"«{it.name}»" for it in lst}))) if len(keys) > 1 else \
                "один товар, но в новой WMS каждое поступление — отдельный ЕИ"
            for it in lst:
                self.find("BLOCK", "дубль ЕИ", f"V: {it.ei} указан в строках {', '.join(map(str, rows))} — {why}; исправьте номер в одной из "
                          "строк", it.ei_row)
        nxt = self.max_seen + 1
        for it in sorted(self.items, key=lambda x: x.rows[0].r):
            if it.n is None and is_empty(it.ei_raw):
                it.n, it.ei, it.ei_src = nxt, f"ЕИ-{nxt:08d}", "NEW"
                nxt += 1
        # 1B-only EIs: stock without an order row
        in_orders = {it.n for it in self.items if it.n}
        self.stock_only = []
        for n, recs in sorted(self.stock_by.items()):
            if n in in_orders:
                continue
            q = sum(r["qty"] for r in recs)
            if q <= EPS:
                continue
            rule = self.s["stock_only"]
            if rule == "skip":
                self.find("WARN", "остаток", f"ЕИ-{n:08d} с остатком {qty_text(q)} есть только в «Наличие» — не переносится (правило)",
                          row=recs[0]["row"], sheet=SHEET_STOCK)
            elif rule == "old":
                r0 = recs[0]
                if not r0["name"] or not r0["unit"]:
                    self.find("BLOCK", "обязательное поле", f"ЕИ-{n:08d}: для переноса как «Старый склад» нужны наименование и единица",
                              row=r0["row"], sheet=SHEET_STOCK)
                    continue
                it = Item("MIGRATE")
                it.stype, it.n, it.ei, it.ei_src = "Старый склад", n, f"ЕИ-{n:08d}", "OLD"
                it.name, it.art, it.unit, it.cat = r0["name"], r0["art"], unit_value(r0["unit"])[0], r0["cat"]
                it.place = r0["place"] or self.s["default_place"]
                it.bal, it.bal_src, it.fact = q, "1B", q
                it.from_stock = recs
                it.rows = []
                if not it.place:
                    self.find("BLOCK", "обязательное поле", f"ЕИ-{n:08d}: не указано место (или задайте место по умолчанию)",
                              row=r0["row"], sheet=SHEET_STOCK)
                self.stock_only.append(it)
                self.items.append(it)
                self.find("INFO", "остаток", f"ЕИ-{n:08d} есть только в «Наличие» — переносится как «Старый склад» с остатком {qty_text(q)} (правило)",
                          row=r0["row"], sheet=SHEET_STOCK)
            else:
                self.find("BLOCK", "остаток", f"ЕИ-{n:08d} с остатком {qty_text(q)} есть в «Наличие», но не в «Заказы» — такой остаток "
                          "потеряется: включите правило «ЕИ только из Наличия» (перенести как «Старый склад» или не переносить)",
                          row=recs[0]["row"], sheet=SHEET_STOCK)
        self.next_ei = max([it.n for it in self.items if it.n] + [0]) + 1
        # old numbers that stay behind (issues and returns of EIs that are not transferred, stock rows not transferred): the core
        # gives new EIs from NEXT_EI — say so when one of them may get such a number
        kept = {it.n for it in self.items if it.n}
        left = ({rec["n"] for side in ("issues", "returns") for rec in self.side[side]["rows"]} | set(self.stock_by)) - kept
        top = max(left | {0})
        if top >= self.next_ei:
            self.find("WARN", "ЕИ", f"старые ЕИ до ЕИ-{top:08d} есть только в выдачах, возвратах или «Наличие» и не переносятся — после "
                      f"переноса NEXT_EI будет {self.next_ei}: новый ЕИ может получить номер старого (на старых этикетках)", row=0,
                      sheet=SHEET_STOCK if top in self.stock_by else SHEET_ISSUES)

    # ------------------------------------------------------------ the current balance of every EI (task §6)
    def balances(self):
        self.suspicious = 0
        self.unknown = 0
        for it in self.items:
            if it.bal is not None:
                continue
            first = it.rows[0] if it.rows else None
            xs = it.xs if it.kind == "MIGRATE" else ([(it.x_row, it.xraw)] if not is_empty(it.xraw) else [])
            fact = it.fact
            why_x = ""
            if xs:
                vals, bad = [], []
                for r, raw in xs:
                    raw, cut = strip_unit(raw, it.unit)
                    if cut:
                        r.norm.append(f"X: «{txt(dict(xs)[r])}» → {raw}")
                    q, why = qty_value(raw, allow_zero=True, max_qty=float("inf"))
                    (vals if q is not None else bad).append((r, q if q is not None else raw, why))
                distinct = sorted({round(q, 3) for _, q, _ in vals})
                if bad:
                    why_x = f"X «{txt(bad[0][1])}» — {bad[0][2]}"
                elif len(distinct) > 1:
                    why_x = "в строках разные X: " + ", ".join(qty_text(x) for x in distinct)
                elif distinct[0] > fact + EPS:
                    why_x = f"X {qty_text(distinct[0])} больше полученного {qty_text(fact)}"
                else:
                    it.bal, it.bal_src = distinct[0], "X"
            place_1b = ""
            if it.bal is None and it.n and it.n in self.stock_by and it.ei_src != "NEW":
                recs = self.stock_by[it.n]
                q = sum(r["qty"] for r in recs)
                names = {name_key(r["name"]) for r in recs if r["name"]}
                arts = {article_key(r["art"]) for r in recs if r["art"]}
                compatible = (not names or name_key(it.name) in names) or (arts and article_key(it.art) in arts)
                if not compatible:
                    self.find("WARN", "остаток", f"{it.ei}: в «Наличие» (1B) записан как «{recs[0]['name']}», а в «Заказы» — «{it.name}»: "
                              "остаток из 1B не использован", first)
                elif q > fact + EPS:
                    self.find("WARN", "остаток", f"{it.ei}: остаток в «Наличие» (1B) {qty_text(q)} больше полученного {qty_text(fact)} — не использован", first)
                else:
                    it.bal, it.bal_src = q, "1B"
            if it.n and it.n in self.stock_by and it.ei_src != "NEW":
                places = {r["place"] for r in self.stock_by[it.n] if r["place"]}
                if len(places) == 1:
                    place_1b = places.pop()
                if it.bal_src == "X":
                    q = sum(r["qty"] for r in self.stock_by[it.n])
                    if abs(q - it.bal) > EPS:
                        self.find("WARN", "остаток", f"{it.ei}: остаток X {qty_text(it.bal)}, а в «Наличие» (1B) {qty_text(q)} — взят X (приоритет)", first)
            if it.bal is None and self.moves and it.n and it.ei_src != "NEW":
                c = round(fact - self.issued.get(it.n, 0.0) + self.returned.get(it.n, 0.0), 3)
                if -EPS <= c <= fact + EPS:
                    it.bal, it.bal_src = max(c, 0.0), "1C/1D"
                else:
                    self.find("WARN", "остаток", f"{it.ei}: по выдачам и возвратам остаток {qty_text(c) if c >= 0 else '−' + qty_text(-c)} — вне "
                              f"0…{qty_text(fact)}, расчёт не использован", first)
            if why_x and it.bal is not None:
                self.suspicious += 1
                self.find("WARN", "остаток", f"{it.ei}: {why_x} — остаток взят из {'листа «Наличие»' if it.bal_src == '1B' else 'выдач и возвратов'}", first)
            if it.bal is None:
                if why_x:
                    self.suspicious += 1
                    self.find("BLOCK", "остаток", f"{it.ei}: подозрительный остаток — {why_x}; исправьте X (или вставьте «Наличие» 1B)", first)
                    continue
                rule = self.s["unknown_balance"]
                self.unknown += 1
                if rule == "zero":
                    it.bal, it.bal_src = 0.0, "RULE0"
                    self.find("WARN", "остаток", f"{it.ei}: Требуется проверить остаток — принят 0 по правилу; после инвентаризации исправьте "
                              "его корректировкой", first)
                elif rule == "fact":
                    it.bal, it.bal_src = fact, "RULEF"
                    self.find("WARN", "остаток", f"{it.ei}: Требуется проверить остаток — принято всё полученное ({qty_text(fact)}) по правилу", first)
                else:
                    self.find("BLOCK", "остаток", f"{it.ei}: Требуется проверить остаток — нет X «Наличие», нет строки в листе «Наличие» (1B), "
                              "нет выдач и возвратов (1C, 1D); укажите остаток или выберите правило в «Настройках»", first)
                    continue
            it.place_now = place_1b if place_1b and place_1b != it.place else ""
            if it.place_now and first is not None:
                first.norm.append(f"текущее место {it.ei} по листу «Наличие»: «{it.place_now}» (место прихода «{it.place}»)")

    # ------------------------------------------------------------ statuses and cancellations (task §4, §7)
    def statuses(self):
        self.status_diff = 0
        self.status_critical = 0
        self.pairs = {}
        for pos in self.positions:
            v = pos.vals
            pos.ordq = v.get(C_H) or 0.0
            pos.rcv = round(sum(it.fact for it in pos.receipts), 3)
            pos.nodoc = sum(it.nodoc for it in pos.receipts)
            pos.edate = v.get(C_Q) or 0.0
            pos.old, pos.old_raw = old_status(pos.row.v[C_W])
            row = pos.row
            if pos.old == OS_CANCELLED:
                if pos.rcv <= EPS:
                    pos.cancel = "ORDER"
                elif pos.rcv < pos.ordq - EPS:
                    pos.cancel = "REST"
                    row.norm.append("W: «Отменено» при частичном приходе → «Частично получено / остаток отменён»")
                else:
                    self.status_critical += 1
                    self.find("BLOCK", "статус", f"W: старый статус «{pos.old_raw}», но позиция получена полностью ({num_text(pos.rcv)} из "
                              f"{num_text(pos.ordq)}) — критическое противоречие: исправьте статус или количество", row)
            elif pos.old == OS_REST_CANCELLED and EPS < pos.rcv < pos.ordq - EPS:
                pos.cancel = "REST"
            if pos.old in RECEIVED_FAMILY and pos.rcv <= EPS:
                self.status_critical += 1
                self.find("BLOCK", "статус", f"W: старый статус «{pos.old_raw}», но фактическое количество F не указано — критическое "
                          "противоречие: укажите F (дату, место, ЕИ) или исправьте статус", row)
            pos.new = position_status(pos.ordq, pos.rcv, pos.nodoc, pos.cancel, pos.edate, self.today)
            key = (pos.old_raw or "(пусто)", pos.new)
            self.pairs[key] = self.pairs.get(key, 0) + 1
            if pos.old is None and pos.old_raw:
                self.find("WARN", "статус", f"W: неизвестный старый статус «{pos.old_raw}» — новый статус «{pos.new}»", row)
            elif pos.old is not None and pos.old != pos.new:
                self.status_diff += 1
                pos.status_note = self.status_reason(pos)
                if not any(m[0] == "BLOCK" and m[1].startswith("W:") for m in row.msgs):
                    lvl = "WARN" if (pos.old in (OS_WAITING, OS_OVERDUE) and pos.rcv > EPS) else "INFO"
                    self.find(lvl, "статус", f"W: «{pos.old_raw}» → «{pos.new}» ({pos.status_note})", row)
            if pos.old is not None and pos.old != pos.new:
                row.recalc = True

    def status_reason(self, pos):
        o, n = pos.old, pos.new
        if o in (OS_WAITING, OS_OVERDUE) and pos.rcv > EPS:
            return f"в старой таблице не учтён приход {num_text(pos.rcv)}"
        if n in (OS_OVERDUE, OS_PARTIAL_OVERDUE) and o in (OS_WAITING, OS_PARTIAL):
            return f"ожидаемая дата {date_text(pos.edate)} прошла"
        if o in (OS_OVERDUE, OS_PARTIAL_OVERDUE) and n in (OS_WAITING, OS_PARTIAL):
            return "ожидаемая дата не прошла или не указана" if pos.edate else "ожидаемая дата (Q) не указана"
        if o == OS_RECEIVED and n == OS_NODOCS:
            return "у прихода нет документа (C) или его даты (O)"
        if o == OS_NODOCS and n == OS_RECEIVED:
            return "документы указаны"
        if n in (OS_RECEIVED, OS_NODOCS):
            return f"получено {num_text(pos.rcv)} из {num_text(pos.ordq)}"
        if n in (OS_PARTIAL, OS_PARTIAL_OVERDUE):
            return f"получено {num_text(pos.rcv)} из {num_text(pos.ordq)}"
        return f"заказано {num_text(pos.ordq)}, получено {num_text(pos.rcv)}"

    # ------------------------------------------------------------ the rows of the new «Заказы», the operations, the OLIDs
    def layout(self):
        t = 1
        self.targets = []
        olid = 0
        ops = []
        by_block = {}
        for pos in self.positions:
            by_block.setdefault(pos.block, []).append(pos)
        for b in sorted(by_block):
            for pos in by_block[b]:
                pos.t = t
                pos.row.t = t
                if pos.merged:
                    pos.merged.t = t
                self.targets.append(("POS", pos, None))
                t += 1
                for d in pos.deliveries:
                    d.t = t
                    self.targets.append(("DEL", pos, d))
                    t += 1
        for pos in self.positions:
            src_it = pos.receipts[0] if pos.receipts and pos.receipts[0].kind == "SRC" else None
            if pos.cancel == "ORDER":
                olid += 1
                pos.olid = olid
                ops.append(dict(op="cancel", t=pos.t, pos=pos))
            elif src_it is not None:
                olid += 1
                pos.olid = olid
                ops.append(dict(op="legacy", t=pos.t, item=src_it))
                for it in pos.receipts[1:]:
                    ops.append(dict(op="legacy_add", t=it.rows[0].t, src=pos.t, item=it))
                if pos.cancel == "REST":
                    ops.append(dict(op="cancel_rest", t=pos.t, pos=pos))
            for it in pos.receipts:
                it.t = it.rows[0].t
                it.olid = pos.olid
        for it in self.items:
            if it.kind == "MIGRATE":
                ops.append(dict(op="migrate", item=it))
        self.ops = ops
        self.next_ol = olid + 1
        for it in self.items:
            if it.kind == "MIGRATE":
                rows = [r.r for r in it.rows]
                it.mark = ("Заказы, стр. " + ", ".join(map(str, rows))) if rows else f"Наличие (1B), стр. {it.from_stock[0]['row']}"

    # ------------------------------------------------------------ the candidate of the release (when it is given)
    def check_candidate(self):
        c = self.candidate
        if not c:
            return
        for p in c.get("problems", []):
            self.find("BLOCK", "книга-кандидат", p, row=0, sheet="книга-кандидат")

    # ------------------------------------------------------------ the summary (sheet «2_Проверка») and the row results
    def summarize(self):
        data = [r for r in self.rows if r.kind not in ("EMPTY", "HEADER")]
        for r in data:
            if r.level == "BLOCK":
                r.result = RESULT_BLOCK
            elif r.level == "WARN":
                r.result = RESULT_WARN
            elif r.fate == "skipped":
                r.result = RESULT_SKIP
            else:
                r.result = RESULT_OK
        for r in self.rows:
            if r.kind == "HEADER":
                r.result = RESULT_SKIP
        B, W = self.blockers(), self.warnings()
        st = {}
        for pos in self.positions:
            st[pos.new] = st.get(pos.new, 0) + 1
        kinds = {}
        for f in self.F:
            if f["level"] in ("BLOCK", "WARN"):
                kinds[f["kind"]] = kinds.get(f["kind"], 0) + 1
        eis = [it for it in self.items if it.n]
        m = []

        def add(key, label, value, level="INFO", hint=""):
            m.append((key, label, value, level, hint))

        n_open = sum(st.get(k, 0) for k in (OS_WAITING, OS_OVERDUE, OS_PARTIAL, OS_PARTIAL_OVERDUE))
        add("rows_total", "Всего строк данных", len(data))
        add("orders", "Заказов (блоков, A = 1 начинает заказ)", len({p.block for p in self.positions}))
        add("positions", "Позиций заказов", len(self.positions))
        add("open", "Открытых позиций (ожидается / частично получено)", n_open)
        add("received_full", "Полностью полученных", st.get(OS_RECEIVED, 0) + st.get(OS_NODOCS, 0))
        add("partial", "Частично полученных", st.get(OS_PARTIAL, 0) + st.get(OS_PARTIAL_OVERDUE, 0))
        add("waiting", "Ожидающихся (ничего не получено)", st.get(OS_WAITING, 0) + st.get(OS_OVERDUE, 0))
        add("overdue", "Просроченных", st.get(OS_OVERDUE, 0) + st.get(OS_PARTIAL_OVERDUE, 0))
        add("cancelled", "Отменённых (и с отменённым остатком)", st.get(OS_CANCELLED, 0) + st.get(OS_REST_CANCELLED, 0))
        add("nodocs", "Получено без документов", st.get(OS_NODOCS, 0))
        add("deliveries", "Дополнительных поступлений (строки)", sum(len(p.deliveries) for p in self.positions))
        add("merged", "Поступлений, перенесённых в строку позиции", sum(1 for p in self.positions if p.merged))
        add("specials", "Специальных приходов (ЕИ: Офис, Производство, Детали, …)", sum(1 for it in self.items if it.kind == "MIGRATE"))
        add("eis", "ЕИ после переноса", len(eis))
        add("eis_old", "из них с прежним номером", sum(1 for it in eis if it.ei_src in ("OLD", "NORM")))
        add("eis_norm", "из них номер записан иначе (427 → ЕИ-00000427)", sum(1 for it in eis if it.ei_src == "NORM"))
        add("eis_new", "новых номеров (в строке не было ЕИ)", sum(1 for it in eis if it.ei_src == "NEW"))
        mx = max([it.n for it in eis if it.ei_src != "NEW"] + [0])
        add("max_ei", "Максимальный старый ЕИ", f"ЕИ-{mx:08d}" if mx else "—")
        add("next_ei", "NEXT_EI после переноса (следующий новый ЕИ)", self.next_ei)
        add("with_stock", "ЕИ с остатком больше 0", sum(1 for it in eis if (it.bal or 0) > EPS))
        add("stock_total", "Остаток по всем ЕИ (сумма, в разных единицах)", qty_text(sum(it.bal or 0 for it in eis)))
        add("bal_x", "Остаток из X «Наличие»", sum(1 for it in eis if it.bal_src == "X"))
        add("bal_1b", "Остаток из листа «Наличие» (1B)", sum(1 for it in eis if it.bal_src == "1B"))
        add("bal_moves", "Остаток по выдачам и возвратам (1C, 1D)", sum(1 for it in eis if it.bal_src == "1C/1D"))
        add("bal_rule", "Остаток по правилу (0 или F)", sum(1 for it in eis if it.bal_src in ("RULE0", "RULEF")),
            "WARN" if any(it.bal_src in ("RULE0", "RULEF") for it in eis) else "OK")
        add("blockers", "Блокирующих замечаний (красный)", len(B), "BLOCK" if B else "OK")
        add("warnings", "Требуют проверки (жёлтый)", len(W), "WARN" if W else "OK")
        add("dup_ei", "Дубли ЕИ", self.dup_eis, "BLOCK" if self.dup_eis else "OK")
        dr = [f for f in self.F if f["kind"] == "дубль строки"]
        add("dup_rows", "Повторяющиеся строки (повторная вставка?)", len(dr),
            "BLOCK" if any(f["level"] == "BLOCK" for f in dr) else ("WARN" if dr else "OK"))
        pc = sum(1 for f in B if f["kind"] == "деталь")
        add("part_conflicts", "Конфликты деталей (артикул ↔ ЕИ)", pc, "BLOCK" if pc else "OK")
        uu = sorted({f["text"].split("«")[1].split("»")[0] for f in W if f["kind"] == "единица" and "неизвестная" in f["text"]})
        add("units_unknown", "Неизвестные единицы", len(uu), "WARN" if uu else "OK", ", ".join(uu[:20]))
        if self.known_places:
            pu = sorted({f["text"].split("«")[1].split("»")[0] for f in W if f["kind"] == "место"})
            add("places_unknown", "Неизвестные места", len(pu), "WARN" if pu else "OK", ", ".join(pu[:20]))
        else:
            places = sorted({it.place for it in eis if it.place})
            add("places_unknown", "Неизвестные места", 0, "INFO", f"список мест не задан — приняты места таблицы ({len(places)})")
        mf = sum(1 for f in B if f["kind"] == "обязательное поле")
        add("missing_fields", "Отсутствующие обязательные поля", mf, "BLOCK" if mf else "OK")
        add("status_diff", "Старый статус ≠ новый рассчитанный", self.status_diff, "WARN" if self.status_diff else "OK")
        add("status_critical", "из них критических противоречий", self.status_critical, "BLOCK" if self.status_critical else "OK")
        add("stock_suspicious", "Подозрительные остатки", self.suspicious, "WARN" if self.suspicious else "OK")
        add("stock_unknown", "Требуется проверить остаток", self.unknown,
            ("BLOCK" if self.s["unknown_balance"] == "stop" else "WARN") if self.unknown else "OK")
        n_ok = sum(1 for r in data if r.result == RESULT_OK)
        n_warn = sum(1 for r in data if r.result == RESULT_WARN)
        n_block = sum(1 for r in data if r.result == RESULT_BLOCK)
        add("rows_auto", "Строк переносится автоматически (зелёные)", n_ok, "OK")
        add("rows_check", "Строк переносится с замечанием (жёлтые)", n_warn, "WARN" if n_warn else "OK")
        add("rows_block", "Строк требуют решения человека (красные)", n_block, "BLOCK" if n_block else "OK")
        add("rows_skipped", "Строк пропущено (пустые, заголовки, по правилу)", sum(1 for r in self.rows if r.fate == "skipped" and r.kind != "EMPTY"))
        add("formulas", "Формул в данных (Transfer заменяет их значениями)", sum(self.st.formulas.values()))
        if self.candidate:
            cp = self.candidate.get("problems", [])
            add("candidate", "Книга-кандидат WMS", "не подходит" if cp else f"{self.candidate.get('core', '')}, схема {self.candidate.get('schema', '')}",
                "BLOCK" if cp else "OK", "; ".join(cp)[:300])
        add("verdict", "ИТОГ", "ПЕРЕНОС НЕВОЗМОЖЕН — исправьте красные строки" if B else "Можно создавать тестовую WMS", "BLOCK" if B else "OK")
        self.metrics = m

    # ------------------------------------------------------------ the plan (JSON): what BUILD executes and VERIFY compares
    def plan_dict(self):
        rows = []
        for kind, pos, d in self.targets:
            if kind == "POS":
                vals = pos.vals
                src_rows = [pos.row.r] + ([pos.merged.r] if pos.merged else [])
                t = pos.t
            else:
                vals = d.vals
                src_rows = [d.r]
                t = d.t
            z = vals.get(C_Z)
            items = [it for it in pos.receipts if it.t == t]
            if items and items[0].bal_src in ("RULE0", "RULEF"):
                z = (z + " " if z else "") + BALANCE_MARK
            out = {COLS[c]: vals.get(c) for c in INPUT_COLS}
            out["Z"] = z
            rows.append(dict(t=t, kind=kind, src_rows=src_rows, pos=pos.t, vals=out))
        items = []
        for it in self.items:
            if not it.n:
                continue
            items.append(dict(kind=it.kind, ei=it.ei, n=it.n, ei_src=it.ei_src, t=it.t, olid=it.olid, fact=it.fact, bal=it.bal,
                              bal_src=it.bal_src, place=it.place, place_now=it.place_now or it.place, name=it.name, art=it.art,
                              unit=it.unit, cat=it.cat, stype=it.stype, state=it.state, nodoc=it.nodoc, mark=it.mark,
                              order=(it.pos.vals.get(C_A) if it.pos else ""), rows=[r.r for r in it.rows],
                              src_t=(it.pos.t if it.pos else None),
                              leg_status=(it.pos.old_raw if it.pos and it.kind == "SRC" else joined(it.rows, C_W)),
                              leg_ctl=joined(it.rows, C_Y), leg_stock=txt(it.xraw) if not isinstance(it.xraw, list) else "",
                              leg_dup=joined(it.rows, C_AB)))
        positions = []
        for pos in self.positions:
            positions.append(dict(t=pos.t, block=pos.block, a=pos.vals.get(C_A), olid=pos.olid, ordq=pos.ordq, rcv=pos.rcv, nodoc=pos.nodoc,
                                  cnt=len(pos.receipts), cancel=pos.cancel, edate=pos.edate, old=pos.old_raw, new=pos.new,
                                  key=(pos.receipts[0].ei if pos.receipts and pos.receipts[0].kind == "SRC" else ""),
                                  deliveries=[d.t for d in pos.deliveries], rows=[pos.row.r] + ([pos.merged.r] if pos.merged else [])))
        ops = []
        for op in self.ops:
            if op["op"] in ("legacy", "legacy_add", "migrate"):
                ops.append(dict(op=op["op"], t=op.get("t"), src=op.get("src"), ei=op["item"].ei))
            else:
                ops.append(dict(op=op["op"], t=op["t"]))
        ei_of, bal_of, new_w = {}, {}, {}
        for it in self.items:
            if it.n:
                for r in it.rows:
                    ei_of[r.r], bal_of[r.r] = it.ei, it.bal
        for pos in self.positions:
            for r in [pos.row] + ([pos.merged] if pos.merged else []):
                new_w[r.r] = pos.new
            for d in pos.deliveries:
                new_w[d.r] = OS_ADD
        mapping = []
        for r in self.rows:
            if r.kind == "EMPTY":
                continue
            mapping.append(dict(row=r.r, kind=r.kind, fate=r.fate, t=r.t, block=r.block, ei=ei_of.get(r.r, ""),
                                result=getattr(r, "result", ""), norm=r.norm, msgs=[m[1] for m in r.msgs],
                                recalc=bool(getattr(r, "recalc", False)), old=[txt(r.orig[c]) for c in (C_W, C_X, C_Y, C_AB)],
                                new_w=new_w.get(r.r, ""), new_x=bal_of.get(r.r)))
        return dict(tool=TOOL_VERSION, created=now().isoformat(timespec="seconds"), staging_hash=self.st.hash,
                    settings=self.s, settings_hash=settings_hash(self.s), rows=rows, items=items, positions=positions, ops=ops,
                    mapping=mapping, next_ei=self.next_ei, next_ol=self.next_ol, blocks=len({p.block for p in self.positions}),
                    today=self.today, blockers=len(self.blockers()), warnings=len(self.warnings()),
                    pairs=[[k[0], k[1], v] for k, v in sorted(self.pairs.items())])


# ================================================================ the files of the check (the frontend loads them into its sheets)

LEVEL_WORD = {"BLOCK": "БЛОКЕР", "WARN": "ПРОВЕРИТЬ", "OK": "ГОТОВО", "INFO": "ИНФО"}


def joined(rows, c):
    """the distinct non-empty old values of column c in the rows of one receipt (as pasted), in row order"""
    out = []
    for r in rows:
        t = txt(r.orig[c])
        if t and t not in out:
            out.append(t)
    return "; ".join(out)


def row_message(r):
    """AE of the row: what stops or needs a look first, then what was normalized"""
    msgs = [m[1] for m in r.msgs if m[0] in ("BLOCK", "WARN")]
    if r.norm:
        msgs.append("нормализовано: " + "; ".join(r.norm))
    return " | ".join(msgs)[:1000]


def write_check_outputs(work, an):
    d = os.path.join(work, "check")
    os.makedirs(d, exist_ok=True)
    write_csv(os.path.join(d, "summary.csv"), ["Показатель", "Значение", "Уровень", "Пояснение", "Ключ"],
              [(label, value if not isinstance(value, int) else float(value), LEVEL_WORD.get(level, level), hint, key)
               for key, label, value, level, hint in an.metrics])
    order = {"BLOCK": 0, "WARN": 1, "INFO": 2}
    sheet_order = {SHEET_ORDERS: 0, SHEET_STOCK: 1, SHEET_ISSUES: 2, SHEET_RETURNS: 3}
    F = sorted(an.F, key=lambda f: (order.get(f["level"], 3), sheet_order.get(f["sheet"], 9), f["row"] or 0))
    write_csv(os.path.join(d, "errors.csv"), ["№", "Лист", "Строка", "Уровень", "Вид", "Описание"],
              [(float(i), f["sheet"], float(f["row"]) if f["row"] else "", LEVEL_WORD[f["level"]], f["kind"], f["text"]) for i, f in enumerate(F, 1)])
    rows = [(float(r.r), r.result, row_message(r)) for r in an.rows if r.kind != "EMPTY"]
    write_csv(os.path.join(d, "rows.csv"), ["Строка", "Результат", "Что не так / что сделано"], rows)
    write_csv(os.path.join(d, "status_pairs.csv"), ["Старый статус", "Новый рассчитанный статус", "Позиций"],
              [(o, n, float(c)) for (o, n), c in sorted(an.pairs.items())])
    plan = an.plan_dict()
    write_text(os.path.join(d, "plan.json"), json.dumps(plan, ensure_ascii=False, indent=1))
    lines = [f"**Данные:** `{an.st.path}` — SHA-256 содержимого `{an.st.hash}`  ", f"**Инструмент:** legacy_transfer {TOOL_VERSION}  ",
             f"**Итог:** {'ПЕРЕНОС НЕВОЗМОЖЕН — блокирующих замечаний ' + str(len(an.blockers())) if an.blockers() else 'можно создавать тестовую WMS'}"
             f"; требуют проверки: {len(an.warnings())}", "", "## Сводка", "", "| Показатель | Значение | Уровень |", "|---|---|---|"]
    for key, label, value, level, hint in an.metrics:
        lines.append(f"| {label} | {value}{(' — ' + hint) if hint else ''} | {LEVEL_WORD.get(level, level)} |")
    lines += ["", "## Старый статус → новый рассчитанный", "", "| Старый | Новый | Позиций |", "|---|---|---|"]
    lines += [f"| {o} | {n} | {c} |" for (o, n), c in sorted(an.pairs.items())]
    lines += ["", "## Замечания", ""]
    lines += [f"- **{LEVEL_WORD[f['level']]}** {f['sheet']} стр. {f['row'] or '—'} ({f['kind']}): {f['text']}" for f in F[:2000]]
    if len(F) > 2000:
        lines.append(f"- … ещё {len(F) - 2000} (полный список — errors.csv)")
    report(os.path.join(d, "CHECK_REPORT.md"), "Проверка переноса старой таблицы", lines)
    write_text(os.path.join(d, "check.json"), json.dumps(dict(blockers=len(an.blockers()), warnings=len(an.warnings()),
                                                            staging_hash=an.st.hash, settings_hash=settings_hash(an.s),
                                                            metrics={k: v for k, _, v, _, _ in an.metrics}), ensure_ascii=False, indent=1))
    return plan


def report(path, title, lines):
    write_text(path, f"# {title}\n\n**Дата:** {now():%d.%m.%Y %H:%M}\n\n" + "\n".join(lines) + "\n")


# ================================================================ state of the steps (WORK/state.json, WORK/state.txt for the frontend)

class State:
    def __init__(self, work):
        self.work = work
        self.path = os.path.join(work, "state.json")
        try:
            self.d = json.load(open(self.path, encoding="utf-8"))
        except (OSError, ValueError):
            self.d = {}

    def get(self, k):
        return self.d.get(k) or {}

    def put(self, k, v, message=""):
        self.d[k] = v
        if message:
            self.d["message"] = message
        self.save()

    def drop(self, *keys):
        for k in keys:
            self.d.pop(k, None)
        self.save()

    def save(self):
        write_text(self.path, json.dumps(self.d, ensure_ascii=False, indent=1))
        c, b, v, p = self.get("check"), self.get("build"), self.get("verify"), self.get("promote")
        flat = dict(check_ok=int(bool(c.get("ok"))), check_time=c.get("time", ""), check_blockers=c.get("blockers", ""),
                    check_warnings=c.get("warnings", ""), build_ok=int(bool(b.get("ok"))), build_time=b.get("time", ""),
                    build_book=b.get("book", ""), verify_ok=int(bool(v.get("ok"))), verify_time=v.get("time", ""),
                    verify_critical=v.get("critical", ""), verify_report=v.get("report", ""), promote_done=int(bool(p.get("ok"))),
                    promote_time=p.get("time", ""), promote_book=p.get("book", ""),
                    promote_ready=int(bool(b.get("ok") and v.get("ok") and not p.get("ok"))),
                    message=str(self.d.get("message", "")).replace("\n", " "))
        write_text(os.path.join(self.work, "state.txt"), "".join(f"{k}={v}\n" for k, v in flat.items()))


def input_hash(st_hash, s):
    return sha256_text(st_hash + "|" + settings_hash(s))


# ================================================================ LibreOffice: reading, the candidate, WMS

class Busy:
    """one engine run per work folder at a time (a second click of a button while the first run works)"""

    def __init__(self, work):
        self.path = os.path.join(work, ".busy")

    def __enter__(self):
        if os.path.exists(self.path):
            try:
                pid = int(open(self.path).read().split()[0])
                os.kill(pid, 0)
                raise StopRun(f"инструмент переноса уже выполняется (процесс {pid}) — дождитесь окончания")
            except (OSError, ValueError, IndexError):
                pass          # a leftover of an interrupted run
        write_text(self.path, f"{os.getpid()}\n")
        return self

    def __exit__(self, *exc):
        try:
            os.remove(self.path)
        except OSError:
            pass


class StopRun(Exception):
    """a stop by the rules: nothing is changed, the message is for the user"""


def office(tag):
    from wmslo import Office
    prof = tempfile.mkdtemp(prefix="wms_legacy_")
    o = Office(f"lt{tag}{os.getpid()}", prof)
    o._legacy_prof = prof
    return o


def office_done(o):
    try:
        o.terminate()
    finally:
        shutil.rmtree(getattr(o, "_legacy_prof", ""), ignore_errors=True)


def core_version_tuple(s):
    m = re.match(r"(\d+)\.(\d+)\.(\d+)", s or "")
    return tuple(int(x) for x in m.groups()) if m else (0, 0, 0)


def inspect_candidate(o, path, allow_test=False):
    """what the transfer needs to know of the candidate book (opened read-only, macros disabled, never stored)"""
    info = dict(path=os.path.abspath(path), problems=[], schema="", core="", mode="", max_qty=MAX_QTY_DEFAULT)
    P = info["problems"]
    if not os.path.isfile(path):
        P.append(f"нет книги-кандидата {path} — положите WMS_LEGACY_TRANSFER.ods в папку выпуска рядом с WMS_PROD_CANDIDATE.ods")
        return info
    jd = os.path.join(os.path.dirname(os.path.abspath(path)), JOURNAL_DIR)
    if os.path.isdir(jd) and any(f.startswith("WMS_journal_") for f in os.listdir(jd)):
        P.append(f"рядом с книгой-кандидатом есть журнал WMS ({jd}) — книгу уже запускали; перенос делается только в чистую книгу: "
                 "распакуйте выпуск заново")
    doc = o.load(os.path.abspath(path), macros=0, hidden=True, ReadOnly=True)
    if doc is None:
        P.append(f"книга-кандидат {path} не открылась")
        return info
    try:
        sh = doc.Sheets
        if not sh.hasByName("_SYS"):
            P.append("в книге-кандидате нет листа _SYS — это не книга WMS")
            return info
        sysd = {}
        for r in sh.getByName("_SYS").getCellRangeByPosition(0, 0, 1, 40).getDataArray():
            if r[0]:
                sysd[str(r[0])] = r[1]
        info.update(schema=str(sysd.get("SCHEMA", "")), core=str(sysd.get("CORE_VERSION", "")), mode=str(sysd.get("MODE", "")),
                    last_seq=sysd.get("LAST_SEQ"), next_ei=sysd.get("NEXT_EI"), sys=sysd, legacy_ops=None)
        # the version of the core is in its Basic code (a book that never started has an empty CORE_VERSION in _SYS)
        try:
            libs = doc.BasicLibraries
            libs.loadLibrary("Standard")
            lib = libs.getByName("Standard")
            m = re.search(r'Public Const WMS_CORE_VERSION = "([^"]+)"', lib.getByName("WmsConfig"))
            if m:
                info["core"] = m.group(1)
            info["legacy_ops"] = lib.hasByName("WmsReceipt") and "Function ReceiptLegacyRow" in lib.getByName("WmsReceipt")
        except Exception:
            pass
        if isinstance(sysd.get("MAX_QTY"), float):
            info["max_qty"] = sysd["MAX_QTY"]
        if info["schema"] != SCHEMA:
            P.append(f"книга-кандидат схемы «{info['schema']}», а этот инструмент переноса работает только со схемой {SCHEMA} (WMS 0.7): "
                     "возьмите книгу-кандидат и WMS_LEGACY_TRANSFER.ods из одного выпуска")
        elif core_version_tuple(info["core"]) < MIN_CORE or info["legacy_ops"] is False:
            P.append(f"ядро книги-кандидата {info['core']} не умеет переносить старую таблицу — нужна книга выпуска "
                     f"{'.'.join(map(str, MIN_CORE))} или новее")
        if info["mode"] != "PROD" and not allow_test:
            P.append(f"книга-кандидат в режиме «{info['mode']}» — нужна рабочая книга-кандидат выпуска (режим PROD)")
        if isinstance(info["last_seq"], float) and info["last_seq"] > 0:
            P.append(f"книга-кандидат уже использовалась: в ней {int(info['last_seq'])} операций — перенос делается только в чистую книгу")
        if not allow_test and isinstance(info["next_ei"], float) and info["next_ei"] != 1:
            P.append(f"в книге-кандидате уже есть ЕИ (NEXT_EI {int(info['next_ei'])}) — нужна чистая книга")
        for name in ("Заказы", "Наличие"):
            if not sh.hasByName(name):
                P.append(f"в книге-кандидате нет листа «{name}»")
                continue
            cur = sh.getByName(name).createCursor()
            cur.gotoEndOfUsedArea(False)
            last = cur.getRangeAddress().EndRow
            if name == "Заказы" and last > 0:
                vals = sh.getByName(name).getCellRangeByPosition(0, 1, 27, min(last, 50)).getDataArray()
                if any(any(v not in ("", None) for v in r) for r in vals):
                    P.append("в книге-кандидате уже есть строки «Заказы» — нужна чистая книга")
            if name == "Наличие" and not allow_test and last > 0:
                vals = sh.getByName(name).getCellRangeByPosition(0, 1, 0, min(last, 50)).getDataArray()
                if any(r[0] for r in vals):
                    P.append("в книге-кандидате уже есть ЕИ в «Наличие» — нужна чистая книга")
    finally:
        try:
            doc.close(True)
        except Exception:
            pass
    return info


def open_wms(o, path):
    import migrate as mg
    return mg.open_wms(o, path)


def close_doc(o, doc, save):
    import migrate as mg
    mg.close_doc(o, doc, save)


def copy_journal(src, dst):
    import migrate as mg
    mg.copy_journal(src, dst)


def B(o, doc, module, func, *args):
    return str(o.basic(doc, module, func, *args) or "")


def sysvals(doc):
    return {str(r[0]): r[1] for r in doc.Sheets.getByName("_SYS").getCellRangeByPosition(0, 0, 1, 40).getDataArray() if r[0]}


def used_last(sheet):
    cur = sheet.createCursor()
    cur.gotoEndOfUsedArea(False)
    return cur.getRangeAddress().EndRow


# ================================================================ BUILD: the test WMS from the clean candidate

def cell_value(v):
    if v is None:
        return ""
    return float(v) if isinstance(v, (int, float)) else str(v)


def origin_of(plan, rows):
    return f"{plan['staging_hash'][:16]}|{SHEET_ORDERS}|{','.join(map(str, rows))}"


def execute(o, doc, plan, prog, fault_at=None, kill_at=None):
    """the plan in the open test book; raises StopRun at the first refusal (the caller throws the copy away)"""
    sh = doc.Sheets.getByName("Заказы")
    if used_last(sh) > 0:
        vals = sh.getCellRangeByPosition(0, 1, 27, min(used_last(sh), 20)).getDataArray()
        if any(any(v not in ("", None) for v in r) for r in vals):
            raise StopRun("в «Заказы» тестовой книги уже есть строки — перенос в неё не выполняется")
    rows = plan["rows"]
    counts = dict(rows=len(rows), positions=len(plan["positions"]), eis=len(plan["items"]), next_ei=plan["next_ei"])
    res = B(o, doc, "WmsMigrate", "LegacyMark", "BEGIN", "\t".join([
        f"STAGING_SHA={plan['staging_hash']}", f"SETTINGS_SHA={plan['settings_hash']}", f"TOOL=legacy_transfer {TOOL_VERSION}",
        f"ROWS={counts['rows']}", f"POSITIONS={counts['positions']}", f"EIS={counts['eis']}",
        f"SOURCE={os.path.basename(plan['settings'].get('source_file') or '') or 'staging'}"]))
    if not res.startswith("OK:"):
        raise StopRun(f"отметка начала переноса: {res}")
    prog(0, 1, "строки заказов", True)
    if rows:
        au = tuple(tuple(cell_value(r["vals"].get(COLS[c])) for c in range(C_A, C_U + 1)) for r in rows)
        za = tuple((cell_value(r["vals"].get("Z")), cell_value(r["vals"].get("AA"))) for r in rows)
        try:
            sh.getCellRangeByPosition(C_A, 1, C_U, len(rows)).setDataArray(au)
            sh.getCellRangeByPosition(C_Z, 1, C_AA, len(rows)).setDataArray(za)
        except Exception as e:
            raise StopRun(f"строки не записаны в «Заказы»: {e}")
    items = {it["ei"]: it for it in plan["items"]}
    ops = plan["ops"]
    for k, op in enumerate(ops, 1):
        if fault_at == k:
            B(o, doc, "WmsCore", "TestSetFault", 2, 1)
        if kill_at == k:
            o.kill()
        it = items.get(op.get("ei"))
        if op["op"] == "legacy":
            res = B(o, doc, "WmsReceipt", "ReceiptLegacyRow", op["t"], it["ei"], qty_text(it["bal"]),
                    it["place_now"] if it["place_now"] != it["place"] else "", origin_of(plan, it["rows"]), it["leg_status"],
                    it["leg_ctl"], it["leg_stock"], it["leg_dup"], it["bal_src"], it["ei_src"])
        elif op["op"] == "legacy_add":
            res = B(o, doc, "WmsReceipt", "ReceiptLegacyAdd", op["t"], op["src"], it["ei"], qty_text(it["bal"]),
                    it["place_now"] if it["place_now"] != it["place"] else "", origin_of(plan, it["rows"]), it["leg_status"],
                    it["leg_ctl"], it["leg_stock"], it["leg_dup"], it["bal_src"], it["ei_src"])
        elif op["op"] == "cancel":
            res = B(o, doc, "WmsReceipt", "OrderCancelRow", op["t"])
        elif op["op"] == "cancel_rest":
            res = B(o, doc, "WmsReceipt", "OrderCancelRestRow", op["t"])
        elif op["op"] == "migrate":
            origin = (f"{plan['staging_hash'][:16]}|{SHEET_ORDERS}|{','.join(map(str, it['rows']))}" if it["rows"]
                      else f"{plan['staging_hash'][:16]}|{SHEET_STOCK}|{it['mark']}")
            res = B(o, doc, "WmsMigrate", "MigrateEI", it["ei"], it["name"], it["art"], it["unit"], qty_text(it["bal"]),
                    it["place_now"] or it["place"], it["cat"], it["stype"], it["mark"], origin)
        else:
            raise StopRun(f"неизвестная операция плана {op['op']}")
        if not res.startswith("OK:"):
            where = f"строка {op['t'] + 1} новой «Заказы»" if op.get("t") is not None else "ЕИ"
            raise StopRun(f"операция {k} из {len(ops)} ({op['op']}, {where}, {op.get('ei') or ''}): {res}")
        prog(k, len(ops), f"операции: {k} из {len(ops)}")
    prog(len(ops), len(ops), "статусы заказов", True)
    rs = B(o, doc, "WmsOrders", "RefreshStatuses")
    res = B(o, doc, "WmsMigrate", "LegacyMark", "END", "\t".join([
        f"STAGING_SHA={plan['staging_hash']}", f"ROWS={counts['rows']}", f"POSITIONS={counts['positions']}", f"EIS={counts['eis']}",
        f"OPS={len(ops)}", f"NEXT_EI={plan['next_ei']}"]))
    if not res.startswith("OK:"):
        raise StopRun(f"отметка конца переноса: {res}")
    return rs


def oracle(doc, jdir):
    import adjust_oracle
    initial = {}
    if sysvals(doc).get("MODE") == "TEST":
        # a test book of the developer (the tests of the transfer): its synthetic registry is the initial stock
        from build_ods import synthetic_registry
        initial = {r[0]: r[4] for r in synthetic_registry()}
    P, info, _ = adjust_oracle.check(doc, jdir, initial)
    return list(P), info


def cmd_build(a):
    work = os.path.abspath(a.work)
    os.makedirs(work, exist_ok=True)
    state = State(work)
    prog = Progress(work, "build")
    t0 = time.time()
    part = os.path.join(work, "test.part")
    done = os.path.join(work, "test")
    lines = []
    with Busy(work):
        s, sp = load_settings(a.settings, overrides(a))
        prog(0, 1, "чтение данных", True)
        o = office("r")
        try:
            st = read_staging(o, a.staging)
            cand = inspect_candidate(o, a.candidate, a.allow_test_candidate)
        finally:
            office_done(o)
        an = Analysis(st, s, max_qty=cand.get("max_qty", MAX_QTY_DEFAULT), candidate=cand)
        an.settings_problems = sp
        an.run()
        plan = write_check_outputs(work, an)
        ih = input_hash(st.hash, s)
        src = s.get("source_file") or ""
        state.put("check", dict(ok=not an.blockers(), time=now().isoformat(timespec="seconds"), blockers=len(an.blockers()),
                                warnings=len(an.warnings()), input=ih, staging=st.hash, source_file=src,
                                source_hash=sha256_file(src) if src and os.path.isfile(src) else ""))
        if an.blockers():
            state.put("build", dict(ok=False, time=now().isoformat(timespec="seconds"), error="блокирующие замечания"),
                      f"Тестовая WMS не создана: блокирующих замечаний {len(an.blockers())} — см. «3_Ошибки»")
            print(f"СТОП: блокирующих замечаний {len(an.blockers())} — тестовая WMS не создаётся (см. check/errors.csv)", file=sys.stderr)
            return 1
        if os.path.exists(part):
            shutil.rmtree(part, ignore_errors=True)
            lines.append("- остаток прерванного создания (test.part) удалён")
        os.makedirs(part)
        book = os.path.join(part, TEST_BOOK)
        cand_hash = sha256_file(a.candidate)
        shutil.copy2(a.candidate, book)
        failed, sc, sc2, P, info, rec, rs, idx = None, "", "", [], {}, None, "", ""
        prog(0, 1, "запуск WMS в копии кандидата", True)
        o = office("b")
        try:
            doc, stl = open_wms(o, book)
            if doc is None or "STATE=CLEAN" not in stl:
                raise StopRun(f"копия кандидата не стала рабочей WMS: {stl[:300]}")
            rs = execute(o, doc, plan, prog, getattr(a, "fault_at", None), getattr(a, "kill_at", None))
            prog(1, 1, "индекс деталей и самопроверка", True)
            idx = B(o, doc, "WmsSpecial", "ArticleIndexRebuild")
            if idx.startswith("ОШИБКА") or "конфликтов «один артикул — несколько ЕИ» 0" not in idx:
                raise StopRun(f"индекс деталей: {idx[:300]}")
            sc = B(o, doc, "WmsCore", "ActionSelfCheck")
            if not sc.startswith(SELF_CHECK_OK):
                raise StopRun("самопроверка WMS: " + "; ".join(x for x in sc.splitlines() if x.startswith("FAIL"))[:600])
            prog(1, 1, "независимая проверка журнала (оракул)", True)
            P, info = oracle(doc, os.path.join(part, JOURNAL_DIR))
            if P:
                raise StopRun("оракул (журнал воспроизведён независимо): " + "; ".join(P[:5]))
            prog(1, 1, "сверка с данными", True)
            rec = reconcile(doc, plan, an)
            if rec["critical"]:
                raise StopRun("сверка с данными: " + "; ".join(rec["critical"][:5]))
            prog(1, 1, "сохранение", True)
            close_doc(o, doc, save=True)
            doc = None
            prog(1, 1, "повторное открытие", True)
            doc, stl2 = open_wms(o, book)
            if doc is None or "STATE=CLEAN" not in stl2:
                raise StopRun(f"после сохранения книга не открылась в рабочем состоянии: {stl2[:300]}")
            sc2 = B(o, doc, "WmsCore", "ActionSelfCheck")
            if not sc2.startswith(SELF_CHECK_OK):
                raise StopRun("самопроверка после повторного открытия: " + sc2.splitlines()[0])
            close_doc(o, doc, save=False)
            doc = None
        except StopRun as e:
            failed = str(e)
        except Exception as e:
            failed = f"ошибка при создании тестовой WMS: {e}"
            traceback.print_exc()
        finally:
            office_done(o)
        if not failed and sha256_file(a.candidate) != cand_hash:
            failed = "книга-кандидат изменилась во время переноса — так быть не должно; тестовая WMS отброшена"
        if failed:
            shutil.rmtree(part, ignore_errors=True)
            report(os.path.join(work, "BUILD_REPORT.md"), "Тестовая WMS НЕ создана", [
                f"**Причина остановки:** {failed}", "",
                "Копия удалена целиком — перенос выполняется только весь или никак. Книга-кандидат и данные не изменены "
                f"(SHA-256 кандидата `{cand_hash}`)."] + lines)
            state.drop("verify")
            state.put("build", dict(ok=False, time=now().isoformat(timespec="seconds"), error=failed), f"Тестовая WMS не создана: {failed}")
            print(f"СТОП: {failed}", file=sys.stderr)
            return 1
        write_text(os.path.join(part, "plan.json"), json.dumps(plan, ensure_ascii=False, indent=1))
        old = os.path.join(work, "test.old")
        shutil.rmtree(old, ignore_errors=True)
        if os.path.exists(done):
            os.replace(done, old)
        os.replace(part, done)
        shutil.rmtree(old, ignore_errors=True)
        final = os.path.join(done, TEST_BOOK)
        # the book is registered at test.part/: its first start from test/ asks to register it there — done now, without
        # an operation (ActionRegisterHere keeps the journal), so that VERIFY and the user open it as a working WMS
        o = office("g")
        try:
            doc, stl = open_wms(o, final)
            ok = doc is not None and "STATE=CLEAN" in stl
            if doc is not None:
                close_doc(o, doc, save=ok)
        finally:
            office_done(o)
        if not ok:
            shutil.rmtree(done, ignore_errors=True)
            state.put("build", dict(ok=False, time=now().isoformat(timespec="seconds"), error=stl[:300]),
                      f"Тестовая WMS не создана: после переноса в папку test/ книга не открылась ({stl[:200]})")
            print(f"СТОП: {stl[:300]}", file=sys.stderr)
            return 1
        th = sha256_file(final)
        secs = time.time() - t0
        m = {k: v for k, _, v, _, _ in an.metrics}
        lines = [f"**Данные:** SHA-256 `{st.hash}`  ", f"**Книга-кандидат:** `{os.path.abspath(a.candidate)}` (SHA-256 `{cand_hash}`) — не изменялась  ",
                 f"**Тестовая WMS:** `{final}` (SHA-256 `{th}`)  ", f"**Время:** {secs:.0f} с", "",
                 f"- строк «Заказы»: {len(plan['rows'])}; заказов {m.get('orders')}, позиций {m.get('positions')}; ЕИ {len(plan['items'])}; "
                 f"операций {len(plan['ops'])}; NEXT_EI {plan['next_ei']}",
                 f"- статусы: {rs}", f"- индекс деталей: {idx}", f"- самопроверка: {sc.splitlines()[0] if sc else '—'}; после повторного "
                 f"открытия: {sc2.splitlines()[0] if sc2 else '—'}", f"- оракул: расхождений 0 (операций {info.get('ops', '—')})",
                 f"- сверка с данными: критических расхождений 0"] + lines
        report(os.path.join(work, "BUILD_REPORT.md"), "Тестовая WMS создана", lines)
        shutil.copy2(os.path.join(work, "BUILD_REPORT.md"), os.path.join(done, "BUILD_REPORT.md"))
        state.drop("verify")
        state.put("build", dict(ok=True, time=now().isoformat(timespec="seconds"), book=final, test_hash=th, input=ih, staging=st.hash,
                                candidate=os.path.abspath(a.candidate), candidate_hash=cand_hash, seconds=round(secs, 1)),
                  f"Тестовая WMS создана: {final} ({secs:.0f} с). Следующий шаг — «СВЕРИТЬ»")
        print(f"OK: тестовая WMS {final} — строк {len(plan['rows'])}, ЕИ {len(plan['items'])}, операций {len(plan['ops'])}, {secs:.0f} с")
        return 0


# ================================================================ reconciliation: the WMS book against the plan (BUILD, VERIFY)

def same(a, b):
    if a in (None, "") and b in (None, ""):
        return True
    if isinstance(a, float) and isinstance(b, (int, float)):
        return abs(a - float(b)) < 1e-6
    if isinstance(b, float) and isinstance(a, (int, float)):
        return abs(float(a) - b) < 1e-6
    return str(a) == str(b)


def shown(v):
    if v in (None, ""):
        return "пусто"
    return txt(v) if isinstance(v, float) else f"«{v}»"


def reconcile(doc, plan, an=None, today=None):
    """critical: every difference between the book and the plan (a row, a value, an EI, a position, a counter)"""
    today = today or today_serial()
    C = []
    sheets = doc.Sheets
    orders = sheets.getByName("Заказы")
    n = len(plan["rows"])
    last = used_last(orders)
    data = orders.getCellRangeByPosition(0, 0, 28, max(last, n, 1)).getDataArray()
    extra = [r for r in range(n + 1, len(data)) if any(v not in ("", None) for v in data[r][:28])]
    if extra:
        C.append(f"«Заказы»: лишние строки {', '.join(str(r + 1) for r in extra[:10])}")
    items_by_t = {}
    for it in plan["items"]:
        if it["t"] is not None and it["kind"] != "MIGRATE":
            items_by_t[it["t"]] = it
    pos_by_t = {p["t"]: p for p in plan["positions"]}
    bad_rows = set()
    for r in plan["rows"]:
        t = r["t"]
        got = data[t] if t < len(data) else [""] * 29
        want = r["vals"]
        for c in INPUT_COLS:
            if not same(got[c], want.get(COLS[c])):
                C.append(f"«Заказы» строка {t + 1} ({COLS[c]}): в книге {shown(got[c])}, по данным {shown(want.get(COLS[c]))}")
                bad_rows.add(t)
        it = items_by_t.get(t)
        pos = pos_by_t.get(r["pos"])
        want_v = it["ei"] if it else ""
        if got[C_V] != want_v:
            C.append(f"«Заказы» строка {t + 1}: ЕИ {shown(got[C_V])}, ожидался {shown(want_v)}")
            bad_rows.add(t)
        if r["kind"] == "DEL":
            want_w, want_ac = OS_ADD, f"OL{pos['olid']}"
        else:
            want_w = position_status(pos["ordq"], pos["rcv"], pos["nodoc"], pos["cancel"], pos["edate"], today)
            want_ac = ""
        if got[C_W] != want_w:
            C.append(f"«Заказы» строка {t + 1}: статус «{got[C_W]}», новый движок должен дать «{want_w}»")
            bad_rows.add(t)
        want_x = it["bal"] if it else ""
        if not same(got[C_X], want_x):
            C.append(f"«Заказы» строка {t + 1}: наличие {shown(got[C_X])}, ожидалось {shown(want_x)}")
            bad_rows.add(t)
        if got[28] != want_ac:
            C.append(f"«Заказы» строка {t + 1}: служебная отметка AC {shown(got[28])}, ожидалась {shown(want_ac)}")
            bad_rows.add(t)
    # «Наличие», _RCV
    stock = sheets.getByName("Наличие")
    rcv = sheets.getByName("_RCV")
    top = max([it["n"] for it in plan["items"]] + [1])
    sd = stock.getCellRangeByPosition(0, 0, 9, max(top, used_last(stock), 1)).getDataArray()
    rd = rcv.getCellRangeByPosition(0, 0, 8, max(top, 1)).getDataArray()
    for it in plan["items"]:
        k = it["n"]
        row = sd[k]
        if it["kind"] == "MIGRATE":
            src = "Перенос" + (f" ({it['mark']})" if it["mark"] else "")
        else:
            src = f"Перенос: заказ {it['order']}"
        want = (it["ei"], it["name"], it["art"], it["unit"], it["bal"], it["place_now"] or it["place"], it["cat"], it["state"], src, it["stype"])
        if not all(same(g, w) for g, w in zip(row, want)):
            C.append(f"{it['ei']}: «Наличие» {tuple(row)} ≠ по данным {want}")
        if it["kind"] != "MIGRATE":
            rr = rd[k]
            want_r = (it["ei"], float(it["olid"]), float(it["t"]), None, "LIVE", "SRC" if it["kind"] == "SRC" else "ADD", it["fact"], float(it["nodoc"]))
            if not all(w is None or same(g, w) for g, w in zip(rr, want_r)):
                C.append(f"{it['ei']}: _RCV {tuple(rr[:8])} ≠ по данным {want_r}")
    for k in range(1, len(sd)):
        q = sd[k][4]
        if isinstance(q, float) and q < -EPS:
            C.append(f"«Наличие» строка {k + 1}: отрицательный остаток {q}")
    # _ORD
    ords = sheets.getByName("_ORD")
    od = ords.getCellRangeByPosition(0, 0, 11, max(plan["next_ol"], 1)).getDataArray()
    for p in plan["positions"]:
        if not p["olid"]:
            continue
        row = od[p["olid"]]
        want = (float(p["olid"]), float(p["t"]), p["key"], None, p["ordq"], p["rcv"], float(p["cnt"]), float(p["nodoc"]), p["cancel"])
        if not all(w is None or same(g, w) for g, w in zip(row, want)):
            C.append(f"позиция OLID {p['olid']} (строка {p['t'] + 1}): _ORD {tuple(row[:9])} ≠ по данным {want}")
    sysd = sysvals(doc)
    if sysd.get("NEXT_EI") != float(plan["next_ei"]):
        C.append(f"NEXT_EI {sysd.get('NEXT_EI')}, по данным {plan['next_ei']}")
    if od[0][11] != float(plan["next_ol"]):
        C.append(f"NEXT_OL {od[0][11]}, по данным {plan['next_ol']}")
    lost = [m["row"] for m in plan["mapping"] if m["fate"] not in ("row", "merged", "special", "skipped")]
    if lost:
        C.append(f"строки данных без места в новой WMS: {', '.join(map(str, lost[:20]))}")
    return dict(critical=C, bad_rows=bad_rows, sys=sysd)


def journal_counts(jdir):
    import journal_oracle
    entries, _ = journal_oracle.read_journal(jdir)
    out = {}
    for e in entries:
        out[e["type"]] = out.get(e["type"], 0) + 1
    return out, entries


# ================================================================ VERIFY: the final reconciliation and the report

CATS = ("Полностью совпало", "Нормализовано", "Пересчитано новым движком", "Требует внимания", "Ошибка")


def categorize(plan, rec):
    """the category of every staging row: the worst of what happened to it"""
    out = {}
    bad_t = rec["bad_rows"] if rec else set()
    for m in plan["mapping"]:
        if m["fate"] == "skipped" and m["kind"] in ("HEADER",):
            continue
        if m["t"] is not None and m["t"] in bad_t:
            cat = 4
        elif m["result"] == RESULT_WARN or (m["fate"] == "skipped" and m["kind"] not in ("HEADER",)):
            cat = 3
        elif m["recalc"]:
            cat = 2
        elif m["norm"]:
            cat = 1
        else:
            cat = 0
        out[m["row"]] = cat
    for it in plan["items"]:
        if it["bal_src"] in ("1B", "1C/1D"):
            for r in it["rows"]:
                if out.get(r, 0) < 2:
                    out[r] = 2
    return out


def cmd_verify(a):
    work = os.path.abspath(a.work)
    state = State(work)
    prog = Progress(work, "verify")
    b = state.get("build")
    with Busy(work):
        if not b.get("ok"):
            raise StopRun("нет созданной тестовой WMS — сначала «СОЗДАТЬ ТЕСТОВУЮ WMS»")
        book = b["book"]
        if not os.path.isfile(book) or sha256_file(book) != b["test_hash"]:
            raise StopRun("тестовая WMS изменилась или удалена после создания — создайте её заново («СОЗДАТЬ ТЕСТОВУЮ WMS»)")
        s, sp = load_settings(a.settings, overrides(a))
        prog(0, 1, "чтение данных", True)
        o = office("v")
        try:
            st = read_staging(o, a.staging)
        finally:
            office_done(o)
        if input_hash(st.hash, s) != b["input"]:
            raise StopRun("данные или правила изменились после создания тестовой WMS — нажмите «СОЗДАТЬ ТЕСТОВУЮ WMS» заново")
        plan = json.load(open(os.path.join(os.path.dirname(book), "plan.json"), encoding="utf-8"))
        prog(0, 1, "открытие тестовой WMS", True)
        o = office("v2")
        sc, P, info, rec, jc = "", [], {}, None, {}
        try:
            doc, stl = open_wms(o, book)
            if doc is None or "STATE=CLEAN" not in stl:
                raise StopRun(f"тестовая WMS не открылась в рабочем состоянии: {stl[:300]}")
            prog(0, 1, "самопроверка", True)
            sc = B(o, doc, "WmsCore", "ActionSelfCheck")
            prog(0, 1, "оракул", True)
            P, info = oracle(doc, os.path.join(os.path.dirname(book), JOURNAL_DIR))
            prog(0, 1, "сверка", True)
            rec = reconcile(doc, plan)
            jc, _ = journal_counts(os.path.join(os.path.dirname(book), JOURNAL_DIR))
            close_doc(o, doc, save=False)
        finally:
            office_done(o)
        crit = list(rec["critical"])
        want_j = dict(LEGACY_RECEIPT=sum(1 for op in plan["ops"] if op["op"] == "legacy"),
                      LEGACY_RECEIPT_ADD=sum(1 for op in plan["ops"] if op["op"] == "legacy_add"),
                      ORDER_CANCEL=sum(1 for op in plan["ops"] if op["op"] == "cancel"),
                      ORDER_CANCEL_REST=sum(1 for op in plan["ops"] if op["op"] == "cancel_rest"),
                      MIGRATE=sum(1 for op in plan["ops"] if op["op"] == "migrate"), LEGACY_BEGIN=1, LEGACY_END=1)
        for k, v in want_j.items():
            if jc.get(k, 0) != v:
                crit.append(f"журнал: операций {k} {jc.get(k, 0)}, по плану {v}")
        if not sc.startswith(SELF_CHECK_OK):
            crit.append("самопроверка: " + "; ".join(x for x in sc.splitlines() if x.startswith("FAIL"))[:400])
        crit += [f"оракул: {x}" for x in P[:20]]
        if sha256_file(book) != b["test_hash"]:
            crit.append("файл тестовой WMS изменился во время сверки")
        cats = categorize(plan, rec)
        rep = write_verify_report(work, plan, cats, crit, sc, info, st, b)
        ok = not crit
        state.put("verify", dict(ok=ok, time=now().isoformat(timespec="seconds"), critical=len(crit), report=rep, test_hash=b["test_hash"],
                                 input=b["input"]),
                  ("Сверка: критических расхождений 0 — можно «СДЕЛАТЬ РАБОЧЕЙ»" if ok else
                   f"Сверка: критических расхождений {len(crit)} — «СДЕЛАТЬ РАБОЧЕЙ» недоступно"))
        print(f"{'OK' if ok else 'СТОП'}: критических расхождений {len(crit)}; отчёт {rep}")
        return 0 if ok else 1


def write_verify_report(work, plan, cats, crit, sc, info, st, b):
    counts = [0] * 5
    for c in cats.values():
        counts[c] += 1
    by_row = {m["row"]: m for m in plan["mapping"]}
    today = today_serial()
    moved = sum(1 for m in plan["mapping"] if m["fate"] == "merged")
    lost = sum(1 for m in plan["mapping"] if m["fate"] not in ("row", "merged", "special", "skipped"))
    head = [("Итог", "МОЖНО СДЕЛАТЬ РАБОЧЕЙ" if not crit else "НЕЛЬЗЯ СДЕЛАТЬ РАБОЧЕЙ — есть критические расхождения"),
            ("Тестовая WMS", b["book"]), ("SHA-256 тестовой WMS", b["test_hash"]), ("SHA-256 данных (staging)", st.hash),
            ("Строк данных", str(len([m for m in plan["mapping"] if m["kind"] not in ("HEADER",)]))),
            ("Строк новой «Заказы»", str(len(plan["rows"]))), ("Заказов", str(plan["blocks"])), ("Позиций", str(len(plan["positions"]))),
            ("ЕИ", str(len(plan["items"]))), ("NEXT_EI", str(plan["next_ei"])),
            ("Поступлений, перенесённых в строку позиции", str(moved)), ("Потерянных строк", str(lost)),
            ("Самопроверка WMS", sc.splitlines()[0] if sc else "—"), ("Оракул (журнал воспроизведён независимо)",
                                                                      f"операций {info.get('ops', '—')}, расхождений {len([x for x in crit if x.startswith('оракул')])}"),
            ("Критических расхождений", str(len(crit)))]
    md = ["| | |", "|---|---|"] + [f"| {k} | {v} |" for k, v in head]
    md += ["", "## Итог по строкам", "", "| Категория | Строк |", "|---|---|"] + [f"| {CATS[i]} | {counts[i]} |" for i in range(5)]
    md += ["", "## Старый статус → новый рассчитанный", "", "| Старый статус | Новый статус | Позиций |", "|---|---|---|"]
    md += [f"| {o} | {n} | {c} |" for o, n, c in plan["pairs"]]
    sections = []
    for ci in (4, 3, 2, 1):
        rows = sorted(r for r, c in cats.items() if c == ci)
        items = []
        for r in rows[:1500]:
            m = by_row[r]
            what = "; ".join(m["msgs"] + m["norm"]) or ("статус пересчитан" if m["recalc"] else "остаток определён не по X")
            items.append((r, m["t"] + 1 if m["t"] is not None else "—", m["ei"], what))
        sections.append((CATS[ci], rows, items))
    for title, rows, items in sections:
        md += ["", f"## {title} ({len(rows)})", ""]
        if title == "Ошибка":
            md += [f"- {x}" for x in crit[:500]] or ["Нет."]
        md += [f"- строка {r} → «Заказы» {t}{(' ' + e) if e else ''}: {w}" for r, t, e, w in items]
        if len(rows) > len(items):
            md.append(f"- … ещё {len(rows) - len(items)} (LEGACY_ROWS.csv)")
    report(os.path.join(work, REPORT_NAME + ".md"), "Отчёт переноса старой таблицы в WMS", md)
    # the same as HTML
    esc = html.escape
    colors = ["#e8f5e9", "#e3f2fd", "#ede7f6", "#fff8e1", "#ffebee"]
    h = ["<!doctype html><html lang=ru><head><meta charset=utf-8><title>Отчёт переноса старой таблицы</title><style>",
         "body{font-family:DejaVu Sans,Arial,sans-serif;margin:24px;color:#1b1b1b;background:#fff}table{border-collapse:collapse;margin:8px 0}"
         "td,th{border:1px solid #ccc;padding:4px 8px;text-align:left;vertical-align:top}th{background:#f2f2f2}h1{font-size:22px}"
         "h2{font-size:18px;margin-top:28px}.ok{color:#1b5e20;font-weight:bold}.bad{color:#b71c1c;font-weight:bold}li{margin:2px 0}",
         "</style></head><body>", f"<h1>Отчёт переноса старой таблицы в WMS</h1><p>Дата: {now():%d.%m.%Y %H:%M}</p><table>"]
    for k, v in head:
        cls = ""
        if k == "Итог":
            cls = " class=ok" if not crit else " class=bad"
        h.append(f"<tr><th>{esc(k)}</th><td{cls}>{esc(v)}</td></tr>")
    h.append("</table><h2>Итог по строкам</h2><table><tr><th>Категория</th><th>Строк</th></tr>")
    for i in range(5):
        h.append(f"<tr style='background:{colors[i]}'><td>{esc(CATS[i])}</td><td>{counts[i]}</td></tr>")
    h.append("</table><h2>Старый статус → новый рассчитанный</h2><table><tr><th>Старый статус</th><th>Новый статус</th><th>Позиций</th></tr>")
    for o_, n_, c_ in plan["pairs"]:
        h.append(f"<tr><td>{esc(o_)}</td><td>{esc(n_)}</td><td>{c_}</td></tr>")
    h.append("</table>")
    for title, rows, items in sections:
        h.append(f"<h2>{esc(title)} ({len(rows)})</h2><ul>")
        if title == "Ошибка":
            h += [f"<li class=bad>{esc(x)}</li>" for x in crit[:500]]
        h += [f"<li>строка {r} → «Заказы» {t}{(' ' + esc(e)) if e else ''}: {esc(w)}</li>" for r, t, e, w in items]
        if len(rows) > len(items):
            h.append(f"<li>… ещё {len(rows) - len(items)} (LEGACY_ROWS.csv)</li>")
        h.append("</ul>")
    h.append("</body></html>")
    write_text(os.path.join(work, REPORT_NAME + ".html"), "\n".join(h))
    write_csv(os.path.join(work, "verify_summary.csv"), ["Показатель", "Значение"],
              head + [(CATS[i], float(counts[i])) for i in range(5)] + [("Отчёт", os.path.join(work, REPORT_NAME + ".html"))])
    # the record of every row of the old table: where it went, its EI, its old W X Y AB beside what the new WMS shows
    write_csv(os.path.join(work, "LEGACY_ROWS.csv"), ["Строка данных", "Вид", "Судьба", "Строка «Заказы»", "Заказ", "ЕИ",
                                                      "Старый статус (W)", "Новый статус", "Старое наличие (X)", "Наличие в WMS",
                                                      "Старый контроль (Y)", "Старый «Возможный дубль» (AB)", "Категория", "Что сделано"],
              [(float(m["row"]), m["kind"], {"row": "перенесена", "merged": "в строку позиции", "special": "ЕИ (специальный приход)",
                                            "skipped": "пропущена"}.get(m["fate"], m["fate"]), float(m["t"] + 1) if m["t"] is not None else "",
                float(m["block"]) if m["block"] else "", m["ei"], m["old"][0], m["new_w"], m["old"][1],
                m["new_x"] if m["new_x"] is not None else "", m["old"][2], m["old"][3], CATS[cats[m["row"]]] if m["row"] in cats else "",
                "; ".join(m["msgs"] + m["norm"])) for m in plan["mapping"]])
    return os.path.join(work, REPORT_NAME + ".html")


# ================================================================ PROMOTE: the working WMS

def cmd_promote(a):
    work = os.path.abspath(a.work)
    state = State(work)
    b, v, p = state.get("build"), state.get("verify"), state.get("promote")
    dst = os.path.abspath(a.to)
    with Busy(work):
        if p.get("ok"):
            raise StopRun(f"перенос уже сделан рабочим {p.get('time')}: {p.get('book')} — повторный перенос не выполняется")
        if not b.get("ok") or not v.get("ok") or v.get("test_hash") != b.get("test_hash"):
            raise StopRun("нет успешной сверки созданной тестовой WMS — «СДЕЛАТЬ РАБОЧЕЙ» недоступно (шаги 3 и 4)")
        book = b["book"]
        if not os.path.isfile(book) or sha256_file(book) != b["test_hash"]:
            raise StopRun("тестовая WMS изменилась после сверки — создайте и сверьте её заново")
        s, _ = load_settings(a.settings, overrides(a))
        if a.staging:
            o = office("p0")
            try:
                st = read_staging(o, a.staging)
            finally:
                office_done(o)
            if input_hash(st.hash, s) != b["input"]:
                raise StopRun("данные или правила изменились после создания тестовой WMS — создайте и сверьте её заново")
        src0 = state.get("check").get("source_file") or ""
        h0 = state.get("check").get("source_hash") or ""
        if src0 and h0 and os.path.isfile(src0) and sha256_file(src0) != h0:
            raise StopRun(f"файл старой таблицы {src0} изменился после проверки — в нём могли появиться новые движения: вставьте лист "
                          "«Заказы» заново и повторите шаги 2–4")
        target = os.path.join(dst, a.name)
        if os.path.exists(target):
            raise StopRun(f"{target} уже существует — рабочая WMS не перезаписывается")
        if os.path.isdir(dst) and os.listdir(dst):
            raise StopRun(f"папка {dst} не пуста — рабочая WMS создаётся в новой (пустой) папке")
        created = not os.path.isdir(dst)
        os.makedirs(dst, exist_ok=True)
        t0 = now()
        try:
            shutil.copy2(book, target)
            copy_journal(os.path.join(os.path.dirname(book), JOURNAL_DIR), os.path.join(dst, JOURNAL_DIR))
            o = office("p")
            try:
                doc, stl = open_wms(o, target)
                if doc is None or "STATE=CLEAN" not in stl:
                    raise StopRun(f"рабочая WMS не открылась в рабочем состоянии: {stl[:300]}")
                sc = B(o, doc, "WmsCore", "ActionSelfCheck")
                if not sc.startswith(SELF_CHECK_OK):
                    raise StopRun("самопроверка рабочей WMS: " + sc.splitlines()[0])
                close_doc(o, doc, save=True)
                doc, stl = open_wms(o, target)
                if doc is None or "STATE=CLEAN" not in stl:
                    raise StopRun(f"рабочая WMS после сохранения: {stl[:300]}")
                close_doc(o, doc, save=False)
            finally:
                office_done(o)
            ph = sha256_file(target)
            rel = os.path.dirname(os.path.abspath(b["candidate"])) if b.get("candidate") else ""
            tb = os.path.join(rel, "WMS_TOOLBOX") if rel else ""
            if tb and os.path.isdir(tb):
                shutil.copytree(tb, os.path.join(dst, "WMS_TOOLBOX"))
            bk = os.path.join(dst, "backup", "release")
            os.makedirs(bk)
            if os.path.isfile(b.get("candidate", "")):
                shutil.copy2(b["candidate"], bk)
            if rel and os.path.isfile(os.path.join(rel, "RELEASE_MANIFEST.csv")):
                shutil.copy2(os.path.join(rel, "RELEASE_MANIFEST.csv"), bk)
            rd = os.path.join(dst, "LEGACY_TRANSFER")
            os.makedirs(rd)
            for f in (REPORT_NAME + ".md", REPORT_NAME + ".html", "LEGACY_ROWS.csv", "BUILD_REPORT.md", os.path.join("check", "CHECK_REPORT.md"),
                      os.path.join("check", "errors.csv"), os.path.join("check", "summary.csv")):
                if os.path.isfile(os.path.join(work, f)):
                    shutil.copy2(os.path.join(work, f), rd)
            src_file = s.get("source_file") or state.get("check").get("source_file") or ""
            src_hash = sha256_file(src_file) if src_file and os.path.isfile(src_file) else ""
            cut = [
                "ПЕРЕХОД НА WMS",
                "",
                "С момента запуска WMS_PROD все новые движения записываются только в новую WMS.",
                "Старая таблица остаётся read-only архивом.",
                "",
                f"Дата и время перехода: {t0:%d.%m.%Y %H:%M:%S}",
                f"Рабочая WMS: {target}",
                f"SHA-256 рабочей WMS на момент перехода: {ph}",
                f"SHA-256 данных старой таблицы (staging): {b.get('staging', '')}",
                f"Файл старой таблицы: {src_file or 'не указан'}" + (f" (SHA-256 {src_hash})" if src_hash else ""),
                f"Отчёт переноса: LEGACY_TRANSFER/{REPORT_NAME}.html (и .md, LEGACY_ROWS.csv)",
                f"Резервная копия выпуска: backup/release/ (книга-кандидат и RELEASE_MANIFEST.csv)",
                f"Тестовая WMS переноса: {book} (SHA-256 {b['test_hash']}) — перенесена в архив папки переноса",
                "",
                "Что дальше: открывайте WMS_PROD.ods в этой папке. Старую таблицу больше не меняйте.",
            ]
            write_text(os.path.join(dst, "CUTOVER_README.txt"), "\n".join(cut) + "\n")
        except Exception as e:
            if created:
                shutil.rmtree(dst, ignore_errors=True)
            else:
                for f in os.listdir(dst):
                    q = os.path.join(dst, f)
                    shutil.rmtree(q, ignore_errors=True) if os.path.isdir(q) else os.remove(q)
            if isinstance(e, StopRun):
                raise
            raise StopRun(f"рабочая WMS не создана: {e}")
        arch = os.path.join(work, "archive")
        os.makedirs(arch, exist_ok=True)
        os.replace(os.path.dirname(book), os.path.join(arch, f"test_{t0:%Y%m%d-%H%M%S}"))
        state.put("promote", dict(ok=True, time=t0.isoformat(timespec="seconds"), book=target, hash=ph),
                  f"Рабочая WMS создана: {target}. С этого момента все движения — только в новой WMS")
        print(f"OK: рабочая WMS — {target}")
        return 0


# ================================================================ CHECK and STATUS commands, CLI

def overrides(a):
    return {k: getattr(a, k, None) for k in DEFAULT_SETTINGS}


def cmd_check(a):
    work = os.path.abspath(a.work)
    os.makedirs(work, exist_ok=True)
    state = State(work)
    prog = Progress(work, "check")
    with Busy(work):
        s, sp = load_settings(a.settings, overrides(a))
        prog(0, 2, "чтение данных", True)
        o = office("c")
        try:
            st = read_staging(o, a.staging)
            cand = inspect_candidate(o, a.candidate, a.allow_test_candidate) if a.candidate else {}
        finally:
            office_done(o)
        prog(1, 2, "проверка", True)
        an = Analysis(st, s, max_qty=cand.get("max_qty", MAX_QTY_DEFAULT) if cand else MAX_QTY_DEFAULT, candidate=cand)
        an.settings_problems = sp
        an.run()
        write_check_outputs(work, an)
        src = s.get("source_file") or ""
        state.put("check", dict(ok=not an.blockers(), time=now().isoformat(timespec="seconds"), blockers=len(an.blockers()),
                                warnings=len(an.warnings()), input=input_hash(st.hash, s), staging=st.hash, source_file=src,
                                source_hash=sha256_file(src) if src and os.path.isfile(src) else ""),
                  (f"Проверка: блокирующих замечаний {len(an.blockers())} — исправьте красные строки" if an.blockers() else
                   f"Проверка: блокирующих замечаний нет (требуют проверки: {len(an.warnings())}) — можно «СОЗДАТЬ ТЕСТОВУЮ WMS»"))
        prog(2, 2, "готово", True)
        m = {k: v for k, _, v, _, _ in an.metrics}
        print(f"{'СТОП' if an.blockers() else 'OK'}: строк {m.get('rows_total')}, заказов {m.get('orders')}, позиций {m.get('positions')}, "
              f"ЕИ {m.get('eis')}, NEXT_EI {m.get('next_ei')}; блокирующих {len(an.blockers())}, требуют проверки {len(an.warnings())}")
        return 1 if an.blockers() else 0


def cmd_status(a):
    work = os.path.abspath(a.work)
    State(work).save() if os.path.isdir(work) else None
    p = os.path.join(work, "state.txt")
    print(open(p, encoding="utf-8").read() if os.path.exists(p) else "нет данных")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="mode", required=True)
    for m in ("check", "build", "verify", "promote", "status"):
        p = sub.add_parser(m)
        p.add_argument("--work", required=True)
        if m in ("check", "build", "verify", "promote"):
            p.add_argument("--staging", required=(m != "promote"))
            p.add_argument("--settings")
            for k in DEFAULT_SETTINGS:
                p.add_argument("--" + k.replace("_", "-"), dest=k)
        if m in ("check", "build"):
            p.add_argument("--candidate", required=(m == "build"))
            p.add_argument("--allow-test-candidate", action="store_true", help=argparse.SUPPRESS)
        if m == "build":
            p.add_argument("--fault-at", type=int, help=argparse.SUPPRESS)       # tests: the k-th operation fails among its writes
            p.add_argument("--kill-at", type=int, help=argparse.SUPPRESS)        # tests: LibreOffice dies before the k-th operation
        if m == "promote":
            p.add_argument("--to", required=True)
            p.add_argument("--name", default="WMS_PROD.ods")
    a = ap.parse_args(argv)
    work = os.path.abspath(a.work)
    try:
        rc = dict(check=cmd_check, build=cmd_build, verify=cmd_verify, promote=cmd_promote, status=cmd_status)[a.mode](a)
    except StopRun as e:
        try:
            State(work).put(a.mode + "_stop", dict(time=now().isoformat(timespec="seconds"), text=str(e)), str(e))
        except OSError:
            pass
        print(f"СТОП: {e}", file=sys.stderr)
        rc = 1
    except Exception as e:
        traceback.print_exc()
        try:
            State(work).put(a.mode + "_error", dict(time=now().isoformat(timespec="seconds"), text=str(e)), f"Ошибка инструмента: {e}")
        except OSError:
            pass
        print(f"ОШИБКА: {e}", file=sys.stderr)
        rc = 2
    finally:
        try:
            Progress(work, a.mode)(1, 1, "готово", True)
        except Exception:
            pass
    return rc


if __name__ == "__main__":
    sys.exit(main())
