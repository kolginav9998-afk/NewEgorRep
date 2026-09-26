"""Dry-run of a future migration of existing EIs into a WMS book (Core Phase 5: validators only — nothing is written).

    python3 tools/migration_dryrun.py INPUT.csv|INPUT.ods|INPUT.xlsx [--book WMS.ods] [--places places.txt]
                                      [--units units.txt] [--report report.md] [--json report.json]

The input is the table the next stage (the migration of the real 500+ EIs, D-056) will load: one row per physical batch.
Columns by the header (order and case do not matter, synonyms accepted):
  ЕИ               the code the batch already carries («123», «ЕИ-123», «EI-00000123»); empty — the migration gives a new one;
  Наименование     required;
  Артикул          required for a part («Детали»: one article — one EI);
  Единица          required (no conversion of units: the unit of an EI is final);
  Количество       the current balance, required (WMS rules: up to 3 decimals, «1,500» is ambiguous);
  Место            the current place;
  Категория, Старая маркировка, Комментарий — optional;
  Тип источника    Поставщик / Офис / Производство / Детали / Старый склад / Иной приход; empty — a blocking finding (D-087),
                   unless the whole transfer names one explicitly: --default-source "Старый склад".
Checks, each finding with its row number:
  1. rows without key data (name, unit, quantity; a part without an article);
  2. EIs: not an EI, the same EI twice, an EI already in the book; the largest EI → the NEXT_EI the book will need;
  3. one article of a part at several EIs (inside the input, or between the input and the parts of the book) — after a
     migration such an article is refused («найдено несколько ЕИ») until it is reviewed;
  4. unknown units (neither used in the book nor in the list); unknown places (when a list or the book is given);
  5. quantities: not a number, negative, more than 3 decimals, ambiguous; zero — a warning;
  6. an unknown or empty kind of source (empty is never guessed — only an explicit --default-source fills it).
The book (optional) is opened read-only with macros disabled and never stored; its file is checked unchanged at the end.
Exit code: 0 — no blocking finding, 1 — blocking findings, 2 — the input or the book could not be read.
"""
import argparse
import csv
import hashlib
import json
import os
import re
import sys
import tempfile

EI_DIGITS = 8
DEFAULT_UNITS = {"шт", "м", "кг", "л", "упак", "пар", "компл", "рул", "м2", "м3", "т", "г", "мл", "см", "мм", "лист", "набор", "пог. м", "бух"}
SOURCES = {"поставщик": "Поставщик", "офис": "Офис", "производство": "Производство", "детали": "Детали", "деталь": "Детали",
           "старый склад": "Старый склад", "иной": "Иной приход", "иной приход": "Иной приход"}
COLUMNS = {
    "ei": ("еи", "ei", "внутренний код", "код еи", "код"),
    "name": ("наименование", "название", "товар", "name"),
    "art": ("артикул", "article", "art", "sku"),
    "unit": ("единица", "единица измерения", "ед", "ед.", "ед. изм.", "ед.изм.", "unit"),
    "qty": ("количество", "остаток", "кол-во", "qty", "количество (остаток)"),
    "place": ("место", "место хранения", "ячейка", "place"),
    "cat": ("категория", "category"),
    "src": ("тип источника", "источник", "тип", "source"),
    "mark": ("старая маркировка", "маркировка", "mark"),
    "note": ("комментарий", "примечание", "comment"),
}
_LATIN = {"А": "A", "В": "B", "Е": "E", "Ё": "E", "К": "K", "М": "M", "Н": "H", "О": "O", "Р": "P", "С": "C", "Т": "T", "У": "Y", "Х": "X"}
_DROP = set("\t\n\r       ­")


# ---------------------------------------------------------------- the rules of WMS (written for this tool)

def article_key(s):
    """the key of a part's article as the index of WMS keeps it: upper case, no spaces, dashes unified, Cyrillic lookalikes
    as Latin (WmsSpecial.ArticleKey)"""
    out = []
    for ch in str(s or "").upper():
        if ch in _DROP:
            continue
        if "‐" <= ch <= "―" or ch == "−":
            ch = "-"
        out.append(_LATIN.get(ch, ch))
    return "".join(out)


def normalize_ei(v):
    """(n, canonical EI, "") or (None, None, reason) — the input forms WMS accepts (WmsIssue.NormalizeEI)"""
    if isinstance(v, float):
        if v != int(v) or not 1 <= v <= 99999999:
            return None, None, f"«{v}» не похоже на номер ЕИ"
        n = int(v)
    else:
        t = str(v).strip().upper()
        if len(t) >= 2 and t[0] in "ЕE" and t[1] in "ИI":
            t = t[2:]
            if t[:1] in ("-", " ", "–"):
                t = t[1:]
        if not t or len(t) > EI_DIGITS or not t.isdigit():
            return None, None, f"«{str(v).strip()}» не похоже на номер ЕИ (пример: ЕИ-00000123 или 123)"
        n = int(t)
        if n < 1:
            return None, None, "номер ЕИ должен быть больше 0"
    return n, f"ЕИ-{n:08d}", ""


def parse_qty(v):
    """(value, "") or (None, reason): a number, or text by the rules of WMS (D-030): digits, one «,» or «.», up to 3 decimals;
    «1,500» (3 digits after the separator, 1–3 before) is ambiguous"""
    if isinstance(v, float):
        if round(v, 3) != v:
            return None, f"«{v}» — больше 3 знаков после запятой"
        return v, ""
    t = str(v).strip().replace(" ", "").replace(" ", "")
    if not t:
        return None, "количество не указано"
    if t.startswith("-"):
        return None, f"«{v}» — отрицательное количество"
    m = re.fullmatch(r"(\d+)(?:[.,](\d+))?", t)
    if not m:
        return None, f"«{v}» — не число"
    ip, fp = m.group(1), m.group(2) or ""
    if len(fp) == 3 and 1 <= len(ip) <= 3:
        return None, f"«{v}» — неоднозначно (тысячи или дробь?): запишите «{ip}{fp}» или «{ip},{fp.rstrip('0') or '0'}»"
    if len(fp.rstrip("0")) > 3:
        return None, f"«{v}» — больше 3 знаков после запятой"
    return float(ip + ("." + fp if fp else "")), ""


def source_of(v):
    """the kind of source as WMS names it; "" when the cell is empty, None when the text is not a kind of source"""
    t = re.sub(r"\s+", " ", str(v or "").strip().lower())
    if not t:
        return ""
    return SOURCES.get(t)


# ---------------------------------------------------------------- reading

def norm_header(h):
    return re.sub(r"\s+", " ", str(h or "").strip().lower()).rstrip(":")


def map_columns(header):
    idx = {}
    for i, h in enumerate(header):
        nh = norm_header(h)
        for key, names in COLUMNS.items():
            if nh in names and key not in idx:
                idx[key] = i
    return idx


def read_csv(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1251"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("кодировка файла не UTF-8 и не Windows-1251")
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=";,\t")
    except csv.Error:
        dialect = csv.excel
        dialect.delimiter = ";"
    return [row for row in csv.reader(text.splitlines(), dialect)]


class Calc:
    """one LibreOffice process for reading ODS/XLSX files (read-only, macros never run, nothing stored)"""

    def __init__(self):
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from wmslo import Office
        self.tmp = tempfile.mkdtemp(prefix="wms_dryrun_")
        self.o = Office("dryrun", self.tmp)

    def open(self, path):
        doc = self.o.load(os.path.abspath(path), macros=0, hidden=True, ReadOnly=True)
        if doc is None:
            raise ValueError(f"не удалось открыть {path}")
        return doc

    @staticmethod
    def used(sheet, ncols=None):
        cur = sheet.createCursor()
        cur.gotoEndOfUsedArea(False)
        a = cur.getRangeAddress()
        return [list(r) for r in sheet.getCellRangeByPosition(0, 0, ncols - 1 if ncols else a.EndColumn, a.EndRow).getDataArray()]

    def close(self, doc):
        try:
            doc.close(True)
        except Exception:
            pass

    def terminate(self):
        self.o.terminate()


def read_book(calc, path):
    """what the dry-run needs from a WMS book: the EIs of «Наличие» (name, article, unit, place, kind of source), the index of
    the parts (_ART), NEXT_EI"""
    doc = calc.open(path)
    try:
        sh = doc.Sheets
        stock = calc.used(sh.getByName("Наличие"), 10) if sh.hasByName("Наличие") else []
        art = calc.used(sh.getByName("_ART"), 3) if sh.hasByName("_ART") else []
        nxt = None
        if sh.hasByName("_SYS"):
            for row in calc.used(sh.getByName("_SYS"), 2):
                if row[0] == "NEXT_EI" and isinstance(row[1], float):
                    nxt = int(row[1])
        eis = {}
        for r, row in enumerate(stock[1:], start=1):
            if row[0]:
                eis[str(row[0])] = dict(row=r + 1, name=row[1], art=row[2], unit=row[3], place=row[5], stype=row[9] if len(row) > 9 else "")
        index = {}
        for row in art[1:]:
            if row[0]:
                index.setdefault(str(row[0]), set()).add(str(row[1]))
        return dict(eis=eis, index=index, next_ei=nxt)
    finally:
        calc.close(doc)


# ---------------------------------------------------------------- the checks

def check(rows, book=None, places=None, units=None, default_source=None):
    """rows: the table (the first row — the header). Returns (findings, summary); a finding: dict(row, level, kind, text)"""
    F = []

    def add(r, level, kind, text):
        F.append(dict(row=r, level=level, kind=kind, text=text))

    if not rows:
        return [dict(row=0, level="BLOCK", kind="файл", text="таблица пуста")], {}
    idx = map_columns(rows[0])
    missing = [k for k in ("name", "unit", "qty") if k not in idx]
    if missing:
        names = {"name": "Наименование", "unit": "Единица", "qty": "Количество"}
        return [dict(row=1, level="BLOCK", kind="заголовок", text="нет колонок: " + ", ".join(names[k] for k in missing))], {}
    known_units = {u.lower() for u in DEFAULT_UNITS | set(units or [])}
    book_places = set()
    if book:
        for e in book["eis"].values():
            if e["unit"]:
                known_units.add(str(e["unit"]).lower())
            if e["place"]:
                book_places.add(str(e["place"]))
    known_places = set(places or []) | book_places if (places or book) else None

    def cell(row, key):
        i = idx.get(key)
        return row[i] if i is not None and i < len(row) else ""

    seen_ei, art_ei, by_src = {}, {}, {}
    n_rows = with_ei = without_ei = 0
    max_ei = 0
    for r, row in enumerate(rows[1:], start=2):
        if all(str(x).strip() == "" for x in row):
            continue
        n_rows += 1
        name, art, unit = str(cell(row, "name")).strip(), str(cell(row, "art")).strip(), str(cell(row, "unit")).strip()
        src = source_of(cell(row, "src"))
        if src == "":
            if default_source:
                src = default_source
            else:
                add(r, "BLOCK", "тип источника", "тип источника не указан — строка не переносится молча как «Старый склад»: укажите тип в таблице "
                    "или явно для всего переноса (--default-source \"Старый склад\")")
                src = None
        elif src is None:
            add(r, "BLOCK", "тип источника", f"тип источника «{cell(row, 'src')}» неизвестен (Поставщик, Офис, Производство, Детали, Старый склад, Иной приход)")
        by_src[src or "?"] = by_src.get(src or "?", 0) + 1
        if not name:
            add(r, "BLOCK", "нет ключевых данных", "нет наименования")
        if not unit:
            add(r, "BLOCK", "нет ключевых данных", "нет единицы измерения")
        elif unit.lower() not in known_units:
            hint = f" (может быть «{unit.lower().rstrip('.')}»?)" if unit.lower().rstrip(".") in known_units else ""
            add(r, "WARN", "неизвестная единица", f"единица «{unit}» не используется в книге и не входит в список{hint}")
        q, why = parse_qty(cell(row, "qty"))
        if q is None:
            add(r, "BLOCK", "количество" if "не указано" not in why else "нет ключевых данных", why)
        elif q == 0:
            add(r, "WARN", "количество", "нулевой остаток — перенос ЕИ без товара")
        place = str(cell(row, "place")).strip()
        if known_places is not None:
            if not place:
                add(r, "WARN", "место", "место не указано")
            elif place not in known_places:
                add(r, "WARN", "неизвестное место", f"место «{place}» не встречается в книге и в списке мест")
        ei_raw = cell(row, "ei")
        canon = None
        if str(ei_raw).strip() == "":
            without_ei += 1
        else:
            n, canon, why = normalize_ei(ei_raw)
            if canon is None:
                add(r, "BLOCK", "ЕИ", why)
            else:
                with_ei += 1
                max_ei = max(max_ei, n)
                if canon in seen_ei:
                    add(r, "BLOCK", "дубль ЕИ", f"{canon} уже в строке {seen_ei[canon]}")
                else:
                    seen_ei[canon] = r
                if book and canon in book["eis"]:
                    add(r, "BLOCK", "ЕИ уже в книге", f"{canon} уже есть в книге («{book['eis'][canon]['name']}», «Наличие» строка {book['eis'][canon]['row']})")
        if src == "Детали":
            key = article_key(art)
            if not key:
                add(r, "BLOCK", "нет ключевых данных", "деталь без артикула — по артикулу WMS находит ЕИ детали")
            else:
                art_ei.setdefault(key, []).append((r, canon or f"(новый ЕИ, строка {r})"))
    # one article of a part — one EI (inside the input and against the book)
    for key, lst in sorted(art_ei.items()):
        eis = sorted({e for _, e in lst})
        if len(eis) > 1:
            add(lst[0][0], "BLOCK", "артикул → несколько ЕИ", f"артикул «{key}» у разных ЕИ: {', '.join(eis)} (строки {', '.join(str(x) for x, _ in lst)}) — "
                "после переноса такой артикул не проводится до разбора")
        if book:
            have = set(book["index"].get(key, set()))
            have |= {e for e, v in book["eis"].items() if v["stype"] == "Детали" and article_key(v["art"]) == key}
            other = sorted(have - set(eis))
            if other:
                add(lst[0][0], "BLOCK", "артикул → несколько ЕИ", f"артикул «{key}» уже у детали книги {', '.join(other)}")
    book_max = 0
    if book:
        for e in book["eis"]:
            n, _, _ = normalize_ei(e)
            if n:
                book_max = max(book_max, n)
    need = max(max_ei, book_max) + 1
    if book and book.get("next_ei"):
        need = max(need, book["next_ei"])
    summary = dict(rows=n_rows, with_ei=with_ei, without_ei=without_ei, max_ei_input=f"ЕИ-{max_ei:08d}" if max_ei else "",
                   max_ei_book=f"ЕИ-{book_max:08d}" if book_max else "", next_ei_needed=need, by_source=by_src,
                   parts_articles=len(art_ei), blocking=sum(1 for f in F if f["level"] == "BLOCK"), warnings=sum(1 for f in F if f["level"] == "WARN"),
                   places_checked=known_places is not None)
    return F, summary


def report_md(path_in, summary, F, book_path, unchanged):
    if book_path is None:
        state = "Книга не задана — проверена только таблица."
    elif unchanged:
        state = f"Книга {os.path.basename(book_path)} открыта только для чтения, без макросов; файл до и после проверки совпадает — книга не изменялась."
    else:
        state = f"ВНИМАНИЕ: файл книги {os.path.basename(book_path)} изменился во время проверки."
    lines = [f"# Предварительная проверка переноса (dry-run): {os.path.basename(path_in)}", "", state, "",
             "| Показатель | Значение |", "|---|---|"]
    names = dict(rows="строк с данными", with_ei="с ЕИ", without_ei="без ЕИ (получат новый)", max_ei_input="максимальный ЕИ таблицы",
                 max_ei_book="максимальный ЕИ книги", next_ei_needed="NEXT_EI после переноса (не ниже)", by_source="по типам источника",
                 parts_articles="артикулов деталей", blocking="блокирующих замечаний", warnings="предупреждений", places_checked="места проверены", default_source="тип источника для пустых строк (задан явно)")
    for k, v in summary.items():
        lines.append(f"| {names.get(k, k)} | {v} |")
    lines += ["", "| Строка | Уровень | Проверка | Замечание |", "|---|---|---|---|"]
    for f in sorted(F, key=lambda x: (x["row"], x["level"])):
        lines.append(f"| {f['row']} | {'БЛОК' if f['level'] == 'BLOCK' else 'предупр.'} | {f['kind']} | {f['text'].replace('|', '¦')} |")
    if not F:
        lines.append("| — | — | — | замечаний нет |")
    return "\n".join(lines) + "\n"


def sha256(path):
    return hashlib.sha256(open(path, "rb").read()).hexdigest()


def main(argv=None):
    ap = argparse.ArgumentParser(description="dry-run переноса ЕИ в книгу WMS: только проверки, без записи")
    ap.add_argument("input")
    ap.add_argument("--book")
    ap.add_argument("--places")
    ap.add_argument("--units")
    ap.add_argument("--default-source", help="тип источника для строк с пустым «Тип источника» (явное правило этого переноса)")
    ap.add_argument("--report")
    ap.add_argument("--json")
    a = ap.parse_args(argv)
    default_source = None
    if a.default_source is not None:
        default_source = source_of(a.default_source)
        if not default_source:
            print(f"ОШИБКА: --default-source «{a.default_source}» — не тип источника (Поставщик, Офис, Производство, Детали, Старый склад, "
                  "Иной приход)", file=sys.stderr)
            return 2
    calc = None
    h0 = sha256(a.book) if a.book else None
    try:
        if a.input.lower().endswith(".csv"):
            rows = read_csv(a.input)
        else:
            calc = Calc()
            doc = calc.open(a.input)
            rows = calc.used(doc.Sheets.getByIndex(0))
            calc.close(doc)
        book = None
        if a.book:
            calc = calc or Calc()
            book = read_book(calc, a.book)
    except Exception as e:
        print(f"ОШИБКА: {e}", file=sys.stderr)
        return 2
    finally:
        if calc:
            calc.terminate()
    places = [x.strip() for x in open(a.places, encoding="utf-8")] if a.places else None
    units = [x.strip() for x in open(a.units, encoding="utf-8")] if a.units else None
    F, summary = check(rows, book, [p for p in places or [] if p] or None, [u for u in units or [] if u], default_source)
    summary["default_source"] = default_source or ""
    unchanged = (sha256(a.book) == h0) if a.book else None
    if a.book and not unchanged:
        F.append(dict(row=0, level="BLOCK", kind="книга", text="файл книги изменился во время проверки"))
    md = report_md(a.input, summary, F, a.book, unchanged)
    if a.report:
        open(a.report, "w", encoding="utf-8").write(md)
    if a.json:
        json.dump(dict(summary=summary, findings=F, book_unchanged=unchanged), open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(md)
    return 1 if summary.get("blocking", 1) or any(f["level"] == "BLOCK" for f in F) else 0


if __name__ == "__main__":
    sys.exit(main())
