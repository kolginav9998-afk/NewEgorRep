"""Migration of the existing EIs of the old warehouse into a WMS book (FINAL WMS MARATHON §4; D-056, D-057, D-087, П-11).

    python3 tools/migrate.py dry-run TABLE --book BOOK [--default-source T] [--places F | --accept-table-places] [--units F]
                                       [--report R]
    python3 tools/migrate.py copy    TABLE --book BOOK --out DIR [the same options]
    python3 tools/migrate.py verify  TABLE --dir DIR [--default-source T]
    python3 tools/migrate.py promote --dir DIR --to TARGET_DIR [--name WMS_PROD.ods]

TABLE — the table of the transfer (CSV UTF-8 / Windows-1251, ODS, XLSX; one physical batch — one row) with the columns of
tools/migration_dryrun.py: ЕИ, Наименование, Артикул, Единица, Количество, Место, Категория, Тип источника, Старая
маркировка, Комментарий. BOOK — the WMS book the EIs go to (normally a new production book); it is never changed.

DRY RUN      the checks of tools/migration_dryrun.py plus the stops of the migration (nothing is written): an EI conflict
             (twice in the table, already in the book), one article of a part at several EIs, an unknown unit, an unknown
             place, an empty source type without --default-source, a negative or wrong quantity, a row without key data.
             Places are known from --places (one per line) or from the book; --accept-table-places accepts every place
             of the table explicitly (listed in the report). The plan: every row keeps its EI; a row without an EI gets a
             new number above every number of the table and of the book; NEXT_EI after the transfer.
MIGRATE COPY only after a clean dry-run (repeated here). The book is closed in WMS (no wms.lock) and has no journal tail.
             DIR/backup/ gets a copy of the book and its journal with SHA-256; DIR gets the working copy (the book and
             its journal); WMS opens the copy (registered there as the working file) and posts every row by one MIGRATE
             operation of the core (WmsMigrate.MigrateEI): the EI keeps its number, a part enters the index, NEXT_EI is
             raised, the journal keeps the origin (table file | SHA-256 | row) and the old marking. Then the index of the
             parts is rebuilt and the self-check must have no errors. Any refusal stops the transfer: DIR is removed as a
             whole and the book is as it was — the whole transfer or nothing (report DIR_MIGRATION_FAILED.md).
VERIFY       the copy against the table: every row once in «Наличие» (number, name, article, unit, quantity, place,
             category, state, source type), one MIGRATE operation per row with the SHA-256 of the table, NEXT_EI above
             every EI, the self-check without errors and the independent oracle (the whole journal replayed by the
             rules of the task, tests/adjust_oracle.py). The SHA-256 of the verified book is recorded.
PROMOTE      only the verified, unchanged copy: the book and its journal to TARGET_DIR (it must not exist) under --name;
             registered there as the working WMS; its state must be «работа разрешена».

Reports: DIR/MIGRATION_REPORT.md, DIR/VERIFY_REPORT.md, TARGET_DIR/PROMOTE_REPORT.md. Exit code: 0 done, 1 a stop by the
rules (nothing changed), 2 the input or the book could not be read / a tool error.
"""
import argparse
import datetime
import hashlib
import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "tests"))
import migration_dryrun as dr  # noqa: E402

JOURNAL_DIR = "WMS_Journal"
LOCK_FILE = "wms.lock"
SELF_CHECK_OK = "САМОПРОВЕРКА WMS: ошибок 0"
STATE_REVIEW = "Требует разбора"
STATE_ACTIVE = "Активен"
# a finding of tools/migration_dryrun.py that is only a warning there but stops a migration (the task, §4)
STOP_KINDS = {"неизвестная единица", "неизвестное место"}


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def now():
    return datetime.datetime.now()


def report(path, title, lines):
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"# {title}\n\n**Дата:** {now():%d.%m.%Y %H:%M}\n\n" + "\n".join(lines) + "\n")


def read_table(path, calc):
    if path.lower().endswith(".csv"):
        return dr.read_csv(path)
    doc = calc.open(path)
    try:
        return calc.used(doc.Sheets.getByIndex(0))
    finally:
        calc.close(doc)


def qty_text(v):
    """the quantity as WMS parses it: a comma, no trailing zeros ("0" for a zero balance)"""
    q, _ = dr.parse_qty(v)
    s = f"{float(q):.3f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def text(v):
    """a cell as text: a whole number without «.0» (an article or a marking typed as a number)"""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v if v is not None else "").strip()


def places_of_table(rows):
    idx = dr.map_columns(rows[0]) if rows else {}
    i = idx.get("place")
    return sorted({text(r[i]) for r in rows[1:] if i is not None and i < len(r) and text(r[i])})


def load_list(path):
    return [x.strip() for x in open(path, encoding="utf-8-sig") if x.strip()]


def checks(a, calc):
    """(rows, book, findings, summary, blocking, places) — the dry-run of tools/migration_dryrun.py with the stops of a migration"""
    rows = read_table(a.table, calc)
    book = dr.read_book(calc, a.book)
    places = load_list(a.places) if a.places else None
    if a.accept_table_places:
        places = sorted(set(places or []) | set(places_of_table(rows)))
    units = load_list(a.units) if a.units else None
    F, summary = dr.check(rows, book, places, units, a.default_source)
    for f in F:
        if f["kind"] in STOP_KINDS:
            f["level"] = "BLOCK"
    if summary:
        summary["blocking"] = sum(1 for f in F if f["level"] == "BLOCK")
        summary["warnings"] = sum(1 for f in F if f["level"] != "BLOCK")
    blocking = [f for f in F if f["level"] == "BLOCK"]
    return rows, book, F, summary, blocking, places


def plan(rows, book, default_source):
    """the rows to post (dict: row, n, ei, name, art, unit, qty, place, cat, stype, mark, new) and NEXT_EI after them; a row
    without an EI gets a number above every number of the table and of the book"""
    idx = dr.map_columns(rows[0])

    def cell(r, k):
        return r[idx[k]] if k in idx and idx[k] < len(r) else ""

    out, used = [], []
    for i, r in enumerate(rows[1:], start=2):
        if all(str(x).strip() == "" for x in r):
            continue
        n = None
        if text(cell(r, "ei")):
            n, _, _ = dr.normalize_ei(cell(r, "ei"))
            used.append(n)
        src = dr.source_of(cell(r, "src"))
        if src == "":
            src = default_source
        out.append(dict(row=i, n=n, name=text(cell(r, "name")), art=text(cell(r, "art")), unit=text(cell(r, "unit")),
                        qty=qty_text(cell(r, "qty")), place=text(cell(r, "place")), cat=text(cell(r, "cat")), stype=src,
                        mark=text(cell(r, "mark")), new=n is None))
    top = max(used + [0])
    if book:
        for e in book["eis"]:
            m, _, _ = dr.normalize_ei(e)
            top = max(top, m or 0)
        top = max(top, (book.get("next_ei") or 1) - 1)
    for p in out:
        if p["n"] is None:
            top += 1
            p["n"] = top
        p["ei"] = f"ЕИ-{p['n']:08d}"
    return out, top + 1


def findings_md(F):
    out = ["| Строка | Уровень | Проверка | Замечание |", "|---:|---|---|---|"]
    for f in sorted(F, key=lambda x: (x["row"], x["level"])):
        out.append(f"| {f['row']} | {'БЛОК' if f['level'] == 'BLOCK' else 'предупр.'} | {f['kind']} | {str(f['text']).replace('|', '¦')} |")
    return out


def head_lines(a, rows, F, blocking, places):
    lines = [f"**Таблица:** `{os.path.abspath(a.table)}` (SHA-256 `{sha256(a.table)}`)  ",
             f"**Книга:** `{os.path.abspath(a.book)}` (SHA-256 `{sha256(a.book)}`)  ",
             f"**Тип источника для пустых строк:** {a.default_source or '— (пустой тип — блокирующая ошибка, D-087)'}  ",
             "**Места:** " + ("приняты из таблицы явно (--accept-table-places): " + ", ".join(places) if a.accept_table_places and places
                              else f"по списку `{a.places}` и книге" if a.places else "по книге (список мест не задан)") + "\n",
             f"Строк таблицы: {max(len(rows) - 1, 0)}; блокирующих замечаний: **{len(blocking)}**; предупреждений: {len(F) - len(blocking)}.", ""]
    if F:
        lines += ["## Замечания", ""] + findings_md(F) + [""]
    return lines


def plan_md(items, next_ei):
    new = sum(1 for p in items if p["new"])
    lines = ["## План переноса", "", f"Строк к переносу: {len(items)}; сохраняют свой номер: {len(items) - new}; новых номеров (строки без ЕИ — "
             f"нужны новые наклейки): {new}; NEXT_EI после переноса: {next_ei}.", "",
             "| Строка | ЕИ | Наименование | Артикул | Ед. | Количество | Место | Тип источника |", "|---:|---|---|---|---|---:|---|---|"]
    for p in items:
        lines.append(f"| {p['row']} | {p['ei']}{' (новый)' if p['new'] else ''} | {p['name']} | {p['art']} | {p['unit']} | {p['qty']} | "
                     f"{p['place']} | {p['stype']} |".replace("\n", " "))
    return lines


def cmd_dry_run(a):
    calc = dr.Calc()
    try:
        h0 = sha256(a.book)
        rows, book, F, summary, blocking, places = checks(a, calc)
    finally:
        calc.terminate()
    lines = head_lines(a, rows, F, blocking, places)
    if sha256(a.book) != h0:
        blocking.append(dict(row=0, level="BLOCK", kind="книга", text="файл книги изменился во время проверки"))
        lines.append("**ВНИМАНИЕ: файл книги изменился во время проверки.**")
    if not blocking:
        items, next_ei = plan(rows, book, a.default_source)
        lines += plan_md(items, next_ei)
    else:
        lines += ["Перенос с такими замечаниями не выполняется: исправьте таблицу (или задайте правило явно) и повторите dry-run."]
    if a.report:
        report(a.report, "DRY RUN переноса", lines)
    print("\n".join(lines[:6]))
    return 1 if blocking else 0


def journal_of(book_path):
    return os.path.join(os.path.dirname(os.path.abspath(book_path)), JOURNAL_DIR)


def copy_journal(src_dir, dst_dir):
    """the journal files without the lock (the lock belongs to a running WMS, never to a copy)"""
    os.makedirs(dst_dir, exist_ok=True)
    if os.path.isdir(src_dir):
        for f in sorted(os.listdir(src_dir)):
            if f != LOCK_FILE and os.path.isfile(os.path.join(src_dir, f)):
                shutil.copy2(os.path.join(src_dir, f), os.path.join(dst_dir, f))


def open_wms(o, path):
    """the book opened with WMS; (doc, state line): registered here when WMS says it is not the working file"""
    doc = o.load(path, macros=4)
    if doc is None:
        return None, "книга не открылась"
    st = o.basic(doc, "WmsCore", "StateDump") or ""
    if "STATE=CLEAN" not in st and "REG=other" in st:
        o.basic(doc, "WmsCore", "ActionRegisterHere", False)
        st = o.basic(doc, "WmsCore", "StateDump") or ""
    return doc, st


def close_doc(o, doc, save):
    if save:
        doc.store()
    if doc.isModified():
        doc.setModified(False)
    try:
        o.dispatch(doc, ".uno:CloseDoc")
    except Exception:
        pass


def cmd_copy(a):
    from wmslo import Office
    out = os.path.abspath(a.out)
    if os.path.exists(out):
        print(f"ОШИБКА: {out} уже существует — перенос выполняется в новую папку", file=sys.stderr)
        return 2
    if os.path.exists(os.path.join(journal_of(a.book), LOCK_FILE)):
        print(f"СТОП: книга открыта в WMS или закрыта аварийно (есть {JOURNAL_DIR}/{LOCK_FILE}) — закройте WMS или снимите блокировку "
              "и повторите", file=sys.stderr)
        return 1
    calc = dr.Calc()
    try:
        rows, book, F, summary, blocking, places = checks(a, calc)
    finally:
        calc.terminate()
    if blocking:
        print(f"СТОП: dry-run нашёл блокирующих замечаний: {len(blocking)} — перенос не начат (см. «dry-run»)", file=sys.stderr)
        return 1
    items, next_ei = plan(rows, book, a.default_source)
    th, bh = sha256(a.table), sha256(a.book)
    name = os.path.basename(a.book)
    t0 = now()
    ts = t0.strftime("%Y%m%d-%H%M%S")
    os.makedirs(os.path.join(out, "backup"))
    bk = os.path.join(out, "backup", f"{os.path.splitext(name)[0]}_{ts}.ods")
    shutil.copy2(a.book, bk)
    copy_journal(journal_of(a.book), os.path.join(out, "backup", JOURNAL_DIR))
    copy = os.path.join(out, name)
    shutil.copy2(a.book, copy)
    copy_journal(journal_of(a.book), os.path.join(out, JOURNAL_DIR))
    prof = tempfile.mkdtemp(prefix="wms_migrate_")
    o = Office(f"migrate{os.getpid()}", prof)
    done, failed, idx, sc, bm, doc = [], None, "", "", "", None
    try:
        doc, st = open_wms(o, copy)
        if doc is None or "STATE=CLEAN" not in st:
            failed = f"копия книги не стала рабочей WMS: {st[:240]}"
        else:
            # the backup «before a migration» of WMS itself (spec §26) in the copy's WMS_Backups/
            bm = str(o.basic(doc, "WmsBackup", "BeforeMigration"))
            if bm.startswith("ОШИБКА"):
                failed = f"резервная копия перед переносом: {bm}"
        if not failed:
            if a.fault_at:
                a.fault_at = int(a.fault_at)
            origin = os.path.basename(a.table)
            for k, p in enumerate(items, start=1):
                if a.fault_at == k:
                    o.basic(doc, "WmsCore", "TestSetFault", 2, 1)       # test books only: the operation fails among its writes
                res = str(o.basic(doc, "WmsMigrate", "MigrateEI", p["ei"], p["name"], p["art"], p["unit"], p["qty"], p["place"], p["cat"],
                                  p["stype"], p["mark"], f"{origin}|{th}|{p['row']}"))
                if not res.startswith("OK:"):
                    failed = f"строка {p['row']} ({p['ei']}): {res}"
                    break
                done.append((p, int(res[3:])))
            if not failed:
                idx = str(o.basic(doc, "WmsSpecial", "ArticleIndexRebuild"))
                sc = str(o.basic(doc, "WmsCore", "ActionSelfCheck") or "")
                if idx.startswith("ОШИБКА") or not sc.startswith(SELF_CHECK_OK):
                    failed = f"после переноса: индекс деталей — {idx[:200]}; самопроверка — {sc.splitlines()[0] if sc else '—'}"
        if doc is not None:
            close_doc(o, doc, save=not failed)
    except Exception as e:
        failed = failed or f"ошибка инструмента: {e}"
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    if failed:
        rep = out.rstrip("/\\") + "_MIGRATION_FAILED.md"
        shutil.rmtree(out, ignore_errors=True)
        same = sha256(a.book) == bh
        report(rep, "Перенос НЕ выполнен", [
            f"**Таблица:** `{os.path.abspath(a.table)}` (SHA-256 `{th}`)  ",
            f"**Книга:** `{os.path.abspath(a.book)}` — {'не изменялась' if same else 'ВНИМАНИЕ: ФАЙЛ ИЗМЕНИЛСЯ'} (SHA-256 `{bh}`)", "",
            f"**Причина остановки:** {failed}", "",
            f"Копия `{out}` удалена целиком (проведено до остановки: {len(done)} из {len(items)}) — перенос выполняется только весь или никак. "
            "Исправьте причину и повторите с «dry-run»."])
        print(f"СТОП: {failed}; копия удалена, книга {'не изменялась' if same else 'ИЗМЕНИЛАСЬ'}. Отчёт: {rep}", file=sys.stderr)
        return 1
    secs = (now() - t0).total_seconds()
    lines = [f"**Таблица:** `{os.path.abspath(a.table)}` (SHA-256 `{th}`)  ",
             f"**Исходная книга:** `{os.path.abspath(a.book)}` (SHA-256 `{bh}`) — не изменялась; резервная копия `backup/{os.path.basename(bk)}` "
             f"и `backup/{JOURNAL_DIR}/`  ",
             f"**Копия:** `{name}` и журнал `{JOURNAL_DIR}/` в этой папке  ",
             f"**Тип источника для пустых строк:** {a.default_source or '—'}  ",
             f"**Места:** {'приняты из таблицы явно' if a.accept_table_places else 'по списку и книге'}\n",
             f"Перенесено ЕИ: **{len(done)}** — одна операция MIGRATE на строку, seq {done[0][1] if done else '—'}…{done[-1][1] if done else '—'}; "
             f"сохранили номер: {sum(1 for p, _ in done if not p['new'])}, новых номеров: {sum(1 for p, _ in done if p['new'])}; NEXT_EI: {next_ei}; "
             f"время {secs:.0f} с.", f"Резервная копия WMS перед переносом: {bm}", f"Индекс деталей: {idx}",
             f"Самопроверка: {sc.splitlines()[0] if sc else '—'}", ""]
    if F:
        lines += ["## Предупреждения dry-run", ""] + findings_md(F) + [""]
    lines += plan_md(items, next_ei) + ["", f"Следующий шаг: `python3 tools/migrate.py verify {a.table} --dir {out}`, затем `promote`."]
    report(os.path.join(out, "MIGRATION_REPORT.md"), "MIGRATE COPY", lines)
    print("\n".join(lines[:6]))
    return 0


def find_book(d):
    books = [f for f in os.listdir(d) if f.lower().endswith(".ods")]
    if len(books) != 1:
        raise ValueError(f"в {d} должна быть одна книга .ods, найдено: {books}")
    return os.path.join(d, books[0])


def cmd_verify(a):
    from wmslo import Office
    import adjust_oracle
    import journal_oracle
    d = os.path.abspath(a.dir)
    book_path = find_book(d)
    calc = dr.Calc()
    try:
        rows = read_table(a.table, calc)
    finally:
        calc.terminate()
    th = sha256(a.table)
    items, _ = plan(rows, None, a.default_source)
    migrated = {p["ei"] for p in items}
    prof = tempfile.mkdtemp(prefix="wms_verify_")
    o = Office(f"mverify{os.getpid()}", prof)
    P, info = [], {}
    try:
        doc, st = open_wms(o, book_path)
        if doc is None or "STATE=CLEAN" not in st:
            P.append(f"WMS не в рабочем состоянии: {st[:200]}")
        if doc is not None:
            sh = doc.Sheets
            stock = sh.getByName("Наличие")
            # rows without an EI got numbers after the table's own: they are found in the journal by their row
            entries, _ = journal_oracle.read_journal(os.path.join(d, JOURNAL_DIR))
            mig = [e for e in entries if e["type"] == "MIGRATE"]
            by_row = {}
            for e in mig:
                org = str(e["fields"].get("ORIGIN", "")).split("|")
                if len(org) != 3 or org[1] != th:
                    P.append(f"seq {e['seq']}: происхождение «{e['fields'].get('ORIGIN')}» — не эта таблица (SHA-256)")
                    continue
                by_row.setdefault(int(org[2]), []).append(e)
            if len(mig) != len(items):
                P.append(f"в журнале операций MIGRATE {len(mig)}, строк таблицы {len(items)}")
            for p in items:
                es = by_row.get(p["row"], [])
                if len(es) != 1:
                    P.append(f"строка {p['row']}: операций MIGRATE {len(es)} (нужна одна)")
                    continue
                ei = es[0]["fields"].get("EI")
                if not p["new"] and ei != p["ei"]:
                    P.append(f"строка {p['row']}: перенесена как {ei}, в таблице {p['ei']} — номер изменён")
                    continue
                n = int(str(ei)[-8:])
                row = stock.getCellRangeByPosition(0, n, 9, n).getDataArray()[0]
                q = float(p["qty"].replace(",", "."))
                want = (ei, p["name"], p["art"], p["unit"], q, p["place"], p["cat"], STATE_REVIEW if p["stype"] == "Иной приход" else STATE_ACTIVE,
                        p["stype"])
                got = (row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[9])
                if got != want:
                    P.append(f"строка {p['row']}: в книге {got}, по таблице {want}")
                if not str(row[8]).startswith("Перенос"):
                    P.append(f"строка {p['row']}: источник «{row[8]}» — не «Перенос…»")
            sysd = {r[0]: r[1] for r in sh.getByName("_SYS").getCellRangeByPosition(0, 0, 1, 40).getDataArray()}
            top = max([int(str(e["fields"].get("EI"))[-8:]) for e in mig] + [0])
            if not isinstance(sysd.get("NEXT_EI"), float) or sysd["NEXT_EI"] <= top:
                P.append(f"NEXT_EI {sysd.get('NEXT_EI')!r} не выше перенесённых номеров (до {top})")
            sc = str(o.basic(doc, "WmsCore", "ActionSelfCheck") or "")
            if not sc.startswith(SELF_CHECK_OK):
                P.append("самопроверка: " + (sc.splitlines()[0] if sc else "—"))
            initial = {}
            if sysd.get("MODE") == "TEST":
                from build_ods import synthetic_registry
                initial = {r[0]: r[4] for r in synthetic_registry()}
            OP, info, _ = adjust_oracle.check(doc, os.path.join(d, JOURNAL_DIR), initial)
            P += [f"оракул: {x}" for x in OP[:20]]
            close_doc(o, doc, save=False)
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    bh = sha256(book_path)
    lines = [f"**Таблица:** `{os.path.abspath(a.table)}` (SHA-256 `{th}`)  ", f"**Книга:** `{book_path}`  ",
             f"**SHA-256 проверенной книги:** `{bh}`\n",
             f"Строк таблицы: {len(items)}; перенесённых ЕИ в журнале: {len(migrated)}; расхождений: **{len(P)}**.",
             f"Оракул (журнал воспроизведён независимо): операций {info.get('ops', '—')}, последняя {info.get('last_seq', '—')}.", ""]
    lines += [f"- {x}" for x in P[:60]] or ["Расхождений нет — копию можно переносить в рабочую папку (`promote`)."]
    report(os.path.join(d, "VERIFY_REPORT.md"), "VERIFY переноса", lines)
    print("\n".join(lines[:5]))
    return 1 if P else 0


def cmd_promote(a):
    from wmslo import Office
    src = os.path.abspath(a.dir)
    dst = os.path.abspath(a.to)
    if os.path.exists(dst):
        print(f"ОШИБКА: {dst} уже существует — рабочая папка создаётся заново", file=sys.stderr)
        return 2
    book = find_book(src)
    vr = os.path.join(src, "VERIFY_REPORT.md")
    vt = open(vr, encoding="utf-8").read() if os.path.exists(vr) else ""
    if "расхождений: **0**" not in vt:
        print("СТОП: нет успешной проверки VERIFY — PROMOTE не выполняется", file=sys.stderr)
        return 1
    if f"`{sha256(book)}`" not in vt:
        print("СТОП: книга изменилась после проверки VERIFY — повторите verify", file=sys.stderr)
        return 1
    os.makedirs(dst)
    target = os.path.join(dst, a.name)
    shutil.copy2(book, target)
    copy_journal(os.path.join(src, JOURNAL_DIR), os.path.join(dst, JOURNAL_DIR))
    prof = tempfile.mkdtemp(prefix="wms_promote_")
    o = Office(f"mpromote{os.getpid()}", prof)
    st = ""
    try:
        doc, st = open_wms(o, target)
        if doc is not None:
            close_doc(o, doc, save="STATE=CLEAN" in st)
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    ok = "STATE=CLEAN" in st
    report(os.path.join(dst, "PROMOTE_REPORT.md"), "PROMOTE переноса", [
        f"**Из:** `{src}` (проверено VERIFY)  ", f"**В:** `{target}` и журнал `{JOURNAL_DIR}/`  ",
        f"**Состояние WMS:** {'работа разрешена' if ok else st[:200]}", "",
        "Рабочая WMS — эта книга. Копия переноса и её резервная копия остаются для истории." if ok else "PROMOTE не завершён: см. состояние."])
    if not ok:
        print(f"СТОП: {st[:200]}", file=sys.stderr)
        return 1
    print(f"OK: рабочая WMS — {target}")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="mode", required=True)
    for m in ("dry-run", "copy"):
        p = sub.add_parser(m)
        p.add_argument("table")
        p.add_argument("--book", required=True)
        p.add_argument("--default-source")
        g = p.add_mutually_exclusive_group()
        g.add_argument("--places")
        g.add_argument("--accept-table-places", action="store_true")
        p.add_argument("--units")
        if m == "dry-run":
            p.add_argument("--report")
        else:
            p.add_argument("--out", required=True)
            p.add_argument("--fault-at", help=argparse.SUPPRESS)          # tests: the k-th transfer fails among its writes
    p = sub.add_parser("verify")
    p.add_argument("table")
    p.add_argument("--dir", required=True)
    p.add_argument("--default-source")
    p = sub.add_parser("promote")
    p.add_argument("--dir", required=True)
    p.add_argument("--to", required=True)
    p.add_argument("--name", default="WMS_PROD.ods")
    a = ap.parse_args(argv)
    if getattr(a, "default_source", None):
        ds = dr.source_of(a.default_source)
        if not ds:
            print(f"ОШИБКА: --default-source «{a.default_source}» — не тип источника (Поставщик, Офис, Производство, Детали, Старый склад, "
                  "Иной приход)", file=sys.stderr)
            return 2
        a.default_source = ds
    try:
        return {"dry-run": cmd_dry_run, "copy": cmd_copy, "verify": cmd_verify, "promote": cmd_promote}[a.mode](a)
    except (ValueError, OSError) as e:
        print(f"ОШИБКА: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
