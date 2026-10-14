"""Upgrade of a working WMS book to this release with all its data and its journal.

PRIME 0.8 — a book of 0.7.0 … 0.7.2 (schema WMS-SYS-4): «Обновление из текущей WMS» (D-091). The book of 0.7 is only read;
the book of 0.8 is made in a new empty folder; nothing becomes the working WMS until PROMOTE; the rollback is the old book.

    python3 tools/upgrade.py check BOOK                          # the checks only, nothing is written anywhere
    python3 tools/upgrade.py build BOOK --to NEW_FOLDER [--rehearsal]
    python3 tools/upgrade.py promote --to NEW_FOLDER             # «СДЕЛАТЬ РАБОЧЕЙ»
    python3 tools/upgrade.py status --to NEW_FOLDER

CHECK    the book closed (no WMS_Journal/wms.lock, no lock file of LibreOffice), schema WMS-SYS-4, no unfinished operation
         (the marker, the before-image), the journal holds no operation beyond LAST_SEQ (a tail), no sheet protected with a
         password other than the WMS one, enough free space; SHA-256 of the book and of every journal file.
BUILD    after a clean CHECK, into NEW_FOLDER (it must not exist or be empty): a byte copy of the book and of the journal
         (without the lock; SHA-256 equal to the originals); on the copy, with macros disabled, the changes of 0.8 — the Basic
         modules, _SYS (schema WMS-SYS-5, UPGRADE_FROM), the registry of the transfer _LEG filled from the journal, the sheet
         «Сверка переноса», the buttons of «Наличие», «Главная» and «Справка» of 0.8, every protection as before. Then VERIFY:
         every cell of every sheet of the old book against the new one (only the cells of the upgrade may differ: the schema,
         the save stamp, UPGRADE_FROM; «Главная» and «Справка» — pages without data — are those of 0.8; the pages «Дашборд»,
         «Поиск», «Карта склада», «Топология» are new), the protections and the visibility of the sheets, the journal byte for
         byte, _LEG against the journal; a working copy of the new book opened by WMS 0.8 («работа разрешена», self-check
         without errors, «остаток = сумма движений», the independent oracle replaying the journal — a discrepancy the old book
         already had is named, a new one stops); the old book and its journal are hashed again (they must be unchanged). Any
         failure: the new book and its journal are removed, the reports stay. Rehearsal (--rehearsal): the same, never promoted.
PROMOTE  the new book unchanged since BUILD, the old book unchanged since BUILD (no operation was posted in it after the
         copy) and closed: WMS 0.8 opens the new book, it becomes the working file in its folder, self-check, saved, opened
         again. The old book is never touched — it is the rollback.
Reports in NEW_FOLDER: UPGRADE_REPORT.md (everything), UPGRADE_SUMMARY.txt (without the data of the warehouse: counts,
hashes, PASS/FAIL — it may be sent to the developer), UPGRADE_STATE.json (and the same as «key=value» lines in
UPGRADE_STATE.txt for the wizard WMS_UPGRADE.ods), progress.txt (the step being done, written first).

0.6 (WMS-SYS-3): the book is upgraded in place, as release 0.7 did it, all or nothing (the old form of the call):
    python3 tools/upgrade.py BOOK [--check]
Exit code: 0 done, 1 a stop by the rules (nothing changed), 2 a tool error (the old book as it was).
"""
import argparse
import datetime
import glob
import hashlib
import json
import os
import shutil
import sys
import tempfile
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
# the oracles: next to this file in a release folder (tools/), in the repository under tests/
sys.path.insert(1, os.path.join(os.path.dirname(HERE), "tests"))

import build_ods as B  # noqa: E402

SCHEMA_06 = "WMS-SYS-3"
SCHEMA_07 = "WMS-SYS-4"
OLD_SCHEMA = SCHEMA_06                 # the in-place upgrade (release 0.7 called it so)
JOURNAL_DIR = "WMS_Journal"
BACKUP_DIR = "WMS_Backups"
LOCK_FILE = "wms.lock"
SELF_CHECK_OK = "САМОПРОВЕРКА WMS: ошибок 0"
IX_FIRST_CAR = 22                   # WmsConfig.IX_CO_MATCH: the first lookup row of the vehicles in _IDX
STATE_NAME = "UPGRADE_STATE.json"
STATE_TEXT = "UPGRADE_STATE.txt"
REPORT_NAME = "UPGRADE_REPORT.md"
SUMMARY_NAME = "UPGRADE_SUMMARY.txt"
PROGRESS_NAME = "progress.txt"
CHUNK = 20000


def test_fault(point):
    """a fault of the tests (WMS_UPGRADE_TEST_FAULT=<point>); inert otherwise"""
    if os.environ.get("WMS_UPGRADE_TEST_FAULT") == point:
        raise RuntimeError(f"тестовый сбой в точке «{point}»")


class StopRun(Exception):
    """a stop by the rules: nothing of the working WMS was changed"""


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def journal_files(jdir):
    return sorted(glob.glob(os.path.join(jdir, "WMS_journal_*.csv")))


def journal_hashes(jdir):
    return {os.path.basename(f): sha256(f) for f in journal_files(jdir)}


def journal_size(jdir):
    return sum(os.path.getsize(f) for f in journal_files(jdir))


def journal_max_seq(jdir, instance):
    """the largest seq of the journal lines of this WMS (J1;seq;time;instance;type;…;END;length)"""
    mx = 0
    for f in journal_files(jdir):
        for raw in open(f, "rb").read().split(b"\n"):
            ln = raw.rstrip(b"\r").decode("utf-8", "replace")
            a = ln.split(";")
            if len(a) > 4 and a[0] == "J1" and a[1].isdigit() and a[3] == instance and ";END;" in ln:
                mx = max(mx, int(a[1]))
    return mx


def sys_values(doc):
    sh = doc.Sheets.getByName("_SYS")
    d = sh.getCellRangeByPosition(0, 0, 1, 40).getDataArray()
    return {str(k): v for k, v in d if k != ""}


def protection_keys(path):
    """{sheet: (protected, key, algorithm)} and the structure protection, read from content.xml (nothing is opened)"""
    import re
    with zipfile.ZipFile(path) as z:
        xml = z.read("content.xml").decode("utf-8")
    out = {}
    for m in re.finditer(r"<table:table ([^>]*)>", xml):
        attrs = dict(re.findall(r'([\w:-]+)="([^"]*)"', m.group(1)))
        name = attrs.get("table:name")
        if name is None:
            continue
        out[name] = (attrs.get("table:protected") == "true", attrs.get("table:protection-key", ""),
                     attrs.get("table:protection-key-digest-algorithm", ""))
    return out


def wms_keys():
    """the protection keys LibreOffice writes for the WMS password (SHA-1 / SHA-256 of UTF-16LE, base64)"""
    import base64
    raw = B.PWD.encode("utf-16-le")
    return {base64.b64encode(hashlib.sha1(raw).digest()).decode(), base64.b64encode(hashlib.sha256(raw).digest()).decode()}


def foreign_protections(path):
    keys = wms_keys()
    return sorted(n for n, (prot, key, _alg) in protection_keys(path).items() if prot and key and key not in keys)


def lo_lock(book):
    d, n = os.path.split(book)
    return os.path.join(d, f".~lock.{n}#")


# ================================================================ the changes of a release in an open book (macros disabled)

def unprotected(sh):
    if sh.isProtected():
        sh.unprotect(B.PWD)


def lift_protections(doc):
    """LibreOffice refuses to change a cell style while any sheet is protected: every protection is lifted for the upgrade
    and put back at the end (the new sheets are protected as in a new book); returns the protected sheets"""
    doc.unprotect(B.PWD)
    sheets = doc.Sheets
    was = []
    for i in range(sheets.getCount()):
        sh = sheets.getByIndex(i)
        if sh.isProtected():
            was.append(sh.Name)
            sh.unprotect(B.PWD)
    return was


def replace_modules(doc, test_book, done):
    doc.BasicLibraries.loadLibrary("Standard")
    lib = doc.BasicLibraries.getByName("Standard")
    old = list(lib.getElementNames())
    for m in old:
        lib.removeByName(m)
    mods = B.MODULES + (B.TEST_MODULES if test_book else [])
    for m in mods:
        folder = B.TEST_SRC if m in B.TEST_MODULES else B.SRC
        B.add_basic(doc, m, B.read_source(os.path.join(folder, m + ".bas")))
    done.append(f"модули Basic: {len(old)} прежних модулей заменены {len(mods)} модулями {B.core_versions()['WMS_PRODUCT_VERSION']}")


def step_07(doc, done):
    """0.6 → 0.7: «Приход авто» and _CAR, NEXT_CAR, the lookups of the vehicles, the separators of «Заказы»"""
    sheets = doc.Sheets
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    sheets.insertNewByName(B.CARS, names.index(B.SPECIAL) + 1)
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    sheets.insertNewByName(B.CAR, names.index(B.ADJ) + 1)
    B.build_cars(doc, sheets.getByName(B.CARS))
    B.build_car_service(doc, sheets.getByName(B.CAR))
    sheets.getByName(B.CAR).IsVisible = False
    done.append(f"листы «{B.CARS}» и {B.CAR} (скрытый) созданы")
    sy = sheets.getByName("_SYS")
    k = B.SYS_KEYS.index("NEXT_CAR")
    if sy.getCellByPosition(0, k).getString() != "":
        raise RuntimeError(f"_SYS: строка {k + 1} занята («{sy.getCellByPosition(0, k).getString()}»)")
    sy.getCellByPosition(0, k).setString("NEXT_CAR")
    sy.getCellByPosition(1, k).setValue(1)
    done.append(f"_SYS: NEXT_CAR = 1 (схема {SCHEMA_06} → {SCHEMA_07})")
    ix = sheets.getByName(B.IDX)
    for i, (label, f) in enumerate(B.idx_car_formulas(), start=IX_FIRST_CAR):
        if ix.getCellByPosition(0, i).getString() not in ("", label):
            raise RuntimeError(f"_IDX: строка {i + 1} занята («{ix.getCellByPosition(0, i).getString()}»)")
        ix.getCellByPosition(0, i).setString(label)
        ix.getCellByPosition(1, i).setFormula(f)
    done.append("_IDX: формулы поиска визитов (строки 23–27)")
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    B.autofilter(doc, names.index(B.CARS), "WMS_CARS", len(B.CAR_HEADERS), start_row=B.CAR_HEAD_ROW)
    osh = sheets.getByName(B.ORDERS)
    B.order_separators(doc, osh)
    rcv = sheets.getByName(B.RCV)
    cur = rcv.createCursor()
    cur.gotoEndOfUsedArea(False)
    rl = cur.getRangeAddress().EndRow
    rows = rcv.getCellRangeByPosition(0, 0, 5, max(rl, 1)).getDataArray()
    cur = osh.createCursor()
    cur.gotoEndOfUsedArea(False)
    ol = cur.getRangeAddress().EndRow
    v = osh.getCellRangeByPosition(21, 0, 21, max(ol, 1)).getDataArray()
    where = {str(x[0]): r for r, x in enumerate(v) if r > 0 and str(x[0]).startswith("ЕИ-")}
    marked = 0
    for r in rows[1:]:
        if r[5] == "ADD" and str(r[0]) in where and isinstance(r[1], float):
            osh.getCellByPosition(B.ORDER_BLOCK_COL, where[str(r[0])]).setString(f"OL{int(r[1])}")
            marked += 1
    done.append(f"«Заказы»: служебная колонка AC и разделители заказов; строк «Ещё поступление» 0.6 отмечено: {marked}")
    return [B.CARS, B.CAR]


def legacy_rows(jdir, instance, last_seq):
    """the rows of _LEG from the journal: every LEGACY_RECEIPT / LEGACY_RECEIPT_ADD / MIGRATE of this WMS applied in the book
    (seq ≤ LAST_SEQ, not abandoned) → {EI number: row A..K}"""
    import journal_oracle
    entries, _ = journal_oracle.read_journal(jdir)
    applied = [e for e in entries if e["inst"] == instance and e["seq"] <= last_seq]
    abandoned = journal_oracle.abandoned_seqs(applied)
    out = {}
    for e in applied:
        if e["seq"] in abandoned or e["type"] not in ("LEGACY_RECEIPT", "LEGACY_RECEIPT_ADD", "MIGRATE"):
            continue
        f = e["fields"]
        ei = f.get("EI", "")
        if len(ei) != 11 or not ei[3:].isdigit():
            raise RuntimeError(f"журнал seq {e['seq']} {e['type']}: ЕИ «{ei}» не распознан")
        n = int(ei[3:])
        if n in out:
            raise RuntimeError(f"журнал: {ei} перенесён дважды (seq {out[n][8]:.0f} и {e['seq']})")
        q = float(f["QTY"])
        if e["type"] == "MIGRATE":
            b, kind, x, src = q, "MIGRATE", "", "MIGRATE"
        else:
            b = float(f["BAL_AFTER"])
            kind = "RECEIPT" if e["type"] == "LEGACY_RECEIPT" else "ADD"
            x, src = f.get("LEGACY_STOCK", ""), f.get("STOCK_SRC", "")
        out[n] = (ei, kind, q, b, round(q - b, 6), x, src, f.get("ORIGIN", ""), float(e["seq"]), 0.0, "")
    return out


def step_08(doc, leg, from_text, done):
    """0.7 → 0.8: _SYS (UPGRADE_FROM, the schema), the registry of the transfer _LEG with its rows, «Сверка переноса», the buttons
    of «Наличие»; returns the new sheets"""
    sheets = doc.Sheets
    sy = sheets.getByName("_SYS")
    k = B.SYS_KEYS.index("UPGRADE_FROM")
    if sy.getCellByPosition(0, k).getString() not in ("",) or sy.getCellByPosition(1, k).getString() != "":
        raise RuntimeError(f"_SYS: строка {k + 1} занята («{sy.getCellByPosition(0, k).getString()}»)")
    sy.getCellByPosition(0, k).setString("UPGRADE_FROM")
    if from_text:
        sy.getCellByPosition(1, k).setString(from_text)
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    for need in (B.LEG, B.LEGREP, B.PACK):
        if need in names:
            raise RuntimeError(f"в книге уже есть лист «{need}»")
    sheets.insertNewByName(B.LEGREP, names.index(B.STOCK) + 1)
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    sheets.insertNewByName(B.LEG, names.index(B.CAR) + 1)
    leg_sh = sheets.getByName(B.LEG)
    B.build_leg(doc, leg_sh)
    B.build_legrep(doc, sheets.getByName(B.LEGREP))
    leg_sh.IsVisible = False
    for n, row in sorted(leg.items()):
        leg_sh.getCellRangeByPosition(0, n, len(row) - 1, n).setDataArray((tuple(x for x in row),))
        # empty texts must stay empty cells (setDataArray writes "" as an empty string)
        for c, x in enumerate(row):
            if x == "":
                leg_sh.getCellByPosition(c, n).clearContents(1 + 2 + 4 + 16)
    done.append(f"реестр переноса {B.LEG}: {len(leg)} ЕИ из журнала (приходы переноса и перенос ЕИ без заказа)")
    # the packing of the EIs («в 1 шт — 4 кан»): empty, filled by the operation EI_PACK of WMS 0.8
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    sheets.insertNewByName(B.PACK, names.index(B.LEG) + 1)
    B.build_pack(doc, sheets.getByName(B.PACK))
    sheets.getByName(B.PACK).IsVisible = False
    done.append(f"лист фасовки {B.PACK} (пустой: фасовку ЕИ задают кнопкой «Фасовка ЕИ»)")
    B.stock_buttons(doc, sheets.getByName(B.STOCK))
    done.append(f"лист «{B.LEGREP}», кнопки «Исправить ЕИ переноса» и «Сверка переноса» на «{B.STOCK}»")
    # the view sheets: «Дашборд» first, «Поиск» and «Карта склада» after «Наличие», «Топология» before «Справка»
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    for need in B.VIEW_SHEETS:
        if need in names:
            raise RuntimeError(f"в книге уже есть лист «{need}»")
    sheets.insertNewByName(B.DASH, 0)
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    sheets.insertNewByName(B.FIND, names.index(B.STOCK) + 1)
    sheets.insertNewByName(B.MAP, names.index(B.STOCK) + 2)
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    sheets.insertNewByName(B.TOPO, names.index(B.HELP))
    B.build_views(doc, sheets)
    B.view_events(sheets)
    B.status_colors(doc, sheets)
    B.where_buttons(doc, sheets)
    B.rebuild_main(doc, sheets.getByName(B.MAIN))
    done.append(f"страницы «{B.DASH}», «{B.FIND}», «{B.MAP}» и лист «{B.TOPO}» (вставьте туда схему склада); «{B.MAIN}» этого выпуска (кнопки к страницам)")
    return [B.LEGREP, B.LEG, B.PACK, B.DASH, B.FIND, B.MAP]


def finish_doc(doc, was_protected, new_sheets, done):
    """«Главная» and «Справка» of this release, the events of the book, every protection as in a new book"""
    sheets = doc.Sheets
    sheets.getByName("_SYS").getCellByPosition(1, B.SYS_KEYS.index("SCHEMA")).setString(B.SCHEMA)
    msh = sheets.getByName(B.MAIN)
    msh.getCellByPosition(1, 0).setString(B.version_text())
    msh.getCellByPosition(0, B.MAIN_HELP_ROW).setString(B.main_help_text())
    hsh = sheets.getByName(B.HELP)
    hsh.getCellRangeByPosition(0, 0, 3, 5000).clearContents(1023)
    B.build_help(hsh)
    done.append(f"«Главная»: версия и подсказка {B.core_versions()['WMS_PRODUCT_VERSION']}; «Справка» {B.core_versions()['WMS_PRODUCT_VERSION']}")
    # the events of the book are bound again (the same handlers; insurance against a changed name)
    B.bind_event(doc.Events, "OnLoad", "WmsCore.OnDocLoad")
    B.bind_event(doc.Events, "OnPrepareUnload", "WmsCore.OnDocPrepareUnload")
    B.bind_event(doc.Events, "OnSave", "WmsCore.OnDocSave")
    B.bind_event(doc.Events, "OnSaveDone", "WmsCore.OnDocSaveDone")
    B.bind_event(doc.Events, "OnSaveFailed", "WmsCore.OnDocSaveDone")
    for name, macro in ((B.ISSUES, "WmsIssue.OnIssuesChange"), (B.ORDERS, "WmsOrders.OnOrdersChange"), (B.RETURNS, "WmsReturn.OnReturnsChange"),
                        (B.SPECIAL, "WmsSpecial.OnSpecialChange"), (B.ADJUST, "WmsAdjust.OnAdjustChange"), (B.RCPT, "WmsIssue.OnRecipientsChange")):
        B.bind_event(sheets.getByName(name).Events, "OnChange", macro)
    for name in was_protected + new_sheets:
        sh = sheets.getByName(name)
        if not sh.isProtected():
            sh.protect(B.PWD)
    doc.protect(B.PWD)
    done.append(f"защита: листов защищено {len(was_protected) + len(new_sheets)} (как было и новые), структура книги защищена")


def upgrade_doc(doc, test_book, schema, leg, from_text):
    """the changes up to this release in an open book of schema `schema` (macros disabled); lines of the report"""
    done = []
    was = lift_protections(doc)
    replace_modules(doc, test_book, done)
    new = []
    if schema == SCHEMA_06:
        new += step_07(doc, done)
    new += step_08(doc, leg, from_text, done)
    finish_doc(doc, was, new, done)
    return done


def allow_insert_rows_all(path):
    """the protection option «insert rows» of the five sheets of rows (lost by a protect() through the API)"""
    def fix(xml):
        for name in (B.ISSUES, B.ORDERS, B.RETURNS, B.SPECIAL, B.ADJUST):
            head = f'<table:table table:name="{name}"'
            i = xml.index(head)
            j = xml.index(">", i) + 1
            if xml.startswith('<loext:table-protection', j) and 'loext:insert-rows="true"' not in xml[j:xml.index("/>", j)]:
                xml = B.allow_insert_rows(xml, name)
        return xml
    B.patch_content(path, fix)
    with zipfile.ZipFile(path) as z:
        n = z.read("content.xml").decode("utf-8").count('loext:insert-rows="true"')
    if n != 5:
        raise RuntimeError(f"параметр защиты «вставка строк» есть у {n} листов из 5")


def stamp_and_views(o, path):
    """the save stamp WMS expects and the views (a document with a view saves them), as tools/build_ods.py does"""
    doc = o.load(path, macros=0, hidden=False)
    ec = doc.getDocumentProperties().EditingCycles
    doc.Sheets.getByName("_SYS").getCellByPosition(1, B.SK_SAVE_STAMP).setValue(ec + 1)
    B.apply_views(doc)
    doc.store()
    doc.close(True)


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
    try:
        doc.close(True)
    except Exception:
        pass


def copy_journal(src_dir, dst_dir):
    """the journal files without the lock (the lock belongs to a running WMS, never to a copy)"""
    os.makedirs(dst_dir, exist_ok=True)
    if os.path.isdir(src_dir):
        for f in sorted(os.listdir(src_dir)):
            if f != LOCK_FILE and os.path.isfile(os.path.join(src_dir, f)):
                shutil.copy2(os.path.join(src_dir, f), os.path.join(dst_dir, f))


# ================================================================ 0.7 → 0.8: CHECK

def check_source(o, book):
    """(stops, warnings, facts) of a book of 0.7; nothing is written (the book is opened read-only, macros disabled)"""
    stops, warns, facts = [], [], {}
    if not os.path.isfile(book):
        return [f"нет файла {book}"], warns, facts
    folder = os.path.dirname(book)
    jdir = os.path.join(folder, JOURNAL_DIR)
    if os.path.exists(os.path.join(jdir, LOCK_FILE)):
        stops.append(f"WMS открыта или закрыта аварийно (есть {JOURNAL_DIR}/{LOCK_FILE}) — закройте книгу; после сбоя откройте её в своей "
                     "версии, дождитесь «работа разрешена» и закройте")
    if os.path.exists(lo_lock(book)):
        stops.append(f"книга открыта в LibreOffice (есть файл блокировки {os.path.basename(lo_lock(book))}) — закройте её")
    try:
        foreign = foreign_protections(book)
    except (zipfile.BadZipFile, KeyError) as e:
        return stops + [f"файл не похож на книгу ODS ({e})"], warns, facts
    if foreign:
        stops.append("листы " + ", ".join(f"«{n}»" for n in foreign) + " защищены не паролем WMS — снимите эту защиту в своей версии WMS "
                     "(«Сервис → Защитить лист», пароль того, кто её ставил) и закройте книгу; иначе обновлённая WMS не сможет проводить по ним")
    facts["book_hash"] = sha256(book)
    facts["journal_hashes"] = journal_hashes(jdir)
    facts["journal_size"] = journal_size(jdir)
    facts["book_size"] = os.path.getsize(book)
    doc = o.load(book, macros=0, hidden=True, ReadOnly=True)
    if doc is None:
        return stops + ["книга не открылась"], warns, facts
    try:
        if not doc.Sheets.hasByName("_SYS"):
            return stops + ["это не книга WMS (нет служебного листа _SYS)"], warns, facts
        sv = sys_values(doc)
        facts.update(schema=sv.get("SCHEMA"), instance=sv.get("INSTANCE_ID"), mode=sv.get("MODE"), last_seq=int(sv.get("LAST_SEQ") or 0),
                     tx=sv.get("TX_STATE"), bi=sv.get("TX_BEFORE_IMAGE"), core=sv.get("CORE_VERSION"), next_ei=int(sv.get("NEXT_EI") or 0))
        if facts["schema"] == B.SCHEMA:
            stops.append(f"книга уже версии {B.core_versions()['WMS_PRODUCT_VERSION']} (схема {B.SCHEMA}) — обновлять нечего")
        elif facts["schema"] == SCHEMA_06:
            stops.append(f"книга версии 0.6 (схема {SCHEMA_06}) — она обновляется на месте: upgrade.py КНИГА")
        elif facts["schema"] != SCHEMA_07:
            stops.append(f"схема книги «{facts['schema']}»: обновляется книга версии 0.7 (схема {SCHEMA_07})")
        if facts["tx"] not in ("NONE", "COMMITTED") or facts["bi"]:
            stops.append(f"в книге незавершённая операция (маркер {facts['tx']}) — откройте книгу в своей версии WMS: запуск её откатит")
        jmax = journal_max_seq(jdir, facts["instance"] or "")
        facts["journal_max"] = jmax
        if jmax > facts["last_seq"]:
            stops.append(f"в журнале {jmax - facts['last_seq']} операций, которых нет в книге (seq {facts['last_seq'] + 1}..{jmax}) — откройте книгу "
                         "в своей версии WMS и нажмите «Восстановить»")
        import journal_oracle
        entries, jinfo = journal_oracle.read_journal(jdir)
        foreign = sum(1 for e in entries if e["inst"] != facts["instance"])
        if foreign:
            stops.append(f"в журнале {foreign} строк другой WMS (другой INSTANCE_ID) — в папке {JOURNAL_DIR} смешаны журналы разных книг; "
                         "оставьте в ней только журнал этой книги")
        if jinfo.get("damaged"):
            warns.append(f"повреждённых строк журнала: {jinfo['damaged']} (обычно оборванная последняя строка после сбоя; WMS их пропускает)")
        facts["journal_lines"] = len(entries)
        if facts["last_seq"] > 0 and jmax < facts["last_seq"]:
            stops.append(f"журнал отстаёт от книги (в книге seq {facts['last_seq']}, в журнале {jmax}) — нужна папка {JOURNAL_DIR} этой книги")
        counts = {}
        for i in range(doc.Sheets.getCount()):
            sh = doc.Sheets.getByIndex(i)
            cur = sh.createCursor()
            cur.gotoEndOfUsedArea(False)
            counts[sh.Name] = cur.getRangeAddress().EndRow
        facts["rows"] = counts
        if facts["mode"] == "TEST":
            warns.append("тестовая книга (режим TEST)")
    finally:
        doc.close(True)
    return stops, warns, facts


# ================================================================ 0.7 → 0.8: VERIFY (old book against the new one)

def sheet_cells(sh, last_row, last_col):
    """(values, formulas) of A1 .. last_col/last_row by chunks"""
    vals, forms = [], []
    for r0 in range(0, last_row + 1, CHUNK):
        r1 = min(last_row, r0 + CHUNK - 1)
        rng = sh.getCellRangeByPosition(0, r0, last_col, r1)
        vals.extend(rng.getDataArray())
        forms.extend(rng.getFormulaArray())
    return vals, forms


def used(sh):
    cur = sh.createCursor()
    cur.gotoEndOfUsedArea(False)
    a = cur.getRangeAddress()
    return a.EndRow, a.EndColumn


def allowed_change(sheet, r, c):
    """the cells the upgrade itself writes"""
    if sheet == "_SYS":
        # the schema, the save stamp of the new file (WMS checks it at its start) and the new key UPGRADE_FROM
        return (r, c) in ((B.SYS_KEYS.index("SCHEMA"), 1), (B.SK_SAVE_STAMP, 1)) or r == B.SYS_KEYS.index("UPGRADE_FROM")
    if sheet == B.MAIN:
        return (r, c) in ((0, 1), (B.MAIN_HELP_ROW, 0))
    return sheet == B.HELP


def compare_books(old, new, leg_expected):
    """every cell of every sheet of the old book against the new one; (problems, statistics)"""
    P, stat = [], dict(sheets=0, cells=0)
    old_names = [old.Sheets.getByIndex(i).Name for i in range(old.Sheets.getCount())]
    new_names = [new.Sheets.getByIndex(i).Name for i in range(new.Sheets.getCount())]
    missing = [n for n in old_names if n not in new_names]
    if missing:
        P.append("в новой книге нет листов: " + ", ".join(missing))
    extra = [n for n in new_names if n not in old_names]
    if sorted(extra) != sorted([B.LEG, B.LEGREP, B.PACK] + B.VIEW_SHEETS + [x for x in (B.CARS, B.CAR) if x in extra]):
        P.append("в новой книге лишние листы: " + ", ".join(extra))
    for name in old_names:
        if name not in new_names:
            continue
        so, sn = old.Sheets.getByName(name), new.Sheets.getByName(name)
        if so.isProtected() != sn.isProtected() and not (sn.isProtected() and name in (B.CARS, B.CAR)):
            P.append(f"«{name}»: защита {so.isProtected()} → {sn.isProtected()}")
        if so.IsVisible != sn.IsVisible:
            P.append(f"«{name}»: видимость {so.IsVisible} → {sn.IsVisible}")
        if name in (B.HELP, B.MAIN):
            # «Справка» and «Главная» of this release: pages without data (WMS writes the state of «Главная» at every start)
            continue
        ro, co = used(so)
        rn, cn = used(sn)
        lr, lc = max(ro, rn), max(co, cn)
        bad = 0
        # chunk by chunk: a book of hundreds of thousands of rows is never held in memory at once
        for r0 in range(0, lr + 1, CHUNK):
            r1 = min(lr, r0 + CHUNK - 1)
            ro_ = so.getCellRangeByPosition(0, r0, lc, r1)
            rn_ = sn.getCellRangeByPosition(0, r0, lc, r1)
            vo, fo = ro_.getDataArray(), ro_.getFormulaArray()
            vn, fn = rn_.getDataArray(), rn_.getFormulaArray()
            for i in range(r1 - r0 + 1):
                if fo[i] == fn[i] and vo[i] == vn[i]:
                    continue
                r = r0 + i
                for c in range(lc + 1):
                    # a formula is compared as its text (a volatile formula such as NOW() may show another value at every load)
                    if str(fo[i][c]).startswith("="):
                        differs = fo[i][c] != fn[i][c]
                    else:
                        differs = vo[i][c] != vn[i][c] or fo[i][c] != fn[i][c]
                    if differs and not allowed_change(name, r, c):
                        bad += 1
                        if bad <= 3:
                            P.append(f"«{name}» {col_name(c)}{r + 1}: было «{fo[i][c]}», стало «{fn[i][c]}»")
        if bad > 3:
            P.append(f"«{name}»: всего отличающихся ячеек {bad}")
        stat["sheets"] += 1
        stat["cells"] += (lr + 1) * (lc + 1)
    # the registry of the transfer against the journal
    if new.Sheets.hasByName(B.LEG):
        sh = new.Sheets.getByName(B.LEG)
        lr, _ = used(sh)
        got = {}
        if lr >= 1:
            for i, row in enumerate(sheet_cells(sh, lr, len(B.LEG_HEADERS) - 1)[0][1:], start=1):
                if row[0] != "":
                    got[i] = tuple(row)
        want = {n: tuple(x for x in row) for n, row in leg_expected.items()}
        if got != want:
            diff = sorted(set(got) ^ set(want)) or [n for n in want if got.get(n) != want[n]]
            P.append(f"{B.LEG}: {len(got)} строк, по журналу {len(want)}; расходятся ЕИ {diff[:5]}")
        stat["leg"] = len(got)
    return P, stat


def col_name(c):
    s = ""
    c += 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def oracle_problems(doc, jdir, initial):
    """the independent replay of the journal against the book (tests/adjust_oracle.py of this release)"""
    import adjust_oracle
    P, info, _ = adjust_oracle.check(doc, jdir, initial, 0, legacy_mn_empty=False)
    return P, info


# ================================================================ 0.7 → 0.8: the steps

class State:
    def __init__(self, folder):
        self.path = os.path.join(folder, STATE_NAME)
        self.d = json.load(open(self.path, encoding="utf-8")) if os.path.isfile(self.path) else {}

    def save(self):
        tmp = self.path + ".tmp"
        json.dump(self.d, open(tmp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        os.replace(tmp, self.path)
        # the same, flat «key=value» lines: the wizard WMS_UPGRADE.ods reads them (Basic has no JSON)
        b, pr = self.d.get("build", {}), self.d.get("promote", {})
        flat = {"kind": self.d.get("kind", ""), "source": self.d.get("source", ""), "started": self.d.get("started", ""),
                "build_ok": "1" if b.get("ok") else "0", "build_time": b.get("time", ""), "build_result": b.get("result", ""),
                "rehearsal": "1" if b.get("rehearsal") or self.d.get("kind") == "rehearsal" else "0",
                "promote_ok": "1" if pr.get("ok") else "0", "promote_time": pr.get("time", ""), "promote_book": pr.get("book", ""),
                "promote_report": pr.get("report", "")}
        tmp = os.path.join(os.path.dirname(self.path), STATE_TEXT + ".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            f.write("".join(f"{k}={str(v).replace(chr(10), ' ')}\n" for k, v in flat.items()))
        os.replace(tmp, os.path.join(os.path.dirname(self.path), STATE_TEXT))


def progress(folder, text):
    try:
        with open(os.path.join(folder, PROGRESS_NAME), "w", encoding="utf-8") as f:
            f.write(f"{now()} {text}\n")
    except OSError:
        pass


def free_space(folder):
    st = os.statvfs(folder)
    return st.f_bavail * st.f_frsize


def write_reports(folder, rep, summary):
    with open(os.path.join(folder, REPORT_NAME), "w", encoding="utf-8") as f:
        f.write("\n".join(rep) + "\n")
    with open(os.path.join(folder, SUMMARY_NAME), "w", encoding="utf-8") as f:
        f.write("\n".join(summary) + "\n")


def release_problem():
    """"" when this folder holds what an upgrade puts into a book (the Basic modules src/basic, the guides docs/ of «Справка»);
    otherwise what is missing — checked before anything is read or copied"""
    miss = [f"src/basic/{m}.bas" for m in B.MODULES if not os.path.isfile(os.path.join(B.SRC, m + ".bas"))]
    miss += [f"docs/{d}" for d in B.HELP_DOCS if not os.path.isfile(os.path.join(B.ROOT, "docs", d))]
    if miss:
        return (f"папка выпуска {B.ROOT} неполная — нет {', '.join(miss[:4])}{' …' if len(miss) > 4 else ''}. Распакуйте архив выпуска "
                "целиком и запускайте обновление из него")
    return ""


def cmd_check(book, out=print):
    from wmslo import Office
    book = os.path.abspath(book)
    why = release_problem()
    if why:
        out(f"СТОП — обновлять нельзя, пока не исправлено:\n- {why}")
        return 1
    prof = tempfile.mkdtemp(prefix="wms_upgrade_")
    o = Office(f"upgchk{os.getpid()}", prof)
    try:
        stops, warns, facts = check_source(o, book)
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    lines = [f"Книга: {book}", f"схема {facts.get('schema')}, ядро {facts.get('core')}, режим {facts.get('mode')}, операций {facts.get('last_seq')}, "
             f"журнал до seq {facts.get('journal_max')}, SHA-256 {facts.get('book_hash', '')[:16]}…"]
    lines += [f"ПРЕДУПРЕЖДЕНИЕ: {w}" for w in warns]
    if stops:
        out("\n".join(lines + ["СТОП — обновлять нельзя, пока не исправлено:"] + [f"- {s}" for s in stops]))
        return 1
    out("\n".join(lines + ["Проверки пройдены: книгу можно обновить — дальше «СОБРАТЬ И СВЕРИТЬ» (build КНИГА --to НОВАЯ_ПАПКА). Книга не изменена."]))
    return 0


def cmd_build(book, to, rehearsal=False, out=print):
    from wmslo import Office
    book = os.path.abspath(book)
    to = os.path.abspath(to)
    if os.path.abspath(os.path.dirname(book)) == to:
        out("СТОП: новая папка совпадает с папкой книги — обновление делается только в новую пустую папку")
        return 1
    if os.path.isdir(to) and os.listdir(to):
        out(f"СТОП: папка {to} не пуста — обновление делается только в новую пустую папку")
        return 1
    why = release_problem()
    if why:
        out(f"СТОП: {why}")
        return 1
    os.makedirs(to, exist_ok=True)
    t0 = time.time()
    rep = [f"# Обновление WMS до версии {B.core_versions()['WMS_PRODUCT_VERSION']}" + (" — РЕПЕТИЦИЯ" if rehearsal else ""), "",
           f"**Старая книга (только чтение):** `{book}`  ", f"**Новая папка:** `{to}`  ", f"**Начато:** {now()}", ""]
    summary = [f"WMS UPGRADE {B.core_versions()['WMS_PRODUCT_VERSION']}{' REHEARSAL' if rehearsal else ''}", f"started {now()}"]
    state = State(to)
    state.d = dict(kind="rehearsal" if rehearsal else "upgrade", source=book, started=now(), build={})
    state.save()
    target = os.path.join(to, os.path.basename(book))
    jdst = os.path.join(to, JOURNAL_DIR)
    prof = tempfile.mkdtemp(prefix="wms_upgrade_")
    o = Office(f"upgbuild{os.getpid()}", prof)
    code = 2
    try:
        progress(to, "1/5 проверки старой книги")
        stops, warns, facts = check_source(o, book)
        rep += ["## 1. Проверки старой книги", "", f"- схема {facts.get('schema')}, ядро {facts.get('core')}, режим {facts.get('mode')}, экземпляр "
                f"{facts.get('instance')}", f"- операций в книге {facts.get('last_seq')}, в журнале до seq {facts.get('journal_max')}",
                f"- SHA-256 книги `{facts.get('book_hash')}`; файлов журнала {len(facts.get('journal_hashes', {}))}"] + [f"- предупреждение: {w}" for w in warns] + [""]
        summary += [f"source schema={facts.get('schema')} core={facts.get('core')} mode={facts.get('mode')} last_seq={facts.get('last_seq')} "
                    f"journal_max={facts.get('journal_max')} journal_files={len(facts.get('journal_hashes', {}))}",
                    f"source sha256={facts.get('book_hash')}"]
        need = 3 * facts.get("book_size", 0) + 2 * facts.get("journal_size", 0) + (50 << 20)
        if not stops and free_space(to) < need:
            stops.append(f"мало места на диске: свободно {free_space(to) >> 20} МБ, нужно не меньше {need >> 20} МБ")
        if stops:
            rep += ["## СТОП", ""] + [f"- {s}" for s in stops]
            summary.append("RESULT=STOP " + " | ".join(stops))
            code = 1
            raise StopRun("; ".join(stops))
        # 2. the copies (byte for byte)
        progress(to, "2/5 копия книги и журнала")
        shutil.copy2(book, target)
        copy_journal(os.path.join(os.path.dirname(book), JOURNAL_DIR), jdst)
        if sha256(target) != facts["book_hash"]:
            raise RuntimeError("копия книги не совпадает с книгой")
        if journal_hashes(jdst) != facts["journal_hashes"]:
            raise RuntimeError("копия журнала не совпадает с журналом")
        test_fault("after_copy")
        rep += ["## 2. Копия", "", f"- `{target}` — SHA-256 совпадает с исходной книгой", f"- `{jdst}` — {len(facts['journal_hashes'])} файлов, "
                "SHA-256 каждого совпадает", ""]
        # 3. the changes of 0.8 on the copy
        progress(to, "3/5 изменения версии 0.8 в копии")
        leg = legacy_rows(jdst, facts["instance"], facts["last_seq"])
        from_text = "|".join([os.path.dirname(book), str(facts["last_seq"]), facts["book_hash"], now(), str(facts["journal_size"])])
        doc = o.load(target, macros=0, hidden=True)
        done = upgrade_doc(doc, facts.get("mode") == "TEST", facts["schema"], leg, from_text)
        doc.store()
        doc.close(True)
        allow_insert_rows_all(target)
        stamp_and_views(o, target)
        rep += ["## 3. Изменения", ""] + [f"- {x}" for x in done] + [""]
        test_fault("after_changes")
        # 4. verification: the old book against the new one, then WMS 0.8 on a working copy of the new one
        progress(to, "4/5 сверка: каждая ячейка старой и новой книги")
        oldd = o.load(book, macros=0, hidden=True, ReadOnly=True)
        newd = o.load(target, macros=0, hidden=True, ReadOnly=True)
        try:
            P, stat = compare_books(oldd, newd, leg)
            P0, inf0 = oracle_problems(oldd, os.path.join(os.path.dirname(book), JOURNAL_DIR), {}) if facts.get("mode") != "TEST" else ([], {})
        finally:
            oldd.close(True)
            newd.close(True)
        rep += ["## 4. Сверка «было / стало»", "", f"- листов сверено {stat['sheets']}, ячеек {stat['cells']}; реестр переноса: {stat.get('leg', 0)} ЕИ",
                f"- расхождений: {len(P)}"] + [f"  - {x}" for x in P[:20]] + [""]
        summary.append(f"compare sheets={stat['sheets']} cells={stat['cells']} legacy_eis={stat.get('leg', 0)} problems={len(P)}")
        if P:
            raise RuntimeError("сверка старой и новой книги: " + "; ".join(P[:3]))
        test_fault("before_wms")
        progress(to, "5/5 сверка: WMS 0.8 открывает новую книгу (рабочая копия проверки)")
        vdir = os.path.join(to, ".verify")
        os.makedirs(vdir)
        vbook = os.path.join(vdir, os.path.basename(book))
        shutil.copy2(target, vbook)
        copy_journal(jdst, os.path.join(vdir, JOURNAL_DIR))
        doc, stl = open_wms(o, vbook)
        if doc is None or "STATE=CLEAN" not in stl:
            raise RuntimeError(f"WMS 0.8 не открыла новую книгу в рабочем состоянии: {stl[:300]}")
        try:
            sc = o.basic(doc, "WmsCore", "ActionSelfCheck") or ""
            bc = o.basic(doc, "WmsLegacy", "BalanceCheck") or ""
            lr = o.basic(doc, "WmsLegacy", "LegacyReport") or ""
            P1, inf1 = oracle_problems(doc, os.path.join(vdir, JOURNAL_DIR), {}) if facts.get("mode") != "TEST" else ([], {})
        finally:
            close_doc(o, doc, save=False)
        shutil.rmtree(vdir, ignore_errors=True)
        new_oracle = [x for x in P1 if x not in P0]
        rep += ["## 5. WMS 0.8 на новой книге", "", f"- запуск: работа разрешена", f"- самопроверка: {sc.splitlines()[0] if sc else '—'}",
                f"- «остаток = сумма движений»: {'все ЕИ сходятся' if not bc else bc}", f"- сверка переноса: {lr}",
                f"- независимый оракул журнала: в старой книге расхождений {len(P0)}, в новой {len(P1)}, новых {len(new_oracle)} "
                f"(операций {inf1.get('ops', '—')})"] + [f"  - уже было в старой книге: {x}" for x in P0[:10]] + [""]
        summary += [f"wms08 selfcheck={'OK' if sc.startswith(SELF_CHECK_OK) else 'FAIL'} balance={'OK' if not bc else 'DIFF'} "
                    f"oracle_old={len(P0)} oracle_new={len(P1)} oracle_added={len(new_oracle)}", f"legacy_report={lr[:120]}"]
        if not sc.startswith(SELF_CHECK_OK):
            raise RuntimeError("самопроверка WMS 0.8 на новой книге: " + sc.splitlines()[0])
        if new_oracle:
            raise RuntimeError("оракул журнала нашёл в новой книге расхождения, которых нет в старой: " + "; ".join(new_oracle[:3]))
        # the old book and its journal were only read
        if sha256(book) != facts["book_hash"] or journal_hashes(os.path.join(os.path.dirname(book), JOURNAL_DIR)) != facts["journal_hashes"]:
            raise RuntimeError("старая книга или её журнал изменились во время обновления — в ней работали? Обновление не принято")
        rh = sha256(target)
        state.d["build"] = dict(ok=True, time=now(), source_hash=facts["book_hash"], source_journal=facts["journal_hashes"], result=target,
                                result_hash=rh, last_seq=facts["last_seq"], instance=facts["instance"], rehearsal=rehearsal)
        state.save()
        secs = round(time.time() - t0)
        rep += ["## Итог", "", f"**{'РЕПЕТИЦИЯ ПРОШЛА' if rehearsal else 'НОВАЯ КНИГА ГОТОВА'}** ({now()}, {secs} с). Старая книга не изменена "
                f"(SHA-256 то же). Новая книга: `{target}` (SHA-256 `{rh}`)."]
        if not rehearsal:
            rep.append("Следующий шаг — «СДЕЛАТЬ РАБОЧЕЙ» (promote) в конце рабочего дня. До этого работайте в старой книге как обычно; "
                       "если после этой сборки в старой книге будут проведены операции, «СДЕЛАТЬ РАБОЧЕЙ» попросит собрать заново.")
        summary += [f"result sha256={rh}", f"RESULT={'REHEARSAL_PASS' if rehearsal else 'BUILD_PASS'} seconds={secs}"]
        code = 0
        out(f"{'Репетиция прошла' if rehearsal else 'Новая книга готова'}: {target}")
    except StopRun as e:
        out(f"СТОП: {e}")
        code = 1
    except Exception as e:
        rep += ["## ОШИБКА", "", f"- {e}", "", "Новая книга и её журнал удалены; старая книга не менялась."]
        summary.append(f"RESULT=FAIL {str(e)[:300]}")
        out(f"ОШИБКА: {e}")
        code = 2
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
        if code != 0:
            for p in (target, jdst, os.path.join(to, ".verify")):
                if os.path.isdir(p):
                    shutil.rmtree(p, ignore_errors=True)
                elif os.path.exists(p):
                    os.remove(p)
            state.d.setdefault("build", {})["ok"] = False
            state.d["build"]["time"] = now()
            state.save()
        progress(to, "готово" if code == 0 else "остановлено")
        write_reports(to, rep, summary)
    return code


def cmd_promote(to, out=print):
    from wmslo import Office
    to = os.path.abspath(to)
    state = State(to)
    b = state.d.get("build", {})
    if state.d.get("promote", {}).get("ok"):
        out(f"СТОП: новая книга уже рабочая с {state.d['promote']['time']}")
        return 1
    if not b.get("ok"):
        out("СТОП: нет успешной сборки новой книги в этой папке — сначала «3. СОБРАТЬ И СВЕРИТЬ» (build)")
        return 1
    if b.get("rehearsal"):
        out("СТОП: это репетиция — её книга рабочей не делается. Для перехода соберите новую книгу в другой папке без --rehearsal")
        return 1
    src, target = state.d["source"], b["result"]
    if not os.path.isfile(target) or sha256(target) != b["result_hash"]:
        out("СТОП: новая книга изменилась после сборки — соберите её заново")
        return 1
    if os.path.exists(os.path.join(os.path.dirname(src), JOURNAL_DIR, LOCK_FILE)) or os.path.exists(lo_lock(src)):
        out("СТОП: старая WMS сейчас открыта — закройте её и повторите")
        return 1
    if not os.path.isfile(src) or sha256(src) != b["source_hash"] or journal_hashes(os.path.join(os.path.dirname(src), JOURNAL_DIR)) != b["source_journal"]:
        out("СТОП: в старой книге работали после сборки (книга или журнал изменились) — соберите новую книгу заново, чтобы ничего не потерять")
        return 1
    prof = tempfile.mkdtemp(prefix="wms_upgrade_")
    o = Office(f"upgprom{os.getpid()}", prof)
    try:
        doc, stl = open_wms(o, target)
        if doc is None or "STATE=CLEAN" not in stl:
            out(f"ОШИБКА: новая книга не открылась в рабочем состоянии: {stl[:300]}")
            return 2
        sc = o.basic(doc, "WmsCore", "ActionSelfCheck") or ""
        lr = o.basic(doc, "WmsLegacy", "LegacyReport") or ""
        if not sc.startswith(SELF_CHECK_OK):
            close_doc(o, doc, save=False)
            out("ОШИБКА: самопроверка новой книги: " + sc.splitlines()[0])
            return 2
        close_doc(o, doc, save=True)
        doc, stl = open_wms(o, target)
        ok = doc is not None and "STATE=CLEAN" in stl
        if doc is not None:
            close_doc(o, doc, save=False)
        if not ok:
            out(f"ОШИБКА: новая книга после сохранения: {stl[:300]}")
            return 2
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    # the tools of this release next to the new working book (the toolbox reads the book of its folder)
    tb_src, tb_dst = os.path.join(os.path.dirname(HERE), "WMS_TOOLBOX"), os.path.join(to, "WMS_TOOLBOX")
    toolbox = ""
    if os.path.isdir(tb_src) and not os.path.exists(tb_dst):
        try:
            shutil.copytree(tb_src, tb_dst)
            toolbox = tb_dst
        except OSError as e:
            toolbox = f"не скопированы ({e}) — скопируйте папку WMS_TOOLBOX выпуска рядом с книгой вручную"
    state.d["promote"] = dict(ok=True, time=now(), book=target, report=lr, toolbox=toolbox)
    state.save()
    with open(os.path.join(to, "ПЕРЕХОД_ВЫПОЛНЕН.txt"), "w", encoding="utf-8") as f:
        f.write(f"Переход выполнен {now()}.\n\nРабочая WMS теперь: {target}\nОткрывайте только её.\n\n"
                f"Старая книга {src} не изменена — это откат. Для работы её больше не открывайте: операции в ней не попадут в новую WMS\n"
                "(новая WMS при запуске предупредит, если в старой книге что-то провели).\n\n"
                f"Сверка переноса в новой книге: {lr}\n" + (f"\nИнструменты выпуска (WMS_TOOLBOX): {toolbox}\n" if toolbox else ""))
    with open(os.path.join(to, REPORT_NAME), "a", encoding="utf-8") as f:
        f.write(f"\n## СДЕЛАТЬ РАБОЧЕЙ\n\n- {now()}: новая книга зарегистрирована рабочей в своей папке, самопроверка без ошибок, сохранена и "
                f"открыта снова.\n- Сверка переноса: {lr}\n")
    with open(os.path.join(to, SUMMARY_NAME), "a", encoding="utf-8") as f:
        f.write(f"PROMOTE=OK {now()}\n")
    out(f"Готово: рабочая WMS — {target}. Старая книга не изменена.")
    return 0


def cmd_status(to, out=print):
    st = State(os.path.abspath(to)).d
    out(json.dumps(st, ensure_ascii=False, indent=1) if st else "нет состояния обновления в этой папке")
    return 0


# ================================================================ 0.6: in place (the form of release 0.7)

def checks(o, book):
    """the reasons to stop (empty: the upgrade may go on) and the facts of a book of 0.6"""
    stops, facts = [], {}
    jdir = os.path.join(os.path.dirname(book), JOURNAL_DIR)
    if os.path.exists(os.path.join(jdir, LOCK_FILE)):
        stops.append(f"книга открыта в WMS или закрыта аварийно (есть {JOURNAL_DIR}/{LOCK_FILE}) — закройте WMS; после сбоя откройте книгу в "
                     "версии 0.6, дождитесь «работа разрешена» (или снимите блокировку) и закройте её")
    doc = o.load(book, macros=0, hidden=True)
    if doc is None:
        return ["книга не открылась"], facts
    try:
        if not doc.Sheets.hasByName("_SYS"):
            return ["это не книга WMS (нет служебного листа _SYS)"], facts
        sv = sys_values(doc)
        facts = dict(schema=sv.get("SCHEMA"), instance=sv.get("INSTANCE_ID"), mode=sv.get("MODE"), last_seq=int(sv.get("LAST_SEQ") or 0),
                     tx=sv.get("TX_STATE"), bi=sv.get("TX_BEFORE_IMAGE"), reg=sv.get("REGISTERED_URL"))
        if facts["schema"] == B.SCHEMA:
            stops.append(f"книга уже версии {B.core_versions()['WMS_PRODUCT_VERSION']} (схема {B.SCHEMA}) — обновлять нечего")
        elif facts["schema"] == SCHEMA_07:
            stops.append(f"книга версии 0.7 (схема {SCHEMA_07}) обновляется в новую папку: upgrade.py build КНИГА --to НОВАЯ_ПАПКА "
                         "(или мастер WMS_UPGRADE.ods)")
        elif facts["schema"] != SCHEMA_06:
            stops.append(f"схема книги «{facts['schema']}»: на месте обновляется только книга версии 0.6 (схема {SCHEMA_06})")
        if facts["tx"] not in ("NONE", "COMMITTED") or facts["bi"]:
            stops.append(f"в книге незавершённая операция (маркер {facts['tx']}) — откройте книгу в версии 0.6: запуск её откатит")
        jmax = journal_max_seq(jdir, facts["instance"] or "")
        facts["journal_max"] = jmax
        if jmax > facts["last_seq"]:
            stops.append(f"в журнале {jmax - facts['last_seq']} операций, которых нет в книге (seq {facts['last_seq'] + 1}..{jmax}) — откройте книгу "
                         "в версии 0.6 и нажмите «Восстановить»")
        if doc.Sheets.hasByName(B.CARS) or doc.Sheets.hasByName(B.CAR):
            stops.append(f"в книге уже есть лист «{B.CARS}» или {B.CAR}")
    finally:
        doc.close(True)
    return stops, facts


def verify(o, book):
    """WMS of this release opens the upgraded book where it is: (ok, state line, self-check)"""
    doc = o.load(book, macros=4)
    if doc is None:
        return False, "книга не открылась", ""
    try:
        st = o.basic(doc, "WmsCore", "StateDump") or ""
        rep = o.basic(doc, "WmsCore", "StartupReport") or ""
        sc = o.basic(doc, "WmsCore", "ActionSelfCheck") or ""
        ok = "STATE=CLEAN" in st and sc.startswith(SELF_CHECK_OK)
        return ok, rep, sc
    finally:
        if doc.isModified():
            doc.store()
        try:
            o.dispatch(doc, ".uno:CloseDoc")
        except Exception:
            pass
        try:
            doc.close(True)
        except Exception:
            pass


def run(book, check_only=False, out=print):
    from wmslo import Office
    book = os.path.abspath(book)
    folder = os.path.dirname(book)
    stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    rep = [f"# Обновление книги WMS до версии {B.core_versions()['WMS_PRODUCT_VERSION']}", "", f"**Книга:** `{book}`  ", f"**Начато:** {now()}", ""]
    report_path = os.path.join(folder, f"UPGRADE_REPORT_{stamp}.md")

    def finish(code, lines):
        rep.extend(lines)
        rep.append("")
        rep.append(f"**Итог:** {['ОБНОВЛЕНО', 'СТОП — ничего не изменено', 'ОШИБКА — книга как была'][code]} ({now()})")
        open(report_path, "w", encoding="utf-8").write("\n".join(rep) + "\n")
        out("\n".join(lines))
        out(f"отчёт: {report_path}")
        return code

    if not os.path.isfile(book):
        return finish(2, [f"ОШИБКА: нет файла {book}"])
    prof = tempfile.mkdtemp(prefix="wms_upgrade_")
    o = Office(f"upgrade{os.getpid()}", prof)
    backup = None
    try:
        stops, facts = checks(o, book)
        rep += ["## 1. Проверки", "", f"- схема: {facts.get('schema')}, экземпляр: {facts.get('instance')}, режим: {facts.get('mode')}",
                f"- последняя операция книги: {facts.get('last_seq')}, журнала: {facts.get('journal_max')}", ""]
        if stops:
            return finish(1, ["## СТОП", ""] + [f"- {s}" for s in stops])
        if check_only:
            return finish(0, ["Проверки пройдены: книгу можно обновить (запуск без --check)."])
        bdir = os.path.join(folder, BACKUP_DIR)
        os.makedirs(bdir, exist_ok=True)
        base = os.path.splitext(os.path.basename(book))[0]
        backup = os.path.join(bdir, f"{base}_{stamp}_preupgrade.ods")
        shutil.copy2(book, backup)
        h0 = sha256(book)
        if sha256(backup) != h0:
            raise RuntimeError("копия книги не совпадает с книгой")
        rep += ["## 2. Резервная копия", "", f"- `{backup}`", f"- SHA-256: `{h0}`", ""]
        work = os.path.join(folder, f".{base}.upgrade.ods")
        shutil.copy2(book, work)
        leg = legacy_rows(os.path.join(folder, JOURNAL_DIR), facts["instance"], facts["last_seq"])
        doc = o.load(work, macros=0, hidden=True)
        done = upgrade_doc(doc, facts.get("mode") == "TEST", facts["schema"], leg, "")
        doc.store()
        doc.close(True)
        allow_insert_rows_all(work)
        stamp_and_views(o, work)
        rep += ["## 3. Изменения", ""] + [f"- {x}" for x in done] + [""]
        copies0 = set(os.listdir(bdir))
        os.replace(work, book)
        ok, state, sc = verify(o, book)
        rep += [f"## 4. Проверка WMS {B.core_versions()['WMS_PRODUCT_VERSION']}", "", f"- запуск: {state[:400]}", "- самопроверка:", "", "```", sc.strip(), "```", ""]
        if not ok:
            shutil.copy2(backup, book)
            gone = sorted(set(os.listdir(bdir)) - copies0)
            for f in gone:
                os.remove(os.path.join(bdir, f))
            return finish(2, [f"ОШИБКА: WMS {B.core_versions()['WMS_PRODUCT_VERSION']} не открыла обновлённую книгу в рабочем состоянии — книга "
                              "возвращена из резервной копии", f"(SHA-256 {sha256(book)}{' = исходной' if sha256(book) == h0 else ' ≠ исходной!'})"]
                          + ([f"удалены копии обновлённой книги: {', '.join(gone)}"] if gone else []))
        return finish(0, [f"Книга обновлена: {book}", f"резервная копия 0.6: {backup}"])
    except Exception as e:
        lines = [f"ОШИБКА: {e}"]
        if backup and os.path.exists(backup):
            try:
                if sha256(book) != sha256(backup):
                    shutil.copy2(backup, book)
                lines.append("книга возвращена из резервной копии")
            except OSError as e2:
                lines.append(f"вернуть книгу не удалось ({e2}) — скопируйте {backup} на место книги вручную")
        return finish(2, lines)
    finally:
        for f in glob.glob(os.path.join(folder, ".*.upgrade.ods")):
            try:
                os.remove(f)
            except OSError:
                pass
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] in ("check", "build", "promote", "status"):
        ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
        ap.add_argument("cmd")
        ap.add_argument("book", nargs="?")
        ap.add_argument("--to")
        ap.add_argument("--rehearsal", action="store_true")
        a = ap.parse_args(argv)
        if a.cmd == "check":
            return cmd_check(a.book) if a.book else 2
        if a.cmd == "build":
            if not a.book or not a.to:
                print("нужно: build КНИГА --to НОВАЯ_ПАПКА")
                return 2
            return cmd_build(a.book, a.to, a.rehearsal)
        if not a.to:
            print(f"нужно: {a.cmd} --to ПАПКА")
            return 2
        return cmd_promote(a.to) if a.cmd == "promote" else cmd_status(a.to)
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("book")
    ap.add_argument("--check", action="store_true", help="только проверки, книга не меняется")
    a = ap.parse_args(argv)
    return run(a.book, a.check)


if __name__ == "__main__":
    sys.exit(main())
