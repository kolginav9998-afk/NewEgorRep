"""Upgrade of a WMS book of release 0.6 (schema WMS-SYS-3) to release 0.7 (WMS-SYS-4) with all its data, its journal and
its registration (M6 PRIME; П-40).

    python3 tools/upgrade.py BOOK [--check]

BOOK — the working WMS book of 0.6, closed (WMS not running: no WMS_Journal/wms.lock next to it).
 1. Checks — nothing is changed on a stop: the schema WMS-SYS-3; no unfinished operation in the book (the marker of _SYS,
    the before-image); the journal holds no operation beyond LAST_SEQ of the book (a tail: open the book in 0.6 and press
    «Восстановить» first); no lock.
 2. A backup of the book: WMS_Backups/<name>_<date>_<time>_preupgrade.ods (the kind WMS rotates), its SHA-256 in the report.
 3. The upgrade on a copy: the Basic modules of 0.7; the sheets «Приход авто» and _CAR; NEXT_CAR in _SYS and the schema
    WMS-SYS-4; the lookups of _IDX; the autofilter of the vehicles; the column AC and the separators of «Заказы» — the
    rows of «Ещё поступление» made by 0.6 get their mark (they never start a block; they stay where they are); the ways
    on «Главная», its version and help; «Справка» of 0.7; the frozen rows; every protection as in a new book of 0.7.
 4. The copy replaces the book; WMS 0.7 opens it where it is: the state must be «работа разрешена» and the self-check must
    have no errors — otherwise the book is put back from the backup and the copies WMS 0.7 made while opening it are removed
    (the whole upgrade or nothing).
--check: step 1 only.
Report: <folder of the book>/UPGRADE_REPORT_<date-time>.md. Exit code: 0 upgraded, 1 a stop by the rules (nothing changed),
2 a tool error (the book as it was).
"""
import argparse
import datetime
import glob
import hashlib
import os
import re
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import build_ods as B  # noqa: E402

OLD_SCHEMA = "WMS-SYS-3"
JOURNAL_DIR = "WMS_Journal"
BACKUP_DIR = "WMS_Backups"
LOCK_FILE = "wms.lock"
SELF_CHECK_OK = "САМОПРОВЕРКА WMS: ошибок 0"
IX_FIRST_CAR = 22                   # WmsConfig.IX_CO_MATCH: the first lookup row of the vehicles in _IDX


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def journal_max_seq(jdir, instance):
    """the largest seq of the journal lines of this WMS (J1;seq;time;instance;type;…;END;length)"""
    mx = 0
    for f in sorted(glob.glob(os.path.join(jdir, "WMS_journal_*.csv"))):
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


def checks(o, book):
    """the reasons to stop (empty: the upgrade may go on) and the facts of the book"""
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
        if facts["schema"] == "WMS-SYS-4":
            stops.append("книга уже версии 0.7 (схема WMS-SYS-4) — обновлять нечего")
        elif facts["schema"] != OLD_SCHEMA:
            stops.append(f"схема книги «{facts['schema']}»: обновляется только книга версии 0.6 (схема {OLD_SCHEMA})")
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


def unprotected(sh):
    if sh.isProtected():
        sh.unprotect(B.PWD)


def upgrade_doc(doc, test_book):
    """the changes of 0.7 in an open book of 0.6 (macros disabled); returns what was done (lines of the report)"""
    done = []
    doc.unprotect(B.PWD)
    sheets = doc.Sheets
    # LibreOffice refuses to change a cell style while any sheet is protected: every protection is lifted for the upgrade
    # and put back at the end (the new sheets are protected as in a new book)
    was_protected = []
    for i in range(sheets.getCount()):
        sh = sheets.getByIndex(i)
        if sh.isProtected():
            was_protected.append(sh.Name)
            sh.unprotect(B.PWD)
    # 1. the Basic modules of 0.7 (the whole library is replaced)
    doc.BasicLibraries.loadLibrary("Standard")
    lib = doc.BasicLibraries.getByName("Standard")
    old = list(lib.getElementNames())
    for m in old:
        lib.removeByName(m)
    mods = B.MODULES + (B.TEST_MODULES if test_book else [])
    for m in mods:
        folder = B.TEST_SRC if m in B.TEST_MODULES else B.SRC
        B.add_basic(doc, m, B.read_source(os.path.join(folder, m + ".bas")))
    done.append(f"модули Basic: {len(old)} модулей 0.6 заменены {len(mods)} модулями 0.7")
    # 2. the sheets of the vehicles
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    sheets.insertNewByName(B.CARS, names.index(B.SPECIAL) + 1)
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    sheets.insertNewByName(B.CAR, names.index(B.ADJ) + 1)
    B.build_cars(doc, sheets.getByName(B.CARS))
    B.build_car_service(doc, sheets.getByName(B.CAR))
    sheets.getByName(B.CAR).IsVisible = False
    done.append(f"листы «{B.CARS}» и {B.CAR} (скрытый) созданы")
    # 3. _SYS: the counter of the visits and the schema
    sy = sheets.getByName("_SYS")
    unprotected(sy)
    k = B.SYS_KEYS.index("NEXT_CAR")
    if sy.getCellByPosition(0, k).getString() != "":
        raise RuntimeError(f"_SYS: строка {k + 1} занята («{sy.getCellByPosition(0, k).getString()}»)")
    sy.getCellByPosition(0, k).setString("NEXT_CAR")
    sy.getCellByPosition(1, k).setValue(1)
    sy.getCellByPosition(1, B.SYS_KEYS.index("SCHEMA")).setString(B.SCHEMA)
    done.append(f"_SYS: NEXT_CAR = 1, схема {OLD_SCHEMA} → {B.SCHEMA}")
    # 4. _IDX: the lookups of the vehicles
    ix = sheets.getByName(B.IDX)
    unprotected(ix)
    for i, (label, f) in enumerate(B.idx_car_formulas(), start=IX_FIRST_CAR):
        if ix.getCellByPosition(0, i).getString() not in ("", label):
            raise RuntimeError(f"_IDX: строка {i + 1} занята («{ix.getCellByPosition(0, i).getString()}»)")
        ix.getCellByPosition(0, i).setString(label)
        ix.getCellByPosition(1, i).setFormula(f)
    done.append("_IDX: формулы поиска визитов (строки 23–27)")
    # 5. the autofilter of the table of the vehicles
    names = [sheets.getByIndex(i).Name for i in range(sheets.getCount())]
    B.autofilter(doc, names.index(B.CARS), "WMS_CARS", len(B.CAR_HEADERS), start_row=B.CAR_HEAD_ROW)
    # 6. «Заказы»: the column AC, the separators; the delivery rows made by 0.6 get their mark
    osh = sheets.getByName(B.ORDERS)
    unprotected(osh)
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
    # 7. «Главная»: the ways, the version, the help; «Справка» of 0.7
    msh = sheets.getByName(B.MAIN)
    unprotected(msh)
    B.add_nav(doc, msh)
    msh.getCellByPosition(1, 0).setString(B.version_text())
    msh.getCellByPosition(0, B.MAIN_HELP_ROW).setString(B.main_help_text())
    msh.getRows().getByIndex(B.MAIN_HELP_ROW).Height = 7000
    hsh = sheets.getByName(B.HELP)
    unprotected(hsh)
    hsh.getCellRangeByPosition(0, 0, 3, 5000).clearContents(1023)
    B.build_help(hsh)
    done.append("«Главная»: переходы на рабочие листы, версия и подсказка 0.7; «Справка» 0.7")
    # 8. protections as in a new book
    for name in was_protected + [B.CARS, B.CAR]:
        sh = sheets.getByName(name)
        if not sh.isProtected():
            sh.protect(B.PWD)
    doc.protect(B.PWD)
    done.append(f"защита: листов защищено {len(was_protected) + 2} (как было и новые), структура книги защищена")
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


def verify(o, book):
    """WMS 0.7 opens the upgraded book where it is: (ok, state line, self-check)"""
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
        # 2. the backup (the kind of WMS: rotated with the other copies before an upgrade)
        bdir = os.path.join(folder, BACKUP_DIR)
        os.makedirs(bdir, exist_ok=True)
        base = os.path.splitext(os.path.basename(book))[0]
        backup = os.path.join(bdir, f"{base}_{stamp}_preupgrade.ods")
        shutil.copy2(book, backup)
        h0 = sha256(book)
        if sha256(backup) != h0:
            raise RuntimeError("копия книги не совпадает с книгой")
        rep += ["## 2. Резервная копия", "", f"- `{backup}`", f"- SHA-256: `{h0}`", ""]
        # 3. the upgrade on a copy next to the book
        work = os.path.join(folder, f".{base}.upgrade.ods")
        shutil.copy2(book, work)
        doc = o.load(work, macros=0, hidden=True)
        test_book = facts.get("mode") == "TEST"
        done = upgrade_doc(doc, test_book)
        doc.store()
        doc.close(True)
        allow_insert_rows_all(work)
        # the save stamp WMS expects and the views (a document with a view saves them), as tools/build_ods.py does
        doc = o.load(work, macros=0, hidden=False)
        ec = doc.getDocumentProperties().EditingCycles
        doc.Sheets.getByName("_SYS").getCellByPosition(1, B.SK_SAVE_STAMP).setValue(ec + 1)
        B.apply_views(doc)
        doc.store()
        doc.close(True)
        rep += ["## 3. Изменения", ""] + [f"- {x}" for x in done] + [""]
        # 4. the copy becomes the book; WMS 0.7 must open it clean
        copies0 = set(os.listdir(bdir))
        os.replace(work, book)
        ok, state, sc = verify(o, book)
        rep += ["## 4. Проверка WMS 0.7", "", f"- запуск: {state[:400]}", "- самопроверка:", "", "```", sc.strip(), "```", ""]
        if not ok:
            shutil.copy2(backup, book)
            # a copy WMS 0.7 made of the upgraded book while opening it (the daily one) would restore a book that failed
            gone = sorted(set(os.listdir(bdir)) - copies0)
            for f in gone:
                os.remove(os.path.join(bdir, f))
            return finish(2, ["ОШИБКА: WMS 0.7 не открыла обновлённую книгу в рабочем состоянии — книга возвращена из резервной копии",
                              f"(SHA-256 {sha256(book)}{' = исходной' if sha256(book) == h0 else ' ≠ исходной!'})"]
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
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("book")
    ap.add_argument("--check", action="store_true", help="только проверки, книга не меняется")
    a = ap.parse_args(argv)
    return run(a.book, a.check)


if __name__ == "__main__":
    sys.exit(main())
