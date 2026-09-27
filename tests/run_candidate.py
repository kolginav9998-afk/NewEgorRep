"""WMS_PROD_CANDIDATE (FINAL WMS MARATHON §5) — automated checks of the release candidate built by tools/build_release.py:
the folder and its manifest, the production book (MODE PROD, first start registers it, the version on «Главная», the
guides on «Справка», the buttons bound to their macros), «Состояние системы», «Проверка перед работой», the check of the
LibreOffice version (a test seam of the test book fakes other versions) and the self-check; the candidate book takes a
transfer (tools/migrate.py dry-run), and the release folder alone — away from the repository — transfers a table into its
own book with its own tools (dry-run → copy → verify with the oracle → promote).

    WMS_TEST_OUT=/tmp/wms_cand python3 tests/run_candidate.py [case ...]

Results: WMS_TEST_OUT/results_candidate.json and WMS_TEST_OUT/TEST_REPORT_candidate.md.
"""
import csv
import hashlib
import os
import shutil
import subprocess
import sys
import time

import uno  # noqa: F401  (LibreOffice Python-UNO)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(ROOT, "tools"))
from harness import OUT, Session, Results, new_wms  # noqa: E402
import build_release  # noqa: E402

R = Results()
CASES = []
REL = os.path.join(OUT, "release")
BOOK = "WMS_PROD_CANDIDATE.ods"
STATUS_ROW = 24                                 # WmsStatus.STATUS_ROW


def case(fn):
    CASES.append(fn)
    return fn


def sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


def release():
    if not os.path.isdir(REL):
        r = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "build_release.py"), REL], capture_output=True, text=True, timeout=900)
        if r.returncode != 0:
            raise RuntimeError(r.stdout + r.stderr)
    return REL


def fresh_candidate(name):
    """a copy of the release folder (the book never opened) for one case"""
    d = os.path.join(OUT, "cases", name)
    shutil.rmtree(d, ignore_errors=True)
    shutil.copytree(release(), d)
    return os.path.join(d, BOOK)


def section(s):
    sh = s.doc.Sheets.getByName("Главная")
    rows = sh.getCellRangeByPosition(0, STATUS_ROW, 1, STATUS_ROW + 30).getDataArray()
    return [(a, b) for a, b in rows if a or b]


@case
def c01_release_folder():
    c = "C01"
    rel = release()
    man = dict()
    files = []
    for row in csv.reader(open(os.path.join(rel, "RELEASE_MANIFEST.csv"), encoding="utf-8"), delimiter=";"):
        if row[0] == "file":
            files.append(row[1:])
        else:
            man[row[0]] = row[1]
    bad = [f[0] for f in files if not os.path.exists(os.path.join(rel, f[0])) or sha(os.path.join(rel, f[0])) != f[2]
           or os.path.getsize(os.path.join(rel, f[0])) != int(f[1])]
    cfg = open(os.path.join(ROOT, "src", "basic", "WmsConfig.bas"), encoding="utf-8").read()
    R.add(c, "папка выпуска: книга WMS_PROD_CANDIDATE.ods, docs (примечания к выпуску, резервные копии и восстановление, контракт), tools "
             "(перенос и его оракул), RELEASE_MANIFEST.csv — версии как в исходниках, размер и SHA-256 каждого файла совпадают",
          len(files) == 1 + len(build_release.DOCS) + len(build_release.TOOLS) + len(build_release.ORACLES) and not bad and man.get("product") and f'WMS_PRODUCT_VERSION = "{man.get("product")}"' in cfg
          and f'WMS_CORE_VERSION = "{man.get("core")}"' in cfg and man.get("schema") == "WMS-SYS-3", f"{man}; файлов {len(files)}; не совпали {bad}")
    rc = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "build_release.py"), rel], capture_output=True, text=True).returncode
    R.add(c, "повторная сборка в существующую папку — отказ (выпуск собирается заново)", rc == 2, str(rc))


@case
def c02_candidate_book():
    c = "C02"
    p = fresh_candidate("c02")
    s = Session(p)
    try:
        s.U("TestUiAuto", 1)
        st = s.state()
        sys_mode = s.sysv("MODE")
        main = s.doc.Sheets.getByName("Главная")
        ver = main.getCellByPosition(1, 0).getString()
        names = [s.doc.Sheets.getByIndex(i).Name for i in range(s.doc.Sheets.getCount()) if s.doc.Sheets.getByIndex(i).IsVisible]
        hlp = s.doc.Sheets.getByName("Справка")
        cur = hlp.createCursor()
        cur.gotoEndOfUsedArea(False)
        text = "\n".join(r[0] for r in hlp.getCellRangeByPosition(0, 0, 0, cur.getRangeAddress().EndRow).getDataArray())
        form = main.getDrawPage().getForms().getByIndex(0)
        btn = {}
        for i in range(form.getCount()):
            ev = form.getScriptEvents(i)
            btn[form.getByIndex(i).Label] = ev[0].ScriptCode.split("Standard.", 1)[1].split("?")[0] if ev else ""
        R.add(c, "книга-кандидат: режим PROD, первый запуск — «работа разрешена» (рабочий файл зарегистрирован), реестр пуст; на «Главной» — "
                 "версия продукта, ядра и схемы; листы по порядку, последний видимый — «Справка»",
              st.get("STATE") == "CLEAN" and st.get("REG") == "own" and sys_mode == "PROD" and "версия 0.6.0" in ver and "WMS-SYS-3" in ver
              and names == ["Главная", "Заказы", "Иной приход", "Выдачи", "Возврат", "Корректировки", "Наличие", "Получатели", "Справка"]
              and s.sysv("NEXT_EI") == 1.0, f"{st}; {sys_mode}; «{ver}»; {names}")
        R.add(c, "«Справка»: примечания к выпуску и руководство по резервным копиям и восстановлению (без разметки Markdown); лист защищён",
              "примечания к выпуску" in text and "резервные копии и восстановление" in text and "Восстановление из резервной копии" in text
              and "**" not in text and "`" not in text and hlp.isProtected(), text[:200])
        R.add(c, "кнопки «Состояние системы» и «Проверка перед работой» на «Главной» привязаны к макросам",
              btn.get("Состояние системы") == "WmsStatus.BtnSystemStatus" and btn.get("Проверка перед работой") == "WmsStatus.BtnPreWorkCheck", str(btn))
        # the test seam of the version check is inert in a production book
        seam = s.B("TestSetLoVersion", "6.4.7.2", module="WmsStatus")
        R.add(c, "тестовый шов проверки версии в рабочей книге не действует", seam.startswith("REFUSED"), seam)
    finally:
        s.close()


@case
def c03_status_and_prework():
    c = "C03"
    p = fresh_candidate("c03")
    s = Session(p)
    try:
        s.U("TestUiAuto", 1)
        s.doc.store()
        mod0 = s.doc.isModified()
        s.B("BtnSystemStatus", module="WmsStatus")
        sec = section(s)
        st = dict(sec[1:])
        R.add(c, "«Состояние системы»: раздел на «Главной» — версии, состояние, экземпляр, рабочий файл (зарегистрирован), проверенная версия "
                 "LibreOffice, журнал, резервные копии (ежедневная копия первого запуска), размер книги; книга не стала «изменённой»",
              sec[0][0] == "СОСТОЯНИЕ СИСТЕМЫ" and "0.6.0-final-core" in st.get("WMS", "") and st.get("состояние", "").startswith("✓")
              and "(зарегистрирован)" in st.get("рабочий файл", "") and "проверенная версия" in st.get("LibreOffice", "")
              and "1 шт." in st.get("резервные копии", "") and "строк движений 0" in st.get("размер книги", "") and s.doc.isModified() == mod0,
              f"{sec}")
        s.B("BtnPreWorkCheck", module="WmsStatus")
        msg = s.U("TestUiLastMessage")
        sec = section(s)
        st = dict(sec[1:])
        R.add(c, "«Проверка перед работой»: ошибок 0; окно называет итог; раздел — WMS, LibreOffice, файл ODS, макросы, запись журнала, "
                 "резервные копии, автосохранение, размер; ничего не записано в учёт",
              msg.startswith("ПРОВЕРКА ПЕРЕД РАБОТОЙ: ошибок 0") and sec[0][0] == "ПРОВЕРКА ПЕРЕД РАБОТОЙ"
              and all(k in st for k in ("WMS", "LibreOffice", "файл", "макросы", "журнал", "резервные копии", "автосохранение", "размер", "защиты"))
              and st["журнал"].startswith("✓") and s.sysv("LAST_SEQ") == 0.0 and not os.listdir(s.jdir) == [], f"{msg}; {sec}")
        prot0 = st.get("защиты", "")
        # no backup for 10 days, then none at all
        bdir = os.path.join(os.path.dirname(p), "WMS_Backups")
        old = time.time() - 10 * 86400
        for f in os.listdir(bdir):
            os.utime(os.path.join(bdir, f), (old, old))
        w1 = dict(section_after(s))["резервные копии"]
        for f in os.listdir(bdir):
            os.remove(os.path.join(bdir, f))
        w2 = dict(section_after(s))["резервные копии"]
        R.add(c, "«Проверка перед работой» предупреждает: последняя копия 10 дней назад; копий нет", w1.startswith("! последняя копия 10") and w2.startswith("! копий нет"),
              f"{w1}; {w2}")
        # the protections: a sheet unprotected, a service sheet shown, the structure unprotected — and put back
        iss = s.doc.Sheets.getByName("Выдачи")
        iss.unprotect("wms")
        s.doc.unprotect("wms")
        s.doc.Sheets.getByName("_SYS").IsVisible = True
        w3 = dict(section_after(s)).get("защиты", "")
        s.doc.Sheets.getByName("_SYS").IsVisible = False
        s.doc.protect("wms")
        iss.protect("wms")
        w4 = dict(section_after(s)).get("защиты", "")
        R.add(c, "«Проверка перед работой» — защиты: листы защищены (кроме «Получателей»), служебные скрыты, структура защищена; снятая защита "
                 "листа, открытый служебный лист, снятая защита структуры — предупреждение с названиями; после восстановления — без замечаний",
              prot0.startswith("✓") and w3.startswith("! ") and "лист «Выдачи» не защищён" in w3 and "служебный лист «_SYS» не скрыт" in w3
              and "структура книги не защищена" in w3 and w4.startswith("✓"), f"{prot0}; {w3}; {w4}")
    finally:
        s.close()


def section_after(s):
    s.B("BtnPreWorkCheck", module="WmsStatus")
    return section(s)[1:]


@case
def c04_libreoffice_versions():
    c = "C04"
    p = new_wms("c04")                              # a test book: the version seam works there
    s = Session(p)
    try:
        s.U("TestUiAuto", 1)
        res = {}
        for v in ("24.2.7.2", "26.2.0.3", "25.8.1.1", "7.6.4.1", "6.4.7.2", ""):
            s.B("TestSetLoVersion", v or "x", module="WmsStatus")
            pre = s.B("PreWorkCheck", module="WmsStatus")
            head = pre.splitlines()[0]
            line = [x for x in pre.splitlines() if "\tLibreOffice\t" in x][0]
            res[v or "x"] = (head, line.split("\t")[0])
        R.add(c, "проверка версии LibreOffice: 24.2 и 26.2 — без замечаний; 25.8 и 7.6 — предупреждение «не проверялась»; 6.4 и неизвестная — "
                 "ошибка проверки перед работой",
              res["24.2.7.2"][1] == "OK" and res["26.2.0.3"][1] == "OK" and res["25.8.1.1"][1] == "WARN" and res["7.6.4.1"][1] == "WARN"
              and res["6.4.7.2"][1] == "FAIL" and res["x"][1] == "FAIL" and "ошибок 1" in res["6.4.7.2"][0], str(res))
        s.B("TestSetLoVersion", "25.8.1.1", module="WmsStatus")
        rep = s.B("WmsStartup")
        s.B("TestSetLoVersion", "", module="WmsStatus")
        rep2 = s.B("WmsStartup")
        R.add(c, "запуск с непроверенной версией: заметка «ВНИМАНИЕ: LibreOffice … не проверялась», работа разрешена (не блокирует); с "
                 "проверенной — заметки нет",
              "ВНИМАНИЕ: LibreOffice 25.8.1.1" in rep and rep.startswith("СОСТОЯНИЕ: работа разрешена") and "ВНИМАНИЕ" not in rep2, f"{rep[:160]} / {rep2[:120]}")
    finally:
        s.close()


@case
def c05_self_check_and_transfer():
    c = "C05"
    p = fresh_candidate("c05")
    s = Session(p)
    try:
        sc = s.B("ActionSelfCheck")
    finally:
        s.close()
    R.add(c, "самопроверка книги-кандидата — ошибок 0", sc.startswith("САМОПРОВЕРКА WMS: ошибок 0"), sc.splitlines()[0])
    d = os.path.dirname(p)
    t = os.path.join(d, "t.csv")
    with open(t, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["ЕИ", "Наименование", "Единица", "Количество", "Место", "Тип источника"])
        w.writerow(["12", "Молоток", "шт", "2", "A-1", "Старый склад"])
    fresh = fresh_candidate("c05b")
    r = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "migrate.py"), "dry-run", t, "--book", fresh, "--accept-table-places"],
                       capture_output=True, text=True, timeout=600)
    R.add(c, "книга-кандидат принимает перенос: dry-run таблицы остатков — без замечаний (перенос — tests/run_migration.py)", r.returncode == 0,
          (r.stdout + r.stderr)[-200:])


@case
def c06_transfer_from_release_folder():
    c = "C06"
    # what the warehouse gets is the release folder alone: its own tools/ must transfer a table into its own book, with
    # nothing of the repository on the path (the oracle of VERIFY included)
    d = os.path.join(OUT, "cases", "c06")
    shutil.rmtree(d, ignore_errors=True)
    rel = os.path.join(d, "release")
    shutil.copytree(release(), rel)
    t = os.path.join(d, "остатки.csv")
    with open(t, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["ЕИ", "Наименование", "Артикул", "Единица", "Количество", "Место", "Категория", "Тип источника", "Старая маркировка"])
        w.writerow(["12", "Молоток", "", "шт", "2", "A-1", "Инструмент", "Старый склад", "инв. 4"])
        w.writerow(["ЕИ-00000030", "Шестерня Z-12", "GR-12", "шт", "5", "D-2", "Детали", "Детали", ""])
        w.writerow(["", "Кабель ВВГ 3х1,5", "", "м", "120,5", "A-1", "Кабель", "Поставщик", ""])
    env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
    tool = os.path.join(rel, "tools", "migrate.py")
    book = os.path.join(rel, BOOK)
    mig, wh = os.path.join(d, "mig"), os.path.join(d, "warehouse")
    steps = []
    for args in (["dry-run", t, "--book", book, "--accept-table-places"],
                 ["copy", t, "--book", book, "--out", mig, "--accept-table-places"],
                 ["verify", t, "--dir", mig],
                 ["promote", "--dir", mig, "--to", wh]):
        r = subprocess.run([sys.executable, tool] + args, capture_output=True, text=True, timeout=900, cwd=d, env=env)
        steps.append((args[0], r.returncode, (r.stdout + r.stderr).strip()[-200:]))
        if r.returncode:
            break
    vr = open(os.path.join(mig, "VERIFY_REPORT.md"), encoding="utf-8").read() if os.path.exists(os.path.join(mig, "VERIFY_REPORT.md")) else ""
    R.add(c, "папка выпуска сама по себе (без репозитория): её tools/migrate.py переносит таблицу в её книгу — dry-run, copy, verify "
             "(расхождений 0, оракул воспроизвёл журнал), promote — коды 0; рабочая книга WMS_PROD.ods",
          [x[1] for x in steps] == [0, 0, 0, 0] and "расхождений: **0**" in vr and "Оракул (журнал воспроизведён независимо): операций 3" in vr
          and os.path.exists(os.path.join(wh, "WMS_PROD.ods")), f"{[(x[0], x[1]) for x in steps]}; {[x[2] for x in steps if x[1]]}")


def main():
    only = set(a.lower() for a in sys.argv[1:])
    t0 = time.time()
    for fn in CASES:
        if only and not any(fn.__name__.startswith(o) for o in only):
            continue
        try:
            fn()
        except Exception as e:
            R.error(fn.__name__, e)
        subprocess.run(["pkill", "-9", "-f", "profile_.*" + str(os.getpid())], capture_output=True)
    R.save(os.path.join(OUT, "results_candidate.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_candidate.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
