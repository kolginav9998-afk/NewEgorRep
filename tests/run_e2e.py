"""The final end-to-end check of the release (FINAL WMS MARATHON §8; M6 PRIME §13–§15): one warehouse from the release
candidate to the tools, and a warehouse of the previous release upgraded by the release folder.

    WMS_TEST_OUT=/tmp/wms_e2e python3 tests/run_e2e.py

release (tools/build_release.py) → migration of a synthetic old stock by the tools of the release folder itself
(tools/migrate.py dry-run → copy → verify → promote; nothing of the repository on the path) → WMS_TOOLBOX of the release
installed next to the book → WMS_PROD.ods: order → receipt → special receipt → refill of a migrated part → move → issue → return → write-off →
inventory correction → vehicles («Приход авто»: two arrivals, one departure) → save → reopen (state, check before work,
self-check, oracles) → backup (WMS and the script to a medium, health check — the scripts of the installed WMS_TOOLBOX) →
snapshot for the tools → WMS_TOOLBOX: DOCTOR, SEARCH (the plate), ANALYTICS (transport), MANAGER («Контроль дня»), LABELS,
RECONCILE, INVENTORY → its batch loaded and posted by WMS → reopen, self-check, oracles, health check. Then E11: a working
book of release 0.6.0 (built by the sources of its commit, used for a day) upgraded to this release by tools/upgrade.py of
the release folder alone. Every step is a check; the book is checked by the independent oracles (the whole journal
replayed, MIGRATE included; the vehicles by tests/car_oracle.py). Synthetic data only.

Results: WMS_TEST_OUT/results_e2e.json and WMS_TEST_OUT/TEST_REPORT_e2e.md.
"""
import csv
import glob
import hashlib
import io
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
from harness import OUT, PROFILES, Session, Results, git_sources, unique  # noqa: E402
from wmslo import Office  # noqa: E402
import car_oracle  # noqa: E402

R = Results()
W = os.path.join(OUT, "warehouse")            # the folder of the warehouse: WMS_PROD.ods, WMS_Journal, WMS_Backups, WMS_Export, WMS_TOOLBOX
D1, D2, D3 = "21.09.2026", "22.09.2026", "23.09.2026"
OLD_COMMIT = "939929a"                        # release 0.6.0 (M5): the warehouse of E11 works on it before the upgrade


def ei(n):
    return f"ЕИ-{n:08d}"


def tool(script, *args, timeout=1800):
    r = subprocess.run([sys.executable, os.path.join(ROOT, "tools", script)] + [str(a) for a in args], capture_output=True, text=True, timeout=timeout)
    return r.returncode, r.stdout + r.stderr


def release_tool(rel, script, *args, timeout=1800):
    """a tool of the release folder as the warehouse runs it: from that folder, nothing of the repository on the path"""
    env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
    r = subprocess.run([sys.executable, os.path.join(rel, "tools", script)] + [str(a) for a in args], capture_output=True, text=True,
                       timeout=timeout, cwd=rel, env=env)
    return r.returncode, r.stdout + r.stderr


def old_stock_table(path):
    """a synthetic old warehouse: 36 physical batches — numbers up to 700 with gaps, three rows without a number, parts,
    one «Иной приход» to be identified, zero balances; the places of the list"""
    rows = []
    kinds = ["Старый склад", "Поставщик", "Офис", "Производство"]
    for i in range(30):
        n = 100 + i * 20
        rows.append([str(n), f"Позиция старого склада {i}", f"OLD-{i:03d}" if i % 4 == 0 else "", ["шт", "м", "кг", "упак"][i % 4],
                     "0" if i == 7 else f"{(i % 9) + 1}", f"S-{i % 6:02d}", ["Крепёж", "Кабель", "Химия"][i % 3], kinds[i % 4], f"инв. {i}" if i % 5 == 0 else "", ""])
    rows += [["710", "Шестерня Z-12", "GR-12", "шт", "8", "D-1", "Детали", "Детали", "", ""],
             ["711", "Шестерня Z-24", "GR-24", "шт", "5", "D-1", "Детали", "Детали", "", ""],
             ["720", "Коробка без маркировки", "", "шт", "2", "Z-1", "", "Иной приход", "", "найдено при переезде"],
             ["", "Стремянка", "", "шт", "1", "S-01", "Инструмент", "Старый склад", "", "без наклейки"],
             ["", "Удлинитель 10 м", "", "шт", "3", "S-02", "Электрика", "Офис", "", ""],
             ["", "Перчатки рабочие", "", "пар", "40", "S-03", "СИЗ", "Поставщик", "", ""]]
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["ЕИ", "Наименование", "Артикул", "Единица", "Количество", "Место", "Категория", "Тип источника", "Старая маркировка", "Комментарий"])
        w.writerows(rows)
    places = os.path.join(os.path.dirname(path), "места.txt")
    open(places, "w", encoding="utf-8").write("\n".join([f"S-{k:02d}" for k in range(6)] + ["D-1", "Z-1", "A-1", "B-7"]) + "\n")
    return rows, places


def step(c, name, ok, detail=""):
    R.add(c, name, bool(ok), detail)
    if not ok:
        raise RuntimeError(f"шаг не выполнен: {name}")


def main():
    t0 = time.time()
    shutil.rmtree(W, ignore_errors=True)
    os.makedirs(W)
    try:
        run()
    except Exception as e:
        R.error("e2e", e)
    subprocess.run(["pkill", "-9", "-f", "profile_.*" + str(os.getpid())], capture_output=True)
    R.save(os.path.join(OUT, "results_e2e.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_e2e.md"), "w", encoding="utf-8") as f:
        f.write("| Шаг | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


def run():
    # ------------------------------------------------------------ E01 release and E02 migration
    rel = os.path.join(OUT, "release")
    shutil.rmtree(rel, ignore_errors=True)
    rc, out = tool("build_release.py", rel)
    cand = os.path.join(rel, "WMS_PROD_CANDIDATE.ods")
    step("E01", "выпуск: папка кандидата (книга, WMS_TOOLBOX, документы, инструменты переноса и обновления с модулями, манифест)",
         rc == 0 and os.path.exists(cand) and os.path.exists(os.path.join(rel, "WMS_TOOLBOX", "WMS_TOOLBOX.ods"))
         and os.path.exists(os.path.join(rel, "docs", "OPERATOR_GUIDE.md")) and os.path.exists(os.path.join(rel, "tools", "upgrade.py"))
         and os.path.exists(os.path.join(rel, "src", "basic", "WmsCar.bas")), out[-160:])
    table = os.path.join(OUT, "остатки.csv")
    rows, places = old_stock_table(table)
    rc1, o1 = release_tool(rel, "migrate.py", "dry-run", table, "--book", cand, "--places", places, "--report", os.path.join(OUT, "dry.md"))
    copy = os.path.join(OUT, "migration")
    shutil.rmtree(copy, ignore_errors=True)
    rc2, o2 = release_tool(rel, "migrate.py", "copy", table, "--book", cand, "--places", places, "--out", copy)
    rc3, o3 = release_tool(rel, "migrate.py", "verify", table, "--dir", copy)
    shutil.rmtree(W, ignore_errors=True)
    rc4, o4 = release_tool(rel, "migrate.py", "promote", "--dir", copy, "--to", W)
    prod = os.path.join(W, "WMS_PROD.ods")
    # the installation (docs/PC_VALIDATION.md §7): WMS_TOOLBOX of the release next to the working book
    shutil.copytree(os.path.join(rel, "WMS_TOOLBOX"), os.path.join(W, "WMS_TOOLBOX"))
    step("E02", f"перенос старого склада ({len(rows)} строк) инструментами папки выпуска: dry-run 0, copy 0, verify 0 (расхождений нет), "
                "promote 0 → рабочая книга WMS_PROD.ods; WMS_TOOLBOX выпуска рядом с ней",
         (rc1, rc2, rc3, rc4) == (0, 0, 0, 0) and os.path.exists(prod), f"{rc1} {rc2} {rc3} {rc4}; {o3[-200:]}")
    # ------------------------------------------------------------ E03 the working day
    s = Session(prod)
    try:
        s.U("TestUiAuto", 1)
        st = s.state()
        top = s.sysv("NEXT_EI")
        s.doc.Sheets.getByName("Получатели").getCellRangeByPosition(0, 1, 1, 2).setDataArray((("ива", "Иванов Иван"), ("пет", "Петров Пётр")))
        s.order_input(1, A="З-1", B="Кабель КГ 3х2,5", H="100", I="м", L="Кабель-Сервис", F="100", N=D1, U="B-7", C="УПД-1", O=D1)
        r_rc = s.click_ord("BtnRcvPost", 1)
        e_rc = int(top)
        s.special_input(1, B="Офис", D="Стол офисный", F="1", G="шт", H=D1, I="A-1", K="Офис")
        r_sp = s.click_spc("BtnSpcPost", 1)
        s.special_input(2, B="Детали", E="gr-12", F="4", G="шт", H=D2)
        r_ref = s.click_spc("BtnSpcPost", 2)
        refill_ei = s.spc(2)[13]
        s.adjust_input(1, B="Перемещение", C="100", J="S-05", K=D2, L="перестановка стеллажей")
        r_mv = s.click_adj("BtnAdjPost", 1)
        s.issue_input(1, ei=str(e_rc), qty="30", date=D2, who="ива")
        r_is = s.post(1)
        s.return_input(1, issue="1", qty="5", date=D3)
        r_rt = s.click_ret("BtnRetPost", 1)
        s.adjust_input(2, B="Списание", C="120", G="1", K=D3, L="брак")
        r_wo = s.click_adj("BtnAdjPost", 2)
        s.adjust_input(3, B="Инвентаризация", C="140", H="9", K=D3, L="пересчёт стеллажа S-02")
        r_inv = s.click_adj("BtnAdjPost", 3)
        res = [r_rc, r_sp, r_ref, r_mv, r_is, r_rt, r_wo, r_inv]
        # the vehicles like the storekeeper: two cells of the panel and ПРИЕХАЛ; the departure of the first one (the engine: the
        # button asks a person when a vehicle leaves within two minutes of its arrival)
        cars = []
        for plate, sup in (("Газель А123ВС", "Кабель-Сервис"), ("КАМАЗ Х456ОР", "Метиз-Опт")):
            for col, v in (("C", plate), ("D", sup)):
                s.goto("Приход авто", f"${col}$2")
                s.o.dispatch(s.doc, ".uno:EnterString", StringName=v)
            s.B("BtnCarArrive", module="WmsCarUi")
            cars.append(s.U("TestUiLastMessage"))
        cars.append(s.B("CarDepart", 1, module="WmsCar"))
        cv = s.doc.Sheets.getByName("Приход авто").getCellRangeByPosition(0, 6, 7, 7).getDataArray()
        step("E03", "рабочий день в WMS_PROD: заказ и приход (новый ЕИ выше перенесённых), иной приход, пополнение перенесённой детали GR-12 в тот "
                    "же ЕИ-00000710, перемещение, выдача, возврат, списание, инвентаризация — всё проведено; «Приход авто»: две машины "
                    "приехали (визиты № 1, № 2), первая уехала",
             st.get("STATE") == "CLEAN" and all(str(x).startswith("OK") for x in res) and refill_ei == ei(710) and s.stock(710) == 12.0
             and s.stock(e_rc) == 75.0 and s.stock_row(100)[5] == "S-05" and all(str(x).startswith("OK") for x in cars)
             and [r_[0] for r_ in cv] == [1.0, 2.0] and [r_[7] for r_ in cv] == ["Уехал", "На территории"],
             f"{res}; пополнен {refill_ei}; {s.stock(710)}; {s.stock(e_rc)}; {cars}; {[r_[7] for r_ in cv]}")
        s.doc.store()
    finally:
        s.close()
    # ------------------------------------------------------------ E04 reopen
    s = Session(prod)
    try:
        s.U("TestUiAuto", 1)
        st = s.state()
        pre = s.B("PreWorkCheck", module="WmsStatus")
        sc = s.B("ActionSelfCheck")
        P, inf, _ = s.adjcheck(0, initial={})
        PC, infc = car_oracle.check(s.doc, s.jdir, 0)
        step("E04", "повторное открытие: «работа разрешена»; проверка перед работой и самопроверка — ошибок 0; оракулы (журнал с переносом "
                    "воспроизведён независимо; визиты машин) — без расхождений",
             st.get("STATE") == "CLEAN" and pre.startswith("ПРОВЕРКА ПЕРЕД РАБОТОЙ: ошибок 0") and sc.startswith("САМОПРОВЕРКА WMS: ошибок 0") and not P
             and not PC and infc.startswith("визитов 2: на территории 1, уехало 1"),
             f"{pre.splitlines()[0]}; {sc.splitlines()[0]}; {P[:3]}; операций {inf.get('ops')}; машины {infc}; {PC[:3]}")
        # ------------------------------------------------------------ E05 backup
        b = s.B("ActionBackupNow")
        # ------------------------------------------------------------ E06 snapshot
        x = s.EX("ExportSnapshot")
        snap = s.EX("TestLastSnapshot")
    finally:
        s.close()
    media = os.path.join(OUT, "носитель")
    shutil.rmtree(media, ignore_errors=True)
    tb_scripts = os.path.join(W, "WMS_TOOLBOX", "backup")
    rb = subprocess.run([sys.executable, os.path.join(tb_scripts, "wms_backup.py"), "--wms", W, "--to", media], capture_output=True, text=True)
    hc = subprocess.run([sys.executable, os.path.join(tb_scripts, "wms_healthcheck.py"), "--wms", W, "--book", "WMS_PROD.ods"], capture_output=True, text=True)
    copies = glob.glob(os.path.join(media, "WMS_BACKUP_*"))
    in_copy = sorted(os.path.relpath(p, copies[0]) for p in glob.glob(os.path.join(copies[0], "**", "*"), recursive=True)) if copies else []
    step("E05", "резервные копии: «Резервная копия» WMS, копия папки на носитель с SHA-256 (backup/wms_backup.py: книга, журнал, копии WMS, "
                "книги WMS_TOOLBOX), проверка перед сменой — код 0",
         "резервная копия" in b and rb.returncode == 0 and hc.returncode == 0 and "WMS_PROD.ods" in in_copy
         and any(x.startswith("WMS_Journal/") for x in in_copy) and "WMS_TOOLBOX/WMS_MANAGER.ods" in in_copy
         and not any(x.endswith(".py") for x in in_copy), f"{b[:80]}; {rb.stdout.strip()[:80]}; {hc.stdout.strip().splitlines()[-1]}")
    man = {r[1]: r[4] for r in csv.reader(io.StringIO(open(os.path.join(snap, "manifest.csv"), encoding="utf-8-sig").read()), delimiter=";") if r and r[0] == "file"}
    okh = all(hashlib.sha256(open(os.path.join(snap, f), "rb").read()).hexdigest() == h for f, h in man.items())
    ncars = len([ln for ln in open(os.path.join(snap, "cars.csv"), encoding="utf-8-sig").read().splitlines()[1:] if ln.strip()]) \
        if os.path.exists(os.path.join(snap, "cars.csv")) else -1
    step("E06", "снимок для инструментов в WMS_Export: manifest, каждый файл совпадает по SHA-256; визиты машин — cars.csv (2) и service/_CAR.csv",
         x.startswith("OK") and okh and len(man) >= 18 and "cars.csv" in man and "service/_CAR.csv" in man and ncars == 2,
         f"{x[:80]}; файлов {len(man)}; визитов {ncars}")
    # ------------------------------------------------------------ E07 the tools
    tbx = os.path.join(W, "WMS_TOOLBOX")
    books = sorted(os.path.basename(p) for p in glob.glob(os.path.join(tbx, "*.ods")))
    step("E07", "WMS_TOOLBOX выпуска рядом с рабочей книгой: 9 инструментов и launcher (снимки — ../WMS_Export, пакеты — ../WMS_Batches)",
         len(books) == 10 and "WMS_TOOLBOX.ods" in books, str(books))

    def open_tool(name):
        o = Office(unique("e2e" + name[4:8].lower()), PROFILES)
        doc = o.load(os.path.join(tbx, name + ".ods"), macros=4)
        o.basic(doc, "TbCommon", "TbTestAuto", True)
        return o, doc

    o, doc = open_tool("WMS_DOCTOR")
    try:
        dr = o.basic(doc, "TbDoctor", "DoctorRun")
        drep = {r_[1]: r_[0] for r_ in doc.Sheets.getByName("Отчёт").getCellRangeByPosition(0, 3, 3, 40).getDataArray() if r_[0]}
    finally:
        o.terminate()
    o, doc = open_tool("WMS_SEARCH")
    try:
        o.basic(doc, "TbSearch", "SearchRefresh")
        card = o.basic(doc, "TbSearch", "CardShow", "710")
        hist = [r_[1] for r_ in doc.Sheets.getByName("Карточка").getCellRangeByPosition(3, 3, 8, 12).getDataArray() if r_[1]]
        sp = o.basic(doc, "TbSearch", "SearchRun", "а 123 вс")
        found = [r_[0] for r_ in doc.Sheets.getByName("Поиск").getCellRangeByPosition(0, 3, 0, 12).getDataArray() if r_[0]]
        vs = o.basic(doc, "TbSearch", "VehicleShow", "х456ор")
    finally:
        o.terminate()
    o, doc = open_tool("WMS_ANALYTICS")
    try:
        an = o.basic(doc, "TbAnalytics", "AnRefresh", 0.0)
    finally:
        o.terminate()
    o, doc = open_tool("WMS_MANAGER")
    try:
        ctl = o.basic(doc, "TbManager", "MgrControl", 0.0)
        cst = {r_[1]: (r_[0], r_[2]) for r_ in doc.Sheets.getByName("Контроль дня").getCellRangeByPosition(0, 3, 2, 20).getDataArray() if r_[0]}
    finally:
        o.terminate()
    o, doc = open_tool("WMS_LABELS")
    try:
        o.basic(doc, "TbLabels", "BtnLblRefresh")
        doc.Sheets.getByName("Этикетки").getCellRangeByPosition(0, 1, 1, 2).setDataArray((("710", ""), (str(e_rc), 2)))
        lb = o.basic(doc, "TbLabels", "LblBuild")
    finally:
        o.terminate()
    names = os.path.join(OUT, "названия.csv")
    open(names, "w", encoding="utf-8").write("Наименование\nшестерня z12\nудлинитель 10м\n")
    rr = subprocess.run([sys.executable, os.path.join(tbx, "reconcile", "wms_reconcile.py"), names, "--export", os.path.join(W, "WMS_Export")],
                        capture_output=True, text=True)
    step("E08", "инструменты на снимке: диагностика — ошибок 0; карточка ЕИ-00000710 с историей (пополнение, перемещения нет); сводка построена; "
                "этикетки (3); сверка старых названий находит перенесённую деталь «уверенно»",
         dr.startswith("OK:ошибок 0") and card.startswith("OK") and any(str(h).startswith("Пополнено") for h in hist) and an.startswith("OK")
         and lb.startswith("OK:этикеток 3") and rr.returncode == 0 and "ЕИ-00000710" in rr.stdout and "уверенно" in rr.stdout,
         f"{dr}; {card}; {hist}; {an[:60]}; {lb[:40]}")
    car_ctl = cst.get("Машины на территории", ("", ""))
    step("E08", "машины в инструментах: диагностика — «транспорт» OK; поиск госномера в другом написании — визит № 1; «История приездов» — "
                "визит № 2; сводка — визитов 2, на территории 1, уехало 1; «Контроль дня» — машина № 2 на территории (внимание)",
         drep.get("транспорт") == "OK" and sp.startswith("OK:найдено") and "Визит № 1" in found and vs.startswith("OK:визитов 1, на территории")
         and "транспорт: визитов 2, на территории 1, уехало 1" in an and "контроль дня" in ctl and car_ctl[0] == "ВНИМАНИЕ" and "№ 2" in str(car_ctl[1]),
         f"{drep.get('транспорт')}; {sp}; {found}; {vs}; {an[-80:]}; {ctl}; {car_ctl}")
    # ------------------------------------------------------------ E09 inventory by the tool → a batch → WMS
    o, doc = open_tool("WMS_INVENTORY")
    try:
        o.basic(doc, "TbInventory", "InvRefresh", False)
        sh = doc.Sheets.getByName("Пересчёт")
        eis = [r_[0] for r_ in sh.getCellRangeByPosition(0, 2, 0, 60).getDataArray()]
        k = eis.index(ei(160))
        book = sh.getCellByPosition(6, 2 + k).getValue()
        sh.getCellByPosition(7, 2 + k).setValue(book + 2)
        bt = o.basic(doc, "TbInventory", "InvBatch", D3)
    finally:
        o.terminate()
    batch = glob.glob(os.path.join(W, "WMS_Batches", "INV-*.csv"))
    s = Session(prod)
    try:
        s.U("TestUiAuto", 1)
        before = s.stock(160)
        # what «Загрузить пакет» does after the file dialog: the whole-batch check and load, then the posting of the rows
        # (the test seam of the file dialog works in test books only; this is the production book)
        msg = s.EX("LoadBatch", uno.systemPathToFileUrl(batch[0]))
        parts = msg[3:].split("|") if msg.startswith("OK:") else []
        # LoadBatch reports sheet rows 1-based (as shown to the operator); PostBatchRows takes 0-based indices
        post = s.EX("PostBatchRows", parts[0], int(parts[1]) - 1, int(parts[2]) - 1, parts[5] == "1") if parts else ""
        msg = f"{msg} → {post}"
        after = s.stock(160)
        s.doc.store()
    finally:
        s.close()
    step("E09", "инвентаризация инструментом: пакет (+2 у ЕИ-00000160) загружен и проведён WMS обычной операцией", bt.startswith("OK") and batch
         and after == before + 2, f"{bt[:80]}; {msg[:100]}; {before} → {after}")
    # ------------------------------------------------------------ E10 the end of the day
    s = Session(prod)
    try:
        st = s.state()
        sc = s.B("ActionSelfCheck")
        P, inf, _ = s.adjcheck(0, initial={})
        PC, infc = car_oracle.check(s.doc, s.jdir, 0)
    finally:
        s.close()
    hc = subprocess.run([sys.executable, os.path.join(tb_scripts, "wms_healthcheck.py"), "--wms", W, "--book", "WMS_PROD.ods"], capture_output=True, text=True)
    step("E10", "конец: открытие — «работа разрешена», самопроверка — ошибок 0, оракулы — без расхождений, проверка перед сменой — код 0",
         st.get("STATE") == "CLEAN" and sc.startswith("САМОПРОВЕРКА WMS: ошибок 0") and not P and not PC and hc.returncode == 0,
         f"{sc.splitlines()[0]}; {P[:3]}; {PC[:3]}; операций {inf.get('ops')}; {hc.stdout.strip().splitlines()[-1]}")
    upgrade_060(rel)


def upgrade_060(rel):
    """E11: the warehouse of release 0.6.0 — its release folder built by the sources of its commit, its candidate book put to
    work (registered, a receipt and an issue) — is upgraded by tools/upgrade.py of this release folder alone"""
    src = os.path.join(OUT, "src_060")
    rel6 = os.path.join(OUT, "release_060")
    shutil.rmtree(src, ignore_errors=True)
    shutil.rmtree(rel6, ignore_errors=True)
    why = git_sources(OLD_COMMIT, src)
    if why:
        R.add("E11", "обновление склада выпуска 0.6.0 инструментом папки выпуска", "SKIP", why)
        return
    env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
    r6 = subprocess.run([sys.executable, os.path.join(src, "tools", "build_release.py"), rel6], capture_output=True, text=True, timeout=1800,
                        cwd=src, env=env)
    w6 = os.path.join(OUT, "warehouse_060")
    shutil.rmtree(w6, ignore_errors=True)
    os.makedirs(w6)
    book = os.path.join(w6, "WMS_PROD.ods")
    shutil.copy2(os.path.join(rel6, "WMS_PROD_CANDIDATE.ods"), book)
    s = Session(book)
    try:
        v0 = s.B("StartupReport")
        s.special_input(1, B="Офис", D="Стеллаж металлический", F="3", G="шт", H=D1, I="A-7", K="Офис")
        r_sp = s.click_spc("BtnSpcPost", 1)
        s.doc.Sheets.getByName("Получатели").getCellRangeByPosition(0, 1, 1, 1).setDataArray((("ива", "Иванов Иван"),))
        s.issue_input(1, ei="1", qty="1", date=D2, who="ива")
        r_is = s.post(1)
        s.doc.store()
    finally:
        s.close()
    rc1, o1 = release_tool(rel, "upgrade.py", book, "--check")
    rc2, o2 = release_tool(rel, "upgrade.py", book)
    backups = glob.glob(os.path.join(w6, "WMS_Backups", "*_preupgrade.ods"))
    s = Session(book)
    try:
        st = s.state()
        v1 = s.B("StartupReport")
        sc = s.B("ActionSelfCheck")
        stock1 = s.stock(1)
        P, inf, _ = s.adjcheck(0, initial={})
        ca = s.B("CarArrive", "Газель А123ВС", "Кабель-Сервис", module="WmsCar")
        cd = s.B("CarDepart", 1, module="WmsCar")
        PC, infc = car_oracle.check(s.doc, s.jdir, 0)
        schema = s.sysv("SCHEMA")
        s.doc.store()
    finally:
        s.close()
    step("E11", "склад на выпуске 0.6.0 (папка выпуска собрана из исходников его commit, книга в работе: иной приход, выдача) обновлён "
                "tools/upgrade.py папки выпуска 0.7.0 (без репозитория): проверка и обновление — код 0, резервная копия *_preupgrade.ods; "
                "WMS 0.7 открывает книгу там же — «работа разрешена», самопроверка без ошибок, остаток и журнал сохранены (оракул), "
                "«Приход авто» работает (оракул машин)",
         r6.returncode == 0 and "ядро 0.6.0" in v0 and r_sp.startswith("OK") and r_is.startswith("OK") and rc1 == 0 and rc2 == 0
         and len(backups) == 1 and st.get("STATE") == "CLEAN" and "ядро 0.7.0" in v1 and sc.startswith("САМОПРОВЕРКА WMS: ошибок 0")
         and stock1 == 2.0 and not P and ca.startswith("OK:") and cd.startswith("OK:") and not PC and schema == "WMS-SYS-4",
         f"{r6.returncode} {(r6.stderr or r6.stdout)[-120:]}; {r_sp}; {r_is}; {rc1}/{rc2} {o2[-200:]}; {st.get('STATE')}; {sc[:50]}; "
         f"остаток {stock1}; {P[:3]}; {ca}; {cd}; {PC[:3]}")


if __name__ == "__main__":
    main()
