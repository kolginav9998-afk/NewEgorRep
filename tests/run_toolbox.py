"""WMS_TOOLBOX (FINAL WMS MARATHON §6–§8) — automated scenarios of the tools on the fixed snapshot fixture.

    WMS_TEST_OUT=/tmp/wms_tb python3 tests/run_toolbox.py [case ...]

The toolbox is built once (tools/build_toolbox.py) into WMS_TEST_OUT; every case copies it and the fixed snapshot
tests/fixtures/WMS_SNAPSHOT_FIXTURE (made once by tests/make_fixture_snapshot.py through WMS itself) into its own folder
as WMS would have them (…/WMS_TOOLBOX next to …/WMS_Export), opens a tool book with its macros and calls the functions its
buttons call (messages recorded, no windows). The expected values are computed here from the CSV files of the fixture,
independently of the Basic code. Batches written by the tools are loaded into a fresh WMS test book with its own loader
(the whole-batch validation of WMS). The scripts run as the user runs them (python3, exit codes, reports).

Results: WMS_TEST_OUT/results_toolbox.json and WMS_TEST_OUT/TEST_REPORT_toolbox.md.
"""
import csv
import glob
import io
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile

import uno  # noqa: F401  (LibreOffice Python-UNO)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(ROOT, "tools"))
from harness import OUT, PROFILES, Session, Results, git_sources, new_wms, unique  # noqa: E402
from wmslo import Office  # noqa: E402

R = Results()
CASES = []
FIXTURE = os.path.join(HERE, "fixtures", "WMS_SNAPSHOT_FIXTURE")
FIXTURE_060 = os.path.join(HERE, "fixtures", "WMS_SNAPSHOT_FIXTURE_060")      # a snapshot of WMS 0.6: contract 1.0, no vehicles
TODAY = 46291.0                                     # 26.09.2026: the fixed «today» of the fixture (analytics, archive)
_built = {}


def case(fn):
    CASES.append(fn)
    return fn


def toolbox():
    if "tb" not in _built:
        shutil.rmtree(os.path.join(OUT, "WMS_TOOLBOX"), ignore_errors=True)
        r = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "build_toolbox.py"), OUT], capture_output=True, text=True, timeout=1200)
        if r.returncode != 0:
            raise RuntimeError(r.stdout + r.stderr)
        _built["tb"] = os.path.join(OUT, "WMS_TOOLBOX")
    return _built["tb"]


def place(name, fixture=FIXTURE):
    """<case>/WMS_TOOLBOX (a copy of the built one) and <case>/WMS_Export/WMS_SNAPSHOT_FIXTURE"""
    d = os.path.join(OUT, "cases", name)
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(os.path.join(d, "WMS_Export"))
    shutil.copytree(toolbox(), os.path.join(d, "WMS_TOOLBOX"))
    shutil.copytree(fixture, os.path.join(d, "WMS_Export", "WMS_SNAPSHOT_FIXTURE"))
    return d


def rehash(snap, fname):
    """the manifest line of fname updated after a deliberate change (a snapshot that WMS itself exported this way)"""
    import hashlib
    p = os.path.join(snap, fname)
    h = hashlib.sha256(open(p, "rb").read()).hexdigest()
    man = os.path.join(snap, "manifest.csv")
    m = open(man, encoding="utf-8").read().splitlines()
    m = [f"file;{fname};{ln.split(';')[2]};{os.path.getsize(p)};{h}" if ln.startswith(f"file;{fname};") else ln for ln in m]
    open(man, "w", encoding="utf-8").write("\n".join(m) + "\n")


class Guard:
    """a runaway LibreOffice (a temporary file growing without end) is killed before it fills the disk: the step fails"""

    def __init__(self, tool, limit_mb=400):
        import threading
        self.tool, self.limit, self.t0, self.hit, self.stop = tool, limit_mb * 1_000_000, time.time(), False, False
        self.th = threading.Thread(target=self.run, daemon=True)
        self.th.start()

    def run(self):
        while not self.stop:
            tot = 0
            for f in glob.glob("/tmp/lu*.tmp/*"):
                try:
                    if os.path.getmtime(f) >= self.t0 - 1:
                        tot += os.path.getsize(f)
                except OSError:
                    pass
            if tot > self.limit:
                self.hit = True
                subprocess.run(["pkill", "-9", "-f", "soffice.bin.*" + self.tool.name], capture_output=True)
                for f in glob.glob("/tmp/lu*.tmp"):
                    try:
                        if os.path.getmtime(f) >= self.t0 - 1:
                            shutil.rmtree(f, ignore_errors=True)
                    except OSError:
                        pass
                return
            time.sleep(0.3)

    def done(self):
        self.stop = True
        return not self.hit


class Tool:
    """one tool book open with its macros in its own LibreOffice"""

    def __init__(self, folder, name):
        self.name = unique("tb" + name[4:8].lower())
        self.o = Office(self.name, PROFILES)
        self.path = os.path.join(folder, "WMS_TOOLBOX", name + ".ods")
        self.doc = self.o.load(self.path, macros=4)
        if self.doc is None:
            raise RuntimeError(f"{name} не открылся")
        self.auto = self.B("TbTestAuto", True, module="TbCommon")

    def B(self, func, *a, module):
        return self.o.basic(self.doc, module, func, *a)

    def sheet(self, name):
        return self.doc.Sheets.getByName(name)

    def rows(self, name, r0, c1, r1=None):
        sh = self.sheet(name)
        if r1 is None:
            cur = sh.createCursor()
            cur.gotoEndOfUsedArea(False)
            r1 = cur.getRangeAddress().EndRow
        if r1 < r0:
            return []
        return [list(r) for r in sh.getCellRangeByPosition(0, r0, c1, r1).getDataArray()]

    def setting(self, r, v):
        self.sheet("Настройки").getCellByPosition(1, r).setString(v)

    def close(self):
        try:
            self.doc.setModified(False)
            self.doc.close(True)
        except Exception:
            pass
        self.o.terminate()


def table(name, fixture=FIXTURE):
    """a table of the fixture as dicts (the keys of the contract)"""
    txt = open(os.path.join(fixture, name + ".csv"), encoding="utf-8-sig").read()
    rows = list(csv.reader(io.StringIO(txt), delimiter=";"))
    return [dict(zip(rows[0], r)) for r in rows[1:]]


def num(s):
    return float(s) if s not in ("", None) else 0.0


# ================================================================ K01 build and compile

@case
def k01_build_and_compile():
    c = "K01"
    tb = toolbox()
    books = sorted(os.path.basename(p) for p in glob.glob(os.path.join(tb, "*.ods")))
    want = ["WMS_ANALYTICS.ods", "WMS_ARCHIVE.ods", "WMS_DOCS.ods", "WMS_DOCTOR.ods", "WMS_IMPORTER.ods", "WMS_INVENTORY.ods", "WMS_LABELS.ods",
            "WMS_MANAGER.ods", "WMS_SEARCH.ods", "WMS_TOOLBOX.ods"]
    scripts = sorted(os.path.relpath(p, tb) for p in glob.glob(os.path.join(tb, "*", "*.py")))
    R.add(c, "сборка WMS_TOOLBOX: 9 книг инструментов и launcher WMS_TOOLBOX.ods, скрипты сверки, резервного копирования, проверки, "
             "TSPL и закономерностей, README, контракт, версия",
          books == want and scripts == ["backup/wms_backup.py", "backup/wms_healthcheck.py", "insights/wms_insights.py", "labels/tspl_labels.py",
                                        "reconcile/wms_reconcile.py"]
          and all(os.path.exists(os.path.join(tb, f)) for f in ("README.md", "EXPORT_CONTRACT.md", "VERSION")), f"{books}; {scripts}")
    d = place("k01")
    res = {}
    for b in want:
        t = Tool(d, b[:-4])
        try:
            res[b] = t.auto
        finally:
            t.close()
    R.add(c, "каждая книга открывается с макросами и её Basic компилируется (вызов общего модуля — «OK»)", all(v == "OK" for v in res.values()), str(res))
    # every button of every book is bound to a Sub of a module of that book
    import re
    bad, nb = [], 0
    for b in want:
        z = zipfile.ZipFile(os.path.join(tb, b))
        content = z.read("content.xml").decode("utf-8")
        subs = set()
        for nm in z.namelist():
            if nm.startswith("Basic/Standard/") and nm.endswith(".xml") and "script-lb" not in nm:
                mod = os.path.basename(nm)[:-4]
                src = z.read(nm).decode("utf-8")
                subs |= {f"{mod}.{m}" for m in re.findall(r"^Sub (\w+)", src, re.M)}
        for macro in re.findall(r"vnd\.sun\.star\.script:Standard\.([\w.]+)\?language=Basic", content):
            nb += 1
            if macro not in subs:
                bad.append(f"{b}: {macro}")
    R.add(c, "каждая кнопка каждой книги привязана к существующей процедуре своей книги; у каждой из 9 книг инструментов — «Все инструменты»",
          nb == 57 and not bad, f"кнопок {nb}; без процедуры {bad}")
    again = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "build_toolbox.py"), OUT], capture_output=True, text=True, timeout=600)
    R.add(c, "повторная сборка в папку с WMS_TOOLBOX — отказ, код 2 (в книгах работающего набора — журнал работы и история актов)",
          again.returncode == 2 and "уже существует" in again.stderr, again.stderr.strip()[:160])


# ================================================================ K02 INVENTORY

@case
def k02_inventory():
    c = "K02"
    d = place("k02")
    t = Tool(d, "WMS_INVENTORY")
    try:
        res = t.B("InvRefresh", False, module="TbInventory")
        rows = [r for r in t.rows("Пересчёт", 2, 9) if r[0]]
        stock = [s for s in table("stock") if s["ei"] and s["state"] != "Приход удалён (сторно)"]
        byei = {r[0]: r for r in rows}
        R.add(c, "«Обновить из снимка»: строка пересчёта на каждый ЕИ снимка (кроме сторнированных) — ЕИ, наименование, артикул, "
                 "единица, место, категория, учётный остаток; лист показывает снимок",
              res.startswith("OK") and len(rows) == len(stock) and all(byei[s["ei"]][6] == num(s["qty"]) and byei[s["ei"]][4] == s["place"]
                                                                        for s in stock)
              and "WMS_SNAPSHOT_FIXTURE" in t.sheet("Пересчёт").getCellByPosition(0, 0).getString(), f"{res}; строк {len(rows)} из {len(stock)}")
        # the storekeeper types the facts: two differences, one equal, one bad value
        sh = t.sheet("Пересчёт")
        pos = {r[0]: 2 + k for k, r in enumerate(rows)}
        sh.getCellByPosition(7, pos["ЕИ-00000201"]).setValue(150)          # book 155 → −5
        sh.getCellByPosition(7, pos["ЕИ-00000206"]).setValue(9)            # book 8 → +1
        sh.getCellByPosition(7, pos["ЕИ-00000001"]).setValue(95)           # book 95 → no difference
        sh.getCellByPosition(9, pos["ЕИ-00000201"]).setString("ящик вскрыт")
        sh.getCellByPosition(7, pos["ЕИ-00000002"]).setString("много")
        bad = t.B("InvBatch", "26.09.2026", module="TbInventory")
        sh.getCellByPosition(7, pos["ЕИ-00000002"]).setValue(240)
        t.doc.calculateAll()
        diff = sh.getCellByPosition(8, pos["ЕИ-00000201"]).getValue()
        ok = t.B("InvBatch", "26.09.2026", module="TbInventory")
        files = glob.glob(os.path.join(d, "WMS_Batches", "INV-*.csv"))
        txt = open(files[0], encoding="utf-8").read() if files else ""
        lines = [ln for ln in txt.splitlines() if ln and not ln.startswith("#")]
        R.add(c, "ошибка ввода (не число) — пакет не создан; затем пакет только из строк с разницей (−5, +1; равная не входит): "
                 "WMS-BATCH-1 «Корректировки», id, источник, снимок; колонки kind, ei, fact, book, date, reason, note; разница на листе",
              bad.startswith("ERR") and "не число" in bad and ok.startswith("OK") and len(files) == 1 and txt.startswith("#WMS-BATCH-1\n")
              and "#target;Корректировки" in txt and "#source;WMS_INVENTORY" in txt and "#snapshot;WMS_SNAPSHOT_FIXTURE" in txt
              and lines[0] == "kind;ei;fact;book;date;reason;note"
              and sorted(lines[1:]) == sorted(["Инвентаризация;ЕИ-00000201;150;155;26.09.2026;Инвентаризация 26.09.2026: ящик вскрыт;",
                                               "Инвентаризация;ЕИ-00000206;9;8;26.09.2026;Инвентаризация 26.09.2026;"]) and diff == -5.0,
              f"{bad[:80]}; {ok[:80]}; {lines}")
    finally:
        t.close()
    # the batch in WMS: the loader of WMS accepts it as a whole (a fresh test book: the rows load; the book balances differ there)
    s = Session(new_wms("k02wms"))
    try:
        s.U("TestUiAuto", 2)                                         # «Провести сейчас?» — «Нет»
        s.EX("TestBatchFile", files[0])
        s.EX("BtnLoadBatch")
        msg = s.U("TestUiLastMessage")
        adj = [s.adj(r) for r in (1, 2)]
        R.add(c, "пакет инвентаризации загружается загрузчиком WMS (проверка целиком пройдена): две строки ввода на «Корректировки» "
                 "с отметкой пакета, ничего не проведено",
              [a[1] for a in adj] == ["Инвентаризация", "Инвентаризация"] and all(a[17].startswith("INV-") for a in adj) and all(a[0] == "" for a in adj),
              f"{msg[:120]}; {[a[:3] for a in adj]}")
    finally:
        s.close()


# ================================================================ K03 SEARCH

@case
def k03_search():
    c = "K03"
    d = place("k03")
    t = Tool(d, "WMS_SEARCH")
    try:
        res = t.B("SearchRefresh", module="TbSearch")
        r1 = t.B("SearchRun", "gr-40", module="TbSearch")
        f1 = [r for r in t.rows("Поиск", 3, 9) if r[0]]
        r2 = t.B("SearchRun", "иванов", module="TbSearch")
        f2 = [r for r in t.rows("Поиск", 3, 9) if r[0]]
        iv = [i for i in table("issues") if i["recipient"].startswith("Иванов")]
        r3 = t.B("SearchRun", "ЕИ-206", module="TbSearch")
        f3 = [r for r in t.rows("Поиск", 3, 9) if r[0]]
        R.add(c, "поиск: артикул «gr-40» (без учёта регистра) — ЕИ-00000206; получатель «иванов» — все его выдачи (с датами); ЕИ «ЕИ-206» — "
                 "одна карточка",
              res.startswith("OK") and r1.startswith("OK") and [r[1] for r in f1] == ["ЕИ-00000206"] and len(f2) == len(iv)
              and all(r[0].startswith("Выдача № ") for r in f2) and f2[0][8] == "15.07.2026" and [r[1] for r in f3] == ["ЕИ-00000206"],
              f"{r1}; {r2}; {r3}; {[r[1] for r in f2]}")
        rc = t.B("CardShow", "201", module="TbSearch")
        card = t.rows("Карточка", 2, 1, 11)
        hist = [r for r in t.rows("Карточка", 3, 8) if r[3]]
        ev = [(h[3], h[4].split(" (")[0], h[5]) for h in hist]
        R.add(c, "карточка ЕИ-00000201: данные «Наличие» и история по датам — пришло 200 (заказ З-101), выдано 20, возвращено 5, выдано 30, "
                 "выдача 1 — «Удалено (сторно)»",
              rc.startswith("OK") and card[0][1] == "ЕИ-00000201" and card[4][1] == 155.0
              and ev == [("2026-07-10", "Пришло", 200.0), ("2026-07-15", "Выдано", -20.0), ("2026-07-30", "Возвращено", 5.0),
                         ("2026-08-20", "Выдано", -30.0), ("2026-09-19", "Выдано", -1.0)] and hist[-1][8] == "Удалено (сторно)",
              f"{rc}; {ev}")
        bad = t.B("CardShow", "999", module="TbSearch")
        R.add(c, "карточка несуществующего ЕИ — отказ «нет в снимке»", bad.startswith("ERR") and "нет в снимке" in bad, bad)
        # M6: vehicles and suppliers — the expected rows computed from cars.csv / orders.csv of the fixture
        cars = [v for v in table("cars") if v["visit_no"]]
        r4 = t.B("SearchRun", "a 123 bc", module="TbSearch")                    # Latin letters, spaces: the plate А123ВС
        f4 = [r for r in t.rows("Поиск", 3, 9) if r[0]]
        want4 = [f"Визит № {v['visit_no']}" for v in cars if v["vehicle_key"] == "А123ВС"]
        r5 = t.B("SearchRun", "метиз", module="TbSearch")
        f5 = [r for r in t.rows("Поиск", 3, 9) if r[0]]
        want5 = sorted([f"Заказ {o['order_no']}" for o in table("orders") if "метиз" in o["supplier"].lower()]
                       + [f"Визит № {v['visit_no']}" for v in cars if "метиз" in v["supplier"].lower()])
        R.add(c, "поиск: госномер в любом написании («a 123 bc») — все визиты этой машины; поставщик «метиз» — его заказы и визиты машин, "
                 "имя поставщика перенесено на лист «Поставщик»",
              r4.startswith("OK") and [r[0] for r in f4] == want4 and len(want4) == 2 and sorted(r[0] for r in f5) == want5
              and "поставщик «Метиз-Опт»" in r5 and t.sheet("Поставщик").getCellByPosition(1, 0).getString() == "Метиз-Опт", f"{r4}; {[r[0] for r in f4]}; {r5}")
        rs = t.B("SupplierShow", "кабель", module="TbSearch")
        sup = t.rows("Поставщик", 0, 8)
        o_rows = [r for r in sup if str(r[0]).startswith("З-")]
        v_rows = [r for r in sup if isinstance(r[0], float) and r[0] > 0]
        want_o = [o["order_no"] for o in table("orders") if "кабель" in o["supplier"].lower()]
        want_v = [float(v["visit_no"]) for v in cars if "кабель" in v["supplier"].lower()]
        closed = [float(v["duration_min"]) for v in cars if "кабель" in v["supplier"].lower() and v["status"] == "Уехал"]
        R.add(c, "«Заказы и машины поставщика»: все строки заказов и все визиты поставщика; итог — заказы, визиты, средняя стоянка",
              rs == f"OK:заказов {len(want_o)}, визитов {len(want_v)}" and [r[0] for r in o_rows] == want_o and [r[0] for r in v_rows] == want_v
              and f"средняя стоянка {int(sum(closed) / len(closed) + 0.5)} мин" in str(sup[1][0]), f"{rs}; {sup[1][0]}")
        rv = t.B("VehicleShow", "х 456 ор 77", module="TbSearch")
        veh = [r for r in t.rows("Машина", 5, 7) if r[0]]
        want_v2 = [float(v["visit_no"]) for v in cars if v["vehicle_key"] == "Х456ОР"]
        rn = t.B("VehicleShow", "Р999РР", module="TbSearch")
        R.add(c, "«История приездов»: госномер с регионом и пробелами — все визиты машины по датам (как записано, поставщик, приезд, выезд, "
                 "минуты, статус); неизвестный номер — «визитов нет»",
              rv == f"OK:визитов {len(want_v2)}" and [r[0] for r in veh] == want_v2 and veh[0][2] == "КАМАЗ Х456ОР" and veh[0][7] == "Уехал"
              and rn == "OK:визитов 0", f"{rv}; {[r[:3] for r in veh]}; {rn}")
    finally:
        t.close()


# ================================================================ K04 DOCTOR

def doctor_run(d):
    t = Tool(d, "WMS_DOCTOR")
    try:
        res = t.B("DoctorRun", module="TbDoctor")
        rows = [r for r in t.rows("Отчёт", 3, 3) if r[0]]
        return res, {r[1]: (r[0], r[2], r[3]) for r in rows}
    finally:
        t.close()


@case
def k04_doctor():
    c = "K04"
    d = place("k04")
    res, rep = doctor_run(d)
    need = ("снимок", "версия схемы", "контракт снимка", "защиты", "журнал: формат", "снимок против журнала", "остатки", "ЕИ без карточки",
            "связи возвратов", "счётчики", "артикулы деталей", "транспорт", "пакеты: незавершённая загрузка", "комплект инструментов", "лишние файлы")
    R.add(c, "аудит снимка-фикстуры: снимок цел, версия схемы и контракта (1.1), защиты, журнал, остатки = сумме движений, ЕИ без карточки нет, "
             "связи возвратов, счётчики, артикулы деталей, визиты машин, пакеты, комплект инструментов, лишних файлов нет — OK; машина, "
             "оставшаяся «На территории» с прошлого дня, и нет резервных копий рядом — предупреждения",
          res.startswith("OK:ошибок 0") and all(rep.get(k, ("",))[0] == "OK" for k in need) and rep.get("резервные копии", ("",))[0] == "WARN"
          and rep.get("транспорт: машины с прошлых дней", ("", ""))[0] == "WARN" and "визит № 8" in rep["транспорт: машины с прошлых дней"][1]
          and "WMS-SYS-4" in rep["версия схемы"][1] and "1.1" in rep["контракт снимка"][1],
          f"{res}; " + "; ".join(f"{k}={v[0]}" for k, v in rep.items()))
    # a snapshot whose table was changed after the export: not used
    snap = os.path.join(d, "WMS_Export", "WMS_SNAPSHOT_FIXTURE")
    with open(os.path.join(snap, "stock.csv"), "a", encoding="utf-8") as f:
        f.write("ЕИ-00099999;лишняя строка;;шт;1;;;;;\n")
    res2, rep2 = doctor_run(d)
    R.add(c, "таблица снимка изменена после экспорта (SHA-256) — «снимок» FAIL, остальные проверки не выполняются",
          rep2.get("снимок", ("",))[0] == "FAIL" and "изменён" in rep2["снимок"][1] and len(rep2) == 1, f"{res2}; {rep2}")
    # a consistent snapshot with a wrong balance (manifest re-hashed: a book with a real inconsistency)
    d2 = place("k04b")
    snap = os.path.join(d2, "WMS_Export", "WMS_SNAPSHOT_FIXTURE")
    p = os.path.join(snap, "stock.csv")
    txt = open(p, encoding="utf-8-sig").read().replace("ЕИ-00000203;Перчатки нитриловые;GL-NIT-M;пар;46;", "ЕИ-00000203;Перчатки нитриловые;GL-NIT-M;пар;47;")
    open(p, "w", encoding="utf-8-sig").write(txt)
    import hashlib
    h = hashlib.sha256(open(p, "rb").read()).hexdigest()
    man = os.path.join(snap, "manifest.csv")
    m = open(man, encoding="utf-8").read().splitlines()
    m = [f"file;stock.csv;{ln.split(';')[2]};{os.path.getsize(p)};{h}" if ln.startswith("file;stock.csv;") else ln for ln in m]
    open(man, "w", encoding="utf-8").write("\n".join(m) + "\n")
    res3, rep3 = doctor_run(d2)
    R.add(c, "остаток ЕИ-00000203 не равен сумме движений (47 при движениях 50 − 4 = 46) — «остатки» FAIL с первым ЕИ и планом",
          rep3.get("остатки", ("",))[0] == "FAIL" and "ЕИ-00000203" in rep3["остатки"][1], f"{res3}; {rep3.get('остатки')}")
    # the protections of the book as WMS wrote them into the manifest (not hashed: written by WMS itself)
    d3 = place("k04c")
    man = os.path.join(d3, "WMS_Export", "WMS_SNAPSHOT_FIXTURE", "manifest.csv")
    m = open(man, encoding="utf-8").read().replace("sheet;Выдачи;1;1", "sheet;Выдачи;0;1").replace("sheet;_SYS;1;0", "sheet;_SYS;1;1") \
        .replace("structure;1", "structure;0")
    open(man, "w", encoding="utf-8").write(m)
    res4, rep4 = doctor_run(d3)
    pr = rep4.get("защиты", ("", "", ""))
    R.add(c, "защиты по manifest снимка: лист «Выдачи» без защиты, служебный «_SYS» не скрыт, структура не защищена — предупреждение с "
             "названиями и планом (защиту восстанавливает ответственный за WMS); «Получатели» без защиты — норма",
          pr[0] == "WARN" and "лист «Выдачи» не защищён" in pr[1] and "служебный лист «_SYS» не скрыт" in pr[1] and "структура книги не защищена" in pr[1]
          and "Получатели" not in pr[1] and "Защитить лист" in pr[2], f"{res4}; {pr}")
    # M6: damaged visits (a departure before the arrival, an empty supplier) in a consistent snapshot
    d5 = place("k04d")
    snap = os.path.join(d5, "WMS_Export", "WMS_SNAPSHOT_FIXTURE")
    p = os.path.join(snap, "cars.csv")
    lines = open(p, encoding="utf-8").read().split("\n")
    lines[1] = lines[1].replace("2026-09-24;Газель А123ВС;Метиз-Опт;2026-09-24 09:10:00;2026-09-24 09:55:00",
                                "2026-09-24;Газель А123ВС;Метиз-Опт;2026-09-24 09:10:00;2026-09-24 08:55:00")
    lines[3] = lines[3].replace(";Лада В789ТТ;СИЗ-Центр;", ";Лада В789ТТ;;")
    open(p, "w", encoding="utf-8").write("\n".join(lines))
    rehash(snap, "cars.csv")
    res5, rep5 = doctor_run(d5)
    tr = rep5.get("транспорт", ("", "", ""))
    R.add(c, "повреждённые визиты: выезд раньше приезда (визит № 1) и пустой «Поставщик» (визит № 3) — «транспорт» FAIL с первым нарушением и планом",
          tr[0] == "FAIL" and "нарушений 2" in tr[1] and "визит № 1: выезд" in tr[1] and "раньше приезда" in tr[1] and "«Приход авто»" in tr[2],
          f"{res5}; {tr}")
    # M6: an unfinished batch import, a batch that WMS never loaded, stray files, a script missing from the set
    d6 = place("k04e")
    snap = os.path.join(d6, "WMS_Export", "WMS_SNAPSHOT_FIXTURE")
    with open(os.path.join(snap, "adjustments.csv"), "a", encoding="utf-8") as f:
        f.write(";Инвентаризация;ЕИ-00000001;;;;;90;95;;2026-09-26;пересчёт;;;;;;INV-TEST-1;\n")
    rehash(snap, "adjustments.csv")
    os.makedirs(os.path.join(d6, "WMS_Batches"))
    with open(os.path.join(d6, "WMS_Batches", "INV-TEST-2.csv"), "w", encoding="utf-8") as f:
        f.write("#WMS-BATCH-1\n#id;INV-TEST-2\n#target;Корректировки\nkind;ei;fact;book;date\nИнвентаризация;ЕИ-00000002;1;2;26.09.2026\n")
    os.makedirs(os.path.join(d6, "WMS_Export", ".WMS_SNAPSHOT_20260926-180000_seq41.part"))
    open(os.path.join(d6, ".WMS_PROD.upgrade.ods"), "wb").write(b"x")
    os.remove(os.path.join(d6, "WMS_TOOLBOX", "insights", "wms_insights.py"))
    res6, rep6 = doctor_run(d6)
    g = lambda k: rep6.get(k, ("", "", ""))
    R.add(c, "пакет загружен, но не проведён (INV-TEST-1); пакет в папке, которого нет в WMS (INV-TEST-2); незавершённый снимок .part и остаток "
             "обновления; нет скрипта insights — предупреждения с планом; ничего не исправлено",
          g("пакеты: незавершённая загрузка")[0] == "WARN" and "INV-TEST-1" in g("пакеты: незавершённая загрузка")[1]
          and g("пакеты: ожидают загрузки")[0] == "WARN" and "INV-TEST-2" in g("пакеты: ожидают загрузки")[1]
          and g("лишние файлы")[0] == "WARN" and ".part" in g("лишние файлы")[1] and ".WMS_PROD.upgrade.ods" in g("лишние файлы")[1]
          and g("комплект инструментов")[0] == "WARN" and "insights/wms_insights.py" in g("комплект инструментов")[1]
          and os.path.exists(os.path.join(d6, ".WMS_PROD.upgrade.ods")) and os.path.exists(os.path.join(d6, "WMS_Batches", "INV-TEST-2.csv")),
          f"{res6}; " + "; ".join(f"{k}={v[0]}: {v[1][:60]}" for k, v in rep6.items() if v[0] != "OK"))

# ================================================================ K05 ANALYTICS

def ser(iso):
    import datetime
    y, m, d = map(int, iso.split("-"))
    return (datetime.date(y, m, d) - datetime.date(1899, 12, 30)).days


def month_counts(y, m):
    """independent counts of the fixture for month y-m: receipts (order receipts + special lines), issues (posted), returns"""
    lo, hi = f"{y:04d}-{m:02d}-01", f"{y:04d}-{m:02d}-31"
    rc = sum(1 for o in table("orders") if o["ei"] and lo <= o["date_received"] <= hi)
    rc += sum(1 for sp in table("special") if sp["line_no"] and lo <= sp["date"] <= hi)
    iss = sum(1 for i in table("issues") if i["issue_no"] and i["status"].startswith("Проведено") and lo <= i["date"] <= hi)
    ret = sum(1 for r in table("returns") if r["return_no"] and r["status"].startswith("Проведено") and lo <= r["date"] <= hi)
    return rc, iss, ret


@case
def k05_analytics():
    c = "K05"
    d = place("k05")
    t = Tool(d, "WMS_ANALYTICS")
    try:
        res = t.B("AnRefresh", float(ser("2026-09-26")), module="TbAnalytics")
        t.doc.calculateAll()
        rows = t.rows("Сводка", 0, 7)
        val = {str(r[0]): r[1:] for r in rows if r[0]}
        stock = [s for s in table("stock") if s["ei"]]
        src = {k: sum(1 for s in stock if s["source_type"] == k) for k in ("Поставщик", "Офис", "Производство", "Детали", "Старый склад", "Иной приход")}
        months = {f"{m:02d}.2026": month_counts(2026, m) for m in (7, 8, 9)}
        got_m = {k: (val[k][0], val[k][1], val[k][2]) for k in months if k in val}
        R.add(c, "«Сводка»: ЕИ в реестре, ЕИ с остатком, источники (ЕИ и остаток) — как в снимке; строки месяцев 07–09.2026: приходы, выдачи "
                 "(только проведённые), возвраты — как посчитано по таблицам",
              res.startswith("OK") and val["ЕИ в реестре"][0] == len(stock) and val["ЕИ с остатком"][0] == sum(1 for s in stock if num(s["qty"]) > 0)
              and all(val[k][0] == v for k, v in src.items()) and got_m == {k: tuple(float(x) for x in v) for k, v in months.items()},
              f"{res[:80]}; источники {[(k, val.get(k, [''])[0]) for k in src]}; месяцы {got_m} ≠? {months}")
        adj = {k: val[k] for k in ("09.2026",)}
        late = val.get("Ожидаемая дата прошла (заказ ждём)", [None])[0]
        R.add(c, "сентябрь: списаний 1, инвентаризаций 2, перемещений 1; заказы: ожидаемая дата прошла у 1 (З-104: снимок WMS до 0.7.2 "
                 "со статусом «Просрочено», ожидаемая дата 20.08.2026 — открытая позиция, считается по дате); отрицательных остатков 0",
              adj["09.2026"][3:6] == [1.0, 2.0, 1.0] and late == 1.0 and "Просрочено" not in val
              and val["Отрицательные остатки (должно быть 0)"][0] == 0.0, f"{adj}; ожидаемая дата прошла {late}")
        top = [r for r in rows if str(r[0]).startswith("Иванов")]
        R.add(c, "получатели: Иванов Иван Андреевич — первым, выдач 4 (проведённых), количество 20 + 25 + 5 + 4",
              bool(top) and top[0][1] == 4.0 and top[0][2] == 54.0, str(top[:1]))
        # M6: the most issued positions (90 days, posted), the vehicles («Транспорт»), the charts, the patterns
        import datetime
        t90 = (datetime.date(2026, 9, 26) - datetime.timedelta(days=89)).isoformat()
        cnt = {}
        for i in table("issues"):
            if i["issue_no"] and i["status"].startswith("Проведено") and i["date"] >= t90:
                cnt[i["ei"]] = cnt.get(i["ei"], 0) + 1
        k = [str(r[0]) for r in rows].index("НАИБОЛЕЕ ВЫДАВАЕМЫЕ ПОЗИЦИИ (90 дней)")
        first = rows[k + 1]
        R.add(c, "наиболее выдаваемые позиции за 90 дней: первой — ЕИ с наибольшим числом проведённых выдач (с наименованием)",
              str(first[0]).split(" «")[0] == max(cnt, key=lambda e: (cnt[e], -int(e[3:]))) and first[1] == float(max(cnt.values())), f"{first}; {cnt}")
        cars = [v for v in table("cars") if v["visit_no"] and v["status"] != "Отменён"]
        closed = sorted(float(v["duration_min"]) for v in cars if v["status"] == "Уехал")
        med = (closed[len(closed) // 2] + closed[(len(closed) - 1) // 2]) / 2
        tv = {str(r[0]): r[1] for r in t.rows("Транспорт", 0, 1) if r[0]}
        sup = {}
        for v in cars:
            if v["status"] == "Уехал":
                sup.setdefault(v["supplier"], []).append(float(v["duration_min"]))
        worst = max(sup, key=lambda x: sum(sup[x]) / len(sup[x]))
        srows = t.rows("Транспорт", 0, 6)
        hdr = [i for i, r in enumerate(srows) if r[0] == "Поставщик"][0]
        R.add(c, "лист «Транспорт»: визитов (без отменённых), отменено, на территории, уехало, средняя / медиана / максимум стоянки, доля дольше "
                 "порога; пик приездов — «Недостаточно данных» (визитов меньше 20); поставщик с самой долгой стоянкой — первым",
              tv.get("Визитов (без отменённых)") == float(len(cars)) and tv.get("Отменено визитов") == 1.0
              and tv.get("На территории (на момент снимка)") == float(sum(1 for v in cars if v["status"] == "На территории"))
              and tv.get("Уехало (закрытых визитов)") == float(len(closed)) and tv.get("Средняя стоянка, мин") == round(sum(closed) / len(closed), 1)
              and tv.get("Медиана стоянки, мин") == med and tv.get("Доля стоянок дольше 120 мин") == f"{sum(1 for x in closed if x > 120) / len(closed):.0%}"
              and str(tv.get("Пик приездов (два часа)", "")).startswith("Недостаточно данных") and srows[hdr + 1][0] == worst,
              f"{tv}; первый поставщик {srows[hdr + 1][0]} (ожидался {worst})")
        ch = t.sheet("Графики").getCharts()
        names = list(ch.getElementNames())
        filled = []
        for n in names:
            a = ch.getByName(n).getRanges()[0]
            src = t.doc.Sheets.getByIndex(a.Sheet)
            data = src.getCellRangeByPosition(a.StartColumn, a.StartRow + 1, a.EndColumn, a.EndRow).getDataArray()
            filled.append(len(data) > 0 and all(str(r[0]) != "" for r in data))
        R.add(c, "лист «Графики»: 8 диаграмм LibreOffice над таблицами (движения по месяцам, категории, источники, места, наиболее выдаваемые, "
                 "получатели, машины по дням и по поставщикам); у каждой — непустой диапазон данных",
              names == ["months", "categories", "sources", "places", "top_issued", "recipients", "cars_days", "cars_suppliers"] and all(filled),
              f"{names}; {filled}")
        ins = t.B("AnInsights", "2026-09-26", module="TbAnalytics")
        outs = sorted(glob.glob(os.path.join(d, "WMS_Reports", "INSIGHTS_*")))
        cons = list(csv.DictReader(open(os.path.join(outs[-1], "consumption.csv"), encoding="utf-8"), delimiter=";")) if outs else []
        found = t.rows("Найденные закономерности", 3, 6)
        crow = [r for r in t.rows("Расход", 3, 13) if r[0]]
        R.add(c, "«Найти закономерности»: движок по снимку (python3) — файлы в ../WMS_Reports/INSIGHTS_…, лист «Найденные закономерности» "
                 "(уровни, основание, уверенность, что сделать; мало истории — «не найдено»), лист «Расход» — каждый ЕИ файла, "
                 "без истории — «Недостаточно данных»",
              ins.startswith("OK:закономерностей") and outs and found[0][:4] == ["Уровень", "Тема", "Объект", "Что обнаружено"]
              and len(crow) == len(cons) and all(r[12] == x["confidence"] for r, x in zip(crow, cons))
              and any(str(r[13]).startswith("Недостаточно данных") for r in crow), f"{ins[:160]}; строк расхода {len(crow)} из {len(cons)}")
    finally:
        t.close()


# ================================================================ K06 MANAGER

@case
def k06_manager():
    c = "K06"
    d = place("k06")
    t = Tool(d, "WMS_MANAGER")
    try:
        r0 = t.B("MgrRefresh", module="TbManager")
        rep = t.sheet("Отчёт")
        rep.getCellByPosition(1, 1).setString("Месяц")
        rep.getCellByPosition(1, 2).setValue(ser("2026-09-15"))
        log = t.sheet("Журнал работы")
        log.getCellRangeByPosition(0, 1, 3, 3).setDataArray(((ser("2026-09-10"), "Разгрузка / погрузка", "фура с кабелем", 2.5),
                                                             (ser("2026-09-11"), "Работа погрузчиком", "перестановка стеллажа", 1.0),
                                                             (ser("2026-08-30"), "Закупки", "не этот месяц", 3.0)))
        r1 = t.B("MgrReport", module="TbManager")
        t.doc.calculateAll()
        rows = [r for r in t.rows("Отчёт", 6, 1) if r[0]]
        val = {str(r[0]): r[1] for r in rows}
        rc, iss, ret = month_counts(2026, 9)
        R.add(c, "отчёт за 09.2026: выдач 3 (сторно не считается), выдано 10,5 + 5 + 4, возвратов 0, списаний 1, инвентаризаций 2, "
                 "перемещений 1; журнал работы: разгрузка 2,5 ч, погрузчик 1 ч, всего 3,5 ч (август не входит)",
              r0.startswith("OK") and r1.startswith("OK") and val.get("Выдач") == iss and val.get("Выдано, количество") == 19.5
              and val.get("Возвратов") == ret and val.get("Списание (операций)") == 1.0 and val.get("Инвентаризация (операций)") == 2.0
              and val.get("Перемещение (операций)") == 1.0 and val.get("Разгрузка / погрузка") == 2.5 and val.get("Работа погрузчиком") == 1.0
              and val.get("Всего часов") == 3.5, f"{r1}; {val}")
        pdf = t.B("MgrPdf", module="TbManager")
        files = glob.glob(os.path.join(d, "WMS_Reports", "Report_2026-09-01_2026-09-30.pdf"))
        R.add(c, "«В PDF» — файл отчёта за период в ../WMS_Reports", pdf.startswith("OK") and len(files) == 1 and os.path.getsize(files[0]) > 1000, pdf)
        # M6: the vehicles of the period, the EIs in movement — independently from the fixture
        cars = [v for v in table("cars") if v["visit_no"] and v["status"] != "Отменён"]
        sep = lambda x: x.startswith("2026-09")
        closed = [float(v["duration_min"]) for v in cars if v["status"] == "Уехал" and sep(v["departed"])]
        moved = {o["ei"] for o in table("orders") if o["ei"] and sep(o["date_received"])}
        moved |= {x["ei"] for x in table("special") if x["line_no"] and sep(x["date"])}
        moved |= {x["ei"] for x in table("issues") if x["issue_no"] and x["status"].startswith("Проведено") and sep(x["date"])}
        moved |= {x["ei"] for x in table("returns") if x["return_no"] and x["status"] != "Удалено (сторно)" and sep(x["date"])}
        moved |= {x["ei"] for x in table("adjustments") if x["adj_no"] and x["status"] != "Удалено (сторно)" and sep(x["date"])}
        R.add(c, "отчёт за 09.2026 (M6): машин приехало / уехало, среднее время на территории, дольше 120 мин, на территории; ЕИ в движении — "
                 "как по таблицам снимка",
              val.get("Машин приехало") == float(sum(1 for v in cars if sep(v["arrived"]))) and val.get("Машин уехало") == float(len(closed))
              and val.get("Среднее время машины на территории, мин") == round(sum(closed) / len(closed), 1)
              and val.get("Дольше 120 мин") == float(sum(1 for x in closed if x > 120)) and val.get("На территории (на момент снимка)") == 1.0
              and val.get("ЕИ в движении (разных ЕИ с операциями)") == float(len(moved)), f"{val}; ЕИ в движении {len(moved)}")
        doc = t.B("MgrDoc", True, module="TbManager")
        odt = glob.glob(os.path.join(d, "WMS_Reports", "Отчёт_руководителю_2026-09-01_2026-09-30.odt"))
        body = zipfile.ZipFile(odt[0]).read("content.xml").decode("utf-8") if odt else ""
        R.add(c, "«Сформировать отчёт руководителю»: документ ODT и PDF в ../WMS_Reports — разделы отчёта (операции, заказы и документы, "
                 "транспорт, работа вне движения товара) и записи «Журнала работы» только этого периода; служебных данных WMS нет",
              doc.startswith("OK:") and odt and os.path.exists(odt[0][:-4] + ".pdf") and "ОТЧЁТ СКЛАДА ЗА 09.2026" in body
              and "ТРАНСПОРТ" in body and "РАБОТА, НЕ ВИДНАЯ В ДВИЖЕНИИ ТОВАРА" in body and "фура с кабелем" in body
              and "перестановка стеллажа" in body and "не этот месяц" not in body and "_SYS" not in body and "WMS-SYS" not in body, doc)
        ctl = t.B("MgrControl", float(ser("2026-09-26")), module="TbManager")
        allr = t.rows("Контроль дня", 3, 3)
        k = next((i for i, r in enumerate(allr) if str(r[0]).startswith("ЧТО ИЗМЕНИЛОСЬ")), len(allr))
        rows = [r for r in allr[:k] if r[0]]
        st = {str(r[1]): (r[0], r[2]) for r in rows}
        R.add(c, "«Контроль дня» 26.09.2026: машина на территории (сегодняшняя — внимание), «Иной приход» ждёт разбора 25 дн. — проблема, "
                 "ожидаемая дата З-104 прошла на 37 дн. (заказ ждём — напоминание, не статус), предупреждения диагностики, резервной копии "
                 "рядом нет; непроведённых строк и приходов без документов нет",
              ctl.startswith("ERR:контроль дня: проблем 1") and st.get("Машины на территории", ("",))[0] == "ВНИМАНИЕ" and "№ 8" in st["Машины на территории"][1]
              and st.get("«Иной приход»: требует разбора", ("",))[0] == "ПРОБЛЕМА" and "ЕИ-00000208" in st["«Иной приход»: требует разбора"][1]
              and st.get("Ожидаемая дата прошла", ("", ""))[0] == "ВНИМАНИЕ" and "З-104" in st["Ожидаемая дата прошла"][1]
              and "на 37 дн." in st["Ожидаемая дата прошла"][1] and "Просроченные заказы" not in st
              and st.get("Диагностика (проверки WMS_DOCTOR)", ("",))[0] == "ВНИМАНИЕ" and st.get("Резервная копия за сегодня", ("",))[0] == "ВНИМАНИЕ"
              and st.get("Непроведённые строки", ("",))[0] == "OK" and st.get("Приходы без документов", ("",))[0] == "OK", f"{ctl}; {st}")
        chg = {str(r[0]): str(r[1]) for r in allr[k + 2:] if r[0]}
        arr = sum(1 for v in cars if v["arrived"].startswith("2026-09-26"))
        dep = sum(1 for v in cars if v["departed"].startswith("2026-09-26"))
        R.add(c, "«Что изменилось сегодня»: пришло, выдано, возвращено, перемещено, списано, скорректировано, машины (приехало, уехало, на "
                 "территории), главные изменения остатков, проблемы",
              all(k in chg for k in ("Пришло", "Выдано", "Возвращено", "Перемещено", "Списано", "Скорректировано (инвентаризация)", "Машины",
                                     "Главные изменения остатков", "Проблемы"))
              and chg["Машины"] == f"приехало {arr}, уехало {dep}, на территории сейчас 1" and "пунктов «Контроля дня»" in chg["Проблемы"], f"{chg}")
        wk = t.B("MgrQuick", "Неделя", module="TbManager")
        R.add(c, "кнопки «Сегодня» / «Неделя» / «Месяц»: отчёт за текущий период по последнему снимку одним нажатием",
              wk.startswith("OK:отчёт за неделю") and t.sheet("Отчёт").getCellByPosition(1, 1).getString() == "Неделя", wk)
    finally:
        t.close()


# ================================================================ K07 LABELS

@case
def k07_labels():
    c = "K07"
    d = place("k07")
    t = Tool(d, "WMS_LABELS")
    try:
        r0 = t.B("TbLoadTable", t.B("TbLatestSnapshot", module="TbCommon"), "stock.csv", "_stock", True, module="TbCommon")
        lst = t.sheet("Этикетки")
        lst.getCellRangeByPosition(0, 1, 1, 3).setDataArray((("201", ""), ("ЕИ-00000206", 2), ("ЕИ-999999", "")))
        bad = t.B("LblBuild", module="TbLabels")
        lst.getCellRangeByPosition(0, 3, 1, 3).setDataArray((("", ""),))
        ok = t.B("LblBuild", module="TbLabels")
        lay = t.sheet("Макет")
        texts = [lay.getCellByPosition(0, r).getString() for r in range(0, 15)]
        shapes = lay.getDrawPage().getCount()
        bars = len(t.B("Code128C", "00000201", module="TbLabels"))
        R.add(c, "ЕИ, которого нет в снимке, — ошибка списка; «Построить этикетки»: 3 этикетки (ЕИ-00000201, ЕИ-00000206 ×2) по одной на "
                 "страницу: ЕИ крупно, наименование, артикул, место; штрихкод Code 128 C прямоугольниками",
              r0 > 0 and bad.startswith("ERR") and "ЕИ-00999999" in bad and ok.startswith("OK:этикеток 3") and texts[0] == "ЕИ-00000201"
              and texts[1] == "Болт М12х40" and texts[2] == "Арт. DIN933-M12" and texts[3] == "Место: A-02-1" and texts[5] == "ЕИ-00000206"
              and texts[10] == "ЕИ-00000206" and shapes > 60 and bars == 6 * 6 + 7, f"{bad[:90]}; {ok}; {texts[:6]}; фигур {shapes}; штрихов {bars}")
        g = Guard(t)
        pdf = t.B("LblPdf", module="TbLabels")
        files = glob.glob(os.path.join(d, "WMS_Labels", "labels_*.pdf"))
        pages = len(re.findall(rb"/Type\s*/Page[^s]", open(files[0], "rb").read())) if files else 0
        R.add(c, "«В PDF»: файл этикеток в ../WMS_Labels — ровно 3 страницы (по этикетке), задание № 1 записано в историю",
              g.done() and pdf.startswith("OK") and "задание № 1" in pdf and len(files) == 1 and os.path.getsize(files[0]) > 1000 and pages == 3,
              f"{pdf}; страниц {pages}")
        # M6: the book with the barcodes of several labels is saved (LibreOffice once wrote an endless file here)
        g = Guard(t)
        t0 = time.time()
        try:
            t.doc.store()
            stored = True
        except Exception as e:
            stored = str(e)[:80]
        R.add(c, "книга с построенными этикетками (штрихкоды нескольких этикеток) сохраняется за секунды — без бесконечного файла",
              g.done() and stored is True and time.time() - t0 < 30 and os.path.getsize(t.path) < 2_000_000,
              f"{stored}; {time.time() - t0:.1f} с; {os.path.getsize(t.path) if os.path.exists(t.path) else 0} байт")
        # M6: the mass selection from the snapshot (filters E2:E10 of «Этикетки»)
        st = [x for x in table("stock") if x["ei"] and x["state"] != "Приход удалён (сторно)"]
        want_a = [x["ei"] for x in st if x["place"].lower().startswith("a-") and num(x["qty"]) > 0]
        lst.getCellByPosition(4, 1).setString("A-")
        pk1 = t.B("LblPick", module="TbLabels")
        got_a = [r[0] for r in t.rows("Этикетки", 1, 1) if r[0]]
        lst.getCellByPosition(4, 1).setString("")
        first = {}
        for o in table("orders"):
            if o["ei"] and o["date_received"]:
                first[o["ei"]] = min(first.get(o["ei"], "9"), o["date_received"])
        for x in table("special"):
            if x["line_no"] and x["ei"]:
                first[x["ei"]] = min(first.get(x["ei"], "9"), x["date"])
        want_d = [x["ei"] for x in st if "2026-09-01" <= first.get(x["ei"], "0") <= "2026-09-26" and num(x["qty"]) > 0]
        lst.getCellByPosition(4, 4).setString("01.09.2026")
        lst.getCellByPosition(4, 5).setValue(ser("2026-09-26"))
        lst.getCellByPosition(4, 9).setValue(2)
        pk0 = t.B("LblPick", module="TbLabels")                      # only the registry loaded: the receipts are needed
        t.B("BtnLblRefresh", module="TbLabels")
        pk2 = t.B("LblPick", module="TbLabels")
        got_d = [(r[0], r[1]) for r in t.rows("Этикетки", 1, 1) if r[0]]
        R.add(c, "«Подобрать ЕИ»: по месту («A-…», только с остатком) и по датам получения 01.09–26.09.2026 (первое поступление ЕИ), по 2 копии — "
                 "список заменён ровно этими ЕИ",
              pk1.startswith("OK:подобрано ЕИ") and got_a == want_a and len(want_a) >= 2 and pk0.startswith("ERR:") and "Обновить из снимка" in pk0
              and pk2.startswith(f"OK:подобрано ЕИ {len(want_d)} × 2")
              and got_d == [(e, 2.0) for e in want_d] and len(want_d) >= 2, f"{pk1}; {got_a}; {pk2}; {got_d}")
        pv = t.B("LblPreview", module="TbLabels")
        lay = t.sheet("Макет")
        R.add(c, "«Предпросмотр»: одна этикетка первого ЕИ списка на листе «Макет» (лист показан), штрихкод на месте",
              pv.startswith("OK:предпросмотр " + want_d[0]) and lay.getCellByPosition(0, 0).getString() == want_d[0]
              and lay.getCellByPosition(0, 5).getString() == "" and lay.getDrawPage().getCount() == bars // 2 + 1
              and t.doc.getCurrentController().getActiveSheet().getName() == "Макет", f"{pv}; фигур {lay.getDrawPage().getCount()}")
        t.B("LblBuild", module="TbLabels")
        g = Guard(t)
        pdf2 = t.B("LblPdf", module="TbLabels")
        hist = os.path.join(d, "WMS_Labels", "history.csv")
        hrows = [ln.split(";") for ln in open(hist, encoding="utf-8").read().splitlines()[1:]] if os.path.exists(hist) else []
        lst.getCellRangeByPosition(0, 1, 1, 20).clearContents(1023)
        rp = t.B("LblReprint", 1, module="TbLabels")
        back = [(r[0], r[1]) for r in t.rows("Этикетки", 1, 1) if r[0]]
        R.add(c, "история печати (../WMS_Labels/history.csv и лист «История печати»): задания № 1 и № 2; «Повторная печать» задания № 1 — список "
                 "(ЕИ-00000201, ЕИ-00000206 ×2) возвращён и построен",
              g.done() and "задание № 2" in pdf2 and [h[0] for h in hrows] == ["1", "1", "2", "2"] and [h[2] for h in hrows[2:]] == want_d[:2]
              and rp.startswith("OK:повтор задания № 1: этикеток 3") and back == [("ЕИ-00000201", 1.0), ("ЕИ-00000206", 2.0)]
              and len([r for r in t.rows("История печати", 1, 4) if r[0]]) == 4, f"{pdf2[-40:]}; {hrows}; {rp}; {back}")
    finally:
        t.close()
    # the thermal printer adapter (TSPL) on the same snapshot
    snap = os.path.join(d, "WMS_Export", "WMS_SNAPSHOT_FIXTURE")
    prn = os.path.join(d, "labels.prn")
    r = subprocess.run([sys.executable, os.path.join(d, "WMS_TOOLBOX", "labels", "tspl_labels.py"), "--snapshot", snap, "--ei", "201,206x2", "--out", prn],
                       capture_output=True, text=True)
    body = open(prn, "rb").read().decode("cp1251") if os.path.exists(prn) else ""
    R.add(c, "labels/tspl_labels.py: команды TSPL (SIZE, GAP, CODEPAGE 1251, TEXT, BARCODE 128, PRINT) для 3 этикеток в Windows-1251",
          r.returncode == 0 and body.count("SIZE 58 mm,40 mm") == 2 and 'BARCODE 16,192,"128"' in body and "PRINT 1,2" in body and "Болт М12х40" in body,
          r.stdout.strip()[:120])

# ================================================================ K08 DOCS

@case
def k08_docs():
    c = "K08"
    d = place("k08")
    t = Tool(d, "WMS_DOCS")
    try:
        r0 = t.B("TbLoadTable", t.B("TbLatestSnapshot", module="TbCommon"), "issues.csv", "_issues", True, module="TbCommon")
        act = t.sheet("Акт")
        act.getCellByPosition(1, 1).setValue(ser("2026-09-26"))
        act.getCellByPosition(1, 2).setString("Кладовщик Смирнов")
        act.getCellByPosition(1, 3).setString("Иванов Иван Андреевич")
        act.getCellByPosition(1, 4).setValue(ser("2026-07-01"))
        act.getCellByPosition(1, 5).setValue(ser("2026-09-30"))
        pick = t.B("DocsPick", module="TbDocs")
        pos = [r for r in t.rows("Акт", 9, 5) if r[0]]
        want = [i for i in table("issues") if i["recipient"] == "Иванов Иван Андреевич" and i["status"].startswith("Проведено")]
        R.add(c, "«Подобрать по выдачам»: позиции акта — проведённые выдачи получателю «Принял» за период (сторнированная не входит)",
              r0 > 0 and pick == f"OK:позиций {len(want)}" and sorted(p[0] for p in pos) == sorted(i["ei"] for i in want), f"{pick}; {[p[0] for p in pos]}")
        doc = t.B("DocsCreate", True, module="TbDocs")
        odt = glob.glob(os.path.join(d, "WMS_Docs", "ACT-00001_2026-09-26.odt"))
        pdf = glob.glob(os.path.join(d, "WMS_Docs", "ACT-00001_2026-09-26.pdf"))
        body = zipfile.ZipFile(odt[0]).read("content.xml").decode("utf-8") if odt else ""
        hist = [r for r in t.rows("История", 1, 5) if r[0]]
        R.add(c, "«Создать документ»: акт № 1 (ODT и PDF) — заголовок, стороны, таблица позиций, подписи; строка в «Истории»; номер следующего акта 2",
              doc.startswith("OK:Акт передачи № 1 от 26.09.2026") and odt and pdf and "Иванов Иван Андреевич" in body and "Болт М12х40" in body
              and "Передал: ____________________ / Кладовщик Смирнов /" in body and len(hist) == 1 and hist[0][4] == float(len(want))
              and t.sheet("Настройки").getCellByPosition(1, 6).getString() == "2", f"{doc}; {odt}; {hist}")
        saved = not t.doc.isModified()
        # the kinds of act: a third party needs the organization, a transfer of parts — an article at every position
        act.getCellByPosition(1, 0).setString("Передача сторонней организации")
        third = t.B("DocsCreate", False, module="TbDocs")
        act.getCellByPosition(1, 0).setString("Передача деталей")
        act.getCellRangeByPosition(0, 9 + len(pos), 5, 9 + len(pos)).setDataArray((("ЕИ-00000204", "Стол офисный", "", "шт", 1, ""),))
        parts = t.B("DocsCreate", False, module="TbDocs")
        act.getCellRangeByPosition(0, 9 + len(pos), 5, 9 + len(pos)).clearContents(1023)
        act.getCellByPosition(1, 0).setString("Передача сторонней организации")
        act.getCellByPosition(1, 6).setString("ООО «Ремсервис»")
        doc2 = t.B("DocsCreate", False, module="TbDocs")
        # the number of an act already made is never reused: the counter set back to 1 — refused, nothing overwritten
        t.setting(6, "1")
        again = t.B("DocsCreate", False, module="TbDocs")
        size1 = os.path.getsize(odt[0]) if odt else 0
    finally:
        t.close()
    R.add(c, "виды актов: «Передача сторонней организации» без организации и «Передача деталей» с позицией без артикула — отказ с причиной; "
             "с организацией — акт № 2 (организация в документе)",
          third.startswith("ERR") and "Организация" in third and parts.startswith("ERR") and "нет артикула" in parts
          and doc2.startswith("OK:Передача сторонней организации № 2"), f"{third}; {parts}; {doc2}")
    t = Tool(d, "WMS_DOCS")
    try:
        hist2 = [r for r in t.rows("История", 1, 5) if r[0]]
        nxt = t.sheet("Настройки").getCellByPosition(1, 6).getString()
    finally:
        t.close()
    body2 = zipfile.ZipFile(glob.glob(os.path.join(d, "WMS_Docs", "ACT-00002_*.odt"))[0]).read("content.xml").decode("utf-8") \
        if glob.glob(os.path.join(d, "WMS_Docs", "ACT-00002_*.odt")) else ""
    R.add(c, "номер и история сохраняются вместе с книгой (после «Создать документ» книга сохранена; открыта заново — 2 акта в «Истории», "
             "следующий номер 3); номер, сброшенный на уже занятый, — отказ, файл акта № 1 не перезаписан",
          saved and len(hist2) == 2 and "ООО «Ремсервис»" in body2 and again.startswith("ERR") and "уже есть" in again
          and size1 == os.path.getsize(odt[0]) and nxt == "3", f"saved {saved}; {hist2}; {again}; next {nxt}")


# ================================================================ K09 ARCHIVE

@case
def k09_archive():
    c = "K09"
    d = place("k09")
    t = Tool(d, "WMS_ARCHIVE")
    try:
        t.setting(5, "1")                                    # keep one month: July and August may go to the archive (26.09.2026)
        an = t.B("ArcAnalyze", float(ser("2026-09-26")), module="TbArchive")
        rows = {r[0]: r[1:] for r in t.rows("Анализ", 5, 2) if r[0]}
        R.add(c, "«Анализ»: месяцы с движениями 2026-07…09, строк по месяцам, можно в архив — до 2026-08 включительно (хранить 1 месяц)",
              an.startswith("OK") and "до 2026-08" in an and sorted(rows) == ["2026-07", "2026-08", "2026-09"] and rows["2026-07"][1] == "да"
              and rows["2026-09"][1] == "нет", f"{an}; {rows}")
        pk = t.B("ArcPackage", "2026-07", "2026-07", module="TbArchive")
        arc = os.path.join(d, "WMS_Archive", "ARCHIVE_2026-07_2026-07")
        iss = open(os.path.join(arc, "issues.csv"), encoding="utf-8-sig").read().splitlines() if os.path.isdir(arc) else []
        want = [i for i in table("issues") if i["date"].startswith("2026-07")]
        again = t.B("ArcPackage", "2026-07", "2026-07", module="TbArchive")
        R.add(c, "архивный пакет 2026-07: строки движений только за июль (выдачи 1 и 4), manifest с SHA-256, проверен; повторный пакет того же "
                 "периода — отказ (не перезаписывается); рабочие файлы не изменены",
              pk.startswith("OK") and len(iss) == 1 + len(want) and os.path.exists(os.path.join(arc, "manifest.csv")) and again.startswith("ERR")
              and "не перезаписываются" in again, f"{pk}; {len(iss)}; {again[:80]}")
        op = t.B("ArcOpen", "file://" + arc + "/", module="TbArchive")
        loaded = t.doc.Sheets.hasByName("Архив issues")
        with open(os.path.join(arc, "issues.csv"), "a", encoding="utf-8") as f:
            f.write("99;;лишняя;;1;;шт;2026-07-01;;;;;;;;;;Проведено\n")
        op2 = t.B("ArcOpen", "file://" + arc + "/", module="TbArchive")
        R.add(c, "«Открыть пакет»: таблицы пакета на листах «Архив …»; изменённый после создания пакет не открывается (SHA-256)",
              op.startswith("OK") and loaded and op2.startswith("ERR") and "SHA-256" in op2, f"{op}; {op2}")
    finally:
        t.close()


# ================================================================ K10 RECONCILE

@case
def k10_reconcile():
    c = "K10"
    d = place("k10")
    names = os.path.join(d, "названия.csv")
    with open(names, "w", encoding="cp1251", newline="") as f:
        f.write("Старое название;Место;Количество\nболт м12 40;A-02-1;150\nкабель кг 3x2.5;;\nшестерня z40;D-9;\nперчатки нитрил;;\nнеизвестная вещь;;\n")
    out = os.path.join(d, "сверка.csv")
    r = subprocess.run([sys.executable, os.path.join(d, "WMS_TOOLBOX", "reconcile", "wms_reconcile.py"), names, "--export", os.path.join(d, "WMS_Export"),
                        "--out", out], capture_output=True, text=True)
    rows = list(csv.reader(io.StringIO(open(out, encoding="utf-8-sig").read()), delimiter=";"))[1:] if os.path.exists(out) else []
    first = {x[1]: (x[2], x[4], int(x[9]) if x[9] else 0) for x in rows if x[3] in ("1", "")}
    R.add(c, "сверка старых названий (файл Windows-1251): «болт м12 40» → ЕИ-00000201, «кабель кг 3x2.5» → ЕИ-00000202, «шестерня z40» → "
             "ЕИ-00000206 — уверенно; «перчатки нитрил» — два похожих ЕИ, «проверить»; «неизвестная вещь» — нет кандидатов; ничего не записано в WMS",
          r.returncode == 0 and first["болт м12 40"][:2] == ("уверенно", "ЕИ-00000201") and first["кабель кг 3x2.5"][:2] == ("уверенно", "ЕИ-00000202")
          and first["шестерня z40"][:2] == ("уверенно", "ЕИ-00000206") and first["перчатки нитрил"][0] == "проверить"
          and first["неизвестная вещь"][0] == "нет кандидатов" and not os.path.exists(os.path.join(d, "WMS_Batches")), f"{r.stdout.strip()}; {first}")
    snap = os.path.join(d, "WMS_Export", "WMS_SNAPSHOT_FIXTURE", "stock.csv")
    with open(snap, "a", encoding="utf-8") as f:
        f.write("ЕИ-00099999;подмена;;шт;1;;;;;\n")
    r2 = subprocess.run([sys.executable, os.path.join(d, "WMS_TOOLBOX", "reconcile", "wms_reconcile.py"), names, "--export", os.path.join(d, "WMS_Export")],
                        capture_output=True, text=True)
    R.add(c, "изменённый снимок (SHA-256 не совпадает) — отказ, код 2", r2.returncode == 2 and "SHA-256" in r2.stderr, r2.stderr.strip()[:120])
    # the report: up to 5 candidates, how each was found; the person marks «да»; --confirm makes the mapping
    head = open(out, encoding="utf-8-sig").readline().strip().split(";") if os.path.exists(out) else []
    ranks = [int(x[3]) for x in rows if x[3]]
    how = {(x[1], x[4]): x[10] for x in rows if x[4]}
    R.add(c, "отчёт: до 5 кандидатов на название, способ — «точное» (то же название), «артикул», «похожее»; колонка «подтвердить» для человека; "
             "колонки «группа», «в группе»",
          head[10:12] == ["способ", "подтвердить"] and head[12:] == ["группа", "в группе"] and ranks and max(ranks) <= 5 and how.get(("кабель кг 3x2.5", "ЕИ-00000202")) == "точное"
          and how.get(("болт м12 40", "ЕИ-00000201")) == "похожее", f"{head}; {how}")
    snap_ok = place("k10b")
    rep = [x[:] for x in rows]
    for x in rep:
        if (x[1], x[3]) in (("болт м12 40", "1"), ("кабель кг 3x2.5", "1")):
            x[11] = "да"
        if x[1] == "перчатки нитрил":
            x[11] = "да"                                        # both candidates confirmed — a problem
        if x[1] == "неизвестная вещь":
            x[4], x[11] = "207", "Да"                           # an EI typed by the person
    marked = os.path.join(d, "сверка_отмечено.csv")
    with open(marked, "w", encoding="cp1251", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(head)
        w.writerows(rep)
    mapping = os.path.join(d, "соответствие.csv")
    r3 = subprocess.run([sys.executable, os.path.join(d, "WMS_TOOLBOX", "reconcile", "wms_reconcile.py"), "--confirm", marked, "--export",
                         os.path.join(snap_ok, "WMS_Export"), "--out", mapping], capture_output=True, text=True)
    mp = {x[1]: x for x in list(csv.reader(io.StringIO(open(mapping, encoding="utf-8-sig").read()), delimiter=";"))[1:]} if os.path.exists(mapping) else {}
    R.add(c, "соответствие по отмеченному отчёту (Windows-1251): подтверждённые ЕИ с данными снимка; ЕИ, вписанный человеком; «без решения»; "
             "два подтверждённых ЕИ у одного названия — ошибка; код 1 (есть ошибки), ничего не записано в WMS",
          r3.returncode == 1 and mp.get("болт м12 40", [""] * 4)[2:4] == ["подтверждено", "ЕИ-00000201"]
          and mp.get("неизвестная вещь", [""] * 5)[2:5] == ["подтверждено", "ЕИ-00000207", "Дрель ударная"]
          and mp.get("шестерня z40", [""] * 3)[2] == "без решения" and mp.get("перчатки нитрил", [""] * 10)[2] == "ошибка"
          and "несколько ЕИ" in mp.get("перчатки нитрил", [""] * 10)[9] and not os.path.exists(os.path.join(d, "WMS_Batches")),
          f"{r3.stdout.strip()}; {list(mp.values())[:5]}")
    # M6: groups of the same / nearly the same old names; the dictionary of the confirmed mappings, reused next time
    d2 = place("k10c")
    names2 = os.path.join(d2, "names.csv")
    with open(names2, "w", encoding="utf-8", newline="") as f:
        f.write("Старое название;Место\nболт м12 40;A-02-1\nБОЛТ  м12 40;\nболт м12 x40;\nкабель кг 3x2.5;\nшестерня z40;D-9\n")
    rep1 = os.path.join(d2, "r1.csv")
    rc = os.path.join(d2, "WMS_TOOLBOX", "reconcile", "wms_reconcile.py")
    r4 = subprocess.run([sys.executable, rc, names2, "--export", os.path.join(d2, "WMS_Export"), "--out", rep1], capture_output=True, text=True)
    rows1 = list(csv.reader(io.StringIO(open(rep1, encoding="utf-8-sig").read()), delimiter=";")) if os.path.exists(rep1) else [[]]
    grp = {x[1]: x[12] for x in rows1[1:] if x[3] in ("1", "")}
    R.add(c, "группы: «болт м12 40» и «БОЛТ  м12 40» (то же после нормализации) — одна группа, «болт м12 x40» — «почти то же» в ней; строки "
             "группы стоят рядом; другие названия без группы",
          r4.returncode == 0 and grp.get("болт м12 40") == "Г1" and grp.get("БОЛТ  м12 40") == "Г1 (то же)" and grp.get("болт м12 x40") == "Г1 (почти то же)"
          and grp.get("кабель кг 3x2.5") == "" and list(dict.fromkeys(x[1] for x in rows1[1:]))[:3] == ["болт м12 40", "БОЛТ  м12 40", "болт м12 x40"]
          and "групп одинаковых / похожих названий 1" in r4.stdout, f"{r4.stdout.strip()}; {grp}")
    for x in rows1[1:]:
        if x[1] in ("болт м12 40", "кабель кг 3x2.5") and x[3] == "1":
            x[11] = "да"
    marked2 = os.path.join(d2, "r1_marked.csv")
    with open(marked2, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f, delimiter=";").writerows(rows1)
    r5 = subprocess.run([sys.executable, rc, "--confirm", marked2, "--export", os.path.join(d2, "WMS_Export"), "--out", os.path.join(d2, "map.csv")],
                        capture_output=True, text=True)
    mp2 = {x[1]: x for x in list(csv.reader(io.StringIO(open(os.path.join(d2, "map.csv"), encoding="utf-8-sig").read()), delimiter=";"))[1:]}
    dic = os.path.join(d2, "WMS_Reconcile", "dictionary.csv")
    drows = list(csv.reader(io.StringIO(open(dic, encoding="utf-8-sig").read()), delimiter=";"))[1:] if os.path.exists(dic) else []
    R.add(c, "--confirm: решение переходит на то же название группы («как «болт м12 40»»), но не на «почти то же»; подтверждённое записано в "
             "словарь ../WMS_Reconcile/dictionary.csv",
          r5.returncode == 0 and mp2["БОЛТ  м12 40"][2:4] == ["подтверждено", "ЕИ-00000201"] and "как «болт м12 40»" in mp2["БОЛТ  м12 40"][9]
          and mp2["болт м12 x40"][2] == "без решения" and sorted((x[0], x[2]) for x in drows) == [("болт м12 40", "ЕИ-00000201"),
                                                                                                 ("кабель кг 3х2,5", "ЕИ-00000202")]
          and "словарь: добавлено" in r5.stdout, f"{r5.stdout.strip()}; {drows}")
    rep2 = os.path.join(d2, "r2.csv")
    r6 = subprocess.run([sys.executable, rc, names2, "--export", os.path.join(d2, "WMS_Export"), "--out", rep2], capture_output=True, text=True)
    rows2 = list(csv.reader(io.StringIO(open(rep2, encoding="utf-8-sig").read()), delimiter=";")) if os.path.exists(rep2) else [[]]
    top1 = {x[1]: x for x in rows2[1:] if x[3] == "1"}
    R.add(c, "повторная сверка: названия из словаря — «подтверждено ранее», ЕИ из словаря первым, «да» уже стоит (человек может снять); "
             "нечёткое совпадение («почти то же») остаётся кандидатом без «да»",
          r6.returncode == 0 and top1["болт м12 40"][2] == "подтверждено ранее" and top1["болт м12 40"][4] == "ЕИ-00000201"
          and top1["болт м12 40"][11] == "да" and top1["болт м12 40"][10].startswith("словарь") and top1["БОЛТ  м12 40"][2] == "подтверждено ранее"
          and top1["кабель кг 3x2.5"][11] == "да" and top1["болт м12 x40"][11] == "" and "из словаря 3" in r6.stdout,
          f"{r6.stdout.strip()}; {[(k, v[2], v[11]) for k, v in top1.items()]}")


# ================================================================ K11 IMPORTER

@case
def k11_importer():
    c = "K11"
    d = place("k11")
    src = os.path.join(d, "накладная.csv")
    with open(src, "w", encoding="utf-8-sig", newline="") as f:
        f.write("Номенклатура;Артикул;Кол-во;Ед. изм.;Дата прихода;Ячейка;Группа;Тип\n"
                "Лента изоляционная;LI-19;10;шт;21.09.2026;E-01;Электрика;Офис\n"
                "Шестерня Z-40;GR-40;2;шт;2026-09-22;D-9;Детали;Детали\n"
                "Хомут 20 мм;XM-20;1,500;шт;22.09.2026;E-02;Крепёж;Офис\n"
                "Клей;KL-1;3;шт;31.13.2026;E-03;Химия;Офис\n")
    t = Tool(d, "WMS_IMPORTER")
    try:
        t.B("TestImpFile", "file://" + src, module="TbImporter")
        ld = t.B("ImpLoad", "file://" + src, module="TbImporter")
        mp = {r[0]: r[2] for r in t.rows("Сопоставление", 1, 3) if r[0]}
        chk = t.B("ImpCheck", module="TbImporter")
        st = [r[-1] for r in t.rows("Таблица", 1, 8) if r[0]]
        nob = t.B("ImpBatch", module="TbImporter")
        R.add(c, "«Загрузить таблицу»: 4 строки; сопоставление по заголовкам (Номенклатура → name, Кол-во → qty, Ед. изм. → unit, Дата прихода → "
                 "date, Ячейка → place, Группа → category, Тип → type); «Проверить»: «1,500» — неоднозначно, «31.13.2026» — не дата; пакет "
                 "с ошибками не создаётся",
              ld.startswith("OK:строк 4") and mp.get("name") == "Номенклатура" and mp.get("qty") == "Кол-во" and mp.get("unit") == "Ед. изм."
              and mp.get("date") == "Дата прихода" and mp.get("place") == "Ячейка" and mp.get("type") == "Тип" and chk.startswith("ERR")
              and st[0] == "OK" and st[1] == "OK" and "неоднозначно" in st[2] and "не дата" in st[3] and nob.startswith("ERR")
              and not glob.glob(os.path.join(d, "WMS_Batches", "*.csv")), f"{ld[:120]}; {chk}; {st}; {nob[:80]}")
        tab = t.sheet("Таблица")
        tab.getCellByPosition(2, 3).setString("1,5")
        tab.getCellByPosition(4, 4).setString("23.09.2026")
        ok = t.B("ImpBatch", module="TbImporter")
        files = glob.glob(os.path.join(d, "WMS_Batches", "SPR-*.csv"))
    finally:
        t.close()
    txt = open(files[0], encoding="utf-8").read() if files else ""
    R.add(c, "после исправления — пакет WMS-BATCH-1 «Иной приход» из 4 строк (количества и даты по правилам WMS)",
          ok.startswith("OK") and len(files) == 1 and "#target;Иной приход" in txt and "#source;WMS_IMPORTER" in txt
          and txt.count("\n") == 6 + 4 and ";1,5;" in txt, f"{ok[:100]}; {txt[:300]}")
    # the batch posted by WMS itself (a fresh test book): four special receipts, the part refills its EI
    s = Session(new_wms("k11wms"))
    try:
        s.U("TestUiAuto", 1)                                          # «Провести сейчас?» — «Да»
        s.special_input(1, B="Детали", D="Шестерня Z-40", E="GR-40", F="6", G="шт", H="03.08.2026", I="D-9")
        s.click_spc("BtnSpcPost", 1)
        s.EX("TestBatchFile", files[0])
        s.EX("BtnLoadBatch")
        msg = s.U("TestUiLastMessage")
        lines = [s.spc(r) for r in range(2, 6)]
        R.add(c, "пакет импортёра проведён WMS: 4 строки «Иного прихода» — 3 новых ЕИ и пополнение детали GR-40 (тот же ЕИ)",
              all(str(x[16]).startswith("Проведено") for x in lines) and lines[1][13] == s.spc(1)[13], f"{msg[:120]}; {[x[16] for x in lines]}")
    finally:
        s.close()
    # repeats: a row repeated in the file, a receipt WMS already has (the snapshot: «Стол офисный», Офис, 1 шт, 05.07.2026)
    d2 = place("k11b")
    src2 = os.path.join(d2, "повторы.csv")
    with open(src2, "w", encoding="utf-8-sig", newline="") as f:
        f.write("Наименование;Артикул;Количество;Единица;Дата;Место;Тип\n"
                "Стол офисный;;1;шт;05.07.2026;O-1;Офис\n"
                "Лента изоляционная;LI-19;10;шт;21.09.2026;E-01;Офис\n"
                "Лента изоляционная;LI-19;10;шт;21.09.2026;E-01;Офис\n"
                "Хомут 20 мм;XM-20;5;шт;22.09.2026;E-02;Офис\n")
    t = Tool(d2, "WMS_IMPORTER")
    try:
        t.B("ImpLoad", "file://" + src2, module="TbImporter")
        chk2 = t.B("ImpCheck", module="TbImporter")
        st2 = [r[-1] for r in t.rows("Таблица", 1, 7) if r[0]]
        pv = [r for r in t.rows("Предпросмотр", 0, 6) if r[0]]
        nob2 = t.B("ImpBatch", module="TbImporter")
        # orders: the order З-101 «Болт М12х40» is in WMS already
        t.setting(5, "Заказы")
        src3 = os.path.join(d2, "заказы.csv")
        with open(src3, "w", encoding="utf-8-sig", newline="") as f:
            f.write("№ заказа;Наименование;Заказано;Ед.;Поставщик\nЗ-101;Болт М12х40;200;шт;Метиз-Опт\nЗ-200;Гайка М12;100;шт;Метиз-Опт\n")
        t.B("ImpLoad", "file://" + src3, module="TbImporter")
        chk3 = t.B("ImpCheck", module="TbImporter")
        st3 = [r[-1] for r in t.rows("Таблица", 1, 5) if r[0]]
    finally:
        t.close()
    R.add(c, "повторы: строка, совпадающая с другой строкой файла, и поступление, которое уже есть в WMS (по снимку: «Иной приход» — тип, "
             "наименование, артикул, количество, дата; «Заказы» — № заказа и наименование) — ошибки с местом первой записи; «Предпросмотр» "
             "— ровно строки пакета; пакет с повторами не создаётся",
          chk2.startswith("ERR") and "повторов 2" in chk2 and "проверены по последнему снимку" in chk2 and st2[0].startswith("уже есть в WMS: «Иной приход», строка 2")
          and st2[1] == "OK" and st2[2].startswith("повтор строки 3 файла") and st2[3] == "OK"
          and pv[0][:3] == ["Тип прихода", "Наименование", "Артикул"] and [r[1] for r in pv[1:]] == ["Лента изоляционная", "Хомут 20 мм"]
          and nob2.startswith("ERR") and not glob.glob(os.path.join(d2, "WMS_Batches", "*.csv"))
          and chk3.startswith("ERR") and st3[0].startswith("уже есть в WMS: «Заказы», строка 2 (заказ З-101)") and st3[1] == "OK",
          f"{chk2}; {st2}; {pv}; {chk3}; {st3}")


# ================================================================ K12 BACKUP / HEALTHCHECK

@case
def k12_backup_healthcheck():
    c = "K12"
    p = new_wms("k12wms")
    s = Session(p)
    try:
        s.issue_input(1, ei="1", qty="1", date="25.09.2026", who="ива")
        s.post(1)
        s.doc.store()
    finally:
        s.close()
    wms = os.path.dirname(p)
    tb = toolbox()
    hc = os.path.join(tb, "backup", "wms_healthcheck.py")
    r = subprocess.run([sys.executable, hc, "--wms", wms], capture_output=True, text=True)
    R.add(c, "проверка перед сменой: книга, защиты, целостность архива книги и последней копии, журнал (совпадает с книгой), блокировки нет, "
             "копии есть, место, права, версии — код 0",
          r.returncode == 0 and "ИТОГ: ошибок 0, предупреждений 0" in r.stdout and "совпадает с книгой" in r.stdout
          and "листы защищены, служебные листы скрыты, структура книги защищена" in r.stdout and "контрольные суммы совпадают" in r.stdout
          and "архив ODS цел" in r.stdout, r.stdout[-400:])
    # a copy of the WMS folder: the sheet «Выдачи» saved without protection, the latest backup damaged
    w2 = os.path.join(OUT, "cases", "k12_copy")
    shutil.rmtree(w2, ignore_errors=True)
    shutil.copytree(wms, w2)
    unprotect_sheet(os.path.join(w2, os.path.basename(p)), "Выдачи")
    bk2 = sorted(glob.glob(os.path.join(w2, "WMS_Backups", "*.ods")), key=os.path.getmtime)[-1]
    data = bytearray(open(bk2, "rb").read())
    data[len(data) // 2:len(data) // 2 + 64] = b"\0" * 64
    open(bk2, "wb").write(bytes(data))
    r4 = subprocess.run([sys.executable, hc, "--wms", w2], capture_output=True, text=True)
    R.add(c, "сохранённая книга с листом «Выдачи» без защиты — предупреждение «защиты»; повреждённая последняя резервная копия — ошибка "
             "«целостность», код 2",
          r4.returncode == 2 and "лист «Выдачи» не защищён" in r4.stdout and "ПРЕДУПРЕЖДЕНИЕ" in r4.stdout and "последняя копия" in r4.stdout
          and any(ln.startswith("ОШИБКА") and "целостность" in ln for ln in r4.stdout.splitlines()), r4.stdout[-600:])
    dest = os.path.join(OUT, "cases", "k12_media")
    shutil.rmtree(dest, ignore_errors=True)
    bk = os.path.join(tb, "backup", "wms_backup.py")
    # the data of people outside the tool books (M6): the reconcile dictionary and the history of the labels
    for rel, txt in (("WMS_Reconcile/dictionary.csv", "старое название;нормализовано;ЕИ\nболт м12;болт м12;ЕИ-00000001\n"),
                     ("WMS_Labels/history.csv", "задание;дата и время;ЕИ;копий;как\n1;27.09.2026 10:00;ЕИ-00000001;2;PDF\n")):
        os.makedirs(os.path.join(wms, os.path.dirname(rel)), exist_ok=True)
        open(os.path.join(wms, rel), "w", encoding="utf-8").write(txt)
    open(os.path.join(wms, "WMS_Labels", "labels_x.pdf"), "wb").write(b"%PDF-1.4")
    b1 = subprocess.run([sys.executable, bk, "--wms", wms, "--to", dest, "--keep", "2"], capture_output=True, text=True)
    time.sleep(1.1)
    subprocess.run([sys.executable, bk, "--wms", wms, "--to", dest, "--keep", "2"], capture_output=True, text=True)
    time.sleep(1.1)
    b3 = subprocess.run([sys.executable, bk, "--wms", wms, "--to", dest, "--keep", "2"], capture_output=True, text=True)
    copies = sorted(x for x in os.listdir(dest) if x.startswith("WMS_BACKUP_"))
    v = subprocess.run([sys.executable, bk, "--wms", wms, "--to", dest, "--verify", os.path.join(dest, copies[-1])], capture_output=True, text=True)
    inside = subprocess.run([sys.executable, bk, "--wms", wms, "--to", os.path.join(wms, "x")], capture_output=True, text=True)
    R.add(c, "копия на носитель: книга, журнал, копии WMS, MANIFEST с SHA-256 (проверка «копия цела»); хранятся 2 последние; копия внутрь папки "
             "WMS — отказ",
          b1.returncode == 0 and b3.returncode == 0 and len(copies) == 2 and "удалены старые копии" in b3.stdout and v.returncode == 0
          and os.path.exists(os.path.join(dest, copies[-1], "WMS_Journal")) and inside.returncode == 2, f"{b1.stdout.strip()[:100]}; {copies}; {v.stdout.strip()}")
    R.add(c, "копия на носитель берёт и данные людей вне книг инструментов: словарь сверки и историю печати этикеток (сами PDF этикеток — нет)",
          os.path.exists(os.path.join(dest, copies[-1], "WMS_Reconcile", "dictionary.csv"))
          and os.path.exists(os.path.join(dest, copies[-1], "WMS_Labels", "history.csv"))
          and not os.path.exists(os.path.join(dest, copies[-1], "WMS_Labels", "labels_x.pdf")), str(copies[-1:]))
    # a damaged journal line and a missing backup folder are errors
    jf = glob.glob(os.path.join(wms, "WMS_Journal", "WMS_journal_*.csv"))[0]
    with open(jf, "a", encoding="utf-8") as f:
        f.write("J1;999;мусор\n")
    shutil.rmtree(os.path.join(wms, "WMS_Backups"))
    r2 = subprocess.run([sys.executable, hc, "--wms", wms], capture_output=True, text=True)
    R.add(c, "повреждённая строка журнала и нет резервных копий — ошибки, код 2",
          r2.returncode == 2 and "повреждённых строк 1" in r2.stdout and "нет резервных копий" in r2.stdout, r2.stdout[-300:])


def unprotect_sheet(book, name):
    """the ODS saved as if the protection of one sheet had been removed (content.xml rewritten, nothing opened)"""
    import re
    tmp = book + ".tmp"
    with zipfile.ZipFile(book) as zi, zipfile.ZipFile(tmp, "w") as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if it.filename == "content.xml":
                data, n = re.subn(rb'(<table:table table:name="' + name.encode() + rb'"[^>]*?) table:protected="true"', rb"\1", data, count=1)
                assert n == 1, "sheet tag not found"
            zo.writestr(it, data)
    os.replace(tmp, book)


# ================================================================ K13 LAUNCHER

@case
def k13_launcher():
    c = "K13"
    d = place("k13")
    wms_book = new_wms("k13wms")                      # a sound WMS folder for «Проверка перед сменой»
    s = Session(wms_book)
    try:
        s.doc.store()
    finally:
        s.close()
    t = Tool(d, "WMS_TOOLBOX")
    try:
        res = t.B("LnRefresh", module="TbLauncher")
        rows = [r for r in t.rows("Инструменты", 3, 3) if r[0]]
        snap = t.sheet("Инструменты").getCellByPosition(1, 0).getString()
        op = t.B("LnOpen", "WMS_SEARCH", module="TbLauncher")
        opened = [c2.getURL() for c2 in _components(t.o) if c2.getURL().endswith("WMS_SEARCH.ods")]
        miss = t.B("LnOpen", "WMS_NONE", module="TbLauncher")
        R.add(c, "launcher: 9 инструментов и 5 скриптов найдены по относительным путям, последний снимок проверен (SHA-256); «Открыть» "
                 "открывает книгу инструмента из своей папки; несуществующий — отказ",
              res.startswith("OK") and len(rows) == 14 and all(r[3] == "есть" for r in rows) and "WMS_SNAPSHOT_FIXTURE" in snap and "проверен" in snap
              and op == "OK:WMS_SEARCH" and len(opened) == 1 and miss.startswith("ERR"), f"{res}; {snap[:90]}; {op}; {miss[:60]}")
        # M6: back from a tool — «Все инструменты» brings up the launcher already open (no second copy)
        srch = [c2 for c2 in _components(t.o) if c2.getURL().endswith("WMS_SEARCH.ods")]
        back = t.o.basic(srch[0], "TbCommon", "TbOpenLauncher") if srch else ""
        lns = [c2 for c2 in _components(t.o) if c2.getURL().endswith("/WMS_TOOLBOX.ods")]
        R.add(c, "«Все инструменты» в книге инструмента: launcher своей папки выходит вперёд (уже открытый — без второй копии)",
              back == "OK:WMS_TOOLBOX" and len(lns) == 1, f"{back}; копий launcher {len(lns)}")
        for c2 in srch:
            c2.setModified(False)
            c2.close(True)
        ct = t.B("LnControl", module="TbLauncher")
        mgr = [c2 for c2 in _components(t.o) if c2.getURL().endswith("WMS_MANAGER.ods")]
        crow = [r for r in mgr[0].Sheets.getByName("Контроль дня").getCellRangeByPosition(0, 0, 3, 12).getDataArray() if r[0]] if mgr else []
        R.add(c, "«Контроль дня» из launcher: открывает «Отчёт руководителю» и сразу выполняет контроль — лист «Контроль дня» заполнен",
              "контроль дня: проблем" in ct and len(mgr) == 1 and str(crow[0][0]).startswith("КОНТРОЛЬ ДНЯ")
              and any(r[1] == "Машины на территории" for r in crow), f"{ct}; {[r[:2] for r in crow[:4]]}")
        for c2 in mgr:
            c2.setModified(False)
            c2.close(True)
        hl = t.B("LnHealth", module="TbLauncher")
        R.add(c, "«Проверка перед сменой» из launcher запускает backup/wms_healthcheck.py для папки WMS (здесь — без книги WMS: ошибка названа)",
              hl.startswith("ERR") and "ПРОВЕРКА WMS" in hl and "нет книги WMS" in hl, hl[:200])
        t.setting(5, os.path.dirname(wms_book))
        hl2 = t.B("LnHealth", module="TbLauncher")
        R.add(c, "«Проверка перед сменой» для исправной папки WMS (Настройки, B6) — «OK» (итог «ошибок 0»)",
              hl2.startswith("OK:") and "ИТОГ: ошибок 0" in hl2, hl2[-200:])
        # WMS_RECONCILE from the launcher: the names → the report (opened for the marks) → the mapping
        names = os.path.join(d, "names.csv")
        with open(names, "w", encoding="utf-8", newline="") as f:
            f.write("Старое название\nболт м12 40\nшестерня z40\n")
        rc = t.B("LnReconcile", uno.systemPathToFileUrl(names), module="TbLauncher")
        report = os.path.join(d, "names_сверка.csv")
        rows = list(csv.reader(io.StringIO(open(report, encoding="utf-8-sig").read()), delimiter=";")) if os.path.exists(report) else []
        for x in rows[1:]:
            if x[3] == "1":
                x[11] = "да"
        with open(report, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f, delimiter=";").writerows(rows)
        cf = t.B("LnConfirm", uno.systemPathToFileUrl(report), module="TbLauncher")
        mapping = os.path.join(d, "names_соответствие.csv")
        mp = list(csv.reader(io.StringIO(open(mapping, encoding="utf-8-sig").read()), delimiter=";"))[1:] if os.path.exists(mapping) else []
        R.add(c, "launcher: «Сверка названий» — отчёт names_сверка.csv рядом с файлом названий; после отметок «да» «Соответствие по сверке» — "
                 "names_соответствие.csv: оба названия подтверждены",
              rc.startswith("OK:названий 2") and len(rows) > 2 and cf.startswith("OK:") and [(x[1], x[2], x[3]) for x in mp]
              == [("болт м12 40", "подтверждено", "ЕИ-00000201"), ("шестерня z40", "подтверждено", "ЕИ-00000206")], f"{rc[:120]}; {cf[:120]}; {mp}")
    finally:
        t.close()


# ================================================================ K14 a snapshot of WMS 0.6 (contract 1.0, no vehicles)

@case
def k14_snapshot_060():
    c = "K14"
    d = place("k14", FIXTURE_060)
    t = Tool(d, "WMS_ANALYTICS")
    try:
        an = t.B("AnRefresh", float(ser("2026-09-26")), module="TbAnalytics")
        tr = t.sheet("Транспорт").getCellByPosition(0, 0).getString()
        charts = list(t.sheet("Графики").getCharts().getElementNames())
        val = {str(r[0]): r[1] for r in t.rows("Сводка", 0, 1) if r[0]}
    finally:
        t.close()
    R.add(c, "снимок WMS 0.6 в «Аналитике»: сводка строится, транспорт — «нет данных … до версии 0.7», 6 диаграмм без транспортных",
          an.startswith("OK:сводка построена") and "нет данных" in an and "до версии 0.7" in tr and len(charts) == 6
          and "cars_days" not in charts and str(val.get("Данные транспорта", "")).startswith("нет в снимке"), f"{an[-80:]}; {tr}; {charts}")
    t = Tool(d, "WMS_MANAGER")
    try:
        rep = t.sheet("Отчёт")
        rep.getCellByPosition(1, 1).setString("Месяц")
        rep.getCellByPosition(1, 2).setValue(ser("2026-09-15"))
        r0 = t.B("MgrRefresh", module="TbManager")
        r1 = t.B("MgrReport", module="TbManager")
        t.doc.calculateAll()
        val = {str(r[0]): r[1] for r in t.rows("Отчёт", 6, 1) if r[0]}
        ctl = t.B("MgrControl", float(ser("2026-09-26")), module="TbManager")
        st = {str(r[1]): r[0] for r in t.rows("Контроль дня", 3, 3) if r[0]}
    finally:
        t.close()
    R.add(c, "снимок WMS 0.6 в «Отчёте руководителю»: отчёт за месяц (выдач 3), транспорт — «нет в снимке»; «Контроль дня» работает, машины — ИНФО",
          r0.startswith("OK") and "cars нет в снимке" in r0 and r1.startswith("OK") and val.get("Выдач") == 3.0
          and str(val.get("Данные транспорта", "")).startswith("нет в снимке") and "контроль дня" in ctl and st.get("Машины на территории") == "ИНФО",
          f"{r0}; {ctl}; {st.get('Машины на территории')}")
    t = Tool(d, "WMS_SEARCH")
    try:
        r0 = t.B("SearchRefresh", module="TbSearch")
        r1 = t.B("SearchRun", "gr-40", module="TbSearch")
        r2 = t.B("SearchRun", "а123вс", module="TbSearch")
        r3 = t.B("SupplierShow", "метиз", module="TbSearch")
        r4 = t.B("VehicleShow", "а123вс", module="TbSearch")
    finally:
        t.close()
    R.add(c, "снимок WMS 0.6 в «Поиске»: поиск ЕИ работает, госномер — ничего не найдено, заказы поставщика есть, машин нет",
          r0.startswith("OK") and r1 == "OK:найдено 1" and r2 == "OK:найдено 0" and r3.startswith("OK:заказов 1, визитов 0") and r4 == "OK:визитов 0",
          f"{r0}; {r1}; {r2}; {r3}; {r4}")
    res, rep = doctor_run(d)
    R.add(c, "снимок WMS 0.6 в «Диагностике»: ошибок нет; схема WMS-SYS-3 и контракт 1.0 — OK с пометкой; транспорт — «нет данных» (OK)",
          res.startswith("OK:ошибок 0") and rep["версия схемы"][0] == "OK" and "WMS-SYS-3" in rep["версия схемы"][1]
          and rep["контракт снимка"][0] == "OK" and "1.0" in rep["контракт снимка"][1] and rep["транспорт"][0] == "OK" and "нет данных" in rep["транспорт"][1],
          f"{res}; " + "; ".join(f"{k}={v[0]}" for k, v in rep.items()))


# ================================================================ K15 WMS_TOOLBOX of release 0.6.0 on a snapshot of 0.7

OLD_COMMIT = "939929a"               # release 0.6.0 (M5)


@case
def k15_old_toolbox_new_snapshot():
    c = "K15"
    src = os.path.join(OUT, "src_060")
    if not os.path.isdir(os.path.join(src, "tools")):
        why = git_sources(OLD_COMMIT, src)
        if why:
            R.add(c, "WMS_TOOLBOX выпуска 0.6.0 на снимке 0.7", "SKIP", why)
            return
    tb6 = os.path.join(OUT, "toolbox_060")
    if not os.path.isdir(os.path.join(tb6, "WMS_TOOLBOX")):
        env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
        r = subprocess.run([sys.executable, os.path.join(src, "tools", "build_toolbox.py"), tb6], capture_output=True, text=True, timeout=1200,
                           cwd=src, env=env)
        if r.returncode != 0:
            raise RuntimeError("WMS_TOOLBOX 0.6.0 не собран: " + (r.stderr or r.stdout)[-300:])
    d = os.path.join(OUT, "cases", "k15")
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(os.path.join(d, "WMS_Export"))
    shutil.copytree(os.path.join(tb6, "WMS_TOOLBOX"), os.path.join(d, "WMS_TOOLBOX"))
    shutil.copytree(FIXTURE, os.path.join(d, "WMS_Export", "WMS_SNAPSHOT_FIXTURE"))
    res, rep_ = doctor_run(d)
    t = Tool(d, "WMS_ANALYTICS")
    try:
        an = t.B("AnRefresh", float(ser("2026-09-26")), module="TbAnalytics")
    finally:
        t.close()
    t = Tool(d, "WMS_SEARCH")
    try:
        sr = t.B("SearchRefresh", module="TbSearch")
        card = t.B("CardShow", "201", module="TbSearch")
    finally:
        t.close()
    t = Tool(d, "WMS_MANAGER")
    try:
        mr = t.B("MgrRefresh", module="TbManager")
        mp = t.B("MgrReport", module="TbManager")
    finally:
        t.close()
    t = Tool(d, "WMS_INVENTORY")
    try:
        inv = t.B("InvRefresh", False, module="TbInventory")
    finally:
        t.close()
    sch = rep_.get("версия схемы", ("", "", ""))
    R.add(c, "WMS_TOOLBOX выпуска 0.6.0 (собран из исходников его commit) на снимке 0.7 (контракт 1.1): диагностика — ошибок 0, схема "
             "WMS-SYS-4 — предупреждение «обновите WMS_TOOLBOX»; аналитика, поиск и карточка ЕИ, отчёт, инвентаризация работают (машины не видит)",
          res.startswith("OK:ошибок 0") and sch[0] == "WARN" and "WMS-SYS-4" in sch[1] and an.startswith("OK:сводка построена") and sr.startswith("OK")
          and card.startswith("OK:ЕИ-00000201") and mr.startswith("OK") and mp.startswith("OK") and inv.startswith("OK:строк пересчёта 208"),
          f"{res}; {sch[:2]}; {an[:60]}; {sr[:60]}; {card[:40]}; {mr[:40]}; {mp}; {inv[:40]}")



# ================================================================ K16 the expected date, not a status (WMS 0.7.2, D-089)

@case
def k16_expected_date():
    c = "K16"
    got = {}
    for name, status, q in (("k16a", "Ожидается", "2026-08-20"), ("k16b", "Ожидается", "2099-10-20")):
        d = place(name)
        snap = os.path.join(d, "WMS_Export", "WMS_SNAPSHOT_FIXTURE")
        p = os.path.join(snap, "orders.csv")
        txt = open(p, encoding="utf-8").read()
        old = "З-104;Герметик силиконовый;;;SIL-300;;;24;шт;;;Химпром;;;;2026-08-01;2026-08-20;;Химия;;D-01;;Просрочено;"
        assert old in txt
        open(p, "w", encoding="utf-8").write(txt.replace(old, old.replace("2026-08-20", q).replace(";Просрочено;", f";{status};")))
        rehash(snap, "orders.csv")
        t = Tool(d, "WMS_ANALYTICS")
        try:
            t.B("AnRefresh", float(ser("2026-09-26")), module="TbAnalytics")
            t.doc.calculateAll()
            val = {str(r[0]): r[1:] for r in t.rows("Сводка", 0, 7) if r[0]}
        finally:
            t.close()
        t = Tool(d, "WMS_MANAGER")
        try:
            t.B("MgrRefresh", module="TbManager")
            t.B("MgrControl", float(ser("2026-09-26")), module="TbManager")
            st = {str(r[1]): (r[0], r[2]) for r in t.rows("Контроль дня", 3, 3) if r[0]}
        finally:
            t.close()
        t = Tool(d, "WMS_SEARCH")
        try:
            t.B("SearchRefresh", module="TbSearch")
            t.B("SupplierShow", "химпром", module="TbSearch")
            head = str(t.sheet("Поставщик").getCellByPosition(0, 1).getString())
        finally:
            t.close()
        got[name] = (val.get("Ожидаемая дата прошла (заказ ждём)", [None])[0], st.get("Ожидаемая дата прошла", ("", "")), head)
    a, b = got["k16a"], got["k16b"]
    R.add(c, "снимок WMS 0.7.2: З-104 «Ожидается» (просрочки как статуса нет), ожидаемая дата 20.08.2026 прошла — аналитика: 1, «Контроль "
             "дня»: внимание «З-104 … на 37 дн.», поиск поставщика: «ожидаемая дата прошла 1»; та же позиция с датой 20.10.2099 — 0, «нет», "
             "«ожидаемая дата прошла 0»: инструменты считают по ожидаемой дате, не по статусу",
          a[0] == 1.0 and a[1][0] == "ВНИМАНИЕ" and "З-104" in a[1][1] and "на 37 дн." in a[1][1] and "ожидаемая дата прошла 1" in a[2]
          and b[0] == 0.0 and b[1] == ("OK", "нет") and "ожидаемая дата прошла 0" in b[2] and "ожидается 1" in b[2], str(got))


def _components(o):
    e = o.desktop.getComponents().createEnumeration()
    out = []
    while e.hasMoreElements():
        out.append(e.nextElement())
    return out


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
    R.save(os.path.join(OUT, "results_toolbox.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_toolbox.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
