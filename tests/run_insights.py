"""WMS_INSIGHTS (M6 PRIME §3–§4, §13) — the pattern engine on synthetic snapshots with known patterns and on a control
snapshot without them. No LibreOffice: the snapshots are written here in the format of the contract (docs/EXPORT_CONTRACT.md,
manifest with SHA-256) and the engine is run as the user runs it (python3, exit codes, files).

    WMS_TEST_OUT=/tmp/wms_ins python3 tests/run_insights.py

Planted patterns (the snapshot «patterns»): a stable consumption with little stock left (days of stock, depletion date, a
forecast with its basis and a high confidence); a sharp acceleration; a stale EI (received 200 days ago, never issued); a
dead stock; two articles issued together (a kit candidate); a regular weekly issue; an unusually large issue; a spike of
write-offs; repeated inventory shortages; a supplier that is systematically late; a supplier whose vehicles stay long; a
vehicle left «на территории» since yesterday; a reorder needed (the stock ends before a new delivery can come). The control
snapshot has the same volume of history without them: the engine must find none of the planted topics. Also: too little
history — «Недостаточно данных», no forecast; a snapshot of contract 1.0 (no vehicles) is read; a changed file is refused;
the same input gives the same output.
Results: WMS_TEST_OUT/results_insights.json.
"""
import csv
import datetime
import hashlib
import io
import os
import random
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from harness import OUT, Results  # noqa: E402

R = Results()
ENGINE = os.path.join(ROOT, "src", "toolbox", "scripts", "insights", "wms_insights.py")
FIXTURE_060 = os.path.join(HERE, "fixtures", "WMS_SNAPSHOT_FIXTURE_060")
TODAY = datetime.date(2026, 9, 26)
COLS = {
    "stock": "ei|name|article|unit|qty|place|category|state|source|source_type",
    "orders": "order_no|name|doc_no|invoice_no|article|qty_fact|qty_doc|qty_ordered|unit|price|amount|supplier|seller|date_received|date_doc|"
              "date_order|date_expected|buyer|category|assignee|place|ei|status|stock|control|note|lead_days|possible_dup",
    "special": "line_no|type|event|name|article|qty|unit|date|place|category|from_who|doc|old_mark|ei|bal_before|bal_after|status|control|note",
    "issues": "issue_no|doc_no|name|article|qty|qty_pct|unit|date|recipient|place|category|ei|returned_legacy|returned_pct_legacy|note|bal_before|"
              "bal_after|status",
    "returns": "return_no|issue_no|ei|name|article|qty|unit|date|from_who|place|category|bal_before|bal_after|status|note",
    "adjustments": "adj_no|kind|ei|name|article|unit|qty|fact|book|place_to|date|reason|place_from|bal_before|bal_after|diff|status|batch|note",
    "cars": "visit_no|date|vehicle|supplier|arrived|departed|duration_min|status|weekday|month|day_no|control|note|vehicle_key",
}


def ei(n):
    return f"ЕИ-{n:08d}"


class Snap:
    def __init__(self):
        self.t = {k: [] for k in COLS}
        self.rcv = []
        self.n_iss = self.n_ret = self.n_adj = self.n_car = 0

    def item(self, n, name, art, qty, place="A-1", cat="Разное", unit="шт"):
        self.t["stock"].append(dict(ei=ei(n), name=name, article=art, unit=unit, qty=qty, place=place, category=cat, state="Активен",
                                    source="тест", source_type="Поставщик"))

    def order(self, n, name, art, qty, supplier, received, ordered=None, expected=None, doc="УПД-1"):
        self.t["orders"].append(dict(order_no=f"З-{n}", name=name, doc_no=doc, article=art, qty_fact=qty, qty_doc=qty, qty_ordered=qty, unit="шт",
                                     supplier=supplier, date_received=received.isoformat(), date_doc=received.isoformat(),
                                     date_order=ordered.isoformat() if ordered else "", date_expected=expected.isoformat() if expected else "",
                                     ei=ei(n), status="Получено"))
        self.rcv.append((ei(n), "LIVE", "SRC"))

    def issue(self, n, d, q, who="Иванов", name="", art=""):
        self.n_iss += 1
        self.t["issues"].append(dict(issue_no=self.n_iss, name=name, article=art, qty=q, date=d.isoformat(), recipient=who, ei=ei(n), status="Проведено"))

    def adjust(self, kind, n, d, diff, book="", fact=""):
        self.n_adj += 1
        self.t["adjustments"].append(dict(adj_no=self.n_adj, kind=kind, ei=ei(n), date=d.isoformat(), diff=diff, book=book, fact=fact, status="Проведено"))

    def car(self, vehicle, key, supplier, arr, dep):
        self.n_car += 1
        self.t["cars"].append(dict(visit_no=self.n_car, date=arr.date().isoformat(), vehicle=vehicle, supplier=supplier,
                                   arrived=arr.strftime("%Y-%m-%d %H:%M:%S"), departed=dep.strftime("%Y-%m-%d %H:%M:%S") if dep else "",
                                   status="Уехал" if dep else "На территории", vehicle_key=key))

    def write(self, path, contract="1.1"):
        shutil.rmtree(path, ignore_errors=True)
        os.makedirs(os.path.join(path, "service"))
        files = []
        for name, cols in COLS.items():
            if name == "cars" and contract == "1.0":
                continue
            keys = cols.split("|")
            with open(os.path.join(path, name + ".csv"), "w", encoding="utf-8", newline="") as fh:
                w = csv.writer(fh, delimiter=";", lineterminator="\n")
                w.writerow(keys)
                for r in self.t[name]:
                    w.writerow([r.get(k, "") for k in keys])
            files.append(name + ".csv")
        with open(os.path.join(path, "service", "_RCV.csv"), "w", encoding="utf-8", newline="") as fh:
            w = csv.writer(fh, delimiter=";", lineterminator="\n")
            w.writerow(["ЕИ", "OLID", "Строка (индекс)", "Антидубль", "Состояние", "Вид"])
            for e, st, kind in self.rcv:
                w.writerow([e, 1, 1, "", st, kind])
        files.append("service/_RCV.csv")
        man = ["key;value", "format;WMS-SNAPSHOT-1", "product_version;0.7.0", f"created;{TODAY.isoformat()}T18:00:00", "last_seq;1",
               f"contract;{contract}"]
        for f in files:
            p = os.path.join(path, f)
            man.append(f"file;{f};1;{os.path.getsize(p)};{hashlib.sha256(open(p, 'rb').read()).hexdigest()}")
        open(os.path.join(path, "manifest.csv"), "w", encoding="utf-8").write("\n".join(man) + "\n")
        return path


def days_ago(k):
    return TODAY - datetime.timedelta(days=k)


def at(k, h, m=0):
    d = days_ago(k)
    return datetime.datetime(d.year, d.month, d.day, h, m)


def patterns_snapshot():
    s = Snap()
    # 1 stable consumption, little stock: 5 every 3 days for 120 days, 20 left (≈ 12 days)
    s.item(1, "Перчатки рабочие", "STAB-1", 20)
    s.order(1, "Перчатки рабочие", "STAB-1", 250, "Надёжный", days_ago(130), days_ago(140), days_ago(131))
    for k in range(120, 0, -3):
        s.issue(1, days_ago(k), 5, "Бригада 1")
    # 2 acceleration: 2 every 7 days for 60 days before, then 2 every day in the last 30 days
    s.item(2, "Лента изоляционная", "ACC-2", 500)
    s.order(2, "Лента изоляционная", "ACC-2", 600, "Надёжный", days_ago(100), days_ago(110), days_ago(101))
    for k in range(90, 30, -7):
        s.issue(2, days_ago(k), 2, "Бригада 2")
    for k in range(29, -1, -1):
        s.issue(2, days_ago(k), 2, "Бригада 3")
    # 3 stale: received 200 days ago, never issued
    s.item(3, "Насос старый", "STALE-3", 10)
    s.order(3, "Насос старый", "STALE-3", 10, "Надёжный", days_ago(200))
    # 4 dead stock: issued once 150 days ago, nothing since
    s.item(4, "Кабель редкий", "DEAD-4", 30)
    s.order(4, "Кабель редкий", "DEAD-4", 40, "Надёжный", days_ago(170))
    s.issue(4, days_ago(150), 10, "Бригада 1")
    # 5, 6 a pair issued together (the same day, the same recipient) 5 times
    s.item(5, "Болт М10", "PAIR-A", 400)
    s.item(6, "Гайка М10", "PAIR-B", 400)
    for k in (80, 60, 45, 30, 12):
        s.issue(5, days_ago(k), 10, "Монтаж")
        s.issue(6, days_ago(k), 10, "Монтаж")
    # 7 weekly: every 7 days exactly, 10 times
    s.item(7, "Фильтр воздушный", "WEEK-7", 100)
    for k in range(70, 0, -7):
        s.issue(7, days_ago(k), 1, "Сервис")
    # 8 an unusually large issue: 2 units ten times, then 30 once
    s.item(8, "Хомут", "BIG-8", 200)
    for k in range(80, 20, -6):
        s.issue(8, days_ago(k), 2, "Бригада 4")
    s.issue(8, days_ago(5), 30, "Бригада 5")
    # 9 write-offs: one a month for five months, then four in the last 30 days
    s.item(9, "Лампа", "WO-9", 100)
    for k in (170, 140, 110, 80, 50):
        s.adjust("Списание", 9, days_ago(k), -1)
    for k in (25, 18, 10, 3):
        s.adjust("Списание", 9, days_ago(k), -2)
    # 10 repeated inventory shortages
    s.item(10, "Краска", "INV-10", 15)
    s.adjust("Инвентаризация", 10, days_ago(60), -3, 20, 17)
    s.adjust("Инвентаризация", 10, days_ago(20), -2, 17, 15)
    # 11 a late supplier: five deliveries 5–7 days after the expected date
    for i, late in enumerate((5, 6, 7, 5, 6)):
        n = 20 + i
        s.item(n, f"Труба {i}", f"LATE-{i}", 50)
        s.order(n, f"Труба {i}", f"LATE-{i}", 50, "Опоздун", days_ago(100 - 15 * i), days_ago(120 - 15 * i), days_ago(100 - 15 * i + late))
    # 12 reorder: fast consumption, a supplier with a long lead time (30 days), 18 days of stock
    s.item(12, "Картридж", "REORD-12", 18)
    s.order(12, "Картридж", "REORD-12", 100, "Далёкий", days_ago(95), days_ago(125), days_ago(96))
    for other in range(3):
        s.item(40 + other, "Картридж старый", f"R-{other}", 0)
        s.order(40 + other, "Картридж старый", f"R-{other}", 10, "Далёкий", days_ago(200 - 30 * other), days_ago(230 - 30 * other), days_ago(201 - 30 * other))
    for k in range(89, -1, -5):
        s.issue(12, days_ago(k), 5, "Офис")
    # vehicles: a supplier whose vehicles stay long, a vehicle left since yesterday, ordinary visits
    for i in range(6):
        s.car("ГАЗ Долгий", f"А{100 + i}АА", "Долгий", at(20 - i, 9), at(20 - i, 12, 10 + i * 5))
    for i in range(20):
        s.car("КАМАЗ", f"В{200 + i}ВВ", "Быстрый", at(25 - i, 10, 15), at(25 - i, 10, 45))
    s.car("Газель забытая", "Е777ЕЕ", "Быстрый", at(1, 16), None)
    return s


def control_snapshot():
    """the same volume of history, no planted pattern: irregular but steady consumption, everything moved lately, no pair more
    than twice, suppliers on time, short vehicle visits"""
    rnd = random.Random(7)
    s = Snap()
    for n in range(1, 13):
        s.item(n, f"Позиция {n}", f"CTRL-{n}", 1000)
        s.order(n, f"Позиция {n}", f"CTRL-{n}", 1500, "Точный", days_ago(150 + n), days_ago(160 + n), days_ago(150 + n))
        k = 120
        while k > 0:
            s.issue(n, days_ago(k), rnd.choice((3, 4, 5, 6)), f"Получатель {rnd.randint(1, 30)}")
            k -= rnd.choice((6, 9, 13, 17, 4))
    for i in range(20):
        s.car("КАМАЗ", f"В{300 + i}ВВ", "Точный", at(30 - i, 9), at(30 - i, 9, 40))
    return s


def transport_snapshot(planted):
    """eight weeks of vehicles. planted: Monday–Thursday 3 vehicles a day, Friday 6; most arrive 09:00–10:59; the vehicles of
    «Медленный» stay about 55 minutes, the others 30–40. Control: 4 vehicles every working day, arrivals spread 08:00–17:59,
    every supplier alike"""
    rnd = random.Random(11)
    s = Snap()
    start = TODAY - datetime.timedelta(days=TODAY.weekday() + 56)          # a Monday eight weeks ago
    nslow = 0
    for k in range(56):
        d = start + datetime.timedelta(days=k)
        wd = d.weekday()
        if wd >= 5:
            continue
        n = (6 if wd == 4 else 3) if planted else 4
        for j in range(n):
            if planted:
                h = rnd.choice((9, 9, 10, 10, 9, 10, 13, 15))
                slow = nslow < 12 and j == 0 and wd in (1, 3)
                sup = "Медленный" if slow else rnd.choice(("Обычный", "Другой", "Третий"))
                dur = rnd.randint(52, 60) if slow else rnd.randint(30, 40)
                nslow += slow
            else:
                h = 8 + (k * 4 + j * 3) % 10
                sup = ("Обычный", "Другой", "Третий", "Медленный")[j]
                dur = rnd.randint(30, 40)
            arr = datetime.datetime(d.year, d.month, d.day, h, rnd.randint(0, 50))
            s.car("Газель", f"Т{k:02d}{j}ТТ", sup, arr, arr + datetime.timedelta(minutes=dur))
    return s


def run_engine(snap, out, *extra):
    return subprocess.run([sys.executable, ENGINE, "--snapshot", snap, "--out", out, "--today", TODAY.isoformat()] + list(extra),
                          capture_output=True, text=True, timeout=300)


def read(out, name):
    p = os.path.join(out, name)
    return list(csv.DictReader(open(p, encoding="utf-8"), delimiter=";")) if os.path.exists(p) else []


def find(ins, topic, needle=""):
    return [x for x in ins if x["topic"] == topic and needle in (x["object"] + " " + x["finding"])]


PLANTED = ["Запас на исходе", "Расход ускорился", "Залежалый ЕИ", "Без движения", "Кандидат в комплект", "Периодичность", "Необычно крупная выдача",
           "Всплеск: списания", "Повторяющиеся расхождения инвентаризации", "Поставщик задерживает", "Долгие стоянки поставщика",
           "Машина с прошлых дней", "Пора заказывать"]


def main():
    base = os.path.join(OUT, "insights")
    os.makedirs(base, exist_ok=True)
    c = "I01"
    snap = patterns_snapshot().write(os.path.join(base, "WMS_SNAPSHOT_PATTERNS"))
    out = os.path.join(base, "out_patterns")
    r = run_engine(snap, out)
    ins = read(out, "insights.csv")
    cons = {x["ei"]: x for x in read(out, "consumption.csv")}
    R.add(c, "движок отработал на снимке с заложенными закономерностями", r.returncode == 0 and ins, r.stdout.strip()[:200] + r.stderr[-200:])
    checks = [
        ("стабильный расход, мало остатка: «Запас на исходе» с датой исчерпания, основанием и уверенностью «высокая»",
         find(ins, "Запас на исходе", "ЕИ-00000001"), lambda h: h[0]["confidence"] == "высокая" and "выдач" in h[0]["basis"] and "закончится" not in h[0]["finding"]),
        ("резкий всплеск расхода: «Расход ускорился»", find(ins, "Расход ускорился", "ЕИ-00000002"), None),
        ("залежалый ЕИ (получен 200 дн. назад, ни разу не выдан)", find(ins, "Залежалый ЕИ", "ЕИ-00000003"), None),
        ("товар без движения 150 дн.", find(ins, "Без движения", "ЕИ-00000004"), None),
        ("пара, выдаваемая вместе, — кандидат в комплект", find(ins, "Кандидат в комплект", "PAIR-A + PAIR-B"), None),
        ("еженедельная выдача — периодичность с датой следующей", find(ins, "Периодичность", "WEEK-7"), lambda h: "еженедельно" in h[0]["finding"]),
        ("аномально крупная выдача относительно своей истории", find(ins, "Необычно крупная выдача", "ЕИ-00000008"), None),
        ("всплеск списаний", find(ins, "Всплеск: списания"), None),
        ("регулярные недостачи инвентаризации — «Критично»", find(ins, "Повторяющиеся расхождения инвентаризации", "ЕИ-00000010"),
         lambda h: h[0]["level"] == "Критично"),
        ("систематически опаздывающий поставщик", find(ins, "Поставщик задерживает", "Опоздун"), None),
        ("долгие стоянки машин поставщика", find(ins, "Долгие стоянки поставщика", "Долгий"), None),
        ("машина с прошлых дней на территории — «Критично»", find(ins, "Машина с прошлых дней", "Газель забытая"), lambda h: h[0]["level"] == "Критично"),
        ("запас закончится раньше новой поставки — «Пора заказывать»", find(ins, "Пора заказывать", "ЕИ-00000012"), lambda h: h[0]["level"] == "Критично"),
    ]
    for name, hits, extra in checks:
        R.add(c, "найдено: " + name, bool(hits) and (extra is None or extra(hits)), "; ".join(f"{h['level']}: {h['object']} — {h['finding'][:90]}" for h in hits[:2]))
    one = cons.get("ЕИ-00000001", {})
    R.add(c, "прогноз расхода ЕИ-1: дней запаса ≈ 12, дата исчерпания, основание (период, выдач, средний расход, медиана)",
          one.get("days_of_stock") in ("11", "12", "13") and one.get("depletion") and "медиана" in one.get("basis", ""), f"{one}")
    R.add(c, "у каждой закономерности есть уровень (Критично / Внимание / Наблюдение), основание и уверенность",
          all(x["level"] in ("Критично", "Внимание", "Наблюдение") and x["basis"] and x["confidence"] for x in ins), f"{len(ins)} закономерностей")
    tr = {x["what"]: x["value"] for x in read(out, "transport_summary.csv")}
    sup = {x["supplier"]: x for x in read(out, "transport_suppliers.csv")}
    R.add(c, "аналитика машин: визиты, средняя/медиана/максимум стоянки, по поставщикам (доля долгих), пиковый час, повторные машины",
          tr.get("визитов всего") == "27" and tr.get("на территории сейчас") == "1" and sup.get("Долгий", {}).get("long_share") == "100%"
          and sup.get("Быстрый", {}).get("median_min") == "30.0" and read(out, "transport_hours.csv") and read(out, "transport_days.csv"), f"{tr}")
    ctl = {x["check"]: x for x in read(out, "control.csv")}
    R.add(c, "«Контроль дня»: машина с прошлых дней — ПРОБЛЕМА; остальные проверки выполнены", ctl.get("Машины на территории", {}).get("status") == "ПРОБЛЕМА"
          and len(ctl) >= 8, "; ".join(f"{k}: {v['status']}" for k, v in ctl.items()))
    ch = {x["what"]: x["detail"] for x in read(out, "changes.csv")}
    R.add(c, "«Что изменилось сегодня»: выдачи за день (ЕИ-2 и ЕИ-7 не сегодня), машины, главные изменения остатков",
          ch.get("Выдачи", "").startswith("1,") and "ЕИ-00000002" in ch.get("Главные изменения остатков", ""), f"{ch}")
    # ------------------------------------------------------------ control snapshot
    c = "I02"
    snap2 = control_snapshot().write(os.path.join(base, "WMS_SNAPSHOT_CONTROL"))
    out2 = os.path.join(base, "out_control")
    r2 = run_engine(snap2, out2)
    ins2 = read(out2, "insights.csv")
    false = [x for x in ins2 if x["topic"] in PLANTED]
    R.add(c, "контрольный снимок без закономерностей: ни одной из 13 заложенных тем не найдено (движок не выдумывает)", r2.returncode == 0 and not false,
          "; ".join(f"{x['topic']}: {x['object']}" for x in false[:5]) or f"прочих наблюдений {len(ins2)}")
    # ------------------------------------------------------------ too little history, determinism, contract 1.0, tampering
    c = "I03"
    s = Snap()
    s.item(1, "Редкая позиция", "RARE-1", 5)
    s.order(1, "Редкая позиция", "RARE-1", 7, "Надёжный", days_ago(20))
    s.issue(1, days_ago(10), 1)
    s.issue(1, days_ago(3), 1)
    snap3 = s.write(os.path.join(base, "WMS_SNAPSHOT_SHORT"))
    out3 = os.path.join(base, "out_short")
    run_engine(snap3, out3)
    row = {x["ei"]: x for x in read(out3, "consumption.csv")}.get("ЕИ-00000001", {})
    R.add(c, "мало истории (2 выдачи): «Недостаточно данных», прогноза нет, закономерностей о запасе нет",
          row.get("confidence") == "недостаточно данных" and row.get("days_of_stock") == "" and row.get("basis", "").startswith("Недостаточно данных")
          and not [x for x in read(out3, "insights.csv") if x["topic"] in ("Запас на исходе", "Скоро закончится", "Пора заказывать")], f"{row}")
    out4 = os.path.join(base, "out_patterns_again")
    run_engine(snap, out4)
    same = all(open(os.path.join(out, f), "rb").read() == open(os.path.join(out4, f), "rb").read()
               for f in ("insights.csv", "consumption.csv", "control.csv", "changes.csv", "transport_suppliers.csv"))
    R.add(c, "детерминированность: тот же снимок и та же дата — те же файлы", same, "")
    if os.path.isdir(FIXTURE_060):
        out5 = os.path.join(base, "out_060")
        r5 = run_engine(FIXTURE_060, out5)
        R.add(c, "снимок контракта 1.0 (WMS 0.6, без машин): движок работает, транспорт — «нет данных»",
              r5.returncode == 0 and not os.path.exists(os.path.join(out5, "transport_summary.csv"))
              and "нет данных транспорта" in open(os.path.join(out5, "control.csv"), encoding="utf-8").read(), r5.stdout.strip()[:160])
    bad = os.path.join(base, "WMS_SNAPSHOT_TAMPERED")
    shutil.copytree(snap, bad, dirs_exist_ok=True)
    with open(os.path.join(bad, "issues.csv"), "a", encoding="utf-8") as fh:
        fh.write("9999;;;;1;;;2026-09-26;x;;;ЕИ-00000001;;;;;;Проведено\n")
    r6 = run_engine(bad, os.path.join(base, "out_bad"))
    R.add(c, "изменённый после экспорта снимок не используется (SHA-256): код 2", r6.returncode == 2 and "SHA-256" in r6.stderr, r6.stderr.strip()[:160])
    # ------------------------------------------------------------ vehicles: relative observations only with enough visits
    c = "I04"
    snap7 = transport_snapshot(True).write(os.path.join(base, "WMS_SNAPSHOT_CARS"))
    out7 = os.path.join(base, "out_cars")
    r7 = run_engine(snap7, out7)
    ins7 = read(out7, "insights.csv")
    slow = find(ins7, "Поставщик дольше обычного", "Медленный")
    R.add(c, "поставщик, чьи машины стоят дольше: «медианное время машины за последние N визитов на X% выше общей медианы» (без порога долгих)",
          r7.returncode == 0 and len(slow) == 1 and "за последние 12 визитов на" in slow[0]["finding"] and "выше общей медианы" in slow[0]["finding"]
          and not find(ins7, "Поставщик дольше обычного", "Обычный") and not find(ins7, "Долгие стоянки поставщика"),
          "; ".join(f"{h['topic']}: {h['object']} — {h['finding']}" for h in ins7 if h["topic"] in ("Поставщик дольше обычного", "Долгие стоянки поставщика")))
    peak = find(ins7, "Пик приездов")
    fri = find(ins7, "День недели")
    R.add(c, "пик приездов 09:00–11:00 и «по пятницам машин больше обычного» — с основанием (недель, визитов)",
          [h["object"] for h in peak] == ["09:00–11:00"] and [h["object"] for h in fri] == ["по пятницам"] and "недель" in fri[0]["basis"],
          "; ".join(f"{h['topic']}: {h['finding']} ({h['basis']})" for h in peak + fri))
    snap8 = transport_snapshot(False).write(os.path.join(base, "WMS_SNAPSHOT_CARS_CONTROL"))
    out8 = os.path.join(base, "out_cars_control")
    r8 = run_engine(snap8, out8)
    ins8 = [x for x in read(out8, "insights.csv") if x["topic"] in ("Поставщик дольше обычного", "Пик приездов", "День недели", "Долгие стоянки поставщика")]
    R.add(c, "контрольный журнал машин (равномерно по дням и часам, поставщики одинаковы): этих наблюдений нет",
          r8.returncode == 0 and not ins8, "; ".join(f"{x['topic']}: {x['object']}" for x in ins8))
    few = Snap()
    for i in range(12):
        few.car("КАМАЗ", f"К{i}", "Медленный" if i < 4 else "Обычный", at(12 - i, 9, 5), at(12 - i, 10, 30 if i < 4 else 0))
    r9 = run_engine(few.write(os.path.join(base, "WMS_SNAPSHOT_CARS_FEW")), os.path.join(base, "out_cars_few"))
    ins9 = [x for x in read(os.path.join(base, "out_cars_few"), "insights.csv") if x["topic"] in ("Поставщик дольше обычного", "Пик приездов", "День недели")]
    R.add(c, "мало визитов (12): сравнений поставщиков, пиков и дней недели нет — только при достаточном числе наблюдений",
          r9.returncode == 0 and not ins9, "; ".join(f"{x['topic']}: {x['object']}" for x in ins9))
    R.save(os.path.join(OUT, "results_insights.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']}; результаты: {OUT}")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
