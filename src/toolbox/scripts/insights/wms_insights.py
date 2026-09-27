#!/usr/bin/env python3
"""WMS_INSIGHTS — the pattern engine of WMS (M6 PRIME §3–§6): it looks for patterns in the history of the warehouse by a
snapshot of WMS («Экспорт для инструментов», docs/EXPORT_CONTRACT.md) and never touches the working book.

    python3 wms_insights.py (--snapshot SNAPSHOT_DIR | --export WMS_EXPORT_DIR) [--out DIR] [--today ГГГГ-ММ-ДД]
                            [--wms WMS_DIR] [--long-min 120]

Deterministic: the same snapshot and the same --today give the same result (default today — the date of the snapshot).
The snapshot is checked first (manifest.csv: the format, every file with its SHA-256); a snapshot of contract 1.0 (no
vehicles) is read as one without transport data. Written into DIR (default: <WMS_DIR>/WMS_Reports/INSIGHTS_<date-time>):
  insights.csv       the found patterns: level (Критично / Внимание / Наблюдение), topic, object, finding, basis (period,
                     observations, mean / median), confidence, advice — a forecast is never made without enough history
                     («Недостаточно данных» in consumption.csv, no finding);
  consumption.csv    every EI with a balance or movements: issued in 7 / 30 / 90 days, the rate per day, the days of stock,
                     the date of depletion if the observed consumption continues, the basis and the confidence;
  transport_*.csv    the vehicles: summary, by day / week / month, by supplier, by hour of arrival, repeated vehicles;
  control.csv        «Контроль дня»: vehicles on the territory, unposted rows, «Иной приход» to identify, receipts without
                     documents, overdue orders, the backup of today, the snapshot, the problems of the data;
  changes.csv        «Что изменилось сегодня»: receipts, issues, returns, moves, write-offs, corrections, vehicles, the
                     largest changes of balances, the problems;
  INSIGHTS_REPORT.md all of it for a person.
Exit code: 0 done, 2 the snapshot could not be read or is not intact.
"""
import argparse
import csv
import datetime
import hashlib
import io
import math
import os
import re
import statistics
import sys
from collections import Counter, defaultdict

LEVELS = ("Критично", "Внимание", "Наблюдение")
ST_POSTED = "Проведено"
ST_STORNO = "Удалено (сторно)"
MIN_OBS = 3                 # issue events in 90 days below this: no forecast («Недостаточно данных»)
MIN_SPAN_DAYS = 14          # the history of an EI must cover at least this many days for a forecast
DEAD_DAYS = 90              # a balance without any movement this long: dead stock
STALE_DAYS = 180            # received this long ago and never issued: a stale EI
CAR_MIN_ALL = 20            # vehicles: closed visits for the overall median, visits for the peak hours
CAR_MIN_SUP = 5             # closed visits of a supplier before it is compared
SUP_LAST = 20               # a supplier is judged by its last visits
CAR_MIN_WEEKDAY = 30        # visits (and four weeks) before the weekdays are compared
WEEKDAYS_PL = ("по понедельникам", "по вторникам", "по средам", "по четвергам", "по пятницам", "по субботам", "по воскресеньям")


def read_text(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"{path}: неизвестная кодировка")


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def latest_snapshot(export_dir):
    snaps = sorted(d for d in os.listdir(export_dir) if d.startswith("WMS_SNAPSHOT_") and os.path.exists(os.path.join(export_dir, d, "manifest.csv")))
    if not snaps:
        raise ValueError(f"в {export_dir} нет снимков WMS — в WMS: «Главная» → «Экспорт для инструментов»")
    return os.path.join(export_dir, snaps[-1])


def manifest(snap):
    man, files = {}, {}
    for row in csv.reader(io.StringIO(read_text(os.path.join(snap, "manifest.csv"))), delimiter=";"):
        if not row:
            continue
        if row[0] == "file" and len(row) >= 5:
            files[row[1]] = row[4]
        elif len(row) >= 2 and row[0] not in ("sheet", "key"):
            man[row[0]] = row[1]
    return man, files


def verify(snap):
    man, files = manifest(snap)
    if man.get("format") != "WMS-SNAPSHOT-1":
        raise ValueError(f"формат снимка «{man.get('format')}», ожидается WMS-SNAPSHOT-1")
    for f, h in files.items():
        p = os.path.join(snap, f)
        if not os.path.exists(p):
            raise ValueError(f"в снимке нет файла {f}")
        if sha256(p) != h:
            raise ValueError(f"{f} снимка изменён после экспорта (SHA-256 не совпадает)")
    return man


def table(snap, name):
    p = os.path.join(snap, name)
    if not os.path.exists(p):
        return None
    return list(csv.DictReader(io.StringIO(read_text(p)), delimiter=";"))


def num(s):
    try:
        return float(str(s).replace(",", "."))
    except (TypeError, ValueError):
        return None


def day(s):
    try:
        return datetime.date.fromisoformat(str(s)[:10])
    except ValueError:
        return None


def dtm(s):
    try:
        return datetime.datetime.strptime(str(s), "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def fmt_q(x):
    if x is None:
        return ""
    r = round(x, 3)
    return (f"{r:.3f}".rstrip("0").rstrip(".") if r != int(r) else str(int(r))).replace(".", ",")


def fmt_d(d):
    return d.strftime("%d.%m.%Y") if d else ""


def unesc(s):
    return re.sub(r"%([0-9A-Fa-f]{2})", lambda m: chr(int(m.group(1), 16)), s)


# ================================================================ the history of the warehouse

class Wh:
    """the snapshot as movements: every receipt, issue, return, correction of every EI, the orders, the vehicles"""

    def __init__(self, snap, today, long_min):
        self.snap = snap
        self.today = today
        self.long_min = long_min
        self.stock = {}
        for r in table(snap, "stock.csv") or []:
            if r.get("ei"):
                self.stock[r["ei"]] = dict(name=r.get("name", ""), article=r.get("article", ""), unit=r.get("unit", ""), qty=num(r.get("qty")) or 0.0,
                                           place=r.get("place", ""), category=r.get("category", ""), state=r.get("state", ""),
                                           stype=r.get("source_type", ""))
        rcv_state = {}
        for r in table(snap, "service/_RCV.csv") or []:
            v = list(r.values())
            if v and str(v[0]).startswith("ЕИ-"):
                rcv_state[v[0]] = v[4] if len(v) > 4 else ""
        self.receipts = []           # (ei, date, qty, supplier, doc, kind)
        self.orders = table(snap, "orders.csv") or []
        for r in self.orders:
            ei = r.get("ei", "")
            if ei.startswith("ЕИ-") and rcv_state.get(ei, "LIVE") != "STORNO" and "КОПИЯ" not in (r.get("control") or "")[:5]:
                self.receipts.append((ei, day(r.get("date_received")), num(r.get("qty_fact")) or 0.0, r.get("supplier", ""), r.get("doc_no", ""), "заказ"))
        self.special = table(snap, "special.csv") or []
        for r in self.special:
            if r.get("line_no") and r.get("ei") and ST_STORNO not in (r.get("status") or "") and "КОПИЯ" not in (r.get("control") or "")[:5]:
                self.receipts.append((r["ei"], day(r.get("date")), num(r.get("qty")) or 0.0, r.get("type", ""), r.get("doc", ""), "иной приход"))
        self.issues = []             # (no, ei, date, qty, recipient, doc)
        self.issue_rows = table(snap, "issues.csv") or []
        for r in self.issue_rows:
            if r.get("issue_no") and (r.get("status") or "").startswith(ST_POSTED):
                q = num(r.get("qty"))
                if q:
                    self.issues.append((r["issue_no"], r.get("ei", ""), day(r.get("date")), q, r.get("recipient", ""), r.get("doc_no", "")))
        self.returns = []
        self.return_rows = table(snap, "returns.csv") or []
        for r in self.return_rows:
            if r.get("return_no") and (r.get("status") or "") != ST_STORNO and "КОПИЯ" not in (r.get("status") or "")[:5]:
                q = num(r.get("qty"))
                if q:
                    self.returns.append((r["return_no"], r.get("ei", ""), day(r.get("date")), q, r.get("from_who", "")))
        self.adjust = []             # (no, kind, ei, date, diff, reason, book, fact)
        self.adjust_rows = table(snap, "adjustments.csv") or []
        for r in self.adjust_rows:
            if r.get("adj_no") and (r.get("status") or "") != ST_STORNO and "КОПИЯ" not in (r.get("status") or "")[:5]:
                self.adjust.append((r["adj_no"], r.get("kind", ""), r.get("ei", ""), day(r.get("date")), num(r.get("diff")) or 0.0, r.get("reason", ""),
                                    num(r.get("book")), num(r.get("fact"))))
        self.cars_rows = table(snap, "cars.csv")      # None: a snapshot of contract 1.0
        self.visits = []
        for r in self.cars_rows or []:
            if not r.get("visit_no"):
                continue
            arr, dep = dtm(r.get("arrived")), dtm(r.get("departed"))
            dur = (dep - arr).total_seconds() / 60 if arr and dep else None
            self.visits.append(dict(no=r["visit_no"], vehicle=r.get("vehicle", ""), key=r.get("vehicle_key", ""), supplier=r.get("supplier", ""),
                                    arr=arr, dep=dep, dur=dur, status=r.get("status", "")))
        self.fixes = Counter()       # EI → corrections and storno of its operations (the journal)
        self.fix_kinds = defaultdict(Counter)
        for f in sorted(os.listdir(os.path.join(snap, "journal"))) if os.path.isdir(os.path.join(snap, "journal")) else []:
            for ln in read_text(os.path.join(snap, "journal", f)).splitlines():
                a = ln.split(";")
                if len(a) < 6 or a[0] != "J1":
                    continue
                typ = a[4]
                if not (typ.endswith("_FIX") or typ.endswith("_DEL")):
                    continue
                d = day(a[2])
                if d is None or (self.today - d).days > 90:
                    continue
                ei = next((unesc(x[3:]) for x in a[5:] if x.startswith("EI=")), "")
                if ei:
                    self.fixes[ei] += 1
                    self.fix_kinds[ei][typ] += 1

    def name(self, ei):
        s = self.stock.get(ei)
        return f"{ei} «{s['name']}»" if s else ei


# ================================================================ statistics

def median(xs):
    return statistics.median(xs) if xs else None


def confidence(n):
    return "высокая" if n >= 12 else ("средняя" if n >= 6 else ("низкая" if n >= MIN_OBS else "недостаточно данных"))


def period(today, days):
    return f"{fmt_d(today - datetime.timedelta(days=days - 1))}–{fmt_d(today)}"


def consumption(w):
    """per EI: issued in 7/30/90 days, rates, days of stock, depletion — with the basis and the confidence"""
    t = w.today
    by_ei = defaultdict(list)
    for no, ei, d, q, who, doc in w.issues:
        if d and d <= t:
            by_ei[ei].append((d, q))
    ret = defaultdict(float)
    for no, ei, d, q, who in w.returns:
        if d and d <= t and (t - d).days < 90:
            ret[ei] += q
    first_move = {}
    for ei, d, q, sup, doc, kind in w.receipts:
        if d and (ei not in first_move or d < first_move[ei]):
            first_move[ei] = d
    rows = []
    for ei in sorted(set(by_ei) | {e for e, s in w.stock.items() if s["qty"] > 0}):
        ev = by_ei.get(ei, [])
        q7 = sum(q for d, q in ev if (t - d).days < 7)
        q30 = sum(q for d, q in ev if (t - d).days < 30)
        q90 = sum(q for d, q in ev if (t - d).days < 90)
        n90 = sum(1 for d, q in ev if (t - d).days < 90)
        n30 = sum(1 for d, q in ev if (t - d).days < 30)
        start = min([d for d, q in ev] + ([first_move[ei]] if ei in first_move else []), default=None)
        span = min(90, (t - start).days + 1) if start else 0
        net90 = max(q90 - ret.get(ei, 0.0), 0.0)
        bal = w.stock.get(ei, {}).get("qty", 0.0)
        r = dict(ei=ei, q7=q7, q30=q30, q90=q90, n30=n30, n90=n90, span=span, bal=bal, rate30=None, rate90=None, days=None, depletion=None,
                 med=median([q for d, q in ev if (t - d).days < 90]), conf=confidence(n90), basis="")
        if n90 >= MIN_OBS and span >= MIN_SPAN_DAYS and net90 > 0:
            r["rate90"] = net90 / span
            r["rate30"] = q30 / min(30, span)
            rate = r["rate30"] if r["rate30"] and n30 >= 2 else r["rate90"]
            r["days"] = bal / rate if rate > 0 else None
            if r["days"] is not None:
                r["depletion"] = t + datetime.timedelta(days=int(r["days"]))
            r["basis"] = (f"{span} дн. ({period(t, span)}): выдач {n90}, выдано {fmt_q(q90)}" + (f", возвращено {fmt_q(ret[ei])}" if ret.get(ei) else "")
                          + f"; средний расход {fmt_q(r['rate90'])} в день, за 30 дн. {fmt_q(r['rate30'])} в день; медиана выдачи {fmt_q(r['med'])}")
        else:
            r["conf"] = "недостаточно данных"
            r["basis"] = f"Недостаточно данных: выдач за 90 дн. {n90} (нужно ≥ {MIN_OBS}), история {span} дн. (нужно ≥ {MIN_SPAN_DAYS})"
        rows.append(r)
    return rows


class Finder:
    def __init__(self, w):
        self.w = w
        self.found = []

    def add(self, level, topic, obj, finding, basis, conf, advice, ei="", value=""):
        self.found.append(dict(level=level, topic=topic, object=obj, finding=finding, basis=basis, confidence=conf, advice=advice, ei=ei, value=value))


def patterns(w, cons, f):
    t = w.today
    # ------------------------------------------------------------ forecasts: days of stock, depletion, reorder
    lead = supplier_lead(w)
    sup_of = {}
    for ei, d, q, sup, doc, kind in w.receipts:
        if kind == "заказ":
            sup_of[ei] = sup
    art_lead = {}
    for ei, s in w.stock.items():
        if ei in sup_of and sup_of[ei] in lead and lead[sup_of[ei]]["median"] is not None:
            art_lead[ei] = lead[sup_of[ei]]
    for r in cons:
        if r["days"] is None:
            continue
        ei = r["ei"]
        lt = art_lead.get(ei)
        if lt is not None and r["days"] <= lt["median"]:
            f.add("Критично", "Пора заказывать", w.name(ei), f"запаса на {r['days']:.0f} дн., а поставщик «{sup_of[ei]}» привозит в среднем за "
                  f"{lt['median']:.0f} дн. — при таком расходе позиция закончится раньше, чем придёт новая партия",
                  r["basis"] + f"; срок поставки: медиана {lt['median']:.0f} дн. по {lt['n']} поставкам", r["conf"],
                  "оформить заказ сейчас", ei, f"{r['days']:.1f}")
        elif r["days"] <= 7:
            f.add("Критично", "Скоро закончится", w.name(ei), f"остаток {fmt_q(r['bal'])} — примерно на {r['days']:.0f} дн., закончится около "
                  f"{fmt_d(r['depletion'])}", r["basis"], r["conf"], "проверить заказ и остаток", ei, f"{r['days']:.1f}")
        elif r["days"] <= 21:
            f.add("Внимание", "Запас на исходе", w.name(ei), f"остаток {fmt_q(r['bal'])} — на {r['days']:.0f} дн., ориентировочно до {fmt_d(r['depletion'])}",
                  r["basis"], r["conf"], "запланировать заказ", ei, f"{r['days']:.1f}")
        # acceleration / deceleration: the last 30 days against the 90-day average
        if r["rate90"] and r["n30"] >= 3 and r["n90"] >= 6:
            k = r["rate30"] / r["rate90"]
            if k >= 1.8:
                f.add("Внимание", "Расход ускорился", w.name(ei), f"за 30 дн. расход {fmt_q(r['rate30'])} в день — в {k:.1f} раза выше среднего за 90 дн.",
                      r["basis"], r["conf"], "проверить причину и запас", ei, f"{k:.2f}")
    for r in cons:
        if r["rate90"] and r["n90"] >= 6 and r["q30"] == 0 and r["span"] >= 60:
            f.add("Наблюдение", "Расход остановился", w.name(r["ei"]), "раньше выдавалась регулярно, за последние 30 дн. выдач нет", r["basis"],
                  r["conf"], "уточнить, нужна ли ещё позиция", r["ei"])
    # ------------------------------------------------------------ no movement, stale, fast movers, placement
    last_move = {}
    for ei, d, q, sup, doc, kind in w.receipts:
        if d:
            last_move[ei] = max(d, last_move.get(ei, d))
    for coll in (w.issues, w.returns):
        for x in coll:
            if x[2]:
                last_move[x[1]] = max(x[2], last_move.get(x[1], x[2]))
    for x in w.adjust:
        if x[3]:
            last_move[x[2]] = max(x[3], last_move.get(x[2], x[3]))
    issued = {x[1] for x in w.issues}
    first_rcv = {}
    for ei, d, q, sup, doc, kind in w.receipts:
        if d:
            first_rcv[ei] = min(d, first_rcv.get(ei, d))
    dead, stale = [], []
    for ei, s in w.stock.items():
        if s["qty"] <= 0 or ei not in last_move:
            continue
        idle = (t - last_move[ei]).days
        if ei in first_rcv and ei not in issued and (t - first_rcv[ei]).days >= STALE_DAYS:
            stale.append((idle, ei))
        elif idle >= DEAD_DAYS:
            dead.append((idle, ei))
    for idle, ei in sorted(stale, reverse=True):
        f.add("Внимание", "Залежалый ЕИ", w.name(ei), f"получен {fmt_d(first_rcv[ei])}, ни разу не выдавался, остаток {fmt_q(w.stock[ei]['qty'])} "
              f"{w.stock[ei]['unit']}", f"с прихода {(t - first_rcv[ei]).days} дн.; последнее движение {fmt_d(last_move[ei])}", "высокая",
              "решить: нужен ли, перемещение/списание", ei, str(idle))
    for idle, ei in sorted(dead, reverse=True):
        f.add("Наблюдение", "Без движения", w.name(ei), f"остаток {fmt_q(w.stock[ei]['qty'])} {w.stock[ei]['unit']} без движения {idle} дн.",
              f"последнее движение {fmt_d(last_move[ei])}", "высокая", "проверить востребованность", ei, str(idle))
    ops90 = Counter(x[1] for x in w.issues if x[2] and (t - x[2]).days < 90)
    # a fast mover stands out: at least twice as many issues as a typical issued EI (not simply the top of a short list)
    typical = median(list(ops90.values())) or 0
    top = [ei for ei, n in ops90.most_common(10) if n >= 5 and n >= 2 * typical]
    for rank, ei in enumerate(top, start=1):
        s = w.stock.get(ei, {})
        f.add("Наблюдение", "Быстроходная позиция", w.name(ei), f"№ {rank} по частоте: {ops90[ei]} выдач за 90 дн.; место «{s.get('place', '')}»",
              f"90 дн. ({period(t, 90)}): выдач {ops90[ei]}", confidence(ops90[ei]),
              "держать ближе к зоне выдачи" if rank <= 5 else "следить за запасом", ei, str(ops90[ei]))
    # ------------------------------------------------------------ anomalies of the own history of an EI
    hist = defaultdict(list)
    for no, ei, d, q, who, doc in w.issues:
        hist[ei].append((d, q, no, who))
    for ei, ev in hist.items():
        if len(ev) < 6:
            continue
        ev.sort(key=lambda x: (x[0] or datetime.date.min, x[2]))
        for i, (d, q, no, who) in enumerate(ev):
            others = [x[1] for j, x in enumerate(ev) if j != i]
            med = statistics.median(others)
            mad = statistics.median([abs(x - med) for x in others]) or (med * 0.1 or 1.0)
            if d and (t - d).days < 90 and q >= 3 * med and (q - med) / mad >= 5:
                f.add("Внимание", "Необычно крупная выдача", w.name(ei), f"выдача № {no} {fmt_d(d)}: {fmt_q(q)} — в {q / med:.1f} раза больше обычной "
                      f"(медиана {fmt_q(med)}) — получатель «{who}»", f"выдач этого ЕИ {len(ev)}, медиана {fmt_q(med)}", confidence(len(ev)),
                      "проверить документ выдачи", ei, fmt_q(q))
    # ------------------------------------------------------------ spikes of write-offs and returns
    spikes(w, f, "Списания", [(x[3], abs(x[4])) for x in w.adjust if x[1] == "Списание"])
    spikes(w, f, "Возвраты", [(x[2], x[3]) for x in w.returns])
    # ------------------------------------------------------------ corrections and storno, inventory differences
    for ei, n in w.fixes.most_common():
        if n >= 3:
            kinds = ", ".join(f"{k} ×{v}" for k, v in sorted(w.fix_kinds[ei].items()))
            f.add("Внимание", "Частые исправления", w.name(ei), f"{n} исправлений и сторно операций за 90 дн. ({kinds})",
                  "журнал WMS в снимке", confidence(n), "разобрать причину ошибок ввода", ei, str(n))
    inv = defaultdict(list)
    for no, kind, ei, d, diff, reason, book, fact in w.adjust:
        if kind == "Инвентаризация" and diff:
            inv[ei].append((d, diff))
    for ei, ev in inv.items():
        if len(ev) >= 2:
            neg = [x for x in ev if x[1] < 0]
            lvl = "Критично" if len(neg) >= 2 and len(neg) == len(ev) else "Внимание"
            f.add(lvl, "Повторяющиеся расхождения инвентаризации", w.name(ei), f"{len(ev)} расхождения: " + ", ".join(f"{fmt_d(d)} {'+' if x > 0 else ''}{fmt_q(x)}"
                  for d, x in sorted(ev, key=lambda z: z[0] or datetime.date.min)) + ("; все — недостача" if lvl == "Критично" else ""),
                  f"инвентаризаций с разницей: {len(ev)}", confidence(len(ev) + 3), "проверить учёт выдач и хранение", ei, str(len(ev)))
    # ------------------------------------------------------------ recipients, pairs, kits, periodicity
    rec = Counter()
    for no, ei, d, q, who, doc in w.issues:
        if d and (t - d).days < 90 and who:
            rec[(who, w.stock.get(ei, {}).get("article") or ei)] += 1
    for (who, art), n in rec.most_common(10):
        if n >= 4:
            f.add("Наблюдение", "Регулярная выдача одному получателю", f"{who} ← {art}", f"{n} выдач за 90 дн.", f"90 дн. ({period(t, 90)})",
                  confidence(n), "рассмотреть норму выдачи или подготовку заранее", "", str(n))
    pairs(w, f)
    periodic(w, f)
    # ------------------------------------------------------------ suppliers
    for sup, lt in sorted(lead.items()):
        if lt["n_delay"] >= 3 and lt["delay_median"] >= 3:
            f.add("Внимание", "Поставщик задерживает", sup, f"фактический приход позже ожидаемого: медиана +{lt['delay_median']:.0f} дн., максимум "
                  f"+{lt['delay_max']:.0f} дн.", f"поставок с ожидаемой датой: {lt['n_delay']}", confidence(lt["n_delay"] + 3),
                  "закладывать запас по сроку или обсудить с поставщиком", "", f"{lt['delay_median']:.0f}")
    # ------------------------------------------------------------ vehicles
    transport_findings(w, f)


def spikes(w, f, what, events):
    """the last 30 days against the months before (at least 3 months of history)"""
    t = w.today
    ev = [(d, q) for d, q in events if d and d <= t]
    if not ev:
        return
    first = min(d for d, q in ev)
    months = (t - first).days // 30
    if months < 3:
        return
    last = [q for d, q in ev if (t - d).days < 30]
    before = [q for d, q in ev if 30 <= (t - d).days < 30 * (months + 1)]
    base = len(before) / months
    if len(last) >= 3 and len(last) >= 2 * max(base, 1):
        f.add("Внимание", f"Всплеск: {what.lower()}", what, f"за 30 дн. {len(last)} ({fmt_q(sum(last))}) — обычно {base:.1f} в месяц",
              f"история {months} мес., событий раньше {len(before)}", confidence(len(ev)), "разобрать причины", "", str(len(last)))


def pairs(w, f):
    """items issued together (the same recipient and day): frequent pairs and kit candidates"""
    t = w.today
    baskets = defaultdict(set)
    for no, ei, d, q, who, doc in w.issues:
        if d and (t - d).days < 180:
            baskets[(d, who)].add(w.stock.get(ei, {}).get("article") or ei)
    cnt, single = Counter(), Counter()
    for items in baskets.values():
        for a in items:
            single[a] += 1
        items = sorted(items)
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                cnt[(items[i], items[j])] += 1
    for (a, b), n in cnt.most_common():
        if n < 3:
            break
        conf = n / min(single[a], single[b])
        kit = conf >= 0.6
        f.add("Наблюдение", "Кандидат в комплект" if kit else "Часто выдаются вместе", f"{a} + {b}", f"вместе {n} раз (в один день одному получателю); "
              f"доля совместных выдач {conf:.0%}", f"180 дн.: выдач {a} {single[a]}, {b} {single[b]}", confidence(n + (3 if kit else 0)),
              "собирать комплектом" if kit else "хранить рядом", "", str(n))


def periodic(w, f):
    """a regular consumption: the intervals between the issues of an article nearly equal (weekly / monthly)"""
    t = w.today
    days = defaultdict(set)
    for no, ei, d, q, who, doc in w.issues:
        if d:
            days[w.stock.get(ei, {}).get("article") or ei].add(d)
    for art, ds in days.items():
        ds = sorted(ds)
        if len(ds) < 6:
            continue
        gaps = [(b - a).days for a, b in zip(ds, ds[1:])]
        m = statistics.mean(gaps)
        if m < 3:
            continue
        cv = statistics.pstdev(gaps) / m
        if cv <= 0.25:
            kind = "еженедельно" if 5 <= m <= 9 else ("ежемесячно" if 25 <= m <= 35 else f"раз в {m:.0f} дн.")
            nxt = ds[-1] + datetime.timedelta(days=round(m))
            f.add("Наблюдение", "Периодичность", art, f"выдаётся {kind}; следующая ожидается около {fmt_d(nxt)}",
                  f"выдач {len(ds)}, интервал {m:.1f} дн. (разброс {cv:.0%})", confidence(len(ds)), "готовить заранее", "", f"{m:.1f}")


def supplier_lead(w):
    """per supplier: the delay of the receipts against the expected date and the lead time (order → receipt)"""
    out = defaultdict(lambda: dict(delays=[], leads=[]))
    for r in w.orders:
        sup = r.get("supplier", "")
        rec, exp, ordd = day(r.get("date_received")), day(r.get("date_expected")), day(r.get("date_order"))
        if not sup or not rec or not r.get("ei") or not r.get("qty_ordered"):
            continue
        if exp:
            out[sup]["delays"].append((rec - exp).days)
        if ordd:
            out[sup]["leads"].append((rec - ordd).days)
    res = {}
    for sup, v in out.items():
        res[sup] = dict(n_delay=len(v["delays"]), delay_median=median(v["delays"]) or 0, delay_max=max(v["delays"], default=0),
                        n=len(v["leads"]), median=median(v["leads"]) if len(v["leads"]) >= 3 else None)
    return {k: v for k, v in res.items() if v["median"] is not None or v["n_delay"]}


# ================================================================ vehicles

def transport(w):
    """the tables of the vehicles: summary, by day / week / month, by supplier, by hour, repeated vehicles"""
    t = w.today
    vs = [v for v in w.visits if v["status"] != "Отменён" and v["arr"]]
    closed = [v for v in vs if v["dur"] is not None]
    durs = [v["dur"] for v in closed]
    summary = [("визитов всего", len(vs)), ("на территории сейчас", sum(1 for v in vs if v["status"] == "На территории")),
               ("приехало сегодня", sum(1 for v in vs if v["arr"].date() == t)),
               ("уехало сегодня", sum(1 for v in closed if v["dep"].date() == t)),
               ("средняя стоянка, мин", round(statistics.mean(durs), 1) if durs else ""),
               ("медиана стоянки, мин", round(statistics.median(durs), 1) if durs else ""),
               ("самая долгая стоянка, мин", round(max(durs), 1) if durs else ""),
               (f"доля стоянок дольше {w.long_min} мин", f"{sum(1 for d in durs if d > w.long_min) / len(durs):.0%}" if durs else ""),
               ("разных машин", len({v["key"] for v in vs})), ("разных поставщиков", len({v["supplier"] for v in vs}))]
    by_day = []
    for k in range(29, -1, -1):
        d = t - datetime.timedelta(days=k)
        ds = [v for v in vs if v["arr"].date() == d]
        dd = [v["dur"] for v in ds if v["dur"] is not None]
        by_day.append((fmt_d(d), len(ds), round(statistics.mean(dd), 1) if dd else ""))
    by_week = []
    monday = t - datetime.timedelta(days=t.weekday())
    for k in range(11, -1, -1):
        a = monday - datetime.timedelta(weeks=k)
        ds = [v for v in vs if a <= v["arr"].date() < a + datetime.timedelta(days=7)]
        by_week.append((f"{fmt_d(a)}", len(ds)))
    by_month = []
    for k in range(11, -1, -1):
        y, m = t.year, t.month - k
        while m < 1:
            y, m = y - 1, m + 12
        ds = [v for v in vs if v["arr"].year == y and v["arr"].month == m]
        by_month.append((f"{m:02d}.{y}", len(ds)))
    sup = defaultdict(list)
    for v in vs:
        sup[v["supplier"]].append(v)
    by_sup = []
    for s, lst in sorted(sup.items(), key=lambda x: -len(x[1])):
        dd = [v["dur"] for v in lst if v["dur"] is not None]
        by_sup.append((s, len(lst), round(statistics.mean(dd), 1) if dd else "", round(statistics.median(dd), 1) if dd else "",
                       f"{sum(1 for d in dd if d > w.long_min) / len(dd):.0%}" if dd else ""))
    hours = Counter(v["arr"].hour for v in vs)
    by_hour = [(f"{h:02d}:00", hours.get(h, 0)) for h in range(24)]
    veh = defaultdict(list)
    for v in vs:
        veh[v["key"]].append(v)
    repeated = [(lst[-1]["vehicle"], k, len(lst), fmt_d(lst[-1]["arr"].date()), ", ".join(sorted({v["supplier"] for v in lst})))
                for k, lst in sorted(veh.items(), key=lambda x: -len(x[1])) if len(lst) >= 2]
    return dict(summary=summary, days=by_day, weeks=by_week, months=by_month, suppliers=by_sup, hours=by_hour, vehicles=repeated)


def transport_findings(w, f):
    t = w.today
    for v in w.visits:
        if v["status"] == "На территории" and v["arr"] and v["arr"].date() < t:
            f.add("Критично", "Машина с прошлых дней", f"визит № {v['no']} «{v['vehicle']}»", f"числится на территории с {v['arr']:%d.%m %H:%M} — "
                  "вероятно, забыли нажать УЕХАЛ", "лист «Приход авто»", "высокая", "отметить выезд или исправить время", "", v["no"])
    # the closed visits in the order of arrival: a supplier is compared by its last SUP_LAST visits with the median of all
    vs = sorted((v for v in w.visits if v["status"] == "Уехал" and v["dur"] is not None and v["arr"]), key=lambda v: (v["arr"], int(v["no"] or 0)))
    overall = statistics.median([v["dur"] for v in vs]) if len(vs) >= CAR_MIN_ALL else None
    sup = defaultdict(list)
    for v in vs:
        sup[v["supplier"]].append(v["dur"])
    for s, dd in sorted(sup.items()):
        if len(dd) < CAR_MIN_SUP:
            continue
        last = dd[-SUP_LAST:]
        share = sum(1 for d in last if d > w.long_min) / len(last)
        med = statistics.median(last)
        basis = f"последних закрытых визитов поставщика {len(last)}" + (f"; общая медиана {overall:.0f} мин по {len(vs)} визитам" if overall else "")
        if share >= 0.3:
            f.add("Внимание", "Долгие стоянки поставщика", s, f"{share:.0%} визитов дольше {w.long_min} мин; медиана {med:.0f} мин"
                  + (f" при общей {overall:.0f} мин" if overall else ""), basis, confidence(len(last)), "разобрать разгрузку/документы этого поставщика",
                  "", f"{share:.2f}")
        elif overall and len(sup) >= 2 and med >= overall * 1.3 and med - overall >= 10:
            f.add("Наблюдение", "Поставщик дольше обычного", s, f"медианное время машины за последние {len(last)} визитов на "
                  f"{(med / overall - 1) * 100:.0f}% выше общей медианы ({med:.0f} мин против {overall:.0f})", basis, confidence(len(last)),
                  "уточнить, что задерживает машины этого поставщика (документы, разгрузка, ожидание)", "", f"{med / overall:.2f}")
    allv = [v for v in w.visits if v["status"] != "Отменён" and v["arr"]]
    if len(allv) >= CAR_MIN_ALL:
        # the busiest two hours against an even spread over a working day of 10 hours (20 % per two hours)
        hours = Counter(v["arr"].hour for v in allv)
        n, _, h = max((hours.get(h, 0) + hours.get(h + 1, 0), hours.get(h, 0), -h) for h in range(23))
        h = -h
        if n >= 10 and n / len(allv) >= 0.4:
            f.add("Наблюдение", "Пик приездов", f"{h:02d}:00–{h + 2:02d}:00", f"пик приездов {h:02d}:00–{h + 2:02d}:00: {n} из {len(allv)} машин "
                  f"({n / len(allv):.0%})", f"визитов {len(allv)}; по часам приезда", confidence(len(allv)),
                  "планировать людей и погрузчик на это время", "", str(n))
    # the weekdays: the mean number of vehicles of a weekday against the other working days (whole weeks of history only)
    days = sorted({v["arr"].date() for v in allv})
    if len(allv) >= CAR_MIN_WEEKDAY and days and (days[-1] - days[0]).days >= 27:
        first, lastd = days[0], days[-1]
        per = Counter(v["arr"].date() for v in allv)
        means = {}
        for wd in range(7):
            dates = [first + datetime.timedelta(days=k) for k in range((lastd - first).days + 1)]
            dates = [d for d in dates if d.weekday() == wd]
            cnt = [per.get(d, 0) for d in dates]
            if dates and sum(cnt):
                means[wd] = (statistics.mean(cnt), sum(cnt), len(dates))
        for wd, (m, tot, nd) in sorted(means.items()):
            others = [x[0] for k, x in means.items() if k != wd]
            if len(others) >= 3 and tot >= 8 and m >= statistics.mean(others) * 1.4:
                f.add("Наблюдение", "День недели", WEEKDAYS_PL[wd], f"{WEEKDAYS_PL[wd]} машин больше обычного: в среднем {m:.1f} против "
                      f"{statistics.mean(others):.1f} в другие дни", f"{nd} недель ({fmt_d(first)}–{fmt_d(lastd)}), визитов {len(allv)}",
                      confidence(nd), "учесть при планировании смен", "", f"{m:.2f}")
    veh = Counter(v["key"] for v in allv)
    for k, n in veh.most_common(5):
        if n >= 4:
            last = max((v for v in allv if v["key"] == k), key=lambda v: v["arr"])
            f.add("Наблюдение", "Частая машина", f"«{last['vehicle']}»", f"{n} визитов, последний {last['arr']:%d.%m.%Y}", f"визитов всего {len(allv)}",
                  confidence(n), "", "", str(n))


# ================================================================ «Контроль дня», «Что изменилось сегодня»

def control(w, man, wms_dir):
    t = w.today
    out = []

    def add(st, what, detail):
        out.append((st, what, detail))
    op = [v for v in w.visits if v["status"] == "На территории"]
    old = [v for v in op if v["arr"] and v["arr"].date() < t]
    if w.cars_rows is None:
        add("ИНФО", "Машины на территории", "в снимке нет данных транспорта (WMS до версии 0.7)")
    else:
        add("ПРОБЛЕМА" if old else ("ВНИМАНИЕ" if op else "OK"), "Машины на территории", (f"{len(op)}: " + "; ".join(f"«{v['vehicle']}» с {v['arr']:%d.%m %H:%M}"
            for v in op[:10]) if op else "нет") + (f" — с прошлых дней: {len(old)}" if old else ""))
    unposted = []
    for r in w.issue_rows:
        if not r.get("issue_no") and (r.get("ei") or r.get("qty")):
            unposted.append("выдачи")
    for r in w.orders:
        if not r.get("ei") and r.get("qty_fact"):
            unposted.append("приходы")
    for r in w.special:
        if not r.get("line_no") and r.get("qty"):
            unposted.append("иной приход")
    for r in w.return_rows:
        if not r.get("return_no") and r.get("qty"):
            unposted.append("возвраты")
    for r in w.adjust_rows:
        if not r.get("adj_no") and r.get("kind"):
            unposted.append("корректировки")
    c = Counter(unposted)
    add("ВНИМАНИЕ" if c else "OK", "Непроведённые строки", ", ".join(f"{k}: {v}" for k, v in c.items()) or "нет")
    rev = [ei for ei, s in w.stock.items() if s["state"] == "Требует разбора"]
    add("ВНИМАНИЕ" if rev else "OK", "Неразобранный «Иной приход»", f"{len(rev)}: " + ", ".join(rev[:10]) if rev else "нет")
    nodoc = [r for r in w.orders if r.get("ei") and (r.get("status") == "Получено без документов" or not (r.get("doc_no") and r.get("date_doc")))]
    add("ВНИМАНИЕ" if nodoc else "OK", "Приходы без документов", f"{len(nodoc)}: " + ", ".join(r["ei"] for r in nodoc[:10]) if nodoc else "нет")
    over = [r for r in w.orders if r.get("status") in ("Просрочено", "Частично получено / просрочено")]
    add("ВНИМАНИЕ" if over else "OK", "Просроченные заказы", f"{len(over)}: " + "; ".join(f"{r.get('order_no')} {r.get('name')}" for r in over[:8]) if over else "нет")
    neg = [ei for ei, s in w.stock.items() if s["qty"] < 0]
    copies = sum(1 for rows, col in ((w.issue_rows, "status"), (w.orders, "control"), (w.special, "control"), (w.return_rows, "status"),
                                     (w.adjust_rows, "status")) for r in rows if (r.get(col) or "").startswith("КОПИЯ"))
    probs = []
    if neg:
        probs.append(f"отрицательные остатки: {len(neg)}")
    if copies:
        probs.append(f"строк КОПИЯ: {copies}")
    add("ПРОБЛЕМА" if probs else "OK", "Проблемы данных (как WMS_DOCTOR)", "; ".join(probs) or "нет")
    created = man.get("created", "")[:10]
    add("OK" if created == t.isoformat() else "ВНИМАНИЕ", "Снимок", f"создан {man.get('created', '')}, целостность проверена (SHA-256)"
        + ("" if created == t.isoformat() else " — снимок не сегодняшний: сделайте «Экспорт для инструментов»"))
    if wms_dir and os.path.isdir(os.path.join(wms_dir, "WMS_Backups")):
        dates = [m.group(1) for f in os.listdir(os.path.join(wms_dir, "WMS_Backups")) for m in [re.search(r"_(\d{4}-\d{2}-\d{2})_\d{6}_", f)] if m]
        latest = max(dates, default="")
        add("OK" if latest == t.isoformat() else "ВНИМАНИЕ", "Резервная копия за сегодня", f"последняя копия {latest or 'нет'}"
            + ("" if latest == t.isoformat() else " — сделайте «Резервная копия» на «Главной»"))
    else:
        add("ИНФО", "Резервная копия за сегодня", "папка WMS_Backups не найдена рядом (укажите --wms)")
    return out


def changes(w, cons):
    t = w.today
    rc = [x for x in w.receipts if x[1] == t]
    iss = [x for x in w.issues if x[2] == t]
    ret = [x for x in w.returns if x[2] == t]
    adj = [x for x in w.adjust if x[3] == t]
    delta = defaultdict(float)
    for ei, d, q, *_ in rc:
        delta[ei] += q
    for no, ei, d, q, *_ in iss:
        delta[ei] -= q
    for no, ei, d, q, *_ in ret:
        delta[ei] += q
    for no, kind, ei, d, diff, *_ in adj:
        delta[ei] += diff
    arrived = [v for v in w.visits if v["arr"] and v["arr"].date() == t and v["status"] != "Отменён"]
    left = [v for v in w.visits if v["dep"] and v["dep"].date() == t and v["status"] != "Отменён"]
    out = [("Приходы", f"ЕИ: {len({x[0] for x in rc})}, позиций заказов: {sum(1 for x in rc if x[5] == 'заказ')}, количество {fmt_q(sum(x[2] for x in rc))}"),
           ("Выдачи", f"{len(iss)}, количество {fmt_q(sum(x[3] for x in iss))}, получателей {len({x[4] for x in iss})}"),
           ("Возвраты", f"{len(ret)}, количество {fmt_q(sum(x[3] for x in ret))}"),
           ("Перемещения", str(sum(1 for x in adj if x[1] == "Перемещение"))),
           ("Списания", f"{sum(1 for x in adj if x[1] == 'Списание')}, количество {fmt_q(sum(-x[4] for x in adj if x[1] == 'Списание'))}"),
           ("Инвентаризационные корректировки", str(sum(1 for x in adj if x[1] == "Инвентаризация"))),
           ("Машины", f"приехало {len(arrived)}, уехало {len(left)}" if w.cars_rows is not None else "нет данных транспорта в снимке")]
    big = sorted(((abs(v), ei, v) for ei, v in delta.items() if v), reverse=True)[:5]
    out.append(("Главные изменения остатков", "; ".join(f"{w.name(ei)} {'+' if v > 0 else ''}{fmt_q(v)}" for a, ei, v in big) or "нет"))
    return out


# ================================================================ output

def write_csv(path, head, rows):
    with open(path, "w", encoding="utf-8", newline="") as fh:
        wr = csv.writer(fh, delimiter=";", lineterminator="\n")
        wr.writerow(head)
        for r in rows:
            wr.writerow(r)


def run(snap, out, today=None, wms_dir=None, long_min=120):
    man = verify(snap)
    if today is None:
        today = day(man.get("created", "")) or datetime.date.today()
    w = Wh(snap, today, long_min)
    cons = consumption(w)
    f = Finder(w)
    patterns(w, cons, f)
    order = {lv: i for i, lv in enumerate(LEVELS)}
    found = sorted(f.found, key=lambda x: (order[x["level"]], x["topic"], x["object"]))
    os.makedirs(out, exist_ok=True)
    write_csv(os.path.join(out, "insights.csv"), ["level", "topic", "object", "finding", "basis", "confidence", "advice", "ei", "value"],
              [[x[k] for k in ("level", "topic", "object", "finding", "basis", "confidence", "advice", "ei", "value")] for x in found])
    write_csv(os.path.join(out, "consumption.csv"), ["ei", "name", "article", "balance", "issued_7", "issued_30", "issued_90", "issues_90", "rate_30",
                                                     "rate_90", "days_of_stock", "depletion", "confidence", "basis"],
              [[r["ei"], w.stock.get(r["ei"], {}).get("name", ""), w.stock.get(r["ei"], {}).get("article", ""), fmt_q(r["bal"]), fmt_q(r["q7"]),
                fmt_q(r["q30"]), fmt_q(r["q90"]), r["n90"], fmt_q(r["rate30"]), fmt_q(r["rate90"]),
                "" if r["days"] is None else f"{r['days']:.0f}", fmt_d(r["depletion"]), r["conf"], r["basis"]] for r in cons])
    tr = transport(w) if w.cars_rows is not None else None
    if tr:
        write_csv(os.path.join(out, "transport_summary.csv"), ["what", "value"], tr["summary"])
        write_csv(os.path.join(out, "transport_days.csv"), ["day", "visits", "avg_min"], tr["days"])
        write_csv(os.path.join(out, "transport_weeks.csv"), ["week_from", "visits"], tr["weeks"])
        write_csv(os.path.join(out, "transport_months.csv"), ["month", "visits"], tr["months"])
        write_csv(os.path.join(out, "transport_suppliers.csv"), ["supplier", "visits", "avg_min", "median_min", "long_share"], tr["suppliers"])
        write_csv(os.path.join(out, "transport_hours.csv"), ["hour", "arrivals"], tr["hours"])
        write_csv(os.path.join(out, "transport_vehicles.csv"), ["vehicle", "key", "visits", "last", "suppliers"], tr["vehicles"])
    ctl = control(w, man, wms_dir)
    write_csv(os.path.join(out, "control.csv"), ["status", "check", "detail"], ctl)
    ch = changes(w, cons)
    write_csv(os.path.join(out, "changes.csv"), ["what", "detail"], ch)
    counts = Counter(x["level"] for x in found)
    md = [f"# Закономерности склада — снимок {os.path.basename(snap.rstrip('/'))}", "",
          f"Дата анализа: {fmt_d(today)}; снимок создан {man.get('created', '')}, LAST_SEQ {man.get('last_seq', '')}, контракт {man.get('contract', '1.0')}.", "",
          "## Найденные закономерности", "", " · ".join(f"{lv}: {counts.get(lv, 0)}" for lv in LEVELS), ""]
    for lv in LEVELS:
        items = [x for x in found if x["level"] == lv]
        if not items:
            continue
        md.append(f"### {lv}")
        md.append("")
        for x in items:
            md.append(f"- **{x['topic']}** — {x['object']}: {x['finding']}. _Основание:_ {x['basis']}; уверенность: {x['confidence']}."
                      + (f" Что сделать: {x['advice']}." if x["advice"] else ""))
        md.append("")
    md += ["## Контроль дня", ""] + [f"- {s} — {c}: {d}" for s, c, d in ctl] + ["", "## Что изменилось сегодня", ""] + [f"- {a}: {b}" for a, b in ch] + [""]
    if tr:
        md += ["## Транспорт", ""] + [f"- {a}: {b}" for a, b in tr["summary"]] + [""]
    else:
        md += ["## Транспорт", "", "В снимке нет данных «Приход авто» (снимок WMS до версии 0.7).", ""]
    open(os.path.join(out, "INSIGHTS_REPORT.md"), "w", encoding="utf-8").write("\n".join(md))
    return found, counts, tr is not None


def main(argv=None):
    ap = argparse.ArgumentParser(description="закономерности склада по снимку WMS (только чтение снимка)")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--snapshot")
    g.add_argument("--export")
    ap.add_argument("--out")
    ap.add_argument("--today", help="ГГГГ-ММ-ДД (по умолчанию — дата снимка)")
    ap.add_argument("--wms", help="папка рабочей WMS (для проверки резервных копий)")
    ap.add_argument("--long-min", type=int, default=120)
    a = ap.parse_args(argv)
    try:
        snap = a.snapshot or latest_snapshot(a.export)
        today = datetime.date.fromisoformat(a.today) if a.today else None
        wms = a.wms or os.path.dirname(os.path.dirname(os.path.abspath(snap)))
        out = a.out or os.path.join(wms, "WMS_Reports", "INSIGHTS_" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S"))
        found, counts, has_tr = run(snap, out, today, wms, a.long_min)
    except (OSError, ValueError, KeyError) as e:
        print(f"ОШИБКА: {e}", file=sys.stderr)
        return 2
    print(f"закономерностей {len(found)} ({', '.join(f'{lv}: {counts.get(lv, 0)}' for lv in LEVELS)}); транспорт: {'есть' if has_tr else 'нет данных'}; "
          f"результат {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
