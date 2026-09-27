#!/usr/bin/env python3
"""WMS_RECONCILE — offline matching of old names written by hand with the existing EIs (FINAL WMS MARATHON §6).

    python3 wms_reconcile.py NAMES --snapshot SNAPSHOT_DIR [--out REPORT.csv] [--top 5]
    python3 wms_reconcile.py NAMES --export WMS_EXPORT_DIR  [...]        (the latest snapshot of the folder)
    python3 wms_reconcile.py --confirm REPORT.csv (--snapshot S | --export E) --out MAPPING.csv

NAMES — the old names: a CSV/TXT file (UTF-8 or Windows-1251), one name per line, or a table with a column «Наименование»
(«Название», «Товар», «Старое название»); a column «Количество» / «Место» / «Артикул» is used when present. SNAPSHOT_DIR —
a snapshot of WMS («Экспорт для инструментов», docs/EXPORT_CONTRACT.md): stock.csv is read, the snapshot is checked by its
manifest (SHA-256) first.

For every old name: up to --top (5) candidates among the EIs of the snapshot with a confidence 0–100 and the way it was
found: «точное» (the normalized names are equal — 100), «артикул» (the article of the EI is in the old record), «похожее»
(words of the name — with the letters ё/е, the Latin and the Cyrillic look-alikes, the separators of sizes unified; the
sizes and codes such as «м12», «3х2,5», «z-40» weigh more; the place of the old record is compared when given). Verdict:
«уверенно» (≥ 85 and ahead of the second by 10), «проверить» (≥ 55), «слабо», «нет кандидатов».

The report is a proposal for a person: nothing is written anywhere, nothing is written off or merged automatically. The
person opens the report (LibreOffice, «;», UTF-8), writes «да» in the column «подтвердить» of the one right candidate of an
old name (or types an EI in «ЕИ» of a row «нет кандидатов» and «да» next to it) and saves it as CSV. --confirm then makes
the mapping: every old name — its confirmed EI (checked against the snapshot: it exists, its receipt is not cancelled),
«без решения» (nothing confirmed) or a problem («подтверждено несколько ЕИ», «ЕИ нет в снимке»). The mapping is a
document for a person too: a decision becomes a WMS operation only by a person (an inventory correction, a move).

Groups (M6 PRIME §9): the old names that are the same after the normalization form a group (column «группа»); a name that
differs only a little (the same sizes and codes, similarity ≥ 0.9) joins the group as «почти то же» — the rows of a group
stand together in the report. In --confirm the decision of a name goes to the unconfirmed names of its group that are the
same after the normalization (never to a «почти то же» name: that stays a candidate for a person).
The dictionary (--dict, default <folder of the snapshots>/../WMS_Reconcile/dictionary.csv) keeps every mapping a person
confirmed (--confirm adds and updates it). A later report proposes the dictionary EI of a name that is the same after the
normalization as «подтверждено ранее» with «да» already written in «подтвердить» (the person may remove it); a fuzzy match
is only ever a candidate. --no-dict: neither read nor write the dictionary.
Exit code: 0 done, 1 --confirm found problems (the mapping is written, the problem rows are marked), 2 the input could not
be read.
"""
import argparse
import csv
import datetime
import difflib
import hashlib
import io
import os
import re
import sys

LOOK = str.maketrans({"a": "а", "e": "е", "o": "о", "p": "р", "c": "с", "x": "х", "y": "у", "k": "к", "m": "м", "t": "т", "b": "в", "h": "н",
                      "ё": "е"})
NAME_COLS = ("наименование", "название", "товар", "старое название", "номенклатура", "name")
QTY_COLS = ("количество", "кол-во", "остаток", "qty")
PLACE_COLS = ("место", "место хранения", "ячейка", "place")
ART_COLS = ("артикул", "код", "article")


def read_text(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"{path}: кодировка не UTF-8 и не Windows-1251")


def norm(s):
    """lower case, ё → е, the Latin look-alikes as Cyrillic (both sides the same way), «х»/«x»/«*» between digits unified,
    the decimal point as a comma, punctuation → spaces"""
    s = str(s or "").lower().replace("ё", "е")
    s = re.sub(r"(?<=\d)\s*[x×*хХ]\s*(?=\d)", "х", s)          # 3x2,5 / 3 х 2.5 → 3х2,5
    s = re.sub(r"(?<=\d)[.,](?=\d)", ",", s)                   # 2.5 → 2,5
    s = re.sub(r"[^\w,]+", " ", s)
    return [w.translate(LOOK) for w in (x.strip(",") for x in s.split()) if w]


def numbers(words):
    """the sizes and codes as numbers (м12х40 → 12, 40; 3х2,5 → 3, 2,5; e27 → 27)"""
    return set(re.findall(r"\d+(?:,\d+)?", " ".join(words)))


def art_key(s):
    return re.sub(r"[\s\-‐‑‒–—−_./]", "", str(s or "").upper()).translate(str.maketrans("АВЕКМНОРСТУХ", "ABEKMHOPCTYX"))


def read_names(path):
    """[(line, name, qty, place, article)]"""
    text = read_text(path)
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return []
    first = lines[0]
    delim = ";" if first.count(";") >= max(first.count(","), first.count("\t")) and ";" in first else ("\t" if "\t" in first else ("," if "," in first else None))
    if delim:
        rows = list(csv.reader(io.StringIO("\n".join(lines)), delimiter=delim))
        head = [h.strip().lower() for h in rows[0]]

        def col(names):
            return next((i for i, h in enumerate(head) if h in names), None)
        ni, qi, pi, ai = col(NAME_COLS), col(QTY_COLS), col(PLACE_COLS), col(ART_COLS)
        if ni is not None:
            out = []
            for k, r in enumerate(rows[1:], start=2):
                def g(i):
                    return r[i].strip() if i is not None and i < len(r) else ""
                if g(ni):
                    out.append((k, g(ni), g(qi), g(pi), g(ai)))
            return out
    # one name per line; a first line that is a column header («Наименование», «Старое название» …) is not a name
    out = [(k, ln.strip(), "", "", "") for k, ln in enumerate(lines, start=1)]
    return out[1:] if out[0][1].strip().lower() in NAME_COLS else out


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


def verify_snapshot(snap):
    man = os.path.join(snap, "manifest.csv")
    fmt = None
    for row in csv.reader(io.StringIO(read_text(man)), delimiter=";"):
        if row and row[0] == "format":
            fmt = row[1]
        if row and row[0] == "file" and row[1] == "stock.csv":
            if sha256(os.path.join(snap, "stock.csv")) != row[4]:
                raise ValueError("stock.csv снимка изменён после экспорта (SHA-256 не совпадает)")
    if fmt != "WMS-SNAPSHOT-1":
        raise ValueError(f"формат снимка «{fmt}», ожидается WMS-SNAPSHOT-1")


def read_stock(snap):
    rows = list(csv.reader(io.StringIO(read_text(os.path.join(snap, "stock.csv"))), delimiter=";"))
    head = rows[0]
    ix = {k: head.index(k) for k in ("ei", "name", "article", "unit", "qty", "place", "category", "state")}
    out = []
    for r in rows[1:]:
        if len(r) > ix["state"] and r[ix["ei"]] and r[ix["state"]] != "Приход удалён (сторно)":
            out.append({k: r[i] for k, i in ix.items()})
    return out


def by_article(old, cand):
    """the article of the candidate is the article of the old record or is written in its name"""
    ak = art_key(cand["article"])
    return bool(ak) and (ak == art_key(old[4]) or (len(ak) >= 4 and ak in art_key(old[1])))


def method(old, cand):
    if norm(old[1]) == norm(cand["name"]):
        return "точное"
    return "артикул" if by_article(old, cand) else "похожее"


def score(old, cand, cw, compact):
    """0–100: the same normalized name — 100; otherwise the words of the old name found in the candidate (exactly, inside
    a compound word, or nearly), the whole string, the sizes and codes (numbers), the article, the place"""
    ow = norm(old[1])
    if not ow:
        return 0
    if ow == norm(cand["name"]):
        return 100
    cws = set(cw)
    matched = 0.0
    for w in ow:
        if w in cws:
            matched += 1.0
        elif len(w) >= 2 and w in compact:
            matched += 0.9
        else:
            best = max((difflib.SequenceMatcher(None, w, c).ratio() for c in cw if abs(len(w) - len(c)) <= 3), default=0.0)
            matched += best if best >= 0.8 else 0.0
    words = matched / len(ow) * 0.6 + matched / max(len(ow), len(cw)) * 0.4
    whole = difflib.SequenceMatcher(None, " ".join(ow), " ".join(cw)).ratio()
    so, sc = numbers(ow), numbers(cw)
    if so:
        sizes = len(so & sc) / len(so)
        s = 0.55 * words + 0.20 * whole + 0.25 * sizes - (0.2 if not so & sc else 0.0)
    else:
        s = 0.65 * words + 0.35 * whole
    if by_article(old, cand):
        s += 0.3
    if old[3] and cand["place"] and old[3].strip().lower() == cand["place"].strip().lower():
        s += 0.05
    return max(0, min(100, round(s * 100)))


def name_key(s):
    return " ".join(norm(s))


def groups(names):
    """[(line, name, …)] → {index: (group no, kind)}: the same normalized name — one group («то же»); a name with the same
    numbers and a similarity ≥ 0.9 to the first name of a group joins it («почти то же»)"""
    out, heads = {}, []                      # heads: (key, numbers, words, group no)
    for i, old in enumerate(names):
        k = name_key(old[1])
        w = norm(old[1])
        nums = numbers(w)
        hit = next((h for h in heads if h[0] == k), None)
        if hit:
            out[i] = (hit[3], "то же")
            continue
        near = None
        for h in heads:
            if h[1] == nums and h[2][:1] and w[:1] and h[2][0][:3] == w[0][:3] and difflib.SequenceMatcher(None, h[0], k).ratio() >= 0.9:
                near = h
                break
        if near:
            out[i] = (near[3], "почти то же")
        else:
            heads.append((k, nums, w, len(heads) + 1))
            out[i] = (len(heads), "")
    return out


def dict_path(a, snap):
    if a.no_dict:
        return None
    if a.dict:
        return a.dict
    exp = a.export or os.path.dirname(os.path.abspath(snap.rstrip("/")))
    return os.path.join(os.path.dirname(os.path.abspath(exp)), "WMS_Reconcile", "dictionary.csv")


def read_dict(path):
    """key → [key, old name, EI, name, confirmed, source]"""
    if not path or not os.path.exists(path):
        return {}
    rows = list(csv.reader(io.StringIO(read_text(path)), delimiter=";"))
    return {r[0]: (r + [""] * len(DICT_HEAD))[:len(DICT_HEAD)] for r in rows[1:] if r and r[0]}


def write_dict(path, d):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".part"
    with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(DICT_HEAD)
        for k in sorted(d):
            w.writerow(d[k])
    os.replace(tmp, path)


def reconcile(names, stock, top=5):
    index = {}
    words, compacts = [], []
    for k, c in enumerate(stock):
        cw = norm(c["name"] + " " + c["article"])
        words.append(cw)
        compacts.append("".join(cw))
        for w in set(cw):
            index.setdefault(w[:3], set()).add(k)
            for n in numbers([w]):
                index.setdefault("#" + n, set()).add(k)
    results = []
    for old in names:
        ow = norm(old[1])
        cand = set()
        for w in ow:
            cand |= index.get(w[:3], set())
            for n in numbers([w]):
                cand |= index.get("#" + n, set())
        if art_key(old[4]):
            cand |= {k for k, c in enumerate(stock) if art_key(c["article"]) == art_key(old[4])}
        scored = sorted(((score(old, stock[k], words[k], compacts[k]), k) for k in cand), key=lambda x: (-x[0], stock[x[1]]["ei"]))[:top]
        scored = [x for x in scored if x[0] >= 40]
        if not scored:
            verdict = "нет кандидатов"
        elif scored[0][0] >= 85 and (len(scored) == 1 or scored[0][0] - scored[1][0] >= 10):
            verdict = "уверенно"
        elif scored[0][0] >= 55:
            verdict = "проверить"
        else:
            verdict = "слабо"
        results.append((old, verdict, [(sc, stock[k]) for sc, k in scored]))
    return results


HEAD = ["строка", "старое название", "вывод", "место кандидата", "ЕИ", "наименование", "артикул", "остаток", "место", "уверенность",
        "способ", "подтвердить", "группа", "в группе"]
NBASE = 12                  # the columns a report of WMS_TOOLBOX 1.0 had (--confirm reads those too)
DICT_HEAD = ["ключ", "старое название", "ЕИ", "наименование", "подтверждено", "источник"]
YES = ("да", "+", "1", "x", "х", "yes")


def ei_key(s):
    """«12», «ЕИ-12», «ЕИ-00000012» → «ЕИ-00000012»; "" when it is not an EI"""
    t = str(s or "").strip().upper().replace("EI-", "ЕИ-").replace("ЕИ ", "ЕИ-")
    if t.startswith("ЕИ-"):
        t = t[3:]
    return f"ЕИ-{int(t):08d}" if t.isdigit() and 0 < int(t) <= 99999999 else ""


def confirm(report, stock):
    """the reviewed report → [(line, old name, decision, EI, stock row or None, note)], number of problems"""
    rows = list(csv.reader(io.StringIO(read_text(report)), delimiter=";"))
    if not rows or [h.strip().lower() for h in rows[0][:len(HEAD)]] != [h.lower() for h in HEAD]:
        raise ValueError(f"{report}: это не отчёт сверки (первая строка — {';'.join(HEAD)})")
    by_ei = {c["ei"]: c for c in stock}
    groups, order = {}, []
    for r in rows[1:]:
        r = (r + [""] * len(HEAD))[:len(HEAD)]
        key = (r[0].strip(), r[1].strip())
        if not key[1]:
            continue
        if key not in groups:
            groups[key] = []
            order.append(key)
        if r[11].strip().lower() in YES:
            groups[key].append(r)
    out, bad = [], 0
    for key in order:
        chosen = groups[key]
        if not chosen:
            out.append((key[0], key[1], "без решения", "", None, "ничего не подтверждено"))
        elif len({ei_key(r[4]) or r[4].strip() for r in chosen}) > 1:
            bad += 1
            out.append((key[0], key[1], "ошибка", "", None, "подтверждено несколько ЕИ: " + ", ".join(r[4].strip() for r in chosen)))
        else:
            e = ei_key(chosen[0][4])
            c = by_ei.get(e)
            if c is None:
                bad += 1
                out.append((key[0], key[1], "ошибка", chosen[0][4].strip(), None, f"ЕИ «{chosen[0][4].strip()}» нет в снимке"))
            else:
                note = "" if not chosen[0][5].strip() or norm(chosen[0][5]) == norm(c["name"]) else f"в отчёте «{chosen[0][5].strip()}»"
                out.append((key[0], key[1], "подтверждено", e, c, note))
    # the same name after the normalization (one group «то же»): the decision goes to its unconfirmed names; different
    # EIs confirmed for the same name — a problem of each of them
    by_key = {}
    for x in out:
        if x[2] == "подтверждено":
            by_key.setdefault(name_key(x[1]), []).append(x)
    res = []
    for x in out:
        same = by_key.get(name_key(x[1]), [])
        eis = {y[3] for y in same}
        if len(eis) > 1:
            bad += 1
            res.append((x[0], x[1], "ошибка", x[3], None, "одинаковые названия подтверждены разными ЕИ: " + ", ".join(sorted(eis))))
        elif x[2] == "без решения" and same:
            y = same[0]
            res.append((x[0], x[1], "подтверждено", y[3], y[4], f"как «{y[1]}» (то же название, строка {y[0]})"))
        else:
            res.append(x)
    return res, bad


def apply_dict(results, stock, dic):
    """a name whose normalized form is in the dictionary with an EI of the snapshot: that EI first, «подтверждено ранее»"""
    by_ei = {c["ei"]: c for c in stock}
    out, used = [], 0
    for old, verdict, cands in results:
        hit = dic.get(name_key(old[1]))
        c = by_ei.get(hit[2]) if hit else None
        if c is None:
            out.append((old, verdict, cands, ""))
            continue
        used += 1
        rest = [(sc, x) for sc, x in cands if x["ei"] != c["ei"]][:4]
        out.append((old, "подтверждено ранее", [(100, c)] + rest, f"словарь (подтверждено {hit[4]})"))
    return out, used


def main(argv=None):
    ap = argparse.ArgumentParser(description="сверка старых названий с ЕИ по снимку WMS (только отчёт)")
    ap.add_argument("names", nargs="?")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--snapshot")
    g.add_argument("--export")
    ap.add_argument("--out")
    ap.add_argument("--top", type=int, default=5)
    ap.add_argument("--confirm", metavar="REPORT", help="отчёт сверки с отметками «да» → соответствие старых названий и ЕИ")
    ap.add_argument("--dict", help="словарь подтверждённых соответствий (по умолчанию ../WMS_Reconcile/dictionary.csv рядом с папкой снимков)")
    ap.add_argument("--no-dict", action="store_true", help="не читать и не пополнять словарь")
    a = ap.parse_args(argv)
    if bool(a.names) == bool(a.confirm):
        ap.error("нужен файл названий или --confirm ОТЧЁТ")
    try:
        snap = a.snapshot or latest_snapshot(a.export)
        verify_snapshot(snap)
        stock = read_stock(snap)
        dpath = dict_path(a, snap)
        dic = read_dict(dpath)
        if a.confirm:
            res, bad = confirm(a.confirm, stock)
        else:
            names = read_names(a.names)
    except (OSError, ValueError, KeyError) as e:
        print(f"ОШИБКА: {e}", file=sys.stderr)
        return 2
    out = io.StringIO()
    w = csv.writer(out, delimiter=";", lineterminator="\n")
    counts = {}
    extra = ""
    if a.confirm:
        w.writerow(["строка", "старое название", "решение", "ЕИ", "наименование", "артикул", "единица", "остаток", "место", "примечание"])
        for line, old, dec, e, c, note in res:
            counts[dec] = counts.get(dec, 0) + 1
            w.writerow([line, old, dec, e] + ([c["name"], c["article"], c["unit"], c["qty"], c["place"]] if c else [""] * 5) + [note])
        if dpath:
            added = changed = 0
            today = datetime.date.today().isoformat()
            for line, old, dec, e, c, note in res:
                if dec != "подтверждено":
                    continue
                k = name_key(old)
                prev = dic.get(k)
                if prev and prev[2] == e:
                    continue
                changed += bool(prev)
                added += not prev
                dic[k] = [k, old, e, c["name"], today, os.path.basename(a.confirm)]
            if added or changed:
                try:
                    write_dict(dpath, dic)
                except OSError as ex:
                    print(f"ОШИБКА: словарь не записан: {ex}", file=sys.stderr)
                    return 2
            extra = f"; словарь: добавлено {added}, изменено {changed}, всего {len(dic)}"
    else:
        res = reconcile(names, stock, a.top)
        used = 0
        if dic:
            res, used = apply_dict(res, stock, dic)
        else:
            res = [(old, verdict, cands, "") for old, verdict, cands in res]
        grp = groups([r[0] for r in res])
        size = {}
        for g, kind in grp.values():
            size[g] = size.get(g, 0) + 1
        order = sorted(range(len(res)), key=lambda i: (grp[i][0], i))
        w.writerow(HEAD)
        for i in order:
            old, verdict, cands, how = res[i]
            g, kind = grp[i]
            gcol = [f"Г{g}" + (f" ({kind})" if kind else ""), size[g]] if size[g] > 1 else ["", ""]
            counts[verdict] = counts.get(verdict, 0) + 1
            if not cands:
                w.writerow([old[0], old[1], verdict, "", "", "", "", "", "", "", "", ""] + gcol)
            for rank, (sc, c) in enumerate(cands, start=1):
                fromdict = how and rank == 1
                w.writerow([old[0], old[1], verdict if rank == 1 else "", rank, c["ei"], c["name"], c["article"], c["qty"], c["place"], sc,
                            how if fromdict else method(old, c), "да" if fromdict else ""] + gcol)
        ng = sum(1 for g, n in size.items() if n > 1)
        extra = (f"; групп одинаковых / похожих названий {ng}" if ng else "") + (f"; из словаря {used}" if used else "")
    text = out.getvalue()
    if a.out:
        open(a.out, "w", encoding="utf-8-sig", newline="").write(text)
    summary = ", ".join(f"{k}: {v}" for k, v in sorted(counts.items()))
    what = "соответствие" if a.confirm else "отчёт"
    print(f"названий {len(res)}; {summary}; снимок {os.path.basename(snap.rstrip('/'))}" + extra + (f"; {what} {a.out}" if a.out else ""))
    if not a.out:
        print(text)
    return 1 if a.confirm and bad else 0


if __name__ == "__main__":
    sys.exit(main())
