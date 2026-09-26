"""Independent check of Core Phase 5 «Все специальные приходы» (written separately from the Basic code on purpose).

check(doc, jdir, initial, …) runs tests/return_oracle.py — it runs tests/receipt_oracle.py, which replays the whole
journal (ordinary receipts, issues, returns and the special receipts SP_RECEIPT / SP_REFILL / SP_FIX / SP_DEL /
SP_IDENTIFY with the rules of the task) — and then compares the replayed special receipts with every view of the book:
  «Иной приход» — every line of the journal in exactly one row (copies aside): № A, type B, event C, name D, article E,
                  quantity F, unit G, date H, place I, category J, «кто передал» K, document L, old marking M, EI N,
                  balances O P, status Q (computed here: Проведено: новый ЕИ / пополнение ЕИ / требует разбора / разобрано,
                  Проведено (исправлено): …, Удалено (сторно)); no posted row without its operation; open rows never
                  carry a posted status;
  _SPR          — one record per line № (dense): event, source, EI, mode NEW/ADD, quantity, LIVE/STORNO, date, the
                  article key of a part, «разобрано как», the number of corrections; the row hint finds its row (a stale
                  hint is allowed: rows inserted above);
  _ART          — the index «article of a part → EI»: exactly the replayed one, no key twice;
  «Наличие»     — the card of every special EI: name, article, unit, place, category, state (Активен / Требует разбора
                  / Приход удалён (сторно)), source «<тип> <событие>» (+ « → <тип>» after «Разобрать»), «Тип источника»;
                  an EI of an ordinary receipt has «Поставщик»;
  counters      — NEXT_SPL and NEXT_OFF … NEXT_OTH above every issued number and not below an ABANDON record.
The invariants S1…S10 (INVARIANTS) are checked by the replay and here. It also derives the history of every EI from the
journal (the task: where it came from, how many times it was refilled, how much came in, was issued, was returned, the
balance, the place) and checks the balance = received − issued + returned.
"""
import datetime

import receipt_oracle as ro
import return_oracle

EPS = 1e-6
NULL_DATE = datetime.date(1899, 12, 30)
ST_DELETED = "Удалено (сторно)"
POSTED_PREFIX, FIXED_PREFIX = "Проведено: ", "Проведено (исправлено): "
K_NEW, K_ADD, K_REVIEW, K_IDENT = "новый ЕИ", "пополнение ЕИ", "требует разбора", "разобрано"

INVARIANTS = {
    "S1": "один артикул детали — один ЕИ (приход известного артикула пополняет его ЕИ, индекс без повторов)",
    "S2": "ЕИ, № строки и № события не переиспользуются, идентичность строки не меняется",
    "S3": "место ЕИ не меняется молча (только отмеченное перемещение)",
    "S4": "ни одной строки дважды: строка проведена один раз и видна один раз",
    "S5": "нет отрицательных остатков",
    "S6": "сторно не оставляет выдач/возвратов без прихода (D-069) и не делает остаток отрицательным",
    "S7": "единица ЕИ не меняется пополнением и при движениях",
    "S8": "иной приход — «Требует разбора» до разбора; разбор сохраняет ЕИ",
    "S9": "событие поступления — одного типа; присоединение только к выданному событию",
    "S10": "антидубль: предупреждение ⇔ есть действующая строка с тем же ключом",
}


def cell_text(v):
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else repr(v)
    return str(v).strip()


def line_status(L):
    if L["state"] == "STORNO":
        return ST_DELETED
    if L["mode"] == "ADD":
        k = K_ADD
    elif L["code"] == "OTH":
        k = K_IDENT if L.get("ident") else K_REVIEW
    else:
        k = K_NEW
    return (FIXED_PREFIX if L.get("fixes") else POSTED_PREFIX) + k


def ei_history(st):
    """EI → origin, refills, received, issued, returned, balance, place — from the replayed journal only"""
    sp, bal, place = st["sp"], st["bal"], st["place"]
    h = {}
    base = st.get("base")
    initial = st.get("initial") or {}
    for ei in bal:
        h[ei] = dict(origin="начальный реестр" if ei in initial else "", refills=0, received=initial.get(ei, 0.0) if base is None else None,
                     issued=0.0, returned=0.0, balance=bal[ei], place=place.get(ei))
    for ei, rec in st["receipts"].items():
        e = h.setdefault(ei, dict(refills=0, issued=0.0, returned=0.0, balance=bal.get(ei), place=place.get(ei)))
        e["origin"] = f"Заказ {rec.get('order')}" if rec.get("order") is not None else "обычный приход"
        e["received"] = rec["qty"] if rec["state"] == "LIVE" else 0.0
    # Final Core: a migrated EI came with its transferred stock (MIGRATE)
    for ei, c in sp["eis"].items():
        if c.get("code") == "MIG":
            e = h.setdefault(ei, dict(refills=0, issued=0.0, returned=0.0, balance=bal.get(ei), place=place.get(ei)))
            e["origin"] = c["src"]
            e["received"] = c.get("migrated", 0.0)
    for n, L in sorted(sp["lines"].items()):
        e = h.setdefault(L["ei"], dict(refills=0, issued=0.0, returned=0.0, balance=bal.get(L["ei"]), place=place.get(L["ei"])))
        if L["mode"] == "NEW":
            e["origin"] = f"{ro.SP_STYPE.get(L['code'], L['code'])} {L['event']}"
            e["received"] = 0.0
        else:
            e["refills"] += 1 if L["state"] == "LIVE" else 0
        if L["state"] == "LIVE" and e.get("received") is not None:
            e["received"] = (e.get("received") or 0.0) + L["qty"]
    for k, iss in st["issues"].items():
        if not iss["deleted"] and iss["ei"] in h:
            h[iss["ei"]]["issued"] += iss["qty"]
    for n, r in st["returns"].items():
        if r["state"] == "LIVE" and r["ei"] in h:
            h[r["ei"]]["returned"] += r["qty"]
    # Final Core: write-offs and inventory corrections change the balance too; moves only the place
    for e in h.values():
        e.setdefault("adjusted", 0.0)
    for n, a in (st.get("adjs") or {}).items():
        if a["state"] == "LIVE" and a["ei"] in h:
            h[a["ei"]]["adjusted"] = h[a["ei"]].get("adjusted", 0.0) + a["delta"]
    return h


def check(doc, jdir, initial, expect_tail=0, today=None, base=None, legacy_mn_empty=True, out=None):
    """out (a dict): receives the replayed state (tests/adjust_oracle.py continues from it)"""
    st = {} if out is None else out
    P, info, rhist = return_oracle.check(doc, jdir, initial, expect_tail, today=today, base=base, legacy_mn_empty=legacy_mn_empty, out=st)
    P = list(P)
    inv_fail = set(x for x in st["inv_fail"] if isinstance(x, str) and x.startswith("S"))

    def bad(msg, inv=None):
        P.append(msg)
        if inv:
            inv_fail.add(inv)

    sp, bal, sysv, place = st["sp"], st["bal"], st["sysv"], st["place"]
    lines, eis, index = sp["lines"], sp["eis"], sp["index"]

    # ------------------------------------------------------------ _SPR
    spr = ro.read(doc, "_SPR", 14)
    book_spr = {}
    for n, row in enumerate(spr):
        if n == 0 or row[0] == "":
            continue
        if not isinstance(row[0], float) or int(row[0]) != n:
            bad(f"_SPR строка {n + 1}: № {row[0]!r} (плотная адресация)", "S4")
            continue
        book_spr[n] = row
    for n, L in lines.items():
        row = book_spr.get(n)
        if row is None:
            bad(f"строка № {n}: нет записи в _SPR", "S4")
            continue
        got = (row[1], row[2], row[3], row[4], row[6], row[11], row[12], int(row[13]) if isinstance(row[13], float) else 0)
        want = (L["event"], L["code"], L["ei"], L["mode"], L["state"], L.get("artkey") or "", L.get("ident") or "", L.get("fixes") or 0)
        if got != want or not isinstance(row[5], float) or abs(row[5] - L["qty"]) > EPS:
            bad(f"строка № {n}: _SPR {got} {row[5]!r} ≠ расчёт {want} {L['qty']}")
        if not L.get("base") and row[8] != float(L["date"]):
            bad(f"строка № {n}: дата в _SPR {row[8]!r} ≠ {L['date']}")
        if L["state"] == "LIVE" and not (isinstance(row[9], float) and row[10]):
            bad(f"строка № {n}: у действующей строки нет ключа антидубля в _SPR", "S10")
    for n in book_spr:
        if n not in lines:
            bad(f"_SPR: запись № {n} без операции в журнале", "S4")

    # ------------------------------------------------------------ _ART
    art = ro.read(doc, "_ART", 3)
    book_idx = {}
    for r, row in enumerate(art):
        if r == 0 or (row[0] == "" and row[1] == ""):
            continue
        k = row[0]
        if k in book_idx:
            bad(f"_ART: артикул «{k}» записан дважды ({book_idx[k][0]}, {row[1]}) — один артикул у нескольких ЕИ", "S1")
            continue
        book_idx[k] = (row[1], row[2])
        if ro.article_key(str(row[2])) != k:
            bad(f"_ART строка {r + 1}: артикул «{row[2]}» не даёт ключ «{k}»", "S1")
    if {k: v[0] for k, v in book_idx.items()} != index:
        extra = sorted(set(book_idx) - set(index))[:5]
        missing = sorted(set(index) - set(book_idx))[:5]
        other = sorted(k for k in set(book_idx) & set(index) if book_idx[k][0] != index[k])[:5]
        bad(f"_ART ≠ расчётному индексу: лишние {extra}, нет {missing}, другой ЕИ {other}", "S1")
    parts = {ei for ei, c in eis.items() if c.get("part")}
    if parts != set(index.values()) or len(set(index.values())) != len(index):
        bad(f"индекс: деталей {len(parts)}, ЕИ в индексе {len(set(index.values()))}, записей {len(index)} — у детали не ровно один артикул", "S1")

    # ------------------------------------------------------------ «Иной приход»
    rows = ro.read(doc, "Иной приход", 19)
    seen, copies, open_rows = {}, 0, 0
    for r, row in enumerate(rows):
        if r == 0:
            continue
        ctl, status = str(row[17]), str(row[16])
        if ctl.startswith("КОПИЯ"):
            copies += 1
            continue
        if row[0] == "":
            if any(str(row[c]) != "" for c in range(1, 13)):
                open_rows += 1
                if status.startswith("Проведено") or status == ST_DELETED:
                    bad(f"«Иной приход» строка {r + 1}: непроведённая строка со статусом «{status}»", "S4")
            continue
        if not isinstance(row[0], float):
            bad(f"«Иной приход» строка {r + 1}: № {row[0]!r} не число и не помечен КОПИЯ", "S4")
            continue
        n = int(row[0])
        if n in seen:
            bad(f"«Иной приход»: № {n} в двух строках без отметки КОПИЯ: {seen[n] + 1} и {r + 1}", "S4")
            continue
        seen[n] = r
        L = lines.get(n)
        if L is None:
            bad(f"«Иной приход» строка {r + 1}: № {n} без операции в журнале", "S4")
            continue
        if L.get("base"):
            if row[13] != L["ei"] or not isinstance(row[5], float) or abs(row[5] - L["qty"]) > EPS or status != line_status(L):
                bad(f"«Иной приход» строка {r + 1} №{n}: {row[13]!r} {row[5]!r} «{status}» ≠ _SPR {L['ei']} {L['qty']} «{line_status(L)}»")
            continue
        want = dict(B=ro.SP_NAME[L["code"]], C=L["event"], D=L["name"], E=L["art"], G=L["unit"], H=float(L["date"]), I=L["place"], J=L["cat"],
                    K=L["who"], L=L["doc"], M=L["mark"], N=L["ei"], Q=line_status(L))
        got = dict(B=row[1], C=row[2], D=cell_text(row[3]), E=cell_text(row[4]), G=row[6], H=row[7], I=cell_text(row[8]), J=cell_text(row[9]),
                   K=cell_text(row[10]), L=cell_text(row[11]), M=cell_text(row[12]), N=row[13], Q=status)
        diff = {k: (got[k], want[k]) for k in want if got[k] != want[k]}
        if not isinstance(row[5], float) or abs(row[5] - L["qty"]) > EPS:
            diff["F"] = (row[5], L["qty"])
        if L["before"] is None:
            if row[14] != "":
                diff["O"] = (row[14], "")
        elif not isinstance(row[14], float) or abs(row[14] - L["before"]) > EPS:
            diff["O"] = (row[14], L["before"])
        if not isinstance(row[15], float) or abs(row[15] - L["after"]) > EPS:
            diff["P"] = (row[15], L["after"])
        if diff:
            bad(f"«Иной приход» строка {r + 1} №{n}: {diff}")
    for n in lines:
        if n not in seen:
            bad(f"строка № {n} есть в журнале, но на листе «Иной приход» её нет", "S4")
    stale = 0
    for n, row in book_spr.items():
        if isinstance(row[7], float) and 0 < int(row[7]) < len(rows) and rows[int(row[7])][0] == float(n):
            continue
        stale += 1

    # ------------------------------------------------------------ «Наличие»: the cards
    stock = ro.read(doc, "Наличие", 10)
    for ei, c in eis.items():
        k = ro.ei_num(ei)
        if k is None or k >= len(stock) or stock[k][0] != ei:
            bad(f"{ei}: нет строки в «Наличие»", 3)
            continue
        row = stock[k]
        if c.get("base"):
            continue
        got = (row[1], row[2], row[3], row[5], row[6], row[7], row[8], row[9])
        want = (c["name"], c["art"], c["unit"], place.get(ei), c["cat"], c["state"], c["src"], c["stype"])
        if got != want:
            bad(f"{ei}: карточка «Наличие» {got} ≠ расчёт {want}", "S3" if got[3] != want[3] else 3)
    if base is None:
        for ei, rec in st["receipts"].items():
            k = ro.ei_num(ei)
            if k is not None and k < len(stock) and stock[k][0] == ei and stock[k][9] != "Поставщик":
                bad(f"{ei} (обычный приход): «Тип источника» {stock[k][9]!r} ≠ «Поставщик»")
        for ei in initial:
            k = ro.ei_num(ei)
            if k is not None and k < len(stock) and stock[k][9] != "":
                bad(f"{ei} (начальный реестр): «Тип источника» {stock[k][9]!r} — ожидалось пусто")
    for ei, b in bal.items():
        if b < -EPS:
            bad(f"{ei}: отрицательный остаток {b}", "S5")

    # ------------------------------------------------------------ counters
    nxt = sysv.get("NEXT_SPL")
    top = max([sp["max_spl"]] + list(sp["abandoned_spl"]))
    if not isinstance(nxt, float) or nxt <= top or nxt < sp["floor_spl"]:
        bad(f"NEXT_SPL {nxt!r}: не выше выданных № строк ({top}) или ниже ABANDON ({sp['floor_spl']})", "S2")
    elif base is None and nxt != max(top + 1, sp["floor_spl"]):
        bad(f"NEXT_SPL {nxt}: ожидалось {max(top + 1, sp['floor_spl'])}", "S2")
    for code, key in ro.SP_COUNTER.items():
        v = sysv.get(key)
        top = max([sp["max_ev"][code]] + list(sp["abandoned_ev"][code]))
        if not isinstance(v, float) or v <= top or v < sp["floor_ev"][code]:
            bad(f"{key} {v!r}: не выше выданных № событий {code} ({top}) или ниже ABANDON ({sp['floor_ev'][code]})", "S2")
        elif base is None and v != max(top + 1, sp["floor_ev"][code]):
            bad(f"{key} {v}: ожидалось {max(top + 1, sp['floor_ev'][code])}", "S2")

    # ------------------------------------------------------------ history of every EI (the task: «ИСТОРИЯ»)
    hist = ei_history(st)
    if base is None:
        for ei, e in hist.items():
            if e.get("received") is None:
                continue
            want = e["received"] - e["issued"] + e["returned"] + e.get("adjusted", 0.0)
            if abs(want - (e["balance"] or 0.0)) > 1e-4:
                bad(f"{ei}: история не сходится: пришло {e['received']} − выдано {e['issued']} + возвращено {e['returned']} = {want}, "
                    f"остаток {e['balance']}", 3)

    live = sum(1 for L in lines.values() if L["state"] == "LIVE")
    info.update(dict(special_ops=sp["ops"], special_lines=len(lines), special_live=live, special_eis=len(eis), parts=len(parts),
                     index=len(index), events={c: len(v) for c, v in sp["events"].items() if v}, special_copies=copies,
                     special_open_rows=open_rows, special_stale_hints=stale,
                     invariants_failed=sorted(set(info.get("invariants_failed", [])) | inv_fail, key=str)))
    return P, info, hist
