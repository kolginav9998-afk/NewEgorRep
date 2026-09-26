"""Independent check of Core Phase 4 «Возвраты» (written separately from the Basic code on purpose).

check(doc, jdir, initial, …) runs tests/receipt_oracle.py — it replays the whole journal (receipts, issues, their
corrections and storno, and the returns: RETURN, RETURN_FIX, RETURN_DEL, with the rules of the task) and compares
«Наличие», «Заказы» (X is the balance of the EI, returns included), «Выдачи» — and then compares the replayed returns
with every view of the book:
  «Возврат»  — every return of the journal in exactly one row (copies aside): № A, issue B, EI C, quantity F, unit G,
               date H, recipient I, place J, balances L M, status N (Проведено / Проведено (исправлено) / Удалено (сторно));
               no posted row without its operation; open rows are not counted anywhere;
  _RET       — one record per return №: issue, EI, quantity, state LIVE / STORNO, a row hint that finds its row;
  _ISS       — per issue: the sum and the number of its live returns (= the replay);
  «Выдачи»   — legacy M «Вернул» / N «Вернул в %» are never written by a return;
  counters   — NEXT_RET above every return № and not below an ABANDON record; numbers never reused.
It also derives the history of a recipient: for each pair (recipient, EI) issued / returned / net — only the registered
movements (a consumable that was not returned is not claimed to be with the person).
"""
import datetime

import receipt_oracle as ro

EPS = 1e-6
NULL_DATE = datetime.date(1899, 12, 30)
ST_POSTED, ST_FIXED, ST_DELETED = "Проведено", "Проведено (исправлено)", "Удалено (сторно)"


def iso_serial(s):
    return float((datetime.date.fromisoformat(s) - NULL_DATE).days) if s else None


def recipient_history(issues, returns):
    """(recipient, EI) → dict(issued, returned, net): live issues and live returns only"""
    h = {}
    for k, iss in issues.items():
        if iss["deleted"]:
            continue
        e = h.setdefault((iss.get("who"), iss["ei"]), dict(issued=0.0, returned=0.0))
        e["issued"] += iss["qty"]
    for n, r in returns.items():
        if r["state"] != "LIVE":
            continue
        e = h.setdefault((r.get("who"), r["ei"]), dict(issued=0.0, returned=0.0))
        e["returned"] += r["qty"]
    for e in h.values():
        e["net"] = round(e["issued"] - e["returned"], 6)
    return h


def check(doc, jdir, initial, expect_tail=0, today=None, base=None, legacy_mn_empty=True, out=None):
    """out (a dict): receives the replayed state of tests/receipt_oracle.py (tests/special_oracle.py continues from it)"""
    st = {} if out is None else out
    P, info = ro.check(doc, jdir, initial, expect_tail, today=today, base=base, out=st)
    P = list(P)

    def bad(msg):
        P.append(msg)

    returns, issues, ret_sum, ret_cnt = st["returns"], st["issues"], st["ret_sum"], st["ret_cnt"]
    sysv = st["sysv"]

    # ------------------------------------------------------------ _RET
    rt = ro.read(doc, "_RET", 6)
    book_ret = {}
    for n, row in enumerate(rt):
        if n == 0 or row[0] == "":
            continue
        if not isinstance(row[0], float) or int(row[0]) != n:
            bad(f"_RET строка {n + 1}: № {row[0]!r} (плотная адресация)")
            continue
        book_ret[n] = row
    for n, r in returns.items():
        row = book_ret.get(n)
        if row is None:
            bad(f"возврат № {n}: нет записи в _RET")
            continue
        got = (int(row[1]) if isinstance(row[1], float) else row[1], row[2], row[4])
        want = (r["issue"], r["ei"], r["state"])
        if got != want or not isinstance(row[3], float) or abs(row[3] - r["qty"]) > EPS:
            bad(f"возврат № {n}: _RET {got} {row[3]!r} ≠ расчёт {want} {r['qty']}")
    for n in book_ret:
        if n not in returns:
            bad(f"_RET: запись № {n} без операции возврата в журнале")

    # ------------------------------------------------------------ _ISS
    isx = ro.read(doc, "_ISS", 4)
    book_iss = {}
    for k, row in enumerate(isx):
        if k == 0 or row[0] == "":
            continue
        if not isinstance(row[0], float) or int(row[0]) != k:
            bad(f"_ISS строка {k + 1}: № {row[0]!r} (плотная адресация)")
            continue
        book_iss[k] = row
    for k in set(ret_sum) | set(ret_cnt):
        row = book_iss.get(k)
        if row is None:
            if ret_cnt.get(k, 0) or abs(ret_sum.get(k, 0.0)) > EPS:
                bad(f"выдача № {k}: есть возвраты, но нет итогов в _ISS")
            continue
        if abs(row[2] - ret_sum.get(k, 0.0)) > EPS or int(row[3]) != ret_cnt.get(k, 0):
            bad(f"выдача № {k}: _ISS возвращено {row[2]} ({int(row[3])}), по журналу {ret_sum.get(k, 0.0)} ({ret_cnt.get(k, 0)})")
    for k in book_iss:
        if k not in ret_sum and k not in ret_cnt:
            bad(f"_ISS: итоги выдачи № {k} без возвратов в журнале")
    for k, q in ret_sum.items():
        iss = issues.get(k)
        if iss is None:
            continue
        if q > iss["qty"] + EPS:
            bad(f"выдача № {k}: возвращено {q} больше выданного {iss['qty']}")
        if iss["deleted"] and ret_cnt.get(k, 0) > 0:
            bad(f"выдача № {k} удалена (сторно), но у неё есть действующие возвраты")

    # ------------------------------------------------------------ «Возврат»
    rows = ro.read(doc, "Возврат", 15)
    seen, copies, open_rows = {}, 0, 0
    for r, row in enumerate(rows):
        if r == 0:
            continue
        ctl = str(row[13])
        if ctl.startswith("КОПИЯ"):
            copies += 1
            continue
        if row[0] == "":
            if any(str(row[c]) != "" for c in (1, 2, 5, 7, 8, 9)):
                open_rows += 1
                if ctl in (ST_POSTED, ST_FIXED, ST_DELETED):
                    bad(f"«Возврат» строка {r + 1}: непроведённая строка со статусом «{ctl}»")
            continue
        if not isinstance(row[0], float):
            bad(f"«Возврат» строка {r + 1}: № {row[0]!r} не число и не помечен КОПИЯ")
            continue
        n = int(row[0])
        if n in seen:
            bad(f"«Возврат»: № {n} в двух строках без отметки КОПИЯ: {seen[n] + 1} и {r + 1}")
            continue
        seen[n] = r
        ret = returns.get(n)
        if ret is None:
            bad(f"«Возврат» строка {r + 1}: № {n} без операции в журнале")
            continue
        want_n = ST_DELETED if ret["state"] == "STORNO" else (ST_FIXED if ret["fixed"] else ST_POSTED)
        if base is not None and ret.get("date") is None:
            # a return of the trusted snapshot: only what the service tables hold
            if row[1] != float(ret["issue"]) or row[2] != ret["ei"] or abs(row[5] - ret["qty"]) > EPS:
                bad(f"«Возврат» строка {r + 1} №{n}: {row[1]!r} {row[2]!r} {row[5]!r} ≠ _RET {ret['issue']} {ret['ei']} {ret['qty']}")
            continue
        got = (row[1], row[2], row[5], row[6], row[7], row[8], row[9], ctl)
        want = (float(ret["issue"]), ret["ei"], ret["qty"], ret.get("unit") or row[6], iso_serial(ret["date"]), ret["who"], ret["place"], want_n)
        if (got[0], got[1], got[3], got[4], got[5], got[6], got[7]) != (want[0], want[1], want[3], want[4], want[5], want[6], want[7]) \
                or not isinstance(got[2], float) or abs(got[2] - want[2]) > EPS:
            bad(f"«Возврат» строка {r + 1} №{n}: {got} ≠ расчёт {want}")
        if ret.get("L") is not None and (not isinstance(row[11], float) or abs(row[11] - ret["L"]) > EPS
                                         or not isinstance(row[12], float) or abs(row[12] - ret["M"]) > EPS):
            bad(f"«Возврат» строка {r + 1} №{n}: остаток до/после {row[11]!r}/{row[12]!r} ≠ {ret['L']}/{ret['M']}")
        iss = issues.get(ret["issue"])
        if ret["state"] == "LIVE" and iss is not None and ret["ei"] != iss["ei"]:
            # a live return belongs to the EI of its issue (a storno keeps the EI of its time as history)
            bad(f"возврат № {n}: ЕИ {ret['ei']} не ЕИ выдачи № {ret['issue']} ({iss['ei']})")
    for n in returns:
        if n not in seen:
            bad(f"возврат № {n} есть в журнале, но строки на листе «Возврат» нет")
    # the row hints of _RET find their rows (a stale hint is allowed: rows inserted above; the operations heal it)
    stale = 0
    for n, row in book_ret.items():
        if isinstance(row[5], float) and int(row[5]) < len(rows) and rows[int(row[5])][0] == float(n):
            continue
        stale += 1

    # ------------------------------------------------------------ «Выдачи» legacy M / N and counters
    if legacy_mn_empty:
        iss_rows = ro.read(doc, "Выдачи", 18)
        for r, row in enumerate(iss_rows):
            if r and isinstance(row[0], float) and (row[12] != "" or row[13] != ""):
                bad(f"«Выдачи» строка {r + 1}: заполнены legacy M/N ({row[12]!r}, {row[13]!r}) — возвраты туда не пишутся")
    nxt = int(sysv["NEXT_RET"])
    if returns and nxt <= max(returns):
        bad(f"NEXT_RET {nxt} не больше максимального № возврата {max(returns)}")
    if nxt < st["min_next_ret"]:
        bad(f"NEXT_RET {nxt} меньше номера, который мог попасть в журнал до отложенного хвоста ({st['min_next_ret']})")

    hist = recipient_history(issues, returns)
    info.update(dict(returns=len(returns), live_returns=sum(1 for r in returns.values() if r["state"] == "LIVE"), return_copies=copies,
                     return_open_rows=open_rows, stale_hints=stale, recipients_pairs=len(hist)))
    return P, info, hist
