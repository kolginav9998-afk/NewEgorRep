"""Independent check of the Final Core corrections (written separately from the Basic code on purpose).

check(doc, jdir, initial, …) runs tests/special_oracle.py (it runs return_oracle and receipt_oracle — the whole journal
is replayed there, the corrections MOVE / WRITE_OFF / INV_ADJ and their *_FIX / *_DEL by receipt_oracle.replay_adjust
with the rules of the task) and then compares the replayed corrections with every view of the book:
  «Корректировки» — every correction of the journal in exactly one row (copies aside): № A, kind B, EI C, the value of
                    its kind (G quantity of a write-off, H actual and I book balance of an inventory, J new place of a
                    move), date K, reason L, place before M, difference P, status Q (Проведено / Проведено (исправлено) /
                    Удалено (сторно)), batch R; no posted row without its operation; open rows are counted nowhere;
  _ADJ            — one record per №: kind, EI, change of the balance, state LIVE / STORNO, places, book / actual balance,
                    date, a row hint that finds its row;
  «Наличие»       — balance and place (checked by receipt_oracle for every EI a correction touched);
  counters        — NEXT_ADJ above every № and not below an ABANDON record; numbers never reused.
"""
import receipt_oracle as ro
import special_oracle

EPS = 1e-6
ST_POSTED, ST_FIXED, ST_DELETED = "Проведено", "Проведено (исправлено)", "Удалено (сторно)"
KIND_NAME = {"MOVE": "Перемещение", "WRITE_OFF": "Списание", "INV_ADJ": "Инвентаризация"}


def check(doc, jdir, initial, expect_tail=0, today=None, base=None, legacy_mn_empty=True):
    st = {}
    P, info, hist = special_oracle.check(doc, jdir, initial, expect_tail, today=today, base=base, legacy_mn_empty=legacy_mn_empty, out=st)
    P = list(P)

    def bad(msg):
        P.append(msg)

    adjs, sysv = st["adjs"], st["sysv"]
    # ------------------------------------------------------------ _ADJ
    rows = ro.read(doc, "_ADJ", 11)
    book = {}
    for n, row in enumerate(rows):
        if n == 0 or row[0] == "":
            continue
        if not isinstance(row[0], float) or int(row[0]) != n:
            bad(f"_ADJ строка {n + 1}: № {row[0]!r} (плотная адресация)")
            continue
        book[n] = row
    for n, a in adjs.items():
        row = book.get(n)
        if row is None:
            bad(f"корректировка № {n}: нет записи в _ADJ")
            continue
        got = (row[1], row[2], row[4])
        want = (a["kind"], a["ei"], a["state"])
        if got != want or not isinstance(row[3], float) or abs(row[3] - a["delta"]) > EPS:
            bad(f"корректировка № {n}: _ADJ {got} {row[3]!r} ≠ расчёт {want} {a['delta']}")
        if a["kind"] == "MOVE" and (row[6], row[7]) != (a["frm"] or "", a["to"] or ""):
            bad(f"перемещение № {n}: _ADJ места {(row[6], row[7])} ≠ {(a['frm'], a['to'])}")
        if a["kind"] == "INV_ADJ" and (not isinstance(row[8], float) or abs(row[8] - a["book"]) > EPS
                                       or not isinstance(row[9], float) or abs(row[9] - a["fact"]) > EPS):
            bad(f"инвентаризация № {n}: _ADJ учётный/факт {row[8]!r}/{row[9]!r} ≠ {a['book']}/{a['fact']}")
        if row[10] != float(ro.iso_serial(a["date"])):
            bad(f"корректировка № {n}: _ADJ дата {row[10]!r} ≠ {a['date']}")
    for n in book:
        if n not in adjs:
            bad(f"_ADJ: корректировка № {n} без операции в журнале")

    # ------------------------------------------------------------ «Корректировки»
    sh = ro.read(doc, "Корректировки", 19)
    seen, copies, open_rows = {}, 0, 0
    for r, row in enumerate(sh):
        if r == 0:
            continue
        ctl = row[16] if isinstance(row[16], str) else ""
        if ctl.startswith("КОПИЯ"):
            copies += 1
            continue
        if row[0] == "":
            if any(row[i] != "" for i in (1, 2, 6, 7, 9, 10, 11)):
                open_rows += 1
                if ctl in (ST_POSTED, ST_FIXED, ST_DELETED):
                    bad(f"«Корректировки» строка {r + 1}: без № , но со статусом проведения «{ctl}»")
            continue
        if not isinstance(row[0], float):
            bad(f"«Корректировки» строка {r + 1}: № {row[0]!r} не число и не помечен КОПИЯ")
            continue
        n = int(row[0])
        if n in seen:
            bad(f"«Корректировки»: № {n} в строках {seen[n] + 1} и {r + 1} без пометки КОПИЯ")
            continue
        seen[n] = r
        a = adjs.get(n)
        if a is None:
            bad(f"«Корректировки» строка {r + 1}: № {n} без операции в журнале")
            continue
        want_st = ST_DELETED if a["state"] == "STORNO" else (ST_FIXED if a["fixes"] else ST_POSTED)
        if ctl != want_st:
            bad(f"корректировка № {n}: статус «{ctl}», ожидался «{want_st}»")
        if (row[1], row[2]) != (KIND_NAME[a["kind"]], a["ei"]):
            bad(f"корректировка № {n}: B C {(row[1], row[2])} ≠ {(KIND_NAME[a['kind']], a['ei'])}")
        if row[10] != float(ro.iso_serial(a["date"])) or row[11] != (a["reason"] or ""):
            bad(f"корректировка № {n}: дата/причина {(row[10], row[11])!r} ≠ {(a['date'], a['reason'])}")
        if (row[17] or "") != (a.get("batch") or ""):
            bad(f"корректировка № {n}: пакет {row[17]!r} ≠ {a.get('batch')!r}")
        if a["kind"] == "MOVE":
            if (row[9], row[12]) != (a["to"], a["frm"] or ""):
                bad(f"перемещение № {n}: J/M {(row[9], row[12])} ≠ {(a['to'], a['frm'])}")
        elif a["kind"] == "WRITE_OFF":
            if not isinstance(row[6], float) or abs(row[6] + a["delta"]) > EPS or not isinstance(row[15], float) or abs(row[15] - a["delta"]) > EPS:
                bad(f"списание № {n}: G/P {(row[6], row[15])} ≠ {(-a['delta'], a['delta'])}")
        else:
            if (not isinstance(row[7], float) or abs(row[7] - a["fact"]) > EPS or not isinstance(row[8], float) or abs(row[8] - a["book"]) > EPS
                    or not isinstance(row[15], float) or abs(row[15] - a["delta"]) > EPS):
                bad(f"инвентаризация № {n}: H/I/P {(row[7], row[8], row[15])} ≠ {(a['fact'], a['book'], a['delta'])}")
        hint = book.get(n)
        if hint is not None and isinstance(hint[5], float) and int(hint[5]) != r:
            pass            # a stale hint after inserted rows is not an error (it is corrected by the next operation)
    for n in adjs:
        if n not in seen:
            bad(f"корректировка № {n}: нет строки на листе «Корректировки»")

    # ------------------------------------------------------------ counter
    v = sysv.get("NEXT_ADJ")
    top = max(list(adjs) + [0])
    if not isinstance(v, float) or v <= top:
        bad(f"NEXT_ADJ {v!r}: не выше выданных № корректировок ({top})")
    elif base is None and v != top + 1 and not _abandon_floor(jdir, sysv, v):
        bad(f"NEXT_ADJ {v}: ожидалось {top + 1}")
    info.update(dict(adjust_ops=sum(1 for e in st["ops"] if e["type"] in ro.ADJUST_TYPES), adjustments=len(adjs),
                     adjust_live=sum(1 for a in adjs.values() if a["state"] == "LIVE"), adjust_copies=copies, adjust_open_rows=open_rows))
    return P, info, hist


def _abandon_floor(jdir, sysv, v):
    """True when an ABANDON record (or a correction of an abandoned tail) raised NEXT_ADJ to v"""
    import journal_oracle
    entries, _ = journal_oracle.read_journal(jdir)
    inst = sysv["INSTANCE_ID"]
    mx = 0
    for e in entries:
        if e["inst"] != inst:
            continue
        f = e["fields"]
        if e["type"] == "ABANDON" and f.get("NEXT_ADJ", "").isdigit():
            mx = max(mx, int(f["NEXT_ADJ"]))
        if e["type"] in ro.ADJUST_TYPES and f.get("ADJ", "").isdigit():
            mx = max(mx, int(f["ADJ"]) + 1)
    return v == mx
