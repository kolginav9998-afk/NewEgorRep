"""Independent check of the issue scenario (written separately from the Basic code on purpose).

check(doc, jdir, initial) compares three views of the same history and reports every difference:
  1. the journal replayed from the initial balances (ISSUE −q; ISSUE_FIX returns the old movement and applies the new
     one; ISSUE_DEL returns the movement) — every BAL_BEFORE/BAL_AFTER must match the running balance;
  2. the «Выдачи» sheet: posted rows (R «Проведено…»), cancelled rows (R «Удалено (сторно)»), copies (R «КОПИЯ…»);
  3. «Наличие»: initial − Σ E of the posted rows of each EI must equal the registry balance, which must also equal the
     replayed journal balance.
No double issue, no negative balance, no repeated №, no journal operation without its row, no row without its operation.
"""
import datetime

import journal_oracle

ISSUE_TYPES = ("ISSUE", "ISSUE_FIX", "ISSUE_DEL")
ST_POSTED = "Проведено"
ST_DELETED = "Удалено (сторно)"
EPS = 1e-6


def serial_to_iso(x):
    return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(x))).isoformat()


def used_rows(sheet):
    cur = sheet.createCursor()
    cur.gotoEndOfUsedArea(False)
    return cur.getRangeAddress().EndRow


def read_issues(doc):
    sh = doc.Sheets.getByName("Выдачи")
    last = max(used_rows(sh), 1)
    return sh.getCellRangeByPosition(0, 0, 17, last).getDataArray()


def read_stock(doc):
    sh = doc.Sheets.getByName("Наличие")
    last = max(used_rows(sh), 1)
    out = {}
    for row in sh.getCellRangeByPosition(0, 1, 8, last).getDataArray():
        if row[0]:
            out[row[0]] = row[4] if isinstance(row[4], float) else None
    return out


def check(doc, jdir, initial, expect_tail=0):
    P = []
    entries, info = journal_oracle.read_journal(jdir)
    sysv = journal_oracle.sys_values(doc)
    inst = sysv["INSTANCE_ID"]
    ours = [e for e in entries if e["inst"] == inst]
    if len(ours) != len(entries):
        P.append(f"строк другого экземпляра в журнале: {len(entries) - len(ours)}")
    seqs = [e["seq"] for e in ours]
    if seqs and seqs != list(range(seqs[0], seqs[0] + len(seqs))):
        P.append("журнал: seq не непрерывны")
    last = int(sysv["LAST_SEQ"])
    mx = max(seqs) if seqs else 0
    if mx - last != expect_tail:
        P.append(f"LAST_SEQ книги {last}, последняя запись журнала {mx}, ожидался хвост {expect_tail}")
    if sysv["TX_STATE"] not in ("COMMITTED", "NONE"):
        P.append(f"маркер {sysv['TX_STATE']}")
    if sysv["TX_BEFORE_IMAGE"]:
        P.append("снимок до изменения не очищен")
    applied = [e for e in ours if e["seq"] <= last]
    abandoned = journal_oracle.abandoned_seqs(applied)
    ops = [e for e in applied if e["type"] in ISSUE_TYPES and e["seq"] not in abandoned]
    # 1. journal replay
    bal = dict(initial)
    state = {}
    for e in ops:
        f = e["fields"]
        no = int(float(f["NO"]))
        if e["type"] == "ISSUE":
            if no in state:
                P.append(f"№ {no} выдан в журнале повторно (seq {e['seq']})")
                continue
            ei, q = f["EI"], float(f["QTY"])
            if abs(bal.get(ei, float("nan")) - float(f["BAL_BEFORE"])) > EPS:
                P.append(f"seq {e['seq']} ISSUE №{no}: BAL_BEFORE {f['BAL_BEFORE']} ≠ расчётный остаток {bal.get(ei)} ({ei})")
            bal[ei] = bal.get(ei, 0.0) - q
            state[no] = dict(ei=ei, qty=q, who=f.get("WHO"), date=f.get("DATE"), deleted=False, fixed=False)
        else:
            prev = state.get(no)
            if prev is None or prev["deleted"]:
                P.append(f"seq {e['seq']} {e['type']} №{no}: нет действующей выдачи")
                continue
            if e["type"] == "ISSUE_FIX":
                if f["OLD_EI"] != prev["ei"] or abs(float(f["OLD_QTY"]) - prev["qty"]) > EPS:
                    P.append(f"seq {e['seq']} ISSUE_FIX №{no}: OLD_* не совпадает с действующей выдачей")
                bal[prev["ei"]] += prev["qty"]
                ei, q = f["EI"], float(f["QTY"])
                if abs(bal.get(ei, float("nan")) - float(f["BAL_BEFORE"])) > EPS:
                    P.append(f"seq {e['seq']} ISSUE_FIX №{no}: BAL_BEFORE {f['BAL_BEFORE']} ≠ расчётный {bal.get(ei)}")
                bal[ei] = bal.get(ei, 0.0) - q
                state[no] = dict(ei=ei, qty=q, who=f.get("WHO"), date=f.get("DATE"), deleted=False, fixed=True)
            else:
                bal[prev["ei"]] += prev["qty"]
                if abs(bal[prev["ei"]] - float(f["BAL_AFTER"])) > EPS:
                    P.append(f"seq {e['seq']} ISSUE_DEL №{no}: BAL_AFTER {f['BAL_AFTER']} ≠ расчётный {bal[prev['ei']]}")
                prev["deleted"] = True
        ei = f["EI"]
        if bal.get(ei, 0.0) < -EPS:
            P.append(f"seq {e['seq']}: отрицательный остаток {ei} = {bal[ei]}")
        if abs(bal.get(ei, 0.0) - float(f["BAL_AFTER"])) > EPS and e["type"] != "ISSUE_DEL":
            P.append(f"seq {e['seq']} {e['type']} №{no}: BAL_AFTER {f['BAL_AFTER']} ≠ расчётный {bal.get(ei)}")
    # 2. the sheet against the journal
    rows = read_issues(doc)
    seen, copies, sheet_sum, nposted, ndeleted = {}, 0, {}, 0, 0
    for r, row in enumerate(rows):
        if r == 0 or not isinstance(row[0], float):
            continue
        ctl = row[17]
        if ctl.startswith("КОПИЯ"):
            copies += 1
            continue
        k = int(row[0])
        if k in seen:
            P.append(f"№ {k} в двух строках: {seen[k] + 1} и {r + 1}")
            continue
        seen[k] = r
        if ctl.startswith(ST_POSTED):
            nposted += 1
            sheet_sum[row[11]] = sheet_sum.get(row[11], 0.0) + (row[4] if isinstance(row[4], float) else float("nan"))
            if not (isinstance(row[15], float) and isinstance(row[16], float) and abs(row[15] - row[4] - row[16]) < EPS):
                P.append(f"строка {r + 1} №{k}: Q ≠ P − E ({row[15]!r} − {row[4]!r} ≠ {row[16]!r})")
        elif ctl == ST_DELETED:
            ndeleted += 1
        else:
            P.append(f"строка {r + 1}: № {k} без статуса проведения ({ctl!r})")
            continue
        st = state.get(k)
        if st is None:
            P.append(f"строка {r + 1}: № {k} есть в книге, но его операции нет в журнале")
            continue
        if st["deleted"] != (ctl == ST_DELETED):
            P.append(f"строка {r + 1} №{k}: статус «{ctl}», а по журналу {'удалена' if st['deleted'] else 'действует'}")
        if row[11] != st["ei"] or not isinstance(row[4], float) or abs(row[4] - st["qty"]) > EPS:
            P.append(f"строка {r + 1} №{k}: {row[11]} {row[4]!r} ≠ журнал {st['ei']} {st['qty']}")
        if row[8] != st["who"]:
            P.append(f"строка {r + 1} №{k}: получатель {row[8]!r} ≠ журнал {st['who']!r}")
        if not isinstance(row[7], float) or serial_to_iso(row[7]) != st["date"]:
            P.append(f"строка {r + 1} №{k}: дата {row[7]!r} ≠ журнал {st['date']}")
    for k in state:
        if k not in seen:
            P.append(f"операция № {k} есть в журнале, но строки с этим № в книге нет")
    # 3. balances: registry = initial − Σ posted rows = journal replay; nothing negative
    book = read_stock(doc)
    for ei, q0 in initial.items():
        exp_sheet = q0 - sheet_sum.get(ei, 0.0)
        got = book.get(ei)
        if got is None:
            P.append(f"{ei}: остаток в «Наличие» не число")
            continue
        if abs(got - exp_sheet) > EPS:
            P.append(f"{ei}: в «Наличие» {got}, по проведённым строкам ожидается {exp_sheet}")
        if abs(got - bal.get(ei, q0)) > EPS:
            P.append(f"{ei}: в «Наличие» {got}, по журналу {bal.get(ei, q0)}")
        if got < -EPS:
            P.append(f"{ei}: отрицательный остаток {got}")
    for ei in sheet_sum:
        if ei not in initial:
            P.append(f"проведена выдача по {ei}, которого нет в реестре")
    if seen and int(sysv["NEXT_NO"]) <= max(seen):
        P.append(f"NEXT_NO {sysv['NEXT_NO']} не больше максимального № {max(seen)}")
    info.update(dict(ops=len(ops), posted=nposted, deleted=ndeleted, copies=copies, last_seq=last, journal_max=mx))
    return P, info
