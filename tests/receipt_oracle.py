"""Independent check of Core Phase 3 «Приход и заказы» (written separately from the Basic code on purpose).

check(doc, jdir, initial, …) replays the journal (receipts, their corrections and storno, order cancellations, the issues
of Phase 2) from the initial balances of the test registry (or from a trusted snapshot of the book, `base`) and compares
the result with every view of the book:
  «Наличие»   — a row per EI, balance = replayed balance; for an EI created by a receipt: name, article, unit, place,
                state (Активен / Приход удалён (сторно)), source «Заказ <№>»;
  _RCV        — one record per receipt EI: position, quantity, state, kind (source / additional row), documents flag;
  _ORD        — one record per order position: ordered, received (sum of live receipts), count, without documents, cancel;
  «Заказы»    — every receipt EI in exactly one row (copies aside); F, H, V, X, W of each row; the status of every position
                computed here from the rules of the task and D-047 (not taken from the book); a partly received
                position shows on its source row (Y) how much is still to come; H only on the source row; cancelled
                positions without EI; open rows show «Ожидается»/«Просрочено» by Q;
  «Выдачи»    — posted / cancelled issues against the journal (as tests/issue_oracle.py);
  counters    — NEXT_EI above every EI, NEXT_OL above every position, no EI created twice.
The ten invariants of the task are checked explicitly (see INVARIANTS).
"""
import datetime

import journal_oracle

RECEIPT_TYPES = ("RECEIPT", "RECEIPT_ADD", "RECEIPT_FIX", "RECEIPT_DEL", "ORDER_CANCEL", "ORDER_CANCEL_REST")
ISSUE_TYPES = ("ISSUE", "ISSUE_FIX", "ISSUE_DEL")
EPS = 1e-6
NULL_DATE = datetime.date(1899, 12, 30)
ORDER_COLS = [chr(65 + i) for i in range(26)] + ["AA", "AB"]
C = {c: i for i, c in enumerate(ORDER_COLS)}

ST_WAITING, ST_OVERDUE, ST_PARTIAL, ST_RECEIVED = "Ожидается", "Просрочено", "Частично получено", "Получено"
ST_NODOCS, ST_CANCELLED, ST_REST = "Получено без документов", "Отменено", "Частично получено / остаток отменён"
ST_ADD, ST_ADD_STORNO = "Дополнительное поступление", "Поступление удалено (сторно)"
ST_PARTIAL_OVERDUE = "Частично получено / просрочено"

INVARIANTS = {
    1: "один ЕИ — одна партия обычного прихода",
    2: "ЕИ никогда не переиспользуется",
    3: "«Наличие» соответствует проведённым движениям",
    4: "сумма приходов позиции считается без H дополнительных строк",
    5: "повторное проведение идемпотентно (ни одной строки дважды)",
    6: "нет отрицательных остатков",
    7: "приход не уменьшен ниже выданного",
    8: "отменённый остаток не считается ожидаемым",
    9: "отменённый заказ не создаёт ЕИ",
    10: "дополнительные УПД не задваивают заказанное количество",
}


def serial(d):
    return (d - NULL_DATE).days


def iso_serial(s):
    return serial(datetime.date.fromisoformat(s)) if s else None


def today_serial(today=None):
    if today is None:
        return serial(datetime.date.today())
    if isinstance(today, str):
        dd, mm, yy = (int(x) for x in today.split("."))
        return serial(datetime.date(yy, mm, dd))
    return int(today)


def ei_num(canon):
    return int(canon[3:]) if isinstance(canon, str) and canon.startswith("ЕИ-") and canon[3:].isdigit() and len(canon) == 11 else None


def canon(n):
    return f"ЕИ-{n:08d}"


def used_rows(sheet):
    cur = sheet.createCursor()
    cur.gotoEndOfUsedArea(False)
    return cur.getRangeAddress().EndRow


def read(doc, name, ncols):
    sh = doc.Sheets.getByName(name)
    last = max(used_rows(sh), 1)
    return sh.getCellRangeByPosition(0, 0, ncols - 1, last).getDataArray()


def num(x):
    return x if isinstance(x, float) else None


DATE_MIN, DATE_MAX = serial(datetime.date(2000, 1, 1)), serial(datetime.date(2099, 12, 31))


def position_status(ordq, rcv, nodoc, cancel, edate, today):
    """the business status of an order position — the rules of the task and D-047, written independently:
    the expected date Q is a valid date (2000–2099) before today → nothing received «Просрочено», partly received
    «Частично получено / просрочено»; a cancelled rest and a fully received position are never overdue"""
    overdue = edate is not None and DATE_MIN <= int(edate) <= DATE_MAX and int(edate) < today
    if cancel == "ORDER":
        return ST_CANCELLED
    if rcv <= EPS:
        if cancel == "REST":
            return ST_CANCELLED
        return ST_OVERDUE if overdue else ST_WAITING
    if rcv >= ordq - EPS:
        return ST_NODOCS if nodoc > 0 else ST_RECEIVED
    if cancel == "REST":
        return ST_REST
    return ST_PARTIAL_OVERDUE if overdue else ST_PARTIAL


def qty_text(x):
    """a quantity as WMS shows it in texts: whole number or up to 3 decimals after a comma"""
    ip, fr = divmod(int(round(x * 1000)), 1000)
    return str(ip) + ("," + f"{fr:03d}".rstrip("0") if fr else "")


def fp_part(v):
    if v == "" or v is None:
        return "E"
    if isinstance(v, float):
        s = repr(v)
        return "N" + (s[:-2] if s.endswith(".0") else s)
    t = str(v).strip()
    out = []
    for ch in t:
        o = ord(ch)
        out.append(f"%{o:02X}" if o < 32 or o == 127 or ch in "%;|=~" else ch)
    return "S" + "".join(out)


def fingerprint(row):
    return "|".join(fp_part(row[C[c]]) for c in "A D B E H I L P".split())


def snapshot_base(doc, initial):
    """the state of a book taken as trusted (benchmark: a generated history without a journal)"""
    base = dict(bal=dict(initial), receipts={}, positions={}, issues={}, cancelled_rows=0)
    stock = read(doc, "Наличие", 9)
    for n, row in enumerate(stock):
        if n and row[0]:
            base["bal"][row[0]] = row[4]
    rcv = read(doc, "_RCV", 9)
    for n, row in enumerate(rcv):
        if n and row[0]:
            base["receipts"][row[0]] = dict(ol=int(row[1]), qty=row[6], state=row[4], kind=row[5], nodoc=int(row[7] or 0), order=None)
    ordd = read(doc, "_ORD", 9)
    for n, row in enumerate(ordd):
        if n and isinstance(row[0], float):
            base["positions"][int(row[0])] = dict(ord=row[4], rcv=row[5], cnt=int(row[6]), nodoc=int(row[7]), cancel=row[8], key=row[2], fp=row[3])
    iss = read(doc, "Выдачи", 18)
    for r, row in enumerate(iss):
        if r and isinstance(row[0], float) and not str(row[17]).startswith("КОПИЯ"):
            base["issues"][int(row[0])] = dict(ei=row[11], qty=row[4], deleted=row[17] == "Удалено (сторно)")
    return base


def check(doc, jdir, initial, expect_tail=0, today=None, base=None):
    P = []
    inv_fail = set()

    def bad(msg, inv=None):
        P.append(msg)
        if inv:
            inv_fail.add(inv)

    tday = today_serial(today)
    entries, info = journal_oracle.read_journal(jdir)
    sysv = journal_oracle.sys_values(doc)
    inst = sysv["INSTANCE_ID"]
    ours = [e for e in entries if e["inst"] == inst]
    if len(ours) != len(entries):
        bad(f"строк другого экземпляра в журнале: {len(entries) - len(ours)}")
    seqs = [e["seq"] for e in ours]
    if seqs and seqs != list(range(seqs[0], seqs[0] + len(seqs))):
        bad("журнал: seq не непрерывны")
    last = int(sysv["LAST_SEQ"])
    mx = max(seqs) if seqs else 0
    if mx - last != expect_tail:
        bad(f"LAST_SEQ книги {last}, последняя запись журнала {mx}, ожидался хвост {expect_tail}")
    if sysv["TX_STATE"] not in ("COMMITTED", "NONE"):
        bad(f"маркер {sysv['TX_STATE']}")
    if sysv["TX_BEFORE_IMAGE"]:
        bad("снимок до изменения не очищен")
    applied = [e for e in ours if e["seq"] <= last]
    abandoned = journal_oracle.abandoned_seqs(applied)
    ops = [e for e in applied if e["seq"] not in abandoned and e["type"] in RECEIPT_TYPES + ISSUE_TYPES]
    min_next_ei = 1
    for e in applied:
        if e["type"] == "ABANDON" and e["fields"].get("NEXT_EI"):
            min_next_ei = max(min_next_ei, int(e["fields"]["NEXT_EI"]))

    # ------------------------------------------------------------ 1. replay
    if base is None:
        bal = dict(initial)
        receipts, positions, issues = {}, {}, {}
    else:
        bal = dict(base["bal"])
        receipts = {k: dict(v) for k, v in base["receipts"].items()}
        positions = {k: dict(v) for k, v in base["positions"].items()}
        issues = {k: dict(v) for k, v in base["issues"].items()}
    created = []                                        # every EI a receipt created, in order
    n_rcpt_ops = 0

    def f_num(f, k):
        return float(f[k]) if f.get(k, "") != "" else None

    for e in ops:
        f, t, sq = e["fields"], e["type"], e["seq"]
        if t in ("RECEIPT", "RECEIPT_ADD"):
            n_rcpt_ops += 1
            ei, ol, q = f["EI"], int(f["OL"]), float(f["QTY"])
            if ei in bal or ei in receipts:
                bad(f"seq {sq} {t}: {ei} уже существовал — ЕИ выдан повторно", 2)
            created.append(ei)
            nodoc = 0 if f.get("DOC", "").strip() and f.get("DOC_DATE") else 1
            rec = dict(ol=ol, qty=q, state="LIVE", kind="SRC" if t == "RECEIPT" else "ADD", nodoc=nodoc, name=f.get("NAME"),
                       art=f.get("ART"), unit=f.get("UNIT"), place=f.get("PLACE"), order=f.get("ORDER"), doc=f.get("DOC", ""),
                       ddate=iso_serial(f.get("DOC_DATE")), rdate=iso_serial(f.get("DATE")), docqty=f_num(f, "DOC_QTY"),
                       price=f_num(f, "PRICE"), supplier=f.get("SUPPLIER"))
            receipts[ei] = rec
            bal[ei] = q
            if abs(float(f["BAL_AFTER"]) - q) > EPS:
                bad(f"seq {sq} {t}: BAL_AFTER {f['BAL_AFTER']} ≠ {q}", 3)
            if t == "RECEIPT":
                if ol in positions:
                    bad(f"seq {sq} RECEIPT: позиция OLID {ol} уже существовала")
                positions[ol] = dict(ord=float(f["ORD_QTY"]), rcv=q, cnt=1, nodoc=nodoc, cancel="", key=ei)
            else:
                p = positions.get(ol)
                if p is None:
                    bad(f"seq {sq} RECEIPT_ADD: нет позиции OLID {ol}")
                    continue
                if p["cancel"] == "ORDER":
                    bad(f"seq {sq} RECEIPT_ADD: поступление по отменённой позиции OLID {ol}", 9)
                p["rcv"] += q
                p["cnt"] += 1
                p["nodoc"] += nodoc
                if abs(float(f["POS_RCV"]) - p["rcv"]) > EPS:
                    bad(f"seq {sq} RECEIPT_ADD: POS_RCV {f['POS_RCV']} ≠ расчётное {p['rcv']}", 4)
        elif t == "RECEIPT_FIX":
            n_rcpt_ops += 1
            ei = f["EI"]
            rec = receipts.get(ei)
            if rec is None or rec["state"] != "LIVE":
                bad(f"seq {sq} RECEIPT_FIX: нет действующего прихода {ei}")
                continue
            old, q = float(f["OLD_QTY"]), float(f["QTY"])
            if abs(old - rec["qty"]) > EPS:
                bad(f"seq {sq} RECEIPT_FIX {ei}: OLD_QTY {old} ≠ расчётный приход {rec['qty']}")
            issued = rec["qty"] - bal[ei]
            if q < issued - EPS:
                bad(f"seq {sq} RECEIPT_FIX {ei}: приход {q} меньше выданного {issued}", 7)
            if abs(float(f["BAL_BEFORE"]) - bal[ei]) > EPS:
                bad(f"seq {sq} RECEIPT_FIX {ei}: BAL_BEFORE {f['BAL_BEFORE']} ≠ {bal[ei]}", 3)
            bal[ei] += q - rec["qty"]
            nodoc = 0 if f.get("DOC", "").strip() and f.get("DOC_DATE") else 1
            p = positions[rec["ol"]]
            p["rcv"] += q - rec["qty"]
            p["nodoc"] += nodoc - rec["nodoc"]
            rec.update(qty=q, nodoc=nodoc, doc=f.get("DOC", ""), ddate=iso_serial(f.get("DOC_DATE")), rdate=iso_serial(f.get("DATE")),
                       docqty=f_num(f, "DOC_QTY"), price=f_num(f, "PRICE"), place=f.get("PLACE"))
            if abs(float(f["BAL_AFTER"]) - bal[ei]) > EPS:
                bad(f"seq {sq} RECEIPT_FIX {ei}: BAL_AFTER {f['BAL_AFTER']} ≠ {bal[ei]}", 3)
            if abs(float(f["POS_RCV"]) - p["rcv"]) > EPS:
                bad(f"seq {sq} RECEIPT_FIX {ei}: POS_RCV {f['POS_RCV']} ≠ {p['rcv']}", 4)
        elif t == "RECEIPT_DEL":
            n_rcpt_ops += 1
            ei = f["EI"]
            rec = receipts.get(ei)
            if rec is None or rec["state"] != "LIVE":
                bad(f"seq {sq} RECEIPT_DEL: нет действующего прихода {ei}")
                continue
            if bal[ei] < rec["qty"] - EPS:
                bad(f"seq {sq} RECEIPT_DEL {ei}: сторно при выданных {rec['qty'] - bal[ei]} — остаток стал бы отрицательным", 6)
            bal[ei] -= rec["qty"]
            rec["state"] = "STORNO"
            p = positions[rec["ol"]]
            p["rcv"] -= rec["qty"]
            p["cnt"] -= 1
            p["nodoc"] -= rec["nodoc"]
            if abs(float(f["BAL_AFTER"]) - bal[ei]) > EPS:
                bad(f"seq {sq} RECEIPT_DEL {ei}: BAL_AFTER {f['BAL_AFTER']} ≠ {bal[ei]}", 3)
        elif t == "ORDER_CANCEL":
            n_rcpt_ops += 1
            ol = int(f["OL"])
            p = positions.get(ol)
            if p is None:
                positions[ol] = dict(ord=float(f["ORD_QTY"]), rcv=0.0, cnt=0, nodoc=0, cancel="ORDER", key="")
            else:
                if p["cnt"] or p["rcv"] > EPS:
                    bad(f"seq {sq} ORDER_CANCEL: у позиции OLID {ol} есть действующие приходы", 9)
                p["cancel"] = "ORDER"
        elif t == "ORDER_CANCEL_REST":
            n_rcpt_ops += 1
            ol = int(f["OL"])
            p = positions.get(ol)
            if p is None or p["cancel"] or not (EPS < p["rcv"] < p["ord"] - EPS):
                bad(f"seq {sq} ORDER_CANCEL_REST: позиция OLID {ol} не частично получена ({p})", 8)
                continue
            p["cancel"] = "REST"
        else:
            # Phase 2 issues
            no = int(float(f["NO"]))
            if t == "ISSUE":
                if no in issues:
                    bad(f"№ {no} выдан в журнале повторно (seq {sq})", 5)
                    continue
                ei, q = f["EI"], float(f["QTY"])
                if abs(bal.get(ei, float("nan")) - float(f["BAL_BEFORE"])) > EPS:
                    bad(f"seq {sq} ISSUE №{no}: BAL_BEFORE {f['BAL_BEFORE']} ≠ расчётный {bal.get(ei)}", 3)
                bal[ei] = bal.get(ei, 0.0) - q
                issues[no] = dict(ei=ei, qty=q, deleted=False, who=f.get("WHO"), date=f.get("DATE"))
            else:
                prev = issues.get(no)
                if prev is None or prev["deleted"]:
                    bad(f"seq {sq} {t} №{no}: нет действующей выдачи")
                    continue
                bal[prev["ei"]] += prev["qty"]
                if t == "ISSUE_FIX":
                    ei, q = f["EI"], float(f["QTY"])
                    if abs(bal.get(ei, float("nan")) - float(f["BAL_BEFORE"])) > EPS:
                        bad(f"seq {sq} ISSUE_FIX №{no}: BAL_BEFORE {f['BAL_BEFORE']} ≠ {bal.get(ei)}", 3)
                    bal[ei] = bal.get(ei, 0.0) - q
                    issues[no] = dict(ei=ei, qty=q, deleted=False, who=f.get("WHO"), date=f.get("DATE"))
                else:
                    prev["deleted"] = True
            ei = f["EI"]
            if abs(bal.get(ei, 0.0) - float(f["BAL_AFTER"])) > EPS and t != "ISSUE_DEL":
                bad(f"seq {sq} {t} №{no}: BAL_AFTER {f['BAL_AFTER']} ≠ {bal.get(ei)}", 3)
        for k in list(bal):
            if bal[k] < -EPS:
                bad(f"seq {sq}: отрицательный остаток {k} = {bal[k]}", 6)
                bal[k] = round(bal[k], 6)
    if len(set(created)) != len(created):
        bad("один и тот же ЕИ создан приходом дважды", 2)

    # ------------------------------------------------------------ 2. «Наличие» and _RCV
    stock = read(doc, "Наличие", 9)
    book_ei = {}
    for n, row in enumerate(stock):
        if n == 0 or not row[0]:
            continue
        if row[0] != canon(n):
            bad(f"«Наличие» строка {n + 1}: «{row[0]}» вместо {canon(n)} (адресация)", 3)
            continue
        book_ei[row[0]] = row
    for ei, b in bal.items():
        row = book_ei.get(ei)
        if row is None:
            bad(f"{ei}: нет строки в «Наличие»", 3)
            continue
        if not isinstance(row[4], float) or abs(row[4] - b) > EPS:
            bad(f"{ei}: в «Наличие» {row[4]!r}, по журналу {b}", 3)
        if isinstance(row[4], float) and row[4] < -EPS:
            bad(f"{ei}: отрицательный остаток {row[4]}", 6)
    for ei in book_ei:
        if ei not in bal:
            bad(f"{ei} есть в «Наличие», но не создан ни начальным реестром, ни приходом", 1)
    rcv = read(doc, "_RCV", 9)
    book_rcv = {}
    for n, row in enumerate(rcv):
        if n and row[0]:
            if row[0] != canon(n):
                bad(f"_RCV строка {n + 1}: «{row[0]}» вместо {canon(n)}")
            book_rcv[row[0]] = row
    for ei, rec in receipts.items():
        row = book_rcv.get(ei)
        if row is None:
            bad(f"{ei}: нет записи прихода в _RCV", 1)
            continue
        want = (float(rec["ol"]), rec["state"], rec["kind"], rec["qty"], float(rec["nodoc"]))
        got = (row[1], row[4], row[5], row[6], row[7])
        if got[:3] != want[:3] or abs((got[3] or 0) - want[3]) > EPS or got[4] != want[4]:
            bad(f"{ei}: _RCV {got} ≠ расчёт {want}")
        if rec["state"] == "LIVE" and not isinstance(row[3], float):
            bad(f"{ei}: у действующего прихода нет ключа антидубля")
        if rec["state"] == "STORNO" and (row[3] != "" or row[8] != ""):
            bad(f"{ei}: у сторнированного прихода остался ключ антидубля")
        srow = book_ei.get(ei)
        if srow is not None and rec.get("name") is not None:
            want_s = (rec["name"], rec["art"] or "", rec["unit"], rec["place"], "Активен" if rec["state"] == "LIVE" else "Приход удалён (сторно)",
                      f"Заказ {rec['order']}")
            got_s = (srow[1], srow[2], srow[3], srow[5], srow[7], srow[8])
            if got_s != want_s:
                bad(f"{ei}: «Наличие» {got_s} ≠ ожидаемое {want_s}", 3)
    for ei in book_rcv:
        if ei not in receipts:
            bad(f"_RCV: запись {ei} без операции прихода в журнале")

    # ------------------------------------------------------------ 3. _ORD
    ordd = read(doc, "_ORD", 12)
    next_ol = int(ordd[0][11])
    book_pos = {}
    for n, row in enumerate(ordd):
        if n and isinstance(row[0], float):
            if int(row[0]) != n:
                bad(f"_ORD строка {n + 1}: OLID {row[0]} (адресация)")
            book_pos[n] = row
    for ol, p in positions.items():
        row = book_pos.get(ol)
        if row is None:
            bad(f"позиция OLID {ol}: нет записи в _ORD")
            continue
        got = (row[4], row[5], int(row[6]), int(row[7]), row[8], row[2])
        want = (p["ord"], p["rcv"], p["cnt"], p["nodoc"], p["cancel"], p["key"])
        if abs(got[0] - want[0]) > EPS or abs(got[1] - want[1]) > EPS or got[2:] != want[2:]:
            bad(f"позиция OLID {ol}: _ORD {got} ≠ расчёт {want}", 4)
        if p["cancel"] == "ORDER" and (p["cnt"] or p["rcv"] > EPS):
            bad(f"позиция OLID {ol}: отменена, но имеет приходы", 9)
    for ol in book_pos:
        if ol not in positions:
            bad(f"_ORD: позиция OLID {ol} без операции в журнале")
    if positions and next_ol <= max(positions):
        bad(f"NEXT_OL {next_ol} не больше максимального OLID {max(positions)}")

    # ------------------------------------------------------------ 4. «Заказы»
    orders = read(doc, "Заказы", 28)
    ei_rows, copies, cancelled_rows, open_rows = {}, 0, [], 0
    for r, row in enumerate(orders):
        if r == 0:
            continue
        v, w, y = row[C["V"]], row[C["W"]], str(row[C["Y"]])
        if y.startswith("КОПИЯ"):
            copies += 1
            continue
        if v:
            if ei_num(v) is None:
                bad(f"«Заказы» строка {r + 1}: «{v}» в V — не ЕИ")
                continue
            if v in ei_rows:
                bad(f"«Заказы»: {v} в двух строках без отметки КОПИЯ: {ei_rows[v] + 1} и {r + 1}", 1)
                continue
            ei_rows[v] = r
        elif w == ST_CANCELLED:
            cancelled_rows.append(r)
        elif any(str(x) != "" for x in row[:21]):
            open_rows += 1
            if w not in ("", ST_WAITING, ST_OVERDUE):
                bad(f"«Заказы» строка {r + 1}: непроведённая строка со статусом «{w}»")
            q = num(row[C["Q"]])
            want = ST_OVERDUE if (q is not None and DATE_MIN <= int(q) <= DATE_MAX and int(q) < tday) else ST_WAITING
            if w and w != want:
                bad(f"«Заказы» строка {r + 1}: статус «{w}», по дате ожидается «{want}»")
            if isinstance(row[C["X"]], float):
                bad(f"«Заказы» строка {r + 1}: у непроведённой строки заполнено «Наличие»")
    for ei, rec in receipts.items():
        r = ei_rows.get(ei)
        if r is None:
            bad(f"{ei}: нет строки прихода на «Заказы»", 1)
            continue
        row = orders[r]
        p = positions.get(rec["ol"])
        if not isinstance(row[C["F"]], float) or abs(row[C["F"]] - rec["qty"]) > EPS:
            bad(f"{ei} строка {r + 1}: F {row[C['F']]!r} ≠ приход {rec['qty']}")
        if not isinstance(row[C["X"]], float) or abs(row[C["X"]] - bal[ei]) > EPS:
            bad(f"{ei} строка {r + 1}: X {row[C['X']]!r} ≠ остаток {bal[ei]}", 3)
        if rec.get("rdate") is not None and row[C["N"]] != float(rec["rdate"]):
            bad(f"{ei} строка {r + 1}: N {row[C['N']]!r} ≠ дата поступления {rec['rdate']}")
        if rec.get("place") is not None and row[C["U"]] != rec["place"]:
            bad(f"{ei} строка {r + 1}: U {row[C['U']]!r} ≠ место {rec['place']}")
        if rec.get("doc") is not None and str(row[C["C"]]).strip() != rec["doc"]:
            bad(f"{ei} строка {r + 1}: C {row[C['C']]!r} ≠ документ {rec['doc']!r}")
        if rec["kind"] == "SRC":
            if p is not None and (not isinstance(row[C["H"]], float) or abs(row[C["H"]] - p["ord"]) > EPS):
                bad(f"{ei} строка {r + 1}: H исходной строки {row[C['H']]!r} ≠ заказано {p['ord']}", 10)
        else:
            if row[C["H"]] != "":
                bad(f"{ei} строка {r + 1}: у дополнительной строки заполнено H {row[C['H']]!r}", 10)
            want_w = ST_ADD if rec["state"] == "LIVE" else ST_ADD_STORNO
            if row[C["W"]] != want_w:
                bad(f"{ei} строка {r + 1}: W «{row[C['W']]}» ≠ «{want_w}»")
            if p is not None and p.get("key") in ei_rows:
                src = orders[ei_rows[p["key"]]]
                for col in "A B D E I L".split():
                    if src[C[col]] != row[C[col]]:
                        bad(f"{ei} строка {r + 1}: {col} {row[C[col]]!r} ≠ исходная строка {src[C[col]]!r}", 10)
    # every position: its status on its source row, H counted once
    cancelled_fp = {}
    for r in cancelled_rows:
        cancelled_fp.setdefault(fingerprint(orders[r]), []).append(r)
    for ol, p in positions.items():
        if p.get("key"):
            r = ei_rows.get(p["key"])
            if r is None:
                bad(f"позиция OLID {ol}: исходная строка ({p['key']}) не найдена")
                continue
        else:
            fp = book_pos.get(ol, [None] * 9)[3]
            cand = cancelled_fp.get(fp, [])
            if len(cand) != 1:
                bad(f"позиция OLID {ol} (отменена без прихода): строк с её отпечатком {len(cand)}", 9)
                continue
            r = cand[0]
        row = orders[r]
        want = position_status(p["ord"], p["rcv"], p["nodoc"], p["cancel"], num(row[C["Q"]]), tday)
        if row[C["W"]] != want:
            bad(f"позиция OLID {ol} строка {r + 1}: статус «{row[C['W']]}», по правилам «{want}»" + (" (инвариант 8)" if p["cancel"] == "REST" else ""),
                8 if p["cancel"] == "REST" else None)
        if p["cancel"] == "REST" and row[C["W"]] in (ST_WAITING, ST_OVERDUE, ST_PARTIAL, ST_PARTIAL_OVERDUE):
            bad(f"позиция OLID {ol}: остаток отменён, но позиция показана ожидаемой", 8)
        if want in (ST_PARTIAL, ST_PARTIAL_OVERDUE):
            rest = f"осталось {qty_text(p['ord'] - p['rcv'])}"
            if rest not in str(row[C["Y"]]):
                bad(f"позиция OLID {ol} строка {r + 1}: «{want}», но Y не показывает «{rest}»: {row[C['Y']]!r}")
        if p["cancel"] == "ORDER" and not p.get("key") and row[C["V"]]:
            bad(f"позиция OLID {ol}: отменённый заказ получил ЕИ", 9)
    if len(cancelled_rows) != sum(1 for p in positions.values() if p["cancel"] == "ORDER" and not p.get("key")):
        bad(f"строк «Отменено» без ЕИ {len(cancelled_rows)}, отменённых позиций без прихода "
            f"{sum(1 for p in positions.values() if p['cancel'] == 'ORDER' and not p.get('key'))}", 9)

    # ------------------------------------------------------------ 5. «Выдачи» (Phase 2)
    iss = read(doc, "Выдачи", 18)
    seen = {}
    for r, row in enumerate(iss):
        if r == 0 or not isinstance(row[0], float) or str(row[17]).startswith("КОПИЯ"):
            continue
        k = int(row[0])
        if k in seen:
            bad(f"«Выдачи»: № {k} в двух строках", 5)
            continue
        seen[k] = r
        st = issues.get(k)
        if st is None:
            bad(f"«Выдачи» строка {r + 1}: № {k} без операции в журнале")
            continue
        if st["deleted"] != (row[17] == "Удалено (сторно)"):
            bad(f"«Выдачи» строка {r + 1} №{k}: статус «{row[17]}» не совпадает с журналом")
        if row[11] != st["ei"] or not isinstance(row[4], float) or abs(row[4] - st["qty"]) > EPS:
            bad(f"«Выдачи» строка {r + 1} №{k}: {row[11]} {row[4]!r} ≠ журнал {st['ei']} {st['qty']}")
    for k in issues:
        if k not in seen:
            bad(f"выдача № {k} есть в журнале, но строки в книге нет")

    # ------------------------------------------------------------ 6. counters
    next_ei = int(sysv["NEXT_EI"])
    all_ei = [ei_num(x) for x in list(bal) + list(book_rcv) + list(ei_rows)]
    top = max([x for x in all_ei if x] or [0])
    if next_ei <= top:
        bad(f"NEXT_EI {next_ei} не больше максимального ЕИ {top}", 2)
    if next_ei < min_next_ei:
        bad(f"NEXT_EI {next_ei} меньше значения из записи ABANDON {min_next_ei}", 2)
    if created and base is None and next_ei != max(max(ei_num(x) for x in created) + 1, min_next_ei, len(initial) + 1):
        bad(f"NEXT_EI {next_ei}: ожидалось {max(ei_num(x) for x in created) + 1}")

    info.update(dict(ops=len(ops), receipt_ops=n_rcpt_ops, receipts=len(receipts), live=sum(1 for x in receipts.values() if x["state"] == "LIVE"),
                     positions=len(positions), open_rows=open_rows, copies=copies, last_seq=last, journal_max=mx,
                     invariants_failed=sorted(inv_fail)))
    return P, info
