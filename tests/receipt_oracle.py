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

Core Phase 4: the returns (RETURN, RETURN_FIX, RETURN_DEL) are replayed here too — they change the balance and the
current place of their EI, and the rules of ISSUE_FIX / ISSUE_DEL depend on them. With `out` (a dict) the replayed
state is handed to tests/return_oracle.py, which compares the returns with «Возврат», _RET and _ISS.

Core Phase 5: the special receipts (SP_RECEIPT — a new EI of Офис / Производство / Детали / Старый склад / Иной,
SP_REFILL — a part added to the EI of its article, SP_FIX, SP_DEL, SP_IDENTIFY) are replayed by replay_special() with
the rules of the task written independently: one article of a part — one EI (the normalized article), a new EI never
reuses a number, the line № and the event № (OFF-/PROD-/DET-/OLD-/OTH-) are monotonic and never reused, joining only an
issued event of the same type, the unit of an EI never changes by a refill, a place changes only by a marked move, a
storno of the last live line of an EI is refused while it has live issues or returns (D-069), a storno of one addition
never makes the balance negative, identification keeps the EI. tests/special_oracle.py compares the replay with
«Иной приход», _SPR, _ART and the cards in «Наличие».

Final Core: the corrections (MOVE, WRITE_OFF, INV_ADJ and their *_FIX / *_DEL) are replayed by replay_adjust() with the
rules of the task written independently: a move keeps the quantity and names the place the EI really had (no silent
place change), a correction or storno of a move only while the EI is still where the move put it; a write-off never
more than the balance; an inventory correction records the difference against the balance the EI really had (BOOK),
never zero; a storno or correction never makes a balance negative; the № is monotonic and never reused; the storno of
the last receipt of an EI is refused while it has live corrections (D-069). tests/adjust_oracle.py compares the replay
with «Корректировки» and _ADJ.
"""
import datetime
import re

import journal_oracle

RECEIPT_TYPES = ("RECEIPT", "RECEIPT_ADD", "RECEIPT_FIX", "RECEIPT_DEL", "ORDER_CANCEL", "ORDER_CANCEL_REST", "LEGACY_RECEIPT", "LEGACY_RECEIPT_ADD")
# M7: a receipt of the old table transferred with its own EI and its current balance (≤ the received quantity: what was
# consumed before the transfer counts as issued); the stock row names the transfer («Перенос: заказ <A>»)
LEGACY_RECEIPTS = ("LEGACY_RECEIPT", "LEGACY_RECEIPT_ADD")
ISSUE_TYPES = ("ISSUE", "ISSUE_FIX", "ISSUE_DEL")
RETURN_TYPES = ("RETURN", "RETURN_FIX", "RETURN_DEL")
SPECIAL_TYPES = ("SP_RECEIPT", "SP_REFILL", "SP_FIX", "SP_DEL", "SP_IDENTIFY")
ADJUST_TYPES = ("MOVE", "MOVE_FIX", "MOVE_DEL", "WRITE_OFF", "WRITE_OFF_FIX", "WRITE_OFF_DEL", "INV_ADJ", "INV_ADJ_FIX", "INV_ADJ_DEL")
MIGRATE_TYPES = ("MIGRATE",)
MIGRATE_STYPES = ("Поставщик", "Офис", "Производство", "Детали", "Старый склад", "Иной приход")
SP_CODES = ("OFF", "PROD", "DET", "OLD", "OTH")
SP_NAME = {"OFF": "Офис", "PROD": "Производство", "DET": "Детали", "OLD": "Старый склад", "OTH": "Иной"}          # «Иной приход».B
SP_STYPE = {"OFF": "Офис", "PROD": "Производство", "DET": "Детали", "OLD": "Старый склад", "OTH": "Иной приход"}  # «Наличие».J
SP_COUNTER = {"OFF": "NEXT_OFF", "PROD": "NEXT_PROD", "DET": "NEXT_DET", "OLD": "NEXT_OLD", "OTH": "NEXT_OTH"}
IDENT_TYPES = ("Поставщик", "Офис", "Производство", "Детали", "Старый склад")
EI_ACTIVE, EI_STORNO, EI_REVIEW = "Активен", "Приход удалён (сторно)", "Требует разбора"
EVENT_RE = re.compile(r"(OFF|PROD|DET|OLD|OTH)-(\d{8})")
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
    base = dict(bal=dict(initial), receipts={}, positions={}, issues={}, cancelled_rows=0, place={})
    stock = read(doc, "Наличие", 9)
    for n, row in enumerate(stock):
        if n and row[0]:
            base["bal"][row[0]] = row[4]
            base["place"][row[0]] = row[5]
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
            base["issues"][int(row[0])] = dict(ei=row[11], qty=row[4], deleted=row[17] == "Удалено (сторно)", who=row[8],
                                               date=(NULL_DATE + datetime.timedelta(days=int(row[7]))).isoformat() if isinstance(row[7], float) else "")
    # Phase 4: the returns and their totals per issue, as the service tables hold them (taken as trusted)
    base["returns"], base["ret_sum"], base["ret_cnt"] = {}, {}, {}
    if doc.Sheets.hasByName("_RET"):
        for n, row in enumerate(read(doc, "_RET", 6)):
            if n and isinstance(row[0], float):
                base["returns"][n] = dict(issue=int(row[1]), ei=row[2], qty=row[3], state=row[4], date=None, place=None, who=None, fixed=False,
                                          L=None, M=None)
        for k, row in enumerate(read(doc, "_ISS", 4)):
            if k and isinstance(row[0], float):
                base["ret_sum"][k] = row[2]
                base["ret_cnt"][k] = int(row[3])
    # Phase 5: the special receipts (lines, index, cards) as the book holds them
    base["sp"] = snapshot_special(doc)
    return base


def check(doc, jdir, initial, expect_tail=0, today=None, base=None, out=None):
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
    ops = [e for e in applied if e["seq"] not in abandoned and e["type"] in RECEIPT_TYPES + ISSUE_TYPES + RETURN_TYPES + SPECIAL_TYPES + ADJUST_TYPES + MIGRATE_TYPES]
    min_next_ei = 1
    min_next_ret = 1
    for e in applied:
        if e["type"] == "ABANDON" and e["fields"].get("NEXT_EI"):
            min_next_ei = max(min_next_ei, int(e["fields"]["NEXT_EI"]))
        if e["type"] == "ABANDON" and e["fields"].get("NEXT_RET"):
            min_next_ret = max(min_next_ret, int(e["fields"]["NEXT_RET"]))
    # numbers an abandoned return took are never issued again either
    for e in applied:
        if e["seq"] in abandoned and e["type"] == "RETURN":
            min_next_ret = max(min_next_ret, int(e["fields"]["RET"]) + 1)
    # Phase 5: the line № and the event № of abandoned special receipts are never issued again; ABANDON raises the counters
    sp = new_special_state() if base is None or "sp" not in base else copy_special_state(base["sp"])
    for e in applied:
        fe = e["fields"]
        if e["seq"] in abandoned and e["type"] in ("SP_RECEIPT", "SP_REFILL"):
            if fe.get("SPL", "").isdigit():
                sp["abandoned_spl"].add(int(fe["SPL"]))
            m = EVENT_RE.fullmatch(fe.get("EVENT", ""))
            if m and fe.get("EVENT_NEW") == "1":
                sp["abandoned_ev"][m.group(1)].add(int(m.group(2)))
        if e["type"] == "ABANDON":
            if fe.get("NEXT_SPL", "").isdigit():
                sp["floor_spl"] = max(sp["floor_spl"], int(fe["NEXT_SPL"]))
            for code, key in SP_COUNTER.items():
                if fe.get(key, "").isdigit():
                    sp["floor_ev"][code] = max(sp["floor_ev"][code], int(fe[key]))

    # ------------------------------------------------------------ 1. replay
    if base is None:
        bal = dict(initial)
        receipts, positions, issues = {}, {}, {}
        place = {}
        returns, ret_sum, ret_cnt = {}, {}, {}
        adjs = {}
    else:
        bal = dict(base["bal"])
        receipts = {k: dict(v) for k, v in base["receipts"].items()}
        positions = {k: dict(v) for k, v in base["positions"].items()}
        issues = {k: dict(v) for k, v in base["issues"].items()}
        place = dict(base.get("place", {}))
        returns = {k: dict(v) for k, v in base.get("returns", {}).items()}
        ret_sum = dict(base.get("ret_sum", {}))
        ret_cnt = dict(base.get("ret_cnt", {}))
        adjs = {k: dict(v) for k, v in base.get("adjs", {}).items()}
    touched_place = set()                               # EIs whose current place a movement set (checked in «Наличие».F)
    n_ret_ops = 0
    created = []                                        # every EI a receipt created, in order
    n_rcpt_ops = 0

    def f_num(f, k):
        return float(f[k]) if f.get(k, "") != "" else None

    for e in ops:
        f, t, sq = e["fields"], e["type"], e["seq"]
        if t in ("RECEIPT", "RECEIPT_ADD") + LEGACY_RECEIPTS:
            n_rcpt_ops += 1
            legacy = t in LEGACY_RECEIPTS
            ei, ol, q = f["EI"], int(f["OL"]), float(f["QTY"])
            if ei in bal or ei in receipts:
                bad(f"seq {sq} {t}: {ei} уже существовал — ЕИ выдан повторно", 2)
            created.append(ei)
            nodoc = 0 if f.get("DOC", "").strip() and f.get("DOC_DATE") else 1
            rec = dict(ol=ol, qty=q, state="LIVE", kind="SRC" if t in ("RECEIPT", "LEGACY_RECEIPT") else "ADD", nodoc=nodoc, name=f.get("NAME"),
                       art=f.get("ART"), unit=f.get("UNIT"), place=f.get("PLACE"), order=f.get("ORDER"), doc=f.get("DOC", ""),
                       ddate=iso_serial(f.get("DOC_DATE")), rdate=iso_serial(f.get("DATE")), docqty=f_num(f, "DOC_QTY"),
                       price=f_num(f, "PRICE"), supplier=f.get("SUPPLIER"), legacy=legacy)
            receipts[ei] = rec
            if legacy:
                b = float(f["BAL_AFTER"])
                if b < -EPS or b > q + EPS:
                    bad(f"seq {sq} {t} {ei}: остаток переноса {b} вне 0…{q} (больше прихода не переносится)", 3)
                if not f.get("ORIGIN"):
                    bad(f"seq {sq} {t} {ei}: нет происхождения (ORIGIN)", "M")
                bal[ei] = b
                place[ei] = f.get("PLACE_NOW") or f.get("PLACE")
                rec["legacy_used"] = q - b          # consumed before the transfer: counts as issued in the history of the EI
            else:
                bal[ei] = q
                place[ei] = f.get("PLACE")
                if abs(float(f["BAL_AFTER"]) - q) > EPS:
                    bad(f"seq {sq} {t}: BAL_AFTER {f['BAL_AFTER']} ≠ {q}", 3)
            if t in ("RECEIPT", "LEGACY_RECEIPT"):
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
            if f.get("PLACE") != f.get("OLD_PLACE"):
                place[ei] = f.get("PLACE")          # the receipt's place was changed: it becomes the current place of the EI
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
            # D-069: no live issue and no live return of the EI may remain behind a cancelled receipt
            live_iss = [k for k, v in issues.items() if not v["deleted"] and v["ei"] == ei]
            live_ret = [n for n, v in returns.items() if v["state"] == "LIVE" and v["ei"] == ei]
            if live_iss or live_ret:
                bad(f"seq {sq} RECEIPT_DEL {ei}: сторно прихода при действующих выдачах {live_iss[:5]} / возвратах {live_ret[:5]} (D-069)", "R")
            live_adj = [n for n, v in adjs.items() if v["state"] == "LIVE" and v["ei"] == ei]
            if live_adj:
                bad(f"seq {sq} RECEIPT_DEL {ei}: сторно прихода при действующих корректировках {live_adj[:5]} (D-069)", "R")
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
        elif t in RETURN_TYPES:
            n_ret_ops += 1
            replay_return(t, f, sq, bal, issues, returns, ret_sum, ret_cnt, place, touched_place, bad)
        elif t in SPECIAL_TYPES:
            replay_special(t, f, sq, bal, sp, place, issues, returns, created, receipts, bad, adjs)
        elif t in ADJUST_TYPES:
            replay_adjust(t, f, sq, bal, adjs, place, touched_place, bad)
        elif t in MIGRATE_TYPES:
            replay_migrate(f, sq, bal, place, sp, created, bad)
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
                # Phase 4: the live returns of an issue stay tied to it
                if ret_cnt.get(no, 0) > 0:
                    if t == "ISSUE_DEL":
                        bad(f"seq {sq} ISSUE_DEL №{no}: сторно выдачи при действующих возвратах ({ret_cnt[no]}, {ret_sum[no]})", "R")
                    elif (f["EI"] != prev["ei"] or f.get("WHO") != prev.get("who") or f.get("DATE") != prev.get("date")
                          or float(f["QTY"]) < ret_sum[no] - EPS):
                        bad(f"seq {sq} ISSUE_FIX №{no}: при действующих возвратах ({ret_sum[no]}) изменены ЕИ/получатель/дата или количество "
                            f"стало меньше возвращённого: {f['EI']} {f.get('WHO')} {f.get('DATE')} {f['QTY']}", "R")
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
            want_s = (rec["name"], rec["art"] or "", rec["unit"], place.get(ei, rec["place"]), "Активен" if rec["state"] == "LIVE" else "Приход удалён (сторно)",
                      f"Перенос: заказ {rec['order']}" if rec.get("legacy") else f"Заказ {rec['order']}")
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

    # ------------------------------------------------------------ 7. the current place of the EIs (Phase 4: a return may move an EI)
    for ei in touched_place:
        row = book_ei.get(ei)
        if row is not None and row[5] != place.get(ei):
            bad(f"{ei}: место в «Наличие» {row[5]!r}, по движениям {place.get(ei)!r}", "R")

    info.update(dict(ops=len(ops), receipt_ops=n_rcpt_ops, receipts=len(receipts), live=sum(1 for x in receipts.values() if x["state"] == "LIVE"),
                     positions=len(positions), open_rows=open_rows, copies=copies, last_seq=last, journal_max=mx, return_ops=n_ret_ops,
                     invariants_failed=sorted(inv_fail, key=str)))
    if out is not None:
        out.update(bal=bal, issues=issues, returns=returns, ret_sum=ret_sum, ret_cnt=ret_cnt, place=place, min_next_ret=min_next_ret,
                   sysv=sysv, ops=ops, book_ei=book_ei, sp=sp, receipts=receipts, initial=initial, base=base, inv_fail=inv_fail, adjs=adjs)
    return P, info


def replay_return(t, f, sq, bal, issues, returns, ret_sum, ret_cnt, place, touched_place, bad):
    """one return operation of the journal, checked against the rules of the task of Phase 4 (written independently):
    a return belongs to a live issue, the same EI and recipient, a date not before the issue, 0 < quantity ≤ what can
    still be returned; a correction keeps the returns of the issue within the issued quantity and the balance ≥ 0; a
    storno never makes the balance negative"""
    n, k, ei = int(f["RET"]), int(f["ISSUE"]), f["EI"]
    iss = issues.get(k)
    if t == "RETURN":
        q = float(f["QTY"])
        if n in returns:
            bad(f"seq {sq} RETURN: № возврата {n} выдан повторно", "R")
            return
        if returns and n <= max(returns):
            bad(f"seq {sq} RETURN: № возврата {n} не больше предыдущего {max(returns)} (нумерация не монотонна)", "R")
        if iss is None or iss["deleted"]:
            bad(f"seq {sq} RETURN №{n}: выдача № {k} {'удалена' if iss else 'не существует'} — возврат запрещён", "R")
            return
        if ei != iss["ei"]:
            bad(f"seq {sq} RETURN №{n}: ЕИ {ei} ≠ ЕИ выдачи № {k} {iss['ei']}", "R")
        if f.get("WHO") != iss.get("who"):
            bad(f"seq {sq} RETURN №{n}: получатель {f.get('WHO')!r} ≠ получатель выдачи {iss.get('who')!r}", "R")
        if iss.get("date") and f.get("DATE", "") < iss["date"]:
            bad(f"seq {sq} RETURN №{n}: дата {f.get('DATE')} раньше даты выдачи {iss['date']}", "R")
        avail = iss["qty"] - ret_sum.get(k, 0.0)
        if not (q > EPS and q <= avail + EPS):
            bad(f"seq {sq} RETURN №{n}: количество {q}, а можно вернуть {avail}", "R")
        if abs(bal.get(ei, float("nan")) - float(f["BAL_BEFORE"])) > EPS:
            bad(f"seq {sq} RETURN №{n}: BAL_BEFORE {f['BAL_BEFORE']} ≠ расчётный {bal.get(ei)}", 3)
        bal[ei] = bal.get(ei, 0.0) + q
        ret_sum[k] = ret_sum.get(k, 0.0) + q
        ret_cnt[k] = ret_cnt.get(k, 0) + 1
        returns[n] = dict(issue=k, ei=ei, qty=q, state="LIVE", date=f.get("DATE"), place=f.get("PLACE"), who=f.get("WHO"), fixed=False,
                          L=float(f["BAL_BEFORE"]), M=float(f["BAL_AFTER"]), unit=f.get("UNIT"))
    else:
        r = returns.get(n)
        if r is None or r["state"] != "LIVE":
            bad(f"seq {sq} {t}: нет действующего возврата № {n}", "R")
            return
        if r["issue"] != k or r["ei"] != ei:
            bad(f"seq {sq} {t} №{n}: выдача/ЕИ {k} {ei} ≠ возврата {r['issue']} {r['ei']}", "R")
        if abs(bal.get(ei, float("nan")) - float(f["BAL_BEFORE"])) > EPS:
            bad(f"seq {sq} {t} №{n}: BAL_BEFORE {f['BAL_BEFORE']} ≠ расчётный {bal.get(ei)}", 3)
        if t == "RETURN_FIX":
            old, q = float(f["OLD_QTY"]), float(f["QTY"])
            if abs(old - r["qty"]) > EPS:
                bad(f"seq {sq} RETURN_FIX №{n}: OLD_QTY {old} ≠ расчётный {r['qty']}", "R")
            if iss is None or iss["deleted"]:
                bad(f"seq {sq} RETURN_FIX №{n}: выдача № {k} не действует", "R")
            elif ret_sum.get(k, 0.0) - r["qty"] + q > iss["qty"] + EPS:
                bad(f"seq {sq} RETURN_FIX №{n}: возвраты выдачи {ret_sum.get(k, 0.0) - r['qty'] + q} больше выданного {iss['qty']}", "R")
            if q <= EPS:
                bad(f"seq {sq} RETURN_FIX №{n}: количество {q} не больше 0", "R")
            if iss and iss.get("date") and f.get("DATE", "") < iss["date"]:
                bad(f"seq {sq} RETURN_FIX №{n}: дата {f.get('DATE')} раньше даты выдачи {iss['date']}", "R")
            bal[ei] += q - r["qty"]
            ret_sum[k] = ret_sum.get(k, 0.0) + q - r["qty"]
            r.update(qty=q, date=f.get("DATE"), place=f.get("PLACE"), who=f.get("WHO"), fixed=True, L=float(f["BAL_BEFORE"]) - old,
                     M=float(f["BAL_AFTER"]))
        else:
            if bal.get(ei, 0.0) < r["qty"] - EPS:
                bad(f"seq {sq} RETURN_DEL №{n}: остаток {bal.get(ei)} меньше возврата {r['qty']} — стал бы отрицательным", 6)
            bal[ei] -= r["qty"]
            ret_sum[k] = ret_sum.get(k, 0.0) - r["qty"]
            ret_cnt[k] = ret_cnt.get(k, 0) - 1
            r["state"] = "STORNO"
    if t != "RETURN_DEL":
        place[ei] = f.get("PLACE")
        touched_place.add(ei)
    if abs(bal[ei] - float(f["BAL_AFTER"])) > EPS:
        bad(f"seq {sq} {t} №{n}: BAL_AFTER {f['BAL_AFTER']} ≠ {bal[ei]}", 3)


# ================================================================ Core Phase 5: special receipts

# Cyrillic capitals that look like Latin ones; spaces of every kind and the soft hyphen are dropped; every dash is "-"
_LATIN = {"А": "A", "В": "B", "Е": "E", "Ё": "E", "К": "K", "М": "M", "Н": "H", "О": "O", "Р": "P", "С": "C", "Т": "T", "У": "Y", "Х": "X"}
_DROP = set("\t\n\r       ­")


def article_key(s):
    """the key of the article of a part (the rule of the task: case, spaces and typical symbols do not make another
    article): upper case, no spaces, dashes unified, Cyrillic lookalikes as Latin; «ABC100» and «ABC-100» stay different"""
    out = []
    for ch in (s or "").upper():
        if ch in _DROP:
            continue
        if "‐" <= ch <= "―" or ch == "−":
            ch = "-"
        out.append(_LATIN.get(ch, ch))
    return "".join(out)


def norm_text(s):
    """names and documents compared by the antidubl: letters and digits only, lower case, ё = е"""
    return "_".join(re.findall(r"[0-9a-zа-я]+", (s or "").strip().lower().replace("ё", "е")))


def dup_key(code, artkey, name, q, dserial, doc):
    """what makes two special lines «the same event»: source + article of a part (otherwise the name) + quantity + date + document"""
    return (code, ("a", artkey) if artkey else ("n", norm_text(name)), round(q, 6), int(dserial or 0), norm_text(doc))


def new_special_state():
    return dict(lines={}, eis={}, index={}, events={c: set() for c in SP_CODES}, max_spl=0, max_ev={c: 0 for c in SP_CODES},
                abandoned_spl=set(), abandoned_ev={c: set() for c in SP_CODES}, floor_spl=1, floor_ev={c: 1 for c in SP_CODES}, ops=0)


def copy_special_state(sp):
    out = new_special_state()
    out["from_base"] = sp.get("from_base", False)
    out["lines"] = {k: dict(v) for k, v in sp["lines"].items()}
    out["eis"] = {k: dict(v) for k, v in sp["eis"].items()}
    out["index"] = dict(sp["index"])
    out["events"] = {c: set(v) for c, v in sp["events"].items()}
    out["max_spl"] = sp["max_spl"]
    out["max_ev"] = dict(sp["max_ev"])
    return out


def snapshot_special(doc):
    """the special receipts of a book taken as trusted (benchmark: a generated history without a journal): _SPR, _ART and
    the cards of their EIs in «Наличие»"""
    sp = new_special_state()
    sp["from_base"] = True
    if not doc.Sheets.hasByName("_SPR"):
        return sp
    stock = read(doc, "Наличие", 10)
    for n, row in enumerate(read(doc, "_SPR", 14)):
        if n == 0 or not isinstance(row[0], float):
            continue
        m = EVENT_RE.fullmatch(str(row[1]))
        code, ei = row[2], row[3]
        if m:
            sp["events"][code].add(int(m.group(2)))
            sp["max_ev"][code] = max(sp["max_ev"][code], int(m.group(2)))
        sp["max_spl"] = max(sp["max_spl"], n)
        sp["lines"][n] = dict(event=row[1], code=code, ei=ei, mode=row[4], qty=row[5], state=row[6], date=row[8], artkey=row[11], ident=row[12],
                              fixes=int(row[13]) if isinstance(row[13], float) else 0, dk=None, base=True)
        k = ei_num(ei)
        if ei not in sp["eis"] and k is not None and k < len(stock):
            c = stock[k]
            sp["eis"][ei] = dict(code=code, stype=c[9], name=c[1], art=c[2], unit=c[3], cat=c[6], state=c[7], src=c[8],
                                 artkey=article_key(c[2]) if c[9] == SP_STYPE["DET"] else "", part=c[9] == SP_STYPE["DET"], base=True)
    for n, row in enumerate(read(doc, "_ART", 3)):
        if n and row[0]:
            sp["index"][row[0]] = row[1]
    return sp


def _live_dependents(ei, issues, returns):
    return ([k for k, v in issues.items() if not v["deleted"] and v["ei"] == ei],
            [k for k, v in returns.items() if v["state"] == "LIVE" and v["ei"] == ei])


def replay_special(t, f, sq, bal, sp, place, issues, returns, created, receipts, bad, adjs=None):
    """one special receipt operation of the journal, checked against the rules of the task of Phase 5 (see the module
    docstring); invariants S1…S10 of tests/special_oracle.py"""
    lines, eis, index = sp["lines"], sp["eis"], sp["index"]
    sp["ops"] += 1
    try:
        n = int(f["SPL"])
    except (KeyError, ValueError):
        bad(f"seq {sq} {t}: нет № строки (SPL)", "S4")
        return
    ei = f.get("EI", "")

    def fnum(k):
        try:
            return float(f[k])
        except (KeyError, ValueError):
            return float("nan")

    if t in ("SP_RECEIPT", "SP_REFILL"):
        code = f.get("SRC", "")
        if code not in SP_CODES:
            bad(f"seq {sq} {t} № {n}: неизвестный источник «{code}»", "S9")
            return
        if n in lines:
            bad(f"seq {sq} {t}: № строки {n} проведён повторно", "S4")
            return
        if n <= sp["max_spl"] or n in sp["abandoned_spl"]:
            bad(f"seq {sq} {t}: № строки {n} не больше выданного ранее ({sp['max_spl']}) или взят отложенной операцией", "S2")
        m = EVENT_RE.fullmatch(f.get("EVENT", ""))
        if not m or m.group(1) != code:
            bad(f"seq {sq} {t} № {n}: событие «{f.get('EVENT')}» не соответствует типу {code}", "S9")
            return
        num = int(m.group(2))
        if f.get("EVENT_NEW") == "1":
            if num in sp["events"][code] or num in sp["abandoned_ev"][code] or num <= sp["max_ev"][code]:
                bad(f"seq {sq} {t} № {n}: событие {f['EVENT']} выдано повторно или не по возрастанию (было до {sp['max_ev'][code]})", "S2")
            sp["events"][code].add(num)
            sp["max_ev"][code] = max(sp["max_ev"][code], num)
        elif f.get("EVENT_NEW") == "0":
            if num not in sp["events"][code] and num not in sp["abandoned_ev"][code]:
                bad(f"seq {sq} {t} № {n}: присоединение к событию {f['EVENT']}, которое WMS не выдавала", "S9")
        else:
            bad(f"seq {sq} {t} № {n}: нет признака нового события EVENT_NEW", "S9")
        q = fnum("QTY")
        if not q > EPS:
            bad(f"seq {sq} {t} № {n}: количество {f.get('QTY')!r} не больше 0", 6)
            return
        dser = iso_serial(f.get("DATE"))
        name, unit, plc = f.get("NAME", ""), f.get("UNIT", ""), f.get("PLACE", "")
        artkey = ""
        if code == "DET":
            artkey = article_key(f.get("ART", ""))
            if not artkey:
                bad(f"seq {sq} {t} № {n}: деталь без артикула", "S1")
            if f.get("ART_KEY", "") != artkey:
                bad(f"seq {sq} {t} № {n}: ключ артикула «{f.get('ART_KEY')}» ≠ нормализованный «{artkey}»", "S1")
        elif f.get("ART_KEY"):
            bad(f"seq {sq} {t} № {n}: ключ артикула у источника {code} (индекс — только у деталей)", "S1")
        if t == "SP_RECEIPT":
            if f.get("MODE") != "NEW":
                bad(f"seq {sq} SP_RECEIPT № {n}: режим {f.get('MODE')!r} ≠ NEW")
            if ei in bal or ei in receipts or ei in eis or ei_num(ei) is None:
                bad(f"seq {sq} SP_RECEIPT № {n}: {ei} уже существовал или не ЕИ — ЕИ выдан повторно", 2)
                return
            created.append(ei)
            if not (name and unit and plc and dser is not None):
                bad(f"seq {sq} SP_RECEIPT № {n}: пустое обязательное поле (наименование {name!r}, единица {unit!r}, место {plc!r}, дата {f.get('DATE')!r})")
            if code == "DET":
                if artkey in index:
                    bad(f"seq {sq} SP_RECEIPT № {n}: артикул «{artkey}» уже у {index[artkey]} — второй ЕИ для одного артикула", "S1")
                index[artkey] = ei
            if abs(fnum("BAL_BEFORE")) > EPS or abs(fnum("BAL_AFTER") - q) > EPS:
                bad(f"seq {sq} SP_RECEIPT № {n}: остаток до/после {f.get('BAL_BEFORE')}/{f.get('BAL_AFTER')} ≠ 0/{q}", 3)
            eis[ei] = dict(code=code, stype=SP_STYPE[code], name=name, art=f.get("ART", ""), unit=unit, cat=f.get("CAT", ""),
                           state=EI_REVIEW if code == "OTH" else EI_ACTIVE, src=f"{SP_STYPE[code]} {f['EVENT']}", artkey=artkey,
                           part=code == "DET", first=n)
            bal[ei] = q
            place[ei] = plc
            before = None
        else:
            if f.get("MODE") != "ADD" or code != "DET":
                bad(f"seq {sq} SP_REFILL № {n}: пополнение существующего ЕИ не у детали ({code}, {f.get('MODE')})", "S1")
                return
            card = eis.get(ei)
            if card is None:
                bad(f"seq {sq} SP_REFILL № {n}: {ei} не создан специальным приходом")
                return
            if index.get(artkey) != ei:
                bad(f"seq {sq} SP_REFILL № {n}: пополнен {ei}, а артикул «{artkey}» по индексу — {index.get(artkey)}", "S1")
            if unit.lower() != card["unit"].lower():
                bad(f"seq {sq} SP_REFILL № {n}: единица «{unit}» ≠ единица ЕИ «{card['unit']}» (пересчёта нет)", "S7")
            if name != card["name"]:
                bad(f"seq {sq} SP_REFILL № {n}: наименование «{name}» не из карточки «{card['name']}»", "S1")
            if card["state"] not in (EI_ACTIVE, EI_STORNO):
                bad(f"seq {sq} SP_REFILL № {n}: пополнение ЕИ в состоянии «{card['state']}»")
            if abs(fnum("BAL_BEFORE") - bal.get(ei, float("nan"))) > EPS:
                bad(f"seq {sq} SP_REFILL № {n}: BAL_BEFORE {f.get('BAL_BEFORE')} ≠ расчётный {bal.get(ei)}", 3)
            if plc != place.get(ei):
                if f.get("PLACE_BEFORE") != place.get(ei):
                    bad(f"seq {sq} SP_REFILL № {n}: место {ei} «{place.get(ei)}» → «{plc}» без отметки перемещения", "S3")
                place[ei] = plc
            elif "PLACE_BEFORE" in f:
                bad(f"seq {sq} SP_REFILL № {n}: отметка перемещения без смены места", "S3")
            bal[ei] = bal.get(ei, 0.0) + q
            if card["state"] == EI_STORNO:
                card["state"] = EI_ACTIVE
            before = fnum("BAL_BEFORE")
        if abs(fnum("BAL_AFTER") - bal[ei]) > EPS:
            bad(f"seq {sq} {t} № {n}: BAL_AFTER {f.get('BAL_AFTER')} ≠ {bal[ei]}", 3)
        dk = dup_key(code, artkey, name, q, dser, f.get("DOC", ""))
        dups = [k for k, L in lines.items() if L["state"] == "LIVE" and L.get("dk") == dk]
        if dups and not f.get("DUP") or f.get("DUP") and not dups and not sp.get("from_base"):
            bad(f"seq {sq} {t} № {n}: антидубль — действующие строки с тем же ключом {dups[:3]}, предупреждение {f.get('DUP')!r}", "S10")
        lines[n] = dict(event=f["EVENT"], code=code, ei=ei, mode=f.get("MODE"), qty=q, state="LIVE", date=dser, place=plc, cat=f.get("CAT", ""),
                        who=f.get("WHO", ""), doc=f.get("DOC", ""), mark=f.get("MARK", ""), name=name, art=f.get("ART", ""), unit=unit,
                        before=before, after=bal[ei], fixes=0, ident="", artkey=artkey, dk=dk, seq=sq)
        sp["max_spl"] = max(sp["max_spl"], n)
        return

    L = lines.get(n)
    if L is None or L["state"] != "LIVE":
        bad(f"seq {sq} {t}: нет действующей строки № {n}", "S4")
        return
    if L["ei"] != ei or L["event"] != f.get("EVENT"):
        bad(f"seq {sq} {t} № {n}: ЕИ/событие {ei} {f.get('EVENT')} ≠ строки {L['ei']} {L['event']} — идентичность не меняется", "S2")
        return
    card = eis.get(ei)
    if card is None:
        bad(f"seq {sq} {t} № {n}: нет карточки {ei}")
        return
    if t == "SP_FIX":
        if f.get("SRC") != L["code"] or f.get("MODE") != L["mode"]:
            bad(f"seq {sq} SP_FIX № {n}: источник/режим {f.get('SRC')} {f.get('MODE')} ≠ {L['code']} {L['mode']}", "S2")
        q1, q2 = fnum("OLD_QTY"), fnum("QTY")
        if abs(q1 - L["qty"]) > EPS:
            bad(f"seq {sq} SP_FIX № {n}: OLD_QTY {q1} ≠ расчётное {L['qty']}")
        if not q2 > EPS:
            bad(f"seq {sq} SP_FIX № {n}: количество {q2} не больше 0", 6)
        if abs(fnum("BAL_BEFORE") - bal[ei]) > EPS:
            bad(f"seq {sq} SP_FIX № {n}: BAL_BEFORE {f.get('BAL_BEFORE')} ≠ {bal[ei]}", 3)
        bal[ei] += q2 - q1
        if bal[ei] < -EPS:
            bad(f"seq {sq} SP_FIX № {n}: исправление сделало остаток {ei} отрицательным ({bal[ei]})", 6)
        if abs(fnum("BAL_AFTER") - bal[ei]) > EPS:
            bad(f"seq {sq} SP_FIX № {n}: BAL_AFTER {f.get('BAL_AFTER')} ≠ {bal[ei]}", 3)
        is_new = L["mode"] == "NEW"
        is_base = bool(L.get("base"))            # a line of the trusted snapshot: its row values are not known here
        if not is_new and any(k in f for k in ("NAME", "UNIT", "CAT", "ART")):
            bad(f"seq {sq} SP_FIX № {n}: у строки пополнения изменены данные карточки ЕИ", "S1")
        if "NAME" in f:
            L["name"] = card["name"] = f["NAME"]
        if "UNIT" in f:
            others = [k for k, x in lines.items() if x["ei"] == ei and x["state"] == "LIVE" and k != n]
            li, lr = _live_dependents(ei, issues, returns)
            if others or li or lr:
                bad(f"seq {sq} SP_FIX № {n}: единица {ei} изменена при других приходах {others[:3]} / выдачах {li[:3]} / возвратах {lr[:3]}", "S7")
            L["unit"] = card["unit"] = f["UNIT"]
        if "CAT" in f:
            L["cat"] = card["cat"] = f["CAT"]
        if "ART" in f:
            if not is_base and f.get("OLD_ART", "") != L["art"]:
                bad(f"seq {sq} SP_FIX № {n}: OLD_ART {f.get('OLD_ART')!r} ≠ {L['art']!r}")
            L["art"] = card["art"] = f["ART"]
            if card["part"]:
                nk = article_key(f["ART"])
                if not nk or f.get("ART_KEY") != nk:
                    bad(f"seq {sq} SP_FIX № {n}: ключ артикула {f.get('ART_KEY')!r} ≠ «{nk}»", "S1")
                if index.get(nk, ei) != ei:
                    bad(f"seq {sq} SP_FIX № {n}: артикул «{nk}» уже у {index[nk]} — исправление перевело деталь на чужой артикул", "S1")
                if nk != card["artkey"]:
                    index.pop(card["artkey"], None)
                    index[nk] = ei
                    card["artkey"] = L["artkey"] = nk
        if not is_base and f.get("OLD_PLACE") != L["place"]:
            bad(f"seq {sq} SP_FIX № {n}: OLD_PLACE {f.get('OLD_PLACE')!r} ≠ {L['place']!r}")
        if f.get("PLACE") != f.get("OLD_PLACE"):
            place[ei] = f.get("PLACE")
        dser = iso_serial(f.get("DATE"))
        L.update(qty=q2, date=dser, place=f.get("PLACE"), doc=f.get("DOC", ""), who=f.get("WHO", ""), mark=f.get("MARK", ""),
                 fixes=L["fixes"] + 1, before=None if is_new else fnum("BAL_BEFORE") - q1, after=bal[ei])
        if is_base:
            L["dk"] = None
        else:
            dk = dup_key(L["code"], L["artkey"] if card["part"] else "", L["name"], q2, dser, L["doc"])
            dups = [k for k, x in lines.items() if k != n and x["state"] == "LIVE" and x.get("dk") == dk]
            if bool(dups) != bool(f.get("DUP")):
                bad(f"seq {sq} SP_FIX № {n}: антидубль — действующие строки с тем же ключом {dups[:3]}, предупреждение {f.get('DUP')!r}", "S10")
            L["dk"] = dk
    elif t == "SP_DEL":
        q = fnum("QTY")
        if abs(q - L["qty"]) > EPS:
            bad(f"seq {sq} SP_DEL № {n}: QTY {q} ≠ приход строки {L['qty']}")
        others = [k for k, x in lines.items() if x["ei"] == ei and x["state"] == "LIVE" and k != n]
        # a migrated EI keeps its transferred stock: a line of it is never «the last receipt»
        last = not others and not (card or {}).get("migrated")
        if (f.get("LAST") == "1") != last:
            bad(f"seq {sq} SP_DEL № {n}: LAST {f.get('LAST')} ≠ расчётному ({'последняя' if last else 'есть другие'} строка ЕИ)", "S6")
        if last:
            li, lr = _live_dependents(ei, issues, returns)
            if li or lr:
                bad(f"seq {sq} SP_DEL № {n}: сторно последнего прихода {ei} при действующих выдачах {li[:5]} / возвратах {lr[:5]} (D-069)", "S6")
            la = [k for k, v in (adjs or {}).items() if v["state"] == "LIVE" and v["ei"] == ei]
            if la:
                bad(f"seq {sq} SP_DEL № {n}: сторно последнего прихода {ei} при действующих корректировках {la[:5]} (D-069)", "S6")
        if bal[ei] < q - EPS:
            bad(f"seq {sq} SP_DEL № {n}: остаток {bal[ei]} меньше прихода строки {q} — стал бы отрицательным", "S6")
        if abs(fnum("BAL_BEFORE") - bal[ei]) > EPS:
            bad(f"seq {sq} SP_DEL № {n}: BAL_BEFORE {f.get('BAL_BEFORE')} ≠ {bal[ei]}", 3)
        bal[ei] -= q
        if abs(fnum("BAL_AFTER") - bal[ei]) > EPS:
            bad(f"seq {sq} SP_DEL № {n}: BAL_AFTER {f.get('BAL_AFTER')} ≠ {bal[ei]}", 3)
        L["state"] = "STORNO"
        if last:
            card["state"] = EI_STORNO
            if abs(bal[ei]) > EPS:
                bad(f"seq {sq} SP_DEL № {n}: после сторно последнего прихода остаток {ei} = {bal[ei]}, а не 0", "S6")
    elif t == "SP_IDENTIFY":
        if L["code"] != "OTH" or L["mode"] != "NEW":
            bad(f"seq {sq} SP_IDENTIFY № {n}: разбор не строки иного прихода ({L['code']}, {L['mode']})", "S8")
        if card["state"] != EI_REVIEW:
            bad(f"seq {sq} SP_IDENTIFY № {n}: {ei} в состоянии «{card['state']}», а не «{EI_REVIEW}»", "S8")
        typ = f.get("TYPE", "")
        if typ not in IDENT_TYPES:
            bad(f"seq {sq} SP_IDENTIFY № {n}: тип «{typ}» недопустим", "S8")
        if not f.get("NAME"):
            bad(f"seq {sq} SP_IDENTIFY № {n}: пустое наименование", "S8")
        if (f.get("OLD_NAME"), f.get("OLD_ART", ""), f.get("OLD_CAT", "")) != (card["name"], card["art"], card["cat"]):
            bad(f"seq {sq} SP_IDENTIFY № {n}: прежняя карточка {f.get('OLD_NAME')!r} {f.get('OLD_ART')!r} {f.get('OLD_CAT')!r} ≠ "
                f"{card['name']!r} {card['art']!r} {card['cat']!r}", "S8")
        if typ == SP_STYPE["DET"]:
            k = article_key(f.get("ART", ""))
            if not k or f.get("ART_KEY") != k:
                bad(f"seq {sq} SP_IDENTIFY № {n}: ключ артикула {f.get('ART_KEY')!r} ≠ «{k}»", "S1")
            if k in index:
                bad(f"seq {sq} SP_IDENTIFY № {n}: артикул «{k}» уже у {index[k]} — второй ЕИ для одного артикула", "S1")
            index[k] = ei
            card["part"] = True
            card["artkey"] = L["artkey"] = k
        card.update(name=f.get("NAME", ""), art=f.get("ART", ""), cat=f.get("CAT", ""), state=EI_ACTIVE, src=card["src"] + " → " + typ, stype=typ)
        L.update(name=f.get("NAME", ""), art=f.get("ART", ""), cat=f.get("CAT", ""), ident=typ)


def replay_adjust(t, f, sq, bal, adjs, place, touched_place, bad):
    """one correction of the journal (Final Core), checked against the rules of the task written independently of the
    Basic code; the invariants of the corrections are reported as "A"."""
    def fnum(k):
        v = f.get(k, "")
        return float(v) if v != "" else None

    n, ei = int(f["ADJ"]), f["EI"]
    kind = t.replace("_FIX", "").replace("_DEL", "")
    if f.get("KIND") != kind:
        bad(f"seq {sq} {t} № {n}: KIND {f.get('KIND')!r} ≠ {kind}", "A")
    if ei not in bal:
        bad(f"seq {sq} {t} № {n}: {ei} не существует", "A")
        return
    if not f.get("REASON"):
        bad(f"seq {sq} {t} № {n}: нет причины / основания", "A")
    if t == kind:
        if n in adjs:
            bad(f"seq {sq} {t}: № {n} выдан повторно", "A")
            return
        if adjs and n <= max(adjs):
            bad(f"seq {sq} {t}: № {n} не больше уже выданного {max(adjs)}", "A")
        a = dict(kind=kind, ei=ei, state="LIVE", delta=0.0, frm=None, to=None, book=None, fact=None, date=f.get("DATE"),
                 reason=f.get("REASON"), batch=f.get("BATCH", ""), fixes=0, row=int(f.get("ROW", 0)))
        if kind == "MOVE":
            # the place of an EI the journal never set (the initial registry, a trusted snapshot) is known from here on
            if ei not in place:
                place[ei] = f.get("FROM", "")
            if f.get("FROM", "") != (place.get(ei) or ""):
                bad(f"seq {sq} MOVE № {n}: FROM {f.get('FROM')!r} ≠ месту {ei} по движениям {place.get(ei)!r} (тихое изменение места)", "A")
            if (f.get("TO") or "").strip().lower() == (f.get("FROM") or "").strip().lower():
                bad(f"seq {sq} MOVE № {n}: новое место совпадает со старым", "A")
            if fnum("BAL") is not None and abs(fnum("BAL") - bal[ei]) > EPS:
                bad(f"seq {sq} MOVE № {n}: BAL {f.get('BAL')} ≠ остатку {bal[ei]} (перемещение не меняет количество)", "A")
            a.update(frm=f.get("FROM", ""), to=f.get("TO", ""))
            place[ei] = f.get("TO", "")
            touched_place.add(ei)
        elif kind == "WRITE_OFF":
            q = fnum("QTY")
            if q is None or q <= 0:
                bad(f"seq {sq} WRITE_OFF № {n}: количество {f.get('QTY')!r}", "A")
                return
            if abs(fnum("BAL_BEFORE") - bal[ei]) > EPS:
                bad(f"seq {sq} WRITE_OFF № {n}: BAL_BEFORE {f.get('BAL_BEFORE')} ≠ {bal[ei]}", 3)
            if q > bal[ei] + EPS:
                bad(f"seq {sq} WRITE_OFF № {n}: списано {q} больше остатка {bal[ei]}", 6)
            bal[ei] -= q
            a["delta"] = -q
        else:
            book, fact, diff = fnum("BOOK"), fnum("FACT"), fnum("DIFF")
            if book is None or abs(book - bal[ei]) > EPS:
                bad(f"seq {sq} INV_ADJ № {n}: учётный остаток {f.get('BOOK')} ≠ остатку {ei} по движениям {bal[ei]}", "A")
            if fact is None or fact < -EPS or diff is None or abs(diff - (fact - (book or 0))) > EPS or abs(diff) < EPS:
                bad(f"seq {sq} INV_ADJ № {n}: факт {f.get('FACT')}, разница {f.get('DIFF')} не сходятся с учётным {f.get('BOOK')} или разница 0", "A")
                return
            bal[ei] += diff
            a.update(delta=diff, book=book, fact=fact)
        adjs[n] = a
    else:
        a = adjs.get(n)
        if a is None or a["state"] != "LIVE" or a["ei"] != ei or a["kind"] != kind:
            bad(f"seq {sq} {t} № {n}: нет действующей корректировки этого вида и ЕИ ({a})", "A")
            return
        if kind == "MOVE":
            if (place.get(ei) or "").strip().lower() != (a["to"] or "").strip().lower():
                bad(f"seq {sq} {t} № {n}: место {ei} {place.get(ei)!r} уже не то, куда его переместили ({a['to']!r})", "A")
            if t == "MOVE_DEL":
                place[ei] = a["frm"]
                a["state"] = "STORNO"
            else:
                if f.get("OLD_TO") != a["to"]:
                    bad(f"seq {sq} MOVE_FIX № {n}: OLD_TO {f.get('OLD_TO')!r} ≠ {a['to']!r}", "A")
                if (f.get("TO") or "").strip().lower() == (a["frm"] or "").strip().lower():
                    bad(f"seq {sq} MOVE_FIX № {n}: исправление вернуло ЕИ на прежнее место (это сторно)", "A")
                place[ei] = f.get("TO", "")
                a["to"] = f.get("TO", "")
                a["fixes"] += 1
            touched_place.add(ei)
        else:
            if fnum("BAL_BEFORE") is not None and abs(fnum("BAL_BEFORE") - bal[ei]) > EPS:
                bad(f"seq {sq} {t} № {n}: BAL_BEFORE {f.get('BAL_BEFORE')} ≠ {bal[ei]}", 3)
            if t.endswith("_DEL"):
                bal[ei] -= a["delta"]
                a["state"] = "STORNO"
            else:
                if kind == "WRITE_OFF":
                    q = fnum("QTY")
                    if q is None or q <= 0 or abs(fnum("OLD_QTY") + a["delta"]) > EPS:
                        bad(f"seq {sq} WRITE_OFF_FIX № {n}: количество {f.get('QTY')} / OLD_QTY {f.get('OLD_QTY')} ≠ {-a['delta']}", "A")
                        return
                    nd = -q
                else:
                    fact, diff = fnum("FACT"), fnum("DIFF")
                    if fnum("BOOK") is None or abs(fnum("BOOK") - a["book"]) > EPS or abs(fnum("OLD_DIFF") - a["delta"]) > EPS:
                        bad(f"seq {sq} INV_ADJ_FIX № {n}: BOOK/OLD_DIFF {f.get('BOOK')}/{f.get('OLD_DIFF')} ≠ {a['book']}/{a['delta']}", "A")
                    if fact is None or diff is None or abs(diff - (fact - a["book"])) > EPS or abs(diff) < EPS:
                        bad(f"seq {sq} INV_ADJ_FIX № {n}: факт {f.get('FACT')}, разница {f.get('DIFF')} не сходятся", "A")
                        return
                    nd = diff
                    a["fact"] = fact
                bal[ei] += nd - a["delta"]
                a["delta"] = nd
                a["fixes"] += 1
            if bal[ei] < -EPS:
                bad(f"seq {sq} {t} № {n}: остаток {ei} стал отрицательным ({bal[ei]})", 6)
            if fnum("BAL_AFTER") is not None and abs(fnum("BAL_AFTER") - bal[ei]) > EPS:
                bad(f"seq {sq} {t} № {n}: BAL_AFTER {f.get('BAL_AFTER')} ≠ {bal[ei]}", 3)
        a["date"] = f.get("DATE", a["date"])
        a["reason"] = f.get("REASON", a["reason"])
    if t == kind and kind != "MOVE" and fnum("BAL_AFTER") is not None and abs(fnum("BAL_AFTER") - bal[ei]) > EPS:
        bad(f"seq {sq} {t} № {n}: BAL_AFTER {f.get('BAL_AFTER')} ≠ {bal[ei]}", 3)


def replay_migrate(f, sq, bal, place, sp, created, bad):
    """one transfer of an existing EI (MIGRATE): the number is kept, never an existing EI, a part enters the index (one
    article — one EI), the card as the transfer wrote it; the invariants of the transfer are reported as "M"."""
    ei = f["EI"]
    if ei in bal or ei in sp["eis"]:
        bad(f"seq {sq} MIGRATE {ei}: ЕИ уже существовал — перенос перенумеровал или слил ЕИ", "M")
        return
    q = float(f["QTY"])
    stype = f.get("STYPE", "")
    if q < 0:
        bad(f"seq {sq} MIGRATE {ei}: отрицательное количество {q}", "M")
    if stype not in MIGRATE_STYPES:
        bad(f"seq {sq} MIGRATE {ei}: тип источника «{stype}» (пустой тип не переносится молча, D-087)", "M")
    if not f.get("ORIGIN"):
        bad(f"seq {sq} MIGRATE {ei}: нет происхождения (ORIGIN)", "M")
    part = stype == "Детали"
    key = article_key(f.get("ART", "")) if part else ""
    if part:
        if not key:
            bad(f"seq {sq} MIGRATE {ei}: деталь без артикула", "M")
        elif key in sp["index"]:
            bad(f"seq {sq} MIGRATE {ei}: артикул «{key}» уже у {sp['index'][key]} — второй ЕИ для одного артикула", "M")
        else:
            sp["index"][key] = ei
    bal[ei] = q
    place[ei] = f.get("PLACE", "")
    created.append(ei)
    sp["eis"][ei] = dict(code="MIG", stype=stype, name=f.get("NAME", ""), art=f.get("ART", ""), unit=f.get("UNIT", ""), cat=f.get("CAT", ""),
                         state=EI_REVIEW if stype == "Иной приход" else EI_ACTIVE,
                         src="Перенос" + (f" ({f['MARK']})" if f.get("MARK") else ""), artkey=key, part=part, first=None, migrated=q)
