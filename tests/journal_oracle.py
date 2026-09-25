"""Independent check of a WMS book against its journal (written separately from the Basic code on purpose).

read_journal(dir)  parses every WMS_journal_*.csv (UTF-8; LF or CRLF), validates J1;...;END;length
check(doc, dir)    book ↔ journal: LAST_SEQ, continuity, instance, marker, and for the test sheet _TST: every journaled
                   TEST operation (not abandoned) is posted exactly once with its inputs, and every posted row is journaled
"""
import glob
import os
import re

SYS_KEYS = ["SCHEMA", "INSTANCE_ID", "MODE", "CORE_VERSION", "LAST_SEQ", "NEXT_EI", "NEXT_NO", "NEXT_RET",
            "JOURNAL_POS", "REGISTERED_URL", "TX_STATE", "TX_SEQ", "TX_TYPE", "TX_TIME", "TX_BEFORE_IMAGE",
            "SAVE_STAMP", "SAVE_SEQ", "MAX_QTY", "KEY_SHEETS"]


def unesc(s):
    return re.sub(r"%([0-9A-Fa-f]{2})", lambda m: chr(int(m.group(1), 16)), s)


def parse_line(ln):
    if not ln.startswith("J1;"):
        return None
    p = ln.rfind(";END;")
    if p < 0:
        return None
    body, slen = ln[:p], ln[p + 5:]
    if not slen.isdigit() or int(slen) != len(body):
        return None
    a = body.split(";")
    if len(a) < 5 or not a[1].isdigit():
        return None
    fields, writes = {}, []
    for f in a[5:]:
        name, _, val = f.partition("=")
        if name == "W":
            p = val.split("|")
            if len(p) != 7:
                return None
            writes.append(dict(kind=p[0], sheet=unesc(p[1]), row=int(p[2]), col=p[3], before=p[4], after=p[5], restorable=p[6] == "1"))
        else:
            fields[name] = unesc(val)
    return dict(seq=int(a[1]), ts=a[2], inst=a[3], type=a[4], fields=fields, writes=writes, raw=ln)


def read_journal(jdir):
    entries, damaged, crlf, files = [], 0, 0, sorted(glob.glob(os.path.join(jdir, "WMS_journal_*.csv")))
    for f in files:
        for raw in open(f, "rb").read().split(b"\n"):
            if raw.endswith(b"\r"):
                crlf += 1
                raw = raw[:-1]
            if not raw:
                continue
            try:
                e = parse_line(raw.decode("utf-8"))
            except UnicodeDecodeError:
                e = None
            if e is None:
                damaged += 1
            else:
                entries.append(e)
    return entries, dict(files=len(files), damaged=damaged, crlf=crlf)


def sys_values(doc):
    sh = doc.Sheets.getByName("_SYS")
    out = {}
    for i, k in enumerate(SYS_KEYS):
        c = sh.getCellByPosition(1, i)
        out[k] = c.getValue() if c.getType().value == "VALUE" else c.getString()
    return out


def abandoned_seqs(entries):
    s = set()
    for e in entries:
        if e["type"] == "ABANDON":
            s.update(range(int(e["fields"]["FROM"]), int(e["fields"]["TO"]) + 1))
    return s


def check(doc, jdir, expect_tail=0):
    P = []
    entries, info = read_journal(jdir)
    sysv = sys_values(doc)
    inst = sysv["INSTANCE_ID"]
    ours = [e for e in entries if e["inst"] == inst]
    if len(ours) != len(entries):
        P.append(f"строк другого экземпляра: {len(entries) - len(ours)}")
    seqs = [e["seq"] for e in ours]
    if seqs:
        if seqs != list(range(seqs[0], seqs[0] + len(seqs))):
            P.append(f"журнал: seq не непрерывны: {seqs[:8]}…")
        if seqs[0] != 1 and ours[0]["type"] != "START":
            P.append(f"журнал начинается с seq {seqs[0]} без START")
    last = int(sysv["LAST_SEQ"])
    mx = max(seqs) if seqs else 0
    if mx - last != expect_tail:
        P.append(f"LAST_SEQ книги {last}, последняя запись журнала {mx}, ожидался хвост {expect_tail}")
    if sysv["TX_STATE"] not in ("COMMITTED", "NONE"):
        P.append(f"маркер {sysv['TX_STATE']}")
    if sysv["TX_BEFORE_IMAGE"]:
        P.append("снимок до изменения не очищен")
    # test sheet
    applied = [e for e in ours if e["seq"] <= last]
    ab = abandoned_seqs(applied)
    ops = [e for e in applied if e["type"] == "TEST" and e["seq"] not in ab]
    posted = {}
    if doc.Sheets.hasByName("_TST"):
        sh = doc.Sheets.getByName("_TST")
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        n = cur.getRangeAddress().EndRow
        rows = sh.getCellRangeByPosition(0, 0, 3, max(n, 1)).getDataArray()
        for r, row in enumerate(rows):
            if r == 0 or not isinstance(row[0], float):
                continue
            if str(row[3]).startswith("КОПИЯ"):
                continue
            if row[3] != "Проведено":
                P.append(f"_TST строка {r + 1}: ключ {row[0]} без статуса «Проведено» ({row[3]!r})")
                continue
            posted.setdefault(int(row[0]), []).append((r, row))
        for no, lst in posted.items():
            if len(lst) > 1:
                P.append(f"_TST: ключ {no} проведён {len(lst)} раз (строки {[x[0] + 1 for x in lst]})")
        for e in ops:
            no = int(float(e["fields"].get("NO", "-1")))
            if no not in posted:
                P.append(f"операция seq {e['seq']} (№{no}) есть в журнале, но не проведена в книге")
                continue
            r, row = posted[no][0]
            if row[1] != e["fields"].get("TEXT"):
                P.append(f"№{no}: текст в книге {row[1]!r} ≠ журнал {e['fields'].get('TEXT')!r}")
            if abs(float(row[2] or 0) - float(e["fields"].get("QTY", "nan"))) > 1e-9:
                P.append(f"№{no}: количество в книге {row[2]!r} ≠ журнал {e['fields'].get('QTY')}")
        journaled = {int(float(e["fields"].get("NO", "-1"))) for e in ops}
        started_later = bool(ours) and ours[0]["type"] == "START"     # history before START lives in another journal
        for no in posted:
            if no not in journaled and not started_later:
                P.append(f"_TST: ключ {no} проведён в книге, но операции нет в журнале")
        if ops and not started_later:
            mxno = max(journaled)
            if int(sysv["NEXT_NO"]) != mxno + 1:
                P.append(f"NEXT_NO {sysv['NEXT_NO']} ≠ максимальный № {mxno} + 1")
    info.update(dict(entries=len(ours), last_seq=last, journal_max=mx, test_ops=len(ops), posted=len(posted), abandoned=len(ab)))
    return P, info
