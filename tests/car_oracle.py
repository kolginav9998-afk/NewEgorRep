"""Independent check of «Приход авто» (M6 PRIME §1) — written separately from the Basic code on purpose.

check(doc, jdir, expect_tail=0) replays the vehicle operations of the journal (CAR_ARRIVE, CAR_DEPART, CAR_FIX,
CAR_CANCEL; abandoned seqs skipped; only seq ≤ LAST_SEQ) with the rules of the task and compares the result with every
view of the book:
  _CAR           — one record per visit № (row = №): key, state, arrival, departure, day, № of the day, the key and the
                   arrival of an open visit only while it is open, fixes, duration, supplier key; nothing after the last
                   visit; the dictionaries hold every vehicle and supplier key once;
  «Приход авто»  — visit n on row 6 + n (1-based): №, date, vehicle and supplier as typed (the last fix), arrival,
                   departure, the fixed duration of a closed visit, status, weekday, month, № of the day, key; no row
                   without its visit;
  _SYS           — NEXT_CAR = the last visit № + 1 (not below an ABANDON record);
  rules          — at most one open visit per vehicle at any moment of the history (checked at every step of the
                   replay); a departure never before its arrival; the visit № grow by one; the № of the day restarts
                   at 1 on a new day; the stock is never touched (no CAR_* operation writes outside _SYS NEXT_CAR,
                   _CAR and «Приход авто»).
The key of a vehicle and of a supplier is computed here independently with the rule of the task: upper case, the Latin
letters that look like the Cyrillic letters of a plate replaced by them, Ё → Е, letters and digits only; a Russian plate
(letter, three digits, two letters; the last one in the text) is the key of its vehicle.
"""
import datetime
import re

import journal_oracle

CAR_TYPES = ("CAR_ARRIVE", "CAR_DEPART", "CAR_FIX", "CAR_CANCEL")
CAR_FIRST = 6                       # 0-based row of visit 1 on «Приход авто»
NULL = datetime.datetime(1899, 12, 30)
TOL = 0.6 / 86400                   # times are whole seconds
S_OPEN, S_GONE, S_CANCEL = "На территории", "Уехал", "Отменён"
STATE_TEXT = {"OPEN": S_OPEN, "CLOSED": S_GONE, "CANCELLED": S_CANCEL}
WEEKDAYS = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]   # datetime.weekday()
LAT, CYR = "ABEKMHOPCTYX", "АВЕКМНОРСТУХ"
PLATE_LETTERS = set(CYR)
ALLOWED_SHEETS = {"_SYS", "_CAR", "Приход авто"}


def search_key(s):
    out = []
    for ch in s.upper():
        if ch in LAT:
            ch = CYR[LAT.index(ch)]
        if ch == "Ё":
            ch = "Е"
        if ch.isdigit() and ch.isascii() or ("A" <= ch <= "Z") or ("А" <= ch <= "Я"):
            out.append(ch)
    return "".join(out)


def plate_key(s):
    t = search_key(s)
    for i in range(len(t) - 6, -1, -1):
        w = t[i:i + 6]
        if w[0] in PLATE_LETTERS and w[1:4].isdigit() and w[4] in PLATE_LETTERS and w[5] in PLATE_LETTERS:
            return w
    return t


def clean(s):
    return " ".join(s.replace(" ", " ").split())


def supplier_key(s):
    t = clean(s).upper().replace("Ё", "Е")
    t = "".join(ch for ch in t if ch not in "«»\"'„“”")
    return clean(t)


def serial(ts):
    """'YYYY-MM-DD HH:MM:SS' → the date-time number of the book"""
    d = datetime.datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
    return (d - NULL).total_seconds() / 86400


def from_serial(x):
    return NULL + datetime.timedelta(days=x)


def used_rows(sheet):
    cur = sheet.createCursor()
    cur.gotoEndOfUsedArea(False)
    return cur.getRangeAddress().EndRow


def replay(entries, last_seq, P):
    """the visits by №: dict(plate, key, sup, skey, state, arr, dep, day, nday, fixes, cancelled_from)"""
    visits, open_keys = {}, {}
    ab = journal_oracle.abandoned_seqs([e for e in entries if e["seq"] <= last_seq])
    abandon_next, nxt = 0, 1
    for e in entries:
        if e["seq"] > last_seq:
            continue
        if e["type"] == "ABANDON":
            v = e["fields"].get("NEXT_CAR")
            if v and v.isdigit():
                abandon_next = max(abandon_next, int(v))
                nxt = max(nxt, int(v))
        if e["type"] not in CAR_TYPES or e["seq"] in ab:
            continue
        f = e["fields"]
        for w in e["writes"]:
            if w["sheet"] not in ALLOWED_SHEETS:
                P.append(f"seq {e['seq']} {e['type']}: запись на лист «{w['sheet']}» — «Приход авто» не должен трогать склад")
        n = int(float(f["CAR"]))
        if e["type"] == "CAR_ARRIVE":
            # the № grow by one; an abandoned operation (ABANDON) may have taken some — they are never reused
            if n != nxt:
                P.append(f"seq {e['seq']}: визит № {n}, ожидался {nxt} (номера растут на 1, отложенные не переиспользуются)")
            nxt = n + 1
            if n in visits:
                P.append(f"seq {e['seq']}: визит № {n} выдан повторно")
                continue
            key = plate_key(f["PLATE"])
            if key != f["KEY"]:
                P.append(f"seq {e['seq']}: ключ машины «{f['KEY']}», по правилу задания «{key}» («{f['PLATE']}»)")
            if key in open_keys:
                P.append(f"seq {e['seq']}: машина {key} открыта второй раз (уже визит № {open_keys[key]})")
            arr = serial(f["TIME"])
            day = int(arr)
            prev = visits[max(visits)] if visits else None
            nday = prev["nday"] + 1 if prev and prev["day"] == day else 1
            visits[n] = dict(plate=clean(f["PLATE"]), key=key, sup=clean(f["SUPPLIER"]), skey=supplier_key(f["SUPPLIER"]), state="OPEN",
                             arr=arr, dep=None, day=day, nday=nday, fixes=0, cancelled_from=None)
            open_keys[key] = n
        elif e["type"] == "CAR_DEPART":
            v = visits.get(n)
            if v is None or v["state"] != "OPEN":
                P.append(f"seq {e['seq']}: выезд визита № {n}, который {'не открыт' if v else 'не существует'}")
                continue
            dep = serial(f["TIME"])
            if dep < v["arr"] - TOL:
                P.append(f"seq {e['seq']}: выезд визита № {n} раньше приезда")
            if abs(round((dep - v["arr"]) * 1440) - int(float(f.get("MIN", "-1")))) > 1:
                P.append(f"seq {e['seq']}: MIN {f.get('MIN')} ≠ длительности {(dep - v['arr']) * 1440:.1f} мин")
            v.update(state="CLOSED", dep=dep)
            open_keys.pop(v["key"], None)
        elif e["type"] == "CAR_FIX":
            v = visits.get(n)
            if v is None or v["state"] == "CANCELLED":
                P.append(f"seq {e['seq']}: исправление визита № {n}, который {'отменён' if v else 'не существует'}")
                continue
            if "PLATE" in f:
                key = plate_key(f["PLATE"])
                if key != f.get("KEY"):
                    P.append(f"seq {e['seq']}: ключ исправленной машины «{f.get('KEY')}», по правилу «{key}»")
                if v["state"] == "OPEN" and key != v["key"]:
                    if key in open_keys:
                        P.append(f"seq {e['seq']}: исправление открыло машину {key} второй раз")
                    open_keys.pop(v["key"], None)
                    open_keys[key] = n
                v.update(plate=clean(f["PLATE"]), key=key)
            if "SUPPLIER" in f:
                v.update(sup=clean(f["SUPPLIER"]), skey=supplier_key(f["SUPPLIER"]))
            arr = serial(f["ARR"])
            if int(arr) != v["day"]:
                P.append(f"seq {e['seq']}: исправление перенесло приезд визита № {n} на другой день")
            v["arr"] = arr
            if v["state"] == "CLOSED":
                if "DEP" not in f:
                    P.append(f"seq {e['seq']}: исправление закрытого визита № {n} без времени выезда")
                else:
                    v["dep"] = serial(f["DEP"])
                if v["dep"] < v["arr"] - TOL:
                    P.append(f"seq {e['seq']}: после исправления выезд визита № {n} раньше приезда")
            elif "DEP" in f:
                P.append(f"seq {e['seq']}: исправление записало выезд открытого визита № {n}")
            v["fixes"] += 1
        elif e["type"] == "CAR_CANCEL":
            v = visits.get(n)
            if v is None or v["state"] == "CANCELLED":
                P.append(f"seq {e['seq']}: отмена визита № {n}, который {'уже отменён' if v else 'не существует'}")
                continue
            if v["state"] == "OPEN":
                open_keys.pop(v["key"], None)
            v.update(cancelled_from=v["state"], state="CANCELLED")
    return visits, abandon_next


def close(a, b):
    return isinstance(a, float) and b is not None and abs(a - b) <= TOL


def check(doc, jdir, expect_tail=0, out=None):
    P = []
    entries, _ = journal_oracle.read_journal(jdir)
    sysv = journal_oracle.sys_values(doc)
    inst = sysv["INSTANCE_ID"]
    entries = [e for e in entries if e["inst"] == inst]
    last_seq = int(sysv["LAST_SEQ"])
    mx = max((e["seq"] for e in entries), default=0)
    if mx - last_seq != expect_tail:
        P.append(f"LAST_SEQ {last_seq}, последняя запись журнала {mx}, ожидался хвост {expect_tail}")
    visits, abandon_next = replay(entries, last_seq, P)
    last = max(visits) if visits else 0
    nxt = int(doc.Sheets.getByName("_SYS").getCellByPosition(1, 26).getValue())
    if nxt < last + 1 or nxt < abandon_next or (nxt != last + 1 and nxt != abandon_next):
        P.append(f"NEXT_CAR {nxt}: последний визит № {last}, отложенный хвост требует ≥ {abandon_next}")
    car = doc.Sheets.getByName("_CAR")
    sh = doc.Sheets.getByName("Приход авто")
    nrows = max(last + 2, 3)
    cr = car.getCellRangeByPosition(0, 0, 20, nrows).getDataArray()
    cv = sh.getCellRangeByPosition(0, 0, 13, CAR_FIRST + last + 1).getDataArray()
    for n in range(1, last + 1):
        v = visits.get(n)
        c = cr[n]
        s = cv[CAR_FIRST + n - 1]
        if v is None:
            # a № of an abandoned operation: no record anywhere
            if n >= abandon_next or any(x != "" for x in c[:13]) or any(x != "" for x in s):
                P.append(f"визит № {n}: нет операции CAR_ARRIVE в журнале" + ("" if n >= abandon_next else " (номер отложенной операции, но строка не пуста)"))
            continue
        st = v["state"]
        want_c = dict(no=float(n), key=v["key"], state=st, okey=v["key"] if st == "OPEN" else "", skey=v["skey"])
        got_c = dict(no=c[0], key=c[1], state=c[2], okey=c[8], skey=c[12])
        if got_c != want_c:
            P.append(f"_CAR визит № {n}: {got_c} ≠ расчёт {want_c}")
        if not close(c[4], v["arr"]) or c[6] != float(v["day"]) or c[7] != float(v["nday"]):
            P.append(f"_CAR визит № {n}: приезд/день/№ за день {c[4]!r}/{c[6]!r}/{c[7]!r} ≠ {v['arr']}/{v['day']}/{v['nday']}")
        if st == "OPEN":
            if c[5] != "" or not close(c[9], v["arr"]) or c[11] != "":
                P.append(f"_CAR визит № {n} (открыт): выезд {c[5]!r}, приезд открытого {c[9]!r}, длительность {c[11]!r}")
        elif st == "CLOSED" or v["cancelled_from"] == "CLOSED":
            if not close(c[5], v["dep"]) or c[9] != "" or not isinstance(c[11], float) or abs(c[11] - (v["dep"] - v["arr"])) > TOL:
                P.append(f"_CAR визит № {n} ({st}): выезд {c[5]!r} ≠ {v['dep']}, длительность {c[11]!r}")
        elif c[9] != "":
            P.append(f"_CAR визит № {n} (отменён на территории): приезд открытого визита не очищен")
        fx = c[10] if c[10] != "" else 0.0
        if fx != float(v["fixes"]):
            P.append(f"_CAR визит № {n}: исправлений {c[10]!r} ≠ {v['fixes']}")
        # the sheet
        d = from_serial(v["arr"])
        want_s = (float(n), float(v["day"]), v["plate"], v["sup"], STATE_TEXT[st], WEEKDAYS[d.weekday()], d.strftime("%Y-%m"),
                  float(v["nday"]), v["key"])
        got_s = (s[0], s[1], s[2], s[3], s[7], s[8], s[9], s[10], s[13])
        if got_s != want_s:
            P.append(f"«Приход авто» строка {CAR_FIRST + n}: {got_s} ≠ расчёт {want_s}")
        if not close(s[4], v["arr"]):
            P.append(f"«Приход авто» строка {CAR_FIRST + n}: приезд {s[4]!r} ≠ {v['arr']}")
        if st == "CLOSED" or v["cancelled_from"] == "CLOSED":
            if not close(s[5], v["dep"]) or not isinstance(s[6], float) or abs(s[6] - (v["dep"] - v["arr"])) > TOL:
                P.append(f"«Приход авто» строка {CAR_FIRST + n}: выезд {s[5]!r} / длительность {s[6]!r} ≠ {v['dep']} / {v['dep'] - v['arr']}")
        elif st == "OPEN":
            f = sh.getCellByPosition(6, CAR_FIRST + n - 1).getFormula()
            if s[5] != "" or f != f"=NOW()-E{CAR_FIRST + n}":
                P.append(f"«Приход авто» строка {CAR_FIRST + n} (на территории): выезд {s[5]!r}, длительность «{f}» — ожидалась формула текущего времени")
        elif s[6] != "":
            P.append(f"«Приход авто» строка {CAR_FIRST + n} (отменён на территории): длительность {s[6]!r} не очищена")
        ctl = s[11]
        if not isinstance(ctl, str) or not ctl.startswith("Приезд записан"):
            P.append(f"«Приход авто» строка {CAR_FIRST + n}: контроль «{ctl}»")
    # nothing after the last visit
    if any(x != "" for x in cr[last + 1][:13]):
        P.append(f"_CAR: запись после последнего визита № {last}: {cr[last + 1][:4]}")
    tail = sh.getCellRangeByPosition(0, CAR_FIRST + last, 13, max(used_rows(sh), CAR_FIRST + last)).getDataArray()
    for i, row in enumerate(tail):
        if any(x != "" for x in row):
            P.append(f"«Приход авто» строка {CAR_FIRST + last + i + 1}: данные без визита {row[:4]}")
            break
    # the dictionaries: the vehicles and the suppliers of the visits, each key once
    np_, ns = int(cr[0][18]), int(cr[0][20])
    dp = car.getCellRangeByPosition(13, 1, 16, max(np_, ns, 1) + 1).getDataArray()
    pkeys = [r[1] for r in dp[:np_]]
    skeys = [r[3] for r in dp[:ns]]
    if len(set(pkeys)) != len(pkeys) or "" in pkeys:
        P.append(f"словарь машин: повторы или пустые ключи ({np_} записей)")
    if len(set(skeys)) != len(skeys) or "" in skeys:
        P.append(f"словарь поставщиков: повторы или пустые ключи ({ns} записей)")
    if dp[np_][1] != "" or dp[ns][3] != "":
        P.append("словарь: запись после его размера")
    for n, v in visits.items():
        if v["key"] not in pkeys:
            P.append(f"визит № {n}: машины {v['key']} нет в словаре")
        if v["skey"] not in skeys:
            P.append(f"визит № {n}: поставщика {v['skey']} нет в словаре")
    info = f"визитов {last}: на территории {sum(1 for v in visits.values() if v['state'] == 'OPEN')}, " \
           f"уехало {sum(1 for v in visits.values() if v['state'] == 'CLOSED')}, отменено {sum(1 for v in visits.values() if v['state'] == 'CANCELLED')}"
    if out is not None:
        out["visits"] = visits
    return P, info


if __name__ == "__main__":
    for s, want in (("Газель А123ВС 77", "А123ВС"), ("a 123 bc", "А123ВС"), ("КАМАЗ 65115 Х456ОР", "Х456ОР"), ("Петрович", "ПЕТРОВИЧ"),
                    ("Scania е001кх 750", "Е001КХ"), ("ГАЗ 3302 А123ВС", "А123ВС")):
        assert plate_key(s) == want, (s, plate_key(s))
    assert re.fullmatch(r"ПЕТРОВИЧ ООО", supplier_key(' «Петрович»  ооо '))
    print("car_oracle self-test OK")
