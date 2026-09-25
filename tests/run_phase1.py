"""Phase 1 reliability skeleton — automated scenarios (CLAUDE_TASK_CORE_PHASE1 "Tests required before completion").

    WMS_TEST_OUT=/tmp/wms_out python3 tests/run_phase1.py [case ...]

Every case builds its own copy of a fresh test book in WMS_TEST_OUT/cases/<case>/ and drives real LibreOffice
processes (headless). Results: WMS_TEST_OUT/results_phase1.json and a Markdown table WMS_TEST_OUT/TEST_REPORT_phase1.md.
"""
import glob
import os
import shutil
import subprocess
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import OUT, Session, Results, new_wms, journal_files, template, unique  # noqa: E402
from wmslo import Office, prop  # noqa: E402
import journal_oracle  # noqa: E402

R = Results()
CASES = []


def case(fn):
    CASES.append(fn)
    return fn


def st(s):
    return s.state().get("STATE"), s.state().get("BLOCK")


def autoinput_of_profile(profile_name):
    o = Office(profile_name, os.path.join(OUT, "profiles"), fresh=False)
    try:
        cp = o.smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", o.ctx)
        na = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (prop("nodepath", "/org.openoffice.Office.Calc/Input"),))
        return bool(na.getByName("AutoInput"))
    finally:
        o.terminate()


def autoinput_now(s):
    cp = s.o.smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", s.o.ctx)
    na = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (prop("nodepath", "/org.openoffice.Office.Calc/Input"),))
    return bool(na.getByName("AutoInput"))


def foreign_line(seq):
    body = f"J1;{seq};2026-01-01T00:00:00;FFFFFFFFFFFFFFFF;TEST;NO=1"
    return f"{body};END;{len(body)}\n"


def rewrite_journal(jdir, fn):
    for f in journal_files(jdir):
        data = open(f, "rb").read()
        open(f, "wb").write(fn(data))


def lines_of(jdir):
    out = []
    for f in journal_files(jdir):
        out += [x for x in open(f, "rb").read().split(b"\n") if x.strip()]
    return out


# ================================================================ T01–T03 normal work

@case
def t01_fresh_start():
    c = "T01"
    s = Session(new_wms("t01"))
    try:
        rep = s.report()
        state, block = st(s)
        lock = s.lock_text() or ""
        bdir = os.path.join(s.dir, "WMS_Backups")
        daily = [f for f in os.listdir(bdir) if f.endswith("_daily.ods")] if os.path.isdir(bdir) else []
        ai = autoinput_now(s)
        R.add(c, "чистый запуск новой WMS: регистрация, блокировка, ежедневная копия, автоввод выключен",
              state == "CLEAN" and "зарегистрирован" in rep and s.state()["SESSION"] in lock and len(daily) == 1 and ai is False,
              f"{rep[:260]}; lock: {lock.strip()[:90]}; копий: {daily}; AutoInput={ai}")
        sc = s.B("ActionSelfCheck")
        R.add(c, "самопроверка чистой WMS", "ошибок 0" in sc.splitlines()[0], sc.replace("\n", " / ")[:600])
    finally:
        s.close()


@case
def t02_normal_operation():
    c = "T02"
    s = Session(new_wms("t02"))
    try:
        s.input_row(1, "Болт М8х30", "1,5")
        t0 = time.time()
        res = s.T("TestPost", 1)
        dt = (time.time() - t0) * 1000
        row, locks, stt = s.tst(1), s.locks(1), s.state()
        ents, info = s.journal()
        e = ents[-1] if ents else {}
        jf = journal_files(s.jdir)
        jpos = s.sysv("JOURNAL_POS")
        ok = (res == "OK:1" and tuple(row) == (1.0, "Болт М8х30", 1.5, "Проведено") and locks == "1111"
              and stt["LAST_SEQ"] == "1" and stt["NEXT_NO"] == "2" and stt["TX"] == "COMMITTED:1"
              and stt["UNDO_LOCKED"] == "False" and stt["UNDO_CAN"] == "False"
              and len(ents) == 1 and e["type"] == "TEST" and e["fields"].get("QTY") == "1.5" and len(e["writes"]) == 6
              and jpos == f"{os.path.basename(jf[-1])}|{os.path.getsize(jf[-1])}|1")
        R.add(c, "операция: план → STARTED → запись → журнал → COMMITTED", ok,
              f"{res}; строка {tuple(row)}; защита {locks}; {stt}; журнал: {e.get('raw', '')[:200]}; JOURNAL_POS {jpos}", timing=round(dt, 1))
        P, inf = s.check()
        R.add(c, "оракул", not P, f"{inf}; {P[:3]}")
        R.add(c, "повторное проведение той же строки", s.T("TestPost", 1).startswith("SKIP"), "")
    finally:
        s.close()


@case
def t03_many_operations():
    c = "T03"
    s = Session(new_wms("t03"))
    try:
        n = 200
        s.input_rows_api(1, [(f"Позиция {i}", f"{i % 7 + 1},25") for i in range(1, n + 1)])
        t0 = time.time()
        res = s.T("TestPostRange", 1, n)
        dt = (time.time() - t0) * 1000
        P, inf = s.check()
        R.add(c, f"{n} операций подряд", res.startswith(f"OK={n};ERR=0") and not P, f"{res}; {dt / n:.1f} мс/операция; {inf}; {P[:3]}",
              timing=round(dt / n, 2))
    finally:
        s.close()


# ================================================================ T04–T06 errors before / after the journal

@case
def t04_error_before_journal_seam():
    c = "T04"
    s = Session(new_wms("t04"))
    try:
        s.input_row(1, "база", "1")
        s.T("TestPost", 1)
        for point, r in ((1, 2), (2, 3)):
            s.input_row(r, f"Гайка {point}", "2,5")
            before, lb = s.tst(r), s.locks(r)
            nj, last, nxt = len(lines_of(s.jdir)), s.sysv("LAST_SEQ"), s.sysv("NEXT_NO")
            s.B("TestSetFault", point, 1)
            res = s.T("TestPost", r)
            after, la = s.tst(r), s.locks(r)
            ok = (res.startswith("ERR-RB") and tuple(after) == tuple(before) and la == lb and s.tst_cell(2, r)[0] == "TEXT"
                  and len(lines_of(s.jdir)) == nj and s.sysv("LAST_SEQ") == last and s.sysv("NEXT_NO") == nxt
                  and s.sysv("TX_STATE") == "COMMITTED" and s.sysv("TX_BEFORE_IMAGE") == "")
            R.add(c, f"ошибка Basic в точке {point} (до журнала) → полный откат", ok,
                  f"{res}; строка до {tuple(before)} после {tuple(after)}; защита {lb}->{la}; журнал {nj}->{len(lines_of(s.jdir))}")
            retry = s.T("TestPost", r)
            R.add(c, f"повтор после отката в точке {point} проводит ровно один раз", retry.startswith("OK:"), retry)
        P, inf = s.check()
        R.add(c, "оракул", not P, f"{inf}; {P[:3]}")
    finally:
        s.close()


@case
def t05_error_before_journal_natural():
    c = "T05"
    s = Session(new_wms("t05"))
    try:
        s.input_row(1, "первая", "1")
        s.T("TestPost", 1)
        s.input_row(2, "сломанный план", "1")
        before = s.tst(2)
        res = s.T("TestPostBrokenPlan", 2)
        R.add(c, "вторая запись плана не выполнима → первая откатывается", res.startswith("ERR-RB") and tuple(s.tst(2)) == tuple(before)
              and len(lines_of(s.jdir)) == 1, f"{res}; строка {tuple(s.tst(2))}")
        jf = journal_files(s.jdir)[-1]
        chat = subprocess.run(["chattr", "+i", jf], capture_output=True, text=True)
        if chat.returncode != 0:
            R.add(c, "журнал недоступен для записи → откат", "SKIP", f"chattr недоступен: {chat.stderr.strip()[:100]}")
        else:
            try:
                s.input_row(3, "журнал закрыт", "2")
                before = s.tst(3)
                res = s.T("TestPost", 3)
                ok = res.startswith("ERR-RB") and tuple(s.tst(3)) == tuple(before) and s.sysv("LAST_SEQ") == 1 and s.sysv("NEXT_NO") == 2
                R.add(c, "журнал недоступен для записи (файл неизменяемый) → операции нет, откат", ok, f"{res}; строка {tuple(s.tst(3))}")
            finally:
                subprocess.run(["chattr", "-i", jf])
            res2 = s.T("TestPost", 3)
            R.add(c, "после восстановления доступа проводится ровно один раз", res2 == "OK:2", res2)
        P, inf = s.check()
        R.add(c, "оракул", not P, f"{inf}; {P[:3]}")
    finally:
        s.close()


@case
def t06_error_after_journal():
    c = "T06"
    s = Session(new_wms("t06"))
    try:
        s.input_row(1, "после журнала", "4")
        s.B("TestSetFault", 3, 1)
        res = s.T("TestPost", 1)
        ok = res.startswith("OK:1") and "после сбоя" in res and s.sysv("LAST_SEQ") == 1 and tuple(s.tst(1))[3] == "Проведено" and len(lines_of(s.jdir)) == 1
        R.add(c, "ошибка после записи журнала → операция доводится, не откатывается", ok, f"{res}; {tuple(s.tst(1))}")
        P, inf = s.check()
        R.add(c, "оракул", not P, f"{inf}; {P[:3]}")
    finally:
        s.close()


# ================================================================ T07–T11 crash, kill, unsaved tail

@case
def t07_crash_before_journal():
    c = "T07"
    p = new_wms("t07")
    s = Session(p)
    s.input_row(1, "a", "1")
    s.input_row(2, "b", "1")
    s.T("TestPost", 1)
    s.T("TestPost", 2)
    s.input_row(3, "Шайба", "3")
    pre = tuple(s.tst(3))
    s.B("TestSetFault", 2, 2)
    res = s.T("TestPost", 3)
    s.kill()
    s2 = Session(p)
    try:
        rep1 = s2.report()
        unl = s2.B("ActionForceUnlock")
        state, block = st(s2)
        row, locks = tuple(s2.tst(3)), s2.locks(3)
        ok = (res == "CRASH-SIM" and "LOCKED" in rep1 and "отменена по снимку" in unl and state == "CLEAN" and row == pre and locks == "0000"
              and s2.sysv("LAST_SEQ") == 2 and s2.sysv("NEXT_NO") == 3 and len(lines_of(s2.jdir)) == 2)
        R.add(c, "книга сохранена посреди операции (до журнала) + kill → откат по маркеру, операции нет нигде", ok,
              f"{res}; запуск: {rep1[:140]}; после «Снять блокировку»: {unl[:260]}; строка {row}, защита {locks}")
        retry = s2.T("TestPost", 3)
        P, inf = s2.check()
        R.add(c, "повтор проводит один раз; оракул", retry == "OK:3" and not P, f"{retry}; {inf}; {P[:3]}")
    finally:
        s2.close()


@case
def t08_crash_after_journal():
    c = "T08"
    p = new_wms("t08")
    s = Session(p)
    s.input_row(1, "a", "1")
    s.T("TestPost", 1)
    s.input_row(2, "после журнала", "5")
    s.B("TestSetFault", 3, 2)
    res = s.T("TestPost", 2)
    s.kill()
    s2 = Session(p)
    try:
        rep = s2.B("ActionForceUnlock")
        state, block = st(s2)
        rec = s2.B("ActionRecover")
        state2, _ = st(s2)
        P, inf = s2.check()
        ok = res == "CRASH-SIM" and "отменена по снимку" in rep and block == "TAIL" and "восстановлено операций: 1" in rec and state2 == "CLEAN" and not P
        R.add(c, "книга сохранена после записи журнала + kill → откат по маркеру, затем «Восстановить» доводит из журнала", ok,
              f"{res}; запуск: {rep[:300]}; восстановление: {rec[:160]}; {inf}; {P[:3]}")
    finally:
        s2.close()


@case
def t09_kill_during_batch():
    c = "T09"
    p = new_wms("t09")
    s = Session(p)
    n = 150
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, n + 1)])
    res = {}

    def run():
        try:
            res["r"] = s.T("TestPostRange", 1, n)
        except Exception as e:
            res["e"] = str(e)[:80]
    th = threading.Thread(target=run)
    th.start()
    t0 = time.time()
    while time.time() - t0 < 60:
        if len(lines_of(s.jdir)) >= 40:
            break
        time.sleep(0.02)
    s.kill()
    th.join(10)
    k = len([x for x in lines_of(s.jdir) if journal_oracle.parse_line(x.decode("utf-8", "replace").rstrip("\r"))])
    s2 = Session(p)
    try:
        rep1 = s2.report()
        unl = s2.B("ActionForceUnlock")
        rec = s2.B("ActionRecover")
        state, _ = st(s2)
        P, inf = s2.check()
        posted = sum(1 for r in range(1, n + 1) if tuple(s2.tst(r))[3] == "Проведено")
        R.add(c, f"kill -9 посреди пакета из {n}: каждая записанная в журнал операция доводится ровно один раз",
              "LOCKED" in rep1 and f"восстановлено операций: {k}" in rec and state == "CLEAN" and posted == k and not P,
              f"в журнале к моменту kill: {k}; вызов: {res}; «Снять блокировку»: {unl[:200]}; «Восстановить»: {rec[:120]}; проведено строк: {posted}; {inf}; {P[:3]}")
        s2.input_row(k + 1, "после восстановления", "2")
        nxt = s2.T("TestPost", k + 1)
        R.add(c, "после восстановления нумерация продолжается", nxt == f"OK:{k + 1}", nxt)
    finally:
        s2.close()


@case
def t10_close_without_save():
    c = "T10"
    p = new_wms("t10")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "2") for i in range(1, 21)])
    r1 = s.T("TestPostRange", 1, 20)
    s.close(save=False)
    s2 = Session(p)
    try:
        rep = s2.report()
        _, block = st(s2)
        blocked_post = s2.T("TestPost", 25)
        rec = s2.B("ActionRecover")
        rec2 = s2.B("ActionRecover")
        P, inf = s2.check()
        R.add(c, "20 проведений → «Не сохранять» → запуск: блок; «Восстановить» возвращает все 20 ровно один раз",
              r1.startswith("OK=20") and block == "TAIL" and blocked_post.startswith("BLOCKED") and "восстановлено операций: 20" in rec
              and "не требуется" in rec2 and not P,
              f"запуск: {rep[:200]}; проведение до восстановления: {blocked_post[:80]}; {rec[:100]}; повторно: {rec2[:60]}; {inf}; {P[:3]}")
        jpos = s2.sysv("JOURNAL_POS")
        pinned = jpos.split("|")[-1] == str(int(s2.sysv("LAST_SEQ")))
    finally:
        s2.close(save=True)
    s3 = Session(p)
    try:
        rep3 = s3.report()
        R.add(c, "после «Восстановить» позиция журнала переставлена на последнюю операцию: следующий запуск читает только новое",
              pinned and st(s3)[0] == "CLEAN" and "не подтвердилась" not in rep3, f"позиция {jpos}; запуск: {rep3[:200]}")
    finally:
        s3.close()


@case
def t11_abandon_tail():
    c = "T11"
    p = new_wms("t11")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 11)])
    s.T("TestPostRange", 1, 10)
    s.close(save=False)
    s2 = Session(p)
    ab = s2.B("ActionAbandonTail")
    state, _ = st(s2)
    last = s2.sysv("LAST_SEQ")
    s2.input_row(1, "новая", "3")
    nxt = s2.T("TestPost", 1)
    P, inf = s2.check()
    R.add(c, "«Отложить хвост журнала»: ABANDON дописан, книга продолжает после него, хвост не применён",
          "отложено операций: 10" in ab and state == "CLEAN" and last == 11 and nxt == "OK:12" and not P,
          f"{ab[:200]}; LAST_SEQ {last}; новая операция {nxt}; {inf}; {P[:3]}")
    s2.close(save=True)
    s3 = Session(p)
    try:
        state3, _ = st(s3)
        ents, _ = s3.journal()
        R.add(c, "после сохранения и перезапуска — чисто; отложенные операции остались в журнале",
              state3 == "CLEAN" and any(e["type"] == "ABANDON" for e in ents) and len(ents) == 12, f"{state3}; записей {len(ents)}")
    finally:
        s3.close()


# ================================================================ T12–T17 journal integrity

@case
def t12_journal_crlf():
    c = "T12"
    p = new_wms("t12")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 6)])
    s.T("TestPostRange", 1, 5)
    s.close(save=True)
    rewrite_journal(os.path.join(os.path.dirname(p), "WMS_Journal"), lambda b: b.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n"))
    s2 = Session(p)
    try:
        rep = s2.report()
        state, _ = st(s2)
        s2.input_row(6, "после CRLF", "1")
        res = s2.T("TestPost", 6)
        P, inf = s2.check()
        R.add(c, "журнал переведён в CRLF → позиция не подтверждается, журнал перечитывается, без ложной тревоги",
              state == "CLEAN" and "не подтвердилась" in rep and "повреждённых" not in rep and res == "OK:6" and not P,
              f"{rep[:260]}; {res}; {inf}; {P[:3]}")
    finally:
        s2.close()


@case
def t13_crlf_crash_shrink():
    c = "T13"
    p = new_wms("t13")
    jdir = os.path.join(os.path.dirname(p), "WMS_Journal")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 6)])
    s.T("TestPostRange", 1, 3)
    s.close(save=True)
    rewrite_journal(jdir, lambda b: b.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n"))
    s = Session(p)
    s.T("TestPost", 4)
    s.close(save=True)
    s = Session(p)
    s.T("TestPost", 5)
    s.kill()
    rewrite_journal(jdir, lambda b: b.replace(b"\r\n", b"\n"))
    s2 = Session(p)
    try:
        unl = s2.B("ActionForceUnlock")
        _, block = st(s2)
        rec = s2.B("ActionRecover")
        P, inf = s2.check()
        R.add(c, "CRLF → операция после сохранения → авария → журнал стал LF (короче): операция не теряется",
              block == "TAIL" and "восстановлено операций: 1" in rec and not P, f"{unl[:260]}; {rec[:120]}; {inf}; {P[:3]}")
    finally:
        s2.close()


@case
def t14_invalid_saved_offset():
    c = "T14"
    p = new_wms("t14")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 8)])
    s.T("TestPostRange", 1, 5)
    jpos = s.sysv("JOURNAL_POS")
    f, off, seq = jpos.split("|")
    out = []
    for bad in (f"{f}|{int(off) - 7}|{seq}", f"{f}|{off}|{int(seq) - 1}", f"{f}|abc|{seq}", "WMS_journal_1999-01.csv|10|5"):
        s.set_sys("JOURNAL_POS", bad)
        s.close(save=True)
        s = Session(p)
        rep = s.report()
        out.append((bad, st(s)[0], "не подтвердилась" in rep))
    try:
        s.input_row(6, "после", "1")
        res = s.T("TestPost", 6)
        P, inf = s.check()
        R.add(c, "неверная сохранённая позиция журнала → журнал перечитывается, состояние верное",
              all(x[1] == "CLEAN" and x[2] for x in out) and res == "OK:6" and not P, f"{out}; {res}; {inf}; {P[:3]}")
    finally:
        s.close()


@case
def t15_journal_gap():
    c = "T15"
    p = new_wms("t15")
    jdir = os.path.join(os.path.dirname(p), "WMS_Journal")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 10)])
    s.T("TestPostRange", 1, 5)
    s.close(save=True)
    rewrite_journal(jdir, lambda b: b"\n".join(x for x in b.split(b"\n") if not x.startswith(b"J1;3;")))
    s = Session(p)
    rep = s.report()
    _, block = st(s)
    post = s.T("TestPost", 6)
    sc = s.B("ActionSelfCheck")
    R.add(c, "разрыв seq в истории журнала → блок, проведение запрещено", block == "JOURNAL_GAP" and post.startswith("BLOCKED"),
          f"{rep[:200]}; {post[:80]}; самопроверка: {sc.splitlines()[0]}")
    s.close()
    # gap inside the unsaved tail
    p = new_wms("t15b")
    jdir = os.path.join(os.path.dirname(p), "WMS_Journal")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 10)])
    s.T("TestPostRange", 1, 3)
    s.close(save=True)
    s = Session(p)
    s.T("TestPostRange", 4, 6)
    s.kill()
    rewrite_journal(jdir, lambda b: b"\n".join(x for x in b.split(b"\n") if not x.startswith(b"J1;5;")))
    s = Session(p)
    try:
        unl = s.B("ActionForceUnlock")
        _, block = st(s)
        rec = s.B("ActionRecover")
        post = s.T("TestPost", 7)
        R.add(c, "разрыв seq в несохранённом хвосте → блок, «Восстановить» не применяет, проведение запрещено",
              block == "JOURNAL_GAP" and "не требуется" in rec and post.startswith("BLOCKED"), f"{unl[:220]}; {rec[:80]}; {post[:60]}")
    finally:
        s.close()


@case
def t16_foreign_writer():
    c = "T16"
    p = new_wms("t16")
    s = Session(p)
    try:
        s.input_row(1, "a", "1")
        s.T("TestPost", 1)
        jf = journal_files(s.jdir)[-1]
        with open(jf, "a", encoding="utf-8") as f:
            f.write(foreign_line(2))
        s.input_row(2, "b", "1")
        before = tuple(s.tst(2))
        r1 = s.T("TestPost", 2)
        r2 = s.T("TestPost", 2)
        ok = "другим процессом" in r1 and r1.startswith("ERR-RB") and r2.startswith("BLOCKED") and tuple(s.tst(2)) == before
        R.add(c, "журнал дописан другим процессом → операция откачена, дальше проведение остановлено", ok, f"{r1[:160]}; {r2[:120]}")
    finally:
        s.close()
    s = Session(p)
    try:
        _, block = st(s)
        R.add(c, "при следующем запуске чужая строка журнала → блок", block == "JOURNAL_FOREIGN", s.report()[:200])
    finally:
        s.close()


@case
def t17_journal_behind():
    c = "T17"
    p = new_wms("t17")
    jdir = os.path.join(os.path.dirname(p), "WMS_Journal")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 6)])
    s.T("TestPostRange", 1, 5)
    s.close(save=True)
    rewrite_journal(jdir, lambda b: b"\n".join(x for x in b.split(b"\n") if not (x.startswith(b"J1;4;") or x.startswith(b"J1;5;"))))
    s = Session(p)
    try:
        _, block = st(s)
        post = s.T("TestPost", 1)
        R.add(c, "книга впереди журнала (журнал неполный) → блок", block == "JOURNAL_BEHIND" and post.startswith("BLOCKED"), s.report()[:260])
    finally:
        s.close()


# ================================================================ T18–T22 lock, second instance, copies and backups

@case
def t18_stale_lock():
    c = "T18"
    p = new_wms("t18")
    s = Session(p)
    s.kill()
    s = Session(p)
    try:
        _, block = st(s)
        post = s.T("TestPost", 1)
        unl = s.B("ActionForceUnlock")
        state, _ = st(s)
        s.input_row(1, "после", "1")
        res = s.T("TestPost", 1)
        R.add(c, "устаревшая блокировка после аварии → блок до «Снять блокировку WMS»",
              block == "LOCKED" and post.startswith("BLOCKED") and state == "CLEAN" and res == "OK:1", f"{post[:120]}; {unl[:160]}; {res}")
    finally:
        s.close()


@case
def t19_second_instance():
    c = "T19"
    p = new_wms("t19")
    a = Session(p, name="A")
    a.input_row(1, "A1", "1")
    a.T("TestPost", 1)
    for f in glob.glob(os.path.join(os.path.dirname(p), ".~lock.*#")):
        os.remove(f)     # as if the second copy were opened from another PC or with «Open» despite LibreOffice's own lock
    b = Session(p, name="B")
    try:
        _, blockB = st(b)
        postB = b.T("TestPost", 2)
        a.input_row(2, "A2", "1")
        postA = a.T("TestPost", 2)
        R.add(c, "второй экземпляр той же WMS: проведение заблокировано, первый работает", blockB == "LOCKED" and postB.startswith("BLOCKED") and postA == "OK:2",
              f"B: {b.report()[:160]}; B проводит: {postB[:60]}; A проводит: {postA}")
        unl = b.B("ActionForceUnlock")
        a.input_row(3, "A3", "1")
        postA2 = a.T("TestPost", 3)
        _, blockB2 = st(b)
        R.add(c, "блокировку сняли во втором экземпляре по ошибке → первый останавливается, второй требует восстановления",
              postA2.startswith("BLOCKED") and blockB2 == "TAIL", f"A: {postA2[:140]}; B после снятия: {unl[:160]}")
    finally:
        a.close()
        b.close()


@case
def t20_old_backup_direct():
    c = "T20"
    p = new_wms("t20")
    s = Session(p)
    s.input_rows_api(1, [("a", "1"), ("b", "1")])
    s.T("TestPostRange", 1, 2)
    s.close(save=True)
    bdir = os.path.join(os.path.dirname(p), "WMS_Backups")
    bk = sorted(glob.glob(os.path.join(bdir, "*_daily.ods")))[0]
    ods = lambda: sorted(f for f in os.listdir(bdir) if f.endswith(".ods"))
    n_backups = len(ods())
    s = Session(bk)
    try:
        rep = s.report()
        _, block = st(s)
        post = s.T("TestPost", 3)
        lock_exists = os.path.exists(os.path.join(os.path.dirname(p), "WMS_Journal", "wms.lock"))
        R.add(c, "старая резервная копия открыта напрямую → не рабочий файл, проведение запрещено, блокировка и копии не тронуты",
              block == "NOT_REGISTERED" and "резервная копия" in rep and post.startswith("BLOCKED") and not lock_exists
              and len(ods()) == n_backups and not os.path.isdir(os.path.join(bdir, "WMS_Journal")),
              f"{rep[:240]}; {post[:60]}")
    finally:
        s.close()


@case
def t21_backup_restored():
    c = "T21"
    p = new_wms("t21")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 11)])
    s.T("TestPostRange", 1, 5)
    s.doc.store()
    bk = s.B("ActionBackupNow")
    s.T("TestPostRange", 6, 10)
    s.close(save=True)
    bdir = os.path.join(os.path.dirname(p), "WMS_Backups")
    manual = sorted(glob.glob(os.path.join(bdir, "*_manual.ods")))[-1]
    shutil.copy(manual, p)
    s = Session(p)
    try:
        _, block = st(s)
        rec = s.B("ActionRecover")
        state, _ = st(s)
        P, inf = s.check()
        R.add(c, "резервная копия (после 5) скопирована поверх рабочей (было 10) → «Восстановить» доводит до 10",
              block == "TAIL" and "восстановлено операций: 5" in rec and state == "CLEAN" and s.sysv("LAST_SEQ") == 10 and s.sysv("NEXT_NO") == 11 and not P,
              f"{bk[:120]}; {rec[:100]}; {inf}; {P[:3]}")
    finally:
        s.close()


@case
def t22_move_and_register():
    c = "T22"
    p = new_wms("t22")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 6)])
    s.T("TestPostRange", 1, 3)
    s.close(save=True)
    nd = os.path.join(OUT, "cases", "t22", "Склад WMS 2026")
    os.makedirs(nd)
    np_ = os.path.join(nd, "Рабочая WMS.ods")
    shutil.copy(p, np_)
    shutil.copytree(os.path.join(os.path.dirname(p), "WMS_Journal"), os.path.join(nd, "WMS_Journal"))
    os.remove(os.path.join(nd, "WMS_Journal", "wms.lock")) if os.path.exists(os.path.join(nd, "WMS_Journal", "wms.lock")) else None
    s = Session(np_)
    try:
        _, block = st(s)
        reg = s.B("ActionRegisterHere")
        state, _ = st(s)
        s.input_row(4, "после переноса", "1")
        res = s.T("TestPost", 4)
        R.add(c, "перенос WMS вместе с журналом в папку с кириллицей и пробелами → «Сделать рабочим файлом»",
              block == "NOT_REGISTERED" and state == "CLEAN" and res == "OK:4", f"{reg[:200]}; {res}")
    finally:
        s.close()
    nd2 = os.path.join(OUT, "cases", "t22", "без журнала")
    os.makedirs(nd2)
    np2 = os.path.join(nd2, "WMS.ods")
    shutil.copy(p, np2)
    s = Session(np2)
    try:
        r1 = s.B("ActionRegisterHere")
        _, block = st(s)
        r2 = s.B("ActionRegisterHere", True)
        state, _ = st(s)
        s.input_row(4, "новый журнал", "1")
        res = s.T("TestPost", 4)
        P, inf = s.check()
        ents, _ = s.journal()
        R.add(c, "перенос без журнала → блок «журнал отстаёт» → «Сделать рабочим файлом» с новым журналом (START)",
              block == "JOURNAL_BEHIND" and state == "CLEAN" and res == "OK:5" and ents[0]["type"] == "START" and not P,
              f"{r1[:120]}; {r2[:160]}; {res}; {inf}; {P[:3]}")
    finally:
        s.close()


# ================================================================ T23–T27 service data, macros off, Undo, reopen, format

@case
def t23_corrupt_marker():
    c = "T23"
    out = []
    # the last case: a well-formed before-image, but TX_SEQ = LAST_SEQ (LAST_SEQ moved while the marker stayed STARTED)
    for i, (state_v, bi) in enumerate((("STARTED", "BI1~V|_TST|zz|1|E"), ("STARTED", "мусор"), ("WHAT", ""),
                                       ("STARTED", "BI1~V|_TST|1|1|E"))):
        p = new_wms(f"t23_{i}")
        s = Session(p)
        s.input_row(1, "a", "1")
        s.T("TestPost", 1)
        s.set_sys("TX_STATE", state_v)
        s.set_sys("TX_BEFORE_IMAGE", bi)
        s.close(save=True)
        s = Session(p)
        _, block = st(s)
        post = s.T("TestPost", 2)
        out.append((state_v, bi[:12], block, post[:40]))
        s.close()
    R.add(c, "повреждённый маркер операции → блок, проведение запрещено", all(x[2] == "MARKER_CORRUPT" and x[3].startswith("BLOCKED") for x in out), str(out))


@case
def t24_saved_without_wms():
    c = "T24"
    p = new_wms("t24")
    s = Session(p)
    s.input_rows_api(1, [("оригинал", "1"), ("второй", "2")])
    s.T("TestPostRange", 1, 2)
    s.close(save=True)
    s = Session(p, macros=0)
    sh = s.doc.Sheets.getByName("_TST")
    sh.getCellRangeByPosition(0, 5, 3, 6).setDataArray(((1.0, "копия", 1.0, "Проведено"), (99.0, "подделка", 1.0, "Проведено")))
    s.doc.store()
    s.close()
    s = Session(p)
    try:
        rep = s.report()
        state, _ = st(s)
        d5, d6, d1 = tuple(s.tst(5))[3], tuple(s.tst(6))[3], tuple(s.tst(1))[3]
        post5, post6 = s.T("TestPost", 5), s.T("TestPost", 6)
        P, inf = s.check()
        R.add(c, "книга сохранена с отключёнными макросами: при запуске обнаружено, поддельные/повторные ключи помечены КОПИЯ",
              "без WMS" in rep and "не выдавались 1" in rep and "повторов 1" in rep and state == "CLEAN" and d5.startswith("КОПИЯ") and d6.startswith("КОПИЯ")
              and d1 == "Проведено" and post5.startswith("SKIP") and post6.startswith("SKIP") and not P,
              f"{rep[:420]}; строка 6: {d5[:60]}; строка 7: {d6[:60]}; {post5}; {inf}; {P[:3]}")
    finally:
        s.close()


@case
def t25_ctrl_z():
    c = "T25"
    s = Session(new_wms("t25"))
    try:
        s.input_row(1, "Ctrl+Z", "2")
        s.T("TestPost", 1)
        row = tuple(s.tst(1))
        for _ in range(5):
            s.ui(".uno:Undo")
        stt = s.state()
        ok1 = tuple(s.tst(1)) == row and stt["LAST_SEQ"] == "1" and stt["UNDO_LOCKED"] == "False"
        s.type_in(1, 2, "ввод после проведения")
        s.ui(".uno:Undo")
        ok2 = tuple(s.tst(2))[1] == ""
        R.add(c, "Ctrl+Z ×5 после проведения ничего не откатывает; Undo не остаётся заблокированным", ok1 and ok2,
              f"строка {tuple(s.tst(1))}; {stt}; отмена нового ввода работает: {ok2}")
    finally:
        s.close()


@case
def t26_save_reopen():
    c = "T26"
    p = new_wms("t26")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 12)])
    s.T("TestPostRange", 1, 10)
    before = {k: s.sysv(k) for k in ("LAST_SEQ", "NEXT_NO", "JOURNAL_POS", "REGISTERED_URL", "INSTANCE_ID", "TX_STATE")}
    sess1 = s.state()["SESSION"]
    s.close(save=True)
    lock_after_close = os.path.exists(os.path.join(os.path.dirname(p), "WMS_Journal", "wms.lock"))
    s = Session(p)
    try:
        rep = s.report()
        after = {k: s.sysv(k) for k in before}
        state, _ = st(s)
        res = s.T("TestPost", 11)
        P, inf = s.check()
        R.add(c, "сохранить → закрыть → открыть: счётчики, маркер, позиция журнала, регистрация сохранены; блокировка снята и создана заново",
              before == after and state == "CLEAN" and not lock_after_close and s.state()["SESSION"] != sess1 and "не подтвердилась" not in rep
              and "без WMS" not in rep and res == "OK:11" and not P, f"{before}; {rep[:200]}; {res}; {inf}; {P[:3]}", timing=round(s.open_s, 2))
    finally:
        s.close()


@case
def t27_ods_guard():
    c = "T27"
    p = new_wms("t27")
    s = Session(p)
    s.input_row(1, "a", "1")
    s.T("TestPost", 1)
    d = os.path.dirname(p)
    import uno
    from wmslo import props
    out = {}
    made = []
    for ext, flt in (("fods", "OpenDocument Spreadsheet Flat XML"), ("xlsx", "Calc MS Excel 2007 XML"), ("xls", "MS Excel 97")):
        try:
            s.doc.storeToURL(uno.systemPathToFileUrl(os.path.join(d, "WMS." + ext)), props(FilterName=flt))
            made.append(ext)
        except Exception as e:
            out[ext] = f"сохранить в {ext} не удалось ({str(e)[:50]})"
    s.close(save=True)
    for ext in made:
        try:
            x = Session(os.path.join(d, "WMS." + ext))
        except Exception as e:
            out[ext] = f"не открылся: {e}"
            continue
        try:
            try:
                rep = x.report()
                post = x.T("TestPost", 2)
                out[ext] = f"{st(x)[1]} / {post[:40]}"
            except Exception as e:
                out[ext] = f"макросы WMS отсутствуют ({type(e).__name__}) → проведение невозможно"
        finally:
            x.close()
    ok = all(("NOT_ODS" in v and "BLOCKED" in v) or "отсутствуют" in v or "не удалось" in v for v in out.values()) and "fods" in made
    R.add(c, "книга не в формате ODS (fods/xlsx/xls) → проведение невозможно", ok, str(out))


# ================================================================ T28–T35 parser, AutoInput, backups, diagnostics, special data, production book, timing

PARSER = [("1,5", 0, 1.5), ("1.5", 0, 1.5), ("0,125", 0, 0.125), ("10,75", 0, 10.75), ("3", 0, 3), (" 2 ", 0, 2),
          ("1,50", 0, 1.5), ("2,1250", 0, 2.125), ("1000,5", 0, 1000.5), ("1,500", 3, None), ("12.345", 3, None),
          ("1,2345", 4, None), ("01.05.2026", 6, None), ("1.5.26", 6, None), ("1,5,0", 6, None), ("1e3", 2, None),
          ("1E3", 2, None), ("-1", 2, None), ("−1", 2, None), ("1 000", 2, None), ("1 000", 2, None), ("1/2", 2, None),
          ("½", 2, None), ("¾", 2, None), ("0", 5, None), ("0,0", 5, None), ("007", 2, None), (",5", 2, None), ("1,", 2, None),
          ("", 1, None), ("abc", 2, None), ("1,5кг", 2, None), ("1234567890", 2, None), ("12.03", 0, 12.03)]


@case
def t28_quantity_parser():
    c = "T28"
    s = Session(new_wms("t28"))
    bad = []
    try:
        for txt, stt, val in PARSER:
            r = s.T("TestParse", txt).split("|", 2)
            if int(r[0]) != stt or (val is not None and abs(float(r[1]) - val) > 1e-12):
                bad.append((txt, r))
        R.add(c, f"строгий разбор количества: {len(PARSER)} вариантов", not bad, f"расхождения: {bad[:5]}")
        msg = s.T("TestParse", "1,500").split("|", 2)[2]
        R.add(c, "сообщение для неоднозначного ввода подсказывает однозначную форму", "1500" in msg and "1,5" in msg, msg)
    finally:
        s.close()
    typed = ["1,5", "1.5", "0,125", "1,500", "2,1250", "01.05.2026", "1e3", "-1", "1 000", "12.03", "3"]
    res = {}
    for loc in ("ru-RU", "en-US"):
        s = Session(new_wms(f"t28_{loc}"), locale=loc)
        try:
            out = []
            for i, t in enumerate(typed, start=1):
                s.type_in(2, i, t)
                cell = s.tst_cell(2, i)
                out.append((t, cell[0], s.T("TestParseCell", i).split("|")[:2]))
            res[loc] = out
        finally:
            s.close()
    same = [a[2] for a in res["ru-RU"]] == [a[2] for a in res["en-US"]]
    all_text = all(x[1] == "TEXT" for loc in res for x in res[loc])
    R.add(c, "ввод с клавиатуры в текстовые ячейки: одинаковый результат в ru-RU и en-US (1,500 не становится 1500)",
          same and all_text, f"ru: {[(a[0], a[2]) for a in res['ru-RU']]}")


@case
def t29_autoinput():
    c = "T29"
    p = new_wms("t29")
    prof = unique("aiprof")
    s = Session(p, profile=prof)
    during = autoinput_now(s)
    s.close()
    after_close = autoinput_of_profile(prof)
    s = Session(p, profile=prof, fresh_profile=False)
    s.kill()
    after_kill = autoinput_of_profile(prof)
    s = Session(p, profile=prof, fresh_profile=False)
    s.B("ActionForceUnlock")
    s.close()
    after_recovery = autoinput_of_profile(prof)
    R.add(c, "AutoInput: выключен при работе WMS, восстановлен при закрытии; после аварии — исходное значение при следующем чистом закрытии",
          during is False and after_close is True and after_recovery is True,
          f"во время работы {during}; после закрытия {after_close}; после kill {after_kill}; после следующего сеанса {after_recovery}")


@case
def t30_backups():
    c = "T30"
    p = new_wms("t30")
    bdir = os.path.join(os.path.dirname(p), "WMS_Backups")
    s = Session(p)
    s.close()
    s = Session(p)
    try:
        daily = len(glob.glob(os.path.join(bdir, "*_daily.ods")))
        t0 = time.time()
        for _ in range(12):
            s.B("ActionBackupNow")
        dt = (time.time() - t0) / 12 * 1000
        manual = len(glob.glob(os.path.join(bdir, "*_manual*.ods")))
        up = s.B("BeforeUpgrade", module="WmsBackup")
        mig = s.B("BeforeMigration", module="WmsBackup")
        kinds = {k: len(glob.glob(os.path.join(bdir, f"*_{k}*.ods"))) for k in ("daily", "manual", "preupgrade", "premigration")}
        R.add(c, "копии: ежедневная один раз в день, ручная, перед обновлением/миграцией, ротация по видам",
              daily == 1 and manual == 10 and kinds == {"daily": 1, "manual": 10, "preupgrade": 1, "premigration": 1},
              f"{kinds}; {up[:80]}; {mig[:80]}", timing=round(dt, 1))
    finally:
        s.close()


@case
def t31_selfcheck():
    c = "T31"
    s = Session(new_wms("t31"))
    try:
        s.input_rows_api(1, [("a", "1"), ("b", "1")])
        s.T("TestPostRange", 1, 2)
        t0 = time.time()
        sc = s.B("ActionSelfCheck")
        dt = (time.time() - t0) * 1000
        with open(journal_files(s.jdir)[-1], "a", encoding="utf-8") as f:
            f.write(foreign_line(3))
        sc2 = s.B("ActionSelfCheck")
        R.add(c, "самопроверка: чистое состояние — 0 ошибок; чужая строка в журнале — FAIL",
              sc.splitlines()[0].startswith("САМОПРОВЕРКА WMS: ошибок 0") and "FAIL журнал" in sc2,
              sc.replace("\n", " / ")[:500] + " || " + [x for x in sc2.splitlines() if x.startswith("FAIL")][0][:160], timing=round(dt, 1))
    finally:
        s.close()


@case
def t32_startup_error_blocks():
    c = "T32"
    p = new_wms("t32")
    s = Session(p, macros=0)
    s.doc.Sheets.getByName("_SYS").getCellByPosition(0, 4).setString("LAST_SEQX")
    s.doc.store()
    s.close()
    s = Session(p)
    try:
        _, block = st(s)
        post = s.T("TestPost", 1)
        R.add(c, "повреждены служебные данные _SYS → блок, проведение запрещено", block == "SYS_CORRUPT" and post.startswith("BLOCKED"), s.report()[:200])
    finally:
        s.close()
    p = new_wms("t32b")
    s = Session(p, macros=0)
    try:
        before = s.state().get("STATE")
        s.input_row(1, "без события открытия", "1")
        post = s.T("TestPost", 1)
        after = s.state().get("STATE")
        R.add(c, "событие открытия не сработало → проведение само выполняет запуск и только потом пишет",
              before == "" and post == "OK:1" and after == "CLEAN", f"до: {before!r}; {post}; после: {after}")
    finally:
        s.close()


SPECIAL = "a;b|c=d~e%f\r\ng\th" + chr(1) + "Жёлудь «»"


@case
def t33_before_image_special_chars():
    c = "T33"
    p = new_wms("t33")
    s = Session(p)
    sh = s.doc.Sheets.getByName("_TST")
    pre_b = "исходное ;|=~% " + chr(10) + chr(9) + chr(2) + " конец"
    sh.getCellByPosition(1, 1).setString(pre_b)
    sh.getCellByPosition(2, 1).setString("7")
    pre = tuple(s.tst(1))
    s.B("TestSetFault", 2, 2)
    res = s.T("TestPostSpecial", 1, SPECIAL)
    s.kill()
    s = Session(p)
    try:
        unl = s.B("ActionForceUnlock")
        row = tuple(s.tst(1))
        R.add(c, "снимок до изменения со спецсимволами: записать → сохранить → закрыть → открыть → откатить точно",
              res == "CRASH-SIM" and row == pre and "отменена по снимку" in unl, f"до {pre!r}; после {row!r}")
        r2 = s.T("TestPostSpecial", 2, SPECIAL)
        ents, _ = s.journal()
        R.add(c, "спецсимволы в журнале: строка печатная, значение восстанавливается точно",
              r2 == "OK:1" and ents[-1]["fields"]["TEXT"] == SPECIAL and all(32 <= ord(ch) or ch in "\n" for ch in ents[-1]["raw"]),
              f"{r2}; {ents[-1]['raw'][:160]}")
    finally:
        s.close()


@case
def t34_production_book():
    c = "T34"
    p = new_wms("t34", kind="PROD")
    s = Session(p)
    try:
        state, _ = st(s)
        fault = s.B("TestSetFault", 1, 1)
        sc = s.B("ActionSelfCheck")
        has_test = s.doc.Sheets.hasByName("_TST") or s.doc.BasicLibraries.getByName("Standard").hasByName("WmsTestOps")
        R.add(c, "производственная книга: чистый запуск, тестовый шов инертен, тестового кода нет",
              state == "CLEAN" and fault.startswith("REFUSED") and "ошибок 0" in sc.splitlines()[0] and not has_test,
              f"{s.report()[:160]}; шов: {fault}; {sc.splitlines()[0]}")
    finally:
        s.close()


@case
def t35_timing_large_journal():
    c = "T35"
    p = new_wms("t35")
    s = Session(p)
    inst = s.sysv("INSTANCE_ID")
    s.close(save=True)
    jdir = os.path.join(os.path.dirname(p), "WMS_Journal")
    per, nfiles = 5000, 12
    n = per * nfiles
    now = time.localtime()
    months = []
    for k in range(nfiles - 1, -1, -1):
        mm = now.tm_year * 12 + now.tm_mon - 1 - k
        months.append(f"WMS_journal_{mm // 12:04d}-{mm % 12 + 1:02d}.csv")
    seq, data = 0, b""
    for fn in months:
        lines = []
        for _ in range(per):
            seq += 1
            body = f"J1;{seq};2026-09-25T10:00:00;{inst};TEST;NO={seq};ROW=0;TEXT=Позиция {seq};QTY=1"
            lines.append(f"{body};END;{len(body)}")
        data = ("\n".join(lines) + "\n").encode("utf-8")
        open(os.path.join(jdir, fn), "wb").write(data)
    s = Session(p, macros=0)
    s.set_sys("LAST_SEQ", n)
    s.set_sys("TX_SEQ", n)
    s.set_sys("TX_STATE", "COMMITTED")
    s.set_sys("NEXT_NO", n + 1)
    s.set_sys("JOURNAL_POS", f"{months[-1]}|{len(data)}|{n}")
    ec = s.doc.getDocumentProperties().EditingCycles
    s.set_sys("SAVE_STAMP", ec + 1)
    s.doc.store()
    s.close()
    ms = lambda r: int(r.split("запуск WMS ")[1].split(" мс")[0]) if "запуск WMS " in r else -1
    # A: the saved position is confirmed → only what follows it
    s = Session(p)
    rep_fast = s.report()
    s.input_row(1, f"после {n}", "1")
    t0 = time.time()
    res = s.T("TestPost", 1)
    op_ms = (time.time() - t0) * 1000
    t0 = time.time()
    sc = s.B("ActionSelfCheck")
    sc_ms = (time.time() - t0) * 1000
    # B: the position does not match the file content (wrong offset) → the file of the position only
    s.set_sys("JOURNAL_POS", f"{months[-1]}|12345|{n + 1}")
    s.close(save=True)
    s = Session(p)
    rep_file = s.report()
    # C: the file of the position is missing → every file
    s.set_sys("JOURNAL_POS", f"WMS_journal_1999-01.csv|10|{n + 1}")
    s.close(save=True)
    s = Session(p)
    try:
        rep_full = s.report()
        ok = ("работа разрешена" in rep_fast and "не подтвердилась" not in rep_fast and res == f"OK:{n + 1}"
              and "работа разрешена" in rep_file and "прочитан файл " + months[-1] in rep_file
              and "работа разрешена" in rep_full and "журнал прочитан целиком" in rep_full)
        R.add(c, f"журнал {n} записей в {nfiles} месячных файлах: запуск по подтверждённой позиции, по файлу позиции, "
                 f"полное чтение; операция; самопроверка",
              ok,
              f"запуск (позиция подтверждена): {ms(rep_fast)} мс; запуск (позиция не подтвердилась — файл месяца): {ms(rep_file)} мс; "
              f"запуск (файла позиции нет — все файлы): {ms(rep_full)} мс; операция: {op_ms:.0f} мс; "
              f"самопроверка (все файлы): {sc_ms:.0f} мс; {sc.splitlines()[0]}",
              timing=dict(entries=n, files=nfiles, start_fast_ms=ms(rep_fast), start_file_ms=ms(rep_file), start_full_ms=ms(rep_full),
                          op_ms=round(op_ms), selfcheck_ms=round(sc_ms)))
    finally:
        s.close()


# ================================================================ T36–T37 append edge cases

@case
def t36_partial_append():
    c = "T36"
    p = new_wms("t36")
    s = Session(p)
    try:
        s.input_row(1, "база", "1")
        s.T("TestPost", 1)
        # 4: half of the line reaches the file, then the write fails → the operation did not happen
        s.input_row(2, "Половина строки", "2")
        before = tuple(s.tst(2))
        s.B("TestSetFault", 4, 1)
        r4 = s.T("TestPost", 2)
        raw = b"".join(open(f, "rb").read() for f in journal_files(s.jdir))
        R.add(c, "запись в журнал оборвалась посреди строки → операции нет, откат",
              r4.startswith("ERR-RB") and tuple(s.tst(2)) == before and s.sysv("LAST_SEQ") == 1 and not raw.endswith(b"\n"),
              f"{r4[:160]}; строка {tuple(s.tst(2))}; конец журнала {raw[-30:]!r}")
        r4b = s.T("TestPost", 2)
        R.add(c, "повтор: обрывок закрыт переводом строки, операция проведена один раз под тем же seq", r4b == "OK:2", r4b)
        # 5: everything but the final LF, then the write fails → the line is complete, the operation is committed
        s.input_row(3, "Без перевода строки", "3")
        s.B("TestSetFault", 5, 1)
        r5 = s.T("TestPost", 3)
        raw = b"".join(open(f, "rb").read() for f in journal_files(s.jdir))
        tail = journal_oracle.parse_line(raw.rsplit(b"\n", 1)[-1].decode("utf-8"))
        R.add(c, "записана вся строка без перевода строки → операция зафиксирована (строку примет любой читатель)",
              r5.startswith("OK:3") and tuple(s.tst(3))[3] == "Проведено" and s.sysv("LAST_SEQ") == 3 and tail is not None and tail["seq"] == 3,
              f"{r5[:160]}; {tuple(s.tst(3))}; конец журнала {raw[-30:]!r}")
        P, inf = s.check()
        R.add(c, "оракул: seq 1..3 непрерывны, обрывок — одна повреждённая строка", not P and inf["damaged"] == 1, f"{inf}; {P[:3]}")
        s.close(save=True)
        s = Session(p)
        rep = s.report()
        state, _ = st(s)
        s.input_row(4, "после перезапуска", "1")
        r = s.T("TestPost", 4)
        P, inf = s.check()
        R.add(c, "сохранить и перезапустить: работа разрешена; следующая операция дописана с новой строки",
              state == "CLEAN" and r == "OK:4" and "повреждённых строк журнала: 1" in rep and not P and inf["damaged"] == 1,
              f"{rep[:300]}; {r}; {inf}; {P[:3]}")
    finally:
        s.close()


@case
def t37_clock_behind_journal():
    c = "T37"
    p = new_wms("t37")
    jdir = os.path.join(os.path.dirname(p), "WMS_Journal")
    s = Session(p)
    s.input_rows_api(1, [(f"Позиция {i}", "1") for i in range(1, 4)])
    s.T("TestPostRange", 1, 2)
    s.close(save=True)
    # as if the two entries had been written while the computer's clock showed December 2099 (then the date was fixed)
    os.rename(journal_files(jdir)[-1], os.path.join(jdir, "WMS_journal_2099-12.csv"))
    s = Session(p)
    try:
        rep = s.report()
        state, _ = st(s)
        r = s.T("TestPost", 3)
        files = [os.path.basename(f) for f in journal_files(jdir)]
        P, inf = s.check()
        R.add(c, "дата компьютера раньше последнего файла журнала → запись продолжается в последнем файле, порядок seq не нарушен",
              state == "CLEAN" and "проверьте часы" in rep and r == "OK:3" and files == ["WMS_journal_2099-12.csv"] and not P,
              f"{rep[:320]}; {r}; файлы {files}; {inf}; {P[:3]}")
        s.close(save=True)
        s = Session(p)
        rep2 = s.report()
        R.add(c, "после сохранения — запуск по подтверждённой позиции", st(s)[0] == "CLEAN" and "не подтвердилась" not in rep2, rep2[:240])
    finally:
        s.close()


# ================================================================ T38 completion after the journal fails (COMMIT_FAILED)

def undo_probe(s, r):
    """ordinary manual input in Calc, then Ctrl+Z: True when the input is recorded and undone (Undo is not locked)"""
    s.type_in(1, r, "ручной ввод")
    typed = tuple(s.tst(r))[1] == "ручной ввод"
    s.ui(".uno:Undo")
    return typed and tuple(s.tst(r))[1] == ""


def book_snapshot(s, rows=5):
    return ([tuple(s.tst(r)) for r in range(1, rows)], [s.locks(r) for r in range(1, rows)],
            [s.sysv(k) for k in ("LAST_SEQ", "NEXT_NO", "TX_STATE", "TX_SEQ", "TX_BEFORE_IMAGE", "JOURNAL_POS", "REGISTERED_URL")])


def ctrl_z_harmless(s, n=5):
    """Ctrl+Z pressed n times changes nothing WMS wrote: rows, cell protection, _SYS"""
    before = book_snapshot(s)
    for _ in range(n):
        s.ui(".uno:Undo")
    return before == book_snapshot(s)


def key_rows(s, key, rows=12):
    return [r for r in range(1, rows) if tuple(s.tst(r))[0] == float(key) and tuple(s.tst(r))[3] == "Проведено"]


def journal_ops(s, no):
    ents, _ = s.journal()
    return [e for e in ents if e["type"] == "TEST" and e["fields"].get("NO") == str(no)]


@case
def t38_commit_failed_undo():
    c = "T38"
    # A/B: the completion fails again when retried (mode 3) → COMMIT_FAILED; A saves the book in that state, B does not
    for variant, point, save in (("A6", 6, True), ("A7", 7, True), ("B6", 6, False)):
        p = new_wms(f"t38_{variant}")
        s = Session(p)
        try:
            s.input_row(1, "база", "1")
            base = s.T("TestPost", 1)
            if not save:
                s.doc.store()
            s.input_row(2, "Кабель ВВГ 3х2,5", "12,5")
            nj = len(journal_ops(s, 2))
            s.B("TestSetFault", point, 3)
            res = s.T("TestPost", 2)
            s.B("TestSetFault", 0, 0)
            stt = s.state()
            jops = journal_ops(s, 2)
            R.add(c, f"{variant}: операция дошла до точки фиксации — строка журнала seq 2 записана до сбоя",
                  base == "OK:1" and nj == 0 and len(jops) == 1 and jops[0]["seq"] == 2,
                  f"до: {nj}; после: {[(e['seq'], e['fields'].get('NO')) for e in jops]}")
            harmless = ctrl_z_harmless(s)
            ok_undo = undo_probe(s, 5)
            blocked = s.T("TestPost", 3)
            stt2 = s.state()
            R.add(c, f"{variant}: ошибка завершения в книге (точка {point}, повтор тоже падает) → COMMIT_FAILED; стек Undo очищен, "
                     f"Ctrl+Z ×5 не меняет записи WMS; Undo не заблокирован — ручной ввод отменяется; проведение запрещено",
                  res.startswith("ERR-CRITICAL") and stt["BLOCK"] == "COMMIT_FAILED" and stt["LAST_SEQ"] == "1"
                  and stt["UNDO_LOCKED"] == "False" and stt["UNDO_CAN"] == "False" and harmless and ok_undo
                  and stt2["UNDO_LOCKED"] == "False" and blocked.startswith("BLOCKED"),
                  f"{res[:150]}; {stt}; Ctrl+Z безвреден: {harmless}; ручной ввод+Ctrl+Z: {ok_undo}; следующее проведение: {blocked[:70]}")
        finally:
            s.close(save=save)
        s = Session(p)
        try:
            rep = s.report()
            _, block = st(s)
            harmless_start = ctrl_z_harmless(s)
            rec = s.B("ActionRecover")
            state, _ = st(s)
            stt = s.state()
            harmless_rec = ctrl_z_harmless(s)
            expect_rec = "уже были в книге: 1" if point == 7 else "восстановлено операций: 1"
            R.add(c, f"{variant}: {'сохранить' if save else 'не сохранять'} → перезапуск → «Восстановить» доводит операцию; "
                     f"Ctrl+Z после запуска и после восстановления ничего не откатывает; Undo не заблокирован",
                  block == "TAIL" and expect_rec in rec and state == "CLEAN" and stt["LAST_SEQ"] == "2" and stt["UNDO_LOCKED"] == "False"
                  and harmless_start and harmless_rec,
                  f"запуск: {rep[:200]}; «Восстановить»: {rec[:110]}; {stt}; Ctrl+Z после запуска / после восстановления безвреден: "
                  f"{harmless_start} / {harmless_rec}")
            ok_undo = undo_probe(s, 5)
            s.input_row(3, "следующая", "1")
            nxt = s.T("TestPost", 3)
            P, inf = s.check()
            once = len(key_rows(s, 2)) == 1 and len(journal_ops(s, 2)) == 1
            R.add(c, f"{variant}: ручной ввод и следующая операция работают; операция seq 2 ровно один раз; "
                     f"журнал, LAST_SEQ и книга согласованы",
                  ok_undo and nxt == "OK:3" and once and not P and inf["last_seq"] == 3 and inf["journal_max"] == 3,
                  f"ручной ввод+Ctrl+Z: {ok_undo}; {nxt}; строк с №2: {len(key_rows(s, 2))}; записей журнала №2: "
                  f"{len(journal_ops(s, 2))}; {inf}; {P[:3]}")
        finally:
            s.close()
    # C: the book is stored in the middle of the completion and the process dies (points 6, 7, 8)
    for point in (6, 7, 8):
        p = new_wms(f"t38_C{point}")
        s = Session(p)
        s.input_row(1, "база", "1")
        s.T("TestPost", 1)
        s.input_row(2, "Авария при завершении", "3")
        s.B("TestSetFault", point, 2)
        res = s.T("TestPost", 2)
        s.kill()
        s = Session(p)
        try:
            unl = s.B("ActionForceUnlock")
            _, block = st(s)
            harmless_start = ctrl_z_harmless(s)
            rec = s.B("ActionRecover") if block == "TAIL" else ""
            state, _ = st(s)
            stt = s.state()
            harmless_rec = ctrl_z_harmless(s)
            ok_undo = undo_probe(s, 5)
            s.input_row(3, "следующая", "1")
            nxt = s.T("TestPost", 3)
            P, inf = s.check()
            once = len(key_rows(s, 2)) == 1 and len(journal_ops(s, 2)) == 1
            ok = (res == "CRASH-SIM" and state == "CLEAN" and stt["UNDO_LOCKED"] == "False" and harmless_start and harmless_rec
                  and ok_undo and nxt == "OK:3" and once and not P)
            R.add(c, f"C{point}: книга сохранена посреди завершения (точка {point}) + kill → после перезапуска операция ровно "
                     f"один раз, всё согласовано; Ctrl+Z ничего не откатывает, Undo не заблокирован",
                  ok, f"{res}; после «Снять блокировку»: {unl[:230]}; «Восстановить»: {rec[:90]}; Ctrl+Z безвреден: "
                      f"{harmless_start}/{harmless_rec}; {nxt}; {inf}; {P[:3]}")
        finally:
            s.close()


# ================================================================ runner

def main():
    only = set(a.lower() for a in sys.argv[1:])
    t0 = time.time()
    template("TEST")
    template("PROD")
    for fn in CASES:
        if only and not any(fn.__name__.startswith(o) for o in only):
            continue
        try:
            fn()
        except Exception as e:
            R.error(fn.__name__, e)
        subprocess.run(["pkill", "-9", "-f", "profile_.*" + str(os.getpid())], capture_output=True)
    R.save(os.path.join(OUT, "results_phase1.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_phase1.md"), "w", encoding="utf-8") as f:
        f.write(f"| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
