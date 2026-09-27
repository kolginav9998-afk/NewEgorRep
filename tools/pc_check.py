#!/usr/bin/env python3
"""Проверка складского ПК перед вводом WMS в работу (FINAL WMS MARATHON §10, M5; docs/PC_VALIDATION.md).

    python3 tools/pc_check.py [--full] [--env-only] [--out ПАПКА] [--display :N]

Запускается из папки исходников WMS (этого репозитория). Рабочую папку WMS не трогает: книги, журналы и отчёты тестов
создаются только в папке результатов (по умолчанию pc_check_results/<ГГГГММДД-ЧЧММСС>/ рядом с исходниками); у каждого
теста свой профиль LibreOffice, открытая у кладовщика WMS не затрагивается.

1. Среда: система, Python, LibreOffice (версия; проверена ли с WMS), Python-UNO, Writer, шрифты с кириллицей, локаль
   ru_RU.UTF-8, свободное место, Xvfb, CUPS (служба, принтеры, принтер по умолчанию), устройства /dev/usb/lp*.
2. Компиляция всех модулей WMS и всех книг WMS_TOOLBOX — в своём невидимом дисплее (Xvfb).
3. Тесты: книга-кандидат (C), перенос (MG), Final Core (Y), инструменты (K), сквозной тест (E); с --full ещё этапы 5…1
   и проверки окон (≈ 40 минут). --env-only — только шаг 1.
4. PC_CHECK_REPORT.md, pc_check.json и архив <папка>.tar.gz (отчёты и журналы тестов, без книг) — прислать разработчику.

Код выхода: 0 — всё прошло; 1 — есть провалы тестов или ошибки среды; 2 — проверка не запустилась.
"""
import argparse
import datetime
import glob
import hashlib
import json
import os
import platform
import re
import shutil
import socket
import subprocess
import sys
import tarfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TESTED_LO = ("24.2", "26.2")

# key, script, title, needs a display
QUICK = [("candidate", "tests/run_candidate.py", "Книга-кандидат (C)", False),
         ("migration", "tests/run_migration.py", "Перенос (MG)", False),
         ("final_core", "tests/run_final_core.py", "Final Core (Y)", False),
         ("toolbox", "tests/run_toolbox.py", "Инструменты (K)", False),
         ("e2e", "tests/run_e2e.py", "Сквозной тест (E)", False),
         ("pc_tools", "tests/run_pc_tools.py", "Пакет проверки ПК и принтера (P)", False)]
FULL = [("phase5", "tests/run_phase5.py", "Этап 5 (X)", False), ("phase4", "tests/run_phase4.py", "Этап 4 (W)", False),
        ("phase3", "tests/run_phase3.py", "Этап 3 (R)", False), ("phase2", "tests/run_phase2.py", "Этап 2 (V)", False),
        ("phase1", "tests/run_phase1.py", "Этап 1 (T)", False),
        ("gui_candidate", "tests/gui_candidate.py", "Окна: кандидат", True), ("gui_final_core", "tests/gui_final_core.py", "Окна: Final Core", True),
        ("gui_phase5", "tests/gui_phase5.py", "Окна: этап 5", True), ("gui_phase4", "tests/gui_phase4.py", "Окна: этап 4", True),
        ("gui_phase3", "tests/gui_phase3.py", "Окна: этап 3", True), ("gui_phase2", "tests/gui_phase2.py", "Окна: этап 2", True)]


def run(cmd, timeout=60, **kw):
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, **kw)
        return r.returncode, (r.stdout or "") + (r.stderr or "")
    except FileNotFoundError:
        return 127, ""
    except subprocess.TimeoutExpired as e:
        return 124, f"превышено время {timeout} с: {e}"


def fingerprint():
    """SHA-256 (16 знаков) исходников, которыми проверяется ПК: src, tools, tests, docs"""
    h = hashlib.sha256()
    pats = ["src/basic/*.bas", "src/toolbox/*.bas", "src/toolbox/scripts/*/*.py", "tools/*.py", "tests/*.py", "tests/basic/*.bas", "docs/*.md"]
    for p in sorted(f for pat in pats for f in glob.glob(os.path.join(ROOT, pat))):
        h.update(os.path.relpath(p, ROOT).encode())
        h.update(open(p, "rb").read())
    return h.hexdigest()[:16]


# ================================================================ 1. the environment

def lo_writer():
    for exe in ("soffice", "libreoffice"):
        path = shutil.which(exe)
        if path:
            prog = os.path.dirname(os.path.realpath(path))
            if glob.glob(os.path.join(prog, "libswlo.so")) or os.path.exists(os.path.join(os.path.dirname(prog), "share", "registry", "writer.xcd")):
                return True
    return False


def environment():
    """[(уровень OK / WARN / FAIL, что, результат, что сделать)]"""
    rows = []
    osr = {}
    if os.path.exists("/etc/os-release"):
        for ln in open("/etc/os-release", encoding="utf-8", errors="replace"):
            k, _, v = ln.strip().partition("=")
            osr[k] = v.strip('"')
    name = osr.get("PRETTY_NAME", platform.platform())
    rows.append(("OK" if osr.get("ID") in ("ubuntu", "debian", "linuxmint") or "ubuntu" in osr.get("ID_LIKE", "") else "WARN", "система",
                 f"{name}; {platform.machine()}; {socket.gethostname()}", "" if osr.get("ID") in ("ubuntu", "debian", "linuxmint") else
                 "WMS проверялась на Linux (Ubuntu); другая система — на ваш риск"))
    rows.append(("OK" if sys.version_info >= (3, 8) else "FAIL", "Python", sys.version.split()[0], "" if sys.version_info >= (3, 8) else "нужен Python 3.8+"))
    rc, out = run(["soffice", "--version"])
    m = re.search(r"(\d+\.\d+)\.\d+(\.\d+)?", out)
    if rc != 0 or not m:
        rows.append(("FAIL", "LibreOffice", "не найден (soffice)", "sudo apt install libreoffice-calc libreoffice-writer"))
    else:
        mm = m.group(1)
        ok = mm in TESTED_LO
        rows.append(("OK" if ok else "WARN", "LibreOffice", out.strip().splitlines()[0][:90],
                     "" if ok else f"с WMS проверены {', '.join(TESTED_LO)}: эта версия проверяется этим пакетом"))
    rc, out = run([sys.executable, "-c", "import uno"])
    rows.append(("OK" if rc == 0 else "FAIL", "Python-UNO", "есть" if rc == 0 else "нет (import uno)", "" if rc == 0 else "sudo apt install python3-uno"))
    w = lo_writer()
    rows.append(("OK" if w else "WARN", "Writer", "установлен" if w else "не найден", "" if w else "sudo apt install libreoffice-writer (акты WMS_DOCS)"))
    rc, out = run(["fc-list", ":lang=ru", "family"])
    fams = sorted({f.split(",")[0].strip() for f in out.splitlines() if f.strip()}) if rc == 0 else []
    good = [f for f in fams if re.match(r"(DejaVu|Liberation|Noto|PT |Carlito|Caladea)", f)]
    rows.append(("OK" if good else "WARN", "шрифты", (", ".join(good[:4]) or "нет шрифтов с кириллицей") + (f" (всего семейств с кириллицей {len(fams)})" if fams else ""),
                 "" if good else "sudo apt install fonts-dejavu fonts-liberation"))
    rc, out = run(["locale", "-a"])
    ru = any(x.lower().replace("-", "") in ("ru_ru.utf8",) for x in out.split())
    rows.append(("OK" if ru else "WARN", "локаль", "ru_RU.UTF-8 есть" if ru else "ru_RU.UTF-8 нет",
                 "" if ru else "sudo locale-gen ru_RU.UTF-8 (даты и числа в окнах LibreOffice по-русски)"))
    free = shutil.disk_usage(ROOT).free / 2 ** 30
    rows.append(("OK" if free >= 2 else ("WARN" if free >= 0.5 else "FAIL"), "место", f"свободно {free:.1f} ГБ ({ROOT})",
                 "" if free >= 2 else "освободите место: WMS и резервные копии растут"))
    xv = shutil.which("Xvfb")
    rows.append(("OK" if xv else "WARN", "Xvfb", "есть" if xv else "нет",
                 "" if xv else "sudo apt install xvfb (без него компиляция и проверки окон пропускаются)"))
    if not shutil.which("lpstat"):
        rows.append(("WARN", "CUPS", "не установлен (нет lpstat)", "sudo apt install cups — печать этикеток через системный принтер"))
    else:
        rc, out = run(["lpstat", "-r"])
        running = "is running" in out or "запущен" in out or "работает" in out
        rc2, pr = run(["lpstat", "-p", "-d"])
        printers = [ln.split()[1] for ln in pr.splitlines() if ln.startswith(("printer ", "принтер "))]
        default = re.search(r"(?:system default destination|назначение по умолчанию)[^:]*:\s*(\S+)", pr)
        rows.append(("OK" if running else "WARN", "CUPS", "служба работает" if running else f"служба не отвечает: {out.strip()[:80]}",
                     "" if running else "sudo systemctl enable --now cups"))
        rows.append(("OK" if printers else "WARN", "принтеры", (", ".join(printers) or "нет принтеров") + (f"; по умолчанию {default.group(1)}" if default else ""),
                     "" if printers else "добавьте принтер: Параметры → Принтеры (или http://localhost:631); принтер без драйвера — TSPL"))
    devs = sorted(glob.glob("/dev/usb/lp*"))
    for d in devs:
        wr = os.access(d, os.W_OK)
        rows.append(("OK" if wr else "WARN", "USB-принтер", f"{d} — {'запись разрешена' if wr else 'нет прав на запись'}",
                     "" if wr else "sudo usermod -aG lp $USER, затем выйти и войти в систему"))
    if not devs:
        rows.append(("OK", "USB-принтер", "устройств /dev/usb/lp* нет (принтер через CUPS или не подключён)", ""))
    return rows


# ================================================================ 2–3. the checks

class Xvfb:
    """a private display for the compile checks and the window tests (nothing appears on the screen of the PC)"""

    def __init__(self, want=None):
        self.proc, self.display = None, want
        if want or not shutil.which("Xvfb"):
            return
        for n in range(97, 89, -1):
            if os.path.exists(f"/tmp/.X11-unix/X{n}") or os.path.exists(f"/tmp/.X{n}-lock"):
                continue
            self.proc = subprocess.Popen(["Xvfb", f":{n}", "-screen", "0", "1600x1000x24", "-nolisten", "tcp"], stdout=subprocess.DEVNULL,
                                         stderr=subprocess.DEVNULL, start_new_session=True)
            for _ in range(50):
                if os.path.exists(f"/tmp/.X11-unix/X{n}"):
                    self.display = f":{n}"
                    return
                time.sleep(0.1)
            self.stop()

    def stop(self):
        if self.proc:
            self.proc.terminate()
            try:
                self.proc.wait(10)
            except subprocess.TimeoutExpired:
                self.proc.kill()
            self.proc = None


def results_of(d):
    """PASS / FAIL / SKIP and the checks that did not pass, from results_*.json of a suite"""
    n, bad = {"PASS": 0, "FAIL": 0, "SKIP": 0}, []
    for f in glob.glob(os.path.join(d, "results_*.json")):
        for it in json.load(open(f, encoding="utf-8")).get("items", []):
            n[it["status"]] = n.get(it["status"], 0) + 1
            if it["status"] != "PASS":
                bad.append(f"{it['case']} {it['status']}: {it['test'][:160]} — {str(it['detail'])[:300]}")
    return n, bad


def step_suite(out, key, script, title, gui, display, timeout=5400):
    t0 = time.time()
    d = os.path.join(out, key)
    env = dict(os.environ, WMS_TEST_OUT=d)
    cmd = [sys.executable, os.path.join(ROOT, script)] + (["--display", display] if gui else [])
    rc, log = run(cmd, timeout=timeout, env=env, cwd=ROOT)
    open(os.path.join(out, key + ".log"), "w", encoding="utf-8").write(log)
    if gui:
        m = re.search(r"GUI: (\d+) из (\d+) PASS", log)
        n = {"PASS": int(m.group(1)), "FAIL": int(m.group(2)) - int(m.group(1)), "SKIP": 0} if m else {"PASS": 0, "FAIL": 1, "SKIP": 0}
        bad = [ln for ln in log.splitlines() if ln.startswith("[FAIL]")][:20] or ([] if m else [log.strip()[-300:]])
    else:
        n, bad = results_of(d)
        if rc != 0 and not n["FAIL"]:
            n["FAIL"] += 1
            bad.append(f"{title}: код {rc} — {log.strip()[-300:]}")
    return dict(key=key, title=title, rc=rc, seconds=round(time.time() - t0), bad=bad, **{k.lower(): v for k, v in n.items()})


def step_compile(out, display):
    steps = []
    t0 = time.time()
    rc, log = run([sys.executable, os.path.join(ROOT, "tools", "basic_compile_check.py"), "--display", display], timeout=1200, cwd=ROOT)
    open(os.path.join(out, "compile_wms.log"), "w", encoding="utf-8").write(log)
    ok = log.startswith("COMPILE OK")
    steps.append(dict(key="compile_wms", title="Компиляция модулей WMS", rc=rc, seconds=round(time.time() - t0), bad=[] if ok else [log.strip()[-400:]],
                      **{"pass": 1 if ok else 0, "fail": 0 if ok else 1, "skip": 0}))
    t0 = time.time()
    tb = os.path.join(out, "toolbox_build")
    shutil.rmtree(tb, ignore_errors=True)
    rc1, log1 = run([sys.executable, os.path.join(ROOT, "tools", "build_toolbox.py"), tb], timeout=1200, cwd=ROOT)
    rc2, log2 = run([sys.executable, os.path.join(ROOT, "tools", "tb_compile_check.py"), os.path.join(tb, "WMS_TOOLBOX"), "--display", display],
                    timeout=1200, cwd=ROOT)
    open(os.path.join(out, "compile_toolbox.log"), "w", encoding="utf-8").write(log1 + "\n" + log2)
    ok = rc1 == 0 and log2.startswith("COMPILE OK")
    steps.append(dict(key="compile_toolbox", title="Компиляция книг WMS_TOOLBOX", rc=rc2, seconds=round(time.time() - t0),
                      bad=[] if ok else [(log1 + log2).strip()[-400:]], **{"pass": 1 if ok else 0, "fail": 0 if ok else 1, "skip": 0}))
    return steps


# ================================================================ 4. the report

def report(out, info):
    env, steps = info["env"], info["steps"]
    nfail = sum(s["fail"] for s in steps) + sum(1 for e in env if e[0] == "FAIL")
    nwarn = sum(1 for e in env if e[0] == "WARN") + sum(s["skip"] for s in steps)
    verdict = "ВСЁ ПРОШЛО" if not nfail and not nwarn else ("ЕСТЬ ПРОВАЛЫ" if nfail else "ПРОШЛО, ЕСТЬ ЗАМЕЧАНИЯ")
    t = [f"# Проверка складского ПК — {info['host']}", "",
         f"**Итог: {verdict}** (провалов {nfail}, замечаний {nwarn})  ",
         f"**Когда:** {info['started']} — {info['finished']}  ",
         f"**Режим:** {info['mode']}; исходники WMS `{info['code_start']}`" + ("" if info['code_start'] == info['code_end'] else f" → `{info['code_end']}` (менялись во время проверки!)") + "  ",
         f"**Папка результатов:** `{out}`", "", "## Среда", "", "| | Проверка | Результат | Что сделать |", "|---|---|---|---|"]
    mark = {"OK": "✓", "WARN": "!", "FAIL": "✗"}
    for lvl, what, res, fix in env:
        t.append(f"| {mark[lvl]} | {what} | {res.replace('|', '¦')} | {fix.replace('|', '¦')} |")
    if steps:
        t += ["", "## Проверки WMS", "", "| Проверка | PASS | FAIL | SKIP | Время |", "|---|---:|---:|---:|---:|"]
        for s in steps:
            t.append(f"| {s['title']} | {s['pass']} | {s['fail']} | {s['skip']} | {s['seconds']} с |")
        bad = [(s["title"], b) for s in steps for b in s["bad"]]
        if bad:
            t += ["", "## Что не прошло", ""] + [f"- **{a}**: {b.replace(chr(10), ' ')}" for a, b in bad[:80]]
    t += ["", "## Дальше", "",
          "1. Принтер этикеток: `python3 tools/printer_check.py` (список), затем `--cups ИМЯ` или `--tspl /dev/usb/lp0` (docs/PC_VALIDATION.md).",
          f"2. Прислать разработчику `{os.path.basename(out)}.tar.gz` (рядом с папкой результатов) и `printer_check.txt`.", ""]
    open(os.path.join(out, "PC_CHECK_REPORT.md"), "w", encoding="utf-8").write("\n".join(t))
    json.dump(info, open(os.path.join(out, "pc_check.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return verdict, nfail, nwarn


def pack(out):
    """<папка>.tar.gz: отчёты, журналы тестов, результаты; без книг ODS и профилей LibreOffice"""
    arc = out.rstrip("/") + ".tar.gz"

    def keep(ti):
        name = ti.name
        if "/profiles" in name or "/profile_" in name or name.endswith(".ods") or "__pycache__" in name:
            return None
        return ti
    with tarfile.open(arc, "w:gz") as tf:
        tf.add(out, arcname=os.path.basename(out), filter=keep)
    return arc


def main(argv=None):
    ap = argparse.ArgumentParser(description="проверка складского ПК перед вводом WMS в работу")
    ap.add_argument("--full", action="store_true", help="ещё этапы 5…1 и проверки окон (≈ 40 минут)")
    ap.add_argument("--env-only", action="store_true", help="только проверка среды (секунды)")
    ap.add_argument("--out", help="папка результатов (по умолчанию pc_check_results/<время>)")
    ap.add_argument("--display", help="готовый дисплей X для компиляции и окон (по умолчанию — свой Xvfb)")
    a = ap.parse_args(argv)
    if not os.path.exists(os.path.join(ROOT, "tools", "build_ods.py")):
        print("ОШИБКА: запускайте из папки исходников WMS (нет tools/build_ods.py)", file=sys.stderr)
        return 2
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    out = os.path.abspath(a.out or os.path.join(ROOT, "pc_check_results", stamp))
    if os.path.exists(out) and os.listdir(out):
        print(f"ОШИБКА: папка {out} не пуста — укажите новую (--out)", file=sys.stderr)
        return 2
    os.makedirs(out, exist_ok=True)
    info = dict(host=socket.gethostname(), started=datetime.datetime.now().strftime("%d.%m.%Y %H:%M"), mode="только среда" if a.env_only else
                ("полная (--full)" if a.full else "быстрая"), code_start=fingerprint(), env=environment(), steps=[])
    for lvl, what, res, _ in info["env"]:
        print(f"{lvl:<5} {what:<12} {res}")
    blocking = [e for e in info["env"] if e[0] == "FAIL" and e[1] in ("LibreOffice", "Python-UNO", "Python")]
    if not a.env_only and not blocking:
        xv = Xvfb(a.display)
        try:
            if xv.display:
                print(f"\nдисплей для компиляции и окон: {xv.display}")
                for s in step_compile(out, xv.display):
                    info["steps"].append(s)
                    print(f"{s['title']}: PASS {s['pass']}, FAIL {s['fail']} ({s['seconds']} с)")
            else:
                info["steps"].append(dict(key="compile", title="Компиляция (нет Xvfb — пропущена)", rc=0, seconds=0, bad=[], **{"pass": 0, "fail": 0, "skip": 1}))
            for key, script, title, gui in QUICK + (FULL if a.full else []):
                if gui and not xv.display:
                    info["steps"].append(dict(key=key, title=title + " (нет Xvfb — пропущено)", rc=0, seconds=0, bad=[], **{"pass": 0, "fail": 0, "skip": 1}))
                    continue
                print(f"{title} …", flush=True)
                s = step_suite(out, key, script, title, gui, xv.display)
                info["steps"].append(s)
                print(f"   PASS {s['pass']}, FAIL {s['fail']}, SKIP {s['skip']} ({s['seconds']} с)")
        finally:
            xv.stop()
    info["finished"] = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    info["code_end"] = fingerprint()
    verdict, nfail, nwarn = report(out, info)
    arc = pack(out)
    print(f"\nИТОГ: {verdict} (провалов {nfail}, замечаний {nwarn})\nотчёт: {os.path.join(out, 'PC_CHECK_REPORT.md')}\nприслать: {arc}")
    if blocking:
        print("ОСТАНОВЛЕНО: без LibreOffice / Python-UNO проверки WMS не запускаются — установите пакеты и повторите", file=sys.stderr)
    return 1 if nfail or blocking else 0


if __name__ == "__main__":
    sys.exit(main())
