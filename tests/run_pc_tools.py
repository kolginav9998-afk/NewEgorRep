"""The check package of the warehouse PC (FINAL WMS MARATHON, M5; docs/PC_VALIDATION.md): tools/pc_check.py (the
environment, the report, the archive, a result folder never reused) and tools/printer_check.py (the list of printers, the
PDF of the test labels made by WMS_LABELS, the TSPL commands, the answer recorded, a failed print named). Nothing is
printed on a real printer: the TSPL goes to a file, a CUPS printer that does not exist is refused.

    WMS_TEST_OUT=/tmp/wms_pc python3 tests/run_pc_tools.py [case ...]

Results: WMS_TEST_OUT/results_pc_tools.json and WMS_TEST_OUT/TEST_REPORT_pc_tools.md.
"""
import json
import os
import shutil
import subprocess
import sys
import tarfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from harness import OUT, Results  # noqa: E402

R = Results()
CASES = []


def case(fn):
    CASES.append(fn)
    return fn


def run(script, *args, env=None):
    e = dict(os.environ, **(env or {}))
    r = subprocess.run([sys.executable, os.path.join(ROOT, "tools", script)] + list(args), capture_output=True, text=True, timeout=1200, env=e,
                       stdin=subprocess.DEVNULL)
    return r.returncode, r.stdout + r.stderr


@case
def p01_pc_check_env():
    c = "P01"
    out = os.path.join(OUT, "cases", "p01")
    shutil.rmtree(out, ignore_errors=True)
    if os.path.exists(out + ".tar.gz"):
        os.remove(out + ".tar.gz")
    rc, log = run("pc_check.py", "--env-only", "--out", out)
    rep = open(os.path.join(out, "PC_CHECK_REPORT.md"), encoding="utf-8").read() if os.path.exists(os.path.join(out, "PC_CHECK_REPORT.md")) else ""
    info = json.load(open(os.path.join(out, "pc_check.json"), encoding="utf-8")) if os.path.exists(os.path.join(out, "pc_check.json")) else {}
    env = {e[1]: e[0] for e in info.get("env", [])}
    names = tarfile.open(out + ".tar.gz").getnames() if os.path.exists(out + ".tar.gz") else []
    R.add(c, "pc_check --env-only: среда (система, Python, LibreOffice, Python-UNO, Writer, шрифты, локаль, место, Xvfb, CUPS, USB-принтер) — "
             "в PC_CHECK_REPORT.md с тем, что сделать; pc_check.json; архив с отчётом; код 0 или 1",
          rc in (0, 1) and "## Среда" in rep and env.get("LibreOffice") in ("OK", "WARN") and env.get("Python-UNO") == "OK"
          and all(k in env for k in ("система", "Python", "Writer", "шрифты", "локаль", "место", "Xvfb", "CUPS", "USB-принтер"))
          and info.get("code_start") == info.get("code_end") and any(n.endswith("PC_CHECK_REPORT.md") for n in names), f"{rc}; {env}; {names}")
    rc2, log2 = run("pc_check.py", "--env-only", "--out", out)
    R.add(c, "папка результатов не используется второй раз — отказ, код 2", rc2 == 2 and "не пуста" in log2, log2.strip()[-120:])


@case
def p02_printer_check():
    c = "P02"
    res = os.path.join(OUT, "cases", "p02")
    shutil.rmtree(res, ignore_errors=True)
    env = {"WMS_PC_RESULTS": res}
    rc0, lst = run("printer_check.py", env=env)
    R.add(c, "printer_check без ключей: принтеры CUPS (или как установить CUPS), устройства /dev/usb/lp*, следующие команды; код 0",
          rc0 == 0 and "--cups" in lst and "--tspl" in lst and ("CUPS" in lst), lst[:200])
    rc1, o1 = run("printer_check.py", "--pdf-only", env=env)
    pdfs = [f for f in os.listdir(res) if f.startswith("test_labels_") and f.endswith(".pdf")] if os.path.isdir(res) else []
    pages = open(os.path.join(res, pdfs[0]), "rb").read().count(b"/Type /Page\n") + open(os.path.join(res, pdfs[0]), "rb").read().count(b"/Type/Page/") \
        if pdfs else 0
    R.add(c, "--pdf-only: PDF тестовых этикеток построен кнопками WMS_LABELS на фикстуре снимка (3 ЕИ), без печати",
          rc1 == 0 and len(pdfs) == 1 and os.path.getsize(os.path.join(res, pdfs[0])) > 5000, f"{o1.strip()[-160:]}; страниц {pages}")
    prn = os.path.join(res, "test.prn")
    rc2, o2 = run("printer_check.py", "--tspl", prn, "--size", "58x30", "--dpmm", "12", "--answer", "да", env=env)
    tspl = open(prn, "rb").read().decode("cp1251") if os.path.exists(prn) else ""
    rec = open(os.path.join(res, "printer_check.txt"), encoding="utf-8").read() if os.path.exists(os.path.join(res, "printer_check.txt")) else ""
    R.add(c, "--tspl в файл (58×30 мм, 12 точек/мм): команды SIZE, CODEPAGE 1251, ЕИ крупно, штрихкод 128 с номером ЕИ; ответ «да» записан в "
             "printer_check.txt",
          rc2 == 0 and "SIZE 58.0 mm,30.0 mm" in tspl and "CODEPAGE 1251" in tspl and '"ЕИ-00000201"' in tspl and '"128"' in tspl
          and '"00000206"' in tspl and "напечатано верно" in rec, f"{o2.strip()[-160:]}; {tspl[:120]!r}")
    rc3, o3 = run("printer_check.py", "--cups", "WMS_TEST_NO_SUCH_PRINTER", "--answer", "да", env=env)
    rec = open(os.path.join(res, "printer_check.txt"), encoding="utf-8").read()
    R.add(c, "--cups с несуществующим принтером (или без CUPS) — отказ с причиной, код 2; неудача записана",
          rc3 == 2 and "ПЕЧАТЬ НЕ УДАЛАСЬ" in rec, o3.strip()[-200:])


def main():
    only = set(a.lower() for a in sys.argv[1:])
    t0 = time.time()
    for fn in CASES:
        if only and not any(fn.__name__.startswith(o) for o in only):
            continue
        try:
            fn()
        except Exception as e:
            R.error(fn.__name__, e)
    R.save(os.path.join(OUT, "results_pc_tools.json"))
    n = {k: sum(1 for i in R.items if i["status"] == k) for k in ("PASS", "FAIL", "SKIP")}
    print(f"\nИТОГО: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} за {time.time() - t0:.0f} с; результаты: {OUT}")
    with open(os.path.join(OUT, "TEST_REPORT_pc_tools.md"), "w", encoding="utf-8") as f:
        f.write("| Кейс | Проверка | Итог | Подробности |\n|---|---|---|---|\n")
        for i in R.items:
            det = str(i["detail"]).replace("|", "¦").replace("\n", " ")[:300]
            f.write(f"| {i['case']} | {i['test']} | {i['status']} | {det} |\n")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
