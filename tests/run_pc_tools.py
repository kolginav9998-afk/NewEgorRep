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


@case
def p03_offline_audit():
    """tools/offline_audit.py: the sources and a production book of WMS — nothing that reaches the network; a sample planted
    with every kind of way out — each one found (the check is not blind)"""
    c = "P03"
    d = os.path.join(OUT, "cases", "p03")
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(os.path.join(d, "book"))
    rcb, ob = run("build_ods.py", os.path.join(d, "book", "WMS_PROD.ods"))
    reps = []
    for name, path in (("src", os.path.join(ROOT, "src")), ("tools", os.path.join(ROOT, "tools")), ("книга", os.path.join(d, "book"))):
        rc, o = run("offline_audit.py", path)
        reps.append((name, rc, o))
    R.add(c, "исходники WMS (src/, tools/) и рабочая книга без тестового модуля: сетевых возможностей не найдено, код 0",
          rcb == 0 and all(rc == 0 and "СЕТЕВЫХ ВОЗМОЖНОСТЕЙ НЕ НАЙДЕНО" in o and "НАЙДЕНО  " not in o for _, rc, o in reps)
          and "проверено файлов (книги, макросы, программы): 1" in reps[2][2],
          "; ".join(f"{n}: код {rc}, {o.strip().splitlines()[1] if len(o.strip().splitlines()) > 1 else o[:80]}" for n, rc, o in reps) + f"; сборка {rcb} {ob[-80:]}")
    bad = os.path.join(d, "bad")
    os.makedirs(bad)
    # the planted sample is put together from pieces: this file itself must not look like a way out when the audit is run on a release
    # folder (it holds tests/)
    H = "ht" + "tps://"
    open(os.path.join(bad, "leak.bas"), "w", encoding="utf-8").write(
        'Sub Leak\n    c = CreateUnoService("com.sun.star.conn" & "ection.Connector")\n'.replace('" & "', "") +
        '    m = CreateUnoService("com.sun.star.system.SimpleSystem" & "Mail")\n'.replace('" & "', "") +
        '    x = CreateObject("MSXML2.XML' + 'HTTP")\n    u = "' + H + 'example.org/upload"\nEnd Sub\n')
    open(os.path.join(bad, "leak.py"), "w", encoding="utf-8").write(
        "imp" + "ort urllib.request\nimp" + "ort socket\ns = sock" + "et.socket()\ns.conn" + 'ect(("10.0.0.5", 80))\n')
    import zipfile
    with zipfile.ZipFile(os.path.join(bad, "leak.ods"), "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("content.xml", '<office:document-content><table:table-cell><text:p><text:a xlink:href="' + H + 'example.org/x">a</text:a></text:p>'
                   '</table:table-cell><table:table-cell table:formula="of:=WEBSERVICE(&quot;' + H + 'example.org/api&quot;)"/>'
                   '<table:table-source xlink:href="file:///другой.ods"/></office:document-content>')
    rc, o = run("offline_audit.py", bad)
    want = ["отправка почты", "сетевое соединение UNO", "HTTP-запрос через Windows", "адрес в коде — " + H + "example.org/upload", "гиперссылок: 1",
            "внешняя ссылка — " + H + "example.org/x", "формула WEBSERVICE", "связь с внешним файлом", "адрес в содержимом — " + H + "example.org/api",
            "сетевой модуль Python", "сетевое соединение —", "sock" + "et.socket"]
    miss = [w for w in want if w not in o]
    R.add(c, "образец с почтой, соединением UNO, HTTP, адресами, гиперссылкой, WEBSERVICE, внешней связью и сетью Python: каждое найдено, код 1",
          rc == 1 and not miss, f"код {rc}; не найдено: {miss}; {o.strip().splitlines()[-1]}")
    # the folder of a release (it holds this very test): nothing found there either
    rc2, o2 = run("offline_audit.py", os.path.join(ROOT, "tests"))
    R.add(c, "папка tests/ (она попадает в архив выпуска) проверкой локальности чиста: сетевых возможностей не найдено, код 0",
          rc2 == 0 and "СЕТЕВЫХ ВОЗМОЖНОСТЕЙ НЕ НАЙДЕНО" in o2, o2.strip().splitlines()[-1] if o2.strip() else "")


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
