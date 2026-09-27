#!/usr/bin/env python3
"""Проверка принтера этикеток на складском ПК (FINAL WMS MARATHON, M5; docs/PC_VALIDATION.md). Печатает только тестовые
этикетки — синтетические ЕИ фикстуры снимка (tests/fixtures); рабочую WMS не открывает.

    python3 tools/printer_check.py                        что видит система: принтеры CUPS, принтер по умолчанию, /dev/usb/lp*
    python3 tools/printer_check.py --pdf-only             только PDF тестовых этикеток (без печати) — посмотреть макет
    python3 tools/printer_check.py --cups ИМЯ [--media M]  тестовые этикетки WMS_LABELS через CUPS: PDF → lp -d ИМЯ
    python3 tools/printer_check.py --tspl /dev/usb/lp0    тестовая этикетка командами TSPL (принтер без драйвера CUPS)
      общие ключи: [--size 58x40] [--gap 2] [--dpmm 8] [--font TSS24.BF2] [--answer да|нет]

Путь PDF — тот же, что у кнопок WMS_LABELS: набор инструментов собирается во временную папку, «Загрузить снимок»,
«Построить этикетки» (ЕИ-00000201, ЕИ-00000206, ЕИ-00000208), «В PDF». После печати скрипт спрашивает, верно ли
напечатано (ЕИ крупно, наименование, артикул, место, штрихкод читается сканером), и записывает ответ в
pc_check_results/printer_check.txt и в последний PC_CHECK_REPORT.md. Код выхода: 0 — отправлено (и ответ «да»),
1 — ответ «нет», 2 — печать не удалась.
"""
import argparse
import datetime
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FIXTURE = os.path.join(ROOT, "tests", "fixtures", "WMS_SNAPSHOT_FIXTURE")
TEST_EIS = ("201", "206", "208")
RESULTS = os.environ.get("WMS_PC_RESULTS") or os.path.join(ROOT, "pc_check_results")     # the tests set their own folder


def run(cmd, timeout=120):
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return r.returncode, (r.stdout or "") + (r.stderr or "")
    except FileNotFoundError:
        return 127, ""


def listing():
    lines = []
    if not shutil.which("lpstat"):
        lines.append("CUPS не установлен: sudo apt install cups (печать через системный принтер)")
    else:
        rc, out = run(["lpstat", "-p", "-d", "-v"])
        lines.append("Принтеры CUPS (lpstat -p -d -v):")
        lines += ["  " + ln for ln in out.strip().splitlines()] or ["  нет принтеров — добавьте: Параметры → Принтеры или http://localhost:631"]
    devs = sorted(glob.glob("/dev/usb/lp*"))
    lines.append("USB-устройства печати: " + (", ".join(f"{d} ({'запись разрешена' if os.access(d, os.W_OK) else 'нет прав: sudo usermod -aG lp $USER'})"
                                                         for d in devs) or "нет (/dev/usb/lp*)"))
    lines += ["", "Дальше:", "  принтер с драйвером CUPS:  python3 tools/printer_check.py --cups ИМЯ_ПРИНТЕРА",
              "  принтер без драйвера (TSPL): python3 tools/printer_check.py --tspl /dev/usb/lp0",
              "  только посмотреть макет:     python3 tools/printer_check.py --pdf-only"]
    return "\n".join(lines)


def label_pdf(work, w, h):
    """the PDF of the test labels made by WMS_LABELS itself (its buttons' functions) on the fixture snapshot"""
    sys.path.insert(0, HERE)
    from wmslo import Office
    import build_toolbox
    base = os.path.join(work, "labels")
    os.makedirs(os.path.join(base, "WMS_Export"))
    shutil.copytree(FIXTURE, os.path.join(base, "WMS_Export", "WMS_SNAPSHOT_FIXTURE"))
    o = Office(f"prncheck{os.getpid()}", work)
    try:
        build_toolbox.build(o, base)
        doc = o.load(os.path.join(base, "WMS_TOOLBOX", "WMS_LABELS.ods"), macros=4)
        o.basic(doc, "TbCommon", "TbTestAuto", True)
        st = doc.Sheets.getByName("Настройки")
        st.getCellByPosition(1, 5).setString(str(w).replace(".", ","))
        st.getCellByPosition(1, 6).setString(str(h).replace(".", ","))
        o.basic(doc, "TbLabels", "BtnLblRefresh")
        doc.Sheets.getByName("Этикетки").getCellRangeByPosition(0, 1, 1, len(TEST_EIS)).setDataArray(tuple((e, 1) for e in TEST_EIS))
        built = o.basic(doc, "TbLabels", "LblBuild")
        pdf = o.basic(doc, "TbLabels", "LblPdf")
        doc.setModified(False)
        doc.close(True)
    finally:
        o.terminate()
    if not str(built).startswith("OK") or not str(pdf).startswith("OK:"):
        raise RuntimeError(f"этикетки не построены: {built}; {pdf}")
    return pdf[3:]


def ask(answer):
    if answer:
        return answer.strip().lower(), ""
    if not sys.stdin.isatty():
        return "", "ответа нет (запуск без терминала)"
    a = input("\nЭтикетка напечатана верно? ЕИ крупно, наименование, артикул, место, штрихкод читается сканером (да/нет): ").strip().lower()
    why = "" if a.startswith("д") else input("Что не так (коротко): ").strip()
    return a, why


def record(text):
    os.makedirs(RESULTS, exist_ok=True)
    with open(os.path.join(RESULTS, "printer_check.txt"), "a", encoding="utf-8") as f:
        f.write(text + "\n")
    reps = sorted(glob.glob(os.path.join(RESULTS, "*", "PC_CHECK_REPORT.md")), key=os.path.getmtime)
    if reps:
        with open(reps[-1], "a", encoding="utf-8") as f:
            f.write("\n## Принтер\n\n" + "\n".join("- " + ln for ln in text.splitlines()) + "\n")
    return os.path.join(RESULTS, "printer_check.txt")


def main(argv=None):
    ap = argparse.ArgumentParser(description="проверка принтера этикеток (печатает только тестовые этикетки)")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--cups", metavar="ПРИНТЕР")
    g.add_argument("--tspl", metavar="УСТРОЙСТВО_ИЛИ_ФАЙЛ")
    g.add_argument("--pdf-only", action="store_true")
    ap.add_argument("--media", help="размер носителя для lp (-o media=…), например w164h113 или Custom.58x40mm")
    ap.add_argument("--size", default="58x40", help="ширина x высота этикетки, мм")
    ap.add_argument("--gap", type=float, default=2)
    ap.add_argument("--dpmm", type=int, default=8, help="точек на мм: 203 dpi — 8, 300 dpi — 12")
    ap.add_argument("--font", default="TSS24.BF2")
    ap.add_argument("--answer", help="ответ без вопроса (да/нет)")
    a = ap.parse_args(argv)
    if not (a.cups or a.tspl or a.pdf_only):
        print(listing())
        return 0
    m = re.fullmatch(r"(\d+(?:[.,]\d+)?)x(\d+(?:[.,]\d+)?)", a.size)
    if not m:
        print("ОШИБКА: --size ШИРИНАxВЫСОТА в мм, например 58x40", file=sys.stderr)
        return 2
    w, h = float(m.group(1).replace(",", ".")), float(m.group(2).replace(",", "."))
    stamp = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    try:
        if a.tspl:
            dev = ["--device", a.tspl] if a.tspl.startswith("/dev/") else ["--out", a.tspl]
            rc, out = run([sys.executable, os.path.join(ROOT, "src", "toolbox", "scripts", "labels", "tspl_labels.py"), "--snapshot", FIXTURE,
                           "--ei", ",".join(TEST_EIS[:2]), "--width", str(w), "--height", str(h), "--gap", str(a.gap), "--dpmm", str(a.dpmm),
                           "--font", a.font] + dev)
            if rc != 0:
                raise RuntimeError(out.strip())
            sent = f"TSPL → {a.tspl}: {out.strip()} (этикетка {w:g}×{h:g} мм, зазор {a.gap:g} мм, {a.dpmm} точек/мм, шрифт {a.font})"
        else:
            work = tempfile.mkdtemp(prefix="wms_printer_")
            pdf = label_pdf(work, w, h)
            keep = os.path.join(RESULTS, f"test_labels_{datetime.datetime.now():%Y%m%d-%H%M%S}.pdf")
            os.makedirs(RESULTS, exist_ok=True)
            shutil.copy2(pdf, keep)
            shutil.rmtree(work, ignore_errors=True)
            if a.pdf_only:
                print(f"OK: PDF тестовых этикеток ({w:g}×{h:g} мм, {len(TEST_EIS)} шт.): {keep}")
                record(f"{stamp}: PDF тестовых этикеток {w:g}×{h:g} мм — {keep} (без печати)")
                return 0
            if not shutil.which("lp"):
                raise RuntimeError("нет команды lp — установите CUPS: sudo apt install cups")
            cmd = ["lp", "-d", a.cups] + (["-o", f"media={a.media}"] if a.media else []) + [keep]
            rc, out = run(cmd)
            if rc != 0:
                raise RuntimeError(f"{' '.join(cmd)}: {out.strip()}")
            sent = f"CUPS → {a.cups}: {out.strip()} ({len(TEST_EIS)} этикетки {w:g}×{h:g} мм; PDF {keep})"
    except Exception as e:
        msg = f"{stamp}: ПЕЧАТЬ НЕ УДАЛАСЬ — {e}"
        print("ОШИБКА: " + str(e), file=sys.stderr)
        record(msg)
        return 2
    print("отправлено: " + sent)
    ans, why = ask(a.answer)
    verdict = "напечатано верно" if ans.startswith("д") else ("НЕ ВЕРНО: " + (why or ans) if ans else why)
    path = record(f"{stamp}: {sent}; ответ: {verdict}")
    print(f"записано: {path}")
    return 0 if ans.startswith("д") or not ans else 1


if __name__ == "__main__":
    sys.exit(main())
