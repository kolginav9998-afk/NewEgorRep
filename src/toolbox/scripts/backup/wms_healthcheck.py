#!/usr/bin/env python3
"""WMS_HEALTHCHECK — проверка рабочей папки WMS перед сменой (FINAL WMS MARATHON §6: WMS_BACKUP / HEALTHCHECK).

    python3 wms_healthcheck.py --wms WMS_DIR [--book WMS_PROD.ods] [--out REPORT.txt]

Только чтение (пробный файл записи создаётся и сразу удаляется). Проверки:
  книга      — файл есть, читается, это ODS; размер; схема, ядро и последняя операция из служебного листа _SYS (один
               потоковый проход по content.xml: память не растёт с книгой);
  целостность — архив ODS книги и последней резервной копии цел (контрольные суммы всех частей);
  защиты     — листы защищены (кроме «Получатели»), служебные листы скрыты, структура книги защищена;
  журнал     — папка WMS_Journal, каждая строка J1 цела (признак конца и длина), seq непрерывен, один экземпляр WMS;
               журнал не отстаёт от книги (иначе — ошибка), хвост после книги — предупреждение («Восстановить» в WMS);
  блокировка — wms.lock: WMS открыта (кем, с какого времени) или осталась после сбоя;
  копии      — WMS_Backups: число, возраст последней (нет копий — ошибка, старше 2 дней — предупреждение);
  место      — свободное место на диске папки WMS (меньше 0,5 ГБ — ошибка, меньше 2 ГБ — предупреждение);
  права      — запись в папку WMS, WMS_Journal, WMS_Backups (пробный файл);
  версии     — LibreOffice (soffice --version) и его Writer (для актов), Python; проверенные версии LibreOffice — 24.2, 26.2;
  снимки     — последний снимок WMS_Export (возраст).
Итог — строки «OK / ПРЕДУПРЕЖДЕНИЕ / ОШИБКА  что  подробности» и «ИТОГ: ошибок N, предупреждений M».
Код выхода: 0 — без замечаний, 1 — есть предупреждения, 2 — есть ошибки (или папку не прочитать).
"""
import argparse
import datetime
import glob
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
import xml.etree.ElementTree as ET

TESTED_LO = ("24.2", "26.2")
TBL = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFF = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
STY = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"
SYS_KEYS = ("SCHEMA", "INSTANCE_ID", "MODE", "CORE_VERSION", "LAST_SEQ")
# the service sheets of WMS (hidden and protected) and the one sheet kept by hand without protection (WmsStatus.Protections)
SERVICE_SHEETS = ("_ORD", "_RCV", "_SPR", "_ART", "_RET", "_ISS", "_ADJ", "_CAR", "_IDX", "_SYS")
FREE_SHEETS = ("Получатели",)


class Report:
    def __init__(self):
        self.lines, self.err, self.warn = [], 0, 0

    def add(self, level, what, text):
        if level == "ОШИБКА":
            self.err += 1
        elif level == "ПРЕДУПРЕЖДЕНИЕ":
            self.warn += 1
        self.lines.append(f"{level:<15} {what:<12} {text}")

    def text(self):
        return "\n".join(self.lines + [f"ИТОГ: ошибок {self.err}, предупреждений {self.warn}"]) + "\n"


TABLE_TAG = re.compile(rb"<table:table\s([^>]*)>")
SHEETS_TAG = re.compile(rb"<office:spreadsheet(\s[^>]*)?>")
ATTR = re.compile(rb'([\w:.-]+)="([^"]*)"')
CHUNK = 1 << 22


def attrs(raw):
    return {k.decode(): v.decode("utf-8", "replace") for k, v in ATTR.findall(raw)}


def read_book(book):
    """one pass over the bytes of content.xml of the ODS (nothing opened in LibreOffice; a few MB of memory whatever the
    size of the book — a book of 250 000 rows, 0.9 GB of XML, is read in seconds): the values SYS_KEYS of the sheet _SYS
    (only that sheet is parsed as XML), every sheet (name, protected, visible), the protection of the structure"""
    hidden, sheets, structure, sysv = set(), [], None, {}
    with zipfile.ZipFile(book) as z, z.open("content.xml") as f:
        # the head up to the body: the namespaces of the root and the table styles that hide a sheet
        head = b""
        while b"<office:body>" not in head:
            chunk = f.read(CHUNK)
            if not chunk:
                raise ET.ParseError("content.xml без office:body")
            head += chunk
        k = head.index(b"<office:body>")
        buf, head = head[k:], head[:k]
        root = re.search(rb"<office:document-content\s([^>]*)>", head)
        ns = b" ".join(re.findall(rb'xmlns:[\w-]+="[^"]*"', root.group(1))) if root else b""
        for m in re.finditer(rb"<style:style\s([^>]*)>(.*?)</style:style>", head, re.S):
            a = attrs(m.group(1))
            if a.get("style:family") == "table" and b'table:display="false"' in m.group(2):
                hidden.add(a.get("style:name"))
        # the body: the tags of the sheets; the sheet _SYS is kept whole
        capture, eof = None, False
        while True:
            if capture is not None:
                e = buf.find(b"</table:table>")
                if e >= 0:
                    capture += buf[:e + len(b"</table:table>")]
                    buf = buf[e + len(b"</table:table>"):]
                    xml = ET.fromstring(b"<r " + ns + b">" + capture + b"</r>")
                    for row in xml.iter(TBL + "table-row"):
                        cells = row.findall(TBL + "table-cell")
                        if len(cells) >= 2:
                            key = "".join(cells[0].itertext())
                            if key in SYS_KEYS:
                                v = cells[1].get(OFF + "value")
                                sysv[key] = v if v is not None else "".join(cells[1].itertext())
                    capture = None
                    continue
                capture += buf[:-16]
                buf = buf[-16:]
            else:
                if structure is None:
                    m = SHEETS_TAG.search(buf)
                    if m:
                        structure = b'table:structure-protected="true"' in (m.group(1) or b"")
                pos = 0
                for m in TABLE_TAG.finditer(buf):
                    a = attrs(m.group(1))
                    name = a.get("table:name", "")
                    sheets.append((name, a.get("table:protected") == "true", a.get("table:style-name") not in hidden))
                    pos = m.end()
                    if name == "_SYS":
                        capture, buf = m.group(0), buf[m.end():]
                        break
                if capture is not None:
                    continue
                # keep the tail from the last «<» after the tags read: it may be the beginning of a tag cut by the chunk
                cut = buf.rfind(b"<")
                buf = buf[cut:] if cut >= pos else b""
            if eof:
                break
            chunk = f.read(CHUNK)
            if not chunk:
                eof = True
            buf += chunk
    if structure is None:
        structure = False
    return sysv, sheets, structure


def sys_values(book):
    """SCHEMA, INSTANCE_ID, MODE, CORE_VERSION, LAST_SEQ of the sheet _SYS"""
    return read_book(book)[0]


def protections(sheets, structure):
    """what is not protected ([] — everything in place)"""
    p = []
    for name, prot, visible in sheets:
        if not prot and name not in FREE_SHEETS:
            p.append(f"лист «{name}» не защищён")
        if name in SERVICE_SHEETS and visible:
            p.append(f"служебный лист «{name}» не скрыт")
    if structure is not True:
        p.append("структура книги не защищена")
    return p


def zip_damage(path):
    """"" when every part of the ODS (zip) reads with its checksum, otherwise what is damaged"""
    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
        return f"часть {bad} повреждена" if bad else ""
    except Exception as e:                      # BadZipFile, zlib.error, EOFError, OSError: the file does not read whole
        return f"{type(e).__name__}: {e}"


def journal(jdir):
    """(lines, damaged, first damaged, max seq, gap, instances)"""
    n = bad = 0
    first_bad = gap = ""
    prev, mx, inst = None, 0, set()
    for f in sorted(glob.glob(os.path.join(jdir, "WMS_journal_*.csv"))):
        for k, raw in enumerate(open(f, "rb").read().split(b"\n"), start=1):
            if not raw:
                continue
            n += 1
            try:
                ln = raw.decode("utf-8")
            except UnicodeDecodeError:
                ln = ""
            p = ln.rfind(";END;")
            ok = ln.startswith("J1;") and p > 0 and ln[p + 5:].isdigit() and int(ln[p + 5:]) == p
            parts = ln.split(";")
            if not ok or len(parts) < 5 or not parts[1].isdigit():
                bad += 1
                first_bad = first_bad or f"{os.path.basename(f)}, строка {k}"
                continue
            seq = int(parts[1])
            inst.add(parts[3])
            if prev is not None and seq != prev + 1 and not gap:
                gap = f"после seq {prev} идёт {seq}"
            prev, mx = seq, max(mx, seq)
    return n, bad, first_bad, mx, gap, inst


def probe(d):
    try:
        p = os.path.join(d, f".wms_probe_{os.getpid()}")
        with open(p, "w") as f:
            f.write("probe")
        os.remove(p)
        return ""
    except OSError as e:
        return str(e)


def lo_writer():
    """True when the Writer module of LibreOffice is installed (WMS_DOCS writes the acts with it)"""
    for exe in ("soffice", "libreoffice"):
        path = shutil.which(exe)
        if path:
            prog = os.path.dirname(os.path.realpath(path))
            if glob.glob(os.path.join(prog, "libswlo.so")) or glob.glob(os.path.join(prog, "swlo.dll")) or \
                    os.path.exists(os.path.join(os.path.dirname(prog), "share", "registry", "writer.xcd")):
                return True
    return False


def lo_version():
    for exe in ("soffice", "libreoffice"):
        try:
            out = subprocess.run([exe, "--version"], capture_output=True, text=True, timeout=60).stdout
            m = re.search(r"(\d+\.\d+)\.\d+", out)
            if m:
                return m.group(0), m.group(1)
        except (OSError, subprocess.TimeoutExpired):
            continue
    return "", ""


def check(wms, book_name=None):
    r = Report()
    wms = os.path.abspath(wms)
    if not os.path.isdir(wms):
        r.add("ОШИБКА", "папка", f"{wms} не найдена")
        return r
    books = [book_name] if book_name else sorted(f for f in os.listdir(wms) if f.lower().endswith(".ods") and not f.startswith(".~lock"))
    book = os.path.join(wms, books[0]) if books else ""
    sysv = {}
    if not book or not os.path.exists(book):
        r.add("ОШИБКА", "книга", f"в {wms} нет книги WMS (.ods)")
    else:
        try:
            sysv, sheets, structure = read_book(book)
            if not sysv:
                r.add("ОШИБКА", "книга", f"{os.path.basename(book)}: нет служебного листа _SYS — это не книга WMS")
            else:
                r.add("OK", "книга", f"{os.path.basename(book)}, {os.path.getsize(book) / 1e6:.1f} МБ; схема {sysv.get('SCHEMA')}, режим {sysv.get('MODE')}, "
                      f"последняя операция {sysv.get('LAST_SEQ')}")
                p = protections(sheets, structure)
                r.add("ПРЕДУПРЕЖДЕНИЕ" if p else "OK", "защиты", "; ".join(p) + " — сообщите ответственному за WMS: защита не даёт изменить учёт "
                      "мимо WMS" if p else "листы защищены, служебные листы скрыты, структура книги защищена")
        except (zipfile.BadZipFile, ET.ParseError, KeyError, OSError, EOFError) as e:
            r.add("ОШИБКА", "книга", f"{os.path.basename(book)} не читается как ODS: {e}")
        why = zip_damage(book)
        r.add("ОШИБКА" if why else "OK", "целостность", f"{os.path.basename(book)}: " + (why + " — восстановите книгу из резервной копии и журнала "
              "(BACKUP_RECOVERY.md)" if why else "все части архива ODS читаются, контрольные суммы совпадают"))
    jdir = os.path.join(wms, "WMS_Journal")
    if not os.path.isdir(jdir):
        r.add("ПРЕДУПРЕЖДЕНИЕ" if sysv.get("LAST_SEQ") in (None, "0", "") else "ОШИБКА", "журнал", "папки WMS_Journal нет" +
              ("" if sysv.get("LAST_SEQ") in (None, "0", "") else " — а в книге есть операции: верните папку журнала"))
    else:
        n, bad, first_bad, mx, gap, inst = journal(jdir)
        last = int(float(sysv.get("LAST_SEQ") or 0))
        if bad:
            r.add("ОШИБКА", "журнал", f"повреждённых строк {bad} (первая — {first_bad}); журнал не правьте, откройте WMS — запуск покажет, что делать")
        elif gap:
            r.add("ОШИБКА", "журнал", f"пропуск номера операции: {gap}")
        elif len(inst) > 1 or (sysv.get("INSTANCE_ID") and inst and sysv.get("INSTANCE_ID") not in inst):
            r.add("ОШИБКА", "журнал", f"в папке журнала строки другой WMS (экземпляры {', '.join(sorted(inst))})")
        elif mx < last:
            r.add("ОШИБКА", "журнал", f"журнал отстаёт от книги: в журнале до seq {mx}, в книге {last} — не работайте до разбора (нужна папка журнала)")
        elif mx > last:
            r.add("ПРЕДУПРЕЖДЕНИЕ", "журнал", f"в журнале операции после сохранённой книги (seq {last + 1}…{mx}): при открытии WMS нажмите «Восстановить»")
        else:
            r.add("OK", "журнал", f"строк {n}, seq до {mx}, совпадает с книгой")
        lock = os.path.join(jdir, "wms.lock")
        if os.path.exists(lock):
            age = (time.time() - os.path.getmtime(lock)) / 3600
            who = open(lock, encoding="utf-8", errors="replace").read().strip()[:120]
            r.add("ПРЕДУПРЕЖДЕНИЕ" if age > 12 else "OK", "блокировка", f"WMS открыта или закрыта аварийно ({who}; {age:.1f} ч назад)" +
                  (" — если WMS нигде не открыта: «Снять блокировку WMS»" if age > 12 else ""))
        else:
            r.add("OK", "блокировка", "WMS сейчас не открыта")
    bdir = os.path.join(wms, "WMS_Backups")
    backs = sorted(glob.glob(os.path.join(bdir, "*.ods")), key=os.path.getmtime) if os.path.isdir(bdir) else []
    if not backs:
        r.add("ОШИБКА", "копии", "в WMS_Backups нет резервных копий — в WMS: «Главная» → «Резервная копия»")
    else:
        age = (time.time() - os.path.getmtime(backs[-1])) / 86400
        r.add("ПРЕДУПРЕЖДЕНИЕ" if age > 2 else "OK", "копии", f"{len(backs)} шт., последняя {os.path.basename(backs[-1])} ({age:.1f} дн. назад)" +
              (" — сделайте копию и перенесите на другой носитель" if age > 2 else ""))
        why = zip_damage(backs[-1])
        r.add("ОШИБКА" if why else "OK", "целостность", f"последняя копия {os.path.basename(backs[-1])}: " + (why + " — сделайте новую копию («Резервная "
              "копия»), проверьте носитель" if why else "архив ODS цел"))
    free = shutil.disk_usage(wms).free / 2 ** 30
    r.add("ОШИБКА" if free < 0.5 else ("ПРЕДУПРЕЖДЕНИЕ" if free < 2 else "OK"), "место", f"свободно {free:.1f} ГБ на диске папки WMS")
    for d in (wms, jdir, bdir):
        if os.path.isdir(d):
            why = probe(d)
            if why:
                r.add("ОШИБКА", "права", f"нет записи в {d}: {why}")
    if not any(ln.startswith("ОШИБКА") and "права" in ln for ln in r.lines):
        r.add("OK", "права", "запись в папки WMS разрешена")
    full, mm = lo_version()
    if not full:
        r.add("ПРЕДУПРЕЖДЕНИЕ", "версии", f"LibreOffice не найден в PATH; Python {sys.version.split()[0]}")
    else:
        r.add("OK" if mm in TESTED_LO else "ПРЕДУПРЕЖДЕНИЕ", "версии", f"LibreOffice {full}" + ("" if mm in TESTED_LO else
              f" — с WMS не проверялась (проверены {', '.join(TESTED_LO)})") + f"; Python {sys.version.split()[0]}")
        if not lo_writer():
            r.add("ПРЕДУПРЕЖДЕНИЕ", "версии", "LibreOffice Writer не установлен — акты WMS_DOCS не создаются (пакет libreoffice-writer)")
    exp = os.path.join(wms, "WMS_Export")
    snaps = sorted(glob.glob(os.path.join(exp, "WMS_SNAPSHOT_*"))) if os.path.isdir(exp) else []
    if snaps:
        age = (time.time() - os.path.getmtime(snaps[-1])) / 86400
        r.add("OK", "снимки", f"последний {os.path.basename(snaps[-1])} ({age:.1f} дн. назад)")
    return r


def main(argv=None):
    ap = argparse.ArgumentParser(description="проверка рабочей папки WMS перед сменой (только чтение)")
    ap.add_argument("--wms", required=True)
    ap.add_argument("--book")
    ap.add_argument("--out")
    a = ap.parse_args(argv)
    r = check(a.wms, a.book)
    head = f"ПРОВЕРКА WMS {datetime.datetime.now():%d.%m.%Y %H:%M} — {os.path.abspath(a.wms)}\n"
    txt = head + r.text()
    if a.out:
        open(a.out, "w", encoding="utf-8").write(txt)
    print(txt, end="")
    return 2 if r.err else (1 if r.warn else 0)


if __name__ == "__main__":
    sys.exit(main())
