#!/usr/bin/env python3
"""WMS_BACKUP — копия рабочей папки WMS на другой носитель (FINAL WMS MARATHON §6: WMS_BACKUP / HEALTHCHECK).

    python3 wms_backup.py --wms WMS_DIR --to DEST_DIR [--keep 8] [--no-backups]

В DEST_DIR создаётся папка WMS_BACKUP_<ГГГГММДД-ЧЧММСС>/: рабочая книга (все .ods папки WMS), журнал WMS_Journal/ (без
файла блокировки) и, если не задано --no-backups, копии WMS_Backups/; MANIFEST.csv — размер и SHA-256 каждого файла.
После записи каждый файл перечитывается и сверяется с манифестом; папка появляется под своим именем только проверенной
(пишется во временную .part). Хранятся --keep последних копий этого вида в DEST_DIR (старые удаляются, другие файлы
не трогаются). Если WMS открыта (есть wms.lock), копия всё равно делается: в ней книга на момент последнего сохранения и
полный журнал — при восстановлении WMS доведёт книгу по журналу («Восстановить»).

Восстановление из такой копии — docs/BACKUP_RECOVERY.md, раздел 5 (книга и WMS_Journal копируются назад вместе).
Код выхода: 0 — копия сделана и проверена, 2 — ошибка (копия не оставлена).
"""
import argparse
import datetime
import hashlib
import os
import shutil
import sys

PREFIX = "WMS_BACKUP_"


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def files_to_copy(wms, with_backups):
    out = [f for f in sorted(os.listdir(wms)) if f.lower().endswith(".ods") and not f.startswith(".~lock") and os.path.isfile(os.path.join(wms, f))]
    for sub in ("WMS_Journal",) + (("WMS_Backups",) if with_backups else ()):
        d = os.path.join(wms, sub)
        if os.path.isdir(d):
            out += [f"{sub}/{f}" for f in sorted(os.listdir(d)) if f != "wms.lock" and os.path.isfile(os.path.join(d, f))]
    return out


def backup(wms, dest, keep=8, with_backups=True):
    wms, dest = os.path.abspath(wms), os.path.abspath(dest)
    if not os.path.isdir(wms):
        raise ValueError(f"папка WMS {wms} не найдена")
    if dest == wms or dest.startswith(wms + os.sep):
        raise ValueError("копию нельзя класть внутрь рабочей папки WMS — выберите другой носитель")
    files = files_to_copy(wms, with_backups)
    if not any(f.lower().endswith(".ods") and "/" not in f for f in files):
        raise ValueError(f"в {wms} нет книги WMS (.ods)")
    os.makedirs(dest, exist_ok=True)
    name = PREFIX + datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    final = os.path.join(dest, name)
    k = 2
    while os.path.exists(final):
        final = os.path.join(dest, f"{name}_{k}")
        k += 1
    part = final + ".part"
    shutil.rmtree(part, ignore_errors=True)
    lines = ["file;size;sha256"]
    try:
        for rel in files:
            src = os.path.join(wms, rel)
            dst = os.path.join(part, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            h0 = sha256(src)
            shutil.copy2(src, dst)
            if sha256(dst) != h0:
                raise OSError(f"{rel}: копия не совпадает с оригиналом (файл менялся во время копирования?)")
            lines.append(f"{rel};{os.path.getsize(dst)};{h0}")
        with open(os.path.join(part, "MANIFEST.csv"), "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines) + "\n")
        os.rename(part, final)
    except Exception:
        shutil.rmtree(part, ignore_errors=True)
        raise
    old = sorted(d for d in os.listdir(dest) if d.startswith(PREFIX) and not d.endswith(".part") and os.path.isdir(os.path.join(dest, d)))
    removed = []
    for d in old[:-keep] if keep > 0 else []:
        shutil.rmtree(os.path.join(dest, d), ignore_errors=True)
        removed.append(d)
    locked = os.path.exists(os.path.join(wms, "WMS_Journal", "wms.lock"))
    return final, len(files), removed, locked


def verify(folder):
    """"" when every file of MANIFEST.csv is present with its SHA-256"""
    for ln in open(os.path.join(folder, "MANIFEST.csv"), encoding="utf-8").read().splitlines()[1:]:
        rel, size, h = ln.rsplit(";", 2)
        p = os.path.join(folder, rel)
        if not os.path.exists(p) or sha256(p) != h:
            return f"{rel}: нет файла или SHA-256 не совпадает"
    return ""


def main(argv=None):
    ap = argparse.ArgumentParser(description="копия рабочей папки WMS на другой носитель с проверкой SHA-256")
    ap.add_argument("--wms", required=True)
    ap.add_argument("--to", required=True)
    ap.add_argument("--keep", type=int, default=8)
    ap.add_argument("--no-backups", action="store_true")
    ap.add_argument("--verify", help="только проверить готовую копию (папку WMS_BACKUP_…)")
    a = ap.parse_args(argv)
    if a.verify:
        why = verify(a.verify)
        print("OK: копия цела" if not why else f"ОШИБКА: {why}")
        return 0 if not why else 2
    try:
        final, n, removed, locked = backup(a.wms, a.to, a.keep, not a.no_backups)
    except (OSError, ValueError) as e:
        print(f"ОШИБКА: {e}", file=sys.stderr)
        return 2
    print(f"OK: {final} — файлов {n}, проверены (SHA-256)" + (f"; удалены старые копии: {', '.join(removed)}" if removed else "") +
          ("; WMS была открыта: в копии книга на момент последнего сохранения и полный журнал" if locked else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
