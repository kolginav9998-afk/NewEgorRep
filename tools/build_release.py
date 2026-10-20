"""The release candidate of WMS (FINAL WMS MARATHON §5): a folder with the production book and what goes with it.

    python3 tools/build_release.py OUT_DIR [--name WMS_PROD_CANDIDATE.ods] [--zip]

OUT_DIR (it must not exist) gets:
  WMS_PROD_CANDIDATE.ods   the production book (tools/build_ods.py without --test: empty registry, no test seams active,
                           MODE PROD) with «Главная» (version, «Состояние системы», «Проверка перед работой») and «Справка»
                           (the guide of the storekeeper, these release notes and the backup / recovery guide);
  WMS_TOOLBOX/             the separate tools (tools/build_toolbox.py): it goes next to the working book;
  WMS_LEGACY_TRANSFER.ods  the transfer of the old table «Заказы» into this WMS (tools/build_transfer.py; M7): its buttons
                           run tools/legacy_transfer.py of this folder on a copy of WMS_PROD_CANDIDATE.ods;
  WMS_UPGRADE.ods          the wizard «Обновление из текущей WMS» (tools/build_upgrade.py; PRIME 0.8): its buttons run
                           tools/upgrade.py of this folder — the working book of 0.7 is only read, the new one is built in a
                           new folder (by default WMS_WORK of this folder) and becomes working by its last button;
  docs/                    OPERATOR_GUIDE.md, RELEASE_NOTES.md, BACKUP_RECOVERY.md, EXPORT_CONTRACT.md, TOOLBOX.md,
                           PC_VALIDATION.md, LEGACY_TRANSFER.md;
  tools/                   migrate.py, migration_dryrun.py, wmslo.py — the transfer of the old warehouse (tools/migrate.py),
                           with the independent oracle its VERIFY replays the journal by (tests/*_oracle.py it imports): the
                           folder alone is enough to transfer, nothing of the repository is needed; upgrade.py with build_ods.py
                           — the upgrade of a working book of the previous release (0.6, WMS-SYS-3) to this one in place;
                           legacy_transfer.py — the engine of WMS_LEGACY_TRANSFER.ods (it imports migrate.py and the oracles);
  src/basic/               the Basic modules of this release: what upgrade.py puts into the upgraded book;
  ПРОЧТИ СНАЧАЛА.txt       1.0: which book to open for what (the upgrade, the transfer, a new warehouse), the check of the PC;
  src/ tools/ tests/       1.0: every source tools/pc_check.py builds and tests from — the PC is checked from the archive;
  RELEASE_MANIFEST.csv     versions and every file of the folder with its size and SHA-256.
--zip: also WMS_RELEASE_<product version>.zip next to OUT_DIR — the folder packed under the name WMS_RELEASE_<version>/
(an existing archive of another version is never touched; the same version is refused).
The book is built, never opened with WMS: its first start registers it where it is put (a new WMS). Binary books are not
kept in Git (D-009) — a release is always built from the sources.
"""
import argparse
import hashlib
import os
import re
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

DOCS = ["OPERATOR_GUIDE.md", "RELEASE_NOTES.md", "BACKUP_RECOVERY.md", "EXPORT_CONTRACT.md", "TOOLBOX.md", "PC_VALIDATION.md",
        "LEGACY_TRANSFER.md"]
TOOLS = ["migrate.py", "migration_dryrun.py", "wmslo.py", "upgrade.py", "build_ods.py", "legacy_transfer.py"]
# the oracle of `migrate.py verify` (tests/adjust_oracle.py) and the modules it imports
ORACLES = ["adjust_oracle.py", "journal_oracle.py", "receipt_oracle.py", "special_oracle.py", "return_oracle.py"]


def versions():
    """WMS_PRODUCT_VERSION, WMS_CORE_VERSION, WMS_SYS_SCHEMA of the sources (src/basic/WmsConfig.bas)"""
    txt = open(os.path.join(ROOT, "src", "basic", "WmsConfig.bas"), encoding="utf-8").read()
    out = {}
    for k in ("WMS_PRODUCT_VERSION", "WMS_CORE_VERSION", "WMS_SYS_SCHEMA"):
        m = re.search(rf'Public Const {k} = "([^"]+)"', txt)
        out[k] = m.group(1) if m else ""
    return out


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def build_release(out, name="WMS_PROD_CANDIDATE.ods"):
    from wmslo import Office
    import build_ods
    import build_toolbox
    import build_transfer
    import build_upgrade
    out = os.path.abspath(out)
    if os.path.exists(out):
        raise ValueError(f"{out} уже существует — выпуск собирается в новую папку")
    os.makedirs(os.path.join(out, "docs"))
    os.makedirs(os.path.join(out, "tools"))
    prof = tempfile.mkdtemp(prefix="wms_release_")
    o = Office(f"release{os.getpid()}", prof)
    try:
        build_ods.build(o, os.path.join(out, name), test=False)
        build_toolbox.build(o, out)
        build_transfer.build(o, out)
        build_upgrade.build(o, out)
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    for d in DOCS:
        shutil.copy2(os.path.join(ROOT, "docs", d), os.path.join(out, "docs", d))
    for t in TOOLS:
        shutil.copy2(os.path.join(HERE, t), os.path.join(out, "tools", t))
    for t in ORACLES:
        shutil.copy2(os.path.join(ROOT, "tests", t), os.path.join(out, "tools", t))
    # the modules the upgrade puts into a book of the previous release (tools/build_ods.py reads them from ../src/basic)
    os.makedirs(os.path.join(out, "src", "basic"))
    for f in sorted(os.listdir(os.path.join(ROOT, "src", "basic"))):
        if f.endswith(".bas"):
            shutil.copy2(os.path.join(ROOT, "src", "basic", f), os.path.join(out, "src", "basic", f))
    # 1.0: the check of the warehouse PC runs from the archive itself (no git, no download of the sources): every source it
    # builds and tests from — src/, tools/, tests/ (with the fixtures), docs/
    copy_sources_for_pc_check(out)
    v = versions()
    write_start_here(out, name, v)
    lines = ["key;value", f"product;{v['WMS_PRODUCT_VERSION']}", f"core;{v['WMS_CORE_VERSION']}", f"schema;{v['WMS_SYS_SCHEMA']}", f"book;{name}"]
    files = []
    for base, dirs, names in os.walk(out):
        dirs[:] = sorted(d for d in dirs if d != "__pycache__")
        files += [os.path.relpath(os.path.join(base, n), out).replace(os.sep, "/") for n in sorted(names)]
    for rel in sorted(files):
        p = os.path.join(out, rel)
        lines.append(f"file;{rel};{os.path.getsize(p)};{sha256(p)}")
    with open(os.path.join(out, "RELEASE_MANIFEST.csv"), "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")
    return out, v


def copy_sources_for_pc_check(out):
    """the sources tools/pc_check.py needs next to the built books (files already there are kept)"""
    def walk(rel, keep):
        for base, dirs, names in os.walk(os.path.join(ROOT, rel)):
            dirs[:] = sorted(d for d in dirs if d != "__pycache__")
            for n in sorted(names):
                if not keep(n):
                    continue
                src = os.path.join(base, n)
                dst = os.path.join(out, os.path.relpath(src, ROOT))
                if not os.path.exists(dst):
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.copy2(src, dst)
    walk("src", lambda n: n.endswith((".bas", ".py")))
    walk("tools", lambda n: n.endswith(".py"))
    walk("tests", lambda n: not n.endswith((".pyc", ".ods")) or "fixtures" in n)
    walk("docs", lambda n: n.endswith(".md"))


START_HERE = "ПРОЧТИ СНАЧАЛА.txt"


def write_start_here(out, name, v):
    """1.0: the first file of the archive — which book to open for what (the empty candidate book was opened instead of the
    upgrade and showed an empty warehouse)"""
    t = f"""WMS {v['WMS_PRODUCT_VERSION']} — С ЧЕГО НАЧАТЬ

1. У ВАС УЖЕ ЕСТЬ РАБОЧАЯ WMS С ДАННЫМИ (0.7.0 … 0.7.2)
   - Закройте рабочую WMS.
   - Откройте WMS_UPGRADE.ods из этой папки с включёнными макросами — мастер «Обновление из текущей WMS».
   - «Выбрать…» — укажите рабочую книгу, затем по порядку: «1. ПРОВЕРИТЬ» → (по желанию «2. РЕПЕТИЦИЯ») →
     «3. СОБРАТЬ И СВЕРИТЬ» → в конце рабочего дня «4. СДЕЛАТЬ РАБОЧЕЙ».
   - Старая книга не меняется: новая со всеми данными и журналом появляется в новой папке, старая запирается
     («ЭТА КНИГА БОЛЬШЕ НЕ РАБОЧАЯ»). Подробно — docs/OPERATOR_GUIDE.md, раздел 11.
   - Книгу промежуточного выпуска 0.8 мастер не переносит (скажет об этом и ничего не изменит) — обновляйте книгу 0.7.2.
   - После перехода: «Обслуживание» → «Копия на флешку…», «Зеркало журнала…»; раз в месяц — «Учебное восстановление».

2. У ВАС СТАРАЯ ТАБЛИЦА «ЗАКАЗЫ» (НЕ WMS)
   - WMS_LEGACY_TRANSFER.ods — перенос таблицы в WMS (docs/LEGACY_TRANSFER.md).

3. НОВЫЙ СКЛАД С НУЛЯ
   - {name} — ПУСТАЯ книга WMS. Скопируйте её в рабочую папку и откройте.
   - Для склада, где уже есть данные, её НЕ открывайте: в ней пусто (данные — через пункт 1 или 2).

ПЕРЕД ПЕРВЫМ ДНЁМ
   - Проверка складского ПК: в этой папке выполните  python3 tools/pc_check.py  (docs/PC_VALIDATION.md).
     Интернет и git для неё не нужны — исходники лежат здесь же.
   - Макросы: Сервис → Параметры → LibreOffice → Безопасность → «Безопасность макросов…» → «Доверенные источники» →
     «Добавить» — папка рабочей WMS. Без этого WMS открывается с надписью «WMS не запущена — макросы выключены».
   - Инструменты (WMS_TOOLBOX) — рядом с рабочей книгой; обновлению их данные переносятся сами.

ДОКУМЕНТЫ: docs/OPERATOR_GUIDE.md — руководство кладовщика; docs/RELEASE_NOTES.md — что нового;
docs/BACKUP_RECOVERY.md — резервные копии и восстановление; RELEASE_MANIFEST.csv — файлы выпуска с SHA-256.
"""
    with open(os.path.join(out, START_HERE), "w", encoding="utf-8", newline="\r\n") as f:
        f.write(t)


def pack(out, version):
    """WMS_RELEASE_<version>.zip next to the folder: every file under WMS_RELEASE_<version>/ (the archive of another
    version stays as it is); the archive is written to a temporary name and then renamed"""
    import zipfile
    top = f"WMS_RELEASE_{version}"
    path = os.path.join(os.path.dirname(out), top + ".zip")
    if os.path.exists(path):
        raise ValueError(f"{path} уже существует — архив этой версии не перезаписывается")
    tmp = path + ".part"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for base, dirs, names in os.walk(out):
            dirs[:] = sorted(d for d in dirs if d != "__pycache__")
            for n in sorted(names):
                p = os.path.join(base, n)
                z.write(p, top + "/" + os.path.relpath(p, out).replace(os.sep, "/"))
    os.replace(tmp, path)
    return path


def main(argv=None):
    ap = argparse.ArgumentParser(description="сборка кандидата в рабочую версию WMS")
    ap.add_argument("out")
    ap.add_argument("--name", default="WMS_PROD_CANDIDATE.ods")
    ap.add_argument("--zip", action="store_true", help="также архив WMS_RELEASE_<версия>.zip рядом с папкой")
    a = ap.parse_args(argv)
    try:
        if a.zip:
            z = os.path.join(os.path.dirname(os.path.abspath(a.out)), f"WMS_RELEASE_{versions()['WMS_PRODUCT_VERSION']}.zip")
            if os.path.exists(z):
                raise ValueError(f"{z} уже существует — архив этой версии не перезаписывается")
        out, v = build_release(a.out, a.name)
        z = pack(out, v["WMS_PRODUCT_VERSION"]) if a.zip else ""
    except (ValueError, OSError) as e:
        print(f"ОШИБКА: {e}", file=sys.stderr)
        return 2
    print(f"OK: {out}/{a.name} — WMS {v['WMS_PRODUCT_VERSION']} (ядро {v['WMS_CORE_VERSION']}, схема {v['WMS_SYS_SCHEMA']})")
    if z:
        print(f"OK: архив {z} ({os.path.getsize(z)} байт)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
