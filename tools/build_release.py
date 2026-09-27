"""The release candidate of WMS (FINAL WMS MARATHON §5): a folder with the production book and what goes with it.

    python3 tools/build_release.py OUT_DIR [--name WMS_PROD_CANDIDATE.ods] [--zip]

OUT_DIR (it must not exist) gets:
  WMS_PROD_CANDIDATE.ods   the production book (tools/build_ods.py without --test: empty registry, no test seams active,
                           MODE PROD) with «Главная» (version, «Состояние системы», «Проверка перед работой») and «Справка»
                           (the guide of the storekeeper, these release notes and the backup / recovery guide);
  WMS_TOOLBOX/             the separate tools (tools/build_toolbox.py): it goes next to the working book;
  docs/                    OPERATOR_GUIDE.md, RELEASE_NOTES.md, BACKUP_RECOVERY.md, EXPORT_CONTRACT.md, TOOLBOX.md,
                           PC_VALIDATION.md;
  tools/                   migrate.py, migration_dryrun.py, wmslo.py — the transfer of the old warehouse (tools/migrate.py),
                           with the independent oracle its VERIFY replays the journal by (tests/*_oracle.py it imports): the
                           folder alone is enough to transfer, nothing of the repository is needed; upgrade.py with build_ods.py
                           — the upgrade of a working book of the previous release (0.6, WMS-SYS-3) to this one in place;
  src/basic/               the Basic modules of this release: what upgrade.py puts into the upgraded book;
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

DOCS = ["OPERATOR_GUIDE.md", "RELEASE_NOTES.md", "BACKUP_RECOVERY.md", "EXPORT_CONTRACT.md", "TOOLBOX.md", "PC_VALIDATION.md"]
TOOLS = ["migrate.py", "migration_dryrun.py", "wmslo.py", "upgrade.py", "build_ods.py"]
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
    v = versions()
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
