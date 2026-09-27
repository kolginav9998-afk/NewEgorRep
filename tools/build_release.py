"""The release candidate of WMS (FINAL WMS MARATHON §5): a folder with the production book and what goes with it.

    python3 tools/build_release.py OUT_DIR [--name WMS_PROD_CANDIDATE.ods]

OUT_DIR (it must not exist) gets:
  WMS_PROD_CANDIDATE.ods   the production book (tools/build_ods.py without --test: empty registry, no test seams active,
                           MODE PROD) with «Главная» (version, «Состояние системы», «Проверка перед работой») and «Справка»
                           (these release notes and the backup / recovery guide);
  docs/                    RELEASE_NOTES.md, BACKUP_RECOVERY.md, EXPORT_CONTRACT.md;
  tools/                   migrate.py, migration_dryrun.py, wmslo.py — the transfer of the old warehouse (tools/migrate.py),
                           with the independent oracle its VERIFY replays the journal by (tests/*_oracle.py it imports): the
                           folder alone is enough to transfer, nothing of the repository is needed;
  RELEASE_MANIFEST.csv     versions and every file with its size and SHA-256.
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

DOCS = ["RELEASE_NOTES.md", "BACKUP_RECOVERY.md", "EXPORT_CONTRACT.md"]
TOOLS = ["migrate.py", "migration_dryrun.py", "wmslo.py"]
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
    out = os.path.abspath(out)
    if os.path.exists(out):
        raise ValueError(f"{out} уже существует — выпуск собирается в новую папку")
    os.makedirs(os.path.join(out, "docs"))
    os.makedirs(os.path.join(out, "tools"))
    prof = tempfile.mkdtemp(prefix="wms_release_")
    o = Office(f"release{os.getpid()}", prof)
    try:
        build_ods.build(o, os.path.join(out, name), test=False)
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    for d in DOCS:
        shutil.copy2(os.path.join(ROOT, "docs", d), os.path.join(out, "docs", d))
    for t in TOOLS:
        shutil.copy2(os.path.join(HERE, t), os.path.join(out, "tools", t))
    for t in ORACLES:
        shutil.copy2(os.path.join(ROOT, "tests", t), os.path.join(out, "tools", t))
    v = versions()
    lines = ["key;value", f"product;{v['WMS_PRODUCT_VERSION']}", f"core;{v['WMS_CORE_VERSION']}", f"schema;{v['WMS_SYS_SCHEMA']}", f"book;{name}"]
    for rel in [name] + [f"docs/{d}" for d in DOCS] + [f"tools/{t}" for t in TOOLS + ORACLES]:
        p = os.path.join(out, rel)
        lines.append(f"file;{rel};{os.path.getsize(p)};{sha256(p)}")
    with open(os.path.join(out, "RELEASE_MANIFEST.csv"), "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")
    return out, v


def main(argv=None):
    ap = argparse.ArgumentParser(description="сборка кандидата в рабочую версию WMS")
    ap.add_argument("out")
    ap.add_argument("--name", default="WMS_PROD_CANDIDATE.ods")
    a = ap.parse_args(argv)
    try:
        out, v = build_release(a.out, a.name)
    except (ValueError, OSError) as e:
        print(f"ОШИБКА: {e}", file=sys.stderr)
        return 2
    print(f"OK: {out}/{a.name} — WMS {v['WMS_PRODUCT_VERSION']} (ядро {v['WMS_CORE_VERSION']}, схема {v['WMS_SYS_SCHEMA']})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
