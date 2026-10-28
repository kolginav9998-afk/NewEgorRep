"""Public Const → Private Const for the constants of src/basic that no other module, tool or test uses.

LibreOffice Basic pays for every Public Const of the libraries at EVERY macro call (measured: 1500 unused constants — +50 ms per call, ≈33 µs each; Private Const cost nothing), so a
constant that is only used inside its own module is declared Private. Run:  python3 tools/private_consts.py [--apply]   (without --apply — only the list). A name that another module,
a tool, a test or another library of src/ mentions (as a word, any case) stays Public — nothing is guessed."""
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DECL = re.compile(r'^(\s*)(Public\s+)?Const\s+(\w+)\s*=', re.M | re.I)


def analyse():
    mods = {}
    for f in sorted(glob.glob(os.path.join(ROOT, "src", "basic", "*.bas"))):
        mods[os.path.basename(f)[:-4]] = open(f, encoding="utf-8").read()
    ext = ""
    for pat in ("tools/*.py", "tests/*.py", "src/toolbox/*.bas", "src/transfer/*.bas", "src/upgrade/*.bas"):
        for f in glob.glob(os.path.join(ROOT, pat)):
            if os.path.basename(f) == "private_consts.py":
                continue
            ext += open(f, encoding="utf-8", errors="replace").read() + "\n"
    out = []
    for m, t in mods.items():
        for x in DECL.finditer(t):
            if not x.group(2):
                continue
            name = x.group(3)
            pat = re.compile(r"\b" + re.escape(name) + r"\b", re.I)
            if any(mm != m and pat.search(tt) for mm, tt in mods.items()) or pat.search(ext):
                continue
            out.append((m, name))
    return mods, out


def main():
    mods, out = analyse()
    print(f"констант Public → Private: {len(out)}")
    if "--apply" not in sys.argv:
        return
    by = {}
    for m, n in out:
        by.setdefault(m, []).append(n)
    for m, names in by.items():
        p = os.path.join(ROOT, "src", "basic", m + ".bas")
        t = mods[m]
        for n in names:
            t, k = re.subn(r"^(\s*)Public(\s+)Const(\s+)" + re.escape(n) + r"\b", r"\1Private\2Const\3" + n, t, flags=re.M | re.I)
            if k != 1:
                raise SystemExit(f"{m}.{n}: объявлений {k}, ожидалось 1")
        open(p, "w", encoding="utf-8", newline="").write(t)
    print("готово")


if __name__ == "__main__":
    main()
