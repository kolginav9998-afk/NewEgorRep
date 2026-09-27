"""Compile check of the toolbox books: every module of every book is called once in a GUI session (Xvfb); a compile
error of a module shows a dialog, whose text is reported (like tools/basic_compile_check.py for the WMS book).

    python3 tools/tb_compile_check.py TOOLBOX_DIR [--display :99]
"""
import argparse
import glob
import os
import sys
import tempfile
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wmslo import Office  # noqa: E402

# a function of every module that runs without a snapshot and without a window
PROBES = {"TbCommon": ("TbInStrRev", ("a/b", "/")), "TbInventory": ("NumText", (1.5,)), "TbAnalytics": ("AnProbe", ()),
          "TbManager": ("MgrProbe", ()), "TbSearch": ("SearchProbe", ()), "TbDoctor": ("DoctorProbe", ()), "TbLabels": ("Code128C", ("00000201",)),
          "TbDocs": ("DocsProbe", ()), "TbArchive": ("FieldOf", ("a;b;c", 1)), "TbImporter": ("DateText", ("2026-09-01", "")),
          "TbLauncher": ("ToolList", ()), "LtMain": ("LtInStrRev", ("a/b", "/"))}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("toolbox")
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    tmp = tempfile.mkdtemp(prefix="wms_tbcc_")
    o = Office("wmstbcc%d" % os.getpid(), tmp, headless=False, display=a.display)
    ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
    bad = []
    try:
        for book in sorted(glob.glob(os.path.join(a.toolbox, "*.ods"))):
            doc = o.load(book, macros=4)
            mods = list(doc.BasicLibraries.getByName("Standard").getElementNames())
            for m in mods:
                fn, args = PROBES.get(m, (None, ()))
                if fn is None:
                    bad.append(f"{os.path.basename(book)}: нет пробы для модуля {m}")
                    continue
                res, dialogs = {}, []

                def call():
                    try:
                        res["r"] = o.basic(doc, m, fn, *args)
                    except Exception as e:
                        res["e"] = str(e)[:200]
                th = threading.Thread(target=call)
                th.start()
                for _ in range(60):
                    time.sleep(0.2)
                    if not th.is_alive():
                        break
                    try:
                        w = ut.getTopFocusWindow()
                        texts = []
                        for cid in w.getChildren():
                            try:
                                for pv in w.getChild(cid).getState():
                                    if pv.Name == "Text" and pv.Value:
                                        texts.append(str(pv.Value))
                            except Exception:
                                pass
                        if texts:
                            dialogs.append(" | ".join(texts))
                            for cid in ("ok", "yes", "close"):
                                if cid in list(w.getChildren()):
                                    w.getChild(cid).executeAction("CLICK", ())
                                    break
                    except Exception:
                        pass
                th.join(20)
                if dialogs or "e" in res or res.get("r") in (None, ""):
                    bad.append(f"{os.path.basename(book)} {m}.{fn}: {res} {dialogs[:2]}")
            doc.setModified(False)
            doc.close(True)
    finally:
        o.terminate()
    print("COMPILE", "OK" if not bad else "FAILED")
    for b in bad:
        print(b[:700])
    sys.exit(0 if not bad else 1)


if __name__ == "__main__":
    main()
