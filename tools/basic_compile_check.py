"""Compile check for the Basic sources: builds a test book, opens it in a GUI session (Xvfb) and calls one function per
module; a compile error of any module disables the whole library and shows a dialog, whose text is reported.

    python3 tools/basic_compile_check.py [--display :99]
"""
import argparse
import os
import sys
import tempfile
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wmslo import Office  # noqa: E402
from build_ods import build  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    tmp = tempfile.mkdtemp(prefix="wms_cc_")
    book = os.path.join(tmp, "cc.ods")
    ob = Office("wmscc_b%d" % os.getpid(), tmp)
    try:
        build(ob, book, test=True)
    finally:
        ob.terminate()
    o = Office("wmscc_g%d" % os.getpid(), tmp, headless=False, display=a.display)
    ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
    res, dialogs = {}, []

    def call():
        try:
            doc = o.load(book, macros=0)
            res["doc"] = doc
            res["r"] = o.basic(doc, "WmsTestOps", "TestEscRoundTrip", "a;b|c=d~e%f")
        except Exception as e:
            res["e"] = str(e)[:300]

    th = threading.Thread(target=call)
    th.start()
    for _ in range(120):
        time.sleep(0.25)
        if not th.is_alive():
            break
        try:
            w = ut.getTopFocusWindow()
            texts = []
            for cid in w.getChildren():
                try:
                    for pv in w.getChild(cid).getState():
                        if pv.Name == "Text" and pv.Value:
                            texts.append(f"{cid}: {pv.Value}")
                except Exception:
                    pass
            if texts:
                dialogs.append(texts)
                for cid in ("ok", "yes", "close"):
                    if cid in list(w.getChildren()):
                        w.getChild(cid).executeAction("CLICK", ())
                        break
        except Exception:
            pass
    th.join(20)
    o.terminate()
    ok = res.get("r") is True
    print("COMPILE", "OK" if ok else "FAILED", res.get("r"), res.get("e", ""))
    for d in dialogs:
        print("DIALOG:", " | ".join(d)[:600])
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
