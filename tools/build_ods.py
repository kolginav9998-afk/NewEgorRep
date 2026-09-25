"""Assemble a WMS workbook from the text sources (MASTER SPEC v0.3 §3: Basic sources are text, the ODS is a build artefact).

    python3 tools/build_ods.py OUT.ods            production skeleton (Phase 1: _SYS only)
    python3 tools/build_ods.py OUT.ods --test     test book: + test module and the synthetic _TST sheet, MODE=TEST

Needs LibreOffice with Python-UNO on the developer machine. Nothing is installed on the warehouse PC.
"""
import argparse
import os
import random
import sys
import tempfile
import uno

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wmslo import Office, add_basic, bind_event, props  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src", "basic")
TEST_SRC = os.path.join(ROOT, "tests", "basic")
PWD = "wms"                      # WmsConfig.PROTECT_PWD
SYS_KEYS = ["SCHEMA", "INSTANCE_ID", "MODE", "CORE_VERSION", "LAST_SEQ", "NEXT_EI", "NEXT_NO", "NEXT_RET",
            "JOURNAL_POS", "REGISTERED_URL", "TX_STATE", "TX_SEQ", "TX_TYPE", "TX_TIME", "TX_BEFORE_IMAGE",
            "SAVE_STAMP", "SAVE_SEQ", "MAX_QTY", "KEY_SHEETS"]        # WmsConfig.SysKeyNames()
SK_SAVE_STAMP = SYS_KEYS.index("SAVE_STAMP")
MODULES = ["WmsConfig", "WmsCore", "WmsJournal", "WmsLock", "WmsBackup", "WmsRecovery", "WmsDiagnostics"]
TEST_MODULES = ["WmsTestOps"]
TEST_SHEET = "_TST"
# _TST: A key (issued by WMS from NEXT_NO), B text (input), C quantity (text input), D control/status
TEST_KEY_SPEC = f"{TEST_SHEET}|0|3|{SYS_KEYS.index('NEXT_NO')}|1"


def read_source(path):
    s = open(path, encoding="utf-8").read()
    return s.replace("\r\n", "\n")


def build(office, out_path, test=False, instance_id=None):
    out_path = os.path.abspath(out_path)
    doc = office.new_calc()
    sheets = doc.Sheets
    sys_sh = sheets.getByIndex(0)
    sys_sh.Name = "_SYS"
    values = {
        "SCHEMA": "WMS-SYS-1", "INSTANCE_ID": instance_id or "%016X" % random.getrandbits(64), "MODE": "TEST" if test else "PROD",
        "CORE_VERSION": "", "LAST_SEQ": 0, "NEXT_EI": 1, "NEXT_NO": 1, "NEXT_RET": 1, "JOURNAL_POS": "", "REGISTERED_URL": "",
        "TX_STATE": "NONE", "TX_SEQ": 0, "TX_TYPE": "", "TX_TIME": "", "TX_BEFORE_IMAGE": "", "SAVE_STAMP": 0, "SAVE_SEQ": 0,
        "MAX_QTY": 100000, "KEY_SHEETS": TEST_KEY_SPEC if test else "",
    }
    for i, k in enumerate(SYS_KEYS):
        sys_sh.getCellByPosition(0, i).setString(k)
        v = values[k]
        if isinstance(v, (int, float)):
            sys_sh.getCellByPosition(1, i).setValue(v)
        else:
            sys_sh.getCellByPosition(1, i).setString(v)
    if test:
        sheets.insertNewByName(TEST_SHEET, 1)
        tst = sheets.getByName(TEST_SHEET)
        for c, h in enumerate(["№", "Текст", "Количество", "Статус"]):
            tst.getCellByPosition(c, 0).setString(h)
        from com.sun.star.util import CellProtection
        unl = CellProtection()
        unl.IsLocked = False
        tst.getCellRangeByPosition(0, 1, 3, 1048575).CellProtection = unl
        # quantity is entered as text and parsed by WMS (spec §12)
        from com.sun.star.lang import Locale
        nf = doc.NumberFormats
        k = nf.queryKey("@", Locale("ru", "RU", ""), False)
        if k == -1:
            k = nf.addNew("@", Locale("ru", "RU", ""))
        tst.getCellRangeByPosition(2, 1, 2, 1048575).NumberFormat = k
        tst.protect(PWD)
    names = MODULES + (TEST_MODULES if test else [])
    for m in names:
        folder = TEST_SRC if m in TEST_MODULES else SRC
        add_basic(doc, m, read_source(os.path.join(folder, m + ".bas")))
    bind_event(doc.Events, "OnLoad", "WmsCore.OnDocLoad")
    bind_event(doc.Events, "OnPrepareUnload", "WmsCore.OnDocPrepareUnload")
    bind_event(doc.Events, "OnSave", "WmsCore.OnDocSave")
    sys_sh.IsVisible = False
    sys_sh.protect(PWD)
    doc.protect(PWD)
    if os.path.exists(out_path):
        os.remove(out_path)
    doc.storeAsURL(uno.systemPathToFileUrl(out_path), props(FilterName="calc8"))
    doc.close(True)
    # the saved file must carry the save stamp WMS expects (EditingCycles after this save), otherwise the first
    # start would report "saved without WMS"; fix it up with macros disabled so no event interferes
    doc = office.load(out_path, macros=0, hidden=True)
    sys_sh = doc.Sheets.getByName("_SYS")
    ec = doc.getDocumentProperties().EditingCycles
    sys_sh.getCellByPosition(1, SK_SAVE_STAMP).setValue(ec + 1)
    doc.store()
    doc.close(True)
    doc = office.load(out_path, macros=0, hidden=True)
    ec = doc.getDocumentProperties().EditingCycles
    stamp = doc.Sheets.getByName("_SYS").getCellByPosition(1, SK_SAVE_STAMP).getValue()
    doc.close(True)
    if stamp != ec:
        raise RuntimeError(f"save stamp {stamp} != EditingCycles {ec}")
    return out_path


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out")
    ap.add_argument("--test", action="store_true", help="test book (test module, _TST sheet, MODE=TEST)")
    a = ap.parse_args()
    prof = tempfile.mkdtemp(prefix="wms_build_")
    o = Office("wmsbuild%d" % os.getpid(), prof)
    try:
        print("built", build(o, a.out, test=a.test))
    finally:
        o.terminate()


if __name__ == "__main__":
    main()
