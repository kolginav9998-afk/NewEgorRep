"""Test harness for the Phase 1 reliability skeleton: builds books from the text sources, drives LibreOffice like a user
(UI dispatch, window close, process kill) and reads the book and journal independently."""
import glob
import json
import os
import shutil
import sys
import tempfile
import time
import traceback

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wmslo import Office, props  # noqa: E402
from build_ods import build, SYS_KEYS, synthetic_registry  # noqa: E402
import journal_oracle  # noqa: E402
import issue_oracle  # noqa: E402

ISSUE_COLS = "ABCDEFGHIJKLMNOPQR"
ORDER_COLS = [chr(65 + i) for i in range(26)] + ["AA", "AB"]
# initial balances of the synthetic EIs of a test book (tools/build_ods.synthetic_registry)
INITIAL_STOCK = {row[0]: row[4] for row in synthetic_registry()}

OUT = os.path.abspath(os.environ.get("WMS_TEST_OUT") or tempfile.mkdtemp(prefix="wms_phase1_"))
PROFILES = os.path.join(OUT, "profiles")
os.makedirs(PROFILES, exist_ok=True)
_templates = {}
_seq = [0]


def template(kind="TEST"):
    """a freshly built, never opened book (built once per run, then copied)"""
    if kind not in _templates:
        p = os.path.join(OUT, f"template_{kind}.ods")
        o = Office(f"tplbuild{kind}{os.getpid()}", PROFILES)
        try:
            build(o, p, test=(kind != "PROD"))
        finally:
            o.terminate()
        _templates[kind] = p
    return _templates[kind]


def new_wms(case, kind="TEST", folder_name=None, book_name="WMS.ods"):
    d = os.path.join(OUT, "cases", case if folder_name is None else folder_name)
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(d)
    p = os.path.join(d, book_name)
    shutil.copy(template(kind), p)
    return p


def unique(name):
    _seq[0] += 1
    return f"{name}{_seq[0]}_{os.getpid()}"


class Session:
    """one LibreOffice process with one WMS book open (macros=4: run as a trusted document; 0: macros disabled)"""

    def __init__(self, path, name="s", macros=4, locale="ru-RU", profile=None, fresh_profile=True):
        self.path = os.path.abspath(path)
        self.dir = os.path.dirname(self.path)
        self.jdir = os.path.join(self.dir, "WMS_Journal")
        self.o = Office(profile or unique(name), PROFILES, locale=locale, fresh=fresh_profile)
        self.stale_lo_lock_removed = False
        t0 = time.time()
        self.doc = self.o.load(self.path, macros=macros)
        if self.doc is None:
            # LibreOffice refuses a file with a stale lock of a killed process in headless mode; a user would press "Open"
            for f in glob.glob(os.path.join(self.dir, ".~lock.*#")):
                os.remove(f)
                self.stale_lo_lock_removed = True
            self.doc = self.o.load(self.path, macros=macros)
        self.open_s = time.time() - t0
        if self.doc is None:
            raise RuntimeError("book did not open")

    # Basic
    def B(self, func, *a, module="WmsCore"):
        return self.o.basic(self.doc, module, func, *a)

    def T(self, func, *a):
        return self.B(func, *a, module="WmsTestOps")

    def report(self):
        return self.B("StartupReport") or ""

    def state(self):
        d = {}
        for kv in (self.B("StateDump") or "").split(";"):
            k, _, v = kv.partition("=")
            d[k] = v
        return d

    # book
    def sysv(self, key):
        c = self.doc.Sheets.getByName("_SYS").getCellByPosition(1, SYS_KEYS.index(key))
        return c.getValue() if c.getType().value == "VALUE" else c.getString()

    def set_sys(self, key, value):
        c = self.doc.Sheets.getByName("_SYS").getCellByPosition(1, SYS_KEYS.index(key))
        if isinstance(value, (int, float)):
            c.setValue(value)
        elif value == "":
            c.clearContents(1 + 2 + 4 + 16)     # setString("") does not clear a text cell
        else:
            c.setString(value)

    def tst(self, r):
        return self.doc.Sheets.getByName("_TST").getCellRangeByPosition(0, r, 3, r).getDataArray()[0]

    def tst_cell(self, c, r):
        cell = self.doc.Sheets.getByName("_TST").getCellByPosition(c, r)
        return cell.getType().value, cell.getString(), cell.getValue()

    def locks(self, r):
        sh = self.doc.Sheets.getByName("_TST")
        return "".join("1" if sh.getCellByPosition(c, r).CellProtection.IsLocked else "0" for c in range(4))

    def type_in(self, c, r, text):
        """enter text like a user (UI path: GoToCell + EnterString on the _TST sheet)"""
        ctl = self.doc.getCurrentController()
        ctl.setActiveSheet(self.doc.Sheets.getByName("_TST"))
        self.o.dispatch(self.doc, ".uno:GoToCell", ToPoint="$_TST.$%s$%d" % ("ABCD"[c], r + 1))
        self.o.dispatch(self.doc, ".uno:EnterString", StringName=text)

    def input_row(self, r, text, qty):
        self.type_in(1, r, text)
        self.type_in(2, r, qty)

    def input_rows_api(self, r0, rows):
        """fast bulk input for batch tests (API, not the UI)"""
        sh = self.doc.Sheets.getByName("_TST")
        sh.getCellRangeByPosition(1, r0, 2, r0 + len(rows) - 1).setDataArray(tuple((t, q) for t, q in rows))

    def check(self, expect_tail=0):
        return journal_oracle.check(self.doc, self.jdir, expect_tail)

    # «Выдачи» (Phase 2)
    def I(self, func, *a):
        return self.B(func, *a, module="WmsIssue")

    def U(self, func, *a):
        return self.B(func, *a, module="WmsUi")

    def iss(self, r):
        """values of «Выдачи» row r (0-based; row 1 is the first data row)"""
        return tuple(self.doc.Sheets.getByName("Выдачи").getCellRangeByPosition(0, r, 17, r).getDataArray()[0])

    def iss_cell(self, col, r):
        cell = self.doc.Sheets.getByName("Выдачи").getCellByPosition(ISSUE_COLS.index(col), r)
        return cell.getType().value, cell.getString(), cell.getValue()

    def iss_locks(self, r):
        sh = self.doc.Sheets.getByName("Выдачи")
        return "".join("1" if sh.getCellByPosition(c, r).CellProtection.IsLocked else "0" for c in range(18))

    def stock(self, n):
        return self.doc.Sheets.getByName("Наличие").getCellByPosition(4, n).getValue()

    def goto(self, sheet, ref):
        ctl = self.doc.getCurrentController()
        ctl.setActiveSheet(self.doc.Sheets.getByName(sheet))
        self.o.dispatch(self.doc, ".uno:GoToCell", ToPoint=f"${sheet}.{ref}")

    def type_iss(self, col, r, text):
        """enter text into «Выдачи» like a user (GoToCell + EnterString): the sheet's change handler runs"""
        self.goto("Выдачи", f"${col}${r + 1}")
        self.o.dispatch(self.doc, ".uno:EnterString", StringName=text)

    def issue_input(self, r, ei=None, qty=None, date=None, who=None, doc=None, note=None):
        for col, v in (("L", ei), ("E", qty), ("H", date), ("I", who), ("B", doc), ("O", note)):
            if v is not None:
                self.type_iss(col, r, v)

    def click(self, button, r=None, col="L"):
        """a button of «Выдачи» for the row under the cursor (the same macro the button is bound to)"""
        if r is not None:
            self.goto("Выдачи", f"${col}${r + 1}")
        return self.U(button)

    def post(self, r):
        return self.I("IssuePostRow", r)

    def main_status(self):
        sh = self.doc.Sheets.getByName("Главная")
        return [sh.getCellByPosition(1, r).getString() for r in range(2, 9)]

    def active_sheet(self):
        return self.doc.getCurrentController().getActiveSheet().getName()

    def icheck(self, expect_tail=0, initial=None):
        return issue_oracle.check(self.doc, self.jdir, INITIAL_STOCK if initial is None else initial, expect_tail)

    # «Заказы» (Phase 3)
    def Rc(self, func, *a):
        return self.B(func, *a, module="WmsReceipt")

    def Od(self, func, *a):
        return self.B(func, *a, module="WmsOrders")

    def OU(self, func, *a):
        return self.B(func, *a, module="WmsOrdersUi")

    def ords(self, r):
        """values of «Заказы» row r (0-based; row 1 is the first data row), A..AB"""
        return tuple(self.doc.Sheets.getByName("Заказы").getCellRangeByPosition(0, r, 27, r).getDataArray()[0])

    def ord(self, r, col):
        return self.ords(r)[ORDER_COLS.index(col)]

    def ord_cell(self, col, r):
        cell = self.doc.Sheets.getByName("Заказы").getCellByPosition(ORDER_COLS.index(col), r)
        return cell.getType().value, cell.getString(), cell.getValue()

    def ord_locks(self, r):
        sh = self.doc.Sheets.getByName("Заказы")
        return "".join("1" if sh.getCellByPosition(c, r).CellProtection.IsLocked else "0" for c in range(28))

    def type_ord(self, col, r, text):
        """enter text into «Заказы» like a user (GoToCell + EnterString): the sheet's change handler runs"""
        self.goto("Заказы", f"${col}${r + 1}")
        self.o.dispatch(self.doc, ".uno:EnterString", StringName=text)

    def order_input(self, r, **cols):
        """cols: column letter → text, typed through the UI in the order given"""
        for col, v in cols.items():
            if v is not None:
                self.type_ord(col, r, v)

    def click_ord(self, button, r=None, col="B", r1=None):
        """a button of «Заказы» for the row (or rows r..r1) under the cursor — the macro the button is bound to"""
        if r is not None:
            ref = f"${col}${r + 1}" if r1 is None else f"$A${r + 1}:$AB${r1 + 1}"
            self.goto("Заказы", ref)
        self.OU(button)
        return self.U("TestUiLastMessage")

    def rcv(self, n):
        return tuple(self.doc.Sheets.getByName("_RCV").getCellRangeByPosition(0, n, 8, n).getDataArray()[0])

    def opos(self, olid):
        return tuple(self.doc.Sheets.getByName("_ORD").getCellRangeByPosition(0, olid, 8, olid).getDataArray()[0])

    def stock_row(self, n):
        return tuple(self.doc.Sheets.getByName("Наличие").getCellRangeByPosition(0, n, 8, n).getDataArray()[0])

    def next_ol(self):
        return self.doc.Sheets.getByName("_ORD").getCellByPosition(11, 0).getValue()

    def rcheck(self, expect_tail=0, initial=None, today=None, base=None):
        import receipt_oracle
        return receipt_oracle.check(self.doc, self.jdir, INITIAL_STOCK if initial is None else initial, expect_tail, today=today, base=base)

    def journal(self):
        return journal_oracle.read_journal(self.jdir)

    def lock_text(self):
        f = os.path.join(self.jdir, "wms.lock")
        return open(f, encoding="utf-8").read() if os.path.exists(f) else None

    def ui(self, cmd, **kw):
        return self.o.dispatch(self.doc, cmd, **kw)

    # leaving
    def close(self, save=False):
        """like a user: optional Save, then close the window (fires OnPrepareUnload); unsaved changes are discarded"""
        if save:
            self.doc.store()
        if self.doc.isModified():
            self.doc.setModified(False)
        try:
            self.o.dispatch(self.doc, ".uno:CloseDoc")
        except Exception:
            pass
        try:
            self.doc.close(True)
        except Exception:
            pass
        self.o.terminate()

    def kill(self):
        self.o.kill()


def journal_files(jdir):
    return sorted(glob.glob(os.path.join(jdir, "WMS_journal_*.csv")))


class Results:
    def __init__(self):
        self.items = []
        self.t0 = time.time()

    def add(self, case, name, ok, detail="", timing=None):
        st = ok if isinstance(ok, str) else ("PASS" if ok else "FAIL")
        self.items.append(dict(case=case, test=name, status=st, detail=detail, timing=timing))
        print(f"[{st}] {case} {name}: {detail[:400]}", flush=True)

    def error(self, case, exc):
        self.add(case, "исключение теста", False, "".join(traceback.format_exception_only(type(exc), exc)).strip()[:600])

    def save(self, path):
        json.dump(dict(items=self.items, seconds=round(time.time() - self.t0, 1)), open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
