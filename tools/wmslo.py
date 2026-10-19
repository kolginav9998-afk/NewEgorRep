"""LibreOffice session helper for the WMS build tool and tests (developer machine only; nothing here runs at the warehouse).

Starts soffice with a private user profile and a UNO pipe, optionally with a GUI display, a UI locale and a macro
security level; loads/creates documents, installs Basic modules, binds document events and calls Basic functions.
"""
import os
import signal
import subprocess
import time
import uno
from com.sun.star.beans import PropertyValue

SOFFICE = os.environ.get("WMS_SOFFICE", "soffice")


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def props(**kw):
    return tuple(prop(k, v) for k, v in kw.items())


def write_profile(profile, locale=None, macro_level=None, extra=""):
    """Pre-seed registrymodifications.xcu of a fresh profile (UI locale, macro security, anything else)."""
    d = os.path.join(profile, "user")
    os.makedirs(d, exist_ok=True)
    items = []
    if locale:
        items.append(f'<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>{locale}</value></prop></item>')
        items.append(f'<item oor:path="/org.openoffice.Office.Linguistic/General"><prop oor:name="DefaultLocale" oor:op="fuse"><value>{locale}</value></prop></item>')
    if macro_level is not None:
        items.append(f'<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>{macro_level}</value></prop></item>')
    if extra:
        items.append(extra)
    f = os.path.join(d, "registrymodifications.xcu")
    body = "\n".join(items)
    if os.path.exists(f):
        s = open(f, encoding="utf-8").read()
        s = s.replace("</oor:items>", body + "\n</oor:items>")
    else:
        s = ('<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
             'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
             + body + "\n</oor:items>\n")
    open(f, "w", encoding="utf-8").write(s)


class Office:
    def __init__(self, name, profile_root, headless=True, display=None, locale="ru-RU", macro_level=0, fresh=True):
        self.name = name
        self.profile = os.path.join(profile_root, "profile_" + name)
        if fresh and os.path.isdir(self.profile):
            import shutil
            shutil.rmtree(self.profile, ignore_errors=True)
        if not os.path.isdir(self.profile):
            write_profile(self.profile, locale=locale, macro_level=macro_level)
        args = [SOFFICE, "--norestore", "--nologo", "--nodefault", "--nofirststartwizard",
                "-env:UserInstallation=" + uno.systemPathToFileUrl(self.profile),
                f"--accept=pipe,name={name};urp;StarOffice.ComponentContext"]
        if headless:
            args[1:1] = ["--headless", "--invisible"]
        env = dict(os.environ)
        if display:
            env["DISPLAY"] = display
            env.setdefault("SAL_USE_VCLPLUGIN", "gen")
        self.proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env, start_new_session=True)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        self.ctx = None
        for _ in range(240):
            try:
                self.ctx = resolver.resolve(f"uno:pipe,name={name};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.25)
        if self.ctx is None:
            raise RuntimeError("LibreOffice did not start")
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        self.dispatcher = self.smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", self.ctx)
        self.toolkit = self.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)

    # ---------------------------------------------------------------- documents
    def new_calc(self):
        return self.desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, props(Hidden=True))

    def load(self, path, macros=4, hidden=False, wait_wms=True, **extra):
        """macros: css.document.MacroExecMode (0 never, 4 always without warning). 1.0: with the macros on, a WMS book is
        returned once its startup has finished (wait_wms)"""
        kw = dict(MacroExecutionMode=macros, Hidden=hidden)
        kw.update(extra)
        doc = self.desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(path)), "_blank", 0, props(**kw))
        self.idle()
        if doc is not None and macros and wait_wms and doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            self.wms_ready(doc)
        return doc

    def wms_state(self, doc, wait=60.0):
        """WMS's state line (WmsCore.StateDump) once its startup has run: on a loaded machine the OnLoad macro Calc
        posts after loading can still be queued — or running, its state half-made («STATE=BLOCKED;BLOCK=;») — when the
        load returns. 1.0 counts the finished startups («STARTS»); a book of an earlier core: while «STATE=» is empty"""
        end = time.time() + wait
        while True:
            st = str(self.basic(doc, "WmsCore", "StateDump") or "")
            busy = ";STARTS=0" in st if ";STARTS=" in st else st.startswith("STATE=;")
            if not busy or time.time() > end:
                return st
            self.idle()
            time.sleep(0.25)

    def wms_ready(self, doc, wait=60.0):
        """after the load of a book with the macros on: waits for the startup of WMS (nothing to wait for in a book without
        WMS — the call of WmsCore fails at once)"""
        try:
            return self.wms_state(doc, wait)
        except Exception:
            return ""

    def idle(self):
        """Let LibreOffice run what it has queued, in particular the document's OnLoad macro, which Calc posts
        asynchronously after loading. Calling a Basic macro through the script provider while that event is being
        dispatched can deadlock LibreOffice (the main thread holds the SolarMutex and waits for the script framework,
        the bridge thread holds the script framework and waits for the SolarMutex); a user never does both at once."""
        try:
            self.toolkit.processEventsToIdle()
        except Exception:
            pass

    def dispatch(self, doc, cmd, **args):
        frame = doc.getCurrentController().getFrame()
        return self.dispatcher.executeDispatch(frame, cmd, "", 0, props(**args))

    def basic(self, doc, module, func, *args):
        sp = doc.getScriptProvider()
        script = sp.getScript(f"vnd.sun.star.script:{lib_of(module)}.{module}.{func}?language=Basic&location=document")
        return script.invoke(tuple(args), (), ())[0]

    # ---------------------------------------------------------------- process
    def pid(self):
        out = subprocess.run(["pgrep", "-f", "profile_" + self.name + "( |$)"], capture_output=True, text=True).stdout.split()
        for p in out:
            try:
                if open(f"/proc/{p}/comm").read().strip() == "soffice.bin":
                    return int(p)
            except OSError:
                pass
        return None

    def terminate(self):
        try:
            self.desktop.terminate()
        except Exception:
            pass
        for _ in range(40):
            if self.pid() is None:
                return
            time.sleep(0.25)
        self.kill()

    def kill(self, sig=signal.SIGKILL):
        p = self.pid()
        if p:
            try:
                os.kill(p, sig)
            except OSError:
                pass
        for _ in range(40):
            if self.pid() is None:
                return
            time.sleep(0.1)


# The Basic of a WMS book lives in two libraries (PRIME 0.8). LibreOffice Basic looks a name of another module up through
# every module of the library: each module, constant and variable of the library makes every such lookup — the core reads
# its constants (WmsConfig) that way all the time — slower (measured: 0.8 in one library posted an issue 28 % slower than
# 0.7.2). The core of the accounting stays in «Standard» with the size of 0.7.2; the pages and every later addition live in
# «WmsPlus»: their names do not slow the core down. A library sees the other one's procedures, constants and globals once
# both are loaded (WmsCore.WmsInit loads WmsPlus).
PLUS_LIB = "WmsPlus"
PLUS_MODULES = ["WmsLegacy", "WmsPrimeUi", "WmsView", "WmsPack", "WmsDash", "WmsFind", "WmsMap", "WmsAsk", "WmsDaily", "WmsSupply", "WmsPrint", "WmsAudit"]


def lib_of(module):
    """the Basic library of a module of a WMS book"""
    return PLUS_LIB if module in PLUS_MODULES else "Standard"


def script_macro(code):
    """«Module.Func» of a script URL of a document macro (any library), or the URL itself when it is not one"""
    for lib in ("Standard", PLUS_LIB):
        head = f"vnd.sun.star.script:{lib}."
        if code.startswith(head):
            return code[len(head):].split("?")[0]
    return code


def add_basic(doc, module, code, lib=None):
    lib_name = lib or lib_of(module)
    libs = doc.BasicLibraries
    if not libs.hasByName(lib_name):
        libs.createLibrary(lib_name)
    libs.loadLibrary(lib_name)
    lib = libs.getByName(lib_name)
    if lib.hasByName(module):
        lib.replaceByName(module, code)
    else:
        lib.insertByName(module, code)


def macro_url(script):
    """the script URL of a document macro «Module.Func» in the library of its module"""
    return f"vnd.sun.star.script:{lib_of(script.split('.')[0])}.{script}?language=Basic&location=document"


def bind_event(events, name, script):
    """Bind a document/sheet event to a Basic macro of the document (the library of its module)."""
    url = macro_url(script)
    uno.invoke(events, "replaceByName", (name, uno.Any("[]com.sun.star.beans.PropertyValue", props(EventType="Script", Script=url))))
