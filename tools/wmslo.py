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

    def load(self, path, macros=4, hidden=False, **extra):
        """macros: css.document.MacroExecMode (0 never, 4 always without warning)."""
        kw = dict(MacroExecutionMode=macros, Hidden=hidden)
        kw.update(extra)
        doc = self.desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(path)), "_blank", 0, props(**kw))
        self.idle()
        return doc

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
        script = sp.getScript(f"vnd.sun.star.script:Standard.{module}.{func}?language=Basic&location=document")
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


def add_basic(doc, module, code):
    libs = doc.BasicLibraries
    if not libs.hasByName("Standard"):
        libs.createLibrary("Standard")
    libs.loadLibrary("Standard")
    lib = libs.getByName("Standard")
    if lib.hasByName(module):
        lib.replaceByName(module, code)
    else:
        lib.insertByName(module, code)


def bind_event(events, name, script):
    """Bind a document/sheet event to a Basic macro of the document's Standard library."""
    url = f"vnd.sun.star.script:Standard.{script}?language=Basic&location=document"
    uno.invoke(events, "replaceByName", (name, uno.Any("[]com.sun.star.beans.PropertyValue", props(EventType="Script", Script=url))))
