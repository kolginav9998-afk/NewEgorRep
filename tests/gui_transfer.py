"""GUI check of WMS_LEGACY_TRANSFER.ods on a virtual screen (Xvfb), used like a user.

    Xvfb :99 &  WMS_TEST_OUT=/tmp/wms_gui python3 tests/gui_transfer.py [--display :99]

1. The main page «LEGACY TRANSFER»: the five steps and the four extra buttons, each bound to its macro, «5. СДЕЛАТЬ РАБОЧЕЙ»
   greyed out while the steps 2–4 are not done.
2. The old table is open in the same LibreOffice: its «Заказы» A:AB (values, formulas, a reference to its own sheet
   «Выдачи») copied with Ctrl+C (.uno:Copy) and pasted with Ctrl+V (.uno:Paste) into A1 of «1_Вставить_Заказы» — the
   question «overwrite?» answered, the data arrive as in the old sheet.
3. «2. ПРОВЕРИТЬ» with the real message window: the formulas become their values (the reference to the old book too),
   the results of the rows in AD:AE, «2_Проверка» filled; «3. СОЗДАТЬ ТЕСТОВУЮ WMS», «4. СВЕРИТЬ» — their messages; the
   button «5» becomes available.
4. «5. СДЕЛАТЬ РАБОЧЕЙ»: the question window, «Нет» — nothing is created; then the promotion (test mode) — WMS_WORK.
5. The old book: not modified in the window, its file unchanged (SHA-256).
Prints one line per step; exits 0 when every step passed.
"""
import argparse
import hashlib
import os
import shutil
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import uno  # noqa: E402
from harness import OUT, PROFILES, unique  # noqa: E402
from wmslo import Office, props  # noqa: E402
from gui_phase4 import step, state, RESULTS  # noqa: E402
from gui_m6 import buttons  # noqa: E402
import run_legacy_transfer as lt  # noqa: E402

TITLE = "WMS LEGACY TRANSFER"


def sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


def texts(w):
    out = []
    for cid in w.getChildren():
        try:
            v = state(w.getChild(cid)).get("Text")
            if v:
                out.append(str(v))
        except Exception:
            pass
    return out


def run_button(o, ut, doc, macro, answer="ok", timeout=900):
    """the macro of a button in a thread; every message window of the book is closed (answer: the button id). The text of
    a Basic message box is not in the UI tree: the book records what it shows (LtTestLastMessage, LtTestLastAsk) — the
    windows seen (title, buttons) and the recorded text of the last message are returned"""
    o.basic(doc, "LtMain", "LtTestAuto", False)
    seen = []
    th = threading.Thread(target=lambda: o.basic(doc, "LtMain", macro), daemon=True)
    th.start()
    t0 = time.time()
    while th.is_alive() and time.time() - t0 < timeout:
        time.sleep(0.3)
        try:
            w = ut.getTopFocusWindow()
            title = str(state(w).get("Text", ""))
            if title.startswith(TITLE):
                kids = list(w.getChildren())
                seen.append((title, " | ".join(texts(w))))
                for b in (answer, "ok", "close"):
                    if b in kids:
                        w.getChild(b).executeAction("CLICK", ())
                        break
        except Exception:
            pass
    th.join(5)
    return seen, str(o.basic(doc, "LtMain", "LtTestLastMessage"))


def paste(o, ut, doc):
    """Ctrl+V into the selected cell; the question «the cells are not empty — replace?» answered «Да»"""
    th = threading.Thread(target=lambda: o.dispatch(doc, ".uno:Paste"), daemon=True)
    th.start()
    asked = []
    for _ in range(120):
        th.join(0.25)
        if not th.is_alive():
            break
        try:
            w = ut.getTopFocusWindow()
            kids = list(w.getChildren())
            t = texts(w)
            if any("содерж" in x or "contain" in x for x in t) or "yes" in kids:
                asked.append(" | ".join(t)[:160])
                if "yes" in kids:
                    w.getChild("yes").executeAction("CLICK", ())
                else:
                    w.executeAction("TYPE", props(KEYCODE="RETURN"))
        except Exception:
            pass
    th.join(10)
    return asked


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--display", default=os.environ.get("DISPLAY", ":99"))
    a = ap.parse_args()
    rel = lt.release()
    cd = os.path.join(OUT, "cases", "gui_transfer")
    shutil.rmtree(cd, ignore_errors=True)
    os.makedirs(cd)
    for f in ("WMS_LEGACY_TRANSFER.ods", "WMS_PROD_CANDIDATE.ods", "RELEASE_MANIFEST.csv"):
        shutil.copy2(os.path.join(rel, f), cd)
    shutil.copytree(os.path.join(rel, "tools"), os.path.join(cd, "tools"))
    shutil.copytree(os.path.join(rel, "WMS_TOOLBOX"), os.path.join(cd, "WMS_TOOLBOX"))
    # the old table: «Заказы» with values and formulas, «Выдачи» that one formula refers to
    old = os.path.join(cd, "СТАРАЯ_ТАБЛИЦА.ods")
    rows = [lt.HDR] + lt.full_rows()
    mk = Office(unique("ltguimk"), PROFILES)
    try:
        d = mk.new_calc()
        d.Sheets.getByIndex(0).Name = "Заказы"
        d.Sheets.insertNewByName("Выдачи", 1)
        d.Sheets.getByName("Выдачи").getCellByPosition(1, 1).setValue(60.0)
        sh = d.Sheets.getByName("Заказы")
        sh.getCellRangeByPosition(0, 0, 27, len(rows) - 1).setDataArray(tuple(tuple(r) for r in rows))
        sh.getCellByPosition(23, 1).setFormula("=$Выдачи.B2")                      # X2 «Наличие» from the old sheet «Выдачи»
        sh.getCellByPosition(10, 1).setFormula("=F2*J2")                           # K2 «Сумма»
        sh.getCellByPosition(22, 2).setFormula('=IF(Q3<TODAY();"Просрочено";"Ожидается")')   # W3 the old status formula
        d.storeToURL(uno.systemPathToFileUrl(old), props(FilterName="calc8"))
        d.close(True)
    finally:
        mk.terminate()
    h_old = sha(old)
    o = Office(unique("ltgui"), PROFILES, headless=False, display=a.display)
    try:
        ut = o.smgr.createInstanceWithContext("com.sun.star.ui.test.UITest", o.ctx)
        doc = o.load(os.path.join(cd, "WMS_LEGACY_TRANSFER.ods"), macros=4)
        o.idle()
        main_sh = doc.Sheets.getByName("LEGACY TRANSFER")
        bt = buttons(doc, main_sh)
        want = {"btnPaste": ("1. ВСТАВИТЬ ДАННЫЕ", "LtMain.BtnLtPaste"), "btnCheck": ("2. ПРОВЕРИТЬ", "LtMain.BtnLtCheck"),
                "btnBuild": ("3. СОЗДАТЬ ТЕСТОВУЮ WMS", "LtMain.BtnLtBuild"), "btnVerify": ("4. СВЕРИТЬ", "LtMain.BtnLtVerify"),
                "btnPromote": ("5. СДЕЛАТЬ РАБОЧЕЙ", "LtMain.BtnLtPromote"), "btnClear": ("Очистить staging", "LtMain.BtnLtClear"),
                "btnErrors": ("Открыть ошибки", "LtMain.BtnLtErrors"), "btnExport": ("Экспорт отчёта", "LtMain.BtnLtExport"),
                "btnBackup": ("Создать резервную копию", "LtMain.BtnLtBackup")}
        got = {k: (v[0], v[1]) for k, v in bt.items()}
        pm = main_sh.getDrawPage().getForms().getByIndex(0).getByName("btnPromote")
        step("главная страница: пять шагов и четыре дополнительные кнопки, каждая со своим макросом; «5. СДЕЛАТЬ РАБОЧЕЙ» недоступна, пока не "
             "выполнены шаги 2–4", got == want and not pm.Enabled, str(got if got != want else pm.Enabled))
        # 1. the paste step: the sheet and A1
        seen, m = run_button(o, ut, doc, "BtnLtPaste")
        act = doc.getCurrentController().getActiveSheet().getName()
        sel = doc.getCurrentController().getSelection().getRangeAddress()
        step("«1. ВСТАВИТЬ ДАННЫЕ»: окно с подсказкой (Ctrl+C в старой таблице, Ctrl+V в A1), открыт лист «1_Вставить_Заказы», выбрана A1",
             seen and "Ctrl+C" in m and act == "1_Вставить_Заказы" and (sel.StartColumn, sel.StartRow) == (0, 0), f"{seen[:1]}; {m[:80]}; {act}")
        # the old table: A:AB copied
        olddoc = o.load(old, macros=0)
        o.idle()
        osh = olddoc.Sheets.getByName("Заказы")
        olddoc.getCurrentController().select(osh.getCellRangeByName(f"A1:AB{len(rows)}"))
        o.dispatch(olddoc, ".uno:Copy")
        doc.getCurrentController().getFrame().getContainerWindow().setFocus()
        doc.getCurrentController().getFrame().activate()
        o.idle()
        tsh = doc.Sheets.getByName("1_Вставить_Заказы")
        doc.getCurrentController().setActiveSheet(tsh)
        doc.getCurrentController().select(tsh.getCellRangeByName("A1"))
        asked = paste(o, ut, doc)
        pasted = [list(r) for r in tsh.getCellRangeByPosition(0, 0, 27, len(rows) - 1).getDataArray()]
        ftext = tsh.getCellByPosition(23, 1).getFormula()
        step("Ctrl+C в старой таблице (A:AB) → Ctrl+V в A1 листа «1_Вставить_Заказы»: вопрос о замене заголовка — «Да»; данные на месте, "
             "формула со ссылкой на лист старой книги стала внешней ссылкой",
             pasted[3][1] == rows[3][1] and pasted[5][21] == rows[5][21] and pasted[1][23] == 60.0 and "file:" in ftext and len(asked) <= 1,
             f"{asked}; {ftext[:90]}; {pasted[1][23]}")
        # 2. ПРОВЕРИТЬ
        seen, m = run_button(o, ut, doc, "BtnLtCheck")
        x2 = tsh.getCellByPosition(23, 1)
        ad = [tuple(r) for r in tsh.getCellRangeByPosition(29, 0, 30, 3).getDataArray()]
        chk = {r[0]: r[1] for r in doc.Sheets.getByName("2_Проверка").getCellRangeByPosition(0, 2, 2, 80).getDataArray() if r[0]}
        step("«2. ПРОВЕРИТЬ»: окно «блокирующих замечаний нет» (формул заменено значениями: 3); X2 — значение 60 вместо ссылки на старую книгу; "
             "результаты строк в AD:AE; «2_Проверка» — итог «Можно создавать тестовую WMS»",
             seen and "блокирующих замечаний нет" in m and "формул заменено значениями: 3" in m and x2.getType().value == "VALUE"
             and x2.getValue() == 60.0 and ad[0] == ("Результат проверки", "Что не так / что сделано") and ad[1][0] in ("Готово", "Проверить")
             and chk.get("ИТОГ") == "Можно создавать тестовую WMS" and not pm.Enabled, f"{seen[-1:]}; {m[:120]}; {ad[:2]}; {chk.get('ИТОГ')}")
        seen, m = run_button(o, ut, doc, "BtnLtBuild")
        step("«3. СОЗДАТЬ ТЕСТОВУЮ WMS»: окно «Тестовая WMS создана и проверена», книга LEGACY_WORK/test/WMS_LEGACY_TEST.ods; кнопка «5» ещё "
             "недоступна", seen and "Тестовая WMS создана" in m and os.path.isfile(os.path.join(cd, "LEGACY_WORK", "test", "WMS_LEGACY_TEST.ods"))
             and not pm.Enabled, f"{seen[-1:]}; {m[:200]}")
        seen, m = run_button(o, ut, doc, "BtnLtVerify")
        step("«4. СВЕРИТЬ»: окно «критических расхождений нет», отчёт LEGACY_TRANSFER_REPORT.html; кнопка «5. СДЕЛАТЬ РАБОЧЕЙ» стала доступна",
             seen and "критических расхождений нет" in m and os.path.isfile(os.path.join(cd, "LEGACY_WORK", "LEGACY_TRANSFER_REPORT.html"))
             and pm.Enabled, f"{seen[-1:]}; {m[:200]}")
        seen, m = run_button(o, ut, doc, "BtnLtPromote", answer="close")
        ask = str(o.basic(doc, "LtMain", "LtTestLastAsk"))
        step("«5. СДЕЛАТЬ РАБОЧЕЙ»: окно-вопрос с правилом перехода (кнопки «Да» / «Нет»); «Нет» — окно «Ничего не изменено», папки WMS_WORK нет",
             len(seen) >= 2 and "~No" in seen[0][1] and "все новые движения записываются только в новую WMS" in ask and "Ничего не изменено" in m
             and not os.path.exists(os.path.join(cd, "WMS_WORK")), f"{seen}; {m[:80]}")
        o.basic(doc, "LtMain", "LtTestAuto", True)
        o.basic(doc, "LtMain", "LtTestConfirm", 1)
        o.basic(doc, "LtMain", "BtnLtPromote")
        msg = str(o.basic(doc, "LtMain", "LtTestLastMessage"))
        step("«Да» — рабочая WMS: WMS_WORK/WMS_PROD.ods, CUTOVER_README.txt; кнопка снова недоступна; на главной — путь рабочей WMS",
             msg.startswith("OK:") and os.path.isfile(os.path.join(cd, "WMS_WORK", "WMS_PROD.ods")) and not pm.Enabled
             and "рабочая WMS" in main_sh.getCellByPosition(2, 7).getString(), msg[:200])
        step("старая таблица: в окне не изменена, её файл не изменился (SHA-256)", not olddoc.isModified() and sha(old) == h_old,
             f"{olddoc.isModified()}; {sha(old)[:12]} / {h_old[:12]}")
        olddoc.close(True)
        doc.setModified(False)
        doc.close(True)
    finally:
        o.terminate()
    print(f"GUI: {sum(RESULTS)} из {len(RESULTS)} PASS")
    sys.exit(0 if all(RESULTS) else 1)


if __name__ == "__main__":
    main()
