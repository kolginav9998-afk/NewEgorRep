"""Builds WMS_UPGRADE.ods — the wizard «Обновление из текущей WMS» (PRIME 0.8, D-091).

    python3 tools/build_upgrade.py OUT_DIR

The book is the frontend of tools/upgrade.py: it lies in the release folder next to the folder tools/. Sheets:
«ОБНОВЛЕНИЕ» (the working book and the new folder with «Выбрать…», the four steps as big buttons with their state:
1. ПРОВЕРИТЬ, 2. РЕПЕТИЦИЯ, 3. СОБРАТЬ И СВЕРИТЬ, 4. СДЕЛАТЬ РАБОЧЕЙ; «Открыть отчёт», «Открыть новую WMS», the help) and
«Отчёт» (UPGRADE_REPORT.md of the new folder, line by line). The Basic module src/upgrade/UpMain.bas.
"""
import argparse
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import uno  # noqa: E402
from wmslo import Office, add_basic, bind_event, props  # noqa: E402
from build_ods import add_button, add_cond, cell_style, read_source  # noqa: E402

BOOK = "WMS_UPGRADE.ods"
MAIN = "ОБНОВЛЕНИЕ"
REPORT = "Отчёт"
ROW_BOOK, ROW_TO, ROW_STEP1, ROW_STATE = 4, 5, 8, 13
STEPS = [("btnCheck", "1. ПРОВЕРИТЬ", "UpMain.BtnUpCheck", "проверка книги: она только читается"),
         ("btnRehearsal", "2. РЕПЕТИЦИЯ", "UpMain.BtnUpRehearsal", "по желанию"),
         ("btnBuild", "3. СОБРАТЬ И СВЕРИТЬ", "UpMain.BtnUpBuild", "новая книга в новой папке"),
         ("btnPromote", "4. СДЕЛАТЬ РАБОЧЕЙ", "UpMain.BtnUpPromote", "в конце рабочего дня")]
EXTRA = [("btnReport", "Открыть отчёт", "UpMain.BtnUpReport"), ("btnOpenNew", "Открыть новую WMS", "UpMain.BtnUpOpenNew")]
HELP = [
    ("Как обновиться", True),
    ("Все ваши данные переходят как есть: заказы, приходы, ЕИ, выдачи, возвраты, корректировки, машины, получатели и журнал. Ничего не "
     "вводится заново.", False),
    ("1. Закройте рабочую WMS. Нажмите «Выбрать…» и укажите её книгу (обычно WMS_WORK/WMS_PROD.ods). Новая папка предлагается рядом — "
     "в ней будет WMS 0.8.", False),
    ("2. «1. ПРОВЕРИТЬ» — книга только читается: WMS закрыта, нет незавершённой операции, журнал полный, нет листов с чужим паролем, "
     "хватает места. Если что-то мешает — мастер скажет, что именно и что сделать.", False),
    ("3. «2. РЕПЕТИЦИЯ» (по желанию) — всё обновление целиком во временной папке, без перехода: можно убедиться заранее.", False),
    ("4. «3. СОБРАТЬ И СВЕРИТЬ» — побайтная копия книги и журнала в новой папке, изменения версии 0.8 в копии, сверка каждой ячейки "
     "старой и новой книги, затем WMS 0.8 открывает копию проверки: «работа разрешена», самопроверка, «остаток = сумма движений», "
     "независимая проверка журнала. Любая ошибка — копия удаляется целиком, остаётся отчёт. После сборки можно спокойно работать "
     "в старой книге до конца дня.", False),
    ("5. «4. СДЕЛАТЬ РАБОЧЕЙ» — в конце рабочего дня, старая WMS закрыта. Если в старой книге после сборки провели операции, мастер "
     "попросит собрать заново, чтобы ничего не потерялось. После перехода работайте только в новой книге.", False),
    ("Откат: старая книга никогда не изменяется — в любой момент можно продолжить работу в ней. Отчёт UPGRADE_SUMMARY.txt в новой "
     "папке не содержит данных склада (только счётчики и контрольные суммы) — его можно прислать разработчику.", False),
]


def build(o, out_dir):
    out_dir = os.path.abspath(out_dir)
    path = os.path.join(out_dir, BOOK)
    doc = o.new_calc()
    try:
        cell_style(doc, "UP_Green", CellBackColor=0xD9EAD3, CharColor=0x1E5631)
        cell_style(doc, "UP_Red", CellBackColor=0xF4CCCC, CharColor=0x8B1A1A)
        cell_style(doc, "UP_Yellow", CellBackColor=0xFFF2CC)
        sheets = doc.Sheets
        sheets.getByIndex(0).Name = MAIN
        sheets.insertNewByName(REPORT, 1)
        m = sheets.getByName(MAIN)
        for c, w in enumerate((600, 7600, 17500, 3600)):
            m.getColumns().getByIndex(c).Width = w
        t = m.getCellByPosition(1, 0)
        t.setString("WMS PRIME 0.8 — обновление вашей рабочей WMS")
        t.CharHeight = 20
        t.CharWeight = 150
        t.CharColor = 0x1F3864
        m.getRows().getByIndex(0).Height = 1100
        s = m.getCellByPosition(1, 1)
        s.setString("Данные переходят как есть. Рабочая книга не изменяется: новая собирается рядом, в новой папке, сверяется ячейка в "
                    "ячейку и становится рабочей только по вашей кнопке.")
        s.CharColor = 0x404040
        s.IsTextWrapped = True
        m.getCellRangeByPosition(1, 1, 3, 1).merge(True)
        m.getRows().getByIndex(1).Height = 900
        m.getRows().getByIndex(2).Height = 300
        sec = m.getCellByPosition(1, 3)
        sec.setString("Что обновляем")
        sec.CharWeight = 150
        sec.CharColor = 0x1F3864
        for r, label in ((ROW_BOOK, "Рабочая книга WMS (0.7)"), (ROW_TO, "Новая папка для WMS 0.8")):
            m.getCellByPosition(1, r).setString(label)
            m.getCellByPosition(1, r).CharWeight = 150
            c = m.getCellByPosition(2, r)
            c.CellBackColor = 0xFFF8E1
            c.VertJustify = 2
            m.getRows().getByIndex(r).Height = 850
            m.getCellByPosition(1, r).VertJustify = 2
        m.getRows().getByIndex(6).Height = 300
        sec = m.getCellByPosition(1, 7)
        sec.setString("Шаги")
        sec.CharWeight = 150
        sec.CharColor = 0x1F3864
        for i, (name, label, macro, hint) in enumerate(STEPS):
            r = ROW_STEP1 + i
            m.getRows().getByIndex(r).Height = 1250
            c = m.getCellByPosition(2, r)
            c.VertJustify = 2
            c.IsTextWrapped = True
            c.setString("— " + hint)
        rng = m.getCellRangeByPosition(2, ROW_STEP1, 2, ROW_STEP1 + 3)
        add_cond(m, rng, f'LEFT($C{ROW_STEP1 + 1};1)="✓"', "UP_Green", 2, ROW_STEP1)
        add_cond(m, rng, f'LEFT($C{ROW_STEP1 + 1};1)="✗"', "UP_Red", 2, ROW_STEP1)
        add_cond(m, rng, f'LEFT($C{ROW_STEP1 + 1};1)="…"', "UP_Yellow", 2, ROW_STEP1)
        m.getRows().getByIndex(12).Height = 300
        m.getCellByPosition(1, ROW_STATE).setString("Состояние")
        m.getCellByPosition(1, ROW_STATE).CharWeight = 150
        m.getCellByPosition(1, ROW_STATE).VertJustify = 1
        st = m.getCellByPosition(2, ROW_STATE)
        st.setString("Закройте рабочую WMS и выберите её книгу («Выбрать…»)")
        st.IsTextWrapped = True
        st.VertJustify = 1
        m.getCellRangeByPosition(2, ROW_STATE, 3, ROW_STATE).merge(True)
        m.getRows().getByIndex(ROW_STATE).Height = 1900
        add_cond(m, m.getCellRangeByPosition(2, ROW_STATE, 2, ROW_STATE), f'LEFT($C{ROW_STATE + 1};1)="✓"', "UP_Green", 2, ROW_STATE)
        add_cond(m, m.getCellRangeByPosition(2, ROW_STATE, 2, ROW_STATE), f'LEFT($C{ROW_STATE + 1};1)="✗"', "UP_Red", 2, ROW_STATE)
        m.getRows().getByIndex(14).Height = 1000
        r0 = 16
        for k, (txt, head) in enumerate(HELP):
            c = m.getCellByPosition(1, r0 + k)
            c.setString(txt)
            c.IsTextWrapped = True
            c.VertJustify = 1
            m.getCellRangeByPosition(1, r0 + k, 3, r0 + k).merge(True)
            if head:
                c.CharWeight = 150
                c.CharColor = 0x1F3864
                m.getRows().getByIndex(r0 + k).Height = 600
            else:
                m.getRows().getByIndex(r0 + k).Height = 450 + 420 * (len(txt) // 150)
        # the buttons (after the row heights: their positions come from the cells)
        for r, name, macro in ((ROW_BOOK, "btnPickBook", "UpMain.BtnUpPickBook"), (ROW_TO, "btnPickFolder", "UpMain.BtnUpPickFolder")):
            cp = m.getCellByPosition(3, r)
            add_button(doc, m, name, "Выбрать…", macro, cp.Position.X + 100, cp.Position.Y + 80, 3400, 690)
        for i, (name, label, macro, hint) in enumerate(STEPS):
            cp = m.getCellByPosition(1, ROW_STEP1 + i)
            add_button(doc, m, name, label, macro, cp.Position.X + 100, cp.Position.Y + 100, 7400, 1050, font=13, bold=True,
                       back=(0xC6E0B4 if name == "btnPromote" else 0xDDEBF7))
        cp = m.getCellByPosition(1, 14)
        for i, (name, label, macro) in enumerate(EXTRA):
            add_button(doc, m, name, label, macro, cp.Position.X + 100 + i * 6400, cp.Position.Y + 150, 6100, 700)
        # «Отчёт»
        rp = sheets.getByName(REPORT)
        rp.getColumns().getByIndex(0).Width = 30000
        h = rp.getCellByPosition(0, 0)
        h.setString("Отчёт обновления")
        h.CharHeight = 16
        h.CharWeight = 150
        h.CharColor = 0x1F3864
        rp.getRows().getByIndex(0).Height = 900
        rp.getCellByPosition(0, 2).setString("Отчёта ещё нет: он появляется после «2. РЕПЕТИЦИЯ» или «3. СОБРАТЬ И СВЕРИТЬ».")
        rp.getCellRangeByPosition(0, 2, 0, 3000).IsTextWrapped = True
        add_button(doc, rp, "btnBack", "← К шагам обновления", "UpMain.BtnUpBack", 21000, 100, 7000, 700)
        src = read_source(os.path.join(ROOT, "src", "upgrade", "UpMain.bas"))
        add_basic(doc, "UpMain", src)
        bind_event(doc.Events, "OnLoad", "UpMain.UpOnLoad")
        ctl = doc.getCurrentController()
        ctl.setActiveSheet(m)
        doc.storeAsURL(uno.systemPathToFileUrl(path), props(FilterName="calc8"))
    finally:
        doc.close(True)
    return path


def version():
    src = read_source(os.path.join(ROOT, "src", "upgrade", "UpMain.bas"))
    return re.search(r'Public Const UP_VERSION = "([^"]+)"', src).group(1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out")
    a = ap.parse_args()
    out = os.path.abspath(a.out)
    if os.path.exists(os.path.join(out, BOOK)):
        print(f"ОШИБКА: {os.path.join(out, BOOK)} уже существует — собирайте в новую папку", file=sys.stderr)
        return 2
    os.makedirs(out, exist_ok=True)
    import tempfile
    import shutil
    prof = tempfile.mkdtemp(prefix="wms_upgrade_book_")
    o = Office(f"upbuild{os.getpid()}", prof)
    try:
        p = build(o, out)
    finally:
        o.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    print(f"OK: {p}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
