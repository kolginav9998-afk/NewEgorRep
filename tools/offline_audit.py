"""Проверка «WMS полностью локальна»: в папке выпуска или рабочей WMS ищется всё, чем программа могла бы выйти в сеть.
Только чтение — ни один файл не меняется.

    python3 tools/offline_audit.py ПАПКА [--out отчёт.txt]

Что проверяется:
  книги .ods  — гиперссылки, связи с внешними файлами и листами, DDE, функции Calc, читающие интернет (WEBSERVICE,
                FILTERXML), адреса http/ftp вне служебных имён формата ODF, макросы Basic внутри книг (сетевые службы
                LibreOffice, почта, браузер, запуск программ — каждая запускаемая команда показывается);
  макросы .bas и программы .py — сетевые модули и вызовы, адреса, запуск внешних программ.

Итог: «СЕТЕВЫХ ВОЗМОЖНОСТЕЙ НЕ НАЙДЕНО» (код 0) или список находок (код 1). Служебные имена формата ODF
(http://openoffice.org/…, http://www.w3.org/…, http://docs.oasis-open.org/…, http://purl.org/dc/…) — это имена разметки
внутри файла, по ним ничего не скачивается; они не считаются находкой.
"""
import argparse
import html
import os
import re
import sys
import zipfile

# the names of the ODF format inside a file (namespaces, RDF types, digest algorithms): identifiers, never fetched
FORMAT_NAMES = ("http://openoffice.org/", "http://www.w3.org/", "http://docs.oasis-open.org/", "http://purl.org/dc/",
                "http://www.oasis-open.org/")
URL = re.compile(r"(?:https?|ftp|ftps|smb|sftp|ws|wss)://[^\s\"'<>)]+", re.I)
# Basic and the API of LibreOffice: what reaches the network, the mail, a browser or another program
BASIC_NET = [
    (re.compile(r"SimpleSystemMail|SimpleCommandMail|SimpleMailClient|MailMerge|com\.sun\.star\.mail", re.I), "отправка почты"),
    (re.compile(r"com\.sun\.star\.(connection|bridge)\.", re.I), "сетевое соединение UNO"),
    (re.compile(r"XMLHttp|WinHttp|MSXML2|InternetExplorer|Inet\.", re.I), "HTTP-запрос через Windows"),
    (re.compile(r"\.uno:OpenHyperlink|\.uno:OpenUrl", re.I), "открытие гиперссылки"),
    (re.compile(r"\bWEBSERVICE\s*\(|\bFILTERXML\s*\(", re.I), "функция Calc, читающая интернет"),
]
PY_NET = [
    (re.compile(r"^\s*(?:import|from)\s+(urllib|http|requests|ftplib|smtplib|poplib|imaplib|telnetlib|xmlrpc|webbrowser|paramiko|"
                r"asyncio|ssl|aiohttp|httpx|websocket)\b", re.M), "сетевой модуль Python"),
    (re.compile(r"\burlopen\s*\(|\bcreate_connection\s*\(|\.connect\s*\(\s*\(", re.M), "сетевое соединение"),
]
# not the network, but a way out of WMS to another program of the computer: shown to be looked at by a person
BASIC_LOOK = [
    (re.compile(r"SystemShellExecute", re.I), "открытие файла программой компьютера (SystemShellExecute) — посмотрите, какого файла"),
    (re.compile(r"CreateObject\s*\(\s*\"WScript\.Shell\"", re.I), "WScript.Shell (Windows: ярлык, папки рабочего стола)"),
]
PY_SOCKET = re.compile(r"socket\.(\w+)")
SHELL = re.compile(r"\bShell\s*\(([^\n]*)", re.I)
SUBPROC = re.compile(r"subprocess\.(?:run|Popen|call|check_output|check_call)\s*\(\s*(\[[^\]]*\]|[^,)\n]+)")


def is_format_name(u):
    return u.startswith(FORMAT_NAMES)


class Audit:
    def __init__(self):
        self.finds, self.looks, self.notes, self.files = [], [], [], 0

    def look(self, where, what):
        self.looks.append(f"ВНИМАНИЕ {where}: {what}")

    def find(self, where, what):
        self.finds.append(f"НАЙДЕНО  {where}: {what}")

    def note(self, where, what):
        self.notes.append(f"сведение {where}: {what}")

    def code(self, where, text, kind):
        for rx, what in (BASIC_NET if kind == "basic" else PY_NET):
            for m in rx.finditer(text):
                self.find(where, f"{what} — «{m.group(0).strip()[:80]}»")
        for u in URL.findall(text):
            if not is_format_name(u) and not u.startswith(("http://localhost", "http://127.0.0.1")):
                self.find(where, f"адрес в коде — {u[:100]}")
            elif u.startswith(("http://localhost", "http://127.0.0.1")):
                self.note(where, f"адрес этого же компьютера в подсказке — {u[:60]}")
        if kind == "basic":
            for rx, what in BASIC_LOOK:
                for m in rx.finditer(text):
                    line = text[text.rfind("\n", 0, m.start()) + 1:text.find("\n", m.end())].strip()
                    self.look(where, f"{what}: «{line[:140]}»")
            for m in SHELL.finditer(text):
                self.note(where, "запуск программы этого компьютера: " + m.group(1).strip()[:160])
        else:
            for f in sorted(set(PY_SOCKET.findall(text))):
                if f == "gethostname":
                    self.note(where, "socket.gethostname — имя этого компьютера для отчёта (без сети)")
                else:
                    self.find(where, f"socket.{f}")
            for m in SUBPROC.finditer(text):
                self.note(where, "запуск программы этого компьютера: " + m.group(1).strip()[:120])

    def book(self, path, rel):
        try:
            z = zipfile.ZipFile(path)
        except zipfile.BadZipFile:
            self.find(rel, "не читается как книга ODF")
            return
        for n in z.namelist():
            if n.endswith("/"):
                continue
            data = z.read(n).decode("utf-8", "replace")
            where = f"{rel} [{n}]"
            if n.startswith("Basic/") and n.endswith(".xml") and "script-l" not in n:
                self.code(where, html.unescape(data), "basic")
                continue
            if not n.endswith(".xml") and not n.endswith(".rdf"):
                continue
            if "<text:a " in data:
                self.find(where, f"гиперссылок: {data.count('<text:a ')}")
            for tag, what in (("table:table-source", "связь с внешним файлом или листом"), ("office:dde-source", "связь DDE"),
                              ("table:dde-link", "связь DDE")):
                if tag in data:
                    self.find(where, what)
            for m in re.finditer(r'xlink:href="([^"]+)"', data):
                h = m.group(1)
                if h.startswith(("#", "vnd.sun.star.script:", "Pictures/", "./", "Basic/", "Configurations2/", "ObjectReplacements/",
                                 "Object ")) or is_format_name(h):
                    continue
                self.find(where, f"внешняя ссылка — {h[:100]}")
            for m in re.finditer(r"(?:of:=|table:formula=\")[^\"]*\b(WEBSERVICE|FILTERXML|HYPERLINK|DDE)\s*\(", data):
                self.find(where, f"формула {m.group(1)}")
            for u in URL.findall(html.unescape(data)):
                if not is_format_name(u):
                    self.find(where, f"адрес в содержимом — {u[:100]}")

    def run(self, root):
        for d, dirs, files in os.walk(root):
            dirs[:] = [x for x in dirs if x not in ("__pycache__", ".git")]
            for fn in sorted(files):
                p = os.path.join(d, fn)
                rel = os.path.relpath(p, root)
                low = fn.lower()
                if low.endswith((".ods", ".ots")):
                    self.files += 1
                    self.book(p, rel)
                elif low.endswith((".bas", ".xba")):
                    self.files += 1
                    self.code(rel, open(p, encoding="utf-8", errors="replace").read(), "basic")
                elif low.endswith(".py"):
                    self.files += 1
                    self.code(rel, open(p, encoding="utf-8", errors="replace").read(), "py")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("folder")
    ap.add_argument("--out")
    a = ap.parse_args()
    au = Audit()
    au.run(os.path.abspath(a.folder))
    head = [f"Проверка локальности WMS: {os.path.abspath(a.folder)}", f"проверено файлов (книги, макросы, программы): {au.files}", ""]
    if au.finds:
        tail = ["ИТОГ: НАЙДЕНО ВОЗМОЖНОСТЕЙ ВЫХОДА В СЕТЬ: " + str(len(au.finds)) + " — см. строки «НАЙДЕНО»"]
    else:
        tail = ["ИТОГ: СЕТЕВЫХ ВОЗМОЖНОСТЕЙ НЕ НАЙДЕНО — ни гиперссылок, ни внешних связей, ни сетевых вызовов, ни адресов в интернете."]
    if au.looks:
        tail.append(f"Мест, где WMS открывает файл или запускает программу этого компьютера (не сеть, посмотрите строки «ВНИМАНИЕ»): "
                    f"{len(sorted(set(au.looks)))}")
    notes = sorted(set(au.notes))
    body = au.finds + ([""] if au.finds else []) + sorted(set(au.looks)) + ([""] if au.looks else [])
    if notes:
        body += ["Для сведения (не сеть):"] + notes
    text = "\n".join(head + body + [""] + tail) + "\n"
    if a.out:
        open(a.out, "w", encoding="utf-8").write(text)
    sys.stdout.write(text)
    return 1 if au.finds else 0


if __name__ == "__main__":
    sys.exit(main())
