#!/usr/bin/env python3
"""WMS_LABELS — адаптер прямой печати на термопринтер этикеток с языком TSPL (TSC, Xprinter, Gprinter и совместимые).

    python3 tspl_labels.py --snapshot SNAPSHOT_DIR --ei 201,206x2 [--width 58 --height 40 --gap 2] [--out labels.prn]
    python3 tspl_labels.py --export WMS_EXPORT_DIR --ei 201 --device /dev/usb/lp0

Основной путь печати этикеток — книга WMS_LABELS.ods через системный принтер (CUPS, обычное окно печати). Этот скрипт —
для принтера, который печатает только командами TSPL: одна этикетка на ЕИ (крупно ЕИ, наименование, артикул, место,
штрихкод Code 128 с номером ЕИ). Текст — в кодировке Windows-1251 (CODEPAGE 1251); шрифт с кириллицей у принтера должен
быть (у большинства — «TSS24.BF2» или загруженный TTF: параметр --font). Размеры — мм, плотность — 203 dpi (8 точек на
мм, параметр --dpmm). Снимок WMS проверяется по manifest (SHA-256) перед печатью. Проверка на настоящем принтере —
пакет проверки ПК (M5): команды печати тестовой этикетки там.
Код выхода: 0 — команды записаны, 2 — ошибка (ЕИ нет в снимке, снимок не цел, устройство недоступно).
"""
import argparse
import csv
import hashlib
import io
import os
import re
import sys


def read_text(p):
    raw = open(p, "rb").read()
    return raw.decode("utf-8-sig")


def sha256(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


def snapshot_stock(snap):
    man = {r[1]: r[4] for r in csv.reader(io.StringIO(read_text(os.path.join(snap, "manifest.csv"))), delimiter=";") if r and r[0] == "file"}
    if man.get("stock.csv") != sha256(os.path.join(snap, "stock.csv")):
        raise ValueError("stock.csv снимка изменён после экспорта (SHA-256 не совпадает с manifest.csv)")
    rows = list(csv.reader(io.StringIO(read_text(os.path.join(snap, "stock.csv"))), delimiter=";"))
    head = rows[0]
    return {r[head.index("ei")]: dict(zip(head, r)) for r in rows[1:] if r and r[head.index("ei")]}


def norm_ei(s):
    t = s.strip().upper().replace("ЕИ", "").replace("EI", "").lstrip("-")
    if not t.isdigit() or not 1 <= int(t) <= 99999999:
        raise ValueError(f"«{s}» — не номер ЕИ")
    return f"ЕИ-{int(t):08d}"


def parse_list(spec):
    """«201,206x2» → [(ЕИ-00000201, 1), (ЕИ-00000206, 2)]"""
    out = []
    for part in [p for p in re.split(r"[,\s;]+", spec) if p]:
        m = re.fullmatch(r"(.+?)(?:[xх*](\d+))?", part, flags=re.I)
        out.append((norm_ei(m.group(1)), int(m.group(2) or 1)))
    return out


def q(s):
    """a text of a TSPL command: quotes doubled out, one line"""
    return str(s).replace('"', "'").replace("\n", " ")[:60]


def label(card, copies, w, h, gap, dpmm, font):
    """the TSPL commands of one label"""
    ei = card["ei"]
    digits = ei[3:]
    x0 = 2 * dpmm
    lines = [f"SIZE {w} mm,{h} mm", f"GAP {gap} mm,0 mm", "DIRECTION 1", "CODEPAGE 1251", "CLS",
             f'TEXT {x0},{2 * dpmm},"{font}",0,2,2,"{q(ei)}"',
             f'TEXT {x0},{9 * dpmm},"{font}",0,1,1,"{q(card.get("name", ""))[:32]}"']
    if card.get("article"):
        lines.append(f'TEXT {x0},{14 * dpmm},"{font}",0,1,1,"Арт. {q(card["article"])[:28]}"')
    if card.get("place"):
        lines.append(f'TEXT {x0},{19 * dpmm},"{font}",0,1,1,"Место: {q(card["place"])[:24]}"')
    lines += [f'BARCODE {x0},{int((h - 16) * dpmm)},"128",{10 * dpmm},1,0,2,4,"{digits}"', f"PRINT 1,{copies}"]
    return "\r\n".join(lines) + "\r\n"


def main(argv=None):
    ap = argparse.ArgumentParser(description="этикетки ЕИ командами TSPL (термопринтер)")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--snapshot")
    g.add_argument("--export")
    ap.add_argument("--ei", required=True, help="ЕИ через запятую; «206x2» — две копии")
    ap.add_argument("--width", type=float, default=58)
    ap.add_argument("--height", type=float, default=40)
    ap.add_argument("--gap", type=float, default=2)
    ap.add_argument("--dpmm", type=int, default=8)
    ap.add_argument("--font", default="TSS24.BF2")
    o = ap.add_mutually_exclusive_group(required=True)
    o.add_argument("--out")
    o.add_argument("--device")
    a = ap.parse_args(argv)
    try:
        snap = a.snapshot
        if a.export:
            snaps = sorted(d for d in os.listdir(a.export) if d.startswith("WMS_SNAPSHOT_"))
            if not snaps:
                raise ValueError(f"в {a.export} нет снимков WMS")
            snap = os.path.join(a.export, snaps[-1])
        stock = snapshot_stock(snap)
        want = parse_list(a.ei)
        missing = [e for e, _ in want if e not in stock]
        if missing:
            raise ValueError(f"нет в снимке: {', '.join(missing)}")
        data = "".join(label(stock[e], n, a.width, a.height, a.gap, a.dpmm, a.font) for e, n in want).encode("cp1251", errors="replace")
        with open(a.out or a.device, "wb") as f:
            f.write(data)
    except (OSError, ValueError) as e:
        print(f"ОШИБКА: {e}", file=sys.stderr)
        return 2
    print(f"OK: этикеток {sum(n for _, n in want)} ({len(want)} ЕИ) → {a.out or a.device}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
