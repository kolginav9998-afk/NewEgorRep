"""What LibreOffice really drew: a minimal reader of the drawing operators of a PDF page exported by LibreOffice (no
third-party modules). Used by tests/run_m6.py to count the separator lines of «Заказы» as rendered.

scan(path) → (lines, rows_y)
  lines  — {y: share of the table width} of the thick horizontal lines drawn in the separator colour (y from the top of
           the page, the same units as rows_y); the table width is the width of the grey header row
  rows_y — the baselines (y from the top) of the texts of the first column below the header row
"""
import re
import zlib

SEPARATOR_RGB = (0x1F / 255, 0x38 / 255, 0x64 / 255)       # build_ods.order_separators: TopBorder colour 1F3864
HEADER_RGB = (0xE7 / 255, 0xE6 / 255, 0xE6 / 255)          # build_ods.header: CellBackColor E7E6E6
TOKEN = re.compile(rb"\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f\s]*>|\[|\]|/[^\s/\[\]()<>]+|[-+]?(?:\d+\.?\d*|\.\d+)|[A-Za-z'\"*]+")


def streams(path):
    data = open(path, "rb").read()
    out = []
    for m in re.finditer(rb"<<(.*?)>>\s*stream\r?\n", data, re.S):
        start = m.end()
        end = data.find(b"endstream", start)
        raw = data[start:end]
        if b"FlateDecode" in m.group(1):
            try:
                raw = zlib.decompressobj().decompress(raw)
            except zlib.error:
                continue
        out.append(raw)
    return out


def page_height(path):
    m = re.search(rb"/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]", open(path, "rb").read())
    return float(m.group(4)) if m else 842.0


def mul(a, b):
    """3x2 matrices [a b c d e f] of PDF: a·b"""
    return [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
            a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]]


def apply(m, x, y):
    return m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]


def close(c, want, eps=0.01):
    return c is not None and all(abs(a - b) <= eps for a, b in zip(c, want))


def scan(path):
    h = page_height(path)
    segs, rects, texts = [], [], []
    for content in streams(path):
        # page contents are text operators; fonts and images are binary
        if not content or b"\x00" in content or sum(1 for ch in content[:4000] if ch > 127) > len(content[:4000]) // 50:
            continue
        if b" Tf" not in content and b" re" not in content:
            continue
        ctm, stroke, fill, width = [1, 0, 0, 1, 0, 0], None, None, 1.0
        stack, ops, path_pts, cur = [], [], [], None
        for tok in TOKEN.findall(content):
            if tok[:1] in b"(<[]/" or tok[:1].isdigit() or tok[:1] in b"-+.":
                if tok[:1] in b"-+." or tok[:1].isdigit():
                    ops.append(float(tok))
                continue
            op = tok.decode("latin-1")
            if op == "q":
                stack.append((ctm[:], stroke, fill, width))
            elif op == "Q" and stack:
                ctm, stroke, fill, width = stack.pop()
            elif op == "cm" and len(ops) >= 6:
                ctm = mul(ops[-6:], ctm)
            elif op == "RG" and len(ops) >= 3:
                stroke = tuple(ops[-3:])
            elif op == "rg" and len(ops) >= 3:
                fill = tuple(ops[-3:])
            elif op == "w" and ops:
                width = ops[-1] * (abs(ctm[3]) or 1)
            elif op == "m" and len(ops) >= 2:
                cur = apply(ctm, ops[-2], ops[-1])
                path_pts = [cur]
            elif op == "l" and len(ops) >= 2:
                path_pts.append(apply(ctm, ops[-2], ops[-1]))
            elif op == "S":
                for a, b in zip(path_pts, path_pts[1:]):
                    segs.append((a, b, width, stroke))
                path_pts = []
            elif op == "re" and len(ops) >= 4:
                x, y, w, hh = ops[-4:]
                p0, p1 = apply(ctm, x, y), apply(ctm, x + w, y + hh)
                rects.append((min(p0[0], p1[0]), min(p0[1], p1[1]), abs(p1[0] - p0[0]), abs(p1[1] - p0[1]), fill))
            elif op in ("Td", "TD") and len(ops) >= 2:
                texts.append(apply(ctm, ops[-2], ops[-1]))
            elif op == "Tm" and len(ops) >= 6:
                texts.append(apply(ctm, ops[-2], ops[-1]))
            elif op in ("n", "f", "f*", "B", "b"):
                path_pts = []
            ops = []
    heads = [r for r in rects if close(r[4], HEADER_RGB) and r[2] > 50]
    if not heads:
        return {}, []
    hx, hy, hw, hh = max(heads, key=lambda r: r[2])[:4]
    lines = {}
    for (a, b, w, col) in segs:
        # the separator colour is used by nothing else on the sheet (the page may be scaled: the width is not compared)
        if close(col, SEPARATOR_RGB) and abs(a[1] - b[1]) < 0.01:
            y = round(h - a[1], 1)
            lines[y] = lines.get(y, 0) + abs(b[0] - a[0]) / hw
    lines = {y: round(v, 3) for y, v in lines.items()}
    xs = [x for x, y in texts if y < hy - 0.5]
    rows_y = []
    if xs:
        x0 = min(xs)
        rows_y = sorted(set(round(h - y, 1) for x, y in texts if y < hy - 0.5 and x < x0 + 8))
    return lines, rows_y
