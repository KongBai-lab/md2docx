# -*- coding: utf-8 -*-
"""生成「墨式」封面 PNG：纯 Python（fontTools 字形轮廓 + 扫描线填充 + zlib PNG）。"""
import math
import os
import struct
import sys
import zlib

_DSH_HOME = os.environ.get("DSH_HOME") or ""
if _DSH_HOME and os.path.isdir(os.path.join(_DSH_HOME, "tools")):
    HERE = os.path.join(_DSH_HOME, "tools")
else:
    try:
        _SCRIPT = os.path.abspath(__file__)
    except NameError:
        _SCRIPT = os.path.abspath(sys.argv[0])
    HERE = os.path.dirname(_SCRIPT)
_PYLIBS = os.path.join(os.path.dirname(HERE), "pylibs")
if os.path.isdir(_PYLIBS):
    sys.path.insert(0, _PYLIBS)

from fontTools.ttLib import TTFont
from fontTools.pens.recordingPen import RecordingPen

W, H = 1240, 1754
SS = 2

_FONT_ALTS = {
    "bold": [
        "/system/fonts/FZHeiT-SC-Bold.ttf",
        "/system/fonts/HarmonyOS_Sans_Bold.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/System/Library/Fonts/PingFang.ttc",
        "C:/Windows/Fonts/msyhbd.ttc",
        "C:/Windows/Fonts/simhei.ttf",
    ],
    "regular": [
        "/system/fonts/FZHeiT-SC-Regular.ttf",
        "/system/fonts/HarmonyOS_Sans.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/PingFang.ttc",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simsun.ttc",
    ],
    "math": [
        "/system/fonts/DejaVuMathTeXGyre.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuMathTeXGyre.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/STIXTwoMath.otf",
        "C:/Windows/Fonts/cambria.ttc",
        "C:/Windows/Fonts/CambriaMath.ttf",
    ],
}


def _pick_font(kind):
    for path in _FONT_ALTS[kind]:
        if os.path.isfile(path):
            return path
    raise SystemExit("找不到可用的字体（%s）。请安装中文字体，或修改 make_cover.py 里的 _FONT_ALTS。" % kind)

BG = (247, 244, 238)
INK = (27, 27, 31)
SUB = (58, 58, 66)
MUTED = (122, 117, 108)
LINE = (185, 178, 166)

CW, CH = W * SS, H * SS
_canvas = bytearray(CW * CH * 3)


def fill_rect(x0, y0, x1, y1, color):
    x0 = max(0, int(x0)); x1 = min(CW, int(x1))
    y0 = max(0, int(y0)); y1 = min(CH, int(y1))
    if x1 <= x0 or y1 <= y0:
        return
    row = bytes(color) * (x1 - x0)
    for y in range(y0, y1):
        off = (y * CW + x0) * 3
        _canvas[off:off + len(row)] = row


def fill_polygons(polygons, color):
    edges = []
    for poly in polygons:
        n = len(poly)
        for i in range(n):
            x0, y0 = poly[i]
            x1, y1 = poly[(i + 1) % n]
            if y0 == y1:
                continue
            edges.append((x0, y0, x1, y1))
    if not edges:
        return
    ymin = max(0, int(math.floor(min(min(e[1], e[3]) for e in edges))))
    ymax = min(CH - 1, int(math.ceil(max(max(e[1], e[3]) for e in edges))))
    for y in range(ymin, ymax + 1):
        yc = y + 0.5
        xs = []
        for (x0, y0, x1, y1) in edges:
            if (y0 <= yc < y1) or (y1 <= yc < y0):
                xs.append(x0 + (yc - y0) * (x1 - x0) / (y1 - y0))
        if len(xs) < 2:
            continue
        xs.sort()
        for i in range(0, len(xs) - 1, 2):
            xa = max(0, int(math.ceil(xs[i])))
            xb = min(CW - 1, int(math.floor(xs[i + 1])))
            for x in range(xa, xb + 1):
                off = (y * CW + x) * 3
                _canvas[off:off + 3] = bytes(color)


def quadratic(p0, c, p1, steps=24):
    pts = []
    for i in range(1, steps + 1):
        t = i / steps
        mt = 1 - t
        pts.append((mt * mt * p0[0] + 2 * mt * t * c[0] + t * t * p1[0],
                    mt * mt * p0[1] + 2 * mt * t * c[1] + t * t * p1[1]))
    return pts


def glyph_polygons(font, glyph_set, glyph_name, size, origin_x, baseline_y):
    """返回以字体单位定义的闭合多边形（像素坐标）。"""
    pen = RecordingPen()
    glyph_set[glyph_name].draw(pen)
    scale = size / font["head"].unitsPerEm
    contours = []
    cur = []
    for op, args in pen.value:
        if op == "moveTo":
            if len(cur) > 2:
                contours.append(cur)
            cur = [args[0]]
        elif op == "lineTo":
            cur.append(args[0])
        elif op == "qCurveTo":
            pts = list(args)
            closed = pts and pts[-1] is None
            if closed:
                pts = pts[:-1]
            prev = cur[-1] if cur else (0, 0)
            for i, ctrl in enumerate(pts[:-1]):
                if i + 1 < len(pts) - 1:
                    nxt = ((ctrl[0] + pts[i + 1][0]) / 2.0,
                           (ctrl[1] + pts[i + 1][1]) / 2.0)
                else:
                    nxt = pts[-1] if not closed else cur[0]
                cur.extend(quadratic(prev, ctrl, nxt))
                prev = nxt
        elif op == "curveTo":
            prev = cur[-1] if cur else (0, 0)
            c1, c2, p1 = args[-3], args[-2], args[-1]
            for i in range(1, 25):
                t = i / 24
                mt = 1 - t
                cur.append((
                    mt ** 3 * prev[0] + 3 * mt * mt * t * c1[0] + 3 * mt * t * t * c2[0] + t ** 3 * p1[0],
                    mt ** 3 * prev[1] + 3 * mt * mt * t * c1[1] + 3 * mt * t * t * c2[1] + t ** 3 * p1[1],
                ))
            prev = p1
        elif op == "closePath":
            if len(cur) > 2:
                contours.append(cur)
            cur = []
    if len(cur) > 2:
        contours.append(cur)
    result = []
    for contour in contours:
        result.append([(origin_x + x * scale, baseline_y - y * scale)
                       for (x, y) in contour])
    return result


_FONTS = {}


def _font(path):
    if path not in _FONTS:
        f = TTFont(path, fontNumber=0)
        _FONTS[path] = (f, f.getGlyphSet(), f.getBestCmap())
    return _FONTS[path]


def text_width(text, font_path, size):
    font, glyphs, cmap = _font(font_path)
    total = 0.0
    for ch in text:
        name = cmap.get(ord(ch))
        if name is None:
            total += size * 0.5
            continue
        total += glyphs[name].width * size / font["head"].unitsPerEm
    return total


def draw_text(text, font_path, size, x, baseline_y, color, align="left"):
    font, glyphs, cmap = _font(font_path)
    upm = font["head"].unitsPerEm
    if align == "center":
        x = x - text_width(text, font_path, size) / 2.0
    cursor = x
    for ch in text:
        name = cmap.get(ord(ch))
        if name is None:
            cursor += size * 0.5
            continue
        polys = glyph_polygons(font, glyphs, name, size, cursor, baseline_y)
        fill_polygons(polys, color)
        cursor += glyphs[name].width * size / upm
    return cursor


def downsample():
    out = bytearray(W * H * 3)
    for y in range(H):
        for x in range(W):
            r = g = b = 0
            base = (y * SS * CW + x * SS) * 3
            for dy in range(SS):
                off = base + dy * CW * 3
                for dx in range(SS):
                    p = off + dx * 3
                    r += _canvas[p]; g += _canvas[p + 1]; b += _canvas[p + 2]
            n = SS * SS
            q = (y * W + x) * 3
            out[q] = r // n; out[q + 1] = g // n; out[q + 2] = b // n
    return out


def write_png(path, width, height, rgb):
    raw = b"".join(b"\x00" + rgb[y * width * 3:(y + 1) * width * 3] for y in range(height))

    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xffffffff)

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(raw, 9))
    png += chunk(b"IEND", b"")
    with open(path, "wb") as fh:
        fh.write(png)


def main():
    font_bold = _pick_font("bold")
    font_reg = _pick_font("regular")
    font_math = _pick_font("math")
    fill_rect(0, 0, CW, CH, BG)
    # 顶部细线
    fill_rect((CW - 440) // 2, 236, (CW + 440) // 2, 240, LINE)
    draw_text("MOSHI  ·  md2docx", font_reg, 40, CW // 2, 300, MUTED, align="center")
    # 主标题
    draw_text("墨式", font_bold, 620, CW // 2, 1060, INK, align="center")
    # 标题下细线
    fill_rect((CW - 720) // 2, 1180, (CW + 720) // 2, 1186, LINE)
    # 副标题
    draw_text("Markdown + LaTeX 文档生成工具", font_reg, 78, CW // 2, 1400, SUB, align="center")
    draw_text("Editable WPS / Word documents with native formulas",
              font_reg, 42, CW // 2, 1500, MUTED, align="center")
    # 公式符号行
    draw_text("∑   ∫   ∂   ∇   α   β", font_math, 96, CW // 2, 1820, MUTED, align="center")
    # 简介小字
    draw_text("从 Markdown 与 LaTeX 出发，一键生成可编辑的 WPS / Word 文档",
              font_reg, 42, CW // 2, 2150, SUB, align="center")
    draw_text("原生公式对象 · 自动排版 · 图片与超链接",
              font_reg, 42, CW // 2, 2250, MUTED, align="center")
    # 底部
    fill_rect((CW - 300) // 2, 3030, (CW + 300) // 2, 3034, LINE)
    draw_text("v1.1.0   ·   2026", font_reg, 38, CW // 2, 3120, MUTED, align="center")

    rgb = downsample()
    out = os.path.join(HERE, "out", "moshi-cover.png")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    write_png(out, W, H, rgb)
    print("封面已生成:", out, os.path.getsize(out), "bytes")


if __name__ == "__main__":
    raise SystemExit(main())
