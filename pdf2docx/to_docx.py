# -*- coding: utf-8 -*-
"""to_docx.py — PDF → WPS/Word 可编辑 .docx（文本 + 内嵌图片）。

链路：
  pdf_read.read_pdf()   行级文本（y / 字号）+ 图片位置
  → pages_to_md()       按字号推标题层级、按 y 排段落，图片按页插入
  → md2docx.py          Markdown → 原生 docx（宋体小四 / OMML 公式）

不追求版式 1:1：PDF 是排版结果，Word 是结构文档。目标是「文字可编辑、层级合理」。
"""
import argparse
import os
import re
import runpy
import struct
import sys
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import pdf_skeleton  # noqa: E402
from pdf_read import read_pdf  # noqa: E402

# md2docx.py 与本目录同级（pdf2docx/../md2docx.py）；可用环境变量 MD2DOCX_PY 覆盖
MD2DOCX = os.environ.get("MD2DOCX_PY") or os.path.join(
    os.path.dirname(HERE), "md2docx.py")

_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")  # XML 1.0 非法控制符

_MIN_W, _MIN_H = 80, 60          # 小于此尺寸的多为图标 / 装饰
_MAX_FULL_W, _MAX_FULL_H = 2000, 2000   # 满版背景图不入正文


def _clean(text):
    """清掉 PDF 文本层夹带的控制字符（NUL 等），保证中间 .md 与产物 docx 都干净。"""
    return _CTRL.sub("", text)


# ───────────────────────────── 图片导出 ─────────────────────────────
def png_chunk(tag, data):
    return (struct.pack(">I", len(data)) + tag + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))


def write_png(path, w, h, pixels, channels=3):
    """最小 PNG 封装：8 bit，RGB(2) / Gray(0)，逐行 filter=0。"""
    color_type = 2 if channels == 3 else 0
    stride = w * channels
    raw = bytearray()
    for y in range(h):
        raw.append(0)
        raw += pixels[y * stride:(y + 1) * stride]
    ihdr = struct.pack(">IIBBBBB", w, h, 8, color_type, 0, 0, 0)
    with open(path, "wb") as f:
        f.write(b"\x89PNG\r\n\x1a\n" + png_chunk(b"IHDR", ihdr)
                + png_chunk(b"IDAT", zlib.compress(bytes(raw), 6))
                + png_chunk(b"IEND", b""))


def _paeth(a, b, c):
    p = a + b - c
    pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
    if pa <= pb and pa <= pc:
        return a
    return b if pb <= pc else c


def _unfilter(data, w, h, bpp):
    """还原 PNG 预测（Predictor ≥ 10）：每行 1 字节 filter + 行数据。"""
    stride = w * bpp
    out = bytearray(stride * h)
    prev = bytearray(stride)
    pos = 0
    for y in range(h):
        if pos >= len(data):
            break
        ft = data[pos]
        pos += 1
        row = bytearray(data[pos:pos + stride])
        pos += stride
        if len(row) < stride:
            row += bytearray(stride - len(row))
        if ft == 1:      # Sub
            for i in range(bpp, stride):
                row[i] = (row[i] + row[i - bpp]) & 0xFF
        elif ft == 2:    # Up
            for i in range(stride):
                row[i] = (row[i] + prev[i]) & 0xFF
        elif ft == 3:    # Average
            for i in range(stride):
                left = row[i - bpp] if i >= bpp else 0
                row[i] = (row[i] + ((left + prev[i]) >> 1)) & 0xFF
        elif ft == 4:    # Paeth
            for i in range(stride):
                left = row[i - bpp] if i >= bpp else 0
                ul = prev[i - bpp] if i >= bpp else 0
                row[i] = (row[i] + _paeth(left, prev[i], ul)) & 0xFF
        out[y * stride:(y + 1) * stride] = row
        prev = row
    return bytes(out)


def _num(body, key):
    m = re.search(rb"/" + key + rb"\s+(\d+)", body)
    return int(m.group(1)) if m else None


def _name(body, key):
    m = re.search(rb"/" + key + rb"\s*/([A-Za-z0-9]+)", body)
    return m.group(1) if m else b""


def _channels(body):
    """按 /ColorSpace 推通道数：DeviceGray→1、DeviceRGB→3、DeviceCMYK→4、ICCBased→N。"""
    cs = _name(body, b"ColorSpace")
    if cs in (b"DeviceRGB", b"CalRGB", b"RGB"):
        return 3
    if cs in (b"DeviceGray", b"CalGray", b"G"):
        return 1
    if cs == b"DeviceCMYK":
        return 4
    if cs == b"ICCBased":
        m = re.search(rb"/ICCBased\s+(\d+)\s+\d+\s+R", body)
        if m:
            return 1  # 解析目标 ICC 流不稳定，先按灰度兜底；正常测试图多为 DCT
    m = re.search(rb"/ColorSpace\s*\[\s*/[A-Za-z]+\s*/([A-Za-z0-9]+)", body)
    if m and m.group(1) == b"DeviceRGB":
        return 3
    return 0


def _png_from_flate(body, w, h):
    """FlateDecode 位图 → PNG。支持 DeviceRGB / DeviceGray 8bit（含 PNG 预测）。"""
    m = pdf_skeleton._mod()
    raw = m["_inflate"](body)
    if not raw:
        return None
    ch = _channels(body)
    if ch not in (1, 3):
        return None
    bpc = _num(body, b"BitsPerComponent") or 8
    if bpc != 8:
        return None
    stride = w * ch
    if not (len(raw) >= stride * h or b"/Predictor" in body):
        return None
    if b"/Predictor" in body:
        raw = _unfilter(raw, w, h, ch)
    if len(raw) < stride * h:
        return None
    return raw[:stride * h], ch


def export_images(pdf_path, out_dir):
    """按页导出 PDF 内嵌图片，返回 {页号: [文件路径]}。

    支持 JPEG(DCTDecode) 原样写出与 FlateDecode 8bit 位图重封装为 PNG；
    JBIG2 / JPEG2000 / CCITT 等扫描件编码跳过（不影响文字提取）。
    """
    m = pdf_skeleton._mod()
    with open(pdf_path, "rb") as f:
        objs = m["pdf_objects"](f.read())
    refs = m["_page_refs"](objs)
    os.makedirs(out_dir, exist_ok=True)
    result = {}
    for idx, num in enumerate(refs, 1):
        xobj = pdf_skeleton._page_xobjects(objs, num)
        if not xobj:
            continue
        files = []
        for name, on in xobj.items():
            body = objs.get(on, b"")
            if b"/Image" not in body[:400]:
                continue
            w, h = _num(body, b"Width"), _num(body, b"Height")
            if not w or not h:
                continue
            if w < _MIN_W or h < _MIN_H:                       # 图标 / 装饰
                continue
            if w >= _MAX_FULL_W and h >= _MAX_FULL_H:          # 满版背景
                continue
            safe = re.sub(r"[^\w.-]", "_", name) or "img"
            base = os.path.join(out_dir, "p%02d_%s" % (idx, safe))
            filt = _name(body, b"Filter")
            path = None
            if filt in (b"DCTDecode", b"DCT"):
                raw = m["_stream_body"](body)
                if raw:
                    with open(base + ".jpg", "wb") as f:
                        f.write(raw.rstrip(b"\r\n"))
                    path = base + ".jpg"
            elif filt in (b"FlateDecode", b"Fl"):
                got = _png_from_flate(body, w, h)
                if got:
                    pixels, ch = got
                    write_png(base + ".png", w, h, pixels, channels=ch)
                    path = base + ".png"
            if path:
                files.append(path)
                if len(files) >= 6:                            # 单页最多 6 张，防爆版
                    break
        if files:
            result[idx] = files
    return result


# ───────────────────────────── PDF → Markdown ─────────────────────────────
def pages_to_md(pages, img_map=None, base_dir=None):
    """每页：字号最大且较短的块作页标题（##），同字号短文作小标题（###），其余成段。

    img_map 里的图片路径是绝对路径；base_dir 给出时改写成相对路径，
    这样中间 .md 可以搬动、也不会把本机目录写进产物。
    """
    lines = []
    for p in pages:
        # 先清洗再排序：PDF 文本层可能夹带 NUL 等控制字符
        texts = sorted((dict(t, text=_clean(t["text"])) for t in p["texts"]),
                       key=lambda t: t["y"])
        texts = [t for t in texts if t["text"].strip()]
        if not texts:
            continue
        max_size = max(t["size"] for t in texts)
        head = None
        for t in texts:
            if t["size"] >= max_size * 0.95 and len(t["text"]) <= 40:
                head = t
                break
        if head is not None:
            lines.append("## " + head["text"].strip())
        for t in texts:
            if t is head:
                continue
            if t["size"] >= max_size * 0.85 and len(t["text"]) <= 28:
                lines.append("### " + t["text"].strip())
            else:
                lines.append(t["text"])
        for path in (img_map or {}).get(p["index"], []):
            ref = path
            if base_dir:
                try:
                    ref = os.path.relpath(path, base_dir).replace(os.sep, "/")
                except ValueError:      # 跨盘符（Windows）
                    ref = path
            lines.append("![图](%s)" % ref)
    return "\n\n".join(lines)


def to_docx(pdf_path, docx_path, with_images=True):
    data = read_pdf(pdf_path)
    img_map = None
    if with_images:
        img_dir = os.path.splitext(docx_path)[0] + "_images"
        try:
            img_map = export_images(pdf_path, img_dir)
        except Exception:
            img_map = None
    md = pages_to_md(data["pages"], img_map, base_dir=os.path.dirname(os.path.abspath(docx_path)))
    md_path = os.path.splitext(docx_path)[0] + ".md"
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md)
    title = None
    for p in data["pages"]:
        if p["texts"]:
            title = sorted(p["texts"], key=lambda t: t["y"])[0]["text"].strip()
            if title:
                break
    if not os.path.isfile(MD2DOCX):
        raise SystemExit("找不到 md2docx.py：%s（可用环境变量 MD2DOCX_PY 指定）" % MD2DOCX)
    sys.argv = ["md2docx.py", md_path, "-o", docx_path]
    if title:
        sys.argv += ["--title", title]
    runpy.run_path(MD2DOCX, run_name="md2docx_lib")["main"]()
    return docx_path


def main(argv=None):
    ap = argparse.ArgumentParser(description="PDF -> WPS/Word docx")
    ap.add_argument("pdf")
    ap.add_argument("-o", "--out", help="输出 .docx（默认与输入同名）")
    ap.add_argument("--no-images", action="store_true", help="不导出内嵌图片")
    args = ap.parse_args(argv)
    out = args.out or os.path.splitext(args.pdf)[0] + ".docx"
    to_docx(args.pdf, out, with_images=not args.no_images)
    print("saved", out, os.path.getsize(out), "bytes")
    return out


if __name__ == "__main__":
    main()
