# -*- coding: utf-8 -*-
"""pdf_blocks.py — 提取 PDF 的「行级文本块」：y 坐标 + 字号 + 文本。

doc_read 只输出纯文本（无位置/字号），而 PDF→Word 需要按字号判断标题层级、
按 y 合并段落。本模块复用 doc_read 的字体/CMap 解码，补上位置信息。
"""
import os
import re
import runpy

DOC_READ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "doc_read.py")

_MOD = None


def _m():
    global _MOD
    if _MOD is None:
        _MOD = runpy.run_path(DOC_READ, run_name="docread_lib")
    return _MOD


_NUM = rb"(-?\d*\.\d+|-?\d+)"
_TF = re.compile(rb"/([A-Za-z0-9]+)\s+" + _NUM + rb"\s+$")
_HEXSTR = re.compile(rb"<([0-9A-Fa-f\s]+)>")
_LIT = re.compile(rb"\(((?:[^()\\]|\\.)*)\)\s*$")
_TJARR = re.compile(rb"\[(.*)\]\s*$", re.S)


def _events(buf):
    ev = []
    for kw in (b"Tf", b"Tm", b"Tj", b"TJ"):
        p = 0
        while True:
            i = buf.find(kw, p)
            if i < 0:
                break
            p = i + len(kw)
            nxt = buf[i + len(kw):i + len(kw) + 1]
            if nxt and (nxt.isalnum() or nxt == 0x2a):
                continue
            # Tj/TJ 前面是 > 或 ]，不要求空白；Tf/Tm 要求前面是空白
            if kw not in (b"Tj", b"TJ") and (i < 1 or buf[i - 1] not in b" \t\r\n"):
                continue
            ev.append((i, kw))
    ev.sort(key=lambda x: x[0])
    return ev


def stream_blocks(buf, get_cmap, get_enc=None, y_gap=5.0):
    """返回 [(y, size, text)]，按内容流顺序。"""
    m = _m()
    cid_text = m["cid_text"]
    _lit_unescape = m["_lit_unescape"]
    _decode_lit = m["_decode_lit"]

    blocks, cur = [], []
    cur_y, cur_size, font, size = None, None, None, 10.0

    def flush():
        if cur:
            t = "".join(cur).strip()
            if t:
                blocks.append((round(cur_y or 0, 1), round(cur_size or size, 1), t))
            cur.clear()

    for pos, kw in _events(buf):
        seg = buf[max(0, pos - 240):pos]
        if kw == b"Tf":
            mm = _TF.search(seg)
            if mm:
                font = mm.group(1).decode("latin-1")
                size = float(mm.group(2))
        elif kw == b"Tm":
            nums = re.findall(rb"-?\d*\.\d+|-?\d+", seg)
            if len(nums) >= 6:
                six = nums[-6:]
                a = abs(float(six[0]))
                f = float(six[5])
                if cur_y is not None and abs(f - cur_y) > y_gap:
                    flush()
                cur_y = f
                cur_size = round(size * a, 1) if a > 0.001 else size
        elif kw == b"Tj":
            h = re.search(rb"<([0-9A-Fa-f\s]+)>\s*$", seg)
            if h:
                cur.append(cid_text(h.group(1), get_cmap(font)))
            else:
                lit = _LIT.search(seg)
                if lit:
                    cur.append(_decode_lit(_lit_unescape(lit.group(1)),
                                          get_enc(font) if get_enc else None))
        elif kw == b"TJ":
            arr = _TJARR.search(seg)
            if arr:
                for h in _HEXSTR.finditer(arr.group(1)):
                    cur.append(cid_text(h.group(1), get_cmap(font)))
                if not _HEXSTR.search(arr.group(1)):
                    for lit in re.finditer(rb"\(((?:[^()\\]|\\.)*)\)", arr.group(1)):
                        cur.append(_decode_lit(_lit_unescape(lit.group(1)),
                                              get_enc(font) if get_enc else None))
    flush()
    return blocks


def extract_blocks(path):
    """按页返回 {"index", "blocks": [(y,size,text)]}。"""
    m = _m()
    data = open(path, "rb").read()
    objs = m["pdf_objects"](data)
    refs = m["_page_refs"](objs)

    font_name = {}
    for body in objs.values():
        if b"/Font" not in body:
            continue
        for blk in m["FONTDICT_RE"].findall(body):
            for mm in m["FONTREF_RE"].finditer(blk):
                font_name.setdefault(mm.group(1).decode("latin1"), int(mm.group(2)))

    cmaps, cmap_by_ref, fallback = {}, {}, {}
    for num, body in objs.items():
        mm = m["TOUNI_RE"].search(body)
        if not mm:
            continue
        ref = int(mm.group(1))
        cm = cmap_by_ref.get(ref)
        if cm is None and ref in objs:
            s = m["_inflate"](objs[ref])
            if s is not None:
                cm = m["parse_cmap"](s)
                cmap_by_ref[ref] = cm
        if cm:
            cmaps[num] = cm
            if len(cm) > len(fallback):
                fallback = cm

    def get_cmap(name):
        on = font_name.get(name)
        if on is not None and on in cmaps:
            return cmaps[on]
        return fallback

    enc_cache = {}

    def font_enc(on):
        if on not in enc_cache:
            enc_cache[on] = m["_font_encoding"](objs, on)
        return enc_cache[on]

    def page_enc(num):
        pf = m["_page_fonts"](objs, num) or font_name
        cache = {}

        def get_enc(name):
            if name in cache:
                return cache[name]
            on = pf.get(name) or font_name.get(name)
            cache[name] = font_enc(on) if on is not None else None
            return cache[name]

        return get_enc

    pages = []
    for idx, num in enumerate(refs, 1):
        content = m["_page_content"](objs, num)
        pages.append({"index": idx, "blocks": stream_blocks(content, get_cmap, page_enc(num))})
    return pages


def _demo():
    import sys
    if len(sys.argv) < 2:
        raise SystemExit("用法: python pdf_blocks.py input.pdf   # 打印每页行级块")
    for p in extract_blocks(sys.argv[1])[:4]:
        print("== p%d ==" % p["index"])
        for y, sz, txt in p["blocks"]:
            print("  y=%-6s sz=%-5s %s" % (y, sz, txt[:60]))


if __name__ == "__main__":
    _demo()
