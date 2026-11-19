# -*- coding: utf-8 -*-
"""pdf_skeleton.py — 从 PDF 内容流提取「版式骨架」。

不依赖 PyMuPDF/poppler：复用 doc_read 的对象解析 + 内容流解压，
再用轻量操作符扫描拿到 文本块(x,y,size,color) / 矩形(x,y,w,h,color)。

用途：把一份分享来的 PDF 当模板，抽出配色/字号/位置规律，供 PPT 重建。
"""
import json
import math
import os
import re
import runpy
import sys

DOC_READ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "doc_read.py")

_NUM = rb"(-?\d*\.\d+|-?\d+)"
_OP = re.compile(
    rb"(?:/([A-Za-z0-9]+)\s+" + _NUM + rb"\s+Tf)"                                              # g0,g1
    rb"|(?:" + _NUM + rb"\s+" + _NUM + rb"\s+" + _NUM + rb"\s+" + _NUM + rb"\s+" + _NUM + rb"\s+" + _NUM + rb"\s+Tm)"  # g2..g7
    rb"|(?:" + _NUM + rb"\s+" + _NUM + rb"\s+" + _NUM + rb"\s+" + _NUM + rb"\s+re)"     # g8..g11
    rb"|(?:" + _NUM + rb"\s+" + _NUM + rb"\s+" + _NUM + rb"\s+(rg|RG))"                  # g12..g15
)

_MOD = None


def _mod():
    global _MOD
    if _MOD is None:
        _MOD = runpy.run_path(DOC_READ, run_name="docread_lib")
    return _MOD


def _hex(rgb):
    return "%02X%02X%02X" % rgb


def _clamp01(v):
    return max(0.0, min(1.0, v))


_ML = re.compile(rb"(-?\d*\.\d+|-?\d+)\s+(-?\d*\.\d+|-?\d+)\s+(m|l)\b")
_END = re.compile(rb"(?:f\*?|S|B)\b")


def parse_paths(content):
    """把矢量路径（m/l 点）聚成外接矩形，补 re 抓不到的图形（图表的柱/线）。"""
    events = []
    for m in _ML.finditer(content):
        events.append((m.start(), 0, (float(m.group(1)), float(m.group(2)))))
    for m in _END.finditer(content):
        events.append((m.start(), 1, None))
    events.sort(key=lambda e: e[0])
    cur, out = [], []
    for _pos, typ, val in events:
        if typ == 0:
            cur.append(val)
        else:
            if len(cur) >= 2:
                xs = [p[0] for p in cur]
                ys = [p[1] for p in cur]
                out.append({"x": round(min(xs), 1), "y": round(min(ys), 1),
                            "w": round(max(xs) - min(xs), 1),
                            "h": round(max(ys) - min(ys), 1)})
            cur = []
    return out


def parse_content(content):
    """返回 {"texts": [...], "rects": [...]}。坐标为 PDF 页面坐标（y 向上）。"""
    size, scale = 1.0, 1.0
    color = (0, 0, 0)
    texts, rects = [], []
    for mm in _OP.finditer(content):
        g = mm.groups()
        if g[1] is not None:                                    # Tf: /F1 28 Tf
            size = float(g[1])
        elif g[2] is not None and g[6] is not None and g[7] is not None:  # Tm
            a = float(g[2])
            e, f = float(g[6]), float(g[7])
            s = abs(a)
            if 0.01 < s < 20:
                scale = s
            texts.append({"x": round(e, 1), "y": round(f, 1),
                          "size": round(size * scale, 1), "color": _hex(color)})
        elif g[8] is not None and g[11] is not None:            # re: x y w h re
            x, y, w, h = (float(g[8]), float(g[9]), float(g[10]), float(g[11]))
            rects.append({"x": round(x, 1), "y": round(y, 1),
                          "w": round(w, 1), "h": round(h, 1), "color": _hex(color)})
        elif g[12] is not None and g[14] is not None:           # rg: r g b rg
            r, gg, b = (float(g[12]), float(g[13]), float(g[14]))
            color = (int(round(_clamp01(r) * 255)), int(round(_clamp01(gg) * 255)),
                     int(round(_clamp01(b) * 255)))
    return {"texts": texts, "rects": rects, "paths": parse_paths(content)}


_CM = re.compile(_NUM + rb"\s+" + _NUM + rb"\s+" + _NUM + rb"\s+" + _NUM + rb"\s+"
                 + _NUM + rb"\s+" + _NUM + rb"\s+cm")
_DO = re.compile(rb"/([A-Za-z0-9_]+)\s+Do")


def _page_xobjects(objs, num):
    body = objs[num]
    m = re.search(rb"/Resources\s+(\d+)\s+\d+\s+R", body)
    res = objs.get(int(m.group(1)), b"") if m else body
    mx = re.search(rb"/XObject\s*<<(.*?)>>", res, re.S)
    if not mx:
        return {}
    return {nm.decode("latin-1"): int(on)
            for nm, on in re.findall(rb"/([A-Za-z0-9_]+)\s+(\d+)\s+\d+\s+R", mx.group(1))}


def _xobject_dims(objs, xobj):
    dims = {}
    for name, on in xobj.items():
        body = objs.get(on, b"")
        mw = re.search(rb"/Width\s+(\d+)", body)
        mh = re.search(rb"/Height\s+(\d+)", body)
        if mw and mh:
            dims[name] = (int(mw.group(1)), int(mh.group(1)))
    return dims


def parse_images(content, dims):
    """按内容流顺序跟踪 cm，遇 Do 记录图片的页面坐标与尺寸。"""
    events = []
    for m in _CM.finditer(content):
        events.append((m.start(), 0, tuple(float(m.group(i)) for i in range(1, 7))))
    for m in _DO.finditer(content):
        events.append((m.start(), 1, m.group(1).decode("latin-1")))
    events.sort(key=lambda e: e[0])
    ctm = (1, 0, 0, 1, 0, 0)
    out = []
    for _pos, typ, val in events:
        if typ == 0:
            ctm = val
            continue
        wh = dims.get(val)
        if not wh:
            continue
        a, _b, _c, d, e, f = ctm
        # PDF 图片的 cm 把「单位正方形」映射到目标框：显示宽 = |a|，高 = |d|
        out.append({"name": val, "x": round(e, 1), "y": round(f, 1),
                    "w": round(abs(a), 1), "h": round(abs(d), 1)})
    return out


def extract(path):
    m = _mod()
    data = open(path, "rb").read()
    objs = m["pdf_objects"](data)
    refs = m["_page_refs"](objs)
    pages = []
    for idx, num in enumerate(refs, 1):
        content = m["_page_content"](objs, num)
        dims = _xobject_dims(objs, _page_xobjects(objs, num))
        pages.append({"index": idx, **parse_content(content),
                      "images": parse_images(content, dims)})
    return pages


def page_stats(page):
    texts = page["texts"]
    if not texts:
        return {}
    sizes = sorted({t["size"] for t in texts}, reverse=True)
    top = [t for t in texts if t["y"] > 300]
    big = max((t["size"] for t in top), default=0)
    titles = [t for t in top if t["size"] == big] if top else []
    if not titles:
        big = max(t["size"] for t in texts)
        titles = [t for t in texts if t["size"] == big]
    footer = [t for t in texts if t["y"] < 60]
    colors = {}
    for t in texts:
        colors[t["color"]] = colors.get(t["color"], 0) + 1
    for r in page["rects"]:
        colors[r["color"]] = colors.get(r["color"], 0) + 1
    return {
        "blocks": len(texts),
        "rects": len(page["rects"]),
        "images": len(page.get("images", [])),
        "sizes": sizes[:6],
        "title": titles[0] if titles else None,
        "title_pos": (titles[0]["x"], titles[0]["y"]) if titles else None,
        "footer": footer[0] if footer else None,
        "top_colors": sorted(colors.items(), key=lambda kv: -kv[1])[:8],
    }


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("-o", "--out")
    ap.add_argument("--pages", type=int, default=0)
    args, _ = ap.parse_known_args(argv)
    pages = extract(args.pdf)
    if args.pages:
        pages = pages[:args.pages]
    for p in pages:
        st = page_stats(p)
        print("p%02d blocks=%-3d rects=%-3d imgs=%-2d sizes=%s title@%s" % (
            p["index"], st.get("blocks", 0), st.get("rects", 0),
            st.get("images", 0), st.get("sizes"), st.get("title_pos")))
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(pages, f, ensure_ascii=False, indent=1)
        print("saved", args.out)
    return pages


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        raise SystemExit("用法: python pdf_skeleton.py input.pdf [-o skeleton.json]")
    main(sys.argv[1:])
