# -*- coding: utf-8 -*-
"""pdf_read.py — PDF 统一读取层（整合 pdf_blocks + pdf_skeleton）。

输出统一 IR：
  {"style": {...}, "pages": [{"role","texts","rects","images","charts","title"}]}

图表识别：先找「柱状图」——同底边、等宽、高度不同的矩形组。
数据优先取柱上方/下方的文本标签；没有标签就只给相对比例。
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)
import pdf_blocks   # noqa: E402
import pdf_skeleton  # noqa: E402


def _merge_texts(blocks, raw_texts):
    """把 pdf_blocks 的 (y,size,text) 与 pdf_skeleton 的 (x,y,size,color) 配对。"""
    out = []
    for y, sz, txt in blocks:
        cand = min(raw_texts, key=lambda r: abs(r["y"] - y) + abs(r["size"] - sz) * 0.4,
                   default=None)
        x = cand["x"] if cand else 0.0
        color = cand["color"] if cand else "000000"
        out.append({"x": x, "y": y, "size": sz, "color": color, "text": txt})
    return out


def _detect_role(page, idx, total):
    texts = page["texts"]
    if not texts:
        return "blank"
    if idx == 1:
        return "cover"
    if idx == total:
        return "end"
    joined = "".join(t["text"] for t in texts)
    sizes = [t["size"] for t in texts]
    mx = max(sizes)
    if "目录" in joined and len(joined) < 100:
        return "toc"
    if mx >= 30 and len(joined) < 50:
        return "section"
    return "content"


def _chart_box(rs, typ):
    x0 = min(r["x"] for r in rs)
    x1 = max(r["x"] + r["w"] for r in rs)
    y0 = min(r["y"] for r in rs)
    y1 = max(r["y"] + r["h"] for r in rs)
    return {"type": typ,
            "box": {"x": round(x0, 1), "y": round(y0, 1),
                    "w": round(x1 - x0, 1), "h": round(y1 - y0, 1)},
            "items": [{"x": round(r["x"], 1), "y": round(r["y"], 1),
                       "w": round(r["w"], 1), "h": round(r["h"], 1)} for r in rs],
            "confidence": "low"}


def _detect_vbar(shapes):
    """竖直柱状图：同底边、等宽、高度不同。"""
    groups = {}
    for r in shapes:
        groups.setdefault((round((r["y"] + r["h"]) / 10), round(r["w"] / 5)), []).append(r)
    out = []
    for rs in groups.values():
        if len(rs) < 3:
            continue
        ws = [r["w"] for r in rs]
        hs = [r["h"] for r in rs]
        if max(ws) - min(ws) > max(ws) * 0.35 or max(hs) - min(hs) < 8:
            continue
        out.append(_chart_box(rs, "vbar"))
    return out


def _detect_hbar(shapes):
    """横向条形图：同左边、等高、宽度不同（p9 的失业率图就是这种）。"""
    groups = {}
    for r in shapes:
        groups.setdefault((round(r["x"] / 10), round(r["h"] / 5)), []).append(r)
    out = []
    for rs in groups.values():
        if len(rs) < 2:
            continue
        hs = [r["h"] for r in rs]
        ws = [r["w"] for r in rs]
        if max(hs) - min(hs) > max(hs) * 0.35 or max(ws) - min(ws) < 8:
            continue
        out.append(_chart_box(rs, "hbar"))
    return out


def _detect_charts(page):
    shapes = [r for r in page.get("rects", []) + page.get("paths", [])
              if r["w"] > 6 and r["h"] > 6]
    if len(shapes) < 2:
        return []
    return _detect_vbar(shapes) + _detect_hbar(shapes)


def _attach_chart_labels(page, charts):
    """把图表区域内的文本挂到图上：柱下方短文本=类别，柱上方数字=数值。"""
    texts = page["texts"]
    for ch in charts:
        bx, by, bw, bh = (ch["box"][k] for k in ("x", "y", "w", "h"))
        inside = [t for t in texts
                  if bx - 10 <= t["x"] <= bx + bw + 10
                  and by - 60 <= t["y"] <= by + bh + 60]
        cats, vals = [], []
        for t in inside:
            s = t["text"].strip()
            if not s:
                continue
            if len(s) <= 12 and any(c.isdigit() for c in s):
                vals.append(s)
            elif len(s) <= 12:
                cats.append(s)
        ch["labels"] = {"categories": cats, "values": vals}
    return charts


def _compute_style(pages):
    colors, sizes = {}, {}
    for p in pages:
        for t in p["texts"]:
            colors[t["color"]] = colors.get(t["color"], 0) + 1
            sizes[t["size"]] = sizes.get(t["size"], 0) + 1
        for r in p["rects"]:
            colors[r["color"]] = colors.get(r["color"], 0) + 1
    top_colors = sorted(colors.items(), key=lambda kv: -kv[1])[:8]
    top_sizes = sorted(sizes.items(), key=lambda kv: -kv[1])[:6]
    return {
        "top_colors": top_colors,
        "top_sizes": top_sizes,
        "title_pt": top_sizes[0][0] if top_sizes else 24,
        "body_pt": sorted(sizes, reverse=True)[1] if len(sizes) > 1 else 16,
    }


def read_pdf(path):
    blocks_pages = pdf_blocks.extract_blocks(path)
    skel_pages = pdf_skeleton.extract(path)
    pages = []
    for b, s in zip(blocks_pages, skel_pages):
        page = {
            "index": b["index"],
            "texts": _merge_texts(b["blocks"], s["texts"]),
            "rects": s["rects"],
            "paths": s.get("paths", []),
            "images": s["images"],
        }
        page["role"] = _detect_role(page, page["index"], len(blocks_pages))
        for im in page["images"]:
            # 尺寸适中的内嵌图可能是位图图表，供后续导出 + 视觉识别
            im["chart_candidate"] = (200 <= im["w"] <= 700 and 100 <= im["h"] <= 400)
        if page["role"] in ("cover", "toc", "end"):
            page["charts"] = []
        else:
            page["charts"] = _attach_chart_labels(page, _detect_charts(page))
        if page["texts"]:
            head = sorted(page["texts"], key=lambda t: t["y"])[0]
            page["title"] = head
        pages.append(page)
    return {"style": _compute_style(pages), "pages": pages}


def summarize(data):
    lines = ["样式: 主色=%s 标题字号=%s 正文字号=%s" % (
        data["style"]["top_colors"][:3], data["style"]["title_pt"], data["style"]["body_pt"])]
    for p in data["pages"]:
        roles = p["role"]
        charts = p["charts"]
        lines.append("p%02d %-8s texts=%-3d rects=%-3d imgs=%-2d charts=%d" % (
            p["index"], roles, len(p["texts"]), len(p["rects"]),
            len(p["images"]), len(charts)))
        for ch in charts:
            lines.append("     chart %s box=%s labels=%s" % (
                ch["type"], ch["box"], ch.get("labels")))
    return "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("用法: python pdf_read.py input.pdf   # 打印每页版式摘要")
    print(summarize(read_pdf(sys.argv[1])))
