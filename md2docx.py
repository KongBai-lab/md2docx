# -*- coding: utf-8 -*-
"""md2docx: Markdown(含 LaTeX 公式) -> .docx，公式为原生 OMML 对象（WPS/Word 可编辑）。

用法:
    python3 md2docx.py input.md -o output.docx [--title "标题"]
"""
import argparse
import os
import re
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile
from xml.sax.saxutils import escape

__version__ = "1.3.2"

_DSH_HOME = os.environ.get("DSH_HOME") or ""
if _DSH_HOME and os.path.isdir(os.path.join(_DSH_HOME, "tools")):
    HERE = os.path.join(_DSH_HOME, "tools")
else:
    try:
        _SCRIPT = os.path.abspath(__file__)
    except NameError:  # 被 in-process exec 时没有 __file__
        _SCRIPT = os.path.abspath(sys.argv[0] if sys.argv else "md2docx.py")
    HERE = os.path.dirname(_SCRIPT)
_PYLIBS = os.path.join(os.path.dirname(HERE), "pylibs")
if os.path.isdir(_PYLIBS):
    sys.path.insert(0, _PYLIBS)

try:
    from latex2mathml.converter import convert as _latex2mathml
    from mathml2omml import convert as _mathml2omml
except Exception as exc:  # pragma: no cover
    raise SystemExit("缺少依赖: %s" % exc)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ---------------- 基础 XML 片段 ----------------
# XML 1.0 不允许的字符：C0 控制符（保留 \t \n \r）、孤立代理对、U+FFFE/U+FFFF。
# PDF 文本层常夹带 U+0000 等控制字符，原样写进 <w:t>/<dc:title> 会让 Word/WPS
# 解析失败（实测 54 个 NUL 导致 document.xml not well-formed）。所有文本出口都过 esc()。
_ILLEGAL_XML_TABLE = {
    c: None
    for c in (
        list(range(0x00, 0x09)) + [0x0B, 0x0C] + list(range(0x0E, 0x20))
        + list(range(0xD800, 0xE000)) + [0xFFFE, 0xFFFF]
    )
}


def strip_illegal_xml(text):
    """剔除 XML 1.0 非法字符，其余原样保留。"""
    return text.translate(_ILLEGAL_XML_TABLE)


def esc(text):
    return escape(strip_illegal_xml(text))


def rpr(bold=False, italic=False, code=False, color=None, sz=None, font=None):
    parts = []
    if font:
        parts.append(
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
            % (font, font, font, font)
        )
    if code:
        parts.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>')
        parts.append('<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>')
    if bold:
        parts.append("<w:b/>")
    if italic:
        parts.append("<w:i/>")
    if color:
        parts.append('<w:color w:val="%s"/>' % color)
    if sz:
        parts.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz))
    return "<w:rPr>" + "".join(parts) + "</w:rPr>" if parts else ""


def run(text, **kw):
    return '<w:r>' + rpr(**kw) + '<w:t xml:space="preserve">' + esc(text) + "</w:t></w:r>"


# ---------------- 图片 / 超链接 / 公式缓存 ----------------
_PX_EMU = 9525            # 96 DPI: 1 px = 9525 EMU
_MAX_IMAGE_W = 5600000    # ~5.9 in
_MAX_IMAGE_H = 7400000    # ~7.8 in
_MEDIA = []               # [(rId, media filename, bytes)]
_EXTERNAL_RELS = []       # [(rId, url)]
_REL_SEQ = [0]
_PIC_SEQ = [0]
_BASE_DIR = [os.getcwd()]
_OMML_CACHE = {}


def _reset_generation():
    _MEDIA.clear()
    _EXTERNAL_RELS.clear()
    _REL_SEQ[0] = 0
    _PIC_SEQ[0] = 0


def _next_rel_id():
    _REL_SEQ[0] += 1
    return "rId%d" % _REL_SEQ[0]


def _image_size(path):
    """返回图片像素尺寸 (w, h)；支持 PNG/JPEG/GIF。"""
    try:
        with open(path, "rb") as fh:
            head = fh.read(32)
            if head[:8] == b"\x89PNG\r\n\x1a\n":
                return struct.unpack(">II", head[16:24])
            if head[:6] in (b"GIF87a", b"GIF89a"):
                return struct.unpack("<HH", head[6:10])
            if head[:2] == b"\xff\xd8":
                fh.seek(2)
                while True:
                    marker = fh.read(2)
                    if len(marker) < 2 or marker[0] != 0xFF:
                        return None
                    code = marker[1]
                    if code in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                                0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                        fh.read(3)
                        h, w = struct.unpack(">HH", fh.read(4))
                        return w, h
                    length = struct.unpack(">H", fh.read(2))[0]
                    fh.seek(max(0, length - 2), 1)
    except Exception:
        return None
    return None


def _strip_jpeg_metadata(data):
    """去掉 JPEG 的 APP1 段（EXIF/XMP），避免携带相机/GPS 等元数据。"""
    if data[:2] != b"\xff\xd8":
        return data
    out = bytearray(data[:2])
    i, n = 2, len(data)
    while i + 4 <= n:
        if data[i] != 0xFF:
            break
        marker = data[i + 1]
        if marker == 0xDA:  # SOS 之后是压缩图像数据
            out += data[i:]
            return bytes(out)
        if marker in (0xD8, 0xD9):
            out += data[i:i + 2]
            i += 2
            continue
        length = struct.unpack(">H", data[i + 2:i + 4])[0]
        end = i + 2 + length
        if end > n:
            break
        if marker != 0xE1:  # 丢弃 APP1（EXIF/XMP）
            out += data[i:end]
        i = end
    out += data[i:]
    return bytes(out)


def _strip_png_metadata(data):
    """去掉 PNG 的 eXIf/tEXt/iTXt/zTXt 文本块。"""
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        return data
    out = bytearray(data[:8])
    i, n = 8, len(data)
    while i + 8 <= n:
        length = struct.unpack(">I", data[i:i + 4])[0]
        ctype = data[i + 4:i + 8]
        end = i + 12 + length
        if end > n:
            break
        if ctype not in (b"eXIf", b"tEXt", b"iTXt", b"zTXt"):
            out += data[i:end]
        i = end
        if ctype == b"IEND":
            break
    return bytes(out)


def _strip_image_metadata(data, ext):
    if ext in ("jpg", "jpeg"):
        return _strip_jpeg_metadata(data)
    if ext == "png":
        return _strip_png_metadata(data)
    return data


def _image_run(path, alt=""):
    """生成行内图片 run；文件缺失或格式不支持时回退为文字提示。"""
    full = path if os.path.isabs(path) else os.path.join(_BASE_DIR[0], path)
    if not os.path.isfile(full):
        return run("[图片缺失: %s]" % (alt or path), color="C00000")
    ext = os.path.splitext(full)[1].lower().lstrip(".")
    if ext not in ("png", "jpg", "jpeg", "gif"):
        return run("[不支持的图片格式: %s]" % path, color="C00000")
    size = _image_size(full) or (300, 200)
    px_w, px_h = max(1, size[0]), max(1, size[1])
    cx, cy = px_w * _PX_EMU, px_h * _PX_EMU
    scale = 1.0
    if cx > _MAX_IMAGE_W:
        scale = _MAX_IMAGE_W / float(cx)
    if cy * scale > _MAX_IMAGE_H:
        scale = _MAX_IMAGE_H / float(cy)
    cx, cy = int(cx * scale), int(cy * scale)
    with open(full, "rb") as fh:
        data = _strip_image_metadata(fh.read(), ext)
    rid = _next_rel_id()
    media_name = "image%d.%s" % (_REL_SEQ[0], "jpeg" if ext in ("jpg", "jpeg") else ext)
    _MEDIA.append((rid, media_name, data))
    _PIC_SEQ[0] += 1
    pid = _PIC_SEQ[0]
    name = escape(alt or media_name, {'"': "&quot;"})
    return (
        '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="%s"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
        % (cx, cy, pid, name, pid, name, rid, cx, cy)
    )


def _hyperlink_runs(text, url):
    """生成真正可点击的超链接。"""
    rid = _next_rel_id()
    _EXTERNAL_RELS.append((rid, url))
    inner = ('<w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>'
             '<w:t xml:space="preserve">' + esc(text) + "</w:t></w:r>")
    return '<w:hyperlink r:id="%s">%s</w:hyperlink>' % (rid, inner)


# 正文里的数学写法分两类：
#   1) 向量/均值等「基字母 + 组合标记」（v⃗ = v+U+20D7、x̄ = x+U+0304），
#      以及显式 LaTeX 命令（\vec{v}、\bar{x}）——统一转成行内 OMML 公式对象，
#      与公式区完全同一种渲染；
#   2) 其它符号区字符（希腊字母、运算符等）——单独 run 用数学字体。
COMBINE_RE = re.compile(
    "([A-Za-z\u00C0-\u024F\u0370-\u03FF][\u0300-\u036F\u20D0-\u20FF]+)"
)
CMD_RE = re.compile(
    r"(\\(?:vec|bar|hat|tilde|dot|ddot|overline|underline)\{[^}]{1,8}\})"
)
_COMBINE_MAP = [
    ("\u20d7", r"\vec{%s}"),
    ("\u20d6", r"\vec{%s}"),
    ("\u0304", r"\bar{%s}"),
    ("\u0305", r"\overline{%s}"),
    ("\u0302", r"\hat{%s}"),
    ("\u0303", r"\tilde{%s}"),
    ("\u0307", r"\dot{%s}"),
    ("\u0308", r"\ddot{%s}"),
]
# —— 正文数学片段识别 ——
# 优先级：组合标记序列 > 字母+上下标字符 > 预组合 macron 字母 > LaTeX 命令 > 符号区
_COMB = r"[A-Za-z\u00C0-\u024F\u0370-\u03FF][\u0300-\u036F\u20D0-\u20FF]+"
_SUBSUP = r"[A-Za-z\u00C0-\u024F\u0370-\u03FF][\u2070-\u209F\u00B2\u00B3\u00B9]+"
_PRECOMB = "[\\u0100\\u0101\\u0112\\u0113\\u012A\\u012B\\u014C\\u014D\\u016A\\u016B\\u0232\\u0233\\u1E20\\u1E21]"
_CMD = r"\\(?:vec|bar|hat|tilde|dot|ddot|overline|underline)\{[^}]{1,8}\}"
_SYMBOLS = ("[\u00AC\u00B0\u00B1\u00B5\u00B7\u00D7\u00F7\u00B2\u00B3\u00B9\u00BC-\u00BE"
            "\u2026\u2032\u2033\u2044"
            "\u0370-\u03FF\u1F00-\u1FFF\u2070-\u209F\u2100-\u214F"
            "\u2190-\u21FF\u2200-\u22FF\u2300-\u23FF\u27C0-\u27EF"
            "\u27F0-\u27FF\u2900-\u297F\u2980-\u29FF\u2A00-\u2AFF"
            "\U0001D400-\U0001D7FF]+")
MATH_FRAGMENT_RE = re.compile(
    "(" + "|".join([_COMB, _SUBSUP, _PRECOMB, _CMD, _SYMBOLS]) + ")"
)
_SYMBOL_CLASS = _SYMBOLS[1:-2]
_SYMBOL_SEQ_RE = re.compile(
    "[" + _SYMBOL_CLASS + "](?:[ \t]*[" + _SYMBOL_CLASS + "])+"
)
MATH_FONT = "Cambria Math"
_PRECOMB_BASE = {
    "\u0100": "A", "\u0101": "a", "\u0112": "E", "\u0113": "e",
    "\u012A": "I", "\u012B": "i", "\u014C": "O", "\u014D": "o",
    "\u016A": "U", "\u016B": "u", "\u0232": "Y", "\u0233": "y",
    "\u1E20": "G", "\u1E21": "g",
}


def _subsup_char(ch):
    o = ord(ch)
    if 0x2070 <= o <= 0x2079:
        return "sup", chr(ord("0") + o - 0x2070)
    if o == 0x207F:
        return "sup", "n"
    if 0x207A <= o <= 0x207E:
        return "sup", "+-=()"[o - 0x207A]
    if 0x2080 <= o <= 0x2089:
        return "sub", chr(ord("0") + o - 0x2080)
    if 0x208A <= o <= 0x208E:
        return "sub", "+-=()"[o - 0x208A]
    if o == 0x00B2:
        return "sup", "2"
    if o == 0x00B3:
        return "sup", "3"
    if o == 0x00B9:
        return "sup", "1"
    return None


def _subsup_tex(frag):
    base, rest = frag[0], frag[1:]
    pieces = []
    kind = None
    buf = []
    for ch in rest:
        info = _subsup_char(ch)
        if info is None:
            continue
        k, v = info
        if k != kind:
            if kind is not None:
                pieces.append(("_{" if kind == "sub" else "^{") + "".join(buf) + "}")
            kind, buf = k, [v]
        else:
            buf.append(v)
    if kind is not None:
        pieces.append(("_{" if kind == "sub" else "^{") + "".join(buf) + "}")
    return base + "".join(pieces)


def combined_to_tex(chunk):
    """把「基字母 + 组合标记」转成 LaTeX；无法映射时返回 None。"""
    base, marks = chunk[0], chunk[1:]
    for mark in marks:
        for known, template in _COMBINE_MAP:
            if mark == known:
                return template % base
    return None


def _fragment_to_tex(frag):
    """数学片段 -> LaTeX；返回 None 表示只按数学字体显示、不转公式对象。"""
    if re.fullmatch(_COMB, frag):
        return combined_to_tex(frag)
    if re.fullmatch(_SUBSUP, frag):
        return _subsup_tex(frag)
    if re.fullmatch(_PRECOMB, frag):
        return "\\bar{%s}" % _PRECOMB_BASE[frag]
    if re.fullmatch(_CMD, frag):
        name = re.match(r"\\([a-z]+)\{", frag).group(1)
        inner = frag[frag.index("{") + 1:frag.rindex("}")]
        return "\\" + name + "{" + inner + "}"
    if re.fullmatch(_SYMBOLS, frag):
        return _symbols_to_tex(frag)
    return None


# —— 整段数学表达式（说明文本中连续书写的公式）——
_EXPR_CHUNK_RE = re.compile(
    r"[A-Za-z0-9\u00C0-\u024F\u0370-\u03FF\u1F00-\u1FFF"
    r"\u2070-\u209F\u0300-\u036F\u20D0-\u20FF"
    r"\u00AC\u00B0\u00B1\u00B5\u00B7\u00D7\u00F7\u00B2\u00B3\u00B9\u00BC-\u00BE"
    r"\u2026\u2032\u2033\u2044"
    r"\u2100-\u214F\u2190-\u21FF\u2200-\u22FF\u2300-\u23FF"
    r"\u27C0-\u27EF\u27F0-\u27FF\u2900-\u297F\u2980-\u29FF\u2A00-\u2AFF"
    r" =+\-*/×÷±≤≥≠≈<>^_()\[\]{}|,.:;]+"
)
_STRONG_OPS = set("=≠≤≥≈×÷±∓→⇒⇔∈∉⊂⊆⊃⊇∪∩∑∏∫∮∂∇∞∝≅∼≡")
_LATEX_SYMBOLS = {
    "α": r"\alpha", "β": r"\beta", "γ": r"\gamma", "δ": r"\delta",
    "ε": r"\epsilon", "ζ": r"\zeta", "η": r"\eta", "θ": r"\theta",
    "ι": r"\iota", "κ": r"\kappa", "λ": r"\lambda", "μ": r"\mu",
    "ν": r"\nu", "ξ": r"\xi", "π": r"\pi", "ρ": r"\rho",
    "σ": r"\sigma", "τ": r"\tau", "υ": r"\upsilon", "φ": r"\phi",
    "χ": r"\chi", "ψ": r"\psi", "ω": r"\omega",
    "Γ": r"\Gamma", "Δ": r"\Delta", "Θ": r"\Theta", "Λ": r"\Lambda",
    "Ξ": r"\Xi", "Π": r"\Pi", "Σ": r"\Sigma", "Υ": r"\Upsilon",
    "Φ": r"\Phi", "Ψ": r"\Psi", "Ω": r"\Omega",
    "×": r"\times", "÷": r"\div", "±": r"\pm", "∓": r"\mp",
    "≤": r"\le", "≥": r"\ge", "≠": r"\ne", "≈": r"\approx",
    "≡": r"\equiv", "≅": r"\cong", "∼": r"\sim", "∝": r"\propto",
    "→": r"\to", "←": r"\leftarrow", "↔": r"\leftrightarrow",
    "⇒": r"\Rightarrow", "⇐": r"\Leftarrow", "⇔": r"\Leftrightarrow",
    "∈": r"\in", "∉": r"\notin", "⊂": r"\subset", "⊆": r"\subseteq",
    "⊃": r"\supset", "⊇": r"\supseteq", "∪": r"\cup", "∩": r"\cap",
    "∅": r"\emptyset", "∞": r"\infty", "∂": r"\partial", "∇": r"\nabla",
    "∑": r"\sum", "∏": r"\prod", "∫": r"\int", "∮": r"\oint",
    "∴": r"\therefore", "∵": r"\because",
    "…": r"\cdots", "′": "'", "″": "''", "°": r"\degree", "√": r"\surd",
    "·": r"\cdot", "⋅": r"\cdot", "−": "-",
    "∀": r"\forall", "∃": r"\exists", "¬": r"\neg",
    "∧": r"\wedge", "∨": r"\vee", "⊕": r"\oplus", "⊗": r"\otimes",
    "⟨": r"\langle", "⟩": r"\rangle",
    "⌈": r"\lceil", "⌉": r"\rceil", "⌊": r"\lfloor", "⌋": r"\rfloor",
    "∖": r"\setminus", "≃": r"\simeq", "⩽": r"\leqslant", "⩾": r"\geqslant",
    "≦": r"\leq", "≧": r"\geq",
    "ℓ": r"\ell", "ℏ": r"\hbar",
    "ℕ": r"\mathbb{N}", "ℤ": r"\mathbb{Z}", "ℚ": r"\mathbb{Q}",
    "ℝ": r"\mathbb{R}", "ℂ": r"\mathbb{C}",
}


# 这些字符保持原样文本（不转公式）：作为数学上标/下标会没有基准字符（空 base），
# 在 Word/WPS 中悬挂错位；角分/角秒/度号本来就是普通正文字符。
_KEEP_TEXT_SYMBOLS = frozenset("′″°")


def _symbols_to_tex(frag):
    """符号片段 -> LaTeX；任一字符无映射时返回 None（回退数学字体文字）。"""
    out = []
    for ch in frag:
        if ch in _KEEP_TEXT_SYMBOLS:
            return None
        tex = _LATEX_SYMBOLS.get(ch)
        if tex is None:
            return None
        out.append(tex)
    return "".join(out)


def _is_math_expr(seg):
    s = seg.strip()
    if len(s) < 3 or len(s) > 60:
        return False
    if not re.search(r"[A-Za-z0-9\u0370-\u03FF]", s):
        return False
    if not any(ch in _STRONG_OPS for ch in s):
        return False
    if re.search(r"https?://|www\.|@", s):
        return False
    return True


def _expr_to_tex(seg):
    """整段表达式 -> LaTeX；遇到无法安全映射的字符返回 None。"""
    out = []
    i = 0
    n = len(seg)
    while i < n:
        rest = seg[i:]
        m = re.match(_CMD, rest)
        if m:
            out.append(m.group(0))
            i += len(m.group(0))
            continue
        m = re.match(r"[A-Za-z\u00C0-\u024F\u0370-\u03FF][\u0300-\u036F\u20D0-\u20FF]+", rest)
        if m:
            tex = combined_to_tex(m.group(0))
            if tex is None:
                return None
            out.append(tex)
            i += len(m.group(0))
            continue
        m = re.match(r"[A-Za-z\u00C0-\u024F\u0370-\u03FF][\u2070-\u209F\u00B2\u00B3\u00B9]+", rest)
        if m:
            out.append(_subsup_tex(m.group(0)))
            i += len(m.group(0))
            continue
        ch = seg[i]
        if ch in _PRECOMB_BASE:
            out.append("\\bar{%s}" % _PRECOMB_BASE[ch])
            i += 1
            continue
        if ch == "√":
            nxt = seg[i + 1:i + 2]
            if not nxt:
                return None
            if nxt == "(":
                depth = 0
                j = i
                while j < n:
                    c = seg[j]
                    if c == "(":
                        depth += 1
                    elif c == ")":
                        depth -= 1
                        if depth == 0:
                            break
                    j += 1
                if depth != 0:
                    return None
                out.append("\\sqrt{" + seg[i + 2:j].strip() + "}")
                i = j + 1
                continue
            if not re.match(r"[A-Za-z0-9\u0370-\u03FF]", nxt):
                return None
            out.append("\\sqrt{" + nxt + "}")
            i += 2
            continue
        if ch in _LATEX_SYMBOLS:
            tex = _LATEX_SYMBOLS[ch]
            out.append(tex)
            nxt = seg[i + 1:i + 2]
            if tex and tex[-1].isalpha() and nxt and nxt.isalnum():
                out.append(" ")
            i += 1
            continue
        if ch in "\\{}":
            return None
        out.append(ch)
        i += 1
    return "".join(out)


def _fragment_chunk(chunk, bold=False, italic=False, color=None, sz=None, stats=None):
    out = []
    for part in MATH_FRAGMENT_RE.split(chunk):
        if not part:
            continue
        if MATH_FRAGMENT_RE.fullmatch(part):
            tex = None if (bold or italic) else _fragment_to_tex(part)
            if tex is not None:
                try:
                    out.append(latex_to_omml(tex))
                    if stats is not None:
                        stats["omml"] += 1
                    continue
                except Exception:
                    pass
            out.append(run(part, bold=bold, italic=italic, color=color, sz=sz,
                           font=MATH_FONT))
        else:
            out.append(run(part, bold=bold, italic=italic, color=color, sz=sz))
    return "".join(out)


def _fragments_runs(text, bold=False, italic=False, color=None, sz=None, stats=None):
    """先合并连续符号序列为一个公式对象，再按数学片段处理其余部分。"""
    if bold or italic:
        return _fragment_chunk(text, bold, italic, color, sz, stats)
    out = []
    pos = 0
    for match in _SYMBOL_SEQ_RE.finditer(text):
        seg = match.group(0)
        compact = re.sub(r"\s+", "", seg)
        tex = _symbols_to_tex(compact)
        if tex is None:
            continue
        try:
            omml = latex_to_omml(tex)
        except Exception:
            continue
        if match.start() > pos:
            out.append(_fragment_chunk(text[pos:match.start()], bold, italic,
                                       color, sz, stats))
        out.append(omml)
        if stats is not None:
            stats["omml"] += 1
        pos = match.end()
    if pos < len(text):
        out.append(_fragment_chunk(text[pos:], bold, italic, color, sz, stats))
    return "".join(out)


def text_runs(text, bold=False, italic=False, color=None, sz=None, stats=None):
    """说明文字 -> run 序列；先识别整段表达式，再按数学片段处理其余部分。"""
    if bold or italic:
        return _fragments_runs(text, bold, italic, color, sz, stats)
    out = []
    pos = 0
    for match in _EXPR_CHUNK_RE.finditer(text):
        seg = match.group(0)
        if not _is_math_expr(seg):
            continue
        tex = _expr_to_tex(seg)
        if tex is None:
            continue
        try:
            omml = latex_to_omml(tex)
        except Exception:
            continue
        if match.start() > pos:
            out.append(_fragments_runs(text[pos:match.start()], bold, italic,
                                       color, sz, stats))
        out.append(omml)
        if stats is not None:
            stats["omml"] += 1
        pos = match.end()
    if pos < len(text):
        out.append(_fragments_runs(text[pos:], bold, italic, color, sz, stats))
    return "".join(out)


def para(inner, style=None, align=None, ppr_extra=""):
    ppr = ""
    if style:
        ppr += '<w:pStyle w:val="%s"/>' % style
    if align:
        ppr += '<w:jc w:val="%s"/>' % align
    ppr += ppr_extra
    open_p = "<w:p>" + ("<w:pPr>" + ppr + "</w:pPr>" if ppr else "")
    return open_p + inner + "</w:p>"


def hr():
    return ('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" '
            'w:space="1" w:color="BBBBBB"/></w:pBdr></w:pPr></w:p>')


# ---------------- LaTeX -> OMML ----------------
_ENV_RE = re.compile(
    r"\\begin\{(aligned|align\*?|eqnarray\*?|gather\*?|split|split\*?)\}(.*?)\\end\{\1\}",
    re.S,
)


def _env_sub(match):
    body = match.group(2).replace("&", " ")
    return r"\begin{array}{l}" + body + r"\end{array}"


_TEX_COMMAND_FIXES = [
    (re.compile(r"\\AA(?![A-Za-z])"), r"\\text{Å}"),
    (re.compile(r"\\textdegree(?![A-Za-z])"), r"^{\\circ}"),
    (re.compile(r"\\ohm(?![A-Za-z])"), r"\\Omega"),
    (re.compile(r"\\micro(?![A-Za-z])"), r"\\mu"),
    (re.compile(r"\\per(?![A-Za-z])"), "/"),
]


# 函数名后紧跟变量时，latex2mathml 不生成间距（\log g -> logg）；
# 在 MathML 层给平级的函数名 mi 后补一个窄空格 mspace。
_MATHML_FUNC_NAMES = ("log", "lg", "ln", "exp", "sin", "cos", "tan", "cot",
                      "sec", "csc", "sinh", "cosh", "tanh", "arcsin", "arccos",
                      "arctan", "max", "min", "sup", "inf", "det", "dim", "lim",
                      "deg", "gcd")
_MATHML_FUNC_RE = re.compile(
    r"(<mi>(?:" + "|".join(_MATHML_FUNC_NAMES) + r")</mi>)"
    r"(?=<mi>[^<]+</mi>|<msub>|<msup>|<mfrac>)")
# 把 (...) 上/下标改写为 {(...)} 上/下标：直接写右括号时 mathml2omml
# 只把 ")" 当作上下标基准（base=')'），包成组后整个括号组才是基准。
_GROUP_SCRIPT_RE = re.compile(r"(?<!\\left)\(([^()]*)\)([_^])")


def _wrap_group_scripts(tex):
    """(...)^{...} / (...)_{...} -> {(...)}^{...}；迭代处理简单嵌套。"""
    for _ in range(3):
        new = _GROUP_SCRIPT_RE.sub(r"{(\1)}\2", tex)
        if new == tex:
            break
        tex = new
    return tex


def preprocess_tex(tex):
    for pattern, repl in _TEX_COMMAND_FIXES:
        tex = pattern.sub(repl, tex)
    tex = _wrap_group_scripts(tex)
    # \mathrm{''/yr} 这类写法里的 LaTeX 双 prime 会被转成无基准的上标；
    # 在正体文本内直接换成 Unicode ″（视觉等价、无空 base）。
    tex = re.sub(r"\\(mathrm|text)\{([^{}]*)\}",
                 lambda m: "\\" + m.group(1) + "{"
                 + m.group(2).replace("''", "\u2033") + "}", tex)
    tex = re.sub(r"\\label\{[^}]*\}", "", tex)
    tex = re.sub(r"\\(nonumber|notag)\b", "", tex)
    tex = _ENV_RE.sub(_env_sub, tex)
    return tex.strip()


def latex_to_omml(tex):
    """返回 <m:oMath>...</m:oMath>；失败抛异常。相同 LaTeX 走缓存。"""
    key = tex.strip()
    cached = _OMML_CACHE.get(key)
    if cached is not None:
        return cached
    mathml = _latex2mathml(preprocess_tex(tex))
    mathml = _MATHML_FUNC_RE.sub(r'\1<mspace width="0.167em"/>', mathml)
    omml = _mathml2omml(mathml)
    # mathml2omml 把 mspace 转成空 m:t run（不占宽度）；改成真正的空格 run，
    # 否则 \log g 在 Word/WPS 中仍渲染成 "logg"。
    omml = omml.replace("<m:t></m:t>", '<m:t xml:space="preserve"> </m:t>')
    omml = omml.replace("<m:t />", '<m:t xml:space="preserve"> </m:t>')
    omml = omml.replace("<m:t/>", '<m:t xml:space="preserve"> </m:t>')
    try:
        _assert_well_formed(omml)
    except ET.ParseError:
        fixed = omml
        for pattern, repl in _OMML_FIXES:
            fixed = pattern.sub(repl, fixed)
        _assert_well_formed(fixed)
        omml = fixed
    omml = _omml_to_acc(omml)
    _assert_well_formed(omml)
    if len(_OMML_CACHE) < 2000:
        _OMML_CACHE[key] = omml
    return omml


def _assert_well_formed(omml):
    ET.fromstring('<w:r xmlns:w="%s" xmlns:m="%s">%s</w:r>' % (W_NS, M_NS, omml))


# mathml2omml 0.0.2 的已知缺陷：mover/munder 生成的 groupChrPr 少了闭合标签。
_OMML_FIXES = [
    (re.compile(r"(<m:groupChr><m:groupChrPr>.*?)</m:groupChr>", re.S),
     r"\1</m:groupChrPr>"),
]

ET.register_namespace("m", M_NS)
ET.register_namespace("w", W_NS)

# mathml2omml 把 \vec / \bar / \hat 等生成为 m:groupChr / m:limUpp；
# Word/WPS 里「手打」的公式是 m:acc（重音）结构。这里统一改写成 m:acc，
# 使行内向量/均值与公式区、以及用户手打的公式字形一致。
_ACCENT_TARGETS = {
    "\u2192": "\u20d7",
    "\u20d7": "\u20d7",
    "\u00af": "\u0304",
    "\u2015": "\u0304",
    "\u0304": "\u0304",
    "^": "\u0302",
    "\u0302": "\u0302",
    "~": "\u0303",
    "\u0303": "\u0303",
    "\u02d9": "\u0307",
    "\u0307": "\u0307",
    "\u00a8": "\u0308",
    "\u0308": "\u0308",
}


def _omml_to_acc(omml):
    """把 groupChr / limUpp 形式的重音改写成 OMML 标准的 m:acc。"""
    M = "{%s}" % M_NS
    try:
        root = ET.fromstring('<root xmlns:m="%s">%s</root>' % (M_NS, omml))
    except ET.ParseError:
        return omml
    for _ in range(3):
        changed = False
        for parent in list(root.iter()):
            for index, child in enumerate(list(parent)):
                if not child.tag.startswith(M):
                    continue
                tag = child.tag[len(M):]
                if tag not in ("groupChr", "limUpp"):
                    continue
                symbol = None
                if tag == "groupChr":
                    chr_el = child.find(".//" + M + "chr")
                    if chr_el is not None:
                        symbol = chr_el.get(M + "val")
                else:
                    lim = child.find(".//" + M + "lim")
                    if lim is not None:
                        text_el = lim.find(".//" + M + "t")
                        if text_el is not None:
                            symbol = (text_el.text or "").strip()
                target = _ACCENT_TARGETS.get(symbol)
                if target is None:
                    continue
                base = child.find(M + "e")
                if base is None:
                    continue
                acc = ET.Element(M + "acc")
                acc_pr = ET.SubElement(acc, M + "accPr")
                chr_el = ET.SubElement(acc_pr, M + "chr")
                chr_el.set(M + "val", target)
                new_e = ET.SubElement(acc, M + "e")
                for sub in list(base):
                    base.remove(sub)
                    new_e.append(sub)
                parent.remove(child)
                parent.insert(index, acc)
                changed = True
        if not changed:
            break
    text = ET.tostring(root, encoding="unicode")
    text = text.split(">", 1)[1] if ">" in text else text
    if "</root>" in text:
        text = text.rsplit("</root>", 1)[0]
    return text


def _omml_limloc_undovr(omml):
    """块级公式：求和/连乘的上下限放在正上/正下，积分的上下限留在右侧（书写习惯）。"""
    M = "{%s}" % M_NS
    integral_chars = set("∫∬∭∮∯∰∱∲∳")
    try:
        root = ET.fromstring('<root xmlns:m="%s">%s</root>' % (M_NS, omml))
    except ET.ParseError:
        return omml
    changed = False
    for nary in root.iter(M + "nary"):
        pr = nary.find(M + "naryPr")
        if pr is None:
            pr = ET.Element(M + "naryPr")
            nary.insert(0, pr)
        chr_el = pr.find(M + "chr")
        symbol = chr_el.get(M + "val") if chr_el is not None else ""
        target = "subSup" if symbol in integral_chars else "undOvr"
        lim = pr.find(M + "limLoc")
        if lim is None:
            lim = ET.Element(M + "limLoc")
            if chr_el is not None:
                pr.insert(list(pr).index(chr_el) + 1, lim)
            else:
                pr.insert(0, lim)
        lim.set(M + "val", target)
        changed = True
    if not changed:
        return omml
    text = ET.tostring(root, encoding="unicode")
    text = text.split(">", 1)[1] if ">" in text else text
    if "</root>" in text:
        text = text.rsplit("</root>", 1)[0]
    return text


_FUNC_LIMLOW = {"lim", "max", "min", "sup", "inf", "limsup", "liminf",
                "argmax", "argmin"}


def _omml_func_limlow(omml):
    """块级公式：把 lim/max/min 等函数的下标从右下（sSub）改为正下方（limLow）。"""
    M = "{%s}" % M_NS
    try:
        root = ET.fromstring('<root xmlns:m="%s">%s</root>' % (M_NS, omml))
    except ET.ParseError:
        return omml
    changed = False
    for parent in list(root.iter()):
        for idx, child in enumerate(list(parent)):
            if child.tag != M + "sSub":
                continue
            e = child.find(M + "e")
            sub = child.find(M + "sub")
            if e is None or sub is None:
                continue
            text = "".join(t.text or "" for t in e.iter()
                           if t.tag == M + "t").strip().lower()
            if text not in _FUNC_LIMLOW:
                continue
            low = ET.Element(M + "limLow")
            new_e = ET.SubElement(low, M + "e")
            for node in list(e):
                e.remove(node)
                new_e.append(node)
            new_lim = ET.SubElement(low, M + "lim")
            for node in list(sub):
                sub.remove(node)
                new_lim.append(node)
            parent.remove(child)
            parent.insert(idx, low)
            changed = True
    if not changed:
        return omml
    text_out = ET.tostring(root, encoding="unicode")
    text_out = text_out.split(">", 1)[1] if ">" in text_out else text_out
    if "</root>" in text_out:
        text_out = text_out.rsplit("</root>", 1)[0]
    return text_out


def omath_para(tex):
    omml = latex_to_omml(tex)
    omml = _omml_limloc_undovr(omml)
    omml = _omml_func_limlow(omml)
    return ('<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'
            + omml + "</m:oMathPara>")


# ---------------- 行内解析 ----------------
_SCRIPT_BASE_RE = re.compile(r"([A-Za-z0-9]+)\$([_^](?:\{[^{}]*\}|[A-Za-z0-9]))\$")
_CJK_SCRIPT_RE = re.compile(r"([^\sA-Za-z0-9$])\$([_^])\{?([0-9+\-−]+)\}?\$")
_UNICODE_SUP = {"0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴",
                "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹",
                "+": "⁺", "-": "⁻", "−": "⁻"}
_UNICODE_SUB = {"0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄",
                "5": "₅", "6": "₆", "7": "₇", "8": "₈", "9": "₉",
                "+": "₊", "-": "₋", "−": "₋"}


def _fix_script_base(text):
    """归一化「$ 紧贴 ^/_」的写法，避免生成 base 为空的上下标。

    1) `SiO$_2$`、`cm$^{-2}$`：把 $ 前紧邻的字母数字串并入公式作基准；
    2) 中文/标点后的 `$^2$`（如「离轴角$^2$」）：转成 Unicode 上/下标字符，
       与正文既有的 10⁶、e⁻ 等写法保持一致（Unicode 无法表示时保持原样）。
    """
    def repl(m):
        base, script = m.group(1), m.group(2)
        if len(base) == 1 or base.isdigit():
            return "$" + base + script + "$"
        return "$\\mathrm{" + base + "}" + script + "$"

    def repl2(m):
        table = _UNICODE_SUP if m.group(2) == "^" else _UNICODE_SUB
        out = "".join(table.get(c, "") for c in m.group(3))
        if len(out) != len(m.group(3)):
            return m.group(0)
        return m.group(1) + out

    text = _SCRIPT_BASE_RE.sub(repl, text)
    return _CJK_SCRIPT_RE.sub(repl2, text)


def _styled_text_runs(text, bold=False, italic=False, stats=None):
    """加粗/斜体文本也要解析 $...$ 行内公式（原实现整段跳过公式转换）。"""
    out = []
    for seg in re.split(r"(\$[^$]+\$)", text):
        if not seg:
            continue
        if len(seg) > 2 and seg.startswith("$") and seg.endswith("$"):
            try:
                out.append(_omml_limloc_undovr(latex_to_omml(seg[1:-1])))
                if stats is not None:
                    stats["omml"] += 1
            except Exception:
                if stats is not None:
                    stats["failed"].append(seg)
                out.append(run(seg, bold=bold, italic=italic, code=True))
        else:
            out.append(text_runs(seg, bold=bold, italic=italic, stats=stats))
    return "".join(out)


_INLINE_RE = re.compile(
    r"(!\[[^\]]*\]\([^)]+\)|\$[^$]+\$|\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`|\[[^\]]+\]\([^)]+\))"
)


def inline(text, stats):
    out = []
    text = _fix_script_base(text)
    for part in _INLINE_RE.split(text):
        if not part:
            continue
        if part.startswith("![") and part.endswith(")"):
            m = re.match(r"!\[([^\]]*)\]\(([^)]+)\)", part)
            if m:
                out.append(_image_run(m.group(2), m.group(1)))
            else:
                out.append(text_runs(part, stats=stats))
        elif len(part) > 2 and part.startswith("$") and part.endswith("$"):
            try:
                out.append(_omml_limloc_undovr(latex_to_omml(part[1:-1])))
                stats["omml"] += 1
            except Exception:
                stats["failed"].append(part)
                out.append(run(part, code=True))
        elif part.startswith("**") and part.endswith("**"):
            out.append(_styled_text_runs(part[2:-2], bold=True, stats=stats))
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            out.append(_styled_text_runs(part[1:-1], italic=True, stats=stats))
        elif part.startswith("`") and part.endswith("`"):
            out.append(run(part[1:-1], code=True))
        elif part.startswith("["):
            m = re.match(r"\[([^\]]+)\]\(([^)]+)\)", part)
            if m:
                out.append(_hyperlink_runs(m.group(1), m.group(2)))
            else:
                out.append(text_runs(part, stats=stats))
        else:
            out.append(text_runs(part, stats=stats))
    return "".join(out)


_LARGE_MARKERS = (r"\frac", r"\sqrt", r"\sum", r"\prod", r"\int", r"\iint",
                  r"\iiint", r"\oint", r"\lim", r"\begin{", r"\left",
                  r"\binom", r"\overline", r"\underline", r"\matrix",
                  r"\pmatrix")


def _is_large_formula(tex):
    """行内公式是否属于「大型公式」（应独立成行显示）。"""
    return any(marker in tex for marker in _LARGE_MARKERS)


def emit_text_block(text, stats, style=None, ppr_extra=""):
    """把段落按大型公式拆分：文字段落 + 独立的居中式公式段落。"""
    segments = []
    for part in _INLINE_RE.split(text):
        if not part:
            continue
        if len(part) > 2 and part.startswith("$") and part.endswith("$"):
            tex = part[1:-1]
            if _is_large_formula(tex):
                segments.append(("large", tex))
                continue
        segments.append(("text", part))
    body = []
    buf = ""
    for kind, value in segments:
        if kind == "text":
            buf += value
        else:
            if buf.strip():
                body.append(para(inline(buf, stats), style=style,
                                 ppr_extra=ppr_extra))
                buf = ""
            try:
                body.append(para(omath_para(value), align="center"))
                stats["omml"] += 1
            except Exception:
                stats["failed"].append("$" + value + "$")
                body.append(para(run("$" + value + "$", code=True)))
    if buf.strip():
        body.append(para(inline(buf, stats), style=style, ppr_extra=ppr_extra))
    return body


def is_block_start(line):
    s = line.strip()
    return (s.startswith("```") or s.startswith("$$") or re.match(r"^#{1,6}\s+", s)
            or s.startswith(">") or re.match(r"^\s*[-*+]\s+", line)
            or re.match(r"^\s*\d+\.\s+", line) or s.startswith("|")
            or re.match(r"^(-{3,}|\*{3,}|_{3,})$", s))


_CAPTION_RE = re.compile(r"^\s*(图|表|Figure|Fig\.?|Table)\s*[\d一二三四五六七八九十]")


def caption_para(text, stats):
    """图片题注：图片下方居中，五号字（10.5pt）。"""
    return para(inline(text, stats), style="Caption", align="center")


def emit_media_block(text, stats, style=None, ppr_extra="", prefix=""):
    """段落含图片时，图片拆成独立居中段落；紧跟的题注居中、五号。"""
    images = _IMAGE_INLINE_RE.findall(text)
    if not images:
        return emit_text_block(prefix + text, stats, style=style, ppr_extra=ppr_extra)
    parts = _IMAGE_INLINE_RE.split(text)
    body = []
    for idx, chunk in enumerate(parts):
        lead = prefix if idx == 0 else ""
        stripped = chunk.strip()
        if stripped:
            if idx > 0 and _CAPTION_RE.match(stripped):
                body.append(caption_para(stripped, stats))
            else:
                body.extend(emit_text_block(lead + chunk, stats, style=style,
                                            ppr_extra=ppr_extra))
        if idx < len(images):
            body.append(para(inline(images[idx], stats), align="center",
                             ppr_extra='<w:keepNext/><w:ind w:firstLine="0" w:firstLineChars="0"/>'))
    return body


_IMAGE_INLINE_RE = re.compile(r"!\[[^\]]*\]\([^)]+\)")


def make_table(rows, stats):
    cols = max(len(r) for r in rows)
    width = max(1, int(9360 / cols))
    grid = "".join('<w:gridCol w:w="%d"/>' % width for _ in range(cols))
    xml = ['<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
           '<w:jc w:val="center"/>'
           '<w:tblBorders>'
           '<w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
           '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
           '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
           "</w:tblBorders></w:tblPr><w:tblGrid>" + grid + "</w:tblGrid>"]
    for ri, row in enumerate(rows):
        xml.append("<w:tr>")
        for ci in range(cols):
            cell = row[ci] if ci < len(row) else ""
            content = inline(cell, stats)
            xml.append(
                '<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>'
                '<w:vAlign w:val="center"/></w:tcPr>'
                + para(content, style=("TableHead" if ri == 0 else None),
                       align="center",
                       ppr_extra='<w:ind w:firstLine="0" w:firstLineChars="0"/>')
                + "</w:tc>"
            )
        xml.append("</w:tr>")
    xml.append("</w:tbl>")
    return "".join(xml)


# ---------------- Markdown 主解析 ----------------
def _collect_headers(md):
    """收集 1-3 级标题，供 [[TOC]] 指令生成目录预览。"""
    headers = []
    in_code = False
    for line in md.splitlines():
        s = line.strip()
        if s.startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", s)
        if m and len(m.group(1)) <= 3:
            headers.append((len(m.group(1)), m.group(2).strip()))
    return headers


PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def _ends_with_page_break(body):
    return bool(body) and body[-1] == PAGE_BREAK


def toc_block(headers, page_break_before=True):
    """目录前后各分页；只在 [[TOC]] 指令处插入，不自动生成。"""
    xml = []
    if page_break_before:
        xml.append(PAGE_BREAK)
    xml.append('<w:p><w:pPr><w:spacing w:after="240"/>'
               '<w:ind w:firstLine="0" w:firstLineChars="0"/><w:jc w:val="center"/>'
               '</w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/>'
               '</w:rPr><w:t>目录</w:t></w:r></w:p>')
    if not headers:
        xml.append(para(run("（文档中没有可列入目录的标题）"), align="center"))
        xml.append(PAGE_BREAK)
        return "".join(xml)
    xml.append('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
               '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
               '<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>')
    for level, text in headers:
        indent = (level - 1) * 360
        xml.append('<w:p><w:pPr><w:spacing w:after="60"/>'
                   '<w:ind w:left="%d" w:firstLine="0" w:firstLineChars="0"/></w:pPr>'
                   '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                   % (indent, esc(text)))
    xml.append('<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    # 目录与正文分页：目录结束后插入分页符，正文另起一页
    xml.append(PAGE_BREAK)
    return "".join(xml)


_DISPLAY_MATH_RE = re.compile(r"\\\[(.*?)\\\]", re.S)


def _normalize_display_math(md):
    """把 LaTeX 显示公式 \\[...\\] 归一化为 md2docx 支持的公式标记。

    - 独占行的 \\[ ... \\]（块级，可跨行）-> $$ ... $$
    - 行内出现的 \\[ ... \\] -> $ ... $（行内公式）
    """
    def repl(m):
        start = m.start()
        line_start = md.rfind("\n", 0, start) + 1
        if md[line_start:start].strip() == "":
            return "$$" + m.group(1) + "$$"
        return "$" + m.group(1) + "$"
    return _DISPLAY_MATH_RE.sub(repl, md)


def convert_markdown(md, stats):
    md = _normalize_display_math(md)
    lines = md.splitlines()
    body = []
    headers = _collect_headers(md)
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()

        if stripped == "":
            i += 1
            continue

        # 目录指令（不自动插入；只有显式 [[TOC]] / [TOC] 才生成）
        if stripped.upper() in ("[[TOC]]", "[TOC]"):
            body.append(toc_block(
                headers,
                page_break_before=bool(body) and not _ends_with_page_break(body)))
            i += 1
            continue

        # 代码块
        if stripped.startswith("```"):
            i += 1
            code_lines = []
            while i < n and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            i += 1
            for cl in code_lines:
                body.append(para(run(cl if cl else " ", code=True), style="Code"))
            continue

        # 块级公式 $$...$$
        if stripped.startswith("$$"):
            first = stripped[2:]
            if first.endswith("$$") and len(first) > 2:
                tex = first[:-2]
                i += 1
            else:
                buf = [first] if first else []
                i += 1
                while i < n and "$$" not in lines[i]:
                    buf.append(lines[i])
                    i += 1
                if i < n:
                    buf.append(lines[i].split("$$")[0])
                    i += 1
                tex = "\n".join(buf)
            try:
                body.append(para(omath_para(tex), align="center",
                                 ppr_extra='<w:ind w:firstLine="0" w:firstLineChars="0"/>'))
                stats["omml"] += 1
            except Exception:
                stats["failed"].append(tex)
                body.append(para(run("[公式转换失败] " + tex, code=True)))
            continue

        # 标题
        m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if m:
            level = len(m.group(1))
            body.append(para(inline(m.group(2), stats), style="Heading%d" % min(level, 3)))
            i += 1
            continue

        # 分隔线
        if re.match(r"^(-{3,}|\*{3,}|_{3,})$", stripped):
            body.append(hr())
            i += 1
            continue

        # 引用
        if stripped.startswith(">"):
            quote_lines = []
            while i < n and lines[i].strip().startswith(">"):
                quote_lines.append(re.sub(r"^>\s?", "", lines[i].strip()))
                i += 1
            quote_text = "\n".join(quote_lines)
            # 引用块内的块级公式 $$...$$ 必须先拆出来；否则 $$ 两端会被拆进
            # 相邻文本段落（表现为前段末尾/后段开头多出一个 $）。
            for k, seg in enumerate(re.split(r"\$\$(.+?)\$\$", quote_text, flags=re.S)):
                if k % 2 == 1:
                    tex = seg.strip()
                    try:
                        body.append(para(omath_para(tex), align="center",
                                         ppr_extra='<w:ind w:firstLine="0" w:firstLineChars="0"/>'))
                        stats["omml"] += 1
                    except Exception:
                        stats["failed"].append(tex)
                        body.append(para(run("[公式转换失败] " + tex, code=True)))
                elif seg.strip():
                    body.extend(emit_text_block(" ".join(seg.split("\n")),
                                                stats, style="Quote"))
            continue

        # 表格
        if stripped.startswith("|"):
            raw_rows = []
            while i < n and lines[i].strip().startswith("|"):
                raw_rows.append(lines[i].strip())
                i += 1
            rows = []
            for idx, raw in enumerate(raw_rows):
                cells = [c.strip() for c in raw.strip("|").split("|")]
                if idx == 1 and all(re.match(r"^:?-{2,}:?$", c) for c in cells if c):
                    continue
                rows.append(cells)
            if rows:
                body.append(make_table(rows, stats))
                body.append(para(""))
            continue

        # 列表
        m = re.match(r"^\s*[-*+]\s+(.*)$", line)
        if m:
            body.extend(emit_media_block(
                m.group(1), stats,
                ppr_extra='<w:ind w:left="720" w:hanging="360" w:firstLine="0" w:firstLineChars="0"/>',
                prefix="• "))
            i += 1
            continue
        m = re.match(r"^\s*(\d+)\.\s+(.*)$", line)
        if m:
            body.extend(emit_media_block(
                m.group(2), stats,
                ppr_extra='<w:ind w:left="720" w:hanging="360" w:firstLine="0" w:firstLineChars="0"/>',
                prefix=m.group(1) + ". "))
            i += 1
            continue

        # 普通段落
        buf = [stripped]
        i += 1
        while i < n and lines[i].strip() != "" and not is_block_start(lines[i]):
            buf.append(lines[i].strip())
            i += 1
        joined = " ".join(buf)
        is_caption = (bool(body) and "<w:drawing>" in body[-1]
                      and _CAPTION_RE.match(joined.strip()))
        if is_caption:
            body.append(caption_para(joined, stats))
        else:
            body.extend(emit_media_block(joined, stats))

    if not body:
        body.append(para(run("")))
    body.append('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
                'w:left="1440" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>')
    return body


# ---------------- DOCX 打包 ----------------
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>
    <w:pPr><w:ind w:firstLineChars="200" w:firstLine="480"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:spacing w:before="240" w:after="240"/><w:ind w:firstLine="0" w:firstLineChars="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:ind w:firstLine="0" w:firstLineChars="0"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="200" w:after="100"/><w:ind w:firstLine="0" w:firstLineChars="0"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:ind w:firstLine="0" w:firstLineChars="0"/><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="CCCCCC"/></w:pBdr><w:ind w:left="720" w:firstLine="0" w:firstLineChars="0"/></w:pPr>
    <w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F7F7F7"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0" w:firstLineChars="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="TableHead"><w:name w:val="Table Head"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:firstLine="0" w:firstLineChars="0"/></w:pPr>
    <w:rPr><w:b/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:spacing w:after="120"/><w:ind w:firstLine="0" w:firstLineChars="0"/><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/><w:color w:val="595959"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="VerbatimChar"><w:name w:val="Verbatim Char"/>
    <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr></w:style>
</w:styles>"""

def content_types_xml():
    defaults = []
    for _rid, name, _data in _MEDIA:
        ext = name.rsplit(".", 1)[-1].lower()
        if ext == "png":
            line = '<Default Extension="png" ContentType="image/png"/>'
        elif ext in ("jpeg", "jpg"):
            line = '<Default Extension="jpeg" ContentType="image/jpeg"/>'
        elif ext == "gif":
            line = '<Default Extension="gif" ContentType="image/gif"/>'
        else:
            continue
        if line not in defaults:
            defaults.append(line)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + "".join(defaults) +
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            "</Types>")

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""

def document_rels_xml():
    rels = ['<Relationship Id="rIdStyles" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>']
    for rid, name, _data in _MEDIA:
        rels.append('<Relationship Id="%s" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                    'Target="media/%s"/>' % (rid, name))
    for rid, url in _EXTERNAL_RELS:
        rels.append('<Relationship Id="%s" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                    'Target="%s" TargetMode="External"/>'
                    % (rid, escape(url, {'"': "&quot;"})))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(rels) + "</Relationships>")


def document_xml(body_xml):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:document xmlns:w="%s" xmlns:m="%s" xmlns:r="%s" '
            'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            "<w:body>%s</w:body></w:document>" % (W_NS, M_NS, R_NS, body_xml))


def core_xml(title):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:title>%s</dc:title><dc:creator>md2docx</dc:creator>"
            '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>'
            "</cp:coreProperties>" % esc(title))


APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>md2docx</Application></Properties>"""


def build_docx(path, body_xml, title):
    parts = [
        ("[Content_Types].xml", content_types_xml()),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document_xml(body_xml)),
        ("word/styles.xml", STYLES_XML),
        ("word/_rels/document.xml.rels", document_rels_xml()),
        ("docProps/core.xml", core_xml(title)),
        ("docProps/app.xml", APP_XML),
    ]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in parts:
            zf.writestr(name, content.encode("utf-8"))
        for _rid, media_name, data in _MEDIA:
            zf.writestr("word/media/" + media_name, data)
    return parts


def _tool_argv():
    """本机 in-process 解释器把 sys.argv 设为 [python3, 脚本, 参数...]。"""
    argv = list(sys.argv)
    if len(argv) >= 2 and argv[1].endswith(".py"):
        return argv[2:]
    if argv and argv[0].endswith(".py"):
        return argv[1:]
    return argv


def _do_audit(path):
    """检查 .docx 的发布风险：元数据、外部链接、批注、修订等。"""
    if not os.path.isfile(path):
        print("文件不存在: " + path, file=sys.stderr)
        return 2
    try:
        zf = zipfile.ZipFile(path)
        names = zf.namelist()
        root = ET.fromstring(zf.read("word/document.xml"))
        W = "{%s}" % W_NS
        NS = {"dc": "http://purl.org/dc/elements/1.1/",
              "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"}
        creator = lastmod = title = ""
        if "docProps/core.xml" in names:
            core = ET.fromstring(zf.read("docProps/core.xml"))
            creator = core.findtext("dc:creator", default="", namespaces=NS) or ""
            lastmod = core.findtext("cp:lastModifiedBy", default="", namespaces=NS) or ""
            title = core.findtext("dc:title", default="", namespaces=NS) or ""
        rels = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        links = [r.get("Target") for r in rels
                 if r.get("TargetMode") == "External" and r.get("Target")]
        media = [n for n in names if n.startswith("word/media/")]
        comments = "word/comments.xml" in names
        revisions = bool(root.findall(".//" + W + "ins")) or bool(root.findall(".//" + W + "del"))
    except Exception as exc:
        print("审计失败: %s" % exc, file=sys.stderr)
        return 2
    print("发布风险检查: " + path)
    print("  元数据: creator=%r, lastModifiedBy=%r, title=%r" % (creator, lastmod, title))
    print("  外部超链接: %d" % len(links))
    for link in links[:10]:
        print("    - " + str(link))
    print("  嵌入图片: %d" % len(media))
    print("  批注: %s | 修订标记: %s" % ("有" if comments else "无", "有" if revisions else "无"))
    warnings = []
    if creator and creator.strip().lower() not in ("md2docx", ""):
        warnings.append("creator 含个人/应用名称，分享前建议清除")
    if lastmod.strip():
        warnings.append("lastModifiedBy 非空，分享前建议清除")
    if links:
        warnings.append("含外部链接，确认不含内网地址或令牌")
    if comments or revisions:
        warnings.append("含批注/修订，分享前建议接受或删除")
    print("  风险提示: " + ("；".join(warnings) if warnings else "未发现明显风险"))
    return 0


def _do_self_test():
    out_dir = os.path.join(HERE, "out")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "self-test.docx")
    md = ("# 自检\n\n行内公式 $E = mc^2$。\n\n"
          "$$\n\\int_0^1 x^2\\,dx = \\frac{1}{3}\n$$\n")
    stats = {"omml": 0, "failed": []}
    _reset_generation()
    body = convert_markdown(md, stats)
    build_docx(out, "".join(body), "self-test")
    try:
        zf = zipfile.ZipFile(out)
        root = ET.fromstring(zf.read("word/document.xml"))
        count = len(root.findall(
            ".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))
        if count < 1:
            raise AssertionError("没有生成公式对象")
        print("md2docx " + __version__)
        print("依赖: latex2mathml / mathml2omml OK")
        print("生成自检文档: " + out)
        print("公式对象: %d" % count)
        print("自检通过")
        return 0
    except Exception as exc:
        print("自检失败: %s" % exc, file=sys.stderr)
        return 1


def _residual_commands(body_xml):
    """扫描公式 m:t 中残留的未识别 LaTeX 命令。"""
    texts = re.findall(r"<m:t[^>]*>(.*?)</m:t>", body_xml, re.S)
    return re.findall(r"\\[A-Za-z]+", "".join(texts))


def _do_backup(zip_path):
    home = os.path.dirname(HERE)
    targets = []
    for name in ("md2docx.py", "make_cover.py", "sample.md", "requirements.txt",
                 "README.md", "SKILL.md", "LICENSE"):
        src = os.path.join(HERE, name)
        if os.path.isfile(src):
            targets.append((src, "tools/" + name))
    skill = os.path.join(home, "skills", "md2docx.md")
    if os.path.isfile(skill):
        targets.append((skill, "skills/md2docx.md"))
    libs = os.path.join(home, "pylibs")
    if os.path.isdir(libs):
        for root, dirs, files in os.walk(libs):
            dirs[:] = [d for d in dirs if d != "__pycache__"]
            for filename in files:
                src = os.path.join(root, filename)
                arc = os.path.relpath(src, home).replace(os.sep, "/")
                targets.append((src, arc))
    readme = ("md2docx \u6062\u590d\u8bf4\u660e\n\n"
              "\u5c06\u672c zip \u89e3\u538b\u5230 $DSH_HOME \u4e0b"
              "\uff08\u4fdd\u6301 tools/ \u3001pylibs/ \u3001skills/ \u7ed3\u6784\uff09\uff0c\n"
              "\u7136\u540e\u8fd0\u884c md2docx.py --self-test \u9a8c\u8bc1\u3002\n")
    root_skill = None
    skill_src = os.path.join(HERE, "SKILL.md")
    if os.path.isfile(skill_src):
        with open(skill_src, "r", encoding="utf-8") as fh:
            root_skill = fh.read()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for src, arc in targets:
            zf.write(src, arc)
        zf.writestr("README-RESTORE.txt", readme)
        if root_skill is not None:
            zf.writestr("SKILL.md", root_skill)
    count = len(targets) + 1 + (1 if root_skill is not None else 0)
    print("已打包: %s（%d 个文件，含根目录 SKILL.md）" % (zip_path, count))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Markdown + LaTeX -> DOCX (native OMML formulas)")
    ap.add_argument("input", nargs="?", help="输入 Markdown 文件")
    ap.add_argument("-o", "--output", help="输出 .docx 路径")
    ap.add_argument("--title", default=None, help="文档标题")
    ap.add_argument("--math-font", default=None,
                    help="正文中数学符号使用的字体（默认 Cambria Math）")
    ap.add_argument("--font-size-pt", type=float, default=12.0,
                    help="正文字号（磅）；默认 12 = 小四")
    ap.add_argument("--version", action="store_true", help="打印版本")
    ap.add_argument("--self-test", action="store_true",
                    help="自检依赖并生成最小 docx")
    ap.add_argument("--backup", default=None, metavar="ZIP",
                    help="把 tools/pylibs/skills 打包成恢复用 zip")
    ap.add_argument("--audit", default=None, metavar="DOCX",
                    help="检查 .docx 的发布风险（元数据/链接/批注/修订）")
    ap.add_argument("--strict", action="store_true",
                    help="存在未识别的 LaTeX 命令时返回非零退出码")
    ap.add_argument("--cover", default=None, metavar="IMAGE",
                    help="可选：封面图片路径；封面单独成页")
    args = ap.parse_args(_tool_argv() if argv is None else argv)

    if args.version:
        print("md2docx " + __version__)
        return 0
    if args.audit:
        return _do_audit(args.audit)
    if args.backup:
        return _do_backup(args.backup)
    if args.self_test:
        return _do_self_test()
    if not args.input or not args.output:
        ap.error("需要输入 Markdown 文件和 -o 输出路径")

    global MATH_FONT, STYLES_XML
    if args.math_font:
        MATH_FONT = args.math_font
    if args.font_size_pt and abs(args.font_size_pt - 12.0) > 1e-6:
        half = int(round(args.font_size_pt * 2))
        STYLES_XML = STYLES_XML.replace('<w:sz w:val="24"/>', '<w:sz w:val="%d"/>' % half)
        STYLES_XML = STYLES_XML.replace('<w:szCs w:val="24"/>', '<w:szCs w:val="%d"/>' % half)

    _BASE_DIR[0] = os.path.dirname(os.path.abspath(args.input))
    _reset_generation()
    with open(args.input, "r", encoding="utf-8") as fh:
        md = fh.read()

    stats = {"omml": 0, "failed": []}
    body = convert_markdown(md, stats)
    if args.cover:
        cover_path = args.cover if os.path.isabs(args.cover) else os.path.join(_BASE_DIR[0], args.cover)
        if os.path.isfile(cover_path):
            body = [para(_image_run(cover_path, "封面"), align="center",
                         ppr_extra='<w:ind w:firstLine="0" w:firstLineChars="0"/>'),
                    PAGE_BREAK] + body
        else:
            print("[warn] 封面文件不存在: " + args.cover, file=sys.stderr)
    title = args.title or os.path.splitext(os.path.basename(args.output))[0]
    body_xml = "".join(body)
    build_docx(args.output, body_xml, title)

    print("已生成: " + args.output)
    print("公式(OMML): %d，图片: %d，超链接: %d，转换失败: %d"
          % (stats["omml"], len(_MEDIA), len(_EXTERNAL_RELS), len(stats["failed"])))
    for tex in stats["failed"]:
        print("  失败公式: " + tex[:80], file=sys.stderr)
    residual = _residual_commands(body_xml)
    print("公式残留命令: %d" % len(residual))
    if residual:
        for cmd in sorted(set(residual))[:10]:
            print("  [warn] 未识别命令: " + cmd, file=sys.stderr)
        if args.strict:
            print("--strict: 存在未识别的 LaTeX 命令，按失败处理", file=sys.stderr)
            return 2
    return 0 if not stats["failed"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
