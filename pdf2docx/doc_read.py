#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""doc_read.py — 按需提取 PDF / PPTX / PPT / DOC 文本（纯标准库，零依赖）

用法:
    python3 doc_read.py <文件> [输出文件|-] [页码，如 1-5,8]
不带输出文件时打印到 stdout。带缓存：同一文件（路径+大小+mtime）二次读取毫秒级。

设计要点:
  * PPTX 只解压 slide XML，跳过全部媒体（188MB 的课件也只读约 2MB）。
  * PDF 走 对象表 → FlateDecode 解压 → ToUnicode CMap → 内容流文本，
    文本层抽取用 bytes.find/struct.unpack 等 C 级原语，避免逐字符 Python 循环。
  * PDF 1.5+ 的对象流（/Type /ObjStm）会被展开并入对象表（见 _expand_objstm）；
    没有 ToUnicode 的文档（扫描件的 OCR 文本层）按 /Pages 页树取页，
    并解码字面串 `(...) Tj`，页号与阅读顺序一致。
  * 简单字体的字面串按**当前字体的 /Encoding** 解码（WinAnsi/MacRoman/Standard，
    字典形式支持 /BaseEncoding + /Differences；/Resources /Font 与 /Encoding
    都可以是间接引用或在对象流里）。字面串的字节是字符代码，不是 UTF-8。
  * OCR 文本层常把一行拆成多个换行的片段，换行处按 Tm/Td/TD/T* 的纵向位移补
    回词间空格（连字符断词不补）；同一行的词自带前导空格，不会重复补。
  * 全程只用标准库（zlib/struct/zipfile/re/hashlib），不需要 pip、不需要联网。
"""
import sys, os, re, zlib, json, html, time, struct, zipfile, hashlib, bisect, tempfile

WS = b' \t\r\n\f'
HEX_RE = re.compile(rb'<([0-9A-Fa-f\s]+)>')
TF_RE = re.compile(rb'/([A-Za-z0-9#+._-]+)\s+[-\d.]+\s+Tf')
NONHEX_RE = re.compile(rb'[^0-9A-Fa-f\s]')
SLIDE_RE = re.compile(r'^ppt/slides/slide(\d+)\.xml$')
FONTDICT_RE = re.compile(rb'/Font\s*<<(.*?)>>', re.S)
FONTREF_RE = re.compile(rb'/([A-Za-z0-9#+._-]+)\s+(\d+)\s+0\s+R')
TOUNI_RE = re.compile(rb'/ToUnicode\s+(\d+)\s+0\s+R')
PAGE_SPEC_RE = re.compile(r'^(\d+)\s*(?:-\s*(\d+))?$')
# 兜底用：对象头正则。Linearized / 交叉引用流结构的 PDF 会让顺序扫描拿不到对象。
_OBJ_HEAD_RE = re.compile(rb'(?<![0-9])(\d+)\s+(\d+)\s+obj\b')
LIT_STR_RE = re.compile(rb'\(((?:\\.|[^\\()])*)\)', re.S)
# TJ 数组的单个 token：<hex> / (字面串) / 数字（千分之一 em 的字距调整）。
# 数字为负且绝对值足够大时表示 TeX 的词间空格，必须还原成空格；
# 否则整篇词粘连（"rotationalbroadening"），无法 grep、引用也会失真。
TJ_TOKEN_RE = re.compile(rb'<([0-9A-Fa-f\s]*)>|\(((?:\\.|[^\\()])*)\)|(-?\d+(?:\.\d+)?)', re.S)
TJ_SPACE_ADV = -120.0          # 字距 <= 此值视为词间空格（TeX 常用 -167/-250/-333）


# ──────────────────────────────── PPTX ────────────────────────────────
def pptx_pages(path):
    z = zipfile.ZipFile(path)
    items = []
    for nm in z.namelist():
        m = SLIDE_RE.match(nm)
        if m:
            items.append((int(m.group(1)), nm))
    items.sort()
    pages = []
    for num, nm in items:
        try:
            xml = z.read(nm).decode('utf-8', 'ignore')
        except Exception:
            continue
        chunks = re.findall(r'<a:t>(.*?)</a:t>', xml, re.S)
        txt = html.unescape(' '.join(chunks))
        pages.append((num, re.sub(r'\s+', ' ', txt).strip()))
    return pages


# ──────────────────────────────── PDF ────────────────────────────────
def _positions(buf, kw, skip_prefixed_by_end=False):
    out = []
    p = 0
    while True:
        i = buf.find(kw, p)
        if i < 0:
            return out
        p = i + 1
        if skip_prefixed_by_end and i >= 3 and buf[i - 3:i] == b'end':
            continue
        out.append(i)


def pdf_objects(data):
    """线性扫描对象表：每个 N 0 obj 的字节区间。O(n log m)，无回溯。"""
    objs = {}
    endobjs = _positions(data, b'endobj')
    streams = _positions(data, b'stream', True)
    endstreams = _positions(data, b'endstream')
    n = len(data)
    i = 0
    while i < n:
        j = data.find(b'obj', i)
        if j < 0:
            break
        i = j + 3
        if j + 3 >= n or data[j + 3] not in WS:
            continue
        k = j
        while k > 0 and data[k - 1] in WS:
            k -= 1
        gen_end = k
        while k > 0 and 48 <= data[k - 1] <= 57:
            k -= 1
        if k == gen_end:
            continue
        while k > 0 and data[k - 1] in WS:
            k -= 1
        num_end = k
        while k > 0 and 48 <= data[k - 1] <= 57:
            k -= 1
        if k == num_end:
            continue
        try:
            num = int(data[k:num_end])
        except ValueError:
            continue
        start = j + 3
        a = bisect.bisect_left(endobjs, start)
        hi = endobjs[a] if a < len(endobjs) else n
        b = bisect.bisect_left(streams, start)
        js = -1
        if b < len(streams) and streams[b] < hi:
            cand = streams[b]
            if data[cand + 6:cand + 7] in (b'\r', b'\n'):
                js = cand
        if js >= 0:
            d = bisect.bisect_left(endstreams, js)
            jend = endstreams[d] if d < len(endstreams) else hi
            objs[num] = data[start:jend]
        else:
            objs[num] = data[start:hi]
    _expand_objstm(objs)
    # 兜底：Linearized / 交叉引用流（PDF-1.6）结构下，顺序扫描可能只拿到极少数对象
    # 甚至 0 个（实测两本 MATLAB 教材：0 个 / 1 个）。判据用「对象太少或没有 Catalog」，
    # 再按 "N G obj … endobj" 对象头正则收集（实测恢复到 4856 / 2761 个对象）。
    if len(objs) < 50 or not any(b"/Catalog" in b[:800] for b in objs.values()):
        added = 0
        for m in _OBJ_HEAD_RE.finditer(data):
            start = m.end()
            end = data.find(b"endobj", start)
            if end < 0:
                continue
            if int(m.group(1)) not in objs:
                objs[int(m.group(1))] = data[start:end]
                added += 1
        if added:
            _expand_objstm(objs)
    return objs


def _expand_objstm(objs):
    """把 /Type /ObjStm 对象流内的压缩对象并入对象表（PDF 1.5+）。

    PDF 1.5+ 常把 /Pages、/Resources、字体字典及其 /ToUnicode 引用压进对象流；
    不展开就读不到页树与 CMap（表现为 pages=0）。对象流内不能含流对象，
    所以内部对象都是普通字典。

    对象流格式: 解压后前 /First 字节是 N 组 "objnum offset"（offset 相对 /First），
    其后依次是各对象的数据，对象之间没有 endobj，只能按 offset 切分。

    优先级: 同一对象号若已有直接对象，保留直接对象（规范要求），对象流仅补缺。
    """
    added = 0
    for onum in list(objs):
        body = objs[onum]
        head = body[:1024]
        if b'/ObjStm' not in head:
            continue
        mn = re.search(rb'/N\s+(\d+)', head)
        mf = re.search(rb'/First\s+(\d+)', head)
        if not (mn and mf):
            continue
        dec = _inflate(body)
        if not dec:
            continue
        try:
            n, first = int(mn.group(1)), int(mf.group(1))
        except ValueError:
            continue
        if n <= 0 or first < 0 or first > len(dec):
            continue
        pairs = re.findall(rb'(\d+)\s+(\d+)', dec[:first])
        if len(pairs) < n:                        # 头部异常时退回全流搜索
            pairs = re.findall(rb'(\d+)\s+(\d+)', dec)[:n]
        pairs = pairs[:n]
        if not pairs:
            continue
        bounds = []
        for num_s, off_s in pairs:
            try:
                bounds.append((int(num_s), first + int(off_s)))
            except ValueError:
                bounds.append((None, None))
        for i, (num, start) in enumerate(bounds):
            if num is None or start is None or num in objs or start >= len(dec):
                continue
            end = len(dec)
            for j in range(i + 1, len(bounds)):
                if bounds[j][1] is not None:
                    end = min(bounds[j][1], len(dec))
                    break
            if start < end:
                objs[num] = dec[start:end]
                added += 1
    return added


def _stream_body(body):
    j = body.find(b'stream')
    if j < 0:
        return None
    s = j + 6
    if body[s:s + 2] == b'\r\n':
        s += 2
    elif body[s:s + 1] in (b'\n', b'\r'):
        s += 1
    return body[s:]


def _inflate(body):
    raw = _stream_body(body)
    if raw is None:
        return None
    if b'/FlateDecode' not in body:
        return raw
    for cand in (raw.rstrip(b'\r\n'), raw):
        try:
            out = zlib.decompressobj().decompress(cand)
            if out:
                return out
        except Exception:
            pass
    return None


def _hex2uni(h):
    try:
        return bytes.fromhex(h.decode('ascii')).decode('utf-16-be')
    except Exception:
        return ''


_BFRANGE_RE = re.compile(
    rb'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(?:\[(.*?)\]|<([0-9A-Fa-f]+)>)', re.S)


def parse_cmap(data):
    """解析 ToUnicode CMap。

    bfrange 有两种写法：数组 <lo> <hi> [<u1> <u2> ...] 与递增 <lo> <hi> <u0>。
    必须用一个正则按出现顺序整体匹配、数组优先。旧实现先用数组正则 findall、
    再用三元组正则 findall 同一个块，会把数组里的连续三元素（如 <0066><006F>
    <0072> 这类 "for" 的码）误当成 <lo><hi><dst> 展开成上万条范围映射，
    覆盖掉正确映射 —— 正文出现「本章回顾 → 星映昡昢」这类系统性错字。
    """
    cmap = {}
    for m in re.finditer(rb'beginbfrange(.*?)endbfrange', data, re.S):
        for lo, hi, arr, dst in _BFRANGE_RE.findall(m.group(1)):
            a = int(lo, 16)
            if arr:                      # findall 的未匹配分组是 b''，不是 None
                for idx, it in enumerate(re.findall(rb'<([0-9A-Fa-f]+)>', arr)):
                    cmap[a + idx] = _hex2uni(it)
            else:
                b, base = int(hi, 16), int(dst, 16)
                for k in range(a, b + 1):
                    cmap[k] = _hex2uni(('%04X' % (base + k - a)).encode())
    for m in re.finditer(rb'beginbfchar(.*?)endbfchar', data, re.S):
        for src, dst in re.findall(rb'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', m.group(1)):
            cmap[int(src, 16)] = _hex2uni(dst)
    return cmap


def cid_text(h, cmap):
    """把 <hex> 串按 CID 批量查表；用 struct.unpack 一次性解出全部 CID。"""
    h = re.sub(rb'\s+', b'', h)
    if not h:
        return ''
    get = cmap.get
    if len(h) % 4 == 0:
        try:
            raw = bytes.fromhex(h.decode('ascii'))
        except Exception:
            return ''
        if len(raw) % 2 == 0 and raw:
            cids = struct.unpack('>%dH' % (len(raw) // 2), raw)
            return ''.join([get(c, '') for c in cids])
    return ''.join([get(int(h[i:i + 2], 16), '') for i in range(0, len(h) - 1, 2)])


def _lit_unescape(raw):
    """解 PDF 字面串的转义（\\n \\r \\t \\b \\f \\( \\) \\\\ \\ddd 与行尾续行）。"""
    if b'\\' not in raw:                     # 绝大多数串没有转义，直接返回
        return raw
    out = bytearray()
    i, n = 0, len(raw)
    while i < n:
        c = raw[i]
        if c == 0x5c and i + 1 < n:
            d = raw[i + 1]
            mp = {0x6e: 10, 0x72: 13, 0x74: 9, 0x62: 8, 0x66: 12}
            if d in mp:
                out.append(mp[d]); i += 2
            elif 0x30 <= d <= 0x37:                 # 1-3 位八进制
                j, v = i + 1, 0
                while j < n and j < i + 4 and 0x30 <= raw[j] <= 0x37:
                    v = v * 8 + (raw[j] - 0x30); j += 1
                out.append(v & 0xff); i = j
            elif d in (0x0a, 0x0d):                 # 反斜杠续行
                i += 2
                if d == 0x0d and i < n and raw[i] == 0x0a:
                    i += 1
            else:
                out.append(d); i += 2
        else:
            out.append(c); i += 1
    return bytes(out)


# ── 简单字体的 /Encoding：字符代码 → 字形名 → Unicode ────────────────────
# PDF 字面串里的字节是**字符代码**，必须按当前字体的 /Encoding 解释：
#   /Encoding 为名字时（WinAnsiEncoding / MacRomanEncoding / StandardEncoding）
#   直接用对应表；为字典时先取 /BaseEncoding，再用 /Differences 覆盖。
_GLYPH = {
    'space': ' ', 'exclam': '!', 'quotedbl': '"', 'numbersign': '#', 'dollar': '$',
    'percent': '%', 'ampersand': '&', 'quotesingle': "'", 'parenleft': '(',
    'parenright': ')', 'asterisk': '*', 'plus': '+', 'comma': ',', 'hyphen': '-',
    'period': '.', 'slash': '/', 'zero': '0', 'one': '1', 'two': '2', 'three': '3',
    'four': '4', 'five': '5', 'six': '6', 'seven': '7', 'eight': '8', 'nine': '9',
    'colon': ':', 'semicolon': ';', 'less': '<', 'equal': '=', 'greater': '>',
    'question': '?', 'at': '@', 'bracketleft': '[', 'backslash': '\\',
    'bracketright': ']', 'asciicircum': '^', 'underscore': '_', 'grave': '`',
    'braceleft': '{', 'bar': '|', 'braceright': '}', 'asciitilde': '~',
    'exclamdown': '¡', 'cent': '¢', 'sterling': '£', 'fraction': '⁄', 'yen': '¥',
    'florin': 'ƒ', 'section': '§', 'currency': '¤', 'quotedblleft': '“',
    'guillemotleft': '«', 'guilsinglleft': '‹', 'guilsinglright': '›',
    'fi': 'ﬁ', 'fl': 'ﬂ', 'ff': 'ﬀ', 'ffi': 'ﬃ', 'ffl': 'ﬄ', 'endash': '–',
    'dagger': '†', 'daggerdbl': '‡', 'periodcentered': '·', 'paragraph': '¶',
    'bullet': '•', 'quotesinglbase': '‚', 'quotedblbase': '„', 'quotedblright': '”',
    'guillemotright': '»', 'ellipsis': '…', 'perthousand': '‰', 'questiondown': '¿',
    'acute': '´', 'circumflex': 'ˆ', 'tilde': '˜', 'macron': '¯', 'breve': '˘',
    'dotaccent': '˙', 'dieresis': '¨', 'ring': '˚', 'cedilla': '¸',
    'hungarumlaut': '˝', 'ogonek': '˛', 'caron': 'ˇ', 'emdash': '—', 'AE': 'Æ',
    'ordfeminine': 'ª', 'Lslash': 'Ł', 'Oslash': 'Ø', 'OE': 'Œ',
    'ordmasculine': 'º', 'ae': 'æ', 'dotlessi': 'ı', 'lslash': 'ł', 'oslash': 'ø',
    'oe': 'œ', 'germandbls': 'ß', 'onesuperior': '¹', 'logicalnot': '¬',
    'mu': 'µ', 'trademark': '™', 'Eth': 'Ð', 'onehalf': '½', 'plusminus': '±',
    'Thorn': 'Þ', 'onequarter': '¼', 'divide': '÷', 'brokenbar': '¦',
    'degree': '°', 'thorn': 'þ', 'threequarters': '¾', 'twosuperior': '²',
    'registered': '®', 'minus': '−', 'eth': 'ð', 'multiply': '×',
    'threesuperior': '³', 'copyright': '©', 'nbspace': '\u00a0',
    'sfthyphen': '\u00ad', 'Euro': '€', 'Scaron': 'Š', 'scaron': 'š',
    'Zcaron': 'Ž', 'zcaron': 'ž', 'Ydieresis': 'Ÿ', 'lozenge': '◊',
    'quoteleft': '‘', 'quoteright': '’', 'notdef': '',
}


def _glyph_char(nm):
    """字形名 → Unicode 字符（支持 /A、/l.alt、/uni00E9、/u1F600）。"""
    base = nm.split('.')[0]
    if not base:
        return ''
    if len(base) == 1:
        return base
    ch = _GLYPH.get(base)
    if ch is not None:
        return ch
    try:
        if base.startswith('uni') and len(base) >= 7:
            return chr(int(base[3:7], 16))
        if base.startswith('u') and 5 <= len(base) <= 7:
            return chr(int(base[1:], 16))
    except ValueError:
        pass
    return ''


def _base_enc(name):
    """BaseEncoding 名字 → {代码: 字符}。未知名按 Latin-1 处理（历史行为）。"""
    try:
        nm = name.decode('latin1')
    except AttributeError:
        nm = name
    out = {}
    if nm == 'MacRomanEncoding':
        for b in range(32, 256):
            out[b] = bytes([b]).decode('mac_roman', 'ignore')
        return out
    for b in range(32, 256):
        try:
            out[b] = bytes([b]).decode('cp1252')
        except UnicodeDecodeError:
            out[b] = bytes([b]).decode('latin1')
    if nm == 'WinAnsiEncoding':
        out[0xA0] = ' '                     # PDF: WinAnsi 0xA0 是空格而非 NBSP
        out[0xAD] = '-'
    return out


diff_tok_re = re.compile(rb'(\d+)\s*|/([^\s/<>\[\]()]+)')


def _enc_from_dict(d):
    """解析 /Encoding 字典：/BaseEncoding + /Differences [65 /A 66 /B ...]。"""
    mb = re.search(rb'/BaseEncoding\s*/(\w+)', d)
    table = dict(_base_enc(mb.group(1) if mb else 'StandardEncoding'))
    md = re.search(rb'/Differences\s*\[(.*?)\]', d, re.S)
    if md:
        code = 0
        for t in diff_tok_re.finditer(md.group(1)):
            if t.group(1):
                code = int(t.group(1))
            else:
                table[code] = _glyph_char(t.group(2).decode('latin1'))
                code += 1
    return table


def _enc_tr(table):
    """{代码: 字符} → str.translate 表（控制码位丢弃，表外码位保留原字节含义）。"""
    tr = {b: '' for b in range(32)}
    for b, ch in table.items():
        if b >= 32:
            tr[b] = ch
    return tr


def _font_encoding(objs, num):
    """字体对象 → str.translate 表；无法确定时返回 None（按 Latin-1 兜底）。"""
    if num is None:
        return None
    body = objs.get(num)
    if not body:
        return None
    head = body[:4096]
    m = re.search(rb'/Encoding\s*(?:/(\w+)|\s*(\d+)\s+\d+\s+R|\s*(<<))', head)
    if not m:
        return None
    if m.group(1):
        return _enc_tr(_base_enc(m.group(1)))
    if m.group(2):
        ed = objs.get(int(m.group(2)))
        return _enc_tr(_enc_from_dict(ed)) if ed else None
    # 直接写在字体字典里的 << ... >>：Differences 内不会出现 >>，可非贪婪匹配
    md = re.search(rb'/Encoding\s*(<<.*?>>)', head, re.S)
    return _enc_tr(_enc_from_dict(md.group(1))) if md else None


def _page_fonts(objs, num):
    """页的 /Resources /Font → {字名: 字体对象号}（Resources/Font 可为间接引用）。"""
    body = objs.get(num)
    depth = 0
    while body is not None and b'/Resources' not in body and depth < 8:
        mp = re.search(rb'/Parent\s+(\d+)\s+\d+\s+R', body)   # 资源可从页树父节点继承
        if not mp:
            break
        body = objs.get(int(mp.group(1)))
        depth += 1
    if body is None:
        return {}
    mr = re.search(rb'/Resources\s+(\d+)\s+\d+\s+R', body)
    res = objs.get(int(mr.group(1))) if mr else body
    if not res:
        return {}
    mf = re.search(rb'/Font\s+(\d+)\s+\d+\s+R', res)
    if mf:
        # /Font 为间接引用时，目标对象本身就是「字名 → 字体对象号」字典，
        # 不能再去找 /Font <<。旧版此处对间接引用必然返回空字典，
        # 导致 arXiv 老论文（2000s LaTeX PDF）字体解析失败、整篇只剩兜底乱码。
        fdict = objs.get(int(mf.group(1)), b'')
    else:
        md = re.search(rb'/Font\s*<<(.*?)>>', res, re.S)
        fdict = md.group(1) if md else b''
    if not fdict:
        return {}
    return {k.decode('latin1'): int(v)
            for k, v in re.findall(rb'/([A-Za-z0-9#+._-]+)\s+(\d+)\s+\d+\s+R', fdict)}


def _decode_lit(raw, tr):
    """字面串字节按字体 Encoding 解码；tr 为 str.translate 表，None 退回 Latin-1。"""
    s = raw.decode('latin1')
    return s.translate(tr) if tr else s


def _lit_before(buf, lo, k):
    """buf[k]==')' 时返回其配对字面串起始 '(' 的位置（跳过被转义的 '('）。"""
    lt = buf.rfind(b'(', lo, k)
    while lt >= 0:
        bs = 0
        j = lt - 1
        while j >= lo and buf[j] == 0x5c:
            bs += 1
            j -= 1
        if not (bs & 1):
            return lt
        lt = buf.rfind(b'(', lo, lt)
    return -1


_QUOTE_OP_RE = re.compile(rb'\)[ \t\r\n]*[\'"]')


def _line_events(buf):
    """预先扫出全部纵向定位算子 → [(位置, 类型, 数值)]，类型 0=Tm(y) 1=Td/TD(ty) 2=T*·'·"。

    简单字体的 OCR 文本层里，同一行的词各自带显式 Tm（y 相同），换行则由 Tm 的
    y 跳变 / Td·TD 的纵向位移 / T* 表示；行内分块之间一般没有定位算子
    （词中截断，如 "( characteristi)" + "(c)"）。

    只对确实出现的算子名（bytes.find）回读数字，等价于一次线性扫描；不做整段
    数字正则——那会在十六进制串的长数字串上退化成 O(n²)。
    """
    ev = []
    for kw, kind in ((b'Tm', 0), (b'Td', 1), (b'TD', 1)):
        p = 0
        while True:
            i = buf.find(kw, p)
            if i < 0:
                break
            p = i + 2
            nxt = buf[i + 2:i + 3]
            if nxt and (nxt.isalnum() or nxt == 0x2a):      # Tm…/Td0 之类不是算子
                continue
            if i < 2 or buf[i - 1] not in WS:
                continue
            parts = buf[i - 100:i].split()
            try:
                if kind == 0:
                    y = float(parts[-1])
                    d = float(parts[-3])            # Tm 的 d（纵向缩放 ≈ 字号）
                else:
                    float(parts[-2])                # 至少要有 tx ty 两个数
                    y = float(parts[-1])
                    d = None
            except (ValueError, IndexError):
                continue
            ev.append((i, kind, y, d))
    p = 0
    while True:
        i = buf.find(b'T*', p)
        if i < 0:
            break
        ev.append((i, 2, 0.0, None))
        p = i + 2
    for m in _QUOTE_OP_RE.finditer(buf):
        ev.append((m.start(), 2, 0.0, None))
    ev.sort()
    return ev


def _need_space(out, chunk):
    """换行处补一个空格：上一段末尾与下一段开头都没有空白、且不是连字符断词。"""
    if not out or not chunk:
        return False
    tail = out[-1]
    if not tail or tail[-1] in ' \t\r\n\f':
        return False
    if chunk[0] in ' \t\r\n\f':
        return False
    return not tail.endswith('-')


def stream_text(buf, get_cmap, get_enc=None, get_kind=None):
    fonts = [(m.start(), m.group(1).decode('latin1')) for m in TF_RE.finditer(buf)]
    ops = []
    p = 0
    while True:
        i = buf.find(b'Tj', p)
        if i < 0:
            break
        ops.append((i, 0))
        p = i + 2
    p = 0
    while True:
        i = buf.find(b'TJ', p)
        if i < 0:
            break
        ops.append((i, 1))
        p = i + 2
    ops.sort()
    events = _line_events(buf) if get_enc else []      # 只有简单字体路径才需要换行判定
    nev = len(events)
    ei = 0
    out = []
    cur = None
    fi = 0
    nf = len(fonts)
    pending_nl = False
    last_y = None
    last_scale = None
    last_cmap_for = None
    cmap = None
    prev_lt = prev_k = -1
    for pos, kind in ops:
        while fi < nf and fonts[fi][0] < pos:
            cur = fonts[fi][1]
            fi += 1
        # 消费本次显示操作之前的定位算子（落在上一条字面串里的视为误匹配，忽略）
        while ei < nev and events[ei][0] < pos:
            epos, ekind, val, escale = events[ei]
            ei += 1
            if prev_lt >= 0 and prev_lt <= epos <= prev_k:
                continue
            if ekind == 0:                                  # Tm
                if escale is not None and last_scale is not None and escale < 0.85 * last_scale:
                    continue                                # 上/下标：不换行，不更新基线
                thr = 0.55 * escale if escale else 4.0
                if thr < 2.5:
                    thr = 2.5
                if last_y is not None and abs(val - last_y) > thr:
                    pending_nl = True
                last_y = val
                if escale:
                    last_scale = escale
            elif ekind == 1:                                # Td/TD: ty 约 1 个行距
                if abs(val) > 0.6:
                    pending_nl = True
            else:                                           # T* / ' / "
                pending_nl = True
        nl, pending_nl = pending_nl, False      # 换行标志由本次显示操作消费掉
        if cur != last_cmap_for:
            cmap = get_cmap(cur)
            last_cmap_for = cur
        k = pos - 1
        while k >= 0 and buf[k] in WS:
            k -= 1
        if k < 0:
            continue
        lo = k - 9000
        if lo < 0:
            lo = 0
        if kind == 0:
            if buf[k:k + 1] == b'>':
                lt = buf.rfind(b'<', lo, k)
                if lt < 0:
                    continue
                h = buf[lt + 1:k]
                if not h or len(h) > 8000 or NONHEX_RE.search(h):
                    continue
                out.append(cid_text(h, cmap))
            elif buf[k:k + 1] == b')':
                # 简单字体（TrueType/Type1/Type3）的字面串：字节是**单字节字符代码**。
                # 有 ToUnicode 时按 CMap 逐字节解码 —— PPT 导出的课件正是这种：
                #   字体带 ToUnicode、正文却用 (...)Tj 字面串；
                #   旧版把条件写成 `and not cmap`，导致这类文档字面串被整段丢弃。
                # 无 ToUnicode 时退回字体 /Encoding（WinAnsi/MacRoman/Differences）。
                lt = _lit_before(buf, lo, k)
                if lt >= 0:
                    prev_lt, prev_k = lt, k
                    raw = _lit_unescape(buf[lt + 1:k])
                    fkind = get_kind(cur) if get_kind else None
                    if cmap and fkind != "cid":
                        chunk = "".join(cmap.get(b, "") for b in raw)
                    elif cmap and fkind == "cid":
                        chunk = ""      # Type0 用两字节代码，字面串形式罕见，交给 <hex> 分支
                    else:
                        tr = get_enc(cur) if get_enc else None
                        chunk = _decode_lit(raw, tr)
                    if nl and _need_space(out, chunk):
                        out.append(' ')
                    out.append(chunk)
        else:
            if buf[k:k + 1] != b']':
                continue
            lb = buf.rfind(b'[', lo, k)
            if lb < 0:
                continue
            seg = buf[lb + 1:k]
            # 按出现顺序逐 token 处理：<hex> / (字面串) / 数字（字距）。
            # TJ 数组里的 (...) 与 Tj 语义相同：简单字体按单字节查 ToUnicode，
            # 无 ToUnicode 退回字体 /Encoding。旧版只在 cmap 为空时解码字面串，
            # 导致带 ToUnicode 的 TeX/LaTeX 文档（arXiv 论文）整页字面串被丢弃、
            # 只读出 0 字符 —— 与 Tj 分支是同一个坑，当时只修了 Tj。
            fkind = get_kind(cur) if get_kind else None
            tr = get_enc(cur) if get_enc else None
            for m in TJ_TOKEN_RE.finditer(seg):
                if m.group(1) is not None:                  # <hex>：CID 串
                    out.append(cid_text(m.group(1), cmap))
                    nl = False
                elif m.group(2) is not None:                # (...) 字面串
                    raw = _lit_unescape(m.group(2))
                    if cmap and fkind != "cid":
                        chunk = "".join(cmap.get(b, "") for b in raw)
                    elif cmap and fkind == "cid":
                        chunk = ""      # Type0 两字节代码，字面串形式罕见，交给 <hex> 分支
                    else:
                        chunk = _decode_lit(raw, tr)
                    if nl and _need_space(out, chunk):
                        out.append(' ')
                    nl = False
                    out.append(chunk)
                else:                                       # 数字：字距调整
                    try:
                        adv = float(m.group(3))
                    except ValueError:
                        continue
                    # 微调字符间距很少超过 -120；更负的是 TeX 词间空格（-167/-250/-333）。
                    if adv <= TJ_SPACE_ADV and out and not out[-1].endswith(' '):
                        out.append(' ')
    return ''.join(out)


def _page_refs(objs):
    """按 /Pages 页树返回页对象号（顺序即阅读顺序）；找不到页树返回 []。"""
    root = None
    for num, body in objs.items():
        if b'/Catalog' not in body[:400]:
            continue
        m = re.search(rb'/Pages\s+(\d+)\s+\d+\s+R', body)
        if m:
            root = int(m.group(1))
            break
    if root is None:                     # 没有 Catalog 时退化为找无 /Parent 的 /Pages
        for num, body in objs.items():
            if re.search(rb'/Type\s*/Pages', body) and b'/Parent' not in body:
                root = num
                break
    if root is None or root not in objs:
        return []
    pages = []
    seen = set()

    def walk(num, depth):
        if depth > 32 or num in seen or num not in objs:
            return
        seen.add(num)
        body = objs[num]
        if re.search(rb'/Type\s*/Page(?!s)', body):
            pages.append(num)
            return
        mk = re.search(rb'/Kids\s*\[(.*?)\]', body, re.S)
        if not mk:
            return
        for k in re.findall(rb'(\d+)\s+\d+\s+R', mk.group(1)):
            walk(int(k), depth + 1)

    walk(root, 0)
    return pages


def _page_content(objs, num):
    """取页对象 /Contents（单引用或数组）并解压拼接为内容流缓冲。"""
    body = objs[num]
    refs = []
    ma = re.search(rb'/Contents\s*\[(.*?)\]', body, re.S)
    if ma:
        refs = [int(x) for x in re.findall(rb'(\d+)\s+\d+\s+R', ma.group(1))]
    else:
        ms = re.search(rb'/Contents\s+(\d+)\s+\d+\s+R', body)
        if ms:
            refs = [int(ms.group(1))]
    chunks = []
    for r in refs:
        b = objs.get(r)
        if b is None:
            continue
        s = _inflate(b)
        if s:
            chunks.append(s)
    return b'\n'.join(chunks)


def ppt_pages(path):
    """老式 .ppt（PowerPoint 97-2003，OLE2 二进制）→ [(页号, 文本)]。

    解析器是同目录的 ppt_read.py（olefile + MS-PPT 记录流）。返回的是模板/幻灯片的
    可见文本；这类老格式没有可靠的页边界，通常整篇作为 1 页返回。
    """
    import importlib.util
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except NameError:                    # in-process exec 时可能没有 __file__
        here = os.path.dirname(os.path.abspath(sys.argv[0] if sys.argv else "."))
    src = os.path.join(here, "ppt_read.py")
    spec = importlib.util.spec_from_file_location("ppt_read_mod", src)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    pages = mod.extract_pages(path)
    return pages if pages else [(1, "")]


def docbin_pages(path):
    """老式 .doc（Word 97-2003，OLE2 二进制）→ [(页号, 文本)]。

    解析器在同目录 docbin_read.py（olefile + FIB/piece-table，三层回退）。
    Word 旧格式没有可靠页边界，通常整篇作为 1 页返回。
    """
    import importlib.util
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        here = os.path.dirname(os.path.abspath(sys.argv[0] if sys.argv else "."))
    src = os.path.join(here, "docbin_read.py")
    spec = importlib.util.spec_from_file_location("docbin_read_mod", src)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    pages, _meta = mod.extract(path)
    return pages if pages else [(1, "")]


def pdf_pages(path):
    data = open(path, 'rb').read()
    objs = pdf_objects(data)
    label_entries = page_labels(objs)
    font_name = {}
    for body in objs.values():
        if b'/Font' not in body:
            continue
        for blk in FONTDICT_RE.findall(body):
            for m in FONTREF_RE.finditer(blk):
                font_name.setdefault(m.group(1).decode('latin1'), int(m.group(2)))
        # /Font 为间接引用（/Font 32 0 R）时，字体字典在目标对象里。
        # 旧版只认内联 /Font << ... >>，arXiv 老论文（2000s LaTeX PDF）因此
        # 拿不到字体名，文本只能走 fallback CMap → 整篇乱码/丢字（实测 6 页共 57 字符）。
        for m in re.finditer(rb'/Font\s+(\d+)\s+\d+\s+R', body):
            target = objs.get(int(m.group(1)))
            if target:
                for mm in FONTREF_RE.finditer(target):
                    font_name.setdefault(mm.group(1).decode('latin1'), int(mm.group(2)))
    cmap_by_ref = {}
    cmaps = {}
    fallback = {}
    for num, body in objs.items():
        m = TOUNI_RE.search(body)
        if not m:
            continue
        ref = int(m.group(1))
        cm = cmap_by_ref.get(ref)
        if cm is None and ref in objs:
            s = _inflate(objs[ref])
            if s is not None:
                cm = parse_cmap(s)
                cmap_by_ref[ref] = cm
        if cm:
            cmaps[num] = cm
            if len(cm) > len(fallback):
                fallback = cm

    def get_cmap(name):
        on = font_name.get(name)
        if on is not None and on in cmaps:
            return cmaps[on]
        # 简单字体（TrueType/Type1/Type3）没有 ToUnicode 时，字面串的字节是字符代码，
        # 应按该字体的 /Encoding 解码（走 stream_text 的字面串分支）。这里必须返回 None：
        # 若返回 fallback（别的字体的 CMap），cmap 非空会让字面串分支被跳过，
        # 导致数字/英文被静默丢弃（实测：PPT 导出的讲义里 20、40、Word 全部丢失）。
        if on is not None and on in objs:
            sub = re.search(rb"/Subtype\s*/(\w+)", objs[on][:2000])
            if sub and sub.group(1) in (b"TrueType", b"Type1", b"Type3"):
                return None
        return fallback

    def font_kind(name):
        """字体子类型：'simple'（TrueType/Type1/Type3，单字节码）/ 'cid'（Type0）/ None。"""
        on = font_name.get(name)
        if on is None or on not in objs:
            return None
        sub = re.search(rb"/Subtype\s*/(\w+)", objs[on][:2000])
        if not sub:
            return None
        st = sub.group(1)
        if st in (b"TrueType", b"Type1", b"Type3"):
            return "simple"
        if st == b"Type0":
            return "cid"
        return None

    enc_cache = {}

    def font_enc(on):
        if on not in enc_cache:
            enc_cache[on] = _font_encoding(objs, on)
        return enc_cache[on]

    def page_enc(num):
        """按页取 /Resources /Font（可为间接引用，也可能来自页树父节点）。"""
        pf = _page_fonts(objs, num) or font_name
        cache = {}

        def get_enc(name):
            if name in cache:
                return cache[name]
            on = pf.get(name)
            if on is None:
                on = font_name.get(name)
            cache[name] = font_enc(on) if on is not None else None
            return cache[name]

        return get_enc

    # 优先按页树取页：页号 = 阅读顺序、页数 = 真实页数、纯插图页也保留。
    # 旧版只在「没有 ToUnicode」时才走页树；有 CMap 的文档走"扫描对象序"启发式，
    # 会漏掉不含 BT 的页（实测 PPT 导出的课件 62 页只读出 32 页）。
    refs = _page_refs(objs)
    if refs:
        tree_pages = []
        for i, num in enumerate(refs, 1):
            txt = stream_text(_page_content(objs, num), get_cmap,
                              page_enc(num), font_kind).strip()
            tree_pages.append((i, txt))
        labels = (expand_labels(label_entries, len(tree_pages))
                  or detect_printed_pages(tree_pages))
        return tree_pages, labels

    pages = []
    seq = 0
    for num in sorted(objs):
        body = objs[num]
        if b'/FlateDecode' not in body:
            continue
        head = body[:600]
        if (b'/Image' in head or b'/DCTDecode' in head
                or b'/FontFile' in head or b'/Metadata' in head):
            continue
        s = _inflate(body)
        if s is None or b'BT' not in s:
            continue
        txt = stream_text(s, get_cmap, None, font_kind).strip()
        if txt:
            seq += 1
            pages.append((seq, txt))
    labels = (expand_labels(label_entries, len(pages))
              or detect_printed_pages(pages))
    return pages, labels


# ──────────────────────────── 页码映射 ────────────────────────────
# PDF 物理页 ≠ 书页（封面/前言/扉页未编号）时的映射：
#   优先 /PageLabels（PDF 标准，精确），无则从页面首尾行识别印刷页码并拟合偏移。
_ROMAN = [(1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'), (100, 'c'), (90, 'xc'),
          (50, 'l'), (40, 'xl'), (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')]


def _to_roman(n, upper=False):
    if n <= 0:
        return str(n)
    s = ''
    for v, sym in _ROMAN:
        while n >= v:
            s += sym
            n -= v
    return s.upper() if upper else s


def _to_alpha(n, upper=False):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord('a') + r) + s
    return s.upper() if upper else s


_LABEL_ENTRY_RE = re.compile(rb'(\d+)\s*<<(.*?)>>', re.S)


def page_labels(objs):
    """解析 Catalog 的 /PageLabels，返回 [{idx, style, start, prefix}]（按 idx 升序）；无则 []。"""
    cat = None
    for num, body in objs.items():
        if b'/Catalog' in body[:500]:
            cat = num
            break
    if cat is None:
        return []
    m = re.search(rb'/PageLabels\s+(\d+)\s+\d+\s+R', objs[cat])
    if not m:
        return []
    pl = int(m.group(1))
    if pl not in objs:
        return []
    body = objs[pl]
    nm = re.search(rb'/Nums\s*\[', body)
    if not nm:
        return []
    start = nm.end() - 1
    depth = 0
    end = None
    for i in range(start, min(len(body), start + 30000)):
        c = body[i:i + 1]
        if c == b'[':
            depth += 1
        elif c == b']':
            depth -= 1
            if depth == 0:
                end = i + 1
                break
    if end is None:
        return []
    entries = []
    for mm in _LABEL_ENTRY_RE.finditer(body[start:end]):
        d = mm.group(2)
        s = re.search(rb'/S\s*/([a-zA-Z])', d)
        st = re.search(rb'/St\s+(\d+)', d)
        pf = re.search(rb'/P\s*\((.*?)(?<!\\)\)', d, re.S)
        entries.append({
            'idx': int(mm.group(1)),
            'style': s.group(1).decode() if s else None,
            'start': int(st.group(1)) if st else 1,
            'prefix': pf.group(1).decode('latin1') if pf else '',
        })
    entries.sort(key=lambda e: e['idx'])
    return entries


def expand_labels(entries, total_pages):
    """把 /PageLabels 条目展开成 {物理页号(1-based): 标签}；无条目返回 {}。"""
    if not entries or total_pages <= 0:
        return {}
    out = {}
    for i, e in enumerate(entries):
        stop = entries[i + 1]['idx'] if i + 1 < len(entries) else total_pages
        n = e['start']
        for pi in range(max(0, e['idx']), min(stop, total_pages)):
            st = e['style']
            if st == 'D':
                lab = str(n)
            elif st == 'r':
                lab = _to_roman(n)
            elif st == 'R':
                lab = _to_roman(n, True)
            elif st == 'a':
                lab = _to_alpha(n)
            elif st == 'A':
                lab = _to_alpha(n, True)
            else:
                lab = ''
            if lab:
                out[pi + 1] = e['prefix'] + lab
            n += 1
    return out


_PAGE_NUM_RE = re.compile(r'^[-–—\s]*(\d{1,4})[-–—\s]*$')


def detect_printed_pages(pages):
    """兜底：从每页首尾各 2 行找印刷页码，取「物理页-印刷页」众数作偏移。

    要求候选数 ≥ max(5, 30% 页数) 且众数占比 ≥ 60%，否则判为无页码信息（返回 {}）。
    """
    if not pages:
        return {}
    cands = []
    for pnum, txt in pages:
        lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
        if not lines:
            continue
        for ln in lines[:2] + lines[-2:]:
            m = _PAGE_NUM_RE.match(ln)
            if m:
                v = int(m.group(1))
                if 1 <= v <= 3000:
                    cands.append((pnum, v))
                    break
    if len(cands) < max(5, int(len(pages) * 0.3)):
        return {}
    diffs = {}
    for p, v in cands:
        diffs[p - v] = diffs.get(p - v, 0) + 1
    off, cnt = max(diffs.items(), key=lambda kv: kv[1])
    if cnt < len(cands) * 0.6:
        return {}
    return {p: str(p - off) for p, _ in pages}


# ──────────────────────────────── 公共 ────────────────────────────────
def parse_pages(spec, total):
    if not spec or spec.lower() == 'all':
        return None
    want = set()
    for chunk in spec.split(','):
        m = PAGE_SPEC_RE.match(chunk.strip())
        if not m:
            continue
        a = int(m.group(1))
        b = int(m.group(2)) if m.group(2) else a
        for i in range(a, min(b, total) + 1):
            if i >= 1:
                want.add(i)
    return want or None


CACHE_DIR = os.environ.get("PDF2DOCX_CACHE") or os.path.join(
    tempfile.gettempdir(), "pdf2docx-cache")
CACHE_VERSION = 11         # 修 arXiv/TeX PDF（TJ 字面串丢失、/Font 间接引用、词间空格还原）


def cache_file_for(src):
    st = os.stat(src)
    key = hashlib.sha1(('%s|%d|%d' % (os.path.abspath(src), st.st_size, int(st.st_mtime))).encode()).hexdigest()[:20]
    os.makedirs(CACHE_DIR, exist_ok=True)
    return os.path.join(CACHE_DIR, key + '.json')


def main():
    if len(sys.argv) < 2:
        sys.stderr.write(__doc__)
        return 2
    src = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 and sys.argv[2] != '-' else None
    spec = sys.argv[3] if len(sys.argv) > 3 else None
    if not os.path.isfile(src):
        sys.stderr.write('doc_read: no such file: %s\n' % src)
        return 2
    ext = os.path.splitext(src)[1].lower()
    if ext not in ('.pdf', '.pptx', '.ppt', '.doc'):
        sys.stderr.write('doc_read: supports .pdf / .pptx / .ppt / .doc (got %s)\n' % ext)
        return 2

    t0 = time.time()
    cached = None
    try:
        cf = cache_file_for(src)
        if os.path.isfile(cf):
            with open(cf, 'r', encoding='utf-8') as fh:
                obj = json.load(fh)
            if obj.get('v') == CACHE_VERSION and isinstance(obj.get('pages'), list):
                cached = obj
    except Exception:
        cached = None

    if cached is not None:
        pages = [(p[0], p[1]) for p in cached['pages']]
        labels = cached.get('labels') or {}
        elapsed = time.time() - t0
        hit = True
    else:
        if ext == '.pptx':
            pages = pptx_pages(src)
            labels = detect_printed_pages(pages)
        elif ext == '.ppt':
            pages = ppt_pages(src)
            labels = {}
        elif ext == '.doc':
            pages = docbin_pages(src)
            labels = {}
        else:
            pages, labels = pdf_pages(src)
        elapsed = time.time() - t0
        hit = False
        try:
            cf = cache_file_for(src)
            with open(cf, 'w', encoding='utf-8') as fh:
                json.dump({'v': CACHE_VERSION, 'src': os.path.abspath(src), 'kind': ext[1:],
                           'spec': spec or '', 'pages': pages, 'labels': labels},
                          fh, ensure_ascii=False)
        except Exception:
            pass

    total_pages = len(pages)
    if spec:
        want = parse_pages(spec, total_pages)
        if want:
            pages = [p for p in pages if p[0] in want]
    def _page_tag(n):
        lab = labels.get(n) if labels else None
        return '--- [page %d | p.%s] ---' % (n, lab) if lab else '--- [page %d] ---' % n
    text = '\n\n'.join(_page_tag(p[0]) + '\n' + p[1] for p in pages)
    with_chars = sum(len(p[1]) for p in pages)

    if out_path:
        with open(out_path, 'w', encoding='utf-8') as fh:
            fh.write(text)
        sys.stderr.write('[doc_read] %s pages=%d chars=%d cache=%s %.2fs -> %s\n'
                         % (ext[1:], total_pages, with_chars, 'HIT' if hit else 'MISS', elapsed, out_path))
    else:
        sys.stderr.write('[doc_read] %s pages=%d chars=%d cache=%s %.2fs\n'
                         % (ext[1:], total_pages, with_chars, 'HIT' if hit else 'MISS', elapsed))
        sys.stdout.write(text)
    return 0


if __name__ == '__main__':
    sys.exit(main())
