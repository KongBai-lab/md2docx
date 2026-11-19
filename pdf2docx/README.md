# PDF → Word（可选模块）

本目录是 `md2docx` 的可选前置：先把 PDF 拆成「行级文本 + 内嵌图片」，排成 Markdown，
再交给上一级目录的 `md2docx.py` 生成 WPS / Word 可编辑的 .docx。

## 管线

```
PDF → pdf_read.py（行级文本 + 图片）→ Markdown → md2docx.py → .docx
```

- **标题层级**：每页字号最大且较短的块 → `##`；同字号短文（≤28 字）→ `###`；其余按 y 排成段落。
- **图片**：按页导出内嵌图，插到该页正文之后（JPEG 原样写出，Flate 位图重新封装为 PNG）。
- **正文样式**：由 md2docx 统一（小四宋体 / 西文 Times New Roman），公式还原为原生 OMML。
- 只追求「文字可编辑、层级合理」，不追求版式 1:1。

## 用法

```bash
# 仓库根目录下执行：PDF → docx
python pdf2docx/to_docx.py input.pdf -o output.docx

# 只要文字，不导出图片
python pdf2docx/to_docx.py input.pdf -o output.docx --no-images

# 只看版式摘要（每页文本块 / 矩形 / 图片数）
python pdf2docx/pdf_read.py input.pdf
```

产物：

| 文件 | 说明 |
|---|---|
| `output.docx` | 最终文档 |
| `output.md` | 中间 Markdown（图片用相对路径引用，可直接改后重跑） |
| `output_images/` | 从 PDF 抽出的图片 |

`to_docx.py` 默认调用**上一级目录**的 `md2docx.py`；两者不在同一仓库时，用环境变量指定：

```bash
MD2DOCX_PY=/path/to/md2docx.py python pdf2docx/to_docx.py input.pdf -o output.docx
```

## 文件

| 文件 | 作用 |
|---|---|
| `to_docx.py` | CLI：PDF → md → docx（含图片导出） |
| `pdf_read.py` | 统一读取层：给出每页 `texts`（x, y, size）与 `images`（x, y, w, h） |
| `pdf_blocks.py` | 行级文本块（y + 字号 + 文本） |
| `pdf_skeleton.py` | 版式骨架（文本块 / 矩形 / 图片位置） |
| `doc_read.py` | 纯标准库 PDF/PPTX 文本解析（对象表 → Flate → ToUnicode CMap） |

## 依赖

- 只用 Python 标准库，无需 pip 安装、无需联网；
- 生成 docx 阶段需要 `md2docx` 的两个依赖（`latex2mathml`、`mathml2omml`），见上级目录 `requirements.txt`。

## 什么时候**不该**用它

只想「读内容 / 答题 / 总结」时不要转 Word：表格会丢、多栏会混排，得到的 docx 看着能用其实不好用。
这类需求直接用文本提取（`doc_read.py` 打印纯文本）就够了。

它的「读」价值在底层：`pdf_read.py` 能给出每页的文本坐标与图片框，适合定位插图位置；
但文本本身可能重复或丢字，**坐标比文字更可信**。

## 限制（如实告知）

- **表格不还原**：PDF 表格是线条 + 散落文字，转出来是普通段落；
- **多栏**：按 y 排序会把栏间内容混排；
- 页眉 / 页脚 / 页码会被当成正文；
- 扫描件抽不出文字（本工具不做 OCR）；
- 字体缺少 `ToUnicode` 或子集化 CID 字体时，可能出现丢字、乱码；
- 图片只支持 `DCTDecode`（JPEG）与 8 bit `FlateDecode`；`JBIG2` / `JPEG2000` / `CCITT` 编码的图直接跳过；
- 图片较多时 docx 体积会明显变大（按原始分辨率嵌入）。
