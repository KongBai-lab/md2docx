---
name: md2docx
description: 把 Markdown（含 LaTeX 公式）生成为 WPS/Word 可编辑的 .docx：原生 OMML 公式、正文小四宋体/西文 Times New Roman、正文首行缩进 2 字符、大型公式独立成行、图片与表格居中、图片题注、[[TOC]] 目录、可点击超链接。适用于生成带公式的 Word/WPS 文档、或把 Markdown/LaTeX 内容转成 .docx。
whenToUse: 用户要求生成 docx/Word/WPS 文档、要求公式可编辑，或需要把 Markdown/LaTeX 内容排版成 Word 文档时。
---

# md2docx — Markdown / LaTeX → 可编辑 DOCX

一个不依赖 Office 的 .docx 生成器：纯 Python 标准库 + 两个纯 Python 依赖，产出标准 OOXML；公式是原生 OMML 对象，在 WPS/Word 中可双击编辑、可「转换为 LaTeX 代码」。

## 特性

- 标记 → Word：标题、段落、粗斜体、链接、列表、引用、代码块、表格
- 公式：`$...$` 行内、`$$...$$` 块级；支持多行 `aligned`、矩阵、分段函数
- 公式对象：OMML（非图片、非文本），与手打公式同款 `m:acc` 重音结构
- 排版：正文小四（12pt）、中文宋体、西文 Times New Roman，段落首行缩进 2 个汉字；大型公式自动独立成行
- 上下限：求和/连乘在正上/正下，积分在符号右侧，lim 在正下方
- 正文符号（希腊字母/运算符/向量/均值）公式化
- 图片：png/jpg/gif，自动缩放，嵌入前剥离 EXIF/XMP 元数据
- 图片与表格：自动居中，表格内容水平/垂直居中；图片与文字同段时会自动拆成独立居中段落
- 图片题注：紧邻图片的「图/表/Figure…」段落自动居中、五号字（10.5pt），图片带 keepNext 避免与题注分页
- 目录：默认不插入；Markdown 中单独一行 `[[TOC]]` 才生成（TOC 域 + 标题预览，目录后自动分页，WPS 中右键更新域可得带页码目录）
- 超链接：真正可点击
- 自检、备份、发布风险审计：`--self-test` / `--backup` / `--audit`

## 安装

### DeepSeek Harness（DSHM，鸿蒙）

1. 把 `md2docx.py`、`make_cover.py`、`sample.md`、`requirements.txt`、`README.md`、`LICENSE` 放到 `$DSH_HOME/tools/`；
2. 把两个依赖（latex2mathml、mathml2omml，版本见 requirements.txt）安装到 `$DSH_HOME/pylibs/`；
3. 把本文件（SKILL.md）重命名为 `md2docx.md`，放到 `$DSH_HOME/skills/`；
4. 在新会话中验证：`--self-test`。

### DSH 桌面版（Windows / macOS / Linux）与普通 Python

```bash
pip install -r requirements.txt
python md2docx.py input.md -o output.docx
```

- Windows：建议 `py -m pip install -r requirements.txt`、`py md2docx.py ...`；
- macOS/Linux：用 `python3`；
- 在 DSH 桌面版里，用终端工具执行上述系统命令。

## 用法

```bash
python md2docx.py input.md -o output.docx [--title 标题] [--font-size-pt 12] [--math-font "Cambria Math"]
python md2docx.py input.md -o output.docx --strict   # 有未识别 LaTeX 命令则失败
python md2docx.py --self-test
python md2docx.py --audit output.docx
python md2docx.py --backup md2docx.zip
```

在 DSH 中通过 python 工具以 file 模式调用脚本，参数为输入 md 与 `-o 输出路径`；生成后用 publish 复制到下载目录。

## 发布 / 分享前的风险规避

- 运行 `--audit <docx>`：检查元数据（creator、lastModifiedBy）、外部链接、批注、修订标记；
- 图片嵌入前会自动剥离 JPEG 的 APP1（EXIF/XMP）段与 PNG 的文本元数据块；敏感照片仍建议先脱敏；
- 备份包只含脚本、依赖与技能文档，不含任何会话/用户数据；
- 分享 docx 前确认外部链接不含内网地址或访问令牌。

## 依赖与许可

- latex2mathml（MIT）、mathml2omml（MIT）—— 纯 Python，随备份包附带许可证文件；
- 本工具脚本：MIT（见 LICENSE）。

## 限制

- 不生成公式自动编号、页眉页脚、脚注；目录只在 Markdown 里单独一行写 `[[TOC]]` 时生成；
- 表格为简单表格（无合并单元格）；
- 复杂 LaTeX（宏包、`\newcommand`、`\label` 等）不支持。
