# 墨式（md2docx）

把 Markdown 与 LaTeX 公式，一键排成 WPS / Word 可编辑的 .docx。

- 公式是**原生 OMML 对象**（WPS/Word 可双击编辑、可「转换为 LaTeX 代码」），不是图片或纯文本；
- 正文默认**小四**、中文**宋体**、西文 **Times New Roman**，段落**首行缩进 2 个汉字**（标题/列表/表格/代码/题注/公式不缩进）；
- **大型公式**（分式、根式、∑/∏/∫/∮、lim、多行环境、矩阵）自动**独立成行居中**；
- 上下限按书写习惯：求和/连乘正上下、积分右侧、lim 正下方；
- 正文里的希腊字母、运算符、向量（v⃗）、均值（x̄）也统一为公式格式；
- 支持标题/列表/表格/引用/代码块/图片（自动剥离 EXIF）/可点击超链接；**图片与表格自动居中，表格内容水平/垂直居中；图片题注（图/表/Figure 开头）位于图片下方、居中显示、五号字**（图片与文字混排时会自动拆成独立居中段落）；
- **目录默认不插入**；只在 Markdown 中单独一行写 `[[TOC]]` 时才生成（TOC 域 + 标题预览）；
- 纯 Python 实现，无 Office 依赖，跨平台。

## 快速开始

### Windows

1. 安装 Python 3.8+（安装时勾选 “Add Python to PATH”）；
2. 打开 PowerShell 或 CMD，进入本目录：

```powershell
pip install -r requirements.txt
python md2docx.py sample.md -o sample.docx
```

3. 用 WPS/Word 打开 `sample.docx`。

> 如果 `python` 不可用，试试 `py md2docx.py ...`。

### macOS / Linux

```bash
pip3 install -r requirements.txt
python3 md2docx.py sample.md -o sample.docx
```

> `sample.md` 里的图片是占位路径（`out/test.png`）。没有该文件时图片会自动跳过，不影响生成；
> 换成自己的图片即可。

### 在 DeepSeek Harness 中

- **DSHM（鸿蒙）**：用 `python` 工具以 file 模式运行脚本，参数为输入 md 与 `-o 输出路径`；生成后用 `publish` 复制到下载目录。
- **DSH 桌面版（Windows/macOS/Linux）**：在终端工具里按上面的系统命令运行。

## 写 Markdown

```markdown
# 标题

行内公式 $E = mc^2$，大型公式独立成行：

$$
\int_0^1 x^2\,dx = \frac{1}{3}
$$

- 列表
- 表格、引用、代码块都支持

![图片](images/photo.png)

[链接](https://example.com)
```

## 目录

默认**不自动插入**目录。需要时在 Markdown 中单独一行写：

```markdown
[[TOC]]
```

生成后会插入「目录」标题 + TOC 域（含标题预览）。在 WPS/Word 中右键目录 →「更新域」即可生成带页码的正式目录。目录结束后会**自动分页**，正文另起一页。

## 公式规则

| 写法 | 效果 |
|---|---|
| `$...$` | 行内公式；分式/根式/∑∫ 等大型公式会自动独立成行 |
| `$$...$$` | 块级公式（居中） |
| `\begin{aligned}...\end{aligned}` | 多行对齐（自动转 array） |
| `v⃗`、`x̄`、`\vec{v}`、`\bar{x}` | 转为行内公式对象 |

## 自检 / 审计 / 备份

```bash
python3 md2docx.py --self-test          # 验证依赖并生成最小 docx
python3 md2docx.py --audit sample.docx  # 发布风险检查（元数据/链接/批注/修订）
python3 md2docx.py --backup md2docx.zip  # 打包 tools + pylibs + skills
python3 md2docx.py input.md -o out.docx --strict  # 有未识别 LaTeX 命令则失败
```

生成时会输出 `公式残留命令: N`，非 0 时在 stderr 打 `[warn] 未识别命令`；交付前建议用 `--strict` 跑一遍。

## 封面生成（可选）

`make_cover.py` 会用系统字体绘制一张 A4 比例封面 PNG（纯 Python，不依赖图像库）：

```bash
python3 make_cover.py
```

字体会自动探测：鸿蒙 `/system/fonts`、Linux Noto CJK / 文泉驿、macOS PingFang、Windows 微软雅黑。找不到时请安装中文字体，或修改脚本里的 `_FONT_ALTS`。

## 版权与致谢

- 本工具：MIT（见 LICENSE）
- 依赖：latex2mathml（MIT）、mathml2omml（MIT），纯 Python，随备份包附带许可证
- 生成的 docx 不嵌入字体，公式由 WPS/Word 的 Cambria Math 渲染

## 常见问题

**Q：WPS 打开后公式显示为方框？**  
公式是 OMML，需要 WPS/Word 2016+；若个别符号缺字，安装 Cambria Math 或换字体（`--math-font`）。

**Q：能生成目录 / 公式编号 / 页眉页脚吗？**  
目录可以（在 Markdown 里单独一行写 `[[TOC]]`）；公式编号、页眉页脚暂不支持。

**Q：图片路径怎么写？**  
相对 Markdown 文件所在目录。

**Q：在 DSH 桌面版里怎么调用？**  
用终端工具执行系统命令；若 DSH 版本没有 python 工具，就用系统 Python（见上方各平台命令）。

## PDF → Word（可选）

`pdf2docx/` 是可选前置：把 PDF 拆成行级文本与内嵌图片，先转成 Markdown，再交给本工具排版。

```bash
python pdf2docx/to_docx.py input.pdf -o output.docx            # 含图片
python pdf2docx/to_docx.py input.pdf -o output.docx --no-images  # 只要文字
```

适合「文字可编辑、层级合理」的场景；表格、多栏、页眉页脚不还原，扫描件不识别。
管线、文件说明与限制见 [pdf2docx/README.md](pdf2docx/README.md)。

