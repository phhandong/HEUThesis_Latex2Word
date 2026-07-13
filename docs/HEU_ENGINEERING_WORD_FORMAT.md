# HEU Engineering LaTeX-to-Word 格式规范

本文档记录 `HeuThesis_Overleaf/examples/book/bachelor/main.tex` 所用
`heuthesisbook + heuthesis` 模板中的论文格式，并给出本项目 Word 后处理的
第一版映射值。

规范优先级：

1. `requirements/Engineering` 最新 Word 要求。
2. `HeuThesis_Overleaf/heuthesis.dtx` 模板内置格式。
3. Pandoc 默认 docx 格式。

## 来源文件

- 入口示例：`HeuThesis_Overleaf/examples/book/bachelor/main.tex`
- 样式核心：`HeuThesis_Overleaf/heuthesis.dtx`
- 表格示例：`HeuThesis_Overleaf/examples/book/bachelor/body/chap04.tex`
- 图片示例：`HeuThesis_Overleaf/examples/book/bachelor/body/chap05.tex`

## 页面

| 项目 | 模板值 | Word 后处理值 |
| --- | --- | --- |
| 纸张 | A4 | A4，21.0cm x 29.7cm |
| 上/下边距 | 2026 版面通知按 28mm | 2.8cm |
| 左/右边距 | 25mm | 2.5cm |
| 页眉/页脚距边界 | 20mm | 2.0cm |
| 正文版芯 | 约 160mm x 241mm | 由 A4 和边距确定 |

## 字体基准

| 区域 | 中文字体 | 英文字体 | 字号 | 行距 |
| --- | --- | --- | --- | --- |
| 正文 | 宋体 | Times New Roman | 小四，12pt | 固定 20.5pt |
| 页眉页脚 | 宋体 | Times New Roman | 五号，10.5pt | 单行 |
| 图题/表题 | 宋体 | Times New Roman | 五号，10.5pt | 约 13.65pt |
| 表格正文 | 宋体 | Times New Roman | 五号，10.5pt | 约 13.65pt |
| 参考文献条目 | 宋体 | Times New Roman | 小四，12pt | 固定 20.5pt |

正文段落首行缩进 2 个中文字符，Word 后处理按 24pt 设置。

## 标题

| 层级 | LaTeX 命令 | Word 样式 | 字体 | 字号 | 行距 | 对齐 | 段前/段后 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 章标题 | `\chapter` | Heading 1 | 黑体 | 小二，18pt | 28.5pt | 居中 | 28.35pt / 28.75pt |
| 节标题 | `\section` | Heading 2 | 黑体 | 15pt | 21pt | 左对齐 | 19.84pt / 19.84pt |
| 小节标题 | `\subsection` | Heading 3 | 黑体 | 14pt | 18pt | 左对齐 | 17.01pt / 17.01pt |
| 四级标题 | `\subsubsection` | Heading 4 | 黑体 | 12pt | 20.5pt | 左对齐 | 8.5pt / 8.5pt |

Word 后处理补全可见标题编号：章标题使用“第1章”，节标题使用“1.1、1.2”，
小节标题使用“1.1.1、1.1.2”。进入新章时节号归零，进入新节时小节号归零；
“本章小结”作为本章最后一节接续编号，前置、结尾部分的专用标题不参与编号。

中文工程硕博分支里，模板标题字体使用黑体，但没有对节标题、小节标题显式
追加 `\bfseries`。因此后处理会移除 Word 默认标题的蓝色和加粗效果，保留
黑色黑体。

前置部分和结尾部分中的专用标题，例如“摘要”“Abstract”“目录”“参考文献”
和声明页标题，使用章标题规格：小二黑体居中。

“学位论文原创性声明”和“学位论文授权使用声明”使用学校最新模板并排在同一页，
两份声明之间不插入分页；授权声明的签字和日期行清除首行/左侧缩进及分页标记，
使用同一段落内的固定制表位保持左右两栏。声明页结束后再分页进入摘要。

## 正文

| 项目 | Word 后处理值 |
| --- | --- |
| 中文字体 | 宋体 |
| 英文字体 | Times New Roman |
| 字号 | 12pt |
| 行距 | 固定 20.5pt |
| 对齐 | 两端对齐 |
| 首行缩进 | 24pt，约 2 字符 |
| 段前段后 | 0pt |

## 公式

| 项目 | 模板格式 | Word 后处理策略 |
| --- | --- | --- |
| 行间公式 | 居中 | 检测到 display math 段落时居中 |
| 公式编号 | 按章编号，例如 `(3-5)` | 保留 Pandoc 结果，普通公式优先转 OMML |
| 公式上下间距 | 约 8pt | display math 段落段前/段后 8pt |
| 行内公式 | 不应撑开正文行距 | 正文仍固定 20.5pt |

## 图题

| 项目 | Word 后处理值 |
| --- | --- |
| 图片位置 | 居中 |
| 图题位置 | 图片下方 |
| 编号 | 按章补全为 `图 5.1 标题` |
| 字体 | 宋体 / Times New Roman |
| 字号 | 五号，10.5pt |
| 行距 | 约 13.65pt |
| 对齐 | 居中 |
| 段前段后 | 0pt |

Pandoc 生成的 `Image Caption` 段落通常只有标题文本，不带“图 5.1”。后处理会
根据当前 Heading 1 章节序号补齐。

### EPS 兼容后备

无扩展名的图片引用优先解析为 PNG、JPEG、SVG 或 PDF。若唯一可用源图是 EPS，
预处理会先查找同名 PNG、JPEG 或 SVG；仍找不到时，对 ImageMagick 生成的栅格
EPS 使用纯标准库解码，并在临时目录生成 PNG 供 Pandoc 嵌入。源 EPS 不会被修改。

当前解码器支持 8-bit 灰度、调色板和 RGB 的原始/RLE 像素流，不执行任意
PostScript。遇到其他 EPS、截断数据或非法像素包时，转换会原子失败、不留下残缺
PNG，保留原始引用，并在转换报告中提示该图片可能无法在部分 Word 环境显示。

## 表题和表格

| 项目 | Word 后处理值 |
| --- | --- |
| 表题位置 | 表格上方 |
| 编号 | 按章补全为 `表 4.1 标题` |
| 表题字体 | 宋体 / Times New Roman，10.5pt |
| 表题对齐 | 居中 |
| 表格正文 | 宋体 / Times New Roman，10.5pt |
| 表格对齐 | 居中 |
| 三线表 | 顶线 1.5pt，表头下线 1pt，底线 1.5pt |

含图片或紧邻“图 N.N”题注的布局表使用 `FigureTable` 样式，统一清除表级和
单元格级边框。完全不含文字、图片、对象或公式的空图片布局表会被删除；普通
空表不会因本规则被删除。复杂 `subfigure`/TikZ 等 Pandoc 不支持的媒体结构仍需
单独的兼容预处理，不能只靠空表清理恢复图片。

## 参考文献

| 项目 | Word 后处理值 |
| --- | --- |
| 标题 | “参考文献”，小二黑体居中 |
| 条目字体 | 宋体 / Times New Roman |
| 条目字号 | 12pt |
| 行距 | 固定 20.5pt |
| 缩进 | 首行不缩进 |
| 编号 | Word 自动编号，数字顺序制，形如 `[1]` |
| 交叉引用 | 条目作为 Word“编号项”列出；正文使用带超链接的 `REF` 域引用编号 |

模板的 `thebibliography` 环境使用 `\normalsize`，因此参考文献条目应与正文同为
小四，而不是五号。

## 页眉页脚

| 区域 | 模板逻辑 | 当前后处理 |
| --- | --- | --- |
| 摘要、Abstract、目录页眉 | 学校名 + 学位论文名，宋体五号居中 | 奇偶页均显示学校名 + 学位论文名，清除段落缩进并居中 |
| 正文奇数页页眉 | 当前章标题，宋体五号居中 | `STYLEREF` 域读取当前“章标题”，宋体五号居中 |
| 正文偶数页页眉 | 学校名 + 学位论文名，宋体五号居中 | 学校名 + 学位论文名，宋体五号居中 |
| 页脚 | 页码居中，五号 | Word PAGE 域居中，五号 |
| 页眉线 | 双线，上细下粗 | 页眉段落底边框：`thickThinSmallGap`，3pt，距文字 1pt，自动颜色 |

摘要、Abstract 和目录分节的奇偶页统一显示学校名和学位论文类型。从第一章起，
正文各分节启用奇偶页不同页眉：奇数页通过 `STYLEREF` 自动跟随章标题，偶数页
固定显示学校名和学位论文类型。所有页眉均清除首行及左右缩进并居中；奇偶页
页眉均使用与 `GraduateEssay.docx` 一致的段落底边框双线，页脚均使用居中的
`PAGE` 域。

## 后处理映射

本项目在 `heuthesis_latex2word/postprocess.py` 中应用以上值：

- 页面：`_format_sections`
- Word 样式：`_format_styles`
- 标题/正文/题注/参考文献：`_format_paragraphs`
- 三线表和表格正文：`_format_tables`
- 页眉页脚和页码域：`_format_headers_and_footers`
