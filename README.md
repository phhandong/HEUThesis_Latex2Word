# HEUThesis LaTeX to Word

把哈尔滨工程大学（Harbin Engineering University）学位论文的 LaTeX 版本转换为可编辑 Word版本。现在优先覆盖工程硕士、工程博士论文，按 `requirements/Engineering` 中的最新 Word 要求做后处理。

## 运行

```powershell
python heuthesis-latex2word.py HeuThesis_Overleaf\examples\book\bachelor\main.tex `
  -o output.docx `
  --degree master
```

启动简易 GUI：

```powershell
python heuthesis-latex2word-gui.py
```

## 依赖

- Python 3.10+
- Pandoc 3.x，并加入 `PATH`
- `python-docx`

```powershell
pip install -r requirements.txt
```

## 当前能力

- 展开 `\input`、`\include`、`\bibliography`、`\graphicspath`。
- 解析 `\heusetup`、中文/英文摘要、关键词、封面元数据。
- 中和 HEU 模板宏，生成 Pandoc 可读 LaTeX。
- 后处理 Word 页面：A4，上/下 2.8cm，左/右 2.5cm，页眉/页脚 2.0cm。
- 设置正文、标题、图表题注、参考文献、页眉页脚、目录域、三线表基础样式。
- 从 BibTeX 生成可编辑参考文献兜底列表。
- 生成 `*.report.md`，记录复杂公式、图片或参考文献降级情况。

详细格式映射见 [`docs/HEU_ENGINEERING_WORD_FORMAT.md`](docs/HEU_ENGINEERING_WORD_FORMAT.md)。

## 已知限制

- 第一版不追求逐像素复刻 PDF。
- 复杂公式如果 Pandoc 无法转换，会保留 TeX 表达并写入报告。
- EPS/TikZ 等图形不保证可编辑；能被 Pandoc 处理时会嵌入，否则报告降级。
- Word 目录域需要在 Word/WPS 中手动更新域刷新页码。

## 配合食用
[HeuThesis_Overleaf](https://github.com/phhandong/HeuThesis_Overleaf)

## TODO 
- Github 自动化流程编译成本地使用的软件
- 完善排版细节：中英文标题、交叉引用、目录、页首
- 适配经管等专业的封面适配
- 加入AI排版功能，提供更加精细的排版检查和修复
