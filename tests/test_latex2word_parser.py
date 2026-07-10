from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from heuthesis_latex2word.bibliography import ReferenceEntry
from heuthesis_latex2word.bibliography import collect_references
from heuthesis_latex2word.latex_parser import expand_project, normalize_soft_linebreaks, parse_heusetup
from heuthesis_latex2word.postprocess import (
    _format_paragraphs,
    _format_styles,
    _format_tables,
    _insert_page_section_breaks,
    _normalize_pandoc_custom_styles,
    _remove_unrelated_custom_styles,
    postprocess_docx,
)
from heuthesis_latex2word.preprocess import (
    COVER_MARKER,
    DECLARATION_MARKER,
    make_pandoc_latex,
)
from heuthesis_latex2word.report import ConversionReport


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "HeuThesis_Overleaf" / "examples" / "book" / "bachelor" / "main.tex"


def test_parse_heusetup_nested_values():
    setup = parse_heusetup(
        r"""
        \heusetup{
          ctitlecover={基于~\LaTeX~的哈尔滨工程大学本硕博\\论文模板使用说明},
          cauthor={马冬梅},
          ckeywords={\TeX, \LaTeX, 论文, 模板},
        }
        """
    )
    assert "ctitlecover" in setup
    assert setup["cauthor"] == "马冬梅"
    assert setup["ckeywords"].startswith(r"\TeX")


def test_expand_sample_project_metadata_and_paths():
    project = expand_project(SAMPLE, degree_override="master")
    assert project.metadata.author_cn == "马冬梅"
    assert project.metadata.degree_label == "工程硕士"
    assert project.metadata.secret_level == "公开"
    assert project.metadata.document_number == "no9527"
    assert project.metadata.classified_index == "TM301.2"
    assert project.metadata.udc == "62-5"
    assert project.metadata.author_en == "Ma Dongmei"
    assert project.metadata.affiliation_en == "School of Power and Energy Engineering"
    assert project.metadata.subject_en == "Power Engineering"
    assert "body\\chap01" not in project.latex
    assert "绪论" in project.latex
    assert any(path.name == "figures" for path in project.resource_paths)
    assert len(project.bibliography_files) == 1


def test_preprocess_neutralizes_heu_frontmatter():
    project = expand_project(SAMPLE, degree_override="doctor")
    latex = make_pandoc_latex(project)
    assert r"\usepackage{heuthesis}" not in latex
    assert COVER_MARKER in latex
    assert DECLARATION_MARKER in latex
    assert "HEU_TOC_PLACEHOLDER" in latex
    assert r"\chapter*{结论}" in latex
    assert r"\chapter*{攻读硕士学位期间发表的论文和取得的科研成果}" in latex
    assert r"\chapter*{致谢}" in latex
    assert r"\includegraphics[scale=1.0]{install-texlive.jpg}" in latex
    assert "HEU_CITE_SUPER_" in latex


def test_collect_bibtex_references():
    project = expand_project(SAMPLE, degree_override="master")
    refs = collect_references(project.bibliography_files)
    assert len(refs) >= 10
    assert refs[0].key
    assert refs[0].index == 1
    assert refs[0].text.startswith("[1]")


def test_normalize_soft_linebreaks_keeps_structural_boundaries():
    text = r"""
\section{标题}
第一行正文
第二行正文

\begin{equation}
a=b
c=d
\end{equation}
\item 列表项
下一段
"""
    normalized, count = normalize_soft_linebreaks(text)
    assert "第一行正文 第二行正文" in normalized
    assert "a=b\nc=d" in normalized
    assert "\\item 列表项\n下一段" in normalized
    assert count == 1


def test_postprocess_adds_cover_styles_bookmarks_and_ref_fields(tmp_path):
    project = expand_project(SAMPLE, degree_override="master")
    doc_path = tmp_path / "sample.docx"
    doc = Document()
    doc.add_paragraph("HEU_COVER_PLACEHOLDER")
    doc.add_paragraph("正文 HEU_CITE_SUPER_RFhNMjAwNQ_END")
    doc.add_paragraph("HEU_REFERENCES_PLACEHOLDER")
    doc.save(doc_path)

    postprocess_docx(
        doc_path,
        project.metadata,
        ConversionReport(),
        references=[ReferenceEntry(key="DXM2005", index=1, text="[1] Test reference.")],
    )

    processed = Document(doc_path)
    style_names = {style.name for style in processed.styles}
    xml = processed._element.xml
    assert "封面题名" in style_names
    assert "参考文献" in style_names
    assert "封面宋体小2" in style_names
    assert "封面2号黑体英文论文名称" in style_names
    assert 'w:name="HEU_REF_DXM2005"' in xml
    assert " REF HEU_REF_DXM2005 \\h " in xml


def test_postprocess_uses_styles_for_text_and_keeps_image_caption_tight():
    doc = Document()
    heading = doc.add_paragraph("样式标题")
    heading.style = "Heading 1"
    body = doc.add_paragraph("样式正文")
    image = doc.add_paragraph()
    image.add_run().add_picture(str(ROOT / "HeuThesis_Overleaf" / "heulogo.jpg"))
    caption = doc.add_paragraph("校徽")
    caption.style = "Caption"

    _format_styles(doc)
    _format_paragraphs(doc)

    assert heading.style.name == "章标题"
    assert heading.text == "第1章 样式标题"
    assert body.style.name == "正文"
    assert caption.style.name == "图表题注"
    assert image.style.name == "图片"
    assert image.style.paragraph_format.space_after.pt == 0
    assert image.style.paragraph_format.line_spacing.pt == 1.0
    assert image.style.paragraph_format.line_spacing_rule == WD_LINE_SPACING.EXACTLY
    assert doc.styles["正文"]._element.find(qn("w:qFormat")) is not None
    assert doc.styles["章标题"]._element.find(qn("w:qFormat")) is not None
    assert doc.styles["图表题注"]._element.find(qn("w:qFormat")) is not None
    assert heading.runs[0].font.size is None
    assert body.runs[0].font.size is None


def test_front_matter_and_chapters_get_required_section_page_numbering():
    doc = Document()
    doc.add_paragraph("封面占位")
    abstract = doc.add_paragraph("摘要")
    abstract.style = "Heading 1"
    doc.add_paragraph("摘要正文")
    toc = doc.add_paragraph("目录")
    toc.style = "Heading 1"
    chapter = doc.add_paragraph("绪论")
    chapter.style = "Heading 1"

    _format_styles(doc)
    _format_paragraphs(doc)
    _insert_page_section_breaks(doc)

    assert abstract.text == "摘要"
    assert toc.text == "目录"
    assert chapter.text == "第1章 绪论"

    xml = doc._element.xml
    assert 'w:type w:val="nextPage"' in xml
    assert 'w:type w:val="oddPage"' in xml
    assert 'w:pgNumType w:fmt="upperRoman" w:start="1"' in xml
    assert 'w:pgNumType w:fmt="decimal" w:start="1"' in xml


def test_inline_math_body_uses_single_line_spacing_style():
    doc = Document()
    para = doc.add_paragraph("含公式的正文")
    para._p.append(OxmlElement("m:oMath"))

    _format_styles(doc)
    _format_paragraphs(doc)

    assert para.style.name == "正文（内嵌公式）"
    assert para.style.paragraph_format.line_spacing == 1.0


def test_image_table_borders_are_removed():
    doc = Document()
    _format_styles(doc)
    table = doc.add_table(rows=1, cols=2)
    for cell in table.rows[0].cells:
        cell.paragraphs[0].add_run().add_picture(str(ROOT / "HeuThesis_Overleaf" / "heulogo.jpg"))

    _format_tables(doc)

    xml = table._tbl.xml
    assert 'w:val="single"' not in xml
    assert xml.count('w:val="nil"') >= 6
    assert {cell.paragraphs[0].style.name for cell in table.rows[0].cells} == {"图片"}


def test_unrelated_custom_styles_are_removed_and_legacy_heu_styles_hidden():
    doc = Document()
    _format_styles(doc)
    used = doc.styles.add_style("Temporary Used Style", WD_STYLE_TYPE.PARAGRAPH)
    doc.styles.add_style("Temporary Unused Style", WD_STYLE_TYPE.PARAGRAPH)
    doc.styles.add_style("HEU Body", WD_STYLE_TYPE.PARAGRAPH)
    doc.add_paragraph("kept", style=used)

    _remove_unrelated_custom_styles(doc)

    style_names = {style.name for style in doc.styles}
    assert "Temporary Used Style" in style_names
    assert "Temporary Unused Style" not in style_names
    assert "HEU Body" not in style_names


def test_pandoc_custom_styles_are_normalized_before_cleanup():
    doc = Document()
    _format_styles(doc)
    doc.styles.add_style("Compact", WD_STYLE_TYPE.PARAGRAPH)
    doc.styles.add_style("Verbatim Char", WD_STYLE_TYPE.CHARACTER)
    para = doc.add_paragraph("compact", style="Compact")
    code = doc.add_paragraph("command: ")
    code.add_run("pytest").style = "Verbatim Char"

    _normalize_pandoc_custom_styles(doc)
    _remove_unrelated_custom_styles(doc)

    style_names = {style.name for style in doc.styles}
    assert para.style.name == "正文"
    assert code.runs[0].style.name == "Default Paragraph Font"
    assert "Compact" not in style_names
    assert "Verbatim Char" not in style_names
