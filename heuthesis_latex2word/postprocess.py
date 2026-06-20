from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph

from .latex_parser import ThesisMetadata
from .report import ConversionReport


A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7
BODY_SIZE_PT = 12.0
BODY_LINE_SPACING_PT = 20.5
FIRST_LINE_INDENT_PT = 24.0
CHAPTER_SIZE_PT = 18.0
CHAPTER_LINE_SPACING_PT = 28.5
CHAPTER_SPACE_BEFORE_PT = 28.35
CHAPTER_SPACE_AFTER_PT = 28.75
SECTION_SIZE_PT = 15.0
SECTION_LINE_SPACING_PT = 21.0
SECTION_SPACE_PT = 19.84
SUBSECTION_SIZE_PT = 14.0
SUBSECTION_LINE_SPACING_PT = 18.0
SUBSECTION_SPACE_PT = 17.01
SUBSUBSECTION_SIZE_PT = 12.0
SUBSUBSECTION_LINE_SPACING_PT = 20.5
SUBSUBSECTION_SPACE_PT = 8.5
CAPTION_SIZE_PT = 10.5
CAPTION_LINE_SPACING_PT = 13.65
TABLE_SIZE_PT = 10.5
TABLE_LINE_SPACING_PT = 13.65

FRONT_MATTER_TITLES = {
    "摘要",
    "Abstract",
    "目录",
    "参考文献",
    "学位论文原创性声明",
    "学位论文授权使用声明",
    "学位论文原创性声明和授权使用声明",
    "插图清单",
    "附表清单",
}

CAPTION_PREFIX_RE = re.compile(r"^[图表]\s*\d+(?:\.\d+)+")
REFERENCE_ENTRY_RE = re.compile(r"^\[\d+\]")


def _set_run_font(run, font_cn: str, font_en: str, size_pt: float, bold: bool | None = None) -> None:
    run.font.name = font_en
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(0, 0, 0)
    if bold is not None:
        run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_cn)
    rfonts.set(qn("w:ascii"), font_en)
    rfonts.set(qn("w:hAnsi"), font_en)


def _set_style_font(style, font_cn: str, font_en: str, size_pt: float, bold: bool = False) -> None:
    style.font.name = font_en
    style.font.size = Pt(size_pt)
    style.font.bold = bold
    style.font.color.rgb = RGBColor(0, 0, 0)
    rpr = style._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_cn)
    rfonts.set(qn("w:ascii"), font_en)
    rfonts.set(qn("w:hAnsi"), font_en)


def _set_paragraph_layout(
    paragraph,
    *,
    alignment: WD_ALIGN_PARAGRAPH | None,
    first_line_indent_pt: float | None,
    line_spacing_pt: float,
    space_before_pt: float,
    space_after_pt: float,
) -> None:
    if alignment is not None:
        paragraph.alignment = alignment
    paragraph.paragraph_format.first_line_indent = (
        None if first_line_indent_pt is None else Pt(first_line_indent_pt)
    )
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    paragraph.paragraph_format.line_spacing = Pt(line_spacing_pt)
    paragraph.paragraph_format.space_before = Pt(space_before_pt)
    paragraph.paragraph_format.space_after = Pt(space_after_pt)


def _set_style_paragraph_layout(
    style,
    *,
    alignment: WD_ALIGN_PARAGRAPH | None,
    first_line_indent_pt: float | None,
    line_spacing_pt: float,
    space_before_pt: float,
    space_after_pt: float,
) -> None:
    if alignment is not None:
        style.paragraph_format.alignment = alignment
    style.paragraph_format.first_line_indent = (
        None if first_line_indent_pt is None else Pt(first_line_indent_pt)
    )
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    style.paragraph_format.line_spacing = Pt(line_spacing_pt)
    style.paragraph_format.space_before = Pt(space_before_pt)
    style.paragraph_format.space_after = Pt(space_after_pt)


def _add_field(paragraph, instruction: str) -> None:
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_begin)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    run._r.append(instr)

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run._r.append(fld_sep)

    text = OxmlElement("w:t")
    text.text = "1"
    run._r.append(text)

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_end)


def _clear_paragraph(paragraph) -> None:
    for run in list(paragraph.runs):
        paragraph._p.remove(run._r)


def _format_sections(doc: Document) -> None:
    for section in doc.sections:
        section.page_width = Cm(A4_WIDTH_CM)
        section.page_height = Cm(A4_HEIGHT_CM)
        section.top_margin = Cm(2.8)
        section.bottom_margin = Cm(2.8)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        section.header_distance = Cm(2.0)
        section.footer_distance = Cm(2.0)
        section.start_type = WD_SECTION_START.NEW_PAGE
        section.odd_and_even_pages_header_footer = True


def _format_styles(doc: Document) -> None:
    styles = doc.styles
    for name in ("Normal", "Body Text", "First Paragraph"):
        if name in styles:
            _set_style_font(styles[name], "宋体", "Times New Roman", BODY_SIZE_PT, False)
            _set_style_paragraph_layout(
                styles[name],
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                first_line_indent_pt=FIRST_LINE_INDENT_PT,
                line_spacing_pt=BODY_LINE_SPACING_PT,
                space_before_pt=0,
                space_after_pt=0,
            )

    for name in ("Heading 1", "Title"):
        if name in styles:
            _set_style_font(styles[name], "黑体", "Times New Roman", CHAPTER_SIZE_PT, False)
            _set_style_paragraph_layout(
                styles[name],
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_indent_pt=0,
                line_spacing_pt=CHAPTER_LINE_SPACING_PT,
                space_before_pt=CHAPTER_SPACE_BEFORE_PT,
                space_after_pt=CHAPTER_SPACE_AFTER_PT,
            )
    if "Heading 2" in styles:
        _set_style_font(styles["Heading 2"], "黑体", "Times New Roman", SECTION_SIZE_PT, False)
        _set_style_paragraph_layout(
            styles["Heading 2"],
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            first_line_indent_pt=0,
            line_spacing_pt=SECTION_LINE_SPACING_PT,
            space_before_pt=SECTION_SPACE_PT,
            space_after_pt=SECTION_SPACE_PT,
        )
    if "Heading 3" in styles:
        _set_style_font(styles["Heading 3"], "黑体", "Times New Roman", SUBSECTION_SIZE_PT, False)
        _set_style_paragraph_layout(
            styles["Heading 3"],
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            first_line_indent_pt=0,
            line_spacing_pt=SUBSECTION_LINE_SPACING_PT,
            space_before_pt=SUBSECTION_SPACE_PT,
            space_after_pt=SUBSECTION_SPACE_PT,
        )
    if "Heading 4" in styles:
        _set_style_font(styles["Heading 4"], "黑体", "Times New Roman", SUBSUBSECTION_SIZE_PT, False)
        _set_style_paragraph_layout(
            styles["Heading 4"],
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            first_line_indent_pt=0,
            line_spacing_pt=SUBSUBSECTION_LINE_SPACING_PT,
            space_before_pt=SUBSUBSECTION_SPACE_PT,
            space_after_pt=SUBSUBSECTION_SPACE_PT,
        )

    for name in ("Caption", "Image Caption", "Table Caption"):
        if name in styles:
            _set_style_font(styles[name], "宋体", "Times New Roman", CAPTION_SIZE_PT, False)
            _set_style_paragraph_layout(
                styles[name],
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_indent_pt=0,
                line_spacing_pt=CAPTION_LINE_SPACING_PT,
                space_before_pt=0,
                space_after_pt=0,
            )


def _format_headers_and_footers(doc: Document, metadata: ThesisMetadata) -> None:
    school_line = f"哈尔滨工程大学{metadata.degree_label}学位论文"
    for section in doc.sections:
        for header in (section.header, section.even_page_header):
            para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            _clear_paragraph(para)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(school_line)
            _set_run_font(run, "宋体", "Times New Roman", 10.5, False)

        footer = section.footer
        para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        _clear_paragraph(para)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_field(para, "PAGE")
        for run in para.runs:
            _set_run_font(run, "宋体", "Times New Roman", 10.5, False)


def _replace_placeholders(doc: Document, report: ConversionReport, references: list[str] | None = None) -> None:
    for para in doc.paragraphs:
        text = para.text.strip()
        if text == "HEU_TOC_PLACEHOLDER":
            _clear_paragraph(para)
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _add_field(para, r'TOC \o "1-3" \h \z \u')
            report.note("已插入 Word 目录域；打开 Word 后请右键更新域以刷新页码。")
        elif text == "HEU_REFERENCES_PLACEHOLDER":
            _clear_paragraph(para)
            if references:
                _format_reference_paragraph(para)
                run = para.add_run(references[0])
                _set_run_font(run, "宋体", "Times New Roman", BODY_SIZE_PT, False)
                parent = para._p.getparent()
                index = parent.index(para._p)
                for ref in references[1:]:
                    new_para = OxmlElement("w:p")
                    parent.insert(index + 1, new_para)
                    index += 1
                    inserted = Paragraph(new_para, para._parent)
                    _format_reference_paragraph(inserted)
                    inserted.add_run(ref)
                    for run in inserted.runs:
                        _set_run_font(run, "宋体", "Times New Roman", BODY_SIZE_PT, False)
                report.note(f"已从 BibTeX 生成 {len(references)} 条可编辑参考文献。")
            else:
                run = para.add_run("参考文献由 Pandoc citeproc 生成；若此处为空，请检查 BibTeX 或转换报告。")
                _set_run_font(run, "宋体", "Times New Roman", BODY_SIZE_PT, False)
                report.warn("参考文献占位仍存在，Pandoc 未能自动生成参考文献列表。")
        elif text.startswith("HEU_LIST_OF_"):
            _clear_paragraph(para)
            run = para.add_run("请在 Word 中根据题注更新清单。")
            _set_run_font(run, "宋体", "Times New Roman", BODY_SIZE_PT, False)


def _is_front_matter_title(text: str) -> bool:
    return text.strip() in FRONT_MATTER_TITLES


def _is_caption_style(style_name: str, kind: str | None = None) -> bool:
    if kind == "figure":
        return style_name in {"Image Caption", "Figure Caption"}
    if kind == "table":
        return style_name == "Table Caption"
    return style_name in {"Caption", "Image Caption", "Figure Caption", "Table Caption"}


def _paragraph_has_drawing(para) -> bool:
    xml = para._element.xml
    return "<w:drawing" in xml or "<v:imagedata" in xml or "<wp:inline" in xml


def _paragraph_has_display_math(para) -> bool:
    return "<m:oMathPara" in para._element.xml


def _replace_text(para, text: str) -> None:
    _clear_paragraph(para)
    para.add_run(text)


def _format_reference_paragraph(para) -> None:
    _set_paragraph_layout(
        para,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent_pt=0,
        line_spacing_pt=BODY_LINE_SPACING_PT,
        space_before_pt=0,
        space_after_pt=0,
    )


def _format_heading_paragraph(para, level: int, text: str) -> None:
    front_title = _is_front_matter_title(text)
    if level == 1 or front_title:
        _set_paragraph_layout(
            para,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent_pt=0,
            line_spacing_pt=CHAPTER_LINE_SPACING_PT,
            space_before_pt=CHAPTER_SPACE_BEFORE_PT,
            space_after_pt=CHAPTER_SPACE_AFTER_PT,
        )
        size = CHAPTER_SIZE_PT
    elif level == 2:
        _set_paragraph_layout(
            para,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            first_line_indent_pt=0,
            line_spacing_pt=SECTION_LINE_SPACING_PT,
            space_before_pt=SECTION_SPACE_PT,
            space_after_pt=SECTION_SPACE_PT,
        )
        size = SECTION_SIZE_PT
    elif level == 3:
        _set_paragraph_layout(
            para,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            first_line_indent_pt=0,
            line_spacing_pt=SUBSECTION_LINE_SPACING_PT,
            space_before_pt=SUBSECTION_SPACE_PT,
            space_after_pt=SUBSECTION_SPACE_PT,
        )
        size = SUBSECTION_SIZE_PT
    else:
        _set_paragraph_layout(
            para,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            first_line_indent_pt=0,
            line_spacing_pt=SUBSUBSECTION_LINE_SPACING_PT,
            space_before_pt=SUBSUBSECTION_SPACE_PT,
            space_after_pt=SUBSUBSECTION_SPACE_PT,
        )
        size = SUBSUBSECTION_SIZE_PT

    for run in para.runs:
        _set_run_font(run, "黑体", "Times New Roman", size, text == "Abstract")


def _format_caption_paragraph(para, *, kind: str, chapter_no: int, caption_no: int) -> None:
    text = para.text.strip()
    if text and not CAPTION_PREFIX_RE.match(text):
        prefix = "图" if kind == "figure" else "表"
        _replace_text(para, f"{prefix} {chapter_no}.{caption_no} {text}")

    _set_paragraph_layout(
        para,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent_pt=0,
        line_spacing_pt=CAPTION_LINE_SPACING_PT,
        space_before_pt=0,
        space_after_pt=0,
    )
    for run in para.runs:
        _set_run_font(run, "宋体", "Times New Roman", CAPTION_SIZE_PT, False)


def _format_body_paragraph(para) -> None:
    if para.text.strip():
        alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        first_line = FIRST_LINE_INDENT_PT
    else:
        alignment = None
        first_line = None
    _set_paragraph_layout(
        para,
        alignment=alignment,
        first_line_indent_pt=first_line,
        line_spacing_pt=BODY_LINE_SPACING_PT,
        space_before_pt=0,
        space_after_pt=0,
    )
    for run in para.runs:
        _set_run_font(run, "宋体", "Times New Roman", BODY_SIZE_PT, None)


def _format_paragraphs(doc: Document, report: ConversionReport | None = None) -> None:
    current_chapter = 0
    figure_counts: dict[int, int] = {}
    table_counts: dict[int, int] = {}
    numbered_figures = 0
    numbered_tables = 0

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else ""
        text = para.text.strip()

        if style_name.startswith("Heading 1"):
            if current_chapter > 0 or not _is_front_matter_title(text):
                current_chapter += 1
            _format_heading_paragraph(para, 1, text)
            continue
        if style_name.startswith("Heading 2"):
            _format_heading_paragraph(para, 2, text)
            continue
        if style_name.startswith("Heading 3"):
            _format_heading_paragraph(para, 3, text)
            continue
        if style_name.startswith("Heading"):
            _format_heading_paragraph(para, 4, text)
            continue

        chapter_no = max(current_chapter, 1)
        if _is_caption_style(style_name, "figure"):
            figure_counts[chapter_no] = figure_counts.get(chapter_no, 0) + 1
            if not CAPTION_PREFIX_RE.match(text):
                numbered_figures += 1
            _format_caption_paragraph(
                para,
                kind="figure",
                chapter_no=chapter_no,
                caption_no=figure_counts[chapter_no],
            )
            continue
        if _is_caption_style(style_name, "table"):
            table_counts[chapter_no] = table_counts.get(chapter_no, 0) + 1
            if not CAPTION_PREFIX_RE.match(text):
                numbered_tables += 1
            _format_caption_paragraph(
                para,
                kind="table",
                chapter_no=chapter_no,
                caption_no=table_counts[chapter_no],
            )
            continue

        if REFERENCE_ENTRY_RE.match(text):
            _format_reference_paragraph(para)
            for run in para.runs:
                _set_run_font(run, "宋体", "Times New Roman", BODY_SIZE_PT, False)
            continue

        if _paragraph_has_drawing(para) and not text:
            _set_paragraph_layout(
                para,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_indent_pt=0,
                line_spacing_pt=BODY_LINE_SPACING_PT,
                space_before_pt=0,
                space_after_pt=0,
            )
            continue

        if _paragraph_has_display_math(para):
            _set_paragraph_layout(
                para,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_indent_pt=0,
                line_spacing_pt=BODY_LINE_SPACING_PT,
                space_before_pt=8,
                space_after_pt=8,
            )
            continue

        _format_body_paragraph(para)

    if report and (numbered_figures or numbered_tables):
        report.note(f"已按章补全题注编号：图 {numbered_figures} 个，表 {numbered_tables} 个。")


def _set_border(parent, edge: str, *, val: str, size_eighths_pt: int = 0) -> None:
    border = OxmlElement(f"w:{edge}")
    border.set(qn("w:val"), val)
    border.set(qn("w:sz"), str(size_eighths_pt))
    border.set(qn("w:space"), "0")
    border.set(qn("w:color"), "auto")
    parent.append(border)


def _format_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    for old in list(tbl_pr.findall(qn("w:tblBorders"))):
        tbl_pr.remove(old)
    tbl_borders = OxmlElement("w:tblBorders")
    _set_border(tbl_borders, "top", val="single", size_eighths_pt=12)
    _set_border(tbl_borders, "bottom", val="single", size_eighths_pt=12)
    for edge in ("left", "right", "insideV", "insideH"):
        _set_border(tbl_borders, edge, val="nil")
    tbl_pr.append(tbl_borders)

    if not table.rows:
        return
    for cell in table.rows[0].cells:
        tc_pr = cell._tc.get_or_add_tcPr()
        for old in list(tc_pr.findall(qn("w:tcBorders"))):
            tc_pr.remove(old)
        tc_borders = OxmlElement("w:tcBorders")
        _set_border(tc_borders, "bottom", val="single", size_eighths_pt=8)
        tc_pr.append(tc_borders)


def _format_tables(doc: Document) -> None:
    for table in doc.tables:
        table.autofit = True
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        _format_table_borders(table)
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _set_paragraph_layout(
                        para,
                        alignment=None,
                        first_line_indent_pt=0,
                        line_spacing_pt=TABLE_LINE_SPACING_PT,
                        space_before_pt=0,
                        space_after_pt=0,
                    )
                    for run in para.runs:
                        _set_run_font(run, "宋体", "Times New Roman", TABLE_SIZE_PT, None)


def postprocess_docx(
    path: str | Path,
    metadata: ThesisMetadata,
    report: ConversionReport,
    references: list[str] | None = None,
) -> None:
    doc = Document(path)
    _format_sections(doc)
    _format_styles(doc)
    _replace_placeholders(doc, report, references=references)
    _format_paragraphs(doc, report)
    _format_tables(doc)
    _format_headers_and_footers(doc, metadata)
    doc.save(path)
    report.note("已应用 HEU Engineering A4 页面、正文、标题、题注、表格、参考文献、页眉页脚与目录域后处理。")
