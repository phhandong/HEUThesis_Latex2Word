from pathlib import Path

from docx import Document

from heuthesis_latex2word.bibliography import ReferenceEntry
from heuthesis_latex2word.bibliography import collect_references
from heuthesis_latex2word.latex_parser import expand_project, normalize_soft_linebreaks, parse_heusetup
from heuthesis_latex2word.postprocess import postprocess_docx
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
    assert "HEU Cover Title" in style_names
    assert "HEU Reference" in style_names
    assert 'w:name="HEU_REF_DXM2005"' in xml
    assert " REF HEU_REF_DXM2005 \\h " in xml
