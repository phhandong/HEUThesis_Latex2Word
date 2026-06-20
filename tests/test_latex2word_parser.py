from pathlib import Path

from heuthesis_latex2word.bibliography import collect_references
from heuthesis_latex2word.latex_parser import expand_project, parse_heusetup
from heuthesis_latex2word.preprocess import (
    COVER_MARKER,
    DECLARATION_MARKER,
    make_pandoc_latex,
)


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
    assert project.metadata.classified_index == "TM301.2"
    assert project.metadata.udc == "62-5"
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


def test_collect_bibtex_references():
    project = expand_project(SAMPLE, degree_override="master")
    refs = collect_references(project.bibliography_files)
    assert len(refs) >= 10
    assert refs[0].startswith("[1]")
