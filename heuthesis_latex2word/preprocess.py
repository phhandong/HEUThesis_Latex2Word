from __future__ import annotations

import re
import base64
from pathlib import Path

from .latex_parser import ExpandedLatexProject, ThesisMetadata

COVER_MARKER = "HEU_COVER_PLACEHOLDER"
HEU_COVER_PLACEHOLDER = COVER_MARKER
DECLARATION_MARKER = "HEU_DECLARATION_PLACEHOLDER"
CITATION_PLACEHOLDER_RE = re.compile(r"HEU_CITE_(TEXT|SUPER)_([A-Za-z0-9_.-]+?)_END")


def _latex_escape(text: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def _paragraphs(text: str) -> str:
    return "\n\n".join(_latex_escape(line.strip()) for line in text.splitlines() if line.strip())


def build_declaration_latex() -> str:
    return r"""
\section*{学位论文原创性声明}

本人郑重声明：本论文的所有工作，是在导师的指导下，由作者本人独立完成的。有关观点、方法、数据和文献的引用已在文中指出，并与参考文献相对应。本论文在撰写过程中，所有由人工智能技术辅助生成的内容，作者均已明确标注，并对其真实性、准确性及合规性承担最终责任。论文中所有研究设计、数据分析、核心观点、结论及创新性研究成果均为作者独立思考完成，不存在将人工智能技术或工具标注为作者或完成人的情况。除文中已注明引用的内容外，本论文不包含任何其他个人或集体已经公开发表的作品成果。对本文的研究做出重要贡献的个人和集体，均已在文中以明确方式标明。本人完全意识到本声明的法律结果由本人承担。

\vspace{2em}

作者签名：\hspace{8em} 日期：\hspace{6em}

\section*{学位论文授权使用声明}

本人完全了解学校保护知识产权的有关规定，即研究生在校攻读学位期间论文工作的知识产权属于哈尔滨工程大学。哈尔滨工程大学有权保留并向国家有关部门或机构送交论文的复印件。本人允许哈尔滨工程大学将论文的部分或全部内容编入有关数据库进行检索，可采用影印、缩印或扫描等复制手段保存和汇编本学位论文，可以公布论文的全部内容。同时本人保证毕业后结合学位论文研究课题再撰写的论文一律注明作者第一署名单位为哈尔滨工程大学。涉密学位论文待解密后适用本声明。

\vspace{2em}

作者签名：\hspace{8em} 导师签名：\hspace{8em}

日期：\hspace{5em} 年\hspace{2em}月\hspace{2em}日\hspace{8em}日期：\hspace{5em} 年\hspace{2em}月\hspace{2em}日

\newpage
"""


def build_abstract_latex(metadata: ThesisMetadata) -> str:
    cn_keywords = "；".join(metadata.keywords_cn)
    en_keywords = "; ".join(metadata.keywords_en)
    return rf"""
\section*{{摘要}}

{_paragraphs(metadata.abstract_cn)}

\textbf{{关键词：}} {_latex_escape(cn_keywords)}

\newpage

\section*{{Abstract}}

{_paragraphs(metadata.abstract_en)}

\textbf{{Keywords:}} {_latex_escape(en_keywords)}

\newpage
"""


def build_toc_latex() -> str:
    return r"""
\section*{目录}

HEU_TOC_PLACEHOLDER

\newpage
"""


def _citation_slug(keys: str) -> str:
    encoded = []
    for key in (key.strip() for key in keys.split(",")):
        if key:
            token = base64.urlsafe_b64encode(key.encode("utf-8")).decode("ascii").rstrip("=")
            encoded.append(token)
    return ".".join(encoded)


def _replace_citations(text: str) -> str:
    def text_cite(match: re.Match[str]) -> str:
        return f"HEU_CITE_TEXT_{_citation_slug(match.group(1))}_END"

    def super_cite(match: re.Match[str]) -> str:
        return f"HEU_CITE_SUPER_{_citation_slug(match.group(1))}_END"

    cite_keys = r"([0-9A-Za-z_.:;,+\-\s]+)"
    text = re.sub(r"\\(?:inlinecite|onlinecite|lcite)\{" + cite_keys + r"\}", text_cite, text)
    return re.sub(r"\\(?:citeup|cite)\{" + cite_keys + r"\}", super_cite, text)


def _neutralize_heu_macros(text: str, metadata: ThesisMetadata, include_auth_scan: Path | None) -> str:
    text = re.sub(r"\\documentclass(?:\[[^\]]*\])?\{heuthesisbook\}", r"\\documentclass{book}", text)
    text = re.sub(r"\\usepackage\{heuthesis\}", "", text)
    text = re.sub(r"\\graphicspath\s*\{(?:[^{}]|\{[^{}]*\})*\}", "", text, flags=re.S)
    text = re.sub(r"\\frontmatter|\\mainmatter|\\backmatter", "", text)
    text = re.sub(r"\\cleardoublepage|\\clearpage", r"\\newpage", text)
    text = re.sub(r"\\makecover\b", COVER_MARKER, text)
    if include_auth_scan:
        auth = (
            r"\section*{学位论文原创性声明和授权使用声明}"
            "\n\n"
            + rf"\begin{{center}}\includegraphics[width=\textwidth,height=0.9\textheight,keepaspectratio]{{{include_auth_scan.resolve().as_posix()}}}\end{{center}}"
            + "\n\n"
            + _latex_escape(f"若扫描页未被 Word 嵌入，请手动插入文件：{include_auth_scan}")
            + "\n\n\\newpage\n"
        )
    else:
        auth = DECLARATION_MARKER
    text = re.sub(r"\\authorization(?:\[[^\]]+\])?", lambda _m: auth, text)
    text = re.sub(r"\\makeabstract\b", lambda _m: build_abstract_latex(metadata), text)
    text = re.sub(r"\\tableofcontents\b", lambda _m: build_toc_latex(), text)
    text = re.sub(r"\\listoffigures\b", r"\\section*{插图清单}\nHEU_LIST_OF_FIGURES_PLACEHOLDER", text)
    text = re.sub(r"\\listoftables\b", r"\\section*{附表清单}\nHEU_LIST_OF_TABLES_PLACEHOLDER", text)
    text = re.sub(r"\\BiChapter\{([^{}]+)\}\{([^{}]+)\}", r"\\chapter{\1}", text)
    text = re.sub(r"\\BiSection\{([^{}]+)\}\{([^{}]+)\}", r"\\section{\1}", text)
    text = re.sub(r"\\BiSubSection\{([^{}]+)\}\{([^{}]+)\}", r"\\subsection{\1}", text)
    text = re.sub(r"\\appendix\{([^{}]+)\}\{([^{}]+)\}(?:\[[^\]]*\])?", r"\\chapter*{附录 \1 \2}", text)
    text = re.sub(r"\\section\*\{本章小结\}(?:\[[^\]]*\])?", r"\\section*{本章小结}", text)
    text = re.sub(r"\\begin\{conclusions\}", r"\\chapter*{结论}", text)
    text = re.sub(r"\\end\{conclusions\}", "", text)
    text = re.sub(
        r"\\begin\{publication\}",
        r"\\chapter*{攻读硕士学位期间发表的论文和取得的科研成果}",
        text,
    )
    text = re.sub(r"\\end\{publication\}", "", text)
    text = re.sub(r"\\begin\{acknowledgements\}", r"\\chapter*{致谢}", text)
    text = re.sub(r"\\end\{acknowledgements\}", "", text)
    text = re.sub(r"\\begin\{publist\}", r"\\begin{itemize}", text)
    text = re.sub(r"\\end\{publist\}", r"\\end{itemize}", text)
    text = re.sub(r"\\(songti|heiti|kaishu|fangsong|xiaosi|wuhao|xiaoer|sanhao|sihao)(?:\[[^\]]*\])?", "", text)
    text = re.sub(r"\\cs\s+([A-Za-z]+)", r"\\textbackslash{}\1", text)
    return _replace_citations(text)


def _resolve_graphics_extensions(text: str, resource_paths: list[Path]) -> str:
    preferred_suffixes = [".png", ".jpg", ".jpeg", ".pdf", ".eps"]

    def replace(match: re.Match[str]) -> str:
        options = match.group(1) or ""
        name = match.group(2)
        if Path(name).suffix:
            return match.group(0)
        for base in resource_paths:
            for suffix in preferred_suffixes:
                candidate = base / f"{name}{suffix}"
                if candidate.exists():
                    return rf"\includegraphics{options}{{{name}{suffix}}}"
        return match.group(0)

    return re.sub(r"\\includegraphics(\[[^\]]*\])?\{([^{}]+)\}", replace, text)


def make_pandoc_latex(project: ExpandedLatexProject, include_auth_scan: Path | None = None) -> str:
    text = _resolve_graphics_extensions(project.latex, project.resource_paths)
    return _neutralize_heu_macros(text, project.metadata, include_auth_scan)
