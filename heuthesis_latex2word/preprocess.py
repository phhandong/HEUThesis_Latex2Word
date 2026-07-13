from __future__ import annotations

import re
import base64
import hashlib
from dataclasses import dataclass, field
from pathlib import Path

from .eps_fallback import convert_imagemagick_eps_to_png
from .latex_parser import ExpandedLatexProject, ThesisMetadata

COVER_MARKER = "HEU_COVER_PLACEHOLDER"
HEU_COVER_PLACEHOLDER = COVER_MARKER
DECLARATION_MARKER = "HEU_DECLARATION_PLACEHOLDER"
CITATION_PLACEHOLDER_RE = re.compile(r"HEU_CITE_(TEXT|SUPER)_([A-Za-z0-9_.-]+?)_END")
EQUATION_MARKER_RE = re.compile(
    r"^HEU_EQUATION_MARK_([A-Za-z0-9_.-]+)_END[\s\u25a1\u25fb\ufffd]*$"
)
EQUATION_REF_PLACEHOLDER_RE = re.compile(
    r"HEU_EQREF_(PLAIN|PAREN)_([A-Za-z0-9_-]+?)_END"
)


@dataclass
class GraphicsFallbackReport:
    converted_eps: list[tuple[Path, Path]] = field(default_factory=list)
    unresolved_eps: list[Path] = field(default_factory=list)

_VERBATIM_BLOCK_RE = re.compile(
    r"\\begin\{(?P<env>lstlisting|verbatim|Verbatim|minted)\}.*?"
    r"\\end\{(?P=env)\}",
    flags=re.S,
)
_INLINE_VERBATIM_RE = re.compile(
    r"\\(?:verb\*?|lstinline(?:\[[^\]]*\])?)(?P<delimiter>[^A-Za-z0-9\s])"
    r".*?(?P=delimiter)",
    flags=re.S,
)
_NUMBERED_EQUATION_RE = re.compile(
    r"\\begin\{(?P<env>equation|align|gather|multline)\}"
    r"(?P<body>.*?)"
    r"\\end\{(?P=env)\}",
    flags=re.S,
)


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


def _marker_token(value: str) -> str:
    return base64.urlsafe_b64encode(value.encode("utf-8")).decode("ascii").rstrip("=")


def _balanced_group(text: str, start: int, opening: str, closing: str) -> tuple[str, int] | None:
    if start >= len(text) or text[start] != opening:
        return None
    depth = 0
    for index in range(start, len(text)):
        char = text[index]
        if char == opening:
            depth += 1
        elif char == closing:
            depth -= 1
            if depth == 0:
                return text[start + 1 : index], index + 1
    return None


def _strip_textcolor_commands(text: str) -> str:
    command = r"\textcolor"
    search_from = 0
    while True:
        start = text.find(command, search_from)
        if start < 0:
            return text
        cursor = start + len(command)
        while cursor < len(text) and text[cursor].isspace():
            cursor += 1
        if cursor < len(text) and text[cursor] == "[":
            option = _balanced_group(text, cursor, "[", "]")
            if option is None:
                search_from = cursor + 1
                continue
            _, cursor = option
            while cursor < len(text) and text[cursor].isspace():
                cursor += 1
        else:
            # Pandoc understands ordinary named colors.  Only the optional
            # color-model form (for example ``\textcolor[rgb]``) needs the
            # fallback used for Word math conversion.
            search_from = cursor
            continue
        color = _balanced_group(text, cursor, "{", "}")
        if color is None:
            search_from = cursor + 1
            continue
        _, cursor = color
        while cursor < len(text) and text[cursor].isspace():
            cursor += 1
        content = _balanced_group(text, cursor, "{", "}")
        if content is None:
            search_from = cursor + 1
            continue
        replacement, end = content
        text = text[:start] + replacement + text[end:]
        search_from = start + len(replacement)


def _normalize_pandoc_math_commands(text: str) -> str:
    text = re.sub(r"\\textnormal\s*\{", r"\\text{", text)
    text = re.sub(
        r"\\underleftrightarrow\s*\{([^{}]+)\}",
        r"\\underset{\\leftrightarrow}{\1}",
        text,
    )
    text = re.sub(
        r"\\underleftarrow\s*\{([^{}]+)\}",
        r"\\underset{\\leftarrow}{\1}",
        text,
    )
    text = re.sub(
        r"\\underrightarrow\s*\{([^{}]+)\}",
        r"\\underset{\\rightarrow}{\1}",
        text,
    )
    text = re.sub(r"\\Bigm\s*\|", r"\\mid ", text)
    text = re.sub(r"\\rm\s*\{", r"\\mathrm{", text)
    return _strip_textcolor_commands(text)


def _prepare_pandoc_math(text: str) -> str:
    chunks: list[tuple[bool, str]] = []
    position = 0
    protected_matches = sorted(
        [*_VERBATIM_BLOCK_RE.finditer(text), *_INLINE_VERBATIM_RE.finditer(text)],
        key=lambda match: match.start(),
    )
    protected_end = -1
    for match in protected_matches:
        if match.start() < protected_end:
            continue
        if match.start() > position:
            chunks.append((False, text[position : match.start()]))
        chunks.append((True, match.group(0)))
        position = match.end()
        protected_end = match.end()
    if position < len(text):
        chunks.append((False, text[position:]))

    equation_labels: set[str] = set()
    prepared: list[tuple[bool, str]] = []

    def mark_equation(match: re.Match[str]) -> str:
        env = match.group("env")
        body = match.group("body")
        labels = re.findall(r"\\label\{([^{}]+)\}", body)
        equation_labels.update(labels)
        body = re.sub(r"\\label\{[^{}]+\}", "", body)
        tokens = ".".join(_marker_token(label) for label in labels) or "NONE"
        return (
            rf"\begin{{{env}}}{body}\end{{{env}}}"
            "\n\n"
            f"HEU_EQUATION_MARK_{tokens}_END"
            "\n\n"
        )

    for is_verbatim, chunk in chunks:
        if not is_verbatim:
            chunk = _normalize_pandoc_math_commands(chunk)
            chunk = _NUMBERED_EQUATION_RE.sub(mark_equation, chunk)
        prepared.append((is_verbatim, chunk))

    def replace_equation_ref(match: re.Match[str]) -> str:
        command, label = match.group(1), match.group(2)
        if label not in equation_labels:
            return match.group(0)
        kind = "PAREN" if command == "eqref" else "PLAIN"
        return f"HEU_EQREF_{kind}_{_marker_token(label)}_END"

    result: list[str] = []
    for is_verbatim, chunk in prepared:
        if not is_verbatim:
            chunk = re.sub(
                r"\\(eqref|ref)\{([^{}]+)\}",
                replace_equation_ref,
                chunk,
            )
        result.append(chunk)
    return "".join(result)


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


def _find_graphic(name: str, resource_paths: list[Path]) -> Path | None:
    path = Path(name)
    if path.is_absolute() and path.exists():
        return path
    for base in resource_paths:
        candidate = base / path
        if candidate.exists():
            return candidate
    return None


def _existing_eps_fallback(source: Path) -> Path | None:
    for suffix in (".png", ".jpg", ".jpeg", ".svg"):
        candidate = source.with_suffix(suffix)
        if candidate.exists():
            return candidate
    return None


def _generated_eps_fallback(
    source: Path,
    generated_assets_dir: Path | None,
    graphics_report: GraphicsFallbackReport | None,
) -> Path | None:
    existing = _existing_eps_fallback(source)
    if existing is not None:
        return existing
    if generated_assets_dir is None:
        return None

    digest = hashlib.sha256(str(source.resolve()).encode("utf-8")).hexdigest()[:12]
    target = generated_assets_dir / f"{source.stem}-{digest}.png"
    if target.exists() or convert_imagemagick_eps_to_png(source, target):
        if graphics_report is not None and (source, target) not in graphics_report.converted_eps:
            graphics_report.converted_eps.append((source, target))
        return target
    if graphics_report is not None and source not in graphics_report.unresolved_eps:
        graphics_report.unresolved_eps.append(source)
    return None


def _resolve_graphics_extensions(
    text: str,
    resource_paths: list[Path],
    *,
    generated_assets_dir: Path | None = None,
    graphics_report: GraphicsFallbackReport | None = None,
) -> str:
    preferred_suffixes = (".png", ".jpg", ".jpeg", ".svg", ".pdf")

    def replace(match: re.Match[str]) -> str:
        options = match.group(1) or ""
        name = match.group(2)
        source_name = Path(name)
        if source_name.suffix and source_name.suffix.lower() != ".eps":
            return match.group(0)

        if source_name.suffix.lower() == ".eps":
            source = _find_graphic(name, resource_paths)
            if source is None:
                return match.group(0)
            fallback = _generated_eps_fallback(source, generated_assets_dir, graphics_report)
            if fallback is None:
                return match.group(0)
            return rf"\includegraphics{options}{{{fallback.resolve().as_posix()}}}"

        for base in resource_paths:
            for suffix in preferred_suffixes:
                candidate = base / f"{name}{suffix}"
                if candidate.exists():
                    return rf"\includegraphics{options}{{{name}{suffix}}}"
            candidate = base / f"{name}.eps"
            if not candidate.exists():
                continue
            fallback = _generated_eps_fallback(candidate, generated_assets_dir, graphics_report)
            if fallback is not None:
                return rf"\includegraphics{options}{{{fallback.resolve().as_posix()}}}"
            return rf"\includegraphics{options}{{{name}.eps}}"
        return match.group(0)

    return re.sub(r"\\includegraphics(\[[^\]]*\])?\{([^{}]+)\}", replace, text)


def make_pandoc_latex(
    project: ExpandedLatexProject,
    include_auth_scan: Path | None = None,
    *,
    generated_assets_dir: Path | None = None,
    graphics_report: GraphicsFallbackReport | None = None,
) -> str:
    text = _resolve_graphics_extensions(
        project.latex,
        project.resource_paths,
        generated_assets_dir=generated_assets_dir,
        graphics_report=graphics_report,
    )
    text = _neutralize_heu_macros(text, project.metadata, include_auth_scan)
    return _prepare_pandoc_math(text)
