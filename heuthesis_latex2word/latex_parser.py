from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path


_COMMENT_SENTINEL = "\0PERCENT\0"


@dataclass
class ThesisMetadata:
    degree: str | None = None
    title_cn: str = ""
    title_cover_cn: str = ""
    title_en: str = ""
    author_cn: str = ""
    author_en: str = ""
    supervisor_cn: str = ""
    supervisor_en: str = ""
    affiliation_cn: str = ""
    subject_cn: str = ""
    student_id: str = ""
    submit_date_cn: str = ""
    oral_date_cn: str = ""
    keywords_cn: list[str] = field(default_factory=list)
    keywords_en: list[str] = field(default_factory=list)
    abstract_cn: str = ""
    abstract_en: str = ""

    @property
    def display_title_cn(self) -> str:
        return self.title_cover_cn or self.title_cn or "论文题目"

    @property
    def running_title_cn(self) -> str:
        return self.title_cn or self.title_cover_cn or "学位论文"

    @property
    def degree_label(self) -> str:
        if self.degree == "doctor":
            return "工程博士"
        return "工程硕士"

    @property
    def document_label(self) -> str:
        if self.degree == "doctor":
            return "专业学位博士学位论文"
        return "专业学位硕士学位论文"


@dataclass
class ExpandedLatexProject:
    root_dir: Path
    main_file: Path
    latex: str
    metadata: ThesisMetadata
    bibliography_files: list[Path] = field(default_factory=list)
    resource_paths: list[Path] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def read_text(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(errors="replace")


def strip_comments(text: str) -> str:
    protected = text.replace(r"\%", _COMMENT_SENTINEL)
    lines: list[str] = []
    for line in protected.splitlines():
        idx = line.find("%")
        if idx >= 0:
            line = line[:idx]
        lines.append(line.replace(_COMMENT_SENTINEL, r"\%"))
    return "\n".join(lines)


def latex_to_plain(text: str) -> str:
    text = text.replace("~", " ")
    text = text.replace(r"\\", "\n")
    replacements = {
        r"\LaTeX": "LaTeX",
        r"\TeX": "TeX",
        r"\XeLaTeX": "XeLaTeX",
        r"\XeTeX": "XeTeX",
        r"\quad": " ",
        r"\ ": " ",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"\\[a-zA-Z]+\*?(?:\[[^\]]*\])?\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\\[a-zA-Z]+\*?", "", text)
    text = text.replace("{", "").replace("}", "")
    return re.sub(r"[ \t]+", " ", text).strip()


def _find_braced_after(text: str, start: int) -> tuple[str, int] | None:
    brace_start = text.find("{", start)
    if brace_start < 0:
        return None
    depth = 0
    escaped = False
    for idx in range(brace_start, len(text)):
        ch = text[idx]
        if escaped:
            escaped = False
            continue
        if ch == "\\":
            escaped = True
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[brace_start + 1 : idx], idx + 1
    return None


def extract_command_block(text: str, command: str) -> tuple[str, tuple[int, int]] | None:
    match = re.search(r"\\" + re.escape(command) + r"\s*\{", text)
    if not match:
        return None
    block = _find_braced_after(text, match.start())
    if not block:
        return None
    value, end = block
    return value, (match.start(), end)


def _split_top_level_commas(text: str) -> list[str]:
    parts: list[str] = []
    start = 0
    depth = 0
    escaped = False
    for idx, ch in enumerate(text):
        if escaped:
            escaped = False
            continue
        if ch == "\\":
            escaped = True
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
        elif ch == "," and depth == 0:
            part = text[start:idx].strip()
            if part:
                parts.append(part)
            start = idx + 1
    tail = text[start:].strip()
    if tail:
        parts.append(tail)
    return parts


def parse_heusetup(text: str) -> dict[str, str]:
    found = extract_command_block(strip_comments(text), "heusetup")
    if not found:
        return {}
    block, _ = found
    result: dict[str, str] = {}
    for item in _split_top_level_commas(block):
        if "=" not in item:
            continue
        key, value = item.split("=", 1)
        key = key.strip()
        value = value.strip()
        if value.startswith("{") and value.endswith("}"):
            value = value[1:-1].strip()
        result[key] = value
    return result


def extract_environment(text: str, env: str) -> str:
    pattern = re.compile(
        r"\\begin\{" + re.escape(env) + r"\}(.*?)\\end\{" + re.escape(env) + r"\}",
        re.S,
    )
    match = pattern.search(text)
    if not match:
        return ""
    return latex_to_plain(match.group(1))


def parse_documentclass_degree(text: str) -> str | None:
    match = re.search(r"\\documentclass(?:\[([^\]]*)\])?\{heuthesisbook\}", text)
    if not match or not match.group(1):
        return None
    options = match.group(1)
    option_match = re.search(r"type\s*=\s*(doctor|master|bachelor|postdoc)", options)
    return option_match.group(1) if option_match else None


def _resolve_tex_path(root: Path, raw_name: str) -> Path:
    name = raw_name.strip()
    path = (root / name).resolve()
    if path.suffix:
        return path
    return path.with_suffix(".tex")


def _parse_graphicspath(text: str, root: Path) -> list[Path]:
    found = extract_command_block(text, "graphicspath")
    if not found:
        fallback = root / "figures"
        return [root, fallback.resolve()] if fallback.exists() else [root]
    block, _ = found
    paths = []
    for item in re.findall(r"\{([^{}]+)\}", block):
        path = (root / item).resolve()
        if path not in paths:
            paths.append(path)
    fallback = root / "figures"
    if fallback.exists() and fallback.resolve() not in paths:
        paths.append(fallback.resolve())
    return [root, *paths]


def _metadata_from_sources(
    setup: dict[str, str], cover_text: str, degree: str | None
) -> ThesisMetadata:
    def value(key: str) -> str:
        return latex_to_plain(setup.get(key, ""))

    metadata = ThesisMetadata(
        degree=degree,
        title_cn=value("ctitle"),
        title_cover_cn=value("ctitlecover"),
        title_en=value("etitle"),
        author_cn=value("cauthor"),
        author_en=value("eauthor"),
        supervisor_cn=value("csupervisor"),
        supervisor_en=value("esupervisor"),
        affiliation_cn=value("caffil"),
        subject_cn=value("csubject"),
        student_id=value("cstudentid"),
        submit_date_cn=value("csubmitdate") or value("cdate"),
        oral_date_cn=value("coralexdate"),
        keywords_cn=[latex_to_plain(x) for x in setup.get("ckeywords", "").split(",") if x.strip()],
        keywords_en=[latex_to_plain(x) for x in setup.get("ekeywords", "").split(",") if x.strip()],
        abstract_cn=extract_environment(cover_text, "cabstract"),
        abstract_en=extract_environment(cover_text, "eabstract"),
    )
    return metadata


def expand_project(main_file: str | Path, degree_override: str | None = None) -> ExpandedLatexProject:
    main_path = Path(main_file).resolve()
    root = main_path.parent
    warnings: list[str] = []
    main_text = read_text(main_path)
    degree = degree_override or parse_documentclass_degree(main_text)
    if degree not in {"master", "doctor"}:
        if degree:
            warnings.append(f"模板 type={degree}，第一版按工程硕士规则输出。")
        degree = "master"

    resource_paths = _parse_graphicspath(main_text, root)
    bibliography_files: list[Path] = []
    cover_texts: list[str] = []
    setup: dict[str, str] = {}
    visited: set[Path] = set()

    def expand_file(path: Path, skip_frontmatter: bool = False) -> str:
        path = path.resolve()
        if path in visited:
            warnings.append(f"跳过重复包含文件: {path}")
            return ""
        visited.add(path)
        if not path.exists():
            warnings.append(f"找不到 LaTeX 包含文件: {path}")
            return ""
        text = read_text(path)
        if "front" in path.parts or "cover" in path.name:
            cover_texts.append(text)
            setup.update(parse_heusetup(text))
        if skip_frontmatter:
            return ""
        return expand_latex_text(text, path.parent)

    def expand_latex_text(text: str, current_dir: Path) -> str:
        text = strip_comments(text)
        lines: list[str] = []
        for line in text.splitlines():
            include_match = re.search(r"\\(?:input|include)\{([^{}]+)\}", line)
            if include_match:
                include_path = _resolve_tex_path(current_dir, include_match.group(1))
                skip = "front" in include_path.parts or "cover" in include_path.name
                expanded = expand_file(include_path, skip_frontmatter=skip)
                if expanded:
                    lines.append(expanded)
                continue
            bib_match = re.search(r"\\bibliography\{([^{}]+)\}", line)
            if bib_match:
                for item in bib_match.group(1).split(","):
                    bib = (current_dir / item.strip()).resolve()
                    if not bib.suffix:
                        bib = bib.with_suffix(".bib")
                    if bib.exists():
                        bibliography_files.append(bib)
                    else:
                        warnings.append(f"找不到参考文献文件: {bib}")
                lines.append(r"\section*{参考文献}")
                lines.append("HEU_REFERENCES_PLACEHOLDER")
                continue
            lines.append(line)
        return "\n".join(lines)

    expanded = expand_latex_text(main_text, root)
    cover_blob = "\n".join(cover_texts)
    metadata = _metadata_from_sources(setup, cover_blob, degree)

    return ExpandedLatexProject(
        root_dir=root,
        main_file=main_path,
        latex=expanded,
        metadata=metadata,
        bibliography_files=bibliography_files,
        resource_paths=resource_paths,
        warnings=warnings,
    )
