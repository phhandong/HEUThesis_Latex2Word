from __future__ import annotations

import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

from .latex_parser import expand_project
from .bibliography import collect_references
from .postprocess import postprocess_docx
from .preprocess import make_pandoc_latex
from .report import ConversionReport
from .word_finalize import finalize_word_fields


@dataclass
class ConversionOptions:
    input_file: Path
    output_file: Path
    degree: str | None = None
    include_auth_scan: Path | None = None
    report_file: Path | None = None
    editable: bool = True


@dataclass
class ConversionResult:
    output_file: Path
    report_file: Path
    report: ConversionReport


def _pandoc_exists() -> bool:
    return shutil.which("pandoc") is not None


def _run_pandoc(
    temp_latex: Path,
    output_file: Path,
    cwd: Path,
    resource_paths: list[Path],
    bibliography_files: list[Path],
    report: ConversionReport,
) -> None:
    if not _pandoc_exists():
        raise RuntimeError("未检测到 pandoc，请先安装 Pandoc 3.x 并加入 PATH。")

    cmd = [
        "pandoc",
        str(temp_latex),
        "--from=latex",
        "--to=docx",
        "--standalone",
        f"--resource-path={';'.join(str(p) for p in resource_paths)}",
        "--output",
        str(output_file),
    ]
    if bibliography_files:
        cmd.append("--citeproc")
        for bib in bibliography_files:
            cmd.extend(["--bibliography", str(bib)])
    completed = subprocess.run(
        cmd,
        cwd=str(cwd),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        timeout=180,
    )
    if completed.stdout.strip():
        report.note(completed.stdout.strip())
    if completed.stderr.strip():
        for line in completed.stderr.splitlines():
            if "WARNING" in line.upper() or "[WARNING]" in line:
                report.warn(line)
            else:
                report.note(line)
    if completed.returncode != 0:
        raise RuntimeError(f"Pandoc 转换失败，退出码 {completed.returncode}: {completed.stderr.strip()}")


def convert_project(options: ConversionOptions) -> ConversionResult:
    report = ConversionReport()
    input_file = Path(options.input_file).absolute()
    output_file = Path(options.output_file).absolute()
    output_file.parent.mkdir(parents=True, exist_ok=True)
    report_file = (
        Path(options.report_file).absolute()
        if options.report_file
        else output_file.with_suffix(".report.md")
    )

    project = expand_project(input_file, degree_override=options.degree)
    report.extend_warnings(project.warnings)
    report.note(f"输入主文件: {project.main_file}")
    report.note(f"输出学位类型: {project.metadata.degree_label}")
    if project.soft_linebreaks_merged:
        report.note(f"已智能合并 LaTeX 正文软换行 {project.soft_linebreaks_merged} 处。")

    with tempfile.TemporaryDirectory(prefix="heuthesis_l2w_") as tmp:
        temp_latex = Path(tmp) / "pandoc_input.tex"
        temp_latex.write_text(
            make_pandoc_latex(project, include_auth_scan=options.include_auth_scan),
            encoding="utf-8",
        )
        _run_pandoc(
            temp_latex=temp_latex,
            output_file=output_file,
            cwd=project.root_dir,
            resource_paths=project.resource_paths,
            bibliography_files=project.bibliography_files,
            report=report,
        )

    references = collect_references(project.bibliography_files)
    if project.bibliography_files and not references:
        report.warn("检测到 BibTeX 文件，但未能解析出参考文献条目。")
    postprocess_docx(output_file, project.metadata, report, references=references)
    finalize_word_fields(output_file, report)
    report.generated_files.append(output_file)
    report.generated_files.append(report_file)
    report_file.write_text(report.to_markdown(), encoding="utf-8")
    return ConversionResult(output_file=output_file, report_file=report_file, report=report)
