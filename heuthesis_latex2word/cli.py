from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .converter import ConversionOptions, convert_project


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="heuthesis-latex2word",
        description="Convert HEU thesis LaTeX projects to editable Word docx.",
    )
    parser.add_argument("input", type=Path, help="HEU LaTeX main.tex")
    parser.add_argument("-o", "--output", type=Path, required=True, help="Output .docx path")
    parser.add_argument("--degree", choices=["master", "doctor"], help="Engineering degree branch")
    parser.add_argument(
        "--mode",
        choices=["editable"],
        default="editable",
        help="Conversion mode. First version supports editable only.",
    )
    parser.add_argument("--editable", action="store_true", help="Compatibility flag; editable is default")
    parser.add_argument("--include-auth-scan", type=Path, help="Optional signed authorization scan path")
    parser.add_argument("--report", type=Path, help="Markdown conversion report path")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        result = convert_project(
            ConversionOptions(
                input_file=args.input,
                output_file=args.output,
                degree=args.degree,
                include_auth_scan=args.include_auth_scan,
                report_file=args.report,
                editable=True,
            )
        )
    except Exception as exc:
        print(f"转换失败: {exc}", file=sys.stderr)
        return 1

    print(f"生成 Word: {result.output_file}")
    print(f"转换报告: {result.report_file}")
    if result.report.warnings:
        print(f"警告: {len(result.report.warnings)} 条，详见报告。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
