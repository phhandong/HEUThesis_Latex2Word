from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable


@dataclass
class ConversionReport:
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    generated_files: list[Path] = field(default_factory=list)

    def warn(self, message: str) -> None:
        self.warnings.append(message)

    def error(self, message: str) -> None:
        self.errors.append(message)

    def note(self, message: str) -> None:
        self.notes.append(message)

    def extend_warnings(self, messages: Iterable[str]) -> None:
        self.warnings.extend(messages)

    @property
    def ok(self) -> bool:
        return not self.errors

    def to_markdown(self) -> str:
        lines: list[str] = ["# HEU LaTeX to Word Conversion Report", ""]
        if self.errors:
            lines.extend(["## Errors", ""])
            lines.extend(f"- {item}" for item in self.errors)
            lines.append("")
        if self.warnings:
            lines.extend(["## Warnings", ""])
            lines.extend(f"- {item}" for item in self.warnings)
            lines.append("")
        if self.notes:
            lines.extend(["## Notes", ""])
            lines.extend(f"- {item}" for item in self.notes)
            lines.append("")
        if self.generated_files:
            lines.extend(["## Generated Files", ""])
            lines.extend(f"- `{path}`" for path in self.generated_files)
            lines.append("")
        if not (self.errors or self.warnings or self.notes):
            lines.append("No issues were reported.")
        return "\n".join(lines).rstrip() + "\n"
