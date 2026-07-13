from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from .latex_parser import latex_to_plain, read_text


@dataclass(frozen=True)
class ReferenceEntry:
    key: str
    index: int
    text: str


def _split_bib_fields(body: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    start = 0
    depth = 0
    in_quote = False
    parts: list[str] = []
    for idx, ch in enumerate(body):
        if ch == '"' and (idx == 0 or body[idx - 1] != "\\"):
            in_quote = not in_quote
        elif not in_quote:
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
            elif ch == "," and depth == 0:
                parts.append(body[start:idx])
                start = idx + 1
    parts.append(body[start:])
    for part in parts:
        if "=" not in part:
            continue
        key, value = part.split("=", 1)
        value = value.strip().strip(",").strip()
        if (value.startswith("{") and value.endswith("}")) or (
            value.startswith('"') and value.endswith('"')
        ):
            value = value[1:-1]
        plain_value = latex_to_plain(value)
        plain_value = re.sub(r"(?<!\\)\$([^$]+)(?<!\\)\$", r"\1", plain_value)
        fields[key.strip().lower()] = plain_value
    return fields


def parse_bibtex(path: Path) -> list[dict[str, str]]:
    text = read_text(path)
    entries: list[dict[str, str]] = []
    pattern = re.compile(r"@(\w+)\s*\{\s*([^,]+)\s*,", re.S)
    pos = 0
    while True:
        match = pattern.search(text, pos)
        if not match:
            break
        # The pattern has already consumed the entry's opening brace.  Keep
        # that outer level in the depth count; otherwise the closing brace of
        # the first braced field is mistaken for the end of the whole entry.
        depth = 1
        in_quote = False
        end = match.end()
        for idx in range(match.end(), len(text)):
            ch = text[idx]
            if ch == '"' and (idx == 0 or text[idx - 1] != "\\"):
                in_quote = not in_quote
            elif not in_quote:
                if ch == "{":
                    depth += 1
                elif ch == "}":
                    depth -= 1
                    if depth == 0:
                        end = idx
                        break
        body = text[match.end() : end]
        fields = _split_bib_fields(body)
        fields["_type"] = match.group(1).lower()
        fields["_key"] = match.group(2).strip()
        entries.append(fields)
        pos = end + 1
    return entries


def _format_authors(entry: dict[str, str]) -> str:
    authors = entry.get("author") or entry.get("editor") or ""
    names = [name.strip() for name in re.split(r"\s+and\s+", authors, flags=re.I) if name.strip()]
    if len(names) <= 1:
        return authors.strip()
    language = entry.get("language", "").strip().lower()
    separator = "，" if language.startswith(("zh", "chi")) else ", "
    return separator.join(names)


def format_reference(entry: dict[str, str], index: int) -> str:
    authors = _format_authors(entry)
    title = entry.get("title", "")
    year = entry.get("year") or entry.get("date", "")
    journal = entry.get("journal") or entry.get("booktitle") or ""
    publisher = entry.get("publisher") or entry.get("school") or entry.get("institution") or ""
    address = entry.get("address", "")
    pages = entry.get("pages", "")
    parts = []
    if authors:
        parts.append(authors)
    if title:
        parts.append(title)
    source = journal or publisher
    if source:
        if address and publisher:
            source = f"{address}: {publisher}"
        parts.append(source)
    if year:
        parts.append(year)
    if pages:
        parts.append(pages)
    body = ". ".join(part for part in parts if part).rstrip(".")
    return f"[{index}] {body}."


def collect_references(paths: list[Path]) -> list[ReferenceEntry]:
    refs: list[ReferenceEntry] = []
    for path in paths:
        for entry in parse_bibtex(path):
            index = len(refs) + 1
            refs.append(
                ReferenceEntry(
                    key=entry.get("_key", ""),
                    index=index,
                    text=format_reference(entry, index),
                )
            )
    return refs
