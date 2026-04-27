#!/usr/bin/env python3
"""Markdown parser shared between docx and pdf generators."""

from __future__ import annotations

import re
from typing import Optional, Tuple

HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
UNORDERED_ITEM_RE = re.compile(r"^(\s*)[-*]\s+(.*)$")
ORDERED_ITEM_RE = re.compile(r"^(\s*)(\d+)\.\s+(.*)$")
TABLE_SEPARATOR_RE = re.compile(r"^\s*\|?(?:\s*:?-{3,}:?\s*\|)+\s*:?-{3,}:?\s*\|?\s*$")
QUOTE_RE = re.compile(r"^\s*>\s?(.*)$")
HR_RE = re.compile(r"^\s*[-*_]{3,}\s*$")

RunSpec = Tuple[str, bool, bool, bool]


def xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _parse_emphasis(text: str, bold: bool = False, italic: bool = False) -> list[RunSpec]:
    runs: list[RunSpec] = []
    plain_buffer: list[str] = []
    index = 0

    def flush_plain() -> None:
        if plain_buffer:
            runs.append(("".join(plain_buffer), bold, italic, False))
            plain_buffer.clear()

    while index < len(text):
        for marker, add_bold, add_italic in (("***", True, True), ("**", True, False), ("*", False, True)):
            if text.startswith(marker, index):
                end = text.find(marker, index + len(marker))
                if end != -1 and end > index + len(marker):
                    flush_plain()
                    inner = text[index + len(marker):end]
                    runs.extend(_parse_emphasis(inner, bold or add_bold, italic or add_italic))
                    index = end + len(marker)
                    break
                plain_buffer.append(marker)
                index += len(marker)
                break
        else:
            plain_buffer.append(text[index])
            index += 1

    flush_plain()
    return runs


def _split_code_spans(text: str) -> list[tuple[str, bool]]:
    parts: list[tuple[str, bool]] = []
    index = 0
    while index < len(text):
        tick = text.find("`", index)
        if tick == -1:
            parts.append((text[index:], False))
            break
        if tick > index:
            parts.append((text[index:tick], False))
        end = text.find("`", tick + 1)
        if end == -1:
            parts.append((text[tick:], False))
            break
        parts.append((text[tick + 1:end], True))
        index = end + 1
    if not parts:
        return [("", False)]
    return parts


def _parse_links_and_emphasis(text: str) -> list[RunSpec]:
    runs: list[RunSpec] = []
    index = 0
    while index < len(text):
        start = text.find("[", index)
        if start == -1:
            runs.extend(_parse_emphasis(text[index:]))
            break
        close_label = text.find("]", start + 1)
        if close_label == -1 or close_label + 1 >= len(text) or text[close_label + 1] != "(":
            runs.extend(_parse_emphasis(text[index:]))
            break
        close_url = text.find(")", close_label + 2)
        if close_url == -1:
            runs.extend(_parse_emphasis(text[index:]))
            break

        runs.extend(_parse_emphasis(text[index:start]))
        label = text[start + 1:close_label]
        url = text[close_label + 2:close_url].strip()
        if url.startswith(("http://", "https://", "mailto:")):
            for run_text, run_bold, run_italic, _ in _parse_emphasis(label):
                runs.append((run_text, run_bold, run_italic, False))
        else:
            runs.extend(_parse_emphasis(text[start:close_url + 1]))

        index = close_url + 1

    if not runs:
        return [("", False, False, False)]
    return runs


def parse_inline_runs(text: str) -> list[RunSpec]:
    runs: list[RunSpec] = []
    for segment, is_code in _split_code_spans(text):
        if is_code:
            runs.append((segment, False, False, True))
            continue
        runs.extend(_parse_links_and_emphasis(segment))

    if not runs:
        return [("", False, False, False)]
    return runs


def split_table_row(row: str) -> list[str]:
    stripped = row.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    return [cell.strip() for cell in stripped.split("|")]


def is_markdown_table_start(lines: list[str], index: int) -> bool:
    if index + 1 >= len(lines):
        return False
    header = lines[index].strip()
    separator = lines[index + 1].strip()
    return "|" in header and bool(TABLE_SEPARATOR_RE.match(separator))


class Relationships:
    def __init__(self) -> None:
        self._next_id = 3
        self._link_to_rid: dict[str, str] = {}

    def get_hyperlink_rid(self, url: str) -> str:
        rid = self._link_to_rid.get(url)
        if rid is not None:
            return rid
        rid = f"rId{self._next_id}"
        self._next_id += 1
        self._link_to_rid[url] = rid
        return rid

    def document_rels_xml(self) -> str:
        rels = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
        ]
        for url, rid in self._link_to_rid.items():
            rels.append(
                "<Relationship "
                f'Id="{rid}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                f'Target="{xml_escape(url)}" '
                'TargetMode="External"/>'
            )
        rels.append("</Relationships>")
        return "".join(rels)


class Block:
    PARAGRAPH = "paragraph"
    HEADING = "heading"
    UNORDERED_LIST = "unordered_list"
    ORDERED_LIST = "ordered_list"
    QUOTE = "quote"
    CODE_BLOCK = "code_block"
    TABLE = "table"
    HR = "hr"


BlockSpec = Tuple[str, str, int, list[str], Optional[list[str]]]


def parse_blocks(lines: list[str], rels: Relationships) -> list[BlockSpec]:
    blocks: list[BlockSpec] = []
    in_code_block = False
    index = 0

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()

        if stripped.startswith("```"):
            in_code_block = not in_code_block
            index += 1
            continue

        if in_code_block:
            blocks.append((Block.CODE_BLOCK, "", 0, [line], None))
            index += 1
            continue

        if HR_RE.match(stripped):
            index += 1
            continue

        if is_markdown_table_start(lines, index):
            header = split_table_row(lines[index])
            index += 2
            body_rows: list[list[str]] = []
            while index < len(lines):
                candidate = lines[index].strip()
                if not candidate or "|" not in candidate or TABLE_SEPARATOR_RE.match(candidate):
                    break
                body_rows.append(split_table_row(lines[index]))
                index += 1
            blocks.append((Block.TABLE, "", 0, header, body_rows))
            continue

        if not stripped:
            blocks.append((Block.PARAGRAPH, "", 0, [""], None))
            index += 1
            continue

        heading_match = HEADING_RE.match(line)
        if heading_match:
            level = len(heading_match.group(1))
            blocks.append((Block.HEADING, heading_match.group(2).strip(), level, [], None))
            index += 1
            continue

        unordered_match = UNORDERED_ITEM_RE.match(line)
        if unordered_match:
            indent = len(unordered_match.group(1).replace("\t", "    "))
            items: list[str] = []
            while index < len(lines):
                u_match = UNORDERED_ITEM_RE.match(lines[index])
                if not u_match:
                    break
                items.append(u_match.group(2))
                index += 1
            blocks.append((Block.UNORDERED_LIST, "", indent // 2, items, None))
            continue

        ordered_match = ORDERED_ITEM_RE.match(line)
        if ordered_match:
            indent = len(ordered_match.group(1).replace("\t", "    "))
            items = []
            while index < len(lines):
                o_match = ORDERED_ITEM_RE.match(lines[index])
                if not o_match:
                    break
                items.append(o_match.group(3))
                index += 1
            blocks.append((Block.ORDERED_LIST, "", indent // 2, items, None))
            continue

        quote_match = QUOTE_RE.match(line)
        if quote_match:
            blocks.append((Block.QUOTE, quote_match.group(1), 0, [], None))
            index += 1
            continue

        blocks.append((Block.PARAGRAPH, stripped, 0, [], None))
        index += 1

    if in_code_block:
        blocks.append((Block.CODE_BLOCK, "", 0, [""], None))

    return blocks