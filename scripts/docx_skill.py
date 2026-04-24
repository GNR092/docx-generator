#!/usr/bin/env python3
"""Generate simple .docx files from plain text or markdown-like input.

Examples:
  python scripts/docx_skill.py --output reporte.docx --title "Reporte" --input reporte.md
  python scripts/docx_skill.py --output nota.docx --title "Nota" --line "Linea 1" --line "Linea 2"
"""

from __future__ import annotations

import argparse
import re
import zipfile
from datetime import date
from pathlib import Path


HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
UNORDERED_ITEM_RE = re.compile(r"^(\s*)[-*]\s+(.*)$")
ORDERED_ITEM_RE = re.compile(r"^(\s*)(\d+)\.\s+(.*)$")


RunSpec = tuple[str, bool, bool, bool]


def xml_escape_attr(text: str) -> str:
    return xml_escape(text).replace('"', "&quot;")


def run(text: str, bold: bool = False, italic: bool = False, code: bool = False) -> str:
    escaped = xml_escape(text)
    props: list[str] = []
    if bold:
        props.append("<w:b/>")
    if italic:
        props.append("<w:i/>")
    if code:
        props.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>')
        props.append('<w:shd w:val="clear" w:color="auto" w:fill="EDEDED"/>')

    if props:
        return (
            "<w:r>"
            f"<w:rPr>{''.join(props)}</w:rPr>"
            f'<w:t xml:space="preserve">{escaped}</w:t>'
            "</w:r>"
        )
    return f'<w:r><w:t xml:space="preserve">{escaped}</w:t></w:r>'


def parse_inline_runs(text: str) -> list[RunSpec]:
    runs: list[RunSpec] = []
    index = 0
    plain_buffer: list[str] = []

    def flush_plain() -> None:
        if plain_buffer:
            runs.append(("".join(plain_buffer), False, False, False))
            plain_buffer.clear()

    while index < len(text):
        if text.startswith("**", index):
            end = text.find("**", index + 2)
            if end != -1 and end > index + 2:
                flush_plain()
                runs.append((text[index + 2:end], True, False, False))
                index = end + 2
                continue
            plain_buffer.append("**")
            index += 2
            continue

        if text[index] == "*":
            end = text.find("*", index + 1)
            if end != -1 and end > index + 1:
                flush_plain()
                runs.append((text[index + 1:end], False, True, False))
                index = end + 1
                continue
            plain_buffer.append("*")
            index += 1
            continue

        if text[index] == "`":
            end = text.find("`", index + 1)
            if end != -1 and end > index + 1:
                flush_plain()
                runs.append((text[index + 1:end], False, False, True))
                index = end + 1
                continue
            plain_buffer.append("`")
            index += 1
            continue

        plain_buffer.append(text[index])
        index += 1

    flush_plain()

    if not runs:
        return [("", False, False, False)]
    return runs


def xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def paragraph(text: str) -> str:
    runs = "".join(
        run(content, bold=bold, italic=italic, code=code)
        for content, bold, italic, code in parse_inline_runs(text)
    )
    return f"<w:p>{runs}</w:p>"


def paragraph_with_props(text: str, ppr: str) -> str:
    runs = "".join(
        run(content, bold=bold, italic=italic, code=code)
        for content, bold, italic, code in parse_inline_runs(text)
    )
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>"


def heading_paragraph(level: int, text: str) -> str:
    sizes = {1: 36, 2: 30, 3: 26, 4: 24, 5: 22, 6: 20}
    size = sizes.get(level, 20)
    ppr = (
        f'<w:pStyle w:val="Heading{level}"/>'
        '<w:spacing w:before="180" w:after="120"/>'
    )
    styled = f"**{text}**"
    body = paragraph_with_props(styled, ppr)
    return body.replace("</w:rPr>", f'<w:sz w:val="{size}"/></w:rPr>', 1)


def list_paragraph(text: str, level: int, marker: str) -> str:
    left = 720 + (max(level, 0) * 360)
    ppr = (
        f'<w:ind w:left="{left}" w:hanging="360"/>'
        '<w:spacing w:after="80"/>'
    )
    return paragraph_with_props(f"{marker} {text}", ppr)


def quote_paragraph(text: str) -> str:
    ppr = (
        '<w:ind w:left="720"/>'
        '<w:spacing w:before="60" w:after="60"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="F7F7F7"/>'
    )
    return paragraph_with_props(text, ppr)


def code_paragraph(text: str) -> str:
    ppr = (
        '<w:ind w:left="720"/>'
        '<w:spacing w:before="40" w:after="40"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="F3F3F3"/>'
    )
    escaped = xml_escape(text)
    return (
        "<w:p>"
        f"<w:pPr>{ppr}</w:pPr>"
        "<w:r><w:rPr>"
        '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>'
        "</w:rPr>"
        f'<w:t xml:space="preserve">{escaped}</w:t>'
        "</w:r>"
        "</w:p>"
    )


def block_paragraph(line: str, in_code_block: bool) -> tuple[str, bool]:
    stripped = line.strip()
    if stripped.startswith("```"):
        return "", not in_code_block

    if in_code_block:
        return code_paragraph(line), in_code_block

    if not stripped:
        return paragraph(""), in_code_block

    heading_match = HEADING_RE.match(line)
    if heading_match:
        level = len(heading_match.group(1))
        return heading_paragraph(level, heading_match.group(2).strip()), in_code_block

    unordered_match = UNORDERED_ITEM_RE.match(line)
    if unordered_match:
        indent = len(unordered_match.group(1).replace("\t", "    "))
        level = indent // 2
        return list_paragraph(unordered_match.group(2), level, "-"), in_code_block

    ordered_match = ORDERED_ITEM_RE.match(line)
    if ordered_match:
        indent = len(ordered_match.group(1).replace("\t", "    "))
        level = indent // 2
        marker = f"{ordered_match.group(2)}."
        return list_paragraph(ordered_match.group(3), level, marker), in_code_block

    if line.lstrip().startswith("> "):
        return quote_paragraph(line.lstrip()[2:]), in_code_block

    return paragraph(line), in_code_block


def build_document_xml(lines: list[str]) -> str:
    blocks: list[str] = []
    in_code_block = False
    for line in lines:
        block, in_code_block = block_paragraph(line, in_code_block)
        if block:
            blocks.append(block)
    if in_code_block:
        blocks.append(code_paragraph(""))

    body = "".join(blocks)
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<w:document "
        "xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" "
        "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
        "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
        "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
        "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" "
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
        "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" "
        "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" "
        "xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" "
        "xmlns:wne=\"http://schemas.microsoft.com/office/2006/wordml\" "
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" "
        "mc:Ignorable=\"w14 wp14\">"
        "<w:body>"
        f"{body}"
        "<w:sectPr>"
        "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
        "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>"
        "<w:cols w:space=\"708\"/>"
        "<w:docGrid w:linePitch=\"360\"/>"
        "</w:sectPr>"
        "</w:body>"
        "</w:document>"
    )


def generate_docx(output: Path, lines: list[str]) -> None:
    content_types = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
        "</Types>"
    )
    rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
        "</Relationships>"
    )
    document = build_document_xml(lines)

    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document)


def load_lines(args: argparse.Namespace) -> list[str]:
    lines: list[str] = []
    if args.title:
        lines.append(f"# {args.title}")
        lines.append(f"Fecha: {date.today().isoformat()}")
        lines.append("")

    if args.input:
        text = args.input.read_text(encoding="utf-8")
        lines.extend(text.splitlines())

    if args.line:
        lines.extend(args.line)

    if not lines:
        lines = ["Documento", f"Fecha: {date.today().isoformat()}"]

    return lines


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="DOCX generator skill script")
    parser.add_argument("--output", required=True, type=Path, help="Output .docx path")
    parser.add_argument("--title", help="Document title")
    parser.add_argument("--input", type=Path, help="Input text/markdown file")
    parser.add_argument(
        "--line",
        action="append",
        help="Inline line content (can be repeated)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    lines = load_lines(args)
    generate_docx(args.output, lines)
    print(args.output)


if __name__ == "__main__":
    main()
