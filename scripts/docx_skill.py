#!/usr/bin/env python3
"""Generate simple .docx files from plain text or markdown-like input.

Examples:
  python scripts/docx_skill.py --output .docs/reporte.docx --title "Reporte" --input .docs/reporte.md
  python scripts/docx_skill.py --output .docs/nota.docx --title "Nota" --line "Linea 1" --line "Linea 2"
"""

from __future__ import annotations

import argparse
import zipfile
from datetime import date
from pathlib import Path


def xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def paragraph(text: str) -> str:
    escaped = xml_escape(text)
    return f'<w:p><w:r><w:t xml:space="preserve">{escaped}</w:t></w:r></w:p>'


def build_document_xml(lines: list[str]) -> str:
    body = "".join(paragraph(line) for line in lines)
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
        lines.append(args.title)
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
