#!/usr/bin/env python3
"""Generate PDF files from markdown using ReportLab (no external dependencies)."""

from __future__ import annotations

import argparse
import shutil
import sys
from datetime import date
from pathlib import Path
from typing import Optional


def check_python() -> bool:
    if not shutil.which("python") and not shutil.which("python3"):
        print("ERROR: Python no esta instalado. Por favor instala Python 3.9+ para usar esta funcion.")
        return False
    return True


def check_reportlab() -> bool:
    try:
        import reportlab
        return True
    except ImportError:
        print("ERROR: La libreria 'reportlab' no esta instalada.")
        print("Por favor ejecuta: pip install reportlab")
        return False


def main() -> None:
    if not check_python() or not check_reportlab():
        sys.exit(1)

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        ListFlowable, ListItem, HRFlowable
    )

    from parsers import (
        parse_blocks, parse_inline_runs, Block, RunSpec
    )

    def build_paragraph_style(name: str, **kwargs) -> ParagraphStyle:
        defaults = {
            "fontName": "Helvetica",
            "fontSize": 11,
            "leading": 15,
            "textColor": colors.black,
            "alignment": TA_LEFT,
            "spaceAfter": 8,
        }
        defaults.update(kwargs)
        return ParagraphStyle(name=name, **defaults)

    def runs_to_markup(specs: list[RunSpec]) -> str:
        parts: list[str] = []
        for text, bold, italic, code in specs:
            if code:
                text = f"<font face='Courier'>{text}</font>"
            elif bold and italic:
                text = f"<b><i>{text}</i></b>"
            elif bold:
                text = f"<b>{text}</b>"
            elif italic:
                text = f"<i>{text}</i>"
            parts.append(text)
        return "".join(parts)

    def make_paragraph(text: str, style: ParagraphStyle) -> Paragraph:
        specs = parse_inline_runs(text)
        markup = runs_to_markup(specs)
        return Paragraph(markup, style)

    def make_heading_paragraph(text: str, level: int) -> Paragraph:
        size_map = {1: 24, 2: 20, 3: 17, 4: 15, 5: 13, 6: 12}
        size = size_map.get(level, 14)
        style = build_paragraph_style(
            f"Heading{level}",
            fontName="Helvetica-Bold",
            fontSize=size,
            leading=size + 4,
            textColor=colors.HexColor("#1a1a1a"),
            spaceBefore=20 if level <= 2 else 14,
            spaceAfter=10,
            alignment=TA_LEFT,
        )
        return Paragraph(text, style)

    def make_quote_paragraph(text: str) -> Paragraph:
        style = build_paragraph_style(
            "Quote",
            fontName="Helvetica-Oblique",
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#555555"),
            leftIndent=20,
            rightIndent=20,
            spaceBefore=8,
            spaceAfter=8,
            alignment=TA_LEFT,
        )
        specs = parse_inline_runs(text)
        markup = runs_to_markup(specs)
        return Paragraph(markup, style)

    def make_code_block(lines: list[str]) -> list[Paragraph]:
        style = build_paragraph_style(
            "CodeBlock",
            fontName="Courier",
            fontSize=9,
            leading=12,
            textColor=colors.HexColor("#333333"),
            leftIndent=20,
            rightIndent=20,
            spaceBefore=6,
            spaceAfter=6,
        )
        return [Paragraph("\n".join(lines), style)]

    def make_table(header: list[str], body_rows: list[list[str]]) -> Table:
        data = [header] + body_rows
        col_count = max(len(h) for h in [header] + body_rows) if body_rows else len(header)
        col_width = (15 * cm) / col_count if col_count > 0 else 3 * cm
        tbl = Table(data, colWidths=[col_width] * col_count)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#34495e")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 10),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 10),
            ("TOPPADDING", (0, 0), (-1, 0), 10),
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#ecf0f1")),
            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 1), (-1, -1), 9),
            ("TOPPADDING", (0, 1), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bdc3c7")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ]))
        return tbl

    def make_unordered_list(items: list[str], style: ParagraphStyle) -> ListFlowable:
        list_items = [ListItem(make_paragraph(text, style)) for text in items]
        return ListFlowable(list_items, bulletType="bullet")

    def make_ordered_list(items: list[str], style: ParagraphStyle) -> ListFlowable:
        list_items = [ListItem(make_paragraph(text, style)) for text in items]
        return ListFlowable(list_items, bulletType="decimal")

    def build_pdf_story(lines: list[str]) -> list:
        from parsers import Relationships
        rels = Relationships()
        blocks = parse_blocks(lines, rels)
        story: list = []

        normal_style = build_paragraph_style(
            "Normal",
            fontName="Helvetica",
            fontSize=11,
            leading=15,
            textColor=colors.black,
            alignment=TA_JUSTIFY,
            spaceAfter=8,
        )

        for block_type, text, level, items, body_rows in blocks:
            if block_type == Block.PARAGRAPH:
                if text:
                    story.append(make_paragraph(text, normal_style))
                else:
                    story.append(Spacer(1, 8))

            elif block_type == Block.HEADING:
                story.append(make_heading_paragraph(text, level))

            elif block_type == Block.QUOTE:
                story.append(make_quote_paragraph(text))

            elif block_type == Block.CODE_BLOCK:
                story.extend(make_code_block(items))

            elif block_type == Block.TABLE:
                if body_rows is not None:
                    story.append(make_table(items, body_rows))
                    story.append(Spacer(1, 12))

            elif block_type == Block.UNORDERED_LIST:
                story.append(make_unordered_list(items, normal_style))
                story.append(Spacer(1, 8))

            elif block_type == Block.ORDERED_LIST:
                story.append(make_ordered_list(items, normal_style))
                story.append(Spacer(1, 8))

            elif block_type == Block.HR:
                story.append(HRFlowable(width="100%", thickness=0.5, color=colors.gray))
                story.append(Spacer(1, 8))

        return story

    def load_lines(args: argparse.Namespace) -> list[str]:
        lines: list[str] = []
        input_has_heading = False

        if args.input:
            text = args.input.read_text(encoding="utf-8")
            input_lines = text.splitlines()
            first_content = next((l.strip() for l in input_lines if l.strip()), "")
            import re
            heading_re = re.compile(r"^(#{1,6})\s+(.*)$")
            input_has_heading = bool(heading_re.match(first_content))
            lines.extend(input_lines)

        if args.title and not input_has_heading:
            if not lines:
                lines.append(f"# {args.title}")
            else:
                lines.insert(0, f"# {args.title}")
            lines.insert(1, f"Fecha: {date.today().isoformat()}")
            lines.insert(2, "")

        if args.line:
            lines.extend(args.line)

        if not lines:
            lines = ["Documento", f"Fecha: {date.today().isoformat()}"]

        return lines

    parser = argparse.ArgumentParser(description="PDF generator skill script")
    parser.add_argument("--output", required=True, type=Path, help="Output .pdf path")
    parser.add_argument("--title", help="Document title")
    parser.add_argument("--input", type=Path, help="Input text/markdown file")
    parser.add_argument("--author", default="OpenCode", help="Document author metadata")
    parser.add_argument("--lang", default="es-ES", help="Document spell-check language (e.g. es-ES, en-US)")
    parser.add_argument("--subject", help="Document subject metadata")
    parser.add_argument("--keywords", help="Comma-separated keywords metadata")
    parser.add_argument(
        "--line",
        action="append",
        help="Inline line content (can be repeated)",
    )
    args = parser.parse_args()
    lines = load_lines(args)

    doc = SimpleDocTemplate(
        str(args.output),
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2.5 * cm,
        bottomMargin=2.5 * cm,
        title=args.title or "Document",
        author=args.author,
        subject=args.subject,
        keywords=args.keywords,
    )

    story = build_pdf_story(lines)
    doc.build(story)
    print(args.output)


if __name__ == "__main__":
    main()