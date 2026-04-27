#!/usr/bin/env python3
"""Generate .docx or .pdf files from plain text or markdown-like input.

Examples:
  python scripts/docx_skill.py --output reporte.docx --title "Reporte" --input reporte.md
  python scripts/docx_skill.py --output reporte.pdf --format pdf --title "Reporte" --input reporte.md
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from datetime import date
from pathlib import Path
from typing import Optional

from parsers import (
    parse_blocks, parse_inline_runs, split_table_row, Block, BlockSpec, RunSpec,
    Relationships, xml_escape
)


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
        return False


HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")


RunSpec = tuple[str, bool, bool, bool, Optional[str]]


def run(
    text: str,
    bold: bool = False,
    italic: bool = False,
    code: bool = False,
    hyperlink: bool = False,
) -> str:
    escaped = xml_escape(text)
    props: list[str] = []
    if bold:
        props.append("<w:b/>")
    if italic:
        props.append("<w:i/>")
    if code:
        props.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>')
        props.append('<w:shd w:val="clear" w:color="auto" w:fill="EDEDED"/>')
    if hyperlink:
        props.append('<w:color w:val="0563C1"/>')
        props.append('<w:u w:val="single"/>')

    if props:
        return (
            "<w:r>"
            f"<w:rPr>{''.join(props)}</w:rPr>"
            f'<w:t xml:space="preserve">{escaped}</w:t>'
            "</w:r>"
        )
    return f'<w:r><w:t xml:space="preserve">{escaped}</w:t></w:r>'


def render_runs(specs: list) -> str:
    pieces: list[str] = []
    current_hyperlink: str | None = None
    hyperlink_runs: list[str] = []

    def flush_link() -> None:
        nonlocal current_hyperlink
        if current_hyperlink is not None:
            pieces.append(f'<w:hyperlink r:id="{current_hyperlink}" w:history="1">{"".join(hyperlink_runs)}</w:hyperlink>')
            hyperlink_runs.clear()
            current_hyperlink = None

    for spec in specs:
        if len(spec) == 5:
            content, bold, italic, code, hyperlink_rid = spec
        else:
            content, bold, italic, code = spec
            hyperlink_rid = None
        run_xml = run(content, bold=bold, italic=italic, code=code, hyperlink=hyperlink_rid is not None)
        if hyperlink_rid is None:
            flush_link()
            pieces.append(run_xml)
            continue
        if current_hyperlink is not None and current_hyperlink != hyperlink_rid:
            flush_link()
        if current_hyperlink is None:
            current_hyperlink = hyperlink_rid
        hyperlink_runs.append(run_xml)

    flush_link()
    return "".join(pieces)


def paragraph(text: str, rels: Relationships) -> str:
    runs = render_runs(parse_inline_runs(text))
    return f"<w:p>{runs}</w:p>"


def paragraph_with_props(text: str, ppr: str, rels: Relationships) -> str:
    runs = render_runs(parse_inline_runs(text))
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>"


def heading_paragraph(level: int, text: str, rels: Relationships) -> str:
    normalized = min(max(level, 1), 6)
    ppr = f'<w:pStyle w:val="Heading{normalized}"/><w:spacing w:before="120" w:after="80"/>'
    return paragraph_with_props(text, ppr, rels)


def list_paragraph(text: str, level: int, num_id: int, rels: Relationships) -> str:
    normalized = min(max(level, 0), 8)
    ppr = (
        "<w:numPr>"
        f'<w:ilvl w:val="{normalized}"/>'
        f'<w:numId w:val="{num_id}"/>'
        "</w:numPr>"
        '<w:spacing w:after="40"/>'
    )
    return paragraph_with_props(text, ppr, rels)


def quote_paragraph(text: str, rels: Relationships) -> str:
    ppr = (
        '<w:pStyle w:val="Quote"/>'
        '<w:ind w:left="720"/>'
        '<w:spacing w:before="60" w:after="60"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="F7F7F7"/>'
    )
    return paragraph_with_props(text, ppr, rels)


def code_paragraph(text: str) -> str:
    ppr = (
        '<w:pStyle w:val="CodeBlock"/>'
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


def table_xml(lines: list[str], start_index: int, rels: Relationships) -> tuple[str, int]:
    header = split_table_row(lines[start_index])
    index = start_index + 2
    body_rows: list[list[str]] = []

    while index < len(lines):
        candidate = lines[index].strip()
        if not candidate or "|" not in candidate:
            break
        from parsers import TABLE_SEPARATOR_RE
        if TABLE_SEPARATOR_RE.match(candidate):
            break
        body_rows.append(split_table_row(lines[index]))
        index += 1

    columns = max(1, len(header))
    col_width = max(1200, 9000 // columns)

    tbl_parts = [
        "<w:tbl>",
        "<w:tblPr><w:tblStyle w:val=\"TableGrid\"/><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>",
        "<w:tblGrid>",
        "".join(f'<w:gridCol w:w="{col_width}"/>' for _ in range(columns)),
        "</w:tblGrid>",
    ]

    def row_xml(cells: list[str], header_row: bool = False) -> str:
        normalized_cells = cells + [""] * (columns - len(cells))
        chunks: list[str] = ["<w:tr>"]
        for cell in normalized_cells[:columns]:
            cell_text = f"**{cell}**" if header_row else cell
            chunks.append(
                "<w:tc>"
                f'<w:tcPr><w:tcW w:w="{col_width}" w:type="dxa"/></w:tcPr>'
                f"{paragraph(cell_text, rels)}"
                "</w:tc>"
            )
        chunks.append("</w:tr>")
        return "".join(chunks)

    tbl_parts.append(row_xml(header, header_row=True))
    for body in body_rows:
        tbl_parts.append(row_xml(body))

    tbl_parts.append("</w:tbl>")
    return "".join(tbl_parts), index


def styles_xml(lang: str = "es-ES") -> str:
    heading_sizes = {1: 36, 2: 30, 3: 26, 4: 24, 5: 22, 6: 20}
    heading_styles = "".join(
        "<w:style w:type=\"paragraph\" "
        f"w:styleId=\"Heading{level}\">"
        f"<w:name w:val=\"heading {level}\"/>"
        "<w:basedOn w:val=\"Normal\"/>"
        "<w:next w:val=\"Normal\"/>"
        "<w:uiPriority w:val=\"9\"/>"
        "<w:qFormat/>"
        f"<w:pPr><w:spacing w:before=\"{120 + (7 - level) * 20}\" w:after=\"80\"/></w:pPr>"
        f"<w:rPr><w:b/><w:sz w:val=\"{heading_sizes[level]}\"/></w:rPr>"
        "</w:style>"
        for level in range(1, 7)
    )

    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
        "<w:docDefaults>"
        f"<w:rPrDefault><w:rPr><w:lang w:val=\"{lang}\"/></w:rPr></w:rPrDefault>"
        "<w:pPrDefault><w:pPr><w:spacing w:line=\"280\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault>"
        "</w:docDefaults>"
        "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">"
        "<w:name w:val=\"Normal\"/>"
        "<w:qFormat/>"
        "</w:style>"
        f"{heading_styles}"
        "<w:style w:type=\"paragraph\" w:styleId=\"Quote\">"
        "<w:name w:val=\"Quote\"/>"
        "<w:basedOn w:val=\"Normal\"/>"
        "<w:pPr><w:ind w:left=\"720\"/><w:spacing w:before=\"80\" w:after=\"80\"/></w:pPr>"
        "<w:rPr><w:i/><w:color w:val=\"555555\"/></w:rPr>"
        "</w:style>"
        "<w:style w:type=\"paragraph\" w:styleId=\"CodeBlock\">"
        "<w:name w:val=\"Code Block\"/>"
        "<w:basedOn w:val=\"Normal\"/>"
        "<w:pPr><w:ind w:left=\"720\"/><w:spacing w:before=\"40\" w:after=\"40\"/></w:pPr>"
        "<w:rPr><w:rFonts w:ascii=\"Consolas\" w:hAnsi=\"Consolas\" w:cs=\"Consolas\"/></w:rPr>"
        "</w:style>"
        "</w:styles>"
    )


def numbering_xml() -> str:
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
        "<w:abstractNum w:abstractNumId=\"0\">"
        "<w:multiLevelType w:val=\"hybridMultilevel\"/>"
        "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"•\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl>"
        "<w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"◦\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1080\" w:hanging=\"360\"/></w:pPr></w:lvl>"
        "<w:lvl w:ilvl=\"2\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"▪\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1440\" w:hanging=\"360\"/></w:pPr></w:lvl>"
        "</w:abstractNum>"
        "<w:abstractNum w:abstractNumId=\"1\">"
        "<w:multiLevelType w:val=\"multilevel\"/>"
        "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl>"
        "<w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerLetter\"/><w:lvlText w:val=\"%2.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1080\" w:hanging=\"360\"/></w:pPr></w:lvl>"
        "<w:lvl w:ilvl=\"2\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/><w:lvlText w:val=\"%3.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1440\" w:hanging=\"360\"/></w:pPr></w:lvl>"
        "</w:abstractNum>"
        "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>"
        "<w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>"
        "</w:numbering>"
    )


def app_xml() -> str:
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" "
        "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
        "<Application>docx-generator skill</Application>"
        "<DocSecurity>0</DocSecurity>"
        "<ScaleCrop>false</ScaleCrop>"
        "<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Title</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>"
        "<TitlesOfParts><vt:vector size=\"1\" baseType=\"lpstr\"><vt:lpstr>Document</vt:lpstr></vt:vector></TitlesOfParts>"
        "<Company></Company>"
        "<LinksUpToDate>false</LinksUpToDate>"
        "<SharedDoc>false</SharedDoc>"
        "<HyperlinksChanged>false</HyperlinksChanged>"
        "<AppVersion>16.0000</AppVersion>"
        "</Properties>"
    )


def core_xml(title: str, author: str, lang: str, subject: str | None, keywords: str | None) -> str:
    now = date.today().isoformat() + "T00:00:00Z"
    subject_xml = f"<dc:subject>{xml_escape(subject)}</dc:subject>" if subject else ""
    keywords_xml = f"<cp:keywords>{xml_escape(keywords)}</cp:keywords>" if keywords else ""
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<cp:coreProperties "
        "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" "
        "xmlns:dcterms=\"http://purl.org/dc/terms/\" "
        "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" "
        "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
        f"<dc:title>{xml_escape(title)}</dc:title>"
        f"<dc:language>{xml_escape(lang)}</dc:language>"
        f"<dc:creator>{xml_escape(author)}</dc:creator>"
        f"<cp:lastModifiedBy>{xml_escape(author)}</cp:lastModifiedBy>"
        f"{subject_xml}"
        f"{keywords_xml}"
        f"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:created>"
        f"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:modified>"
        "</cp:coreProperties>"
    )


def build_document_xml(lines: list[str], rels: Relationships) -> str:
    blocks = parse_blocks(lines, rels)
    body_parts: list[str] = []

    for block_type, text, level, items, body_rows in blocks:
        if block_type == Block.PARAGRAPH:
            if text:
                body_parts.append(paragraph(text, rels))
            else:
                body_parts.append(paragraph("", rels))

        elif block_type == Block.HEADING:
            body_parts.append(heading_paragraph(level, text, rels))

        elif block_type == Block.UNORDERED_LIST:
            for item in items:
                body_parts.append(list_paragraph(item, level, 1, rels))

        elif block_type == Block.ORDERED_LIST:
            for item in items:
                body_parts.append(list_paragraph(item, level, 2, rels))

        elif block_type == Block.QUOTE:
            body_parts.append(quote_paragraph(text, rels))

        elif block_type == Block.CODE_BLOCK:
            for line in items:
                body_parts.append(code_paragraph(line))

        elif block_type == Block.TABLE:
            if body_rows is not None:
                tbl, _ = table_xml(lines, lines.index(f"|{'|'.join(header := items)}|"), rels)
                body_parts.append(tbl)

        elif block_type == Block.HR:
            pass

    body = "".join(body_parts)
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
        "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
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


def generate_docx(
    output: Path,
    lines: list[str],
    title: str,
    author: str,
    lang: str,
    subject: str | None,
    keywords: str | None,
) -> None:
    content_types = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
        "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
        "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>"
        "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>"
        "</Types>"
    )
    rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
        "</Relationships>"
    )
    relationships = Relationships()
    document = build_document_xml(lines, relationships)

    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("docProps/core.xml", core_xml(title=title, author=author, lang=lang, subject=subject, keywords=keywords))
        docx.writestr("docProps/app.xml", app_xml())
        docx.writestr("word/document.xml", document)
        docx.writestr("word/styles.xml", styles_xml(lang))
        docx.writestr("word/numbering.xml", numbering_xml())
        docx.writestr("word/_rels/document.xml.rels", relationships.document_rels_xml())


def load_lines(args: argparse.Namespace) -> list[str]:
    lines: list[str] = []
    input_has_heading = False

    if args.input:
        text = args.input.read_text(encoding="utf-8")
        input_lines = text.splitlines()
        first_content = next((l.strip() for l in input_lines if l.strip()), "")
        input_has_heading = bool(HEADING_RE.match(first_content))
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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="DOCX/PDF generator skill script")
    parser.add_argument("--output", required=True, type=Path, help="Output path (.docx or .pdf)")
    parser.add_argument("--format", choices=["docx", "pdf"], help="Output format (default: inferred from --output extension)")
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
    return parser.parse_args()


def detect_format(output: Path, format_arg: str | None) -> str:
    if format_arg in ("docx", "pdf"):
        return format_arg
    ext = output.suffix.lower()
    if ext == ".pdf":
        return "pdf"
    return "docx"


def run_pdf(args: argparse.Namespace, lines: list[str]) -> None:
    if not check_reportlab():
        print("ERROR: La libreria 'reportlab' no esta instalada.")
        print("Por favor ejecuta: pip install reportlab")
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
    from parsers import Block

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

    def runs_to_markup(specs: list) -> str:
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
        from parsers import parse_inline_runs as pr
        specs = pr(text)
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
        from parsers import parse_inline_runs as pr
        markup = runs_to_markup(pr(text))
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

    from parsers import Relationships as Rels, parse_blocks as pb
    rels = Rels()
    blocks = pb(lines, rels)
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

    doc.build(story)
    print(args.output)


def main() -> None:
    args = parse_args()
    fmt = detect_format(args.output, args.format)

    if fmt == "pdf":
        if not check_python() or not check_reportlab():
            sys.exit(1)
        lines = load_lines(args)
        run_pdf(args, lines)
    else:
        lines = load_lines(args)
        generate_docx(
            args.output,
            lines,
            title=args.title or "Documento",
            author=args.author,
            lang=args.lang,
            subject=args.subject,
            keywords=args.keywords,
        )
        print(args.output)


if __name__ == "__main__":
    main()