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
from typing import Optional, Tuple


HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
UNORDERED_ITEM_RE = re.compile(r"^(\s*)[-*]\s+(.*)$")
ORDERED_ITEM_RE = re.compile(r"^(\s*)(\d+)\.\s+(.*)$")
TABLE_SEPARATOR_RE = re.compile(r"^\s*\|?(?:\s*:?-{3,}:?\s*\|)+\s*:?-{3,}:?\s*\|?\s*$")
QUOTE_RE = re.compile(r"^\s*>\s?(.*)$")
HR_RE = re.compile(r"^\s*[-*_]{3,}\s*$")


RunSpec = Tuple[str, bool, bool, bool, Optional[str]]


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
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>",
        ]
        for url, rid in self._link_to_rid.items():
            rels.append(
                "<Relationship "
                f"Id=\"{rid}\" "
                "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" "
                f"Target=\"{xml_escape_attr(url)}\" "
                "TargetMode=\"External\"/>"
            )
        rels.append("</Relationships>")
        return "".join(rels)


def xml_escape_attr(text: str) -> str:
    return xml_escape(text).replace('"', "&quot;")


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


def _parse_emphasis(text: str, bold: bool = False, italic: bool = False) -> list[RunSpec]:
    runs: list[RunSpec] = []
    plain_buffer: list[str] = []
    index = 0

    def flush_plain() -> None:
        if plain_buffer:
            runs.append(("".join(plain_buffer), bold, italic, False, None))
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


def _parse_links_and_emphasis(text: str, rels: Relationships) -> list[RunSpec]:
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
            rid = rels.get_hyperlink_rid(url)
            for run_text, run_bold, run_italic, _, _ in _parse_emphasis(label):
                runs.append((run_text, run_bold, run_italic, False, rid))
        else:
            runs.extend(_parse_emphasis(text[start:close_url + 1]))

        index = close_url + 1

    if not runs:
        return [("", False, False, False, None)]
    return runs


def parse_inline_runs(text: str, rels: Relationships) -> list[RunSpec]:
    runs: list[RunSpec] = []
    for segment, is_code in _split_code_spans(text):
        if is_code:
            runs.append((segment, False, False, True, None))
            continue
        runs.extend(_parse_links_and_emphasis(segment, rels))

    if not runs:
        return [("", False, False, False, None)]
    return runs


def xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def render_runs(specs: list[RunSpec]) -> str:
    pieces: list[str] = []
    current_hyperlink: str | None = None
    hyperlink_runs: list[str] = []

    def flush_link() -> None:
        nonlocal current_hyperlink
        if current_hyperlink is not None:
            pieces.append(f'<w:hyperlink r:id="{current_hyperlink}" w:history="1">{"".join(hyperlink_runs)}</w:hyperlink>')
            hyperlink_runs.clear()
            current_hyperlink = None

    for content, bold, italic, code, hyperlink_rid in specs:
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
    runs = render_runs(parse_inline_runs(text, rels))
    return f"<w:p>{runs}</w:p>"


def paragraph_with_props(text: str, ppr: str, rels: Relationships) -> str:
    runs = render_runs(parse_inline_runs(text, rels))
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


def split_table_row(row: str) -> list[str]:
    stripped = row.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    return [cell.strip() for cell in stripped.split("|")]


def table_xml(lines: list[str], start_index: int, rels: Relationships) -> tuple[str, int]:
    header = split_table_row(lines[start_index])
    index = start_index + 2
    body_rows: list[list[str]] = []

    while index < len(lines):
        candidate = lines[index].strip()
        if not candidate or "|" not in candidate or TABLE_SEPARATOR_RE.match(candidate):
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
            text = f"**{cell}**" if header_row else cell
            chunks.append(
                "<w:tc>"
                f'<w:tcPr><w:tcW w:w="{col_width}" w:type="dxa"/></w:tcPr>'
                f"{paragraph(text, rels)}"
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


def is_markdown_table_start(lines: list[str], index: int) -> bool:
    if index + 1 >= len(lines):
        return False
    header = lines[index].strip()
    separator = lines[index + 1].strip()
    return "|" in header and bool(TABLE_SEPARATOR_RE.match(separator))


def block_paragraphs(lines: list[str], rels: Relationships) -> list[str]:
    blocks: list[str] = []
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
            blocks.append(code_paragraph(line))
            index += 1
            continue

        if HR_RE.match(stripped):
            index += 1
            continue

        if is_markdown_table_start(lines, index):
            table, index = table_xml(lines, index, rels)
            blocks.append(table)
            continue

        if not stripped:
            blocks.append(paragraph("", rels))
            index += 1
            continue

        heading_match = HEADING_RE.match(line)
        if heading_match:
            level = len(heading_match.group(1))
            blocks.append(heading_paragraph(level, heading_match.group(2).strip(), rels))
            index += 1
            continue

        unordered_match = UNORDERED_ITEM_RE.match(line)
        if unordered_match:
            indent = len(unordered_match.group(1).replace("\t", "    "))
            blocks.append(list_paragraph(unordered_match.group(2), indent // 2, 1, rels))
            index += 1
            continue

        ordered_match = ORDERED_ITEM_RE.match(line)
        if ordered_match:
            indent = len(ordered_match.group(1).replace("\t", "    "))
            blocks.append(list_paragraph(ordered_match.group(3), indent // 2, 2, rels))
            index += 1
            continue

        quote_match = QUOTE_RE.match(line)
        if quote_match:
            blocks.append(quote_paragraph(quote_match.group(1), rels))
            index += 1
            continue

        blocks.append(paragraph(line, rels))
        index += 1

    if in_code_block:
        blocks.append(code_paragraph(""))

    return blocks


def build_document_xml(lines: list[str], rels: Relationships) -> str:
    body = "".join(block_paragraphs(lines, rels))
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
    parser = argparse.ArgumentParser(description="DOCX generator skill script")
    parser.add_argument("--output", required=True, type=Path, help="Output .docx path")
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


def main() -> None:
    args = parse_args()
    lines = load_lines(args)
    title = args.title or "Documento"
    generate_docx(
        args.output,
        lines,
        title=title,
        author=args.author,
        lang=args.lang,
        subject=args.subject,
        keywords=args.keywords,
    )
    print(args.output)


if __name__ == "__main__":
    main()
