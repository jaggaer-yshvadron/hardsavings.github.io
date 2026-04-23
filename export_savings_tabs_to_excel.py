from __future__ import annotations

from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape
import re
import zipfile


ROOT = Path("/home/yshvadro/Hard Savings")
SOURCE_HTML = ROOT / "research" / "savings_analysis_compiled.html"
OUTPUT_XLSX = ROOT / "research" / "hard_soft_savings_tabs.xlsx"

BLUE = "FF1D4ED8"
RED = "FFB91C1C"
PURPLE = "FF7C3AED"


@dataclass
class TextRun:
    text: str
    color: str | None = None
    bold: bool = False


@dataclass
class RichText:
    runs: list[TextRun] = field(default_factory=list)

    def add_text(self, text: str, color: str | None = None, bold: bool = False) -> None:
        clean = re.sub(r"\s+", " ", text)
        if not clean.strip():
            return
        if self.runs and not self.runs[-1].text.endswith(("\n", " ", "•")) and not clean.startswith((" ", ".", ",", ";", ":", ")", "]")):
            clean = " " + clean.lstrip()
        if self.runs and self.runs[-1].color == color and self.runs[-1].bold == bold:
            self.runs[-1].text += clean
        else:
            self.runs.append(TextRun(clean, color=color, bold=bold))

    def add_raw(self, text: str, color: str | None = None, bold: bool = False) -> None:
        if not text:
            return
        if self.runs and self.runs[-1].color == color and self.runs[-1].bold == bold:
            self.runs[-1].text += text
        else:
            self.runs.append(TextRun(text, color=color, bold=bold))

    def newline(self) -> None:
        if not self.runs:
            return
        if not self.runs[-1].text.endswith("\n"):
            self.runs[-1].text += "\n"

    def ensure_bullet(self) -> None:
        if self.runs and not self.runs[-1].text.endswith("\n"):
            self.newline()
        self.add_raw("• ")

    def plain_text(self) -> str:
        return "".join(run.text for run in self.runs).strip()


@dataclass
class TableData:
    headers: list[RichText] = field(default_factory=list)
    rows: list[list[RichText]] = field(default_factory=list)


@dataclass
class SectionData:
    title: str = ""
    note: RichText = field(default_factory=RichText)
    table: TableData = field(default_factory=TableData)
    trailing_heading: str | None = None
    trailing_items: list[RichText] = field(default_factory=list)


class SectionParser(HTMLParser):
    VOID_TAGS = {"br"}
    COLOR_CLASSES = {
        "metric-blue": BLUE,
        "formula-metric": BLUE,
        "metric-red": RED,
        "formula-metric-red": RED,
        "metric-purple": PURPLE,
    }

    def __init__(self, section_id: str) -> None:
        super().__init__(convert_charrefs=True)
        self.section_id = section_id
        self.section = SectionData()
        self._in_section = False
        self._section_depth = 0
        self._style_stack: list[tuple[str | None, bool]] = [(None, False)]
        self._heading_tag: str | None = None
        self._heading_text: list[str] = []
        self._in_note = False
        self._in_table = False
        self._current_row: list[RichText] | None = None
        self._current_cell: RichText | None = None
        self._current_cell_tag: str | None = None
        self._after_trailing_h3 = False
        self._collecting_trailing_ul = False
        self._current_trailing_item: RichText | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr_map = dict(attrs)

        if not self._in_section:
            if tag == "section" and attr_map.get("id") == self.section_id:
                self._in_section = True
                self._section_depth = 1
            return

        is_void = tag in self.VOID_TAGS
        if not is_void:
            self._section_depth += 1

        classes = set((attr_map.get("class") or "").split())
        parent_color, parent_bold = self._style_stack[-1]
        color = parent_color
        bold = parent_bold or tag == "strong"
        for cls, mapped in self.COLOR_CLASSES.items():
            if cls in classes:
                color = mapped
                bold = True
        if not is_void:
            self._style_stack.append((color, bold))

        if tag in {"h2", "h3"} and not self._in_table and not self._in_note:
            self._heading_tag = tag
            self._heading_text = []
        elif tag == "div" and "note" in classes:
            self._in_note = True
        elif tag == "br":
            if self._in_note:
                self.section.note.newline()
            elif self._current_cell is not None:
                self._current_cell.newline()
            elif self._current_trailing_item is not None:
                self._current_trailing_item.newline()
        elif tag == "table":
            self._in_table = True
        elif tag == "tr" and self._in_table:
            self._current_row = []
        elif tag in {"th", "td"} and self._in_table:
            self._current_cell = RichText()
            self._current_cell_tag = tag
        elif tag == "li":
            if self._current_cell is not None:
                self._current_cell.ensure_bullet()
            elif self._collecting_trailing_ul:
                self._current_trailing_item = RichText()
                self._current_trailing_item.ensure_bullet()
        elif tag == "ul" and self._after_trailing_h3 and not self._in_table:
            self._collecting_trailing_ul = True

    def handle_endtag(self, tag: str) -> None:
        if not self._in_section:
            return

        if tag == "section":
            self._section_depth -= 1
            if self._section_depth == 0:
                self._in_section = False
            return

        if tag == self._heading_tag:
            text = re.sub(r"\s+", " ", "".join(self._heading_text)).strip()
            if tag == "h2" and text:
                self.section.title = text
                self._after_trailing_h3 = False
            elif tag == "h3" and text:
                self.section.trailing_heading = text
                self._after_trailing_h3 = True
            self._heading_tag = None
            self._heading_text = []
        elif tag == "div" and self._in_note:
            self._in_note = False
        elif tag in {"th", "td"} and self._current_cell is not None and self._current_row is not None:
            self._current_row.append(self._current_cell)
            self._current_cell = None
            self._current_cell_tag = None
        elif tag == "tr" and self._current_row is not None:
            if self._current_row:
                if self.section.table.headers:
                    self.section.table.rows.append(self._current_row)
                else:
                    self.section.table.headers = self._current_row
            self._current_row = None
        elif tag == "table":
            self._in_table = False
        elif tag == "li" and self._current_trailing_item is not None:
            self.section.trailing_items.append(self._current_trailing_item)
            self._current_trailing_item = None
        elif tag == "ul" and self._collecting_trailing_ul:
            self._collecting_trailing_ul = False
            self._after_trailing_h3 = False

        self._section_depth -= 1
        if self._style_stack:
            self._style_stack.pop()

    def handle_data(self, data: str) -> None:
        if not self._in_section:
            return
        color, bold = self._style_stack[-1]
        if self._heading_tag is not None:
            self._heading_text.append(data)
        elif self._in_note:
            self.section.note.add_text(data, color=color, bold=bold)
        elif self._current_cell is not None:
            self._current_cell.add_text(data, color=color, bold=bold)
        elif self._current_trailing_item is not None:
            self._current_trailing_item.add_text(data, color=color, bold=bold)


def parse_section(html: str, section_id: str) -> SectionData:
    match = re.search(
        rf'<section\b[^>]*id="{re.escape(section_id)}"[^>]*>(.*?)</section>',
        html,
        flags=re.DOTALL,
    )
    if not match:
        raise ValueError(f"Could not find section {section_id!r} in {SOURCE_HTML}")

    parser = SectionParser(section_id)
    parser.feed(f'<section id="{section_id}">{match.group(1)}</section>')
    parser.close()
    return parser.section


def col_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def xml_text(value: str) -> str:
    return escape(value).replace("\n", "&#10;")


def rich_text_inline(value: RichText) -> str:
    if not value.runs:
        return "<is><t></t></is>"
    if len(value.runs) == 1 and not value.runs[0].color and not value.runs[0].bold:
        return f'<is><t xml:space="preserve">{xml_text(value.runs[0].text)}</t></is>'

    parts = ["<is>"]
    for run in value.runs:
        parts.append("<r>")
        if run.color or run.bold:
            parts.append("<rPr>")
            if run.bold:
                parts.append("<b/>")
            if run.color:
                parts.append(f'<color rgb="{run.color}"/>')
            parts.append("</rPr>")
        parts.append(f'<t xml:space="preserve">{xml_text(run.text)}</t>')
        parts.append("</r>")
    parts.append("</is>")
    return "".join(parts)


def make_cell(ref: str, rich: RichText, style_id: int) -> str:
    return f'<c r="{ref}" t="inlineStr" s="{style_id}">{rich_text_inline(rich)}</c>'


def make_plain_cell(ref: str, text: str, style_id: int) -> str:
    return make_cell(ref, RichText([TextRun(text)]), style_id)


def row_xml(row_num: int, values: Iterable[str], height: float | None = None) -> str:
    attrs = [f'r="{row_num}"']
    if height is not None:
        attrs.append(f'ht="{height}"')
        attrs.append('customHeight="1"')
    return f"<row {' '.join(attrs)}>{''.join(values)}</row>"


def build_sheet_xml(section: SectionData) -> str:
    widths = [18, 34, 40, 28, 34, 28]
    row_num = 1
    rows: list[str] = []
    merges = [f"A1:F1", f"A2:F3"]

    rows.append(
        row_xml(
            row_num,
            [make_plain_cell("A1", section.title, 1)],
            height=24,
        )
    )
    row_num += 1

    rows.append(
        row_xml(
            row_num,
            [make_cell("A2", section.note, 2)],
            height=36,
        )
    )
    row_num += 1
    rows.append(row_xml(row_num, [], height=12))
    row_num += 1
    rows.append(row_xml(row_num, [], height=8))
    row_num += 1

    header_cells = [
        make_cell(f"{col_letter(idx)}{row_num}", cell, 3)
        for idx, cell in enumerate(section.table.headers, start=1)
    ]
    rows.append(row_xml(row_num, header_cells, height=28))
    row_num += 1

    for data_row in section.table.rows:
        cells = [
            make_cell(f"{col_letter(idx)}{row_num}", cell, 4)
            for idx, cell in enumerate(data_row, start=1)
        ]
        rows.append(row_xml(row_num, cells, height=84))
        row_num += 1

    if section.trailing_heading:
        merges.append(f"A{row_num}:F{row_num}")
        rows.append(
            row_xml(
                row_num,
                [make_plain_cell(f"A{row_num}", section.trailing_heading, 5)],
                height=22,
            )
        )
        row_num += 1
        for item in section.trailing_items:
            merges.append(f"A{row_num}:F{row_num}")
            rows.append(row_xml(row_num, [make_cell(f"A{row_num}", item, 4)], height=34))
            row_num += 1

    cols_xml = "".join(
        f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>'
        for idx, width in enumerate(widths, start=1)
    )
    merge_xml = "".join(f"<mergeCell ref=\"{ref}\"/>" for ref in merges)
    max_row = row_num - 1
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:F{max_row}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>{cols_xml}</cols>
  <sheetData>{''.join(rows)}</sheetData>
  <mergeCells count="{len(merges)}">{merge_xml}</mergeCells>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font>
      <sz val="11"/>
      <color rgb="FF1F2937"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color rgb="FF111827"/>
      <name val="Arial"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FF111827"/>
      <name val="Arial"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFCE7EF"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF6F1E7"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFD6D3D1"/></left>
      <right style="thin"><color rgb="FFD6D3D1"/></right>
      <top style="thin"><color rgb="FFD6D3D1"/></top>
      <bottom style="thin"><color rgb="FFD6D3D1"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="top" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="top" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center" wrapText="1"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"""


def write_xlsx(path: Path, hard: SectionData, soft: SectionData) -> None:
    sheet1 = build_sheet_xml(hard)
    sheet2 = build_sheet_xml(soft)

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
""",
        )
        zf.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
""",
        )
        zf.writestr(
            "docProps/core.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Hard and Soft Savings Tabs</dc:title>
  <dc:creator>Codex</dc:creator>
</cp:coreProperties>
""",
        )
        zf.writestr(
            "docProps/app.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>
""",
        )
        zf.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Hard Savings" sheetId="1" r:id="rId1"/>
    <sheet name="Soft Savings" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
""",
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
""",
        )
        zf.writestr("xl/styles.xml", styles_xml())
        zf.writestr("xl/worksheets/sheet1.xml", sheet1)
        zf.writestr("xl/worksheets/sheet2.xml", sheet2)


def main() -> None:
    html = SOURCE_HTML.read_text(encoding="utf-8")
    hard = parse_section(html, "panel-hard")
    soft = parse_section(html, "panel-soft")
    write_xlsx(OUTPUT_XLSX, hard, soft)
    print(f"Exported {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
