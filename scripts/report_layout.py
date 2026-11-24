"""Let filled report sections flow naturally without template writing-space floors.

The blank form reserves large handwriting areas using minimum row heights.
Those floors are inappropriate after the cells contain typed report content:
Word/WPS may move the next row to a fresh page when its minimum cannot fit.
Only filled content rows are normalized; the cover and teacher/signature area
retain their template geometry.
"""

from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph


def _set_on_off(properties, tag: str, enabled: bool) -> None:
    node = properties.find(qn(tag))
    if node is None:
        node = OxmlElement(tag)
        properties.append(node)
    node.set(qn("w:val"), "1" if enabled else "0")


def _compact_separator(element, doc) -> None:
    paragraph = Paragraph(element, doc._body)
    formatting = paragraph.paragraph_format
    formatting.space_before = Pt(0)
    formatting.space_after = Pt(0)
    formatting.line_spacing = Pt(1)
    formatting.keep_with_next = False
    formatting.keep_together = False
    formatting.page_break_before = False
    properties = element.get_or_add_pPr()
    run_properties = properties.find(qn("w:rPr"))
    if run_properties is None:
        run_properties = OxmlElement("w:rPr")
        properties.append(run_properties)
    for name in ("w:sz", "w:szCs"):
        size = run_properties.find(qn(name))
        if size is None:
            size = OxmlElement(name)
            run_properties.append(size)
        size.set(qn("w:val"), "2")


def normalize_report_layout(doc, content_cells) -> dict[str, int]:
    """Normalize the four filled report cells before physical pagination.

    Pass purpose, materials, steps and results cells, in their document order.
    Existing text, borders, fonts and page margins are preserved. Paragraphs
    remain free to break across pages. Do not apply keep-with-next to headings
    inside these large table cells: WPS can treat that as a row-level chain and
    push the entire body past the cover page. Images remain indivisible and
    need separate size fitting.
    """
    removed_heights = 0
    seen_rows = set()
    for cell in content_cells:
        row = cell._tc.getparent()
        if row.tag != qn("w:tr"):
            raise ValueError("A report content cell must belong to a table row")
        if row not in seen_rows:
            seen_rows.add(row)
            properties = row.get_or_add_trPr()
            for height in list(properties.findall(qn("w:trHeight"))):
                properties.remove(height)
                removed_heights += 1
            # Explicit false also overrides a table style's inherited setting.
            _set_on_off(properties, "w:cantSplit", False)
        paragraphs = cell.paragraphs
        for paragraph in paragraphs:
            formatting = paragraph.paragraph_format
            formatting.keep_with_next = False
            formatting.keep_together = False
            formatting.page_break_before = False

    compacted_separators = 0
    children = list(doc.element.body)
    for index, element in enumerate(children):
        if element.tag != qn("w:p") or index == 0:
            continue
        if children[index - 1].tag != qn("w:tbl"):
            continue
        next_tag = children[index + 1].tag if index + 1 < len(children) else None
        if next_tag not in (qn("w:tbl"), qn("w:sectPr"), None):
            continue
        # Only the mandatory/separating empty paragraphs, never a break,
        # drawing, field, bookmark or section boundary supplied by the user.
        if any(child.tag != qn("w:pPr") for child in element):
            continue
        properties = element.find(qn("w:pPr"))
        if properties is not None and properties.find(qn("w:sectPr")) is not None:
            continue
        _compact_separator(element, doc)
        compacted_separators += 1

    return {"removed_row_heights": removed_heights,
            "compacted_separators": compacted_separators}
