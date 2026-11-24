"""Regression checks for compact report layout, including native Word/WPS pages.

Temporary fixtures never write to the user's desktop or change the template.
"""

from __future__ import annotations

import io
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from fill_report import (
    Figs, _pagination, as_item, fill_results, fill_review, fill_section, find_cell,
    remove_template_check,
)
from report_layout import normalize_report_layout


CONTENT_KEYS = (
    "目的与要求",
    "主要材料与仪器设备",
    "内容、方法与步骤",
    "结果与问题讨论",
)
TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "Template.docx"


class ReportLayoutRegressionTests(unittest.TestCase):
    def setUp(self):
        self.doc = Document(str(TEMPLATE))
        figures = Figs()
        for key, count in zip(CONTENT_KEYS[:3], (10, 13, 30)):
            fill_section(
                find_cell(self.doc, key),
                [as_item(f"{key}测试项目第{index + 1}项。") for index in range(count)],
                figures,
                allow_images=False,
            )
        fill_results(
            find_cell(self.doc, CONTENT_KEYS[3]),
            [as_item("实验结果符合预期。")],
            [],
            figures,
        )
        fill_review(find_cell(self.doc, "教师评阅"), "报告结构完整。", "")
        self.cells = [find_cell(self.doc, key) for key in CONTENT_KEYS]

    def test_handwriting_floors_removed_without_altering_report_content(self):
        """The screenshot's steps row must fit remaining space without a floor."""
        original_text = [node.text for node in self.doc.element.iter(qn("w:t"))]
        original_cover = [
            row.xml for row in self.doc.tables[0]._tbl.findall(qn("w:tr"))[:4]
        ]
        original_review = find_cell(self.doc, "教师评阅")._tc.getparent().xml
        original_sections = [section._sectPr.xml for section in self.doc.sections]
        original_heights = sum(
            len(cell._tc.getparent().findall("w:trPr/w:trHeight", cell._tc.nsmap))
            for cell in self.cells
        )
        self.assertGreater(original_heights, 0, "Fixture must exercise template row floors")

        changes = normalize_report_layout(self.doc, self.cells)

        self.assertEqual(changes["removed_row_heights"], original_heights)
        self.assertEqual(original_text, [node.text for node in self.doc.element.iter(qn("w:t"))])
        self.assertEqual(original_cover, [
            row.xml for row in self.doc.tables[0]._tbl.findall(qn("w:tr"))[:4]
        ])
        self.assertEqual(original_review, find_cell(self.doc, "教师评阅")._tc.getparent().xml)
        self.assertEqual(original_sections, [section._sectPr.xml for section in self.doc.sections])
        for cell in self.cells:
            row_properties = cell._tc.getparent().find(qn("w:trPr"))
            self.assertIsNone(row_properties.find(qn("w:trHeight")))
            self.assertEqual(row_properties.find(qn("w:cantSplit")).get(qn("w:val")), "0")
            self.assertFalse(cell.paragraphs[0].paragraph_format.keep_with_next)
            for paragraph in cell.paragraphs[1:]:
                self.assertFalse(paragraph.paragraph_format.keep_with_next)
                self.assertFalse(paragraph.paragraph_format.keep_together)
                self.assertFalse(paragraph.paragraph_format.page_break_before)

        # The flags must survive DOCX serialization, which is what Word reads.
        saved = io.BytesIO()
        self.doc.save(saved)
        saved.seek(0)
        reopened = Document(saved)
        for key in CONTENT_KEYS:
            self.assertIsNone(
                find_cell(reopened, key)._tc.getparent().find("w:trPr/w:trHeight", reopened.element.nsmap)
            )

    def test_reused_template_pagination_locks_are_released(self):
        """A reused form's paragraph/row locks cannot reintroduce the blank band."""
        for cell in self.cells:
            row_properties = cell._tc.getparent().get_or_add_trPr()
            lock = OxmlElement("w:cantSplit")
            row_properties.append(lock)
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.keep_with_next = True
                paragraph.paragraph_format.keep_together = True
                paragraph.paragraph_format.page_break_before = True

        normalize_report_layout(self.doc, self.cells)
        for cell in self.cells:
            self.assertEqual(
                cell._tc.getparent().find("w:trPr/w:cantSplit", cell._tc.nsmap).get(qn("w:val")),
                "0",
            )
            self.assertFalse(cell.paragraphs[0].paragraph_format.keep_with_next)
            self.assertFalse(cell.paragraphs[-1].paragraph_format.keep_with_next)
            self.assertFalse(cell.paragraphs[-1].paragraph_format.keep_together)
            self.assertFalse(cell.paragraphs[0].paragraph_format.page_break_before)
        once = self.doc.element.xml
        changes = normalize_report_layout(self.doc, self.cells)
        self.assertEqual(changes["removed_row_heights"], 0)
        self.assertEqual(once, self.doc.element.xml, "Repeated normalization must be idempotent")

    def test_only_empty_table_separators_are_compacted(self):
        """Mandatory blank paragraphs shrink, but a meaningful field is preserved."""
        body = self.doc.element.body
        between = next(
            child for child in body
            if child.tag == qn("w:p")
            and child.getprevious() is not None
            and child.getprevious().tag == qn("w:tbl")
            and child.getnext() is not None
            and child.getnext().tag == qn("w:tbl")
        )
        field = OxmlElement("w:fldSimple")
        field.set(qn("w:instr"), "PAGE")
        between.append(field)
        original_field_paragraph = between.xml

        changes = normalize_report_layout(self.doc, self.cells)

        self.assertEqual(between.xml, original_field_paragraph)
        self.assertEqual(changes["compacted_separators"], 1)
        tail = list(body)[-2]
        spacing = tail.find("w:pPr/w:spacing", tail.nsmap)
        self.assertEqual(spacing.get(qn("w:line")), "20")
        self.assertEqual(spacing.get(qn("w:lineRule")), "exact")

    def test_native_headings_use_first_page_remaining_space(self):
        """Catch WPS's table keep-next chain, which XML-only tests cannot see."""
        remove_template_check(self.doc)
        normalize_report_layout(self.doc, self.cells)
        prefix = "JRLayout_"
        expected = {}
        used_ids = [
            int(node.get(qn("w:id")))
            for node in self.doc.element.iter(qn("w:bookmarkStart"))
            if (node.get(qn("w:id")) or "").isdigit()
        ]
        next_id = max(used_ids, default=-1) + 1
        for index, cell in enumerate(self.cells[:3]):
            name = prefix + ("Purpose", "Materials", "Steps")[index]
            expected[name] = 1
            paragraph = cell.paragraphs[0]._p
            start = OxmlElement("w:bookmarkStart")
            start.set(qn("w:id"), str(next_id + index))
            start.set(qn("w:name"), name)
            end = OxmlElement("w:bookmarkEnd")
            end.set(qn("w:id"), str(next_id + index))
            offset = 1 if len(paragraph) and paragraph[0].tag == qn("w:pPr") else 0
            paragraph.insert(offset, start)
            paragraph.insert(offset + 1, end)
        with tempfile.TemporaryDirectory(prefix="jr-layout-native-") as folder:
            scratch = Path(folder)
            fixture = scratch / "filled_report.docx"
            self.doc.save(fixture)
            measured = _pagination(fixture, prefix, scratch)
        self.assertEqual(measured["positions"], expected,
                         "Text sections must use page 1 instead of leaving a cover-only page")


if __name__ == "__main__":
    unittest.main()
