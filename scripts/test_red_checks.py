"""Regression checks. Native tests require installed Word/WPS.

Run: python scripts/test_red_checks.py
Geometry tests only: python -m unittest test_red_checks.CheckGeometryTests
"""

import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.image.image import Image as DocxImage
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm

import fill_report


STAMP_PREFIX = "JRTrainingCheck_Page_"


def page_setup(width_mm=210, height_mm=297, *, left_mm=25, right_mm=25,
               top_mm=25, bottom_mm=25, gutter_mm=0, gutter_pos=0,
               mirror=False):
    return {
        "width": int(Mm(width_mm)), "height": int(Mm(height_mm)),
        "left": int(Mm(left_mm)), "right": int(Mm(right_mm)),
        "top": int(Mm(top_mm)), "bottom": int(Mm(bottom_mm)),
        "gutter": int(Mm(gutter_mm)), "gutter_pos": gutter_pos,
        "mirror": mirror,
    }


def check_anchors(doc):
    result = {}
    for anchor in doc.element.iter(qn("wp:anchor")):
        properties = anchor.find(qn("wp:docPr"))
        if properties is None or not properties.get("name", "").startswith(STAMP_PREFIX):
            continue
        name = properties.get("name")
        if name in result:
            raise AssertionError("Duplicate check: " + name)
        extent = anchor.find(qn("wp:extent"))
        horizontal = anchor.find(qn("wp:positionH"))
        vertical = anchor.find(qn("wp:positionV"))
        result[name] = {
            "left": int(horizontal.find(qn("wp:posOffset")).text),
            "top": int(vertical.find(qn("wp:posOffset")).text),
            "width": int(extent.get("cx")), "height": int(extent.get("cy")),
        }
    return result


class GeometryAssertions:
    def assert_fitted_check(self, geometry, setup, page=1, tolerance=1):
        left, right = setup["left"], setup["right"]
        if setup["mirror"] and page % 2 == 0:
            left, right = right, left
        top, bottom = setup["top"], setup["bottom"]
        if setup["gutter_pos"] == 1:
            top += setup["gutter"]
        elif setup["gutter_pos"] == 2 or (setup["mirror"] and page % 2 == 0):
            right += setup["gutter"]
        else:
            left += setup["gutter"]
        content_width = setup["width"] - left - right
        content_height = setup["height"] - top - bottom
        self.assertGreater(geometry["width"], 0)
        self.assertGreater(geometry["height"], 0)
        self.assertGreaterEqual(geometry["left"], left - tolerance)
        self.assertGreaterEqual(geometry["top"], top - tolerance)
        self.assertLessEqual(geometry["left"] + geometry["width"], setup["width"] - right + tolerance)
        self.assertLessEqual(geometry["top"] + geometry["height"], setup["height"] - bottom + tolerance)
        self.assertLessEqual(geometry["width"], content_width * 0.60 + tolerance)
        self.assertLessEqual(geometry["height"], content_height * 0.55 + tolerance)
        image = DocxImage.from_file(str(fill_report.check_image_path()))
        self.assertAlmostEqual(geometry["width"] / geometry["height"],
                               image.px_width / image.px_height, places=5)


class CheckGeometryTests(GeometryAssertions, unittest.TestCase):
    def test_page_sizes_orientations_and_unequal_margins(self):
        fixtures = {
            "A4": page_setup(),
            "A5": page_setup(148, 210, left_mm=15, right_mm=15,
                             top_mm=15, bottom_mm=15),
            "A4 landscape": page_setup(297, 210),
            "unequal margins": page_setup(left_mm=43, right_mm=11,
                                          top_mm=36, bottom_mm=18),
        }
        for label, setup in fixtures.items():
            with self.subTest(page=label):
                self.assert_fitted_check(fill_report._check_geometry(setup, 1), setup)

    def test_check_changes_size_when_available_page_area_changes(self):
        small = fill_report._check_geometry(page_setup(148, 210), 1)
        large = fill_report._check_geometry(page_setup(210, 297), 1)
        self.assertLess(small["width"], large["width"])
        self.assertLess(small["height"], large["height"])

    def test_mirrored_margins_and_gutters_stay_inside_content(self):
        for gutter_pos in (0, 1, 2):
            for mirror in (False, True):
                setup = page_setup(left_mm=37, right_mm=16, gutter_mm=9,
                                   gutter_pos=gutter_pos, mirror=mirror)
                for page in (1, 2):
                    with self.subTest(gutter_pos=gutter_pos, mirror=mirror, page=page):
                        self.assert_fitted_check(fill_report._check_geometry(setup, page),
                                                 setup, page)

    def test_invalid_page_dimensions_stop_stamping(self):
        for label, changes in (
            ("missing width", {"width": None}),
            ("zero page width", {"width": 0}),
            ("negative margin", {"top": -1}),
            ("no writable width", {"left": int(Mm(210))}),
            ("invalid height", {"height": "unknown"}),
        ):
            setup = page_setup()
            setup.update(changes)
            with self.subTest(case=label), self.assertRaises(fill_report.ReportError):
                fill_report._check_geometry(setup, 1)


class CheckRemovalTests(unittest.TestCase):
    def test_removes_tagged_and_legacy_checks_preserving_content_pictures(self):
        doc = Document()
        for paragraph in (doc.add_paragraph("Body content"),
                          doc.sections[0].header.paragraphs[0]):
            ordinary = paragraph.add_run().add_picture(str(fill_report.check_image_path()),
                                                       width=Inches(0.3))
            ordinary._inline.find(qn("wp:docPr")).set("name", "SourcePicture")
            dynamic = paragraph.add_run().add_picture(str(fill_report.check_image_path()),
                                                      width=Inches(2))
            dynamic._inline.find(qn("wp:docPr")).set("name", STAMP_PREFIX + "1")
            paragraph.add_run().add_picture(str(fill_report.check_image_path()),
                                             width=5_466_080, height=6_363_335)
        fill_report.remove_template_check(doc)
        self.assertEqual(doc.paragraphs[0].text, "Body content")
        for root in (doc.element, doc.sections[0].header._element):
            self.assertEqual([node.get("name") for node in root.iter(qn("wp:docPr"))],
                             ["SourcePicture"])


class LastPageChecks(unittest.TestCase):
    def test_physical_pages(self):
        for page_count, pictures in ((1, True), (3, False), (3, True)):
            with self.subTest(page_count=page_count, pictures=pictures):
                doc = Document()
                for page in range(1, page_count + 1):
                    if page > 1:
                        doc.add_page_break()
                    doc.add_paragraph("Test page " + str(page))
                    doc.add_paragraph("Discussion and comment remain unchanged.")
                    if pictures:
                        # A deliberately small image is source content, not a stamp.
                        doc.add_picture(str(fill_report.check_image_path()), width=Inches(0.3))
                text_before = [p.text for p in doc.paragraphs]
                fill_report.stamp_check_except_last_page(doc)
                self.assertEqual(text_before, [p.text for p in doc.paragraphs])
                self.assertCountEqual(check_anchors(doc),
                                      [STAMP_PREFIX + str(n) for n in range(1, page_count)])
                self.assertEqual(len(list(doc.element.iter(qn("wp:inline")))),
                                 page_count if pictures else 0)
                for node in doc.element.iter(qn("w:bookmarkStart")):
                    self.assertFalse(node.get(qn("w:name"), "").startswith("JRCheck_"))
                # Stamping a reused report must replace its variable-size checks.
                if page_count == 3 and pictures:
                    first_checks = check_anchors(doc)
                    fill_report.stamp_check_except_last_page(doc)
                    self.assertEqual(check_anchors(doc), first_checks)
                    self.assertEqual(text_before, [p.text for p in doc.paragraphs])
                    self.assertEqual(len(list(doc.element.iter(qn("wp:inline")))), page_count)
                with tempfile.TemporaryDirectory(prefix="jr-check-test-") as folder:
                    doc.save(Path(folder) / "report.docx")


class MixedSectionChecks(GeometryAssertions, unittest.TestCase):
    def test_native_page_sizes_and_physical_last_page_with_number_restarts(self):
        doc = Document()
        fixtures = [
            page_setup(),
            page_setup(148, 210, left_mm=15, right_mm=19, top_mm=16, bottom_mm=20),
            page_setup(297, 210, left_mm=22, right_mm=31, top_mm=17, bottom_mm=21),
            page_setup(148, 210, left_mm=15, right_mm=15, top_mm=15, bottom_mm=15),
        ]
        for page, setup in enumerate(fixtures, 1):
            section = doc.sections[0] if page == 1 else doc.add_section(WD_SECTION_START.NEW_PAGE)
            section.page_width, section.page_height = setup["width"], setup["height"]
            section.left_margin, section.right_margin = setup["left"], setup["right"]
            section.top_margin, section.bottom_margin = setup["top"], setup["bottom"]
            numbering = OxmlElement("w:pgNumType")
            numbering.set(qn("w:start"), "1")
            section._sectPr.append(numbering)
            doc.add_paragraph("Physical page " + str(page) + ": display numbering restarts at 1.")
            doc.add_paragraph("This content must stay on its original page.")
        text_before = [paragraph.text for paragraph in doc.paragraphs]
        fill_report.stamp_check_except_last_page(doc)
        checks = check_anchors(doc)
        self.assertCountEqual(checks, [STAMP_PREFIX + str(page) for page in (1, 2, 3)])
        for page in (1, 2, 3):
            # Word stores page lengths in twentieths of a point. Allow the
            # resulting rounding when comparing its geometry with millimetres.
            self.assert_fitted_check(checks[STAMP_PREFIX + str(page)], fixtures[page - 1],
                                     page, tolerance=1270)
        self.assertNotEqual(checks[STAMP_PREFIX + "1"]["width"],
                            checks[STAMP_PREFIX + "2"]["width"])
        self.assertNotEqual(checks[STAMP_PREFIX + "2"]["height"],
                            checks[STAMP_PREFIX + "3"]["height"])
        self.assertEqual(text_before, [paragraph.text for paragraph in doc.paragraphs])


if __name__ == "__main__":
    unittest.main()
