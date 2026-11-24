# -*- coding: utf-8 -*-
"""Fill the fixed training-report template. The blank template is never modified.

The check is a page-positioned transparent picture in the document body.
Word/WPS measures physical pages first; the final page is never stamped.
No repeating header check is used, including for a single-page report.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import re
import shutil
import sys
import subprocess
import uuid
import tempfile
import winreg
from pathlib import Path

try:
    from docx import Document
    from docx.image.image import Image as DocxImage
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from docx.table import _Cell
    from report_layout import normalize_report_layout
except ImportError:
    print(
        "ERROR: python-docx is missing. Run: python -m pip install python-docx",
        file=sys.stderr,
    )
    sys.exit(1)


MAX_REVIEW_WIDTH = 70
MAX_IMAGE_WIDTH_EMU = 13 * 360000
MAX_IMAGE_HEIGHT_EMU = 15 * 360000
WEEKDAYS = "一二三四五六日"
# Used only to recognize the original template ink, never for new placement.
LEGACY_CHECK_WIDTH_EMU = 5_466_080
LEGACY_CHECK_HEIGHT_EMU = 6_363_335
CHECK_NAME_PREFIX = "JRTrainingCheck_Page_"


class ReportError(Exception):
    def __init__(self, message: str, code: int = 1):
        super().__init__(message)
        self.code = code


class Figs:
    def __init__(self) -> None:
        self.n = 1


def find_desktop() -> Path:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        ) as key:
            raw, _ = winreg.QueryValueEx(key, "Desktop")
        desktop = Path(os.path.expandvars(raw))
        if desktop.is_dir():
            return desktop
    except OSError:
        pass
    home = Path.home()
    for candidate in (
        home / "Desktop",
        home / "OneDrive" / "Desktop",
        home / "OneDrive" / "桌面",
    ):
        if candidate.is_dir():
            return candidate
    fallback = home / "Desktop"
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback


def find_template() -> Path:
    bundled = Path(__file__).resolve().parent.parent / "assets" / "Template.docx"
    desktop_template = find_desktop() / "Template.docx"
    if desktop_template.is_file():
        return desktop_template
    if bundled.is_file():
        return bundled
    raise ReportError(
        "Template.docx was not found. Put the blank template on the desktop, or keep assets/Template.docx inside this skill."
    )


def sanitize_stem(name: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]', "", name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().rstrip(".")
    if len(cleaned) > 80:
        cleaned = cleaned[:80].rstrip()
    if cleaned.lower() == "template":
        return ""
    return cleaned


def unique_docx(desktop: Path, stem: str) -> Path:
    path = desktop / f"{stem}.docx"
    if not path.exists():
        return path
    index = 2
    while True:
        candidate = desktop / f"{stem}_{index}.docx"
        if not candidate.exists():
            return candidate
        index += 1


def display_width(text: str) -> float:
    width = 0.0
    for char in text:
        width += 1.0 if ord(char) > 127 else 0.5
    return width


def strip_number(text: str) -> str:
    return re.sub(r"^(\d+[.、．)]|（\d+）|\(\d+\))\s*", "", text.strip()).strip()


def cell_plain_text(cell) -> str:
    return "".join(paragraph.text or "" for paragraph in cell.paragraphs)


def iter_cells(doc):
    """Walk real table cells. row.cells drops some horizontally merged cells.

    Hold every cell before deduplicating. lxml can recycle element ids if a
    previous cell is no longer referenced, which would skip later cells.
    """
    cells = []
    for table in doc.tables:
        for tc in table._tbl.iter(qn("w:tc")):
            cells.append(_Cell(tc, table))
    unique = []
    seen = set()
    for cell in cells:
        key = id(cell._tc)
        if key in seen:
            continue
        seen.add(key)
        unique.append(cell)
    return unique


def find_cell(doc, keyword: str):
    for cell in iter_cells(doc):
        first = cell.paragraphs[0].text if cell.paragraphs else ""
        if keyword in first:
            return cell
    raise ReportError(f"Template section not found: {keyword}")


def paragraph_has_drawing(paragraph) -> bool:
    return any(element.tag.endswith("drawing") for element in paragraph._p.iter())


def delete_paragraph(paragraph) -> None:
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def set_run_font(run, size_pt: int = 12) -> None:
    run.bold = False
    run.italic = False
    run.font.size = Pt(size_pt)
    run.font.name = "宋体"
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        r_fonts.set(qn(f"w:{attr}"), "宋体")
    for tag in ("sz", "szCs"):
        node = r_pr.find(qn(f"w:{tag}"))
        if node is None:
            node = OxmlElement(f"w:{tag}")
            r_pr.append(node)
        node.set(qn("w:val"), str(int(size_pt * 2)))


def apply_paragraph_format(
    paragraph,
    *,
    indent_chars: int = 2,
    center: bool = False,
    exact: bool = True,
    before: int = 0,
    after: int = 0,
) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    jc = p_pr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        p_pr.append(jc)
    jc.set(qn("w:val"), "center" if center else "left")

    indent = p_pr.find(qn("w:ind"))
    if indent is None:
        indent = OxmlElement("w:ind")
        p_pr.append(indent)
    indent.set(qn("w:left"), "0")
    indent.set(qn("w:leftChars"), "0")
    indent.set(qn("w:right"), "0")
    indent.set(qn("w:rightChars"), "0")
    if center or indent_chars <= 0:
        indent.set(qn("w:firstLine"), "0")
        indent.set(qn("w:firstLineChars"), "0")
    else:
        indent.set(qn("w:firstLineChars"), str(indent_chars * 100))
        indent.set(qn("w:firstLine"), str(indent_chars * 12 * 20))

    spacing = p_pr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        p_pr.append(spacing)
    spacing.set(qn("w:before"), str(before))
    spacing.set(qn("w:after"), str(after))
    if exact:
        spacing.set(qn("w:line"), "440")
        spacing.set(qn("w:lineRule"), "exact")
    else:
        # Images sit in the paragraph. Exact 22-pt spacing would crop them.
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")


def replace_runs(paragraph, text: str):
    parent = paragraph._p
    for child in list(parent):
        if child.tag == qn("w:r"):
            parent.remove(child)
    run = paragraph.add_run(text)
    set_run_font(run)
    return run


def append_value(paragraph, value: str) -> None:
    run = paragraph.add_run(value.strip())
    set_run_font(run)


def rewrite_label_value(paragraph, new_value: str) -> None:
    """Replace the text after a label and keep the label itself."""
    match = re.match(r"^(\s*\S+?[:：]\s*)", paragraph.text or "")
    prefix = match.group(1) if match else (paragraph.text or "")
    replace_runs(paragraph, prefix + new_value.strip())


def format_time(text: str) -> tuple[str, str]:
    raw = re.sub(r"^时间[:：]\s*", "", text.strip())
    match = re.search(
        r"(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?",
        raw,
    )
    if not match:
        return f"时间：{raw}", raw
    year, month, day = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
    try:
        weekday_index = dt.date(year, month, day).weekday()
    except ValueError:
        return f"时间：{raw}", raw
    weekday_match = re.search(r"星期([一二三四五六日天])", raw)
    if weekday_match:
        weekday = "日" if weekday_match.group(1) == "天" else weekday_match.group(1)
    else:
        weekday = WEEKDAYS[weekday_index]
    cover = f"时间：{year}年{month}月{day}日 星期{weekday}"
    foot = f"{year}年{month}月{day}日"
    return cover, foot


def as_item(raw) -> dict:
    if isinstance(raw, str):
        return {"text": raw, "children": [], "images": [], "code": None, "table": None}
    if not isinstance(raw, dict):
        raise ReportError("Each point must be a string or an object")
    children = []
    for child in raw.get("children") or []:
        if isinstance(child, str):
            children.append(child)
        elif isinstance(child, dict):
            if child.get("images"):
                raise ReportError("Do not attach images to sub-points. Attach them to the parent point.")
            children.append(child.get("text") or "")
        else:
            raise ReportError("A sub-point has the wrong format")
    images = []
    for image in raw.get("images") or []:
        if isinstance(image, str):
            images.append({"path": image, "caption": ""})
        elif isinstance(image, dict):
            images.append(
                {
                    "path": image.get("path") or "",
                    "caption": image.get("caption") or "",
                    "lead": image.get("lead") or "",
                }
            )
        else:
            raise ReportError("An image entry has the wrong format")
    code = raw.get("code")
    if isinstance(code, bool) and code:
        code_text = raw.get("text") or ""
    elif isinstance(code, str):
        code_text = code
    else:
        code_text = None
    table = raw.get("table")
    if table is not None and not isinstance(table, dict):
        raise ReportError("table must be an object with header and rows")
    return {
        "text": raw.get("text") or "",
        "children": children,
        "images": images,
        "code": code_text,
        "table": table,
    }


def require_items(spec: dict, key: str) -> list[dict]:
    raw = spec.get(key)
    if not isinstance(raw, list) or not raw:
        raise ReportError(f"Missing section content: {key}")
    return [as_item(item) for item in raw]


def optional_items(spec: dict, key: str) -> list[dict]:
    raw = spec.get(key) or []
    if not isinstance(raw, list):
        raise ReportError(f"{key} must be an array")
    return [as_item(item) for item in raw]


def normalize_review(text: str) -> str:
    review = re.sub(r"[\r\n]+", "", (text or "")).strip()
    if not review:
        raise ReportError("Missing teacher comment")
    width = display_width(review)
    if width > MAX_REVIEW_WIDTH:
        raise ReportError(
            f"The teacher comment is longer than two lines (width {width:.0f}, limit {MAX_REVIEW_WIDTH}). Shorten it and run again.",
            code=2,
        )
    return review


def fitted_width(path: Path) -> int:
    image = DocxImage.from_file(str(path))
    scale = min(MAX_IMAGE_WIDTH_EMU / image.width, MAX_IMAGE_HEIGHT_EMU / image.height)
    return max(1, int(image.width * scale))


def add_image(cell, path: Path, caption: str, figs: Figs) -> None:
    if not path.is_file():
        raise ReportError(f"Image not found: {path}")
    if not caption.strip():
        raise ReportError(f"Image is missing a caption: {path}")
    paragraph = cell.add_paragraph()
    apply_paragraph_format(paragraph, indent_chars=0, center=True, exact=False, before=120, after=0)
    run = paragraph.add_run()
    run.add_picture(str(path), width=fitted_width(path))
    caption_paragraph = cell.add_paragraph()
    apply_paragraph_format(
        caption_paragraph, indent_chars=0, center=True, exact=True, before=40, after=80
    )
    caption_run = caption_paragraph.add_run(f"图{figs.n} {caption.strip()}")
    set_run_font(caption_run)
    figs.n += 1


def set_code_font(run, size_pt: float = 10.5) -> None:
    run.bold = False
    run.italic = False
    run.font.size = Pt(size_pt)
    run.font.name = "Times New Roman"
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), "宋体")
    r_fonts.set(qn("w:cs"), "Times New Roman")
    for tag in ("sz", "szCs"):
        node = r_pr.find(qn(f"w:{tag}"))
        if node is None:
            node = OxmlElement(f"w:{tag}")
            r_pr.append(node)
        node.set(qn("w:val"), str(int(size_pt * 2)))


def add_code_block(cell, code_text: str) -> None:
    """Render code lines as bordered, shaded paragraphs like the lab guide."""
    lines = (code_text or "").strip("\n").splitlines()
    if not lines:
        return
    for line in lines:
        paragraph = cell.add_paragraph()
        p_pr = paragraph._p.get_or_add_pPr()
        p_bdr = OxmlElement("w:pBdr")
        for edge, color, width in (
            ("top", "BFD3E6", "4"),
            ("left", "5B9BD5", "12"),
            ("bottom", "BFD3E6", "4"),
            ("right", "BFD3E6", "4"),
        ):
            node = OxmlElement(f"w:{edge}")
            node.set(qn("w:val"), "single")
            node.set(qn("w:color"), color)
            node.set(qn("w:sz"), width)
            node.set(qn("w:space"), "4")
            p_bdr.append(node)
        p_pr.append(p_bdr)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F5F8FC")
        p_pr.append(shd)
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:before"), "20")
        spacing.set(qn("w:after"), "20")
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")
        p_pr.append(spacing)
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "360")
        ind.set(qn("w:leftChars"), "0")
        ind.set(qn("w:right"), "120")
        ind.set(qn("w:rightChars"), "0")
        ind.set(qn("w:firstLine"), "0")
        ind.set(qn("w:firstLineChars"), "0")
        p_pr.append(ind)
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "left")
        p_pr.append(jc)
        run = paragraph.add_run(line if line.strip() else " ")
        set_code_font(run)


def add_table(cell, table_data: dict) -> None:
    header = table_data.get("header") or []
    rows = table_data.get("rows") or []
    if not header and not rows:
        raise ReportError("A table entry is empty")
    n_cols = len(header) if header else max(len(r) for r in rows)
    n_rows = len(rows) + (1 if header else 0)
    table = cell.add_table(rows=n_rows, cols=n_cols)
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "000000")
        borders.append(node)
    tbl_pr.append(borders)
    data = [header] + rows if header else rows
    for ri, row in enumerate(data):
        for ci in range(n_cols):
            value = row[ci] if ci < len(row) else ""
            table_cell = table.cell(ri, ci)
            table_cell.text = str(value)
            for paragraph in table_cell.paragraphs:
                for run in paragraph.runs:
                    set_code_font(run)
                p_pr = paragraph._p.get_or_add_pPr()
                jc = OxmlElement("w:jc")
                jc.set(qn("w:val"), "center")
                p_pr.append(jc)
                spacing = OxmlElement("w:spacing")
                spacing.set(qn("w:line"), "240")
                spacing.set(qn("w:lineRule"), "auto")
                spacing.set(qn("w:before"), "0")
                spacing.set(qn("w:after"), "0")
                p_pr.append(spacing)


def add_item(cell, item: dict, number: int, figs: Figs, allow_images: bool) -> None:
    text = strip_number(item["text"])
    if not text:
        raise ReportError("A point is empty")
    paragraph = cell.add_paragraph()
    apply_paragraph_format(paragraph, indent_chars=2, exact=True)
    run = paragraph.add_run(f"{number}. {text}")
    set_run_font(run)
    for index, child in enumerate(item["children"], start=1):
        child_text = strip_number(child)
        if not child_text:
            raise ReportError("A sub-point is empty")
        child_paragraph = cell.add_paragraph()
        apply_paragraph_format(child_paragraph, indent_chars=2, exact=True)
        child_run = child_paragraph.add_run(f"（{index}）{child_text}")
        set_run_font(child_run)
    if item.get("code"):
        add_code_block(cell, item["code"])
    if item.get("table"):
        add_table(cell, item["table"])
    if not item["images"]:
        return
    if not allow_images:
        raise ReportError("Images can only be placed in the steps section or the results section")
    # A caption under a picture is not the sentence between pictures. Two pictures
    # in a row with nothing but captions between them read as a screenshot dump.
    if len(item["images"]) >= 2:
        for image in item["images"]:
            if not (image.get("lead") or "").strip():
                raise ReportError(
                    "Pictures are stacked without text. Add a lead sentence before each picture in that point."
                )
    for image in item["images"]:
        lead = strip_number(image.get("lead") or "")
        if lead:
            lead_paragraph = cell.add_paragraph()
            apply_paragraph_format(lead_paragraph, indent_chars=2, exact=True)
            lead_run = lead_paragraph.add_run(lead)
            set_run_font(lead_run)
        add_image(cell, Path(image["path"]), image["caption"], figs)


def clear_after_heading(cell, keep: int) -> None:
    paragraphs = list(cell.paragraphs)
    if len(paragraphs) < keep:
        raise ReportError("A template section is missing its heading")
    for paragraph in paragraphs[keep:]:
        if paragraph_has_drawing(paragraph):
            raise ReportError("A template section already contains a picture, so filling stopped")
        delete_paragraph(paragraph)


def fill_section(cell, items: list[dict], figs: Figs, allow_images: bool) -> None:
    clear_after_heading(cell, 1)
    for index, item in enumerate(items, start=1):
        add_item(cell, item, index, figs, allow_images)


def add_subheading(cell, text: str, before: int = 120) -> None:
    paragraph = cell.add_paragraph()
    apply_paragraph_format(paragraph, indent_chars=0, exact=True, before=before, after=0)
    run = paragraph.add_run(text)
    set_run_font(run)


def fill_results(cell, items: list[dict], discussion: list[dict], figs: Figs) -> None:
    paragraphs = list(cell.paragraphs)
    has_subheading = len(paragraphs) >= 2 and "实验结果与分析" in paragraphs[1].text
    clear_after_heading(cell, 2 if has_subheading else 1)
    if not has_subheading:
        add_subheading(cell, "实验结果与分析：", before=0)
    for index, item in enumerate(items, start=1):
        add_item(cell, item, index, figs, allow_images=True)
    if not discussion:
        return
    add_subheading(cell, "问题讨论：")
    for index, item in enumerate(discussion, start=1):
        add_item(cell, item, index, figs, allow_images=False)


def fill_cover(doc, meta: dict) -> None:
    values = {
        "course": (meta.get("course") or "").strip(),
        "lab": (meta.get("lab") or "").strip(),
        "experiment": (meta.get("experiment") or "").strip(),
        "student": (meta.get("student") or "").strip(),
        "time": (meta.get("time") or "").strip(),
        "teacher": (meta.get("teacher") or "").strip(),
        "student_id": (meta.get("student_id") or "").strip(),
        "major": (meta.get("major") or "").strip(),
        "class_name": (meta.get("class_name") or "").strip(),
        "college": (meta.get("college") or "").strip(),
    }
    found = set()
    for cell in iter_cells(doc):
        label = re.sub(r"\s+", "", cell_plain_text(cell))
        if any(
            key in label
            for key in ("目的与要求", "主要材料", "方法与步骤", "结果与问题", "教师评阅", "教师签名")
        ):
            continue
        paragraph = cell.paragraphs[0]
        if label.startswith("课程名称"):
            found.add("course")
            if values["course"]:
                append_value(paragraph, values["course"])
        elif label.startswith("实验（训）室名称") or label.startswith("实验(训)室名称"):
            found.add("lab")
            if values["lab"]:
                append_value(paragraph, values["lab"])
        elif label.startswith("实验（训）名称") or label.startswith("实验(训)名称"):
            found.add("experiment")
            if values["experiment"]:
                append_value(paragraph, values["experiment"])
        elif label.startswith("学生姓名"):
            found.add("student")
            if values["student"]:
                append_value(paragraph, values["student"])
        elif label.startswith("学号"):
            found.add("student_id")
            if values["student_id"]:
                append_value(paragraph, values["student_id"])
        elif label.startswith("专业"):
            found.add("major")
            if values["major"]:
                rewrite_label_value(paragraph, values["major"])
        elif label.startswith("班级"):
            found.add("class_name")
            if values["class_name"]:
                rewrite_label_value(paragraph, values["class_name"])
        elif label.startswith("院（部）") or label.startswith("院(部)"):
            found.add("college")
            if values["college"]:
                rewrite_label_value(paragraph, values["college"])
        elif label.startswith("时间"):
            found.add("time")
            if values["time"]:
                cover, _ = format_time(values["time"])
                replace_runs(paragraph, cover)
        elif label.startswith("指导教师"):
            found.add("teacher")
            if values["teacher"]:
                append_value(paragraph, values["teacher"])
    missing = [name for name, value in values.items() if value and name not in found]
    if missing:
        raise ReportError("Template is missing cover fields: " + ", ".join(missing))


def fill_review(cell, review: str, time_raw: str) -> None:
    """Write the comment under the teacher-comment label and leave the stamp box alone."""
    paragraphs = list(cell.paragraphs)
    label_index = next(
        (index for index, paragraph in enumerate(paragraphs) if "教师评阅" in paragraph.text),
        None,
    )
    if label_index is None:
        raise ReportError("Template is missing the teacher-comment cell")
    target = None
    for paragraph in paragraphs[label_index + 1 :]:
        if paragraph_has_drawing(paragraph):
            continue
        text = paragraph.text.strip()
        if text.startswith("教师签名") or re.sub(r"\s+", "", text) == "年月日":
            break
        target = paragraph
        break
    if target is None:
        label = paragraphs[label_index]
        target = cell.add_paragraph()
        label._p.addnext(target._p)
    apply_paragraph_format(target, indent_chars=2, exact=True)
    replace_runs(target, review)

    if not time_raw.strip():
        return
    _, foot = format_time(time_raw)
    for paragraph in cell.paragraphs:
        if re.sub(r"\s+", "", paragraph.text) == "年月日":
            replace_runs(paragraph, foot)
            return
    raise ReportError("Template is missing the date line at the end")


def check_image_path() -> Path:
    path = Path(__file__).resolve().parent.parent / "assets" / "red-check.png"
    if not path.is_file():
        raise ReportError("assets/red-check.png is missing, so the red check cannot be stamped.")
    return path


def remove_template_check(doc) -> None:
    """Remove only check-sized drawings/ink from the body and all headers.

    A reused template may already contain a repeating check or body copies.
    Leave unrelated header content, score boxes, signatures and images alone.
    """
    alternate = "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
    roots = [doc.element.body]
    for part in doc.part.package.parts:
        if part.content_type.endswith("wordprocessingml.header+xml"):
            roots.append(part.element)
    for root in roots:
        for extent in list(root.iter(qn("wp:extent"))):
            drawing = next((node for node in extent.iterancestors() if node.tag == qn("w:drawing")), None)
            if drawing is None:
                continue
            tagged = any(node.get("name", "").startswith(CHECK_NAME_PREFIX)
                         for node in drawing.iter(qn("wp:docPr")))
            legacy = (extent.get("cx") == str(LEGACY_CHECK_WIDTH_EMU)
                      and extent.get("cy") == str(LEGACY_CHECK_HEIGHT_EMU))
            if not tagged and not legacy:
                continue
            target = next((node for node in drawing.iterancestors() if node.tag == alternate), drawing)
            parent = target.getparent()
            if parent is not None:
                parent.remove(target)
                if parent.tag == qn("w:r") and len(parent) == 0:
                    parent.getparent().remove(parent)


def _check_geometry(setup: dict, page: int) -> dict:
    """Fit a proportional check inside this physical page's writable area.

    Word/WPS supplies section dimensions after layout, including mixed paper
    sizes. Keep a visible check without covering nearly the entire sheet.
    Returned values and setup lengths are EMU, not pixels or points.
    """
    try:
        values = {key: int(setup[key])
                  for key in ("width", "height", "left", "right", "top", "bottom")}
        gutter = max(0, int(setup.get("gutter", 0)))
        gutter_pos = int(setup.get("gutter_pos", 0))
        left, right = values["left"], values["right"]
        top, bottom = values["top"], values["bottom"]
        if min(left, right, top, bottom) < 0:
            raise ValueError("negative page margin")
        if setup.get("mirror") and page % 2 == 0:
            left, right = right, left
        if gutter_pos == 1:
            top += gutter
        elif gutter_pos == 2 or (setup.get("mirror") and page % 2 == 0):
            right += gutter
        else:
            left += gutter
        usable_width = values["width"] - left - right
        usable_height = values["height"] - top - bottom
        if min(usable_width, usable_height) <= 0:
            raise ValueError("empty writable area")
    except (KeyError, TypeError, ValueError, OverflowError) as err:
        raise ReportError("Word/WPS returned invalid page dimensions for the red check.") from err
    picture = DocxImage.from_file(str(check_image_path()))
    scale = min(usable_width * 0.60 / picture.px_width,
                usable_height * 0.55 / picture.px_height)
    width = max(1, round(picture.px_width * scale))
    height = max(1, round(picture.px_height * scale))
    return {"width": width, "height": height,
            "left": left + (usable_width - width) // 2,
            "top": top + (usable_height - height) // 2}


def _inline_to_page_anchor(inline, left: int, top: int, *, layout_in_cell: str = "1", z: str = "251661312"):
    anchor = OxmlElement("wp:anchor")
    for attr, value in (
        ("distT", "0"),
        ("distB", "0"),
        ("distL", "0"),
        ("distR", "0"),
        ("simplePos", "0"),
        ("relativeHeight", z),
        ("behindDoc", "0"),
        ("locked", "0"),
        ("layoutInCell", layout_in_cell),
        ("allowOverlap", "1"),
    ):
        anchor.set(attr, value)
    simple = OxmlElement("wp:simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    anchor.append(simple)
    for axis, origin, offset in (
        ("wp:positionH", "page", left),
        ("wp:positionV", "page", top),
    ):
        node = OxmlElement(axis)
        node.set("relativeFrom", origin)
        pos = OxmlElement("wp:posOffset")
        pos.text = str(offset)
        node.append(pos)
        anchor.append(node)
    for tag in ("wp:extent", "wp:effectExtent"):
        child = inline.find(qn(tag))
        if child is not None:
            anchor.append(child)
    anchor.append(OxmlElement("wp:wrapNone"))
    for tag in ("wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"):
        child = inline.find(qn(tag))
        if child is not None:
            anchor.append(child)
    inline.getparent().replace(inline, anchor)


def _pagination(docx_path: Path, prefix: str, temp_dir: Path) -> dict:
    """Read actual Word/WPS pagination; OOXML itself has no physical-page model."""
    helper = Path(__file__).with_name("paginate_docx.ps1")
    if not helper.is_file():
        raise ReportError("scripts/paginate_docx.ps1 is missing.")
    result_path = temp_dir / "pagination.json"
    result_path.unlink(missing_ok=True)
    powershell = shutil.which("pwsh.exe") or shutil.which("powershell.exe")
    if not powershell:
        raise ReportError("PowerShell and Microsoft Word or WPS Writer are required to exclude the final page.")
    try:
        result = subprocess.run(
            [powershell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(helper),
             "-DocumentPath", str(docx_path.resolve()), "-ResultPath", str(result_path.resolve()),
             "-MarkerPrefix", prefix],
            capture_output=True, timeout=120,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.TimeoutExpired) as err:
        raise ReportError("Could not determine the last page in Word/WPS. No finished report was written.") from err
    if result.returncode or not result_path.is_file():
        detail = (result.stderr or result.stdout).decode(errors="replace").strip()
        raise ReportError("Word/WPS pagination failed. No finished report was written. " + detail)
    try:
        data = json.loads(result_path.read_text(encoding="utf-8-sig"))
        pages = int(data["pages"])
        positions = data["positions"]
        if pages < 1 or not isinstance(positions, dict) or not positions:
            raise ValueError("invalid pagination")
        if any(not isinstance(page, int) or not 1 <= page <= pages for page in positions.values()):
            raise ValueError("invalid physical page")
        page_setup = data["page_setup"]
        if not isinstance(page_setup, dict) or not all(str(page) in page_setup for page in positions.values()):
            raise ValueError("missing page dimensions")
    except (OSError, ValueError, KeyError, TypeError) as err:
        raise ReportError("Word/WPS returned incomplete pagination. No finished report was written.") from err
    return data


def stamp_check_except_last_page(doc) -> None:
    """Stamp one body check on every physical page except the final page.

    Temporary zero-length bookmarks identify existing main-story paragraphs.
    They do not change the layout and are removed before the final save.
    A second pagination pass verifies all check anchors after stamping.
    """
    from docx.text.paragraph import Paragraph

    remove_template_check(doc)
    prefix = "JRCheck_" + uuid.uuid4().hex[:8] + "_"
    bookmark_ids = [int(node.get(qn("w:id"))) for node in doc.element.iter(qn("w:bookmarkStart"))
                    if (node.get(qn("w:id")) or "").isdigit()]
    next_id = max(bookmark_ids, default=-1) + 1
    markers = []
    for paragraph in doc.element.body.iter(qn("w:p")):
        if any(node.tag == qn("w:txbxContent") for node in paragraph.iterancestors()):
            continue
        name = prefix + str(len(markers))
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(next_id))
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(next_id))
        next_id += 1
        offset = 1 if len(paragraph) and paragraph[0].tag == qn("w:pPr") else 0
        paragraph.insert(offset, start)
        paragraph.insert(offset + 1, end)
        markers.append((name, paragraph, start, end))
    if not markers:
        raise ReportError("No body paragraphs were available for page checks.")
    try:
        with tempfile.TemporaryDirectory(prefix="jr-training-pages-") as folder:
            temp_dir = Path(folder)
            staged = temp_dir / "report.docx"
            doc.save(str(staged))
            before = _pagination(staged, prefix, temp_dir)
            if set(before["positions"]) != {name for name, _, _, _ in markers}:
                raise ReportError("Word/WPS could not locate every paragraph. No finished report was written.")
            chosen = {}
            for name, paragraph, _, _ in markers:
                page = before["positions"][name]
                if page == before["pages"]:
                    continue
                # Prefer a picture paragraph so the overlay paints over screenshots.
                if page not in chosen or any(node.tag == qn("w:drawing") for node in paragraph.iter()):
                    chosen[page] = (name, paragraph)
            if set(chosen) != set(range(1, before["pages"])):
                raise ReportError("A page has no safe paragraph anchor. No finished report was written.")
            for page, (name, paragraph_elm) in chosen.items():
                geometry = _check_geometry(before["page_setup"][str(page)], page)
                paragraph = Paragraph(paragraph_elm, doc._body)
                run = paragraph.add_run()
                run.font.size = Pt(1)
                run.add_picture(str(check_image_path()), width=geometry["width"], height=geometry["height"])
                inline = run._r.find(qn("w:drawing")).find(qn("wp:inline"))
                if inline is None:
                    raise ReportError("Could not place the body red check.")
                inline.find(qn("wp:docPr")).set("name", CHECK_NAME_PREFIX + str(page))
                _inline_to_page_anchor(inline, geometry["left"], geometry["top"],
                                       layout_in_cell="0", z="251665000")
            doc.save(str(staged))
            after = _pagination(staged, prefix, temp_dir)
            if (before["pages"] != after["pages"] or before["positions"] != after["positions"]
                    or before["page_setup"] != after["page_setup"]):
                raise ReportError("Pagination changed while placing checks. No finished report was written.")
            expected_checks = {"JRTrainingCheck_Page_" + str(page): page for page in chosen}
            if after.get("checks") != expected_checks:
                raise ReportError("Word/WPS could not confirm the check pages. No finished report was written.")
    finally:
        for _, _, start, end in markers:
            for node in (start, end):
                if node.getparent() is not None:
                    node.getparent().remove(node)


def build(spec: dict, desktop: Path | None = None) -> Path:
    if not isinstance(spec, dict):
        raise ReportError("The spec file must be a JSON object")
    meta = spec.get("meta") or {}
    if not isinstance(meta, dict):
        raise ReportError("meta must be an object")

    template = find_template()
    desktop = Path(desktop) if desktop is not None else find_desktop()
    desktop.mkdir(parents=True, exist_ok=True)
    stem = sanitize_stem((meta.get("experiment") or "").strip()) or "实训报告"
    output = unique_docx(desktop, stem)
    if output.resolve() == template.resolve():
        raise ReportError("Refusing to overwrite the blank template")

    purpose = require_items(spec, "purpose")
    materials = require_items(spec, "materials")
    steps = require_items(spec, "steps")
    results = require_items(spec, "results")
    discussion = optional_items(spec, "discussion")
    review = normalize_review(spec.get("review") or "")

    fd, temp_name = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    temp_path = Path(temp_name)
    shutil.copyfile(template, temp_path)
    try:
        doc = Document(str(temp_path))
        experiment = (meta.get("experiment") or "").strip()
        if experiment:
            doc.core_properties.title = experiment
        fill_cover(doc, meta)
        figs = Figs()
        fill_section(find_cell(doc, "目的与要求"), purpose, figs, allow_images=False)
        fill_section(find_cell(doc, "主要材料与仪器设备"), materials, figs, allow_images=False)
        fill_section(find_cell(doc, "内容、方法与步骤"), steps, figs, allow_images=True)
        fill_results(find_cell(doc, "结果与问题讨论"), results, discussion, figs)
        fill_review(find_cell(doc, "教师评阅"), review, (meta.get("time") or ""))
        normalize_report_layout(doc, [find_cell(doc, key) for key in (
            "目的与要求", "主要材料与仪器设备", "内容、方法与步骤", "结果与问题讨论"
        )])
        stamp_check_except_last_page(doc)
        doc.save(str(output))
    finally:
        temp_path.unlink(missing_ok=True)
    return output


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    args = list(sys.argv[1:] if argv is None else argv)
    if len(args) != 1:
        print("Usage: python fill_report.py <spec.json>", file=sys.stderr)
        return 1
    spec_path = Path(args[0])
    try:
        spec = json.loads(spec_path.read_text(encoding="utf-8-sig"))
        template = find_template()
        output = build(spec)
    except ReportError as err:
        print(f"ERROR: {err}", file=sys.stderr)
        return err.code
    except json.JSONDecodeError as err:
        print(f"ERROR: The spec file is not valid JSON: {err}", file=sys.stderr)
        return 1
    except OSError as err:
        print(f"ERROR: Could not read or write a file: {err}", file=sys.stderr)
        return 1
    print(f"TEMPLATE: {template}")
    print(f"WROTE: {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
