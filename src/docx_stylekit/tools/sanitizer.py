from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Optional

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

from .image_paragraphs import fix_image_paragraph_spacing


MANDATORY_STYLES = ["ImageParagraph", "PageNumber", "InfoTable"]


def _copy_docx(src: Path) -> Path:
    tmp_dir = Path(tempfile.mkdtemp())
    dst = tmp_dir / src.name
    shutil.copy2(src, dst)
    return dst


def _replace_part(docx_path: Path, part_name: str, content: bytes):
    tmp_zip = docx_path.with_suffix(".tmp")
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
        replaced = False
        for item in zin.infolist():
            if item.filename == part_name:
                zout.writestr(item, content)
                replaced = True
            else:
                zout.writestr(item, zin.read(item.filename))
        if not replaced:
            zout.writestr(part_name, content)
    docx_path.unlink()
    tmp_zip.rename(docx_path)


def _extract_part(docx_path: Path, part_name: str) -> Optional[bytes]:
    with zipfile.ZipFile(docx_path) as z:
        if part_name in z.namelist():
            return z.read(part_name)
    return None


def _ensure_paragraph_style(
    doc: Document,
    name: str,
    *,
    base_on: str = "Normal",
    font_name: str = "Times New Roman",
    size_pt: float = 16,
    bold: bool = False,
    italic: bool = False,
    align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER,
    line_spacing: float = 1.0,
) -> None:
    try:
        style = doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        try:
            style.base_style = doc.styles[base_on]
        except KeyError:
            pass

    style.font.name = font_name
    style.font.size = Pt(size_pt)
    style.font.bold = bold
    style.font.italic = italic
    pf = style.paragraph_format
    pf.alignment = align
    pf.line_spacing = line_spacing
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Pt(0)


def _ensure_table_style(doc: Document, name: str):
    try:
        doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, WD_STYLE_TYPE.TABLE)
        base_style = None
        for candidate in ("Table Grid", "Normal Table"):
            try:
                base_style = doc.styles[candidate]
                break
            except KeyError:
                continue
        if base_style:
            style.base_style = base_style


def _ensure_required_styles(doc: Document):
    _ensure_paragraph_style(
        doc,
        "Title",
        base_on="Normal",
        font_name="Times New Roman",
        size_pt=18,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _ensure_paragraph_style(
        doc,
        "ImageParagraph",
        base_on="Normal",
        font_name="Times New Roman",
        size_pt=16,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _ensure_paragraph_style(
        doc,
        "PageNumber",
        base_on="Normal",
        font_name="Times New Roman",
        size_pt=14,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _ensure_table_style(doc, "InfoTable")


def _detect_heading_level(text: str) -> Optional[int]:
    if not text:
        return None
    stripped = text.strip()
    # Numeric patterns like 1., 1.1, 1.1.1
    numeric_match = re.match(r"^(\d+(?:\.\d+)*)[\.\s]", stripped)
    if numeric_match:
        segments = numeric_match.group(1).split('.')
        return len(segments)
    # Chinese numerals
    if re.match(r"^第[一二三四五六七八九十百千]+章", stripped):
        return 1
    if re.match(r"^[一二三四五六七八九十]+、", stripped):
        return 1
    if re.match(r"^（[一二三四五六七八九十]+）", stripped) or re.match(r"^\([一二三四五六七八九十]+\)", stripped):
        return 2
    if re.match(r"^（\d+）", stripped) or re.match(r"^\(\d+\)", stripped):
        return 2
    return None


TITLE_KEYWORDS = [
    "说明书",
    "方案",
    "规划",
    "报告",
    "手册",
    "计划",
    "分析",
    "项目",
]


def _extract_orig_heading_level(style_name: str) -> Optional[int]:
    if not style_name:
        return None
    m = re.match(r"Heading\s*(\d+)", style_name, re.IGNORECASE)
    if m:
        return int(m.group(1))
    m = re.match(r"标题\s*(\d+)", style_name)
    if m:
        return int(m.group(1))
    return None


def _map_style_name(
    raw_name: str,
    text: str,
    non_empty_index: int,
    previous_level: Optional[int],
) -> str:
    raw_name = raw_name or ""
    normalized = raw_name.lower()
    stripped = text.strip()
    original_level = _extract_orig_heading_level(raw_name)
    pattern_level = _detect_heading_level(stripped)

    if original_level and not pattern_level:
        return f"Heading {original_level}"

    if pattern_level:
        if original_level:
            if previous_level is not None:
                target = previous_level + 1
                diff_pattern = abs(pattern_level - target)
                diff_original = abs(original_level - target)
                level = pattern_level if diff_pattern <= diff_original else original_level
            else:
                level = pattern_level if pattern_level <= original_level else original_level
        else:
            level = pattern_level
        return f"Heading {max(1, min(level, 9))}"

    if original_level:
        return f"Heading {original_level}"

    if non_empty_index == 0 and stripped:
        if len(stripped) <= 80 and any(keyword in stripped for keyword in TITLE_KEYWORDS):
            return "Title"
        if len(stripped) <= 40:
            return "Title"
    if stripped and len(stripped) <= 25 and any(stripped.endswith(k) for k in TITLE_KEYWORDS):
        return "Title"
    return "Normal"


def _clear_run_formatting(paragraph):
    for run in paragraph.runs:
        rpr = run._element.find(qn("w:rPr"))
        if rpr is not None:
            run._element.remove(rpr)


def _clear_paragraph_formatting(paragraph):
    ppr = paragraph._element.find(qn("w:pPr"))
    if ppr is not None:
        for child in list(ppr):
            if child.tag in {
                qn("w:ind"),
                qn("w:spacing"),
                qn("w:jc"),
                qn("w:keepNext"),
                qn("w:keepLines"),
                qn("w:pageBreakBefore"),
            }:
                ppr.remove(child)


def _iter_paragraphs(doc: Document):
    for paragraph in doc.paragraphs:
        yield paragraph
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph


def sanitize_docx(
    raw_docx: Path,
    template_docx: Optional[Path] = None,
    *,
    output_path: Optional[Path] = None,
) -> Path:
    raw_docx = Path(raw_docx)
    template_docx = Path(template_docx) if template_docx else None

    working_copy = _copy_docx(raw_docx)

    if template_docx and template_docx.exists():
        styles_xml = _extract_part(template_docx, "word/styles.xml")
        if styles_xml:
            _replace_part(working_copy, "word/styles.xml", styles_xml)
        numbering_xml = _extract_part(template_docx, "word/numbering.xml")
        if numbering_xml:
            _replace_part(working_copy, "word/numbering.xml", numbering_xml)

    doc = Document(str(working_copy))
    original_doc = Document(str(raw_docx))
    _ensure_required_styles(doc)

    paragraphs = list(_iter_paragraphs(doc))
    original_paragraphs = list(_iter_paragraphs(original_doc))
    original_styles = [p.style.name if p.style else "" for p in original_paragraphs]

    non_empty_counter = 0
    previous_heading_level: Optional[int] = None
    for idx, paragraph in enumerate(paragraphs):
        original_style = original_styles[idx] if idx < len(original_styles) else ""
        style_name = _map_style_name(original_style, paragraph.text, non_empty_counter, previous_heading_level)
        if paragraph._element.xpath(".//w:drawing"):
            paragraph.style = doc.styles["ImageParagraph"]
        else:
            try:
                paragraph.style = doc.styles[style_name]
            except KeyError:
                paragraph.style = doc.styles.get(style_name) or doc.styles["Normal"]
        _clear_run_formatting(paragraph)
        _clear_paragraph_formatting(paragraph)
        if paragraph.text.strip():
            if style_name.startswith("Heading "):
                try:
                    previous_heading_level = int(style_name.split()[1])
                except ValueError:
                    previous_heading_level = None
            elif style_name == "Title":
                previous_heading_level = 0
            non_empty_counter += 1

    try:
        doc.styles["InfoTable"]
        for table in doc.tables:
            try:
                table.style = doc.styles["InfoTable"]
            except KeyError:
                continue
    except KeyError:
        pass

    doc.save(str(working_copy))
    fix_image_paragraph_spacing(working_copy, working_copy)

    destination = Path(output_path) if output_path else raw_docx
    shutil.copy2(working_copy, destination)
    return destination


def qn(tag: str) -> str:
    from docx.oxml.ns import qn as _qn

    return _qn(tag)
