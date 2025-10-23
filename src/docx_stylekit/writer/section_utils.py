# src/docx_stylekit/writer/section_utils.py
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

def apply_section_layout(section, layout: dict):
    if not layout:
        return
    m = layout.get("marginsCm") or {}
    if "top" in m and m["top"] is not None:
        section.top_margin = Cm(m["top"])
    if "bottom" in m and m["bottom"] is not None:
        section.bottom_margin = Cm(m["bottom"])
    if "left" in m and m["left"] is not None:
        section.left_margin = Cm(m["left"])
    if "right" in m and m["right"] is not None:
        section.right_margin = Cm(m["right"])
    if "header" in m and m["header"] is not None:
        section.header_distance = Cm(m["header"])
    if "footer" in m and m["footer"] is not None:
        section.footer_distance = Cm(m["footer"])

    # 方向
    if layout.get("orientation") == "landscape":
        section.orientation = 1  # WD_ORIENT.LANDSCAPE（避免导入枚举）
    # 首页不同/奇偶页不同
    if layout.get("titleFirstPageDifferent") is not None:
        section.different_first_page_header_footer = bool(layout["titleFirstPageDifferent"])
    if layout.get("evenOddDifferent") is not None:
        # python-docx 无直接属性；需要 settings.xml 支持。此处略过或留扩展点。
        pass

    # 页码起始/格式：由页眉页脚域与 settings/sectPr 配合，这里仅设置 startAt（可选）
    if layout.get("startAt") is not None:
        pg_num_type = section._sectPr.get_or_add_pgNumType()
        pg_num_type.set(qn('w:start'), str(int(layout["startAt"])))

    # 垂直对齐（vAlign）
    if layout.get("verticalAlign"):
        v = section._sectPr.find(qn('w:vAlign'))
        if v is None:
            v = OxmlElement('w:vAlign')
            section._sectPr.append(v)
        v.set(qn('w:val'), layout["verticalAlign"])

def add_page_number_field(paragraph, align="center"):
    # 设置段落对齐
    align_map = {"left":0, "center":1, "right":2}
    paragraph.alignment = align_map.get(align, 1)
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = ""
    r.append(t)
    fld.append(r)
    paragraph._p.append(fld)

def add_toc_field(paragraph, levels=(1,3)):
    start = min(levels) if isinstance(levels, (list,tuple)) and levels else 1
    end = max(levels) if isinstance(levels, (list,tuple)) and levels else 3
    instr = f'TOC \\o "{start}-{end}" \\h \\z \\u'
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), instr)
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = "目录（打开后更新）"
    r.append(t)
    fld.append(r)
    paragraph._p.append(fld)
