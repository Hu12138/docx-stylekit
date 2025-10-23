# src/docx_stylekit/writer/style_store.py
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def _set_rfonts(rPr, eastAsia=None, ascii_=None):
    rfonts = rPr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rPr.append(rfonts)
    if eastAsia:
        rfonts.set(qn('w:eastAsia'), eastAsia)
    if ascii_:
        rfonts.set(qn('w:ascii'), ascii_)
        rfonts.set(qn('w:hAnsi'), ascii_)

def _ensure_child(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el

def _pt_to_halfpt(v):
    # python-docx sz uses half-points
    return int(round(float(v) * 2))

def _pt_to_twentieth(v):
    # OOXML spacing (before/after/line exact/atLeast) uses twentieth of a point
    return int(round(float(v) * 20))

def apply_font_rpr(rPr, font: dict):
    if font is None:
        return
    _set_rfonts(rPr, font.get("eastAsia"), font.get("ascii"))
    if "sizePt" in font and font["sizePt"]:
        sz = _ensure_child(rPr, 'w:sz')
        sz.set(qn('w:val'), str(_pt_to_halfpt(font["sizePt"])))
    if "bold" in font:
        if font["bold"]:
            _ensure_child(rPr, 'w:b')
        else:
            b = rPr.find(qn('w:b'))
            if b is not None:
                rPr.remove(b)
    if "italic" in font:
        if font["italic"]:
            _ensure_child(rPr, 'w:i')
        else:
            i = rPr.find(qn('w:i'))
            if i is not None:
                rPr.remove(i)
    if "color" in font and font["color"]:
        color = _ensure_child(rPr, 'w:color')
        color.set(qn('w:val'), font["color"].lstrip('#').upper())

def apply_paragraph_ppr(pPr, para: dict):
    if para is None:
        return
    if "align" in para and para["align"]:
        jc = _ensure_child(pPr, 'w:jc')
        jc.set(qn('w:val'), para["align"])
    # 行距
    if para.get("lineSpacingMultiple"):
        spacing = _ensure_child(pPr, 'w:spacing')
        # line = 倍数 * 240 （OOXML 规则）
        spacing.set(qn('w:line'), str(int(para["lineSpacingMultiple"] * 240)))
        spacing.set(qn('w:lineRule'), 'auto')
    if para.get("lineExactPt"):
        spacing = _ensure_child(pPr, 'w:spacing')
        spacing.set(qn('w:line'), str(_pt_to_twentieth(para["lineExactPt"])))
        spacing.set(qn('w:lineRule'), 'exact')
    if para.get("lineAtLeastPt"):
        spacing = _ensure_child(pPr, 'w:spacing')
        spacing.set(qn('w:line'), str(_pt_to_twentieth(para["lineAtLeastPt"])))
        spacing.set(qn('w:lineRule'), 'atLeast')
    # 段前后
    if para.get("spaceBeforePt") is not None:
        spacing = _ensure_child(pPr, 'w:spacing')
        spacing.set(qn('w:before'), str(_pt_to_twentieth(para["spaceBeforePt"])))
    if para.get("spaceAfterPt") is not None:
        spacing = _ensure_child(pPr, 'w:spacing')
        spacing.set(qn('w:after'), str(_pt_to_twentieth(para["spaceAfterPt"])))
    # 缩进（字符数 → 大约 1 字符 = 2 字符宽度的 1/2 cm；此处用 cm 更可靠）
    ind = None
    if para.get("leftIndentCm") is not None or para.get("rightIndentCm") is not None or para.get("firstLineChars") is not None or para.get("hangingChars") is not None:
        ind = _ensure_child(pPr, 'w:ind')
    if para.get("leftIndentCm") is not None:
        ind.set(qn('w:left'), str(int(para["leftIndentCm"] * 567)))  # cm→twips (1cm≈567)
    if para.get("rightIndentCm") is not None:
        ind.set(qn('w:right'), str(int(para["rightIndentCm"] * 567)))
    # 首行/悬挂（字符数 → 这里简化按每字符 2 个汉字宽，估 0.74cm/字；可按需校准）
    if para.get("firstLineChars") is not None:
        ind.set(qn('w:firstLine'), str(int(para["firstLineChars"] * 0.74 * 567)))
    if para.get("hangingChars") is not None:
        ind.set(qn('w:hanging'), str(int(para["hangingChars"] * 0.74 * 567)))
    if para.get("outlineLevel") is not None:
        ol = _ensure_child(pPr, 'w:outlineLvl')
        ol.set(qn('w:val'), str(int(para["outlineLevel"])))
    if para.get("keepNext") is not None:
        if para["keepNext"]:
            _ensure_child(pPr, 'w:keepNext')
        else:
            kn = pPr.find(qn('w:keepNext'))
            if kn is not None:
                pPr.remove(kn)

class StyleResolver:
    """
    解析/创建样式：
    - 先用文档内已有样式（通常来自模板 DOCX）
    - 其次用 JSON stylesInline 动态创建
    - prefer_json_styles/$override 控制是否覆盖同名样式字段
    """
    def __init__(self, document, styles_inline: dict, prefer_json_styles: bool = False):
        self.document = document
        self.styles_inline = styles_inline or {}
        self.prefer_json_styles = prefer_json_styles

    def _doc_style_by_name(self, name):
        try:
            return self.document.styles[name]
        except KeyError:
            return None

    def ensure_style(self, name: str, expected_type: str):
        """
        返回 python-docx 的 style 对象；必要时依据 JSON 定义创建或受控覆盖。
        expected_type: 'paragraph' | 'character' | 'table'
        """
        st = self._doc_style_by_name(name)
        json_def = self.styles_inline.get(name)

        # 创建：文档无、JSON 提供
        if st is None and json_def:
            st = self._create_style_from_json(name, json_def)
            return st

        # 覆盖：文档有、JSON 也有，且允许覆盖
        if st is not None and json_def:
            if json_def.get("$override") or self.prefer_json_styles:
                self._apply_json_to_style(st, json_def)
            return st

        return st  # 文档已有；或找不到（返回 None）

    def _create_style_from_json(self, name: str, jd: dict):
        stype = jd.get("type")
        from docx.enum.style import WD_STYLE_TYPE
        mapping = {
            "paragraph": WD_STYLE_TYPE.PARAGRAPH,
            "character": WD_STYLE_TYPE.CHARACTER,
            "table": WD_STYLE_TYPE.TABLE,
        }
        if stype not in mapping:
            raise ValueError("Unsupported style type for JSON inline style: %s" % stype)
        st = self.document.styles.add_style(name, mapping[stype])
        self._apply_json_to_style(st, jd)
        return st

    def _apply_json_to_style(self, st, jd: dict):
        # basedOn
        if jd.get("basedOn"):
            base = self._doc_style_by_name(jd["basedOn"])
            if base is not None:
                st.base_style = base
        # font / paragraph
        s_el = st._element
        rPr = s_el.get_or_add_rPr()
        pPr = s_el.get_or_add_pPr()
        apply_font_rpr(rPr, jd.get("font"))
        apply_paragraph_ppr(pPr, jd.get("paragraph"))
        # numbering（简化：仅建立 outlineLevel，真正 numPr 绑定建议由模板完成或后续扩展）
        # 如需强绑多级编号，可在此写入 w:numPr (抽象编号/级别)，此处暂留扩展点。
