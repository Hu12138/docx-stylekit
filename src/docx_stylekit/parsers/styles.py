from ..utils.xml import parse_bytes, findall, find, attr
from ..utils.units import halfpoints_to_pt
from ..constants import NS, CN_FONT_SIZE_PT

def closest_cn_size_name(pt):
    # 最近邻匹配中文字号名
    items = sorted(CN_FONT_SIZE_PT.items(), key=lambda kv: abs(kv[1]-pt))
    return items[0][0] if items else None

def parse_styles(xml_bytes):
    """
    返回：
    {
      "paragraph_styles": {
          "Normal": {
            "name": "正文",
            "based_on": None,
            "rPr": {"eastAsia":"仿宋_GB2312","ascii":"Times New Roman","size_pt":16.0,"bold":False,"italic":False,"underline":"none","color":{"theme":"text1"|"#112233"}},
            "pPr": {"alignment":"justify","line":{"rule":"1.5"/"single"/"double"/"exact"/"at_least","value_pt":None},"space_before_pt":0,"space_after_pt":0,
                    "indent":{"first_line":{"type":"chars/cm/none","value":2},"left_cm":0,"right_cm":0},
                    "outline_level": null}
          }, ...
      },
      "character_styles": {...},
      "table_styles": {...},
      "doc_defaults": {...}
    }
    """
    if not xml_bytes:
        return {"paragraph_styles": {}, "character_styles": {}, "table_styles": {}, "doc_defaults": {}}

    root = parse_bytes(xml_bytes)
    out = {"paragraph_styles": {}, "character_styles": {}, "table_styles": {}, "doc_defaults": {}}

    # docDefaults
    rdef = find(root, ".//w:docDefaults/w:rPrDefault/w:rPr")
    pdef = find(root, ".//w:docDefaults/w:pPrDefault/w:pPr")
    out["doc_defaults"] = {
        "rPr": _read_rPr(rdef),
        "pPr": _read_pPr(pdef),
    }

    # styles
    for st in findall(root, ".//w:style"):
        st_type = attr(st, "{%s}type" % NS["w"])
        st_id = attr(st, "{%s}styleId" % NS["w"])
        st_name = attr(find(st, "w:name"), "{%s}val" % NS["w"])
        based_on = attr(find(st, "w:basedOn"), "{%s}val" % NS["w"])
        link = attr(find(st, "w:link"), "{%s}val" % NS["w"])

        item = {
            "name": st_name,
            "based_on": based_on,
            "link_char_style": link,
            "rPr": _read_rPr(find(st, "w:rPr")),
            "pPr": _read_pPr(find(st, "w:pPr")),
        }

        if st_type == "paragraph":
            out["paragraph_styles"][st_id] = item
        elif st_type == "character":
            out["character_styles"][st_id] = item
        elif st_type == "table":
            out["table_styles"][st_id] = item

    return out

def _read_rPr(node):
    if node is None:
        return {}
    rFonts = find(node, "w:rFonts")
    color = find(node, "w:color")
    sz = find(node, "w:sz")
    out = {
        "eastAsia": attr(rFonts, "{%s}eastAsia" % NS["w"]) if rFonts is not None else None,
        "ascii": attr(rFonts, "{%s}ascii" % NS["w"]) if rFonts is not None else None,
        "bold": find(node, "w:b") is not None,
        "italic": find(node, "w:i") is not None,
        "underline": attr(find(node, "w:u"), "{%s}val" % NS["w"]) if find(node, "w:u") is not None else "none",
        "size_pt": halfpoints_to_pt(float(attr(sz, "{%s}val" % NS["w"], 0))) if sz is not None else None,
        "size_cn": None,
        "color": None,
    }
    if out["size_pt"]:
        out["size_cn"] = closest_cn_size_name(out["size_pt"])
    if color is not None:
        val = attr(color, "{%s}val" % NS["w"])
        theme = attr(color, "{%s}themeColor" % NS["w"])
        out["color"] = ({"hex": f"#{val}"} if val else None) or ({"theme": theme} if theme else None)
    return out

def _read_pPr(node):
    if node is None:
        return {}
    spacing = find(node, "w:spacing")
    ind = find(node, "w:ind")
    jc = find(node, "w:jc")
    outline = find(node, "w:outlineLvl")

    # 行距
    line = None
    if spacing is not None:
        line_rule = attr(spacing, "{%s}lineRule" % NS["w"])
        line_val = attr(spacing, "{%s}line" % NS["w"])
        before = attr(spacing, "{%s}before" % NS["w"], 0)
        after = attr(spacing, "{%s}after" % NS["w"], 0)
        if line_rule in ("auto", None):
            if line_val:
                lv = int(line_val)
                if lv == 240: kind = "single"
                elif lv == 360: kind = "1.5"
                elif lv == 480: kind = "double"
                else: kind = "auto"
                line = {"rule": kind}
            else:
                line = {"rule": "single"}
        elif line_rule == "exact":
            line = {"rule": "exact", "value_pt": int(line_val)/20.0 if line_val else None}
        elif line_rule == "atLeast":
            line = {"rule": "at_least", "value_pt": int(line_val)/20.0 if line_val else None}
    else:
        before = 0
        after = 0

    # 缩进（优先 firstLineChars）
    first_line = None
    if ind is not None:
        flc = attr(ind, "{%s}firstLineChars" % NS["w"])
        fl = attr(ind, "{%s}firstLine" % NS["w"])
        if flc:
            first_line = {"type": "chars", "value": int(flc)/100.0}
        elif fl:
            # 以 cm 近似
            from ..utils.units import twips_to_cm
            first_line = {"type": "cm", "value": twips_to_cm(int(fl))}
        left = attr(ind, "{%s}left" % NS["w"], 0)
        right = attr(ind, "{%s}right" % NS["w"], 0)
        from ..utils.units import twips_to_cm
        left_cm = twips_to_cm(int(left))
        right_cm = twips_to_cm(int(right))
    else:
        first_line = None
        left_cm = 0
        right_cm = 0

    return {
        "alignment": attr(jc, "{%s}val" % NS["w"]) if jc is not None else None,
        "line": line,
        "space_before_pt": int(before)/20.0 if isinstance(before, str) else 0,
        "space_after_pt": int(after)/20.0 if isinstance(after, str) else 0,
        "indent": {"first_line": first_line, "left_cm": left_cm, "right_cm": right_cm},
        "outline_level": int(attr(outline, "{%s}val" % NS["w"])) if outline is not None else None,
        "keep_next": find(node, "w:keepNext") is not None,
    }
