from ..utils.xml import parse_bytes, find, findall, attr

def parse_theme(xml_bytes):
    """
    返回:
    {
      "colors": {"accent1":"#112233", ... , "text1":"#000000","background1":"#FFFFFF"},
      "fonts":  {"major":{"latin":"Times New Roman","ea":"仿宋_GB2312"},
                 "minor":{"latin":"Times New Roman","ea":"仿宋_GB2312"}}
    }
    """
    out = {"colors": {}, "fonts": {"major": {}, "minor": {}}}
    if not xml_bytes:
        return out
    root = parse_bytes(xml_bytes)

    # 颜色
    clr_scheme = find(root, ".//a:themeElements/a:clrScheme")
    if clr_scheme is not None:
        for child in clr_scheme:
            key = child.tag.split("}")[1]  # accent1, text1, background1...
            srgb = find(child, ".//a:srgbClr")
            if srgb is not None:
                out["colors"][key] = f'#{attr(srgb, "val", "000000")}'.upper()

    # 字体
    font_scheme = find(root, ".//a:themeElements/a:fontScheme")
    for group, slot in (("major", ".//a:majorFont"), ("minor", ".//a:minorFont")):
        node = find(font_scheme, slot) if font_scheme is not None else None
        if node is not None:
            latin = find(node, "a:latin")
            ea = find(node, "a:ea")
            if latin is not None:
                out["fonts"][group]["latin"] = attr(latin, "typeface")
            if ea is not None:
                out["fonts"][group]["ea"] = attr(ea, "typeface")

    return out
