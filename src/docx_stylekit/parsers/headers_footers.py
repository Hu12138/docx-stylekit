from ..utils.xml import parse_bytes, findall

def detect_page_field(xml_bytes):
    """
    简单识别是否包含 PAGE 域（fldSimple 或 instrText）
    """
    if not xml_bytes:
        return {"has_page": False, "patterns": []}
    root = parse_bytes(xml_bytes)
    patterns = []

    # <w:fldSimple w:instr="PAGE">
    for fld in findall(root, ".//w:fldSimple"):
        instr = fld.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instr")
        if instr and "PAGE" in instr:
            patterns.append(instr)

    # 分散域：<w:r><w:instrText>PAGE</w:instrText></w:r>
    for it in findall(root, ".//w:instrText"):
        if it.text and "PAGE" in it.text:
            patterns.append(it.text)

    return {"has_page": bool(patterns), "patterns": list(set(patterns))}
