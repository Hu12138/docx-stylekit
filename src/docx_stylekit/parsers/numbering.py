from ..utils.xml import parse_bytes, findall, find, attr

def parse_numbering(xml_bytes):
    """
    返回：
    {
      "abstract": {
         absId: {
            "levels": {
               0: {"fmt":"decimal","text":"%1","pStyle":"Heading1"},
               1: {"fmt":"decimal","text":"%1.%2","pStyle":"Heading2"},
               ...
            }
         }
      },
      "nums": { numId: absId }
    }
    """
    if not xml_bytes:
        return {"abstract": {}, "nums": {}}
    root = parse_bytes(xml_bytes)
    out = {"abstract": {}, "nums": {}}

    for absn in findall(root, ".//w:abstractNum"):
        abs_id = attr(absn, "{%s}abstractNumId" % absn.nsmap["w"])
        lvls = {}
        for lvl in findall(absn, "w:lvl"):
            val = int(attr(lvl, "{%s}ilvl" % lvl.nsmap["w"]))
            fmt = attr(find(lvl, "w:numFmt"), "{%s}val" % lvl.nsmap["w"])
            txt = attr(find(lvl, "w:lvlText"), "{%s}val" % lvl.nsmap["w"])
            pstyle = attr(find(lvl, "w:pStyle"), "{%s}val" % lvl.nsmap["w"])
            lvls[val] = {"fmt": fmt, "text": txt, "pStyle": pstyle}
        out["abstract"][abs_id] = {"levels": lvls}

    for num in findall(root, ".//w:num"):
        num_id = attr(num, "{%s}numId" % num.nsmap["w"])
        abs_ref = attr(find(num, "w:abstractNumId"), "{%s}val" % num.nsmap["w"])
        out["nums"][num_id] = abs_ref

    return out
