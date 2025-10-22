from ..utils.xml import parse_bytes, find, findall, attr
from ..utils.units import twips_to_cm
from ..constants import NS

def parse_sections(xml_bytes):
    """
    抽取每个节的页面设置、页码起始等（取第一节作为默认）。
    返回：
    {
      "sections": [
        {"pgSz":{"w_cm":21.0,"h_cm":29.7,"orient":"portrait"},
         "pgMar":{"top":2.5,"bottom":2.5,"left":2.8,"right":2.8,"header":1.5,"footer":1.75,"gutter":0},
         "titlePg": true/false,
         "pgNumStart": 1 or None,
         "headerRefs":[{"type":"default","rId":"rId7"}, ...],
         "footerRefs":[...]
        },
        ...
      ]
    }
    """
    out = {"sections": []}
    if not xml_bytes:
        return out
    root = parse_bytes(xml_bytes)

    for sectPr in findall(root, ".//w:sectPr"):
        pgSz = find(sectPr, "w:pgSz")
        pgMar = find(sectPr, "w:pgMar")
        titlePg = find(sectPr, "w:titlePg") is not None
        pgNum = find(sectPr, "w:pgNumType")
        pgStart = attr(pgNum, "{%s}start" % NS["w"]) if pgNum is not None else None

        w_cm = twips_to_cm(attr(pgSz, "{%s}w" % NS["w"], 11906))
        h_cm = twips_to_cm(attr(pgSz, "{%s}h" % NS["w"], 16838))
        orient = attr(pgSz, "{%s}orient" % NS["w"], "portrait")

        def m(name, default):
            return twips_to_cm(attr(pgMar, "{%s}%s" % (NS["w"], name), default))

        mar = {
            "top": m("top", 1440),
            "bottom": m("bottom", 1440),
            "left": m("left", 1701),
            "right": m("right", 1701),
            "header": m("header", 851),
            "footer": m("footer", 992),
            "gutter": m("gutter", 0),
        }

        hrefs = []
        frefs = []
        for hr in findall(sectPr, "w:headerReference"):
            hrefs.append({"type": attr(hr, "{%s}type" % NS["w"], "default"), "rId": attr(hr, "{%s}id" % NS["r"])})
        for fr in findall(sectPr, "w:footerReference"):
            frefs.append({"type": attr(fr, "{%s}type" % NS["w"], "default"), "rId": attr(fr, "{%s}id" % NS["r"])})

        out["sections"].append({
            "pgSz": {"w_cm": w_cm, "h_cm": h_cm, "orient": orient},
            "pgMar": mar,
            "titlePg": titlePg,
            "pgNumStart": int(pgStart) if pgStart else None,
            "headerRefs": hrefs,
            "footerRefs": frefs,
        })
    return out
