# OOXML 命名空间
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "rels": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# 中文字号名 ↔ pt 对照（取常用）
CN_FONT_SIZE_PT = {
    "初号": 42.0, "小初": 36.0, "一号": 26.0, "小一": 24.0,
    "二号": 22.0, "小二": 18.0, "三号": 16.0, "小三": 15.0,
    "四号": 14.0, "小四": 12.0, "五号": 10.5, "小五": 9.0,
    "六号": 7.5,  "小六": 6.5,  "七号": 5.5,  "八号": 5.0,
}

# 行距换算（auto模式下）
LINE_AUTO_MAP = {
    "single": 240,   # 单倍
    "1.5":    360,   # 1.5倍
    "double": 480,   # 两倍
}
