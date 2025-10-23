# src/docx_stylekit/writer/docx_writer.py
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def add_page_number_run(paragraph):
    """
    在段落中插入 PAGE 域： { PAGE }
    """
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = ""
    r.append(t)
    fld.append(r)
    paragraph._p.append(fld)

def add_toc(paragraph, levels=(1,3)):
    """
    插入 TOC 域（1..levels）；Word 打开后需手动 F9 更新。
    """
    start, end = levels if isinstance(levels, (list, tuple)) and len(levels) == 2 else (1, 3)
    instr = f'TOC \\o "{start}-{end}" \\h \\z \\u'
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), instr)
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = "目录（打开后更新）"
    r.append(t)
    fld.append(r)
    paragraph._p.append(fld)

def write_blocks(doc: Document, blocks: list, style_catalog: dict):
    for b in blocks:
        btype = b.get("type")
        if btype == "pageBreak":
            p = doc.add_paragraph()
            p.paragraph_format.page_break_before = True
            continue

        if btype in ("paragraph", "caption"):
            style = b.get("styleRef", "Normal")
            p = doc.add_paragraph(style=style if style in style_catalog.get("paragraph", []) else None)
            if b.get("pageBreakBefore"):
                p.paragraph_format.page_break_before = True
            for r in b.get("runs", []):
                run = p.add_run(r.get("text", ""))
                cstyle = r.get("charStyleRef")
                if cstyle and cstyle in style_catalog.get("character", []):
                    run.style = cstyle
            continue

        if btype == "heading":
            level = int(b.get("level", 1))
            style = b.get("styleRef", f"Heading{level}")
            p = doc.add_paragraph(b.get("text", ""), style=style if style in style_catalog.get("paragraph", []) else None)
            continue

        if btype == "list":
            # 简化实现：使用段落+项目符号/编号由模板样式控制；
            # 如果需要强制项目符号，需要更深入的 numPr 写入，这里依赖样式。
            ordered = bool(b.get("ordered", False))
            style = b.get("styleRef", "Normal")
            for it in b.get("items", []):
                p = doc.add_paragraph(style=style if style in style_catalog.get("paragraph", []) else None)
                for r in it.get("runs", []):
                    run = p.add_run(r.get("text", ""))
                    cstyle = r.get("charStyleRef")
                    if cstyle and cstyle in style_catalog.get("character", []):
                        run.style = cstyle
            continue

        if btype == "table":
            columns = b.get("columns", [])
            ncols = len(columns)
            header = b.get("header", [])
            rows = b.get("rows", [])

            # 先估计总行数：header行数 + 数据行数
            total_rows = len(header) + len(rows)
            if total_rows == 0:
                total_rows = 1
            table = doc.add_table(rows=total_rows, cols=max(1, ncols))
            tstyle = b.get("styleRef")
            if tstyle and tstyle in style_catalog.get("table", []):
                table.style = tstyle

            # 写头部
            r_idx = 0
            for hrow in header:
                for c_idx, cell in enumerate(hrow):
                    _write_cell_blocks(table.cell(r_idx, c_idx), cell.get("blocks", []), style_catalog)
                r_idx += 1
            # 写数据
            for drow in rows:
                for c_idx, cell in enumerate(drow):
                    _write_cell_blocks(table.cell(r_idx, c_idx), cell.get("blocks", []), style_catalog)
                r_idx += 1
            # 列宽与对齐（python-docx 对列宽支持有限；通常保持样式统一）
            continue

        # 其它类型（figure/variable 等）可逐步扩展

def _write_cell_blocks(cell, blocks, style_catalog):
    # 清空默认段落
    cell.text = ""
    for b in blocks:
        if b.get("type") == "paragraph":
            p = cell.add_paragraph(style=b.get("styleRef", "Normal"))
            for r in b.get("runs", []):
                run = p.add_run(r.get("text", ""))
                cstyle = r.get("charStyleRef")
                if cstyle and cstyle in style_catalog.get("character", []):
                    run.style = cstyle
        elif b.get("type") == "heading":
            level = int(b.get("level", 1))
            style = b.get("styleRef", f"Heading{level}")
            cell.add_paragraph(b.get("text", ""), style=style)
        elif b.get("type") == "caption":
            p = cell.add_paragraph(style=b.get("styleRef", "Caption"))
            p.add_run(b.get("text", ""))
        elif b.get("type") == "pageBreak":
            p = cell.add_paragraph()
            p.paragraph_format.page_break_before = True

def render_to_docx(template_json: dict, template_docx_path: str, output_path: str):
    """
    template_json: expand_document() 的结果（已展开控制结构）
    template_docx_path: 样式模板DOCX（含 Heading/Normal/TableGrid 等）
    output_path: 生成的DOCX
    """
    doc_cfg = template_json.get("doc", {})
    catalog = doc_cfg.get("styleCatalog", {"paragraph": [], "character": [], "table": []})
    doc = Document(template_docx_path)

    # TOC（如需要，在模板第二页或自定义位置插入）
    if doc_cfg.get("toc", {}).get("required"):
        p = doc.add_paragraph()
        add_toc(p, levels=(1, max(doc_cfg["toc"].get("levels", [1,3]))))

    write_blocks(doc, doc_cfg.get("blocks", []), catalog)

    # 页眉页脚页码域：若模板未设置而 JSON 指定 pattern=PAGE，则插入
    hf = doc_cfg.get("headersFooters", {})
    if "footer" in hf:
        for comp in hf["footer"]:
            if comp.get("type") == "pageNumber":
                for section in doc.sections:
                    fp = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
                    fp.alignment = {"left":0, "center":1, "right":2}.get(comp.get("align","center"), 1)
                    # 仅在空时插入一个 PAGE 域，避免重复
                    if not fp.text.strip():
                        add_page_number_run(fp)

    doc.save(output_path)
