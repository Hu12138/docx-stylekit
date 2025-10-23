# src/docx_stylekit/writer/docx_writer.py
import json
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from .style_store import StyleResolver
from .section_utils import apply_section_layout, add_page_number_field, add_toc_field

def _clear_cell(cell):
    while cell._tc.getchildren():
        cell._tc.remove(cell._tc.getchildren()[-1])

def _set_cell_shading(cell, color_hex: str):
    if not color_hex:
        return
    color = color_hex.lstrip("#").upper()
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tc_pr.append(shd)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)

def _set_cell_border(cell, border: dict):
    if not border:
        return
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn('w:tcBorders'))
    if tc_borders is None:
        tc_borders = OxmlElement('w:tcBorders')
        tc_pr.append(tc_borders)
    for edge in ("top", "bottom", "left", "right"):
        cfg = border.get(edge)
        if not cfg:
            continue
        el = tc_borders.find(qn(f'w:{edge}'))
        if el is None:
            el = OxmlElement(f'w:{edge}')
            tc_borders.append(el)
        el.set(qn('w:val'), cfg.get("style", "single"))
        if "color" in cfg:
            el.set(qn('w:color'), cfg["color"].lstrip("#").upper())
        if "size" in cfg:
            el.set(qn('w:sz'), str(int(cfg["size"])))

def _apply_table_format(table, fmt: dict):
    if not fmt:
        return
    header_cfg = fmt.get("header")
    if header_cfg:
        for cell in table.rows[0].cells:
            _set_cell_shading(cell, header_cfg.get("fill"))
            _set_cell_border(cell, header_cfg.get("border"))
            for p in cell.paragraphs:
                for run in p.runs:
                    font = run.font
                    if header_cfg.get("bold") is not None:
                        font.bold = bool(header_cfg.get("bold"))
                    if header_cfg.get("color"):
                        font.color.rgb = RGBColor.from_string(header_cfg["color"].lstrip("#"))
    banding = fmt.get("bandedRows")
    if banding:
        alt_cfg = fmt.get("alternate")
        for idx, row in enumerate(table.rows[1:], start=1):
            if idx % 2 == 1 and alt_cfg:
                for cell in row.cells:
                    _set_cell_shading(cell, alt_cfg.get("fill"))
    table_border = fmt.get("tableBorder")
    if table_border:
        tbl = table._tbl
        tbl_pr = tbl.tblPr
        if tbl_pr is None:
            tbl_pr = OxmlElement('w:tblPr')
            tbl.append(tbl_pr)
        borders = tbl_pr.find(qn('w:tblBorders'))
        if borders is None:
            borders = OxmlElement('w:tblBorders')
            tbl_pr.append(borders)
        for edge in ("top","bottom","left","right","insideH","insideV"):
            cfg = table_border.get(edge)
            if not cfg:
                continue
            el = borders.find(qn(f'w:{edge}'))
            if el is None:
                el = OxmlElement(f'w:{edge}')
                borders.append(el)
            el.set(qn('w:val'), cfg.get("style","single"))
            if "color" in cfg:
                el.set(qn('w:color'), cfg["color"].lstrip("#").upper())
            if "size" in cfg:
                el.set(qn('w:sz'), str(int(cfg["size"])))

def _write_cell_blocks(cell, blocks, resolver: StyleResolver):
    _clear_cell(cell)
    for b in blocks or []:
        btype = b.get("type")
        if btype == "paragraph":
            st = resolver.ensure_style(b.get("styleRef", "Normal"), "paragraph")
            p = cell.add_paragraph(style=st.name if st else None)
            for r in b.get("runs", []):
                run = p.add_run(r.get("text", ""))
                cstyle = resolver.ensure_style(r.get("charStyleRef"), "character") if r.get("charStyleRef") else None
                if cstyle:
                    run.style = cstyle
        elif btype == "heading":
            level = int(b.get("level", 1))
            st = resolver.ensure_style(b.get("styleRef", f"Heading {level}"), "paragraph")
            cell.add_paragraph(b.get("text", ""), style=st.name if st else None)
        elif btype == "caption":
            st = resolver.ensure_style(b.get("styleRef", "Caption"), "paragraph")
            p = cell.add_paragraph(style=st.name if st else None)
            p.add_run(b.get("text", ""))

def write_blocks(doc: Document, blocks: list, resolver: StyleResolver, fail_on_unknown_style: bool = True):
    for b in blocks or []:
        # 页面模板调用：新建节 + 应用布局 + 渲染模板内部 blocks
        if "useTemplate" in b:
            tpl_name = b["useTemplate"]
            page_templates = getattr(doc, "_page_templates_cfg", {})  # 在 render_to_docx 中挂上
            tpl = page_templates.get(tpl_name)
            if not tpl:
                raise ValueError(f"useTemplate 指向的页面模板不存在：{tpl_name}")
            # 新建节
            section = doc.add_section()
            apply_section_layout(section, tpl.get("layout"))
            # 局部变量合并：已在 expand_document 合并到 b["variables"]，此处仅透传
            # 渲染模板 blocks
            write_blocks(doc, tpl.get("blocks", []), resolver, fail_on_unknown_style)
            continue

        btype = b.get("type")
        if btype == "pageBreak":
            p = doc.add_paragraph()
            p.paragraph_format.page_break_before = True
            continue

        if btype == "toc":
            p = doc.add_paragraph()
            doc_toc = getattr(doc, "_toc_levels", [1,3])
            add_toc_field(p, doc_toc)
            continue

        if btype in ("paragraph", "caption"):
            stname = b.get("styleRef", "Normal")
            st = resolver.ensure_style(stname, "paragraph")
            if fail_on_unknown_style and st is None:
                raise ValueError(f"未知样式（段落）：{stname}")
            p = doc.add_paragraph(style=st.name if st else None)
            if b.get("pageBreakBefore"):
                p.paragraph_format.page_break_before = True
            for r in b.get("runs", []):
                run = p.add_run(r.get("text", ""))
                cstyle = resolver.ensure_style(r.get("charStyleRef"), "character") if r.get("charStyleRef") else None
                if cstyle:
                    run.style = cstyle
            if btype == "caption" and not b.get("runs"):
                p.add_run(b.get("text", ""))
            continue

        if btype == "heading":
            level = int(b.get("level", 1))
            stname = b.get("styleRef", f"Heading {level}")
            st = resolver.ensure_style(stname, "paragraph")
            if fail_on_unknown_style and st is None:
                raise ValueError(f"未知样式（标题）：{stname}")
            doc.add_paragraph(b.get("text", ""), style=st.name if st else None)
            continue

        if btype == "list":
            ordered = bool(b.get("ordered", False))
            stname = b.get("styleRef", "Normal")
            st = resolver.ensure_style(stname, "paragraph")
            if fail_on_unknown_style and st is None:
                raise ValueError(f"未知样式（列表段落）：{stname}")
            for it in b.get("items", []):
                p = doc.add_paragraph(style=st.name if st else None)
                for r in it.get("runs", []):
                    run = p.add_run(r.get("text", ""))
                    cstyle = resolver.ensure_style(r.get("charStyleRef"), "character") if r.get("charStyleRef") else None
                    if cstyle:
                        run.style = cstyle
            continue

        if btype == "table":
            columns = b.get("columns", [])
            ncols = max(1, len(columns))
            header = b.get("header", [])
            rows = b.get("rows", [])
            total_rows = len(header) + len(rows)
            if total_rows == 0:
                total_rows = 1
            table = doc.add_table(rows=total_rows, cols=ncols)
            tstyle = resolver.ensure_style(b.get("styleRef"), "table") if b.get("styleRef") else None
            if tstyle:
                table.style = tstyle
            r_idx = 0
            for hrow in header:
                for c_idx, cell in enumerate(hrow):
                    _write_cell_blocks(table.cell(r_idx, c_idx), cell.get("blocks", []), resolver)
                r_idx += 1
            for drow in rows:
                for c_idx, cell in enumerate(drow):
                    _write_cell_blocks(table.cell(r_idx, c_idx), cell.get("blocks", []), resolver)
                r_idx += 1
            _apply_table_format(table, b.get("format"))
            continue

        # 其它类型（figure 等）可按需扩展
# src/docx_stylekit/writer/docx_writer.py （续）
import os
import json
import yaml
from .style_store import StyleResolver
from .section_utils import apply_section_layout, add_page_number_field, add_toc_field

def render_to_docx(template_json: dict,
                   template_docx_path: str,
                   styles_yaml: dict = None,
                   output_path: str = "output.docx",
                   prefer_json_styles: bool = False,
                   fail_on_unknown_style: bool = True):
    """
    template_json: expand_document() 的结果（已展开变量/循环/条件；保留 useTemplate）
    template_docx_path: 样式/编号/页眉页脚基础骨架（推荐提供）
    styles_yaml: 合并后的 YAML（dict）。如传入路径字符串则会自动读取。
    """
    doc_cfg = template_json.get("doc", {})
    # 读取 JSON 内联样式 / 页面模板 / TOC 级别
    styles_inline = doc_cfg.get("stylesInline", {}) or {}
    page_templates = doc_cfg.get("pageTemplates", {}) or {}
    toc_levels = doc_cfg.get("toc", {}).get("levels", [1,3])

    # 打开模板 DOCX
    doc = Document(template_docx_path)
    # 挂载配置供 writer 使用
    doc._page_templates_cfg = page_templates
    doc._toc_levels = toc_levels

    # 全局 pageSetup（第一节）
    ps = doc_cfg.get("pageSetup", {})
    if ps:
        first = doc.sections[0]
        layout = {
            "marginsCm": ps.get("marginsCm"),
            "orientation": ps.get("orientation"),
            "titleFirstPageDifferent": ps.get("titleFirstPageDifferent"),
            "evenOddDifferent": ps.get("evenOddDifferent"),
            "startAt": ps.get("pageNumbering", {}).get("startAt"),
        }
        apply_section_layout(first, layout)

    # YAML 样式库：此实现依赖模板 DOCX 自带样式；YAML 用于校验/提示（如需从 YAML 动态创建，可在此扩展）
    if isinstance(styles_yaml, str) and os.path.exists(styles_yaml):
        with open(styles_yaml, "r", encoding="utf-8") as f:
            _ = yaml.safe_load(f)  # 预留：可用于验证 styleCatalog 白名单

    # 构建解析器（支持 JSON 动态新增样式 & 受控覆盖）
    resolver = StyleResolver(doc, styles_inline, prefer_json_styles=prefer_json_styles)

    # TOC（顶层 doc.toc.required 也可以在 blocks 里单独放 type:"toc" 控制位置）
    # 若需要固定在文档开头：可以在此插入；这里尊重 blocks 中的显式位置。

    # 内容写入
    write_blocks(doc, doc_cfg.get("blocks", []), resolver, fail_on_unknown_style=fail_on_unknown_style)

    # 页眉页脚页码（若 JSON 指定 pageNumber 项，模板未内置时可插入）
    hf = doc_cfg.get("headersFooters", {})
    if "footer" in hf:
        for comp in hf["footer"]:
            if comp.get("type") == "pageNumber":
                for section in doc.sections:
                    fp = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
                    if not fp.text.strip():
                        add_page_number_field(fp, align=comp.get("align", "center"))

    doc.save(output_path)
