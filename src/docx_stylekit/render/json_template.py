# src/docx_stylekit/render/json_template.py
from copy import deepcopy
import re

_VAR_PATTERN = re.compile(r"\{([A-Za-z0-9_.]+)\}")

def _get_var(vars_dict, dotted):
    cur = vars_dict
    for part in dotted.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        else:
            return None
    return cur

def substitute_text(text: str, vars_dict: dict) -> str:
    def repl(m):
        key = m.group(1)
        val = _get_var(vars_dict, key)
        return "" if val is None else str(val)
    return _VAR_PATTERN.sub(repl, text or "")

def expand_blocks(blocks: list, vars_dict: dict) -> list:
    """
    输入含有 repeat/conditional/variable 的 blocks，
    输出为“纯内容块”列表（仅 paragraph/heading/list/table/caption/pageBreak 等）。
    """
    out = []
    for b in blocks or []:
        btype = b.get("type")
        if btype == "repeat":
            arr = _get_var(vars_dict, b.get("for", ""))
            as_name = b.get("as", "item")
            template = b.get("template", [])
            if isinstance(arr, list):
                for item in arr:
                    local_vars = {**vars_dict, as_name: item}
                    out.extend(expand_blocks(template, local_vars))
            continue
        if btype == "conditional":
            cond_key = b.get("if")
            truthy = False
            if isinstance(cond_key, str):
                val = _get_var(vars_dict, cond_key)
                truthy = bool(val)
            elif isinstance(cond_key, bool):
                truthy = cond_key
            branch = b.get("then", []) if truthy else b.get("else", [])
            out.extend(expand_blocks(branch, vars_dict))
            continue
        if btype == "variable":
            # 单变量段：转 paragraph
            text = substitute_text(b.get("text", ""), vars_dict)
            out.append({
                "type": "paragraph",
                "styleRef": b.get("styleRef", "Normal"),
                "runs": [{"text": text, "charStyleRef": b.get("charStyleRef")}]
            })
            continue

        nb = deepcopy(b)
        # 段落 runs/标题 text/列表 items/table 单元内文本做变量替换
        if "text" in nb:
            nb["text"] = substitute_text(nb.get("text", ""), vars_dict)
        if "runs" in nb and isinstance(nb["runs"], list):
            for r in nb["runs"]:
                r["text"] = substitute_text(r.get("text", ""), vars_dict)
        if btype == "list":
            # 展开 items 内的 runs
            items = []
            for it in nb.get("items", []):
                eit = deepcopy(it)
                if "runs" in eit:
                    for r in eit["runs"]:
                        r["text"] = substitute_text(r.get("text", ""), vars_dict)
                items.append(eit)
            nb["items"] = items
        if btype == "table":
            # header/rows: 每个 cell.blocks 递归处理
            def _proc_rows(rows):
                out_rows = []
                for row in rows or []:
                    out_row = []
                    for cell in row:
                        ec = deepcopy(cell)
                        ec["blocks"] = expand_blocks(ec.get("blocks", []), vars_dict)
                        out_row.append(ec)
                    out_rows.append(out_row)
                return out_rows
            nb["header"] = _proc_rows(nb.get("header", []))
            # rows 可能含 repeat 模式：兼容 expand_blocks 之前的 repeat 设计
            if isinstance(nb.get("rows"), dict) and "repeat" in nb["rows"]:
                # { "repeat": { "for": "ARR", "as": "item", "template": [ row, ... ] } }
                repeat = nb["rows"]["repeat"]
                arr = _get_var(vars_dict, repeat.get("for", ""))
                as_name = repeat.get("as", "item")
                tpl = repeat.get("template", [])
                rrows = []
                if isinstance(arr, list):
                    for item in arr:
                        local_vars = {**vars_dict, as_name: item}
                        # template 是一组行，每行是一组 cell
                        for tpl_row in tpl:
                            rrow = []
                            for cell in tpl_row:
                                ec = deepcopy(cell)
                                ec["blocks"] = expand_blocks(ec.get("blocks", []), local_vars)
                                rrow.append(ec)
                            rrows.append(rrow)
                nb["rows"] = rrows
            else:
                nb["rows"] = _proc_rows(nb.get("rows", []))

        out.append(nb)
    return out

def expand_document(template_json: dict) -> dict:
    doc = deepcopy(template_json.get("doc", {}))
    vars_dict = doc.get("variables", {})
    blocks = doc.get("blocks", [])
    doc["blocks"] = expand_blocks(blocks, vars_dict)
    return {"doc": doc}
