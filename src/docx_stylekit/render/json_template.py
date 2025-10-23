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

def _merge_vars(base: dict, override: dict) -> dict:
    if not override:
        return dict(base or {})
    out = dict(base or {})
    out.update(override or {})
    return out

def expand_blocks(blocks: list, vars_dict: dict) -> list:
    """
    展开 repeat / conditional / variable 的变量替换。
    保留 useTemplate（由 writer 在渲染时处理），但合并其 variables。
    """
    out = []
    for b in blocks or []:
        btype = b.get("type")
        # 控制块：repeat
        if btype == "repeat":
            arr = _get_var(vars_dict, b.get("for", ""))
            as_name = b.get("as", "item")
            template = b.get("template", [])
            if isinstance(arr, list):
                for item in arr:
                    local_vars = {**vars_dict, as_name: item}
                    out.extend(expand_blocks(template, local_vars))
            continue

        # 控制块：conditional
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

        # 语法糖：variable → paragraph
        if btype == "variable":
            text = substitute_text(b.get("text", ""), vars_dict)
            out.append({
                "type": "paragraph",
                "styleRef": b.get("styleRef", "Normal"),
                "runs": [{"text": text, "charStyleRef": b.get("charStyleRef")}]
            })
            continue

        # useTemplate：仅合并 variables；交由 writer 执行新建 section + 渲染 blocks
        if "useTemplate" in b:
            nb = deepcopy(b)
            nb["variables"] = _merge_vars(vars_dict, b.get("variables") or {})
            out.append(nb)
            continue

        # 普通块的文本替换
        nb = deepcopy(b)
        if "text" in nb:
            nb["text"] = substitute_text(nb.get("text", ""), vars_dict)
        if "runs" in nb and isinstance(nb["runs"], list):
            for r in nb["runs"]:
                r["text"] = substitute_text(r.get("text", ""), vars_dict)

        # 列表项 runs 替换
        if btype == "list":
            items = []
            for it in nb.get("items", []):
                eit = deepcopy(it)
                if "runs" in eit:
                    for r in eit["runs"]:
                        r["text"] = substitute_text(r.get("text", ""), vars_dict)
                items.append(eit)
            nb["items"] = items

        # 表格 cell.blocks 递归替换
        if btype == "table":
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
            if isinstance(nb.get("rows"), dict) and "repeat" in nb["rows"]:
                repeat = nb["rows"]["repeat"]
                arr = _get_var(vars_dict, repeat.get("for", ""))
                as_name = repeat.get("as", "item")
                tpl = repeat.get("template", [])
                rrows = []
                if isinstance(arr, list):
                    for item in arr:
                        local_vars = {**vars_dict, as_name: item}
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
