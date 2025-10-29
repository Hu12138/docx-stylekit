from __future__ import annotations

from dataclasses import dataclass
import re
from typing import List, Optional, Dict, Any

from markdown_it import MarkdownIt
from markdown_it.token import Token


@dataclass
class InlineContext:
    strong: bool = False
    emphasis: bool = False
    code: bool = False

    def char_style(self) -> Optional[str]:
        if self.code:
            return "Code"
        if self.strong:
            return "Strong"
        if self.emphasis:
            return "Emphasis"
        return None


def markdown_to_template(markdown_text: str, *, title: Optional[str] = None) -> Dict[str, Any]:
    """
    将 Markdown 文本转换为 docx-stylekit JSON 模板结构。
    支持的 Markdown 要素：标题、段落、无序/有序列表、代码块、行内强调、表格。
    """
    markdown_text = _normalize_markdown_text(markdown_text)
    md = MarkdownIt("commonmark").enable("table")
    tokens = md.parse(markdown_text)
    has_explicit_headings = any(t.type == "heading_open" for t in tokens)
    blocks = _tokens_to_blocks(tokens, allow_heading_heuristics=not has_explicit_headings)

    derived_title = title or _extract_title_from_blocks(blocks)
    template = {
        "doc": {
            "meta": {
                "title": derived_title or "Markdown Document",
                "lang": "zh-CN",
                "version": "1.0.0",
                "createdAt": "2024-01-01T00:00:00Z",
            },
            "styleCatalog": {
                "paragraph": [
                    "Normal",
                    "Heading 1",
                    "Heading 2",
                    "Heading 3",
                    "Heading 4",
                    "Heading 5",
                    "Heading 6",
                    "Heading 7",
                    "Heading 8",
                    "Heading 9",
                    "List Paragraph",
                    "Quote",
                ],
                "character": [
                    "Default Paragraph Font",
                    "Strong",
                    "Emphasis",
                    "Code",
                ],
                "table": [
                    "Normal Table",
                    "InfoTable",
                ],
            },
            "pageSetup": {
                "paper": "A4",
                "orientation": "portrait",
                "marginsCm": {
                    "top": 2.5,
                    "bottom": 2.5,
                    "left": 2.8,
                    "right": 2.8,
                    "header": 1.5,
                    "footer": 1.75,
                },
                "titleFirstPageDifferent": False,
                "evenOddDifferent": False,
                "pageNumbering": {
                    "position": "footer",
                    "startAt": 1,
                    "format": "decimal",
                },
            },
            "numbering": {
                "bindHeadings": True,
                "preset": "decimal-dot",
                "levels": {
                    "1": {"styleRef": "Heading 1"},
                    "2": {"styleRef": "Heading 2"},
                    "3": {"styleRef": "Heading 3"},
                    "4": {"styleRef": "Heading 4"},
                    "5": {"styleRef": "Heading 5"},
                },
            },
            "toc": {
                "required": False,
                "levels": [1, 3],
                "note": "Word 中可通过更新域刷新目录。",
            },
            "blocks": blocks,
        }
    }
    return template


def _normalize_markdown_text(text: str) -> str:
    """
    处理一些来源不规范的 Markdown：
    - 有些文件保存为单行，使用 \"\\n\" 字符串而非真实换行，需要替换回真正的换行；
    - 同时处理 \"\\r\\n\" / \"\\t\" 等常见转义。
    """
    if "\n" not in text and "\\n" in text:
        text = text.replace("\\r\\n", "\n")
        text = text.replace("\\n", "\n")
    if "\t" not in text and "\\t" in text:
        text = text.replace("\\t", "\t")
    return text


def _extract_title_from_blocks(blocks: List[Dict[str, Any]]) -> Optional[str]:
    for block in blocks:
        if block.get("type") == "heading":
            return block.get("text")
    return None


def _tokens_to_blocks(tokens: List[Token], *, allow_heading_heuristics: bool) -> List[Dict[str, Any]]:
    blocks: List[Dict[str, Any]] = []
    idx = 0
    length = len(tokens)
    while idx < length:
        token = tokens[idx]
        ttype = token.type

        if ttype == "heading_open":
            level = int(token.tag[1])
            inline = tokens[idx + 1]
            text = inline.content.strip()
            blocks.append({
                "type": "heading",
                "level": level,
                "styleRef": f"Heading {min(level, 9)}",
                "text": text,
            })
            idx += 3  # heading_open, inline, heading_close
            continue

        if ttype == "paragraph_open":
            inline = tokens[idx + 1]
            runs = _convert_inline(inline.children or [])
            text_content = "".join(run["text"] for run in runs).strip()
            heading_level = _detect_heading_level(text_content, runs, allow_heading_heuristics)
            if heading_level:
                blocks.append({
                    "type": "heading",
                    "level": heading_level,
                    "styleRef": f"Heading {min(heading_level, 9)}",
                    "text": text_content,
                })
            else:
                blocks.append({
                    "type": "paragraph",
                    "styleRef": "Normal",
                    "runs": runs or [{"text": inline.content.strip()}],
                })
            idx += 3  # paragraph_open, inline, paragraph_close
            continue

        if ttype in {"bullet_list_open", "ordered_list_open"}:
            ordered = ttype == "ordered_list_open"
            items, idx = _parse_list(tokens, idx)
            blocks.append({
                "type": "list",
                "ordered": ordered,
                "styleRef": "List Paragraph",
                "items": items,
            })
            continue

        if ttype == "blockquote_open":
            quote_runs, idx = _parse_blockquote(tokens, idx)
            blocks.append({
                "type": "paragraph",
                "styleRef": "Quote",
                "runs": quote_runs,
            })
            continue

        if ttype == "fence":
            code_text = token.content.rstrip()
            blocks.append({
                "type": "paragraph",
                "styleRef": "Normal",
                "runs": [{"text": code_text, "charStyleRef": "Code"}],
            })
            idx += 1
            continue

        if ttype == "code_block":
            code_text = token.content.rstrip()
            blocks.append({
                "type": "paragraph",
                "styleRef": "Normal",
                "runs": [{"text": code_text, "charStyleRef": "Code"}],
            })
            idx += 1
            continue

        if ttype == "table_open":
            table_block, idx = _parse_table(tokens, idx)
            blocks.append(table_block)
            continue

        # fallthrough: skip token
        idx += 1

    return blocks


def _convert_inline(inline_tokens: List[Token]) -> List[Dict[str, Any]]:
    runs: List[Dict[str, Any]] = []
    ctx = InlineContext()
    buffer: List[str] = []

    def flush():
        if not buffer:
            return
        text = "".join(buffer)
        buffer.clear()
        run: Dict[str, Any] = {"text": text}
        style = ctx.char_style()
        if style:
            run["charStyleRef"] = style
        runs.append(run)

    for token in inline_tokens:
        if token.type == "text":
            buffer.append(token.content)
        elif token.type == "code_inline":
            flush()
            runs.append({"text": token.content, "charStyleRef": "Code"})
        elif token.type == "softbreak":
            buffer.append("\n")
        elif token.type == "hardbreak":
            buffer.append("\n")
        elif token.type == "strong_open":
            flush()
            ctx.strong = True
        elif token.type == "strong_close":
            flush()
            ctx.strong = False
        elif token.type == "em_open":
            flush()
            ctx.emphasis = True
        elif token.type == "em_close":
            flush()
            ctx.emphasis = False
        else:
            # 未显式处理的 inline 类型，先直接拼接
            if token.content:
                buffer.append(token.content)

    flush()
    return runs


def _parse_list(tokens: List[Token], idx: int):
    list_open = tokens[idx]
    ordered = list_open.type == "ordered_list_open"
    items = []
    idx += 1
    while idx < len(tokens):
        token = tokens[idx]
        if token.type == "list_item_open":
            idx += 1
            item_runs: List[Dict[str, Any]] = []
            while tokens[idx].type != "list_item_close":
                current = tokens[idx]
                if current.type == "paragraph_open":
                    inline = tokens[idx + 1]
                    para_runs = _convert_inline(inline.children or [])
                    item_runs.extend(para_runs or [{"text": inline.content.strip()}])
                    idx += 3
                    continue
                if current.type in {"bullet_list_open", "ordered_list_open"}:
                    sub_items, idx = _parse_list(tokens, idx)
                    sub_marker = "• " if not current.type.startswith("ordered") else "1. "
                    flat_text = "; ".join(
                        "".join(run["text"] for run in item["runs"])
                        for item in sub_items
                    )
                    item_runs.append({"text": f"{sub_marker}{flat_text}"})
                    continue
                if current.type == "fence":
                    item_runs.append({"text": current.content.rstrip(), "charStyleRef": "Code"})
                    idx += 1
                    continue
                # skip unknown
                idx += 1
            items.append({"runs": item_runs or [{"text": ""}]})
            idx += 1  # skip list_item_close
            continue
        if token.type in {"bullet_list_close", "ordered_list_close"}:
            idx += 1
            break
        idx += 1
    return items, idx


def _parse_blockquote(tokens: List[Token], idx: int):
    runs: List[Dict[str, Any]] = []
    idx += 1
    while idx < len(tokens):
        token = tokens[idx]
        if token.type == "paragraph_open":
            inline = tokens[idx + 1]
            runs.extend(_convert_inline(inline.children or []))
            idx += 3
            continue
        if token.type == "blockquote_close":
            idx += 1
            break
        idx += 1
    return runs or [{"text": ""}], idx


def _parse_table(tokens: List[Token], idx: int):
    header: List[List[Dict[str, Any]]] = []
    rows: List[List[Dict[str, Any]]] = []
    idx += 1
    while idx < len(tokens):
        token = tokens[idx]
        if token.type == "thead_open":
            hdr_rows, idx = _parse_table_rows(tokens, idx + 1, "thead_close")
            header = hdr_rows
            continue
        if token.type == "tbody_open":
            body_rows, idx = _parse_table_rows(tokens, idx + 1, "tbody_close")
            rows.extend(body_rows)
            continue
        if token.type == "table_close":
            idx += 1
            break
        idx += 1

    table_block = {
        "type": "table",
        "styleRef": "Normal Table",
        "header": header,
        "rows": rows,
    }
    return table_block, idx


def _parse_table_rows(tokens: List[Token], idx: int, closing_type: str):
    rows: List[List[Dict[str, Any]]] = []
    current_row: List[Dict[str, Any]] = []
    while idx < len(tokens):
        token = tokens[idx]
        if token.type == "tr_open":
            current_row = []
            idx += 1
            continue
        if token.type in {"th_open", "td_open"}:
            inline = tokens[idx + 1]
            cell_runs = _convert_inline(inline.children or [])
            cell_block = {
                "blocks": [
                    {
                        "type": "paragraph",
                        "styleRef": "TableBase",
                        "runs": cell_runs or [{"text": inline.content.strip()}],
                    }
                ]
            }
            current_row.append(cell_block)
            idx += 3  # cell_open, inline, cell_close
            continue
        if token.type == "tr_close":
            rows.append(current_row)
            idx += 1
            continue
        if token.type == closing_type:
            idx += 1
            break
        idx += 1
    return rows, idx


CHINESE_NUMERAL = "一二三四五六七八九十百千万"
_HEADING_PATTERNS = [
    (1, re.compile(rf"^第\s*[{CHINESE_NUMERAL}]+\s*[章节篇部]")),
    (1, re.compile(rf"^[{CHINESE_NUMERAL}]+\s*[、．.]")),
    (1, re.compile(r"^[0-9]+\s*[、．.]")),
    (1, re.compile(r"^[IVXLC]+\s*[、．.]", re.IGNORECASE)),
    (2, re.compile(rf"^（\s*[{CHINESE_NUMERAL}]+\s*）")),
    (2, re.compile(rf"^\(\s*[{CHINESE_NUMERAL}]+\s*\)")),
    (2, re.compile(r"^[0-9]+(\.[0-9]+)+\s")),
    (3, re.compile(r"^（\s*[0-9]+\s*）")),
    (3, re.compile(r"^\(\s*[0-9]+\s*\)")),
]


def _detect_heading_level(text: str, runs: List[Dict[str, Any]], allow_heading_heuristics: bool) -> Optional[int]:
    if not allow_heading_heuristics:
        return None
    stripped = text.strip()
    if not stripped:
        return None
    if any(run.get("charStyleRef") for run in runs):
        return None
    if len(stripped) > 80:
        return None
    if "\n" in stripped:
        return None
    for level, pattern in _HEADING_PATTERNS:
        if pattern.match(stripped):
            return level
    return None
