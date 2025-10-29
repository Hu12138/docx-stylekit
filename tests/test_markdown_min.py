from pathlib import Path
import subprocess
import sys


def test_markdown_to_docx(tmp_path):
    md_path = tmp_path / "sample.md"
    md_path.write_text("# 测试标题\n\n这是第一段内容。\n\n- 列表项一\n- 列表项二", encoding="utf-8")
    out_docx = tmp_path / "output.docx"
    cmd = [
        sys.executable,
        "-m",
        "docx_stylekit.cli",
        "markdown",
        str(md_path),
        "-o",
        str(out_docx),
    ]
    subprocess.run(cmd, check=True)
    assert out_docx.exists()
    assert out_docx.stat().st_size > 0


def test_markdown_table_conversion():
    from docx_stylekit.convert.markdown import markdown_to_template

    md = """| A | B |
| --- | --- |
| 1 | 2 |
"""
    data = markdown_to_template(md)
    blocks = data["doc"]["blocks"]
    assert blocks[0]["type"] == "table"
    assert len(blocks[0]["header"]) == 1
    assert len(blocks[0]["rows"]) == 1


def test_markdown_normalize_literal_newlines():
    from docx_stylekit.convert.markdown import markdown_to_template

    md = "第一段\\n\\n| A | B |\\n| --- | --- |\\n| 1 | 2 |"
    data = markdown_to_template(md)
    blocks = data["doc"]["blocks"]
    assert blocks[0]["type"] == "paragraph"
    assert blocks[1]["type"] == "table"


def test_headline_detection():
    from docx_stylekit.convert.markdown import markdown_to_template

    md = "一、企业基本情况\n\n（ 一 ）公司概况\n\n（1）小节内容\n\n正文。"
    data = markdown_to_template(md)
    blocks = data["doc"]["blocks"]
    assert blocks[0]["type"] == "heading"
    assert blocks[0]["level"] == 1
    assert blocks[1]["type"] == "heading"
    assert blocks[1]["level"] == 2


def test_headline_heuristic_disabled_when_markdown_headings_present():
    from docx_stylekit.convert.markdown import markdown_to_template

    md = "# 标题\n\n正文段落包含 一、这样的标记但不应被识别。"
    data = markdown_to_template(md)
    # 第二个块应保持段落
    assert data["doc"]["blocks"][1]["type"] == "paragraph"
