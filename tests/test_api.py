from pathlib import Path
from docx import Document

from docx_stylekit.api import (
    observe_docx,
    merge_yaml,
    diff_yaml,
    render_from_markdown,
    render_from_json,
)


def test_observe_docx(tmp_path):
    sample_docx = tmp_path / "sample.docx"
    doc = Document()
    doc.add_paragraph("API observe test")
    doc.save(sample_docx)
    output = tmp_path / "observed.yaml"
    data = observe_docx(sample_docx, output=output)
    assert "styles" in data
    assert output.exists()
    assert output.stat().st_size > 0


def test_merge_and_diff(tmp_path):
    enterprise = Path("examples/enterprise_baseline.yaml")
    observed = Path("sampleObserved.yaml")
    merged = merge_yaml(enterprise, observed)
    assert "styles_observed" in merged
    diffs = diff_yaml(enterprise, observed)
    assert isinstance(diffs, list)
    assert diffs


def test_render_markdown_to_bytes():
    content = "# 一级标题\n\n正文段落。"
    doc_bytes = render_from_markdown(content, return_bytes=True)
    assert isinstance(doc_bytes, (bytes, bytearray))
    assert doc_bytes[:2] == b"PK"  # DOCX zip header

def test_render_markdown_from_path(tmp_path):
    md_path = tmp_path / "sample.md"
    md_path.write_text("# 目录\n\n内容段落。", encoding="utf-8")
    out_path = tmp_path / "sample.docx"
    render_from_markdown(str(md_path), output_path=out_path)
    assert out_path.exists()
    assert out_path.stat().st_size > 0


def test_render_json_to_path(tmp_path):
    template_json = {
        "doc": {
            "blocks": [
                {
                    "type": "paragraph",
                    "styleRef": "Normal",
                    "runs": [{"text": "Hello from API"}],
                }
            ]
        }
    }
    out_path = tmp_path / "api_render.docx"
    result_path = render_from_json(template_json, output_path=out_path)
    assert result_path == out_path
    assert out_path.exists()
    assert out_path.stat().st_size > 0
