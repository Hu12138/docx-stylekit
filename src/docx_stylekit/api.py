from __future__ import annotations

import json
import tempfile
import yaml
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from importlib import resources

from .convert.markdown import markdown_to_template
from .diff.differ import dict_diff
from .merge.merger import merge_enterprise_with_observed
from .render.json_template import expand_document
from .writer.docx_writer import render_to_docx
from .utils.io import load_yaml
from .model.observed import create_observed_skeleton
from .io.docx_zip import DocxZip
from .io.rels import parse_document_rels
from .parsers.theme import parse_theme
from .parsers.styles import parse_styles
from .parsers.numbering import parse_numbering
from .parsers.document import parse_sections
from .parsers.headers_footers import detect_page_field
from .emit.observed_yaml import emit_observed_yaml
from .tools.image_paragraphs import fix_image_paragraph_spacing


BytesLike = Union[bytes, bytearray, memoryview]
PathLike = Union[str, Path]
YamlLike = Union[Dict[str, Any], PathLike, BytesLike]
JsonLike = Union[Dict[str, Any], PathLike, BytesLike]


def _ensure_path(path: Optional[PathLike]) -> Optional[Path]:
    if path is None or isinstance(path, Path):
        return path
    return Path(path)


def _load_yaml_any(source: YamlLike) -> Dict[str, Any]:
    if isinstance(source, dict):
        return source
    if isinstance(source, (bytes, bytearray, memoryview)):
        return yaml.safe_load(bytes(source).decode("utf-8"))
    return load_yaml(source)


def _load_json_any(source: JsonLike) -> Dict[str, Any]:
    if isinstance(source, dict):
        return source
    if isinstance(source, (bytes, bytearray, memoryview)):
        return json.loads(bytes(source))
    with open(source, "r", encoding="utf-8") as f:
        return json.load(f)


def _write_output(data: Dict[str, Any], output: Optional[PathLike], *, as_yaml: bool = True) -> Optional[Path]:
    if output is None:
        return None
    path = _ensure_path(output)
    if not path:
        return None
    if as_yaml:
        emit_observed_yaml(data, path)
    else:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    return path


def observe_docx(docx: PathLike, *, output: Optional[PathLike] = None) -> Dict[str, Any]:
    dz = DocxZip(docx)
    parts = dz.parts()
    observed = create_observed_skeleton()

    if dz.has(parts["theme"]):
        observed["theme"] = parse_theme(dz.read_xml(parts["theme"]))
    if dz.has(parts["styles"]):
        observed["styles"] = parse_styles(dz.read_xml(parts["styles"]))
    if dz.has(parts["numbering"]):
        observed["numbering"] = parse_numbering(dz.read_xml(parts["numbering"]))
    if dz.has(parts["document"]):
        observed["page_setup"] = parse_sections(dz.read_xml(parts["document"]))

    headers = {}
    for hp in dz.list_headers():
        headers[hp] = detect_page_field(dz.read_xml(hp))
    footers = {}
    for fp in dz.list_footers():
        footers[fp] = detect_page_field(dz.read_xml(fp))
    observed["headers_footers"] = {"headers": headers, "footers": footers}

    if dz.has(parts["doc_rels"]):
        observed["rels_document"] = parse_document_rels(dz.read_xml(parts["doc_rels"]))

    dz.close()
    _write_output(observed, output, as_yaml=True)
    return observed


def merge_yaml(
    enterprise: YamlLike,
    observed: YamlLike,
    *,
    output: Optional[PathLike] = None,
) -> Dict[str, Any]:
    ent = _load_yaml_any(enterprise)
    obs = _load_yaml_any(observed)
    merged = merge_enterprise_with_observed(ent, obs)
    _write_output(merged, output, as_yaml=True)
    return merged


def diff_yaml(left: YamlLike, right: YamlLike) -> List[Dict[str, Any]]:
    a = _load_yaml_any(left)
    b = _load_yaml_any(right)
    return dict_diff(a, b)


def render_from_json(
    template: JsonLike,
    *,
    template_docx: Optional[PathLike] = None,
    styles_yaml: Optional[YamlLike] = None,
    output_path: Optional[PathLike] = None,
    prefer_json_styles: bool = False,
    fail_on_unknown_style: bool = True,
    keep_template_content: bool = False,
    return_bytes: bool = False,
) -> Union[Path, bytes]:
    data = _load_json_any(template)
    prepared = expand_document(_merge_with_default(data) if template_docx is None else data)
    styles_resolved: Optional[Dict[str, Any]] = None
    if styles_yaml:
        styles_resolved = _load_yaml_any(styles_yaml)
    tmp_template_path: Optional[Path] = None
    if template_docx is not None and isinstance(template_docx, (bytes, bytearray, memoryview)):
        tmp_template = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tmp_template.write(bytes(template_docx))
        tmp_template.flush()
        tmp_template.close()
        tmp_template_path = Path(tmp_template.name)
        template_path = tmp_template_path
    else:
        template_path = _ensure_path(template_docx)

    tmp_output_path: Optional[Path] = None
    if return_bytes:
        tmp_output = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tmp_output.close()
        output_path = Path(tmp_output.name)
        tmp_output_path = output_path
    else:
        if output_path is None:
            raise ValueError("output_path is required when return_bytes is False")
        output_path = _ensure_path(output_path)
        if output_path is None:
            raise ValueError("output_path is required when return_bytes is False")
        output_path.parent.mkdir(parents=True, exist_ok=True)

    render_to_docx(
        prepared,
        template_docx_path=template_path,
        styles_yaml=styles_resolved,
        output_path=output_path,
        prefer_json_styles=prefer_json_styles,
        fail_on_unknown_style=fail_on_unknown_style,
        clear_existing_content=not keep_template_content,
    )
    try:
        if return_bytes:
            assert tmp_output_path is not None
            data_bytes = tmp_output_path.read_bytes()
            tmp_output_path.unlink(missing_ok=True)
            return data_bytes
        assert output_path is not None
        return output_path
    finally:
        if tmp_template_path:
            tmp_template_path.unlink(missing_ok=True)


def render_from_markdown(
    markdown: Union[str, PathLike, BytesLike],
    *,
    template_docx: Optional[PathLike] = None,
    styles_yaml: Optional[YamlLike] = None,
    output_path: Optional[PathLike] = None,
    prefer_json_styles: bool = False,
    fail_on_unknown_style: bool = True,
    keep_template_content: bool = False,
    return_bytes: bool = False,
    title: Optional[str] = None,
) -> Union[Path, bytes]:
    if isinstance(markdown, (bytes, bytearray, memoryview)):
        text = bytes(markdown).decode("utf-8")
    elif isinstance(markdown, str):
        candidate = Path(markdown)
        if candidate.exists():
            text = candidate.read_text(encoding="utf-8")
        else:
            text = markdown
    else:
        text = Path(markdown).read_text(encoding="utf-8")
    json_template = markdown_to_template(text, title=title)
    return render_from_json(
        json_template,
        template_docx=template_docx,
        styles_yaml=styles_yaml,
        output_path=output_path,
        prefer_json_styles=prefer_json_styles,
        fail_on_unknown_style=fail_on_unknown_style,
        keep_template_content=keep_template_content,
        return_bytes=return_bytes,
    )


def _merge_with_default(data: Dict[str, Any]) -> Dict[str, Any]:
    from docx_stylekit.utils.dicts import deep_merge
    resource_path = resources.files("docx_stylekit.data").joinpath("default_render_template.yaml")
    with resources.as_file(resource_path) as path:
        default_profile = load_yaml(path)
    return deep_merge(default_profile, data)


def fix_image_paragraphs(
    docx: PathLike,
    *,
    output_path: Optional[PathLike] = None,
) -> Path:
    input_path = _ensure_path(docx)
    if input_path is None:
        raise ValueError("input path is required")
    destination = _ensure_path(output_path)
    return fix_image_paragraph_spacing(input_path, destination)
