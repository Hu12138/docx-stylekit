import json
from pathlib import Path

import click
import yaml
from importlib import resources
from colorama import Fore, Style
from .io.docx_zip import DocxZip
from .io.rels import parse_document_rels
from .parsers.theme import parse_theme
from .parsers.styles import parse_styles
from .parsers.numbering import parse_numbering
from .parsers.document import parse_sections
from .parsers.headers_footers import detect_page_field
from .model.observed import create_observed_skeleton
from .emit.observed_yaml import emit_observed_yaml
from .merge.merger import merge_enterprise_with_observed
from .diff.differ import dict_diff
from .emit.report import print_diff_report
from .utils.io import load_yaml, dump_yaml
from .utils.dicts import deep_merge
from .convert.markdown import markdown_to_template
from .render.json_template import expand_document
from .writer.docx_writer import render_to_docx


def _load_default_render_profile(template_docx_path):
    if template_docx_path:
        return {}
    try:
        default_profile_path = resources.files("docx_stylekit.data").joinpath("default_render_template.yaml")
        with resources.as_file(default_profile_path) as p:
            return load_yaml(p)
    except (FileNotFoundError, ModuleNotFoundError):
        return {}


def _prepare_render_input(data, template_docx_path):
    default_profile = _load_default_render_profile(template_docx_path)
    if default_profile:
        data = deep_merge(default_profile, data)
    return data


def _load_styles_yaml(styles):
    if not styles:
        return None
    if isinstance(styles, dict):
        return styles
    if isinstance(styles, str) and styles.endswith((".yml", ".yaml")) and Path(styles).exists():
        with open(styles, "r", encoding="utf-8") as yf:
            return yaml.safe_load(yf)
    return styles


def _render_from_data(
    data,
    template_docx_path,
    styles,
    output,
    prefer_json_styles,
    fail_on_unknown_style,
    keep_template_content,
):
    prepared = _prepare_render_input(data, template_docx_path)
    expanded = expand_document(prepared)
    styles_yaml = _load_styles_yaml(styles)
    render_to_docx(
        expanded,
        template_docx_path=template_docx_path,
        styles_yaml=styles_yaml,
        output_path=output,
        prefer_json_styles=prefer_json_styles,
        fail_on_unknown_style=fail_on_unknown_style,
        clear_existing_content=not keep_template_content,
    )

@click.group()
def main():
    """docx-stylekit: Observe → Merge → Diff → Validate"""

@main.command()
@click.argument("docx_path", type=click.Path(exists=True))
@click.option("-o", "--output", default="observed.yaml", help="Output YAML path.")
def observe(docx_path, output):
    """从DOCX解析样式/编号/页面设置，生成 observed.yaml"""
    dz = DocxZip(docx_path)
    parts = dz.parts()
    observed = create_observed_skeleton()

    # theme
    if dz.has(parts["theme"]):
        observed["theme"] = parse_theme(dz.read_xml(parts["theme"]))

    # styles
    if dz.has(parts["styles"]):
        observed["styles"] = parse_styles(dz.read_xml(parts["styles"]))

    # numbering
    if dz.has(parts["numbering"]):
        observed["numbering"] = parse_numbering(dz.read_xml(parts["numbering"]))

    # sections
    if dz.has(parts["document"]):
        observed["page_setup"] = parse_sections(dz.read_xml(parts["document"]))

    # headers/footers
    headers = {}
    for hp in dz.list_headers():
        info = detect_page_field(dz.read_xml(hp))
        headers[hp] = info
    footers = {}
    for fp in dz.list_footers():
        info = detect_page_field(dz.read_xml(fp))
        footers[fp] = info
    observed["headers_footers"] = {"headers": headers, "footers": footers}

    # rels (可选：用于 header/footer 关联校验)
    if dz.has(parts["doc_rels"]):
        _rels = parse_document_rels(dz.read_xml(parts["doc_rels"]))
        observed["rels_document"] = _rels

    dz.close()
    emit_observed_yaml(observed, output)
    click.echo(Fore.GREEN + f"observed.yaml generated at: {output}" + Style.RESET_ALL)

@main.command()
@click.argument("enterprise_yaml", type=click.Path(exists=True))
@click.argument("observed_yaml", type=click.Path(exists=True))
@click.option("-o", "--output", default="merged.yaml", help="Merged YAML path.")
def merge(enterprise_yaml, observed_yaml, output):
    """合并企业基线与观测到的模板，产出 merged.yaml"""
    ent = load_yaml(enterprise_yaml)
    obs = load_yaml(observed_yaml)
    merged = merge_enterprise_with_observed(ent, obs)
    dump_yaml(merged, output)
    click.echo(Fore.GREEN + f"merged.yaml generated at: {output}" + Style.RESET_ALL)

@main.command()
@click.argument("left_yaml", type=click.Path(exists=True))
@click.argument("right_yaml", type=click.Path(exists=True))
@click.option("--fmt", type=click.Choice(["text","json"]), default="text")
def diff(left_yaml, right_yaml, fmt):
    """对比两份 YAML（可用于企业基线 vs 观测）"""
    a = load_yaml(left_yaml)
    b = load_yaml(right_yaml)
    diffs = dict_diff(a, b)
    print_diff_report(diffs, fmt=fmt)

@main.command()
@click.argument("json_template", type=click.Path(exists=True))
@click.option("--template", "-t", type=click.Path(exists=True), required=False,
              help="样式模板 DOCX（包含企业样式/编号/页眉页脚）。若省略，则使用内置默认模板。")
@click.option("--styles", "-s", type=click.Path(exists=False), required=False,
              help="合并后的 YAML（merged.yaml）。可选，用于校验/对照。")
@click.option("-o", "--output", type=click.Path(), default="output.docx", help="输出 DOCX 路径")
@click.option("--prefer-json-styles/--no-prefer-json-styles", default=False, help="允许 JSON 覆盖同名 YAML 样式字段")
@click.option("--fail-on-unknown-style/--no-fail-on-unknown-style", default=True, help="未知样式是否直接失败（默认 true）")
@click.option("--keep-template-content/--wipe-template-content", default=False,
              help="是否保留模板 DOCX 原有正文内容（默认不保留，仅使用样式/布局）")
def render(json_template, template, styles, output, prefer_json_styles, fail_on_unknown_style, keep_template_content):
    """读取 JSON 模版（含内容+内联样式+页面模板），渲染为 DOCX"""
    with open(json_template, "r", encoding="utf-8") as f:
        data = json.load(f)
    _render_from_data(
        data,
        template_docx_path=template,
        styles=styles,
        output=output,
        prefer_json_styles=prefer_json_styles,
        fail_on_unknown_style=fail_on_unknown_style,
        keep_template_content=keep_template_content,
    )
    click.echo(Fore.GREEN + f"DOCX generated at: {output}" + Style.RESET_ALL)


@main.command()
@click.argument("markdown_path", type=click.Path(exists=True))
@click.option("--template", "-t", type=click.Path(exists=True), required=False,
              help="样式模板 DOCX（包含企业样式/编号/页眉页脚）。若省略，则使用内置默认模板。")
@click.option("--styles", "-s", type=click.Path(exists=False), required=False,
              help="合并后的 YAML（merged.yaml），供样式校验使用。")
@click.option("-o", "--output", type=click.Path(), default="output.docx", help="输出 DOCX 路径")
@click.option("--title", type=str, required=False, help="覆盖 Markdown 文档标题。")
@click.option("--prefer-json-styles/--no-prefer-json-styles", default=False, help="允许 JSON 覆盖同名 YAML 样式字段")
@click.option("--fail-on-unknown-style/--no-fail-on-unknown-style", default=True, help="未知样式是否直接失败（默认 true）")
@click.option("--keep-template-content/--wipe-template-content", default=False,
              help="是否保留模板 DOCX 原有正文内容（默认不保留，仅使用样式/布局）")
def markdown(markdown_path, template, styles, output, title, prefer_json_styles, fail_on_unknown_style, keep_template_content):
    """将 Markdown 文件转换为 DOCX（内部先转 JSON，再复用 render 流程）"""
    text = Path(markdown_path).read_text(encoding="utf-8")
    data = markdown_to_template(text, title=title)
    _render_from_data(
        data,
        template_docx_path=template,
        styles=styles,
        output=output,
        prefer_json_styles=prefer_json_styles,
        fail_on_unknown_style=fail_on_unknown_style,
        keep_template_content=keep_template_content,
    )
    click.echo(Fore.GREEN + f"DOCX generated at: {output}" + Style.RESET_ALL)
