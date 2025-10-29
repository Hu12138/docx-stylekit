import click
from colorama import Fore, Style
from .emit.report import print_diff_report
from .api import (
    observe_docx,
    merge_yaml,
    diff_yaml,
    render_from_json,
    render_from_markdown,
)

@click.group()
def main():
    """docx-stylekit: Observe → Merge → Diff → Validate"""

@main.command()
@click.argument("docx_path", type=click.Path(exists=True))
@click.option("-o", "--output", default="observed.yaml", help="Output YAML path.")
def observe(docx_path, output):
    """从DOCX解析样式/编号/页面设置，生成 observed.yaml"""
    observe_docx(docx_path, output=output)
    click.echo(Fore.GREEN + f"observed.yaml generated at: {output}" + Style.RESET_ALL)

@main.command()
@click.argument("enterprise_yaml", type=click.Path(exists=True))
@click.argument("observed_yaml", type=click.Path(exists=True))
@click.option("-o", "--output", default="merged.yaml", help="Merged YAML path.")
def merge(enterprise_yaml, observed_yaml, output):
    """合并企业基线与观测到的模板，产出 merged.yaml"""
    merge_yaml(enterprise_yaml, observed_yaml, output=output)
    click.echo(Fore.GREEN + f"merged.yaml generated at: {output}" + Style.RESET_ALL)

@main.command()
@click.argument("left_yaml", type=click.Path(exists=True))
@click.argument("right_yaml", type=click.Path(exists=True))
@click.option("--fmt", type=click.Choice(["text","json"]), default="text")
def diff(left_yaml, right_yaml, fmt):
    """对比两份 YAML（可用于企业基线 vs 观测）"""
    diffs = diff_yaml(left_yaml, right_yaml)
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
    render_from_json(
        json_template,
        template_docx=template,
        styles_yaml=styles,
        output_path=output,
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
    render_from_markdown(
        markdown_path,
        template_docx=template,
        styles_yaml=styles,
        output_path=output,
        prefer_json_styles=prefer_json_styles,
        fail_on_unknown_style=fail_on_unknown_style,
        keep_template_content=keep_template_content,
        title=title,
    )
    click.echo(Fore.GREEN + f"DOCX generated at: {output}" + Style.RESET_ALL)


if __name__ == "__main__":
    main()
