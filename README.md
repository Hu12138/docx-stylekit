# docx-stylekit

> 版本：0.2.0

`docx-stylekit` 通过解析 OOXML，把 Word 模板抽取成结构化 YAML；同时支持把 YAML/JSON/Markdown 渲染回规范的 DOCX，适合构建企业文档标准化、模板校验等工具链。

## 安装

```bash
pip install docx-stylekit
```

开发环境可直接本地安装：

```bash
git clone https://github.com/your-org/docx-stylekit.git
cd docx-stylekit
pip install -e .
```

## 命令行使用

```bash
# 解析 DOCX → observed.yaml
docx-stylekit observe examples/sample.docx -o observed.yaml

# 合并企业基线与观测结果
docx-stylekit merge examples/enterprise_baseline.yaml observed.yaml -o merged.yaml

# Diff 两份 YAML
docx-stylekit diff examples/enterprise_baseline.yaml observed.yaml

# Markdown → DOCX（可选 --template / --styles）
docx-stylekit markdown doc/测试用例.md -o doc/output.docx

# 调整图片段落（行距/对齐/缩进）
docx-stylekit fix-images doc/测试用例.docx -o doc/测试用例_图片优化.docx
```

`docx-stylekit markdown` 默认会应用内置模板的样式（标题、页码等已设为中文规范）；若需要企业模板，可加 `-t 企业模板.docx`。

## Python API

CLI 所有能力均可通过代码调用，方便集成 Flask/FastAPI：

```python
from docx_stylekit import (
    observe_docx,
    merge_yaml,
    diff_yaml,
    render_from_markdown,
    fix_image_paragraphs,
)

# 解析 DOCX
observed = observe_docx("examples/sample.docx")

# 合并基线
merged = merge_yaml("examples/enterprise_baseline.yaml", observed)

# Diff
diffs = diff_yaml("examples/enterprise_baseline.yaml", observed)

# Markdown → DOCX（返回字节流，可直接用于 HTTP Response）
docx_bytes = render_from_markdown("# 标题\n\n内容", return_bytes=True)

# 规范图片段落（单倍行距、居中、零缩进）
fix_image_paragraphs("报告初稿.docx", output_path="报告初稿_图片调整.docx")
```

更多 API 说明见 `src/docx_stylekit/api.py`，包括传入/输出 `bytes`、模板样式覆盖等选项。

## 仓库结构

```
docx-stylekit/
├─ src/docx_stylekit/
│  ├─ api.py                 # 对外公开的 Python API
│  ├─ cli.py                 # 命令行入口
│  ├─ convert/               # Markdown → JSON 模板
│  ├─ parsers/               # 解析 theme/styles/document 等
│  ├─ writer/                # 基于 python-docx 写回 DOCX
│  ├─ data/default_render_template.yaml  # 内置样式与页码设置
│  └─ ...
├─ tests/                    # 单元测试（pytest）
└─ examples/                 # 示例 YAML 与 DOCX
```

## 开发

```bash
pip install -e ".[dev]"
python -m pytest
```

欢迎根据业务场景扩展校验规则、渲染模板或对接 Web 服务。*** End Patch
