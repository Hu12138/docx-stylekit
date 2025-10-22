简要写明：

安装：pip install -e .

用法：

docx-stylekit observe examples/sample.docx -o observed.yaml

docx-stylekit merge examples/enterprise_baseline.yaml observed.yaml -o merged.yaml

docx-stylekit diff examples/enterprise_baseline.yaml observed.yaml

docx-stylekit/
├─ pyproject.toml
├─ README.md
├─ LICENSE
├─ .gitignore
├─ examples/
│  ├─ enterprise_baseline.yaml        # 你的企业基线模板（来自我们上一版）
│  └─ sample.docx                     # 外部模板样例（自备）
├─ src/
│  └─ docx_stylekit/
│     ├─ __init__.py
│     ├─ cli.py                       # 命令行入口：observe / merge / diff / validate
│     ├─ constants.py                 # 命名空间、字号映射等常量
│     ├─ utils/
│     │  ├─ units.py                  # twips/pt/cm 转换、字号名映射
│     │  ├─ xml.py                    # XML 辅助：读/写/查找、属性读取
│     │  └─ io.py                     # 文件读写、YAML load/dump、安全解压
│     ├─ io/
│     │  ├─ docx_zip.py               # 读 docx zip & 取出各 XML 部件
│     │  └─ rels.py                   # 关系解析（rels）
│     ├─ parsers/
│     │  ├─ theme.py                  # 解析 theme1.xml → 主题色/字体
│     │  ├─ styles.py                 # 解析 styles.xml → 段落/字符/表格样式
│     │  ├─ numbering.py              # 解析 numbering.xml → 多级编号绑定
│     │  ├─ document.py               # 解析 document.xml → 分节/页面/表格摘要/直改统计
│     │  └─ headers_footers.py        # 解析 header*/footer* → 页码域、样式
│     ├─ model/
│     │  ├─ observed.py               # “观测数据”的Python结构 & 合并前标准化
│     │  └─ schema.py                 # YAML 映射路径与模式（轻量）
│     ├─ emit/
│     │  ├─ observed_yaml.py          # 把观测数据转成 observed.yaml（结构与基线一致）
│     │  └─ report.py                 # 生成 diff 报告（文本/JSON）
│     ├─ merge/
│     │  └─ merger.py                 # enterprise_baseline.yaml + observed.yaml → merged.yaml
│     └─ diff/
│        └─ differ.py                 # 结构化 diff，供 report.py 使用
└─ tests/
   ├─ test_observe_min.py
   ├─ test_merge_min.py
   └─ assets/
      └─ small.docx
