下面这份“项目总结 + 提示词 + 下一步路线 + 测试样例设计”可直接放进 README 顶部，给团队和后续贡献者一眼看懂。

# 项目总结（docx-stylekit）

**docx-stylekit** 通过解包 `.docx`（ZIP+OOXML），解析 `theme/styles/numbering/document/headers-footers` 等 XML，**把任意 Word 模板的版式样式抽取为机器可读的 YAML（observed.yaml）**，并与**企业基线 YAML** 合并生成 **merged.yaml**，为后续“自动规范化/修复 docx”提供稳定输入。

已实现：

* CLI：`observe`（解析模板 → observed.yaml）、`merge`（基线+观测 → merged.yaml）、`diff`（两份 YAML 差异）。
* 解析器：主题色/字体、段落/字符/表格样式、多级编号、分节与页面设置、页眉页脚页码域的基础检测。
* 归一化：字号 pt↔中文字号名、行距（single/1.5/double/exact/atLeast）、缩进（字符/厘米）、边距（twips↔cm）。
* 工程骨架：可扩展的 parsers / merge / diff / emit 模块、requirements 与示例基线 YAML。

# 项目提示词（放 README 开头）

> **提示词（电梯稿）**
> 给我一个 DOCX 模板，我要得到一份**客观的样式/版式指纹（observed.yaml）**，再把它和公司的**统一规范（enterprise_baseline.yaml）**融合成**可复用的工程模板（merged.yaml）**。
> 我不直接改文档外观，而是**先抽象成 YAML 标准**，后续一切“修复/生成/验证”都以 YAML 为唯一真相。
> 核心目标：**样式优先、直改最少、统一编号、版式可审计、结果可回滚**。

# 接下来要做什么（路线图）

1. **Validate 子命令**

   * 读 merged.yaml + observed.yaml，按 `policy/validation` 执行规则校验（如“禁止 run 直改”，“必须绑定 Heading1..3 编号”，“封面首页不同”等），输出红黄绿报告。
2. **更智能的 Merge 策略**

   * 把 `observed.styles` 映射到企业 YAML 的 `styles.*`（default / options）而非仅存档；
   * 对“核心样式”（Normal/Heading1..3/TOC/Caption/TableGrid）提供**可配置优先级**（企业优先 or 模板优先），并生成差异注记。
3. **rels 一致性 & 目录域检查**

   * 校验 header/footer 与 `document.xml.rels` 的引用完整性；
   * 识别 TOC 指令（\o、\h、\z、\u），在校验里提示“需在 Word 内更新域”。
4. **Report 增强**

   * 生成 Markdown/HTML 报告：页面设置摘要、样式覆盖率、直改占比、编号绑定图、表格策略等。
5. **（可选）Writer 阶段**

   * 基于 merged.yaml 生成/修复 docx：写回 `styles.xml/numbering.xml/theme/settings/sectPr/header/footer`，实现“一键规范化”。

# 测试样例设计（测试集 + 期望 + 验收）

为保证解析与合并的稳定性，设计**最小可复现**的 8 份样例 DOCX（均放 `tests/assets/`），每个样例配一段**期望断言**（可写成 pytest，现先描述）。

| 样例文件                          | 目的       | 文档特征                            | 观测期望（observed.yaml 关键点）                                                    | 验收要点                   |
| ----------------------------- | -------- | ------------------------------- | -------------------------------------------------------------------------- | ---------------------- |
| `A_basic_theme.docx`          | 主题解析     | 含 theme1：accent1/major/minor 字体 | `theme.colors.accent1` 有十六进制；`fonts.major.latin/ea` 正确                     | 颜色/字体不为空               |
| `B_styles_heading.docx`       | 标题/正文样式  | Normal、Heading1..3 定义齐全         | `paragraph_styles.Heading1.pPr.outline_level=0`；字号映射到“三号/小三”等              | 最近邻字号名正确               |
| `C_numbering_binding.docx`    | 多级编号绑定   | numbering.xml 绑定 Heading1..3    | `numbering.abstract.*.levels[0..2].pStyle=HeadingX`                        | 3 级均绑定                 |
| `D_sections_margins.docx`     | 分节与边距    | 两个节：封面（首页不同）、正文（页码起始=1）         | 第一节 `titlePg=true`、第二节 `pgNumStart=1`；边距 cm 正确                             | twips→cm 转换正确          |
| `E_headers_footers_page.docx` | 页码域检测    | 页脚中 `PAGE` 域                    | `headers_footers.footers.*.has_page=true`                                  | 识别 fldSimple/instrText |
| `F_tables_policy.docx`        | 表格策略     | 表格样式统一，标题行重复、行不跨页               | `tables.header_row_repeat=true`、`row_cant_split=true`（在 observed 的表格统计中体现） | 标志位正确                  |
| `G_direct_formatting.docx`    | 直改检测     | 正文大量 run 直改（颜色、大小）              | 在后续 validate 中触发“直改占比”告警（现阶段记录为备注）                                         | 能统计 run 直改（占比/计数）      |
| `H_no_theme_hardcolor.docx`   | 无主题+硬编码色 | 样式中使用 `w:color val="FF0000"`    | `styles.*.rPr.color.hex="#FF0000"`                                         | 颜色以 hex 落盘             |

> 如暂不实现“直改统计”，先在 `observed.styles` 中保留 rPr/pPr 原样信息，待 validate 阶段补充统计逻辑。

**差异测试（diff）**

* 准备一份企业基线 `enterprise_baseline.yaml` 与 `A_basic_theme.docx` 的 observed 对比：

  * 若基线默认字体不同，应在 `diff` 输出 `styles.Normal.rPr.ascii` 的 changed 项；
  * 合并后（merge）不改变 default，仅把观测值进入 `*_observed` 或 options（视策略而定）。

**合并测试（merge）**

* 输入：`enterprise_baseline.yaml` + `B_styles_heading.observed.yaml`
* 期望：`merged.yaml` 中

  * 基线 default 保持不变；
  * `styles_observed.paragraph_styles.Heading1` 有观测信息；
  * `numbering_observed` 挂载原始抽取结构；
  * `page_setup_observed.sections[0].pgMar.left≈2.8`（随样例设置）。

**端到端冒烟**

```bash
docx-stylekit observe tests/assets/B_styles_heading.docx -o /tmp/obs.yaml
docx-stylekit merge examples/enterprise_baseline.yaml /tmp/obs.yaml -o /tmp/merged.yaml
docx-stylekit diff examples/enterprise_baseline.yaml /tmp/obs.yaml --fmt text
```

验证：三个命令均成功退出；`/tmp/obs.yaml` 与 `/tmp/merged.yaml` 体积>0 且包含主题/样式/编号/分节键。

# 验收标准（阶段性 Done 的定义）

* **解析覆盖率**：对上表 8 个样例，`observe` 能抽到对应关键字段（≥95% 覆盖）。
* **映射正确率**：字号、行距、边距换算误差 ≤ 1%。
* **幂等性**：相同输入执行两次，`observed.yaml`、`merged.yaml` 内容完全一致。
* **健壮性**：缺少 `numbering.xml` 或 `theme1.xml` 时，能优雅降级（字段缺省而非崩溃）。

---

需要的话，我可以把这套测试描述转成 **pytest 测试用例骨架**（生成临时 observed.yaml 后断言关键路径），以及一份**空白样例文档制作指引**（如何在 Word 里快速造出每种“带坑”特性）。
