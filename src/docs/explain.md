下面是**新版 JSON 模版标准的字段解释文档**（面向使用 LLM 生成 JSON、以及渲染器实现）。它与我们现有的 YAML 样式库**可协同工作**：渲染时**先查 YAML**，再用 JSON 的 `stylesInline` 做**增补/覆盖（可控）**，并支持 `pageTemplates` 实现封面/摘要/参考文献等**页面级**结构。

---

# 总览

* 根：`{ "doc": { ... } }`
* 关键能力：

  * **样式来源双通道**：`styleCatalog`（白名单）、YAML 样式库（外部）、`stylesInline`（JSON 内联样式）
  * **页面模板**：`pageTemplates` + `blocks[].useTemplate`
  * **内容结构**：`blocks`（段落/标题/表格/列表/目录/分页/条件/循环/变量）
  * **版式**：`pageSetup`（全局纸张/边距/页码）、`pageTemplates.*.layout`（分节级设置）
  * **编号**：`numbering`（标题多级编号绑定）
  * **页眉页脚**：`headersFooters`
  * **变量**：`variables`
  * **渲染策略**：`renderPolicy`
  * **默认模版**：当 CLI `render` 未指定 `--template` 时，会自动合并 `docx_stylekit.data/default_render_template.yaml`，提供基础页面设置、表格样式与段落样式。

---

# 解析与优先级（非常重要）

1. `styleRef` 解析顺序：
   **YAML 样式库**（外部） → **JSON `stylesInline`** → 否则错误/回退（按策略）。
2. 覆盖规则：

   * 默认：**YAML 优先**。
   * 仅当 `stylesInline[样式名]."$override": true` 或 CLI 开关 `--prefer-json-styles` 时，允许 JSON **覆盖同名样式的字段**（建议只覆盖必要字段，如字号/对齐）。
3. 分节版式：

   * 全局：`pageSetup`；
   * 局部（分节）：`pageTemplates.*.layout` 在调用时生效（`blocks[].useTemplate`）。
4. TOC（目录）：

   * `doc.toc` 控制是否需要目录；
   * `blocks` 中也可插入 `{"type":"toc"}`；Word 打开后需更新域。

---

# 字段详解

## 1. `doc.meta`（文档元数据）

| 字段          | 类型                | 必填 | 说明             |
| ----------- | ----------------- | -- | -------------- |
| `title`     | string            | ✓  | 文档标题           |
| `lang`      | string            | ✓  | 语言代码，如 `zh-CN` |
| `version`   | string            | ✓  | 模版版本           |
| `createdAt` | string(date-time) | ✓  | 创建时间（ISO 8601） |

> 仅用于记录与审计，不影响排版。

---

## 2. `doc.styleCatalog`（样式白名单）

| 字段          | 类型       | 必填 | 说明                                              |
| ----------- | -------- | -- | ----------------------------------------------- |
| `paragraph` | string[] | ✓  | 允许使用的段落样式名集合（如 `["Normal","Heading 1","正文文字"]`） |
| `character` | string[] | ✓  | 允许使用的字符样式名集合                                    |
| `table`     | string[] | ✓  | 允许使用的表格样式名集合                                    |

> 渲染前可用它做“风格门禁”。不定义样式细节，只列可用名称。

---

## 3. `doc.stylesInline`（JSON 内联样式库，可新增/可覆盖）

* 用于**新增** YAML 中没有的样式（如 `CoverTitle`、`ReferenceItem`），或在需要时**受控覆盖**同名 YAML 样式的字段。

**样式对象 `StyleDef`**

| 字段          | 类型                                    | 必填 | 说明                                                  |
| ----------- | ------------------------------------- | -- | --------------------------------------------------- |
| `type`      | enum(`paragraph`,`character`,`table`) | ✓  | 样式类型                                                |
| `basedOn`   | string                                |    | 继承的样式名（可指向 YAML 或 `stylesInline` 已定义的样式，如 `Normal`） |
| `font`      | `FontProps`                           |    | 字体属性（见下）                                            |
| `paragraph` | `ParagraphProps`                      |    | 段落属性（见下，`type=paragraph`时适用）                        |
| `numbering` | `NumberingLink`                       |    | 多级编号绑定（常用于标题）                                       |
| `$override` | boolean                               |    | 允许覆盖同名 YAML 样式字段（默认 false）                          |

**`FontProps`**

* `eastAsia`（中文/东亚字体），`ascii`（西文字体），`sizePt`（pt），`bold`，`italic`，`color`（#RRGGBB）

**`ParagraphProps`（常用）**

* `align`：`left/center/right/both`
* `lineSpacingMultiple`（倍数行距）、`lineExactPt`（固定值 pt）、`lineAtLeastPt`（最小值 pt）
* `spaceBeforePt`、`spaceAfterPt`（段前/段后 pt）
* `firstLineChars`（首行缩进 **字符数**）、`hangingChars`（悬挂缩进字符数）
* `leftIndentCm`、`rightIndentCm`（左右缩进 cm）
* `outlineLevel`（0..9；标题级别）、`keepNext`（与下段同页）

**`NumberingLink`**

* `link`：编号方案名（可对应 YAML 中的多级编号方案），
* `level`：级别（0..8，对应 Heading1..Heading9）

> 建议：将**封面/参考文献**等样式定义在这里，命名清晰（如 `CoverTitle`、`CoverAuthor`、`ReferenceItem`）。

---

## 4. `doc.pageTemplates`（页面级模板）

用于定义“封面/摘要/参考文献”等**分节模板**。在 `blocks` 中用 `{"useTemplate":"cover_page"}` 调用。

**结构 `PageTemplate`**

| 字段       | 类型        | 必填 | 说明                         |
| -------- | --------- | -- | -------------------------- |
| `layout` | `Layout`  |    | 分节级版式设置（可为空使用全局）           |
| `blocks` | `Block[]` | ✓  | 模板内的内容块（同 `doc.blocks` 规范） |

**`Layout` 常用字段**

* `marginsCm.{top,bottom,left,right[,header,footer]}`：页边距（cm）
* `orientation`：`portrait/landscape`
* `verticalAlign`：`top/center/both/bottom`（页面垂直对齐；封面常用 `center`）
* `pageNumbering`：`none/roman/arabic`（本节页码方案）
* `titleFirstPageDifferent`、`evenOddDifferent`：首页不同/奇偶页不同（布尔）
* `startAt`：本节页码起始

> 调用模板时会**新起一个 Section**，应用 `layout` 再渲染其 `blocks`。

---

## 5. `doc.pageSetup`（全局页面设置）

| 字段                                                | 类型      | 必填 | 说明                                      |
| ------------------------------------------------- | ------- | -- | --------------------------------------- |
| `paper`                                           | string  | ✓  | 纸张规格（如 `A4`）                            |
| `orientation`                                     | enum    | ✓  | 全局横/纵向                                  |
| `marginsCm.{top,bottom,left,right,header,footer}` | number  | ✓  | 全局边距（cm）                                |
| `titleFirstPageDifferent`                         | boolean | ✓  | 封面/首页不同页眉页脚                             |
| `evenOddDifferent`                                | boolean | ✓  | 奇偶页不同页眉页脚                               |
| `pageNumbering.{position,startAt,format}`         | object  | ✓  | 页码位置（header/footer）、起始号、格式（如 `decimal`） |

> 分节模板中的 `layout` 会**覆盖**全局设置（仅作用于该节）。

---

## 6. `doc.numbering`（标题多级编号绑定）

| 字段             | 类型      | 必填 | 说明                                                                  |
| -------------- | ------- | -- | ------------------------------------------------------------------- |
| `bindHeadings` | boolean | ✓  | 是否将编号绑定到标题样式                                                        |
| `preset`       | string  | ✓  | 预设方案名（例如 `decimal-dot`）                                             |
| `levels`       | object  | ✓  | `{"1":{"styleRef":"Heading 1"}, "2":{"styleRef":"Heading 2"}, ...}` |

> 渲染器依据 YAML/模板中的 numbering 或创建临时编号；JSON `stylesInline[].numbering` 也可单独声明绑定。

---

## 7. `doc.toc`（目录）

| 字段         | 类型      | 必填 | 说明                     |
| ---------- | ------- | -- | ---------------------- |
| `required` | boolean | ✓  | 是否需要目录                 |
| `levels`   | int[]   | ✓  | 目录包含的标题级别（如 `[1,2,3]`） |
| `note`     | string  |    | 备注（如“打开后更新目录域”）        |

> 也可以在 `blocks` 中放置 `{ "type": "toc" }` 来决定目录插入位置。

---

## 8. `doc.headersFooters`（页眉页脚）

数组元素 `HeaderFooterItem`：

| 字段         | 类型                            | 必填 | 说明                                   |
| ---------- | ----------------------------- | -- | ------------------------------------ |
| `type`     | enum(`text`,`pageNumber`)     | ✓  | 内容类型                                 |
| `styleRef` | string                        | ✓  | 样式引用（段落或字符样式）                        |
| `align`    | enum(`left`,`center`,`right`) | ✓  | 对齐                                   |
| `text`     | string                        |    | 文本（当 `type=text`）                    |
| `pattern`  | string                        |    | 页码域格式（当 `type=pageNumber`，常用 `PAGE`） |

> 若模板 DOCX 已有页码域，渲染器可选择**不重复插入**。

---

## 9. `doc.variables`（变量）

* `{"KEY": "value"}`，供文本替换 `{KEY}` 使用。
* `repeat/conditional` 中也可引用嵌套变量（如 `{item.NAME}`）。

---

## 10. `doc.blocks`（内容块数组）

**通用字段**

* `type`：`paragraph | heading | list | table | figure | caption | pageBreak | conditional | repeat | variable | toc`
* `styleRef`：样式名（遵循“YAML → stylesInline”的解析顺序）
* `pageBreakBefore`：布尔；在该块前强制分页

**常见块类型**

1. `paragraph`

```json
{ "type":"paragraph", "styleRef":"正文文字", "runs":[{"text":"段落文本","charStyleRef":"Emphasis"}] }
```

* `runs[].text`（可含 `{VAR}`）、`runs[].charStyleRef`（字符样式）

2. `heading`

```json
{ "type":"heading", "level":2, "styleRef":"Heading 2", "text":"研究方法" }
```

* 不要在 `text` 中手写编号；编号由 `numbering` 控制

3. `list`

```json
{ "type":"list", "ordered": true, "styleRef":"Normal", "items":[ { "runs":[{"text":"项1"}] } ] }
```

* 简化处理：外观依赖样式；如需精准 `numPr` 可按项目扩展

4. `table`

```json
{
  "type":"table", "styleRef":"TableGrid",
  "columns":[{"widthPct":40,"align":"center"}, {"widthPct":60,"align":"left"}],
  "header":[ [ { "blocks":[{ "type":"paragraph", "styleRef":"Normal", "runs":[{"text":"列1"}] } ] },
               { "blocks":[{ "type":"paragraph", "styleRef":"Normal", "runs":[{"text":"列2"}] } ] } ] ],
  "rows":[ [ { "blocks":[{ "type":"paragraph", "styleRef":"Normal", "runs":[{"text":"A"}] } ] },
             { "blocks":[{ "type":"paragraph", "styleRef":"Normal", "runs":[{"text":"B"}] } ] } ] ],
  "format": {
    "header": {
      "fill": "#DCE6F1",
      "color": "#1F3864",
      "bold": true,
      "verticalAlign": "center"
    },
    "cell": {
      "verticalAlign": "center"
    },
    "tableBorder": {
      "top": { "style": "single", "color": "#305496", "size": 6 },
      "insideH": { "style": "single", "color": "#9BAED9", "size": 4 }
    }
  }
}
```

* `columns[].widthPct` 会自动映射为列宽（结合页面可用宽度）；未声明时使用 Word 默认列宽，如提供 `renderDefaults.table.columns` 则按该配置推断。
* `format` 允许快速定义表头底色、交替行配色、垂直对齐与边框；若未指定，则回退到 `renderDefaults.table.format`（如存在）。

5. `caption`

```json
{ "type":"caption", "styleRef":"Caption", "text":"表 1 统计结果" }
```

6. `pageBreak`

```json
{ "type":"pageBreak" }
```

7. `toc`

```json
{ "type":"toc" }
```

8. `variable`（语法糖：等价于一个段落只有一段文本的场景）

```json
{ "type":"variable", "styleRef":"Normal", "text":"{COMPANY}" }
```

9. `conditional`（条件块）

```json
{ "type":"conditional", "if":"HAS_RISK", "then":[ ... ], "else":[ ... ] }
```

10. `repeat`（循环块）

```json
{ "type":"repeat", "for":"SHAREHOLDERS", "as":"item", "template":[ ...blocks... ] }
```

11. `useTemplate`（调用页面模板）

```json
{ "useTemplate":"cover_page", "variables":{ "TITLE":"XX大学论文", "AUTHOR":"张三" } }
```

* 渲染时创建新节 → 应用 `pageTemplates.cover_page.layout` → 渲染其 `blocks`

---

## 11. `doc.renderDefaults`（内置渲染默认值）

| 字段              | 类型    | 说明                                                                 |
| ----------------- | ----- | -------------------------------------------------------------------- |
| `table.styleRef`  | string | 未显式指定 `styleRef` 的表格使用的默认表格样式（如 `InfoTable`）                  |
| `table.columns`   | array  | 默认列宽/对齐定义（`widthPct` + `align`）。表格未声明 `columns` 时按此推断。           |
| `table.format`    | object | 表格视觉风格（见下）；可通过块内 `format` 覆盖指定字段。                                  |

**`table.format` 常用键**

| 字段             | 说明                                                                                       |
| ---------------- | ------------------------------------------------------------------------------------------ |
| `header.fill`    | 表头底色，十六进制色值（如 `#DCE6F1`）                                                               |
| `header.color`   | 表头文字颜色                                                                               |
| `header.bold`    | 表头文字是否加粗                                                                             |
| `header.verticalAlign` | 表头单元格垂直对齐：`top/center/middle/bottom/both`                                          |
| `header.border.*` | 表头底部/左右边框样式，`style/color/size`                                                     |
| `alternate.fill` | 奇偶行底色；启用 `bandedRows=true` 时生效                                                     |
| `alternate.verticalAlign` | 交替行垂直对齐设置                                                                    |
| `cell.verticalAlign` | 普通数据单元格垂直对齐（未指定时使用 Word 默认）                                             |
| `tableBorder.*`  | 整体表格边框线配置（包括 `top/bottom/left/right/insideH/insideV`）                              |

> CLI 未指定 `--template` 时，会自动合并 `default_render_template.yaml`，该文件示例化了蓝灰配色的 InfoTable 表格样式与宋体/Times New Roman 组合的段落样式。

---

## 12. `doc.constraints`（生成约束，渲染前可校验）

| 字段                           | 类型       | 作用                     |
| ---------------------------- | -------- | ---------------------- |
| `forbidDirectFormatting`     | boolean  | 禁止“直改”指令（如在文本里描述字号/颜色） |
| `allowedParagraphStyles`     | string[] | 允许的段落样式集合              |
| `allowedCharacterStyles`     | string[] | 允许的字符样式集合              |
| `maxHeadingLevel`            | int      | 最大标题级别（如 3）            |
| `noManualNumbersForHeadings` | boolean  | 标题文本中不得手写编号            |

---

## 13. `doc.renderPolicy`（可选的渲染策略位）

| 字段                   | 类型      | 默认    | 说明                                              |
| -------------------- | ------- | ----- | ----------------------------------------------- |
| `preferJsonStyles`   | boolean | false | JSON 样式允许覆盖 YAML（等同 CLI `--prefer-json-styles`） |
| `failOnUnknownStyle` | boolean | true  | 找不到样式时立即失败（false 时回退到 `Normal` 并记日志）            |

---

# 与 Word/OOXML 的对应关系（速查）

| JSON 字段                     | OOXML/Word 对应            | 备注                |
| --------------------------- | ------------------------ | ----------------- |
| `stylesInline[].font`       | `w:rPr`                  | 字体/字号/加粗/颜色       |
| `stylesInline[].paragraph`  | `w:pPr`                  | 对齐/行距/段前后/缩进/大纲级别 |
| `stylesInline[].numbering`  | `w:numPr`                | 绑定抽象编号/级别         |
| `pageSetup`                 | `w:sectPr`(文档默认节)        | 纸张/边距/页码方案        |
| `pageTemplates.*.layout`    | `w:sectPr`(新节)           | 分节级覆盖             |
| `headersFooters`            | header/footer + `PAGE` 域 | 结构简化版             |
| `toc`/`blocks[].type="toc"` | `TOC` 域                  | Word 内更新域         |

---
