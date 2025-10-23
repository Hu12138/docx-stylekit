| 字段                                      | 类型                 | 必需 | 说明                           |
| --------------------------------------- | ------------------ | -- | ---------------------------- |
| `doc`                                   | object             | 是  | 整个文档模版根对象                    |
| `doc.meta`                              | object             | 是  | 文档元信息                        |
| `doc.meta.title`                        | string             | 是  | 文档标题（用于命名/展示）                |
| `doc.meta.lang`                         | string             | 是  | 语言代码，如 `zh-CN`               |
| `doc.meta.version`                      | string             | 是  | 模版版本号                        |
| `doc.meta.createdAt`                    | string (date-time) | 是  | 模版创建时间 ISO 格式                |
| `doc.styleCatalog`                      | object             | 是  | 文档可用样式目录，供生成引擎校验             |
| `doc.styleCatalog.paragraph`            | array[string]      | 是  | 段落样式名称列表                     |
| `doc.styleCatalog.character`            | array[string]      | 是  | 字符样式名称列表                     |
| `doc.styleCatalog.table`                | array[string]      | 是  | 表格样式名称列表                     |
| `doc.pageSetup`                         | object             | 是  | 页面设置（纸张、边距、页码等）              |
| `doc.pageSetup.paper`                   | string             | 是  | 纸张型号，如 `"A4"`                |
| `doc.pageSetup.orientation`             | string             | 是  | `"portrait"` 或 `"landscape"` |
| `doc.pageSetup.marginsCm`               | object             | 是  | 各边距/页眉/页脚 尺寸（单位 cm）          |
| `doc.pageSetup.titleFirstPageDifferent` | boolean            | 是  | 是否封面首页与正文不同页眉页脚              |
| `doc.pageSetup.evenOddDifferent`        | boolean            | 是  | 是否奇偶页页眉/页脚不同                 |
| `doc.pageSetup.pageNumbering`           | object             | 是  | 页码设置                         |
| `doc.pageSetup.pageNumbering.position`  | string             | 是  | `"header"` 或 `"footer"`      |
| `doc.pageSetup.pageNumbering.startAt`   | integer            | 是  | 页码起始数字                       |
| `doc.pageSetup.pageNumbering.format`    | string             | 是  | 编号格式标识（如 `"decimal"`）        |
| `doc.numbering`                         | object             | 是  | 多级编号体系设置                     |
| `doc.numbering.bindHeadings`            | boolean            | 是  | 是否将编号绑定至 Heading 样式          |
| `doc.numbering.preset`                  | string             | 是  | 预设编号风格 ID                    |
| `doc.numbering.levels`                  | object             | 是  | 每级编号绑定到哪个样式                  |
| `doc.toc`                               | object             | 否  | 目录配置（可选）                     |
| `doc.headersFooters`                    | object             | 否  | 页眉页脚配置                       |
| `doc.variables`                         | object             | 否  | 模版变量键值对                      |
| `doc.blocks`                            | array[object]      | 是  | 文档内容块数组                      |
| `doc.constraints`                       | object             | 否  | 模版约束规则（生成后用于校验）              |
