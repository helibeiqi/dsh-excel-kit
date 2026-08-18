# dsh-excel-kit

**dsh（DeepSeek Harness）下"最稳的 Excel 只读分析工具"。**

一句话价值：无论多大的 xlsx，都能以流式方式产出紧凑、可用的分析结果（describe / filter / pivot）——既不把整个工作簿读进内存，也绝不写入或格式化任何一个单元格。

> **版本锁定警告**：本插件基于 **DSH `0.1.0-rc.6`** 构建并验证。请**不要盲目** `npm install dsh-excel-kit@latest`——请锁定兼容区间 `>=0.1.0-rc.6 <0.2.0`（或与你运行宿主声明一致的版本）。dsh 插件 ABI 在不同版本间可能变化。

---

## 核心特性（3–5 条）

1. **流式、大文件零崩溃** —— 基于 `yauzl`（流式解压 zip）+ `sax`（流式 XML）自研 xlsx 读取器，**不调用 `XLSX.readFile`**，不做全表加载，100MB+ 文件不会 OOM。
2. **三个专注的只读工具** —— `excel_describe` / `excel_filter` / `excel_pivot`，返回**紧凑 JSON**（聚合、示例值、命中行），绝不把原始行倾泻给模型。
3. **Spill 集成** —— 超大工具结果经 `ctx.spillStore` 持久化，value 中只保留紧凑摘要 + `spilled` 定位信息。
4. **日期类型识别** —— 依据 `xl/styles.xml` 的 `numFmtId`，把 Excel 序列日期转成 ISO 字符串（`2024-01-01` / `2024-01-01T10:30:00.000Z`）。
5. **可取消且并发安全** —— 每次长扫描都转发 `exec.signal`（AbortSignal）；只读工具声明 `isConcurrencySafe`。

---

## 工作原理

本插件是一个标准 dsh 插件：入口导出 `name` / `inject` / `apply`，向 `ctx.tools` 注册工具。底层把 xlsx 当作 zip 打开，只流式读取需要的条目，逐行组装解析。

### 能力接缝：dsh 能力 ⇄ 本插件适配

| dsh 能力 | 本插件如何适配 |
| --- | --- |
| `ctx.tools.register`（dsh-tools） | 以 `defineTool` 形态注册 `excel_describe` / `excel_filter` / `excel_pivot`：`parameters` 为 **MCP 风格顶层 JSON Schema**（`{ type: 'object', properties, required }`），`output.schema` 声明 `{ content, structuredContent }` 包装，`output.render` 投影 `content`，`isConcurrencySafe`，异步 `execute(args, exec)` 返回 `{ content: ContentBlock[], structuredContent: <结果> }`。 |
| `exec.signal`（AbortSignal） | 转发进流式读取器（`streamSheet`、sharedStrings 加载）；中止即停解压/解析。 |
| `ctx.spillStore`（dsh-spill） | 结果超过阈值（默认 32KB）时经 `ctx.spillStore.saveText({owner, source, suggestedName, content})` 落盘；返回值变为 `{ spilled: { locator, bytes, retrievalHint }, tool, summary }`。无真实 spillStore 时使用同接口的内存回退。 |
| `package.json` 的 `dsh.bundle` | 声明 `cordis.patch.yml`，`dsh plugin --profile <name> add dsh-excel-kit` 时把 `excel-kit` 条目注入 profile 配置。 |
| `ctx.excel`（潜在扩展服务） | 导出 `XlsxStreamReader`，未来可把同一流式引擎以 provider 服务形式跨插件复用（当前尚未注册为服务）。 |

### 流式读取器（护城河）

- `yauzl` 以 `lazyEntries` 打开文件并读取中央目录——不做全文件解压。
- `xl/sharedStrings.xml`（如存在）解析进**分块 spill 字符串数组**：字符串累积到 8MB 一块即落盘临时文件，随机访问走小规模 LRU 缓存。超大共享字符串表内存仍然有界。
- `xl/worksheets/sheetN.xml` 用 `sax` 解析，**逐 `<row>`** 组装对象；单元格覆盖 `inlineStr`、`s`（SST 索引）、数值（含 `E` 科学计数法）、`t="str"`、`t="b"`、`t="d"` 与空单元格。
- `xl/styles.xml` 的 `numFmtId` 识别日期列（内置 14–22 / 45–47，以及含 `y`/`m`/`d` 的自定义格式）；日期单元格输出 ISO 字符串。
- 行数据经回调逐行消费，绝不累积全表数组。

---

## 工具一览

| 工具 | 用途 | 关键参数 | 返回 |
| --- | --- | --- | --- |
| `excel_describe` | 剖析工作簿/sheet | `file_path`(必填)、`sheet`、`max_rows`、`sample`(默认3) | sheets、行数列数、每列类型分布、非空计数、空值率、数值 min/max/mean、示例值 |
| `excel_filter` | 按条件过滤行 | `file_path`(必填)、`sheet`、`conditions`(必填)、`columns`、`limit`(默认100，上限500) | matched/returned、投影后的行 |
| `excel_pivot` | 分组聚合 | `file_path`(必填)、`sheet`、`rows`(必填)、`values`(必填)、`limit`(默认50) | 分组、嵌套键如 `dept|grade`、聚合值 |

条件操作符：`eq | ne | gt | gte | lt | lte | contains | in | between`。

聚合操作符：`count | sum | mean | min | max`。

> **返回契约**：每个工具的 `execute` 解析为 `{ content: [{type:'text', text}], structuredContent: <结果> }`（dsh-tools rc.6 registry 形态）。下方 JSON 载荷即 `structuredContent`；`content` 是给模型看的人类可读渲染。结果超过 spill 阈值（默认 32KB）时，`structuredContent` 变为 `{ spilled: { locator, bytes, retrievalHint }, tool, summary }`。

---

## 安装

需要已安装 dsh（`0.1.0-rc.6` 兼容）。在插件包目录内：

```bash
# 1) 在 dsh 内把插件加入 profile
dsh plugin --profile <name> add dsh-excel-kit

# 2) 或手动安装后通过 dsh.bundle patch 生效
npm install dsh-excel-kit
```

`dsh.bundle.patch`（`cordis.patch.yml`）会插入：

```yaml
- insert:
    - id: excel-kit
      name: 'dsh-excel-kit'
```

**不要盲目装 `@latest`**——请锁定兼容区间（见文首警告）。

---

## 配置说明

插件遵循 dsh 的 `name` / `inject` / `apply` 契约，**不提供插件级 `Config`**（非 schemastery 的普通 Config 对象会让 cordis loader 调用 `Config.validate()` 而崩溃）。行为由源码内常量控制，本文档列出：

| 调优项 | 默认 | 说明 |
| --- | --- | --- |
| `spillThreshold` | `32 * 1024`（32KB） | 结果序列化字节数超过该阈值时启用 `ctx.spillStore`。 |
| filter `limit` | `100`（硬上限 `500`） | `excel_filter` 最多返回的行数。 |
| pivot `limit` | `50` | `excel_pivot` 最多返回的分组数。 |

按 profile 的覆盖需求应写在 profile 的 `cordis.patch.yml` 层（如禁用插件条目），而不是插件 `Config`。

---

## 使用示例

### 1) excel_describe

```json
{
  "tool": "excel_describe",
  "arguments": {
    "file_path": "/data/reports/2026-08-sales.xlsx",
    "sheet": "Sheet1",
    "sample": 3
  }
}
```

紧凑结果（节选）：

```json
{
  "file_path": "/data/reports/2026-08-sales.xlsx",
  "sheet": "Sheet1",
  "sheets": ["Sheet1", "Sheet2"],
  "total_rows": 10001,
  "total_cols": 6,
  "columns": [
    { "col": 0, "header": "id", "total": 10001, "non_empty": 10000, "empty_rate": 0.0001,
      "types": { "number": 10000 }, "numeric": { "min": 1, "max": 10000, "mean": 5000.5, "sum": 50005000 },
      "samples": [1, 2, 3] },
    { "col": 4, "header": "order_date", "total": 10001, "non_empty": 10000, "empty_rate": 0.0001,
      "types": { "date": 10000 }, "samples": ["2026-08-01", "2026-08-02", "2026-08-03"] }
  ]
}
```

### 2) excel_filter

```json
{
  "tool": "excel_filter",
  "arguments": {
    "file_path": "/data/reports/2026-08-sales.xlsx",
    "conditions": [
      { "column": "region", "op": "in", "values": ["华东", "华南"] },
      { "column": "amount", "op": "gte", "value": 5000 }
    ],
    "columns": ["id", "region", "amount"],
    "limit": 100
  }
}
```

结果结构：

```json
{
  "file_path": "/data/reports/2026-08-sales.xlsx",
  "sheet": "Sheet1",
  "columns": ["id", "region", "amount"],
  "matched": 823,
  "returned": 100,
  "truncated": true,
  "rows": [
    { "row": 4, "values": { "id": 3, "region": "华东", "amount": 12800 } }
  ]
}
```

### 3) excel_pivot

```json
{
  "tool": "excel_pivot",
  "arguments": {
    "file_path": "/data/reports/2026-08-sales.xlsx",
    "rows": ["region", "channel"],
    "values": [
      { "column": "amount", "agg": "sum" },
      { "column": "amount", "agg": "mean" },
      { "column": "id", "agg": "count" }
    ]
  }
}
```

结果结构：

```json
{
  "file_path": "/data/reports/2026-08-sales.xlsx",
  "sheet": "Sheet1",
  "rows": ["region", "channel"],
  "values": [{ "column": "amount", "agg": "sum" }, { "column": "amount", "agg": "mean" }, { "column": "id", "agg": "count" }],
  "groups": 8,
  "truncated": false,
  "data": [
    { "key": "华东|线上", "group": { "region": "华东", "channel": "线上" },
      "values": { "sum:amount": 1234567.89, "mean:amount": 5234.56, "count:id": 236 } }
  ]
}
```

---

## 验证记录

| 检查项 | 命令 | 断言目标 | 结果 |
| --- | --- | --- | --- |
| 构建 | `npm run build` | `tsc` strict 通过 | **PASS**（2026-08-18 已修复）：`npx tsc --noEmit` 0 错误；tsconfig 已加 `noEmitOnError:true` 防止带错 emit |
| 单元测试 | `npm test` | 20+ 断言：日期转换、条件匹配、describe/filter/pivot 正确性、SST/日期/空单元格、SheetJS 交叉验证 | **PASS** 19/19（tsx --test，约 9.3s） |
| 生成 fixtures | `npm run gen:fixtures` | small.xlsx ~1MB + large.xlsx ~100MB 流式生成 | **PASS** small.xlsx 1,725,588 B（约 1.65 MiB）；large.xlsx 139,715,025 B（约 133.2 MiB） |
| 集成测试 | `npm run test:integration` | 1MB vs 100MB 耗时对比；峰值 RSS 增量 `< 800MB`；三工具在 100MB 上可用；返回紧凑 JSON | **PASS** 6/6。describe large 26.9s / RSS +153.2MB / JSON 2215B；filter large 20.9s / +63.6MB（matched 63,228）；pivot large 21.8s / +36.2MB；SST 分块 spill 用例 34.3s / +5.9MB |

数字由 QA（2026-08-18）独立复验后记录。

---

## 已知限制

- **设计上只读**：v0.1.0 严格禁止写入/格式化，没有 `set_cell`，不做样式修改。
- SheetJS CE 仅用于测试交叉验证；生产读取路径从不调用 `XLSX.readFile`。
- `excel_filter` / `excel_pivot` 按**表头名**（第一行）引用列；空表头回退为 `colN`（1 起）。
- `mean`/`sum`/`min`/`max` 只作用于数值（数字或数字字符串）；非数值单元格在数值聚合中忽略。`count` 统计组内该列的非空值个数。
- `excel_pivot` 的分组键用 `|` 拼接；值本身含 `|` 时可能与嵌套键冲突（实际场景罕见）。
- 仅支持 zip 格式的 `.xlsx`；不支持旧 `.xls`、`.xlsb`、加密文件。
- 流有界，但 `excel_describe` 默认会扫全部行——这正是它的用途（流式而非免费）。

---

## License

MIT — 见 [LICENSE](./LICENSE)。
