# dsh-excel-kit

**A read-only Excel analysis toolkit for dsh (DeepSeek Harness).**

One sentence of value: stream any xlsx, however large, into compact, useful analysis (describe / filter / pivot) — without ever loading the whole workbook into memory, and without ever writing or formatting a cell.

> **Version lock warning**: this plugin is built and verified against **DSH `0.1.0-rc.6`**. Do **NOT** `npm install dsh-excel-kit@latest` blindly — use the locked range `>=0.1.0-rc.6 <0.2.0` (or whatever the host you run declares). The dsh plugin ABI may change between versions.

---

## Features

1. **Streaming, big-file safe** — a home-grown xlsx reader built on `yauzl` (streaming zip) + `sax` (streaming XML). No `XLSX.readFile`, no whole-workbook load, no OOM on 100 MB+ files.
2. **Three focused read tools** — `excel_describe`, `excel_filter`, `excel_pivot`. They return **compact JSON** (aggregates, samples, matched rows) — never a firehose of raw rows.
3. **Spill integration** — oversized tool results are persisted through `ctx.spillStore` and replaced in the value with a compact summary + `spilled` locator.
4. **Date-aware typing** — `numFmtId` from `xl/styles.xml` is used to convert Excel serial dates into ISO strings (`2024-01-01` / `2024-01-01T10:30:00.000Z`).
5. **Cancellable & concurrency-safe** — every long scan forwards `exec.signal` (AbortSignal); read-only tools declare `isConcurrencySafe`.

---

## How it works

The plugin is a normal dsh plugin: entry exports `name` / `inject` / `apply` and registers tools onto `ctx.tools`. Under the hood it opens the xlsx as a zip, streams only the entries it needs, and parses rows one at a time.

### Capability seams: dsh ⇄ this plugin

| dsh capability | How this plugin adapts it |
| --- | --- |
| `ctx.tools.register` (dsh-tools) | Registers `excel_describe` / `excel_filter` / `excel_pivot` via a `defineTool` shape: `parameters` are **MCP-style top-level JSON Schema** (`{ type: 'object', properties, required }`), `output.schema` declares the `{ content, structuredContent }` wrapper, `output.render` projects `content`, `isConcurrencySafe`, and async `execute(args, exec)` returns `{ content: ContentBlock[], structuredContent: <result> }`. |
| `exec.signal` (AbortSignal) | Forwarded into the streaming reader (`streamSheet`, sharedStrings load); aborts stop decompression/parsing promptly. |
| `ctx.spillStore` (dsh-spill) | Results larger than the threshold (default 32 KB) are persisted via `ctx.spillStore.saveText({owner, source, suggestedName, content})`; the returned value becomes `{ spilled: { locator, bytes, retrievalHint }, tool, summary }`. A same-interface in-memory fallback is used when no real spillStore exists. |
| `dsh.bundle` in `package.json` | Declares `cordis.patch.yml` so `dsh plugin --profile <name> add dsh-excel-kit` injects the `excel-kit` entry into the profile config. |
| `ctx.excel` (potential future service) | The `XlsxStreamReader` is exported so a future provider service can reuse the same streaming engine cross-plugin. (Not yet registered as a service.) |

### The streaming reader

- `yauzl` opens the file with `lazyEntries` and reads the central directory — no whole-file decompression.
- `xl/sharedStrings.xml` (when present) is parsed into a **chunked, spilled string array**: strings accumulate into a chunk, and when a chunk crosses 8 MB it is flushed to a temp file; random access goes through a small LRU cache of on-disk chunks. Memory stays bounded even for huge shared-string tables.
- `xl/worksheets/sheetN.xml` is parsed with `sax` and rows are assembled **one `<row>` at a time**; cells handle `inlineStr`, `s` (SST index), numeric (incl. `E` scientific notation), `t="str"`, `t="b"`, `t="d"` and empty cells.
- `xl/styles.xml` `numFmtId` is used to detect date columns (built-in 14–22 / 45–47, and custom formats containing `y`/`m`/`d`); date cells are emitted as ISO strings.
- Row data is handed to a callback and never accumulated into a full-table array.

---

## Tools at a glance

| Tool | Purpose | Key parameters | Returns |
| --- | --- | --- | --- |
| `excel_describe` | Profile a workbook/sheet | `file_path` (req), `sheet`, `max_rows`, `sample` (default 3) | sheets, rows, cols, per-column type distribution, non-empty count, empty rate, numeric min/max/mean, samples |
| `excel_filter` | Filter rows by conditions | `file_path` (req), `sheet`, `conditions` (req), `columns`, `limit` (default 100, cap 500) | matched/returned, projected rows |
| `excel_pivot` | Group-and-aggregate | `file_path` (req), `sheet`, `rows` (req), `values` (req), `limit` (default 50) | groups, nested keys like `dept|grade`, aggregated values |

Condition operators: `eq | ne | gt | gte | lt | lte | contains | in | between`.

Aggregation operators: `count | sum | mean | min | max`.

> **Return contract**: every tool's `execute` resolves to `{ content: [{type:'text', text}], structuredContent: <result> }` (dsh-tools rc.6 registry shape). The JSON payloads below are the `structuredContent` values; `content` is the human-readable rendering shown by the model. When a result exceeds the spill threshold (32 KB default), `structuredContent` becomes `{ spilled: { locator, bytes, retrievalHint }, tool, summary }`.

---

## Install

Requires a dsh installation (`0.1.0-rc.6` compatible). From the plugin's package directory:

```bash
# 1) inside dsh, add the plugin to a profile
dsh plugin --profile <name> add dsh-excel-kit

# 2) or install manually then apply via dsh.bundle patch
npm install dsh-excel-kit
```

The `dsh.bundle.patch` (`cordis.patch.yml`) inserts:

```yaml
- insert:
    - id: excel-kit
      name: 'dsh-excel-kit'
```

**Do not install `@latest` blindly** — pin to the compatible range (see the warning at the top).

---

## Configuration

The plugin follows the dsh `name` / `inject` / `apply` contract and ships **no plugin-level `Config`** (a plain, non-schemastery `Config` object breaks the cordis loader, which calls `Config.validate()`). Behavior is tuned by constants in source and documented in this README:

| Tuning knob | Default | Meaning |
| --- | --- | --- |
| `spillThreshold` | `32 * 1024` (32 KB) | Result serialization threshold (bytes) above which `ctx.spillStore` is used. |
| filter `limit` | `100` (hard cap `500`) | Max rows returned by `excel_filter`. |
| pivot `limit` | `50` | Max groups returned by `excel_pivot`. |

Per-profile overrides belong in the profile's `cordis.patch.yml` layer (e.g. disabling the plugin entry), not in a plugin `Config`.

---

## Usage examples

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

Compact result (truncated):

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

Result shape:

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

Result shape:

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

## Verification record

| Check | Command | Assertion target | Result |
| --- | --- | --- | --- |
| Build | `npm run build` | `tsc` strict passes | **PASS** (fixed 2026-08-18): `npx tsc --noEmit` 0 errors; `noEmitOnError:true` added to tsconfig to prevent emitting on error |
| Unit | `npm test` | 20+ assertions: date conversion, condition matching, describe/filter/pivot correctness, SST/date/empty cells, SheetJS cross-check | **PASS** 19/19 (tsx --test, ~9.3s) |
| Fixtures | `npm run gen:fixtures` | small.xlsx ~1 MB + large.xlsx ~100 MB generated streaming | **PASS** small.xlsx 1,725,588 B (~1.65 MiB); large.xlsx 139,715,025 B (~133.2 MiB) |
| Integration | `npm run test:integration` | 1 MB vs 100 MB timing; peak RSS delta `< 800 MB`; three tools usable on 100 MB; compact JSON results | **PASS** 6/6. describe large 26.9s / RSS +153.2 MB / JSON 2215 B; filter large 20.9s / +63.6 MB (matched 63,228); pivot large 21.8s / +36.2 MB; SST-chunk spill case 34.3s / +5.9 MB |

Numbers recorded 2026-08-18.

---

## Known limitations

- **Read-only by design**: v0.1.0 strictly forbids writing or formatting. No `set_cell`, no style mutation.
- SheetJS CE is a dependency for test cross-validation only; the production reader never calls `XLSX.readFile`.
- `excel_filter` and `excel_pivot` reference columns by **header name** (first row). Empty header cells fall back to `colN` (1-based).
- `mean`/`sum`/`min`/`max` operate on numeric values (numbers or numeric strings); non-numeric cells are ignored for numeric aggs. `count` counts non-empty values of the column within a group.
- Group keys in `excel_pivot` join values with `|`; a value containing `|` may collide with a nested key (rare in practice).
- Only the first sheet's `xlsx` (`.xlsx`, zip-based) format is supported — not legacy `.xls`, `.xlsb`, or password-encrypted files.
- Streams are bounded, but `excel_describe` with default settings scans every row (that's the point — it is streaming, not free).

---

## License

MIT — see [LICENSE](./LICENSE).
