/**
 * dsh-excel-kit — excel_pivot 工具
 *
 * 按 rows 分组列聚合，返回紧凑 JSON。多分组列用嵌套键（如 "dept|grade"）。
 */
import * as fs from 'fs';
import { XlsxStreamReader } from '../stream/xlsx-stream-reader';
import { AggOp, CellValue, PivotValueSpec } from '../stream/types';
import { accumAgg, finalizeAgg, initAggState, isNumericValue, AggState } from '../utils';

export const PIVOT_LIMIT_HARD_CAP = 500;

export interface PivotArgs {
  file_path: string;
  sheet?: string;
  rows: string[];
  values: PivotValueSpec[];
  limit?: number;
}

export const pivotParameters = {
  type: 'object',
  properties: {
    file_path: {
      type: 'string',
      description: 'xlsx 文件绝对路径',
    },
    sheet: {
      type: 'string',
      description: 'sheet 名称，缺省为第一个 sheet',
    },
    rows: {
      type: 'array',
      description: '分组列表头名，可多个；多分组列用嵌套键如 "dept|grade"',
      items: { type: 'string' },
    },
    values: {
      type: 'array',
      description: '聚合配置 [{column, agg: count|sum|mean|min|max}]',
      items: {
        type: 'object',
        properties: {
          column: { type: 'string' },
          agg: { type: 'string', enum: ['count', 'sum', 'mean', 'min', 'max'] },
        },
        required: ['column', 'agg'],
      },
    },
    limit: {
      type: 'number',
      description: `最大分组数，默认 50，硬上限 ${PIVOT_LIMIT_HARD_CAP}`,
    },
  },
  required: ['file_path', 'rows', 'values'],
} as const;

export interface PivotResult {
  file_path: string;
  sheet: string;
  rows: string[];
  values: PivotValueSpec[];
  groups: number;
  truncated: boolean;
  data: {
    key: string;
    group: Record<string, CellValue>;
    values: Record<string, number | null>;
  }[];
}

function buildHeaderIndex(header: CellValue[]): Map<string, number> {
  const map = new Map<string, number>();
  header.forEach((v, idx) => {
    const key = v == null || v === '' ? `col${idx + 1}` : String(v);
    if (!map.has(key)) map.set(key, idx);
  });
  return map;
}

export async function executePivot(
  args: PivotArgs,
  signal?: AbortSignal,
): Promise<PivotResult> {
  if (!args.file_path) throw new Error('file_path is required');
  if (!fs.existsSync(args.file_path)) throw new Error(`File not found: ${args.file_path}`);
  const rowsCols: string[] = Array.isArray(args.rows) ? args.rows : [];
  const valueSpecs: PivotValueSpec[] = Array.isArray(args.values) ? args.values : [];
  if (rowsCols.length === 0) throw new Error('rows must be a non-empty array');
  if (valueSpecs.length === 0) throw new Error('values must be a non-empty array');
  const validAggs = new Set<AggOp>(['count', 'sum', 'mean', 'min', 'max']);
  for (const v of valueSpecs) {
    if (!validAggs.has(v.agg)) throw new Error(`Invalid agg: ${v.agg}`);
  }
  const limit = Math.max(1, Math.min(args.limit ?? 50, PIVOT_LIMIT_HARD_CAP));

  const reader = new XlsxStreamReader(args.file_path);
  try {
    const sheets = await reader.getSheetNames();
    if (sheets.length === 0) throw new Error('Workbook has no sheets');
    const sheet = args.sheet ?? sheets[0];
    if (!sheets.includes(sheet)) {
      throw new Error(`Sheet "${sheet}" not found. Available: ${sheets.join(', ')}`);
    }

    const headerRow = await reader.getHeaderRow(sheet, signal);
    const headerValues: CellValue[] = [];
    for (const c of headerRow?.cells ?? []) headerValues[c.col] = c.value;
    const headerMap = buildHeaderIndex(headerValues);

    const rowIdxs = rowsCols.map((name) => {
      const idx = headerMap.get(name);
      if (idx == null) {
        throw new Error(`Row column "${name}" not found. Available: ${[...headerMap.keys()].join(', ')}`);
      }
      return idx;
    });
    const valueColIdxs = valueSpecs.map((v) => {
      const idx = headerMap.get(v.column);
      if (idx == null) {
        throw new Error(`Value column "${v.column}" not found. Available: ${[...headerMap.keys()].join(', ')}`);
      }
      return idx;
    });

    // 分组聚合状态：key → (group 值数组, agg states)
    const groups = new Map<string, { keyVals: CellValue[]; states: AggState[] }>();
    let truncated = false;
    let firstRow = true;

    await reader.streamSheet(sheet, {
      signal,
      rowHandler: (row) => {
        // 第一行为表头，不参与分组
        if (firstRow) {
          firstRow = false;
          return true;
        }
        const vals: CellValue[] = [];
        for (const c of row.cells) vals[c.col] = c.value;
        // 分组键：全部空则跳过
        const keyVals = rowIdxs.map((i) => vals[i] ?? null);
        if (keyVals.every((v) => v == null || v === '')) return true;
        const key = keyVals.map((v) => (v == null || v === '' ? '' : String(v))).join('|');
        let entry = groups.get(key);
        if (!entry) {
          if (groups.size >= limit) {
            truncated = true;
            return true; // 超限分组不再纳入，但继续扫以保持流程一致
          }
          entry = { keyVals, states: valueSpecs.map(() => initAggState()) };
          groups.set(key, entry);
        }
        valueSpecs.forEach((spec, vi) => {
          const v = vals[valueColIdxs[vi]];
          const numeric = spec.agg !== 'count' ? isNumericValue(v) : true;
          accumAgg(entry!.states[vi], v, numeric);
        });
        return true;
      },
    });

    const data = [...groups.entries()].map(([key, entry]) => {
      const group: Record<string, CellValue> = {};
      rowsCols.forEach((name, i) => {
        group[name] = entry.keyVals[i];
      });
      const aggValues: Record<string, number | null> = {};
      valueSpecs.forEach((spec, vi) => {
        aggValues[`${spec.agg}:${spec.column}`] = finalizeAgg(entry.states[vi], spec.agg);
      });
      return { key, group, values: aggValues };
    });

    return {
      file_path: args.file_path,
      sheet,
      rows: rowsCols,
      values: valueSpecs,
      groups: data.length,
      truncated,
      data,
    };
  } finally {
    await reader.dispose();
  }
}

export function formatPivot(r: PivotResult): string {
  const lines: string[] = [];
  lines.push(`File: ${r.file_path}`);
  lines.push(`Sheet: ${r.sheet}   groups=${r.groups}${r.truncated ? ' (truncated)' : ''}`);
  lines.push(`rows=[${r.rows.join(', ')}]  values=[${r.values.map((v) => `${v.agg}:${v.column}`).join(', ')}]`);
  lines.push('---');
  for (const d of r.data) {
    lines.push(`${d.key}: ${JSON.stringify(d.values)}`);
  }
  return lines.join('\n');
}
