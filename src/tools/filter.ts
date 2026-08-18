/**
 * dsh-excel-kit — excel_filter 工具
 *
 * 按条件过滤行，返回指定列、紧凑 JSON。limit 默认 100，硬上限 500。
 */
import * as fs from 'fs';
import { XlsxStreamReader } from '../stream/xlsx-stream-reader';
import { CellValue, FilterCondition, RowMap } from '../stream/types';
import { matchCondition } from '../utils';

export const FILTER_LIMIT_HARD_CAP = 500;

export interface FilterArgs {
  file_path: string;
  sheet?: string;
  conditions: FilterCondition[];
  columns?: string[];
  limit?: number;
}

export const filterParameters = {
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
    conditions: {
      type: 'array',
      description:
        '过滤条件数组，[{column, op: eq|ne|gt|gte|lt|lte|contains|in|between, value?, values?}]，column 为表头名',
      items: {
        type: 'object',
        properties: {
          column: { type: 'string' },
          op: {
            type: 'string',
            enum: ['eq', 'ne', 'gt', 'gte', 'lt', 'lte', 'contains', 'in', 'between'],
          },
          value: {},
          values: { type: 'array' },
        },
        required: ['column', 'op'],
      },
    },
    columns: {
      type: 'array',
      description: '仅返回这些表头列（缺省返回全部列）',
      items: { type: 'string' },
    },
    limit: {
      type: 'number',
      description: `最大返回行数，默认 100，硬上限 ${FILTER_LIMIT_HARD_CAP}`,
    },
  },
  required: ['file_path', 'conditions'],
} as const;

export interface FilterResult {
  file_path: string;
  sheet: string;
  columns: string[];
  matched: number;
  returned: number;
  truncated: boolean;
  rows: { row: number; values: RowMap }[];
}

/** 构建表头名（含空表头回退 colN）→ 列索引 */
function buildHeaderIndex(header: CellValue[]): Map<string, number> {
  const map = new Map<string, number>();
  header.forEach((v, idx) => {
    const key = v == null || v === '' ? `col${idx + 1}` : String(v);
    if (!map.has(key)) map.set(key, idx);
  });
  return map;
}

export async function executeFilter(
  args: FilterArgs,
  signal?: AbortSignal,
): Promise<FilterResult> {
  if (!args.file_path) throw new Error('file_path is required');
  if (!fs.existsSync(args.file_path)) throw new Error(`File not found: ${args.file_path}`);
  const conditions: FilterCondition[] = Array.isArray(args.conditions) ? args.conditions : [];
  if (conditions.length === 0) throw new Error('conditions must be a non-empty array');
  const limit = Math.max(1, Math.min(args.limit ?? 100, FILTER_LIMIT_HARD_CAP));

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

    // 校验条件列名存在
    for (const cond of conditions) {
      if (!headerMap.has(cond.column)) {
        throw new Error(
          `Condition column "${cond.column}" not found in header. Available: ${[...headerMap.keys()].join(', ')}`,
        );
      }
    }
    // 校验 op 合法性
    const validOps = new Set(['eq', 'ne', 'gt', 'gte', 'lt', 'lte', 'contains', 'in', 'between']);
    for (const cond of conditions) {
      if (!validOps.has(cond.op)) throw new Error(`Invalid op: ${cond.op}`);
    }

    // 投影列
    const projectCols: number[] = [];
    const projectNames: string[] = [];
    if (args.columns && args.columns.length > 0) {
      for (const name of args.columns) {
        const idx = headerMap.get(name);
        if (idx == null) {
          throw new Error(`Projection column "${name}" not found. Available: ${[...headerMap.keys()].join(', ')}`);
        }
        projectCols.push(idx);
        projectNames.push(name);
      }
    } else {
      projectCols.push(...headerValues.map((_, i) => i));
      projectNames.push(
        ...headerValues.map((v, i) => (v == null || v === '' ? `col${i + 1}` : String(v))),
      );
    }

    const rows: { row: number; values: RowMap }[] = [];
    let matched = 0;
    let firstRow = true;

    await reader.streamSheet(sheet, {
      signal,
      rowHandler: (row) => {
        // 第一行为表头，不参与匹配
        if (firstRow) {
          firstRow = false;
          return true;
        }
        const vals: CellValue[] = [];
        for (const c of row.cells) vals[c.col] = c.value;
        let ok = true;
        for (const cond of conditions) {
          const idx = headerMap.get(cond.column)!;
          if (!matchCondition(vals[idx], cond)) {
            ok = false;
            break;
          }
        }
        if (!ok) return true;
        matched++;
        if (rows.length >= limit) return true; // 继续全扫以获得准确 matched
        const values: RowMap = {};
        projectCols.forEach((ci, pi) => {
          values[projectNames[pi]] = vals[ci] ?? null;
        });
        rows.push({ row: row.row, values });
        return true;
      },
    });

    return {
      file_path: args.file_path,
      sheet,
      columns: projectNames,
      matched,
      returned: rows.length,
      truncated: matched > rows.length,
      rows,
    };
  } finally {
    await reader.dispose();
  }
}

export function formatFilter(r: FilterResult): string {
  const lines: string[] = [];
  lines.push(`File: ${r.file_path}`);
  lines.push(`Sheet: ${r.sheet}   matched=${r.matched} returned=${r.returned}${r.truncated ? ' (truncated)' : ''}`);
  lines.push('---');
  for (const row of r.rows) {
    lines.push(`row ${row.row}: ${JSON.stringify(row.values)}`);
  }
  return lines.join('\n');
}
