/**
 * dsh-excel-kit — excel_describe 工具
 *
 * 返回紧凑 JSON：sheet 列表、行数、列数、每列类型分布/非空计数/空值率、
 * 数值列 min/max/mean、示例值。绝不返回海量原始行。
 */
import * as fs from 'fs';
import { XlsxStreamReader } from '../stream/xlsx-stream-reader';
import { CellValue, ColumnProfile, DescribeResult } from '../stream/types';
import {
  inferValueType,
  isNumericValue,
  formatValue,
} from '../utils';

export interface DescribeArgs {
  file_path: string;
  sheet?: string;
  max_rows?: number;
  sample?: number;
}

export const describeParameters = {
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
    max_rows: {
      type: 'number',
      description: '最多扫描行数，缺省全扫（仍为流式）',
    },
    sample: {
      type: 'number',
      description: '每列示例值条数，缺省 3',
    },
  },
  required: ['file_path'],
} as const;

interface ColAcc {
  total: number;
  nonEmpty: number;
  types: Record<string, number>;
  numeric: { sum: number; min: number; max: number; cnt: number } | null;
  samples: CellValue[];
}

export async function executeDescribe(
  args: DescribeArgs,
  signal?: AbortSignal,
): Promise<DescribeResult> {
  if (!args.file_path) throw new Error('file_path is required');
  if (!fs.existsSync(args.file_path)) {
    throw new Error(`File not found: ${args.file_path}`);
  }
  const sample = Math.max(0, Math.floor(args.sample ?? 3));
  const maxRows = args.max_rows != null ? Math.max(0, Math.floor(args.max_rows)) : undefined;

  const reader = new XlsxStreamReader(args.file_path);
  try {
    const sheets = await reader.getSheetNames();
    if (sheets.length === 0) throw new Error('Workbook has no sheets');
    const sheet = args.sheet ?? sheets[0];
    if (!sheets.includes(sheet)) {
      throw new Error(`Sheet "${sheet}" not found. Available: ${sheets.join(', ')}`);
    }

    const headerRow = await reader.getHeaderRow(sheet, signal);
    const headerCells = headerRow?.cells ?? [];
    const headerMap = new Map<number, string | null>();
    for (const c of headerCells) headerMap.set(c.col, c.value == null ? null : String(c.value));

    const accs: ColAcc[] = [];
    const bump = (col: number) => {
      while (accs.length <= col) {
        accs.push({
          total: 0,
          nonEmpty: 0,
          types: {},
          numeric: null,
          samples: [],
        });
      }
    };

    let rowsSeen = 0;
    let maxCol = 0;
    await reader.streamSheet(sheet, {
      maxRows,
      signal,
      rowHandler: (row) => {
        rowsSeen++;
        if (row.cells.length > 0) maxCol = Math.max(maxCol, row.cells[row.cells.length - 1].col + 1);
        // 第一行为表头：只用于列骨架，不参与列统计
        if (rowsSeen === 1) return true;
        for (const cell of row.cells) {
          bump(cell.col);
          const acc = accs[cell.col];
          acc.total++;
          const v = cell.value;
          if (v == null || v === '') continue;
          acc.nonEmpty++;
          const t = inferValueType(v);
          acc.types[t] = (acc.types[t] ?? 0) + 1;
          if (acc.samples.length < sample) acc.samples.push(v);
          if (isNumericValue(v)) {
            const n = typeof v === 'number' ? v : Number(v);
            if (Number.isNaN(n)) continue;
            if (!acc.numeric) acc.numeric = { sum: 0, min: n, max: n, cnt: 0 };
            acc.numeric.sum += n;
            if (n < acc.numeric.min) acc.numeric.min = n;
            if (n > acc.numeric.max) acc.numeric.max = n;
            acc.numeric.cnt++;
          }
        }
        return true;
      },
    });

    const totalCols = maxCol;
    const columns: ColumnProfile[] = [];
    for (let i = 0; i < totalCols; i++) {
      const acc = accs[i];
      if (!acc) continue;
      const numeric =
        acc.numeric && acc.numeric.cnt > 0
          ? {
              min: acc.numeric.min,
              max: acc.numeric.max,
              mean: acc.numeric.sum / acc.numeric.cnt,
              sum: acc.numeric.sum,
            }
          : undefined;
      // 不把 numeric 作为 own property 写入（避免 undefined 字段导致
      // dsh-tools 的 canonical lossless-JSON 校验失败：JSON.stringify
      // 会跳过 undefined 字段，使 round-trip 后该 own property 消失）。
      const profile: ColumnProfile = {
        col: i,
        header: headerMap.get(i) ?? null,
        total: acc.total,
        nonEmpty: acc.nonEmpty,
        emptyRate: acc.total > 0 ? round(1 - acc.nonEmpty / acc.total) : 0,
        types: acc.types,
        samples: acc.samples,
      };
      if (numeric) profile.numeric = numeric;
      columns.push(profile);
    }

    return {
      path: args.file_path,
      sheets,
      sheet,
      totalRows: rowsSeen,
      totalCols,
      columns,
      rawTotalCols: totalCols,
    };
  } finally {
    await reader.dispose();
  }
}

function round(n: number): number {
  return Math.round(n * 10000) / 10000;
}

/** 渲染为人类可读文本（ContentBlock text） */
export function formatDescribe(r: DescribeResult): string {
  const lines: string[] = [];
  lines.push(`File: ${r.path}`);
  lines.push(`Sheet: ${r.sheet}  (sheets: ${r.sheets.join(', ')})`);
  lines.push(`Rows: ${r.totalRows}   Cols: ${r.totalCols}`);
  lines.push('---');
  for (const col of r.columns) {
    const header = col.header ?? `col${col.col + 1}`;
    const types = Object.entries(col.types)
      .map(([k, v]) => `${k}:${v}`)
      .join(' ');
    const num = col.numeric
      ? ` [min=${col.numeric.min} max=${col.numeric.max} mean=${round(col.numeric.mean)}]`
      : '';
    const samples = col.samples.map((s) => `"${formatValue(s)}"`).join(', ');
    lines.push(
      `#${col.col + 1} ${header} | nonEmpty=${col.nonEmpty}/${col.total} emptyRate=${col.emptyRate} | ${types}${num} | e.g. ${samples}`,
    );
  }
  return lines.join('\n');
}
