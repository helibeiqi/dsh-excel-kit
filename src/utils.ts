/**
 * dsh-excel-kit — 通用工具：条件匹配 / 聚合 / 格式化（基于 lodash）
 */
import _ from 'lodash';
import { CellValue, FilterCondition, AggOp } from './stream/types';

/** 数值比较用 epsilon（浮点容差） */
const EPS = 1e-9;

function toNumber(v: CellValue): number | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v);
    return _.isNaN(n) ? null : n;
  }
  return null;
}

function toComparable(v: CellValue): string | number | null {
  if (v == null) return null;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = Number(v);
  if (v !== '' && !_.isNaN(n)) return n;
  return String(v);
}

/** 匹配一行单元格是否满足条件（column 为列索引） */
export function matchCondition(
  value: CellValue,
  cond: FilterCondition,
): boolean {
  if (value == null || value === '') {
    // 空值只可能命中 eq/ne/in 中的显式空匹配
    if (cond.op === 'eq') return cond.value == null || cond.value === '';
    if (cond.op === 'ne') return cond.value != null && cond.value !== '';
    if (cond.op === 'in') return (cond.values ?? []).some((v) => v == null || v === '');
    return false;
  }

  switch (cond.op) {
    case 'eq': {
      const a = toComparable(value);
      const b = toComparable(cond.value as CellValue);
      if (a == null || b == null) return String(value) === String(cond.value ?? '');
      return a === b;
    }
    case 'ne': {
      const a = toComparable(value);
      const b = toComparable(cond.value as CellValue);
      if (a == null || b == null) return String(value) !== String(cond.value ?? '');
      return a !== b;
    }
    case 'gt': {
      const a = toNumber(value);
      const b = toNumber(cond.value as CellValue);
      return a != null && b != null && a > b + EPS;
    }
    case 'gte': {
      const a = toNumber(value);
      const b = toNumber(cond.value as CellValue);
      return a != null && b != null && a >= b - EPS;
    }
    case 'lt': {
      const a = toNumber(value);
      const b = toNumber(cond.value as CellValue);
      return a != null && b != null && a < b - EPS;
    }
    case 'lte': {
      const a = toNumber(value);
      const b = toNumber(cond.value as CellValue);
      return a != null && b != null && a <= b + EPS;
    }
    case 'contains': {
      const needle = String(cond.value ?? '');
      return String(value).toLowerCase().includes(needle.toLowerCase());
    }
    case 'in': {
      const list = (cond.values ?? []).map((v) => toComparable(v));
      const a = toComparable(value);
      return list.some((v) => v === a) || list.some((v) => String(v) === String(value));
    }
    case 'between': {
      const a = toNumber(value);
      const [lo, hi] = (cond.values ?? []).map((v) => toNumber(v));
      return a != null && lo != null && hi != null && a >= lo - EPS && a <= hi + EPS;
    }
    default:
      return false;
  }
}

/** 多条件 AND 匹配 */
export function matchAllConditions(
  values: CellValue[],
  conditions: FilterCondition[],
): boolean {
  for (const cond of conditions) {
    const v = cond.column == null ? null : values[Number(cond.column)];
    if (!matchCondition(v, cond)) return false;
  }
  return true;
}

// ---------------------------------------------------------------------------
// 聚合
// ---------------------------------------------------------------------------

export interface AggState {
  count: number;
  sum: number;
  mean: number;
  min: number | null;
  max: number | null;
}

export function initAggState(): AggState {
  return { count: 0, sum: 0, mean: 0, min: null, max: null };
}

/** 累加一个数值（非数值行 count 仍累加——count 统计行数） */
export function accumAgg(state: AggState, value: CellValue, numeric: boolean): void {
  state.count += 1;
  if (!numeric) return;
  const n = toNumber(value);
  if (n == null) return;
  state.sum += n;
  if (state.min == null || n < state.min) state.min = n;
  if (state.max == null || n > state.max) state.max = n;
}

export function finalizeAgg(state: AggState, agg: AggOp): number | null {
  switch (agg) {
    case 'count':
      return state.count;
    case 'sum':
      return round6(state.sum);
    case 'mean':
      return state.count > 0 ? round6(state.sum / state.count) : null;
    case 'min':
      return state.min == null ? null : round6(state.min);
    case 'max':
      return state.max == null ? null : round6(state.max);
    default:
      return null;
  }
}

/** 保留 6 位小数，去浮点噪声 */
function round6(n: number): number {
  return _.round(n, 6);
}

// ---------------------------------------------------------------------------
// 格式化
// ---------------------------------------------------------------------------

const MAX_SAMPLE_CHARS = 80;

export function formatValue(v: CellValue): string {
  if (v == null) return '(empty)';
  if (typeof v === 'string') {
    return v.length > MAX_SAMPLE_CHARS ? `${v.slice(0, MAX_SAMPLE_CHARS)}…` : v;
  }
  return String(v);
}

export function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

/** 判定值是否为数值（describe 类型分布用） */
export function inferValueType(v: CellValue): 'string' | 'number' | 'boolean' | 'date' | 'other' | 'empty' {
  if (v == null) return 'empty';
  if (typeof v === 'number') return 'number';
  if (typeof v === 'boolean') return 'boolean';
  if (typeof v === 'string') {
    // ISO 日期字符串（由日期列转换而来）
    if (/^\d{4}-\d{2}-\d{2}(T.*)?$/.test(v) && !Number.isNaN(Date.parse(v))) {
      return 'date';
    }
    return 'string';
  }
  return 'other';
}

/** 判断值是否为数值列可统计（纯数字或数字字符串） */
export function isNumericValue(v: CellValue): boolean {
  return (
    _.isNumber(v) ||
    (typeof v === 'string' && v.trim() !== '' && !_.isNaN(Number(v)))
  );
}
