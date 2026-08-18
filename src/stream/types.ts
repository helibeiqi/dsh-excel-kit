/**
 * dsh-excel-kit — 流式 xlsx 读取器类型定义
 */

/** 单元格原始值类型 */
export type CellValue = string | number | boolean | null;

/** 流式解析产出的单元格 */
export interface StreamCell {
  /** 列号（0-based），来自列字母如 B → 1 */
  col: number;
  value: CellValue;
}

/** 流式解析产出的行（不累积全表） */
export interface StreamRow {
  /** 行号（1-based，来自 r 属性） */
  row: number;
  cells: StreamCell[];
}

/** 行回调：返回 false 可提前终止解析 */
export type RowHandler = (row: StreamRow) => boolean | void;

/** 描述文件时用的列统计信息（紧凑） */
export interface ColumnProfile {
  col: number;
  /** 表头（第一行该列的值） */
  header: string | null;
  total: number;
  nonEmpty: number;
  emptyRate: number;
  /** 类型分布：{ string: n, number: n, boolean: n, date: n, other: n } */
  types: Record<string, number>;
  /** 数值列统计（仅当存在 number 类型时） */
  numeric?: { min: number; max: number; mean: number; sum: number };
  /** 示例值（最多 sample 条） */
  samples: CellValue[];
}

/** describe 输出（紧凑，绝不含全量行） */
export interface DescribeResult {
  path: string;
  sheets: string[];
  sheet: string;
  totalRows: number;
  totalCols: number;
  columns: ColumnProfile[];
  /** 原始列数（未裁剪） */
  rawTotalCols: number;
}

/** 单元格 → 行映射的便捷结构（用于 filter/pivot） */
export type RowMap = Record<string, CellValue>;

/** 过滤条件 */
export interface FilterCondition {
  column: string;
  op:
    | 'eq'
    | 'ne'
    | 'gt'
    | 'gte'
    | 'lt'
    | 'lte'
    | 'contains'
    | 'in'
    | 'between';
  value?: CellValue;
  values?: CellValue[];
}

/** 聚合操作符 */
export type AggOp = 'count' | 'sum' | 'mean' | 'min' | 'max';

/** pivot values 配置 */
export interface PivotValueSpec {
  column: string;
  agg: AggOp;
}

/** 解析选项 */
export interface ParseOptions {
  /** 仅解析指定 sheet（名称） */
  sheet?: string;
  /** 仅解析到前 maxRows 行（描述文件默认全扫但流式） */
  maxRows?: number;
  /** 是否仅取表头（第一行），供 describe 建列骨架 */
  headerOnly?: boolean;
}
