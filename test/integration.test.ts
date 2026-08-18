/**
 * dsh-excel-kit — 集成测试：1MB vs 100MB 耗时/内存对比
 *
 * 断言目标：
 *  - 100MB 文件读取成功（describe/filter/pivot 均可用）
 *  - 峰值 RSS 增量 < 800MB（有界内存，不整表驻留）
 *  - 三工具返回紧凑结果（JSON 体积可控，绝不返回海量原始行）
 *
 * 运行方式：node --import tsx --test test/integration.test.ts
 */
import { before, test } from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import { performance } from 'node:perf_hooks';
import { ensureFixtures } from './fixtures/generate';
import { executeDescribe } from '../src/tools/describe';
import { executeFilter } from '../src/tools/filter';
import { executePivot } from '../src/tools/pivot';

const RSS_DELTA_LIMIT = 800 * 1024 * 1024; // 800MB

interface MeasureResult<T> {
  elapsedMs: number;
  rssDelta: number;
  result: T;
}

/** 测量执行时间 + 峰值 RSS 增量 */
async function measure<T>(fn: () => Promise<T>): Promise<MeasureResult<T>> {
  const base = process.memoryUsage().rss;
  let peak = base;
  const timer = setInterval(() => {
    peak = Math.max(peak, process.memoryUsage().rss);
  }, 20);
  const t0 = performance.now();
  const result = await fn();
  const elapsedMs = performance.now() - t0;
  clearInterval(timer);
  const final = process.memoryUsage().rss;
  const rssDelta = Math.max(peak, final) - base;
  return { elapsedMs, rssDelta, result };
}

function mb(n: number): string {
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

let SMALL: string;
let LARGE: string;
let smallBytes = 0;
let largeBytes = 0;

before(async () => {
  const f = await ensureFixtures();
  SMALL = f.small;
  LARGE = f.large;
  smallBytes = f.smallBytes;
  largeBytes = f.largeBytes;
  console.log(`fixtures: small=${mb(smallBytes)}, large=${mb(largeBytes)}`);
});

test('large.xlsx 存在且确实是大文件（>=90MB）', () => {
  assert.ok(largeBytes >= 90 * 1024 * 1024, `large.xlsx 实际 ${mb(largeBytes)}`);
  assert.ok(fs.existsSync(LARGE));
});

test('1MB vs 100MB：excel_describe 耗时与内存', async () => {
  const small = await measure(() => executeDescribe({ file_path: SMALL, sheet: 'Sheet1' }));
  const large = await measure(() => executeDescribe({ file_path: LARGE, sheet: 'Sheet1' }));

  console.log(`describe small: ${small.elapsedMs.toFixed(0)}ms, rss+${mb(small.rssDelta)}`);
  console.log(`describe large: ${large.elapsedMs.toFixed(0)}ms, rss+${mb(large.rssDelta)}`);

  // 100MB 读取成功
  const l = large.result;
  assert.ok(l.totalRows > 100_000, `totalRows=${l.totalRows}`);
  assert.ok(l.totalRows < 2_000_000, `totalRows=${l.totalRows}`);
  assert.equal(l.totalCols, 11);
  assert.equal(l.sheets.length, 1);

  // emp_code 列为 SST 唯一串 → 非空计数 = 数据行数
  const empCode = l.columns.find((c) => c.header === 'emp_code');
  assert.ok(empCode, 'emp_code column exists');
  assert.equal(empCode!.nonEmpty, l.totalRows - 1);

  // 紧凑返回：结果 JSON 体积很小（不返回原始行）
  const jsonBytes = Buffer.byteLength(JSON.stringify(l), 'utf8');
  console.log(`describe large JSON: ${mb(jsonBytes)} (${jsonBytes} B)`);
  assert.ok(jsonBytes < 20 * 1024, `describe JSON 应紧凑，实际 ${jsonBytes} B`);

  // 峰值 RSS 增量低于阈值
  assert.ok(large.rssDelta < RSS_DELTA_LIMIT, `rssDelta=${mb(large.rssDelta)} >= ${mb(RSS_DELTA_LIMIT)}`);
  // 耗时合理（流式解析不应异常缓慢）
  assert.ok(large.elapsedMs < 180_000, `elapsed=${large.elapsedMs}ms`);
});

test('large.xlsx excel_filter 可用且返回紧凑结果', async () => {
  const m = await measure(() =>
    executeFilter({
      file_path: LARGE,
      conditions: [{ column: 'cnt', op: 'gte', value: 400 }],
      columns: ['emp_id', 'emp_code', 'cnt', 'grade'],
      limit: 100,
    }),
  );
  console.log(`filter large: ${m.elapsedMs.toFixed(0)}ms, rss+${mb(m.rssDelta)}, matched=${m.result.matched}, returned=${m.result.returned}`);
  assert.ok(m.result.matched > 0, '应匹配到行');
  assert.equal(m.result.returned, 100);
  assert.equal(m.result.truncated, true);
  for (const row of m.result.rows) {
    assert.ok(row.values.cnt >= 400);
    assert.deepEqual(Object.keys(row.values).sort(), ['cnt', 'emp_code', 'emp_id', 'grade']);
  }
  assert.ok(m.rssDelta < RSS_DELTA_LIMIT, `rssDelta=${mb(m.rssDelta)}`);
});

test('large.xlsx excel_pivot 可用且返回紧凑结果', async () => {
  const m = await measure(() =>
    executePivot({
      file_path: LARGE,
      rows: ['grade'],
      values: [{ column: 'cnt', agg: 'sum' }],
    }),
  );
  console.log(`pivot large: ${m.elapsedMs.toFixed(0)}ms, rss+${mb(m.rssDelta)}, groups=${m.result.groups}`);
  assert.equal(m.result.groups, 4); // A/B/C/D
  assert.equal(m.result.truncated, false);
  const d = m.result.data[0];
  assert.ok(d.values['sum:cnt']! > 0);
  assert.ok(m.rssDelta < RSS_DELTA_LIMIT, `rssDelta=${mb(m.rssDelta)}`);
});

test('large.xlsx 多分组 pivot', async () => {
  const r = await executePivot({
    file_path: LARGE,
    rows: ['grade', 'active'],
    values: [{ column: 'emp_id', agg: 'count' }],
  });
  assert.equal(r.groups, 8); // 4 grade × 2 active
  for (const d of r.data) {
    assert.equal(d.key, `${d.group.grade}|${d.group.active}`);
  }
});

test('大文件上 SST 分块 spill 生效（SST 唯一串>8MB）', async () => {
  // large.xlsx 的 SST 包含数十万条 40 字符唯一串 → 必然触发分块落盘
  const base = 'EMPLOYEE-CODE-';
  const suffix = String(1).padStart(12, '0');
  const pad = 'X'.repeat(Math.max(0, 40 - base.length - suffix.length));
  const code1 = (base + pad + suffix).slice(0, 40);
  const m = await measure(() =>
    executeFilter({
      file_path: LARGE,
      conditions: [{ column: 'emp_code', op: 'eq', value: code1 }],
      columns: ['emp_id'],
    }),
  );
  console.log(`filter sst lookup: ${m.elapsedMs.toFixed(0)}ms, rss+${mb(m.rssDelta)}`);
  assert.equal(m.result.matched, 1);
  assert.equal(m.result.rows[0].values.emp_id, 1);
  assert.ok(m.rssDelta < RSS_DELTA_LIMIT, `rssDelta=${mb(m.rssDelta)}`);
});
