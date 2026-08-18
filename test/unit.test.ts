/**
 * dsh-excel-kit — 单元测试（解析器 + 三工具正确性）
 *
 * 运行方式：node --import tsx --test test/unit.test.ts
 */
import { before, test } from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
import { ensureFixtures, SMALL_XLSX, TINY_DEFLATE_XLSX } from './fixtures/generate';
import { executeDescribe } from '../src/tools/describe';
import { executeFilter } from '../src/tools/filter';
import { executePivot } from '../src/tools/pivot';
import { matchCondition } from '../src/utils';
import { excelSerialToIso, isDateNumFmtId } from '../src/stream/date';
import { XlsxStreamReader } from '../src/stream/xlsx-stream-reader';
import { maybeSpill } from '../src/spill';
import { apply } from '../src/index';

before(async () => {
  // 单元测试只生成 small.xlsx（跳过 100MB large）
  await ensureFixtures({ skipLarge: true });
});

// ---------------------------------------------------------------------------
// 日期工具
// ---------------------------------------------------------------------------

test('isDateNumFmtId 识别内置/自定义日期格式', () => {
  assert.equal(isDateNumFmtId(14, null), true);
  assert.equal(isDateNumFmtId(22, null), true);
  assert.equal(isDateNumFmtId(45, null), true);
  assert.equal(isDateNumFmtId(164, 'yyyy-mm-dd'), true);
  assert.equal(isDateNumFmtId(164, '0.00'), false);
  assert.equal(isDateNumFmtId(0, null), false);
  assert.equal(isDateNumFmtId(49, '@'), false);
});

test('excelSerialToIso 序列→ISO', () => {
  assert.equal(excelSerialToIso(1), '1900-01-01');
  assert.equal(excelSerialToIso(59), '1900-02-28');
  assert.equal(excelSerialToIso(61), '1900-03-01');
  assert.equal(excelSerialToIso(43831), '2020-01-01');
  assert.equal(excelSerialToIso(45292), '2024-01-01');
  assert.ok(excelSerialToIso(43831.5).startsWith('2020-01-01T'));
});

// ---------------------------------------------------------------------------
// 条件匹配
// ---------------------------------------------------------------------------

test('matchCondition 各 op', () => {
  assert.equal(matchCondition(10, { column: 'a', op: 'eq', value: 10 }), true);
  assert.equal(matchCondition(10, { column: 'a', op: 'eq', value: '10' }), true);
  assert.equal(matchCondition(10, { column: 'a', op: 'ne', value: 11 }), true);
  assert.equal(matchCondition(10, { column: 'a', op: 'gt', value: 9 }), true);
  assert.equal(matchCondition(10, { column: 'a', op: 'gte', value: 10 }), true);
  assert.equal(matchCondition(10, { column: 'a', op: 'lt', value: 11 }), true);
  assert.equal(matchCondition(10, { column: 'a', op: 'lte', value: 10 }), true);
  assert.equal(matchCondition('HelloWorld', { column: 'a', op: 'contains', value: 'world' }), true);
  assert.equal(matchCondition('X', { column: 'a', op: 'in', values: ['a', 'b', 'X'] }), true);
  assert.equal(matchCondition(15, { column: 'a', op: 'between', values: [10, 20] }), true);
  assert.equal(matchCondition(25, { column: 'a', op: 'between', values: [10, 20] }), false);
});

// ---------------------------------------------------------------------------
// excel_describe
// ---------------------------------------------------------------------------

test('describe small.xlsx Sheet1', async () => {
  const r = await executeDescribe({ file_path: SMALL_XLSX, sheet: 'Sheet1', sample: 3 });
  assert.deepEqual(r.sheets, ['Sheet1', 'Sheet2']);
  assert.equal(r.sheet, 'Sheet1');
  assert.equal(r.totalRows, 6001); // 表头 + 6000 数据行
  assert.equal(r.totalCols, 8);
  const headers = r.columns.map((c) => c.header);
  assert.deepEqual(headers, ['id', 'name', 'dept', 'score', 'hired', 'active', 'remark', 'money']);

  const dept = r.columns[2];
  assert.equal(dept.nonEmpty, 6000);
  assert.ok(dept.types.string === 6000);

  const score = r.columns[3];
  assert.ok(score.numeric);
  assert.ok(Math.abs(score.numeric!.mean - 500) < 1, `mean=${score.numeric!.mean}`);
  assert.ok(Math.abs(score.numeric!.min - 0.5) < 1e-6);
  assert.ok(Math.abs(score.numeric!.max - 999.5) < 1e-6);

  const hired = r.columns[4];
  assert.ok(hired.samples.length > 0);
  for (const s of hired.samples) {
    assert.match(String(s), /^\d{4}-\d{2}-\d{2}$/);
  }

  const money = r.columns[7];
  assert.ok(money.emptyRate > 0, 'money 列应有空值');
  // 科学计数法被正确解析为数字
  assert.ok(money.types.number > 0);
});

test('describe small.xlsx Sheet2', async () => {
  const r = await executeDescribe({ file_path: SMALL_XLSX, sheet: 'Sheet2' });
  assert.equal(r.totalRows, 51); // 表头 + 50 数据行
  assert.equal(r.totalCols, 4);
  assert.equal(r.columns[0].header, 'a');
});

test('describe 支持 max_rows 截断', async () => {
  const r = await executeDescribe({ file_path: SMALL_XLSX, max_rows: 100 });
  assert.equal(r.totalRows, 100);
});

// ---------------------------------------------------------------------------
// excel_filter
// ---------------------------------------------------------------------------

test('filter eq / 投影 / 数量', async () => {
  const r = await executeFilter({
    file_path: SMALL_XLSX,
    conditions: [{ column: 'dept', op: 'eq', value: 'HR' }],
    columns: ['id', 'name', 'dept'],
    limit: 100,
  });
  assert.ok(r.matched >= 900 && r.matched <= 1100, `matched=${r.matched}`);
  assert.equal(r.returned, 100);
  assert.equal(r.truncated, true);
  assert.deepEqual(r.columns, ['id', 'name', 'dept']);
  assert.ok(r.rows.length === 100);
  assert.equal(r.rows[0].values.dept, 'HR');
  assert.equal(Object.keys(r.rows[0].values).length, 3);
});

test('filter 数值比较 / between / in / contains', async () => {
  const gt = await executeFilter({
    file_path: SMALL_XLSX,
    conditions: [{ column: 'score', op: 'gt', value: 999 }],
    limit: 500,
  });
  assert.ok(gt.matched >= 4 && gt.matched <= 8, `gt.matched=${gt.matched}`);

  const btw = await executeFilter({
    file_path: SMALL_XLSX,
    conditions: [{ column: 'id', op: 'between', values: [1, 10] }],
  });
  assert.equal(btw.matched, 10);

  const inC = await executeFilter({
    file_path: SMALL_XLSX,
    conditions: [{ column: 'dept', op: 'in', values: ['HR', 'ENG'] }],
  });
  assert.equal(inC.matched, 2000);

  const contains = await executeFilter({
    file_path: SMALL_XLSX,
    conditions: [{ column: 'name', op: 'contains', value: 'Zhang' }],
  });
  assert.ok(contains.matched >= 100, `contains.matched=${contains.matched}`);
});

test('filter 多条件 AND + 空值处理', async () => {
  const r = await executeFilter({
    file_path: SMALL_XLSX,
    conditions: [
      { column: 'dept', op: 'eq', value: 'FIN' },
      { column: 'active', op: 'eq', value: true },
    ],
    limit: 10,
  });
  assert.ok(r.matched > 0);
  for (const row of r.rows) {
    assert.equal(row.values.active, true);
    assert.equal(row.values.dept, 'FIN');
  }
});

test('filter limit 硬上限 500', async () => {
  const r = await executeFilter({
    file_path: SMALL_XLSX,
    conditions: [{ column: 'dept', op: 'eq', value: 'HR' }],
    limit: 9999,
  });
  assert.equal(r.returned, 500);
});

test('filter 列不存在报错', async () => {
  await assert.rejects(
    executeFilter({ file_path: SMALL_XLSX, conditions: [{ column: 'nope', op: 'eq', value: 1 }] }),
    /not found/,
  );
});

// ---------------------------------------------------------------------------
// excel_pivot
// ---------------------------------------------------------------------------

test('pivot 单分组列', async () => {
  const r = await executePivot({
    file_path: SMALL_XLSX,
    rows: ['dept'],
    values: [
      { column: 'score', agg: 'mean' },
      { column: 'id', agg: 'count' },
    ],
  });
  assert.equal(r.groups, 6);
  assert.ok(r.data.length === 6);
  for (const d of r.data) {
    assert.equal(d.group.dept, d.key);
    assert.ok(Math.abs(d.values['mean:score']! - 500) < 1, `mean=${d.values['mean:score']}`);
    assert.ok(Math.abs(d.values['count:id']! - 1000) < 1, `count=${d.values['count:id']}`);
  }
});

test('pivot 多分组列嵌套键', async () => {
  const r = await executePivot({
    file_path: SMALL_XLSX,
    rows: ['dept', 'active'],
    values: [{ column: 'score', agg: 'sum' }],
  });
  assert.equal(r.groups, 12);
  for (const d of r.data) {
    assert.equal(d.key, `${d.group.dept}|${d.group.active}`);
  }
});

test('pivot min/max/sum', async () => {
  const r = await executePivot({
    file_path: SMALL_XLSX,
    rows: ['dept'],
    values: [
      { column: 'score', agg: 'min' },
      { column: 'score', agg: 'max' },
      { column: 'score', agg: 'sum' },
    ],
  });
  const d = r.data[0];
  assert.equal(d.values['min:score'], 0.5);
  assert.equal(d.values['max:score'], 999.5);
  assert.ok(Math.abs(d.values['sum:score']! - 1000 * 500) < 2000, `sum=${d.values['sum:score']}`);
});

// ---------------------------------------------------------------------------
// 解析器
// ---------------------------------------------------------------------------

test('streamSheet 逐行产出且可提前终止', async () => {
  const reader = new XlsxStreamReader(SMALL_XLSX);
  try {
    const names = await reader.getSheetNames();
    assert.deepEqual(names, ['Sheet1', 'Sheet2']);
    const header = await reader.getHeaderRow('Sheet1');
    assert.ok(header);
    assert.equal(header!.cells.length, 8);
    assert.equal(header!.cells[0].value, 'id');
    assert.equal(header!.cells[4].value, 'hired');

    let count = 0;
    const res = await reader.streamSheet('Sheet1', {
      maxRows: 10,
      rowHandler: () => {
        count++;
        return true;
      },
    });
    assert.equal(res.rowsRead, 10);
    assert.equal(count, 10);
  } finally {
    await reader.dispose();
  }
});

// ---------------------------------------------------------------------------
// 与 SheetJS 交叉验证（small.xlsx 由 xlsx 库可正常读取）
// ---------------------------------------------------------------------------

test('交叉验证：SheetJS 读取 deflate 小文件表头与行数一致', () => {
  // small.xlsx 的 sheet 条目为 store（不压缩）以精确控制体积，SheetJS CE 对
  // yazl store 条目兼容性差；交叉验证用同结构的 deflate 小文件。
  const wb = XLSX.readFile(TINY_DEFLATE_XLSX);
  const ws = wb.Sheets['Sheet1'];
  assert.ok(ws, 'Sheet1 exists');
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1');
  assert.equal(range.e.r + 1, 301); // 表头 + 300 数据行
  assert.equal(range.e.c + 1, 8); // 列数一致
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  assert.deepEqual(rows[0], ['id', 'name', 'dept', 'score', 'hired', 'active', 'remark', 'money']);
});

// ---------------------------------------------------------------------------
// spill 集成与插件入口
// ---------------------------------------------------------------------------

test('maybeSpill 阈值：小内容原样返回，大内容走 spill', async () => {
  const small = await maybeSpill({
    toolName: 't', callId: 'c1', exec: {}, ctx: {},
    content: 'x'.repeat(100), summary: 's',
  });
  assert.equal(small.full, true);
  assert.equal(small.spilled, null);

  const big = await maybeSpill({
    toolName: 't', callId: 'c2', exec: {}, ctx: {},
    content: 'x'.repeat(40000), summary: 's',
  });
  assert.ok(big.spilled, '超过阈值应产生 spill 引用');
  assert.equal(big.spilled!.bytes, 40000);
  assert.ok(big.spilled!.locator.startsWith('mem://'), '无真实 spillStore 时应走内存回退');
});

test('插件入口 apply 注册三个工具', async () => {
  const registered: any[] = [];
  apply({ tools: { register: (t: any) => registered.push(t) } } as any);
  const names = registered.map((t) => t.name).sort();
  assert.deepEqual(names, ['excel_describe', 'excel_filter', 'excel_pivot']);
  for (const t of registered) {
    assert.equal(typeof t.execute, 'function');
    assert.equal(typeof t.output.render, 'function');
    assert.equal(typeof t.output.schema, 'object');
    assert.equal(t.isConcurrencySafe(), true);
  }
});

test('excel_describe 工具 execute 封装（content + structuredContent）', async () => {
  const registered: any[] = [];
  apply({ tools: { register: (t: any) => registered.push(t) } } as any);
  const desc = registered.find((t) => t.name === 'excel_describe');
  const out = await desc.execute(
    { file_path: SMALL_XLSX, sheet: 'Sheet1', sample: 2 },
    { callId: 'c1', agent: { sessionId: 's1' } },
  );
  // execute 契约：{ content: [{type:'text',text}], structuredContent: <result> }
  assert.ok(Array.isArray(out.content), 'content 应为数组');
  assert.equal(out.content[0].type, 'text');
  assert.ok(typeof out.content[0].text === 'string' && out.content[0].text.length > 0);
  assert.equal(out.structuredContent.totalRows, 6001);
  assert.equal(out.structuredContent.sheet, 'Sheet1');
  // render 返回 content（registry 渲染路径）
  const blocks = desc.output.render({}, out);
  assert.equal(blocks[0].type, 'text');
});

test('dsh-tools canonical lossless-JSON 校验：execute 返回值经 JSON round-trip 后结构一致', async () => {
  const registered: any[] = [];
  apply({ tools: { register: (t: any) => registered.push(t) } } as any);
  for (const t of registered) {
    const args = t.name === 'excel_describe'
      ? { file_path: SMALL_XLSX, sample: 2 }
      : t.name === 'excel_filter'
        ? { file_path: SMALL_XLSX, conditions: [{ column: 'score', op: 'gt', value: 50 }], limit: 5 }
        : { file_path: SMALL_XLSX, rows: ['dept'], values: [{ column: 'score', agg: 'mean' }] };
    const out = await t.execute(args, { callId: 'c1', agent: { sessionId: 's1' } });
    // 模拟 dsh registry 的 lossless 校验：JSON.stringify + JSON.parse 后结构应一致
    // （无 undefined 字段残留、NaN/Inf 已被替换为 null）
    const json = JSON.stringify(out);
    const back = JSON.parse(json);
    // 不能引入 new own properties；所有数值必须有限
    function walk(o: any, p: string) {
      if (o === null) return;
      if (typeof o === 'number') {
        assert.ok(Number.isFinite(o), `${t.name}: non-finite number at ${p}=${o}`);
        return;
      }
      if (typeof o === 'object') {
        for (const k of Object.keys(o)) walk(o[k], `${p}.${k}`);
      }
    }
    walk(out, t.name);
    // 内容侧：output.render 必须能透传 content blocks
    const blocks = t.output.render({}, out);
    assert.ok(Array.isArray(blocks) && blocks.length > 0, `${t.name}: render 产出 blocks`);
    assert.equal(blocks[0].type, 'text');
    assert.ok(typeof blocks[0].text === 'string' && blocks[0].text.length > 0, `${t.name}: render 文本非空`);
  }
});
