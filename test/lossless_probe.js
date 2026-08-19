// Lossless JSON 复现：跑 executeDescribe 并做 round-trip 校验，定位失败字段
const { executeDescribe } = require('../lib/tools/describe');
const { executeFilter } = require('../lib/tools/filter');
const { executePivot } = require('../lib/tools/pivot');

async function probe(name, fn) {
  try {
    const r = await fn();
    const json = JSON.stringify(r);
    const back = JSON.parse(json);
    const a = JSON.stringify(r);
    const b = JSON.stringify(back);
    const lossy = a !== b;
    const nonFinite = [];
    const undef = [];
    function walk(o, p = '') {
      if (o === null) return;
      if (typeof o === 'number') {
        if (Number.isNaN(o) || !Number.isFinite(o)) nonFinite.push(p + '=' + String(o));
        return;
      }
      if (typeof o === 'undefined') { undef.push(p); return; }
      if (typeof o !== 'object') return;
      for (const k of Object.keys(o)) walk(o[k], p + '.' + k);
    }
    walk(r, 'root');
    console.log(`[${name}] lossy=${lossy} jsonBytes=${json.length} nonFinite=${nonFinite.length} undef=${undef.length}`);
    if (nonFinite.length) console.log('  nonFinite:', nonFinite.slice(0, 10).join(' | '));
    if (undef.length) console.log('  undef:', undef.slice(0, 10).join(' | '));
    if (lossy) {
      // 找出哪里 round-trip 不一致（值丢失）
      for (const k of Object.keys(r)) {
        if (JSON.stringify(r[k]) !== JSON.stringify(back[k])) {
          console.log(`  diff field: ${k}`);
        }
      }
    }
  } catch (e) {
    console.log(`[${name}] EXCEPTION: ${e.message}`);
  }
}

(async () => {
  const path = require('path');
  // 相对脚本位置定位 fixture，避免硬编码本机路径（换机器可跑）
  const XLSX = path.join(__dirname, 'fixtures', 'small.xlsx');
  await probe('describe(small)', () => executeDescribe({ file_path: XLSX }));
  await probe('filter(small,age>20)', () => executeFilter({ file_path: XLSX, conditions: [{ column: 'age', op: 'gt', value: 20 }], limit: 10 }));
  await probe('pivot(small,dept)', () => executePivot({ file_path: XLSX, rows: ['dept'], values: [{ column: 'age', agg: 'mean' }] }));
})();
