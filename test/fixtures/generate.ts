/**
 * dsh-excel-kit — 测试 fixture 流式生成器
 *
 * 用 yazl 从零构造合法 xlsx（zip），绝不在内存中整表驻留：
 *  - sheet XML / sharedStrings XML 先流式写入临时文件（带背压），再用 zipfile.addFile 打进 zip
 *  - large.xlsx 目标 ~100MB（store 模式，加速生成），含超大 sharedStrings（触发分块 spill）
 *  - small.xlsx ~1MB，覆盖字符串(SST)/inlineStr/数字/科学计数法/布尔/日期格式列/空单元格/多 sheet
 */
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as yazl from 'yazl';

const FIXTURES_DIR = path.join(__dirname, '..', 'fixtures');
export const SMALL_XLSX = path.join(FIXTURES_DIR, 'small.xlsx');
export const LARGE_XLSX = path.join(FIXTURES_DIR, 'large.xlsx');
/** 供 SheetJS 交叉验证的 deflate 小文件（SheetJS 对 yazl store 条目兼容性差） */
export const TINY_DEFLATE_XLSX = path.join(FIXTURES_DIR, 'tiny-deflate.xlsx');

// ---------------------------------------------------------------------------
// 常量
// ---------------------------------------------------------------------------

const NAMES = [
  'Zhang Wei', 'Li Na', 'Wang Fang', 'Liu Yang', 'Chen Jing', 'Yang Fan',
  'Huang Lei', 'Zhao Min', 'Wu Qiang', 'Zhou Yu', 'Xu Jie', 'Sun Li',
  'Ma Chao', 'Zhu Dan', 'Hu Bin', 'Guo Xin', 'He Lin', 'Gao Yuan',
  'Lin Xiao', 'Luo Qi', 'Zheng Hao', 'Liang Bo', 'Xie Nan', 'Song Tao',
  'Tang Wei', 'Han Mei', 'Cao Jun', 'Xu Lei', 'Deng Li', 'Feng Xiao',
  'Ceng Xin', 'Peng Yu', 'Xiao Gang', 'Tian Jing', 'Dong Fang', 'Pan Yue',
  'Yuan Jing', 'Jiang Tao', 'Cai Wen', 'Yu Min', 'Du Juan', 'Ding Lei',
  'Ren Jing', 'Rao Yi', 'Su Man', 'Wei Ran', 'Jia Lin', 'Fan Yi',
  'Shi Jing', 'Bai Yun',
] as const;

const DEPTS = ['HR', 'ENG', 'SALES', 'OPS', 'FIN', 'R&D'] as const;
const GRADES = ['A', 'B', 'C', 'D'] as const;

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';

// ---------------------------------------------------------------------------
// 工具
// ---------------------------------------------------------------------------

function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** 列字母（1-based） */
function colName(n: number): string {
  let s = '';
  let v = n;
  while (v > 0) {
    const rem = (v - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    v = Math.floor((v - 1) / 26);
  }
  return s;
}

/** 带背压的顺序写文件 */
function createWriter(file: string) {
  const ws = fs.createWriteStream(file);
  return {
    write(s: string): Promise<void> {
      return new Promise((resolve) => {
        if (ws.write(s)) resolve();
        else ws.once('drain', resolve);
      });
    },
    end(): Promise<void> {
      return new Promise((resolve, reject) => {
        ws.end((err) => (err ? reject(err) : resolve()));
      });
    },
  };
}

// ---------------------------------------------------------------------------
// 静态 XML 片段
// ---------------------------------------------------------------------------

function contentTypesXml(withSst: boolean): string {
  const sst = withSst
    ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${NS_CONTENT_TYPES}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${sst}</Types>`;
}

function rootRelsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_PKG_REL}">
<Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

function workbookXml(sheets: { name: string; rId: string }[]): string {
  const s = sheets
    .map((x) => `<sheet name="${esc(x.name)}" sheetId="${x.rId.replace('rId', '')}" r:id="${x.rId}"/>`)
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">
<sheets>${s}</sheets>
</workbook>`;
}

function workbookRelsXml(rels: { id: string; target: string }[]): string {
  const s = rels
    .map((r) => `<Relationship Id="${r.id}" Type="${NS_REL}/worksheet" Target="${r.target}"/>`)
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_PKG_REL}">${s}</Relationships>`;
}

function stylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS_MAIN}">
<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0"/></cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0"/>
<xf numFmtId="164" applyNumberFormat="1"/>
</cellXfs>
</styleSheet>`;
}

const SHEET_HEAD = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">
<sheetData>`;

const SHEET_FOOT = `</sheetData>
</worksheet>`;

/** 表头行：inlineStr */
function headerRow(headers: string[]): string {
  const cells = headers
    .map((h, i) => `<c r="${colName(i + 1)}1" t="inlineStr"><is><t>${esc(h)}</t></is></c>`)
    .join('');
  return `<row r="1">${cells}</row>\n`;
}

function sstXmlHead(count: number): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="${NS_MAIN}" count="${count}" uniqueCount="${count}">`;
}

function sstXmlFoot(): string {
  return `</sst>`;
}

function si(s: string): string {
  return `<si><t>${esc(s)}</t></si>`;
}

// ---------------------------------------------------------------------------
// 数据行生成
// ---------------------------------------------------------------------------

/** small.xlsx 数据行（r 为数据行号，XML 行号为 r+1） */
function smallDataRow(r: number): string {
  const row = r + 1;
  const nameIdx = r % NAMES.length;
  const deptIdx = r % DEPTS.length;
  // 分数：按 6 行块取分（每块内各 dept 同分），使每个 dept 组都完整覆盖 0.5..999.5
  const score = ((Math.floor((r - 1) / DEPTS.length) % 1000) + 0.5).toFixed(1);
  const serial = 43831 + (r % 400); // 2020-01-01 起
  // active：按 6 行块翻转，与 dept 解耦（dept*active 组合可达 12 组）
  const active = Math.floor(r / DEPTS.length) % 2;
  const remark = r % 10 === 0 ? '' : `note-${r % 100}`;
  const money = r % 100 === 0 ? `1.5E${(r % 5) + 1}` : ((r * 3.14) % 1000).toFixed(2);
  const cells: string[] = [];
  cells.push(`<c r="A${row}"><v>${r}</v></c>`); // id
  cells.push(`<c r="B${row}" t="s"><v>${nameIdx}</v></c>`); // name
  cells.push(`<c r="C${row}" t="s"><v>${NAMES.length + deptIdx}</v></c>`); // dept
  cells.push(`<c r="D${row}"><v>${score}</v></c>`); // score
  cells.push(`<c r="E${row}" s="1"><v>${serial}</v></c>`); // hired (date)
  cells.push(`<c r="F${row}" t="b"><v>${active}</v></c>`); // active
  if (remark === '') {
    cells.push(`<c r="G${row}" t="inlineStr"><is></is></c>`); // 显式空 inlineStr
  } else {
    cells.push(`<c r="G${row}" t="inlineStr"><is><t>${esc(remark)}</t></is></c>`);
  }
  if (r % 7 === 0) {
    cells.push(`<c r="H${row}"/>`); // 显式空数字单元格
  } else {
    cells.push(`<c r="H${row}"><v>${money}</v></c>`);
  }
  return `<row r="${row}">${cells.join('')}</row>\n`;
}

/** small.xlsx Sheet2 数据（纯数字小表） */
function sheet2DataRow(r: number): string {
  const row = r + 1;
  return `<row r="${row}"><c r="A${row}"><v>${r}</v></c><c r="B${row}"><v>${r * 2}</v></c><c r="C${row}"><v>${r * r}</v></c><c r="D${row}" t="s"><v>${NAMES.length + (r % DEPTS.length)}</v></c></row>\n`;
}

function uniqueCode(r: number): string {
  const base = 'EMPLOYEE-CODE-';
  const suffix = String(r).padStart(12, '0');
  const pad = 'X'.repeat(Math.max(0, 40 - base.length - suffix.length));
  return (base + pad + suffix).slice(0, 40);
}

/** large.xlsx 数据行 */
function largeDataRow(r: number, gradeSstOffset: number): string {
  const row = r + 1;
  const score = ((r % 10000) + 0.25).toFixed(2);
  const serial = 43831 + (r % 500);
  const active = r % 3 === 0 ? 1 : 0;
  const gradeIdx = r % GRADES.length;
  const n1 = ((r * 1.5) % 100).toFixed(3);
  const n2 = ((r * 0.7) % 50).toFixed(3);
  const n3 = ((r * 2.3) % 200).toFixed(3);
  const n4 = ((r * 0.11) % 10).toFixed(4);
  const cells: string[] = [];
  cells.push(`<c r="A${row}"><v>${r}</v></c>`); // emp_id
  cells.push(`<c r="B${row}" t="s"><v>${r - 1}</v></c>`); // emp_code (SST 唯一)
  cells.push(`<c r="C${row}"><v>${score}</v></c>`); // score
  cells.push(`<c r="D${row}"><v>${n1}</v></c>`); // v1
  cells.push(`<c r="E${row}" s="1"><v>${serial}</v></c>`); // hired (date)
  cells.push(`<c r="F${row}" t="b"><v>${active}</v></c>`); // active
  cells.push(`<c r="G${row}"><v>${n2}</v></c>`); // v2
  cells.push(`<c r="H${row}"><v>${n3}</v></c>`); // v3
  cells.push(`<c r="I${row}"><v>${n4}</v></c>`); // v4
  cells.push(`<c r="J${row}"><v>${(r * 7) % 500}</v></c>`); // cnt
  cells.push(`<c r="K${row}" t="s"><v>${gradeSstOffset + gradeIdx}</v></c>`); // grade
  return `<row r="${row}">${cells.join('')}</row>\n`;
}

// ---------------------------------------------------------------------------
// 生成主流程
// ---------------------------------------------------------------------------

async function writeSstTemp(items: string[]): Promise<string> {
  const file = path.join(os.tmpdir(), `dsh-excel-kit-sst-${Date.now()}-${Math.random().toString(36).slice(2)}.xml`);
  const w = createWriter(file);
  await w.write(sstXmlHead(items.length));
  for (const it of items) await w.write(si(it));
  await w.write(sstXmlFoot());
  await w.end();
  return file;
}

/** 流式写 sheet 到临时文件；genDataRow(r) 产数据行（r 从 1 开始），rows 为数据行数 */
async function writeSheetTemp(
  genDataRow: (r: number) => string,
  headers: string[],
  rows: number,
): Promise<string> {
  const file = path.join(os.tmpdir(), `dsh-excel-kit-sheet-${Date.now()}-${Math.random().toString(36).slice(2)}.xml`);
  const w = createWriter(file);
  await w.write(SHEET_HEAD);
  await w.write(headerRow(headers));
  for (let r = 1; r <= rows; r++) await w.write(genDataRow(r));
  await w.write(SHEET_FOOT);
  await w.end();
  return file;
}

interface ZipEntry {
  name: string;
  buffer?: Buffer;
  file?: string;
  compress?: boolean;
}

async function assembleZip(outPath: string, entries: ZipEntry[]): Promise<void> {
  const zf = new yazl.ZipFile();
  const outStream = zf.outputStream;
  const writeStream = fs.createWriteStream(outPath);
  outStream.pipe(writeStream);
  for (const e of entries) {
    if (e.buffer) zf.addBuffer(e.buffer, e.name);
    else if (e.file) zf.addFile(e.file, e.name, { compress: e.compress ?? true });
  }
  zf.end();
  await new Promise<void>((resolve, reject) => {
    writeStream.on('close', resolve);
    writeStream.on('error', reject);
  });
}

function ensureDir(): void {
  fs.mkdirSync(FIXTURES_DIR, { recursive: true });
}

function fileSize(p: string): number {
  try {
    return fs.statSync(p).size;
  } catch {
    return 0;
  }
}

// ---------------------------------------------------------------------------
// 公开 API
// ---------------------------------------------------------------------------

export interface GenerateOptions {
  largeTargetBytes?: number;
  smallRows?: number;
  largeRows?: number;
  force?: boolean;
  /** 跳过 large.xlsx 生成（单元测试用，避免 100MB） */
  skipLarge?: boolean;
}

/** 生成 small.xlsx（~1MB，含多 sheet / 日期 / SST / 科学计数法 / 空单元格） */
async function generateSmall(rows: number): Promise<string> {
  ensureDir();
  const sstItems = [...NAMES, ...DEPTS].map((s) => s);
  const sstFile = await writeSstTemp(sstItems);
  const sheet1File = await writeSheetTemp(smallDataRow, ['id', 'name', 'dept', 'score', 'hired', 'active', 'remark', 'money'], rows);
  const sheet2File = await writeSheetTemp(sheet2DataRow, ['a', 'b', 'c', 'dept'], 50);

  const sstBuf = fs.readFileSync(sstFile);
  await assembleZip(SMALL_XLSX, [
    { name: '[Content_Types].xml', buffer: Buffer.from(contentTypesXml(true), 'utf8') },
    { name: '_rels/.rels', buffer: Buffer.from(rootRelsXml(), 'utf8') },
    { name: 'xl/workbook.xml', buffer: Buffer.from(workbookXml([{ name: 'Sheet1', rId: 'rId1' }, { name: 'Sheet2', rId: 'rId2' }]), 'utf8') },
    { name: 'xl/_rels/workbook.xml.rels', buffer: Buffer.from(workbookRelsXml([
      { id: 'rId1', target: 'worksheets/sheet1.xml' },
      { id: 'rId2', target: 'worksheets/sheet2.xml' },
    ]), 'utf8') },
    { name: 'xl/styles.xml', buffer: Buffer.from(stylesXml(), 'utf8') },
    { name: 'xl/sharedStrings.xml', buffer: sstBuf },
    { name: 'xl/worksheets/sheet1.xml', file: sheet1File, compress: false },
    { name: 'xl/worksheets/sheet2.xml', file: sheet2File, compress: false },
  ]);

  for (const f of [sstFile, sheet1File, sheet2File]) {
    await fs.promises.unlink(f).catch(() => {});
  }
  return SMALL_XLSX;
}

/** 生成 large.xlsx（目标 ~100MB store 模式；超大 SST 触发分块 spill） */
async function generateLarge(targetBytes: number, maxRows: number | undefined): Promise<string> {
  ensureDir();
  // 以首行长度为基准估算行数
  const probeRows = 32;
  let probe = '';
  for (let r = 1; r <= probeRows; r++) probe += largeDataRow(r, 999_999_999);
  const avgLen = Buffer.byteLength(probe, 'utf8') / probeRows;
  const rows = maxRows ?? Math.min(Math.floor(targetBytes / avgLen), 2_000_000);
  const gradeSstOffset = rows; // SST = emp codes(rows) + grades(4)
  // 先写 sheet
  const sheetFile = await writeSheetTemp(
    (r) => largeDataRow(r, gradeSstOffset),
    ['emp_id', 'emp_code', 'score', 'v1', 'hired', 'active', 'v2', 'v3', 'v4', 'cnt', 'grade'],
    rows,
  );
  // 再流式写 SST（emp codes + grades）
  const sstFile = await writeSstTempStream(rows);

  await assembleZip(LARGE_XLSX, [
    { name: '[Content_Types].xml', buffer: Buffer.from(contentTypesXml(true), 'utf8') },
    { name: '_rels/.rels', buffer: Buffer.from(rootRelsXml(), 'utf8') },
    { name: 'xl/workbook.xml', buffer: Buffer.from(workbookXml([{ name: 'Sheet1', rId: 'rId1' }]), 'utf8') },
    { name: 'xl/_rels/workbook.xml.rels', buffer: Buffer.from(workbookRelsXml([{ id: 'rId1', target: 'worksheets/sheet1.xml' }]), 'utf8') },
    { name: 'xl/styles.xml', buffer: Buffer.from(stylesXml(), 'utf8') },
    // store 模式：不压缩，使 large.xlsx 实际字节数接近 ~100MB，加速生成与读取
    { name: 'xl/sharedStrings.xml', file: sstFile, compress: false },
    { name: 'xl/worksheets/sheet1.xml', file: sheetFile, compress: false },
  ]);

  for (const f of [sheetFile, sstFile]) {
    await fs.promises.unlink(f).catch(() => {});
  }
  return LARGE_XLSX;
}

/** SST（uniqueCode 1..rows + GRADES）流式写入临时文件 */
async function writeSstTempStream(rows: number): Promise<string> {
  const file = path.join(os.tmpdir(), `dsh-excel-kit-sst-${Date.now()}-${Math.random().toString(36).slice(2)}.xml`);
  const w = createWriter(file);
  const count = rows + GRADES.length;
  await w.write(sstXmlHead(count));
  for (let r = 1; r <= rows; r++) await w.write(si(uniqueCode(r)));
  for (const g of GRADES) await w.write(si(g));
  await w.write(sstXmlFoot());
  await w.end();
  return file;
}

/** 生成 tiny-deflate.xlsx：与 small 相同结构，但 deflate 压缩（SheetJS 交叉验证用） */
async function generateTinyDeflate(rows: number): Promise<string> {
  ensureDir();
  const sstItems = [...NAMES, ...DEPTS].map((s) => s);
  const sstFile = await writeSstTemp(sstItems);
  const sheet1File = await writeSheetTemp(smallDataRow, ['id', 'name', 'dept', 'score', 'hired', 'active', 'remark', 'money'], rows);

  const sstBuf = fs.readFileSync(sstFile);
  await assembleZip(TINY_DEFLATE_XLSX, [
    { name: '[Content_Types].xml', buffer: Buffer.from(contentTypesXml(true), 'utf8') },
    { name: '_rels/.rels', buffer: Buffer.from(rootRelsXml(), 'utf8') },
    { name: 'xl/workbook.xml', buffer: Buffer.from(workbookXml([{ name: 'Sheet1', rId: 'rId1' }]), 'utf8') },
    { name: 'xl/_rels/workbook.xml.rels', buffer: Buffer.from(workbookRelsXml([{ id: 'rId1', target: 'worksheets/sheet1.xml' }]), 'utf8') },
    { name: 'xl/styles.xml', buffer: Buffer.from(stylesXml(), 'utf8') },
    { name: 'xl/sharedStrings.xml', buffer: sstBuf },
    { name: 'xl/worksheets/sheet1.xml', file: sheet1File }, // deflate（默认 compress:true）
  ]);

  for (const f of [sstFile, sheet1File]) {
    await fs.promises.unlink(f).catch(() => {});
  }
  return TINY_DEFLATE_XLSX;
}

/** 确保 fixtures 存在（缺失或 force 时生成）；返回实际路径与字节数 */
export async function ensureFixtures(opts: GenerateOptions = {}): Promise<{ small: string; large: string; smallBytes: number; largeBytes: number }> {
  const {
    largeTargetBytes = 100 * 1024 * 1024,
    smallRows = 6000,
    largeRows,
    force = false,
    skipLarge = false,
  } = opts;

  if (force || fileSize(SMALL_XLSX) === 0) {
    await generateSmall(smallRows);
  }
  if (force || fileSize(TINY_DEFLATE_XLSX) === 0) {
    await generateTinyDeflate(300);
  }
  if (!skipLarge && (force || fileSize(LARGE_XLSX) === 0)) {
    await generateLarge(largeTargetBytes, largeRows);
  }
  return {
    small: SMALL_XLSX,
    large: LARGE_XLSX,
    smallBytes: fileSize(SMALL_XLSX),
    largeBytes: fileSize(LARGE_XLSX),
  };
}

// CLI 入口：npm run gen:fixtures [--force] [--large-rows N]
if (require.main === module) {
  (async () => {
    const force = process.argv.includes('--force');
    const lrIdx = process.argv.indexOf('--large-rows');
    const largeRows = lrIdx >= 0 ? parseInt(process.argv[lrIdx + 1], 10) : undefined;
    const t0 = Date.now();
    const r = await ensureFixtures({ force, largeRows });
    console.log(`small.xlsx: ${(r.smallBytes / 1024 / 1024).toFixed(1)} MB`);
    console.log(`large.xlsx: ${(r.largeBytes / 1024 / 1024).toFixed(1)} MB`);
    console.log(`done in ${((Date.now() - t0) / 1000).toFixed(1)}s`);
  })().catch((e) => {
    console.error(e);
    process.exit(1);
  });
}
