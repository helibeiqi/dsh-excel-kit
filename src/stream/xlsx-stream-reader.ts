/**
 * dsh-excel-kit — 自研 xlsx 流式读取器
 *
 * 原理：yauzl 流式解压 zip（lazyEntries，逐 entry openReadStream）+ sax 流式 XML 解析。
 * 禁止 XLSX.readFile 全量加载；行数据逐行产出、不累积全量数组。
 *
 * 关键能力：
 *  - xl/sharedStrings.xml 分块 spill 到临时文件（内存有界），随机访问带 LRU 缓存
 *  - xl/worksheets/sheetN.xml 逐 <row> 组装对象，支持 inlineStr / s(SST索引) /
 *    数值(含 E 科学计数法) / t="str" / 空单元格
 *  - xl/styles.xml numFmtId 识别日期列，日期单元格输出 ISO 字符串
 *  - 支持 AbortSignal 取消
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { randomUUID } from 'crypto';
import { Readable } from 'stream';
import * as yauzl from 'yauzl';
import sax from 'sax';
import { CellValue, StreamCell, StreamRow } from './types';
import { excelSerialToIso, isDateNumFmtId } from './date';

// ---------------------------------------------------------------------------
// 错误类型
// ---------------------------------------------------------------------------

export class AbortError extends Error {
  constructor() {
    super('Excel read aborted');
    this.name = 'AbortError';
  }
}

export class SheetNotFoundError extends Error {
  constructor(sheet: string, available: string[]) {
    super(`Sheet "${sheet}" not found. Available sheets: ${available.join(', ') || '(none)'}`);
    this.name = 'SheetNotFoundError';
  }
}

// ---------------------------------------------------------------------------
// 共享字符串分块 spill 数组
// ---------------------------------------------------------------------------

/** 每块目标大小（字节） */
const SST_CHUNK_TARGET = 8 * 1024 * 1024;
/** 内存中缓存的已落盘块数（LRU） */
const SST_CACHE_CHUNKS = 4;

class SpilledStringArray {
  private current: string[] = [];
  private currentBytes = 0;
  private chunks: { file: string; count: number }[] = [];
  private writeChain: Promise<void> = Promise.resolve();
  private cache: { chunk: number; strings: string[] }[] = [];
  private tmpPrefix: string;
  private disposed = false;

  constructor(tmpDir?: string) {
    const dir = tmpDir ?? os.tmpdir();
    this.tmpPrefix = path.join(dir, `dsh-excel-kit-sst-${randomUUID().replace(/-/g, '')}-`);
  }

  /** 同步入队一条字符串；达到块阈值时异步落盘 */
  push(s: string): void {
    const bytes = Buffer.byteLength(s, 'utf8');
    if (this.currentBytes > 0 && this.currentBytes + bytes > SST_CHUNK_TARGET) {
      this.flush();
    }
    this.current.push(s);
    this.currentBytes += bytes;
  }

  private flush(): void {
    if (this.current.length === 0) return;
    const strings = this.current;
    const file = `${this.tmpPrefix}${this.chunks.length}.json`;
    this.current = [];
    this.currentBytes = 0;
    this.chunks.push({ file, count: strings.length });
    // 串行化落盘，保证文件可见顺序
    this.writeChain = this.writeChain.then(() =>
      fs.promises.writeFile(file, JSON.stringify(strings), 'utf8'),
    );
  }

  /** 结束写入：flush 剩余块并等待全部落盘 */
  async finalize(): Promise<void> {
    this.flush();
    await this.writeChain;
  }

  /** 随机访问（索引有序递增时命中内存缓存，性能接近数组） */
  async get(index: number): Promise<string> {
    await this.writeChain;
    if (index < 0) throw new RangeError(`SST index out of range: ${index}`);
    let offset = 0;
    for (let c = 0; c < this.chunks.length; c++) {
      const count = this.chunks[c].count;
      if (index < offset + count) {
        const strings = await this.loadChunk(c);
        return strings[index - offset];
      }
      offset += count;
    }
    throw new RangeError(`SST index out of range: ${index}`);
  }

  private async loadChunk(c: number): Promise<string[]> {
    for (let i = 0; i < this.cache.length; i++) {
      if (this.cache[i].chunk === c) {
        // LRU：移到末尾
        const [hit] = this.cache.splice(i, 1);
        this.cache.push(hit);
        return hit.strings;
      }
    }
    const strings = JSON.parse(
      await fs.promises.readFile(this.chunks[c].file, 'utf8'),
    ) as string[];
    this.cache.push({ chunk: c, strings });
    if (this.cache.length > SST_CACHE_CHUNKS) this.cache.shift();
    return strings;
  }

  /** 删除临时文件 */
  async dispose(): Promise<void> {
    if (this.disposed) return;
    this.disposed = true;
    await this.writeChain;
    for (const c of this.chunks) {
      await fs.promises.unlink(c.file).catch(() => {
        /* ignore */
      });
    }
  }
}

// ---------------------------------------------------------------------------
// 内部类型
// ---------------------------------------------------------------------------

interface CellXf {
  numFmtId: number;
  numFmtCode: string | null;
  isDate: boolean;
}

const DEFAULT_XF: CellXf = { numFmtId: 0, numFmtCode: null, isDate: false };

interface PreparedMeta {
  sheets: { name: string; target?: string }[];
  cellXfs: CellXf[];
}

interface XmlHandlers {
  onOpenTag?: (name: string, attrs: Record<string, string>) => void;
  onCloseTag?: (name: string) => void;
  onText?: (text: string) => void;
}

interface RowCtx {
  rowNum: number;
  cells: { col: number; value: CellValue | { __sst: number } }[];
}

/** 流式行回调：返回 false 提前终止 */
export type StreamRowHandler = (row: StreamRow) => boolean | void;

export interface StreamSheetOptions {
  rowHandler: StreamRowHandler;
  /** 只取前 maxRows 行（未设则全扫，仍为流式） */
  maxRows?: number;
  /** 仅产出一行（表头），用于建列骨架 */
  headerOnly?: boolean;
  /** 跳过第一行（表头行）不交给 rowHandler，也不计入 rowsRead */
  skipFirstRow?: boolean;
  signal?: AbortSignal;
}

export interface StreamSheetResult {
  rowsRead: number;
  endedEarly: boolean;
}

// ---------------------------------------------------------------------------
// zip 辅助
// ---------------------------------------------------------------------------

function openZip(filePath: string): Promise<yauzl.ZipFile> {
  return new Promise((resolve, reject) => {
    yauzl.open(filePath, { lazyEntries: true, autoClose: false }, (err, zip) => {
      if (err) reject(err);
      else resolve(zip);
    });
  });
}

function openEntryStream(
  zip: yauzl.ZipFile,
  entry: yauzl.Entry,
): Promise<Readable> {
  return new Promise((resolve, reject) => {
    zip.openReadStream(entry, (err, stream) => {
      if (err) reject(err);
      else resolve(stream);
    });
  });
}

/**
 * 打开 zip 边遍历边匹配 entry；命中后立即 openReadStream（yauzl 要求
 * openReadStream 必须在 entry 事件回调中调用，否则流报 closed），
 * 交给 fn 处理；完成后关闭 zip。返回是否找到。
 */
async function withEntry(
  filePath: string,
  entryName: string,
  fn: (rs: Readable, entry: yauzl.Entry) => Promise<void>,
): Promise<boolean> {
  const zip = await openZip(filePath);
  return new Promise<boolean>((resolve, reject) => {
    let settled = false;
    const finish = (err?: Error, found?: boolean) => {
      if (settled) return;
      settled = true;
      try {
        zip.close();
      } catch {
        /* ignore */
      }
      if (err) reject(err);
      else resolve(found ?? false);
    };
    zip.on('error', (e) => finish(e));
    zip.on('entry', (entry) => {
      const name = entry.fileName.replace(/\\/g, '/');
      if (name !== entryName) {
        zip.readEntry();
        return;
      }
      zip.openReadStream(entry, (err, rs) => {
        if (err || !rs) return finish(err ?? new Error('no read stream'));
        fn(rs, entry).then(
          () => finish(undefined, true),
          (e) => finish(e),
        );
      });
    });
    zip.on('end', () => finish(undefined, false));
    zip.readEntry();
  });
}

function normalizeTarget(t: string): string {
  let s = t.replace(/\\/g, '/');
  if (s.startsWith('/')) s = s.slice(1);
  if (!s.startsWith('xl/')) s = `xl/${s}`;
  return s;
}

/** 通用流式 XML 解析（sax parser，非 strict 以容错；tag/attr 名统一转小写） */
function parseXmlStream(rs: Readable, handlers: XmlHandlers): Promise<void> {
  return new Promise((resolve, reject) => {
    const parser = sax.parser(false, {});
    parser.onopentag = (node) => {
      const attrs: Record<string, string> = {};
      const raw = node.attributes as Record<string, unknown>;
      for (const k of Object.keys(raw)) {
        attrs[k.toLowerCase()] = raw[k] == null ? '' : String(raw[k]);
      }
      handlers.onOpenTag?.(node.name.toLowerCase(), attrs);
    };
    parser.onclosetag = (name) => handlers.onCloseTag?.(name.toLowerCase());
    parser.ontext = (t) => handlers.onText?.(t);
    parser.oncdata = (t) => handlers.onText?.(t);
    parser.onend = () => resolve();
    parser.onerror = (e) => reject(e instanceof Error ? e : new Error(String(e)));
    rs.on('error', (e) => reject(e));
    rs.on('data', (chunk: Buffer) => {
      parser.write(typeof chunk === 'string' ? chunk : chunk.toString('utf8'));
    });
    rs.on('end', () => parser.end());
  });
}

function colIndexFromRef(ref: string | undefined): number | null {
  if (!ref) return null;
  const m = /^([A-Z]+)/i.exec(ref);
  if (!m) return null;
  let n = 0;
  for (const ch of m[1].toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw new AbortError();
}

// ---------------------------------------------------------------------------
// 主读取器
// ---------------------------------------------------------------------------

export class XlsxStreamReader {
  private prepared?: PreparedMeta;
  private sstPromise: Promise<SpilledStringArray | null> | null = null;

  constructor(
    private filePath: string,
    private tmpDir?: string,
  ) {}

  async getSheetNames(): Promise<string[]> {
    const meta = await this.prepare();
    return meta.sheets.map((s) => s.name);
  }

  async dispose(): Promise<void> {
    if (this.sstPromise) {
      const sst = await this.sstPromise;
      if (sst) await sst.dispose();
    }
  }

  private async prepare(): Promise<PreparedMeta> {
    if (this.prepared) return this.prepared;
    const meta = await this.loadMeta();
    this.prepared = meta;
    return meta;
  }

  private async loadMeta(): Promise<PreparedMeta> {
    // workbook.xml → sheet 列表
    const sheets: { name: string; rId: string }[] = [];
    await withEntry(this.filePath, 'xl/workbook.xml', async (rs) => {
      await parseXmlStream(rs, {
        onOpenTag(name, attrs) {
          if (name === 'sheet') {
            const nm = attrs['name'];
            const rid = attrs['r:id'] ?? attrs['rId'];
            if (nm != null) sheets.push({ name: nm, rId: rid ?? '' });
          }
        },
      });
    });

    // workbook.xml.rels → rId → target
    const rels = new Map<string, string>();
    await withEntry(this.filePath, 'xl/_rels/workbook.xml.rels', async (rs) => {
      await parseXmlStream(rs, {
        onOpenTag(name, attrs) {
          if (name === 'relationship') {
            const type = attrs['type'] ?? '';
            if (type.endsWith('/worksheet')) {
              const id = attrs['id'];
              const target = attrs['target'] ?? '';
              if (id && target) rels.set(id, normalizeTarget(target));
            }
          }
        },
      });
    });

    const sheetList = sheets.map((s) => ({ name: s.name, target: rels.get(s.rId) }));
    const cellXfs = await this.loadCellXfs();
    return { sheets: sheetList, cellXfs };
  }

  private async loadCellXfs(): Promise<CellXf[]> {
    const customFmts = new Map<number, string>();
    const xfs: CellXf[] = [];
    let inCellXfs = false;
    await withEntry(this.filePath, 'xl/styles.xml', async (rs) => {
      await parseXmlStream(rs, {
        onOpenTag(name, attrs) {
          if (name === 'numfmt') {
            const id = parseInt(attrs['numfmtid'] ?? '', 10);
            if (!Number.isNaN(id) && attrs['formatcode'] != null) {
              customFmts.set(id, attrs['formatcode']);
            }
          } else if (name === 'cellxfs') {
            inCellXfs = true;
          } else if (name === 'xf' && inCellXfs) {
            const id = parseInt(attrs['numfmtid'] ?? '0', 10) || 0;
            const code = customFmts.get(id) ?? null;
            xfs.push({
              numFmtId: id,
              numFmtCode: code,
              isDate: isDateNumFmtId(id, code),
            });
          }
        },
        onCloseTag(name) {
          if (name === 'cellxfs') inCellXfs = false;
        },
      });
    });
    return xfs;
  }

  private getXf(s: string | undefined): CellXf {
    if (s == null) return DEFAULT_XF;
    const i = parseInt(s, 10);
    if (Number.isNaN(i) || !this.prepared) return DEFAULT_XF;
    return this.prepared.cellXfs[i] ?? DEFAULT_XF;
  }

  // -------------------------------------------------------------------------
  // sharedStrings
  // -------------------------------------------------------------------------

  /** 懒加载 sharedStrings（有界内存：分块 spill 到临时文件）。不存在返回 null */
  getSharedStrings(signal?: AbortSignal): Promise<SpilledStringArray | null> {
    if (!this.sstPromise) {
      this.sstPromise = this.loadSharedStrings(signal);
    }
    return this.sstPromise;
  }

  private async loadSharedStrings(signal?: AbortSignal): Promise<SpilledStringArray | null> {
    const arr = new SpilledStringArray(this.tmpDir);
    let found = false;
    await withEntry(this.filePath, 'xl/sharedStrings.xml', async (rs) => {
      found = true;
      await new Promise<void>((resolve, reject) => {
        const parser = sax.parser(false, {});
        let inT = false;
        let cur = '';
        let paused = false;
        let finished = false;
        const finish = (err?: Error) => {
          if (finished) return;
          finished = true;
          if (err) reject(err);
          else resolve();
        };
        const abortHandler = () => {
          finished = true;
          reject(new AbortError());
          rs.destroy();
        };
        signal?.addEventListener('abort', abortHandler, { once: true });
        parser.onopentag = (node) => {
          if (node.name.toLowerCase() === 't') {
            inT = true;
            cur = '';
          }
        };
        parser.ontext = (t) => {
          if (inT) cur += t;
        };
        parser.oncdata = (t) => {
          if (inT) cur += t;
        };
        parser.onclosetag = (name) => {
          if (name.toLowerCase() === 't' && inT) {
            inT = false;
            // 同步入队；仅当触发落盘时才产生微任务
            const text = cur;
            cur = '';
            arr.push(text);
            if (paused) {
              paused = false;
              rs.resume();
            }
          }
        };
        parser.onend = () => finish();
        parser.onerror = (e) => finish(e instanceof Error ? e : new Error(String(e)));
        rs.on('error', (e) => finish(e));
        rs.on('data', (chunk: Buffer) => {
          if (finished) return;
          rs.pause();
          paused = true;
          try {
            parser.write(typeof chunk === 'string' ? chunk : chunk.toString('utf8'));
          } catch (e) {
            finish(e instanceof Error ? e : new Error(String(e)));
            return;
          }
          if (paused) paused = false;
          if (!finished) rs.resume();
        });
        rs.on('end', () => parser.end());
      });
    });
    if (!found) {
      await arr.dispose();
      return null;
    }
    await arr.finalize();
    return arr;
  }

  // -------------------------------------------------------------------------
  // sheet 流式解析
  // -------------------------------------------------------------------------

  /** 解析指定 sheet 到目标对象；sheet 不存在抛 SheetNotFoundError */
  async streamSheet(
    sheet: string,
    opts: StreamSheetOptions,
  ): Promise<StreamSheetResult> {
    throwIfAborted(opts.signal);
    const meta = await this.prepare();
    const info = meta.sheets.find((s) => s.name === sheet);
    if (!info || !info.target) {
      throw new SheetNotFoundError(sheet, meta.sheets.map((s) => s.name));
    }
    const sst = await this.getSharedStrings(opts.signal);
    throwIfAborted(opts.signal);

    let rowsRead = 0;
    let endedEarly = false;

    // 解析 sheet：以 withEntry 打开目标条目（openReadStream 在 entry 事件内调用）
    const parseSheet = async (rs: Readable): Promise<void> => {
      await new Promise<void>((resolve, reject) => {
        const parser = sax.parser(false, {});
        let curRow: RowCtx | null = null;
        let curCell: {
          col: number;
          t?: string;
          s?: string;
          text: string;
        } | null = null;
        let inV = false;
        let inT = false;
        let lastCol = -1;
        let stop = false;
        let finished = false;
        // 行链：串行化逐行产出（SST 解析为异步，必须保证输出行序）
        let rowChain: Promise<void> = Promise.resolve();

        const finish = (err?: Error) => {
          if (finished) return;
          if (err) {
            finished = true;
            reject(err);
            return;
          }
          // 成功路径：等行链全部产出完毕再 resolve，避免末块行丢失/竞争
          rowChain
            .then(() => {
              if (finished) return;
              finished = true;
              resolve();
            })
            .catch(() => {
              if (finished) return;
              finished = true;
              resolve();
            });
        };

        const abortHandler = () => {
          stop = true;
          finish(new AbortError());
          rs.destroy();
        };
        opts.signal?.addEventListener('abort', abortHandler, { once: true });

        const emitRow = (row: StreamRow): boolean => {
          throwIfAborted(opts.signal);
          if (opts.skipFirstRow && rowsRead === 0) {
            // 跳过表头行：不计入 rowsRead，不交给 rowHandler
            rowsRead++;
            return true;
          }
          const cap = opts.headerOnly ? 1 : opts.maxRows ?? Infinity;
          if (rowsRead >= cap) return false;
          rowsRead++;
          return opts.rowHandler(row) !== false;
        };

        const finalizeCell = (
          cell: { t?: string; s?: string; text: string },
        ): CellValue | { __sst: number } => {
          const t = cell.t;
          const txt = cell.text;
          if (t === 's') {
            const idx = parseInt(txt, 10);
            if (Number.isNaN(idx)) return null;
            return { __sst: idx };
          }
          if (t === 'b') return txt === '1';
          if (t === 'inlineStr') return txt === '' ? null : txt;
          if (t === 'str') return txt === '' ? null : txt;
          if (t === 'd') return txt === '' ? null : txt;
          if (t === 'e') return null;
          if (txt === '') return null;
          const num = Number(txt);
          if (Number.isNaN(num)) return txt; // 未标类型的文本容错
          const xf = this.getXf(cell.s);
          if (xf.isDate) return excelSerialToIso(num);
          return num;
        };

        const handleRow = (rowCtx: RowCtx): void => {
          // 串行化：本行必须在前一行产出完成后才产出
          const p = rowChain.then(async () => {
            if (stop || finished) return;
            const cells: StreamCell[] = [];
            for (const c of rowCtx.cells) {
              if (c.value && typeof c.value === 'object' && '__sst' in c.value) {
                cells.push({ col: c.col, value: sst ? await sst.get(c.value.__sst) : null });
              } else {
                cells.push({ col: c.col, value: c.value });
              }
            }
            if (stop || finished) return;
            const cont = emitRow({ row: rowCtx.rowNum, cells });
            if (cont === false) {
              stop = true;
              endedEarly = true;
              finish();
              rs.destroy();
            }
          });
          rowChain = p.catch((e: Error) => {
            stop = true;
            finish(e);
            rs.destroy();
          });
        };

        parser.onopentag = (node) => {
          const name = node.name.toLowerCase();
          const attrs = node.attributes as Record<string, string>;
          const rawAttrs: Record<string, string> = {};
          for (const k of Object.keys(attrs)) {
            rawAttrs[k.toLowerCase()] = attrs[k];
          }
          if (name === 'row') {
            const r = parseInt(rawAttrs['r'] ?? '', 10);
            curRow = { rowNum: Number.isNaN(r) ? 0 : r, cells: [] };
          } else if (name === 'c' && curRow) {
            const col = colIndexFromRef(rawAttrs['r']);
            const c = col == null ? lastCol + 1 : col;
            lastCol = c;
            curCell = { col: c, t: rawAttrs['t'], s: rawAttrs['s'], text: '' };
            inV = false;
            inT = false;
          } else if (name === 'v' && curCell) {
            inV = true;
          } else if (name === 't' && curCell) {
            inT = true;
          }
        };
        parser.ontext = (text) => {
          if (curCell && (inV || inT)) curCell.text += text;
        };
        parser.oncdata = (text) => {
          if (curCell && (inV || inT)) curCell.text += text;
        };
        parser.onclosetag = (rawName) => {
          const name = rawName.toLowerCase();
          if (name === 'v') inV = false;
          else if (name === 't') inT = false;
          else if (name === 'c' && curCell && curRow) {
            curRow.cells.push({ col: curCell.col, value: finalizeCell(curCell) });
            curCell = null;
          } else if (name === 'row' && curRow) {
            const row = curRow;
            curRow = null;
            // rowNum 缺省时按出现顺序补
            if (!row.rowNum) row.rowNum = rowsRead + 1;
            handleRow(row);
          }
        };
        parser.onend = () => finish();
        parser.onerror = (e) => {
          if (!stop) finish(e instanceof Error ? e : new Error(String(e)));
        };
        rs.on('error', (e) => {
          if (!stop) finish(e);
        });
        rs.on('data', (chunk: Buffer) => {
          if (stop || finished) return;
          rs.pause();
          try {
            parser.write(typeof chunk === 'string' ? chunk : chunk.toString('utf8'));
          } catch (e) {
            finish(e instanceof Error ? e : new Error(String(e)));
            return;
          }
          if (finished || stop) return;
          // 等本块所有行（含 SST 异步解析）串行产出后再恢复源流
          const tail = rowChain.then(() => {
            if (!finished && !stop) rs.resume();
          });
          tail.catch(() => {
            /* rowChain 自身已捕获错误 */
          });
        });
        rs.on('end', () => parser.end());
      });
    };

    const found = await withEntry(
      this.filePath,
      info.target ?? '',
      parseSheet,
    );
    if (!found) {
      throw new SheetNotFoundError(sheet, meta.sheets.map((s) => s.name));
    }
    return { rowsRead, endedEarly };
  }

  /** 便捷：取第一行（表头），返回 CellValue 数组 */
  async getHeaderRow(sheet: string, signal?: AbortSignal): Promise<StreamRow | null> {
    let header: StreamRow | null = null;
    await this.streamSheet(sheet, {
      headerOnly: true,
      signal,
      rowHandler: (row) => {
        header = row;
        return false;
      },
    });
    return header;
  }
}
