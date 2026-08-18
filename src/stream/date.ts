/**
 * dsh-excel-kit — numFmtId 日期识别
 *
 * 通过 xl/styles.xml 中 cellXfs 的 numFmtId 判断列是否为日期格式。
 * 内置格式 14-22、45-47 为日期；自定义格式 id>=164 且格式串含 y/m/d 关键字的视为日期。
 */

/** 内置日期 numFmtId 集合 */
const BUILTIN_DATE_FMT_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, // 日期/时间
  45, 46, 47, // 长日期 / 短日期等
]);

/** 日期判定：numFmtId 匹配内置，或自定义格式串含日期关键字 */
export function isDateNumFmtId(
  numFmtId: number | null | undefined,
  numFmtCode: string | null | undefined,
): boolean {
  if (numFmtId == null) return false;
  if (BUILTIN_DATE_FMT_IDS.has(numFmtId)) return true;
  // 自定义格式：id>=164 或格式码含 y/m/d
  if (numFmtId >= 164) {
    if (numFmtCode) return containsDateToken(numFmtCode);
    return false;
  }
  if (numFmtCode) return containsDateToken(numFmtCode);
  return false;
}

/** 判断格式串是否含日期 token（yyyy/mm/dd、yy、m月、d日、h:mm 等） */
function containsDateToken(fmt: string): boolean {
  // 去掉引号内文字（如 "年" "月" "日" 是字面量，不应误判；但 yyyy"年"mm"月" 仍含 y/m 数字符号）
  const stripped = fmt.replace(/"[^"]*"/g, '');
  // 日期符号：y / m / d 作为 token（m 至少出现一次且后跟非字母，或 y/d 出现）
  return /(^|[^a-zA-Z])y{1,4}([^a-zA-Z]|$)/i.test(stripped)
    || /(^|[^a-zA-Z])d{1,2}([^a-zA-Z]|$)/i.test(stripped)
    || /(^|[^a-zA-Z])m{1,2}([^a-zA-Z]|$)/i.test(stripped);
}

/** 把 Excel 序列日期数转 ISO 字符串（1900 日期系统，处理虚构 1900-02-29） */
export function excelSerialToIso(serial: number): string {
  if (serial < 0 || !Number.isFinite(serial)) return String(serial);
  const msPerDay = 24 * 60 * 60 * 1000;
  // Excel serial 1 = 1900-01-01，且虚构 1900-02-29（serial 60）。
  // 1900 年 1-2 月（serial<=59）基准偏移 +1；>=61（1900-03-01 起）直接用基准 1899-12-30。
  const corrected = serial < 61 ? serial + 1 : serial;
  // 基准：1899-12-30
  const base = Date.UTC(1899, 11, 30);
  const d = new Date(base + corrected * msPerDay);
  if (isPureDate(serial)) {
    // 纯日期输出 YYYY-MM-DD（仍是 ISO 8601）
    return d.toISOString().slice(0, 10);
  }
  return d.toISOString();
}

/** 判断序列数是否为纯日期（无时间部分） */
export function isPureDate(serial: number): boolean {
  return Number.isInteger(serial);
}
