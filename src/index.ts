/**
 * dsh-excel-kit — dsh 插件入口
 *
 * 遵循 dsh 0.1.0-rc.6 插件范式（dsh-beisen 实证契约）：
 *   name / inject / apply 三个导出（不导出 Config——
 *   cordis loader 会对导出的 Config 调用 validate()，非 schemastery schema
 *   对象会导致启动崩溃；无 Config 时 cordis 直接跳过配置校验）。
 */
import { ExcelKitService, PluginContext } from './service';

export const name = 'dsh-excel-kit';

export const inject = ['tools'] as const;

export function apply(
  ctx: PluginContext,
  config?: { enableTools?: boolean; spillThreshold?: number },
): void {
  // 预留：config.enableTools=false 时跳过注册（当前默认注册）
  if (config?.enableTools === false) return;
  new ExcelKitService(ctx).register();
}

export { ExcelKitService, defineTool } from './service';
export { XlsxStreamReader } from './stream/xlsx-stream-reader';
export { maybeSpill, SPILL_THRESHOLD } from './spill';
export {
  executeDescribe,
  formatDescribe,
} from './tools/describe';
export { executeFilter, formatFilter } from './tools/filter';
export { executePivot, formatPivot } from './tools/pivot';
