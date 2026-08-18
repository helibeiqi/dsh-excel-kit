/**
 * dsh-excel-kit — ExcelKitService
 *
 * 在 ctx.tools 上注册 excel_describe / excel_filter / excel_pivot 三个只读工具。
 *
 * 契约（与 dsh-tools rc.6 registry 对齐，实测自 dsh-beisen 插件）：
 *  - parameters 为 MCP 风格顶层 JSON Schema（type:'object' / properties / required）
 *  - execute 返回 { content: ContentBlock[], structuredContent: <result> }
 *  - output.schema 固定声明 content + structuredContent 形态
 *  - 结果序列化超阈值时经 ctx.spillStore 持久化，structuredContent 变为 spill 元信息
 */
import {
  describeParameters,
  executeDescribe,
  formatDescribe,
  DescribeArgs,
} from './tools/describe';
import {
  filterParameters,
  executeFilter,
  formatFilter,
  FilterArgs,
  FilterResult,
} from './tools/filter';
import {
  pivotParameters,
  executePivot,
  formatPivot,
  PivotArgs,
  PivotResult,
} from './tools/pivot';
import { DescribeResult } from './stream/types';
import { maybeSpill, SpilledRef } from './spill';

// ---------------------------------------------------------------------------
// 局部结构类型（与 dsh 宿主实际注入的类型结构兼容，宿主运行时注入真实实现）
// ---------------------------------------------------------------------------

export interface ContentBlock {
  type: 'text';
  text: string;
}

export interface ToolExec {
  callId?: string;
  name?: string;
  arguments?: Record<string, unknown>;
  signal?: AbortSignal;
  agent?: { sessionId?: string } | null;
}

/** execute 的返回契约：content 供渲染，structuredContent 为结构化结果 */
export interface ToolResult {
  content: ContentBlock[];
  structuredContent: unknown;
}

export interface ToolOutputSpec {
  schema: Record<string, unknown>;
  render: (args: Record<string, unknown>, value: unknown) => ContentBlock[];
  presentationMeta?: (args: Record<string, unknown>, value: unknown) => Record<string, unknown>;
}

export interface DefineToolInput {
  name: string;
  description: string;
  parameters: Record<string, unknown>;
  output: ToolOutputSpec;
  isConcurrencySafe?: () => boolean;
  execute: (args: Record<string, unknown>, exec: ToolExec) => Promise<unknown>;
}

export interface PluginContext {
  tools: {
    register: (tool: unknown) => void;
  };
  /** cordis 可选服务访问（未声明 inject 的服务属性访问会抛错，必须用 get） */
  get?: (name: string) => unknown;
  spillStore?: {
    saveText: (input: {
      owner: { sessionId: string };
      source: { toolName: string; callId: string; label: string };
      suggestedName: string;
      content: string;
    }) => Promise<{ locator: string; bytes: number; retrievalHint: string }>;
  };
  session?: { sessionId?: string };
}

/** 本地 defineTool（宿主 dsh 会注入同构实现；此处做最小适配） */
export function defineTool(input: DefineToolInput): DefineToolInput {
  return input;
}

/** 统一的 output.schema：content + structuredContent 包装 */
const TOOL_OUTPUT_SCHEMA = {
  type: 'object',
  properties: {
    content: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          type: { type: 'string' },
          text: { type: 'string' },
        },
      },
    },
    structuredContent: {},
  },
  required: ['content'],
  additionalProperties: false,
} as const;

/** spill 分支的 structuredContent 形态：spilled 为定位对象（locator/bytes/retrievalHint） */
interface SpilledStructured {
  spilled: SpilledRef;
  tool: string;
  summary: string;
}

function isSpilledValue(v: unknown): v is SpilledStructured {
  return (
    typeof v === 'object' &&
    v !== null &&
    typeof (v as { spilled?: unknown }).spilled === 'object' &&
    (v as { spilled?: unknown }).spilled !== null &&
    typeof (v as { summary?: unknown }).summary === 'string'
  );
}

/**
 * 将值规整为 dsh-tools "canonical lossless JSON" 兼容形态：
 *  - 剔除值为 undefined 的 own property（JSON.stringify 会丢，导致 round-trip 后结构变化）
 *  - NaN / ±Infinity 替换为 null（JSON.stringify 将其序列化为 null，round-trip 不等）
 * 这层防御覆盖所有三个工具，避免单个工具遗漏引发同种 invalid output。
 */
function toLossless(v: unknown): unknown {
  if (v === null) return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v !== 'object') return v;
  if (Array.isArray(v)) return v.map(toLossless);
  const out: Record<string, unknown> = {};
  for (const k of Object.keys(v as Record<string, unknown>)) {
    const val = (v as Record<string, unknown>)[k];
    if (val === undefined) continue;
    out[k] = toLossless(val);
  }
  return out;
}

/**
 * 组装工具结果：序列化超阈值时 spill 到 ctx.spillStore，
 * 否则 structuredContent 携带完整结果、content 携带人类可读渲染。
 */
async function finalizeResult<T>(opts: {
  toolName: string;
  exec: ToolExec;
  ctx: PluginContext;
  value: T;
  summary: string;
  format: (value: T) => string;
}): Promise<ToolResult> {
  // 先做 lossless 规整，避免 undefined 字段 / NaN 触发 dsh 校验失败
  const value = toLossless(opts.value) as T;
  const content = JSON.stringify(value);
  // cordis 可选服务访问：未在 inject 声明的服务属性直接读取会抛
  // "cannot get property without inject"，必须经 ctx.get() 获取（服务不存在时为 undefined）
  const spillStore =
    (opts.ctx.get?.('spillStore') as PluginContext['spillStore'] | undefined) ??
    opts.ctx.spillStore;
  const res = await maybeSpill({
    toolName: opts.toolName,
    callId: opts.exec?.callId ?? '',
    exec: opts.exec,
    ctx: opts.ctx,
    store: spillStore,
    content,
    summary: opts.summary,
  });
  if (res.full) {
    return {
      content: [{ type: 'text', text: opts.format(value) }],
      structuredContent: value,
    };
  }
  return {
    content: [{ type: 'text', text: res.summary }],
    structuredContent: {
      spilled: res.spilled, // SpilledRef：{ locator, bytes, retrievalHint }
      tool: opts.toolName,
      summary: res.summary,
    },
  };
}

export class ExcelKitService {
  constructor(private ctx: PluginContext) {}

  register(): void {
    const { ctx } = this;

    ctx.tools.register(
      defineTool({
        name: 'excel_describe',
        description:
          'Read-only Excel workbook description: sheet list, row/column counts, per-column type distribution, non-empty counts, empty rate, numeric min/max/mean and sample values. Streaming, big-file safe.',
        parameters: describeParameters as Record<string, unknown>,
        output: {
          schema: TOOL_OUTPUT_SCHEMA as Record<string, unknown>,
          render: (_args, value) => (value as ToolResult).content,
          presentationMeta: (_args, value) => {
            const v = value as ToolResult;
            const sc = v?.structuredContent as { sheet?: string; spilled?: unknown } | undefined;
            if (isSpilledValue(sc)) return { title: 'excel_describe (spilled)' };
            return { title: `describe: ${sc?.sheet ?? ''}` };
          },
        },
        isConcurrencySafe: () => true,
        async execute(args, exec) {
          const result = await executeDescribe(
            args as unknown as DescribeArgs,
            exec?.signal,
          );
          return finalizeResult({
            toolName: 'excel_describe',
            exec,
            ctx,
            value: result,
            summary: `describe ${result.sheet}: ${result.totalRows} rows × ${result.totalCols} cols`,
            format: (v) => formatDescribe(v as DescribeResult),
          });
        },
      }),
    );

    ctx.tools.register(
      defineTool({
        name: 'excel_filter',
        description:
          'Read-only Excel row filtering: conditions [{column, op: eq|ne|gt|gte|lt|lte|contains|in|between, value|values}], optional column projection, limit (default 100, hard cap 500). Streaming, big-file safe.',
        parameters: filterParameters as Record<string, unknown>,
        output: {
          schema: TOOL_OUTPUT_SCHEMA as Record<string, unknown>,
          render: (_args, value) => (value as ToolResult).content,
          presentationMeta: (_args, value) => {
            const v = value as ToolResult;
            const sc = v?.structuredContent as { matched?: number; spilled?: unknown } | undefined;
            if (isSpilledValue(sc)) return { title: 'excel_filter (spilled)' };
            return { title: `filter matched ${sc?.matched ?? 0} rows` };
          },
        },
        isConcurrencySafe: () => true,
        async execute(args, exec) {
          const result = await executeFilter(
            args as unknown as FilterArgs,
            exec?.signal,
          );
          return finalizeResult({
            toolName: 'excel_filter',
            exec,
            ctx,
            value: result,
            summary: `filter ${result.sheet}: matched ${result.matched}, returned ${result.returned}`,
            format: (v) => formatFilter(v as FilterResult),
          });
        },
      }),
    );

    ctx.tools.register(
      defineTool({
        name: 'excel_pivot',
        description:
          'Read-only Excel pivot: group by rows columns (multiple allowed, nested key like "dept|grade"), aggregate values [{column, agg: count|sum|mean|min|max}], limit (default 50). Streaming, big-file safe.',
        parameters: pivotParameters as Record<string, unknown>,
        output: {
          schema: TOOL_OUTPUT_SCHEMA as Record<string, unknown>,
          render: (_args, value) => (value as ToolResult).content,
          presentationMeta: (_args, value) => {
            const v = value as ToolResult;
            const sc = v?.structuredContent as { groups?: number; spilled?: unknown } | undefined;
            if (isSpilledValue(sc)) return { title: 'excel_pivot (spilled)' };
            return { title: `pivot ${sc?.groups ?? 0} groups` };
          },
        },
        isConcurrencySafe: () => true,
        async execute(args, exec) {
          const result = await executePivot(
            args as unknown as PivotArgs,
            exec?.signal,
          );
          return finalizeResult({
            toolName: 'excel_pivot',
            exec,
            ctx,
            value: result,
            summary: `pivot ${result.sheet}: ${result.groups} groups`,
            format: (v) => formatPivot(v as PivotResult),
          });
        },
      }),
    );
  }
}
