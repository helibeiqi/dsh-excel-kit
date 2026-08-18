/**
 * dsh-excel-kit — SpillStore 适配器
 *
 * ctx.spillStore 是 dsh-spill 提供的"超大工具文本持久化接缝"。
 * 工具结果序列化超过阈值时调用 saveText 落盘；value 仅保留紧凑摘要 + spilled 元信息。
 *
 * 兼容策略：typeof ctx.spillStore?.saveText === 'function' 才使用真实实现；
 * 否则回退到内存实现（同接口），保证单测/无 dsh 运行时也能工作。
 */
import { formatBytes } from './utils';

export const SPILL_THRESHOLD = 32 * 1024; // 32KB

export interface SpilledRef {
  locator: string;
  bytes: number;
  retrievalHint: string;
}

export interface SpillResult {
  spilled: SpilledRef | null;
  summary: string;
  // 若非空，表示结果未溢出，summary 即完整正文
  full?: boolean;
}

export interface Spillable {
  spillStore: {
    saveText: (input: {
      owner: { sessionId: string };
      source: { toolName: string; callId: string; label: string };
      suggestedName: string;
      content: string;
    }) => Promise<{ locator: string; bytes: number; retrievalHint: string }>;
  };
}

/** 内存回退实现（仅用于无 ctx.spillStore 的环境，如单测） */
class InMemorySpillStore {
  private items = new Map<string, { content: string; bytes: number }>();
  private seq = 0;

  async saveText(input: {
    owner: { sessionId: string };
    source: { toolName: string; callId: string; label: string };
    suggestedName: string;
    content: string;
  }): Promise<{ locator: string; bytes: number; retrievalHint: string }> {
    this.seq += 1;
    const locator = `mem://${input.suggestedName}-${this.seq}`;
    const bytes = Buffer.byteLength(input.content, 'utf8');
    this.items.set(locator, { content: input.content, bytes });
    return {
      locator,
      bytes,
      retrievalHint: `(in-memory fallback) content stored under locator ${locator}`,
    };
  }
}

/** 从 exec.agent / ctx.session 尽力解析 sessionId（ctx 服务经 get() 取，避免 inject 报错） */
export function resolveSessionId(
  exec: { agent?: { sessionId?: string } | null } | undefined,
  ctx:
    | { get?: (name: string) => unknown; session?: { sessionId?: string } }
    | undefined,
): string {
  const fromAgent = exec?.agent?.sessionId;
  if (fromAgent) return fromAgent;
  const session = (ctx?.get?.('session') as { sessionId?: string } | undefined) ?? ctx?.session;
  const fromCtx = session?.sessionId;
  if (fromCtx) return fromCtx;
  return 'dsh-excel-kit';
}

/**
 * 结果溢出处理器：
 *  - 序列化内容 ≤ 阈值：原样返回（full=true）
 *  - 超阈值：saveText 落盘，返回紧凑摘要 + spilled 引用
 */
export async function maybeSpill(args: {
  toolName: string;
  callId: string;
  exec: { agent?: { sessionId?: string } | null };
  ctx: { session?: { sessionId?: string } } | { spillStore?: unknown };
  store?: unknown;
  content: string;
  summary: string;
}): Promise<SpillResult> {
  const bytes = Buffer.byteLength(args.content, 'utf8');
  if (bytes <= SPILL_THRESHOLD) {
    return { spilled: null, summary: args.content, full: true };
  }
  const store = (args.store ?? (args.ctx as { spillStore?: unknown }).spillStore) as
    | Spillable['spillStore']
    | undefined;
  if (store && typeof store.saveText === 'function') {
    const saved = await store.saveText({
      owner: { sessionId: resolveSessionId(args.exec, args.ctx as { session?: { sessionId?: string } }) },
      source: { toolName: args.toolName, callId: args.callId, label: args.toolName },
      suggestedName: `${args.toolName}-${args.callId || 'unknown'}`,
      content: args.content,
    });
    return {
      spilled: {
        locator: saved.locator,
        bytes: saved.bytes,
        retrievalHint: saved.retrievalHint,
      },
      summary: `${args.summary} [spilled ${formatBytes(saved.bytes)} to spillStore]`,
    };
  }
  // 无真实 spillStore：内存回退（内容仍驻留，但接口一致）
  const mem = new InMemorySpillStore();
  const saved = await mem.saveText({
    owner: { sessionId: resolveSessionId(args.exec, args.ctx as { session?: { sessionId?: string } }) },
    source: { toolName: args.toolName, callId: args.callId, label: args.toolName },
    suggestedName: `${args.toolName}-${args.callId || 'unknown'}`,
    content: args.content,
  });
  return {
    spilled: {
      locator: saved.locator,
      bytes: saved.bytes,
      retrievalHint: saved.retrievalHint,
    },
    summary: `${args.summary} [spilled ${formatBytes(saved.bytes)} (in-memory fallback)]`,
  };
}
