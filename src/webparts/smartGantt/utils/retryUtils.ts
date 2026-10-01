// Shared throttling helpers for SharePoint (PnPjs) and Microsoft Graph calls:
// retry with exponential backoff on 429/503/504 (honoring Retry-After) and a
// small concurrency-limited runner.

/** HTTP status carried by a PnPjs HttpRequestError, a Graph error, or (as a last resort) its message. */
export function getErrorStatus(e: unknown): number | undefined {
  const err = e as { status?: number; statusCode?: number; response?: { status?: number } } | null;
  const status = err?.status ?? err?.statusCode ?? err?.response?.status;
  if (typeof status === 'number') return status;
  const msg = e instanceof Error ? e.message : '';
  const m = /\b(429|503|504)\b/.exec(msg);
  return m ? parseInt(m[1], 10) : undefined;
}

export function isThrottleError(e: unknown): boolean {
  const s = getErrorStatus(e);
  return s === 429 || s === 503 || s === 504;
}

// Retry-After is either delta-seconds or an HTTP date. Returns ms, or undefined.
function retryAfterMs(e: unknown): number | undefined {
  const err = e as {
    response?: { headers?: { get?: (n: string) => string | null } };
    headers?: Record<string, string> | { get?: (n: string) => string | null };
  } | null;
  let raw: string | null | undefined;
  try {
    raw = err?.response?.headers?.get?.('Retry-After');
    if (!raw && err?.headers) {
      const h = err.headers as { get?: (n: string) => string | null } & Record<string, string>;
      raw = typeof h.get === 'function' ? h.get('Retry-After') : (h['retry-after'] || h['Retry-After']);
    }
  } catch { /* header access is best effort */ }
  if (!raw) return undefined;
  const secs = Number(raw);
  if (!isNaN(secs)) return Math.max(0, secs * 1000);
  const when = Date.parse(raw);
  return isNaN(when) ? undefined : Math.max(0, when - Date.now());
}

export interface IRetryOptions {
  /** Retries after the first attempt. Default 4. */
  retries?: number;
  /** First backoff delay in ms (doubles each retry). Default 500. */
  baseDelayMs?: number;
  /** Upper bound for any single wait. Default 20000. */
  maxDelayMs?: number;
}

/** Run `fn`, retrying throttling/transient failures (429/503/504) with backoff. Other errors propagate immediately. */
export async function withRetry<T>(fn: () => Promise<T>, opts: IRetryOptions = {}): Promise<T> {
  const retries = opts.retries ?? 4;
  const base = opts.baseDelayMs ?? 500;
  const cap = opts.maxDelayMs ?? 20000;
  for (let attempt = 0; ; attempt++) {
    try {
      return await fn();
    } catch (e) {
      if (attempt >= retries || !isThrottleError(e)) throw e;
      const backoff = base * Math.pow(2, attempt) * (0.75 + Math.random() * 0.5);
      const wait = Math.min(cap, Math.max(backoff, retryAfterMs(e) ?? 0));
      await new Promise<void>(resolve => setTimeout(resolve, wait));
    }
  }
}

/**
 * Map `items` through `worker` with at most `limit` running at once. Results
 * keep input order. A rejection from `worker` rejects the whole run, so give
 * the worker its own catch when failures should be tolerated.
 */
export async function runLimited<T, R>(
  items: T[],
  limit: number,
  worker: (item: T, index: number) => Promise<R>
): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let next = 0;
  const lane = async (): Promise<void> => {
    while (next < items.length) {
      const i = next++;
      results[i] = await worker(items[i], i);
    }
  };
  const lanes: Promise<void>[] = [];
  for (let i = 0; i < Math.max(1, Math.min(limit, items.length)); i++) lanes.push(lane());
  await Promise.all(lanes);
  return results;
}
