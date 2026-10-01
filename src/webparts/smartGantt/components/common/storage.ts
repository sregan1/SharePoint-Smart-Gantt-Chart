// localStorage helpers for per-viewer conveniences (remembered sort, filters,
// collapsed columns). Storage can be unavailable or throw (private windows,
// blocked site data), so every access is wrapped and the UI must render
// correctly without it.

export function readStored<T>(key: string, fallback: T): T {
  try {
    const raw = window.localStorage.getItem(key);
    if (raw === null) return fallback;
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

export function writeStored(key: string, value: unknown): void {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Ignore — persistence is best-effort.
  }
}
