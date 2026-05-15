// Cross-runtime diagnostic log.
//
// All Outlook runtimes in this add-in are served from the same origin
// (tools.iusehalo.com) so they share window.localStorage. We deliberately use
// localStorage instead of Office.context.roamingSettings here because the
// roamingSettings bag is loaded ONCE when each runtime starts and never
// auto-refreshes — meaning the task pane reads its own stale in-memory copy
// and can't see writes that the launch-event runtime made after the task pane
// opened. localStorage is synchronous, immediate, and visible across runtimes
// at the same origin.
//
// The launch-event runtime has a parallel ES5 implementation in
// public/launchevent.js — they MUST agree on the storage key and entry shape.

export type LogLevel = "info" | "warn" | "error";

export interface LogEntry {
  /** ISO timestamp the entry was recorded. */
  ts: string;
  level: LogLevel;
  /** Free-form tag identifying which surface wrote the entry (e.g. "on-send", "auth", "api"). */
  source: string;
  /** Short human-readable message. Truncated to MAX_MESSAGE_LEN before storage. */
  message: string;
  /** Optional structured payload. Kept small — entire log is bounded by MAX_BYTES. */
  data?: Record<string, unknown>;
}

export const DIAG_LOG_KEY = "halo.diagLog.v1";
const MAX_ENTRIES = 200;
const MAX_MESSAGE_LEN = 500;
// localStorage has a 5 MB cap per origin — we're nowhere near that, but the
// log isn't worth scaling without bound. Cap at ~64 KB.
const MAX_BYTES = 64_000;

function read(): LogEntry[] {
  try {
    const raw = window.localStorage.getItem(DIAG_LOG_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function write(entries: LogEntry[]): void {
  try {
    window.localStorage.setItem(DIAG_LOG_KEY, JSON.stringify(entries));
  } catch {
    /* swallow — out-of-quota or private mode */
  }
}

function trim(entries: LogEntry[]): LogEntry[] {
  let trimmed = entries.length > MAX_ENTRIES ? entries.slice(-MAX_ENTRIES) : entries;
  while (trimmed.length > 1 && JSON.stringify(trimmed).length > MAX_BYTES) {
    trimmed = trimmed.slice(1);
  }
  return trimmed;
}

export function logEvent(
  level: LogLevel,
  source: string,
  message: string,
  data?: Record<string, unknown>,
): void {
  try {
    const entry: LogEntry = {
      ts: new Date().toISOString(),
      level,
      source,
      message: message.length > MAX_MESSAGE_LEN ? message.slice(0, MAX_MESSAGE_LEN) + "…" : message,
    };
    if (data) entry.data = data;
    write(trim([...read(), entry]));
  } catch {
    // Logging must never throw into caller logic.
  }
}

export function getEvents(): LogEntry[] {
  return read();
}

export function clearEvents(): void {
  try {
    window.localStorage.removeItem(DIAG_LOG_KEY);
  } catch {
    /* swallow */
  }
}

/** Trigger a browser download of the current log as JSON. */
export function downloadEvents(): void {
  const blob = new Blob([JSON.stringify(read(), null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `halo-outlook-diagnostics-${new Date().toISOString().replace(/[:.]/g, "-")}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
