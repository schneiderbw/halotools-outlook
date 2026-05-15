// Cross-runtime diagnostic log.
//
// All Outlook runtimes in this add-in share Office.context.roamingSettings,
// so we use it as a ring-buffered append-only log: the launch-event handler
// can write entries even though its console isn't reachable from the task
// pane, and the Settings screen reads them back for inspection / download.
//
// The launch-event runtime has a parallel ES5 implementation in
// public/launchevent.js — they MUST agree on the storage key and entry shape.

import { storage } from "@iusehalo/halo-api";

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
const MAX_ENTRIES = 100;
const MAX_MESSAGE_LEN = 500;
// roamingSettings has a ~32 KB total cap shared with tokens/config/defaults.
// Keep the log well under that so we never blow out persistence on a write.
const MAX_BYTES = 16_000;

function read(): LogEntry[] {
  try {
    const raw = storage().get<LogEntry[]>(DIAG_LOG_KEY);
    return Array.isArray(raw) ? raw : [];
  } catch {
    return [];
  }
}

function trim(entries: LogEntry[]): LogEntry[] {
  let trimmed = entries.length > MAX_ENTRIES ? entries.slice(-MAX_ENTRIES) : entries;
  // Drop the oldest entries until we're under the byte budget. Most logs never
  // hit this — it's the safety net against a single huge data payload.
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
    const next = trim([...read(), entry]);
    void storage().set(DIAG_LOG_KEY, next);
  } catch {
    // Logging must never throw into caller logic.
  }
}

export function getEvents(): LogEntry[] {
  return read();
}

export async function clearEvents(): Promise<void> {
  try {
    await storage().set(DIAG_LOG_KEY, []);
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
