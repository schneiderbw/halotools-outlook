// Typed request/response contract between the popup, options, and content
// scripts on one side, and the background service worker on the other.
//
// The service worker is the only context that talks to Halo directly —
// keeping the network code there lets us:
//   - reuse one cached client across multiple content scripts on different tabs
//   - keep refresh-token logic in a single place
//   - avoid duplicating CORS / OAuth surface on every page

import type { HaloUser, HaloTicket } from "./types";

export type Request =
  | { kind: "ping" }
  | { kind: "signIn" }
  | { kind: "signOut" }
  | { kind: "getAuthStatus" }
  | { kind: "lookupEmail"; email: string }
  | { kind: "search"; query: string }
  | { kind: "openInHalo"; path: string };

export interface LookupHit {
  user: HaloUser;
  recentTickets: HaloTicket[];
}

export type Response<R extends Request = Request> = R extends { kind: "ping" }
  ? { ok: true; pong: true }
  : R extends { kind: "signIn" }
  ? { ok: true } | { ok: false; error: string }
  : R extends { kind: "signOut" }
  ? { ok: true }
  : R extends { kind: "getAuthStatus" }
  ? { ok: true; configured: boolean; signedIn: boolean }
  : R extends { kind: "lookupEmail" }
  ? { ok: true; hit: LookupHit | null } | { ok: false; error: string }
  : R extends { kind: "search" }
  ? { ok: true; users: HaloUser[] } | { ok: false; error: string }
  : R extends { kind: "openInHalo" }
  ? { ok: true }
  : never;

export function send<R extends Request>(req: R): Promise<Response<R>> {
  return chrome.runtime.sendMessage(req) as Promise<Response<R>>;
}
