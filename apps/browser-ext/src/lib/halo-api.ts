// HaloPSA REST client.
//
// COPIED + ADAPTED from ../../../src/lib/halo-api.ts. The shape is identical;
// the only changes are:
//   - async getConfig() (chrome.storage is promise-based) — so `call` is async
//     at every step, no synchronous `getConfig()` call
//   - the network-error message says "extension" instead of "add-in"
//   - only the read paths we use from the extension are exposed; create/update
//     paths and reference-data caches stay on the Outlook side. If the
//     extension grows to support those, port them over here at that point.
//
// All Halo calls go through `call()` so we get bearer auth, 401-retry-on-
// refresh, and a uniform network-failure message in one place.

import { getAccessToken, refresh, NotAuthenticatedError } from "./auth";
import { getConfig, getTokens } from "./storage";
import type { HaloClient, HaloUser, HaloTicket } from "./types";

class HaloApiError extends Error {
  constructor(public status: number, public body: string) {
    super(`Halo API ${status}: ${body}`);
    this.name = "HaloApiError";
  }
}

async function call<T>(path: string, init: RequestInit = {}, retried = false): Promise<T> {
  const cfg = await getConfig();
  if (!cfg) throw new NotAuthenticatedError("No tenant config");

  const token = await getAccessToken();
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${token}`);
  if (!headers.has("Accept")) headers.set("Accept", "application/json");
  if (init.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }

  let res: Response;
  try {
    res = await fetch(`${cfg.haloBaseUrl}/api${path}`, { ...init, headers });
  } catch (e) {
    // Fetch only throws on network-level failures (CORS preflight rejection,
    // offline, DNS, TLS). Surface the most likely cause first.
    throw new HaloApiError(
      0,
      `Network call to ${cfg.haloBaseUrl} failed. Most common cause: this extension's origin (chrome-extension://<id>) is not on this Halo Connect app's CORS allowed origins list. Original error: ${(e as Error).message}`,
    );
  }

  // 401 → one retry after forced refresh.
  if (res.status === 401 && !retried) {
    const tokens = await getTokens();
    if (tokens) {
      await refresh(tokens.refreshToken);
      return call<T>(path, init, true);
    }
  }

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new HaloApiError(res.status, body);
  }

  if (res.status === 204) return undefined as unknown as T;
  return (await res.json()) as T;
}

// ---------- Read paths ----------

export async function findUserByEmail(email: string): Promise<HaloUser | undefined> {
  const q = new URLSearchParams({ search: email, includeinactive: "false" });
  const res = await call<{ users: HaloUser[] } | HaloUser[]>(`/Users?${q}`);
  const arr = Array.isArray(res) ? res : res.users;
  return arr.find((u) => u.emailaddress?.toLowerCase() === email.toLowerCase()) ?? arr[0];
}

/** Broad user search for the popup picker — returns all matches. */
export async function searchUsers(query: string, limit = 25): Promise<HaloUser[]> {
  const q = new URLSearchParams({
    search: query,
    includeinactive: "false",
    count: String(limit),
  });
  const res = await call<{ users: HaloUser[] } | HaloUser[]>(`/Users?${q}`);
  return Array.isArray(res) ? res : res.users;
}

export async function findClientByDomain(domain: string): Promise<HaloClient | undefined> {
  const q = new URLSearchParams({ search: domain, includeinactive: "false" });
  const res = await call<{ clients: HaloClient[] } | HaloClient[]>(`/Client?${q}`);
  const arr = Array.isArray(res) ? res : res.clients;
  return arr[0];
}

/** Broad client search for the popup picker. */
export async function searchClients(query: string, limit = 25): Promise<HaloClient[]> {
  const q = new URLSearchParams({
    search: query,
    includeinactive: "false",
    count: String(limit),
  });
  const res = await call<{ clients: HaloClient[] } | HaloClient[]>(`/Client?${q}`);
  return Array.isArray(res) ? res : res.clients;
}

export async function listOpenTicketsForClient(clientId: number): Promise<HaloTicket[]> {
  const q = new URLSearchParams({
    client_id: String(clientId),
    open_only: "true",
    pageinate: "false",
  });
  const res = await call<{ tickets: HaloTicket[] } | HaloTicket[]>(`/Tickets?${q}`);
  return Array.isArray(res) ? res : res.tickets;
}

export async function listRecentTicketsForUser(
  userId: number,
  limit = 10,
): Promise<HaloTicket[]> {
  const q = new URLSearchParams({
    user_id: String(userId),
    orderbydesc: "dateoccurred",
    count: String(limit),
    pageinate: "false",
  });
  const res = await call<{ tickets: HaloTicket[] } | HaloTicket[]>(`/Tickets?${q}`);
  return Array.isArray(res) ? res : res.tickets;
}

export { HaloApiError };
