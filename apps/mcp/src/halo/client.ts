// Server-side HaloPSA REST client. Stateless: pass a HaloAuth in to every call.
//
// Unlike the browser version (src/lib/halo-api.ts in the parent repo), there is no
// Office context, no roamingSettings, and no refresh-token loop — the MCP server
// receives an already-valid access token (or client credentials) per request.
//
// Supported auth modes:
//   - { baseUrl, accessToken }                                — bearer-only
//   - { baseUrl, clientId, clientSecret }                     — Halo Connect client_credentials grant
//
// In the second mode we exchange credentials for a token on demand and cache it
// per-(baseUrl, clientId) until expiry.

import type {
  HaloUser,
  HaloClient,
  HaloTicket,
  HaloAction,
  HaloCannedText,
  HaloCannedTextGroup,
  HaloCRMNote,
  HaloFeedItem,
  HaloFeedResponse,
  CreateActionPayload,
  CreateTicketPayload,
  CreateCRMNotePayload,
  CRMScope,
} from "./types.js";

export class HaloApiError extends Error {
  constructor(public status: number, public body: string) {
    super(`Halo API ${status}: ${body}`);
    this.name = "HaloApiError";
  }
}

export interface HaloAuthToken {
  baseUrl: string;
  accessToken: string;
}

export interface HaloAuthClientCreds {
  baseUrl: string;
  clientId: string;
  clientSecret: string;
  /** Optional scope; Halo accepts "all" / "edit:tickets" etc. Defaults to "all". */
  scope?: string;
}

export type HaloAuth = HaloAuthToken | HaloAuthClientCreds;

interface CachedToken {
  accessToken: string;
  expiresAt: number;
}

const tokenCache = new Map<string, CachedToken>();

function cacheKey(a: HaloAuthClientCreds): string {
  return `${a.baseUrl}::${a.clientId}`;
}

async function clientCredentialsToken(a: HaloAuthClientCreds): Promise<string> {
  const key = cacheKey(a);
  const cached = tokenCache.get(key);
  if (cached && cached.expiresAt > Date.now() + 30_000) return cached.accessToken;

  // Halo Connect's token endpoint sits at the auth host, not the API host.
  // Convention: same origin as baseUrl, /auth/token. Some installs nest auth
  // under /authorize — we try /auth/token first which is the documented public path.
  const tokenUrl = `${a.baseUrl.replace(/\/$/, "")}/auth/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: a.clientId,
    client_secret: a.clientSecret,
    scope: a.scope ?? "all",
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new HaloApiError(res.status, `client_credentials grant failed: ${text}`);
  }
  const json = (await res.json()) as { access_token: string; expires_in?: number };
  const expiresIn = json.expires_in ?? 3600;
  tokenCache.set(key, {
    accessToken: json.access_token,
    expiresAt: Date.now() + expiresIn * 1000,
  });
  return json.access_token;
}

async function getBearer(auth: HaloAuth): Promise<string> {
  if ("accessToken" in auth) return auth.accessToken;
  return clientCredentialsToken(auth);
}

async function call<T>(
  auth: HaloAuth,
  path: string,
  init: RequestInit = {},
): Promise<T> {
  const token = await getBearer(auth);
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${token}`);
  if (!headers.has("Accept")) headers.set("Accept", "application/json");
  if (init.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }

  const base = auth.baseUrl.replace(/\/$/, "");
  let res: Response;
  try {
    res = await fetch(`${base}/api${path}`, { ...init, headers });
  } catch (e) {
    throw new HaloApiError(
      0,
      `Network call to ${base} failed: ${(e as Error).message}`,
    );
  }

  // On 401 with client-credentials auth, invalidate and retry once.
  if (res.status === 401 && "clientId" in auth) {
    tokenCache.delete(cacheKey(auth));
    const retryToken = await getBearer(auth);
    const retryHeaders = new Headers(headers);
    retryHeaders.set("Authorization", `Bearer ${retryToken}`);
    res = await fetch(`${base}/api${path}`, { ...init, headers: retryHeaders });
  }

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new HaloApiError(res.status, body);
  }
  if (res.status === 204) return undefined as unknown as T;
  return (await res.json()) as T;
}

// ---------- Read paths ----------

export async function findUserByEmail(
  auth: HaloAuth,
  email: string,
): Promise<HaloUser | undefined> {
  const q = new URLSearchParams({ search: email, includeinactive: "false" });
  const res = await call<{ users: HaloUser[] } | HaloUser[]>(auth, `/Users?${q}`);
  const arr = Array.isArray(res) ? res : res.users;
  return arr.find((u) => u.emailaddress?.toLowerCase() === email.toLowerCase()) ?? arr[0];
}

export async function searchTickets(
  auth: HaloAuth,
  query: string,
  limit = 25,
): Promise<HaloTicket[]> {
  const q = new URLSearchParams({
    search: query,
    pageinate: "false",
    count: String(limit),
  });
  const res = await call<{ tickets: HaloTicket[] } | HaloTicket[]>(
    auth,
    `/Tickets?${q}`,
  );
  return Array.isArray(res) ? res : res.tickets;
}

export async function listOpenTicketsForClient(
  auth: HaloAuth,
  clientId: number,
): Promise<HaloTicket[]> {
  const q = new URLSearchParams({
    client_id: String(clientId),
    open_only: "true",
    pageinate: "false",
    includedetails: "true",
    includeagentdetails: "true",
  });
  const res = await call<{ tickets: HaloTicket[] } | HaloTicket[]>(
    auth,
    `/Tickets?${q}`,
  );
  return Array.isArray(res) ? res : res.tickets;
}

export async function listOpenTicketsForUser(
  auth: HaloAuth,
  userId: number,
): Promise<HaloTicket[]> {
  const q = new URLSearchParams({
    user_id: String(userId),
    open_only: "true",
    pageinate: "false",
    includedetails: "true",
    includeagentdetails: "true",
  });
  const res = await call<{ tickets: HaloTicket[] } | HaloTicket[]>(
    auth,
    `/Tickets?${q}`,
  );
  return Array.isArray(res) ? res : res.tickets;
}

/** Returns count of open tickets for a user. Best-effort, swallows errors. */
export async function getOpenTicketCount(
  auth: HaloAuth,
  userId: number,
): Promise<number> {
  try {
    const tix = await listOpenTicketsForUser(auth, userId);
    return tix.length;
  } catch {
    return 0;
  }
}

// ---------- Canned text ----------

export async function listCannedText(auth: HaloAuth): Promise<HaloCannedText[]> {
  const q = new URLSearchParams({
    showall: "true",
    entity: "0",
    access_control_level: "2",
  });
  const res = await call<HaloCannedText[] | { canned_texts: HaloCannedText[] }>(
    auth,
    `/CannedText?${q}`,
  );
  return Array.isArray(res) ? res : res.canned_texts ?? [];
}

export async function listCannedTextGroups(
  auth: HaloAuth,
): Promise<HaloCannedTextGroup[]> {
  const q = new URLSearchParams({
    lookupid: "45",
    showallcodes: "true",
    access_control_level: "2",
  });
  const res = await call<HaloCannedTextGroup[]>(auth, `/Lookup?${q}`);
  return (Array.isArray(res) ? res : []).filter(
    (g) => g.valueint1 == null || g.valueint1 === 0,
  );
}

export async function searchCannedText(
  auth: HaloAuth,
  query: string,
  groupId?: number,
): Promise<HaloCannedText[]> {
  const all = await listCannedText(auth);
  const scoped = groupId == null ? all : all.filter((c) => c.group_id === groupId);
  const needle = query.trim().toLowerCase();
  if (!needle) return scoped.slice(0, 50);
  return scoped
    .filter((c) => `${c.name ?? ""} ${c.text ?? ""}`.toLowerCase().includes(needle))
    .slice(0, 50);
}

// ---------- Write paths ----------

export async function appendAction(
  auth: HaloAuth,
  payload: CreateActionPayload,
): Promise<HaloAction> {
  const res = await call<HaloAction[]>(auth, "/Actions", {
    method: "POST",
    body: JSON.stringify([payload]),
  });
  return res[0];
}

export async function createTicket(
  auth: HaloAuth,
  payload: CreateTicketPayload,
): Promise<HaloTicket> {
  const res = await call<HaloTicket[]>(auth, "/Tickets", {
    method: "POST",
    body: JSON.stringify([payload]),
  });
  return res[0];
}

// ---------- CRM notes ----------

export async function createCRMNote(
  auth: HaloAuth,
  payload: CreateCRMNotePayload,
): Promise<HaloCRMNote> {
  const res = await call<{ actions?: HaloCRMNote[] } | HaloCRMNote[]>(
    auth,
    "/CRMNote",
    { method: "POST", body: JSON.stringify([payload]) },
  );
  const arr = Array.isArray(res) ? res : res.actions ?? [];
  return arr[0];
}

// ---------- Activity feed ----------

export async function listFeed(
  auth: HaloAuth,
  scope: CRMScope,
  count = 20,
): Promise<HaloFeedItem[]> {
  const q = new URLSearchParams({ count: String(count) });
  if (scope.client_id) q.set("related_client_id", String(scope.client_id));
  if (scope.site_id) q.set("related_site_id", String(scope.site_id));
  if (scope.user_id) q.set("related_user_id", String(scope.user_id));
  const res = await call<HaloFeedResponse | HaloFeedItem[]>(auth, `/Feed?${q}`);
  if (Array.isArray(res)) return res;
  return res.feed ?? [];
}
