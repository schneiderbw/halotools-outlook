import { getAccessToken, refresh, NotAuthenticatedError } from "./auth";
import { getConfig, getTokens } from "./config";
import { storage } from "./storage";
import type {
  HaloClient,
  HaloUser,
  HaloTicket,
  HaloAction,
  HaloTicketType,
  HaloStatus,
  HaloAgent,
  CreateTicketPayload,
  CreateActionPayload,
  CreateContactPayload,
  UpdateTicketPayload,
} from "../types/halo";

class HaloApiError extends Error {
  constructor(public status: number, public body: string) {
    super(`Halo API ${status}: ${body}`);
    this.name = "HaloApiError";
  }
}

async function call<T>(path: string, init: RequestInit = {}, retried = false): Promise<T> {
  const cfg = getConfig();
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
    // Fetch only throws on network-level failures (CORS preflight rejection, offline, DNS, TLS).
    // Surface the most likely cause first since CORS misconfiguration is the dominant failure mode
    // for SPAs talking to a Halo tenant.
    throw new HaloApiError(
      0,
      `Network call to ${cfg.haloBaseUrl} failed. Most common cause: the add-in's origin (https://tools.iusehalo.com) is not on this Halo Connect app's CORS allowed origins list. Original error: ${(e as Error).message}`,
    );
  }

  // 401 → one retry after forced refresh.
  if (res.status === 401 && !retried) {
    const tokens = getTokens();
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

/** Broad user search for the manual picker — returns all matches. */
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

/** Broad client search for the manual picker. */
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

export async function findTicketsByConversationId(
  conversationId: string,
  customFieldName = "CFOutlookConversationId",
): Promise<HaloTicket[]> {
  const q = new URLSearchParams({
    [`field_${customFieldName}`]: conversationId,
    pageinate: "false",
    includedetails: "true",
  });
  const res = await call<{ tickets: HaloTicket[] } | HaloTicket[]>(`/Tickets?${q}`);
  const all = Array.isArray(res) ? res : res.tickets;
  // Halo silently ignores filter params for custom fields it doesn't recognize
  // (or that no ticket has populated yet), returning every ticket in the tenant.
  // Verify each ticket actually carries the matching custom field value before
  // claiming it belongs to this conversation.
  return all.filter((t) =>
    t.customfields?.some(
      (cf) => cf.name === customFieldName && String(cf.value) === conversationId,
    ),
  );
}

// ---------- Reference data (cached in-memory for the session) ----------

let _ticketTypesCache: HaloTicketType[] | undefined;
let _agentsCache: HaloAgent[] | undefined;
let _statusesCache: HaloStatus[] | undefined;

export async function listTicketTypes(force = false): Promise<HaloTicketType[]> {
  if (_ticketTypesCache && !force) return _ticketTypesCache;
  const res = await call<{ tickettypes: HaloTicketType[] } | HaloTicketType[]>(
    "/TicketType?includeinactive=false",
  );
  _ticketTypesCache = (Array.isArray(res) ? res : res.tickettypes).filter((t) => !t.inactive);
  return _ticketTypesCache;
}

export async function listAgents(force = false): Promise<HaloAgent[]> {
  if (_agentsCache && !force) return _agentsCache;
  const res = await call<{ agents: HaloAgent[] } | HaloAgent[]>(
    "/Agent?includeinactive=false",
  );
  _agentsCache = (Array.isArray(res) ? res : res.agents).filter((a) => !a.inactive);
  return _agentsCache;
}

export async function listStatuses(force = false): Promise<HaloStatus[]> {
  if (_statusesCache && !force) return _statusesCache;
  const res = await call<{ statuses: HaloStatus[] } | HaloStatus[]>(
    "/Status?includeinactive=false",
  );
  _statusesCache = (Array.isArray(res) ? res : res.statuses).filter((s) => !s.inactive);
  return _statusesCache;
}

export function clearReferenceCache() {
  _ticketTypesCache = undefined;
  _agentsCache = undefined;
  _statusesCache = undefined;
}

// ---------- Current user → Halo agent ----------

const CURRENT_AGENT_KEY = "halo.currentAgentId.v1";

/**
 * Resolve the current Outlook user to a Halo agent. Cached in storage on success.
 * Returns undefined if no agent has a matching email — surfaces a non-fatal "Assign to me unavailable".
 */
export async function getCurrentAgent(
  outlookEmail: string,
): Promise<HaloAgent | undefined> {
  const cachedId = storage().get<number>(CURRENT_AGENT_KEY);
  if (cachedId) {
    const agents = await listAgents();
    const cached = agents.find((a) => a.id === cachedId);
    if (cached) return cached;
  }
  const agents = await listAgents();
  const matched = agents.find(
    (a) => a.email?.toLowerCase() === outlookEmail.toLowerCase(),
  );
  if (matched) await storage().set(CURRENT_AGENT_KEY, matched.id);
  return matched;
}

// ---------- Write paths ----------

export async function appendAction(payload: CreateActionPayload): Promise<HaloAction> {
  const res = await call<HaloAction[]>("/Actions", {
    method: "POST",
    body: JSON.stringify([payload]),
  });
  return res[0];
}

export async function createTicket(payload: CreateTicketPayload): Promise<HaloTicket> {
  const res = await call<HaloTicket[]>("/Tickets", {
    method: "POST",
    body: JSON.stringify([payload]),
  });
  return res[0];
}

/** Apply a partial update to an existing ticket (status / agent / priority / custom fields). */
export async function updateTicket(payload: UpdateTicketPayload): Promise<HaloTicket> {
  const res = await call<HaloTicket[]>("/Tickets", {
    method: "POST",
    body: JSON.stringify([payload]),
  });
  return res[0];
}

/**
 * Stamp the Outlook conversation + message IDs on a ticket so future emails
 * in the same conversation can find it via findTicketsByConversationId.
 * No-ops if either ID is missing. Errors are non-fatal — the caller's primary
 * action (e.g., appending an email) already succeeded.
 */
export async function stampOutlookThreadFields(
  ticketId: number,
  conversationId: string | undefined,
  internetMessageId: string | undefined,
): Promise<void> {
  const customfields: Array<{ name: string; value: string }> = [];
  if (conversationId) {
    customfields.push({ name: "CFOutlookConversationId", value: conversationId });
  }
  if (internetMessageId) {
    customfields.push({ name: "CFOutlookInternetMessageId", value: internetMessageId });
  }
  if (customfields.length === 0) return;
  await updateTicket({ id: ticketId, customfields });
}

/** Full client record — includes assigned account manager and other fields not in list results. */
export async function getClientDetails(clientId: number): Promise<HaloClient> {
  return await call<HaloClient>(`/Client/${clientId}`);
}

/**
 * Asynchronous stats for the contact dossier: open ticket count and last activity time.
 * Both calls are best-effort — any failure degrades gracefully to a zero count so the
 * dossier still renders the rest of its data.
 */
export async function getContactStats(
  userId: number,
): Promise<{ openTicketCount: number; lastActivityAt?: string }> {
  let openTicketCount = 0;
  let lastActivityAt: string | undefined;

  try {
    const q = new URLSearchParams({
      user_id: String(userId),
      open_only: "true",
      count: "true",
      pageinate: "false",
    });
    const res = await call<{ count?: number; tickets?: HaloTicket[] } | HaloTicket[]>(
      `/Tickets?${q}`,
    );
    if (Array.isArray(res)) {
      openTicketCount = res.length;
    } else if (typeof res.count === "number") {
      openTicketCount = res.count;
    } else if (Array.isArray(res.tickets)) {
      openTicketCount = res.tickets.length;
    }
  } catch {
    /* swallow — stats are decorative */
  }

  try {
    const q = new URLSearchParams({
      user_id: String(userId),
      count: "1",
      orderbydesc: "datetime",
      pageinate: "false",
    });
    const res = await call<{ actions?: HaloAction[] } | HaloAction[]>(`/Actions?${q}`);
    const arr = Array.isArray(res) ? res : res.actions ?? [];
    lastActivityAt = arr[0]?.datetime;
  } catch {
    /* swallow — stats are decorative */
  }

  return { openTicketCount, lastActivityAt };
}

/** Create a new contact (HaloPSA "user"). Mirrors createTicket's array-wrapped POST shape. */
export async function createContact(payload: CreateContactPayload): Promise<HaloUser> {
  const res = await call<HaloUser[]>("/Users", {
    method: "POST",
    body: JSON.stringify([payload]),
  });
  return res[0];
}

export { HaloApiError };
