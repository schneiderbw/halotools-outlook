// Per-request auth context. Stdio mode populates this once at startup from env
// vars; HTTP mode populates it per request from the Authorization header.
//
// Implemented with AsyncLocalStorage so tool handlers can read it without
// having auth threaded through every signature.

import { AsyncLocalStorage } from "node:async_hooks";
import type { HaloAuth } from "./client.js";

const als = new AsyncLocalStorage<HaloAuth>();

export function withHaloAuth<T>(auth: HaloAuth, fn: () => Promise<T>): Promise<T> {
  return als.run(auth, fn);
}

export function getHaloAuth(): HaloAuth {
  const auth = als.getStore();
  if (!auth) {
    throw new Error(
      "No Halo auth in current context. In stdio mode, set HALO_BASE_URL and either HALO_ACCESS_TOKEN or (HALO_CLIENT_ID + HALO_CLIENT_SECRET). In HTTP mode, send Authorization: Bearer halo:<base-url>:<token>.",
    );
  }
  return auth;
}

/**
 * Parse the bearer token format used by the hosted HTTP transport.
 *
 * Format:
 *   halo:<base-url>:<access-token>
 *
 * Examples:
 *   halo:https://acme.halopsa.com:eyJhbGciOi...
 *
 * Or with client credentials (less common over HTTP since it implies the client
 * has the client_secret, but supported for completeness):
 *   haloc:<base-url>:<client-id>:<client-secret>
 *
 * Both URL and token parts may contain `:` (URL has `https://`), so we anchor
 * on the prefix and the next two known separators rather than splitting blindly.
 */
export function parseBearerToken(token: string): HaloAuth {
  if (token.startsWith("halo:")) {
    // halo:<base-url>:<access-token>
    // The base URL itself contains ":" (https://), so we find the LAST ":"
    // before what we expect to be the token. Heuristic: a Halo URL ends in
    // .com / .net / .co.uk / etc., and is followed by an optional path. We
    // split on the first ":" that comes after a "//" + at least one ".".
    const rest = token.slice("halo:".length);
    const sep = findUrlTokenBoundary(rest);
    if (sep < 0) throw new Error("Malformed bearer token: expected halo:<base-url>:<access-token>");
    return {
      baseUrl: rest.slice(0, sep),
      accessToken: rest.slice(sep + 1),
    };
  }
  if (token.startsWith("haloc:")) {
    const rest = token.slice("haloc:".length);
    const sep1 = findUrlTokenBoundary(rest);
    if (sep1 < 0) throw new Error("Malformed bearer token: expected haloc:<base-url>:<client-id>:<client-secret>");
    const baseUrl = rest.slice(0, sep1);
    const credsPart = rest.slice(sep1 + 1);
    const sep2 = credsPart.indexOf(":");
    if (sep2 < 0) throw new Error("Malformed bearer token: expected client-id:client-secret after base URL");
    return {
      baseUrl,
      clientId: credsPart.slice(0, sep2),
      clientSecret: credsPart.slice(sep2 + 1),
    };
  }
  throw new Error(
    "Unsupported bearer token format. Expected halo:<base-url>:<access-token> or haloc:<base-url>:<client-id>:<client-secret>.",
  );
}

/**
 * Find the boundary between a base URL and what follows.
 * Walks past the scheme, past the host (next `/` or `:` after the host portion).
 */
function findUrlTokenBoundary(s: string): number {
  // Expect s to start with http:// or https://
  const schemeEnd = s.indexOf("://");
  if (schemeEnd < 0) {
    // Not a URL — just split on first `:`.
    return s.indexOf(":");
  }
  // After scheme, scan for the next `:` that delimits the trailing token.
  // The host may have an optional path (after `/`). The URL ends at the
  // last `/` (no trailing slash) or at the `:` directly following the host.
  // We pick: the first `:` that comes AFTER the first character that follows
  // the scheme — but only counting characters after the scheme separator.
  let i = schemeEnd + 3;
  // Walk through host + optional path. The URL portion ends at a `:` that is
  // not followed by `//` (i.e. not another scheme — but Halo URLs don't nest).
  while (i < s.length) {
    if (s[i] === ":") return i;
    i++;
  }
  return -1;
}

/** Load auth from process env (stdio mode). Returns undefined if not configured. */
export function loadEnvAuth(): HaloAuth | undefined {
  const baseUrl = process.env.HALO_BASE_URL;
  if (!baseUrl) return undefined;
  const accessToken = process.env.HALO_ACCESS_TOKEN;
  if (accessToken) return { baseUrl, accessToken };
  const clientId = process.env.HALO_CLIENT_ID;
  const clientSecret = process.env.HALO_CLIENT_SECRET;
  if (clientId && clientSecret) {
    return {
      baseUrl,
      clientId,
      clientSecret,
      scope: process.env.HALO_SCOPE,
    };
  }
  return undefined;
}
