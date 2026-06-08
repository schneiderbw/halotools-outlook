// Resolve the public origin we live at, so that OAuth metadata and the
// redirect URIs we hand to Claude/Halo all point at a single canonical URL.
//
// Order of precedence:
//   1. MCP_PUBLIC_ORIGIN env var (e.g. https://tools.iusehalo.com) — set this
//      in production so behind-proxy weirdness doesn't matter
//   2. X-Forwarded-Proto + X-Forwarded-Host (set by nginx)
//   3. Host header + assumed http

import type { IncomingMessage } from "node:http";

export function getPublicOrigin(req: IncomingMessage): string {
  const env = process.env.MCP_PUBLIC_ORIGIN;
  if (env) return env.replace(/\/+$/, "");
  const proto = (req.headers["x-forwarded-proto"] as string | undefined) ?? "http";
  const host =
    (req.headers["x-forwarded-host"] as string | undefined) ??
    req.headers.host ??
    "localhost";
  return `${proto}://${host}`;
}

/** The full URL we register with Halo as the redirect_uri. Same for every
 *  tool (Outlook, MCP, future tools) — they're discriminated by the state prefix. */
export function getCallbackUrl(req: IncomingMessage): string {
  return `${getPublicOrigin(req)}/auth/callback`;
}
