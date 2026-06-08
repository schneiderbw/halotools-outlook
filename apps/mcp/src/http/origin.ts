// Resolve the public origin we live at, so that OAuth metadata and the
// redirect URIs we hand to Claude/Halo all point at a single canonical URL.
//
// MCP clients (Claude, Cursor, ChatGPT desktop) refuse to do OAuth over plain
// HTTP, and the hosted deployment is always behind Cloudflare → EasyPanel →
// nginx → node. So we hardcode the proto to https. Local dev that genuinely
// needs http (e.g. testing against `http://localhost:3001` without TLS) can
// still set MCP_PUBLIC_ORIGIN to override.

import type { IncomingMessage } from "node:http";

export function getPublicOrigin(req: IncomingMessage): string {
  const env = process.env.MCP_PUBLIC_ORIGIN;
  if (env) return env.replace(/\/+$/, "");
  const host =
    (req.headers["x-forwarded-host"] as string | undefined) ??
    req.headers.host ??
    "localhost";
  return `https://${host}`;
}

/** The full URL we register with Halo as the redirect_uri. Same for every
 *  tool (Outlook, MCP, future tools) — they're discriminated by the state prefix. */
export function getCallbackUrl(req: IncomingMessage): string {
  return `${getPublicOrigin(req)}/auth/callback`;
}
