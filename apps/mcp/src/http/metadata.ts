// MCP OAuth discovery metadata.
//
// We act as BOTH the protected resource and the authorization server. The
// "authorization" we do is mostly bouncing Claude through Halo's real OAuth
// page, but to MCP clients we look like a normal RFC 9728 + RFC 8414 OAuth
// server with PKCE.

import type { ServerResponse } from "node:http";

function writeJson(res: ServerResponse, status: number, body: unknown): void {
  const json = JSON.stringify(body);
  res.writeHead(status, {
    "Content-Type": "application/json",
    "Cache-Control": "no-store",
    "Access-Control-Allow-Origin": "*",
  });
  res.end(json);
}

export function emitProtectedResourceMetadata(
  res: ServerResponse,
  issuer: string,
): void {
  writeJson(res, 200, {
    resource: issuer,
    authorization_servers: [issuer],
    scopes_supported: ["all"],
    bearer_methods_supported: ["header"],
  });
}

export function emitAuthorizationServerMetadata(
  res: ServerResponse,
  issuer: string,
): void {
  writeJson(res, 200, {
    issuer,
    authorization_endpoint: `${issuer}/authorize`,
    token_endpoint: `${issuer}/token`,
    registration_endpoint: `${issuer}/register`,
    response_types_supported: ["code"],
    grant_types_supported: ["authorization_code", "refresh_token"],
    code_challenge_methods_supported: ["S256"],
    token_endpoint_auth_methods_supported: ["none"],
    scopes_supported: ["all"],
  });
}

/** Minimal RFC 7591 Dynamic Client Registration — echoes the request back as
 *  an issued client. We don't actually maintain a client registry: every
 *  client_id is accepted on /token, and the only real authorization happens
 *  when the user logs in to Halo. This exists so MCP clients that demand DCR
 *  before continuing don't get stuck. */
export async function handleRegistration(
  body: string,
  res: ServerResponse,
): Promise<void> {
  let parsed: Record<string, unknown> = {};
  try {
    parsed = body ? (JSON.parse(body) as Record<string, unknown>) : {};
  } catch {
    writeJson(res, 400, { error: "invalid_client_metadata" });
    return;
  }
  const clientId = `mcp-client-${Math.random().toString(36).slice(2, 10)}`;
  writeJson(res, 201, {
    ...parsed,
    client_id: clientId,
    client_id_issued_at: Math.floor(Date.now() / 1000),
    token_endpoint_auth_method: "none",
  });
}
