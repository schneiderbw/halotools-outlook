// Streamable HTTP MCP transport with OAuth bridge to Halo.
//
// URL layout (every endpoint under /mcp/t/<config>/ where <config> encodes
// {halo, clientId}):
//
//   POST   /mcp/t/<config>/                                       JSON-RPC
//   GET    /mcp/t/<config>/.well-known/oauth-protected-resource   metadata
//   GET    /mcp/t/<config>/.well-known/oauth-authorization-server metadata
//   GET    /mcp/t/<config>/authorize                              OAuth bounce
//   POST   /mcp/t/<config>/token                                  OAuth token
//   POST   /mcp/t/<config>/register                               RFC 7591 (stub)
//
//   GET    /auth/callback                                         shared callback
//   GET    /health                                                healthcheck

import http from "node:http";
import { randomUUID } from "node:crypto";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";

import { createHaloMcpServer } from "../server.js";
import { withRequestAuth } from "../halo/context.js";

import { parseTenantPath } from "./tenant.js";
import { getPublicOrigin } from "./origin.js";
import {
  emitProtectedResourceMetadata,
  emitAuthorizationServerMetadata,
  handleRegistration,
} from "./metadata.js";
import { handleAuthorize, handleToken } from "./oauth.js";
import { handleAuthCallback } from "./callback.js";

function readBody(req: http.IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    req.on("data", (c: Buffer) => chunks.push(c));
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function writeJson(
  res: http.ServerResponse,
  status: number,
  body: unknown,
  extraHeaders: Record<string, string> = {},
): void {
  res.writeHead(status, {
    "Content-Type": "application/json",
    "Cache-Control": "no-store",
    "Access-Control-Allow-Origin": "*",
    ...extraHeaders,
  });
  res.end(JSON.stringify(body));
}

function notFound(res: http.ServerResponse): void {
  res.writeHead(404, { "Content-Type": "text/plain" });
  res.end("Not found");
}

function methodNotAllowed(res: http.ServerResponse, allowed: string[]): void {
  res.writeHead(405, {
    Allow: allowed.join(", "),
    "Content-Type": "text/plain",
  });
  res.end("Method not allowed");
}

function preflight(res: http.ServerResponse): void {
  res.writeHead(204, {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Authorization, Content-Type, Mcp-Session-Id",
    "Access-Control-Max-Age": "86400",
  });
  res.end();
}

export function createHttpServer(): http.Server {
  return http.createServer(async (req, res) => {
    const started = Date.now();
    const method = req.method ?? "?";
    const url = req.url ?? "?";
    const origin = (req.headers.origin as string | undefined) ?? "-";
    const ua = (req.headers["user-agent"] as string | undefined) ?? "-";
    res.on("finish", () => {
      process.stderr.write(
        `mcp ${method} ${url} → ${res.statusCode} (${Date.now() - started}ms) origin=${origin} ua=${ua.slice(0, 80)}\n`,
      );
    });
    try {
      await route(req, res);
    } catch (err) {
      process.stderr.write(`halo-mcp-server: handler error: ${(err as Error).stack ?? err}\n`);
      if (!res.headersSent) {
        writeJson(res, 500, { error: "internal_error", message: (err as Error).message });
      }
    }
  });
}

async function route(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
  if (req.method === "OPTIONS") {
    preflight(res);
    return;
  }

  const url = new URL(req.url ?? "/", `http://${req.headers.host ?? "localhost"}`);
  const path = url.pathname;

  if (path === "/health" && req.method === "GET") {
    writeJson(res, 200, { status: "ok" });
    return;
  }

  if (path === "/auth/callback" && req.method === "GET") {
    await handleAuthCallback(req, res);
    return;
  }

  const tenantPath = parseTenantPath(path);
  if (!tenantPath) {
    notFound(res);
    return;
  }

  const { configBlob, tenant, rest } = tenantPath;
  const issuer = `${getPublicOrigin(req)}/mcp/t/${configBlob}`;

  // Metadata endpoints — readable without auth.
  if (rest === ".well-known/oauth-protected-resource" && req.method === "GET") {
    emitProtectedResourceMetadata(res, issuer);
    return;
  }
  if (rest === ".well-known/oauth-authorization-server" && req.method === "GET") {
    emitAuthorizationServerMetadata(res, issuer);
    return;
  }

  if (rest === "authorize") {
    if (req.method !== "GET") return methodNotAllowed(res, ["GET"]);
    handleAuthorize(req, res, configBlob, tenant);
    return;
  }

  if (rest === "token") {
    if (req.method !== "POST") return methodNotAllowed(res, ["POST"]);
    const body = await readBody(req);
    await handleToken(body, req, res, tenant);
    return;
  }

  if (rest === "register") {
    if (req.method !== "POST") return methodNotAllowed(res, ["POST"]);
    const body = await readBody(req);
    await handleRegistration(body, res);
    return;
  }

  // Anything else under /mcp/t/<config>/ — the JSON-RPC transport. Handles
  // POST (client → server) and GET (SSE stream back) and DELETE (session end).
  if (rest === "" || rest === "/") {
    await handleMcpTransport(req, res, configBlob, tenant);
    return;
  }

  notFound(res);
}

async function handleMcpTransport(
  req: http.IncomingMessage,
  res: http.ServerResponse,
  configBlob: string,
  tenant: { halo: string; clientId: string },
): Promise<void> {
  const authHeader = req.headers["authorization"];
  if (typeof authHeader !== "string" || !authHeader.startsWith("Bearer ")) {
    const issuer = `${getPublicOrigin(req)}/mcp/t/${configBlob}`;
    res.writeHead(401, {
      "Content-Type": "application/json",
      "WWW-Authenticate": `Bearer resource_metadata="${issuer}/.well-known/oauth-protected-resource"`,
      "Access-Control-Allow-Origin": "*",
    });
    res.end(JSON.stringify({ error: "unauthorized", message: "Bearer token required." }));
    return;
  }

  const accessToken = authHeader.slice("Bearer ".length).trim();
  if (!accessToken) {
    writeJson(res, 401, { error: "unauthorized", message: "Empty bearer token." });
    return;
  }

  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: () => randomUUID(),
    enableJsonResponse: false,
  });
  const server = createHaloMcpServer();

  res.on("close", () => {
    transport.close().catch(() => undefined);
    server.close().catch(() => undefined);
  });

  await server.connect(transport);
  await withRequestAuth(
    {
      baseUrl: tenant.halo,
      accessToken,
      clientId: tenant.clientId,
    },
    () => transport.handleRequest(req, res),
  );
}
