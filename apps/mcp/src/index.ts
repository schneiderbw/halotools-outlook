#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { randomUUID } from "node:crypto";
import http from "node:http";
import { createHaloMcpServer } from "./server.js";
import {
  loadEnvAuth,
  parseBearerToken,
  withHaloAuth,
} from "./halo/context.js";
import type { HaloAuth } from "./halo/client.js";

// ---------- transport selection ----------

function pickTransport(argv: string[]): "stdio" | "http" {
  if (argv.includes("--http")) return "http";
  if (argv.includes("--stdio")) return "stdio";
  if (process.env.MCP_TRANSPORT === "http") return "http";
  return "stdio";
}

async function runStdio(): Promise<void> {
  const auth = loadEnvAuth();
  if (!auth) {
    process.stderr.write(
      "halo-mcp-server: missing Halo config. Set HALO_BASE_URL and either HALO_ACCESS_TOKEN or HALO_CLIENT_ID + HALO_CLIENT_SECRET.\n",
    );
    process.exit(1);
  }

  const server = createHaloMcpServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);

  // Wrap every incoming JSON-RPC message in the stdio auth scope. The high-level
  // McpServer routes via Server.setRequestHandler under the hood — by installing
  // an onmessage shim BEFORE connect() returns control to the transport, we
  // would race the SDK's own handler. So instead we lean on the fact that in
  // stdio mode there's exactly one HALO_* config: run the entire process under
  // the same auth context.
  await withHaloAuth(auth, () => new Promise<void>(() => { /* run forever */ }));
}

async function runHttp(): Promise<void> {
  const port = Number(process.env.PORT ?? 3001);

  // Stateless: every POST is independent. Each request carries its own
  // Authorization header naming a (tenant, token) pair, so we cannot pool
  // transports — or McpServer instances — across requests.
  const httpServer = http.createServer(async (req, res) => {
    // Health check.
    if (req.url === "/health" && req.method === "GET") {
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ status: "ok" }));
      return;
    }

    if (!req.url?.startsWith("/mcp")) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }

    // Parse Authorization: Bearer <token>.
    const authHeader = req.headers["authorization"];
    let haloAuth: HaloAuth | undefined;
    if (typeof authHeader === "string" && authHeader.startsWith("Bearer ")) {
      try {
        haloAuth = parseBearerToken(authHeader.slice("Bearer ".length));
      } catch (e) {
        res.writeHead(401, { "Content-Type": "application/json" });
        res.end(
          JSON.stringify({ error: "invalid_bearer", message: (e as Error).message }),
        );
        return;
      }
    } else {
      // Fall back to env vars (useful for single-tenant hosted deployments).
      haloAuth = loadEnvAuth();
    }

    if (!haloAuth) {
      res.writeHead(401, { "Content-Type": "application/json" });
      res.end(
        JSON.stringify({
          error: "unauthorized",
          message:
            "Send Authorization: Bearer halo:<base-url>:<halo-access-token> or haloc:<base-url>:<client-id>:<client-secret>.",
        }),
      );
      return;
    }

    // Per-request transport and server (fully stateless). Generating a session
    // ID per call lets the SDK still issue an Mcp-Session-Id response header,
    // which some clients expect, but we don't retain any state across requests.
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
    await withHaloAuth(haloAuth, () => transport.handleRequest(req, res));
  });

  httpServer.listen(port, () => {
    process.stderr.write(
      `halo-mcp-server: listening on http://0.0.0.0:${port}/mcp\n`,
    );
  });

  // Graceful shutdown.
  const shutdown = () => {
    httpServer.close(() => process.exit(0));
    setTimeout(() => process.exit(1), 5000).unref();
  };
  process.on("SIGTERM", shutdown);
  process.on("SIGINT", shutdown);
}

async function main(): Promise<void> {
  const transport = pickTransport(process.argv);
  if (transport === "http") {
    await runHttp();
  } else {
    await runStdio();
  }
}

main().catch((err) => {
  process.stderr.write(`halo-mcp-server: fatal: ${err?.stack ?? err}\n`);
  process.exit(1);
});
