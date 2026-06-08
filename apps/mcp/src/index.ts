#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { randomUUID } from "node:crypto";
import http from "node:http";
import { createHaloMcpServer } from "./server.js";
import {
  installRequestStorage,
  loadEnvAuth,
  parseBearerToken,
  withRequestAuth,
  type RequestAuth,
} from "./halo/context.js";

// halo-api reads tenant + token via its storage adapter; install ours (which is
// AsyncLocalStorage-backed) before any tool can run.
installRequestStorage();

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
      "halo-mcp-server: missing Halo config. Set HALO_BASE_URL and HALO_ACCESS_TOKEN.\n",
    );
    process.exit(1);
  }

  const server = createHaloMcpServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);

  // Stdio mode is single-tenant for the life of the process — wrap the whole
  // run in one auth scope so every tool call resolves the same Halo config.
  await withRequestAuth(auth, () => new Promise<void>(() => { /* run forever */ }));
}

async function runHttp(): Promise<void> {
  const port = Number(process.env.PORT ?? 3001);

  // Stateless: every POST is independent. Each request carries its own bearer
  // token naming a (tenant, access-token) pair, so we mint a fresh transport
  // and McpServer per request and run it inside its own auth scope.
  const httpServer = http.createServer(async (req, res) => {
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

    const authHeader = req.headers["authorization"];
    let haloAuth: RequestAuth | undefined;
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
      haloAuth = loadEnvAuth();
    }

    if (!haloAuth) {
      res.writeHead(401, { "Content-Type": "application/json" });
      res.end(
        JSON.stringify({
          error: "unauthorized",
          message:
            "Send Authorization: Bearer halo:<base-url>:<halo-access-token>.",
        }),
      );
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
    await withRequestAuth(haloAuth, () => transport.handleRequest(req, res));
  });

  httpServer.listen(port, () => {
    process.stderr.write(
      `halo-mcp-server: listening on http://0.0.0.0:${port}/mcp\n`,
    );
  });

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
