# halo-mcp-server

Model Context Protocol server exposing HaloPSA as a tool for AI assistants (Claude Desktop, ChatGPT desktop, Cursor, etc.).

This is a sibling deployment to the Outlook add-in in the parent repo. The add-in is browser-side; this server is Node-side and runs either as a local subprocess (stdio) or as a hosted service (HTTP+SSE via the Streamable HTTP transport).

## What it exposes

| Tool                    | Purpose                                                                     |
| ----------------------- | --------------------------------------------------------------------------- |
| `findContact`           | Look up a Halo contact by email; returns name, client, phone, open count.   |
| `listOpenTickets`       | List open tickets for a client and/or user.                                 |
| `searchTickets`         | Free-text search across Halo tickets.                                       |
| `createTicket`          | Open a new ticket against a client (and optional user).                     |
| `appendActionToTicket`  | Append a note / action to an existing ticket. Default outcome `"Note"`.     |
| `logNote`               | Create a CRM note against a client / user / site.                           |
| `searchCannedText`      | Search saved canned-text snippets, optionally scoped to a group.            |
| `getActivityFeed`       | Merged activity feed (actions, notes, status changes) for a client/user.    |

## Install / build

From inside `mcp/`:

```bash
npm install
npm run build
```

Produces `dist/index.js`. The binary alias `halo-mcp` resolves to `dist/index.js`.

## Run modes

```bash
node dist/index.js --stdio    # default; for local Claude Desktop / Cursor
node dist/index.js --http     # listens on PORT (default 3001) at /mcp
```

`MCP_TRANSPORT=http` env var has the same effect as `--http`.

## Configuration

### Stdio mode (local desktop integration)

Set these env vars before launching:

| Var                   | Required          | Notes                                                      |
| --------------------- | ----------------- | ---------------------------------------------------------- |
| `HALO_BASE_URL`       | yes               | e.g. `https://acme.halopsa.com`                            |
| `HALO_ACCESS_TOKEN`   | one of            | Pre-issued bearer token.                                   |
| `HALO_CLIENT_ID`      | one of (with secret) | Halo Connect application client ID.                     |
| `HALO_CLIENT_SECRET`  | with `HALO_CLIENT_ID` | Halo Connect application client secret. Server-side only. |
| `HALO_SCOPE`          | no                | Defaults to `all`.                                         |

With client credentials, the server exchanges them for a token at `<base>/auth/token` and caches it until expiry.

### HTTP mode (hosted multi-tenant)

Each request must include `Authorization: Bearer <token>`. Token format:

- `halo:<base-url>:<halo-access-token>` — pre-issued bearer token (recommended).
- `haloc:<base-url>:<client-id>:<client-secret>` — server exchanges for a token.

Examples:

```
Authorization: Bearer halo:https://acme.halopsa.com:eyJhbGciOiJIUzI1NiJ9.AbC...
Authorization: Bearer haloc:https://acme.halopsa.com:abcd1234:s3cr3t
```

**Security notes for HTTP mode**

- This server has no rate limit and no token validation beyond passing the value straight to Halo.
- Put a reverse proxy (nginx / Cloudflare) in front for TLS, rate limiting, and IP allow-listing.
- The `haloc:` form puts a client secret on the wire (encrypted in transit, but parsed server-side) — prefer `halo:` with rotating short-lived tokens where possible.
- Falls back to `HALO_*` env vars if no Authorization header is sent. Disable that fallback in shared deployments by leaving the env vars unset.

If you only need it for one Halo tenant, just set the env vars and skip per-request auth.

## Adding to Claude Desktop

In `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or the Windows equivalent:

```json
{
  "mcpServers": {
    "halo": {
      "command": "node",
      "args": ["/absolute/path/to/halotools-outlook/mcp/dist/index.js", "--stdio"],
      "env": {
        "HALO_BASE_URL": "https://acme.halopsa.com",
        "HALO_CLIENT_ID": "your-halo-connect-client-id",
        "HALO_CLIENT_SECRET": "your-halo-connect-client-secret"
      }
    }
  }
}
```

If you publish the package to npm as `halo-mcp-server`, you can use `npx` instead:

```json
{
  "mcpServers": {
    "halo": {
      "command": "npx",
      "args": ["-y", "halo-mcp-server", "--stdio"],
      "env": {
        "HALO_BASE_URL": "https://acme.halopsa.com",
        "HALO_ACCESS_TOKEN": "<token>"
      }
    }
  }
}
```

Restart Claude Desktop; tools appear in the MCP panel.

## Adding to Cursor

Same shape under `.cursor/mcp.json` in the workspace, or in user settings under `mcp.servers`.

## Hosted deployment

The repo's pattern is EasyPanel. Build:

```bash
docker build -t halo-mcp-server -f Dockerfile .
docker run --rm -p 3001:3001 halo-mcp-server
```

For `tools.iusehalo.com/mcp`, drop a second EasyPanel service that builds this `Dockerfile`. Route `/mcp` to it via nginx; keep the existing add-in at `/outlook/`. Health check is `GET /health`.

## Project layout

```
mcp/
├── package.json
├── tsconfig.json
├── Dockerfile
├── src/
│   ├── index.ts          # entrypoint; stdio / http transport switch
│   ├── server.ts         # McpServer construction + tool registration
│   ├── halo/
│   │   ├── client.ts     # stateless HaloPSA REST client (server-side fetch)
│   │   ├── types.ts      # Halo response / payload shapes
│   │   └── context.ts    # AsyncLocalStorage-based per-request auth context
│   └── tools/            # one file per MCP tool
│       ├── findContact.ts
│       ├── listOpenTickets.ts
│       ├── searchTickets.ts
│       ├── createTicket.ts
│       ├── appendActionToTicket.ts
│       ├── logNote.ts
│       ├── searchCannedText.ts
│       ├── getActivityFeed.ts
│       └── index.ts
└── dist/                 # built output (after `npm run build`)
```

## Limitations / deferred

- No tests in this skeleton.
- No tool for attaching files to tickets — actions currently log notes only.
- No tool for updating an existing ticket's status / assignee — append-only.
- No prompt or resource exports; tools only.
- The HTTP transport is stateless; long-running streaming jobs are not supported.
