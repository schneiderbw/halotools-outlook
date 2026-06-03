# HaloPSA for Outlook

Modern Outlook task-pane add-in for HaloPSA. Surfaces sender context and ticket history, and lets agents log emails to tickets (append or create) without leaving Outlook.

## Installation

Go to https://tools.iusehalo.com and follow the instructions 🙂

The below details are for security review or self-hosting only.

## Architecture

- **Pure SPA** — OAuth2 + PKCE public client, no backend. Direct browser → HaloPSA API.
- **Multi-tenant** — each MSP registers their own Halo Connect application; the user enters their HaloPSA URL + Client ID on first run.
- **Hosting** — static, nginx, served from `https://tools.iusehalo.com/outlook/` on EasyPanel.
- **Token storage** — `Office.context.roamingSettings` (per-mailbox, roams across devices).
- **Stack** — Vite + React 18 + TypeScript + Fluent UI v9.

## v1 features

- Per-tenant config screen on first launch (HaloPSA URL + Client ID)
- PKCE OAuth2 sign-in via Office Dialog API; refresh-token rotation client-side
- Sender lookup: contact (exact email) → falls back to client (domain match)
- Manual overrides: search dialogs for contact and client when auto-match is wrong
- Open tickets for the matched client, with thread match by `CFOutlookConversationId`
- **Ticket quick actions** (per row): open in Halo, assign to me, change status
- **Append email to ticket** — HTML body, attachments, internal-note toggle, time entry (minutes)
- **Create ticket from email** — HTML body, ticket type picker, attachments, dedup warning when thread match exists
- **Smart button emphasis** — Append is primary when a thread match is present, Create otherwise
- Settings screen — default ticket type, default action outcome, attachments-by-default
- **CORS / network failure detection** — surfaces "likely a CORS misconfig" hint instead of cryptic fetch errors
- **Refresh button** on the dashboard for manual re-resolution
- Code-split bundle (Dashboard loads lazy after auth)
- Multi-tenant safe — no shared secrets, no shared backend

## Repo layout

```
.
├── manifest.json              # Unified JSON manifest
├── index.html                 # Task pane entry
├── auth/callback.html         # OAuth dialog landing
├── public/assets/             # Icons (svg source + 16/32/64/80/128 png)
├── src/
│   ├── main.tsx               # React bootstrap (awaits Office.onReady)
│   ├── App.tsx                # Config → Auth → Dashboard state machine
│   ├── lib/
│   │   ├── pkce.ts            # RFC 7636 verifier/challenge
│   │   ├── auth.ts            # Authorize-via-dialog + token exchange + refresh
│   │   ├── halo-api.ts        # Halo REST wrapper with 401-retry-on-refresh
│   │   ├── storage.ts         # roamingSettings (with localStorage fallback for dev)
│   │   ├── config.ts          # Per-tenant config + token storage
│   │   ├── defaults.ts        # Per-user preferences
│   │   └── office.ts          # Email context + body + attachment helpers
│   ├── components/
│   │   ├── ConfigScreen.tsx   # First-run: Halo URL + Client ID
│   │   ├── AuthScreen.tsx     # Sign-in trigger
│   │   ├── Dashboard.tsx      # Main pane — resolves sender, lists tickets
│   │   ├── ContactCard.tsx    # Sender + match status + override pickers
│   │   ├── TicketList.tsx     # Open & thread-matched tickets
│   │   ├── LogActions.tsx     # Append + Create dialogs
│   │   ├── SettingsScreen.tsx # Defaults, sign out, switch tenant
│   │   └── SearchPicker.tsx   # Generic Halo entity picker
│   └── types/halo.ts
├── Dockerfile + nginx.conf    # EasyPanel deployment
└── package.json + vite.config.ts + tsconfig.json
```

## Setup on the HaloPSA side (per tenant)

Each customer's HaloPSA admin does this once on their own instance.

1. **Custom fields** — add two ticket custom fields:
   - `CFOutlookConversationId` (Text, indexed if possible)
   - `CFOutlookInternetMessageId` (Text)
2. **Halo Connect application** — Config → Integrations → Halo Connect → API → New:
   - Application name: `Outlook Add-in`
   - Authentication Method: **Authorization Code** (PKCE)
   - Public client: **yes** (no secret required)
   - Login Type: **Agent**
   - Redirect URI: `https://tools.iusehalo.com/outlook/auth/callback.html`
   - Permissions: `all` (or scope down to read tickets/clients/users + edit tickets/actions/attachments)
   - CORS allowed origins: `https://tools.iusehalo.com`
3. Copy the **Client ID** out for users.

## Setup on the Outlook side (per user)

1. Open the add-in in Outlook — first launch shows the config screen.
2. Enter the HaloPSA URL and the Client ID from above.
3. Click **Sign in with HaloPSA** — a dialog opens to Halo's authorize page, user authenticates with their own agent creds, dialog returns the code to the task pane.

## Local development

```bash
npm install
npm run dev          # vite dev server on port 3000
npm run typecheck    # tsc --noEmit
npm run build        # production bundle to ./dist
```

Outlook sideload requires HTTPS:

- **`office-addin-dev-certs`** for a trusted localhost cert, then enable `server.https` in `vite.config.ts`.
- **ngrok / Cloudflare tunnel** pointing at `http://localhost:3000`, then update the manifest URLs to the tunnel hostname for the dev manifest only.

For the dev manifest, also add a localhost-equivalent redirect URI to the Halo Connect app (e.g. `https://localhost:3000/auth/callback.html`).

## Production deployment (EasyPanel)

Static-only nginx image. Deploy as an App on EasyPanel:

1. New App service pointed at this repo.
2. Build type: Dockerfile.
3. Domain: `tools.iusehalo.com` with TLS via Let's Encrypt.
4. No env vars required — multi-tenant config lives client-side per user.

Final URLs:

- Add-in: `https://tools.iusehalo.com/outlook/`
- OAuth callback: `https://tools.iusehalo.com/outlook/auth/callback.html`
- Icons: `https://tools.iusehalo.com/outlook/assets/icon-{16,32,64,80,128}.png`

## Installing the add-in

For testing: Outlook → Get Add-ins → My add-ins → Add a custom add-in → From URL → `https://tools.iusehalo.com/outlook/manifest.json`.

For wider rollout: push via Microsoft 365 admin center → Integrated apps, or via Intune for managed devices.

For public listing: submit to AppSource (partner.microsoft.com).

## Open items still worth verifying before sideload

- **Halo custom-field query syntax** — `halo-api.ts` uses `?field_CFOutlookConversationId=...`; varies by Halo version. Falls back gracefully (`.catch(() => [])`).
- **Halo OAuth endpoint paths** — assumed `/auth/authorize` and `/auth/token`; standard but worth confirming on the target tenant.
- **Halo attachment payload shape** — `attachments: [{ filename, data_base64, contenttype, isimage }]` is the assumed inline shape on `/api/Actions` and `/api/Tickets`. If the tenant version expects `data` or `base64` instead of `data_base64`, adjust in `LogActions.tsx::toHaloAttachment`.
- **Action outcomes** — `"Email Received"` is the default for Append; users can override per-tenant in Settings → Defaults.
- **Manifest schema validation** — run `npx office-addin-manifest validate manifest.json` against a network with access to `developer.microsoft.com` before submission.

## Roadmap (post-v1)

- **v1.5 — Halo intake fix**: rules / Rewst workflow on the Halo side that reads `CFOutlookConversationId` + `CFOutlookInternetMessageId` from inbound replies to thread to existing tickets, replacing Halo's subject-only match.
- **v2 — detection layer**: thin EasyPanel backend that does server-side classification (sales vs support vs spam), fuzzy contact matching, and shared sender → client caching firm-wide.
- **v3 — compose surface**: add a compose-time runtime so the add-in also assists when writing outbound mail (template lookup, send-as-ticket-action, etc.).
- **More tools at `tools.iusehalo.com/`**: public MCP server, helper agents, landing page for the broader Halo community.
