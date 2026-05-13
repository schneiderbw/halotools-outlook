# Project notes for Claude Code

Quick context for picking this up in a new session.

## What this is

An Outlook task-pane add-in for HaloPSA. Pure SPA (no backend). OAuth2 + PKCE public client. Multi-tenant — each MSP brings their own HaloPSA URL and Halo Connect Client ID.

Deployed at `https://tools.iusehalo.com/outlook/` on EasyPanel (static nginx).

## Architecture invariants

- **No backend.** All Halo calls go from the browser directly. CORS is configured per-tenant on the Halo Connect application. Don't introduce a backend without explicit ask.
- **Multi-tenant.** Halo URL + Client ID live in `Office.context.roamingSettings`, scoped to the user's mailbox. Nothing about a specific MSP is in the build artifact.
- **Tokens never leave the client.** Refresh tokens stay in `roamingSettings`. Auth is `signIn()` in `src/lib/auth.ts` via the Office Dialog API.
- **Read surface only.** v1 is `mailRead` context. Compose support requires a new runtime in `manifest.json` plus a different React shell (selection of recipients vs sender).

## Repo layout

```
src/
├── lib/         # Framework-agnostic logic (auth, halo-api, storage, office helpers, pkce)
├── components/  # Fluent UI React components
├── types/       # Halo response shapes
└── App.tsx      # Top-level state machine: Config → Auth → Dashboard
```

## Conventions

- Fluent UI v9 only — don't pull in Fluent v8, MUI, Chakra, etc. The task pane should feel native to Outlook.
- All Halo API calls go through `src/lib/halo-api.ts::call`. Adds bearer auth, 401-retry-on-refresh, and network-failure messaging.
- `Office.context.roamingSettings` is wrapped by `src/lib/storage.ts` — use the wrapper, not the raw API. It also falls back to `localStorage` outside Outlook for dev.
- Components avoid persistent state. Storage lives in `lib/config.ts` (tenant + tokens) and `lib/defaults.ts` (user preferences).
- HaloPSA SQL constraints (Report Center, Runbooks): no single-line comments, no trailing semicolons, one statement, no variables. *Not directly relevant to this repo but applies to any SQL emitted alongside it.*

## Common dev commands

```bash
npm install
npm run dev          # vite dev server on port 3000
npm run typecheck    # tsc --noEmit
npm run build        # production bundle to ./dist
npm run validate-manifest    # office-addin-manifest validate manifest.json
```

Manifest validation needs network access to `developer.microsoft.com` for the schema fetch.

## Gotchas

- **Halo API payload variations** — Halo's REST endpoints vary slightly between versions. Three places to watch:
  1. Custom-field query params: `?field_X=` vs `?customfield_X=` depending on tenant version. Currently using `field_X`.
  2. Inline attachment field name: `data_base64` vs `data` vs `base64`. Currently using `data_base64`.
  3. Action `outcome` strings are tenant-configurable (`"Email Received"` is the assumed default).
  If something breaks at runtime, try the alternates before assuming a deeper bug.
- **Office.js loads from MS CDN**, not from npm. The `<script>` tag in `index.html` and `auth/callback.html` pulls it. `@types/office-js` provides the TS types.
- **Office Dialog API for OAuth** — `displayDialogAsync` requires HTTPS even for localhost dev. Use `office-addin-dev-certs` or a tunnel.
- **Bundle splitting** — Dashboard is lazy-loaded so the initial config/auth flow stays under 100 KB gzipped. Don't import Dashboard statically from `App.tsx`.
- **roamingSettings has a ~32 KB total cap.** Tokens + tenant config + defaults fit comfortably, but don't dump large objects in there.
- **No localStorage in artifacts/built code other than the dev fallback.** roamingSettings is the storage of record.

## Things explicitly deferred (don't add without asking)

- Backend proxy / API gateway — pure SPA is the design.
- Compose surface support (mailCompose runtime) — separate runtime, separate UI flow.
- Server-side detection / classification layer — v2 work.
- Halo webhook handling — needs a backend.
- Time-tracking beyond per-action `time_taken` minutes — full timer / charge code logic is out of scope.

## Likely first tasks after handoff

1. Replace the placeholder developer info in `manifest.json` (privacyUrl, termsOfUseUrl, etc.) once you have the real URLs.
2. Test against a real Halo tenant — settle the three payload variations listed under Gotchas.
3. Set up the EasyPanel deployment for `tools.iusehalo.com`.
4. Register the Halo Connect app in the Rising Tide tenant and run the first end-to-end auth.
5. v1.5: write a Rewst workflow / Halo intake rule that reads `CFOutlookConversationId` and `CFOutlookInternetMessageId` from inbound replies and threads them to existing tickets, replacing Halo's subject-only matching for the sales mailbox.
