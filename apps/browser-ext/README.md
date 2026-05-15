# HaloPSA Browser Extension

Chromium-first (Chrome + Edge) MV3 extension that surfaces HaloPSA contact +
ticket context wherever an agent is browsing. Sibling to the Outlook add-in at
the repo root — same OAuth2 + PKCE model, same `Halo Connect` per-tenant setup,
no backend.

## What it does (v1)

- **Toolbar popup** — search Halo for a contact by name or email, jump to the
  customer page or start a new ticket.
- **Inline email badges** — every visible email address on every page (except
  Gmail / Outlook.com — too noisy) gets a tiny red Halo badge if the address
  exists as a Halo user. Click it to open `/customer?userid=N` in your tenant.
- **Options page** — wizard for Halo URL + Client ID + connect. OAuth runs via
  `chrome.identity.launchWebAuthFlow`.

Tokens live in `chrome.storage.local`. Nothing leaves the browser.

## Build

```bash
cd browser-ext
npm install
npm run build
```

Output goes to `browser-ext/dist/`. `@crxjs/vite-plugin` is the MV3-aware
bundler glue; it reads `manifest.json` as the source of truth and rewrites
script paths.

## Load unpacked (development)

1. `npm run build`
2. Open `chrome://extensions/` (or `edge://extensions/`).
3. Toggle **Developer mode** (top right).
4. Click **Load unpacked** → select `browser-ext/dist/`.
5. The Options page opens automatically on first install.

For dev with HMR: `npm run dev` then reload the unpacked extension after the
first build.

## Halo Connect configuration (admin side)

For each MSP tenant that uses this extension:

1. In HaloPSA, go to **Configuration → Integrations → Halo Connect → Custom
   Applications**.
2. Create a new application (or reuse the one used for the Outlook add-in —
   they can share an OAuth client).
3. **Authentication Method:** Authorization Code (PKCE).
4. **Login Type:** Agent.
5. **Redirect URI:** the value shown on the extension's Options page after
   load. It takes the form:

   ```
   https://<EXTENSION_ID>.chromiumapp.org/
   ```

   - During local-unpacked development, `EXTENSION_ID` is a random ID Chrome
     generates the first time you load the unpacked extension. It persists for
     the life of that install — uninstalling and reloading mints a new one.
     Find the ID at the top of the row on `chrome://extensions/`.
   - For Chrome Web Store production builds, the Store assigns a stable
     extension ID. Use that in production.
   - To pin a stable ID across machines for dev/CI, add a `"key"` field to
     `manifest.json` containing the public key portion of a generated
     keypair — Chrome derives the ID deterministically from it. Not required
     for v1; document only if needed.

6. **CORS Whitelist:** add `chrome-extension://<EXTENSION_ID>` (no trailing
   slash, no path).
7. **Scope:** at minimum `all` (the extension passes `scope=all` by default).
8. Copy the **Client ID** into the extension's Options page along with the
   Halo base URL (e.g. `https://halo.example.com`).
9. Click **Connect to Halo** → browser auth window opens → sign in.

## Manifest V3 permissions

| permission          | why                                                                            |
| ------------------- | ------------------------------------------------------------------------------ |
| `storage`           | persist Halo tenant config + OAuth tokens in `chrome.storage.local`            |
| `identity`          | `chrome.identity.launchWebAuthFlow` for the OAuth dance                        |
| `tabs`              | open Halo deep-links (customer page, new ticket) in a new tab                  |
| `host_permissions: <all_urls>` | content script needs to inspect text on every page for emails       |

The `<all_urls>` host permission is the broadest grant in the manifest. It's
the only way the content script can find email addresses on arbitrary pages,
but it should be **narrowed before publishing**. Two options:

- Switch to `activeTab` + a user gesture model — heavy UX cost, drops the
  "passive badge on every page" feature.
- Move to an **optional** `host_permissions` set with a per-site allowlist
  managed in the options page. This is the recommended v2 path.

`gmail.com`, `outlook.live.com`, `outlook.office.com`, and `outlook.office365.com`
are excluded via `exclude_matches` to avoid duplicate-looking badges in the
user's own inbox.

## Architecture

```
browser-ext/
├── manifest.json            # MV3 manifest (source of truth)
├── vite.config.ts           # vite + @crxjs/vite-plugin
├── tsconfig.json
├── public/
│   └── icons/               # 16/32/48/128 placeholder red disc PNGs — swap in the real Halo bubble
└── src/
    ├── popup/               # Fluent UI v9 + React toolbar popup
    ├── options/             # Fluent UI v9 + React options page
    ├── background/          # MV3 service worker — owns OAuth + Halo calls + lookup cache
    ├── content/             # plain TS content script, no React
    └── lib/                 # copied from ../../src/, adapted for chrome.storage
        ├── auth.ts          # chrome.identity-based OAuth2 + PKCE
        ├── halo-api.ts      # bearer auth + 401 retry, identical to Outlook side
        ├── messages.ts      # typed SW <-> UI request/response contract
        ├── pkce.ts          # verbatim copy from Outlook lib
        ├── storage.ts       # chrome.storage.local wrapper
        └── types.ts         # subset of ../../src/types/halo.ts
```

The `lib/` files were **copied**, not symlinked or imported, so the extension
remains a standalone package and the upstream Outlook side stays untouched.
When upstream changes a Halo type or a quirk in `halo-api.ts`, mirror it here.

## Future Firefox support

`@crxjs/vite-plugin` has Firefox support on its roadmap; the manifest itself
needs a `browser_specific_settings.gecko` block and Firefox uses the `browser`
namespace polyfill rather than `chrome` (the `chrome` global is aliased in
Firefox 109+, so most calls work as-is). When we add Firefox:

1. Add `webextension-polyfill` if needed for older Firefox.
2. Ship a parallel `manifest.firefox.json` with the gecko ID and any
   permission-name differences.
3. The OAuth redirect URL pattern differs (`<UUID>.extensions.allizom.org`) —
   document the alternate URL on the options page when running in Firefox.

## Things deferred (don't add without asking)

- Ticket-creation from popup beyond the deep-link to Halo's new-ticket page.
- Backend / proxy of any kind — same constraint as the Outlook side.
- Page-aware classifiers (e.g. "you're on a known vendor's status page,
  surface related tickets") — v2.
- Per-site allowlist for the content script — see `host_permissions` note.
