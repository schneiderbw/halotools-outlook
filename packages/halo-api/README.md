# @iusehalo/halo-api

Framework-agnostic HaloPSA client used by the Halo Outlook add-in and any other surface (Teams app, MCP server, browser extension) that needs to talk to HaloPSA from a browser-like environment. Provides a fetch wrapper with 401-retry-on-refresh, an OAuth2 Authorization Code + PKCE flow, RFC 5322 Message-ID ticket threading via `findTicketsForEmail`, and storage-adapter-based persistence of tenant config + tokens. No React, no UI, no Office.js — the consumer plugs in a `Storage` adapter and a `DialogOpener`.

```ts
import {
  setStorage,
  localStorageStorage,
  setConfig,
  signIn,
  searchUsers,
  type DialogOpener,
} from "@iusehalo/halo-api";

setStorage(localStorageStorage());

await setConfig({
  haloBaseUrl: "https://halo.example.com",
  clientId: "abc-123",
});

const opener: DialogOpener = {
  async open(url) {
    // Open `url` in a popup / Office dialog / browser tab, wait for the
    // OAuth redirect, parse code/state/error from the query, and return them.
    return { code: "...", state: "..." };
  },
};

await signIn(opener, {
  redirectUri: "https://your.app/auth/callback.html",
});

const users = await searchUsers("alice");
```
