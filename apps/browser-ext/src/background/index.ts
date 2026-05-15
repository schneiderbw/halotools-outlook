// MV3 service worker. Owns:
//   - the OAuth flow (chrome.identity.launchWebAuthFlow)
//   - all Halo API calls (so the access token / refresh logic lives in
//     one place and so content scripts on third-party origins never see
//     CORS preflight failures — the SW fetches with the extension's own
//     host_permissions)
//   - an in-memory email-lookup cache for the content-script badges
//
// Stateless across SW restarts except for chrome.storage.local — Chrome
// can kill the SW at any time and Vite/crxjs will rebuild this file as
// an ES module entry on every change.

import { signIn, signOut, isAuthenticated } from "../lib/auth";
import { getConfig } from "../lib/storage";
import {
  findUserByEmail,
  listRecentTicketsForUser,
  searchUsers,
} from "../lib/halo-api";
import type { Request, LookupHit } from "../lib/messages";

// 10-minute in-memory cache of email → lookup result. The content
// script debounces and dedupes its own requests, but cross-tab traffic
// hits the SW often enough that caching here is cheap insurance.
//
// `null` means "we looked, didn't find anyone" — still worth caching to
// avoid re-querying for an unknown address on every page load.
const TTL_MS = 10 * 60 * 1000;
const lookupCache = new Map<string, { at: number; hit: LookupHit | null }>();

// Domains that almost never produce useful results in Halo (the user is
// rarely a customer of Gmail). Matches the content script's exclude list
// in spirit; we keep a duplicate here so server-side calls are blocked
// even if the content script somehow asks.
const SKIP_DOMAINS = new Set([
  "gmail.com",
  "googlemail.com",
  "outlook.com",
  "hotmail.com",
  "live.com",
  "yahoo.com",
  "icloud.com",
]);

function normaliseEmail(email: string): string | null {
  const e = email.trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) return null;
  const domain = e.split("@")[1];
  if (SKIP_DOMAINS.has(domain)) return null;
  return e;
}

async function handleLookup(email: string): Promise<LookupHit | null> {
  const norm = normaliseEmail(email);
  if (!norm) return null;

  const cached = lookupCache.get(norm);
  if (cached && Date.now() - cached.at < TTL_MS) return cached.hit;

  const user = await findUserByEmail(norm);
  if (!user) {
    lookupCache.set(norm, { at: Date.now(), hit: null });
    return null;
  }
  // recentTickets is best-effort — a failure shouldn't block the badge.
  let recentTickets: LookupHit["recentTickets"] = [];
  try {
    recentTickets = await listRecentTicketsForUser(user.id, 5);
  } catch {
    /* swallow — decorative */
  }
  const hit: LookupHit = { user, recentTickets };
  lookupCache.set(norm, { at: Date.now(), hit });
  return hit;
}

// chrome.runtime.onMessage's listener must return `true` to keep the
// channel open for async sendResponse. Wrap each handler so we can
// `await` cleanly and still satisfy the protocol.
chrome.runtime.onMessage.addListener((message: Request, _sender, sendResponse) => {
  (async () => {
    try {
      switch (message.kind) {
        case "ping":
          sendResponse({ ok: true, pong: true });
          return;
        case "signIn":
          await signIn();
          sendResponse({ ok: true });
          return;
        case "signOut":
          await signOut();
          // Clear cached lookups so the next sign-in doesn't show data
          // from the previous tenant.
          lookupCache.clear();
          sendResponse({ ok: true });
          return;
        case "getAuthStatus": {
          const cfg = await getConfig();
          const signedIn = await isAuthenticated();
          sendResponse({ ok: true, configured: !!cfg, signedIn });
          return;
        }
        case "lookupEmail": {
          const hit = await handleLookup(message.email);
          sendResponse({ ok: true, hit });
          return;
        }
        case "search": {
          const users = await searchUsers(message.query, 15);
          sendResponse({ ok: true, users });
          return;
        }
        case "openInHalo": {
          const cfg = await getConfig();
          if (!cfg) {
            sendResponse({ ok: true });
            return;
          }
          await chrome.tabs.create({ url: `${cfg.haloBaseUrl}${message.path}` });
          sendResponse({ ok: true });
          return;
        }
        default: {
          // Exhaustiveness check — TS will error here if a new kind is
          // added to Request without a handler.
          const _exhaustive: never = message;
          void _exhaustive;
          sendResponse({ ok: false, error: "Unknown message kind" });
        }
      }
    } catch (e) {
      sendResponse({ ok: false, error: (e as Error).message });
    }
  })();
  return true;
});

// On install, open the options page so the admin can configure Halo
// before they hit the popup and find a useless empty form.
chrome.runtime.onInstalled.addListener((details) => {
  if (details.reason === "install") {
    chrome.runtime.openOptionsPage();
  }
});
