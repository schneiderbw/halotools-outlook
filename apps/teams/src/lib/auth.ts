// OAuth2 Authorization Code + PKCE flow for HaloPSA, run entirely client-side —
// Teams variant. Mirrors the Outlook add-in's auth.ts but uses
// microsoftTeams.authentication.authenticate() instead of Office's
// displayDialogAsync() to open the consent popup.
//
// Flow:
//   1. generate verifier + challenge + state
//   2. call microsoftTeams.authentication.authenticate({ url: <halo>/auth/authorize })
//      → Teams opens a popup that lands on Halo's consent page
//   3. user authenticates on Halo
//   4. Halo redirects to our /teams/auth/callback.html in the popup
//   5. callback page calls microsoftTeams.authentication.notifySuccess(JSON.stringify({code, state}))
//   6. the parent tab's authenticate() promise resolves with that JSON string
//   7. parent exchanges code+verifier for tokens at <halo>/auth/token
//   8. tokens persisted to localStorage (Teams tabs don't have roamingSettings)

import { authentication, app as teamsApp } from "@microsoft/teams-js";
import { generateState, generateVerifier, challengeFromVerifier } from "./pkce";
import {
  getConfig,
  getTokens,
  setTokens,
  clearTokens,
  type StoredTokens,
  type TenantConfig,
} from "./config";

const REDIRECT_PATH = "/teams/auth/callback.html";
const APP_ORIGIN = "https://tools.iusehalo.com";

function redirectUri(): string {
  return APP_ORIGIN + REDIRECT_PATH;
}

interface DialogResultMessage {
  code?: string;
  state?: string;
  error?: string;
  errorDescription?: string;
}

/** Trigger the OAuth dance. Resolves when tokens are obtained and stored. */
export async function signIn(): Promise<StoredTokens> {
  const cfg = getConfig();
  if (!cfg) throw new Error("HaloPSA tenant not configured. Run config first.");

  const verifier = generateVerifier();
  const challenge = await challengeFromVerifier(verifier);
  const state = generateState();

  const authorizeUrl = buildAuthorizeUrl(cfg, challenge, state);
  const message = await openAuthPopup(authorizeUrl);

  if (message.error) {
    throw new Error(`${message.error}: ${message.errorDescription ?? ""}`.trim());
  }
  if (!message.code || message.state !== state) {
    throw new Error("Authorization response missing code or state mismatch.");
  }

  const tokens = await exchangeCodeForTokens(cfg, message.code, verifier);
  await setTokens(tokens);
  return tokens;
}

/** Return a valid access token, refreshing if needed. Throws if no tokens. */
export async function getAccessToken(): Promise<string> {
  const tokens = getTokens();
  if (!tokens) throw new NotAuthenticatedError();

  // Refresh 60s before expiry to absorb clock drift.
  if (Date.now() >= tokens.expiresAt - 60_000) {
    const refreshed = await refresh(tokens.refreshToken);
    return refreshed.accessToken;
  }
  return tokens.accessToken;
}

/** Force a refresh using the stored refresh token. */
export async function refresh(refreshToken: string): Promise<StoredTokens> {
  const cfg = getConfig();
  if (!cfg) throw new Error("No tenant config.");

  const body = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    client_id: cfg.clientId,
  });

  const res = await fetch(`${cfg.haloBaseUrl}/auth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!res.ok) {
    await clearTokens();
    throw new NotAuthenticatedError(`Refresh failed: ${res.status}`);
  }

  const json = await res.json();
  const tokens = tokensFromResponse(json, refreshToken);
  await setTokens(tokens);
  return tokens;
}

export async function signOut(): Promise<void> {
  await clearTokens();
}

export function isAuthenticated(): boolean {
  return !!getTokens();
}

export class NotAuthenticatedError extends Error {
  constructor(msg = "Not authenticated") {
    super(msg);
    this.name = "NotAuthenticatedError";
  }
}

// ---------- internals ----------

function buildAuthorizeUrl(cfg: TenantConfig, codeChallenge: string, state: string): string {
  const params = new URLSearchParams({
    response_type: "code",
    client_id: cfg.clientId,
    redirect_uri: redirectUri(),
    scope: cfg.scope ?? "all",
    state,
    code_challenge: codeChallenge,
    code_challenge_method: "S256",
  });
  return `${cfg.haloBaseUrl}/auth/authorize?${params.toString()}`;
}

/**
 * Open the Halo consent page via Teams' authentication popup and wait for the
 * callback page to call notifySuccess/notifyFailure.
 *
 * If we're not running inside Teams (e.g. a developer hits the tab in a normal
 * browser tab during local dev), fall back to window.open + postMessage so
 * the flow is testable without sideloading.
 */
async function openAuthPopup(url: string): Promise<DialogResultMessage> {
  if (isInsideTeams()) {
    try {
      const result = await authentication.authenticate({
        url,
        width: 480,
        height: 600,
      });
      try {
        return JSON.parse(result) as DialogResultMessage;
      } catch {
        return { error: "bad_result", errorDescription: result };
      }
    } catch (e) {
      // authenticate() rejects with a string in many SDK versions; surface as error.
      const msg = e instanceof Error ? e.message : String(e);
      // The SDK throws "CancelledByUser" when the user closes the popup —
      // normalize so callers can show a friendlier message.
      if (msg === "CancelledByUser") return { error: "cancelled" };
      return { error: msg };
    }
  }
  // Browser dev fallback.
  return openAuthPopupViaWindowOpen(url);
}

function isInsideTeams(): boolean {
  // teamsJs sets a flag once initialized; otherwise sniff the URL/referrer.
  // app.isInitialized is the cheapest source-of-truth.
  try {
    return teamsApp.isInitialized();
  } catch {
    return false;
  }
}

/** Dev-only fallback for running the tab in a regular browser. */
function openAuthPopupViaWindowOpen(url: string): Promise<DialogResultMessage> {
  return new Promise((resolve, reject) => {
    const popup = window.open(url, "halo-auth", "width=480,height=640");
    if (!popup) {
      reject(new Error("Popup blocked. Allow popups for this site."));
      return;
    }

    function onMessage(ev: MessageEvent) {
      const data = ev.data as
        | { source?: string; ok?: boolean; payload?: DialogResultMessage }
        | undefined;
      if (!data || data.source !== "halo-teams-auth") return;
      window.removeEventListener("message", onMessage);
      try {
        popup?.close();
      } catch {
        /* ignore */
      }
      const payload = data.payload ?? {};
      resolve(payload);
    }

    window.addEventListener("message", onMessage);
  });
}

async function exchangeCodeForTokens(
  cfg: TenantConfig,
  code: string,
  verifier: string,
): Promise<StoredTokens> {
  const body = new URLSearchParams({
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri(),
    client_id: cfg.clientId,
    code_verifier: verifier,
  });

  const res = await fetch(`${cfg.haloBaseUrl}/auth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Token exchange failed (${res.status}): ${text}`);
  }
  return tokensFromResponse(await res.json());
}

function tokensFromResponse(
  json: {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
    scope?: string;
  },
  fallbackRefresh?: string,
): StoredTokens {
  const refresh_token = json.refresh_token ?? fallbackRefresh;
  if (!refresh_token) throw new Error("Token response missing refresh_token");
  return {
    accessToken: json.access_token,
    refreshToken: refresh_token,
    expiresAt: Date.now() + json.expires_in * 1000,
    scope: json.scope,
  };
}
