// OAuth2 Authorization Code + PKCE flow for HaloPSA, run inside the
// extension's service worker via chrome.identity.launchWebAuthFlow.
//
// ADAPTED from ../../../src/lib/auth.ts. The Outlook side uses
// Office.context.ui.displayDialogAsync; chrome.identity.launchWebAuthFlow
// is the extension-world equivalent — it opens the authorize URL in a
// modal, watches the redirect, and resolves with the final redirect URL
// when it matches `https://<EXT_ID>.chromiumapp.org/` (the special URL
// Chrome reserves for this purpose).
//
// Flow:
//   1. generate verifier + challenge + state
//   2. compute redirect_uri = chrome.identity.getRedirectURL()
//   3. launchWebAuthFlow(authorize URL) → user signs in on Halo
//   4. Halo redirects to redirect_uri with ?code=...&state=...
//   5. exchange code + verifier for tokens at {halo}/auth/token
//   6. tokens persisted to chrome.storage.local
//
// Note: this module MUST run in the service worker or options page —
// chrome.identity is not available from the popup window after Chrome 116
// changes (the popup closes when the auth window opens).

import { generateState, generateVerifier, challengeFromVerifier } from "./pkce";
import {
  getConfig,
  getTokens,
  setTokens,
  clearTokens,
  type StoredTokens,
  type TenantConfig,
} from "./storage";

/** Chrome assigns each extension `https://<id>.chromiumapp.org/` as a redirect target. */
function redirectUri(): string {
  return chrome.identity.getRedirectURL();
}

/** Trigger the OAuth dance. Resolves when tokens are obtained and stored. */
export async function signIn(): Promise<StoredTokens> {
  const cfg = await getConfig();
  if (!cfg) throw new Error("HaloPSA tenant not configured. Run config first.");

  const verifier = generateVerifier();
  const challenge = await challengeFromVerifier(verifier);
  const state = generateState();

  const authorizeUrl = buildAuthorizeUrl(cfg, challenge, state);
  const finalUrl = await launchAuthFlow(authorizeUrl);

  const params = new URLSearchParams(new URL(finalUrl).search);
  const error = params.get("error");
  if (error) {
    throw new Error(`${error}: ${params.get("error_description") ?? ""}`.trim());
  }
  const code = params.get("code");
  const returnedState = params.get("state");
  if (!code || returnedState !== state) {
    throw new Error("Authorization response missing code or state mismatch.");
  }

  const tokens = await exchangeCodeForTokens(cfg, code, verifier);
  await setTokens(tokens);
  return tokens;
}

/** Return a valid access token, refreshing if needed. Throws if no tokens. */
export async function getAccessToken(): Promise<string> {
  const tokens = await getTokens();
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
  const cfg = await getConfig();
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
    // Refresh token likely revoked or expired — wipe and force re-auth.
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

export async function isAuthenticated(): Promise<boolean> {
  return !!(await getTokens());
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

function launchAuthFlow(url: string): Promise<string> {
  return new Promise((resolve, reject) => {
    chrome.identity.launchWebAuthFlow({ url, interactive: true }, (responseUrl) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message ?? "Auth flow failed"));
        return;
      }
      if (!responseUrl) {
        reject(new Error("Sign-in cancelled."));
        return;
      }
      resolve(responseUrl);
    });
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
