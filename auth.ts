// OAuth2 Authorization Code + PKCE flow for HaloPSA, run entirely client-side.
//
// Why client-side: Halo Connect supports public clients (no secret) when using PKCE.
// Combined with per-app CORS allowlist on the Halo side, the SPA talks directly to Halo.
//
// Flow:
//   1. generate verifier + challenge + state
//   2. open Office Dialog at {halo}/auth/authorize with PKCE + state
//   3. user authenticates on Halo
//   4. Halo redirects to our callback page in the dialog
//   5. callback page posts the code back to the parent via Office.context.ui.messageParent
//   6. parent exchanges code+verifier for tokens at {halo}/auth/token
//   7. tokens persisted to roamingSettings

import { generateState, generateVerifier, challengeFromVerifier } from "./pkce";
import {
  getConfig,
  getTokens,
  setTokens,
  clearTokens,
  type StoredTokens,
  type TenantConfig,
} from "./config";

const REDIRECT_PATH = "/outlook/auth/callback.html";
const ADDIN_ORIGIN = "https://tools.iusehalo.com";

function redirectUri(): string {
  return ADDIN_ORIGIN + REDIRECT_PATH;
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
  const message = await openAuthDialog(authorizeUrl);

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

function openAuthDialog(url: string): Promise<DialogResultMessage> {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      url,
      { height: 60, width: 30, promptBeforeOpen: false },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(asyncResult.error?.message ?? "Failed to open auth dialog"));
          return;
        }
        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          try {
            // messageProperty is the JSON string posted from the callback page
            const data: DialogResultMessage = JSON.parse(
              (arg as { message: string }).message,
            );
            resolve(data);
          } catch (e) {
            reject(new Error(`Bad message from auth dialog: ${(e as Error).message}`));
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          // 12006 = user closed the dialog
          const ev = arg as { error: number };
          dialog.close();
          if (ev.error === 12006) {
            reject(new Error("Sign-in cancelled."));
          } else {
            reject(new Error(`Auth dialog error: ${ev.error}`));
          }
        });
      },
    );
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
