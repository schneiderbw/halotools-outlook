import { storage } from "./storage";

// Each MSP's admin registers their own Halo Connect app on their own HaloPSA instance,
// then enters the resulting Client ID + Halo base URL on first launch.
// We never see the secret; we only do PKCE public-client flows.

const KEY = "halo.tenantConfig.v1";

export interface TenantConfig {
  /** Base URL of the user's HaloPSA instance, e.g. https://halo.example.com (no trailing slash) */
  haloBaseUrl: string;
  /** OAuth Client ID for the Halo Connect Application the admin registered on their tenant */
  clientId: string;
  /** Optional scope override; defaults to "all" for v1 simplicity */
  scope?: string;
}

export interface StoredTokens {
  accessToken: string;
  refreshToken: string;
  /** Epoch millis when accessToken expires. */
  expiresAt: number;
  scope?: string;
}

const TOKENS_KEY = "halo.tokens.v1";

export function getConfig(): TenantConfig | undefined {
  return storage().get<TenantConfig>(KEY);
}

export async function setConfig(cfg: TenantConfig): Promise<void> {
  // Normalize: strip trailing slash, ensure https://
  const cleaned: TenantConfig = {
    ...cfg,
    haloBaseUrl: cfg.haloBaseUrl.replace(/\/+$/, ""),
  };
  if (!/^https:\/\//i.test(cleaned.haloBaseUrl)) {
    throw new Error("HaloPSA URL must start with https://");
  }
  await storage().set(KEY, cleaned);
}

export async function clearConfig(): Promise<void> {
  await storage().remove(KEY);
  await storage().remove(TOKENS_KEY);
}

export function getTokens(): StoredTokens | undefined {
  return storage().get<StoredTokens>(TOKENS_KEY);
}

export async function setTokens(t: StoredTokens): Promise<void> {
  await storage().set(TOKENS_KEY, t);
}

export async function clearTokens(): Promise<void> {
  await storage().remove(TOKENS_KEY);
}
