// chrome.storage.local adapter — the extension-side equivalent of the
// Outlook add-in's roamingSettings wrapper (../../../src/lib/storage.ts).
//
// Why chrome.storage.local instead of sync:
//   - tokens stay on the device (less surface area than syncing them to
//     the user's Google account)
//   - local has a much larger quota (~10 MB vs 100 KB)
//   - the same surface is available from popup, options page, content
//     script, AND service worker — localStorage is not available in the
//     MV3 service worker.
//
// All accessors are async to mirror chrome.storage.local's promise API.

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

const CONFIG_KEY = "halo.tenantConfig.v1";
const TOKENS_KEY = "halo.tokens.v1";

async function get<T>(key: string): Promise<T | undefined> {
  const obj = await chrome.storage.local.get(key);
  return obj[key] as T | undefined;
}

async function set<T>(key: string, value: T): Promise<void> {
  await chrome.storage.local.set({ [key]: value });
}

async function remove(key: string): Promise<void> {
  await chrome.storage.local.remove(key);
}

export function getConfig(): Promise<TenantConfig | undefined> {
  return get<TenantConfig>(CONFIG_KEY);
}

export async function setConfig(cfg: TenantConfig): Promise<void> {
  const cleaned: TenantConfig = {
    ...cfg,
    haloBaseUrl: cfg.haloBaseUrl.replace(/\/+$/, ""),
  };
  if (!/^https:\/\//i.test(cleaned.haloBaseUrl)) {
    throw new Error("HaloPSA URL must start with https://");
  }
  await set(CONFIG_KEY, cleaned);
}

export async function clearConfig(): Promise<void> {
  await remove(CONFIG_KEY);
  await remove(TOKENS_KEY);
}

export function getTokens(): Promise<StoredTokens | undefined> {
  return get<StoredTokens>(TOKENS_KEY);
}

export async function setTokens(t: StoredTokens): Promise<void> {
  await set(TOKENS_KEY, t);
}

export async function clearTokens(): Promise<void> {
  await remove(TOKENS_KEY);
}
