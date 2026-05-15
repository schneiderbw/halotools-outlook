// Storage for the Teams app — pure localStorage.
//
// Unlike the Outlook add-in, Teams tabs don't have Office.context.roamingSettings.
// Teams *does* expose a per-app per-user "tab settings" object on configurable
// channel tabs, but that's specific to the tab instance and not the user, so
// localStorage scoped to the tab's origin (https://tools.iusehalo.com) is the
// right cache for tenant config + Halo OAuth tokens.
//
// Trade-off: if the user wipes site data or switches devices, they'll re-auth.
// Acceptable for v1; refresh tokens still last weeks once issued.

export interface Storage {
  get<T = unknown>(key: string): T | undefined;
  set<T = unknown>(key: string, value: T): Promise<void>;
  remove(key: string): Promise<void>;
}

class LocalStorageImpl implements Storage {
  get<T>(key: string): T | undefined {
    try {
      const raw = localStorage.getItem(key);
      if (raw == null) return undefined;
      try {
        return JSON.parse(raw) as T;
      } catch {
        return raw as unknown as T;
      }
    } catch {
      // localStorage can throw in sandboxed iframes with cookies disabled.
      // Surface as "not stored" so the app falls through to the config screen
      // rather than crashing.
      return undefined;
    }
  }
  async set<T>(key: string, value: T): Promise<void> {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch {
      /* quota or sandbox failure — caller will re-auth next session */
    }
  }
  async remove(key: string): Promise<void> {
    try {
      localStorage.removeItem(key);
    } catch {
      /* ignore */
    }
  }
}

let _storage: Storage | undefined;
export function storage(): Storage {
  if (!_storage) _storage = new LocalStorageImpl();
  return _storage;
}
