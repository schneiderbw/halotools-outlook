// Storage abstraction over Office.context.roamingSettings.
//
// roamingSettings is per-mailbox and roams with the Outlook profile across devices.
// Total size limit is ~32KB across all settings combined — plenty for our config + tokens.
// Note: writes are async and require saveAsync() to persist.

const HAS_OFFICE = () =>
  typeof Office !== "undefined" && Office.context && Office.context.roamingSettings;

export interface Storage {
  get<T = unknown>(key: string): T | undefined;
  set<T = unknown>(key: string, value: T): Promise<void>;
  remove(key: string): Promise<void>;
}

class RoamingStorage implements Storage {
  get<T>(key: string): T | undefined {
    const v = Office.context.roamingSettings.get(key);
    return v === undefined ? undefined : (v as T);
  }
  set<T>(key: string, value: T): Promise<void> {
    Office.context.roamingSettings.set(key, value);
    return this.persist();
  }
  remove(key: string): Promise<void> {
    Office.context.roamingSettings.remove(key);
    return this.persist();
  }
  private persist(): Promise<void> {
    return new Promise((resolve, reject) => {
      Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error?.message ?? "roamingSettings save failed"));
      });
    });
  }
}

class LocalFallback implements Storage {
  get<T>(key: string): T | undefined {
    const raw = localStorage.getItem(key);
    if (raw == null) return undefined;
    try {
      return JSON.parse(raw) as T;
    } catch {
      return raw as unknown as T;
    }
  }
  async set<T>(key: string, value: T): Promise<void> {
    localStorage.setItem(key, JSON.stringify(value));
  }
  async remove(key: string): Promise<void> {
    localStorage.removeItem(key);
  }
}

// Lazy singleton. Picks roaming if Office is present, otherwise localStorage (for dev outside Outlook).
let _storage: Storage | undefined;
export function storage(): Storage {
  if (!_storage) _storage = HAS_OFFICE() ? new RoamingStorage() : new LocalFallback();
  return _storage;
}
