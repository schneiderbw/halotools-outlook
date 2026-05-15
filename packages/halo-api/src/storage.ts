// Storage abstraction for tenant config + tokens.
//
// The package itself is environment-agnostic — consumers pick an adapter and
// install it via setStorage(). Outlook uses roamingSettings; browser-extension
// or Teams consumers can pass localStorage / chrome.storage / their own.

export interface Storage {
  get<T = unknown>(key: string): T | undefined;
  set<T = unknown>(key: string, value: T): Promise<void>;
  remove(key: string): Promise<void>;
}

// ---- Office.context.roamingSettings adapter ----

// Minimal structural type so the package doesn't need to depend on @types/office-js.
// Consumers in an Office context pass the real Office.context.roamingSettings — it
// satisfies this interface.
export interface RoamingSettingsLike {
  get(key: string): unknown;
  set(key: string, value: unknown): void;
  remove(key: string): void;
  saveAsync(
    cb: (result: {
      status: unknown;
      error?: { message?: string };
    }) => void,
  ): void;
}

/**
 * Wrap Office.context.roamingSettings as a Storage.
 * Pass an object that quacks like roamingSettings. In Outlook code this is just
 * `Office.context.roamingSettings`.
 *
 * The saveAsync result.status comparison is done by reference against the
 * Office enum value the caller provides (or by truthy/falsy convention if the
 * caller hasn't passed `succeededStatus`). To avoid pulling Office types here,
 * we accept any value and trust the caller to forward the SDK's success status.
 */
export function roamingSettingsStorage(
  roaming: RoamingSettingsLike,
  succeededStatus?: unknown,
): Storage {
  const persist = (): Promise<void> =>
    new Promise((resolve, reject) => {
      roaming.saveAsync((result) => {
        const ok =
          succeededStatus !== undefined
            ? result.status === succeededStatus
            : // Office.AsyncResultStatus.Succeeded === "succeeded" in current SDKs.
              // Fall back to that string when the caller didn't pass the enum.
              result.status === "succeeded" || result.status === 0;
        if (ok) resolve();
        else reject(new Error(result.error?.message ?? "roamingSettings save failed"));
      });
    });

  return {
    get<T>(key: string): T | undefined {
      const v = roaming.get(key);
      return v === undefined ? undefined : (v as T);
    },
    set<T>(key: string, value: T): Promise<void> {
      roaming.set(key, value);
      return persist();
    },
    remove(key: string): Promise<void> {
      roaming.remove(key);
      return persist();
    },
  };
}

// ---- Browser localStorage adapter ----

export function localStorageStorage(): Storage {
  return {
    get<T>(key: string): T | undefined {
      const raw = localStorage.getItem(key);
      if (raw == null) return undefined;
      try {
        return JSON.parse(raw) as T;
      } catch {
        return raw as unknown as T;
      }
    },
    async set<T>(key: string, value: T): Promise<void> {
      localStorage.setItem(key, JSON.stringify(value));
    },
    async remove(key: string): Promise<void> {
      localStorage.removeItem(key);
    },
  };
}

// ---- Module-scoped current adapter ----

let _storage: Storage | undefined;

/** Install the storage adapter the package will use. Call this once at startup. */
export function setStorage(adapter: Storage): void {
  _storage = adapter;
}

/** Read the currently-installed adapter. Throws if no adapter has been set. */
export function storage(): Storage {
  if (!_storage) {
    throw new Error(
      "@iusehalo/halo-api: no storage adapter installed. Call setStorage(...) before using the API.",
    );
  }
  return _storage;
}
