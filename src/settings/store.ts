export interface AddinSettings {
  port: number;
  token: string;
  styleId: string;
}

export interface KVStore {
  get(key: string): string | null;
  set(key: string, value: string): void;
}

export interface SettingsStore {
  load(): AddinSettings;
  save(patch: Partial<AddinSettings>): AddinSettings;
  clear(): void;
}

const DEFAULTS: AddinSettings = { port: 8474, token: "", styleId: "apa" };
const KEY = "notees.addin.settings.v1";

export function createSettingsStore(kv: KVStore): SettingsStore {
  return {
    load() {
      try {
        const raw = kv.get(KEY);
        if (!raw) return { ...DEFAULTS };
        const parsed = JSON.parse(raw) as Partial<AddinSettings>;
        const port = Number(parsed.port);
        return {
          port: Number.isInteger(port) && port > 0 && port < 65536 ? port : DEFAULTS.port,
          token: typeof parsed.token === "string" ? parsed.token : DEFAULTS.token,
          styleId: typeof parsed.styleId === "string" && parsed.styleId ? parsed.styleId : DEFAULTS.styleId,
        };
      } catch {
        return { ...DEFAULTS };
      }
    },
    save(patch) {
      const next = { ...this.load(), ...patch };
      kv.set(KEY, JSON.stringify(next));
      return next;
    },
    clear() {
      kv.set(KEY, JSON.stringify(DEFAULTS));
    },
  };
}
