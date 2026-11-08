import { describe, expect, it } from "vitest";
import { createSettingsStore, type KVStore } from "../src/settings/store";

function memoryKV(seed?: string): KVStore & { dump(): string | null } {
  const m = new Map<string, string>();
  if (seed !== undefined) m.set("notees.addin.settings.v1", seed);
  return {
    get: (k) => m.get(k) ?? null,
    set: (k, v) => void m.set(k, v),
    dump: () => m.get("notees.addin.settings.v1") ?? null,
  };
}

describe("createSettingsStore", () => {
  it("returns defaults on empty store", () => {
    expect(createSettingsStore(memoryKV()).load()).toEqual({ port: 8377, token: "", styleId: "notees-default" });
  });

  it("round-trips a save", () => {
    const kv = memoryKV();
    const s = createSettingsStore(kv);
    const next = s.save({ port: 9000, token: "tok-abc" });
    expect(next).toEqual({ port: 9000, token: "tok-abc", styleId: "notees-default" });
    expect(s.load()).toEqual(next);
    expect(JSON.parse(kv.dump()!)).toEqual(next);
  });

  it("falls back to defaults on corrupted JSON", () => {
    expect(createSettingsStore(memoryKV("{not json")).load()).toEqual({ port: 8377, token: "", styleId: "notees-default" });
  });

  it("rejects out-of-range ports per-field, keeping valid fields", () => {
    const kv = memoryKV(JSON.stringify({ port: 70000, token: "keep-me", styleId: "apa" }));
    expect(createSettingsStore(kv).load()).toEqual({ port: 8377, token: "keep-me", styleId: "apa" });
  });

  it("rejects non-integer and non-positive ports", () => {
    for (const bad of [0, -1, 65536, 3.14, "abc"]) {
      const kv = memoryKV(JSON.stringify({ port: bad }));
      expect(createSettingsStore(kv).load().port).toBe(8377);
    }
  });

  it("survives a token round-trip and preserves it on unrelated saves", () => {
    const s = createSettingsStore(memoryKV());
    s.save({ token: "Bearer-token-123" });
    expect(s.save({ port: 8474 }).token).toBe("Bearer-token-123");
  });

  it("save merges onto stored state and clear resets to defaults", () => {
    const kv = memoryKV();
    const s = createSettingsStore(kv);
    s.save({ styleId: "mla" });
    expect(s.save({ port: 9001 })).toEqual({ port: 9001, token: "", styleId: "mla" });
    s.clear();
    expect(s.load()).toEqual({ port: 8377, token: "", styleId: "notees-default" });
    expect(JSON.parse(kv.dump()!)).toEqual({ port: 8377, token: "", styleId: "notees-default" });
  });
});
