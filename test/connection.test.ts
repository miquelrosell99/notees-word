import { describe, expect, it } from "vitest";
import type { Bridge, BridgeResponse } from "../src/bridge/types";
import { BridgeOfflineError } from "../src/bridge/client";
import { isRejectedCode, probeBridge } from "../src/ui/connection";

type ListStylesResponse = BridgeResponse<{ styles: { id: string; label: string }[] }>;

function fakeBridge(listStyles: () => Promise<ListStylesResponse>): Bridge {
  return {
    search: () => Promise.reject(new Error("not used")),
    renderCitation: () => Promise.reject(new Error("not used")),
    renderBibliography: () => Promise.reject(new Error("not used")),
    listStyles,
  };
}

describe("isRejectedCode", () => {
  it("classifies 401/403/origin codes as rejected", () => {
    for (const code of ["unauthorized", "forbidden", "origin_denied"]) {
      expect(isRejectedCode(code)).toBe(true);
    }
  });

  it("does not classify command-level codes as rejected", () => {
    for (const code of ["not_found", "unknown_style", "unknown", "shape"]) {
      expect(isRejectedCode(code)).toBe(false);
    }
  });
});

describe("probeBridge", () => {
  it("pairs on ok:true and carries the style list", async () => {
    const bridge = fakeBridge(() =>
      Promise.resolve({ ok: true, styles: [{ id: "apa", label: "APA" }] })
    );
    const outcome = await probeBridge(bridge);
    expect(outcome).toEqual({
      state: "paired",
      bridge,
      styles: [{ id: "apa", label: "APA" }],
    });
  });

  it("maps BridgeOfflineError to offline", async () => {
    const outcome = await probeBridge(
      fakeBridge(() => Promise.reject(new BridgeOfflineError(new Error("ECONNREFUSED"))))
    );
    expect(outcome).toEqual({ state: "offline" });
  });

  it("maps unauthorized / forbidden / origin_denied to rejected", async () => {
    for (const code of ["unauthorized", "forbidden", "origin_denied"]) {
      const outcome = await probeBridge(
        fakeBridge(() => Promise.resolve({ ok: false, error: { code, message: "no" } }))
      );
      expect(outcome).toEqual({ state: "rejected" });
    }
  });

  it("treats ok:false not_found as paired — the bridge is alive", async () => {
    const bridge = fakeBridge(() =>
      Promise.resolve({ ok: false, error: { code: "not_found", message: "no styles" } })
    );
    const outcome = await probeBridge(bridge);
    expect(outcome).toEqual({ state: "paired", bridge, styles: [] });
  });

  it("treats any other ok:false code as paired", async () => {
    const bridge = fakeBridge(() =>
      Promise.resolve({ ok: false, error: { code: "unknown_style", message: "x" } })
    );
    const outcome = await probeBridge(bridge);
    expect(outcome).toEqual({ state: "paired", bridge, styles: [] });
  });

  it("maps any other thrown error to offline", async () => {
    const outcome = await probeBridge(
      fakeBridge(() => Promise.reject(new Error("boom")))
    );
    expect(outcome).toEqual({ state: "offline" });
  });
});
