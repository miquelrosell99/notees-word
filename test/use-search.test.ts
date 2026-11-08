import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { BridgeOfflineError, BridgeShapeError } from "../src/bridge/client";
import type { BridgeResponse, BridgeResult } from "../src/bridge/types";
import { SearchEngine } from "../src/ui/use-search";

type SearchResponse = BridgeResponse<{ results: BridgeResult[] }>;

const RESULTS: BridgeResult[] = [
  {
    objectId: "node-1",
    title: "Doe on things",
    authors: ["Doe, Jane"],
    year: 2020,
    citekey: "doe2020things",
  },
];

function makeEngine(searchFn: (q: string) => Promise<SearchResponse>): SearchEngine {
  return new SearchEngine({ searchFn, debounceMs: 300 });
}

describe("SearchEngine", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("stays idle for an empty query and never calls searchFn", async () => {
    const searchFn = vi.fn((_q: string): Promise<SearchResponse> => Promise.resolve({ ok: true, results: [] }));
    const engine = makeEngine(searchFn);
    engine.setQuery("   ");
    await vi.advanceTimersByTimeAsync(1_000);
    expect(searchFn).not.toHaveBeenCalled();
    expect(engine.getView().state).toBe("idle");
  });

  it("debounces: typing 'doe' calls searchFn once after 300 ms and lands in results", async () => {
    const searchFn = vi.fn((_q: string): Promise<SearchResponse> => Promise.resolve({ ok: true, results: RESULTS }));
    const engine = makeEngine(searchFn);
    engine.setQuery("doe");
    expect(searchFn).not.toHaveBeenCalled();
    await vi.advanceTimersByTimeAsync(299);
    expect(searchFn).not.toHaveBeenCalled();
    await vi.advanceTimersByTimeAsync(1);
    expect(searchFn).toHaveBeenCalledTimes(1);
    expect(searchFn).toHaveBeenCalledWith("doe");
    expect(engine.getView()).toEqual({
      state: "results",
      query: "doe",
      results: RESULTS,
      error: null,
    });
  });

  it("lands in error carrying the code on ok:false", async () => {
    const engine = makeEngine(() =>
      Promise.resolve({ ok: false, error: { code: "not_found", message: "no matches" } }),
    );
    engine.setQuery("zzz");
    await vi.advanceTimersByTimeAsync(300);
    const view = engine.getView();
    expect(view.state).toBe("error");
    expect(view.error).toEqual({ kind: "bridge", code: "not_found", message: "no matches" });
  });

  it("treats the synthetic { code: 'shape' } outcome as a shape error, not a server error", async () => {
    const engine = makeEngine(() =>
      Promise.resolve({ ok: false, error: { code: "shape", message: "results not an array" } }),
    );
    engine.setQuery("doe");
    await vi.advanceTimersByTimeAsync(300);
    expect(engine.getView().error).toEqual({ kind: "shape", message: "results not an array" });
  });

  it("cancels earlier calls on rapid typing — only the last query is applied", async () => {
    const calls: string[] = [];
    const engine = makeEngine((q) => {
      calls.push(q);
      return Promise.resolve({ ok: true, results: RESULTS });
    });
    engine.setQuery("d");
    await vi.advanceTimersByTimeAsync(100);
    engine.setQuery("do");
    await vi.advanceTimersByTimeAsync(100);
    engine.setQuery("doe");
    await vi.advanceTimersByTimeAsync(300);
    expect(calls).toEqual(["doe"]);
    expect(engine.getView().query).toBe("doe");
  });

  it("ignores a superseded in-flight response", async () => {
    const deferreds: Array<{ q: string; resolve: (r: SearchResponse) => void }> = [];
    const engine = makeEngine(
      (q) =>
        new Promise<SearchResponse>((resolve) => {
          deferreds.push({ q, resolve });
        }),
    );
    engine.setQuery("doe");
    await vi.advanceTimersByTimeAsync(300);
    engine.setQuery("smith");
    await vi.advanceTimersByTimeAsync(300);
    expect(deferreds.map((d) => d.q)).toEqual(["doe", "smith"]);

    // Late-arriving first response must not clobber the newer query.
    deferreds[0].resolve({ ok: true, results: RESULTS });
    await vi.advanceTimersByTimeAsync(0);
    expect(engine.getView().query).toBe("smith");
    expect(engine.getView().state).toBe("loading");

    deferreds[1].resolve({ ok: true, results: [] });
    await vi.advanceTimersByTimeAsync(0);
    expect(engine.getView()).toEqual({ state: "results", query: "smith", results: [], error: null });
  });

  it("maps a thrown BridgeOfflineError to a dedicated offline error", async () => {
    const engine = makeEngine(() => Promise.reject(new BridgeOfflineError(new Error("ECONNREFUSED"))));
    engine.setQuery("doe");
    await vi.advanceTimersByTimeAsync(300);
    expect(engine.getView().state).toBe("error");
    expect(engine.getView().error).toEqual({ kind: "offline" });
  });

  it("maps a thrown BridgeShapeError to a shape error", async () => {
    const engine = makeEngine(() => Promise.reject(new BridgeShapeError("missing ok flag")));
    engine.setQuery("doe");
    await vi.advanceTimersByTimeAsync(300);
    expect(engine.getView().error).toEqual({ kind: "shape", message: "missing ok flag" });
  });
});
