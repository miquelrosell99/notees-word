import { useState } from "react";
import { createBridge } from "../bridge/client";
import { createSettingsStore } from "../settings/store";
import {
  bodyBibliographyWriter,
  describeBibliographyOutcome,
  insertOrReplaceBibliography,
} from "../word/bibliography";
import { describeRefreshOutcome, refreshCitations, wordFieldAccess } from "../word/refresh";
import { useConnection } from "./connection";
import { SearchPanel } from "./search-panel";
import { SettingsPanel } from "./settings-panel";
import "./app.css";

const store = createSettingsStore({
  get: (key) => localStorage.getItem(key),
  set: (key, value) => localStorage.setItem(key, value),
});

const bridgeFactory = (port: number, token: string) => createBridge({ port, token });

export function App() {
  const conn = useConnection({ createBridge: bridgeFactory, store });
  const [refreshing, setRefreshing] = useState(false);
  const [refreshNote, setRefreshNote] = useState<string | null>(null);
  const [refreshIsError, setRefreshIsError] = useState(false);
  const [bibliographyBusy, setBibliographyBusy] = useState(false);
  const [bibliographyNote, setBibliographyNote] = useState<string | null>(null);
  const [bibliographyIsError, setBibliographyIsError] = useState(false);

  async function onRefresh() {
    const bridge = conn.bridge;
    if (bridge === null || refreshing) return;
    setRefreshing(true);
    setRefreshNote(null);
    setRefreshIsError(false);
    try {
      const { access, writable } = wordFieldAccess();
      if (!writable) {
        // Word for the web et al.: fields are read-only there — report the
        // degradation up front; the document is never touched.
        setRefreshNote(
          "Refresh needs the Word desktop app — citation fields are read-only in this host.",
        );
        setRefreshIsError(true);
        return;
      }
      const outcome = await refreshCitations(bridge, access);
      setRefreshNote(describeRefreshOutcome(outcome));
    } catch (e) {
      setRefreshNote(e instanceof Error ? e.message : String(e));
      setRefreshIsError(true);
    } finally {
      setRefreshing(false);
    }
  }

  async function onInsertBibliography() {
    const bridge = conn.bridge;
    if (bridge === null || bibliographyBusy) return;
    setBibliographyBusy(true);
    setBibliographyNote(null);
    setBibliographyIsError(false);
    try {
      const { access, writable } = wordFieldAccess();
      if (!writable) {
        // Same read-only degradation as refresh: the document is never touched.
        setBibliographyNote(
          "Inserting a bibliography needs the Word desktop app — fields are read-only in this host.",
        );
        setBibliographyIsError(true);
        return;
      }
      const outcome = await insertOrReplaceBibliography(bridge, conn.settings.styleId, access, bodyBibliographyWriter());
      setBibliographyNote(describeBibliographyOutcome(outcome));
      // "no citations yet" is guidance, not a failure of the insert itself.
      setBibliographyIsError(outcome.status !== "inserted" && outcome.status !== "replaced" && outcome.status !== "no-citations");
    } catch (e) {
      setBibliographyNote(e instanceof Error ? e.message : String(e));
      setBibliographyIsError(true);
    } finally {
      setBibliographyBusy(false);
    }
  }

  if (conn.connection !== "paired") {
    return (
      <main data-testid="app-shell" className="pane">
        <header className="pane-header">
          <h1 className="pane-title">Notees</h1>
        </header>
        <SettingsPanel connection={conn} />
      </main>
    );
  }

  return (
    <main data-testid="app-shell" className="pane">
      <header className="pane-header">
        <h1 className="pane-title">Notees</h1>
        <span className="badge" data-testid="citation-count">
          citations: 0
        </span>
      </header>
      <section className="pane-body">
        <div className="refresh-row">
          <button
            type="button"
            className="secondary"
            onClick={() => void onRefresh()}
            disabled={refreshing || conn.bridge === null}
          >
            {refreshing ? "Refreshing…" : "Refresh citations"}
          </button>
          {refreshNote !== null && (
            <p className={refreshIsError ? "form-error" : "pane-note"} role={refreshIsError ? "alert" : "status"}>
              {refreshNote}
            </p>
          )}
          <button
            type="button"
            className="secondary"
            onClick={() => void onInsertBibliography()}
            disabled={bibliographyBusy || conn.bridge === null}
          >
            {bibliographyBusy ? "Inserting bibliography…" : "Insert bibliography"}
          </button>
          {bibliographyNote !== null && (
            <p
              className={bibliographyIsError ? "form-error" : "pane-note"}
              role={bibliographyIsError ? "alert" : "status"}
            >
              {bibliographyNote}
            </p>
          )}
        </div>
        {conn.bridge !== null ? (
          <SearchPanel bridge={conn.bridge} styleId={conn.settings.styleId} />
        ) : (
          <p className="pane-placeholder">
            Connected to the Notees bridge ({conn.styles.length} citation styles
            available).
          </p>
        )}
      </section>
    </main>
  );
}
