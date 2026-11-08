import { createBridge } from "../bridge/client";
import { createSettingsStore } from "../settings/store";
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
