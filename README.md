# Notees for Word

An Office.js task-pane add-in for Word: search your Notees graph, insert live
citations as Word fields, refresh them in place, and build a bibliography —
all offline-first. The pane talks **only** to the `notees serve` CLI bridge on
`http://127.0.0.1:<port>` (your own machine); it holds no server credentials
and makes no other network calls.

The add-in's statics (`dist/`) are plain files served by the Notees sync
server at `https://<host>/word/` (the server hosts whatever directory its
`NOTEES_WORD_ADDIN_DIR` points at).

## Build

```sh
pnpm install
pnpm build
```

`pnpm build` produces `dist/` with this layout — the exact tree the sync
server serves at `/word/`:

```
dist/
  index.html                         task pane entry (relative asset URLs)
  manifest.xml                       classic XML manifest, ${HOST} resolved
  assets/icon-32.png                 manifest IconUrl
  assets/icon-80.png                 manifest HighResolutionIconUrl
  assets/index-*.js / index-*.css    bundled pane
  lib/office.js                      vendored Office.js (no CDN dependency)
  lib/o15apptofilemappingtable.js    required by office.js at runtime
```

The manifest's `SourceLocation`/icon URLs are written as
`https://<host>/word/…` at build time: `scripts/copy-manifest.mjs` replaces
every `${HOST}` placeholder with the value of `NOTEES_WORD_HOST`
(default `https://localhost:8378`). Set it to the public origin where your
sync server serves `/word/`, e.g.:

```sh
NOTEES_WORD_HOST=https://<host> pnpm build
```

## How `/word/` is served (matters for the manifest URLs)

The sync server serves `NOTEES_WORD_ADDIN_DIR` read-only at `/word/*`:

- **No directory index**: `GET /word/` (empty path) returns **404**. Every URL
  in the manifest and docs must name a file explicitly — `/word/index.html`,
  not `/word/`. The manifest's `SourceLocation` is `${HOST}/word/index.html`
  accordingly.
- Explicit content types for `.html`, `.js`, `.css`, `.json`, `.xml`, `.svg`,
  `.png`, `.ico`; anything else is `application/octet-stream`.
- Word fetches `manifest.xml` and both icons from these URLs during
  sideloading, so the statics must already be reachable at `NOTEES_WORD_HOST`
  before you sideload.

## Sideloading

The manifest is the **classic XML** manifest (not the unified manifest for
Microsoft 365) because the shared-folder catalog — the primary desktop
sideload path — does not support unified manifests.

### Windows desktop Word — shared-folder catalog

1. Put `dist/` (or at least `manifest.xml`) on a folder Word can reach: a
   local directory or a network share `\\<host>\share`.
2. Trust the catalog (per-user, no admin rights):
   registry key `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs`
   — add a string value whose name and data are the folder path
   (`C:\…` or `\\<host>\share`).
3. Restart Word, then **Insert → My Add-ins → SHARED FOLDER** → Notees.
   (Choose the *Shared Folder* tab, not the Office Store.)

### Word for the web — admin-center deployment

Word for the web cannot sideload from a network share — deployment goes
through the Microsoft 365 admin center:

1. Admin center → **Settings → Integrated apps → Upload custom apps**, upload
   `dist/manifest.xml`, review the requested permissions, and consent.
2. Assign the app to the users who need it. It then appears under
   **Insert → Add-ins** in Word for the web.

A per-session "Upload my add-in" shortcut exists in Word for the web, but only
when the organization's admin has explicitly allowed custom-app sideloading —
treat the admin-center upload as the supported path.

**Capability caveat**: on hosts without the WordApi 1.5 field APIs, the pane
degrades to reading fields only — citation counts and refresh status work,
but inserting, refreshing, and restyling fields need a desktop Word.
The pane says so explicitly instead of failing silently.

### Word for Mac — wef folder

Copy `dist/manifest.xml` into the Word for Mac sideload folder:

```
~/Library/Containers/com.Microsoft/Data/documents/wef
```

Restart Word; the add-in appears in the task-pane picker.

## Pairing the pane with your workspace

The pane does nothing until it is paired with a `notees serve` bridge:

1. Run `notees serve --port 8474` (the pane's default port; any port works).
2. Copy the pairing token it prints.
3. Open the Notees pane in Word → settings → enter the port and token →
   **Connect**. The token is stored in the pane's local storage, so pairing
   persists across Word sessions.

## Development

```sh
pnpm dev        # vite dev server
pnpm test       # vitest unit suite
pnpm typecheck  # tsc --noEmit
pnpm build      # full dist/ (icons + bundle + office.js + manifest)
```

Layout: `src/bridge/` (typed fetch client for the bridge protocol),
`src/citation/` (pure, unit-tested field/OOXML/CSL logic), `src/word/` (thin
Office.js adapter), `src/ui/` (React task pane), `scripts/` (build helpers).

Host-only behavior that vitest cannot cover is checked by hand — see
[docs/smoke-checklist.md](docs/smoke-checklist.md) (sideloading, field codes,
refresh, bibliography, style picker, read-only hosts).
