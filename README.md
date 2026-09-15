# Font Manager for PowerPoint (add-in)

> **AI-assisted project.** Built with [Claude](https://claude.com/claude-code),
> directed and reviewed by a human. The whole path — get the deck bytes, scan
> them, probe what's installed, fetch a missing font and save it — is **verified
> running inside real PowerPoint for Mac** (16.112, WKWebView). PowerPoint for
> Windows and PowerPoint on the web are in the manifest's host list but have
> **not been tested** — see [Status](#status).

An Office task-pane add-in that does, inside PowerPoint, what
[pptx-font-manager](https://github.com/stoatworks-labs/pptx-font-manager) does in
a browser: work out which fonts a deck **actually** uses, check which of those
are installed on the machine looking at it, and offer the missing ones as real
downloads, metric-compatible substitutes, or a sidecar `.zip`.

One add-in covers **PowerPoint 2016+ (Windows), 2019+ (Mac), on the web, and
Microsoft 365** — the manifest validator confirms the set.

## Install

The add-in is hosted at **https://pptx-font-manager-addin.stoatworks-labs.com**.
Installing it means handing PowerPoint one file, `manifest.xml`, that points at
that origin — nothing else lands on your machine, and every deploy reaches every
installed copy.

1. Download [`manifest.xml`](https://pptx-font-manager-addin.stoatworks-labs.com/manifest.xml)
   (the same file is attached to every [release](https://github.com/stoatworks-labs/pptx-font-manager-addin/releases)).
2. **PowerPoint for Mac:** copy it into
   `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` (create
   `wef` if needed), quit PowerPoint completely and reopen it, then
   **Home ▸ Add-ins ▸ Developer Add-ins ▸ Font Manager**. After that first
   launch a *Font Manager* button stays on the Home tab.
3. **PowerPoint on the web, and Microsoft 365 on Windows:**
   **Home ▸ Add-ins ▸ More Add-ins ▸ My Add-ins ▸ Upload My Add-in**, pick the
   file. Older Windows builds without *Upload My Add-in* sideload from a
   [trusted folder catalog](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) —
   `scripts/sideload-windows.ps1 -Hosted` sets one up (untested so far).

The install page at the origin's root says the same thing.

<!-- downloads:start -->
<!-- downloads:end -->

## Status

- **Mac: proven.** On PowerPoint for Mac 16.112, from the hosted build: the pane
  reads the open deck, scans it, reports installed vs missing correctly against
  a python-pptx test deck, offers Adobe links / substitutes / Google downloads
  for the right rows, and a download raises the native Save dialog and writes a
  valid TrueType file.
- **Windows: untested.** The WebView2 path is the one the add-in exists for (a
  browser tab can't answer "is this font installed" reliably on Windows — see
  pptx-font-manager's notes) and it has never been run, because no machine here
  has desktop PowerPoint for Windows. The pane carries a Windows-specific note
  about fonts installed after launch; whether it fires correctly is unverified.
- **Web: untested.** The manifest supports it and the CSP allows the Office
  hosts to frame the pane, but nobody has uploaded the manifest to PowerPoint on
  the web yet.

## Why an add-in, and its one honest limit

A browser tab can't see the deck you have open, and it can't read the machine's
font list without a permission prompt. An add-in gets the deck bytes directly
(`getFileAsync`) — no upload, no drag-drop — and sits where you're working.

What it **cannot** do that a browser also can't: query the OS font list
directly. Office.js exposes no installed-font API, and the task pane is a webview
(WebView2 on Windows, WKWebView on Mac), so the "is it installed here?" check is
the same **canvas width-probe** the web app uses. It's accurate for *is it
present*, and it can't install anything — the add-in offers downloads and a
bundle, never an install.

## How the scanner is reused

`src/core`, `src/lib`, `src/platform` and `src/data` are vendored from
pptx-font-manager under [`vendor/pptx-font-manager`](vendor/pptx-font-manager)
via `git subtree`, and imported unchanged through the `@fm/*` alias. Refresh them:

```bash
npm run core:pull    # git subtree pull from stoatworks-labs/pptx-font-manager main
```

The add-in imports the canvas inventory straight from `@fm/platform/fontcheck`,
never through `@fm/platform` — the latter pulls in the Tauri desktop bridge the
add-in doesn't want.

## Develop

```bash
npm install
npm run certs         # one-time: trust the localhost dev cert (Office needs HTTPS)
npm run dev           # vite dev server on https://localhost:3000
npm run validate      # Microsoft's schema check of manifest.xml
npm test              # vitest — getFileAsync reassembly, and that manifest.xml is the production one
```

Sideload a **dev** copy (the scripts rewrite the hosted origin to
`https://localhost:3000` on the way in — `manifest.xml` itself is never edited):

```bash
npm run sideload:mac          # drops a dev manifest in the PowerPoint wef folder (then quit + reopen PPT)
npm run sideload:mac:hosted   # the production manifest instead, to test the live build
npm run sideload:windows      # registers a trusted catalog folder (then restart PPT); -Hosted for production
```

Dev and hosted copies share the add-in `<Id>`, so PowerPoint treats them as one
add-in: sideload one or the other, never both. On Mac the custom
**Home ▸ Font Manager** button may only appear after you launch the add-in once
via **Home ▸ Add-ins ▸ Developer Add-ins ▸ Font Manager** — a known wef-sideload
quirk. After that first launch it sticks.

### Two version numbers, on purpose

- `package.json` `version` is the **release** version: the tag, the GitHub
  release, the number in the pane's footer and on the install page.
- `manifest.xml` `<Version>` is the **manifest** version. Office requires it to
  be ≥ 1.0 (the validator rejects `0.x`), and hosts cache manifests by it. Bump
  it when the manifest itself changes in a way a host must notice — URLs,
  permissions, the ribbon, icons — not on every release.

## Hosting and deploying

`manifest.xml` in this repo **is the production manifest**: every URL in it
points at `https://pptx-font-manager-addin.stoatworks-labs.com`, and
`test/manifest.test.ts` fails if any URL points elsewhere or if that hostname
stops matching `wrangler.toml`. `npm run build` emits `dist/` with the task pane,
the install page, the icons, `_headers` and a copy of the manifest, and
`.github/workflows/deploy.yml` runs `wrangler deploy` on every push to `main` and
then checks the live origin is serving that build. `cf-run npm run deploy` does
the same from a laptop.

The hostname is part of the add-in's identity — the manifest hard-codes it, so
moving it breaks every installed copy until the user sideloads a new manifest.

The CSP in [`public/_headers`](public/_headers) is deliberately not the fleet's
usual one: Office.js must load from `appsforoffice.microsoft.com`, PowerPoint on
the web frames the pane from Microsoft's Office hosts (so no `X-Frame-Options:
DENY`), and font downloads need `raw.githubusercontent.com` and
`cdn.jsdelivr.net` in `connect-src`. `npm run serve:dist` serves the built site
with those headers applied so a too-narrow policy fails locally rather than in
production.

**Releasing:** bump `package.json` (and the lockfile), commit, tag `vX.Y.Z`, push
the tag. `release.yml` checks the tag against `package.json`, validates the
manifest and attaches it to a published GitHub release. Then run
`stoatworks-backend/release/gen-downloads.py --repo pptx-font-manager-addin` to
refresh the Download block above and the website's `downloads.json`.

## Findings

Measured on PowerPoint for Mac 16.112 on 2026-09-12:

- **`getFileAsync(Compressed)` works and round-trips the deck** — a 9,165,464-byte
  deck came back as three 4 MB slices, reassembled exactly, scanned in 8 ms, and
  the canvas probe ran in 14 ms. The scanner found the 4 real fonts and ignored
  43 script-fallback names — identical to the CLI result.
- **It's slow** — that read took ~17 s. The pane scans behind a progress line and
  never blocks; don't expect it to be instant on a large deck.
- **It re-serialises the document.** On a deck PowerPoint considers slightly
  corrupt it raises its "couldn't read some content… and removed it" repair
  dialog, and returns the *repaired* bytes. Fine for font scanning; just be aware
  the bytes aren't always identical to the file on disk.
- **File delivery works on Mac.** A blob + `<a download>` click raises a native
  **Save As** dialog in PowerPoint for Mac's WKWebView. WKWebView blocks the
  `window.open` popup fallback, but the anchor path — the one
  [`src/download.ts`](src/download.ts) uses first — works, so the fallback isn't
  needed there. It stays for any host that blocks the anchor.

## Layout

```
manifest.xml          the PRODUCTION manifest — what users sideload, served at /manifest.xml
index.html            install page at the origin's root
taskpane.html         task-pane shell
src/taskpane.ts       Office bootstrap + scan → resolve → render orchestration
src/office-file.ts    getFileAsync + slice reassembly (unit-tested)
src/render.ts         DOM rendering of the result + per-row actions
src/actions.ts        the fix half — fetch real / substitute / build bundle (reuses @fm core)
src/download.ts       hand a produced file to the user (blob, with a webview fallback)
public/icons/         ribbon icons (rendered from pptx-font-manager's icon.svg)
public/_headers       CSP and caching for the Cloudflare assets runtime
wrangler.toml         the Worker and its custom domain
scripts/              sideload for Mac (wef) and Windows (trusted catalog); serve-dist.py
test/                 getFileAsync reassembly; manifest.xml is the production manifest
vendor/pptx-font-manager/  the scanner, via git subtree
```
