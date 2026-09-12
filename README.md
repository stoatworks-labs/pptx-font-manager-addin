# Font Manager for PowerPoint (add-in)

> **AI-assisted project.** Built with [Claude](https://claude.com/claude-code),
> directed and reviewed by a human. The checker path — get the deck bytes, scan
> them, probe what's installed — is **verified running inside real PowerPoint for
> Mac** (16.112, WKWebView). The download/bundle path reuses the web app's
> proven code, and its in-webview *delivery* is now verified on Mac too (it
> raises a native Save dialog) — see [Findings](#findings).

An Office task-pane add-in that does, inside PowerPoint, what
[pptx-font-manager](https://github.com/stoatworks-labs/pptx-font-manager) does in
a browser: work out which fonts a deck **actually** uses, check which of those
are installed on the machine looking at it, and offer the missing ones as real
downloads, metric-compatible substitutes, or a sidecar `.zip`.

One add-in covers **PowerPoint 2016+ (Windows), 2019+ (Mac), on the web, and
Microsoft 365** — the manifest validator confirms the set.

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
npm run validate      # schema-check manifest.xml
npm test              # vitest — getFileAsync reassembly
```

Sideload:

```bash
npm run sideload:mac       # drops manifest in the PowerPoint wef folder (then quit + reopen PPT)
npm run sideload:windows   # registers a trusted catalog folder (then restart PPT)
```

On Mac the custom **Home ▸ Font Manager** button may only appear after you launch
the add-in once via **Home ▸ Add-ins ▸ Developer Add-ins ▸ Font Manager** — a
known wef-sideload quirk. After that first launch it sticks.

For a hosted build, `npm run build` emits `dist/`; serve it over HTTPS and
replace the `https://localhost:3000` origin throughout `manifest.xml`. The host
must send a CSP that allows `raw.githubusercontent.com` and `cdn.jsdelivr.net`
in `connect-src` (that's where font files are fetched) — same as the web app.

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
- **Verified: file delivery works on Mac.** A blob + `<a download>` click raises
  a native **Save As** dialog in PowerPoint for Mac's WKWebView (tested
  2026-09-12: the pane offered `font-manager-delivery-test.txt` into ~/Downloads).
  WKWebView blocks the `window.open` popup fallback, but the anchor path — the one
  [`src/download.ts`](src/download.ts) uses first — works, so the fallback isn't
  needed there. It stays for any host that blocks the anchor.

## Layout

```
taskpane.html         task-pane shell
src/taskpane.ts       Office bootstrap + scan → resolve → render orchestration
src/office-file.ts    getFileAsync + slice reassembly (unit-tested)
src/render.ts         DOM rendering of the result + per-row actions
src/actions.ts        the fix half — fetch real / substitute / build bundle (reuses @fm core)
src/download.ts       hand a produced file to the user (blob, with a webview fallback)
manifest.xml          add-in manifest (Windows + Mac + web + M365)
scripts/              sideload for Mac (wef) and Windows (trusted catalog)
vendor/pptx-font-manager/  the scanner, via git subtree
```
