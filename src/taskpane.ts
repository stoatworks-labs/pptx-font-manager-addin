// Font Manager for PowerPoint — task pane entry point.
//
// Reuses the pptx-font-manager scanner and resolver verbatim (the @fm/* alias
// resolves into the vendored subtree). The only add-in-specific parts are how
// the deck bytes arrive (getFileAsync, not a file input) and how a produced
// file leaves (a task-pane download, not a browser one).
import { scanPptx } from '@fm/core/scan'
import type { ScanResult } from '@fm/core/types'
import { defaultInventory, fontListSnapshottedAtLaunch } from '@fm/platform/fontcheck'
import { resolveAll, summarize, type ResolvedFont } from '@fm/lib/resolve'

import { getPresentationBytes } from './office-file'
import { deliverFile } from './download'
import { renderRows, renderSummary, type RowHandlers } from './render'
import { buildDeckBundle, fetchReal, fetchSubstitute } from './actions'

const $ = (id: string) => document.getElementById(id)!
const scanBtn = () => $('scan') as HTMLButtonElement
const bundleBtn = () => $('bundle') as HTMLButtonElement

let scan: ScanResult | null = null
let resolved: ResolvedFont[] = []
let deckName = 'presentation.pptx'

function setBusy(msg: string | null) {
  $('status').textContent = msg ?? ''
  scanBtn().disabled = !!msg
  bundleBtn().disabled = !!msg || resolved.length === 0
}

function setError(msg: string) {
  const box = $('error')
  box.textContent = msg
  box.hidden = !msg
}

function deckNameFromOffice(): string {
  const url = (Office.context.document as { url?: string }).url
  if (url) {
    const base = url.split(/[\\/]/).pop()
    if (base) return decodeURIComponent(base)
  }
  return 'presentation.pptx'
}

const handlers: RowHandlers = {
  getReal: (r) => void deliverOne(() => fetchReal(r), r.font.name),
  getSubstitute: (r, family) => void deliverOne(() => fetchSubstitute(r, family), family),
}

async function deliverOne(fetcher: () => Promise<Array<{ filename: string; data: Uint8Array }>>, label: string) {
  setError('')
  setBusy(`Downloading ${label}…`)
  try {
    const faces = await fetcher()
    if (faces.length === 0) {
      setError(`Nothing downloadable was found for ${label}.`)
      return
    }
    let blocked = false
    for (const f of faces) {
      const outcome = deliverFile(f.data, f.filename, 'font/ttf')
      if (outcome === 'blocked') blocked = true
    }
    if (blocked) setError('Your browser blocked the download. Use “Build font bundle” instead, or try Chrome/Edge.')
  } catch (e) {
    setError((e as Error).message)
  } finally {
    setBusy(null)
  }
}

async function runScan() {
  setError('')
  setBusy('Reading the presentation from PowerPoint…')
  $('results').hidden = true
  try {
    deckName = deckNameFromOffice()
    const bytes = await getPresentationBytes((p) => {
      setBusy(`Reading the presentation… ${(p.bytes / 1_048_576).toFixed(1)} MB (${p.slicesReceived}/${p.sliceCount})`)
    })

    setBusy('Scanning fonts…')
    scan = scanPptx(bytes)
    const inventory = defaultInventory() // canvas probing — no permission prompt
    resolved = resolveAll(scan.fonts, inventory)
    const summary = summarize(resolved)

    renderSummary($('summary'), summary, scan.ignoredFallbacks.length)
    renderRows($('rows'), resolved, handlers)
    renderWindowsNote(summary.missing)
    $('results').hidden = false
  } catch (e) {
    setError('Could not scan the presentation: ' + (e as Error).message)
  } finally {
    setBusy(null)
  }
}

/**
 * On Windows, Chromium reads the OS font list once per browser process — and
 * PowerPoint's task pane is that browser. A font installed after PowerPoint
 * launched will keep reading as missing here until PowerPoint is fully
 * relaunched (which restarts the webview). PowerPoint itself sees it, so the
 * deck is fine; only this check is stale. See the pptx-font-manager AGENTS.md.
 */
function renderWindowsNote(missing: number) {
  const box = $('winnote')
  if (missing > 0 && fontListSnapshottedAtLaunch()) {
    box.textContent =
      'Windows: if you install one of these now, this pane will keep showing it as missing until you quit and reopen PowerPoint. PowerPoint itself will use the font straight away.'
    box.hidden = false
  } else {
    box.hidden = true
  }
}

async function makeBundle() {
  setError('')
  setBusy('Building the font bundle…')
  try {
    const build = await buildDeckBundle(deckName, resolved, (label) => setBusy(label))
    if (build.included === 0) {
      setError('Nothing could be bundled — every font is either already installed, embedded, or unavailable to share.')
      return
    }
    const outcome = deliverFile(build.zip, build.filename, 'application/zip')
    if (outcome === 'blocked') {
      setError('Your browser blocked the download. Try Chrome, Edge, or PowerPoint on the web.')
    } else {
      setBusy(null)
      $('status').textContent = `Bundle ready: ${build.included} file(s)${build.unavailable.length ? `, ${build.unavailable.length} could not be included (see MANIFEST.txt)` : ''}.`
      return
    }
  } catch (e) {
    setError((e as Error).message)
  } finally {
    if (scanBtn().disabled) setBusy(null)
  }
}

$('version').textContent = `Font Manager v${__APP_VERSION__}`

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    $('status').textContent = 'This add-in runs in PowerPoint.'
    return
  }
  $('status').textContent = ''
  scanBtn().disabled = false
  scanBtn().addEventListener('click', () => void runScan())
  bundleBtn().addEventListener('click', () => void makeBundle())
})
