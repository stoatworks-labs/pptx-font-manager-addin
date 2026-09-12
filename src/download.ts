// Handing a file to the user from inside a task pane.
//
// VERIFIED on Mac (2026-09-12): the blob + <a download> click raises a native
// Save As dialog in PowerPoint for Mac's WKWebView, and works in Chromium
// (Windows WebView2) too. WKWebView blocks the window.open popup, so that
// fallback is inert there — but the anchor path runs first and succeeds, so the
// fallback only matters for a host that blocks the anchor instead. See README
// §Findings.

export type DeliveryOutcome = 'downloaded' | 'opened-in-window' | 'blocked'

export function deliverFile(
  data: Uint8Array,
  filename: string,
  mime = 'application/octet-stream',
): DeliveryOutcome {
  // Copy into a fresh buffer so a detached/pooled view can't corrupt the Blob.
  const copy = new Uint8Array(data.length)
  copy.set(data)
  const url = URL.createObjectURL(new Blob([copy], { type: mime }))

  try {
    const a = document.createElement('a')
    if ('download' in a) {
      a.href = url
      a.download = filename
      document.body.appendChild(a)
      a.click()
      a.remove()
      setTimeout(() => URL.revokeObjectURL(url), 10_000)
      return 'downloaded'
    }
  } catch {
    /* fall through to the window path */
  }

  // Fallback: a new window/tab the user can save from. Office task panes allow
  // window.open to the same-origin blob.
  const win = window.open(url, '_blank')
  if (win) {
    setTimeout(() => URL.revokeObjectURL(url), 60_000)
    return 'opened-in-window'
  }

  URL.revokeObjectURL(url)
  return 'blocked'
}
