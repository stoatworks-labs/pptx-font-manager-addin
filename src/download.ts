// Handing a file to the user from inside a task pane.
//
// UNKNOWN, flagged for verification: the standard blob + <a download> click is
// what the web app uses and works in Chromium (Windows WebView2). WKWebView on
// Mac has historically blocked programmatic downloads, so this may be inert in
// Mac PowerPoint. We try the anchor first and fall back to opening the blob in
// a new window (which the user can then save). Until this is tested on Mac,
// treat a bundle download there as unproven — see README §Findings.

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
