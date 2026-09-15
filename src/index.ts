// The install page at the origin's root. Static apart from stamping the
// release version, so the page and the pane can never disagree about it.
document.getElementById('version')!.textContent = `v${__APP_VERSION__}`
