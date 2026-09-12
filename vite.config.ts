import { defineConfig } from 'vite'
import { homedir } from 'node:os'
import { join, resolve } from 'node:path'
import { readFileSync, existsSync } from 'node:fs'

// src/core, src/lib, src/platform and src/data are vendored from
// stoatworks-labs/pptx-font-manager via git subtree (see README). The @fm alias
// points at that tree so the scanner is imported unchanged. Refresh it with
// `npm run core:pull`.
const FM = resolve(import.meta.dirname, 'vendor/pptx-font-manager/src')

// office-addin-dev-certs writes a trusted localhost cert here; Office requires
// HTTPS even in development.
const certDir = join(homedir(), '.office-addin-dev-certs')
const certFile = join(certDir, 'localhost.crt')
const keyFile = join(certDir, 'localhost.key')
const httpsReady = existsSync(certFile) && existsSync(keyFile)

export default defineConfig({
  // Reference vendored files by relative path with no leading slash.
  resolve: { alias: { '@fm': FM } },
  server: {
    port: 3000,
    strictPort: true,
    host: 'localhost',
    https: httpsReady ? { cert: readFileSync(certFile), key: readFileSync(keyFile) } : undefined,
  },
  build: {
    outDir: 'dist',
    sourcemap: true,
    rollupOptions: {
      input: { taskpane: resolve(import.meta.dirname, 'taskpane.html') },
    },
  },
})
