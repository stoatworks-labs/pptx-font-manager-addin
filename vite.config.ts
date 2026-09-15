import { defineConfig, type Plugin } from 'vite'
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

// The release version, shown in the task pane's footer and on the install page.
// Distinct from the manifest's <Version>, which Office requires to be >= 1.0 and
// which only moves when the manifest itself changes — see manifest.xml's header.
const pkg = JSON.parse(readFileSync(resolve(import.meta.dirname, 'package.json'), 'utf8')) as { version: string }

// Ship manifest.xml at the root of the built site. It is the file a user
// sideloads, and it hard-codes this origin, so the origin is the natural place
// to fetch it from (https://<host>/manifest.xml). The repo copy is the source of
// truth; this just puts the same bytes into dist/ so one `wrangler deploy`
// publishes page, pane and manifest together.
function shipManifest(): Plugin {
  return {
    name: 'ship-manifest',
    apply: 'build',
    generateBundle() {
      this.emitFile({
        type: 'asset',
        fileName: 'manifest.xml',
        source: readFileSync(resolve(import.meta.dirname, 'manifest.xml')),
      })
    },
  }
}

export default defineConfig({
  // Reference vendored files by relative path with no leading slash.
  resolve: { alias: { '@fm': FM } },
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
  plugins: [shipManifest()],
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
      input: {
        // The task pane PowerPoint loads, and the install page a visitor sees
        // at the origin's root.
        taskpane: resolve(import.meta.dirname, 'taskpane.html'),
        index: resolve(import.meta.dirname, 'index.html'),
      },
    },
  },
})
