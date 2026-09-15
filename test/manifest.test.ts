import { describe, it, expect } from 'vitest'
import { readFileSync, existsSync } from 'node:fs'
import { resolve } from 'node:path'

// manifest.xml is the PRODUCTION manifest and the release artefact (see its
// header). These checks are what stop a dev copy — or a manifest pointing at a
// hostname wrangler.toml no longer deploys to — from being committed and shipped
// verbatim to every user. The Office validator (`npm run validate`) checks the
// schema; it does not know which origin is the right one.

const ROOT = resolve(import.meta.dirname, '..')
const manifest = readFileSync(resolve(ROOT, 'manifest.xml'), 'utf8')
const wrangler = readFileSync(resolve(ROOT, 'wrangler.toml'), 'utf8')

const deployedHost = /^pattern\s*=\s*"([^"]+)"/m.exec(wrangler)?.[1]
const urls = [...manifest.matchAll(/DefaultValue="(https?:[^"]+)"/g)].map((m) => new URL(m[1]!))

describe('manifest.xml is the production manifest', () => {
  it('names a deployed origin in wrangler.toml to compare against', () => {
    expect(deployedHost).toBeTruthy()
  })

  it('carries no localhost URL — the dev copy is derived by the sideload scripts, never committed', () => {
    // The header comment is allowed to *talk about* localhost; the markup is not.
    const markup = manifest.replace(/<!--[\s\S]*?-->/g, '')
    expect(markup).not.toMatch(/localhost/i)
  })

  it('points every URL at the deployed origin or the GitHub repo, over https', () => {
    expect(urls.length).toBeGreaterThan(0)
    for (const u of urls) {
      expect(u.protocol).toBe('https:')
      expect([deployedHost, 'github.com']).toContain(u.hostname)
    }
  })

  it('references only icons that exist in public/, at the paths Vite will serve them from', () => {
    const icons = urls.filter((u) => u.hostname === deployedHost && u.pathname.startsWith('/icons/'))
    expect(icons.length).toBeGreaterThanOrEqual(3)
    for (const u of icons) {
      expect(existsSync(resolve(ROOT, 'public', u.pathname.slice(1))), u.pathname).toBe(true)
    }
  })

  it('loads the task pane from the deployed origin', () => {
    const panes = urls.filter((u) => u.pathname === '/taskpane.html')
    expect(panes.length).toBeGreaterThanOrEqual(2) // DefaultSettings + the ribbon button
    for (const u of panes) expect(u.hostname).toBe(deployedHost)
  })

  it('has a manifest <Version> of at least 1.0, which Office requires', () => {
    const v = /<Version>([^<]+)<\/Version>/.exec(manifest)?.[1]
    expect(v).toBeTruthy()
    expect(Number(v!.split('.')[0])).toBeGreaterThanOrEqual(1)
  })

  it('keeps the sideload scripts pointed at the same hosted origin', () => {
    const mac = readFileSync(resolve(ROOT, 'scripts/sideload-mac.sh'), 'utf8')
    const win = readFileSync(resolve(ROOT, 'scripts/sideload-windows.ps1'), 'utf8')
    expect(mac).toContain(`https://${deployedHost}`)
    expect(win).toContain(`https://${deployedHost}`)
  })
})
