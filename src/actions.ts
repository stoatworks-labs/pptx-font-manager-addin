// The "fix" half, reusing pptx-font-manager's core verbatim.
//
// This is a browser-class host (WKWebView / WebView2): no Tauri, and no
// font-install capability. So the add-in offers exactly what the web build
// offers a browser — real downloads from Google Fonts / Fontsource, curated
// metric-compatible substitutes, and a sidecar .zip — and never claims to
// install anything. The desktop/Local-Font-Access branches of App.tsx are
// deliberately absent.

import { BASIC_LATIN_TOTAL } from '@fm/core/sfnt'
import { buildBundle, bundleFilename, type BundleEntry } from '@fm/core/bundle'
import { adobeBundleNote } from '@fm/core/adobe'
import { fetchGoogleFaces } from '@fm/core/google'
import { downloadPlan, fetchPlanFaces, type ResolvedFont } from '@fm/lib/resolve'

export interface BundleBuild {
  zip: Uint8Array
  filename: string
  included: number
  unavailable: Array<{ name: string; reason: string }>
}

/** Fetch a single missing font's real files (Google or Fontsource). */
export async function fetchReal(r: ResolvedFont): Promise<Array<{ filename: string; data: Uint8Array }>> {
  const plan = downloadPlan(r)
  if (!plan) return []
  return fetchPlanFaces(plan, r.font.weight, r.font.italic)
}

/** Fetch a curated metric-compatible / similar substitute by family. */
export async function fetchSubstitute(
  r: ResolvedFont,
  family: string,
): Promise<Array<{ filename: string; data: Uint8Array }>> {
  return fetchGoogleFaces(family, [r.font.weight], { italics: r.font.italic })
}

/**
 * Build the sidecar bundle for a whole deck. Progress is reported per font so
 * the ~seconds of network fetches don't look like a hang.
 *
 * Mirrors App.tsx's makeBundle minus the desktop paths: embedded faces that
 * were extractable go in; compressed-embedded and Adobe fonts are recorded as
 * unavailable with the reason; downloadable fonts are fetched; installed fonts
 * are skipped (a browser add-in can't read their files to include them).
 */
export async function buildDeckBundle(
  deckName: string,
  resolved: ResolvedFont[],
  onProgress?: (label: string) => void,
): Promise<BundleBuild> {
  const entries: BundleEntry[] = []
  const unavailable: Array<{ name: string; reason: string }> = []

  for (const r of resolved) {
    const font = r.font

    if (font.embedded?.extracted?.length) {
      for (const face of font.embedded.extracted) {
        entries.push({
          filename: face.filename,
          data: face.data,
          family: font.family,
          source: 'embedded',
          license: 'Embedded in the presentation — terms unknown',
          redistributable: false,
          provenance: `extracted from ${face.part}`,
          partialCoverage:
            face.coverage && face.coverage.basicLatin < BASIC_LATIN_TOTAL
              ? { basicLatin: face.coverage.basicLatin, total: BASIC_LATIN_TOTAL }
              : undefined,
        })
      }
      continue
    }
    if (font.embedded) {
      unavailable.push({
        name: font.name,
        reason:
          'Embedded in the presentation, but the data is MicroType Express compressed ' +
          '(PowerPoint’s own format) and cannot be unpacked into an installable file. ' +
          'It travels with the deck, so it will render wherever the .pptx goes.',
      })
      continue
    }

    const plan = downloadPlan(r)
    if (plan) {
      onProgress?.(`Fetching ${plan.family}…`)
      try {
        const faces = await fetchPlanFaces(plan, font.weight, font.italic)
        const label = plan.source === 'google' ? 'Google Fonts' : 'Fontsource'
        for (const f of faces) {
          entries.push({
            filename: f.filename,
            data: f.data,
            family: f.family,
            source: plan.source,
            license: f.license,
            redistributable: true,
            provenance: `${label} (${f.license})${f.variable ? ', variable font' : ''}`,
          })
        }
      } catch (e) {
        unavailable.push({ name: font.name, reason: (e as Error).message })
      }
      continue
    }

    // Missing and not downloadable. An Adobe font gets its licence explanation;
    // anything else is simply not available to bundle (the user may still have
    // it locally, which a browser add-in cannot read into the zip).
    if (r.state === 'missing') {
      unavailable.push({
        name: font.name,
        reason: r.adobe
          ? adobeBundleNote(r.adobe)
          : 'No redistributable copy was found. If you have this font, install it on the target machine directly.',
      })
    }
  }

  const zip = buildBundle({ deckName, entries, unavailable })
  return { zip, filename: bundleFilename(deckName), included: entries.length, unavailable }
}
