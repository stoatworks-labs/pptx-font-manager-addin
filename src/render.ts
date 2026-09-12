// Rendering the scan result into the task pane. Pure DOM; all side effects go
// through the handlers the orchestrator passes in.
import type { ResolvedFont, ResolveSummary } from '@fm/lib/resolve'

export interface RowHandlers {
  /** Download the real font (Google / Fontsource) for one missing row. */
  getReal: (r: ResolvedFont) => void
  /** Download a curated substitute family for one missing row. */
  getSubstitute: (r: ResolvedFont, family: string) => void
}

const el = <K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Partial<Record<string, string>> = {},
  ...kids: (Node | string)[]
): HTMLElementTagNameMap[K] => {
  const node = document.createElement(tag)
  for (const [k, v] of Object.entries(attrs)) if (v != null) node.setAttribute(k, v)
  for (const kid of kids) node.append(kid)
  return node
}

const TIER_LABEL: Record<string, string> = {
  slide: 'used on a slide',
  inherited: 'from a layout or theme',
  elsewhere: 'notes / charts only',
}

function statusChip(r: ResolvedFont): HTMLElement {
  if (r.state === 'embedded') return el('span', { class: 'chip emb' }, 'embedded in deck')
  if (r.state === 'installed') return el('span', { class: 'chip ok' }, 'installed')
  if (r.state === 'family-installed') return el('span', { class: 'chip warn' }, 'family only')
  return el('span', { class: 'chip miss' }, 'missing')
}

/** Build the per-row action area for a missing font. */
function actions(r: ResolvedFont, h: RowHandlers): HTMLElement {
  const wrap = el('div', { class: 'actions' })
  if (r.state !== 'missing' && r.state !== 'family-installed') return wrap

  if (r.google?.downloadable || r.fontsource) {
    const src = r.google?.downloadable ? 'Google Fonts' : 'Fontsource'
    const btn = el('button', { class: 'get', title: `Download ${r.matchedFamily ?? r.font.family} from ${src}` }, `Download (${src})`)
    btn.addEventListener('click', () => h.getReal(r))
    wrap.append(btn)
    return wrap
  }

  if (r.adobe) {
    wrap.append(
      el('a', { class: 'link', href: r.adobe.url, target: '_blank', rel: 'noopener' }, 'Adobe Fonts — activate ↗'),
    )
    return wrap
  }

  if (r.substitutes) {
    for (const s of r.substitutes.substitutes.slice(0, 2)) {
      const btn = el(
        'button',
        { class: s.metric ? 'sub metric' : 'sub', title: s.metric ? 'Same widths — keeps your line breaks' : 'Similar look — may reflow' },
        `${s.metric ? 'Substitute' : 'Similar'}: ${s.family}`,
      )
      btn.addEventListener('click', () => h.getSubstitute(r, s.family))
      wrap.append(btn)
    }
    return wrap
  }

  if (r.suggestions.length) {
    wrap.append(el('span', { class: 'muted' }, 'closest: ' + r.suggestions.map((s) => s.family).join(', ')))
  }
  return wrap
}

export function renderSummary(container: HTMLElement, s: ResolveSummary, ignoredFallbacks: number) {
  container.replaceChildren()
  const line = (n: number, label: string, cls: string) =>
    n > 0 ? el('span', { class: `stat ${cls}` }, el('b', {}, String(n)), ' ' + label) : ''
  const parts: (Node | string)[] = [
    line(s.total, 'used', 'total'),
    line(s.installed, 'installed', 'ok'),
    line(s.missing, 'missing', 'miss'),
    line(s.embedded, 'embedded', 'emb'),
  ].filter(Boolean)
  container.append(...parts)
  if (s.fixable > 0) container.append(el('div', { class: 'note' }, `${s.fixable} of the missing can be downloaded as the real font.`))
  if (s.substitutable > 0)
    container.append(
      el('div', { class: 'note' }, `${s.substitutable} have a stand-in (${s.metricSubstitutable} keep your exact line breaks).`),
    )
  if (ignoredFallbacks > 0)
    container.append(el('div', { class: 'muted small' }, `${ignoredFallbacks} script-fallback names ignored (they are theme noise, not fonts the deck uses).`))
}

export function renderRows(container: HTMLElement, resolved: ResolvedFont[], h: RowHandlers) {
  container.replaceChildren()
  const visible = resolved.filter((r) => r.font.tier !== 'elsewhere')
  for (const r of visible) {
    const row = el('div', { class: 'row' })
    row.append(
      el(
        'div',
        { class: 'name' },
        el('div', { class: 'fam' }, r.font.name),
        el('div', { class: 'tier' }, TIER_LABEL[r.font.tier] ?? r.font.tier),
      ),
      el('div', { class: 'state' }, statusChip(r)),
      actions(r, h),
    )
    container.append(row)
  }
  const hidden = resolved.length - visible.length
  if (hidden > 0)
    container.append(el('div', { class: 'muted small' }, `${hidden} more only in notes, charts or unused layouts — not shown.`))
}
