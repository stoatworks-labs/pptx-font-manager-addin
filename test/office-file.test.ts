import { describe, it, expect } from 'vitest'
import { concatSlices } from '../src/office-file'

describe('concatSlices — getFileAsync reassembly', () => {
  it('joins ordered slices into one buffer, preserving bytes and order', () => {
    const a = new Uint8Array([1, 2, 3])
    const b = new Uint8Array([4, 5])
    const c = new Uint8Array([6, 7, 8, 9])
    const out = concatSlices([a, b, c])
    expect(Array.from(out)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9])
    expect(out.length).toBe(9)
  })

  it('handles a single slice and empty slices', () => {
    expect(Array.from(concatSlices([new Uint8Array([42])]))).toEqual([42])
    expect(concatSlices([]).length).toBe(0)
    expect(Array.from(concatSlices([new Uint8Array([]), new Uint8Array([7]), new Uint8Array([])]))).toEqual([7])
  })

  it('reassembles a realistic multi-megabyte, 4 MB-sliced payload exactly', () => {
    // Mirrors the Mac run: a ~9 MB file arrives as three 4 MB(-ish) slices.
    const sizes = [4 * 1024 * 1024, 4 * 1024 * 1024, 1_165_464]
    const slices = sizes.map((n, i) => new Uint8Array(n).fill(i + 1))
    const out = concatSlices(slices)
    expect(out.length).toBe(sizes.reduce((a, b) => a + b, 0))
    // boundaries land where expected
    expect(out[0]).toBe(1)
    expect(out[sizes[0]! - 1]).toBe(1)
    expect(out[sizes[0]!]).toBe(2)
    expect(out[sizes[0]! + sizes[1]!]).toBe(3)
    expect(out[out.length - 1]).toBe(3)
  })
})
