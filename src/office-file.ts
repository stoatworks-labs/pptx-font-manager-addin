// Getting the open presentation out of PowerPoint as raw .pptx bytes.
//
// `Office.context.document.getFileAsync(Compressed)` writes a temporary copy of
// the live document and hands it back in slices (≤4 MB each, or ≤64 KB on
// iPad). We reassemble them into the Uint8Array the scanner wants.
//
// Two things learned by running this on real Mac PowerPoint (see README §Findings):
//   1. It is SLOW — ~17 s for a ~9 MB deck — so it must run behind a progress
//      indicator and never block the first paint.
//   2. It RE-SERIALISES the document. On a deck PowerPoint considers slightly
//      corrupt, that triggers its "couldn't read some content… and removed it"
//      repair dialog, and the bytes returned are the *repaired* copy. That is
//      fine for font scanning (the theme and font tables survive) but it means
//      the bytes are not always identical to the file on disk.

export interface FileProgress {
  slicesReceived: number
  sliceCount: number
  bytes: number
}

/** Concatenate ordered slices into one buffer. Pure — unit tested. */
export function concatSlices(slices: Uint8Array[]): Uint8Array {
  let total = 0
  for (const s of slices) total += s.length
  const out = new Uint8Array(total)
  let off = 0
  for (const s of slices) {
    out.set(s, off)
    off += s.length
  }
  return out
}

/** Read the whole open presentation as .pptx bytes. `onProgress` is optional. */
export function getPresentationBytes(onProgress?: (p: FileProgress) => void): Promise<Uint8Array> {
  return new Promise((fulfil, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4 * 1024 * 1024 },
      (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(res.error.message))
          return
        }
        readAllSlices(res.value, onProgress).then(fulfil, reject)
      },
    )
  })
}

function readAllSlices(file: Office.File, onProgress?: (p: FileProgress) => void): Promise<Uint8Array> {
  return new Promise((fulfil, reject) => {
    const { sliceCount } = file
    const slices: Uint8Array[] = new Array(sliceCount)
    let received = 0
    let bytes = 0
    let settled = false

    const fail = (msg: string) => {
      if (settled) return
      settled = true
      file.closeAsync(() => reject(new Error(msg)))
    }

    for (let i = 0; i < sliceCount; i++) {
      file.getSliceAsync(i, (res) => {
        if (settled) return
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          fail(`Could not read slice ${i} of the presentation: ${res.error.message}`)
          return
        }
        // Office returns slice data as a Uint8Array in the browser runtimes and
        // occasionally as a plain number[] — accept both.
        const raw = res.value.data as unknown
        const slice = raw instanceof Uint8Array ? raw : Uint8Array.from(raw as number[])
        slices[i] = slice
        received++
        bytes += slice.length
        onProgress?.({ slicesReceived: received, sliceCount, bytes })
        if (received === sliceCount) {
          const out = concatSlices(slices)
          settled = true
          file.closeAsync(() => fulfil(out))
        }
      })
    }
  })
}
