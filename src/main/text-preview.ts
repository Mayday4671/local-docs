import { fileKind } from '../shared/file-types'

/** Text preview is bounded and never rewrites the source bytes. */
export function plainTextPreview(bytes: Uint8Array): string | null {
  if (bytes.byteLength > 5 * 1024 ** 2) return null
  let text: string
  try {
    if (bytes[0] === 0xff && bytes[1] === 0xfe)
      text = new TextDecoder('utf-16le', { fatal: true }).decode(bytes)
    else if (bytes[0] === 0xfe && bytes[1] === 0xff)
      text = new TextDecoder('utf-16be', { fatal: true }).decode(bytes)
    else {
      try {
        text = new TextDecoder('utf-8', { fatal: true }).decode(bytes)
      } catch {
        text = new TextDecoder('gb18030', { fatal: true }).decode(bytes)
      }
    }
    // A misleading text extension must not flood the preview/index with binary data.
    return /[\x00-\x08\x0e-\x1f]/.test(text) ? null : text
  } catch {
    return null
  }
}
export function plainTextIndex(extension: string, bytes: Uint8Array): string {
  return fileKind(extension) === 'text' ? (plainTextPreview(bytes) ?? '') : ''
}
