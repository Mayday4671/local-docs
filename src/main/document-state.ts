import type { DocumentState, AnnotationInput } from '../shared/types'
const id = (v: unknown) =>
  typeof v === 'string' && /^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/.test(v)
const text = (v: unknown, max: number) => typeof v === 'string' && v.length <= max
export function validateAnnotation(a: AnnotationInput): void {
  if (
    !a ||
    !text(a.quote, 8000) ||
    !a.quote.trim() ||
    !text(a.prefix, 80) ||
    !text(a.suffix, 80) ||
    !text(a.scope, 200) ||
    !a.scope ||
    !text(a.note, 5000) ||
    !Number.isSafeInteger(a.offset) ||
    a.offset < 0 ||
    a.offset > 10_000_000 ||
    !['yellow', 'green', 'blue', 'pink'].includes(a.color) ||
    (a.id !== undefined && !id(a.id))
  )
    throw new Error('标记内容或位置无效。')
  const l = a.location
  if (
    l &&
    !(l.kind === 'paragraph'
      ? text(l.part, 100) && Number.isSafeInteger(l.index) && l.index >= 0 && l.index <= 50000
      : l.kind === 'cell' && text(l.sheet, 100) && /^[A-Z]{1,3}[1-9]\d{0,6}$/.test(l.address))
  )
    throw new Error('标记位置无效。')
}
export function validateState(s: DocumentState, documentId: string): DocumentState {
  if (
    !s ||
    !Array.isArray(s.annotations) ||
    s.annotations.length > 500 ||
    !Array.isArray(s.attachments) ||
    s.attachments.length > 200
  )
    throw new Error('文档标记或附件数据无效。')
  const used = new Set<string>()
  for (const a of s.annotations) {
    validateAnnotation(a)
    if (
      !id(a.id) ||
      used.has(a.id) ||
      a.documentId !== documentId ||
      !Number.isSafeInteger(a.revision) ||
      a.revision < 1 ||
      !Number.isFinite(Date.parse(a.createdAt))
    )
      throw new Error('标记记录无效。')
    used.add(a.id)
  }
  used.clear()
  for (const a of s.attachments) {
    if (
      !a ||
      !id(a.id) ||
      used.has(a.id) ||
      !text(a.name, 180) ||
      !/^[a-f0-9]{64}$/.test(a.hash) ||
      !Number.isSafeInteger(a.size) ||
      a.size < 1 ||
      a.size > 10 * 1024 ** 2 ||
      !['image/png', 'image/jpeg', 'image/webp'].includes(a.mime) ||
      !['.png', '.jpg', '.webp'].includes(a.extension) ||
      { 'image/png': '.png', 'image/jpeg': '.jpg', 'image/webp': '.webp' }[a.mime] !== a.extension
    )
      throw new Error('图片附件数据无效。')
    used.add(a.id)
  }
  if (
    s.draft !== null &&
    (!s.draft ||
      !text(s.draft.text, 5 * 1024 ** 2) ||
      Buffer.byteLength(s.draft.text) > 5 * 1024 ** 2 ||
      !['markdown', 'office'].includes(s.draft.kind) ||
      !Number.isSafeInteger(s.draft.revision) ||
      s.draft.revision < 1 ||
      !Number.isFinite(Date.parse(s.draft.updatedAt)))
  )
    throw new Error('草稿数据无效。')
  if (s.draft?.kind === 'office') {
    let changes: unknown
    try {
      changes = s.draft.text ? JSON.parse(s.draft.text) : []
    } catch {
      throw new Error('Office 草稿数据无效。')
    }
    const keys = new Set<string>()
    if (!Array.isArray(changes) || changes.length > 2000) throw new Error('Office 草稿数据无效。')
    for (const c of changes) {
      if (
        !c ||
        typeof c.key !== 'string' ||
        !c.key ||
        c.key.length > 300 ||
        keys.has(c.key) ||
        typeof c.text !== 'string' ||
        c.text.length > 32767 ||
        /[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(c.text)
      )
        throw new Error('Office 草稿数据无效。')
      keys.add(c.key)
    }
  }
  return s
}
export const emptyState = (): DocumentState => ({ annotations: [], attachments: [], draft: null })
