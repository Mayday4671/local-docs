import { extname } from 'node:path'
import type { Category, DocumentRecord, ThemePreference, DocumentState } from '../shared/types'
import { validateState } from './document-state'

export const MAX_OBJECT_BYTES = 100 * 1024 * 1024
export const MAX_BACKUP_BYTES = 10 * 1024 * 1024 * 1024
export const MAX_MANIFEST_BYTES = 64 * 1024 * 1024
export const MAX_ENTRIES = 100_000
export interface BackupDocument extends DocumentRecord {
  blobHash: string
  state?: DocumentState
}
export interface BackupVersion {
  id: string
  documentId: string
  blobHash: string
  createdAt: string
  reason: string
}
export interface BackupManifest {
  format: 'local-docs-backup'
  version: 1 | 2
  createdAt: string
  theme: ThemePreference
  categories: Category[]
  documents: BackupDocument[]
  versions: BackupVersion[]
  objects: { hash: string; size: number }[]
}

function requireValid(condition: unknown): asserts condition {
  if (!condition) throw new Error('备份清单不完整或包含无效数据，未修改当前文档库。')
}
const isRecord = (v: unknown): v is Record<string, unknown> =>
  !!v && typeof v === 'object' && !Array.isArray(v)
const isId = (v: unknown): v is string =>
  typeof v === 'string' && /^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/.test(v)
const isHash = (v: unknown): v is string => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
const isDate = (v: unknown): v is string =>
  typeof v === 'string' &&
  /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(v) &&
  Number.isFinite(Date.parse(v)) &&
  new Date(v).toISOString() === v
const nullableDate = (v: unknown) => v === null || isDate(v)
const isName = (v: unknown): v is string =>
  typeof v === 'string' &&
  v.trim() === v &&
  v.length > 0 &&
  v.length <= 189 &&
  !/[<>:"/\\|?*\x00-\x1f]/.test(v) &&
  !/[. ]$/.test(v) &&
  !/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(v)
const isSize = (v: unknown): v is number =>
  typeof v === 'number' && Number.isSafeInteger(v) && v >= 0 && v <= MAX_OBJECT_BYTES

/** Parse only the portable data format. No SQL or archive paths are trusted. */
export function validateManifest(value: unknown): BackupManifest {
  requireValid(isRecord(value))
  requireValid(
    value.format === 'local-docs-backup' &&
      (value.version === 1 || value.version === 2) &&
      isDate(value.createdAt),
  )
  requireValid(['light', 'dark', 'system'].includes(String(value.theme)))
  for (const key of ['categories', 'documents', 'versions', 'objects'])
    requireValid(Array.isArray(value[key]) && value[key].length <= MAX_ENTRIES)
  const m = value as unknown as BackupManifest
  const categories = new Map<string, Category>()
  const siblings = new Set<string>()
  for (const c of m.categories) {
    requireValid(
      isRecord(c) && isId(c.id) && isName(c.name) && (c.parentId === null || isId(c.parentId)),
    )
    const key = `${c.parentId ?? ''}/${c.name}`
    requireValid(!categories.has(c.id) && !siblings.has(key))
    categories.set(c.id, c)
    siblings.add(key)
  }
  // Bound nesting and reject cycles; topological order is used during restore.
  const ordered: Category[] = []
  const depths = new Map<string, number>()
  for (const c of m.categories) {
    const chain: Category[] = []
    const visiting = new Set<string>()
    let current: Category | undefined = c
    while (current && !depths.has(current.id)) {
      requireValid(!visiting.has(current.id) && chain.length < 128)
      visiting.add(current.id)
      chain.push(current)
      requireValid(current.parentId === null || categories.has(current.parentId))
      current = current.parentId === null ? undefined : categories.get(current.parentId)
    }
    let depth = current ? depths.get(current.id)! : 0
    for (const item of chain.reverse()) {
      requireValid(++depth <= 128)
      depths.set(item.id, depth)
      ordered.push(item)
    }
  }
  const objects = new Map<string, number>()
  let total = 0
  for (const o of m.objects) {
    requireValid(isRecord(o) && isHash(o.hash) && isSize(o.size) && !objects.has(o.hash))
    total += o.size
    requireValid(total <= MAX_BACKUP_BYTES)
    objects.set(o.hash, o.size)
  }
  const documents = new Map<string, BackupDocument>()
  const referenced = new Set<string>()
  for (const d of m.documents) {
    requireValid(isRecord(d) && isId(d.id) && isName(d.name) && !documents.has(d.id))
    requireValid(
      ['.md', '.markdown', '.docx', '.xlsx'].includes(d.extension) &&
        extname(d.name).toLowerCase() === d.extension,
    )
    requireValid(d.categoryId === null || categories.has(d.categoryId))
    requireValid(
      typeof d.favorite === 'boolean' &&
        isSize(d.size) &&
        isHash(d.blobHash) &&
        objects.get(d.blobHash) === d.size,
    )
    requireValid(
      isDate(d.createdAt) &&
        isDate(d.updatedAt) &&
        nullableDate(d.openedAt) &&
        nullableDate(d.deletedAt),
    )
    requireValid(Number.isSafeInteger(d.revision) && d.revision >= 1)
    requireValid(
      Array.isArray(d.tags) &&
        d.tags.length <= 30 &&
        d.tags.every((t) => typeof t === 'string' && t.trim().length > 0 && t.length <= 40),
    )
    requireValid(typeof d.notes === 'string' && d.notes.length <= 5000)
    if (['.md', '.markdown'].includes(d.extension)) requireValid(d.size <= 5 * 1024 * 1024)
    documents.set(d.id, d)
    referenced.add(d.blobHash)
    if (m.version === 2) {
      requireValid(d.state !== undefined)
      const state = validateState(d.state!, d.id)
      if (state.draft)
        requireValid(
          state.draft.kind === (['.md', '.markdown'].includes(d.extension) ? 'markdown' : 'office'),
        )
      for (const a of state.attachments) {
        requireValid(objects.get(a.hash) === a.size)
        referenced.add(a.hash)
      }
      if (state.attachments.length) requireValid(['.md', '.markdown'].includes(d.extension))
    } else requireValid(d.state === undefined)
  }
  const versionIds = new Set<string>()
  for (const v of m.versions) {
    requireValid(isRecord(v) && isId(v.id) && !versionIds.has(v.id) && documents.has(v.documentId))
    requireValid(isHash(v.blobHash) && objects.has(v.blobHash) && isDate(v.createdAt))
    requireValid(typeof v.reason === 'string' && v.reason.length > 0 && v.reason.length <= 200)
    if (['.md', '.markdown'].includes(documents.get(v.documentId)!.extension))
      requireValid(objects.get(v.blobHash)! <= 5 * 1024 * 1024)
    versionIds.add(v.id)
    referenced.add(v.blobHash)
  }
  requireValid(referenced.size === objects.size)
  return { ...m, categories: ordered }
}
