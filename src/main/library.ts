import { createHash, randomUUID } from 'node:crypto'
import {
  closeSync,
  constants,
  copyFileSync,
  existsSync,
  fsyncSync,
  lstatSync,
  mkdirSync,
  openSync,
  readFileSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path'
import { backup, DatabaseSync } from 'node:sqlite'
import { fileKind, canEditFile } from '../shared/file-types'
import { plainTextPreview, plainTextIndex } from './text-preview'
import { emptyState, validateAnnotation, validateState } from './document-state'
import { newOfficeFile } from './office-edit'
import { validateManifest, type BackupManifest, type BackupVersion } from './backup-format'
import type {
  Category,
  DocumentContent,
  DocumentRecord,
  LibrarySnapshot,
  VersionRecord,
  OfficeData,
  SearchHit,
  ThemePreference,
  Annotation,
  AnnotationInput,
  DocumentState,
  Draft,
  Attachment,
} from '../shared/types'

const MAX_FILE_SIZE = 100 * 1024 * 1024
const MAX_TEXT_SIZE = 5 * 1024 * 1024

type Row = Record<string, string | number | null>

function nameOf(value: unknown): string {
  if (typeof value !== 'string') throw new Error('请输入名称。')
  const name = value.trim()
  if (
    !name ||
    name.length > 180 ||
    /[<>:"/\\|?*\x00-\x1f]/.test(name) ||
    /[. ]$/.test(name) ||
    /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(name)
  ) {
    throw new Error('名称不能为空，且不能包含路径、特殊字符或 Windows 保留名称。')
  }
  return name
}

function idOf(value: unknown): string {
  if (typeof value !== 'string' || !/^[\da-f-]{36}$/.test(value))
    throw new Error('无效的文档或分类编号。')
  return value
}

function decodeText(bytes: Uint8Array): string {
  if (bytes.byteLength > MAX_TEXT_SIZE) throw new Error('Markdown 文件最大支持 5 MB。')
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes)
  } catch {
    throw new Error('请先将 Markdown 文件转换为 UTF-8 编码，再导入。')
  }
}

export class Library {
  private readonly db: DatabaseSync
  private readonly objectsPath: string

  constructor(readonly root: string) {
    if (!isAbsolute(root)) throw new Error('文档库必须使用绝对路径。')
    this.objectsPath = join(root, 'objects')
    mkdirSync(this.objectsPath, { recursive: true })
    this.db = new DatabaseSync(join(root, 'library.sqlite'))
    this.db.exec(
      'PRAGMA foreign_keys = ON; PRAGMA journal_mode = WAL; PRAGMA synchronous = FULL; PRAGMA busy_timeout = 5000;',
    )
    const version = this.db.prepare('PRAGMA user_version').get() as { user_version: number }
    if (version.user_version > 5) {
      this.db.close()
      throw new Error('此文档库由较新版本创建，请升级应用。')
    }
    if (version.user_version === 0)
      this.db.exec(`
      BEGIN IMMEDIATE;
      CREATE TABLE categories (id TEXT PRIMARY KEY, name TEXT NOT NULL UNIQUE);
      CREATE TABLE documents (
        id TEXT PRIMARY KEY, name TEXT NOT NULL, extension TEXT NOT NULL,
        category_id TEXT REFERENCES categories(id) ON DELETE SET NULL,
        favorite INTEGER NOT NULL DEFAULT 0, size INTEGER NOT NULL,
        blob_hash TEXT NOT NULL, search_text TEXT NOT NULL DEFAULT '',
        created_at TEXT NOT NULL, updated_at TEXT NOT NULL, opened_at TEXT, deleted_at TEXT,
        revision INTEGER NOT NULL DEFAULT 1
      );
      CREATE TABLE versions (
        id TEXT PRIMARY KEY, document_id TEXT NOT NULL REFERENCES documents(id),
        blob_hash TEXT NOT NULL, created_at TEXT NOT NULL, reason TEXT NOT NULL
      );
      CREATE INDEX documents_category ON documents(category_id);
      CREATE INDEX versions_document ON versions(document_id);
      PRAGMA user_version = 1;
      COMMIT;
    `)
    if (version.user_version < 2)
      this.db.exec(`
      BEGIN IMMEDIATE;
      CREATE TABLE office_cache (
        document_id TEXT PRIMARY KEY REFERENCES documents(id) ON DELETE CASCADE,
        blob_hash TEXT NOT NULL, data TEXT, error TEXT
      );
      PRAGMA user_version = 2;
      COMMIT;
    `)
    if (version.user_version < 3) {
      // Rebuild the old globally-unique category names while keeping document references.
      this.db.exec('PRAGMA foreign_keys = OFF;')
      try {
        this.db.exec(`
          BEGIN IMMEDIATE;
          CREATE TABLE categories_new (
            id TEXT PRIMARY KEY, name TEXT NOT NULL,
            parent_id TEXT REFERENCES categories_new(id) ON DELETE RESTRICT
          );
          INSERT INTO categories_new (id, name) SELECT id, name FROM categories;
          DROP TABLE categories;
          ALTER TABLE categories_new RENAME TO categories;
          CREATE UNIQUE INDEX category_siblings ON categories(COALESCE(parent_id, ''), name);
          ALTER TABLE documents ADD COLUMN tags TEXT NOT NULL DEFAULT '[]';
          ALTER TABLE documents ADD COLUMN notes TEXT NOT NULL DEFAULT '';
          PRAGMA user_version = 3;
          COMMIT;
        `)
      } catch (error) {
        this.db.exec('ROLLBACK;')
        this.db.close()
        throw error
      }
      this.db.exec('PRAGMA foreign_keys = ON;')
    }
    if (version.user_version < 4) {
      // Old Excel caches contain unformatted numeric values. Rebuild only derived data.
      this.db.exec(`
        BEGIN IMMEDIATE;
        CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
        DELETE FROM office_cache;
        PRAGMA user_version = 4;
        COMMIT;
      `)
    }
    if (version.user_version < 5)
      this.db.exec(`BEGIN IMMEDIATE;
      CREATE TABLE IF NOT EXISTS document_state (document_id TEXT PRIMARY KEY REFERENCES documents(id) ON DELETE CASCADE, data TEXT NOT NULL);
      PRAGMA user_version = 5; COMMIT;`)
  }

  getTheme(): ThemePreference {
    const value = this.db.prepare("SELECT value FROM settings WHERE key = 'theme'").get()?.value
    return value === 'light' || value === 'dark' ? value : 'system'
  }

  setTheme(theme: ThemePreference): void {
    if (!['system', 'light', 'dark'].includes(theme)) throw new Error('无效的主题选项。')
    this.db
      .prepare(
        "INSERT INTO settings (key, value) VALUES ('theme', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
      )
      .run(theme)
  }

  close(): void {
    this.db.close()
  }

  /** SQLite's online backup includes committed WAL pages; never copy a live database file. */
  async copyDatabase(destination: string): Promise<void> {
    await backup(this.db, destination)
  }

  checkIntegrity(): void {
    const rows = this.db.prepare('PRAGMA quick_check').all()
    if (rows.length !== 1 || rows[0].quick_check !== 'ok') throw new Error('文档库数据库校验失败。')
  }

  backupManifest(): BackupManifest {
    return this.transaction(() => {
      const snapshot = this.snapshot()
      const hashes = this.db.prepare('SELECT id, blob_hash FROM documents').all() as {
        id: string
        blob_hash: string
      }[]
      const byId = new Map(hashes.map((row) => [row.id, row.blob_hash]))
      const documents = snapshot.documents.map((d) => ({
        ...d,
        blobHash: byId.get(d.id)!,
        state: this.documentState(d.id),
      }))
      const versions = this.db
        .prepare(
          'SELECT id, document_id AS documentId, blob_hash AS blobHash, created_at AS createdAt, reason FROM versions ORDER BY rowid',
        )
        .all() as unknown as BackupVersion[]
      const allHashes = new Set([
        ...documents.map((d) => d.blobHash),
        ...versions.map((v) => v.blobHash),
        ...documents.flatMap((d) => d.state.attachments.map((a) => a.hash)),
      ])
      const objects = [...allHashes].map((hash) => {
        const info = lstatSync(this.objectPath(hash))
        if (!info.isFile() || info.isSymbolicLink())
          throw new Error('文档内容缺失或存储异常，无法备份。')
        return { hash, size: info.size }
      })
      return validateManifest({
        format: 'local-docs-backup',
        version: 2,
        createdAt: new Date().toISOString(),
        theme: this.getTheme(),
        categories: snapshot.categories,
        documents,
        versions,
        objects,
      })
    })
  }

  backupObjectPath(hash: string): string {
    return this.objectPath(hash)
  }

  installBackupObject(hash: string, bytes: Buffer): void {
    if (bytes.length > MAX_FILE_SIZE || createHash('sha256').update(bytes).digest('hex') !== hash)
      throw new Error('备份文件内容校验失败，未恢复文档库。')
    this.putObject(bytes)
  }

  /** Objects must already be installed and flushed. SQLite rolls the entire replacement back on failure. */
  applyBackup(value: BackupManifest): void {
    const m = validateManifest(value)
    this.transaction(() => {
      this.db.exec('DELETE FROM office_cache; DELETE FROM versions; DELETE FROM documents;')
      // Child rows must be removed before their parents (ON DELETE RESTRICT).
      const old = this.snapshot().categories
      while (old.length) {
        const parents = new Set(old.map((c) => c.parentId))
        const leaves = old.filter((c) => !parents.has(c.id))
        if (!leaves.length) throw new Error('当前分类结构异常，已取消恢复。')
        for (const c of leaves) {
          this.db.prepare('DELETE FROM categories WHERE id = ?').run(c.id)
          old.splice(
            old.findIndex((x) => x.id === c.id),
            1,
          )
        }
      }
      this.db.exec('DELETE FROM settings;')
      for (const c of m.categories)
        this.db.prepare('INSERT INTO categories VALUES (?, ?, ?)').run(c.id, c.name, c.parentId)
      const insert = this.db.prepare(
        `INSERT INTO documents (id, name, extension, category_id, favorite, size, blob_hash, search_text, created_at, updated_at, opened_at, deleted_at, revision, tags, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      )
      for (const d of m.documents) {
        const text = ['.md', '.markdown'].includes(d.extension)
          ? decodeText(readFileSync(this.objectPath(d.blobHash)))
          : fileKind(d.extension) === 'text' && d.size <= MAX_TEXT_SIZE
            ? plainTextIndex(d.extension, readFileSync(this.objectPath(d.blobHash)))
            : ''
        insert.run(
          d.id,
          d.name,
          d.extension,
          d.categoryId,
          Number(d.favorite),
          d.size,
          d.blobHash,
          text,
          d.createdAt,
          d.updatedAt,
          d.openedAt,
          d.deletedAt,
          d.revision,
          JSON.stringify(d.tags),
          d.notes,
        )
        this.writeState(d.id, d.state ?? emptyState())
      }
      const insertVersion = this.db.prepare('INSERT INTO versions VALUES (?, ?, ?, ?, ?)')
      for (const v of m.versions)
        insertVersion.run(v.id, v.documentId, v.blobHash, v.createdAt, v.reason)
      this.setTheme(m.theme)
    })
  }

  private transaction<T>(operation: () => T): T {
    this.db.exec('BEGIN IMMEDIATE')
    try {
      const result = operation()
      this.db.exec('COMMIT')
      return result
    } catch (error) {
      this.db.exec('ROLLBACK')
      throw error
    }
  }

  documentState(id: string): DocumentState {
    this.row(id)
    const row = this.db.prepare('SELECT data FROM document_state WHERE document_id = ?').get(id)
    return row ? validateState(JSON.parse(String(row.data)), id) : emptyState()
  }
  private writeState(id: string, state: DocumentState): void {
    validateState(state, id)
    this.db
      .prepare(
        'INSERT INTO document_state VALUES (?, ?) ON CONFLICT(document_id) DO UPDATE SET data = excluded.data',
      )
      .run(id, JSON.stringify(state))
  }
  annotations(id?: string): Annotation[] {
    if (id !== undefined) return this.documentState(id).annotations
    return (
      this.db
        .prepare(
          'SELECT s.data FROM document_state s JOIN documents d ON d.id=s.document_id WHERE d.deleted_at IS NULL',
        )
        .all() as { data: string }[]
    ).flatMap((r) => (JSON.parse(r.data) as DocumentState).annotations)
  }
  saveAnnotation(id: string, input: AnnotationInput, revision: number): Annotation {
    const row = this.row(id)
    if (row.deleted_at || row.revision !== revision)
      throw new Error('文档已变化或进入回收站，请重新打开后标记。')
    validateAnnotation(input)
    const state = this.documentState(id)
    const old = input.id ? state.annotations.find((a) => a.id === input.id) : undefined
    if (input.id && !old) throw new Error('标记不存在。')
    const annotation: Annotation = {
      ...input,
      id: old?.id ?? randomUUID(),
      documentId: id,
      revision,
      createdAt: old?.createdAt ?? new Date().toISOString(),
    }
    state.annotations = [...state.annotations.filter((a) => a.id !== annotation.id), annotation]
    this.writeState(id, state)
    return annotation
  }
  removeAnnotation(id: string, annotationId: string): void {
    if (this.row(id).deleted_at) throw new Error('请先恢复文档。')
    const state = this.documentState(id)
    state.annotations = state.annotations.filter((a) => a.id !== idOf(annotationId))
    this.writeState(id, state)
  }
  saveDraft(id: string, text: string, revision: number, kind: Draft['kind']): void {
    const row = this.row(id)
    if (row.deleted_at || row.revision !== revision)
      throw new Error('文档已变化，草稿未覆盖旧版本，请先保存或重新打开。')
    if (
      !canEditFile(String(row.extension)) ||
      (kind === 'markdown') !== ['.md', '.markdown'].includes(String(row.extension))
    )
      throw new Error('草稿类型与文档不匹配。')
    const state = this.documentState(id)
    state.draft = { text, revision, kind, updatedAt: new Date().toISOString() }
    this.writeState(id, state)
  }
  discardDraft(id: string): void {
    const state = this.documentState(id)
    state.draft = null
    this.writeState(id, state)
  }
  addImage(id: string, bytes: Buffer, mime: Attachment['mime'], filename = '图片'): Attachment {
    const row = this.row(id)
    if (row.deleted_at || !['.md', '.markdown'].includes(String(row.extension)))
      throw new Error('请在 Markdown 文档中添加图片。')
    const state = this.documentState(id)
    if (bytes.length > 10 * 1024 ** 2 || !bytes.length) throw new Error('图片最大支持 10 MB。')
    const attachment: Attachment = {
      id: randomUUID(),
      name: nameOf(filename),
      hash: createHash('sha256').update(bytes).digest('hex'),
      size: bytes.length,
      mime,
      extension: { 'image/png': '.png', 'image/jpeg': '.jpg', 'image/webp': '.webp' }[
        mime
      ] as Attachment['extension'],
    }
    state.attachments.push(attachment)
    validateState(state, id)
    this.putObject(bytes)
    this.writeState(id, state)
    return attachment
  }
  attachmentPath(id: string, attachmentId: string): { path: string; mime: string } {
    const a = this.documentState(id).attachments.find((a) => a.id === idOf(attachmentId))
    if (!a) throw new Error('附件不存在。')
    return { path: this.objectPath(a.hash), mime: a.mime }
  }
  setFavorites(ids: string[], favorite: boolean): void {
    if (!Array.isArray(ids) || !ids.length || ids.length > 10000 || typeof favorite !== 'boolean')
      throw new Error('无效的收藏操作。')
    for (const id of ids) if (this.row(id).deleted_at) throw new Error('请先恢复回收站文件。')
    this.transaction(() => {
      const set = this.db.prepare('UPDATE documents SET favorite=? WHERE id=?')
      for (const id of new Set(ids)) set.run(Number(favorite), id)
    })
  }
  saveOfficeBytes(id: string, bytes: Buffer, revision: number): DocumentRecord {
    const row = this.row(id)
    if (row.deleted_at || row.revision !== revision)
      throw new Error('文件已变化，请重新打开后保存。')
    if (!['.docx', '.xlsx'].includes(String(row.extension)) || bytes.length > MAX_FILE_SIZE)
      throw new Error('无效的 Office 内容。')
    const hash = this.putObject(bytes)
    if (hash === row.blob_hash) {
      this.discardDraft(id)
      return this.record(row)
    }
    this.transaction(() => {
      const now = new Date().toISOString()
      this.db
        .prepare('INSERT INTO versions VALUES (?, ?, ?, ?, ?)')
        .run(randomUUID(), id, hash, now, 'Office 编辑保存')
      this.db
        .prepare(
          'UPDATE documents SET blob_hash=?, size=?, search_text=?, updated_at=?, revision=revision+1 WHERE id=?',
        )
        .run(hash, bytes.length, '', now, id)
      this.db.prepare('DELETE FROM office_cache WHERE document_id=?').run(id)
      this.discardDraft(id)
    })
    return this.record(this.row(id))
  }

  private objectPath(hash: string): string {
    if (!/^[a-f0-9]{64}$/.test(hash)) throw new Error('文档存储信息不正确。')
    return join(this.objectsPath, hash)
  }

  // Immutable content is flushed before committing its pointer in SQLite.
  // An interrupted operation can leave an unreferenced object, never a replaced original.
  private putObject(bytes: Buffer): string {
    const hash = createHash('sha256').update(bytes).digest('hex')
    const path = this.objectPath(hash)
    if (existsSync(path)) {
      const actual = createHash('sha256').update(readFileSync(path)).digest('hex')
      if (actual !== hash) throw new Error('文档内容校验失败，请保留文档库并检查磁盘。')
      return hash
    }
    const handle = openSync(path, 'wx')
    try {
      writeFileSync(handle, bytes)
      fsyncSync(handle)
    } finally {
      closeSync(handle)
    }
    return hash
  }

  private row(id: string): Row {
    const row = this.db.prepare('SELECT * FROM documents WHERE id = ?').get(idOf(id)) as
      Row | undefined
    if (!row) throw new Error('文档不存在。')
    return row
  }

  private record(row: Row): DocumentRecord {
    return {
      id: String(row.id),
      name: String(row.name),
      extension: String(row.extension),
      categoryId: row.category_id as string | null,
      favorite: Boolean(row.favorite),
      size: Number(row.size),
      createdAt: String(row.created_at),
      updatedAt: String(row.updated_at),
      openedAt: row.opened_at as string | null,
      deletedAt: row.deleted_at as string | null,
      revision: Number(row.revision),
      tags: JSON.parse(String(row.tags)),
      notes: String(row.notes),
    }
  }

  private category(id: string | null): string | null {
    if (id === null) return null
    if (!this.db.prepare('SELECT id FROM categories WHERE id = ?').get(idOf(id)))
      throw new Error('分类不存在。')
    return id
  }

  snapshot(): LibrarySnapshot {
    return {
      documents: (
        this.db.prepare('SELECT * FROM documents ORDER BY updated_at DESC, id').all() as Row[]
      ).map((row) => this.record(row)),
      categories: this.db
        .prepare('SELECT id, name, parent_id AS parentId FROM categories ORDER BY rowid')
        .all() as unknown as Category[],
      storagePath: this.root,
      annotationCounts: Object.fromEntries(
        (
          this.db.prepare('SELECT document_id, data FROM document_state').all() as {
            document_id: string
            data: string
          }[]
        )
          .map((r) => [r.document_id, (JSON.parse(r.data) as DocumentState).annotations.length])
          .filter(([, count]) => Number(count) > 0),
      ),
    }
  }

  search(query: string): string[] {
    if (typeof query !== 'string' || query.length > 200) throw new Error('搜索词最多 200 个字符。')
    const pattern = `%${query.trim().replace(/[!%_]/g, '!$&')}%`
    return (
      this.db
        .prepare(
          "SELECT id FROM documents WHERE name LIKE ? ESCAPE '!' OR search_text LIKE ? ESCAPE '!' OR tags LIKE ? ESCAPE '!' OR notes LIKE ? ESCAPE '!' OR EXISTS (SELECT 1 FROM document_state s, json_each(s.data, '$.annotations') a WHERE s.document_id = documents.id AND (json_extract(a.value, '$.quote') LIKE ? ESCAPE '!' OR json_extract(a.value, '$.note') LIKE ? ESCAPE '!'))",
        )
        .all(pattern, pattern, pattern, pattern, pattern, pattern) as { id: string }[]
    ).map((row) => row.id)
  }

  officeSource(id: string): { path: string; hash: string; extension: string } {
    const row = this.row(id)
    if (!['.docx', '.xlsx'].includes(String(row.extension)))
      throw new Error('该文件不是 Office 文档。')
    return {
      path: this.objectPath(String(row.blob_hash)),
      hash: String(row.blob_hash),
      extension: String(row.extension),
    }
  }

  officeCache(id: string, hash: string): { data: OfficeData | null; error: string | null } | null {
    const row = this.db
      .prepare('SELECT data, error FROM office_cache WHERE document_id = ? AND blob_hash = ?')
      .get(id, hash) as { data: string | null; error: string | null } | undefined
    return row ? { data: row.data ? JSON.parse(row.data) : null, error: row.error } : null
  }

  cacheOffice(id: string, hash: string, data: OfficeData | null, error: string | null): void {
    if (this.row(id).blob_hash !== hash) return
    this.transaction(() => {
      this.db
        .prepare(
          'INSERT OR REPLACE INTO office_cache (document_id, blob_hash, data, error) VALUES (?, ?, ?, ?)',
        )
        .run(id, hash, data ? JSON.stringify(data) : null, error)
      this.db
        .prepare('UPDATE documents SET search_text = ? WHERE id = ?')
        .run(data?.blocks.map((block) => block.text).join('\n') ?? '', id)
    })
  }

  pendingOfficeIds(): string[] {
    return (
      this.db
        .prepare(
          `SELECT d.id FROM documents d LEFT JOIN office_cache c ON c.document_id = d.id
      WHERE d.extension IN ('.docx', '.xlsx') AND (c.blob_hash IS NULL OR c.blob_hash <> d.blob_hash)`,
        )
        .all() as { id: string }[]
    ).map((row) => row.id)
  }

  searchResults(query: string): SearchHit[] {
    const ids = this.search(query)
    const needle = query.trim().toLocaleLowerCase()
    const snippet = (text: string): string => {
      const at = Math.max(0, text.toLocaleLowerCase().indexOf(needle))
      const start = Math.max(0, at - 30)
      return `${start ? '…' : ''}${text.slice(start, at + needle.length + 80)}${at + needle.length + 80 < text.length ? '…' : ''}`
    }
    return ids.map((id) => {
      const row = this.row(id)
      if (String(row.name).toLocaleLowerCase().includes(needle))
        return { id, source: 'name', snippet: String(row.name) }
      const metadata = [...JSON.parse(String(row.tags)), String(row.notes)].join(' · ')
      if (metadata.toLocaleLowerCase().includes(needle))
        return { id, source: 'metadata', snippet: snippet(metadata) }
      const annotation = this.annotations(id).find((a) =>
        `${a.quote}\n${a.note}`.toLocaleLowerCase().includes(needle),
      )
      if (annotation)
        return {
          id,
          source: 'metadata',
          snippet: `标记：${snippet(annotation.quote + ' ' + annotation.note)}`,
          location: annotation.location,
        }
      const block = this.officeCache(id, String(row.blob_hash))?.data?.blocks.find((block) =>
        block.text.toLocaleLowerCase().includes(needle),
      )
      return {
        id,
        source: 'content',
        snippet: snippet(block?.text ?? String(row.search_text)),
        ...(block ? { location: block.location } : {}),
      }
    })
  }

  createCategory(name: string, parentId: string | null = null): Category {
    this.category(parentId)
    this.checkCategoryDepth(parentId, 1)
    const category = { id: randomUUID(), name: nameOf(name), parentId }
    if (
      this.db
        .prepare('SELECT id FROM categories WHERE name = ? AND parent_id IS ?')
        .get(category.name, parentId)
    )
      throw new Error('已有同名分类。')
    this.db
      .prepare('INSERT INTO categories (id, name, parent_id) VALUES (?, ?, ?)')
      .run(category.id, category.name, parentId)
    return category
  }

  private checkCategoryDepth(parentId: string | null, subtreeDepth: number): void {
    let depth = subtreeDepth
    while (parentId !== null) {
      if (++depth > 128) throw new Error('分类层级最多支持 128 层。')
      const row = this.db.prepare('SELECT parent_id FROM categories WHERE id = ?').get(parentId)
      if (!row) throw new Error('分类不存在。')
      parentId = row.parent_id as string | null
    }
  }

  private categorySubtree(id: string): { id: string; depth: number }[] {
    idOf(id)
    this.category(id)
    return this.db
      .prepare(
        `WITH RECURSIVE subtree(id, depth) AS (
      SELECT id, 1 FROM categories WHERE id = ?
      UNION ALL SELECT c.id, s.depth + 1 FROM categories c JOIN subtree s ON c.parent_id = s.id
    ) SELECT * FROM subtree ORDER BY depth`,
      )
      .all(id) as { id: string; depth: number }[]
  }

  updateCategory(id: string, patch: { name: string; parentId: string | null }): void {
    if (!patch || typeof patch !== 'object') throw new Error('无效的分类修改。')
    const name = nameOf(patch.name)
    const parentId = this.category(patch.parentId)
    const subtree = this.categorySubtree(id)
    if (subtree.some((c) => c.id === parentId)) throw new Error('不能将分类移动到自身或子分类中。')
    this.checkCategoryDepth(parentId, Math.max(...subtree.map((c) => c.depth)))
    if (
      this.db
        .prepare('SELECT id FROM categories WHERE name = ? AND parent_id IS ? AND id <> ?')
        .get(name, parentId, id)
    )
      throw new Error('目标位置已有同名分类，请修改名称。')
    this.db
      .prepare('UPDATE categories SET name = ?, parent_id = ? WHERE id = ?')
      .run(name, parentId, id)
  }

  deleteCategory(id: string, destination: string | null): void {
    this.category(destination)
    const subtree = this.categorySubtree(id)
    if (subtree.some((c) => c.id === destination))
      throw new Error('文件存放位置不能是将要删除的分类。')
    this.transaction(() => {
      const move = this.db.prepare('UPDATE documents SET category_id = ? WHERE category_id = ?')
      const remove = this.db.prepare('DELETE FROM categories WHERE id = ?')
      for (const c of [...subtree].reverse()) {
        move.run(destination, c.id)
        remove.run(c.id)
      }
    })
  }

  moveDocuments(ids: string[], categoryId: string | null): void {
    if (!Array.isArray(ids) || !ids.length || ids.length > 10000)
      throw new Error('请选择 1–10000 份文件。')
    this.category(categoryId)
    const unique = [...new Set(ids)]
    for (const id of unique)
      if (this.row(id).deleted_at) throw new Error('请先恢复回收站中的文件。')
    this.transaction(() => {
      const move = this.db.prepare('UPDATE documents SET category_id = ? WHERE id = ?')
      for (const id of unique) move.run(categoryId, id)
    })
  }

  /** Folder imports never replace existing documents. Repeated identical copies may be skipped. */
  importBytes(
    filename: string,
    bytes: Buffer,
    categoryId: string | null,
    duplicates: 'skip' | 'keep',
  ): DocumentRecord | null {
    this.category(categoryId)
    const name = nameOf(filename)
    const extension = extname(name).toLowerCase()
    if (bytes.length > MAX_FILE_SIZE) throw new Error('文件超过 100 MB。')
    if (extension === '.md' || extension === '.markdown') decodeText(bytes)
    if (duplicates !== 'skip' && duplicates !== 'keep') throw new Error('无效的重复文件处理方式。')
    const rows = this.db
      .prepare(
        'SELECT name, blob_hash FROM documents WHERE category_id IS ? AND deleted_at IS NULL',
      )
      .all(categoryId) as { name: string; blob_hash: string }[]
    const hash = createHash('sha256').update(bytes).digest('hex')
    const stem = extension ? name.slice(0, -extension.length) : name
    // Generated copy names also count as duplicates on retry, even when the original name had different content.
    const isCopyName = (candidate: string) => {
      const lower = candidate.toLocaleLowerCase()
      if (lower === name.toLocaleLowerCase()) return true
      if (!lower.endsWith(extension)) return false
      const copy = (extension ? candidate.slice(0, -extension.length) : candidate).match(
        /（([2-9]|[1-9]\d+)）$/,
      )
      if (!copy) return false
      const suffix = `${copy[0]}${extension}`
      return lower === `${stem.slice(0, 180 - suffix.length)}${suffix}`.toLocaleLowerCase()
    }
    if (duplicates === 'skip' && rows.some((r) => r.blob_hash === hash && isCopyName(r.name)))
      return null
    const used = new Set(rows.map((r) => r.name.toLocaleLowerCase()))
    let uniqueName = name
    for (let n = 2; used.has(uniqueName.toLocaleLowerCase()); n++) {
      const suffix = `（${n}）${extension}`
      uniqueName = `${stem.slice(0, 180 - suffix.length)}${suffix}`
    }
    return this.insert(uniqueName, extension, bytes, categoryId, '文件夹导入')
  }

  documentPath(id: string): string {
    return this.objectPath(String(this.row(id).blob_hash))
  }

  private insert(
    name: string,
    extension: string,
    bytes: Buffer,
    categoryId: string | null,
    reason: string,
  ): DocumentRecord {
    this.category(categoryId)
    const text = ['.md', '.markdown'].includes(extension)
      ? decodeText(bytes)
      : plainTextIndex(extension, bytes)
    const hash = this.putObject(bytes)
    const id = randomUUID()
    const now = new Date().toISOString()
    this.transaction(() => {
      this.db
        .prepare(
          `INSERT INTO documents (id, name, extension, category_id, size, blob_hash, search_text, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        )
        .run(id, name, extension, categoryId, bytes.byteLength, hash, text, now, now)
      this.db
        .prepare('INSERT INTO versions VALUES (?, ?, ?, ?, ?)')
        .run(randomUUID(), id, hash, now, reason)
    })
    return this.record(this.row(id))
  }

  importFile(path: string, categoryId: string | null): DocumentRecord {
    if (!isAbsolute(path)) throw new Error('请选择本地文件。')
    const info = lstatSync(path)
    if (!info.isFile() || info.isSymbolicLink()) throw new Error('请选择普通文件。')
    if (info.size > MAX_FILE_SIZE) throw new Error('当前开发版单文件最大支持 100 MB。')
    const extension = extname(path).toLowerCase()
    if (['.md', '.markdown'].includes(extension) && info.size > MAX_TEXT_SIZE)
      throw new Error('Markdown 文件最大支持 5 MB。')
    const bytes = readFileSync(path)
    if (bytes.byteLength > MAX_FILE_SIZE) throw new Error('文件超过 100 MB。')
    return this.insert(nameOf(basename(path)), extension, bytes, categoryId, '首次导入')
  }

  createMarkdown(name: string, categoryId: string | null): DocumentRecord {
    const safe = nameOf(name)
    const filename = /\.(md|markdown)$/i.test(safe) ? safe : `${safe}.md`
    return this.insert(
      filename,
      extname(filename).toLowerCase(),
      Buffer.from(''),
      categoryId,
      '新建文档',
    )
  }

  createOffice(
    name: string,
    extension: '.docx' | '.xlsx',
    categoryId: string | null,
  ): DocumentRecord {
    if (extension !== '.docx' && extension !== '.xlsx') throw new Error('请选择 Word 或 Excel。')
    const safe = nameOf(name)
    const filename = safe.toLowerCase().endsWith(extension) ? safe : `${safe}${extension}`
    return this.insert(
      nameOf(filename),
      extension,
      newOfficeFile(extension),
      categoryId,
      '新建文档',
    )
  }

  readDocument(id: string): DocumentContent {
    const row = this.row(id)
    const text = ['.md', '.markdown'].includes(String(row.extension))
      ? decodeText(readFileSync(this.objectPath(String(row.blob_hash))))
      : fileKind(String(row.extension)) === 'text' && Number(row.size) <= MAX_TEXT_SIZE
        ? plainTextPreview(readFileSync(this.objectPath(String(row.blob_hash))))
        : null
    return { document: this.record(row), text }
  }

  readFilePreview(id: string): Uint8Array {
    const row = this.row(id)
    if (!['pdf', 'image'].includes(fileKind(String(row.extension))))
      throw new Error('此文件不使用图像预览。')
    if (Number(row.size) > 30 * 1024 ** 2)
      throw new Error('文件已保存。超过 30 MB 的 PDF / 图片请导出后使用本地软件阅读。')
    const bytes = readFileSync(this.objectPath(String(row.blob_hash)))
    return new Uint8Array(bytes)
  }

  openDocument(id: string): DocumentContent {
    const content = this.readDocument(id)
    if (content.document.deletedAt) throw new Error('请先从回收站恢复文档。')
    this.db
      .prepare('UPDATE documents SET opened_at = ? WHERE id = ?')
      .run(new Date().toISOString(), id)
    return content
  }

  saveMarkdown(id: string, text: string, revision: number): DocumentRecord {
    const row = this.row(id)
    if (row.deleted_at) throw new Error('请先从回收站恢复文档。')
    if (!['.md', '.markdown'].includes(String(row.extension)))
      throw new Error('Office 文档请使用对应的内容编辑保存。')
    if (!Number.isInteger(revision) || row.revision !== revision)
      throw new Error('文件已发生变化，请重新打开，避免覆盖较新的内容。')
    if (typeof text !== 'string' || Buffer.byteLength(text, 'utf8') > MAX_TEXT_SIZE)
      throw new Error('Markdown 内容最大支持 5 MB。')
    const bytes = Buffer.from(text, 'utf8')
    const hash = this.putObject(bytes)
    if (hash === row.blob_hash) {
      this.discardDraft(id)
      return this.record(row)
    }
    const now = new Date().toISOString()
    this.transaction(() => {
      this.db
        .prepare('INSERT INTO versions VALUES (?, ?, ?, ?, ?)')
        .run(randomUUID(), id, hash, now, '手动保存')
      this.db
        .prepare(
          'UPDATE documents SET blob_hash = ?, search_text = ?, size = ?, updated_at = ?, revision = revision + 1 WHERE id = ?',
        )
        .run(hash, text, bytes.byteLength, now, id)
      this.discardDraft(id)
    })
    return this.record(this.row(id))
  }

  updateDocument(
    id: string,
    patch: {
      name?: string
      categoryId?: string | null
      favorite?: boolean
      tags?: string[]
      notes?: string
    },
  ): void {
    const row = this.row(id)
    if (row.deleted_at) throw new Error('请先恢复文档。')
    if (!patch || typeof patch !== 'object') throw new Error('无效的修改。')
    const name = patch.name === undefined ? String(row.name) : nameOf(patch.name)
    if (extname(name).toLowerCase() !== row.extension)
      throw new Error('重命名时请保留原文件扩展名。')
    const category =
      patch.categoryId === undefined ? row.category_id : this.category(patch.categoryId)
    if (patch.favorite !== undefined && typeof patch.favorite !== 'boolean')
      throw new Error('无效的收藏状态。')
    const favorite = patch.favorite === undefined ? row.favorite : Number(patch.favorite)
    if (
      patch.tags !== undefined &&
      (!Array.isArray(patch.tags) ||
        patch.tags.length > 30 ||
        patch.tags.some((tag) => typeof tag !== 'string' || !tag.trim() || tag.length > 40))
    )
      throw new Error('最多添加 30 个标签，每个标签 1–40 个字符。')
    if (patch.notes !== undefined && (typeof patch.notes !== 'string' || patch.notes.length > 5000))
      throw new Error('备注最多 5000 个字符。')
    const tags =
      patch.tags === undefined
        ? row.tags
        : JSON.stringify([...new Set(patch.tags.map((tag) => tag.trim()))])
    const notes = patch.notes === undefined ? row.notes : patch.notes
    this.db
      .prepare(
        'UPDATE documents SET name = ?, category_id = ?, favorite = ?, tags = ?, notes = ? WHERE id = ?',
      )
      .run(name, category, favorite, tags, notes, id)
  }

  trashDocument(id: string): void {
    this.row(id)
    this.db
      .prepare('UPDATE documents SET deleted_at = COALESCE(deleted_at, ?) WHERE id = ?')
      .run(new Date().toISOString(), id)
  }

  restoreDocument(id: string): void {
    this.row(id)
    this.db.prepare('UPDATE documents SET deleted_at = NULL WHERE id = ?').run(id)
  }

  /** Only trash can be purged. Commit metadata first, then remove exclusively owned objects. */
  purgeDocuments(ids: string[]): { deleted: number; pendingCleanup: number } {
    if (!Array.isArray(ids) || !ids.length || ids.length > 10000)
      throw new Error('请选择 1–10000 份回收站文件。')
    const unique = new Set(ids.map(idOf))
    const paths = this.transaction(() => {
      for (const id of unique)
        if (!this.row(id).deleted_at) throw new Error('只能彻底删除回收站中的文件。')
      const candidates = new Set<string>(),
        retained = new Set<string>()
      const remember = (id: string, hash: string) =>
        (unique.has(id) ? candidates : retained).add(hash)
      for (const row of this.db.prepare('SELECT id, blob_hash FROM documents').all())
        remember(String(row.id), String(row.blob_hash))
      for (const row of this.db.prepare('SELECT document_id, blob_hash FROM versions').all())
        remember(String(row.document_id), String(row.blob_hash))
      for (const row of this.db.prepare('SELECT document_id, data FROM document_state').all()) {
        const id = String(row.document_id)
        for (const attachment of validateState(JSON.parse(String(row.data)), id).attachments)
          remember(id, attachment.hash)
      }
      // Validate paths before committing. Shared contents, history and attachments remain intact.
      const orphanPaths = [...candidates]
        .filter((hash) => !retained.has(hash))
        .map((hash) => this.objectPath(hash))
      const versions = this.db.prepare('DELETE FROM versions WHERE document_id = ?')
      const documents = this.db.prepare('DELETE FROM documents WHERE id = ?')
      for (const id of unique) {
        versions.run(id)
        documents.run(id) // office_cache and document_state are removed by foreign keys.
      }
      return orphanPaths
    })
    let pendingCleanup = 0
    for (const path of paths) {
      try {
        unlinkSync(path)
      } catch (error) {
        // A locked object must not roll back metadata that now refers to already removed bytes.
        if ((error as NodeJS.ErrnoException).code !== 'ENOENT') pendingCleanup++
      }
    }
    return { deleted: unique.size, pendingCleanup }
  }

  exportDocument(id: string, destination: string): void {
    const row = this.row(id)
    const target = resolve(destination)
    const root = resolve(this.root)
    if (
      target.toLowerCase() === root.toLowerCase() ||
      target
        .toLowerCase()
        .startsWith(`${root.toLowerCase()}${process.platform === 'win32' ? '\\' : '/'}`)
    ) {
      throw new Error('请选择文档库外的导出位置。')
    }
    if (existsSync(target)) throw new Error('该位置已有同名文件，请换个名称导出。')
    this.exportAttachments(id, dirname(target))
    // Never overwrite an existing file, including the original import source.
    copyFileSync(this.objectPath(String(row.blob_hash)), target, constants.COPYFILE_EXCL)
  }

  exportAttachments(id: string, directory: string): void {
    const images = this.documentState(id).attachments
    if (!images.length) return
    const folder = join(directory, 'attachments')
    mkdirSync(folder, { recursive: true })
    const info = lstatSync(folder)
    if (!info.isDirectory() || info.isSymbolicLink())
      throw new Error('附件导出位置被占用，请选择其他文件夹。')
    for (const a of images) {
      const path = join(folder, `${a.id}${a.extension}`)
      if (existsSync(path)) {
        const info = lstatSync(path)
        if (
          info.isSymbolicLink() ||
          !info.isFile() ||
          createHash('sha256').update(readFileSync(path)).digest('hex') !== a.hash
        )
          throw new Error('附件导出位置存在不同内容，未覆盖已有文件。')
      } else copyFileSync(this.objectPath(a.hash), path, constants.COPYFILE_EXCL)
    }
  }

  versions(id: string): VersionRecord[] {
    this.row(id)
    return this.db
      .prepare(
        'SELECT id, created_at AS createdAt, reason FROM versions WHERE document_id = ? ORDER BY created_at DESC, rowid DESC',
      )
      .all(id) as unknown as VersionRecord[]
  }

  versionSource(id: string, versionId: string) {
    const row = this.row(id)
    const version = this.db
      .prepare(
        'SELECT id, blob_hash AS hash, created_at AS createdAt, reason FROM versions WHERE id = ? AND document_id = ?',
      )
      .get(idOf(versionId), id) as (VersionRecord & { hash: string }) | undefined
    if (!version) throw new Error('历史版本不存在。')
    const extension = String(row.extension)
    const path = this.objectPath(version.hash)
    return {
      version: { id: version.id, createdAt: version.createdAt, reason: version.reason },
      hash: version.hash,
      extension,
      path,
      text: ['.md', '.markdown'].includes(extension)
        ? decodeText(readFileSync(path))
        : fileKind(extension) === 'text'
          ? plainTextPreview(readFileSync(path))
          : null,
    }
  }

  restoreVersion(id: string, versionId: string, revision: number): DocumentContent {
    const row = this.row(id)
    if (row.deleted_at) throw new Error('请先恢复文档。')
    if (row.revision !== revision) throw new Error('文件已变化，请重新打开后恢复版本。')
    const version = this.db
      .prepare('SELECT blob_hash FROM versions WHERE id = ? AND document_id = ?')
      .get(idOf(versionId), id) as { blob_hash: string } | undefined
    if (!version) throw new Error('历史版本不存在。')
    const bytes = readFileSync(this.objectPath(version.blob_hash))
    const text = ['.md', '.markdown'].includes(String(row.extension))
      ? decodeText(bytes)
      : fileKind(String(row.extension)) === 'text'
        ? plainTextIndex(String(row.extension), bytes)
        : (this.officeCache(id, version.blob_hash)
            ?.data?.blocks.map((block) => block.text)
            .join('\n') ?? '')
    const now = new Date().toISOString()
    this.transaction(() => {
      this.db
        .prepare('INSERT INTO versions VALUES (?, ?, ?, ?, ?)')
        .run(randomUUID(), id, version.blob_hash, now, '从历史恢复')
      this.db
        .prepare(
          'UPDATE documents SET blob_hash = ?, search_text = ?, size = ?, updated_at = ?, revision = revision + 1 WHERE id = ?',
        )
        .run(version.blob_hash, text, bytes.byteLength, now, id)
    })
    return this.readDocument(id)
  }
}
