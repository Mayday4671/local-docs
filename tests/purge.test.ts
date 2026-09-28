import { beforeEach, afterEach, expect, test } from 'vitest'
import {
  existsSync,
  mkdtempSync,
  mkdirSync,
  readFileSync,
  realpathSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { DatabaseSync } from 'node:sqlite'
import { Library } from '../src/main/library'
import { TransferService } from '../src/main/transfer-service'

let root: string, library: Library
beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'local-docs-purge-test-'))
  library = new Library(join(root, 'library'))
})
afterEach(() => {
  library.close()
  const path = realpathSync(root)
  if (!path.startsWith(join(realpathSync(tmpdir()), 'local-docs-purge-test-')))
    throw Error('Unsafe cleanup')
  rmSync(path, { recursive: true, force: true })
})
const png = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  'base64',
)

test('彻底删除当前文件、历史、草稿、标记及独占附件；源文件与旧备份保留，重启不恢复记录', async () => {
  const source = join(root, '原文件.md')
  writeFileSync(source, '删除测试正文')
  const doc = library.importFile(source, null)
  const original = library.documentPath(doc.id)
  const image = library.addImage(doc.id, png, 'image/png')
  const imagePath = library.attachmentPath(doc.id, image.id).path
  library.saveMarkdown(doc.id, '修改后的正文', 1)
  const current = library.documentPath(doc.id)
  library.saveDraft(doc.id, '恢复草稿', 2, 'markdown')
  library.saveAnnotation(
    doc.id,
    {
      quote: '正文',
      prefix: '',
      suffix: '',
      offset: 0,
      scope: 'markdown',
      color: 'yellow',
      note: '备注',
    },
    2,
  )
  const transfer = new TransferService(library, library.root),
    backup = join(root, '旧备份.localdocs-backup')
  await transfer.backup(backup)
  const backupBytes = readFileSync(backup)
  library.trashDocument(doc.id)
  expect(library.purgeDocuments([doc.id, doc.id])).toEqual({ deleted: 1, pendingCleanup: 0 })
  expect([original, current, imagePath].map(existsSync)).toEqual([false, false, false])
  expect(readFileSync(source, 'utf8')).toBe('删除测试正文')
  expect(readFileSync(backup)).toEqual(backupBytes)
  expect(library.backupManifest().documents).toEqual([])
  const db = new DatabaseSync(join(library.root, 'library.sqlite'))
  for (const table of ['versions', 'document_state', 'office_cache'])
    expect(db.prepare(`SELECT COUNT(*) AS n FROM ${table}`).get()?.n).toBe(0)
  db.close()
  library.close()
  library = new Library(join(root, 'library'))
  expect(library.snapshot().documents).toEqual([])
  expect(() => library.restoreDocument(doc.id)).toThrow('不存在')
})

test('共享正文、其他文件的历史、附件及回收站引用保留，最后一个引用删除后才清理', () => {
  const a = library.importBytes('甲.md', Buffer.from('共享正文'), null, 'keep')!
  const b = library.importBytes('乙.md', Buffer.from('共享正文'), null, 'keep')!
  const shared = library.documentPath(a.id)
  const history = library.versions(b.id)[0]
  library.saveMarkdown(b.id, '乙的新正文', 1)
  const imageA = library.addImage(a.id, png, 'image/png')
  const imageB = library.addImage(b.id, png, 'image/png')
  const imagePath = library.attachmentPath(a.id, imageA.id).path
  library.trashDocument(b.id)
  library.trashDocument(a.id)
  library.purgeDocuments([a.id])
  expect(readFileSync(shared, 'utf8')).toBe('共享正文')
  expect(readFileSync(library.attachmentPath(b.id, imageB.id).path)).toEqual(png)
  library.restoreDocument(b.id)
  library.restoreVersion(b.id, history.id, 2)
  expect(library.readDocument(b.id).text).toBe('共享正文')
  library.trashDocument(b.id)
  library.purgeDocuments([b.id])
  expect(existsSync(shared)).toBe(false)
  expect(existsSync(imagePath)).toBe(false)
})

test('拒绝普通文件、无效/缺失编号；混合批次不会部分删除', () => {
  const a = library.createMarkdown('甲', null),
    b = library.createMarkdown('乙', null)
  library.trashDocument(a.id)
  for (const ids of [
    [],
    [a.id, b.id],
    [a.id, '../escape'],
    [a.id, '00000000-0000-0000-0000-000000000000'],
  ]) {
    expect(() => library.purgeDocuments(ids)).toThrow()
    expect(library.snapshot().documents).toHaveLength(2)
    expect(existsSync(library.documentPath(a.id))).toBe(true)
  }
})

test('数据库删除失败时整体回滚，文件内容不被提前清除', () => {
  const a = library.createMarkdown('甲', null),
    b = library.createMarkdown('乙', null)
  library.trashDocument(a.id)
  library.trashDocument(b.id)
  const before = library.backupManifest()
  const db = new DatabaseSync(join(library.root, 'library.sqlite'))
  db.exec(
    `CREATE TRIGGER refuse_delete BEFORE DELETE ON documents WHEN OLD.id = '${b.id}' BEGIN SELECT RAISE(ABORT, 'test failure'); END`,
  )
  db.close()
  expect(() => library.purgeDocuments([a.id, b.id])).toThrow('test failure')
  expect(library.backupManifest().documents).toEqual(before.documents)
  expect(library.backupManifest().versions).toEqual(before.versions)
  library.checkIntegrity()
})

test('物理文件清理失败明确返回，不误报数据库回滚，也不递归删除异常路径', () => {
  const doc = library.createMarkdown('失败路径', null)
  const path = library.documentPath(doc.id)
  renameSync(path, join(root, '保留的内容'))
  mkdirSync(path)
  writeFileSync(join(path, '不应删除.txt'), '保留')
  library.trashDocument(doc.id)
  expect(library.purgeDocuments([doc.id])).toEqual({ deleted: 1, pendingCleanup: 1 })
  expect(library.snapshot().documents).toHaveLength(0)
  expect(readFileSync(join(path, '不应删除.txt'), 'utf8')).toBe('保留')
})
