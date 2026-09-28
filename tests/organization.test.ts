import { afterEach, beforeEach, expect, test, vi } from 'vitest'
import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  realpathSync,
  renameSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Library } from '../src/main/library'
import { FolderImportService } from '../src/main/folder-import-service'
import { TransferService } from '../src/main/transfer-service'
import { wordFixture, sheetFixture } from './fixtures/office'

let root: string, source: string, library: Library, folders: FolderImportService
beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'local-docs-organization-test-'))
  source = join(root, '资料包')
  mkdirSync(source)
  library = new Library(join(root, 'data', 'library'))
  folders = new FolderImportService(library, join(root, 'data'))
})
afterEach(() => {
  vi.restoreAllMocks()
  library.close()
  const target = realpathSync(root)
  if (!target.startsWith(join(realpathSync(tmpdir()), 'local-docs-organization-test-')))
    throw new Error('Unsafe test cleanup')
  rmSync(target, { recursive: true, force: true })
})
function file(name: string, text: string | Buffer = '# 合成测试资料') {
  const path = join(source, name)
  writeFileSync(path, text)
  return path
}

test('分类改名与移动拒绝循环、重名；不改变文档内容、元数据和历史', () => {
  const a = library.createCategory('工作'),
    b = library.createCategory('生活')
  const child = library.createCategory('设计', a.id)
  const other = library.createCategory('设计', b.id)
  const doc = library.createMarkdown('说明', child.id)
  library.saveMarkdown(doc.id, '已保存版本', 1)
  library.updateDocument(doc.id, { tags: ['重要'], notes: '保留备注', favorite: true })
  const before = library.readDocument(doc.id),
    versions = library.versions(doc.id)
  expect(() => library.updateCategory(a.id, { name: '工作', parentId: child.id })).toThrow('子分类')
  expect(() => library.updateCategory(child.id, { name: '设计', parentId: b.id })).toThrow('同名')
  expect(() => library.updateCategory(child.id, { name: '../路径', parentId: null })).toThrow(
    '名称',
  )
  library.updateCategory(child.id, { name: '归档设计', parentId: b.id })
  expect(library.snapshot().categories).toContainEqual({
    ...child,
    name: '归档设计',
    parentId: b.id,
  })
  expect(library.snapshot().categories).toContainEqual(other)
  expect(library.readDocument(doc.id)).toEqual(before)
  expect(library.versions(doc.id)).toEqual(versions)
  library.backupManifest()
})

test('删除整棵分类树时转移文件并保留回收站状态、历史和其他分类', () => {
  const parent = library.createCategory('项目'),
    child = library.createCategory('资料', parent.id)
  const target = library.createCategory('归档')
  const doc = library.createMarkdown('说明', child.id),
    trash = library.createMarkdown('旧稿', parent.id)
  library.saveMarkdown(doc.id, '正文', 1)
  library.trashDocument(trash.id)
  const before = library.snapshot(),
    versions = library.versions(doc.id)
  expect(() => library.deleteCategory(parent.id, child.id)).toThrow('将要删除')
  expect(library.snapshot()).toEqual(before)
  library.deleteCategory(parent.id, target.id)
  expect(library.snapshot().categories).toEqual([target])
  expect(library.readDocument(doc.id).document.categoryId).toBe(target.id)
  expect(library.readDocument(doc.id).text).toBe('正文')
  expect(library.readDocument(trash.id).document).toMatchObject({
    categoryId: target.id,
    deletedAt: before.documents.find((d) => d.id === trash.id)!.deletedAt,
  })
  expect(library.versions(doc.id)).toEqual(versions)
  library.deleteCategory(target.id, null)
  expect(library.snapshot().documents.every((d) => d.categoryId === null)).toBe(true)
  expect(() => library.deleteCategory(null as unknown as string, null)).toThrow()
  library.close()
  library = new Library(join(root, 'data', 'library'))
  expect(library.readDocument(doc.id).text).toBe('正文')
  expect(library.versions(doc.id)).toEqual(versions)
})

test('批量移动先校验所有文件，失败时不出现部分移动', () => {
  const category = library.createCategory('目标')
  const a = library.createMarkdown('甲', null),
    b = library.createMarkdown('乙', null)
  const before = library.snapshot()
  expect(() => library.moveDocuments([a.id, 'bad'], category.id)).toThrow()
  expect(library.snapshot()).toEqual(before)
  library.trashDocument(b.id)
  expect(() => library.moveDocuments([a.id, b.id], category.id)).toThrow('恢复')
  expect(library.readDocument(a.id).document.categoryId).toBeNull()
  library.restoreDocument(b.id)
  library.moveDocuments([a.id, b.id, a.id], category.id)
  expect(library.snapshot().documents.every((d) => d.categoryId === category.id)).toBe(true)
  expect(library.versions(a.id)).toHaveLength(1)
})

test('分类创建和移动遵守备份可恢复的层级上限', () => {
  let parent: string | null = null
  for (let i = 0; i < 128; i++) parent = library.createCategory(`层${i}`, parent).id
  expect(() => library.createCategory('过深', parent)).toThrow('128')
  const a = library.createCategory('新分类')
  expect(() => library.updateCategory(a.id, { name: a.name, parentId: parent })).toThrow('128')
  expect(library.backupManifest().categories).toHaveLength(129)
})

test('文件夹预览只读；文档、文本与空目录保留，原文件字节不变，备份可恢复', async () => {
  mkdirSync(join(source, '子分类', '空目录'), { recursive: true })
  const word = Buffer.from(wordFixture()),
    sheet = Buffer.from(sheetFixture())
  const wordPath = file('说明.docx', word),
    sheetPath = file('数据.xlsx', sheet)
  file('子分类/笔记.markdown', '| 项目 | 状态 |\n| --- | --- |\n| 合成资料 | 完成 |')
  file('说明.txt', '文本也能导入')
  const before = library.snapshot()
  const preview = await folders.preview(source)
  expect(preview).toMatchObject({ files: 4, folders: 3, skipped: 0 })
  expect(library.snapshot()).toEqual(before)
  const parent = library.createCategory('目标位置')
  const result = await folders.import(preview.token, parent.id, 'skip')
  expect(result).toMatchObject({ skipped: 0, failures: [], cancelled: false })
  expect(result.imported).toHaveLength(4)
  const snap = library.snapshot()
  expect(snap.categories.find((c) => c.name === '资料包')!.parentId).toBe(parent.id)
  const child = snap.categories.find((c) => c.name === '子分类')!
  expect(snap.categories.find((c) => c.name === '空目录')!.parentId).toBe(child.id)
  for (const doc of snap.documents) {
    expect(library.versions(doc.id)).toHaveLength(1)
    if (doc.extension === '.docx') expect(readFileSync(library.documentPath(doc.id))).toEqual(word)
    if (doc.extension === '.xlsx') expect(readFileSync(library.documentPath(doc.id))).toEqual(sheet)
  }
  expect(readFileSync(wordPath)).toEqual(word)
  expect(readFileSync(sheetPath)).toEqual(sheet)
  const transfer = new TransferService(library, join(root, 'data'))
  const backup = join(root, '备份.localdocs-backup')
  await transfer.backup(backup)
  const stored = await transfer.preview(backup)
  library.deleteCategory(parent.id, null)
  await transfer.restore(stored.token)
  expect(library.snapshot()).toEqual(snap)
})

test('重复内容可跳过，冲突文件加编号；重试不会重复生成已编号的副本', async () => {
  const sourceFile = file('笔记.md', '第一版')
  const preview = await folders.preview(source)
  const first = await folders.import(preview.token, null, 'skip')
  let p = await folders.preview(source)
  expect(await folders.import(p.token, null, 'skip')).toMatchObject({ imported: [], skipped: 1 })
  writeFileSync(sourceFile, '不同内容的第二版')
  p = await folders.preview(source)
  expect((await folders.import(p.token, null, 'skip')).imported).toHaveLength(1)
  expect(
    library
      .snapshot()
      .documents.map((d) => d.name)
      .sort(),
  ).toEqual(['笔记.md', '笔记（2）.md'].sort())
  p = await folders.preview(source)
  expect(await folders.import(p.token, null, 'skip')).toMatchObject({ imported: [], skipped: 1 })
  p = await folders.preview(source)
  expect((await folders.import(p.token, null, 'keep')).imported).toHaveLength(1)
  expect(library.readDocument(first.imported[0]).text).toBe('第一版')
  expect(library.snapshot().categories).toHaveLength(1)
  const longName = `${'长'.repeat(177)}.md`
  library.importBytes(longName, Buffer.from('a'), first.rootCategoryId, 'skip')
  library.importBytes(longName, Buffer.from('b'), first.rootCategoryId, 'skip')
  expect(library.importBytes(longName, Buffer.from('b'), first.rootCategoryId, 'skip')).toBeNull()
})

test('错误编码、消失或预览后修改的文件单独失败，其他文件继续导入', async () => {
  file('编码.md', Buffer.from([0xff, 0xfe]))
  const changed = file('变化.md'),
    missing = file('消失.md')
  file('正常.md', '正常内容')
  const p = await folders.preview(source)
  writeFileSync(changed, '已变化的新内容')
  renameSync(missing, join(root, '移走.md'))
  const r = await folders.import(p.token, null, 'skip')
  expect(r.imported).toHaveLength(1)
  expect(r.failures).toHaveLength(3)
  expect(r.failures.map((f) => f.reason).join(' ')).toMatch(/UTF-8/)
  expect(library.snapshot().documents[0].name).toBe('正常.md')
})

test('不导入应用存储目录或目录链接，扫描深度有界，拒绝过期预览', async () => {
  file('普通.md')
  symlinkSync(join(root, 'data'), join(source, '链接'), 'junction')
  let deep = source
  for (let i = 0; i < 33; i++) {
    deep = join(deep, '下级')
    mkdirSync(deep)
  }
  const p = await folders.preview(source)
  expect(p.warnings.some((w) => w.reason.includes('链接'))).toBe(true)
  expect(p.warnings.some((w) => w.reason.includes('32'))).toBe(true)
  expect(p.files).toBe(1)
  folders.discardPreview()
  await expect(folders.import(p.token, null, 'skip')).rejects.toThrow('失效')
  await expect(folders.preview(join(root, 'data'))).rejects.toThrow('自身')
  const ancestor = await folders.preview(root)
  expect(ancestor.warnings.some((w) => w.name === 'data')).toBe(true)
})

test('预览后目录被替换为链接时，不读取链接目标中的文件', async () => {
  const sub = join(source, '子目录')
  mkdirSync(sub)
  file('子目录/笔记.md', '批准的内容')
  const p = await folders.preview(source)
  const outside = join(root, '其他目录')
  mkdirSync(outside)
  writeFileSync(join(outside, '笔记.md'), '不能导入的内容')
  renameSync(sub, join(source, '移走的目录'))
  symlinkSync(outside, sub, 'junction')
  const r = await folders.import(p.token, null, 'skip')
  expect(r.imported).toHaveLength(0)
  expect(r.failures.length).toBeGreaterThan(0)
  expect(library.snapshot().documents).toEqual([])
})

test('取消扫描不更改库；取消导入保留已完成文件，重试跳过它们', async () => {
  for (let i = 0; i < 20; i++) file(`${i}.md`, `合成文件 ${i}`)
  const pending = folders.preview(source)
  folders.cancel()
  await expect(pending).rejects.toThrow('取消')
  expect(library.snapshot().categories).toEqual([])
  const p = await folders.preview(source)
  const original = library.importBytes.bind(library)
  const spy = vi.spyOn(library, 'importBytes').mockImplementation((...args) => {
    const doc = original(...args)
    folders.cancel()
    return doc
  })
  const r = await folders.import(p.token, null, 'skip')
  expect(r.cancelled).toBe(true)
  expect(r.imported).toHaveLength(1)
  expect(folders.getStatus()).toBeNull()
  spy.mockRestore()
  const retry = await folders.preview(source)
  const done = await folders.import(retry.token, null, 'skip')
  expect(done.imported).toHaveLength(19)
  expect(done.skipped).toBe(1)
  expect(library.snapshot().documents).toHaveLength(20)
})
