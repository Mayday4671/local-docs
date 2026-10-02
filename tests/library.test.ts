import { afterEach, beforeEach, describe, expect, test } from 'vitest'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { DatabaseSync } from 'node:sqlite'
import { Library } from '../src/main/library'

let directory: string
let library: Library
beforeEach(() => {
  directory = mkdtempSync(join(tmpdir(), 'local-docs-test-'))
  library = new Library(join(directory, 'library'))
})
afterEach(() => {
  library.close()
  const target = realpathSync(directory)
  const prefix = join(realpathSync(tmpdir()), 'local-docs-test-')
  if (!target.startsWith(prefix))
    throw new Error('Refusing to remove a directory outside test temp scope')
  rmSync(target, { recursive: true, force: true })
})

describe('文档库持久化与文件保护', () => {
  test('PDF 分块读取覆盖 100 MB，校验身份、范围与截断，且不改动文档和源文件', () => {
    const bytes = Buffer.alloc(100 * 1024 ** 2, 65)
    bytes.write('%PDF-1.4\n')
    bytes.write('%%EOF', bytes.length - 5)
    const source = join(directory, '大文件.pdf')
    writeFileSync(source, bytes)
    const doc = library.importFile(source, null)
    const info = library.pdfInfo(doc.id)
    const before = library.snapshot(),
      versions = library.versions(doc.id)
    expect(info.size).toBe(bytes.length)
    expect(Buffer.from(library.readPdfRange(doc.id, info.hash, 0, 1024 ** 2))).toEqual(
      bytes.subarray(0, 1024 ** 2),
    )
    expect(
      Buffer.from(
        library.readPdfRange(doc.id, info.hash, bytes.length - 5, bytes.length),
      ).toString(),
    ).toBe('%%EOF')
    for (const [begin, end] of [
      [-1, 3],
      [0, 0],
      [1, 0],
      [0.5, 1],
      [0, Infinity],
      [0, 1024 ** 2 + 1],
      [bytes.length, bytes.length + 1],
    ])
      expect(() => library.readPdfRange(doc.id, info.hash, begin, end)).toThrow('读取范围')
    expect(() => library.readPdfRange(doc.id, 'different', 0, 1)).toThrow('文档已更改')
    expect(() => library.pdfInfo('../outside.pdf')).toThrow('编号')
    expect(() => library.pdfInfo(library.createMarkdown('其他类型', null).id)).toThrow('不是 PDF')
    expect(library.readDocument(doc.id).document).toEqual(before.documents[0])
    expect(library.versions(doc.id)).toEqual(versions)
    expect(readFileSync(source).equals(bytes)).toBe(true)
    writeFileSync(library.documentPath(doc.id), 'truncated')
    expect(() => library.readPdfRange(doc.id, info.hash, 0, 1)).toThrow('不完整')
  })
  test('历史读取不改变当前文件、最近使用、草稿或版本列表，拒绝跨文件版本', () => {
    const doc = library.createMarkdown('版本对比', null)
    const original = library.versions(doc.id)[0]
    const saved = library.saveMarkdown(doc.id, '保存后的正文', doc.revision)
    library.saveDraft(doc.id, '尚未保存的正文', saved.revision, 'markdown')
    const snapshot = library.snapshot(),
      versions = library.versions(doc.id),
      draft = library.documentState(doc.id).draft
    const old = library.versionSource(doc.id, original.id)
    const current = library.versionSource(doc.id, versions[0].id)
    expect(old.text).toBe('')
    expect(current.text).toBe('保存后的正文')
    expect(old.hash).not.toBe(current.hash)
    expect(library.snapshot()).toEqual(snapshot)
    expect(library.versions(doc.id)).toEqual(versions)
    expect(library.documentState(doc.id).draft).toEqual(draft)
    const other = library.createMarkdown('其他文档', null)
    expect(() => library.versionSource(other.id, original.id)).toThrow('历史版本不存在')
    expect(() => library.versionSource(doc.id, '../objects/file')).toThrow('编号')
    expect(() => library.versionSource(doc.id, '00000000-0000-0000-0000-000000000000')).toThrow(
      '历史版本不存在',
    )
  })
  test('主题即时保存、重启恢复且拒绝无效主题', () => {
    expect(library.getTheme()).toBe('system')
    library.setTheme('dark')
    expect(() => library.setTheme('invalid' as 'dark')).toThrow('主题')
    expect(library.getTheme()).toBe('dark')
    library.close()
    library = new Library(join(directory, 'library'))
    expect(library.getTheme()).toBe('dark')
    library.setTheme('light')
    expect(library.getTheme()).toBe('light')
    library.setTheme('system')
    expect(library.getTheme()).toBe('system')
  })
  test('多级分类允许不同父目录同名，标签备注可检索且重启保留', () => {
    const work = library.createCategory('工作')
    const life = library.createCategory('生活')
    const first = library.createCategory('资料', work.id)
    expect(library.createCategory('资料', life.id).parentId).toBe(life.id)
    expect(() => library.createCategory('资料', work.id)).toThrow('同名')
    expect(() => library.createCategory('无效父级', 'bad')).toThrow()
    const doc = library.createMarkdown('设计说明', first.id)
    library.updateDocument(doc.id, { tags: ['待整理', '端口', '端口'], notes: '确认甲方要求' })
    expect(library.searchResults('端口')).toMatchObject([{ id: doc.id, source: 'metadata' }])
    expect(library.search('甲')).toEqual([doc.id])
    expect(() => library.updateDocument(doc.id, { tags: [null as unknown as string] })).toThrow(
      '标签',
    )
    library.close()
    library = new Library(join(directory, 'library'))
    expect(
      library.snapshot().categories.find((category) => category.id === first.id)?.parentId,
    ).toBe(work.id)
    expect(library.readDocument(doc.id).document.tags).toEqual(['待整理', '端口'])
    expect(library.readDocument(doc.id).document.revision).toBe(1)
    expect(library.versions(doc.id)).toHaveLength(1)
  })

  test('升级 0.1 旧库保留分类关联、内容和历史，新增字段默认安全', () => {
    const category = library.createCategory('已有资料')
    const doc = library.createMarkdown('旧版笔记', category.id)
    library.saveMarkdown(doc.id, '需要保留的原文', doc.revision)
    library.close()
    const old = new DatabaseSync(join(directory, 'library', 'library.sqlite'))
    old.exec(`DROP TABLE settings; DROP INDEX category_siblings; ALTER TABLE categories DROP COLUMN parent_id;
      ALTER TABLE documents DROP COLUMN tags; ALTER TABLE documents DROP COLUMN notes;
      PRAGMA user_version = 2;`)
    old.close()
    library = new Library(join(directory, 'library'))
    expect(library.readDocument(doc.id)).toMatchObject({
      text: '需要保留的原文',
      document: { categoryId: category.id, tags: [], notes: '' },
    })
    expect(library.snapshot().categories).toContainEqual({ ...category, parentId: null })
    expect(library.versions(doc.id)).toHaveLength(2)
    const verify = new DatabaseSync(join(directory, 'library', 'library.sqlite'))
    expect(verify.prepare('PRAGMA foreign_key_check').all()).toEqual([])
    expect(verify.prepare('PRAGMA user_version').get()?.user_version).toBe(5)
    verify.close()
  })
  test('导入、编辑、导出不改变源文件，关闭后内容和分类仍在', () => {
    const source = join(directory, '笔记.md')
    writeFileSync(source, '# 原始文档\n网络配置')
    const category = library.createCategory('工作资料')
    const imported = library.importFile(source, category.id)
    library.saveMarkdown(imported.id, '# 新的内容', imported.revision)
    expect(readFileSync(source, 'utf8')).toBe('# 原始文档\n网络配置')
    library.close()
    library = new Library(join(directory, 'library'))
    expect(library.readDocument(imported.id).text).toBe('# 新的内容')
    expect(library.snapshot().documents[0].categoryId).toBe(category.id)
    const exported = join(directory, '导出.md')
    library.exportDocument(imported.id, exported)
    expect(readFileSync(exported, 'utf8')).toBe('# 新的内容')
  })

  test('Office 文件按字节保管和导出，不假装能够编辑', () => {
    const source = join(directory, '资料.docx')
    const bytes = Buffer.from([80, 75, 3, 4, 0, 255, 17, 38])
    writeFileSync(source, bytes)
    const doc = library.importFile(source, null)
    expect(library.readDocument(doc.id).text).toBeNull()
    expect(() => library.saveMarkdown(doc.id, '不应写入', 1)).toThrow('Office')
    const target = join(directory, '副本.docx')
    library.exportDocument(doc.id, target)
    expect(readFileSync(target)).toEqual(bytes)
  })

  test('短中文和通配符按字面搜索，编辑后立即更新搜索内容', () => {
    const first = library.createMarkdown('报表_100%', null)
    library.saveMarkdown(first.id, '端口配置与联网说明', 1)
    const second = library.createMarkdown('普通笔记', null)
    expect(library.search('端')).toEqual([first.id])
    expect(library.search('端口')).toEqual([first.id])
    expect(library.search('%')).toEqual([first.id])
    expect(library.search('_')).toEqual([first.id])
    library.saveMarkdown(first.id, '替换正文', 2)
    expect(library.search('端口')).toEqual([])
    expect(library.search('普通')).toEqual([second.id])
  })

  test('预览不计入最近使用，真正打开才会记录', () => {
    const doc = library.createMarkdown('笔记', null)
    library.readDocument(doc.id)
    expect(library.snapshot().documents[0].openedAt).toBeNull()
    library.openDocument(doc.id)
    expect(library.snapshot().documents[0].openedAt).not.toBeNull()
  })

  test('回收站保留内容和分类，并阻止继续编辑已删除的文档', () => {
    const category = library.createCategory('项目')
    const doc = library.createMarkdown('记录', category.id)
    library.saveMarkdown(doc.id, '可恢复的内容', 1)
    library.trashDocument(doc.id)
    expect(library.readDocument(doc.id).text).toBe('可恢复的内容')
    expect(() => library.saveMarkdown(doc.id, '不能写入', 2)).toThrow('恢复')
    expect(() => library.openDocument(doc.id)).toThrow('恢复')
    library.restoreDocument(doc.id)
    expect(library.readDocument(doc.id).document).toMatchObject({
      categoryId: category.id,
      deletedAt: null,
    })
  })

  test('历史恢复保留恢复前内容，并拒绝覆盖新版本', () => {
    const doc = library.createMarkdown('记录', null)
    library.saveMarkdown(doc.id, '第一版', 1)
    const firstVersion = library.versions(doc.id)[0]
    library.saveMarkdown(doc.id, '第二版', 2)
    expect(() => library.saveMarkdown(doc.id, '过期窗口', 2)).toThrow('文件已发生变化')
    expect(() => library.restoreVersion(doc.id, firstVersion.id, 2)).toThrow('文件已变化')
    const restored = library.restoreVersion(doc.id, firstVersion.id, 3)
    expect(restored.text).toBe('第一版')
    expect(restored.document.revision).toBe(4)
    expect(library.versions(doc.id)).toHaveLength(4)
    const secondVersion = library.versions(doc.id)[1]
    expect(library.restoreVersion(doc.id, secondVersion.id, 4).text).toBe('第二版')
  })

  test('重复保存相同内容不消耗版本', () => {
    const doc = library.createMarkdown('记录', null)
    const saved = library.saveMarkdown(doc.id, '内容', 1)
    expect(library.saveMarkdown(doc.id, '内容', saved.revision).revision).toBe(saved.revision)
    expect(library.versions(doc.id)).toHaveLength(2)
  })

  test('导出不能覆盖现有文件，也不能写入内部文档库', () => {
    const doc = library.createMarkdown('记录', null)
    const existing = join(directory, '原文件.md')
    writeFileSync(existing, '不能覆盖')
    expect(() => library.exportDocument(doc.id, existing)).toThrow()
    expect(readFileSync(existing, 'utf8')).toBe('不能覆盖')
    expect(() => library.exportDocument(doc.id, resolve(library.root, 'library.sqlite'))).toThrow(
      '文档库外',
    )
  })

  test('拒绝非法名称、无效分类、类型变更和跨文档历史恢复', () => {
    expect(() => library.createMarkdown('../外部路径', null)).toThrow('名称')
    expect(() => library.createMarkdown('CON', null)).toThrow('名称')
    expect(() => library.createMarkdown('笔记', '123')).toThrow('编号')
    const category = library.createCategory('资料')
    expect(() => library.createCategory('资料')).toThrow('同名')
    const a = library.createMarkdown('A', category.id)
    const b = library.createMarkdown('B', null)
    expect(() => library.updateDocument(a.id, { name: 'A.exe' })).toThrow('扩展名')
    expect(() => library.restoreVersion(a.id, library.versions(b.id)[0].id, 1)).toThrow('不存在')
  })

  test('错误 Markdown 编码拒绝导入，其他文件只保存不执行', () => {
    const invalid = join(directory, '编码.md')
    writeFileSync(invalid, Buffer.from([0xff, 0xfe, 0x81]))
    expect(() => library.importFile(invalid, null)).toThrow('UTF-8')
    const unsupported = join(directory, '程序.exe')
    writeFileSync(unsupported, 'test')
    const stored = library.importFile(unsupported, null)
    expect(library.readDocument(stored.id).text).toBeNull()
    expect(readFileSync(library.documentPath(stored.id), 'utf8')).toBe('test')
    expect(() => library.saveDraft(stored.id, '[]', 1, 'office')).toThrow('草稿类型')
    expect(library.snapshot().documents).toHaveLength(1)
  })
})
