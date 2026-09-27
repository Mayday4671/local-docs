import { beforeEach, afterEach, expect, test } from 'vitest'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { unzipSync, strFromU8 } from 'fflate'
import { Library } from '../src/main/library'
import { TransferService } from '../src/main/transfer-service'
import { validateManifest } from '../src/main/backup-format'
import { officeEditModel, applyOfficeChanges, newOfficeFile } from '../src/main/office-edit'
import { wordFixture, sheetFixture, officeZip } from './fixtures/office'
import { parseOffice } from '../src/main/office-parser'
import { DatabaseSync } from 'node:sqlite'
import type { AnnotationInput } from '../src/shared/types'
let root: string, library: Library
beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'local-docs-content-test-'))
  library = new Library(join(root, 'data', 'library'))
})
afterEach(() => {
  library.close()
  const path = realpathSync(root)
  if (!path.startsWith(join(realpathSync(tmpdir()), 'local-docs-content-test-')))
    throw new Error('Unsafe cleanup')
  rmSync(path, { recursive: true, force: true })
})
const mark: AnnotationInput = {
  quote: '重点内容',
  prefix: '',
  suffix: '',
  offset: 0,
  scope: 'markdown',
  color: 'yellow',
  note: '稍后核对',
}
const png = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  'base64',
)

test('收藏批处理原子校验，标记独立持久化，不修改正文和版本', () => {
  const a = library.createMarkdown('甲', null),
    b = library.createMarkdown('乙', null)
  library.saveMarkdown(a.id, '重点内容', 1)
  const content = library.readDocument(a.id),
    history = library.versions(a.id)
  expect(() => library.setFavorites([a.id, 'bad'], true)).toThrow()
  expect(library.readDocument(a.id).document.favorite).toBe(false)
  library.setFavorites([a.id, b.id], true)
  const annotation = library.saveAnnotation(a.id, mark, 2)
  expect(library.snapshot().annotationCounts?.[a.id]).toBe(1)
  expect(() => library.saveAnnotation(a.id, mark, 1)).toThrow('变化')
  expect(() => library.saveAnnotation(b.id, { ...mark, id: annotation.id }, 1)).toThrow('不存在')
  library.saveAnnotation(a.id, { ...annotation, note: '已核对', color: 'green' }, 2)
  library.close()
  library = new Library(join(root, 'data', 'library'))
  expect(library.annotations(a.id)[0]).toMatchObject({ note: '已核对', color: 'green' })
  expect(library.readDocument(a.id).text).toBe(content.text)
  expect(library.versions(a.id)).toEqual(history)
  library.trashDocument(a.id)
  expect(library.annotations()).toHaveLength(0)
  library.restoreDocument(a.id)
  expect(library.annotations()).toHaveLength(1)
  library.removeAnnotation(a.id, annotation.id)
  expect(library.annotations(a.id)).toEqual([])
})

test('恢复草稿不消耗版本，重启可恢复，正式保存清理草稿，拒绝过期覆盖', () => {
  const doc = library.createMarkdown('草稿', null)
  library.saveDraft(doc.id, '首次草稿', 1, 'markdown')
  library.saveDraft(doc.id, '最新草稿', 1, 'markdown')
  expect(library.versions(doc.id)).toHaveLength(1)
  expect(library.readDocument(doc.id).text).toBe('')
  library.close()
  library = new Library(join(root, 'data', 'library'))
  expect(library.documentState(doc.id).draft?.text).toBe('最新草稿')
  library.saveMarkdown(doc.id, '最新草稿', 1)
  expect(library.documentState(doc.id).draft).toBeNull()
  expect(library.versions(doc.id)).toHaveLength(2)
  expect(() => library.saveDraft(doc.id, '过期', 1, 'markdown')).toThrow('变化')
  expect(() => library.saveDraft(doc.id, 'x', 2, 'office')).toThrow('类型')
})

test('附件与标记、草稿、收藏随备份恢复，附件随单文件和分类导出；旧备份仍可恢复', async () => {
  const doc = library.createMarkdown('图片笔记', null)
  const image = library.addImage(doc.id, png, 'image/png', '样本.png')
  library.saveMarkdown(doc.id, `重点内容\n![测试](attachments/${image.id}.png)`, 1)
  library.saveAnnotation(doc.id, mark, 2)
  library.saveDraft(doc.id, '未完成草稿', 2, 'markdown')
  library.setFavorites([doc.id], true)
  const before = library.documentState(doc.id)
  const transfer = new TransferService(library, join(root, 'data'))
  const path = join(root, '测试.localdocs-backup')
  await transfer.backup(path)
  library.discardDraft(doc.id)
  const preview = await transfer.preview(path)
  await transfer.restore(preview.token)
  expect(library.documentState(doc.id)).toEqual(before)
  expect(readFileSync(library.attachmentPath(doc.id, image.id).path)).toEqual(png)
  library.exportDocument(doc.id, join(root, '导出.md'))
  expect(readFileSync(join(root, 'attachments', `${image.id}.png`))).toEqual(png)
  const exported = await transfer.export(root, { kind: 'all' })
  const manifest = JSON.parse(readFileSync(join(exported.path, '导出清单.json'), 'utf8'))
  const relative = manifest.files[0].path
  expect(
    readFileSync(join(exported.path, relative.replace(/[^/\\]+$/, `attachments/${image.id}.png`))),
  ).toEqual(png)
  const legacy = library.backupManifest()
  legacy.version = 1
  for (const d of legacy.documents) delete d.state
  legacy.objects = legacy.objects.filter(
    (o) =>
      legacy.documents.some((d) => d.blobHash === o.hash) ||
      legacy.versions.some((v) => v.blobHash === o.hash),
  )
  library.applyBackup(validateManifest(legacy))
  expect(library.documentState(doc.id)).toEqual({ annotations: [], attachments: [], draft: null })
})

test('拒绝伪造标记、附件路径和不完整的新版备份', () => {
  const doc = library.createMarkdown('标记校验', null)
  expect(() => library.saveAnnotation(doc.id, { ...mark, scope: '', offset: -1 }, 1)).toThrow()
  expect(() => library.saveAnnotation(doc.id, { ...mark, color: 'red' as 'yellow' }, 1)).toThrow()
  const manifest = library.backupManifest()
  manifest.documents[0].state!.annotations.push({
    ...mark,
    id: 'invalid',
    documentId: doc.id,
    revision: 1,
    createdAt: new Date().toISOString(),
  })
  expect(() => validateManifest(manifest)).toThrow()
  expect(() => library.attachmentPath(doc.id, '../file')).toThrow()
})

test('Word 跨文字片段修改保留粗体、图片、页眉页脚、表格和其他压缩条目', () => {
  const original = wordFixture(),
    model = officeEditModel(original, '.docx', 1)
  const field = model.fields.find((f) => f.text === '请检查端口配置与连接方式。')!
  const modified = applyOfficeChanges(original, '.docx', [
    { key: field.key, text: '请核对端口与网络连接。' },
  ])
  expect(
    parseOffice(modified, '.docx').blocks.some((b) => b.text === '请核对端口与网络连接。'),
  ).toBe(true)
  const before = unzipSync(original),
    after = unzipSync(modified)
  for (const key of Object.keys(before).filter((k) => k !== 'word/document.xml'))
    expect(after[key]).toEqual(before[key])
  const xml = strFromU8(after['word/document.xml'])
  expect(xml).toContain('<w:b/>')
  expect(xml).toContain('<w:tbl>')
  expect(xml).toContain('<w:drawing>')
  expect(() =>
    applyOfficeChanges(original, '.docx', [
      { key: model.fields.find((f) => !f.editable)!.key, text: '不可改' },
    ]),
  ).toThrow('不可编辑')
  expect(() =>
    applyOfficeChanges(original, '.docx', [{ key: field.key, text: '换行\n不可改' }]),
  ).toThrow('单段')
  expect(applyOfficeChanges(original, '.docx', [])).toEqual(Buffer.from(original))
})

test('Excel 修改真实工作表及单元格，保留其他内容，清除旧公式结果且不冒充重算', () => {
  const original = sheetFixture(),
    model = officeEditModel(original, '.xlsx', 1)
  const field = model.fields.find((f) => f.label === '设备清单 · D120')!
  const numeric = model.fields.find((f) => f.label === '预算 · B2')!
  const modified = applyOfficeChanges(original, '.xlsx', [
    { key: field.key, text: '第三个端口 & <文字>' },
    { key: numeric.key, text: '120' },
  ])
  const data = parseOffice(modified, '.xlsx')
  expect(data.sheets[0].cells.find((c) => c.address === 'D120')!.text).toBe('第三个端口 & <文字>')
  expect(data.sheets[1].cells.find((c) => c.address === 'B2')!.text).toBe('120')
  expect(data.sheets[1].cells.find((c) => c.address === 'D2')).toMatchObject({
    formula: 'SUM(B2:C2)',
    text: '（无缓存结果）',
  })
  expect(data.sheets[0].merges).toEqual(['A1:B1'])
  expect(unzipSync(modified)['xl/sharedStrings.xml']).toEqual(
    unzipSync(original)['xl/sharedStrings.xml'],
  )
  expect(() =>
    applyOfficeChanges(original, '.xlsx', [{ key: numeric.key, text: '=HYPERLINK("x")' }]),
  ).toThrow('公式')
  expect(() =>
    applyOfficeChanges(original, '.xlsx', [
      { key: model.fields.find((f) => !f.editable)!.key, text: '1' },
    ]),
  ).toThrow('不可编辑')
})

test('Office 编辑保存、历史恢复、原件保护和旧版本拒绝', () => {
  const source = join(root, '原文.docx'),
    bytes = Buffer.from(wordFixture())
  writeFileSync(source, bytes)
  const doc = library.importFile(source, null),
    model = officeEditModel(bytes, '.docx', 1)
  const field = model.fields.find((f) => f.editable)!
  library.saveDraft(doc.id, JSON.stringify([{ key: field.key, text: '新的标题' }]), 1, 'office')
  const first = library.versions(doc.id)[0]
  const result = library.saveOfficeBytes(
    doc.id,
    applyOfficeChanges(bytes, '.docx', [{ key: field.key, text: '新的标题' }]),
    1,
  )
  expect(result.revision).toBe(2)
  expect(library.documentState(doc.id).draft).toBeNull()
  expect(() => library.saveOfficeBytes(doc.id, bytes, 1)).toThrow('变化')
  expect(library.versions(doc.id)).toHaveLength(2)
  library.restoreVersion(doc.id, first.id, 2)
  expect(readFileSync(library.documentPath(doc.id))).toEqual(bytes)
  expect(readFileSync(source)).toEqual(bytes)
})

test('新建 Word 和 Excel 可编辑、保存并导出；Word 可追加多段正文', () => {
  for (const extension of ['.docx', '.xlsx'] as const) {
    const doc = library.createOffice('新建资料', extension, null)
    const bytes = readFileSync(library.documentPath(doc.id))
    const model = officeEditModel(bytes, extension, 1)
    const field = model.fields[0]
    const changes = [{ key: field.key, text: '新建后保存的内容' }]
    if (extension === '.docx')
      changes.push({ key: model.fields.at(-1)!.key, text: '第二段\n第三段' })
    const updated = applyOfficeChanges(bytes, extension, changes)
    library.saveOfficeBytes(doc.id, updated, 1)
    const data = parseOffice(updated, extension)
    if (extension === '.docx')
      expect(data.blocks.map((b) => b.text)).toEqual(['新建后保存的内容', '第二段', '第三段'])
    else {
      expect(data.sheets[0].cells).toHaveLength(160)
      expect(data.sheets[0].cells[0].text).toBe('新建后保存的内容')
    }
    const output = join(root, '导出' + extension)
    library.exportDocument(doc.id, output)
    expect(readFileSync(output)).toEqual(updated)
    expect(library.versions(doc.id)).toHaveLength(2)
  }
})

test('Excel 数组公式结果区域也保持只读', () => {
  const parts = unzipSync(sheetFixture())
  const xml = strFromU8(parts['xl/worksheets/sheet9.xml']).replace(
    '<c r="B2"><v>100</v></c>',
    '<c r="B2"><f t="array" ref="B2:C2">{100,200}</f><v>100</v></c>',
  )
  const model = officeEditModel(
    officeZip({ ...parts, 'xl/worksheets/sheet9.xml': xml }),
    '.xlsx',
    1,
  )
  expect(model.fields.find((f) => f.label === '预算 · C2')!.editable).toBe(false)
})

test('整段修订和内容控件禁止改写，新增段落保持在页设置之前', () => {
  const parts = unzipSync(newOfficeFile('.docx'))
  const xml = strFromU8(parts['word/document.xml']).replace(
    '<w:sectPr>',
    '<w:ins w:id="1"><w:p><w:r><w:t>待审核</w:t></w:r></w:p></w:ins><w:sdt><w:sdtContent><w:p><w:r><w:t>锁定控件</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr>',
  )
  const bytes = officeZip({ ...parts, 'word/document.xml': xml })
  const model = officeEditModel(bytes, '.docx', 1)
  expect(model.fields.find((f) => f.text === '待审核')!.editable).toBe(false)
  expect(model.fields.find((f) => f.text === '锁定控件')!.editable).toBe(false)
  const result = applyOfficeChanges(bytes, '.docx', [
    { key: 'word/document.xml:append', text: '附录' },
  ])
  const after = strFromU8(unzipSync(result)['word/document.xml'])
  expect(after.indexOf('附录')).toBeLessThan(after.indexOf('<w:sectPr>'))
  expect(after).toContain('<w:ins w:id="1"><w:p><w:r><w:t>待审核</w:t>')
})

test('Excel 保留布尔类型和既有计算选项，重算标记位于合法位置', () => {
  const parts = unzipSync(sheetFixture())
  const wb = strFromU8(parts['xl/workbook.xml']).replace(
    '</workbook>',
    '<calcPr calcMode="manual" iterate="1" iterateCount="200"/><extLst/></workbook>',
  )
  const bytes = officeZip({ ...parts, 'xl/workbook.xml': wb })
  const model = officeEditModel(bytes, '.xlsx', 1)
  const field = model.fields.find((f) => f.label === '设备清单 · C2')!
  const result = applyOfficeChanges(bytes, '.xlsx', [{ key: field.key, text: 'FALSE' }])
  const files = unzipSync(result),
    changed = strFromU8(files['xl/workbook.xml'])
  expect(changed).toContain('iterateCount="200"')
  expect(changed).toContain('calcMode="manual"')
  expect(changed).toContain('fullCalcOnLoad="1"')
  expect(changed.indexOf('<calcPr')).toBeLessThan(changed.indexOf('<extLst'))
  expect(strFromU8(files['xl/worksheets/sheet2.xml'])).toContain('<c r="C2" t="b"><v>0</v></c>')
})

test('Office 草稿拒绝错误结构、重复键和错误文档类型的备份', () => {
  const doc = library.createOffice('草稿验证', '.docx', null)
  expect(() => library.saveDraft(doc.id, '{}', 1, 'office')).toThrow('草稿')
  expect(() =>
    library.saveDraft(doc.id, '[{"key":"a","text":"b"},{"key":"a","text":"c"}]', 1, 'office'),
  ).toThrow('草稿')
  const manifest = library.backupManifest()
  manifest.documents[0].state!.draft = {
    kind: 'markdown',
    text: '错误类型',
    revision: 1,
    updatedAt: new Date().toISOString(),
  }
  expect(() => validateManifest(manifest)).toThrow('无效')
  expect(() => validateManifest({ ...library.backupManifest(), version: '2' })).toThrow('无效')
})

test('真实 schema v4 升级到 v5 保留文件、收藏与历史', () => {
  const doc = library.createMarkdown('升级保留', null)
  library.saveMarkdown(doc.id, '升级前的正文', 1)
  library.setFavorites([doc.id], true)
  const before = library.readDocument(doc.id),
    versions = library.versions(doc.id)
  library.close()
  const db = new DatabaseSync(join(root, 'data', 'library', 'library.sqlite'))
  db.exec('DROP TABLE document_state; PRAGMA user_version=4;')
  db.close()
  library = new Library(join(root, 'data', 'library'))
  expect(library.readDocument(doc.id)).toEqual(before)
  expect(library.versions(doc.id)).toEqual(versions)
  expect(library.documentState(doc.id)).toEqual({ annotations: [], attachments: [], draft: null })
})
