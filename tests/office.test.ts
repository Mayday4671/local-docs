import { describe, expect, test } from 'vitest'
import { parseOffice } from '../src/main/office-parser'
import { officeZip, wordFixture, sheetFixture } from './fixtures/office'
import { mkdtempSync, readFileSync, writeFileSync, realpathSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { DatabaseSync } from 'node:sqlite'
import { Library } from '../src/main/library'
import { formattedSheetFixture, wideWordFixture } from './fixtures/formatting'

describe('Office 阅读与检索', () => {
  test('Word 编码换行和中文字符可读取，宽表与合并标题不丢失', () => {
    const data = parseOffice(wideWordFixture(), '.docx')
    expect(data.blocks.some((block) => block.text.includes('第二行中文'))).toBe(true)
    expect(data.blocks.some((block) => block.text === '合并标题')).toBe(true)
    expect(data.blocks.some((block) => block.text === '第 8 列\n多行文字')).toBe(true)
  })
  test('Excel 日期、百分比、千分位、补零和公式缓存按原数字格式显示', () => {
    const cells = parseOffice(formattedSheetFixture(), '.xlsx').sheets[1].cells
    const get = (address: string) => cells.find((cell) => cell.address === address)!
    expect(get('B2')).toMatchObject({ text: '2024-01-01', rawValue: '45292' })
    expect(get('B3').text).toBe('12.50%')
    expect(get('B4').text).toBe('12,345.60')
    expect(get('B5').text).toBe('00042')
    expect(get('B6')).toMatchObject({ text: '25.00%', formula: '1/4' })
    expect(get('B7').text).toBe('第一行\n第二行')
    expect(get('B8').text).toBe('00042')
    expect(get('B9').text).toBe('#DIV/0!')
    expect(get('B10').text).toBe('1/1/24')
    expect(get('C10').text).toBe('')
    expect(get('B11')).toMatchObject({ text: '', rawValue: '2', formula: '1+1' })
    expect(get('C11').text).toBe('')
    expect(
      parseOffice(formattedSheetFixture(true), '.xlsx').sheets[1].cells.find(
        (cell) => cell.address === 'B2',
      )?.text,
    ).toBe('1904-01-01')
  })

  test('0.2 的旧预览缓存升级后重新解析，保留文件、分类、标签和历史', () => {
    const root = mkdtempSync(join(tmpdir(), 'local-docs-office-'))
    let library: Library | undefined
    try {
      const dataRoot = join(root, 'library')
      library = new Library(dataRoot)
      const path = join(root, '旧表.XLSX')
      writeFileSync(path, formattedSheetFixture())
      const category = library.createCategory('资料')
      const doc = library.importFile(path, category.id)
      library.updateDocument(doc.id, { tags: ['已有标签'] })
      const source = library.officeSource(doc.id)
      library.cacheOffice(
        doc.id,
        source.hash,
        { kind: 'xlsx', blocks: [], sheets: [], warnings: [] },
        null,
      )
      library.close()
      library = undefined
      const old = new DatabaseSync(join(dataRoot, 'library.sqlite'))
      old.exec('DROP TABLE settings; PRAGMA user_version = 3;')
      old.close()
      library = new Library(dataRoot)
      expect(library.pendingOfficeIds()).toEqual([doc.id])
      expect(library.readDocument(doc.id).document).toMatchObject({
        categoryId: category.id,
        tags: ['已有标签'],
        revision: 1,
      })
      expect(library.versions(doc.id)).toHaveLength(1)
      library.cacheOffice(
        doc.id,
        source.hash,
        parseOffice(readFileSync(source.path), '.xlsx'),
        null,
      )
      expect(library.searchResults('2024-01-01')[0]?.location).toEqual({
        kind: 'cell',
        sheet: '预算',
        address: 'B2',
      })
      expect(readFileSync(source.path)).toEqual(readFileSync(path))
      expect(library.pendingOfficeIds()).toEqual([])
    } finally {
      library?.close()
      const target = realpathSync(root)
      if (!target.startsWith(join(realpathSync(tmpdir()), 'local-docs-office-')))
        throw new Error('Unexpected cleanup target')
      rmSync(target, { recursive: true, force: true })
    }
  })
  test('DOCX 提取分段文本、表格、页眉页脚，合并不同格式的文字并忽略删除修订', () => {
    const result = parseOffice(wordFixture(), '.docx')
    expect(result.kind).toBe('docx')
    expect(result.blocks[1]).toEqual({
      text: '请检查端口配置与连接方式。',
      location: { kind: 'paragraph', part: '正文', index: 1 },
    })
    expect(result.blocks.some((block) => block.text === '交换机')).toBe(true)
    expect(
      result.blocks.some(
        (block) => block.location.kind === 'paragraph' && block.location.part === '页眉 1',
      ),
    ).toBe(true)
    expect(result.blocks.some((block) => block.text.includes('已经删除'))).toBe(false)
    expect(result.blocks.some((block) => block.text.includes('<script>'))).toBe(true)
  })

  test('XLSX 按真实关系读取多个工作表，保留稀疏坐标、共享字符串、公式缓存与合并区信息', () => {
    const result = parseOffice(sheetFixture(), '.xlsx')
    expect(result.sheets.map((sheet) => sheet.name)).toEqual(['设备清单', '预算'])
    expect(result.sheets[0].cells.find((cell) => cell.address === 'A2')?.text).toBe('核心交换机')
    expect(result.blocks.find((block) => block.text === '二号机端口')?.location).toEqual({
      kind: 'cell',
      sheet: '设备清单',
      address: 'D120',
    })
    expect(result.sheets[0].cells.find((cell) => cell.address === 'AA120')?.column).toBe(27)
    expect(result.sheets[0].merges).toEqual(['A1:B1'])
    expect(result.sheets[1].hidden).toBe(true)
    expect(result.sheets[1].cells.find((cell) => cell.address === 'D2')).toMatchObject({
      text: '300',
      formula: 'SUM(B2:C2)',
    })
    expect(result.sheets[1].cells.find((cell) => cell.address === 'E2')?.text).toBe(
      '（无缓存结果）',
    )
  })

  test('损坏文件、异常路径、XML 实体和压缩炸弹不会继续解析', () => {
    expect(() => parseOffice(new Uint8Array([80, 75]), '.docx')).toThrow('损坏')
    expect(() => parseOffice(officeZip({ '../bad.xml': '<x/>' }), '.docx')).toThrow('损坏')
    expect(() =>
      parseOffice(
        officeZip({ '[Content_Types].xml': '<!DOCTYPE x [<!ENTITY x "unsafe">]><x/>' }),
        '.docx',
      ),
    ).toThrow('XML')
    expect(() => parseOffice(officeZip({ '[Content_Types].xml': '<x><y></x>' }), '.docx')).toThrow(
      'XML',
    )
    const bomb = officeZip({ '[Content_Types].xml': '<Types/>' })
    const view = new DataView(bomb.buffer)
    for (let i = 0; i < bomb.length - 46; i++)
      if (view.getUint32(i, true) === 0x02014b50) {
        view.setUint32(i + 24, 100 * 1024 * 1024, true)
        break
      }
    expect(() => parseOffice(bomb, '.docx')).toThrow('超出')
  })

  test('缓存升级和正文检索不修改文件、历史版本或最近使用，重启后索引仍然有效', () => {
    const root = mkdtempSync(join(tmpdir(), 'local-docs-office-'))
    let library: Library | undefined
    try {
      const dataRoot = join(root, 'library')
      library = new Library(dataRoot)
      const source = join(root, '工作簿.xlsx')
      writeFileSync(source, sheetFixture())
      const doc = library.importFile(source, null)
      library.close()
      library = undefined
      // Recreate the previous v1 schema to exercise an in-place migration with existing data.
      const old = new DatabaseSync(join(dataRoot, 'library.sqlite'))
      old.exec(
        `DROP TABLE settings; DROP TABLE office_cache; DROP INDEX category_siblings; ALTER TABLE categories DROP COLUMN parent_id; ALTER TABLE documents DROP COLUMN tags; ALTER TABLE documents DROP COLUMN notes; PRAGMA user_version = 1;`,
      )
      old.close()
      library = new Library(dataRoot)
      expect(library.pendingOfficeIds()).toEqual([doc.id])
      const info = library.officeSource(doc.id)
      library.cacheOffice(doc.id, info.hash, parseOffice(readFileSync(info.path), '.xlsx'), null)
      expect(library.searchResults('端口')[0]).toMatchObject({
        id: doc.id,
        source: 'content',
        location: { kind: 'cell', sheet: '设备清单', address: 'D120' },
      })
      expect(library.readDocument(doc.id).document).toMatchObject({ revision: 1, openedAt: null })
      expect(library.versions(doc.id)).toHaveLength(1)
      const exported = join(root, '导出.xlsx')
      library.exportDocument(doc.id, exported)
      expect(readFileSync(exported)).toEqual(readFileSync(source))
      library.close()
      library = new Library(dataRoot)
      expect(library.pendingOfficeIds()).toEqual([])
      expect(library.search('端')).toEqual([doc.id])
      library.restoreVersion(doc.id, library.versions(doc.id)[0].id, 1)
      expect(library.search('端')).toEqual([doc.id])
      library.cacheOffice(doc.id, 'stale-hash', null, '过期结果')
      expect(library.search('端')).toEqual([doc.id])
    } finally {
      library?.close()
      const target = realpathSync(root)
      if (!target.startsWith(join(realpathSync(tmpdir()), 'local-docs-office-')))
        throw new Error('Unexpected cleanup target')
      rmSync(target, { recursive: true, force: true })
    }
  })
})
