import { describe, expect, test } from 'vitest'
import { compareVersions } from '../src/shared/version-diff'
import type { OfficeData, VersionContent } from '../src/shared/types'

const md = (text: string, hash = text): VersionContent => ({
  version: { id: 'version', createdAt: '', reason: '' },
  extension: '.md',
  hash,
  text,
  office: null,
})
const office = (data: OfficeData, hash = JSON.stringify(data)): VersionContent => ({
  ...md('', hash),
  extension: `.${data.kind}`,
  text: null,
  office: data,
})
describe('历史版本内容对比', () => {
  test('对齐插入和删除，不把后续内容全部视为修改', () => {
    const result = compareVersions(
      md('标题\n删除的段\n相同\n旧值\n结尾'),
      md('标题\n相同\n新值\n结尾\n新增的段'),
    )
    expect(result.counts).toEqual({ same: 3, removed: 1, added: 1, modified: 1 })
    expect(result.rows.find((r) => r.kind === 'modified')).toMatchObject({
      before: { text: '旧值', label: '第 4 行' },
      after: { text: '新值', label: '第 3 行' },
    })
  })
  test('重复文本、空行、末尾换行和换行编码', () => {
    const result = compareVersions(md('a\na\nb\n'), md('a\nb\n'))
    expect(result.counts).toEqual({ same: 3, removed: 1, added: 0, modified: 0 })
    expect(compareVersions(md('a\r\nb'), md('a\nb')).counts.same).toBe(2)
    expect(compareVersions(md('a'), md('a\n')).counts.added).toBe(1)
    expect(compareVersions(md(''), md('a')).counts.added).toBe(1)
    expect(compareVersions(md('a'), md('')).counts.removed).toBe(1)
  })
  test('任意方向和相同版本都能对比，输入内容不变', () => {
    const a = md('一\n二'),
      b = md('零\n一\n二')
    const snapshot = JSON.stringify([a, b])
    expect(compareVersions(a, b).counts.added).toBe(1)
    expect(compareVersions(b, a).counts.removed).toBe(1)
    expect(compareVersions(a, a).identical).toBe(true)
    expect(JSON.stringify([a, b])).toBe(snapshot)
  })
  test('较大文档利用内容锚点对齐，不丢行；大量重写退化配对有提示', () => {
    const a = Array.from({ length: 4000 }, (_, i) => `段落${i}`)
    const b = ['新增', ...a.slice(0, 2000), '中间新增', ...a.slice(2000), '尾部新增']
    expect(compareVersions(md(a.join('\n')), md(b.join('\n'))).counts).toEqual({
      same: 4000,
      added: 3,
      removed: 0,
      modified: 0,
    })
    const result = compareVersions(
      md(Array(2000).fill('旧').join('\n')),
      md(Array(2000).fill('新').join('\n')),
    )
    expect(result.counts.modified).toBe(2000)
    expect(result.warnings.join()).toContain('按位置配对')
  })
  test('Word 按部件和段落内容对齐，正文插入不影响页眉和后续段落', () => {
    const word = (values: string[]): OfficeData => ({
      kind: 'docx',
      sheets: [],
      warnings: [],
      blocks: [
        ...values.map((text, index) => ({
          text,
          location: { kind: 'paragraph' as const, part: 'word/document.xml', index },
        })),
        { text: '相同页眉', location: { kind: 'paragraph', part: 'word/header1.xml', index: 0 } },
      ],
    })
    const result = compareVersions(
      office(word(['标题', '表格文字'])),
      office(word(['标题', '新增', '表格文字'])),
    )
    expect(result.counts).toEqual({ same: 3, added: 1, removed: 0, modified: 0 })
    expect(result.rows[2]).toMatchObject({
      before: { label: '正文 · 段落 2' },
      after: { label: '正文 · 段落 3' },
    })
  })
  test('随机重复行和交叉调整后仍完整保留两侧原文顺序', () => {
    let seed = 73
    const random = () => {
      seed = (seed * 16807) % 2147483647
      return seed
    }
    for (let run = 0; run < 150; run++) {
      const a = Array.from({ length: (random() % 50) + 1 }, () => `值${random() % 8}`)
      const b = Array.from({ length: (random() % 50) + 1 }, () => `值${random() % 8}`)
      const { rows } = compareVersions(md(a.join('\n')), md(b.join('\n')))
      expect(rows.flatMap((r) => (r.before ? [r.before.text] : []))).toEqual(a)
      expect(rows.flatMap((r) => (r.after ? [r.after.text] : []))).toEqual(b)
    }
    expect(() => compareVersions(md('\n'.repeat(100000)), md(''))).toThrow('10 万行')
  })
  test('Excel 比较真实地址、公式、原始值、工作表和合并区域', () => {
    const data: OfficeData = {
      kind: 'xlsx',
      blocks: [],
      warnings: [],
      sheets: [
        {
          name: '明细',
          hidden: false,
          merges: [],
          cells: [
            {
              address: 'AA120',
              row: 120,
              column: 27,
              text: '100',
              rawValue: '100.1',
              formula: 'SUM(A1:A3)',
            },
          ],
        },
      ],
    }
    const next = structuredClone(data)
    next.sheets[0].hidden = true
    next.sheets[0].merges = ['A1:B1']
    next.sheets[0].cells[0].rawValue = '100.2'
    next.sheets[0].cells[0].formula = 'SUM(A1:A4)'
    next.sheets[0].cells.push({ address: 'A2', row: 2, column: 1, text: '新增' })
    next.sheets.push({ name: '空表', cells: [], hidden: false, merges: [] })
    const result = compareVersions(office(data), office(next))
    expect(result.counts).toEqual({ same: 0, added: 2, removed: 0, modified: 2 })
    expect(result.rows[2]).toMatchObject({
      after: { label: '明细 · AA120', text: '100\n原始值：100.2\n公式：=SUM(A1:A4)' },
    })
  })
  test('只有样式等变化时不误报文件完全相同；解析警告保留', () => {
    const data: OfficeData = {
      kind: 'docx',
      sheets: [],
      blocks: [],
      warnings: ['部分内容无法解析'],
    }
    const result = compareVersions(office(data, 'hash-one'), office(data, 'hash-two'))
    expect(result.identical).toBe(false)
    expect(result.counts.modified).toBe(0)
    expect(result.warnings).toEqual(['部分内容无法解析'])
    expect(() => compareVersions(office(data), md(''))).toThrow('同一类型')
  })
})
