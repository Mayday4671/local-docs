import { expect, test } from 'vitest'
import { strFromU8, unzipSync } from 'fflate'
import { officeEditModel, applyOfficeChanges } from '../src/main/office-edit'
import { parseOffice } from '../src/main/office-parser'
import { officeZip, sheetFixture, wordFixture } from './fixtures/office'
import { tabStopWordFixture } from './fixtures/word-editing'

test('段落制表位是格式，标题和表格普通文字可编辑，原格式保持不变', () => {
  const bytes = tabStopWordFixture(),
    model = officeEditModel(bytes, '.docx', 1)
  const title = model.fields.find((f) => f.text === '年度项目记录')!
  const cell = model.fields.find((f) => f.text === '示例项目')!
  expect(title.editable).toBe(true)
  expect(cell.editable).toBe(true)
  expect(title.anchor).toBeTruthy()
  const output = applyOfficeChanges(bytes, '.docx', [
    { key: title.key, text: '可编辑的年度记录' },
    { key: cell.key, text: '更新项目' },
  ])
  const before = strFromU8(unzipSync(bytes)['word/document.xml'])
  const after = strFromU8(unzipSync(output)['word/document.xml'])
  expect(after).toBe(
    before
      .replace('<w:t>年度项目记录</w:t>', '<w:t xml:space="preserve">可编辑的年度记录</w:t>')
      .replace('<w:t>示例项目</w:t>', '<w:t xml:space="preserve">更新项目</w:t>'),
  )
})

test('无文字节点的空白表格段落可以输入、保存、清空和再次输入', () => {
  const bytes = tabStopWordFixture(),
    model = officeEditModel(bytes, '.docx', 1)
  const field = model.fields.find((f) => f.text === '' && f.editable && !f.key.endsWith(':append'))!
  expect(field.anchor).toBeTruthy()
  const output = applyOfficeChanges(bytes, '.docx', [{ key: field.key, text: '备注 & <详情>' }])
  const after = strFromU8(unzipSync(output)['word/document.xml'])
  expect(after).toContain('备注 &amp; &lt;详情&gt;')
  expect(after.match(/<w:tabs>/g)).toHaveLength(3)
  expect(after).toContain('<w:tblBorders>')
  const cleared = applyOfficeChanges(output, '.docx', [{ key: field.key, text: '' }])
  const again = applyOfficeChanges(cleared, '.docx', [{ key: field.key, text: '再次输入' }])
  expect(officeEditModel(again, '.docx', 4).fields.find((f) => f.key === field.key)?.text).toBe(
    '再次输入',
  )
})

test('正文实际制表符、换行与受保护段落仍不可改写', () => {
  const files = unzipSync(tabStopWordFixture())
  for (const token of ['<w:tab/>', '<w:br/>', '<w:fldChar w:fldCharType="begin"/>']) {
    const bytes = officeZip({
      ...files,
      'word/document.xml': strFromU8(files['word/document.xml']).replace(
        '<w:t>年度项目记录</w:t>',
        `<w:t>年度项目记录</w:t>${token}`,
      ),
    })
    const field = officeEditModel(bytes, '.docx', 1).fields.find((f) => f.text === '年度项目记录')!
    expect(field.editable).toBe(false)
    expect(() =>
      applyOfficeChanges(bytes, '.docx', [{ key: field.key, text: '不允许修改' }]),
    ).toThrow('不可编辑')
  }
})

test('Word 页面使用独立位置锚点，重复文本也不串段，保存不混入预览锚点', () => {
  const files = unzipSync(wordFixture())
  const xml = strFromU8(files['word/document.xml']).replace('交换机', '设备')
  const bytes = officeZip({ ...files, 'word/document.xml': xml })
  const model = officeEditModel(bytes, '.docx', 1)
  const duplicates = model.fields.filter((f) => f.text === '设备')
  expect(duplicates).toHaveLength(2)
  expect(duplicates[0].anchor).not.toBe(duplicates[1].anchor)
  const preview = strFromU8(unzipSync(model.layoutBytes!)['word/document.xml'])
  for (const field of duplicates) expect(preview).toContain(`w:name="${field.anchor}"`)
  const saved = applyOfficeChanges(bytes, '.docx', [{ key: duplicates[1].key, text: '新设备' }])
  const output = strFromU8(unzipSync(saved)['word/document.xml'])
  expect(output).not.toContain(duplicates[1].anchor)
  expect(
    parseOffice(saved, '.docx')
      .blocks.filter((b) => ['设备', '新设备'].includes(b.text))
      .map((b) => b.text),
  ).toEqual(['设备', '新设备'])
  expect(unzipSync(saved)['word/styles.xml']).toEqual(files['word/styles.xml'])
  expect(unzipSync(saved)['word/media/sample.png']).toEqual(files['word/media/sample.png'])
})

test('Excel 空白格可创建，行列保持顺序，保留合并、样式与非目标工作表', () => {
  const original = sheetFixture()
  const model = officeEditModel(original, '.xlsx', 1)
  expect(model.sheets?.[0]).toMatchObject({
    name: '设备清单',
    part: 'xl/worksheets/sheet2.xml',
    merges: ['A1:B1'],
  })
  const saved = applyOfficeChanges(original, '.xlsx', [
    { key: 'xl/worksheets/sheet2.xml:B2', text: '新格子' },
    { key: 'xl/worksheets/sheet2.xml:E100', text: '123.5' },
    { key: 'xl/worksheets/sheet2.xml:A3', text: '00042' },
    { key: 'xl/worksheets/sheet2.xml:A1', text: '合并标题' },
  ])
  const data = parseOffice(saved, '.xlsx').sheets[0]
  expect(data.cells.find((c) => c.address === 'B2')?.text).toBe('新格子')
  expect(data.cells.find((c) => c.address === 'E100')?.text).toBe('123.5')
  expect(data.cells.find((c) => c.address === 'A3')?.text).toBe('00042')
  expect(data.merges).toEqual(['A1:B1'])
  const xml = strFromU8(unzipSync(saved)['xl/worksheets/sheet2.xml'])
  expect(xml.indexOf('r="B2"')).toBeLessThan(xml.indexOf('r="C2"'))
  expect(xml.indexOf('r="3"')).toBeLessThan(xml.indexOf('r="100"'))
  expect(xml.indexOf('r="100"')).toBeLessThan(xml.indexOf('r="120"'))
  expect(unzipSync(saved)['xl/sharedStrings.xml']).toEqual(
    unzipSync(original)['xl/sharedStrings.xml'],
  )
})

test('Excel 拒绝合并从属格、越界、未知工作表和受保护空白格', () => {
  const bytes = sheetFixture()
  for (const key of [
    'xl/worksheets/sheet2.xml:B1',
    'xl/worksheets/sheet2.xml:XFE1',
    'xl/worksheets/sheet2.xml:A1048577',
    'xl/worksheets/sheet2.xml:Z0',
    'xl/worksheets/missing.xml:A1',
  ])
    expect(() => applyOfficeChanges(bytes, '.xlsx', [{ key, text: '错误位置' }])).toThrow(
      '不可编辑',
    )
  const files = unzipSync(bytes)
  const xml = strFromU8(files['xl/worksheets/sheet2.xml']).replace(
    '</worksheet>',
    '<sheetProtection sheet="1"/></worksheet>',
  )
  const protectedFile = officeZip({ ...files, 'xl/worksheets/sheet2.xml': xml })
  expect(() =>
    applyOfficeChanges(protectedFile, '.xlsx', [
      { key: 'xl/worksheets/sheet2.xml:B2', text: '禁止' },
    ]),
  ).toThrow('不可编辑')
  const arrayXml = strFromU8(files['xl/worksheets/sheet9.xml']).replace(
    '<f>SUM(B2:C2)</f>',
    '<f t="array" ref="D2:D4">SUM(B2:C2)</f>',
  )
  expect(() =>
    applyOfficeChanges(officeZip({ ...files, 'xl/worksheets/sheet9.xml': arrayXml }), '.xlsx', [
      { key: 'xl/worksheets/sheet9.xml:D3', text: '禁止' },
    ]),
  ).toThrow('不可编辑')
})

test('完全空白的工作表、空行和旧 dimension 可正常追加格子', () => {
  const files = unzipSync(sheetFixture())
  for (const content of ['<sheetData/>', '<sheetData><row r="1"/><row r="2"/></sheetData>']) {
    const bytes = officeZip({
      ...files,
      'xl/worksheets/sheet2.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/>${content}</worksheet>`,
    })
    const saved = applyOfficeChanges(bytes, '.xlsx', [
      { key: 'xl/worksheets/sheet2.xml:B2', text: '你好' },
      { key: 'xl/worksheets/sheet2.xml:C5', text: '3' },
    ])
    expect(parseOffice(saved, '.xlsx').sheets[0].cells.map((c) => c.address)).toEqual(['B2', 'C5'])
    expect(strFromU8(unzipSync(saved)['xl/worksheets/sheet2.xml'])).not.toContain('<dimension')
  }
})
