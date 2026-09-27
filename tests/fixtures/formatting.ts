import { unzipSync, strFromU8 } from 'fflate'
import { officeZip, sheetFixture, wordFixture } from './office.ts'

// Synthetic fixtures only; never copy personal documents into regression tests.
export const markdownFixture = [
  '# 格式检查样本',
  '',
  '| 项目 | 值 | |',
  '| :--- | ---: | --- |',
  '| 名称 | 演示资料 | |',
  '| 状态 | **正常** | |',
  '| 分隔符 | 左\\|右 | |',
  '',
  '- [x] 已验证',
  '- [ ] 待处理',
  '',
  '~~旧说明~~，**重点**，*备注*。',
  '',
  '> 引用说明',
  '',
  '```text',
  'A    B',
  '  保留缩进',
  '```',
  '',
  '| ' + Array.from({ length: 12 }, (_, i) => `长表列 ${i + 1}`).join(' | ') + ' |',
  '| ' + Array(12).fill('---').join(' | ') + ' |',
  '| ' + Array(12).fill('中文内容').join(' | ') + ' |',
  '',
  '<script>window.__unsafe = true</script>',
  '',
  '![远程图片](https://example.invalid/tracker.png)',
  '',
  '[外部链接](https://example.invalid/)',
].join('\r\n')

export function formattedSheetFixture(date1904 = false): Uint8Array {
  const files = Object.fromEntries(
    Object.entries(unzipSync(sheetFixture())).map(([name, bytes]) => [name, strFromU8(bytes)]),
  )
  files['xl/workbook.xml'] = files['xl/workbook.xml'].replace(
    '<sheets>',
    `<workbookPr date1904="${date1904 ? '1' : '0'}"/><sheets>`,
  )
  files['xl/styles.xml'] =
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="3"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="00000"/><numFmt numFmtId="166" formatCode=";;;"/></numFmts>
    <cellXfs count="7"><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="10"/><xf numFmtId="4"/><xf numFmtId="165"/><xf numFmtId="14"/><xf numFmtId="166"/></cellXfs>
  </styleSheet>`
  files['xl/worksheets/sheet9.xml'] =
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>格式检查</t></is></c><c r="B1" t="inlineStr"><is><t>显示结果</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>日期</t></is></c><c r="B2" s="1"><v>${date1904 ? '0' : '45292'}</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>百分比</t></is></c><c r="B3" s="2"><v>0.125</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>千分位</t></is></c><c r="B4" s="3"><v>12345.6</v></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>编号</t></is></c><c r="B5" s="4"><v>42</v></c></row>
    <row r="6"><c r="A6" t="inlineStr"><is><t>公式结果</t></is></c><c r="B6" s="2"><f>1/4</f><v>0.25</v></c></row>
    <row r="7"><c r="A7" t="inlineStr"><is><t>多行</t></is></c><c r="B7" t="inlineStr"><is><t xml:space="preserve">第一行&#10;第二行</t></is></c></row>
    <row r="8"><c r="A8" t="inlineStr"><is><t>文本编号</t></is></c><c r="B8" t="inlineStr"><is><t>00042</t></is></c></row>
    <row r="9"><c r="A9" t="inlineStr"><is><t>错误结果</t></is></c><c r="B9" t="e"><v>#DIV/0!</v></c></row>
    <row r="10"><c r="B10" s="5"><v>45292</v></c><c r="C10" s="2"/></row>
    <row r="11"><c r="B11" s="6"><f>1+1</f><v>2</v></c><c r="C11" t="str"><f>IF(TRUE,&quot;&quot;,0)</f><v></v></c></row>
  </sheetData></worksheet>`
  return officeZip(files)
}

export function wideWordFixture(): Uint8Array {
  const files: Record<string, string | Uint8Array> = unzipSync(wordFixture())
  const cells = Array.from(
    { length: 8 },
    (_, i) =>
      `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>第 ${i + 1} 列</w:t><w:br/><w:t>多行文字</w:t></w:r></w:p></w:tc>`,
  ).join('')
  const table = `<w:tbl><w:tblPr><w:tblW w:w="19200" w:type="dxa"/></w:tblPr><w:tblGrid>${Array(8).fill('<w:gridCol w:w="2400"/>').join('')}</w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="8"/></w:tcPr><w:p><w:r><w:t>合并标题</w:t></w:r></w:p></w:tc></w:tr><w:tr>${cells}</w:tr></w:tbl>`
  files['word/document.xml'] = strFromU8(files['word/document.xml'] as Uint8Array)
    .replace('<w:sectPr>', table + '<w:sectPr>')
    .replace('配置与连接方式。', '配置与连接方式。&#10;第二行&#x4E2D;文')
  return officeZip(files)
}
