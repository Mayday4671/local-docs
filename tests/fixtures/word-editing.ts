import { unzipSync, strFromU8 } from 'fflate'
import { officeZip, wordFixture } from './office.ts'

// Reproduces paragraph tab-stop formatting and empty table paragraphs without user data.
export function tabStopWordFixture(): Uint8Array {
  const files = unzipSync(wordFixture())
  const properties =
    '<w:pPr><w:tabs><w:tab w:val="center" w:pos="4320"/></w:tabs><w:jc w:val="center"/><w:rPr><w:sz w:val="32"/></w:rPr></w:pPr>'
  const table = `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="5600"/></w:tblGrid>
    <w:tr><w:tc><w:p><w:r><w:t>项目名称</w:t></w:r></w:p></w:tc><w:tc><w:p>${properties}<w:r><w:t>示例项目</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:p><w:r><w:t>补充备注</w:t></w:r></w:p></w:tc><w:tc><w:p>${properties}</w:p></w:tc></w:tr></w:tbl>`
  const xml = strFromU8(files['word/document.xml'])
    .replace('<w:pPr><w:pStyle w:val="Title"/></w:pPr>', properties)
    .replace('项目网络说明', '年度项目记录')
    .replace(/<w:tbl>[\s\S]*?<\/w:tbl>/, table)
  return officeZip({ ...files, 'word/document.xml': xml })
}
