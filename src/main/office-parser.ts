import { strFromU8, unzipSync } from 'fflate'
import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { posix } from 'node:path'
import { format as formatNumber } from 'ssf'
import type { OfficeData, OfficeSheet, SheetCell, TextBlock } from '../shared/types'

const MAX_BYTES = 20 * 1024 * 1024
const MAX_EXPANDED = 64 * 1024 * 1024
const MAX_XML = 8 * 1024 * 1024
const MAX_CELLS = 50000
const MAX_TEXT = 2 * 1024 * 1024
const parser = new XMLParser({
  htmlEntities: true,
  ignoreAttributes: false,
  removeNSPrefix: true,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
})
const ordered = new XMLParser({
  htmlEntities: true,
  ignoreAttributes: false,
  removeNSPrefix: true,
  preserveOrder: true,
  parseTagValue: false,
  trimValues: false,
})
// The parser never resolves a filesystem path or fetches external relationships.
function xml(files: Record<string, Uint8Array>, name: string, inOrder = false): any {
  const bytes = files[name]
  if (!bytes) throw new Error(`文件缺少必要的内容：${name}`)
  if (bytes.byteLength > MAX_XML) throw new Error('文档内容过大，当前无法预览。')
  const text = strFromU8(bytes)
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('文档包含不支持的 XML 声明。')
  if (XMLValidator.validate(text) !== true) throw new Error('文档内部 XML 已损坏。')
  return (inOrder ? ordered : parser).parse(text)
}
const list = <T>(value: T | T[] | undefined): T[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value]
const textOf = (value: any): string =>
  typeof value === 'string' ? value : (value?.['#text'] ?? '')
const richText = (value: any): string =>
  value?.t !== undefined
    ? textOf(value.t)
    : list<any>(value?.r)
        .map((run) => textOf(run.t))
        .join('')

function wordText(nodes: any[]): string {
  let text = ''
  for (const node of nodes)
    for (const [key, value] of Object.entries(node)) {
      if (key === 't') text += (value as any[]).map((item) => item['#text'] ?? '').join('')
      else if (key === 'tab') text += '\t'
      else if (key === 'br' || key === 'cr') text += '\n'
      else if (Array.isArray(value) && !['del', 'instrText', 'drawing', 'pict'].includes(key))
        text += wordText(value)
    }
  return text
}

function paragraphs(nodes: any[], part: string, blocks: TextBlock[]): void {
  let index = 0
  const visit = (items: any[]) => {
    for (const node of items)
      for (const [key, value] of Object.entries(node)) {
        if (key === 'p') {
          const text = wordText(value as any[])
          if (text.trim()) blocks.push({ text, location: { kind: 'paragraph', part, index } })
          index++
        } else if (Array.isArray(value) && key !== 'del') visit(value)
      }
  }
  visit(nodes)
}

function columnIndex(address: string): number {
  let value = 0
  for (const ch of address.replace(/[0-9]/g, '')) value = value * 26 + ch.charCodeAt(0) - 64
  return value
}

export function parseOffice(bytes: Uint8Array, extension: string): OfficeData {
  if (bytes.byteLength > MAX_BYTES)
    throw new Error('当前仅预览 20 MB 以内的 Office 文件；原文件仍可导出。')
  if (extension !== '.docx' && extension !== '.xlsx') throw new Error('不支持的 Office 文件类型。')
  let expanded = 0
  let count = 0
  const names = new Set<string>()
  let files: Record<string, Uint8Array>
  try {
    files = unzipSync(bytes, {
      filter(entry) {
        expanded += entry.originalSize
        if (
          ++count > 2000 ||
          expanded > MAX_EXPANDED ||
          entry.originalSize > MAX_EXPANDED ||
          names.has(entry.name)
        )
          throw new Error('解压资源超过限制')
        if (
          entry.name.includes('\\') ||
          entry.name.startsWith('/') ||
          entry.name.split('/').includes('..')
        )
          throw new Error('异常文件路径')
        names.add(entry.name)
        return true
      },
    })
  } catch {
    throw new Error('文件已损坏、已加密，或解压资源超出预览限制；仍可原样导出。')
  }
  xml(files, '[Content_Types].xml')
  const blocks: TextBlock[] = []
  const sheets: OfficeSheet[] = []
  const warnings: string[] = []
  if (extension === '.docx') {
    paragraphs(xml(files, 'word/document.xml', true), '正文', blocks)
    for (const name of Object.keys(files)
      .filter((name) => /^word\/(header|footer)\d+\.xml$/.test(name))
      .sort()) {
      paragraphs(
        xml(files, name, true),
        name.includes('header')
          ? `页眉 ${name.match(/\d+/)?.[0]}`
          : `页脚 ${name.match(/\d+/)?.[0]}`,
        blocks,
      )
    }
    warnings.push('只读预览；复杂分页、域和修订可能与 Word 显示不同。')
  } else {
    const workbook = xml(files, 'xl/workbook.xml').workbook
    const relationships = list<any>(
      xml(files, 'xl/_rels/workbook.xml.rels').Relationships?.Relationship,
    )
    const strings = files['xl/sharedStrings.xml']
      ? list<any>(xml(files, 'xl/sharedStrings.xml').sst?.si).map(richText)
      : []
    const styles = files['xl/styles.xml'] ? xml(files, 'xl/styles.xml').styleSheet : undefined
    const formats = new Map<number, string>(
      list<any>(styles?.numFmts?.numFmt).map((item) => [
        Number(item['@_numFmtId']),
        String(item['@_formatCode']),
      ]),
    )
    const cellStyles = list<any>(styles?.cellXfs?.xf)
    const date1904 = ['1', 'true'].includes(workbook?.workbookPr?.['@_date1904'])
    let cellCount = 0
    for (const sheet of list<any>(workbook?.sheets?.sheet)) {
      if (sheets.length >= 100) throw new Error('工作表超过 100 张，当前无法预览。')
      const rel = relationships.find((item) => item['@_Id'] === sheet['@_id'])
      if (
        !rel ||
        rel['@_TargetMode'] === 'External' ||
        !String(rel['@_Type']).endsWith('/worksheet')
      ) {
        warnings.push(`未读取非普通工作表：${sheet['@_name']}`)
        continue
      }
      const target = String(rel['@_Target'])
      const name = target.startsWith('/')
        ? posix.normalize(target.slice(1))
        : posix.normalize(posix.join('xl', target))
      if (!name.startsWith('xl/') || name.includes(':')) throw new Error('工作表路径无效。')
      const source = xml(files, name).worksheet
      const cells: SheetCell[] = []
      const sheetName = String(sheet['@_name'])
      for (const row of list<any>(source?.sheetData?.row))
        for (const cell of list<any>(row.c)) {
          if (++cellCount > MAX_CELLS) throw new Error('非空单元格超过 50,000 个，当前无法预览。')
          const address = String(cell['@_r'] ?? '')
          if (!/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(address)) throw new Error('单元格坐标无效。')
          const column = columnIndex(address)
          const rowNumber = Number(address.replace(/^[A-Z]+/, ''))
          if (column > 16384 || rowNumber > 1048576) throw new Error('单元格坐标超出 Excel 范围。')
          const value = textOf(cell.v)
          let text =
            cell['@_t'] === 's'
              ? strings[Number(value)]
              : cell['@_t'] === 'inlineStr'
                ? richText(cell.is)
                : cell['@_t'] === 'b'
                  ? value === '1'
                    ? 'TRUE'
                    : 'FALSE'
                  : value
          if (text === undefined) throw new Error('共享字符串引用无效。')
          const formula = cell.f === undefined ? undefined : textOf(cell.f)
          const numeric = !cell['@_t'] || cell['@_t'] === 'n'
          if (numeric && value.trim() && Number.isFinite(Number(value))) {
            const formatId = Number(cellStyles[Number(cell['@_s'] ?? 0)]?.['@_numFmtId'] ?? 0)
            try {
              text = formatNumber(formats.get(formatId) ?? formatId, Number(value), { date1904 })
            } catch {
              text = value
              if (!warnings.includes('部分数字格式无法识别，已保留原始值。'))
                warnings.push('部分数字格式无法识别，已保留原始值。')
            }
          }
          if (cell.v === undefined && cell.is === undefined && formula !== undefined)
            text = '（无缓存结果）'
          const parsed = {
            address,
            row: rowNumber,
            column,
            text,
            ...(numeric && value !== text ? { rawValue: value } : {}),
            ...(formula === undefined ? {} : { formula }),
          }
          cells.push(parsed)
          if (text || formula)
            blocks.push({
              text: formula ? `${text}\n公式：=${formula}` : text,
              location: { kind: 'cell', sheet: sheetName, address },
            })
        }
      cells.sort((a, b) => a.row - b.row || a.column - b.column)
      sheets.push({
        name: sheetName,
        cells,
        hidden: ['hidden', 'veryHidden'].includes(sheet['@_state']),
        merges: list<any>(source?.mergeCells?.mergeCell).map((merge) => String(merge['@_ref'])),
      })
    }
    if (!sheets.length) throw new Error('工作簿没有可读取的普通工作表。')
    warnings.push(
      '只读数据视图；按数字格式显示日期和数值，公式使用文件中保存的结果，不重新计算。图表、条件格式和复杂样式尚未呈现。',
    )
  }
  if (
    blocks.length > MAX_CELLS ||
    blocks.reduce((size, block) => size + block.text.length, 0) > MAX_TEXT
  )
    throw new Error('提取文本超过当前预览限制。')
  return { kind: extension === '.docx' ? 'docx' : 'xlsx', blocks, sheets, warnings }
}
