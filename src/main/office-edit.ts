import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { posix } from 'node:path'
import { randomUUID } from 'node:crypto'
import { parseOffice } from './office-parser'
import type {
  OfficeChange,
  OfficeEditField,
  OfficeEditModel,
  OfficeEditSheet,
} from '../shared/types'

const parser = new XMLParser({ ignoreAttributes: false, parseTagValue: false, trimValues: false })
const escape = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const decode = (s: string): string => String(parser.parse(`<t>${s}</t>`).t ?? '')
const arr = <T>(x: T | T[] | undefined): T[] => (x === undefined ? [] : Array.isArray(x) ? x : [x])
type Field = OfficeEditField & { part: string; start: number; end: number; xml: string }
function source(
  bytes: Uint8Array,
  extension: string,
): {
  files: Record<string, Uint8Array>
  fields: Field[]
  kind: 'docx' | 'xlsx'
  sheets: OfficeEditSheet[]
} {
  // Reuse the reader's package, XML and expansion limits before any edit.
  const data = parseOffice(bytes, extension)
  const files = unzipSync(bytes)
  if (Object.keys(files).some((n) => n.startsWith('_xmlsignatures/')))
    throw new Error('带数字签名的文档暂不支持编辑，原件可导出。')
  const fields: Field[] = []
  const sheets: OfficeEditSheet[] = []
  if (extension === '.docx') {
    const protectedDocument = /<w:documentProtection\b/.test(
      strFromU8(files['word/settings.xml'] || new Uint8Array()),
    )
    for (const part of Object.keys(files).filter(
      (n) => n === 'word/document.xml' || /^word\/(header|footer)\d+\.xml$/.test(n),
    )) {
      const xml = strFromU8(files[part])
      // Structured controls and revision wrappers can enclose whole paragraphs.
      const protectedRanges: [number, number][] = [],
        open: number[] = []
      for (const tag of xml.matchAll(/<(\/?)(?:w:del|w:ins|w:sdt|w:moveFrom|w:moveTo)\b[^>]*>/g)) {
        if (tag[0].endsWith('/>')) continue
        if (!tag[1]) open.push(tag.index!)
        else {
          const start = open.pop()
          if (start !== undefined) protectedRanges.push([start, tag.index! + tag[0].length])
        }
      }
      let index = 0
      for (const match of xml.matchAll(/<w:p(?=[\s>])[^>]*>[\s\S]*?<\/w:p>/g)) {
        const fragment = match[0]
        const tokens = [...fragment.matchAll(/<w:t(?=[\s>])[^>]*>([\s\S]*?)<\/w:t>/g)]
        const text = tokens.map((t) => decode(t[1])).join('')
        // A tab stop in paragraph properties is formatting, not a tab in the text.
        // Keep properties verbatim; inspect only the paragraph's content for unsupported nodes.
        const content = fragment.replace(/<w:pPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:pPr>)/g, '')
        const complex =
          protectedDocument ||
          protectedRanges.some(([start, end]) => match.index! >= start && match.index! < end) ||
          /<w:(drawing|pict|object|fldChar|instrText|del|ins|sdt|sym|tab|br|cr|footnoteReference|endnoteReference|txbxContent)\b|<mc:|<!\[CDATA\[/.test(
            content,
          )
        fields.push({
          key: `${part}:${index}`,
          label: `${part === 'word/document.xml' ? '正文' : part.includes('header') ? '页眉' : '页脚'} · 段落 ${++index}`,
          text,
          editable: !complex,
          ...(complex
            ? { reason: '此段含图片、域、修订、换行或受保护结构，请使用阅读视图。' }
            : {}),
          part,
          start: match.index!,
          end: match.index! + fragment.length,
          xml: fragment,
        })
      }
      if (part === 'word/document.xml' && xml.includes('</w:body>')) {
        const section = /<w:sectPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:sectPr>)\s*<\/w:body>/.exec(xml)
        const start = section?.index ?? xml.indexOf('</w:body>')
        fields.push({
          key: `${part}:append`,
          label: '正文 · 末尾新增段落',
          text: '',
          editable: !protectedDocument,
          ...(protectedDocument ? { reason: '此文档受保护，不能新增段落。' } : {}),
          part,
          start,
          end: start,
          xml: '',
        })
      }
    }
  } else {
    const workbook = parser.parse(strFromU8(files['xl/workbook.xml'])).workbook
    const rels = arr<any>(
      parser.parse(strFromU8(files['xl/_rels/workbook.xml.rels'])).Relationships.Relationship,
    )
    for (const sheet of arr<any>(workbook.sheets.sheet)) {
      const rel = rels.find((r) => r['@_Id'] === sheet['@_r:id'])
      if (!rel || rel['@_TargetMode'] === 'External') continue
      const part = rel['@_Target'].startsWith('/')
        ? posix.normalize(rel['@_Target'].slice(1))
        : posix.normalize(posix.join('xl', rel['@_Target']))
      if (!files[part]) continue
      const xml = strFromU8(files[part])
      const parsedSheet = data.sheets.find((s) => s.name === sheet['@_name'])
      const point = (address: string) => {
        const m = /^([A-Z]+)(\d+)$/.exec(address)
        return m
          ? [Number(m[2]), [...m[1]].reduce((v, c) => v * 26 + c.charCodeAt(0) - 64, 0)]
          : null
      }
      const formulaRanges = [
        ...xml.matchAll(/<f\b[^>]*\bref=["']([A-Z]+\d+)(?::([A-Z]+\d+))?["'][^>]*>/g),
      ].map((m) => [point(m[1]), point(m[2] || m[1])])
      if (parsedSheet)
        sheets.push({
          ...parsedSheet,
          part,
          editable: !/<sheetProtection\b/.test(xml),
          protectedRanges: [
            ...xml.matchAll(/<f\b[^>]*\bref=["']([A-Z]+\d+(?::[A-Z]+\d+)?)["'][^>]*>/g),
          ].map((m) => m[1]),
        })
      for (const match of xml.matchAll(/<c(?=[\s>])[^>]*(?:\/>|>[\s\S]*?<\/c>)/g)) {
        const address = match[0].match(/\br=["']([A-Z]+\d+)["']/)?.[1]
        const cell = parsedSheet?.cells.find((c) => c.address === address)
        if (!cell) continue
        const complex =
          /<sheetProtection\b/.test(xml) ||
          /<f\b|\bcm=|\bvm=/.test(match[0]) ||
          (parsedSheet?.merges.some(
            (range) =>
              cell.address !== range.split(':')[0] && inRange([cell.row, cell.column], range),
          ) ??
            false) ||
          formulaRanges.some(
            ([start, end]) =>
              start &&
              end &&
              cell.row >= start[0] &&
              cell.row <= end[0] &&
              cell.column >= start[1] &&
              cell.column <= end[1],
          )
        const type = match[0].match(/\bt=["']([^"']*)["']/)?.[1]
        // Preserve text-like cells, leading zeroes and numeric formats when editing values.
        const text = cell.formula !== undefined ? `=${cell.formula}` : (cell.rawValue ?? cell.text)
        fields.push({
          key: `${part}:${address}`,
          label: `${sheet['@_name']} · ${address}`,
          text,
          editable: !complex,
          ...(complex
            ? {
                reason:
                  '公式单元格或受保护单元格暂不改写；修改输入值后请在支持重算的软件中重新计算。',
              }
            : {}),
          part,
          start: match.index!,
          end: match.index! + match[0].length,
          xml: match[0],
        })
        if (type === 's' || type === 'inlineStr') fields.at(-1)!.text = cell.text
      }
    }
  }
  if (!fields.length && !sheets.length) throw new Error('这份文档没有当前可编辑的段落或单元格。')
  return { files, fields, kind: data.kind, sheets }
}
export function officeEditModel(
  bytes: Uint8Array,
  extension: string,
  revision: number,
): OfficeEditModel {
  const s = source(bytes, extension)
  // Rendering-only anchors map paragraphs by identity, never by text or DOM order.
  // These bytes are never used for saving/exporting the document.
  if (s.kind === 'docx') {
    const prefix = `ld${randomUUID().replace(/-/g, '')}`
    for (const part of new Set(s.fields.map((f) => f.part))) {
      const partFields = s.fields.filter((f) => f.part === part && !f.key.endsWith(':append'))
      let xml = strFromU8(s.files[part])
      for (const field of partFields.sort((a, b) => b.start - a.start)) {
        if (!field.editable) continue
        field.anchor = `${prefix}_${s.fields.indexOf(field)}`
        const markerId = 2147483000 - s.fields.indexOf(field)
        const fragment = field.xml.replace(
          '</w:p>',
          `<w:bookmarkStart w:id="${markerId}" w:name="${field.anchor}"/><w:bookmarkEnd w:id="${markerId}"/></w:p>`,
        )
        xml = xml.slice(0, field.start) + fragment + xml.slice(field.end)
      }
      s.files[part] = strToU8(xml)
    }
  }
  return {
    kind: s.kind,
    revision,
    fields: s.fields.map(({ part: _p, start: _s, end: _e, xml: _x, ...f }) => f),
    ...(s.kind === 'docx' ? { layoutBytes: zipSync(s.files) } : { sheets: s.sheets }),
  }
}
function wordPatch(xml: string, old: string, next: string) {
  if (/[\r\n\t]/.test(next)) throw new Error('当前段落编辑请保持单段文字，不插入换行或制表符。')
  if (!/<w:t(?=[\s>])[^>]*>[\s\S]*?<\/w:t>/.test(xml)) {
    // Empty table cells often contain only a paragraph and its formatting properties.
    // A new run inherits the paragraph style without altering any surrounding structure.
    return xml.replace('</w:p>', `<w:r><w:t xml:space="preserve">${escape(next)}</w:t></w:r></w:p>`)
  }
  let prefix = 0,
    suffix = 0
  while (prefix < old.length && prefix < next.length && old[prefix] === next[prefix]) prefix++
  while (
    suffix < old.length - prefix &&
    suffix < next.length - prefix &&
    old[old.length - 1 - suffix] === next[next.length - 1 - suffix]
  )
    suffix++
  const removedEnd = old.length - suffix,
    inserted = next.slice(prefix, next.length - suffix)
  let position = 0,
    added = false
  return xml.replace(
    /<w:t(?=[\s>])([^>]*)>([\s\S]*?)<\/w:t>/g,
    (_all, attrs: string, encoded: string) => {
      const text = decode(encoded),
        start = position,
        end = start + text.length
      position = end
      let value = text.slice(0, Math.max(0, Math.min(text.length, prefix - start)))
      if (!added && prefix <= end) {
        value += inserted
        added = true
      }
      value += text.slice(Math.max(0, Math.min(text.length, removedEnd - start)))
      return `<w:t${attrs.replace(/\s+xml:space=["'][^"']*["']/g, '')} xml:space="preserve">${escape(value)}</w:t>`
    },
  )
}
export function applyOfficeChanges(
  bytes: Uint8Array,
  extension: string,
  changes: OfficeChange[],
): Buffer {
  if (!Array.isArray(changes) || changes.length > 2000)
    throw new Error('单次最多修改 2000 个段落或单元格。')
  const s = source(bytes, extension),
    seen = new Set<string>(),
    replacements = new Map<string, { field: Field; xml: string }[]>(),
    insertedCells = new Map<string, { address: string; xml: string }[]>()
  for (const change of changes) {
    if (
      !change ||
      typeof change.key !== 'string' ||
      typeof change.text !== 'string' ||
      change.text.length > 32767 ||
      /[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(change.text) ||
      seen.has(change.key)
    )
      throw new Error('编辑内容无效或重复。')
    seen.add(change.key)
    let field = s.fields.find((f) => f.key === change.key)
    if (!field && s.kind === 'xlsx') {
      const sheet = s.sheets.find((sh) => change.key.startsWith(`${sh.part}:`))
      const address = sheet && change.key.slice(sheet.part.length + 1)
      const point = address && cellPoint(address)
      if (
        sheet?.editable &&
        point &&
        !sheet.protectedRanges.some((range) => inRange(point, range)) &&
        !sheet.merges.some((range) => inRange(point, range) && address !== range.split(':')[0])
      ) {
        field = {
          key: change.key,
          label: address!,
          text: '',
          editable: true,
          part: sheet.part,
          start: -1,
          end: -1,
          xml: `<c r="${address}"/>`,
        }
      }
    }
    if (!field || !field.editable) throw new Error('该位置不可编辑，请重新打开文档。')
    if (field.text === change.text) continue
    let xml: string
    if (s.kind === 'docx')
      xml = field.key.endsWith(':append')
        ? change.text
            .split(/\r\n|\r|\n/)
            .map(
              (line) =>
                `<w:p><w:r><w:t xml:space="preserve">${escape(line.replace(/\t/g, '    '))}</w:t></w:r></w:p>`,
            )
            .join('')
        : wordPatch(field.xml, field.text, change.text)
    else {
      if (change.text.startsWith('=')) throw new Error('当前版本不改写公式，请输入普通文字或数值。')
      const tag = field.xml
        .match(/^<c\b([^>]*?)(?:\/?>)/)![1]
        .replace(/\s+t=["'][^"']*["']/g, '')
        .replace(/\/$/, '')
      const wasText = /\bt=["'](?:s|inlineStr|str)["']/.test(field.xml)
      const numeric =
        !wasText &&
        /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(change.text) &&
        Number.isFinite(Number(change.text))
      const ext = field.xml.match(/<extLst\b[\s\S]*?<\/extLst>/)?.[0] || ''
      const boolean = /\bt=["']b["']/.test(field.xml) && /^(?:true|false|0|1)$/i.test(change.text)
      xml = boolean
        ? `<c${tag} t="b"><v>${/^(?:true|1)$/i.test(change.text) ? '1' : '0'}</v>${ext}</c>`
        : numeric
          ? `<c${tag}><v>${escape(change.text)}</v>${ext}</c>`
          : `<c${tag} t="inlineStr"><is><t xml:space="preserve">${escape(change.text)}</t></is>${ext}</c>`
    }
    const list = replacements.get(field.part) || []
    if (field.start === -1) {
      const cells = insertedCells.get(field.part) || []
      cells.push({ address: field.key.slice(field.part.length + 1), xml })
      insertedCells.set(field.part, cells)
    } else list.push({ field, xml })
    replacements.set(field.part, list)
  }
  if (!replacements.size) return Buffer.from(bytes)
  for (const [part, list] of replacements) {
    let xml = strFromU8(s.files[part])
    for (const { field, xml: replacement } of list.sort((a, b) => b.field.start - a.field.start))
      xml = xml.slice(0, field.start) + replacement + xml.slice(field.end)
    for (const cell of insertedCells.get(part) || [])
      xml = insertSheetCell(xml, cell.address, cell.xml)
    if (insertedCells.has(part)) {
      // Optional cached used-range must not hide newly inserted rows/columns.
      xml = xml.replace(/<dimension\b[^>]*(?:\/>|>[\s\S]*?<\/dimension>)/, '')
    }
    if (XMLValidator.validate(xml) !== true) throw new Error('编辑后内容校验失败，原文件未改变。')
    s.files[part] = strToU8(xml)
  }
  if (s.kind === 'xlsx') {
    // Invalidate cached results, without claiming that this app calculates formulas.
    for (const part of Object.keys(s.files).filter((p) => /^xl\/worksheets\/[^/]+\.xml$/.test(p))) {
      const xml = strFromU8(s.files[part]).replace(/<c(?=[\s>])[^>]*>[\s\S]*?<\/c>/g, (cell) =>
        /<f\b/.test(cell) ? cell.replace(/<v(?:\s[^>]*)?>[\s\S]*?<\/v>|<v\s*\/>/g, '') : cell,
      )
      s.files[part] = strToU8(xml)
    }
    let workbook = strFromU8(s.files['xl/workbook.xml'])
    const calc = /<calcPr\b([^>]*?)(?:\/>|>[\s\S]*?<\/calcPr>)/
    const existing = calc.exec(workbook)
    const attrs = (existing?.[1] || '').replace(
      /\s+(?:fullCalcOnLoad|forceFullCalc)=["'][^"']*["']/g,
      '',
    )
    const updated = `<calcPr${attrs} fullCalcOnLoad="1" forceFullCalc="1"/>`
    if (existing) workbook = workbook.replace(calc, updated)
    else
      workbook = workbook.replace(
        /<(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/workbook>/,
        (token) => updated + token,
      )
    s.files['xl/workbook.xml'] = strToU8(workbook)
  }
  const result = zipSync(s.files)
  parseOffice(result, extension)
  return Buffer.from(result)
}

function cellPoint(address: string): [number, number] | null {
  const m = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(address)
  if (!m) return null
  const point: [number, number] = [
    Number(m[2]),
    [...m[1]].reduce((v, c) => v * 26 + c.charCodeAt(0) - 64, 0),
  ]
  return point[0] <= 1048576 && point[1] <= 16384 ? point : null
}
function inRange(point: [number, number], range: string): boolean {
  const [a, b = a] = range.split(':')
  const start = cellPoint(a),
    end = cellPoint(b)
  return (
    !!start &&
    !!end &&
    point[0] >= start[0] &&
    point[0] <= end[0] &&
    point[1] >= start[1] &&
    point[1] <= end[1]
  )
}
function insertSheetCell(xml: string, address: string, cell: string): string {
  const [row, column] = cellPoint(address)!
  const data = /<sheetData\b[^>]*(?:\/>|>[\s\S]*?<\/sheetData>)/.exec(xml)
  if (!data) throw new Error('当前工作表缺少数据区域，无法新增单元格。')
  let content = data[0].replace(/\/>$/, '></sheetData>')
  const rows = [...content.matchAll(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g)]
  const same = rows.find((m) => Number(m[0].match(/\br=["'](\d+)["']/)?.[1]) === row)
  if (same) {
    let value = same[0].replace(/\/>$/, '></row>')
    const next = [
      ...value.matchAll(/<c\b[^>]*\br=["']([A-Z]+\d+)["'][^>]*(?:\/>|>[\s\S]*?<\/c>)/g),
    ].find((m) => cellPoint(m[1])![1] > column)
    const at = next?.index ?? value.search(/<extLst\b|<\/row>/)
    value = value.slice(0, at) + cell + value.slice(at)
    content = content.slice(0, same.index!) + value + content.slice(same.index! + same[0].length)
  } else {
    const next = rows.find((m) => Number(m[0].match(/\br=["'](\d+)["']/)?.[1]) > row)
    const at = next?.index ?? content.indexOf('</sheetData>')
    content = content.slice(0, at) + `<row r="${row}">${cell}</row>` + content.slice(at)
  }
  return xml.slice(0, data.index) + content + xml.slice(data.index + data[0].length)
}

export function newOfficeFile(extension: '.docx' | '.xlsx'): Buffer {
  const relNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
  const officeRelationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const target = extension === '.docx' ? 'word/document.xml' : 'xl/workbook.xml'
  const parts: Record<string, string> = {
    '_rels/.rels': `<Relationships xmlns="${relNamespace}"><Relationship Id="r1" Type="${officeRelationship}/officeDocument" Target="${target}"/></Relationships>`,
  }
  if (extension === '.docx') {
    parts['[Content_Types].xml'] =
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    parts[target] =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>新建文档</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>'
  } else {
    parts['[Content_Types].xml'] =
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    parts[target] =
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationship}"><sheets><sheet name="工作表1" sheetId="1" r:id="rSheet"/></sheets></workbook>`
    parts['xl/_rels/workbook.xml.rels'] =
      `<Relationships xmlns="${relNamespace}"><Relationship Id="rSheet" Type="${officeRelationship}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`
    parts['xl/worksheets/sheet1.xml'] =
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      Array.from(
        { length: 20 },
        (_, r) =>
          `<row r="${r + 1}">${Array.from({ length: 8 }, (_, c) => `<c r="${String.fromCharCode(65 + c)}${r + 1}"/>`).join('')}</row>`,
      ).join('') +
      '</sheetData></worksheet>'
  }
  return Buffer.from(
    zipSync(Object.fromEntries(Object.entries(parts).map(([path, xml]) => [path, strToU8(xml)]))),
  )
}
