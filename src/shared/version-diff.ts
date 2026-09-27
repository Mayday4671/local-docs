import type { VersionContent } from './types'

export interface DiffValue {
  label: string
  text: string
}
export interface DiffRow {
  before?: DiffValue
  after?: DiffValue
  kind: 'same' | 'added' | 'removed' | 'modified'
}
export interface VersionDiff {
  rows: DiffRow[]
  counts: Record<DiffRow['kind'], number>
  identical: boolean
  warnings: string[]
}
const row = (before?: DiffValue, after?: DiffValue): DiffRow => ({
  before,
  after,
  kind: !before ? 'added' : !after ? 'removed' : before.text === after.text ? 'same' : 'modified',
})
const partName = (part: string) =>
  part === 'word/document.xml'
    ? '正文'
    : part.includes('header')
      ? `页眉 ${part.match(/\d+/)?.[0] ?? ''}`
      : part.includes('footer')
        ? `页脚 ${part.match(/\d+/)?.[0] ?? ''}`
        : part.includes('footnotes')
          ? '脚注'
          : part.includes('endnotes')
            ? '尾注'
            : part

// Bounded LCS handles repeated lines; unique anchors keep large documents responsive.
function align(
  before: DiffValue[],
  after: DiffValue[],
  budget: { cells: number; approximate: boolean },
): DiffRow[] {
  const rows: DiffRow[] = []
  const visit = (a: number, endA: number, b: number, endB: number, depth = 0) => {
    while (a < endA && b < endB && before[a].text === after[b].text)
      rows.push(row(before[a++], after[b++]))
    let suffix = 0
    while (a < endA && b < endB && before[endA - 1].text === after[endB - 1].text) {
      endA--
      endB--
      suffix++
    }
    const n = endA - a,
      m = endB - b
    const pair = (fromA: number, toA: number, fromB: number, toB: number) => {
      for (let i = 0; i < Math.max(toA - fromA, toB - fromB); i++)
        rows.push(
          row(
            fromA + i < toA ? before[fromA + i] : undefined,
            fromB + i < toB ? after[fromB + i] : undefined,
          ),
        )
    }
    if (!n || !m) pair(a, endA, b, endB)
    else if ((n + 1) * (m + 1) <= budget.cells) {
      budget.cells -= (n + 1) * (m + 1)
      const width = m + 1,
        lengths = new Uint32Array((n + 1) * width)
      for (let i = n - 1; i >= 0; i--)
        for (let j = m - 1; j >= 0; j--)
          lengths[i * width + j] =
            before[a + i].text === after[b + j].text
              ? lengths[(i + 1) * width + j + 1] + 1
              : Math.max(lengths[(i + 1) * width + j], lengths[i * width + j + 1])
      let i = 0,
        j = 0,
        startA = 0,
        startB = 0
      while (i < n && j < m) {
        if (before[a + i].text === after[b + j].text) {
          pair(a + startA, a + i, b + startB, b + j)
          rows.push(row(before[a + i++], after[b + j++]))
          startA = i
          startB = j
        } else if (lengths[(i + 1) * width + j] >= lengths[i * width + j + 1]) i++
        else j++
      }
      pair(a + startA, endA, b + startB, endB)
    } else {
      const unique = (items: DiffValue[], start: number, end: number) => {
        const map = new Map<string, number>()
        for (let i = start; i < end; i++) map.set(items[i].text, map.has(items[i].text) ? -1 : i)
        return map
      }
      const left = unique(before, a, endA),
        right = unique(after, b, endB)
      const candidates: [number, number][] = []
      for (const [text, index] of left) {
        const other = right.get(text)
        if (index >= 0 && other !== undefined && other >= 0) candidates.push([index, other])
      }
      const tails: number[] = [],
        previous = new Int32Array(candidates.length).fill(-1)
      candidates.forEach((candidate, index) => {
        let lo = 0,
          hi = tails.length
        while (lo < hi) {
          const mid = (lo + hi) >>> 1
          if (candidates[tails[mid]][1] < candidate[1]) lo = mid + 1
          else hi = mid
        }
        if (lo) previous[index] = tails[lo - 1]
        tails[lo] = index
      })
      const anchors: [number, number][] = []
      for (let index = tails.at(-1) ?? -1; index >= 0; index = previous[index])
        anchors.push(candidates[index])
      anchors.reverse()
      if (anchors.length && depth < 20) {
        for (const [x, y] of anchors) {
          visit(a, x, b, y, depth + 1)
          rows.push(row(before[x], after[y]))
          a = x + 1
          b = y + 1
        }
        visit(a, endA, b, endB, depth + 1)
      } else {
        budget.approximate = true
        pair(a, endA, b, endB)
      }
    }
    for (let i = 0; i < suffix; i++) rows.push(row(before[endA + i], after[endB + i]))
  }
  visit(0, before.length, 0, after.length)
  return rows
}

export function compareVersions(before: VersionContent, after: VersionContent): VersionDiff {
  if (before.extension !== after.extension) throw new Error('请选择同一类型的文件版本。')
  const warnings = new Set([...(before.office?.warnings ?? []), ...(after.office?.warnings ?? [])])
  const budget = { cells: 1000000, approximate: false }
  let rows: DiffRow[] = []
  if (before.text !== null && after.text !== null) {
    const lines = (text: string) => (text === '' ? [] : text.replace(/\r\n?/g, '\n').split('\n'))
    const a = lines(before.text),
      b = lines(after.text)
    if (a.length + b.length > 100000)
      throw new Error('文档行数过多，当前支持对比合计 10 万行以内的版本。')
    const label = (text: string, i: number) => ({ label: `第 ${i + 1} 行`, text })
    rows = align(a.map(label), b.map(label), budget)
  } else if (before.office?.kind === 'docx' && after.office?.kind === 'docx') {
    const groups = (version: VersionContent) => {
      const map = new Map<string, DiffValue[]>()
      for (const block of version.office!.blocks)
        if (block.location.kind === 'paragraph') {
          const { part, index } = block.location
          if (!map.has(part)) map.set(part, [])
          map.get(part)!.push({ label: `${partName(part)} · 段落 ${index + 1}`, text: block.text })
        }
      return map
    }
    const a = groups(before),
      b = groups(after)
    for (const part of new Set([...a.keys(), ...b.keys()]))
      rows.push(...align(a.get(part) ?? [], b.get(part) ?? [], budget))
  } else if (before.office?.kind === 'xlsx' && after.office?.kind === 'xlsx') {
    const a = new Map(before.office.sheets.map((sheet) => [sheet.name, sheet]))
    const b = new Map(after.office.sheets.map((sheet) => [sheet.name, sheet]))
    const metadata = (sheet: NonNullable<ReturnType<typeof a.get>>, order: number): DiffValue => ({
      label: `${sheet.name} · 工作表`,
      text: `顺序：${order + 1}；${sheet.hidden ? '隐藏' : '可见'}；合并区域：${[...sheet.merges].sort().join('、') || '无'}`,
    })
    for (const name of new Set([...a.keys(), ...b.keys()])) {
      const left = a.get(name),
        right = b.get(name)
      rows.push(
        row(
          left && metadata(left, [...a.keys()].indexOf(name)),
          right && metadata(right, [...b.keys()].indexOf(name)),
        ),
      )
      const x = new Map(left?.cells.map((cell) => [cell.address, cell]) ?? [])
      const y = new Map(right?.cells.map((cell) => [cell.address, cell]) ?? [])
      const addresses = [...new Set([...x.keys(), ...y.keys()])].sort((c, d) => {
        const first = x.get(c) ?? y.get(c)!,
          second = x.get(d) ?? y.get(d)!
        return first.row - second.row || first.column - second.column
      })
      const value = (cell: ReturnType<typeof x.get>): DiffValue | undefined =>
        cell && {
          label: `${name} · ${cell.address}`,
          text: `${cell.text}${cell.rawValue !== undefined ? `\n原始值：${cell.rawValue}` : ''}${cell.formula !== undefined ? `\n公式：=${cell.formula}` : ''}`,
        }
      for (const address of addresses) rows.push(row(value(x.get(address)), value(y.get(address))))
    }
  } else throw new Error('此版本没有可供比较的内容。')
  if (budget.approximate)
    warnings.add('部分大段内容变化过多，已按位置配对；新增、删除和修改数量仅供参考。')
  const counts = { same: 0, added: 0, removed: 0, modified: 0 }
  for (const item of rows) counts[item.kind]++
  return { rows, counts, identical: before.hash === after.hash, warnings: [...warnings] }
}
