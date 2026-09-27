import { useMemo, useRef, useState } from 'react'
import type { KeyboardEvent } from 'react'
import { ChevronLeft, ChevronRight, Search } from 'lucide-react'
import type { OfficeChange, OfficeEditModel } from '../../shared/types'

const ROWS = 40,
  COLS = 12
function columnName(value: number): string {
  let result = ''
  for (; value > 0; value = Math.floor((value - 1) / 26))
    result = String.fromCharCode(65 + ((value - 1) % 26)) + result
  return result
}
function point(address: string) {
  const m = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(address.toUpperCase())
  if (!m) return null
  const row = Number(m[2]),
    col = [...m[1]].reduce((v, c) => v * 26 + c.charCodeAt(0) - 64, 0)
  return row <= 1048576 && col <= 16384 ? { row, col } : null
}
const addressOf = (row: number, col: number) => `${columnName(col)}${row}`
function range(value: string) {
  const [a, b = a] = value.split(':'),
    start = point(a),
    end = point(b)
  return start && end ? { start, end } : null
}

export function SheetEditor({
  model,
  changes,
  onChange,
  busy,
}: {
  model: OfficeEditModel
  changes: OfficeChange[]
  onChange: (key: string, value: string, original: string) => void
  busy: boolean
}) {
  const sheets = model.sheets || []
  const [sheetIndex, setSheetIndex] = useState(0),
    [selected, setSelected] = useState('A1')
  const [rowPage, setRowPage] = useState(0),
    [colPage, setColPage] = useState(0)
  const [filter, setFilter] = useState(''),
    [hint, setHint] = useState('点击单元格输入；Enter 向下、Tab 向右。公式只读，当前不重新计算。')
  const [focused, setFocused] = useState('')
  const revert = useRef('')
  const grid = useRef<HTMLDivElement>(null)
  const sheet = sheets[Math.min(sheetIndex, sheets.length - 1)]
  const fields = useMemo(() => new Map(model.fields.map((f) => [f.key, f])), [model])
  const changed = useMemo(() => new Map(changes.map((c) => [c.key, c.text])), [changes])
  const cells = useMemo(() => new Map(sheet?.cells.map((c) => [c.address, c])), [sheet])
  const merges = useMemo(() => sheet?.merges.map(range).filter((r) => r !== null) || [], [sheet])
  const protectedRanges = useMemo(
    () => sheet?.protectedRanges.map(range).filter((r) => r !== null) || [],
    [sheet],
  )
  if (!sheet) return <p>这份工作簿没有可以打开的工作表。</p>
  const contains = (r: NonNullable<ReturnType<typeof range>>, row: number, col: number) =>
    row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col
  const keyOf = (address: string) => `${sheet.part}:${address}`
  const original = (address: string) => fields.get(keyOf(address))?.text ?? ''
  const valueOf = (address: string) => changed.get(keyOf(address)) ?? original(address)
  const editable = (address: string) => {
    const p = point(address)!
    return (
      sheet.editable &&
      fields.get(keyOf(address))?.editable !== false &&
      !protectedRanges.some((r) => contains(r, p.row, p.col)) &&
      !merges.some(
        (r) => contains(r, p.row, p.col) && address !== addressOf(r.start.row, r.start.col),
      )
    )
  }
  const edit = (address: string, value: string) => {
    if (busy || !editable(address)) return
    if (value.startsWith('=')) {
      setHint('当前支持文字和数值编辑，尚不支持新增或修改公式。')
      return
    }
    onChange(keyOf(address), value, original(address))
  }
  const go = (row: number, col: number, focus = true) => {
    row = Math.max(1, Math.min(1048576, row))
    col = Math.max(1, Math.min(16384, col))
    const merge = merges.find((r) => contains(r, row, col))
    if (merge) {
      row = merge.start.row
      col = merge.start.col
    }
    const address = addressOf(row, col)
    setSelected(address)
    setRowPage(Math.floor((row - 1) / ROWS))
    setColPage(Math.floor((col - 1) / COLS))
    if (focus)
      requestAnimationFrame(() => {
        const input = grid.current?.querySelector<HTMLTextAreaElement>(
          `[data-address="${address}"]`,
        )
        input?.focus()
        input?.scrollIntoView({ block: 'nearest', inline: 'nearest' })
      })
  }
  const move = (e: KeyboardEvent<HTMLTextAreaElement>, address: string) => {
    if (e.nativeEvent.isComposing) return
    const p = point(address)!,
      merge = merges.find((r) => contains(r, p.row, p.col))
    let row = p.row,
      col = p.col
    if (e.key === 'Enter' && !e.altKey) row = e.shiftKey ? row - 1 : (merge?.end.row ?? row) + 1
    else if (e.key === 'Tab') col = e.shiftKey ? col - 1 : (merge?.end.col ?? col) + 1
    else if (e.key === 'ArrowDown') row = (merge?.end.row ?? row) + 1
    else if (e.key === 'ArrowUp') row--
    else if (e.key === 'Escape') {
      e.preventDefault()
      edit(address, revert.current)
      e.currentTarget.blur()
      return
    } else return
    e.preventDefault()
    go(row, col)
  }
  const startRow = rowPage * ROWS + 1,
    endRow = Math.min(startRow + ROWS - 1, 1048576)
  const startCol = colPage * COLS + 1,
    endCol = Math.min(startCol + COLS - 1, 16384)
  const maxRow = Math.max(
    1,
    ...sheet.cells.map((c) => c.row),
    ...changes
      .filter((c) => c.key.startsWith(`${sheet.part}:`))
      .map((c) => point(c.key.slice(sheet.part.length + 1))?.row || 1),
  )
  const maxCol = Math.max(
    1,
    ...sheet.cells.map((c) => c.column),
    ...changes
      .filter((c) => c.key.startsWith(`${sheet.part}:`))
      .map((c) => point(c.key.slice(sheet.part.length + 1))?.col || 1),
  )
  const find = () => {
    const p = point(filter.trim())
    if (p) {
      go(p.row, p.col)
      setHint(`已定位 ${filter.toUpperCase()}`)
      return
    }
    const query = filter.trim().toLocaleLowerCase()
    const match =
      query &&
      sheet.cells.find((c) =>
        (changed.get(keyOf(c.address)) ?? c.text).toLocaleLowerCase().includes(query),
      )
    const newMatch =
      query &&
      changes.find(
        (c) => c.key.startsWith(`${sheet.part}:`) && c.text.toLocaleLowerCase().includes(query),
      )
    const target = match
      ? { row: match.row, col: match.column }
      : newMatch
        ? point(newMatch.key.slice(sheet.part.length + 1))
        : null
    if (target) {
      go(target.row, target.col)
      setHint('已定位匹配单元格。')
    } else setHint('当前工作表没有找到匹配内容。')
  }
  return (
    <div className="sheet-edit-workspace">
      <div className="office-context-toolbar">
        <form
          className="office-find"
          onSubmit={(e) => {
            e.preventDefault()
            find()
          }}
        >
          <Search size={16} />
          <input
            aria-label="查找编辑位置"
            placeholder="查找内容或输入坐标，如 D120"
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
          />
          <button className="secondary" type="submit">
            定位
          </button>
        </form>
        <span className="office-edit-status">
          {sheet.editable ? '文字与数值编辑' : '此工作表受保护，只读'}
        </span>
      </div>
      <div className="sheet-edit-formula">
        <strong aria-label="当前单元格">{selected}</strong>
        <span aria-hidden="true">fx</span>
        <input
          aria-label="单元格内容"
          value={valueOf(selected)}
          readOnly={busy || !editable(selected)}
          maxLength={32767}
          onChange={(e) => edit(selected, e.target.value)}
          placeholder="选择单元格后，可在这里输入内容"
        />
      </div>
      <div className="sheet-edit-scroll" ref={grid}>
        <table className="sheet-edit-grid" aria-label={`编辑工作表 ${sheet.name}`}>
          <colgroup>
            <col style={{ width: 48 }} />
            {Array.from({ length: endCol - startCol + 1 }, (_, i) => (
              <col key={i} style={{ width: 150 }} />
            ))}
          </colgroup>
          <thead>
            <tr>
              <th aria-label="行号" />
              {Array.from({ length: endCol - startCol + 1 }, (_, i) => (
                <th key={i}>{columnName(startCol + i)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: endRow - startRow + 1 }, (_, i) => {
              const row = startRow + i
              return (
                <tr key={row}>
                  <th>{row}</th>
                  {Array.from({ length: endCol - startCol + 1 }, (_, j) => {
                    const col = startCol + j,
                      merge = merges.find((r) => contains(r, row, col))
                    if (
                      merge &&
                      (row !== Math.max(startRow, merge.start.row) ||
                        col !== Math.max(startCol, merge.start.col))
                    )
                      return null
                    const address = merge
                      ? addressOf(merge.start.row, merge.start.col)
                      : addressOf(row, col)
                    const key = keyOf(address),
                      cell = cells.get(address),
                      allowed = editable(address)
                    const display =
                      focused === address
                        ? valueOf(address)
                        : (changed.get(key) ?? cell?.text ?? '')
                    return (
                      <td
                        key={col}
                        rowSpan={merge ? Math.min(endRow, merge.end.row) - row + 1 : undefined}
                        colSpan={merge ? Math.min(endCol, merge.end.col) - col + 1 : undefined}
                        className={`${selected === address ? 'cell-active' : ''} ${changed.has(key) ? 'cell-changed' : ''} ${!allowed ? 'cell-readonly' : ''}`}
                      >
                        <textarea
                          data-address={address}
                          aria-label={`编辑 ${sheet.name} · ${address}`}
                          title={allowed ? address : fields.get(key)?.reason || '受保护单元格'}
                          readOnly={busy || !allowed}
                          value={display}
                          maxLength={32767}
                          rows={Math.min(4, Math.max(1, display.split('\n').length))}
                          onFocus={() => {
                            setSelected(address)
                            setFocused(address)
                            revert.current = valueOf(address)
                            if (!allowed)
                              setHint(fields.get(key)?.reason || '此处受保护，保留原样。')
                          }}
                          onBlur={() => setFocused('')}
                          onChange={(e) => edit(address, e.target.value)}
                          onKeyDown={(e) => move(e, address)}
                        />
                      </td>
                    )
                  })}
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
      <div className="sheet-edit-footer">
        <div className="reader-tabs" role="tablist" aria-label="工作表">
          {sheets.map((s, i) => (
            <button
              role="tab"
              aria-selected={i === sheetIndex}
              className={i === sheetIndex ? 'active' : ''}
              key={s.part}
              onClick={() => {
                setSheetIndex(i)
                setSelected('A1')
                setRowPage(0)
                setColPage(0)
                setFocused('')
              }}
            >
              {s.name}
              {s.hidden ? '（隐藏）' : ''}
            </button>
          ))}
        </div>
        <div className="sheet-edit-pages">
          <button
            className="icon-button"
            aria-label="上一组行"
            disabled={!rowPage}
            onClick={() => setRowPage(rowPage - 1)}
          >
            <ChevronLeft size={16} />
          </button>
          <span>
            行 {startRow}–{endRow}
          </span>
          <button
            className="icon-button"
            aria-label="下一组行"
            disabled={endRow >= 1048576}
            onClick={() => setRowPage(rowPage + 1)}
          >
            <ChevronRight size={16} />
          </button>
          <button
            className="icon-button"
            aria-label="上一组列"
            disabled={!colPage}
            onClick={() => setColPage(colPage - 1)}
          >
            <ChevronLeft size={16} />
          </button>
          <span>
            列 {columnName(startCol)}–{columnName(endCol)}
          </span>
          <button
            className="icon-button"
            aria-label="下一组列"
            disabled={endCol >= 16384}
            onClick={() => setColPage(colPage + 1)}
          >
            <ChevronRight size={16} />
          </button>
        </div>
      </div>
      <div className="office-edit-status" role="status">
        {hint}
        <span>已有区域至 {addressOf(maxRow, maxCol)} · 未保存的格子以蓝色底色标出</span>
      </div>
    </div>
  )
}
