import { useEffect, useRef, useState } from 'react'
import { ChevronLeft, ChevronRight, FileSearch, LoaderCircle } from 'lucide-react'
import type {
  ContentLocation,
  OfficeData,
  OfficePreview,
  OfficeSheet,
  SheetCell,
} from '../../shared/types'
import './office.css'

export function locationLabel(location?: ContentLocation): string {
  if (!location) return ''
  return location.kind === 'cell'
    ? `${location.sheet} · ${location.address}`
    : `${location.part} · 段落 ${location.index + 1}`
}

export function Highlight({ text, query = '' }: { text: string; query?: string }) {
  const index = query ? text.toLocaleLowerCase().indexOf(query.toLocaleLowerCase()) : -1
  if (index < 0) return <>{text}</>
  return (
    <>
      {text.slice(0, index)}
      <mark>{text.slice(index, index + query.length)}</mark>
      {text.slice(index + query.length)}
    </>
  )
}

function WordLayout({ bytes, compact = false }: { bytes: Uint8Array; compact?: boolean }) {
  const host = useRef<HTMLDivElement>(null)
  const [state, setState] = useState('正在生成文档预览…')
  useEffect(() => {
    let cancelled = false
    const root = host.current!
    const shadow = root.shadowRoot ?? root.attachShadow({ mode: 'open' })
    const target = document.createElement('div')
    shadow.replaceChildren(target)
    const blockNavigation = (event: Event) => event.preventDefault()
    target.addEventListener('click', blockNavigation)
    void import('docx-preview')
      .then(({ renderAsync }) =>
        renderAsync(bytes, target, target, {
          className: 'word-page',
          inWrapper: true,
          ignoreFonts: true,
          renderAltChunks: false,
          renderChanges: false,
          renderComments: false,
          renderHeaders: true,
          renderFooters: true,
          useBase64URL: true,
          ignoreWidth: compact,
          ignoreHeight: compact,
        }),
      )
      .then(() => {
        if (cancelled) return
        // Links are informational in the reader. Imported documents never navigate the app.
        target.querySelectorAll('a').forEach((anchor) => anchor.removeAttribute('href'))
        const style = document.createElement('style')
        style.textContent = compact
          ? ':host{color-scheme:light;color:#18233f}.word-page-wrapper{background:white!important;padding:0!important}section.word-page{width:100%!important;min-height:650px;padding:40px 26px!important;box-sizing:border-box;box-shadow:none!important;overflow-wrap:anywhere;margin:0!important}img{max-width:100%}article{width:100%!important}'
          : ':host{color-scheme:light;color:#18233f}.word-page-wrapper{background:var(--surface-muted,#edf2f8)!important;padding:24px!important}section.word-page{max-width:100%;box-sizing:border-box;overflow-wrap:anywhere}img{max-width:100%}'
        target.append(style)
        target.querySelectorAll('table').forEach((table) => {
          const scroller = document.createElement('div')
          scroller.style.cssText = 'overflow-x:auto;max-width:100%;margin:8px 0'
          scroller.tabIndex = 0
          scroller.setAttribute('role', 'region')
          scroller.setAttribute('aria-label', 'Word 表格')
          table.before(scroller)
          scroller.append(table)
        })
        setState('')
      })
      .catch(() => {
        if (!cancelled) setState('版式预览失败，请切换到“正文”。原文件仍然完整保留。')
      })
    return () => {
      cancelled = true
      target.removeEventListener('click', blockNavigation)
      target.remove()
    }
  }, [bytes, compact])
  return (
    <>
      <div className="word-layout-status" role="status">
        {state}
      </div>
      <div ref={host} className="word-layout" aria-label="Word 版式预览" />
    </>
  )
}

function WordText({
  data,
  query,
  location,
}: {
  data: OfficeData
  query: string
  location?: ContentLocation
}) {
  const initial = data.blocks.findIndex(
    (block) =>
      block.location.kind === 'paragraph' &&
      location?.kind === 'paragraph' &&
      block.location.part === location.part &&
      block.location.index === location.index,
  )
  const [page, setPage] = useState(Math.floor(Math.max(0, initial) / 100))
  const scroll = useRef<HTMLDivElement>(null)
  useEffect(() => {
    scroll.current?.querySelector('.located')?.scrollIntoView({ block: 'center' })
  }, [page])
  return (
    <>
      <div className="reader-pagination">
        <span>{data.blocks.length} 段可检索文本</span>
        <button
          className="icon-button"
          aria-label="上一页正文"
          disabled={!page}
          onClick={() => setPage(page - 1)}
        >
          <ChevronLeft size={16} />
        </button>
        <span>
          {page + 1} / {Math.max(1, Math.ceil(data.blocks.length / 100))}
        </span>
        <button
          className="icon-button"
          aria-label="下一页正文"
          disabled={(page + 1) * 100 >= data.blocks.length}
          onClick={() => setPage(page + 1)}
        >
          <ChevronRight size={16} />
        </button>
      </div>
      <div className="word-text" ref={scroll}>
        {data.blocks.slice(page * 100, (page + 1) * 100).map((block, index) => (
          <section key={index} className={initial === page * 100 + index ? 'located' : ''}>
            <small>{locationLabel(block.location)}</small>
            <p
              data-mark-scope={JSON.stringify(block.location)}
              data-location={JSON.stringify(block.location)}
            >
              <Highlight text={block.text} query={query} />
            </p>
          </section>
        ))}
        {!data.blocks.length && <p>没有可提取的文本，试试版式预览。</p>}
      </div>
    </>
  )
}

const columnName = (index: number): string => {
  let text = ''
  for (let value = index; value > 0; value = Math.floor((value - 1) / 26))
    text = String.fromCharCode(65 + ((value - 1) % 26)) + text
  return text
}

function SheetGrid({
  sheet,
  location,
  query,
}: {
  sheet: OfficeSheet
  location?: ContentLocation
  query: string
}) {
  const rows = [...new Set(sheet.cells.map((cell) => cell.row))]
  const columns = [...new Set(sheet.cells.map((cell) => cell.column))].sort((a, b) => a - b)
  const target =
    location?.kind === 'cell' && location.sheet === sheet.name
      ? sheet.cells.find((cell) => cell.address === location.address)
      : undefined
  const [rowPage, setRowPage] = useState(
    Math.floor(Math.max(0, rows.indexOf(target?.row ?? 0)) / 50),
  )
  const [columnPage, setColumnPage] = useState(
    Math.floor(Math.max(0, columns.indexOf(target?.column ?? 0)) / 12),
  )
  const [selected, setSelected] = useState<SheetCell | undefined>(target ?? sheet.cells[0])
  const shownRows = rows.slice(rowPage * 50, (rowPage + 1) * 50)
  const shownColumns = columns.slice(columnPage * 12, (columnPage + 1) * 12)
  const byAddress = new Map(sheet.cells.map((cell) => [cell.address, cell]))
  const scroll = useRef<HTMLDivElement>(null)
  useEffect(() => {
    scroll.current?.querySelector('.located')?.scrollIntoView({ block: 'center', inline: 'center' })
  }, [])
  return (
    <>
      <div className="formula-bar">
        <strong>{selected?.address ?? '—'}</strong>
        <span>
          {selected?.formula !== undefined
            ? `=${selected.formula || '（共享公式）'}`
            : selected
              ? (selected.rawValue ?? selected.text)
              : '选择单元格查看内容'}
        </span>
      </div>
      <div className="sheet-grid-scroll" ref={scroll}>
        <table className="sheet-grid" aria-label={`${sheet.name} 工作表`}>
          <thead>
            <tr>
              <th aria-label="行号" />
              {shownColumns.map((column) => (
                <th key={column}>{columnName(column)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {shownRows.map((row) => (
              <tr key={row}>
                <th>{row}</th>
                {shownColumns.map((column) => {
                  const address = `${columnName(column)}${row}`
                  const cell = byAddress.get(address)
                  return (
                    <td
                      key={column}
                      className={`${cell && cell === target ? 'located' : ''} ${selected?.address === address ? 'selected-cell' : ''}`}
                    >
                      <button
                        data-cell-mark="true"
                        data-mark-scope={JSON.stringify({
                          kind: 'cell',
                          sheet: sheet.name,
                          address,
                        })}
                        data-location={JSON.stringify({ kind: 'cell', sheet: sheet.name, address })}
                        aria-label={`${address} ${cell?.text ?? '空白'}`}
                        onClick={() => setSelected(cell ?? { address, row, column, text: '' })}
                        title={cell?.formula ? `=${cell.formula}` : cell?.text}
                      >
                        <Highlight text={cell?.text || ''} query={query} />
                      </button>
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
        {!sheet.cells.length && <p className="reader-empty">此工作表没有数据。</p>}
      </div>
      <div className="reader-pagination">
        <span>按有内容的行列分页 · 原始坐标</span>
        <button
          className="icon-button"
          aria-label="上一组行"
          disabled={!rowPage}
          onClick={() => setRowPage(rowPage - 1)}
        >
          <ChevronLeft size={16} />
        </button>
        <span>
          行 {rowPage + 1}/{Math.max(1, Math.ceil(rows.length / 50))}
        </span>
        <button
          className="icon-button"
          aria-label="下一组行"
          disabled={(rowPage + 1) * 50 >= rows.length}
          onClick={() => setRowPage(rowPage + 1)}
        >
          <ChevronRight size={16} />
        </button>
        <button
          className="icon-button"
          aria-label="上一组列"
          disabled={!columnPage}
          onClick={() => setColumnPage(columnPage - 1)}
        >
          <ChevronLeft size={16} />
        </button>
        <span>
          列 {columnPage + 1}/{Math.max(1, Math.ceil(columns.length / 12))}
        </span>
        <button
          className="icon-button"
          aria-label="下一组列"
          disabled={(columnPage + 1) * 12 >= columns.length}
          onClick={() => setColumnPage(columnPage + 1)}
        >
          <ChevronRight size={16} />
        </button>
      </div>
      {sheet.merges.length > 0 && (
        <p className="merge-note">
          合并区域：{sheet.merges.slice(0, 10).join('、')}
          {sheet.merges.length > 10 ? '…' : ''}。当前按原始单元格显示。
        </p>
      )}
    </>
  )
}

export function OfficeReader({
  id,
  revision,
  compact = false,
  query = '',
  location,
}: {
  id: string
  revision: number
  compact?: boolean
  query?: string
  location?: ContentLocation
}) {
  const [preview, setPreview] = useState<OfficePreview | null>(null)
  const [mode, setMode] = useState<'layout' | 'text'>(
    location?.kind === 'paragraph' ? 'text' : 'layout',
  )
  const [sheetIndex, setSheetIndex] = useState(0)
  useEffect(() => {
    let cancelled = false
    setPreview(null)
    void window
      .localDocs!.readOffice(id, true)
      .then((value) => {
        if (cancelled) return
        setPreview(value)
        setMode(location?.kind === 'paragraph' ? 'text' : 'layout')
        setSheetIndex(
          Math.max(
            0,
            value.data?.sheets.findIndex(
              (sheet) => location?.kind === 'cell' && sheet.name === location.sheet,
            ) ?? 0,
          ),
        )
      })
      .catch((error) => {
        if (!cancelled) setPreview({ data: null, error: String(error) })
      })
    return () => {
      cancelled = true
    }
  }, [id, revision, location, compact])
  if (!preview)
    return (
      <div className="office-loading" role="status">
        <LoaderCircle size={23} />
        <p>正在读取文档内容…</p>
      </div>
    )
  if (!preview.data)
    return (
      <div className="office-read-error">
        <FileSearch size={26} />
        <strong>暂时无法预览</strong>
        <p>{preview.error}</p>
        <small>文件副本已保留，可以原样导出。</small>
      </div>
    )
  const data = preview.data
  if (compact && data.kind === 'docx' && preview.bytes)
    return <WordLayout bytes={preview.bytes} compact />
  return (
    <section
      className={`office-reader ${compact ? 'office-compact' : ''}`}
      aria-label={data.kind === 'docx' ? 'Word 阅读器' : 'Excel 阅读器'}
    >
      <div className="reader-toolbar">
        <span className="reader-badge">阅读视图</span>
        {data.kind === 'docx' ? (
          <div className="reader-tabs">
            <button className={mode === 'layout' ? 'active' : ''} onClick={() => setMode('layout')}>
              版式
            </button>
            <button className={mode === 'text' ? 'active' : ''} onClick={() => setMode('text')}>
              正文
            </button>
          </div>
        ) : (
          <div className="reader-tabs" role="tablist" aria-label="工作表">
            {data.sheets.map((sheet, index) => (
              <button
                key={index}
                role="tab"
                aria-selected={sheetIndex === index}
                className={sheetIndex === index ? 'active' : ''}
                onClick={() => setSheetIndex(index)}
              >
                {sheet.name}
                {sheet.hidden ? '（隐藏表）' : ''}
              </button>
            ))}
          </div>
        )}
      </div>
      <p className="reader-notice">{data.warnings.join(' ')}</p>
      {data.kind === 'docx' ? (
        mode === 'layout' && preview.bytes ? (
          <div className="word-layout-scroll">
            <WordLayout bytes={preview.bytes} />
          </div>
        ) : (
          <WordText key={`${id}:${revision}`} data={data} location={location} query={query} />
        )
      ) : (
        <SheetGrid
          key={`${id}:${revision}:${sheetIndex}`}
          sheet={data.sheets[sheetIndex]}
          location={location}
          query={query}
        />
      )}
    </section>
  )
}
