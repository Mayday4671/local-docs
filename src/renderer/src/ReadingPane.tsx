import { useEffect, useRef, useState } from 'react'
import { Highlighter, Search } from 'lucide-react'
import type {
  Annotation,
  AnnotationInput,
  Attachment,
  ContentLocation,
  DocumentRecord,
} from '../../shared/types'
import { fileKind, fileTypeName } from '../../shared/file-types'
import { FileReader } from './FileReader'
import { Markdown } from './Markdown'
import { OfficeReader, locationLabel } from './OfficeReader'

function resolveOffset(text: string, a: AnnotationInput): number {
  const candidates: number[] = []
  for (let at = text.indexOf(a.quote); at >= 0; at = text.indexOf(a.quote, at + 1))
    candidates.push(at)
  if (candidates.length === 1) return candidates[0]
  const exact = candidates.filter(
    (at) =>
      text.slice(Math.max(0, at - a.prefix.length), at) === a.prefix &&
      text.slice(at + a.quote.length, at + a.quote.length + a.suffix.length) === a.suffix,
  )
  if (exact.length === 1) return exact[0]
  return -1 // Repeated or modified text must never silently jump to the wrong passage.
}
function rangeFor(element: Element, start: number, length: number): Range | null {
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT)
  let node: Node | null,
    offset = 0,
    range: Range | null = null
  while ((node = walker.nextNode())) {
    const end = offset + (node.textContent?.length || 0)
    if (!range && start < end) {
      range = document.createRange()
      range.setStart(node, Math.max(0, start - offset))
    }
    if (range && start + length <= end) {
      range.setEnd(node, start + length - offset)
      return range
    }
    offset = end
  }
  return null
}
export function ReadingPane({
  doc,
  text,
  query = '',
  location,
  compact = false,
  onChanged,
  disabled = false,
  attachments,
  showMarks = false,
}: {
  doc: DocumentRecord
  text: string | null
  query?: string
  location?: ContentLocation
  compact?: boolean
  onChanged?: () => void
  disabled?: boolean
  attachments?: Attachment[]
  showMarks?: boolean
}) {
  const root = useRef<HTMLDivElement>(null)
  const [marks, setMarks] = useState<Annotation[]>([]),
    [images, setImages] = useState<Attachment[]>([])
  const [selection, setSelection] = useState<AnnotationInput | null>(null),
    [editing, setEditing] = useState<AnnotationInput | null>(null)
  const [panel, setPanel] = useState(showMarks),
    [error, setError] = useState(''),
    [busy, setBusy] = useState(false)
  const [jump, setJump] = useState<Annotation | null>(null),
    [find, setFind] = useState(''),
    [findIndex, setFindIndex] = useState(0),
    [found, setFound] = useState(0)
  const [paint, setPaint] = useState(0)
  const markdown = ['.md', '.markdown'].includes(doc.extension)
  const kind = fileKind(doc.extension)
  const textLike = markdown || kind === 'text'
  const canMark = textLike || kind === 'office'
  useEffect(() => {
    if (textLike) {
      setFind(query)
      setFindIndex(0)
      setJump(null)
    }
  }, [query, textLike, doc.id])
  const highlightId = useRef(`reading-${Math.random().toString(36).slice(2)}`).current
  useEffect(() => {
    if (showMarks) setPanel(true)
  }, [showMarks])
  useEffect(() => {
    let cancelled = false
    void Promise.all([window.localDocs!.annotations(doc.id), window.localDocs!.attachments(doc.id)])
      .then(([m, a]) => {
        if (!cancelled) {
          setMarks(m)
          setImages(a)
        }
      })
      .catch((e) => {
        if (!cancelled) setError(String(e))
      })
    setSelection(null)
    setEditing(null)
    setJump(null)
    return () => {
      cancelled = true
    }
  }, [doc.id, doc.revision])
  useEffect(() => {
    let timer: ReturnType<typeof setTimeout>
    const observer = new MutationObserver(() => {
      clearTimeout(timer)
      timer = setTimeout(() => setPaint((v) => v + 1), 50)
    })
    observer.observe(root.current!, { childList: true, subtree: true, characterData: true })
    return () => {
      clearTimeout(timer)
      observer.disconnect()
    }
  }, [])
  useEffect(() => {
    const registry = (CSS as unknown as { highlights: Map<string, unknown> }).highlights
    const HighlightClass = (window as unknown as { Highlight: new (...ranges: Range[]) => unknown })
      .Highlight
    if (!registry || !HighlightClass) return
    const groups: Record<string, Range[]> = { yellow: [], green: [], blue: [], pink: [], find: [] }
    let jumpRange: Range | null = null
    const scopes = [...root.current!.querySelectorAll<HTMLElement>('[data-mark-scope]')]
    for (const a of marks) {
      const scope = scopes.find((s) => s.dataset.markScope === a.scope)
      if (!scope) continue
      const offset = resolveOffset(scope.textContent || '', a)
      const range = offset < 0 ? null : rangeFor(scope, offset, a.quote.length)
      if (range) {
        groups[a.color].push(range)
        if (a.id === jump?.id) jumpRange = range
      } else if (scope && a.id === jump?.id)
        setError('原文已变化或存在重复片段，暂时无法准确定位；标记内容仍保留。')
    }
    if (find.trim())
      for (const scope of scopes) {
        const content = (scope.textContent || '').toLocaleLowerCase(),
          needle = find.toLocaleLowerCase()
        for (
          let at = content.indexOf(needle);
          at >= 0 && groups.find.length < 1000;
          at = content.indexOf(needle, at + needle.length)
        ) {
          const range = rangeFor(scope, at, needle.length)
          if (range) groups.find.push(range)
        }
      }
    setFound(groups.find.length)
    const target = jumpRange ?? groups.find[findIndex % (groups.find.length || 1)]
    target?.startContainer.parentElement?.scrollIntoView({ block: 'center', inline: 'nearest' })
    const style = document.createElement('style')
    style.textContent = Object.entries({
      yellow: '#ffe083',
      green: '#9ee3bc',
      blue: '#a9d4ff',
      pink: '#ffb8d4',
      find: '#ffa857',
    })
      .map(
        ([color, value]) =>
          `::highlight(${highlightId}-${color}){background:${value};color:#162035}`,
      )
      .join('')
    document.head.append(style)
    for (const [color, ranges] of Object.entries(groups))
      registry.set(`${highlightId}-${color}`, new HighlightClass(...ranges))
    return () => {
      style.remove()
      for (const color of Object.keys(groups)) registry.delete(`${highlightId}-${color}`)
    }
  }, [marks, paint, jump, find, findIndex, highlightId])
  const capture = () => {
    const selected = window.getSelection()
    if (!selected || selected.isCollapsed || !selected.rangeCount) return
    const range = selected.getRangeAt(0),
      parent =
        range.startContainer.nodeType === Node.TEXT_NODE
          ? range.startContainer.parentElement
          : (range.startContainer as Element)
    const scope = parent?.closest<HTMLElement>('[data-mark-scope]')
    if (!scope || !root.current?.contains(scope) || !scope.contains(range.endContainer)) return
    const quote = selected.toString()
    if (!quote.trim() || quote.length > 8000) return
    const preceding = document.createRange()
    preceding.selectNodeContents(scope)
    preceding.setEnd(range.startContainer, range.startOffset)
    const offset = preceding.toString().length,
      all = scope.textContent || ''
    setSelection({
      quote,
      offset,
      prefix: all.slice(Math.max(0, offset - 80), offset),
      suffix: all.slice(offset + quote.length, offset + quote.length + 80),
      scope: scope.dataset.markScope!,
      location: scope.dataset.location ? JSON.parse(scope.dataset.location) : undefined,
      color: 'yellow',
      note: '',
    })
  }
  const save = async () => {
    if (!editing || busy || disabled || doc.deletedAt) return
    setBusy(true)
    setError('')
    try {
      await window.localDocs!.saveAnnotation(doc.id, editing, doc.revision)
      setMarks(await window.localDocs!.annotations(doc.id))
      setEditing(null)
      setSelection(null)
      setPanel(true)
      onChanged?.()
    } catch (e) {
      setError(String(e))
    } finally {
      setBusy(false)
    }
  }
  return (
    <div className={`reading-pane ${compact ? 'compact' : ''}`}>
      {canMark && (
        <div className="reading-tools">
          <button
            className="secondary"
            disabled={disabled || !selection || !!doc.deletedAt}
            title={disabled ? '请先保存文档修改再添加标记' : '选中文字或单元格后标记'}
            onClick={() => setEditing(selection)}
          >
            <Highlighter size={15} />
            标记所选内容
          </button>
          <button className="secondary" aria-pressed={panel} onClick={() => setPanel(!panel)}>
            标记 {marks.length}
          </button>
          <label className="reader-find">
            <Search size={14} />
            <input
              aria-label="文内查找"
              placeholder="查找当前页"
              value={find}
              onChange={(e) => {
                setFind(e.target.value)
                setFindIndex(0)
                setJump(null)
              }}
            />
          </label>
          {find && (
            <button
              className="secondary"
              disabled={!found}
              onClick={() => setFindIndex((v) => v + 1)}
            >
              {found ? `${(findIndex % found) + 1}/${found}` : '0'} 下一处
            </button>
          )}
        </div>
      )}
      {doc.extension === '.docx' && (
        <p className="reading-hint">标记文字或查找时，请切换到“正文”。</p>
      )}
      {error && (
        <p className="annotation-error" role="alert">
          {error}
        </p>
      )}
      {editing && (
        <form
          className="mark-editor"
          onSubmit={(e) => {
            e.preventDefault()
            void save()
          }}
        >
          <blockquote>{editing.quote}</blockquote>
          <label>
            标记颜色
            <select
              aria-label="标记颜色"
              value={editing.color}
              onChange={(e) =>
                setEditing({ ...editing, color: e.target.value as AnnotationInput['color'] })
              }
            >
              <option value="yellow">黄色</option>
              <option value="green">绿色</option>
              <option value="blue">蓝色</option>
              <option value="pink">粉色</option>
            </select>
          </label>
          <textarea
            aria-label="标记备注"
            placeholder="添加备注（可选）"
            maxLength={5000}
            value={editing.note}
            onChange={(e) => setEditing({ ...editing, note: e.target.value })}
          />
          <div>
            <button className="primary" disabled={busy || disabled || !!doc.deletedAt}>
              保存标记
            </button>
            <button
              type="button"
              className="secondary"
              disabled={busy}
              onClick={() => setEditing(null)}
            >
              取消
            </button>
          </div>
        </form>
      )}
      {panel && (
        <aside className="annotation-list" aria-label="内容标记">
          {!marks.length && <p>选中文字或单元格后，点击“标记所选内容”。</p>}
          {marks.map((a) => (
            <article
              key={a.id}
              style={{
                borderLeftColor: {
                  yellow: '#e4b94b',
                  green: '#52ac79',
                  blue: '#558fda',
                  pink: '#da77a0',
                }[a.color],
              }}
            >
              <button
                className="mark-quote"
                onClick={() => {
                  setJump({ ...a })
                  setError('')
                  setFind('')
                  if (textLike) {
                    const scope = root.current!.querySelector('[data-mark-scope="markdown"]')
                    if (scope && resolveOffset(scope.textContent || '', a) < 0)
                      setError('原文已变化或存在重复片段，暂时无法准确定位；标记内容仍保留。')
                  }
                }}
              >
                {a.quote}
              </button>
              <small>
                {locationLabel(a.location) || `${fileTypeName(doc.extension)} 正文`}
                {a.revision !== doc.revision ? ' · 标记后文档有更新' : ''}
              </small>
              {a.note && <p>{a.note}</p>}
              <div>
                <button
                  disabled={disabled || !!doc.deletedAt || busy}
                  onClick={() => setEditing(a)}
                >
                  修改备注 / 颜色
                </button>
                <button
                  disabled={disabled || !!doc.deletedAt || busy}
                  onClick={async () => {
                    setBusy(true)
                    try {
                      await window.localDocs!.removeAnnotation(doc.id, a.id)
                      setMarks(await window.localDocs!.annotations(doc.id))
                      onChanged?.()
                    } catch (e) {
                      setError(String(e))
                    } finally {
                      setBusy(false)
                    }
                  }}
                >
                  删除标记
                </button>
              </div>
            </article>
          ))}
        </aside>
      )}
      <div
        ref={root}
        className="reading-content"
        onMouseUp={capture}
        onKeyUp={capture}
        onClick={(e) => {
          const scope = (e.target as Element).closest<HTMLElement>('[data-cell-mark]')
          if (scope && root.current?.contains(scope)) {
            const quote = scope.textContent || ''
            if (quote.trim())
              setSelection({
                quote,
                offset: 0,
                prefix: '',
                suffix: '',
                scope: scope.dataset.markScope!,
                location: JSON.parse(scope.dataset.location!),
                color: 'yellow',
                note: '',
              })
          }
        }}
      >
        {markdown ? (
          <Markdown text={text || ''} documentId={doc.id} attachments={attachments ?? images} />
        ) : kind === 'office' ? (
          <OfficeReader
            key={`${doc.id}:${jump?.id || ''}`}
            id={doc.id}
            revision={doc.revision}
            compact={compact}
            query={query}
            location={jump?.location || location}
          />
        ) : (
          <FileReader doc={doc} text={text} />
        )}
      </div>
    </div>
  )
}
