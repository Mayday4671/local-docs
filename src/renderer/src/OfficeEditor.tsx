import { useEffect, useMemo, useRef, useState } from 'react'
import type {
  DocumentRecord,
  OfficeChange,
  OfficeEditModel,
  ContentLocation,
} from '../../shared/types'
import { ReadingPane } from './ReadingPane'
import { WordEditor } from './WordEditor'
import { SheetEditor } from './SheetEditor'

export function OfficeEditor({
  doc,
  text,
  onChange,
  busy,
  onChanged,
  query = '',
  location,
  draftStatus = '',
  initialMode = 'read',
}: {
  doc: DocumentRecord
  text: string
  onChange: (text: string) => void
  busy: boolean
  onChanged: () => void
  query?: string
  location?: ContentLocation
  draftStatus?: string
  initialMode?: 'read' | 'edit'
}) {
  const [model, setModel] = useState<OfficeEditModel | null>(null)
  const [error, setError] = useState('')
  const [mode, setMode] = useState<'read' | 'edit'>(initialMode)
  const [hasEdited, setHasEdited] = useState(initialMode === 'edit')
  useEffect(() => {
    let cancelled = false
    setModel(null)
    setError('')
    void window
      .localDocs!.officeEdit(doc.id)
      .then((m) => {
        if (!cancelled) setModel(m)
      })
      .catch((e) => {
        if (!cancelled) setError(String(e))
      })
    return () => {
      cancelled = true
    }
  }, [doc.id, doc.revision])
  const changes = useMemo<OfficeChange[]>(() => {
    try {
      const value = text ? JSON.parse(text) : []
      return Array.isArray(value) ? value : []
    } catch {
      return []
    }
  }, [text])
  const latest = useRef(changes)
  latest.current = changes
  const update = (key: string, value: string, original: string) => {
    const next = latest.current.filter((c) => c.key !== key)
    if (value !== original) next.push({ key, text: value })
    latest.current = next
    onChange(next.length ? JSON.stringify(next) : '')
  }
  return (
    <section className="office-editor">
      <div className="office-edit-toolbar">
        <span className="office-mode-label">{mode === 'edit' ? '编辑中' : '阅读中'}</span>
        <span className="office-mode-status" role="status">
          {changes.length
            ? `${changes.length} 处未保存修改 · ${draftStatus}`
            : mode === 'edit'
              ? '修改后点击顶部保存'
              : '可阅读内容和添加标记'}
          {mode === 'read' && changes.length > 0 ? ' · 当前阅读的是已保存版本' : ''}
        </span>
        <button
          className="secondary"
          disabled={busy}
          onClick={() => {
            if (mode === 'read') setHasEdited(true)
            setMode(mode === 'edit' ? 'read' : 'edit')
          }}
        >
          {mode === 'edit' ? '阅读与标记' : hasEdited ? '返回编辑' : '编辑'}
        </button>
      </div>
      {mode === 'read' ? (
        <ReadingPane
          doc={doc}
          text={null}
          query={query}
          location={location}
          onChanged={onChanged}
        />
      ) : (
        <>
          {error && <p role="alert">{error}</p>}
          {!model && !error && <p role="status">正在打开编辑器…</p>}
          {model &&
            (model.kind === 'docx' ? (
              <WordEditor model={model} changes={changes} onChange={update} busy={busy} />
            ) : (
              <SheetEditor model={model} changes={changes} onChange={update} busy={busy} />
            ))}
        </>
      )}
    </section>
  )
}
