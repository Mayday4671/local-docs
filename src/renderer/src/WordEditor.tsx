import { useEffect, useRef, useState } from 'react'
import { Search, Plus } from 'lucide-react'
import type { OfficeChange, OfficeEditField, OfficeEditModel } from '../../shared/types'

type MappedParagraph = { element: HTMLElement; field: OfficeEditField; original: Node[] }

// Reapply edits to the original runs so unchanged bold/colour/link spans survive.
function paint(paragraph: MappedParagraph, value: string) {
  const { element, field, original } = paragraph
  element.replaceChildren(...original.map((n) => n.cloneNode(true)))
  let prefix = 0,
    suffix = 0
  while (
    prefix < field.text.length &&
    prefix < value.length &&
    field.text[prefix] === value[prefix]
  )
    prefix++
  while (
    suffix < field.text.length - prefix &&
    suffix < value.length - prefix &&
    field.text.at(-suffix - 1) === value.at(-suffix - 1)
  )
    suffix++
  const endRemoved = field.text.length - suffix
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT)
  let node: Node | null,
    offset = 0,
    inserted = false
  while ((node = walker.nextNode())) {
    const old = node.textContent || '',
      start = offset,
      end = start + old.length
    offset = end
    let next = old.slice(0, Math.max(0, Math.min(old.length, prefix - start)))
    if (!inserted && prefix <= end) {
      next += value.slice(prefix, value.length - suffix)
      inserted = true
    }
    next += old.slice(Math.max(0, Math.min(old.length, endRemoved - start)))
    node.textContent = next
  }
  if (!inserted && value) element.append(document.createTextNode(value))
}

export function WordEditor({
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
  const host = useRef<HTMLDivElement>(null)
  const mapped = useRef<MappedParagraph[]>([])
  const live = useRef({ changes, onChange, busy })
  live.current = { changes, onChange, busy }
  const [status, setStatus] = useState('正在生成文档页面…')
  const [hint, setHint] = useState('点击页面中的文字直接修改；表格和图片保留在原来的位置。')
  const [filter, setFilter] = useState('')
  const [zoom, setZoom] = useState(100)
  const append = model.fields.find((f) => f.key.endsWith(':append'))
  const appendRef = useRef<HTMLTextAreaElement>(null)
  const [showAppend, setShowAppend] = useState(!!changes.find((c) => c.key === append?.key))
  useEffect(() => {
    let cancelled = false
    const shadow = host.current!.shadowRoot ?? host.current!.attachShadow({ mode: 'open' })
    const target = document.createElement('div')
    shadow.replaceChildren(target)
    mapped.current = []
    setStatus('正在生成文档页面…')
    const blockLink = (e: Event) => {
      if ((e.target as Element).closest('a')) e.preventDefault()
    }
    target.addEventListener('click', blockLink)
    void import('docx-preview')
      .then(({ renderAsync }) =>
        renderAsync(model.layoutBytes, target, target, {
          className: 'editable-word-page',
          inWrapper: true,
          ignoreFonts: true,
          renderAltChunks: false,
          renderChanges: false,
          renderComments: false,
          renderHeaders: true,
          renderFooters: true,
          useBase64URL: true,
        }),
      )
      .then(() => {
        if (cancelled) return
        target.querySelectorAll('a').forEach((a) => a.removeAttribute('href'))
        const style = document.createElement('style')
        style.textContent = `:host{display:block;color-scheme:light;color:#18233f}.editable-word-page-wrapper{padding:24px!important;background:transparent!important;min-width:max-content}section.editable-word-page{box-shadow:0 2px 12px #18233f20!important}p[data-editable]{cursor:text;min-height:1em;border-radius:2px;white-space:pre-wrap}p[data-editable]:hover{outline:1px dashed #8db9ee;outline-offset:3px}p[data-editable]:focus{outline:2px solid #087eff;outline-offset:3px}p[data-changed]{background:#eef6ff}p[data-found]{outline:2px solid #d2a236;outline-offset:3px}img{max-width:100%}`
        target.append(style)
        for (const field of model.fields) {
          if (!field.anchor) continue
          for (const marker of target.querySelectorAll(`[id="${field.anchor}"]`)) {
            const element = marker.closest('p')
            // Refuse ambiguous or altered renderings rather than editing the wrong paragraph.
            if (!element || element.textContent !== field.text) continue
            const paragraph: MappedParagraph = {
              element,
              field,
              original: [...element.childNodes].map((n) => n.cloneNode(true)),
            }
            mapped.current.push(paragraph)
            element.dataset.editable = 'true'
            element.contentEditable = live.current.busy ? 'false' : 'plaintext-only'
            element.setAttribute('role', 'textbox')
            element.setAttribute('aria-label', `编辑 ${field.label}`)
            element.setAttribute('aria-multiline', 'false')
            element.title = '直接修改文字；保留现有文字格式'
            element.tabIndex = 0
            const commit = () => {
              if (live.current.busy) return
              const value = (element.textContent || '').replace(/[\r\n\t]/g, ' ').slice(0, 32767)
              live.current.onChange(field.key, value, field.text)
              element.toggleAttribute('data-changed', value !== field.text)
            }
            let composing = false
            element.addEventListener('compositionstart', () => {
              composing = true
            })
            element.addEventListener('compositionend', () => {
              composing = false
              commit()
            })
            element.addEventListener('input', () => {
              if (!composing) commit()
            })
            element.addEventListener('beforeinput', (event) => {
              if (
                ['insertParagraph', 'insertLineBreak'].includes((event as InputEvent).inputType)
              ) {
                event.preventDefault()
                setHint('当前支持在原段落内改字；需要新段落时，点击“在末尾新增段落”。')
              }
            })
            element.addEventListener('paste', (event) => {
              event.preventDefault()
              const text =
                event.clipboardData
                  ?.getData('text/plain')
                  .replace(/[\r\n\t]+/g, ' ')
                  .slice(0, 32767) || ''
              document.execCommand('insertText', false, text)
            })
            element.addEventListener('drop', (event) => event.preventDefault())
            element.addEventListener('blur', () => {
              commit()
              const value = (element.textContent || '').replace(/[\r\n\t]/g, ' ').slice(0, 32767)
              paint(paragraph, value)
            })
            element.addEventListener('focus', () => setHint(`${field.label} · 可直接修改文字`))
            const value = live.current.changes.find((c) => c.key === field.key)?.text
            if (value !== undefined) {
              paint(paragraph, value)
              element.dataset.changed = 'true'
            }
          }
        }
        target.querySelectorAll('p:not([data-editable])').forEach((p) => {
          p.setAttribute('title', '此处暂不支持直接编辑，原有内容与格式会保留。')
          p.addEventListener('click', () =>
            setHint('此处暂不支持直接编辑；可编辑文字和空白段落悬停时会显示蓝色边框。'),
          )
        })
        mapped.current
          .find((p) => p.field.key.startsWith('word/document.xml:'))
          ?.element.focus({ preventScroll: true })
        setStatus(mapped.current.length ? '' : '这份文档没有可直接修改的普通段落，页面仍可阅读。')
      })
      .catch(() => {
        if (!cancelled) setStatus('文档页面未能打开，请切换“阅读与标记”查看；原文件完整保留。')
      })
    return () => {
      cancelled = true
      mapped.current = []
      target.remove()
      target.removeEventListener('click', blockLink)
    }
  }, [model])
  useEffect(() => {
    for (const p of mapped.current) {
      p.element.contentEditable = busy ? 'false' : 'plaintext-only'
      const value = changes.find((c) => c.key === p.field.key)?.text ?? p.field.text
      if (
        (p.element.getRootNode() as ShadowRoot).activeElement !== p.element &&
        p.element.textContent !== value
      )
        paint(p, value)
      p.element.toggleAttribute('data-changed', value !== p.field.text)
    }
  }, [changes, busy])
  const find = () => {
    let found = false
    for (const p of mapped.current) {
      const match =
        !!filter &&
        (p.element.textContent || '').toLocaleLowerCase().includes(filter.toLocaleLowerCase())
      p.element.toggleAttribute('data-found', match)
      if (match && !found) {
        p.element.scrollIntoView({ block: 'center' })
        found = true
      }
    }
    if (filter)
      setHint(found ? '已定位匹配文字，黄色边框为查找结果。' : '可编辑文字中没有找到匹配内容。')
  }
  return (
    <div className="word-edit-workspace">
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
            placeholder="在文档中查找…"
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
          />
          <button className="secondary" type="submit">
            查找
          </button>
        </form>
        <button
          className="secondary"
          disabled={busy || !append?.editable}
          onClick={() => {
            setShowAppend(true)
            requestAnimationFrame(() => {
              appendRef.current?.scrollIntoView({ block: 'center' })
              appendRef.current?.focus()
            })
          }}
        >
          <Plus size={16} />
          在末尾新增段落
        </button>
        <label className="office-zoom">
          缩放
          <select
            aria-label="文档缩放"
            value={zoom}
            onChange={(e) => setZoom(Number(e.target.value))}
          >
            {[75, 90, 100, 125, 150].map((v) => (
              <option key={v} value={v}>
                {v}%
              </option>
            ))}
          </select>
        </label>
      </div>
      <div className="word-edit-scroll">
        {status && (
          <p role="status" className="office-edit-status">
            {status}
          </p>
        )}
        <div
          className="word-edit-layout"
          ref={host}
          style={{ zoom: `${zoom}%` }}
          aria-label="Word 页面编辑"
        />
        {showAppend && append && (
          <div className="word-append">
            <label>
              末尾新增段落
              <textarea
                ref={appendRef}
                aria-label={`编辑 ${append.label}`}
                placeholder="输入新段落，按 Enter 换段…"
                disabled={busy || !append.editable}
                value={changes.find((c) => c.key === append.key)?.text ?? ''}
                maxLength={32767}
                onChange={(e) => onChange(append.key, e.target.value, '')}
              />
            </label>
          </div>
        )}
      </div>
      <div className="office-edit-status" role="status">
        {hint}
      </div>
    </div>
  )
}
