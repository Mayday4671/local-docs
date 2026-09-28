import { ZoomSelect } from './ZoomSelect'
import { useCtrlWheelZoom } from './useCtrlWheelZoom'
import { useEffect, useRef, useState } from 'react'
import {
  getDocument,
  GlobalWorkerOptions,
  type PDFDocumentProxy,
  type RenderTask,
} from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import type { DocumentRecord } from '../../shared/types'
import { FileFallback } from './FileReader'

GlobalWorkerOptions.workerSrc = workerUrl
const assets = new URL('./pdf-assets/', document.baseURI).href

/** Local bytes only. No links, forms, document actions or embedded JavaScript are activated. */
export default function PdfReader({ doc }: { doc: DocumentRecord }) {
  const [pdf, setPdf] = useState<PDFDocumentProxy | null>(null)
  const [page, setPage] = useState(1),
    [zoom, setZoom] = useState(100)
  const [error, setError] = useState(''),
    [loading, setLoading] = useState(true)
  const [text, setText] = useState('')
  const canvasHost = useRef<HTMLDivElement>(null)
  const reader = useRef<HTMLDivElement>(null)
  useCtrlWheelZoom(reader, setZoom, !!pdf && !error)
  const [availableWidth, setAvailableWidth] = useState(0)
  useEffect(() => {
    if (!pdf || !canvasHost.current) return
    const host = canvasHost.current
    const observer = new ResizeObserver(() =>
      setAvailableWidth(Math.max(100, host.clientWidth - 32)),
    )
    observer.observe(host)
    return () => observer.disconnect()
  }, [pdf])
  useEffect(() => {
    let active = true,
      task: ReturnType<typeof getDocument> | undefined
    void window
      .localDocs!.readFilePreview(doc.id)
      .then(async (bytes) => {
        if (!active) return
        task = getDocument({
          data: Uint8Array.from(bytes),
          cMapUrl: `${assets}cmaps/`,
          standardFontDataUrl: `${assets}standard_fonts/`,
          wasmUrl: `${assets}wasm/`,
          iccUrl: `${assets}iccs/`,
          useWorkerFetch: true,
          enableXfa: false,
          maxImageSize: 16_000_000,
          canvasMaxAreaInBytes: 64_000_000,
        })
        const value = await task.promise
        if (active) setPdf(value)
      })
      .catch((e: unknown) => {
        if (active) {
          setError(
            e instanceof Error && e.name === 'PasswordException'
              ? '这是加密 PDF，请导出后使用本地 PDF 软件输入密码阅读。'
              : `无法预览 PDF，原文件仍完整保留。${String(e)}`,
          )
          setLoading(false)
        }
      })
    return () => {
      active = false
      void task?.destroy().catch(() => {})
    }
  }, [doc.id])
  useEffect(() => {
    if (!pdf || !availableWidth) return
    let active = true,
      render: RenderTask | undefined
    setLoading(true)
    setText('')
    void pdf
      .getPage(page)
      .then(async (value) => {
        if (!active) return
        const fit = Math.min(1, availableWidth / value.getViewport({ scale: 1 }).width)
        const base = value.getViewport({ scale: (zoom / 100) * fit })
        const scale =
          (zoom / 100) *
          fit *
          Math.min(
            window.devicePixelRatio || 1,
            2,
            Math.sqrt(16_000_000 / (base.width * base.height)),
          )
        const viewport = value.getViewport({ scale })
        const canvas = document.createElement('canvas')
        canvas.width = Math.ceil(viewport.width)
        canvas.height = Math.ceil(viewport.height)
        canvas.style.width = `${base.width}px`
        canvas.style.maxWidth = 'none'
        canvas.style.height = 'auto'
        canvas.setAttribute('aria-label', `PDF 第 ${page} 页`)
        canvasHost.current?.replaceChildren(canvas)
        render = value.render({ canvas, viewport })
        await render.promise
        if (active) setLoading(false)
        const content = await value.getTextContent()
        if (active)
          setText(
            content.items
              .map((item) => ('str' in item ? item.str + (item.hasEOL ? '\n' : ' ') : ''))
              .join(''),
          )
      })
      .catch((e) => {
        if (active) {
          setError(`此页无法显示：${String(e)}`)
          setLoading(false)
        }
      })
    return () => {
      active = false
      render?.cancel()
    }
  }, [pdf, page, zoom, availableWidth])
  if (error) return <FileFallback doc={doc} message={error} />
  return (
    <div className="pdf-reader" ref={reader}>
      <div className="pdf-tools">
        <button
          className="secondary"
          disabled={!pdf || page <= 1}
          onClick={() => setPage(page - 1)}
        >
          上一页
        </button>
        <span aria-live="polite">{pdf ? `${page} / ${pdf.numPages} 页` : '正在读取 PDF…'}</span>
        <button
          className="secondary"
          disabled={!pdf || page >= pdf.numPages}
          onClick={() => setPage(page + 1)}
        >
          下一页
        </button>
        <ZoomSelect label="PDF 缩放" value={zoom} onChange={setZoom} disabled={!pdf} />
      </div>
      {loading && (
        <p className="pdf-loading" role="status">
          正在渲染页面…
        </p>
      )}
      <div className="pdf-canvas" ref={canvasHost} />
      {text.trim() && (
        <details className="pdf-text">
          <summary>本页文字（可复制）</summary>
          <pre>{text}</pre>
        </details>
      )}
    </div>
  )
}
