import { ZoomSelect } from './ZoomSelect'
import { useCtrlWheelZoom } from './useCtrlWheelZoom'
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type SetStateAction,
} from 'react'
import {
  getDocument,
  GlobalWorkerOptions,
  PDFDataRangeTransport,
  type PDFDocumentProxy,
  type PDFPageProxy,
  type RenderTask,
} from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import type { DocumentRecord } from '../../shared/types'
import { FileFallback } from './FileReader'

GlobalWorkerOptions.workerSrc = workerUrl
const assets = new URL('./pdf-assets/', document.baseURI).href
type Size = { width: number; height: number }
type PageBox = Size & { top: number }
const GAP = 16

class LocalPdfRange extends PDFDataRangeTransport {
  private stopped = false
  constructor(
    private id: string,
    private hash: string,
    size: number,
    private failed: (error: unknown) => void,
  ) {
    super(size, new Uint8Array(), true)
  }
  requestDataRange(begin: number, end: number) {
    void this.read(begin, end).catch((error) => {
      if (!this.stopped) this.failed(error)
    })
  }
  private async read(begin: number, end: number) {
    const bytes = new Uint8Array(end - begin)
    for (let offset = begin; offset < end; offset += 1024 ** 2) {
      if (this.stopped) return
      const part = await window.localDocs!.readPdfRange(
        this.id,
        this.hash,
        offset,
        Math.min(end, offset + 1024 ** 2),
      )
      bytes.set(part, offset - begin)
    }
    if (!this.stopped) this.onDataRange(begin, bytes)
  }
  abort() {
    this.stopped = true
  }
}

function pageAt(boxes: PageBox[], position: number) {
  let low = 0,
    high = boxes.length - 1
  while (low < high) {
    const mid = Math.ceil((low + high) / 2)
    if (boxes[mid].top <= position) low = mid
    else high = mid - 1
  }
  return low
}

function PdfPage({
  pdf,
  number,
  scale,
  onSize,
}: {
  pdf: PDFDocumentProxy
  number: number
  scale: number
  onSize: (number: number, size: Size) => void
}) {
  const host = useRef<HTMLDivElement>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  useEffect(() => {
    let active = true,
      render: RenderTask | undefined,
      renderedPage: PDFPageProxy | undefined
    setLoading(true)
    setError('')
    void pdf
      .getPage(number)
      .then(async (value) => {
        if (!active) return
        renderedPage = value
        const natural = value.getViewport({ scale: 1 })
        onSize(number, { width: natural.width, height: natural.height })
        const base = value.getViewport({ scale })
        const pixelRatio = Math.min(
          window.devicePixelRatio || 1,
          2,
          Math.sqrt(8_000_000 / (base.width * base.height)),
        )
        const viewport = value.getViewport({ scale: scale * pixelRatio })
        const canvas = document.createElement('canvas')
        canvas.width = Math.ceil(viewport.width)
        canvas.height = Math.ceil(viewport.height)
        canvas.style.width = `${base.width}px`
        canvas.style.height = `${base.height}px`
        canvas.setAttribute('aria-label', `PDF 第 ${number} 页`)
        host.current?.replaceChildren(canvas)
        render = value.render({ canvas, viewport })
        await render.promise
        if (active) {
          canvas.dataset.rendered = 'true'
          setLoading(false)
        }
      })
      .catch((e) => {
        if (active) {
          setError(`第 ${number} 页无法显示：${String(e)}`)
          setLoading(false)
        }
      })
    return () => {
      active = false
      render?.cancel()
      // Once a page leaves the viewport, release decoded images as well as its canvas.
      void render?.promise
        .catch(() => {})
        .then(() => renderedPage?.cleanup())
        .catch(() => {})
    }
  }, [pdf, number, scale, onSize])
  return (
    <>
      <div className="pdf-canvas" ref={host} />
      {loading && (
        <p className="pdf-loading" role="status">
          正在渲染页面…
        </p>
      )}
      {error && (
        <p className="pdf-page-error" role="alert">
          {error}
        </p>
      )}
    </>
  )
}

/** Continuous local reading. Only pages around the viewport have live canvases. */
export default function PdfReader({ doc }: { doc: DocumentRecord }) {
  const [pdf, setPdf] = useState<PDFDocumentProxy | null>(null)
  const [sizes, setSizes] = useState<Record<number, Size>>({})
  const [page, setPage] = useState(1),
    [zoom, setZoom] = useState(100)
  const [error, setError] = useState(''),
    [text, setText] = useState('')
  const [view, setView] = useState({ width: 0, height: 0, top: 0 })
  const reader = useRef<HTMLDivElement>(null),
    scroll = useRef<HTMLDivElement>(null)
  const anchor = useRef<{ top: number; left: number } | null>(null)
  const capturePosition = useCallback(() => {
    const el = scroll.current
    if (el) anchor.current = { top: el.scrollTop, left: el.scrollLeft }
  }, [])
  const changeZoom = useCallback(
    (value: SetStateAction<number>) => {
      capturePosition()
      setZoom(value)
    },
    [capturePosition],
  )
  useCtrlWheelZoom(reader, changeZoom, !!pdf && !error)
  const onSize = useCallback(
    (number: number, size: Size) => {
      capturePosition()
      setSizes((current) =>
        current[number]?.width === size.width && current[number]?.height === size.height
          ? current
          : { ...current, [number]: size },
      )
    },
    [capturePosition],
  )

  useEffect(() => {
    let active = true,
      task: ReturnType<typeof getDocument> | undefined
    const failed = (e: unknown) => {
      if (!active) return
      setError(
        e instanceof Error && e.name === 'PasswordException'
          ? '这是加密 PDF，请导出后使用本地 PDF 软件输入密码阅读。'
          : `无法预览 PDF，原文件仍完整保留。${String(e)}`,
      )
      void task?.destroy().catch(() => {})
    }
    void window
      .localDocs!.pdfInfo(doc.id)
      .then(async ({ size, hash }) => {
        if (!active) return
        task = getDocument({
          range: new LocalPdfRange(doc.id, hash, size, failed),
          rangeChunkSize: 256 * 1024,
          disableStream: true,
          disableAutoFetch: true,
          cMapUrl: `${assets}cmaps/`,
          standardFontDataUrl: `${assets}standard_fonts/`,
          wasmUrl: `${assets}wasm/`,
          iccUrl: `${assets}iccs/`,
          useWorkerFetch: true,
          enableXfa: false,
          canvasMaxAreaInBytes: 32_000_000,
        })
        const value = await task.promise
        const first = (await value.getPage(1)).getViewport({ scale: 1 })
        if (active) {
          setSizes({ 1: { width: first.width, height: first.height } })
          setPdf(value)
        }
      })
      .catch(failed)
    return () => {
      active = false
      void task?.destroy().catch(() => {})
    }
  }, [doc.id])

  useEffect(() => {
    const el = scroll.current
    if (!el) return
    const update = () => {
      capturePosition()
      setView({ width: el.clientWidth, height: el.clientHeight, top: el.scrollTop })
    }
    const observer = new ResizeObserver(update)
    observer.observe(el)
    update()
    return () => observer.disconnect()
  }, [capturePosition])

  const scale = sizes[1] ? ((Math.max(100, view.width - 2 * GAP) / sizes[1].width) * zoom) / 100 : 1
  const layout = useMemo(() => {
    let top = GAP
    return Array.from({ length: pdf?.numPages || 0 }, (_, index) => {
      const size = sizes[index + 1] || sizes[1]
      const box = { width: size.width * scale, height: size.height * scale, top }
      top += box.height + GAP
      return box
    })
  }, [pdf, sizes, scale])
  const width = layout.reduce((max, box) => Math.max(max, box.width + GAP * 2), view.width)
  const previousLayout = useRef<{ boxes: PageBox[]; width: number } | null>(null)
  // Preserve the top visible page and its relative position during zoom and lazy size discovery.
  useLayoutEffect(() => {
    const el = scroll.current,
      previous = previousLayout.current
    if (!el || !layout.length) return
    if (previous?.boxes.length) {
      const top = anchor.current?.top ?? view.top
      const left = anchor.current?.left ?? el.scrollLeft
      const index = pageAt(previous.boxes, top + GAP + 1)
      const old = previous.boxes[index],
        next = layout[index]
      const fraction = (top + GAP - old.top) / old.height
      el.scrollTop = next.top + fraction * next.height - GAP
      el.scrollLeft = ((left + el.clientWidth / 2) * width) / previous.width - el.clientWidth / 2
    }
    previousLayout.current = { boxes: layout, width }
    anchor.current = null
    setView((current) => ({ ...current, top: el.scrollTop }))
  }, [layout, width])

  useEffect(() => {
    if (!pdf) return
    let active = true
    setText('')
    void pdf
      .getPage(page)
      .then((value) => value.getTextContent())
      .then((content) => {
        if (active)
          setText(
            content.items
              .map((item) => ('str' in item ? item.str + (item.hasEOL ? '\n' : ' ') : ''))
              .join(''),
          )
      })
      .catch(() => {
        if (active) setText('本页文字暂时无法提取。')
      })
    return () => {
      active = false
    }
  }, [pdf, page])

  function goTo(number: number) {
    if (!scroll.current || !layout[number - 1]) return
    scroll.current.scrollTop = layout[number - 1].top - GAP
    setView((current) => ({ ...current, top: scroll.current!.scrollTop }))
    setPage(number)
  }
  if (error) return <FileFallback doc={doc} message={error} />
  const last = layout.at(-1)
  return (
    <div className="pdf-reader" ref={reader}>
      <div className="pdf-tools">
        <button className="secondary" disabled={!pdf || page <= 1} onClick={() => goTo(page - 1)}>
          上一页
        </button>
        <span aria-live="polite">{pdf ? `${page} / ${pdf.numPages} 页` : '正在读取 PDF…'}</span>
        <button
          className="secondary"
          disabled={!pdf || page >= pdf.numPages}
          onClick={() => goTo(page + 1)}
        >
          下一页
        </button>
        <ZoomSelect label="PDF 缩放" value={zoom} onChange={changeZoom} disabled={!pdf} />
      </div>
      <div
        className="pdf-scroll"
        ref={scroll}
        tabIndex={0}
        aria-label="PDF 连续阅读区域"
        onScroll={(event) => {
          const top = event.currentTarget.scrollTop
          setView((current) => ({ ...current, top }))
          if (layout.length) setPage(pageAt(layout, top + GAP + 1) + 1)
        }}
      >
        {!pdf && (
          <p className="pdf-loading" role="status">
            正在读取 PDF…
          </p>
        )}
        <div
          className="pdf-pages"
          style={{
            width,
            height: last ? last.top + Math.max(last.height, view.height - GAP) + GAP : 0,
          }}
        >
          {layout.map((box, index) =>
            box.top + box.height >= view.top - 300 && box.top <= view.top + view.height + 300 ? (
              <div
                className="pdf-page"
                data-page-number={index + 1}
                key={index}
                style={{
                  top: box.top,
                  left: (width - box.width) / 2,
                  width: box.width,
                  height: box.height,
                }}
              >
                <PdfPage pdf={pdf!} number={index + 1} scale={scale} onSize={onSize} />
              </div>
            ) : null,
          )}
        </div>
      </div>
      {pdf && (
        <details className="pdf-text">
          <summary>本页文字（可复制）</summary>
          <pre>{text || '本页暂无可复制文字。'}</pre>
        </details>
      )}
    </div>
  )
}
