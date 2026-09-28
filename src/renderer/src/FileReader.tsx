import { lazy, Suspense, useEffect, useState } from 'react'
import { Download, FileText } from 'lucide-react'
import { fileKind } from '../../shared/file-types'
import type { DocumentRecord } from '../../shared/types'

const PdfReader = lazy(() => import('./PdfReader'))
export function FileFallback({ doc, message }: { doc: DocumentRecord; message?: string }) {
  const [status, setStatus] = useState('')
  return (
    <div className="file-fallback">
      <FileText size={40} />
      <strong>{doc.name}</strong>
      <p>{message || '文件已完整保存，此格式暂不支持内置预览。'}</p>
      <p>可导出原文件，用电脑上的对应软件打开。</p>
      <button
        className="secondary"
        onClick={async () => {
          try {
            if (await window.localDocs!.exportDocument(doc.id)) setStatus('已导出原文件')
          } catch (error) {
            setStatus(String(error))
          }
        }}
      >
        <Download size={16} />
        导出原文件
      </button>
      {status && <p role="status">{status}</p>}
    </div>
  )
}
function ImageReader({ doc }: { doc: DocumentRecord }) {
  const [url, setUrl] = useState(''),
    [error, setError] = useState('')
  useEffect(() => {
    let active = true,
      objectUrl = ''
    setUrl('')
    setError('')
    void window
      .localDocs!.readFilePreview(doc.id)
      .then((bytes) => {
        if (!active) return
        const mime = (
          {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.webp': 'image/webp',
            '.gif': 'image/gif',
            '.bmp': 'image/bmp',
          } as Record<string, string>
        )[doc.extension]
        objectUrl = URL.createObjectURL(new Blob([Uint8Array.from(bytes)], { type: mime }))
        setUrl(objectUrl)
      })
      .catch((e) => {
        if (active) setError(String(e))
      })
    return () => {
      active = false
      if (objectUrl) URL.revokeObjectURL(objectUrl)
    }
  }, [doc.id, doc.revision, doc.extension])
  if (error) return <FileFallback doc={doc} message={error} />
  return url ? (
    <div className="image-reader">
      <img src={url} alt={doc.name} onError={() => setError('图片无法解码，原文件仍完整保留。')} />
    </div>
  ) : (
    <p className="muted">正在读取图片…</p>
  )
}
export function FileReader({ doc, text }: { doc: DocumentRecord; text: string | null }) {
  const kind = fileKind(doc.extension)
  if (kind === 'text')
    return text !== null ? (
      <pre className="plain-text-reader" data-mark-scope="text">
        {text}
      </pre>
    ) : (
      <FileFallback
        doc={doc}
        message="文件已保存。文本超过 5 MB、编码无法识别或包含二进制内容，暂不提供文本预览。"
      />
    )
  if (kind === 'pdf')
    return (
      <Suspense fallback={<p className="muted">正在加载 PDF 阅读器…</p>}>
        <PdfReader key={`${doc.id}:${doc.revision}`} doc={doc} />
      </Suspense>
    )
  if (kind === 'image') return <ImageReader doc={doc} />
  return <FileFallback doc={doc} />
}
