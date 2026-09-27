import ReactMarkdown from 'react-markdown'
import { memo } from 'react'
import remarkGfm from 'remark-gfm'
import type { Attachment } from '../../shared/types'

// The preview and editor must interpret exactly the same Markdown dialect.
export const Markdown = memo(function Markdown({
  text,
  documentId,
  attachments = [],
}: {
  text: string
  documentId?: string
  attachments?: Attachment[]
}) {
  return (
    <div className="markdown-body" data-mark-scope="markdown">
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        skipHtml
        urlTransform={(url) =>
          /^attachments\/[a-f0-9-]{36}\.(png|jpg|webp)$/.test(url) ? url : ''
        }
        components={{
          table: ({ children }) => (
            <div
              className="markdown-table-scroll"
              tabIndex={0}
              role="region"
              aria-label="Markdown 表格"
            >
              <table>{children}</table>
            </div>
          ),
          // Imported documents cannot fetch remote images or navigate the app.
          img: ({ alt, src }) => {
            const a = attachments.find((a) => src === `attachments/${a.id}${a.extension}`)
            return a && documentId ? (
              <img
                className="managed-image"
                src={`app://local/attachment/${documentId}/${a.id}`}
                alt={alt || a.name}
              />
            ) : (
              <span className="attachment-note">[图片：{alt || '请先插入本地图片附件'}]</span>
            )
          },
          a: ({ children, href }) => (
            <span className="document-link" title={href}>
              {children}
            </span>
          ),
        }}
      >
        {text || '*文档暂时没有内容。*'}
      </ReactMarkdown>
    </div>
  )
})
