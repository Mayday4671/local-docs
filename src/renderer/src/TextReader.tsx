import { memo, useEffect, useMemo, useState } from 'react'
import { MAX_SYNTAX_LENGTH, textLanguage, type SyntaxToken } from '../../shared/text-language'

export const TextReader = memo(function TextReader({
  extension,
  text,
}: {
  extension: string
  text: string
}) {
  const language = useMemo(() => textLanguage(extension, text), [extension, text])
  const [result, setResult] = useState<{
    text: string
    language: string
    tokens: SyntaxToken[] | null
  } | null>(null)
  useEffect(() => {
    if (!language.language || text.length > MAX_SYNTAX_LENGTH || !text) return
    let worker: Worker | undefined
    let timer: ReturnType<typeof setTimeout> | undefined
    let active = true
    const complete = (tokens: SyntaxToken[] | null) => {
      if (!active) return
      active = false
      clearTimeout(timer)
      worker?.terminate()
      setResult({ text, language: language.language!, tokens })
    }
    try {
      worker = new Worker(new URL('./text-syntax.worker.ts', import.meta.url), { type: 'module' })
      worker.onmessage = (event: MessageEvent<SyntaxToken[] | null>) => complete(event.data)
      worker.onerror = () => complete(null)
      timer = setTimeout(() => complete(null), 5000)
      worker.postMessage({ language: language.language, text })
    } catch {
      complete(null)
    }
    return () => {
      active = false
      clearTimeout(timer)
      worker?.terminate()
    }
  }, [language.language, text])
  const current = result?.text === text && result.language === language.language ? result : null
  const large = !!language.language && text.length > MAX_SYNTAX_LENGTH
  return (
    <div className="text-reader">
      <div className="text-reader-meta">
        <span className="text-language" data-language={language.language || 'plaintext'}>
          {language.label}
        </span>
        <span>{language.source === 'content' ? '按内容识别' : '只读'}</span>
        {(large || (current && !current.tokens)) && <span>当前按普通文本显示</span>}
      </div>
      <pre className="plain-text-reader" data-mark-scope="text">
        {current?.tokens
          ? current.tokens.map((token, i) =>
              token.classes.length ? (
                <span key={i} className={token.classes.join(' ')}>
                  {token.text}
                </span>
              ) : (
                token.text
              ),
            )
          : text}
      </pre>
    </div>
  )
})
