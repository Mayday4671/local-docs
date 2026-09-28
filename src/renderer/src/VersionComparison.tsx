import { Select } from './Select'
import { useEffect, useMemo, useRef, useState } from 'react'
import { X } from 'lucide-react'
import type { VersionContent, VersionRecord } from '../../shared/types'
import { compareVersions, type DiffValue, type VersionDiff } from '../../shared/version-diff'

const PAGE_SIZE = 100
const kindName = { added: '新增', removed: '删除', modified: '修改', same: '未变' }
export function VersionComparison({
  documentId,
  name,
  versions,
  selectedId,
  onClose,
}: {
  documentId: string
  name: string
  versions: VersionRecord[]
  selectedId: string
  onClose: () => void
}) {
  const [beforeId, setBeforeId] = useState(
    selectedId === versions[0].id ? (versions[1]?.id ?? selectedId) : selectedId,
  )
  const [afterId, setAfterId] = useState(versions[0].id)
  const [onlyChanges, setOnlyChanges] = useState(versions.length > 1)
  const [page, setPage] = useState(0)
  const [result, setResult] = useState<VersionDiff | null>(null)
  const [error, setError] = useState('')
  const [retry, setRetry] = useState(0)
  const dialog = useRef<HTMLDialogElement>(null)
  const scroll = useRef<HTMLDivElement>(null)
  const cache = useRef(new Map<string, Promise<VersionContent>>())
  const close = () => {
    dialog.current?.close()
    onClose()
  }
  useEffect(() => {
    if (!dialog.current!.open) dialog.current!.showModal()
  }, [])
  useEffect(() => {
    let cancelled = false
    setResult(null)
    setError('')
    setPage(0)
    const read = (id: string) => {
      if (!cache.current.has(id)) {
        if (cache.current.size >= 6) cache.current.delete(cache.current.keys().next().value!)
        const request = window.localDocs!.readVersion(documentId, id).catch((e) => {
          cache.current.delete(id)
          throw e
        })
        cache.current.set(id, request)
      }
      return cache.current.get(id)!
    }
    void Promise.all([read(beforeId), read(afterId)])
      .then(([a, b]) => {
        if (!cancelled) setResult(compareVersions(a, b))
      })
      .catch((e) => {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e))
      })
    return () => {
      cancelled = true
    }
  }, [documentId, beforeId, afterId, retry])
  const rows = useMemo(
    () => result?.rows.filter((row) => !onlyChanges || row.kind !== 'same') ?? [],
    [result, onlyChanges],
  )
  const pages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE))
  useEffect(() => {
    scroll.current?.scrollTo(0, 0)
  }, [page, result, onlyChanges])
  const label = (version: VersionRecord, index: number) =>
    `版本 ${versions.length - index}${index === 0 ? '（最新保存）' : ''} · ${new Date(version.createdAt).toLocaleString('zh-CN')} · ${version.reason}`
  const options = versions.map((version, index) => (
    <option key={version.id} value={version.id}>
      {label(version, index)}
    </option>
  ))
  const value = (item?: DiffValue) =>
    item ? (
      <>
        <span className="diff-location">{item.label}</span>
        <pre>{item.text || <em>空行 / 空值</em>}</pre>
      </>
    ) : (
      <span className="diff-absent">此版本无对应内容</span>
    )
  return (
    <dialog
      ref={dialog}
      className="version-comparison"
      aria-labelledby="comparison-title"
      onCancel={(event) => {
        event.preventDefault()
        close()
      }}
    >
      <header className="comparison-header">
        <div>
          <h2 id="comparison-title">历史版本对比</h2>
          <p title={name}>{name} · 只读查看，不包含未保存的修改</p>
        </div>
        <button className="icon-button" aria-label="关闭版本对比" onClick={close}>
          <X size={22} />
        </button>
      </header>
      <div className="comparison-selectors">
        <label>
          对比前
          <Select
            aria-label="对比前版本"
            value={beforeId}
            onChange={(e) => setBeforeId(e.target.value)}
          >
            {options}
          </Select>
        </label>
        <label>
          对比后
          <Select
            aria-label="对比后版本"
            value={afterId}
            onChange={(e) => setAfterId(e.target.value)}
          >
            {options}
          </Select>
        </label>
      </div>
      <div className="comparison-summary">
        <span role="status">
          {result
            ? `新增 ${result.counts.added} · 删除 ${result.counts.removed} · 修改 ${result.counts.modified}`
            : error
              ? '对比未完成'
              : '正在读取历史版本…'}
        </span>
        <label>
          <input
            type="checkbox"
            checked={onlyChanges}
            onChange={(e) => {
              setOnlyChanges(e.target.checked)
              setPage(0)
            }}
          />
          仅看差异
        </label>
      </div>
      <p className="comparison-scope">
        {name.toLowerCase().endsWith('.docx')
          ? '对比段落文字（含表格内文字、页眉页脚等）。空白段落、字体、排版和图片变化暂不展开。'
          : name.toLowerCase().endsWith('.xlsx')
            ? '对比工作表、单元格值与公式；公式结果为文件中保存的结果。单元格样式、图片和图表变化暂不展开。'
            : '按文本行对比内容，忽略 Windows / Unix 换行符差别；图片显示为引用地址。'}
      </p>
      <div className="comparison-content" ref={scroll}>
        {error && (
          <div className="comparison-empty" role="alert">
            <p>{error}</p>
            <button className="secondary" onClick={() => setRetry((value) => value + 1)}>
              重新读取
            </button>
          </div>
        )}
        {result && (
          <>
            {result.warnings.length > 0 && (
              <div className="comparison-warnings">
                {result.warnings.map((warning, i) => (
                  <p key={i}>{warning}</p>
                ))}
              </div>
            )}
            {result.counts.added + result.counts.removed + result.counts.modified === 0 && (
              <div className="comparison-empty">
                <strong>
                  {result.identical
                    ? '两个版本的文件内容相同'
                    : '可对比的文字和数据没有差异，但文件仍有其他变化'}
                </strong>
                {!result.identical && <p>差别可能来自格式、图片、空白结构或文件元数据。</p>}
                {onlyChanges && <p>取消“仅看差异”可查看版本内容。</p>}
              </div>
            )}
            {rows.length > 0 && (
              <table className="comparison-table" aria-label="版本差异">
                <thead>
                  <tr>
                    <th>变化</th>
                    <th>对比前</th>
                    <th>对比后</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE).map((item, index) => (
                    <tr key={`${page}-${index}`} data-kind={item.kind}>
                      <th scope="row">{kindName[item.kind]}</th>
                      <td className="diff-before">{value(item.before)}</td>
                      <td className="diff-after">{value(item.after)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </>
        )}
      </div>
      <footer className="comparison-footer">
        <span>
          {result
            ? `共 ${rows.length} 项 · ${page + 1} / ${pages} 页`
            : '读取历史不会恢复或覆盖文档'}
        </span>
        <button
          className="secondary"
          disabled={page === 0 || !result}
          onClick={() => setPage(page - 1)}
        >
          上一页
        </button>
        <button
          className="secondary"
          disabled={page + 1 >= pages || !result}
          onClick={() => setPage(page + 1)}
        >
          下一页
        </button>
      </footer>
    </dialog>
  )
}
