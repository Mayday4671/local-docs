import { Select } from './Select'
import { useEffect, useRef, useState, type ReactNode } from 'react'
import { FileText, Folder, X } from 'lucide-react'
import type {
  Category,
  FolderImportPreview,
  FolderImportResult,
  LibrarySnapshot,
} from '../../shared/types'
import { categoryPath, fileSize } from './LibraryWorkspace'

export function categoryDescendants(categories: Category[], id: string): Set<string> {
  const ids = new Set([id])
  let size = 0
  while (ids.size !== size) {
    size = ids.size
    for (const c of categories) if (c.parentId && ids.has(c.parentId)) ids.add(c.id)
  }
  return ids
}
function CategorySelect({
  categories,
  value,
  onChange,
  label,
  rootLabel,
  exclude,
}: {
  categories: Category[]
  value: string | null
  onChange: (id: string | null) => void
  label: string
  rootLabel: string
  exclude?: Set<string>
}) {
  const choices = categories
    .filter((c) => !exclude?.has(c.id))
    .map((c) => ({
      id: c.id,
      path: categoryPath(categories, c.id)
        .map((x) => x.name)
        .join(' / '),
    }))
    .sort((a, b) => a.path.localeCompare(b.path, 'zh-CN'))
  return (
    <label className="form-label">
      {label}
      <Select
        aria-label={label}
        value={value || ''}
        onChange={(e) => onChange(e.target.value || null)}
      >
        <option value="">{rootLabel}</option>
        {choices.map((c) => (
          <option key={c.id} value={c.id}>
            {c.path}
          </option>
        ))}
      </Select>
    </label>
  )
}
function Dialog({
  title,
  busy = false,
  onClose,
  children,
}: {
  title: string
  busy?: boolean
  onClose: () => void
  children: ReactNode
}) {
  const ref = useRef<HTMLElement>(null)
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null
    ref.current?.focus()
    return () => previous?.focus()
  }, [])
  return (
    <div
      className="modal-backdrop organization-backdrop"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget && !busy) onClose()
      }}
    >
      <section
        ref={ref}
        tabIndex={-1}
        className="modal organization-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={title}
        onKeyDown={(e) => {
          // The picker handles its own Escape/Tab before the containing dialog.
          if (e.target instanceof Element && e.target.closest('select:open')) return
          if (e.key === 'Escape' && !busy) {
            e.stopPropagation()
            onClose()
          }
          if (e.key !== 'Tab') return
          const controls = [
            ...ref.current!.querySelectorAll<HTMLElement>(
              'button:not(:disabled), input:not(:disabled), select:not(:disabled), summary',
            ),
          ].filter((el) => el.getClientRects().length)
          const first = controls[0],
            last = controls.at(-1)
          if (
            e.shiftKey &&
            (document.activeElement === first || document.activeElement === ref.current)
          ) {
            e.preventDefault()
            last?.focus()
          } else if (
            !e.shiftKey &&
            (document.activeElement === last || document.activeElement === ref.current)
          ) {
            e.preventDefault()
            first?.focus()
          }
        }}
      >
        <div className="modal-heading">
          <h2>{title}</h2>
          <button
            className="icon-button"
            aria-label="关闭整理对话框"
            disabled={busy}
            onClick={onClose}
          >
            <X size={18} />
          </button>
        </div>
        {children}
      </section>
    </div>
  )
}
function useAction() {
  const [busy, setBusy] = useState(false),
    [error, setError] = useState('')
  const run = async (action: () => Promise<void>) => {
    if (busy) return
    setBusy(true)
    setError('')
    try {
      await action()
    } catch (error) {
      setError(
        error instanceof Error
          ? error.message.replace(/^Error invoking remote method '[^']+': Error: /, '')
          : String(error),
      )
    } finally {
      setBusy(false)
    }
  }
  return {
    busy,
    run,
    error: error ? (
      <p className="organization-error" role="alert">
        {error}
      </p>
    ) : null,
  }
}

export function CategoryDialog({
  category,
  snapshot,
  onClose,
  onSave,
  onDelete,
}: {
  category: Category
  snapshot: LibrarySnapshot
  onClose: () => void
  onSave: (name: string, parentId: string | null) => Promise<void>
  onDelete: (destination: string | null) => Promise<void>
}) {
  const [name, setName] = useState(category.name),
    [parent, setParent] = useState(category.parentId)
  const [deleting, setDeleting] = useState(false),
    [destination, setDestination] = useState<string | null>(null)
  const { busy, run, error } = useAction()
  const excluded = categoryDescendants(snapshot.categories, category.id)
  const affected = snapshot.documents.filter((d) => d.categoryId && excluded.has(d.categoryId))
  return (
    <Dialog title={deleting ? '删除分类并保留文件' : '管理分类'} busy={busy} onClose={onClose}>
      {deleting ? (
        <>
          <p className="restore-warning">
            将删除“{category.name}”及其子分类，共 {excluded.size} 个分类。分类结构删除后无法撤销。
          </p>
          <p>
            其中 {affected.filter((d) => !d.deletedAt).length} 份文件将转移到下方位置；另有{' '}
            {affected.filter((d) => d.deletedAt).length}{' '}
            份回收站文件会更新所属分类，仍留在回收站。内容、标签和历史版本都会保留。
          </p>
          <fieldset disabled={busy}>
            <CategorySelect
              categories={snapshot.categories}
              value={destination}
              onChange={setDestination}
              label="文件转移到"
              rootLabel="不指定分类（在全部文件中查看）"
              exclude={excluded}
            />
          </fieldset>
          {error}
          <div className="modal-actions">
            <button className="secondary" disabled={busy} onClick={() => setDeleting(false)}>
              返回
            </button>
            <button
              className="danger"
              disabled={busy}
              onClick={() => void run(() => onDelete(destination))}
            >
              删除分类，保留文件
            </button>
          </div>
        </>
      ) : (
        <form
          onSubmit={(e) => {
            e.preventDefault()
            void run(() => onSave(name, parent))
          }}
        >
          <fieldset disabled={busy}>
            <label className="form-label">
              分类名称
              <input
                aria-label="分类名称"
                required
                maxLength={180}
                value={name}
                onChange={(e) => setName(e.target.value)}
              />
            </label>
            <CategorySelect
              categories={snapshot.categories}
              value={parent}
              onChange={setParent}
              label="移动到上级分类"
              rootLabel="顶级分类"
              exclude={excluded}
            />
          </fieldset>
          <p className="muted">修改名称或位置后，子分类与文件会一起保留。</p>
          {error}
          <div className="modal-actions">
            <button
              type="button"
              className="danger category-delete"
              disabled={busy}
              onClick={() => setDeleting(true)}
            >
              删除分类…
            </button>
            <button type="button" className="secondary" disabled={busy} onClick={onClose}>
              取消
            </button>
            <button className="primary" disabled={busy || !name.trim()}>
              保存分类
            </button>
          </div>
        </form>
      )}
    </Dialog>
  )
}

export function MoveDialog({
  ids,
  categories,
  initial,
  onClose,
  onMove,
}: {
  ids: string[]
  categories: Category[]
  initial: string | null
  onClose: () => void
  onMove: (destination: string | null) => Promise<void>
}) {
  const [destination, setDestination] = useState(initial)
  const { busy, run, error } = useAction()
  return (
    <Dialog title="移动文件到分类" busy={busy} onClose={onClose}>
      <p>已选择 {ids.length} 份文件。移动只改变分类，原文件和历史版本都会保留。</p>
      <fieldset disabled={busy}>
        <CategorySelect
          categories={categories}
          value={destination}
          onChange={setDestination}
          label="目标分类"
          rootLabel="不指定分类（在全部文件中查看）"
        />
      </fieldset>
      {error}
      <div className="modal-actions">
        <button className="secondary" disabled={busy} onClick={onClose}>
          取消
        </button>
        <button
          className="primary"
          disabled={busy}
          onClick={() => void run(() => onMove(destination))}
        >
          确认移动
        </button>
      </div>
    </Dialog>
  )
}

export function FolderPreviewDialog({
  preview,
  categories,
  initial,
  onClose,
  onImport,
}: {
  preview: FolderImportPreview
  categories: Category[]
  initial: string | null
  onClose: () => void
  onImport: (parent: string | null, duplicates: 'skip' | 'keep') => void
}) {
  const [parent, setParent] = useState(initial),
    [duplicates, setDuplicates] = useState<'skip' | 'keep'>('skip')
  return (
    <Dialog title="导入文件夹" onClose={onClose}>
      <p className="folder-source">{preview.sourcePath}</p>
      <div className="folder-stats">
        <span>
          <strong>{preview.files}</strong> 份文件
        </span>
        <span>
          <strong>{preview.folders}</strong> 个分类
        </span>
        <span>{fileSize(preview.bytes)}</span>
      </div>
      <p className="muted">
        保留“{preview.rootName}”及子目录结构，包含空文件夹。复制文档、PDF、文本等所有普通
        文件，原文件保持不变。
      </p>
      <CategorySelect
        categories={categories}
        value={parent}
        onChange={setParent}
        label="导入到"
        rootLabel="顶级分类"
      />
      <label className="form-label">
        重复文件
        <Select
          aria-label="重复文件"
          value={duplicates}
          onChange={(e) => setDuplicates(e.target.value as 'skip' | 'keep')}
        >
          <option value="skip">跳过同分类中同名且内容相同的文件</option>
          <option value="keep">全部保留为新副本</option>
        </Select>
      </label>
      <p className="muted">同名分类会复用。同名但内容不同的文件会自动加编号，不会覆盖已有文件。</p>
      <details className="folder-details">
        <summary>查看目录预览（最多显示 100 项）</summary>
        <ul>
          {preview.entries.map((e, i) => (
            <li key={i}>
              {e.kind === 'folder' ? <Folder size={15} /> : <FileText size={15} />}
              <span>{e.path}</span>
            </li>
          ))}
        </ul>
      </details>
      {preview.skipped > 0 && (
        <details className="folder-details">
          <summary>已跳过 {preview.skipped} 项（最多显示 100 项原因）</summary>
          <ul>
            {preview.warnings.map((w, i) => (
              <li key={i}>
                <span>
                  {w.name}：{w.reason}
                </span>
              </li>
            ))}
          </ul>
        </details>
      )}
      <p className="muted">导入过程中可以取消，已导入的文件会保留。再次导入时可跳过重复项。</p>
      <div className="modal-actions">
        <button className="secondary" onClick={onClose}>
          取消
        </button>
        <button className="primary" onClick={() => onImport(parent, duplicates)}>
          开始导入
        </button>
      </div>
    </Dialog>
  )
}

export function FolderResultDialog({
  result,
  onClose,
  onView,
}: {
  result: FolderImportResult
  onClose: () => void
  onView: () => void
}) {
  return (
    <Dialog title={result.cancelled ? '导入已取消' : '文件夹导入完成'} onClose={onClose}>
      <div className="folder-stats">
        <span>
          <strong>{result.imported.length}</strong> 份已导入
        </span>
        <span>{result.skipped} 项跳过</span>
        <span>{result.failures.length} 项失败</span>
      </div>
      <p>
        {result.cancelled
          ? '已完成的文件和分类已保留；尚未处理的项目没有导入。'
          : '文件已保存到本地文档库，原文件没有改变。'}
      </p>
      {result.failures.length > 0 && (
        <details className="folder-details" open>
          <summary>未能导入的项目</summary>
          <ul>
            {result.failures.map((f, i) => (
              <li key={i}>
                <span>
                  {f.name}：{f.reason}
                </span>
              </li>
            ))}
          </ul>
        </details>
      )}
      {(result.cancelled || result.failures.length > 0) && (
        <p className="muted">修正问题后，可重新选择同一文件夹导入，默认跳过已经导入的相同内容。</p>
      )}
      <div className="modal-actions">
        <button className="secondary" onClick={onClose}>
          关闭
        </button>
        <button className="primary" disabled={!result.imported.length} onClick={onView}>
          查看本批文件
        </button>
      </div>
    </Dialog>
  )
}
