import { useEffect, useState, type ReactNode, type RefObject } from 'react'
import {
  BookOpen,
  Check,
  ChevronDown,
  ChevronRight,
  Clock3,
  Copy,
  Download,
  FilePlus2,
  FileText,
  Folder,
  FolderOpen,
  Home,
  Highlighter,
  Info,
  LayoutGrid,
  List,
  MoreVertical,
  PanelRightClose,
  PanelRightOpen,
  Pencil,
  Plus,
  RotateCcw,
  Search,
  Settings,
  Star,
  Tag,
  Trash2,
  X,
} from 'lucide-react'
import type {
  Category,
  DocumentContent,
  DocumentRecord,
  LibraryApi,
  LibrarySnapshot,
  LibraryView,
  SearchHit,
  ExportScope,
} from '../../shared/types'
import { fileKind, canEditFile, fileTypeName } from '../../shared/file-types'
import { Highlight, locationLabel } from './OfficeReader'
import { ReadingPane } from './ReadingPane'
import { ResizableSplit } from './ResizableSplit'

type Patch = Parameters<LibraryApi['updateDocument']>[1]
export const fileType = (doc: DocumentRecord) => fileTypeName(doc.extension)
export const fileSize = (size: number) =>
  size < 1024
    ? `${size} B`
    : size < 1048576
      ? `${(size / 1024).toFixed(0)} KB`
      : `${(size / 1048576).toFixed(1)} MB`
export const fileDate = (date: string) => {
  const value = new Date(date)
  const pad = (part: number) => String(part).padStart(2, '0')
  return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())} ${pad(value.getHours())}:${pad(value.getMinutes())}`
}
export function categoryPath(categories: Category[], id: string | null): Category[] {
  const path: Category[] = []
  const visited = new Set<string>()
  while (id && !visited.has(id)) {
    visited.add(id)
    const category = categories.find((item) => item.id === id)
    if (!category) break
    path.unshift(category)
    id = category.parentId
  }
  return path
}
function FileIcon({ doc, large = false }: { doc: DocumentRecord; large?: boolean }) {
  if (!canEditFile(doc.extension))
    return (
      <span
        className={`document-icon extra-file-icon ${fileKind(doc.extension)} ${large ? 'large' : ''}`}
        aria-hidden="true"
      >
        <FileText />
        <small>{doc.extension.slice(1).toUpperCase().slice(0, 5) || 'FILE'}</small>
      </span>
    )
  return (
    <img
      className={`document-icon ${fileType(doc).toLowerCase()} ${large ? 'large' : ''}`}
      src={`./icons/${fileType(doc).toLowerCase()}.svg`}
      alt=""
    />
  )
}

interface Props {
  snapshot: LibrarySnapshot
  documents: DocumentRecord[]
  selected?: DocumentRecord
  preview: DocumentContent | null
  view: LibraryView
  viewLabel: string
  query: string
  type: string
  matches: SearchHit[] | null
  busy: boolean
  loading: boolean
  previewOpen: boolean
  batch: string[] | null
  batchOnly: boolean
  searchRef: RefObject<HTMLInputElement | null>
  onNavigate: (view: LibraryView) => void
  onQuery: (query: string) => void
  onType: (type: string) => void
  onSelect: (id: string) => void
  onPreviewToggle: () => void
  onImport: () => void
  onImportFolder: () => void
  onManageCategory: (category: Category) => void
  onMoveDocuments: (ids: string[]) => void
  selectionEpoch: number
  onCreateDocument: () => void
  onCreateOffice: (extension: '.docx' | '.xlsx') => void
  onCreateCategory: (parentId: string | null) => void
  onSettings: () => void
  onOpen: (doc: DocumentRecord, mode?: 'read' | 'edit') => void
  onRename: () => void
  onExport: (doc: DocumentRecord) => void
  onExportLibrary: (scope: ExportScope) => void
  onBatch: () => void
  onDismissBatch: () => void
  onAction: (action: () => Promise<void>) => Promise<void>
  onNotice: (notice: string) => void
  themeControl: ReactNode
}

function MetadataForm({ doc, onSave }: { doc: DocumentRecord; onSave: (patch: Patch) => void }) {
  const [tags, setTags] = useState(doc.tags.join('，'))
  const [notes, setNotes] = useState(doc.notes)
  return (
    <form
      className="metadata-form"
      onSubmit={(event) => {
        event.preventDefault()
        onSave({
          tags: [
            ...new Set(
              tags
                .split(/[,，]/)
                .map((tag) => tag.trim())
                .filter(Boolean),
            ),
          ],
          notes,
        })
      }}
    >
      <label className="form-label">
        标签
        <input
          aria-label="文档标签"
          value={tags}
          onChange={(event) => setTags(event.target.value)}
          placeholder="例如：项目资料，待整理"
          disabled={Boolean(doc.deletedAt)}
        />
      </label>
      <p className="muted">用逗号分隔，添加后可在顶部搜索。</p>
      <div className="tag-chips">
        {doc.tags.map((tag) => (
          <span key={tag}>
            <Tag size={13} />
            {tag}
          </span>
        ))}
      </div>
      <label className="form-label">
        备注
        <textarea
          aria-label="文档备注"
          value={notes}
          maxLength={5000}
          onChange={(event) => setNotes(event.target.value)}
          rows={6}
          placeholder="记录这份资料的用途或整理说明"
          disabled={Boolean(doc.deletedAt)}
        />
      </label>
      <button className="primary" type="submit" disabled={Boolean(doc.deletedAt)}>
        保存标签与备注
      </button>
    </form>
  )
}

export function LibraryWorkspace(p: Props) {
  const [layout, setLayout] = useState<'list' | 'grid'>('list')
  const [tab, setTab] = useState<'preview' | 'info' | 'tags'>('preview')
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set())
  const [checked, setChecked] = useState<Set<string>>(new Set())
  const [menu, setMenu] = useState<'files' | 'document' | null>(null)
  const [sort, setSort] = useState<'updated' | 'name' | 'size'>('updated')
  useEffect(() => {
    if (!menu) return
    const outside = (event: PointerEvent) => {
      if (event.target instanceof Element && !event.target.closest('.menu-anchor')) setMenu(null)
    }
    const escape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setMenu(null)
    }
    window.addEventListener('pointerdown', outside)
    window.addEventListener('keydown', escape)
    return () => {
      window.removeEventListener('pointerdown', outside)
      window.removeEventListener('keydown', escape)
    }
  }, [menu])
  const categoryId = p.view.startsWith('category:') ? p.view.slice(9) : null
  const currentPath = categoryPath(p.snapshot.categories, categoryId)
  const active = p.snapshot.documents.filter((doc) => !doc.deletedAt)
  const documents = [...p.documents].sort((a, b) =>
    sort === 'name' ? a.name.localeCompare(b.name, 'zh-CN') : sort === 'size' ? b.size - a.size : 0,
  )
  const checkedVisible = documents.filter((doc) => checked.has(doc.id))
  useEffect(() => {
    setChecked(new Set())
    setMenu(null)
  }, [p.view, p.query, p.type, p.selectionEpoch])
  useEffect(() => {
    setTab('preview')
    setMenu(null)
  }, [p.selected?.id])
  useEffect(() => {
    setCollapsed((value) => {
      const next = new Set(value)
      for (const category of categoryPath(p.snapshot.categories, categoryId))
        next.delete(category.id)
      return next
    })
  }, [categoryId, p.snapshot.categories])
  const update = (doc: DocumentRecord, patch: Patch) =>
    void p.onAction(async () => {
      await window.localDocs!.updateDocument(doc.id, patch)
      p.onNotice('已保存文档信息')
    })
  const trash = (docs: DocumentRecord[]) =>
    void p.onAction(async () => {
      for (const doc of docs) await window.localDocs!.trashDocument(doc.id)
      setChecked(new Set())
      p.onNotice('已移入回收站，可随时恢复')
    })
  const restore = (doc: DocumentRecord) =>
    void p.onAction(async () => {
      await window.localDocs!.restoreDocument(doc.id)
      p.onNotice('已恢复文档')
    })
  const toggleCheck = (id: string) =>
    setChecked((old) => {
      const next = new Set(old)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
  const tree = (parentId: string | null = null, level = 0): ReactNode =>
    p.snapshot.categories
      .filter((item) => item.parentId === parentId)
      .map((category) => {
        const children = p.snapshot.categories.some((item) => item.parentId === category.id)
        const count = active.filter((doc) => doc.categoryId === category.id).length
        return (
          <div key={category.id}>
            <div
              className={`category-row ${categoryId === category.id ? 'active' : ''}`}
              style={{ marginLeft: `calc(${level} * var(--tree-step))`, paddingLeft: '10px' }}
            >
              {children ? (
                <button
                  className="tree-toggle icon-button"
                  aria-label={`${collapsed.has(category.id) ? '展开' : '折叠'} ${category.name}`}
                  aria-expanded={!collapsed.has(category.id)}
                  onClick={() =>
                    setCollapsed((old) => {
                      const next = new Set(old)
                      if (next.has(category.id)) next.delete(category.id)
                      else next.add(category.id)
                      return next
                    })
                  }
                >
                  {collapsed.has(category.id) ? (
                    <ChevronRight size={15} />
                  ) : (
                    <ChevronDown size={15} />
                  )}
                </button>
              ) : (
                <span className="tree-indent" />
              )}
              <button
                className="category-name"
                onClick={() => p.onNavigate(`category:${category.id}`)}
              >
                <Folder size={21} fill="currentColor" />
                <span>{category.name}</span>
                {count > 0 && <em>{count}</em>}
              </button>
              <button
                className="icon-button category-manage"
                aria-label={`管理分类 ${category.name}`}
                title="管理分类"
                disabled={p.busy}
                onClick={() => p.onManageCategory(category)}
              >
                <MoreVertical size={16} />
              </button>
            </div>
            {!collapsed.has(category.id) && tree(category.id, level + 1)}
          </div>
        )
      })
  const navItems = [
    { id: 'all', name: '全部文件', Icon: Folder, count: active.length },
    {
      id: 'recent',
      name: '最近使用',
      Icon: Clock3,
      count: active.filter((doc) => doc.openedAt).length,
    },
    {
      id: 'favorites',
      name: '收藏',
      Icon: Star,
      count: active.filter((doc) => doc.favorite).length,
    },
    {
      id: 'trash',
      name: '回收站',
      Icon: Trash2,
      count: p.snapshot.documents.filter((doc) => doc.deletedAt).length,
    },
    {
      id: 'marked',
      name: '内容标记',
      Icon: Highlighter,
      count: active.filter((d) => p.snapshot.annotationCounts?.[d.id]).length,
    },
  ] as const
  const excerpt = (doc: DocumentRecord) => {
    const hit = p.matches?.find((item) => item.id === doc.id && item.source !== 'name')
    return hit ? (
      <span className="search-excerpt">
        <small>
          {hit.source === 'metadata' ? '标签 / 备注' : locationLabel(hit.location) || '正文命中'}
        </small>
        <Highlight text={hit.snippet} query={p.query.trim()} />
      </span>
    ) : null
  }
  return (
    <div className="library-shell">
      <header className="library-topbar">
        <div className="library-brand">
          <img src="./icons/app.svg" alt="" />
          <div>
            <strong>我的文档库</strong>
            <span>本地文档管理工具 · 简单 · 高效</span>
          </div>
        </div>
        <label className="global-search">
          <Search size={20} />
          <input
            ref={p.searchRef}
            value={p.query}
            maxLength={200}
            onChange={(event) => p.onQuery(event.target.value)}
            placeholder="搜索文件名、内容或标签…"
            aria-label="搜索文档"
          />
          {p.query && (
            <button className="icon-button" aria-label="清空搜索" onClick={() => p.onQuery('')}>
              <X size={16} />
            </button>
          )}
        </label>
        <div className="library-top-actions">
          <button className="primary" disabled={p.busy || p.view === 'trash'} onClick={p.onImport}>
            <Plus size={21} />
            添加文件
          </button>
          <button
            className="secondary"
            disabled={p.busy}
            onClick={() => p.onCreateCategory(categoryId)}
          >
            <FolderOpen size={20} />
            新建文件夹
          </button>
          <div className="view-switch" role="group" aria-label="视图方式">
            <button
              aria-label="网格视图"
              aria-pressed={layout === 'grid'}
              onClick={() => setLayout('grid')}
            >
              <LayoutGrid size={21} />
            </button>
            <button
              aria-label="列表视图"
              aria-pressed={layout === 'list'}
              onClick={() => setLayout('list')}
            >
              <List size={22} />
            </button>
          </div>
          {p.themeControl}
        </div>
      </header>
      <aside className="library-sidebar">
        <nav aria-label="文档导航">
          {navItems.map(({ id, name, Icon, count }) => (
            <button
              key={id}
              className={`nav-item ${p.view === id || (id === 'all' && categoryId) ? 'active' : ''}`}
              onClick={() => p.onNavigate(id)}
            >
              <Icon size={22} fill={id === 'all' ? 'currentColor' : 'none'} />
              <span>{name}</span>
              <em>{count}</em>
            </button>
          ))}
        </nav>
        <div className="category-heading">
          <span>我的分类</span>
          <button
            className="icon-button"
            aria-label="新建分类"
            title="新建顶级分类"
            disabled={p.busy}
            onClick={() => p.onCreateCategory(null)}
          >
            <Plus size={22} />
          </button>
        </div>
        <div className="category-list">{tree()}</div>
        <button className="nav-item settings-nav" onClick={p.onSettings}>
          <Settings size={22} />
          <span>设置</span>
        </button>
      </aside>
      <main className={`library-content ${p.previewOpen ? '' : 'without-preview'}`}>
        <ResizableSplit
          className="library-split"
          preferenceKey="local-docs:library-preview-ratio"
          defaultRatio={0.616}
          minLeft={330}
          minRight={330}
          leftLabel="文件列表"
          rightLabel="文档预览"
          secondaryVisible={p.previewOpen}
        >
          <section className="file-panel" aria-label="文件列表">
            <header className="files-toolbar">
              <div className="breadcrumb">
                <button aria-label="全部文件" onClick={() => p.onNavigate('all')}>
                  <Home size={18} fill="currentColor" />
                </button>
                {currentPath.length ? (
                  currentPath.map((category, index) => (
                    <span key={category.id}>
                      <ChevronRight size={15} />
                      {index === currentPath.length - 1 ? (
                        <h1>{category.name}</h1>
                      ) : (
                        <button onClick={() => p.onNavigate(`category:${category.id}`)}>
                          {category.name}
                        </button>
                      )}
                    </span>
                  ))
                ) : (
                  <h1>{p.batchOnly ? '本批添加的文件' : p.viewLabel}</h1>
                )}
              </div>
              <div className="menu-anchor">
                <button
                  className="icon-button"
                  aria-label="列表选项"
                  onClick={() => setMenu(menu === 'files' ? null : 'files')}
                >
                  <MoreVertical size={20} />
                </button>
                {menu === 'files' && (
                  <div className="popover">
                    <button
                      onClick={() => {
                        p.onCreateDocument()
                        setMenu(null)
                      }}
                      disabled={p.view === 'trash'}
                    >
                      <FilePlus2 size={17} />
                      新建 Markdown
                    </button>
                    <button
                      disabled={p.busy || p.view === 'trash'}
                      onClick={() => {
                        p.onCreateOffice('.docx')
                        setMenu(null)
                      }}
                    >
                      <FilePlus2 size={17} />
                      新建 Word
                    </button>
                    <button
                      disabled={p.busy || p.view === 'trash'}
                      onClick={() => {
                        p.onCreateOffice('.xlsx')
                        setMenu(null)
                      }}
                    >
                      <FilePlus2 size={17} />
                      新建 Excel
                    </button>
                    <button
                      disabled={p.busy}
                      onClick={() => {
                        p.onImportFolder()
                        setMenu(null)
                      }}
                    >
                      <FolderOpen size={17} />
                      导入文件夹
                    </button>
                    <button
                      disabled={p.busy}
                      onClick={() => {
                        p.onExportLibrary(
                          categoryId ? { kind: 'category', id: categoryId } : { kind: 'all' },
                        )
                        setMenu(null)
                      }}
                    >
                      <Download size={17} />
                      {categoryId ? '导出此分类（含子分类）' : '导出整个文档库'}
                    </button>
                    <button
                      onClick={() => {
                        p.onPreviewToggle()
                        setMenu(null)
                      }}
                    >
                      {p.previewOpen ? <PanelRightClose size={17} /> : <PanelRightOpen size={17} />}
                      {p.previewOpen ? '收起预览' : '展开预览'}
                    </button>
                    <label>
                      类型
                      <select
                        aria-label="文件类型"
                        value={p.type}
                        onChange={(event) => p.onType(event.target.value)}
                      >
                        <option value="all">全部类型</option>
                        <option>Word</option>
                        <option>Excel</option>
                        <option>Markdown</option>
                        {['PDF', 'TXT', 'SQL', '文本', '图片', '其他'].map((type) => (
                          <option key={type}>{type}</option>
                        ))}
                      </select>
                    </label>
                    <label>
                      排序
                      <select
                        aria-label="排序方式"
                        value={sort}
                        onChange={(event) => setSort(event.target.value as typeof sort)}
                      >
                        <option value="updated">
                          {p.view === 'recent' ? '最近打开' : '最近修改'}
                        </option>
                        <option value="name">名称</option>
                        <option value="size">大小</option>
                      </select>
                    </label>
                  </div>
                )}
              </div>
            </header>
            {p.batch && (
              <div className="batch-note">
                <Check size={15} />
                <span>已添加 {p.batch.length} 份文件</span>
                <button onClick={p.onBatch}>{p.batchOnly ? '返回全部文件' : '查看本批文件'}</button>
                <button
                  className="icon-button"
                  aria-label="关闭导入提示"
                  onClick={p.onDismissBatch}
                >
                  <X size={15} />
                </button>
              </div>
            )}
            {checkedVisible.length > 0 && (
              <div className="selection-toolbar">
                <span>已选择 {checkedVisible.length} 项</span>
                <button
                  className="text-button"
                  disabled={p.busy || p.view === 'trash'}
                  onClick={() =>
                    void p.onAction(async () => {
                      await window.localDocs!.setFavorites(
                        checkedVisible.map((d) => d.id),
                        true,
                      )
                      p.onNotice('已收藏选中的文件')
                    })
                  }
                >
                  <Star size={15} />
                  批量收藏
                </button>
                {p.view === 'favorites' && (
                  <button
                    className="text-button"
                    disabled={p.busy}
                    onClick={() =>
                      void p.onAction(async () => {
                        await window.localDocs!.setFavorites(
                          checkedVisible.map((d) => d.id),
                          false,
                        )
                        setChecked(new Set())
                        p.onNotice('已取消收藏')
                      })
                    }
                  >
                    取消所选收藏
                  </button>
                )}
                <button
                  className="text-button"
                  disabled={p.busy || p.view === 'trash'}
                  onClick={() => p.onMoveDocuments(checkedVisible.map((d) => d.id))}
                >
                  <FolderOpen size={15} />
                  移动到分类
                </button>
                <button
                  className="text-button"
                  disabled={p.busy || p.view === 'trash'}
                  onClick={() => trash(checkedVisible)}
                >
                  <Trash2 size={15} />
                  移入回收站
                </button>
                <button
                  className="icon-button"
                  aria-label="取消选择"
                  onClick={() => setChecked(new Set())}
                >
                  <X size={15} />
                </button>
              </div>
            )}
            <div className="files-scroll">
              {p.loading ? (
                <div className="empty-state">
                  <p>正在读取文档…</p>
                </div>
              ) : !documents.length ? (
                <div className="empty-state">
                  <FolderOpen size={54} strokeWidth={1.1} />
                  <h2>
                    {p.query || p.type !== 'all'
                      ? '没有找到匹配的文件'
                      : p.view === 'trash'
                        ? '回收站是空的'
                        : '这里还没有文件'}
                  </h2>
                  <p>
                    {p.query
                      ? '试试其他关键词，或清除搜索。'
                      : '点击添加，或直接拖入 PDF、TXT、SQL 等文件。'}
                  </p>
                  {p.view !== 'trash' && (
                    <button className="primary" onClick={p.onImport} disabled={p.busy}>
                      <Plus size={18} />
                      添加第一份文件
                    </button>
                  )}
                </div>
              ) : layout === 'list' ? (
                <table className="file-table">
                  <colgroup>
                    <col className="check-column" />
                    <col className="name-column" />
                    <col className="type-column" />
                    <col className="date-column" />
                    <col className="size-column" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th>
                        <input
                          type="checkbox"
                          aria-label="选择全部文件"
                          checked={checkedVisible.length === documents.length}
                          onChange={(event) =>
                            setChecked(
                              new Set(event.target.checked ? documents.map((doc) => doc.id) : []),
                            )
                          }
                        />
                      </th>
                      <th>名称</th>
                      <th>类型</th>
                      <th>修改时间</th>
                      <th>大小</th>
                    </tr>
                  </thead>
                  <tbody>
                    {documents.map((doc) => (
                      <tr
                        key={doc.id}
                        tabIndex={0}
                        aria-selected={p.selected?.id === doc.id}
                        className={p.selected?.id === doc.id ? 'selected' : ''}
                        onClick={() => p.onSelect(doc.id)}
                        onDoubleClick={() => !p.busy && p.onOpen(doc)}
                        onKeyDown={(event) => {
                          if (event.target !== event.currentTarget) return
                          if (event.key === 'Enter' && !p.busy) p.onOpen(doc)
                          if (event.key === ' ') {
                            event.preventDefault()
                            p.onSelect(doc.id)
                          }
                        }}
                      >
                        <td
                          onClick={(event) => event.stopPropagation()}
                          onDoubleClick={(event) => event.stopPropagation()}
                        >
                          <input
                            type="checkbox"
                            aria-label={`选择 ${doc.name}`}
                            checked={checked.has(doc.id)}
                            onChange={() => toggleCheck(doc.id)}
                          />
                        </td>
                        <td>
                          <div className="file-name">
                            <FileIcon doc={doc} />
                            <div>
                              <strong>
                                <Highlight text={doc.name} query={p.query.trim()} />
                              </strong>
                              {excerpt(doc)}
                            </div>
                            <button
                              className="icon-button file-favorite"
                              aria-label={`${doc.favorite ? '取消收藏' : '收藏'} ${doc.name}`}
                              aria-pressed={doc.favorite}
                              disabled={p.busy || !!doc.deletedAt}
                              onClick={(e) => {
                                e.stopPropagation()
                                update(doc, { favorite: !doc.favorite })
                              }}
                              onDoubleClick={(e) => e.stopPropagation()}
                            >
                              <Star size={16} fill={doc.favorite ? 'currentColor' : 'none'} />
                            </button>
                          </div>
                        </td>
                        <td>{fileType(doc)}</td>
                        <td className="date-cell">{fileDate(doc.updatedAt)}</td>
                        <td className="size-cell">{fileSize(doc.size)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <div className="file-grid">
                  {documents.map((doc) => (
                    <button
                      key={doc.id}
                      className={`file-card ${p.selected?.id === doc.id ? 'selected' : ''}`}
                      onClick={() => p.onSelect(doc.id)}
                      onDoubleClick={() => !p.busy && p.onOpen(doc)}
                    >
                      <FileIcon doc={doc} large />
                      <strong>{doc.name}</strong>
                      <span>
                        {fileType(doc)} · {fileSize(doc.size)}
                      </span>
                    </button>
                  ))}
                </div>
              )}
            </div>
            <footer className="list-footer">
              <span>
                共 {documents.length} 个文件{p.query ? ` · 搜索“${p.query}”` : ''}
              </span>
              {!p.previewOpen && (
                <button className="icon-button" aria-label="展开预览" onClick={p.onPreviewToggle}>
                  <PanelRightOpen size={18} />
                </button>
              )}
            </footer>
          </section>
          {p.previewOpen && (
            <aside className="preview-panel" aria-label="文档预览">
              {p.selected ? (
                <>
                  <div className="preview-info">
                    <FileIcon doc={p.selected} large />
                    <div>
                      <h2>{p.selected.name}</h2>
                      <span>{fileSize(p.selected.size)}</span>
                      <p>修改时间：{fileDate(p.selected.updatedAt)}</p>
                    </div>
                    <div className="menu-anchor">
                      <button
                        className="icon-button"
                        aria-label="文档选项"
                        onClick={() => setMenu(menu === 'document' ? null : 'document')}
                      >
                        <MoreVertical size={20} />
                      </button>
                      {menu === 'document' && (
                        <div className="popover">
                          <button
                            onClick={() => {
                              p.onRename()
                              setMenu(null)
                            }}
                            disabled={Boolean(p.selected.deletedAt)}
                          >
                            <Pencil size={16} />
                            重命名
                          </button>
                          <button
                            onClick={() => {
                              update(p.selected!, { favorite: !p.selected!.favorite })
                              setMenu(null)
                            }}
                            disabled={Boolean(p.selected.deletedAt)}
                          >
                            <Star size={16} />
                            {p.selected.favorite ? '取消收藏' : '收藏'}
                          </button>
                          <button
                            onClick={() => {
                              p.onExport(p.selected!)
                              setMenu(null)
                            }}
                          >
                            <Download size={16} />
                            导出
                          </button>
                          <button
                            onClick={() => {
                              p.onPreviewToggle()
                              setMenu(null)
                            }}
                          >
                            <PanelRightClose size={16} />
                            收起预览
                          </button>
                        </div>
                      )}
                    </div>
                  </div>
                  <div className="preview-tabs" role="tablist" aria-label="文档详情">
                    {(
                      [
                        { id: 'preview', name: '预览', Icon: FileText },
                        { id: 'info', name: '信息', Icon: Info },
                        { id: 'tags', name: '标签', Icon: Tag },
                      ] as const
                    ).map(({ id, name, Icon }) => (
                      <button
                        key={id}
                        role="tab"
                        aria-selected={tab === id}
                        onClick={() => setTab(id)}
                      >
                        <Icon size={17} />
                        {name}
                      </button>
                    ))}
                  </div>
                  <div
                    className={`preview-body ${tab === 'preview' ? 'preview-paper-frame' : ''}`}
                    role="tabpanel"
                    aria-label={tab === 'preview' ? '预览' : tab === 'info' ? '信息' : '标签'}
                  >
                    {tab === 'preview' ? (
                      p.preview ? (
                        <ReadingPane
                          key={p.selected.id}
                          doc={p.selected}
                          text={p.preview.text}
                          compact
                          showMarks={p.view === 'marked'}
                          query={p.query.trim()}
                          onChanged={() => void p.onAction(async () => {})}
                        />
                      ) : (
                        <p className="muted">正在读取…</p>
                      )
                    ) : tab === 'info' ? (
                      <div className="document-details">
                        <dl>
                          <dt>文件名称</dt>
                          <dd>{p.selected.name}</dd>
                          <dt>类型</dt>
                          <dd>
                            {fileType(p.selected)} · {p.selected.extension}
                          </dd>
                          <dt>大小</dt>
                          <dd>{fileSize(p.selected.size)}</dd>
                          <dt>创建时间</dt>
                          <dd>{fileDate(p.selected.createdAt)}</dd>
                          <dt>修改时间</dt>
                          <dd>{fileDate(p.selected.updatedAt)}</dd>
                        </dl>
                        <label className="form-label">
                          所在分类
                          <select
                            aria-label="所在分类"
                            value={p.selected.categoryId || ''}
                            disabled={p.busy || Boolean(p.selected.deletedAt)}
                            onChange={(event) =>
                              update(p.selected!, { categoryId: event.target.value || null })
                            }
                          >
                            <option value="">不指定分类</option>
                            {p.snapshot.categories.map((category) => (
                              <option key={category.id} value={category.id}>
                                {categoryPath(p.snapshot.categories, category.id)
                                  .map((item) => item.name)
                                  .join(' / ')}
                              </option>
                            ))}
                          </select>
                        </label>
                        <button
                          className="secondary"
                          disabled={p.busy || Boolean(p.selected.deletedAt)}
                          onClick={() => update(p.selected!, { favorite: !p.selected!.favorite })}
                        >
                          <Star size={17} fill={p.selected.favorite ? 'currentColor' : 'none'} />
                          {p.selected.favorite ? '取消收藏' : '收藏'}
                        </button>
                        <p className="muted">文件副本保存在文档库中，导入来源不受影响。</p>
                      </div>
                    ) : (
                      <MetadataForm
                        key={`${p.selected.id}:${JSON.stringify(p.selected.tags)}:${p.selected.notes}`}
                        doc={p.selected}
                        onSave={(patch) => update(p.selected!, patch)}
                      />
                    )}
                  </div>
                  <div className="preview-actions">
                    {p.selected.deletedAt ? (
                      <button
                        className="primary"
                        disabled={p.busy}
                        onClick={() => restore(p.selected!)}
                      >
                        <RotateCcw size={18} />
                        恢复文档
                      </button>
                    ) : (
                      <>
                        <button
                          className="primary"
                          disabled={p.busy}
                          onClick={() => p.onOpen(p.selected!)}
                        >
                          <FolderOpen size={19} />
                          打开
                        </button>
                        {canEditFile(p.selected.extension) && (
                          <button
                            className="secondary"
                            disabled={p.busy}
                            title={`编辑 ${fileType(p.selected)}`}
                            onClick={() => p.onOpen(p.selected!, 'edit')}
                          >
                            <Pencil size={18} />
                            编辑
                          </button>
                        )}
                      </>
                    )}
                    <button
                      className="secondary copy-path"
                      aria-label="复制路径"
                      disabled={p.busy}
                      title="复制文档库内副本的实际存储路径"
                      onClick={() =>
                        void p.onAction(async () => {
                          await window.localDocs!.copyDocumentPath(p.selected!.id)
                          p.onNotice('已复制文档库内副本路径')
                        })
                      }
                    >
                      <Copy size={18} />
                      <span>复制路径</span>
                    </button>
                    {!p.selected.deletedAt && (
                      <button
                        className="danger-outline"
                        aria-label="移入回收站"
                        disabled={p.busy}
                        onClick={() => trash([p.selected!])}
                      >
                        <Trash2 size={18} />
                        删除
                      </button>
                    )}
                  </div>
                </>
              ) : (
                <>
                  <div className="preview-empty-heading">
                    文档预览
                    <button
                      className="icon-button"
                      aria-label="关闭预览"
                      onClick={p.onPreviewToggle}
                    >
                      <X size={17} />
                    </button>
                  </div>
                  <div className="preview-empty">
                    <BookOpen size={45} strokeWidth={1.2} />
                    <h3>选择文件以预览</h3>
                    <p>单击查看内容，双击打开文档。</p>
                  </div>
                </>
              )}
            </aside>
          )}
        </ResizableSplit>
      </main>
    </div>
  )
}
