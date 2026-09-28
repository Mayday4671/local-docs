import { useCallback, useEffect, useRef, useState, type FormEvent } from 'react'
import { ArrowLeft, Download, HardDrive, History, Library, Save, Star, X } from 'lucide-react'
import { ThemeSwitcher } from './ThemeSwitcher'
import { StorageActions, TransferDialogs } from './StorageTools'
import {
  CategoryDialog,
  MoveDialog,
  FolderPreviewDialog,
  FolderResultDialog,
  categoryDescendants,
} from './OrganizationDialogs'
import { canEditFile, canCompareFile, fileKind, fileTypeName } from '../../shared/file-types'
import { FileDropZone } from './FileDropZone'
import { OfficeEditor } from './OfficeEditor'
import { VersionComparison } from './VersionComparison'
import { EditorSplit } from './EditorSplit'
import { ReadingPane } from './ReadingPane'
import { LibraryWorkspace, categoryPath } from './LibraryWorkspace'
import type {
  DocumentContent,
  DocumentRecord,
  LibrarySnapshot,
  LibraryView,
  VersionRecord,
  SearchHit,
  ContentLocation,
  ThemePreference,
  BackupPreview,
  ExportScope,
  OperationStatus,
  TransferResult,
  Category,
  FolderImportPreview,
  FolderImportResult,
  Draft,
  Attachment,
} from '../../shared/types'

const api = window.localDocs
const emptySnapshot: LibrarySnapshot = { documents: [], categories: [], storagePath: '' }
const views = [
  { id: 'all', label: '全部文件', icon: Library },
  { id: 'recent', label: '最近使用' },
  { id: 'favorites', label: '收藏' },
  { id: 'marked', label: '内容标记' },
] as const
const isMarkdown = (doc: DocumentRecord) => ['.md', '.markdown'].includes(doc.extension)
const typeName = (doc: DocumentRecord) => fileTypeName(doc.extension)
export function App() {
  const [theme, setTheme] = useState<ThemePreference>('system')
  const [themeReady, setThemeReady] = useState(false)
  const [themeSaving, setThemeSaving] = useState(false)
  useEffect(() => {
    void api
      ?.getTheme()
      .then(setTheme)
      .catch((error) => setError(String(error)))
      .finally(() => setThemeReady(true))
  }, [])
  const [snapshot, setSnapshot] = useState(emptySnapshot)
  const [loading, setLoading] = useState(true)
  const [view, setView] = useState<LibraryView>('all')
  const [query, setQuery] = useState('')
  const [matches, setMatches] = useState<SearchHit[] | null>(null)
  const [searching, setSearching] = useState(false)
  const [type, setType] = useState('all')
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [preview, setPreview] = useState<DocumentContent | null>(null)
  const [previewOpen, setPreviewOpen] = useState(true)
  const [busy, setBusy] = useState(false)
  const [transferring, setTransferring] = useState(false)
  const [transferProgress, setTransferProgress] = useState<OperationStatus | null>(null)
  const [backupPreview, setBackupPreview] = useState<BackupPreview | null>(null)
  const [transferResult, setTransferResult] = useState<TransferResult | null>(null)
  const [managedCategory, setManagedCategory] = useState<Category | null>(null)
  const [moveIds, setMoveIds] = useState<string[] | null>(null)
  const [selectionEpoch, setSelectionEpoch] = useState(0)
  const [folderPreview, setFolderPreview] = useState<FolderImportPreview | null>(null)
  const [folderResult, setFolderResult] = useState<FolderImportResult | null>(null)
  const organizing = !!(managedCategory || moveIds || folderPreview || folderResult)
  const [error, setError] = useState('')
  const [noticeState, setNoticeState] = useState({ text: '', sequence: 0 })
  const notice = noticeState.text
  const setNotice = useCallback((text: string) => {
    setNoticeState((previous) => ({ text, sequence: previous.sequence + 1 }))
  }, [])
  useEffect(() => {
    // Each completed action gets its own timer, even when the message repeats.
    // Errors and progress stay visible; an older timer cannot dismiss a newer notice.
    if (!noticeState.text || busy || error) return
    const timer = setTimeout(() => {
      setNoticeState((previous) =>
        previous.sequence === noticeState.sequence ? { ...previous, text: '' } : previous,
      )
    }, 3000)
    return () => clearTimeout(timer)
  }, [noticeState, busy, error])
  const [batch, setBatch] = useState<string[] | null>(null)
  const [batchOnly, setBatchOnly] = useState(false)
  const [editor, setEditor] = useState<{
    doc: DocumentRecord
    text: string
    savedText: string
    location?: ContentLocation
    initialMode?: 'read' | 'edit'
    readOnlyText?: string | null
  } | null>(null)
  const [history, setHistory] = useState<VersionRecord[] | null>(null)
  const [comparedVersion, setComparedVersion] = useState<string | null>(null)
  const [recovery, setRecovery] = useState<Draft | null>(null)
  const [draftStatus, setDraftStatus] = useState('')
  const [editorImages, setEditorImages] = useState<Attachment[]>([])
  const draftWrite = useRef<Promise<unknown>>(Promise.resolve())
  const sourceRef = useRef<HTMLTextAreaElement>(null)
  const [modal, setModal] = useState<
    'document' | 'category' | 'rename' | 'settings' | 'close' | null
  >(null)
  const [name, setName] = useState('')
  const [newExtension, setNewExtension] = useState<'.docx' | '.xlsx' | null>(null)
  const [createParent, setCreateParent] = useState<string | null>(null)
  const searchRef = useRef<HTMLInputElement>(null)
  const selected = snapshot.documents.find(
    (doc) => doc.id === selectedId && (view === 'trash' ? Boolean(doc.deletedAt) : !doc.deletedAt),
  )
  const dirty = Boolean(editor && editor.text !== editor.savedText)
  useEffect(() => {
    if (!editor || busy || recovery) return
    if (!dirty) {
      if (!draftStatus) return
      let active = true
      void draftWrite.current
        .then(() => api!.discardDraft(editor.doc.id))
        .then(() => {
          if (active) setDraftStatus('')
        })
        .catch((e) => {
          if (active) setDraftStatus(`草稿清理失败：${String(e)}`)
        })
      return () => {
        active = false
      }
    }
    let active = true
    setDraftStatus('等待保存恢复草稿…')
    const timer = setTimeout(() => {
      const task = api!.saveDraft(
        editor.doc.id,
        editor.text,
        editor.doc.revision,
        isMarkdown(editor.doc) ? 'markdown' : 'office',
      )
      draftWrite.current = task.catch(() => {})
      void task
        .then(() => {
          if (active) setDraftStatus('恢复草稿已自动保存')
        })
        .catch((e) => {
          if (active) setDraftStatus(`草稿保存失败：${String(e)}`)
        })
    }, 800)
    return () => {
      active = false
      clearTimeout(timer)
    }
  }, [editor, dirty, busy, recovery])
  const categoryId = view.startsWith('category:') ? view.slice(9) : null
  const viewLabel =
    view === 'trash'
      ? '回收站'
      : views.find((item) => item.id === view)?.label ||
        snapshot.categories.find((item) => item.id === categoryId)?.name ||
        '全部文件'

  const refresh = useCallback(async () => {
    if (api) setSnapshot(await api.snapshot())
  }, [])
  useEffect(() => {
    void refresh()
      .catch((error) => setError(String(error)))
      .finally(() => setLoading(false))
  }, [refresh])

  useEffect(() => {
    let cancelled = false
    if (!query.trim()) {
      setMatches(null)
      setSearching(false)
      return
    }
    setMatches([])
    setSearching(true)
    const timeout = setTimeout(() => {
      void api
        ?.searchResults(query)
        .then((ids) => {
          if (!cancelled) setMatches(ids)
        })
        .catch((error) => {
          if (!cancelled) setError(String(error))
        })
        .finally(() => {
          if (!cancelled) setSearching(false)
        })
    }, 180)
    return () => {
      cancelled = true
      clearTimeout(timeout)
    }
  }, [query, snapshot])

  useEffect(() => {
    let cancelled = false
    setPreview((current) => (current?.document.id === selectedId ? current : null))
    if (selectedId)
      void api
        ?.readDocument(selectedId)
        .then((value) => {
          if (!cancelled) setPreview(value)
        })
        .catch((error) => {
          if (!cancelled) setError(String(error))
        })
    return () => {
      cancelled = true
    }
  }, [selectedId, snapshot])

  useEffect(() => {
    const listener = (event: BeforeUnloadEvent) => {
      if (dirty) {
        event.preventDefault()
        event.returnValue = ''
      }
    }
    window.addEventListener('beforeunload', listener)
    return () => window.removeEventListener('beforeunload', listener)
  }, [dirty])

  const perform = async (action: () => Promise<void>) => {
    setBusy(true)
    setError('')
    try {
      await action()
      await refresh()
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

  useEffect(() => {
    if (!transferring) return
    let stopped = false
    const poll = () =>
      void api
        ?.operationStatus()
        .then((value) => {
          if (!stopped) setTransferProgress(value)
        })
        .catch(() => {})
    poll()
    const timer = setInterval(poll, 250)
    return () => {
      stopped = true
      clearInterval(timer)
    }
  }, [transferring])
  const transfer = async (action: () => Promise<void>) => {
    if (busy) return
    setBusy(true)
    setTransferring(true)
    setError('')
    setNotice('')
    setTransferProgress(null)
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
      setTransferring(false)
      setTransferProgress(null)
    }
  }
  const backupLibrary = () =>
    void transfer(async () => {
      const result = await api!.createBackup()
      if (result) {
        setTransferResult(result)
        setNotice('完整备份已保存')
        setModal('settings')
      }
    })
  const previewBackup = () =>
    void transfer(async () => {
      const result = await api!.previewBackup()
      if (result) setBackupPreview(result)
    })
  const exportLibrary = (scope: ExportScope) =>
    void transfer(async () => {
      const result = await api!.exportLibrary(scope)
      if (result) {
        setTransferResult(result)
        setNotice(`已按分类导出 ${result.files} 份文件`)
        setModal('settings')
      }
    })
  const restoreBackup = () => {
    if (!backupPreview) return
    const token = backupPreview.token
    setBackupPreview(null)
    void transfer(async () => {
      const result = await api!.restoreBackup(token)
      setSelectedId(null)
      setPreview(null)
      setEditor(null)
      setHistory(null)
      setView('all')
      setQuery('')
      setType('all')
      setBatch(null)
      setBatchOnly(false)
      setMatches(null)
      setTheme(await api!.getTheme())
      await refresh()
      setTransferResult(result)
      setNotice('文档库已恢复，恢复前的资料已自动备份')
      setModal('settings')
    })
  }

  const changeTheme = async (value: ThemePreference) => {
    setThemeSaving(true)
    try {
      await api!.setTheme(value)
      setTheme(value)
    } catch (error) {
      setError(error instanceof Error ? error.message : String(error))
    } finally {
      setThemeSaving(false)
    }
  }
  const themeControl = (
    <ThemeSwitcher
      theme={theme}
      disabled={!themeReady || themeSaving}
      onChange={(value) => void changeTheme(value)}
    />
  )

  const open = (doc: DocumentRecord, initialMode: 'read' | 'edit' = 'read') => {
    if (doc.deletedAt) return
    void perform(async () => {
      const content = await api!.openDocument(doc.id)
      setRecovery(await api!.draft(doc.id))
      setEditorImages(await api!.attachments(doc.id))
      setDraftStatus('')
      setEditor({
        doc: content.document,
        readOnlyText: content.text,
        text: content.text ?? '',
        savedText: content.text ?? '',
        location: matches?.find((hit) => hit.id === doc.id)?.location,
        initialMode,
      })
      setHistory(null)
    })
  }
  const save = () => {
    if (!editor || !canEditFile(editor.doc.extension) || !dirty || busy) return
    void perform(async () => {
      await draftWrite.current
      const doc = isMarkdown(editor.doc)
        ? await api!.saveMarkdown(editor.doc.id, editor.text, editor.doc.revision)
        : await api!.saveOffice(
            editor.doc.id,
            editor.text ? JSON.parse(editor.text) : [],
            editor.doc.revision,
          )
      const text = isMarkdown(doc) ? editor.text : ''
      setEditor({ doc, text, savedText: text })
      setDraftStatus('')
      setRecovery(null)
      setNotice('已保存到本地文档库')
      if (history) setHistory(await api!.versions(doc.id))
    })
  }

  useEffect(() => {
    const listener = (event: KeyboardEvent) => {
      if (comparedVersion) {
        if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's')
          event.preventDefault()
        return
      }
      if (organizing || transferring || backupPreview || modal) return
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's' && editor) {
        event.preventDefault()
        save()
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k' && !editor) {
        event.preventDefault()
        searchRef.current?.focus()
      }
    }
    window.addEventListener('keydown', listener)
    return () => window.removeEventListener('keydown', listener)
  })

  const navigate = (next: LibraryView) => {
    setView(next)
    setBatchOnly(false)
    setSelectedId(null)
  }
  const importFiles = (files?: File[]) =>
    void perform(async () => {
      const result = files
        ? await window.localFileImport!(files, categoryId)
        : await api!.importFiles(categoryId)
      if (!result.imported.length && !result.failures.length) return
      setNotice(
        `已添加 ${result.imported.length} 份文件${result.failures.length ? `，${result.failures.length} 份未能添加` : ''}`,
      )
      if (result.imported.length) {
        setView(categoryId ? `category:${categoryId}` : 'all')
        setQuery('')
        setType('all')
        setBatchOnly(false)
        setBatch(result.imported)
        setSelectedId(result.imported[0])
        setPreviewOpen(true)
      }
      if (result.failures.length)
        setError(result.failures.map((item) => `${item.name}：${item.reason}`).join('\n'))
    })
  const showNameModal = (kind: 'document' | 'category' | 'rename') => {
    setName(kind === 'rename' ? selected?.name || '' : '')
    setModal(kind)
  }
  const submitName = (event: FormEvent) => {
    event.preventDefault()
    void perform(async () => {
      if (modal === 'category') {
        const category = await api!.createCategory(name, createParent)
        navigate(`category:${category.id}`)
      } else if (modal === 'rename' && selected) await api!.updateDocument(selected.id, { name })
      else if (modal === 'document') {
        const doc = newExtension
          ? await api!.createOffice(name, newExtension, categoryId)
          : await api!.createMarkdown(name, categoryId)
        setSelectedId(doc.id)
        setPreviewOpen(true)
        const content = await api!.openDocument(doc.id)
        setEditor({ doc: content.document, text: '', savedText: '', initialMode: 'edit' })
        setRecovery(null)
        setEditorImages([])
        setDraftStatus('')
        setHistory(null)
      }
      setModal(null)
    })
  }
  const exportDoc = (doc: DocumentRecord) =>
    void perform(async () => {
      if (await api!.exportDocument(doc.id)) setNotice('已导出文档副本')
    })

  let documents = snapshot.documents.filter((doc) => {
    if (view === 'trash' ? !doc.deletedAt : Boolean(doc.deletedAt)) return false
    if (view === 'favorites' && !doc.favorite) return false
    if (view === 'marked' && !snapshot.annotationCounts?.[doc.id]) return false
    if (view === 'recent' && !doc.openedAt) return false
    if (categoryId && doc.categoryId !== categoryId) return false
    if (type !== 'all' && typeName(doc) !== type) return false
    if (matches && !matches.some((hit) => hit.id === doc.id)) return false
    if (batchOnly && !batch?.includes(doc.id)) return false
    return true
  })
  if (view === 'recent')
    documents = [...documents].sort((a, b) => (b.openedAt || '').localeCompare(a.openedAt || ''))

  if (!api)
    return (
      <div className="environment-error">
        <Library size={40} />
        <h1>请在桌面应用中打开</h1>
        <p>我的文档库需要桌面窗口来访问本地文件。</p>
        <p>开发时请在项目目录运行 npm run dev。</p>
      </div>
    )

  return (
    <>
      <FileDropZone
        disabled={
          busy ||
          loading ||
          !!editor ||
          !!modal ||
          organizing ||
          transferring ||
          !!backupPreview ||
          view === 'trash'
        }
        destination={categoryId ? viewLabel : '全部文件'}
        onFiles={importFiles}
        onBlocked={() => setNotice('请返回文件列表并完成当前操作后再拖入文件')}
      />
      <div
        inert={transferring || !!backupPreview || organizing || !!modal}
        style={{ display: editor ? 'none' : undefined }}
      >
        <LibraryWorkspace
          snapshot={snapshot}
          documents={documents}
          selected={selected}
          preview={preview}
          view={view}
          viewLabel={viewLabel}
          query={query}
          type={type}
          matches={matches}
          busy={busy}
          loading={loading || searching}
          previewOpen={previewOpen}
          batch={batch}
          batchOnly={batchOnly}
          searchRef={searchRef}
          onNavigate={navigate}
          onQuery={setQuery}
          onType={setType}
          onSelect={(id) => {
            setSelectedId(id)
            setPreviewOpen(true)
          }}
          onPreviewToggle={() => setPreviewOpen(!previewOpen)}
          onImport={() => importFiles()}
          onImportFolder={() =>
            void transfer(async () => {
              const preview = await api.previewFolder()
              if (preview) setFolderPreview(preview)
            })
          }
          onManageCategory={setManagedCategory}
          onMoveDocuments={setMoveIds}
          selectionEpoch={selectionEpoch}
          onCreateDocument={() => {
            setNewExtension(null)
            showNameModal('document')
          }}
          onCreateOffice={(extension) => {
            setNewExtension(extension)
            showNameModal('document')
          }}
          onCreateCategory={(parentId) => {
            setCreateParent(parentId)
            showNameModal('category')
          }}
          onSettings={() => setModal('settings')}
          onOpen={open}
          onRename={() => showNameModal('rename')}
          onExport={exportDoc}
          onExportLibrary={exportLibrary}
          onBatch={() => {
            setView('all')
            setQuery('')
            setType('all')
            setBatchOnly(!batchOnly)
          }}
          onDismissBatch={() => {
            setBatch(null)
            setBatchOnly(false)
          }}
          onAction={perform}
          onNotice={setNotice}
          themeControl={themeControl}
        />
      </div>
      {editor && (
        <div
          className="editor-page"
          inert={transferring || !!backupPreview || organizing || !!modal}
        >
          <header className="editor-header">
            <button
              className="secondary"
              disabled={busy}
              onClick={() => {
                if (dirty) setModal('close')
                else {
                  setEditor(null)
                  setHistory(null)
                }
              }}
            >
              <ArrowLeft size={16} />
              返回列表
            </button>
            <div className="editor-title">
              <strong>{editor.doc.name}</strong>
              <span>
                {isMarkdown(editor.doc)
                  ? dirty
                    ? '有未保存的修改'
                    : '已保存在本地'
                  : dirty
                    ? '有未保存的修改'
                    : canEditFile(editor.doc.extension)
                      ? '阅读 / 内容编辑 · 原件保留'
                      : '本地阅读 · 原文件保留'}
              </span>
            </div>
            <button
              className="secondary"
              disabled={busy}
              onClick={() =>
                void perform(async () =>
                  setHistory(history ? null : await api!.versions(editor.doc.id)),
                )
              }
            >
              <History size={16} />
              历史版本
            </button>
            <button className="secondary" disabled={busy} onClick={() => exportDoc(editor.doc)}>
              <Download size={16} />
              导出
            </button>
            <button
              className="secondary"
              aria-label="收藏当前文档"
              aria-pressed={editor.doc.favorite}
              disabled={busy}
              onClick={() =>
                void perform(async () => {
                  await api!.setFavorites([editor.doc.id], !editor.doc.favorite)
                  setEditor({ ...editor, doc: { ...editor.doc, favorite: !editor.doc.favorite } })
                })
              }
            >
              <Star size={17} fill={editor.doc.favorite ? 'currentColor' : 'none'} />
              {editor.doc.favorite ? '已收藏' : '收藏'}
            </button>
            {canEditFile(editor.doc.extension) && (
              <button className="primary" disabled={busy || !dirty} onClick={save}>
                <Save size={16} />
                保存
              </button>
            )}
            {themeControl}
          </header>
          {recovery && (
            <div className="draft-recovery" role="status">
              <span>
                发现 {new Date(recovery.updatedAt).toLocaleString('zh-CN')} 的恢复草稿
                {recovery.revision !== editor.doc.revision
                  ? '（正式版本已更新，请核对后恢复）'
                  : ''}
              </span>
              <button
                className="primary"
                disabled={
                  busy || (!isMarkdown(editor.doc) && recovery.revision !== editor.doc.revision)
                }
                onClick={() => {
                  setEditor({ ...editor, text: recovery.text })
                  setRecovery(null)
                  setDraftStatus('草稿已恢复，保存后生成正式版本')
                }}
              >
                恢复草稿
              </button>
              <button
                className="secondary"
                disabled={busy}
                onClick={() =>
                  void perform(async () => {
                    await api!.discardDraft(editor.doc.id)
                    setRecovery(null)
                  })
                }
              >
                丢弃草稿
              </button>
            </div>
          )}
          <div className="editor-workspace">
            {isMarkdown(editor.doc) ? (
              <EditorSplit>
                <section className="editor-source">
                  <div className="editor-section-title">
                    Markdown 编辑<span>{draftStatus || 'Ctrl + S 保存'}</span>
                  </div>
                  <div className="markdown-toolbar">
                    <button
                      className="secondary"
                      disabled={busy || !!recovery}
                      onClick={() =>
                        void perform(async () => {
                          const image = await api!.addImage(editor.doc.id)
                          if (!image) return
                          setEditorImages(await api!.attachments(editor.doc.id))
                          const at = sourceRef.current?.selectionStart ?? editor.text.length
                          const end = sourceRef.current?.selectionEnd ?? at
                          setEditor({
                            ...editor,
                            text:
                              editor.text.slice(0, at) +
                              `\n![图片](attachments/${image.id}${image.extension})\n` +
                              editor.text.slice(end),
                          })
                        })
                      }
                    >
                      插入图片
                    </button>
                    <span>可粘贴截图 · 图片随备份和导出保存</span>
                  </div>
                  <textarea
                    ref={sourceRef}
                    aria-label="Markdown 内容"
                    spellCheck={false}
                    value={editor.text}
                    disabled={busy || !!recovery}
                    onPaste={(event) => {
                      const file = [...event.clipboardData.files].find((f) =>
                        f.type.startsWith('image/'),
                      )
                      if (!file) return
                      event.preventDefault()
                      if (file.size > 10 * 1024 ** 2) {
                        setError('图片最大支持 10 MB。')
                        return
                      }
                      const at = event.currentTarget.selectionStart,
                        end = event.currentTarget.selectionEnd
                      void perform(async () => {
                        const image = await api!.addImage(
                          editor.doc.id,
                          new Uint8Array(await file.arrayBuffer()),
                        )
                        if (!image) return
                        setEditorImages(await api!.attachments(editor.doc.id))
                        setEditor({
                          ...editor,
                          text:
                            editor.text.slice(0, at) +
                            `\n![图片](attachments/${image.id}${image.extension})\n` +
                            editor.text.slice(end),
                        })
                      })
                    }}
                    onChange={(event) => setEditor({ ...editor, text: event.target.value })}
                    placeholder="# 从这里开始记录"
                  />
                </section>
                <section className="editor-preview">
                  <div className="editor-section-title">阅读预览</div>
                  <ReadingPane
                    doc={editor.doc}
                    text={editor.text}
                    query={query.trim()}
                    attachments={editorImages}
                    disabled={dirty || busy || !!recovery}
                    onChanged={() => void refresh()}
                  />
                </section>
              </EditorSplit>
            ) : fileKind(editor.doc.extension) === 'office' ? (
              <OfficeEditor
                key={editor.doc.id}
                initialMode={editor.initialMode}
                doc={editor.doc}
                query={query.trim()}
                location={editor.location}
                draftStatus={draftStatus}
                text={editor.text}
                busy={busy || !!recovery}
                onChange={(text) => setEditor({ ...editor, text })}
                onChanged={() => void refresh()}
              />
            ) : (
              <section className="file-reader-page">
                <ReadingPane
                  doc={editor.doc}
                  text={editor.readOnlyText ?? null}
                  query={query.trim()}
                  disabled={busy}
                  onChanged={() => void refresh()}
                />
              </section>
            )}
            {history && (
              <aside className="history-panel">
                <h3>
                  历史版本
                  <button
                    className="icon-button"
                    aria-label="关闭历史版本"
                    onClick={() => setHistory(null)}
                  >
                    <X size={15} />
                  </button>
                </h3>
                <p>查看差异可对比两个已保存版本。恢复版本前，请先保存当前修改。</p>
                {history.map((version, index) => (
                  <div className="history-item" key={version.id}>
                    <strong>{version.reason}</strong>
                    <span>{new Date(version.createdAt).toLocaleString('zh-CN')}</span>
                    {canCompareFile(editor.doc.extension) && (
                      <button
                        className="secondary"
                        disabled={busy}
                        onClick={() => setComparedVersion(version.id)}
                      >
                        {history.length > 1 ? '查看差异' : '查看内容'}
                      </button>
                    )}
                    <button
                      className="secondary"
                      disabled={busy || dirty || index === 0}
                      onClick={() =>
                        void perform(async () => {
                          const result = await api!.restoreVersion(
                            editor.doc.id,
                            version.id,
                            editor.doc.revision,
                          )
                          setEditor({
                            doc: result.document,
                            readOnlyText: result.text,
                            text: result.text ?? '',
                            savedText: result.text ?? '',
                          })
                          setHistory(await api!.versions(editor.doc.id))
                          setNotice('已恢复历史版本，恢复前的内容仍然保留')
                        })
                      }
                    >
                      {index === 0 ? '当前版本' : '恢复此版本'}
                    </button>
                  </div>
                ))}
              </aside>
            )}
          </div>
        </div>
      )}
      {editor && history && comparedVersion && (
        <VersionComparison
          documentId={editor.doc.id}
          name={editor.doc.name}
          versions={history}
          selectedId={comparedVersion}
          onClose={() => setComparedVersion(null)}
        />
      )}
      {(error || notice || busy) && (
        <div className={`message ${error ? 'error' : ''}`} role={error ? 'alert' : 'status'}>
          <span>{busy ? '正在处理…' : error || notice}</span>
          {!busy && (
            <button
              className="icon-button"
              aria-label="关闭提示"
              onClick={() => {
                setError('')
                setNotice('')
              }}
            >
              <X size={16} />
            </button>
          )}
        </div>
      )}
      {modal && (
        <div
          className="modal-backdrop"
          inert={transferring || !!backupPreview}
          onMouseDown={(event) => {
            if (event.target === event.currentTarget && !busy) setModal(null)
          }}
        >
          <section
            className={`modal ${modal === 'settings' ? 'storage-modal' : ''}`}
            role="dialog"
            aria-modal="true"
            aria-labelledby="modal-title"
          >
            <div className="modal-heading">
              <h2 id="modal-title">
                {modal === 'settings'
                  ? '设置与存储'
                  : modal === 'close'
                    ? '还有未保存的修改'
                    : modal === 'category'
                      ? '新建分类'
                      : modal === 'rename'
                        ? '重命名文档'
                        : newExtension
                          ? `新建 ${newExtension === '.docx' ? 'Word' : 'Excel'}`
                          : '新建 Markdown'}
              </h2>
              <button
                className="icon-button"
                disabled={busy}
                aria-label="关闭对话框"
                onClick={() => setModal(null)}
              >
                <X size={18} />
              </button>
            </div>
            {modal === 'settings' ? (
              <>
                <div className="storage-card">
                  <HardDrive size={23} />
                  <div>
                    <strong>文档库位置</strong>
                    <p>{snapshot.storagePath}</p>
                  </div>
                </div>
                <p className="muted">版本 0.9.0 · 测试版</p>
                <p>原文件不会随导入而移动或删除。文档副本、分类和历史版本保存在上述目录。</p>
                <StorageActions
                  onBackup={backupLibrary}
                  onRestore={previewBackup}
                  onExport={() => exportLibrary({ kind: 'all' })}
                  result={transferResult}
                />
                <p className="muted">
                  迁移到另一台电脑：创建完整备份，在新电脑的本应用中选择这份备份并恢复。
                </p>
              </>
            ) : modal === 'close' ? (
              <>
                <p>可以保留恢复草稿，下次打开时继续；正式保存的历史版本不受影响。</p>
                <div className="modal-actions">
                  <button className="secondary" onClick={() => setModal(null)}>
                    继续编辑
                  </button>
                  <button
                    className="secondary"
                    disabled={busy}
                    onClick={() =>
                      void perform(async () => {
                        await draftWrite.current
                        await api!.saveDraft(
                          editor!.doc.id,
                          editor!.text,
                          editor!.doc.revision,
                          isMarkdown(editor!.doc) ? 'markdown' : 'office',
                        )
                        setModal(null)
                        setEditor(null)
                        setHistory(null)
                        setRecovery(null)
                      })
                    }
                  >
                    保留草稿并返回
                  </button>
                  <button
                    className="danger"
                    disabled={busy}
                    onClick={() =>
                      void perform(async () => {
                        await draftWrite.current
                        await api!.discardDraft(editor!.doc.id)
                        setModal(null)
                        setEditor(null)
                        setHistory(null)
                        setRecovery(null)
                      })
                    }
                  >
                    放弃修改并返回
                  </button>
                </div>
              </>
            ) : (
              <form onSubmit={submitName}>
                <label className="form-label">
                  {modal === 'category' ? '分类名称' : '文件名称'}
                  <input
                    autoFocus
                    required
                    maxLength={180}
                    value={name}
                    onChange={(event) => setName(event.target.value)}
                    placeholder={modal === 'category' ? '例如：工作资料' : '例如：项目笔记'}
                  />
                </label>
                {modal === 'category' && (
                  <label className="form-label">
                    上级分类
                    <select
                      aria-label="上级分类"
                      value={createParent || ''}
                      onChange={(event) => setCreateParent(event.target.value || null)}
                    >
                      <option value="">顶级分类</option>
                      {snapshot.categories.map((category) => (
                        <option key={category.id} value={category.id}>
                          {categoryPath(snapshot.categories, category.id)
                            .map((item) => item.name)
                            .join(' / ')}
                        </option>
                      ))}
                    </select>
                  </label>
                )}
                <p className="muted">
                  {modal === 'category'
                    ? '先建一个分类，之后随时可以把文件移进来。'
                    : modal === 'rename'
                      ? '请保留文件原有的扩展名。'
                      : '新文档将保存在当前分类；未指定分类的文件可在“全部文件”中查看。'}
                </p>
                <div className="modal-actions">
                  <button
                    className="secondary"
                    type="button"
                    disabled={busy}
                    onClick={() => setModal(null)}
                  >
                    取消
                  </button>
                  <button className="primary" disabled={busy || !name.trim()} type="submit">
                    {modal === 'rename' ? '保存名称' : '创建'}
                  </button>
                </div>
              </form>
            )}
          </section>
        </div>
      )}
      <TransferDialogs
        active={transferring}
        progress={transferProgress}
        preview={backupPreview}
        onCancel={() => void api.cancelOperation()}
        onDismissPreview={() => {
          setBackupPreview(null)
          void transfer(async () => api.discardBackupPreview())
        }}
        onRestore={restoreBackup}
      />
      {managedCategory && (
        <CategoryDialog
          key={managedCategory.id}
          category={managedCategory}
          snapshot={snapshot}
          onClose={() => setManagedCategory(null)}
          onSave={async (name, parentId) => {
            await api.updateCategory(managedCategory.id, { name, parentId })
            await refresh()
            setManagedCategory(null)
            setNotice('分类已更新，文件和子分类已保留')
          }}
          onDelete={async (destination) => {
            const removed = categoryDescendants(snapshot.categories, managedCategory.id)
            await api.deleteCategory(managedCategory.id, destination)
            if (categoryId && removed.has(categoryId))
              navigate(destination ? `category:${destination}` : 'all')
            await refresh()
            setManagedCategory(null)
            setNotice('分类已删除，文件和历史版本已保留')
          }}
        />
      )}
      {moveIds && (
        <MoveDialog
          ids={moveIds}
          categories={snapshot.categories}
          initial={categoryId}
          onClose={() => setMoveIds(null)}
          onMove={async (destination) => {
            await api.moveDocuments(moveIds, destination)
            await refresh()
            setNotice(`已移动 ${moveIds.length} 份文件`)
            setMoveIds(null)
            setSelectionEpoch((value) => value + 1)
            setQuery('')
            setType('all')
            navigate(destination ? `category:${destination}` : 'all')
          }}
        />
      )}
      {folderPreview && (
        <FolderPreviewDialog
          preview={folderPreview}
          categories={snapshot.categories}
          initial={categoryId}
          onClose={() => {
            setFolderPreview(null)
            void perform(async () => api.discardFolderPreview())
          }}
          onImport={(parent, duplicates) => {
            const token = folderPreview.token
            setFolderPreview(null)
            void transfer(async () => {
              const result = await api.importFolder(token, parent, duplicates)
              await refresh()
              setBatch(result.imported.length ? result.imported : null)
              setBatchOnly(false)
              setFolderResult(result)
            })
          }}
        />
      )}
      {folderResult && (
        <FolderResultDialog
          result={folderResult}
          onClose={() => setFolderResult(null)}
          onView={() => {
            setView('all')
            setQuery('')
            setType('all')
            setBatchOnly(true)
            setSelectedId(folderResult.imported[0] || null)
            setPreviewOpen(true)
            setSelectionEpoch((value) => value + 1)
            setFolderResult(null)
          }}
        />
      )}
    </>
  )
}
