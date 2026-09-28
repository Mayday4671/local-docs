import {
  app,
  BrowserWindow,
  clipboard,
  dialog,
  ipcMain,
  nativeTheme,
  nativeImage,
  net,
  protocol,
  session,
  shell,
} from 'electron'
import { basename, dirname, isAbsolute, join, relative, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { lstatSync, readFileSync, statSync } from 'node:fs'
import { setImmediate as yieldToEvents } from 'node:timers/promises'
import { officeEditModel, applyOfficeChanges } from './office-edit'
import { Library } from './library'
import { OfficeService } from './office-service'
import { TransferService } from './transfer-service'
import { FolderImportService } from './folder-import-service'
import { StorageLocation } from './storage-location'
import { IPC, type ImportResult, type LibraryApi } from '../shared/types'

app.setName('我的文档库')
app.setAppUserModelId('com.mayday.localdocs')
app.setPath(
  'userData',
  process.env.LOCAL_DOCS_DATA_DIR || join(app.getPath('appData'), 'local-docs'),
)
protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true } },
])

const locked = app.requestSingleInstanceLock()
let window: BrowserWindow | null = null
let library: Library | undefined
let activeOffice: OfficeService | undefined
let transfers: TransferService | undefined
let folderImports: FolderImportService | undefined
let storage: StorageLocation | undefined
if (!locked) app.quit()

function windowColors() {
  return nativeTheme.shouldUseDarkColors
    ? { color: '#141c2b', symbolColor: '#e4ebf7', height: 32 }
    : { color: '#f5f8fc', symbolColor: '#18233f', height: 32 }
}
nativeTheme.on('updated', () => {
  if (!window || window.isDestroyed()) return
  const colors = windowColors()
  window.setTitleBarOverlay(colors)
  window.setBackgroundColor(colors.color)
})

function createWindow(): void {
  const mainWindow = new BrowserWindow({
    width: 1536,
    height: 1024,
    minWidth: 1000,
    minHeight: 650,
    show: false,
    title: '我的文档库',
    backgroundColor: windowColors().color,
    titleBarStyle: 'hidden',
    titleBarOverlay: windowColors(),
    autoHideMenuBar: true,
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  window = mainWindow
  mainWindow.on('close', (event) => {
    if (transfers?.active || folderImports?.active || storage?.active) {
      event.preventDefault()
      void dialog.showMessageBox(mainWindow, {
        type: 'info',
        message: '文件导入、备份、恢复、导出或存储迁移仍在进行。',
        detail: '请等待完成，或在操作窗口中取消后再关闭应用。',
      })
    }
  })
  mainWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  mainWindow.webContents.on('will-navigate', (event) => event.preventDefault())
  mainWindow.webContents.on('will-attach-webview', (event) => event.preventDefault())
  mainWindow.webContents.on('will-prevent-unload', (event) => {
    const choice = dialog.showMessageBoxSync(mainWindow, {
      type: 'question',
      buttons: ['继续编辑', '放弃修改并关闭'],
      defaultId: 0,
      cancelId: 0,
      title: '还有未保存的修改',
      message: '当前修改尚未保存。',
      detail:
        '正式文件保留上次保存内容；已自动保存的草稿可在下次打开时恢复。刚输入的内容可能尚未写入草稿。',
    })
    if (choice === 1) event.preventDefault()
  })
  mainWindow.on('closed', () => {
    window = null
  })
  mainWindow.once('ready-to-show', () => {
    if (process.env.LOCAL_DOCS_SMOKE !== '1') mainWindow.show()
  })
  const devUrl = process.env.ELECTRON_RENDERER_URL
  if (!app.isPackaged && devUrl) void mainWindow.loadURL(devUrl)
  else void mainWindow.loadURL('app://local/index.html')
}

function registerApi(
  store: Library,
  office: OfficeService,
  transfer: TransferService,
  folders: FolderImportService,
): void {
  let queue: Promise<unknown> = Promise.resolve()
  let generation = 0
  const importPaths = async (paths: string[], categoryId: string | null): Promise<ImportResult> => {
    if (
      !Array.isArray(paths) ||
      paths.length > 500 ||
      paths.some((p) => typeof p !== 'string' || !isAbsolute(p))
    )
      throw new Error('请选择本地文件，每次最多 500 份。')
    paths = [...new Set(paths)]
    let totalBytes = 0
    const result: ImportResult = { imported: [], failures: [] }
    for (const path of paths) {
      try {
        const size = lstatSync(path).size
        if (size + totalBytes > 2 * 1024 ** 3)
          throw new Error('本批已达到 2 GB 上限，请分批导入剩余文件。')
        const imported = store.importFile(path, categoryId)
        totalBytes += imported.size
        result.imported.push(imported.id)
        if (['.docx', '.xlsx'].includes(imported.extension)) await office.ensure(imported.id)
      } catch (error) {
        result.failures.push({
          name: basename(path),
          reason: error instanceof Error ? error.message : '导入失败',
        })
      }
      await yieldToEvents()
    }
    return result
  }
  const api: LibraryApi = {
    storageInfo: async () => storage!.info(store),
    previewStorageLocation: async () => {
      storage!.discardPreview()
      const choice = await dialog.showOpenDialog(window!, {
        title: '选择新的文档库位置（空文件夹）',
        defaultPath: dirname(store.root),
        properties: ['openDirectory', 'createDirectory'],
        buttonLabel: '选择此位置',
      })
      return choice.canceled || !choice.filePaths[0]
        ? null
        : storage!.prepare(store, choice.filePaths[0])
    },
    discardStoragePreview: async () => storage!.discardPreview(),
    moveStorage: async (token) => {
      await transfer.discardPreview()
      folders.discardPreview()
      const replacement = await storage!.move(store, token)
      const original = store
      office.dispose()
      store = library = replacement
      office = activeOffice = new OfficeService(store)
      // The large temporary files and recovery backups follow the selected storage volume.
      transfer = transfers = new TransferService(store, store.root)
      folders = folderImports = new FolderImportService(store, store.root)
      generation++
      original.close()
      return storage!.info(store)
    },
    createOffice: async (name, extension, categoryId) =>
      store.createOffice(name, extension, categoryId),
    annotations: async (id) => store.annotations(id),
    saveAnnotation: async (id, value, revision) => store.saveAnnotation(id, value, revision),
    removeAnnotation: async (id, annotationId) => store.removeAnnotation(id, annotationId),
    saveDraft: async (id, text, revision, kind) => store.saveDraft(id, text, revision, kind),
    draft: async (id) => store.documentState(id).draft,
    discardDraft: async (id) => store.discardDraft(id),
    attachments: async (id) => store.documentState(id).attachments,
    setFavorites: async (ids, favorite) => store.setFavorites(ids, favorite),
    officeEdit: async (id) => {
      const source = store.officeSource(id)
      return officeEditModel(
        readFileSync(source.path),
        source.extension,
        store.readDocument(id).document.revision,
      )
    },
    saveOffice: async (id, changes, revision) => {
      const source = store.officeSource(id)
      const bytes = applyOfficeChanges(readFileSync(source.path), source.extension, changes)
      const result = store.saveOfficeBytes(id, bytes, revision)
      await office.ensure(id)
      return result
    },
    addImage: async (id, pasted) => {
      let bytes: Buffer,
        name = '粘贴图片.png'
      if (pasted !== undefined) {
        if (!(pasted instanceof Uint8Array) || pasted.byteLength > 10 * 1024 ** 2)
          throw new Error('请选择 10 MB 以内的图片。')
        bytes = Buffer.from(pasted)
      } else {
        const choice = await dialog.showOpenDialog(window!, {
          title: '插入本地图片',
          properties: ['openFile'],
          filters: [{ name: '图片', extensions: ['png', 'jpg', 'jpeg', 'webp'] }],
        })
        if (choice.canceled || !choice.filePaths[0]) return null
        if (statSync(choice.filePaths[0]).size > 10 * 1024 ** 2)
          throw new Error('图片最大支持 10 MB。')
        bytes = readFileSync(choice.filePaths[0])
        name = basename(choice.filePaths[0])
      }
      const img = nativeImage.createFromBuffer(bytes),
        size = img.getSize()
      if (img.isEmpty() || size.width * size.height > 16_000_000)
        throw new Error('图片无法读取或超过 1600 万像素。')
      return store.addImage(id, img.toPNG(), 'image/png', name)
    },
    previewFolder: async () => {
      folders.discardPreview()
      const choice = await dialog.showOpenDialog(window!, {
        title: '选择要导入的资料文件夹',
        properties: ['openDirectory'],
        buttonLabel: '预览导入',
      })
      return choice.canceled || !choice.filePaths[0] ? null : folders.preview(choice.filePaths[0])
    },
    discardFolderPreview: async () => folders.discardPreview(),
    importFolder: async (token, parentId, duplicates) =>
      folders.import(token, parentId, duplicates),
    updateCategory: async (id, patch) => store.updateCategory(id, patch),
    deleteCategory: async (id, destination) => store.deleteCategory(id, destination),
    moveDocuments: async (ids, categoryId) => store.moveDocuments(ids, categoryId),
    createBackup: async () => {
      const choice = await dialog.showSaveDialog(window!, {
        title: '备份整个文档库',
        defaultPath: `我的文档库-${new Date().toISOString().slice(0, 10)}.localdocs-backup`,
        filters: [{ name: '文档库完整备份', extensions: ['localdocs-backup'] }],
        buttonLabel: '创建备份',
      })
      return choice.canceled || !choice.filePath ? null : transfer.backup(choice.filePath)
    },
    previewBackup: async () => {
      const choice = await dialog.showOpenDialog(window!, {
        title: '选择文档库备份',
        properties: ['openFile'],
        filters: [{ name: '文档库完整备份', extensions: ['localdocs-backup'] }],
      })
      return choice.canceled || !choice.filePaths[0] ? null : transfer.preview(choice.filePaths[0])
    },
    discardBackupPreview: async () => transfer.discardPreview(),
    restoreBackup: async (token) => {
      const result = await transfer.restore(token)
      folders.discardPreview()
      generation++
      nativeTheme.themeSource = store.getTheme()
      return result
    },
    exportLibrary: async (scope) => {
      const choice = await dialog.showOpenDialog(window!, {
        title:
          scope?.kind === 'category' ? '选择分类导出位置（含子分类）' : '选择整个文档库的导出位置',
        properties: ['openDirectory', 'createDirectory'],
        buttonLabel: '导出到此处',
      })
      return choice.canceled || !choice.filePaths[0]
        ? null
        : transfer.export(choice.filePaths[0], scope)
    },
    operationStatus: async () =>
      storage!.getStatus() ?? transfer.getStatus() ?? folders.getStatus(),
    cancelOperation: async () => {
      transfer.cancel()
      folders.cancel()
      storage!.cancel()
    },
    revealTransferResult: async () => {
      if (transfer.lastResultPath) shell.showItemInFolder(transfer.lastResultPath)
    },
    getTheme: async () => store.getTheme(),
    setTheme: async (theme) => {
      store.setTheme(theme)
      nativeTheme.themeSource = theme
    },
    snapshot: async () => store.snapshot(),
    search: async (query) => {
      await office.indexPending()
      return store.search(query)
    },
    searchResults: async (query) => {
      await office.indexPending()
      return store.searchResults(query)
    },
    readOffice: async (id, includeLayout) => office.read(id, includeLayout === true),
    importFiles: async (categoryId) => {
      const choice = await dialog.showOpenDialog(window!, {
        title: '添加到我的文档库',
        properties: ['openFile', 'multiSelections'],
        filters: [
          { name: '所有文件', extensions: ['*'] },
          {
            name: '常用文档',
            extensions: ['docx', 'xlsx', 'md', 'markdown', 'pdf', 'txt', 'sql', 'csv', 'json'],
          },
        ],
      })
      return choice.canceled
        ? { imported: [], failures: [] }
        : importPaths(choice.filePaths, categoryId)
    },
    importDroppedFiles: importPaths,
    readFilePreview: async (id) => store.readFilePreview(id),
    createMarkdown: async (name, categoryId) => store.createMarkdown(name, categoryId),
    createCategory: async (name, parentId) => store.createCategory(name, parentId ?? null),
    copyDocumentPath: async (id) => clipboard.writeText(store.documentPath(id)),
    readDocument: async (id) => store.readDocument(id),
    openDocument: async (id) => store.openDocument(id),
    saveMarkdown: async (id, text, revision) => store.saveMarkdown(id, text, revision),
    updateDocument: async (id, patch) => store.updateDocument(id, patch),
    trashDocument: async (id) => store.trashDocument(id),
    restoreDocument: async (id) => store.restoreDocument(id),
    exportDocument: async (id) => {
      const document = store.readDocument(id).document
      const choice = await dialog.showSaveDialog(window!, {
        title: '导出文档副本',
        defaultPath: document.name,
        buttonLabel: '导出',
      })
      if (choice.canceled || !choice.filePath) return false
      try {
        store.exportDocument(id, choice.filePath)
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'EEXIST')
          throw new Error('该位置已有同名文件，请换个名称导出。')
        throw error
      }
      return true
    },
    versions: async (id) => store.versions(id),
    readVersion: (id, versionId) => office.readVersion(id, versionId),
    restoreVersion: async (id, versionId, revision) =>
      store.restoreVersion(id, versionId, revision),
  }
  for (const key of Object.keys(IPC) as (keyof LibraryApi)[]) {
    ipcMain.handle(IPC[key], (event, ...args: unknown[]) => {
      if (
        !window ||
        event.sender !== window.webContents ||
        event.senderFrame !== window.webContents.mainFrame
      )
        throw new Error('拒绝来自非主窗口的请求。')
      const url = new URL(event.senderFrame.url)
      const devUrl = !app.isPackaged && process.env.ELECTRON_RENDERER_URL
      const trusted = devUrl
        ? url.origin === new URL(devUrl).origin
        : url.protocol === 'app:' && url.host === 'local'
      if (!trusted) throw new Error('拒绝来自未知页面的请求。')
      const invoke = () => (api[key] as (...values: unknown[]) => Promise<unknown>)(...args)
      if (key === 'operationStatus' || key === 'cancelOperation') return invoke()
      if (transfer.active || folders.active || storage!.active)
        throw new Error('请等待文件导入、备份、恢复、导出或存储迁移完成。')
      const acceptedGeneration = generation
      const result = queue.then(() => {
        if (generation !== acceptedGeneration)
          throw new Error('文档库位置或内容已切换，请重新操作。')
        return invoke()
      })
      queue = result.catch(() => {})
      return result
    })
  }
}

if (locked) {
  app.on('second-instance', () => {
    if (window?.isMinimized()) window.restore()
    window?.focus()
  })
  void app
    .whenReady()
    .then(async () => {
      session.defaultSession.setPermissionRequestHandler((_contents, _permission, callback) =>
        callback(false),
      )
      session.defaultSession.setPermissionCheckHandler(() => false)
      const rendererRoot = resolve(__dirname, '../renderer')
      protocol.handle('app', (request) => {
        const url = new URL(request.url)
        if (url.host !== 'local') return new Response('Not found', { status: 404 })
        const attachment = url.pathname.match(/^\/attachment\/([a-f0-9-]{36})\/([a-f0-9-]{36})$/)
        if (attachment) {
          try {
            const image = library!.attachmentPath(attachment[1], attachment[2])
            return new Response(new Uint8Array(readFileSync(image.path)), {
              headers: {
                'Content-Type': image.mime,
                'Cache-Control': 'no-store',
                'X-Content-Type-Options': 'nosniff',
              },
            })
          } catch {
            return new Response('Not found', { status: 404 })
          }
        }
        let pathname: string
        try {
          pathname = decodeURIComponent(url.pathname)
        } catch {
          return new Response('Invalid path', { status: 400 })
        }
        const path = resolve(rendererRoot, `.${pathname}`)
        const relativePath = relative(rendererRoot, path)
        if (relativePath.startsWith('..') || isAbsolute(relativePath))
          return new Response('Forbidden', { status: 403 })
        return net.fetch(pathToFileURL(path).toString())
      })
      const testInstallation =
        process.env.LOCAL_DOCS_SMOKE === '1' && process.env.LOCAL_DOCS_TEST_INSTALL_DIR
      storage = new StorageLocation(
        app.getPath('userData'),
        testInstallation || (app.isPackaged ? dirname(app.getPath('exe')) : app.getAppPath()),
        process.env.LOCAL_DOCS_DATA_DIR && !testInstallation
          ? join(app.getPath('userData'), 'library')
          : undefined,
      )
      try {
        library = await storage.open()
      } catch (initialError) {
        let error = initialError
        while (!library) {
          const choice = await dialog.showMessageBox({
            type: 'error',
            title: '文档库位置不可用',
            message: error instanceof Error ? error.message : String(error),
            detail:
              '可以重新选择已有文档库。首次使用时也可选择有写入权限的空文件夹。原有资料不会被删除。',
            buttons: ['重新选择位置', '退出'],
            defaultId: 0,
            cancelId: 1,
          })
          if (choice.response !== 0) {
            app.quit()
            return
          }
          const folder = await dialog.showOpenDialog({
            title: '选择文档库位置',
            properties: ['openDirectory', 'createDirectory'],
          })
          if (folder.canceled || !folder.filePaths[0]) {
            app.quit()
            return
          }
          try {
            library = await storage.recoverLocation(folder.filePaths[0])
          } catch (nextError) {
            error = nextError
          }
        }
      }
      nativeTheme.themeSource = library.getTheme()
      activeOffice = new OfficeService(library)
      transfers = new TransferService(library, library.root)
      folderImports = new FolderImportService(library, library.root)
      registerApi(library, activeOffice, transfers, folderImports)
      createWindow()
    })
    .catch((error) => {
      dialog.showErrorBox('文档库启动失败', error instanceof Error ? error.message : String(error))
      app.quit()
    })
  app.on('activate', () => {
    if (!window && library) createWindow()
  })
  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })
  app.on('will-quit', () => {
    activeOffice?.dispose()
    library?.close()
  })
  app.on('before-quit', (event) => {
    if (transfers?.active || folderImports?.active || storage?.active) event.preventDefault()
  })
}
