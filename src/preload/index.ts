import { contextBridge, ipcRenderer, webUtils } from 'electron'
import { IPC, type LibraryApi } from '../shared/types'

const api: Omit<LibraryApi, 'importDroppedFiles'> = {
  storageInfo: () => ipcRenderer.invoke(IPC.storageInfo),
  previewStorageLocation: () => ipcRenderer.invoke(IPC.previewStorageLocation),
  discardStoragePreview: () => ipcRenderer.invoke(IPC.discardStoragePreview),
  moveStorage: (token) => ipcRenderer.invoke(IPC.moveStorage, token),
  createOffice: (name, extension, categoryId) =>
    ipcRenderer.invoke(IPC.createOffice, name, extension, categoryId),
  annotations: (id) => ipcRenderer.invoke(IPC.annotations, id),
  saveAnnotation: (id, annotation, revision) =>
    ipcRenderer.invoke(IPC.saveAnnotation, id, annotation, revision),
  removeAnnotation: (id, annotationId) =>
    ipcRenderer.invoke(IPC.removeAnnotation, id, annotationId),
  saveDraft: (id, text, revision, kind) =>
    ipcRenderer.invoke(IPC.saveDraft, id, text, revision, kind),
  draft: (id) => ipcRenderer.invoke(IPC.draft, id),
  discardDraft: (id) => ipcRenderer.invoke(IPC.discardDraft, id),
  addImage: (id, pasted) => ipcRenderer.invoke(IPC.addImage, id, pasted),
  attachments: (id) => ipcRenderer.invoke(IPC.attachments, id),
  setFavorites: (ids, favorite) => ipcRenderer.invoke(IPC.setFavorites, ids, favorite),
  officeEdit: (id) => ipcRenderer.invoke(IPC.officeEdit, id),
  saveOffice: (id, changes, revision) => ipcRenderer.invoke(IPC.saveOffice, id, changes, revision),
  previewFolder: () => ipcRenderer.invoke(IPC.previewFolder),
  discardFolderPreview: () => ipcRenderer.invoke(IPC.discardFolderPreview),
  importFolder: (token, parentId, duplicates) =>
    ipcRenderer.invoke(IPC.importFolder, token, parentId, duplicates),
  updateCategory: (id, patch) => ipcRenderer.invoke(IPC.updateCategory, id, patch),
  deleteCategory: (id, destination) => ipcRenderer.invoke(IPC.deleteCategory, id, destination),
  moveDocuments: (ids, categoryId) => ipcRenderer.invoke(IPC.moveDocuments, ids, categoryId),
  createBackup: () => ipcRenderer.invoke(IPC.createBackup),
  previewBackup: () => ipcRenderer.invoke(IPC.previewBackup),
  discardBackupPreview: () => ipcRenderer.invoke(IPC.discardBackupPreview),
  restoreBackup: (token) => ipcRenderer.invoke(IPC.restoreBackup, token),
  exportLibrary: (scope) => ipcRenderer.invoke(IPC.exportLibrary, scope),
  operationStatus: () => ipcRenderer.invoke(IPC.operationStatus),
  cancelOperation: () => ipcRenderer.invoke(IPC.cancelOperation),
  revealTransferResult: () => ipcRenderer.invoke(IPC.revealTransferResult),
  getTheme: () => ipcRenderer.invoke(IPC.getTheme),
  setTheme: (theme) => ipcRenderer.invoke(IPC.setTheme, theme),
  snapshot: () => ipcRenderer.invoke(IPC.snapshot),
  search: (query) => ipcRenderer.invoke(IPC.search, query),
  searchResults: (query) => ipcRenderer.invoke(IPC.searchResults, query),
  readOffice: (id, includeLayout) => ipcRenderer.invoke(IPC.readOffice, id, includeLayout),
  readFilePreview: (id) => ipcRenderer.invoke(IPC.readFilePreview, id),
  importFiles: (categoryId) => ipcRenderer.invoke(IPC.importFiles, categoryId),
  createMarkdown: (name, categoryId) => ipcRenderer.invoke(IPC.createMarkdown, name, categoryId),
  createCategory: (name, parentId) => ipcRenderer.invoke(IPC.createCategory, name, parentId),
  copyDocumentPath: (id) => ipcRenderer.invoke(IPC.copyDocumentPath, id),
  readDocument: (id) => ipcRenderer.invoke(IPC.readDocument, id),
  openDocument: (id) => ipcRenderer.invoke(IPC.openDocument, id),
  saveMarkdown: (id, text, revision) => ipcRenderer.invoke(IPC.saveMarkdown, id, text, revision),
  updateDocument: (id, patch) => ipcRenderer.invoke(IPC.updateDocument, id, patch),
  trashDocument: (id) => ipcRenderer.invoke(IPC.trashDocument, id),
  restoreDocument: (id) => ipcRenderer.invoke(IPC.restoreDocument, id),
  purgeDocuments: (ids) => ipcRenderer.invoke(IPC.purgeDocuments, ids),
  exportDocument: (id) => ipcRenderer.invoke(IPC.exportDocument, id),
  versions: (id) => ipcRenderer.invoke(IPC.versions, id),
  readVersion: (id, versionId) => ipcRenderer.invoke(IPC.readVersion, id, versionId),
  restoreVersion: (id, versionId, revision) =>
    ipcRenderer.invoke(IPC.restoreVersion, id, versionId, revision),
}
contextBridge.exposeInMainWorld('localDocs', Object.freeze(api))

// Only native File objects cross this bridge; renderer code never supplies arbitrary paths.
contextBridge.exposeInMainWorld('localFileImport', (files: File[], categoryId: string | null) => {
  if (!Array.isArray(files) || files.length > 500) throw new Error('每次最多拖入 500 份文件。')
  const paths = files.map((file) => webUtils.getPathForFile(file))
  if (paths.some((path) => !path)) throw new Error('请从电脑的文件夹拖入文件。')
  return ipcRenderer.invoke(IPC.importDroppedFiles, paths, categoryId)
})
