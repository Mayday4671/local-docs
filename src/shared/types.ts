export type LibraryView =
  'all' | 'recent' | 'favorites' | 'marked' | 'uncategorized' | 'trash' | `category:${string}`
export interface Category {
  id: string
  name: string
  parentId: string | null
}
export interface DocumentRecord {
  id: string
  name: string
  extension: string
  categoryId: string | null
  favorite: boolean
  size: number
  createdAt: string
  updatedAt: string
  openedAt: string | null
  deletedAt: string | null
  revision: number
  tags: string[]
  notes: string
}
export interface LibrarySnapshot {
  documents: DocumentRecord[]
  categories: Category[]
  storagePath: string
  annotationCounts?: Record<string, number>
}
export interface DocumentContent {
  document: DocumentRecord
  text: string | null
}
export type ContentLocation =
  | { kind: 'paragraph'; part: string; index: number }
  | { kind: 'cell'; sheet: string; address: string }
export interface TextBlock {
  text: string
  location: ContentLocation
}
export interface SheetCell {
  address: string
  row: number
  column: number
  text: string
  rawValue?: string
  formula?: string
}
export interface OfficeSheet {
  name: string
  cells: SheetCell[]
  hidden: boolean
  merges: string[]
}
export interface OfficeData {
  kind: 'docx' | 'xlsx'
  blocks: TextBlock[]
  sheets: OfficeSheet[]
  warnings: string[]
}
export interface OfficePreview {
  data: OfficeData | null
  error: string | null
  bytes?: Uint8Array
}
export interface SearchHit {
  id: string
  source: 'name' | 'content' | 'metadata'
  snippet: string
  location?: ContentLocation
}
export interface ImportResult {
  imported: string[]
  failures: { name: string; reason: string }[]
}
export interface VersionRecord {
  id: string
  createdAt: string
  reason: string
}
export interface VersionContent {
  version: VersionRecord
  extension: string
  hash: string
  text: string | null
  office: OfficeData | null
}
export interface LibraryApi {
  createOffice(
    name: string,
    extension: '.docx' | '.xlsx',
    categoryId: string | null,
  ): Promise<DocumentRecord>
  annotations(id?: string): Promise<Annotation[]>
  saveAnnotation(id: string, annotation: AnnotationInput, revision: number): Promise<Annotation>
  removeAnnotation(id: string, annotationId: string): Promise<void>
  saveDraft(id: string, text: string, revision: number, kind: Draft['kind']): Promise<void>
  draft(id: string): Promise<Draft | null>
  discardDraft(id: string): Promise<void>
  addImage(id: string, pasted?: Uint8Array): Promise<Attachment | null>
  attachments(id: string): Promise<Attachment[]>
  setFavorites(ids: string[], favorite: boolean): Promise<void>
  officeEdit(id: string): Promise<OfficeEditModel>
  saveOffice(id: string, changes: OfficeChange[], revision: number): Promise<DocumentRecord>
  previewFolder(): Promise<FolderImportPreview | null>
  discardFolderPreview(): Promise<void>
  importFolder(
    token: string,
    parentId: string | null,
    duplicates: 'skip' | 'keep',
  ): Promise<FolderImportResult>
  updateCategory(id: string, patch: { name: string; parentId: string | null }): Promise<void>
  deleteCategory(id: string, destination: string | null): Promise<void>
  moveDocuments(ids: string[], categoryId: string | null): Promise<void>
  createBackup(): Promise<TransferResult | null>
  previewBackup(): Promise<BackupPreview | null>
  discardBackupPreview(): Promise<void>
  restoreBackup(token: string): Promise<TransferResult>
  exportLibrary(scope: ExportScope): Promise<TransferResult | null>
  operationStatus(): Promise<OperationStatus | null>
  cancelOperation(): Promise<void>
  revealTransferResult(): Promise<void>
  getTheme(): Promise<ThemePreference>
  setTheme(theme: ThemePreference): Promise<void>
  snapshot(): Promise<LibrarySnapshot>
  search(query: string): Promise<string[]>
  searchResults(query: string): Promise<SearchHit[]>
  readOffice(id: string, includeLayout?: boolean): Promise<OfficePreview>
  importFiles(categoryId: string | null): Promise<ImportResult>
  createMarkdown(name: string, categoryId: string | null): Promise<DocumentRecord>
  createCategory(name: string, parentId?: string | null): Promise<Category>
  copyDocumentPath(id: string): Promise<void>
  readDocument(id: string): Promise<DocumentContent>
  openDocument(id: string): Promise<DocumentContent>
  saveMarkdown(id: string, text: string, revision: number): Promise<DocumentRecord>
  updateDocument(
    id: string,
    patch: {
      name?: string
      categoryId?: string | null
      favorite?: boolean
      tags?: string[]
      notes?: string
    },
  ): Promise<void>
  trashDocument(id: string): Promise<void>
  restoreDocument(id: string): Promise<void>
  exportDocument(id: string): Promise<boolean>
  versions(id: string): Promise<VersionRecord[]>
  readVersion(id: string, versionId: string): Promise<VersionContent>
  restoreVersion(id: string, versionId: string, revision: number): Promise<DocumentContent>
}

export const IPC = {
  createOffice: 'library:create-office',
  annotations: 'library:annotations',
  saveAnnotation: 'library:save-annotation',
  removeAnnotation: 'library:remove-annotation',
  saveDraft: 'library:save-draft',
  draft: 'library:draft',
  discardDraft: 'library:discard-draft',
  addImage: 'library:add-image',
  attachments: 'library:attachments',
  setFavorites: 'library:set-favorites',
  officeEdit: 'library:office-edit',
  saveOffice: 'library:save-office',
  previewFolder: 'folder:preview',
  discardFolderPreview: 'folder:discard',
  importFolder: 'folder:import',
  updateCategory: 'library:update-category',
  deleteCategory: 'library:delete-category',
  moveDocuments: 'library:move-documents',
  createBackup: 'transfer:backup',
  previewBackup: 'transfer:preview',
  discardBackupPreview: 'transfer:discard',
  restoreBackup: 'transfer:restore',
  exportLibrary: 'transfer:export',
  operationStatus: 'transfer:status',
  cancelOperation: 'transfer:cancel',
  revealTransferResult: 'transfer:reveal',
  getTheme: 'settings:get-theme',
  setTheme: 'settings:set-theme',
  snapshot: 'library:snapshot',
  search: 'library:search',
  searchResults: 'library:search-results',
  readOffice: 'library:read-office',
  importFiles: 'library:import',
  createMarkdown: 'library:create-markdown',
  createCategory: 'library:create-category',
  copyDocumentPath: 'library:copy-path',
  readDocument: 'library:read',
  openDocument: 'library:open',
  saveMarkdown: 'library:save',
  updateDocument: 'library:update',
  trashDocument: 'library:trash',
  restoreDocument: 'library:restore',
  exportDocument: 'library:export',
  versions: 'library:versions',
  readVersion: 'library:read-version',
  restoreVersion: 'library:restore-version',
} as const

export type ThemePreference = 'system' | 'light' | 'dark'

export type ExportScope = { kind: 'all' } | { kind: 'category'; id: string }
export interface TransferResult {
  path: string
  files: number
  safetyBackupPath?: string
}

export interface AnnotationInput {
  id?: string
  quote: string
  prefix: string
  suffix: string
  offset: number
  scope: string
  location?: ContentLocation
  color: 'yellow' | 'green' | 'blue' | 'pink'
  note: string
}
export interface Annotation extends AnnotationInput {
  id: string
  documentId: string
  revision: number
  createdAt: string
}
export interface Draft {
  text: string
  revision: number
  updatedAt: string
  kind: 'markdown' | 'office'
}
export interface Attachment {
  id: string
  name: string
  hash: string
  size: number
  mime: 'image/png' | 'image/jpeg' | 'image/webp'
  extension: '.png' | '.jpg' | '.webp'
}
export interface DocumentState {
  annotations: Annotation[]
  attachments: Attachment[]
  draft: Draft | null
}
export interface OfficeChange {
  key: string
  text: string
}
export interface OfficeEditField {
  key: string
  label: string
  text: string
  editable: boolean
  reason?: string
  anchor?: string
}
export interface OfficeEditSheet extends OfficeSheet {
  part: string
  editable: boolean
  protectedRanges: string[]
}
export interface OfficeEditModel {
  fields: OfficeEditField[]
  kind: 'docx' | 'xlsx'
  revision: number
  layoutBytes?: Uint8Array
  sheets?: OfficeEditSheet[]
}
export interface BackupPreview {
  token: string
  createdAt: string
  files: number
  trash: number
  categories: number
  versions: number
  bytes: number
}
export interface OperationStatus {
  kind: 'backup' | 'preview' | 'restore' | 'export' | 'folder-scan' | 'folder-import'
  phase: string
  completed: number
  total: number
  cancellable: boolean
}

export interface FolderImportPreview {
  token: string
  rootName: string
  sourcePath: string
  files: number
  folders: number
  bytes: number
  skipped: number
  warnings: { name: string; reason: string }[]
  entries: { path: string; kind: 'file' | 'folder' }[]
}
export interface FolderImportResult extends ImportResult {
  rootCategoryId: string | null
  skipped: number
  cancelled: boolean
}
