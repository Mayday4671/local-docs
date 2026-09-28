import type { LibraryApi } from '../../shared/types'
declare global {
  interface Window {
    localDocs?: Omit<LibraryApi, 'importDroppedFiles'>
    localFileImport?: (
      files: File[],
      categoryId: string | null,
    ) => ReturnType<LibraryApi['importDroppedFiles']>
  }
}
