/** Shared by import, search and the UI; unknown files are stored without executing them. */
const textExtensions = new Set([
  '.txt',
  '.sql',
  '.csv',
  '.tsv',
  '.json',
  '.jsonl',
  '.xml',
  '.yaml',
  '.yml',
  '.log',
  '.ini',
  '.conf',
  '.cfg',
  '.properties',
  '.html',
  '.htm',
  '.css',
  '.scss',
  '.less',
  '.js',
  '.jsx',
  '.ts',
  '.tsx',
  '.mjs',
  '.cjs',
  '.vue',
  '.svelte',
  '.py',
  '.java',
  '.c',
  '.cpp',
  '.cc',
  '.h',
  '.hpp',
  '.cs',
  '.go',
  '.rs',
  '.rb',
  '.php',
  '.sh',
  '.bat',
  '.cmd',
  '.ps1',
  '.toml',
  '.env',
  '.gitignore',
  '.dockerignore',
  '.rst',
  '.tex',
])
export function fileKind(extension: string) {
  if (['.md', '.markdown'].includes(extension)) return 'markdown'
  if (['.docx', '.xlsx'].includes(extension)) return 'office'
  if (extension === '.pdf') return 'pdf'
  if (textExtensions.has(extension)) return 'text'
  if (['.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp'].includes(extension)) return 'image'
  return 'other'
}
export const canEditFile = (extension: string) =>
  ['markdown', 'office'].includes(fileKind(extension))
export const canCompareFile = (extension: string) =>
  canEditFile(extension) || fileKind(extension) === 'text'
export function fileTypeName(extension: string): string {
  if (extension === '.docx') return 'Word'
  if (extension === '.xlsx') return 'Excel'
  if (extension === '.txt') return 'TXT'
  if (extension === '.sql') return 'SQL'
  return {
    markdown: 'Markdown',
    office: 'Office',
    pdf: 'PDF',
    text: '文本',
    image: '图片',
    other: '其他',
  }[fileKind(extension)]
}
