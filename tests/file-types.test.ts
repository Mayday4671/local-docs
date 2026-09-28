import { expect, test } from 'vitest'
import { plainTextPreview } from '../src/main/text-preview'
import { canEditFile, fileKind } from '../src/shared/file-types'

test('文本识别 UTF-8、UTF-16 双字节序和 GBK 中文，保留空格换行及代码原文', () => {
  const text = '中文\r\n  SELECT 1; <script>alert(1)</script>'
  expect(plainTextPreview(Buffer.from(text))).toBe(text)
  expect(plainTextPreview(Buffer.from('\ufeff' + text, 'utf16le'))).toBe(text)
  expect(plainTextPreview(Buffer.from('\ufeff' + text, 'utf16le').swap16())).toBe(text)
  expect(plainTextPreview(Buffer.from([0xd6, 0xd0, 0xce, 0xc4]))).toBe('中文')
  expect(plainTextPreview(Buffer.from(''))).toBe('')
})
test('二进制伪装文本与过大文本不进入正文索引或预览', () => {
  expect(plainTextPreview(Buffer.from([0, 1, 2, 3]))).toBeNull()
  expect(plainTextPreview(Buffer.alloc(5 * 1024 ** 2 + 1, 65))).toBeNull()
})
test('按实际格式选择阅读器，新增格式不误入 Office 编辑器', () => {
  for (const ext of ['.txt', '.sql', '.csv', '.json', '.html', '.py', '.java'])
    expect(fileKind(ext)).toBe('text')
  expect(fileKind('.pdf')).toBe('pdf')
  expect(fileKind('.png')).toBe('image')
  for (const ext of ['.pdf', '.txt', '.sql', '.zip', '.exe', ''])
    expect(canEditFile(ext)).toBe(false)
  for (const ext of ['.md', '.markdown', '.docx', '.xlsx']) expect(canEditFile(ext)).toBe(true)
})
