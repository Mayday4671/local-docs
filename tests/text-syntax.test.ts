import { expect, test } from 'vitest'
import { textLanguage, MAX_SYNTAX_LENGTH } from '../src/shared/text-language'
import { syntaxTokens } from '../src/renderer/src/text-syntax'
import { plainTextPreview } from '../src/main/text-preview'

test('扩展名优先选择语言，SQL 和不完整 JSON 都能识别', () => {
  expect(textLanguage('.SQL', '{"message": "中文"}').language).toBe('sql')
  expect(textLanguage('.json', '{"message": ').language).toBe('json')
  expect(textLanguage('.jsonl', '{"id":1}\n{"id":2}').label).toBe('JSON Lines')
})

test('普通文本仅在结构明确时推断语言，避免给自然语言误加代码颜色', () => {
  expect(textLanguage('.txt', '{"中文": [true, 2]}')).toMatchObject({
    language: 'json',
    source: 'content',
  })
  expect(textLanguage('.txt', '-- 中文说明\nSELECT name FROM documents;').language).toBe('sql')
  expect(textLanguage('.txt', '<?xml version="1.0"?><root>中文</root>').language).toBe('xml')
  expect(textLanguage('.txt', '#!/usr/bin/env python3\nprint("中文")').language).toBe('python')
  for (const text of [
    '普通中文说明',
    'SELECT an option from the menu',
    '{未完成的普通说明',
    'true',
  ])
    expect(textLanguage('.txt', text).language).toBeNull()
  expect(textLanguage('.log', '{"event":true}').language).toBeNull()
})

test('SQL 与 JSON 的真实词法类别可区分，中文、缩进和 CRLF 一字不改', () => {
  const sql = '-- 中文\r\n\tSELECT "name", 42 FROM users WHERE name = \'测试\';\r\n'
  const tokens = syntaxTokens('sql', sql)!
  expect(tokens.map((token) => token.text).join('')).toBe(sql)
  expect(tokens.some((t) => t.classes.includes('hljs-keyword') && /SELECT/i.test(t.text))).toBe(
    true,
  )
  expect(tokens.some((t) => t.classes.includes('hljs-comment') && t.text.includes('中文'))).toBe(
    true,
  )
  expect(tokens.some((t) => t.classes.includes('hljs-string') && t.text.includes('测试'))).toBe(
    true,
  )
  const json = '{\r\n  "中文": "<script>window.injected = true</script>", "n": 42, "ok": true\r\n}'
  const jsonTokens = syntaxTokens('json', json)!
  expect(jsonTokens.map((token) => token.text).join('')).toBe(json)
  expect(jsonTokens.some((t) => t.classes.includes('hljs-attr') && t.text.includes('中文'))).toBe(
    true,
  )
  expect(jsonTokens.some((t) => t.classes.includes('hljs-number') && t.text === '42')).toBe(true)
  expect(jsonTokens.some((t) => t.classes.includes('hljs-literal') && t.text === 'true')).toBe(true)
})

test('常见代码和配置格式都有可用的本地语法，HTML 只返回原文和样式', () => {
  const examples = {
    '.xml': '<root>中文</root>',
    '.html': '<script>window.injected = true</script>',
    '.yaml': 'name: 中文\nitems: [1, 2]',
    '.ini': '[section]\nname=中文',
    '.properties': 'name=中文',
    '.js': 'const name = "中文";',
    '.tsx': 'const a: string = "中文";',
    '.css': 'body { color: red; }',
    '.scss': '$color: red;',
    '.less': '@color: red;',
    '.py': 'def hello():\n  return "中文"',
    '.java': 'class A { int n = 1; }',
    '.cpp': '#include <string>\nint main() { return 0; }',
    '.cs': 'class A { string name = "中文"; }',
    '.go': 'package main\nfunc main() {}',
    '.rs': 'fn main() { let n = 1; }',
    '.rb': 'puts "中文"',
    '.php': '<?php echo "中文"; ?>',
    '.sh': 'echo "中文"',
    '.bat': '@echo off\necho 中文',
    '.ps1': 'Write-Host "中文"',
    '.tex': '\\section{中文}',
  }
  for (const [extension, text] of Object.entries(examples)) {
    const language = textLanguage(extension, text).language!
    const tokens = syntaxTokens(language, text)
    expect(tokens, extension).not.toBeNull()
    expect(tokens!.map((token) => token.text).join(''), extension).toBe(text)
    expect(
      tokens!.some((token) => token.classes.length > 0),
      extension,
    ).toBe(true)
    expect(tokens!.every((token) => token.classes.every((c) => /^[\w-]+$/.test(c)))).toBe(true)
  }
})

test('不同字符编码先正确解码，再识别语言，原件字节不改变', () => {
  const text = '{"说明":"中文"}'
  const gbk = Buffer.concat([
    Buffer.from('{"name":"'),
    Buffer.from([0xd6, 0xd0, 0xce, 0xc4]),
    Buffer.from('"}'),
  ])
  for (const bytes of [Buffer.from(text), Buffer.from('\ufeff' + text, 'utf16le'), gbk]) {
    const before = Buffer.from(bytes)
    const decoded = plainTextPreview(bytes)!
    expect(decoded).toContain('中文')
    expect(textLanguage('.txt', decoded).language).toBe('json')
    expect(
      syntaxTokens('json', decoded)!
        .map((t) => t.text)
        .join(''),
    ).toBe(decoded)
    expect(bytes).toEqual(before)
  }
})

test('大文本、过多着色节点和未知语言退回原文，限制额外开销', () => {
  expect(syntaxTokens('json', ' '.repeat(MAX_SYNTAX_LENGTH + 1))).toBeNull()
  expect(syntaxTokens('json', '[' + '0,'.repeat(15_000) + '0]')).toBeNull()
  expect(syntaxTokens('missing-language', '中文')).toBeNull()
})
