import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile, readFile } from 'node:fs/promises'
import { resolve, join } from 'node:path'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'text-syntax-'))
const sql = "-- 中文查询\r\nSELECT name, 42 FROM documents WHERE name = '中文';\r\n"
const json =
  '{\r\n  "名称": "中文",\r\n  "count": 42, "enabled": true,\r\n  "example": "<script>window.injected = true</script>"\r\n}'
const gbk = Buffer.concat([
  Buffer.from('-- '),
  Buffer.from([0xd6, 0xd0, 0xce, 0xc4]),
  Buffer.from('\r\nSELECT 42 FROM documents;'),
])
const fixtures = {
  '查询.sql': { bytes: Buffer.from(sql), text: sql, language: 'sql' },
  '配置.json': { bytes: Buffer.from(json), text: json, language: 'json' },
  'UTF16.json': { bytes: Buffer.from('\ufeff' + json, 'utf16le'), text: json, language: 'json' },
  'UTF16BE.json': {
    bytes: Buffer.from('\ufeff' + json, 'utf16le').swap16(),
    text: json,
    language: 'json',
  },
  'GBK.sql': { bytes: gbk, text: '-- 中文\r\nSELECT 42 FROM documents;', language: 'sql' },
  '自动JSON.txt': { bytes: Buffer.from(json), text: json, language: 'json' },
  '自动SQL.txt': { bytes: Buffer.from(sql), text: sql, language: 'sql' },
  '说明.txt': { bytes: Buffer.from('普通中文说明\r\n  保留缩进'), language: 'plaintext' },
  '空白.sql': { bytes: Buffer.alloc(0), text: '', language: 'sql' },
  '不完整.json': { bytes: Buffer.from('{"名称": "中文",'), language: 'json' },
  '设置.yaml': { bytes: Buffer.from('name: 中文\ncount: 42\nenabled: true\n'), language: 'yaml' },
  '脚本.js': {
    bytes: Buffer.from('const name = "中文";\nconsole.log(name);'),
    language: 'javascript',
  },
  '页面.html': { bytes: Buffer.from('<script>window.injected = true</script>'), language: 'xml' },
  '大文本.sql': {
    bytes: Buffer.from('-- 中文\n' + 'SELECT 42 FROM documents;\n'.repeat(15_000)),
    language: 'sql',
  },
}
for (const [name, fixture] of Object.entries(fixtures)) {
  fixture.text ??= fixture.bytes.toString()
  await writeFile(join(root, name), fixture.bytes)
}
let app
const errors = [],
  network = [],
  consoles = []
try {
  const executablePath = process.env.LOCAL_DOCS_TEST_EXECUTABLE
  const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: join(root, 'data') }
  delete env.ELECTRON_RUN_AS_NODE
  app = await electron.launch({
    executablePath,
    args: executablePath ? [] : ['.'],
    cwd: executablePath ? root : undefined,
    env,
  })
  const page = await app.firstWindow()
  page.setDefaultTimeout(20_000)
  await page.emulateMedia({ colorScheme: null })
  page.on('pageerror', (e) => errors.push(String(e)))
  page.on('request', (r) => {
    if (/^https?:/.test(r.url())) network.push(r.url())
  })
  page.on('console', (m) => {
    if (m.type() === 'error') consoles.push(m.text())
  })
  await app.evaluate(({ BrowserWindow, session }) => {
    BrowserWindow.getAllWindows()[0].showInactive()
    BrowserWindow.getAllWindows()[0].setSize(1500, 950)
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ['http://*/*', 'https://*/*'] },
      (_, callback) => callback({ cancel: true }),
    )
  })
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  await app.evaluate(
    ({ dialog }, filePaths) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths })
    },
    Object.keys(fixtures).map((name) => join(root, name)),
  )
  await page.getByRole('button', { name: '添加文件', exact: true }).click()
  await expect(page.locator('.file-table tbody tr')).toHaveCount(Object.keys(fixtures).length)
  await page.evaluate(() => window.localDocs.setTheme('light'))
  const pre = page.locator('.preview-panel .plain-text-reader')
  for (const [name, fixture] of Object.entries(fixtures)) {
    await page
      .getByRole('row')
      .filter({ has: page.getByText(name, { exact: true }) })
      .click()
    await expect(page.locator('.preview-panel .text-language')).toHaveAttribute(
      'data-language',
      fixture.language,
    )
    await expect(pre).toHaveJSProperty('textContent', fixture.text)
    if (name === '大文本.sql') {
      await expect(page.locator('.preview-panel .text-reader-meta')).toContainText(
        '当前按普通文本显示',
      )
      await expect(pre.locator('span')).toHaveCount(0)
    } else if (fixture.language !== 'plaintext' && fixture.text) {
      await expect(pre.locator('[class*="hljs-"]').first()).toBeVisible()
    }
    expect(await page.evaluate(() => window.injected)).toBeUndefined()
  }
  await page
    .getByRole('row')
    .filter({ has: page.getByText('查询.sql', { exact: true }) })
    .dblclick()
  const editor = page.locator('.editor-page')
  const reader = editor.locator('.plain-text-reader')
  await expect(reader.locator('.hljs-keyword').first()).toBeVisible()
  await expect(reader).toHaveJSProperty('textContent', sql)
  await editor.getByLabel('文内查找', { exact: true }).fill('name, 42 FROM')
  await expect(editor.getByRole('button', { name: '1/1 下一处', exact: true })).toBeVisible()
  await editor.getByLabel('文内查找', { exact: true }).fill('')
  const quote = 'SELECT name, 42 FROM documents'
  const copied = await reader.evaluate((element, quote) => {
    const at = element.textContent.indexOf(quote)
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT)
    const range = document.createRange()
    let node,
      offset = 0,
      started = false
    while ((node = walker.nextNode())) {
      const end = offset + node.textContent.length
      if (!started && at < end) {
        range.setStart(node, at - offset)
        started = true
      }
      if (started && at + quote.length <= end) {
        range.setEnd(node, at + quote.length - offset)
        break
      }
      offset = end
    }
    const selection = window.getSelection()
    selection.removeAllRanges()
    selection.addRange(range)
    element.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }))
    return selection.toString()
  }, quote)
  expect(copied).toBe(quote)
  await editor.getByRole('button', { name: '标记所选内容', exact: true }).click()
  await editor.getByLabel('标记备注', { exact: true }).fill('跨语法颜色的内容标记')
  await editor.getByRole('button', { name: '保存标记', exact: true }).click()
  await expect(page.getByRole('complementary', { name: '内容标记' })).toContainText(
    '跨语法颜色的内容标记',
  )
  await page.screenshot({ path: join(results, 'text-syntax-sql-light.png') })
  const light = await reader
    .locator('.hljs-keyword')
    .first()
    .evaluate((el) => getComputedStyle(el).color)
  await page.evaluate(() => window.localDocs.setTheme('dark'))
  await expect
    .poll(() =>
      reader
        .locator('.hljs-keyword')
        .first()
        .evaluate((el) => getComputedStyle(el).color),
    )
    .not.toBe(light)
  await page.screenshot({ path: join(results, 'text-syntax-sql-dark.png') })
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  // Rapid switching must not apply a late SQL result to the next file.
  for (const name of ['查询.sql', '配置.json', '说明.txt', '配置.json'])
    await page
      .getByRole('row')
      .filter({ has: page.getByText(name, { exact: true }) })
      .click()
  await expect(pre.locator('.hljs-attr').first()).toBeVisible()
  await expect(pre).toHaveJSProperty('textContent', json)
  await page.screenshot({ path: join(results, 'text-syntax-json-dark.png') })
  const snapshot = await page.evaluate(() => window.localDocs.snapshot())
  for (const doc of snapshot.documents) {
    const exported = join(root, 'export-' + doc.name)
    await app.evaluate(({ dialog }, filePath) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath })
    }, exported)
    await page.evaluate((id) => window.localDocs.exportDocument(id), doc.id)
    expect(await readFile(exported)).toEqual(fixtures[doc.name].bytes)
    expect(await readFile(join(root, doc.name))).toEqual(fixtures[doc.name].bytes)
  }
  expect(await page.evaluate(() => window.localDocs.search('中文'))).not.toHaveLength(0)
  expect(network).toEqual([])
  expect(errors).toEqual([])
  expect(
    consoles.filter((m) => /Content Security Policy|Failed to fetch|Failed to load|Worker/.test(m)),
  ).toEqual([])
  console.log(
    `Offline syntax, encodings, themes, exact text, cross-token search/marks, rapid switching and byte-identical exports passed: ${root}`,
  )
} finally {
  if (app) await app.close()
}
