import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile, readFile } from 'node:fs/promises'
import { resolve, join } from 'node:path'
import { pdfFixture } from '../tests/fixtures/pdf.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'file-import-'))
const fixtures = {
  '阅读.pdf': pdfFixture(),
  '说明.txt': Buffer.from('中文文本\r\n  保留缩进\n独有搜索词'),
  '查询.sql': Buffer.from(
    '-- 离线 SQL\nSELECT * FROM synthetic_table;\n<script>window.injected = true</script>',
  ),
  '配置.json': Buffer.from('{"offline":true}'),
  '图片.png': Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/lYcAAAAASUVORK5CYII=',
    'base64',
  ),
  '归档.zip': Buffer.from([0x50, 0x4b, 0, 1, 2, 3]),
  '拖入.txt': Buffer.from('拖入的合成文本'),
  '拖入.sql': Buffer.from('SELECT 2;'),
  '损坏.pdf': Buffer.from('not a pdf'),
  无扩展名: Buffer.from('original binary archive placeholder'),
}
for (const [name, bytes] of Object.entries(fixtures)) await writeFile(join(root, name), bytes)
await mkdir(join(root, '文件夹'))
const errors = [],
  network = [],
  consoleErrors = []
let app
try {
  const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: join(root, 'data') }
  delete env.ELECTRON_RUN_AS_NODE
  const executablePath = process.env.LOCAL_DOCS_TEST_EXECUTABLE
  app = await electron.launch({
    executablePath,
    args: executablePath ? [] : ['.'],
    env,
    cwd: executablePath ? root : undefined,
  })
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  const page = await app.firstWindow()
  page.setDefaultTimeout(20000)
  page.on('pageerror', (e) => errors.push(String(e)))
  page.on('console', (m) => {
    if (m.type() === 'error') consoleErrors.push(m.text())
  })
  page.on('request', (r) => {
    if (/^https?:/.test(r.url())) network.push(r.url())
  })
  await app.evaluate(({ BrowserWindow, session }) => {
    BrowserWindow.getAllWindows()[0].showInactive()
    BrowserWindow.getAllWindows()[0].setSize(1500, 950)
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ['http://*/*', 'https://*/*'] },
      (_, cb) => cb({ cancel: true }),
    )
  })
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  await expect(page.getByRole('button', { name: /^未分类/ })).toHaveCount(0)
  await app.evaluate(
    ({ dialog }, paths) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
    },
    Object.keys(fixtures)
      .filter((n) => !n.startsWith('拖入'))
      .map((n) => join(root, n)),
  )
  await page.getByRole('button', { name: '添加文件', exact: true }).click()
  await expect(page.locator('.file-table tbody tr')).toHaveCount(8)
  await page.getByRole('row', { name: /说明.txt/ }).click()
  await expect(page.locator('.preview-panel .plain-text-reader')).toHaveText(
    fixtures['说明.txt'].toString(),
  )
  await expect(page.getByRole('button', { name: '编辑', exact: true })).toHaveCount(0)
  await page.getByRole('row', { name: /查询.sql/ }).dblclick()
  await expect(page.locator('.editor-page .plain-text-reader')).toContainText(
    '<script>window.injected = true</script>',
  )
  expect(await page.evaluate(() => window.injected)).toBeUndefined()
  await expect(
    page.locator('.editor-page').getByRole('button', { name: '保存', exact: true }),
  ).toHaveCount(0)
  await page.screenshot({ path: join(results, 'file-import-sql.png') })
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  const hits = await page.evaluate(() => window.localDocs.search('独有搜索词'))
  expect(hits).toHaveLength(1)
  await page.getByRole('row', { name: /阅读.pdf/ }).dblclick()
  const pdf = page.locator('.editor-page .pdf-reader')
  await expect(pdf.getByText('1 / 2 页', { exact: true })).toBeVisible()
  await expect(pdf.locator('canvas')).toBeVisible()
  await expect(pdf.locator('pre')).toContainText('Offline PDF - page one')
  await pdf.getByRole('button', { name: '下一页' }).click()
  await expect(pdf.locator('pre')).toContainText('离线预览测试')
  await pdf.getByLabel('PDF 缩放').selectOption('150')
  await expect(pdf.locator('canvas')).toBeVisible()
  await expect(pdf.getByRole('status')).toHaveCount(0)
  expect(
    await pdf.locator('canvas').evaluate((c) => {
      const data = c.getContext('2d').getImageData(0, 0, c.width, c.height).data
      return data.some((v, i) => i % 4 !== 3 && v < 100)
    }),
  ).toBe(true)
  await page.screenshot({ path: join(results, 'file-import-pdf.png') })
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('row', { name: /损坏.pdf/ }).click()
  await expect(page.locator('.preview-panel .file-fallback')).toContainText('原文件仍完整保留')
  await page.getByRole('row', { name: /图片.png/ }).click()
  await expect(page.locator('.image-reader img')).toBeVisible()
  expect(await page.locator('.image-reader img').evaluate((i) => i.naturalWidth)).toBe(1)
  await page.getByRole('row', { name: /归档.zip/ }).click()
  await expect(page.locator('.preview-panel .file-fallback')).toContainText(
    '此格式暂不支持内置预览',
  )

  const category = await page.evaluate(() => window.localDocs.createCategory('拖入目标', null))
  await page.reload()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  await page.getByRole('button', { name: /^拖入目标/ }).click()
  const cdp = await page.context().newCDPSession(page)
  const drag = async (names) => {
    const data = { items: [], files: names.map((n) => join(root, n)), dragOperationsMask: 1 }
    await cdp.send('Input.dispatchDragEvent', { type: 'dragEnter', x: 650, y: 350, data })
    await cdp.send('Input.dispatchDragEvent', { type: 'dragOver', x: 650, y: 350, data })
    await expect(page.locator('.file-drop-overlay')).toContainText('拖入目标')
    await page.screenshot({ path: join(results, 'file-import-drop.png') })
    await cdp.send('Input.dispatchDragEvent', { type: 'drop', x: 650, y: 350, data })
  }
  await drag(['拖入.txt', '拖入.sql', '文件夹'])
  await expect(page.locator('.file-table tbody tr')).toHaveCount(2)
  const snapshot = await page.evaluate(() => window.localDocs.snapshot())
  const dropped = snapshot.documents.filter((d) => d.name.startsWith('拖入'))
  expect(dropped).toHaveLength(2)
  expect(dropped.every((d) => d.categoryId === category.id)).toBe(true)
  expect(snapshot.documents).toHaveLength(10)
  await expect(page.getByRole('alert')).toContainText('请选择普通文件')
  await expect(page.locator('.file-drop-overlay')).toHaveCount(0)
  for (const d of snapshot.documents) {
    const exported = join(root, `export-${d.name}`)
    await app.evaluate(({ dialog }, filePath) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath })
    }, exported)
    expect(await page.evaluate((id) => window.localDocs.exportDocument(id), d.id)).toBe(true)
    expect(await readFile(exported)).toEqual(fixtures[d.name])
    expect(await readFile(join(root, d.name))).toEqual(fixtures[d.name])
  }
  await page.evaluate(() => window.localDocs.setTheme('dark'))
  await page.screenshot({ path: join(results, 'file-import-dark.png') })
  expect(network).toEqual([])
  expect(errors).toEqual([])
  expect(
    consoleErrors.filter((m) => /Content Security Policy|Failed to fetch|Failed to load/.test(m)),
  ).toEqual([])
  console.log(
    `File picker, native file drop, category routing, offline PDF/CJK, text, image, fallback, search and byte-identical exports passed: ${root}`,
  )
} finally {
  if (app) await app.close()
}
