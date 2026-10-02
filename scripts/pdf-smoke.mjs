import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile, readFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { createHash } from 'node:crypto'
import { largePdfFixture, pdfFixture } from '../tests/fixtures/pdf.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'pdf-reader-'))
const large = largePdfFixture(),
  originalHash = createHash('sha256').update(large).digest('hex')
expect(large.length).toBeGreaterThan(30 * 1024 ** 2)
const paths = [
  join(root, '连续阅读.pdf'),
  join(root, '中文.pdf'),
  join(root, '损坏.pdf'),
  join(root, '接近上限.pdf'),
]
await writeFile(paths[0], large)
await writeFile(paths[1], pdfFixture())
await writeFile(paths[2], 'broken PDF')
const nearLimit = largePdfFixture(133)
expect(nearLimit.length).toBeGreaterThan(99 * 1024 ** 2)
expect(nearLimit.length).toBeLessThanOrEqual(100 * 1024 ** 2)
await writeFile(paths[3], nearLimit)
const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: join(root, 'data') }
delete env.ELECTRON_RUN_AS_NODE
const executablePath = process.env.LOCAL_DOCS_TEST_EXECUTABLE
const app = await electron.launch({ executablePath, args: executablePath ? [] : ['.'], env })
const page = await app.firstWindow()
await page.emulateMedia({ colorScheme: null })
page.setDefaultTimeout(20000)
const errors = [],
  network = []
page.on('pageerror', (error) => errors.push(String(error)))
page.on('request', (request) => {
  if (/^https?:/.test(request.url())) network.push(request.url())
})
const reader = page.locator('.pdf-reader:visible')
const viewport = reader.locator('.pdf-scroll')
const zoom = reader.getByLabel('PDF 缩放')
const canvas = (n) => reader.locator(`canvas[aria-label="PDF 第 ${n} 页"][data-rendered="true"]`)
async function wheel(dy, ctrl = false, dx = 0) {
  const box = await viewport.boundingBox()
  await page.mouse.move(box.x + box.width / 2, box.y + Math.min(100, box.height / 2))
  if (ctrl) await page.keyboard.down('Control')
  try {
    await page.mouse.wheel(dx, dy)
  } finally {
    if (ctrl) await page.keyboard.up('Control')
  }
}
try {
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  await app.evaluate(({ BrowserWindow, dialog, session }, paths) => {
    const win = BrowserWindow.getAllWindows()[0]
    win.setContentSize(1280, 800)
    win.showInactive()
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ['http://*/*', 'https://*/*'] },
      (_, cb) => cb({ cancel: true }),
    )
  }, paths)
  await page.getByRole('button', { name: '添加文件', exact: true }).click()
  await expect(page.locator('.file-table tbody tr')).toHaveCount(4)
  const baseline = await page.evaluate(() => window.localDocs.snapshot())
  const doc = baseline.documents.find((d) => d.name === '连续阅读.pdf')
  await page.getByRole('row', { name: /连续阅读.pdf/ }).dblclick()
  await expect(canvas(1)).toBeVisible()
  await expect(reader.locator('.pdf-tools')).toContainText('1 / 48 页')
  const baseWidth = (await canvas(1).boundingBox()).width
  expect(baseWidth).toBeGreaterThan(1000)
  expect(await viewport.evaluate((el) => el.scrollWidth - el.clientWidth)).toBeLessThanOrEqual(1)
  // Actual continuous wheel scrolling crosses page boundaries without using page buttons.
  await wheel(1800)
  await expect.poll(() => viewport.evaluate((el) => el.scrollTop)).toBeGreaterThan(1500)
  await expect(reader.locator('.pdf-tools')).not.toContainText('1 / 48 页')
  await reader.getByRole('button', { name: '上一页', exact: true }).click()
  await expect(reader.locator('.pdf-tools')).toContainText('1 / 48 页')
  await wheel(-100, true)
  await expect(zoom).toHaveValue('110')
  await expect(canvas(1)).toBeVisible()
  await expect
    .poll(async () => (await canvas(1).boundingBox()).width)
    .toBeGreaterThan(baseWidth * 1.09)
  await zoom.selectOption('200')
  await expect(canvas(1)).toBeVisible()
  await expect
    .poll(() => viewport.evaluate((el) => el.scrollWidth / el.clientWidth))
    .toBeGreaterThan(1.9)
  const left = await viewport.evaluate((el) => el.scrollLeft)
  await wheel(0, false, 250)
  await expect.poll(() => viewport.evaluate((el) => el.scrollLeft)).toBeGreaterThan(left)
  // Scroll far enough to evict earlier canvases; distant, mixed-size pages load on demand.
  await zoom.selectOption('100')
  await viewport.evaluate((el) => {
    el.scrollTop = el.scrollHeight
  })
  await expect(canvas(48)).toBeVisible()
  await expect.poll(() => reader.locator('canvas').count()).toBeLessThan(8)
  await expect(canvas(1)).toHaveCount(0)
  await expect(reader.locator('.pdf-tools')).toContainText('48 / 48 页')
  await reader.getByRole('button', { name: '上一页', exact: true }).click()
  await expect(reader.locator('.pdf-tools')).toContainText('47 / 48 页')
  await wheel(-100, true)
  await expect(zoom).toHaveValue('110')
  await expect(reader.locator('.pdf-tools')).toContainText('47 / 48 页')
  await expect(reader.getByRole('alert')).toHaveCount(0)
  await zoom.selectOption('50')
  await expect(reader.locator('.pdf-tools')).toContainText('47 / 48 页')
  await expect(canvas(47)).toBeVisible()
  await expect(canvas(48)).toBeVisible()
  await page.screenshot({ path: join(results, 'pdf-continuous-full.png') })
  await viewport.evaluate((el) => {
    el.scrollTop = 0
  })
  await zoom.selectOption('100')
  await expect(reader.locator('.pdf-tools')).toContainText('1 / 48 页')
  for (let n = 2; n <= 11; n++) {
    await reader.getByRole('button', { name: '下一页', exact: true }).click()
    await expect(reader.locator('.pdf-tools')).toContainText(`${n} / 48 页`)
    await expect(canvas(n)).toBeVisible()
    if (n === 7 || n === 11) {
      const box = await canvas(n).boundingBox()
      expect(box.width).toBeGreaterThan(box.height)
    }
  }
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  await expect(canvas(1)).toBeVisible()
  await page.getByRole('button', { name: '切换主题', exact: true }).click()
  await page.getByRole('menuitemradio', { name: '深色', exact: true }).click()
  const sidebarWidth = (await canvas(1).boundingBox()).width
  await wheel(-100, true)
  await expect(zoom).toHaveValue('110')
  await expect
    .poll(async () => (await canvas(1).boundingBox()).width)
    .toBeGreaterThan(sidebarWidth * 1.09)
  await wheel(600)
  await expect.poll(() => viewport.evaluate((el) => el.scrollTop)).toBeGreaterThan(400)
  await expect(reader.locator('.pdf-tools')).not.toContainText('1 / 48 页')
  await zoom.selectOption('200')
  await expect.poll(() => viewport.evaluate((el) => el.scrollWidth > el.clientWidth)).toBe(true)
  await page.screenshot({ path: join(results, 'pdf-continuous-sidebar-dark.png') })
  await page.getByRole('row', { name: /中文.pdf/ }).click()
  await expect(canvas(1)).toBeVisible()
  await reader.getByRole('button', { name: '下一页', exact: true }).click()
  await expect(reader.locator('pre')).toContainText('离线预览测试')
  await expect(canvas(2)).toBeVisible()
  await page.getByRole('row', { name: /损坏.pdf/ }).click()
  await expect(page.locator('.preview-panel .file-fallback')).toContainText('原文件仍完整保留')
  await page.getByRole('row', { name: /连续阅读.pdf/ }).click()
  await expect(canvas(1)).toBeVisible()
  expect(
    await app.evaluate(({ BrowserWindow }) =>
      BrowserWindow.getAllWindows()[0].webContents.getZoomFactor(),
    ),
  ).toBe(1)
  const exported = join(root, 'export.pdf')
  await app.evaluate(({ dialog }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath })
  }, exported)
  await page.evaluate((id) => window.localDocs.exportDocument(id), doc.id)
  for (const file of [paths[0], exported])
    expect(
      createHash('sha256')
        .update(await readFile(file))
        .digest('hex'),
    ).toBe(originalHash)
  const after = await page.evaluate(() => window.localDocs.snapshot())
  expect(after.documents.map((d) => [d.id, d.revision])).toEqual(
    baseline.documents.map((d) => [d.id, d.revision]),
  )
  expect(network).toEqual([])
  expect(errors).toEqual([])
  await page.getByRole('row', { name: /接近上限.pdf/ }).dblclick()
  await expect(page.locator('.editor-page')).toBeVisible()
  await expect(canvas(1)).toBeVisible()
  await viewport.evaluate((el) => {
    el.scrollTop = el.scrollHeight
  })
  await expect(canvas(133)).toBeVisible()
  await expect(reader.locator('.pdf-tools')).toContainText('133 / 133 页')
  expect(errors).toEqual([])
  console.log(
    `PDF 36 MB / 48 pages and 99 MB / 133 pages: offline range reading, continuous wheel, current page, lazy canvases, mixed/rotated pages, Ctrl wheel width, horizontal scrolling, sidebar/fullscreen, CJK, corrupt fallback and original/export hashes passed: ${root}`,
  )
} catch (error) {
  console.log(
    await reader
      .evaluate((el) => {
        const scroll = el.querySelector('.pdf-scroll')
        return {
          top: scroll.scrollTop,
          height: scroll.clientHeight,
          scrollHeight: scroll.scrollHeight,
          pages: [...el.querySelectorAll('.pdf-page')].map((p) => ({
            n: p.dataset.pageNumber,
            top: p.offsetTop,
            height: p.offsetHeight,
          })),
          tools: el.querySelector('.pdf-tools').textContent,
        }
      })
      .catch(() => null),
  )
  await page.screenshot({ path: join(results, 'pdf-reader-failure.png') }).catch(() => {})
  throw error
} finally {
  await app.close()
}
