import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile, readFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { pdfFixture } from '../tests/fixtures/pdf.ts'
import { wordFixture } from '../tests/fixtures/office.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'zoom-'))
const pdfBytes = Buffer.from(pdfFixture()),
  wordBytes = Buffer.from(wordFixture())
const paths = [join(root, '滚轮测试.pdf'), join(root, '滚轮测试.docx')]
await writeFile(paths[0], pdfBytes)
await writeFile(paths[1], wordBytes)
const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: join(root, 'data') }
delete env.ELECTRON_RUN_AS_NODE
const executablePath = process.env.LOCAL_DOCS_TEST_EXECUTABLE
const app = await electron.launch({ executablePath, args: executablePath ? [] : ['.'], env })
const page = await app.firstWindow()
page.setDefaultTimeout(15000)
await page.emulateMedia({ colorScheme: null })
const errors = []
page.on('pageerror', (error) => errors.push(String(error)))
page.on('console', (message) => {
  if (message.type() === 'error') errors.push(message.text())
})
async function wheel(target, delta, ctrl = true) {
  const box = await target.boundingBox()
  await page.mouse.move(box.x + Math.min(100, box.width / 2), box.y + Math.min(150, box.height / 2))
  if (ctrl) await page.keyboard.down('Control')
  try {
    await page.mouse.wheel(0, delta)
  } finally {
    if (ctrl) await page.keyboard.up('Control')
  }
}
async function singleLineOptions(name, screenshot) {
  const select = page.getByRole('combobox', { name, exact: true })
  await select.click()
  await expect(page.locator('select:open')).toHaveCount(1)
  const options = await select.locator('option').evaluateAll((items) =>
    items.map((item) => {
      const range = document.createRange()
      range.selectNodeContents(item)
      return {
        label: item.textContent.trim(),
        lines: new Set([...range.getClientRects()].map((rect) => Math.round(rect.top))).size,
        height: item.getBoundingClientRect().height,
      }
    }),
  )
  for (const option of options) {
    expect(option.label).toMatch(/^\d+%$/)
    expect(option.lines).toBe(1)
    expect(option.height).toBeLessThan(40)
  }
  await page.screenshot({ path: join(results, screenshot) })
  await page.keyboard.press('Escape')
  await expect(page.locator('select:open')).toHaveCount(0)
}
const waitForPdf = async () => {
  await expect(page.locator('.pdf-reader:visible canvas').first()).toBeVisible()
  await expect(page.locator('.pdf-reader:visible [role="status"]')).toHaveCount(0)
}
try {
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  await app.evaluate(({ BrowserWindow, dialog }, paths) => {
    const win = BrowserWindow.getAllWindows()[0]
    win.setContentSize(1280, 800)
    win.showInactive()
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
  }, paths)
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  await page.evaluate(async () => {
    await window.localDocs.importFiles(null)
    await window.localDocs.setTheme('dark')
  })
  await page.reload()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  const baseline = await page.evaluate(() => window.localDocs.snapshot())
  const browserZoom = await app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()[0].webContents.getZoomFactor(),
  )
  await page.getByRole('row', { name: /滚轮测试.pdf/ }).click()
  await waitForPdf()
  const zoom = page.getByRole('combobox', { name: 'PDF 缩放', exact: true })
  const content = page.locator('.pdf-reader:visible .pdf-scroll')
  const width = (await page.locator('.pdf-reader:visible canvas').first().boundingBox()).width
  await wheel(content, -100)
  await expect(zoom).toHaveValue('110')
  await waitForPdf()
  await expect
    .poll(
      async () =>
        (await page.locator('.pdf-reader:visible canvas').first().boundingBox())?.width || 0,
    )
    .toBeGreaterThan(width)
  // At a scrollbar boundary, rendering must settle instead of repeatedly replacing the canvas.
  const stable = await page.locator('.pdf-reader:visible').evaluate(async (el) => {
    const canvas = el.querySelector('canvas')
    for (let frame = 0; frame < 30; frame++) {
      await new Promise(requestAnimationFrame)
      if (el.querySelector('canvas') !== canvas || el.querySelector('[role="status"]')) return false
    }
    return true
  })
  expect(stable).toBe(true)
  await wheel(content, 100)
  await expect(zoom).toHaveValue('100')
  await zoom.selectOption('125')
  await wheel(content, -100)
  await expect(zoom).toHaveValue('135')
  await singleLineOptions('PDF 缩放', 'zoom-pdf-sidebar-dark.png')

  // The preview uses the same reader as the full page, and the browser UI stays at its original size.
  expect(
    await app.evaluate(({ BrowserWindow }) =>
      BrowserWindow.getAllWindows()[0].webContents.getZoomFactor(),
    ),
  ).toBe(browserZoom)
  await page.getByRole('button', { name: '打开', exact: true }).click()
  await expect(page.locator('.editor-page')).toBeVisible()
  await waitForPdf()
  await page.locator('.editor-page').getByRole('button', { name: '下一页', exact: true }).click()
  await expect(page.locator('.pdf-tools:visible')).toContainText('2 / 2 页')
  await waitForPdf()
  await wheel(content, -100)
  await expect(zoom).toHaveValue('110')
  await expect(page.locator('.pdf-tools:visible')).toContainText('2 / 2 页')
  await wheel(content, -10000)
  await expect(zoom).toHaveValue('300')
  await wheel(content, 10000)
  await expect(zoom).toHaveValue('50')
  await waitForPdf()

  // Pixel/line/page units and small trackpad deltas use one bounded zoom state.
  await zoom.selectOption('100')
  for (const [deltaMode, deltaY, expected] of [
    [1, -3, '110'],
    [2, -1, '120'],
    [0, -50, '120'],
    [0, -50, '130'],
  ]) {
    await page.locator('.pdf-reader:visible').evaluate(
      (el, { deltaMode, deltaY }) => {
        const event = new WheelEvent('wheel', {
          ctrlKey: true,
          deltaMode,
          deltaY,
          bubbles: true,
          cancelable: true,
        })
        el.dispatchEvent(event)
        if (!event.defaultPrevented) throw Error('Ctrl wheel did not suppress browser zoom')
      },
      { deltaMode, deltaY },
    )
    await expect(zoom).toHaveValue(expected)
  }
  await zoom.selectOption('200')
  await waitForPdf()
  const scrollBefore = await content.evaluate((el) => el.scrollTop)
  await wheel(content, 180, false)
  await expect.poll(() => content.evaluate((el) => el.scrollTop)).toBeGreaterThan(scrollBefore)
  await expect(zoom).toHaveValue('200')
  await page.getByRole('button', { name: '切换主题', exact: true }).click()
  await page.getByRole('menuitemradio', { name: '浅色', exact: true }).click()
  await singleLineOptions('PDF 缩放', 'zoom-pdf-full-light.png')
  await page.getByRole('button', { name: '返回列表', exact: true }).click()

  await page.getByRole('row', { name: /滚轮测试.docx/ }).click()
  await page.getByRole('button', { name: '编辑', exact: true }).click()
  const title = page.getByLabel('编辑 正文 · 段落 1', { exact: true })
  await title.fill('滚轮缩放保留未保存的文字')
  await wheel(page.locator('.word-edit-scroll'), -100)
  const wordZoom = page.getByRole('combobox', { name: '文档缩放', exact: true })
  await expect(wordZoom).toHaveValue('110')
  await expect(title).toHaveText('滚轮缩放保留未保存的文字')
  await expect(page.locator('.word-edit-layout')).toHaveCSS('zoom', '1.1')
  await wordZoom.selectOption('125')
  await wheel(page.locator('.word-edit-scroll'), 100)
  await expect(wordZoom).toHaveValue('115')
  await expect(title).toHaveText('滚轮缩放保留未保存的文字')
  await singleLineOptions('文档缩放', 'zoom-word-light.png')
  expect(
    await app.evaluate(({ BrowserWindow }) =>
      BrowserWindow.getAllWindows()[0].webContents.getZoomFactor(),
    ),
  ).toBe(browserZoom)
  const after = await page.evaluate(() => window.localDocs.snapshot())
  expect(after.documents.map((doc) => [doc.id, doc.revision])).toEqual(
    baseline.documents.map((doc) => [doc.id, doc.revision]),
  )
  expect(await readFile(paths[0])).toEqual(pdfBytes)
  expect(await readFile(paths[1])).toEqual(wordBytes)
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('button', { name: '放弃修改并返回', exact: true }).click()
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  expect(errors).toEqual([])
  console.log(
    `Ctrl wheel zoom passed: PDF sidebar/full page, bounds, native scrolling, delta modes, synchronized single-line percentages, Word drafts, unchanged browser zoom and originals. ${root}`,
  )
} catch (error) {
  await page.screenshot({ path: join(results, 'zoom-failure.png') }).catch(() => {})
  throw error
} finally {
  await app.close()
}
