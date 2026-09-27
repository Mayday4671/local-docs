import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { wordFixture, sheetFixture } from '../tests/fixtures/office.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const userData = await mkdtemp(join(results, 'library-split-'))
const errors = []
let app
async function launch() {
  const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: userData }
  delete env.ELECTRON_RUN_AS_NODE
  const executablePath = process.env.LOCAL_DOCS_TEST_EXECUTABLE
  app = await electron.launch({
    executablePath,
    args: executablePath ? [] : ['.'],
    cwd: executablePath ? userData : undefined,
    env,
  })
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  const page = await app.firstWindow()
  page.setDefaultTimeout(15000)
  page.on('pageerror', (error) => errors.push(String(error)))
  await page.emulateMedia({ colorScheme: null })
  await app.evaluate(({ BrowserWindow }) => {
    BrowserWindow.getAllWindows()[0].setContentSize(1536, 1000)
    BrowserWindow.getAllWindows()[0].showInactive()
  })
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  return page
}
const divider = (page) => page.getByRole('separator', { name: '调整文件列表与文档预览宽度' })
async function sizes(page) {
  return page.locator('.library-split').evaluate((root) => {
    const left = root.querySelector('.file-panel').getBoundingClientRect(),
      right = root.querySelector('.preview-panel')?.getBoundingClientRect()
    return {
      left: left.width,
      right: right?.width,
      ratio: right ? left.width / (left.width + right.width) : 1,
      overflow: document.documentElement.scrollWidth > innerWidth + 1,
    }
  })
}
async function dragTo(page, ratio) {
  const bar = await divider(page).boundingBox(),
    root = await page.locator('.library-split').boundingBox()
  await page.mouse.move(bar.x + bar.width / 2, bar.y + bar.height / 2)
  await page.mouse.down()
  await page.mouse.move(
    root.x + (root.width - bar.width) * ratio + bar.width / 2,
    bar.y + bar.height / 2,
    { steps: 12 },
  )
  await page.mouse.up()
}
async function previewToggle(page, name) {
  await page.getByRole('button', { name: '列表选项', exact: true }).click()
  await page.getByRole('button', { name, exact: true }).click()
}
async function actionsFit(page) {
  expect(
    await page.locator('.preview-actions').evaluate((bar) => {
      const panel = bar.closest('.preview-panel').getBoundingClientRect()
      return [...bar.querySelectorAll('button')].every((button) => {
        const box = button.getBoundingClientRect()
        return box.left >= panel.left && box.right <= panel.right && box.bottom <= panel.bottom
      })
    }),
  ).toBe(true)
}
try {
  let page = await launch()
  await dragTo(page, 0.7)
  await divider(page).dblclick()
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.616, 2)
  const md = await page.evaluate(async () => {
    const doc = await window.localDocs.createMarkdown('首页笔记', null)
    await window.localDocs.saveMarkdown(
      doc.id,
      '# 首页调整检查\n\n' +
        Array.from({ length: 70 }, (_, i) => `段落 ${i + 1}：保持阅读位置。`).join('\n\n'),
      1,
    )
    await window.localDocs.setTheme('light')
    return doc
  })
  const paths = [join(userData, '首页Word.docx'), join(userData, '首页Excel.xlsx')]
  await writeFile(paths[0], wordFixture())
  await writeFile(paths[1], sheetFixture())
  await app.evaluate(({ dialog }, paths) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
  }, paths)
  await page.evaluate(() => window.localDocs.importFiles(null))
  await page.reload()
  await page.getByRole('row', { name: /首页笔记.md/ }).click()
  await page.getByRole('checkbox', { name: '选择 首页笔记.md', exact: true }).check()
  await expect(page.locator('.reading-content .markdown-body')).toBeVisible()
  await page.locator('.reading-content').evaluate((e) => {
    e.scrollTop = 300
  })
  const scroll = await page.locator('.reading-content').evaluate((e) => e.scrollTop),
    sidebar = await page.locator('.library-sidebar').boundingBox()
  await dragTo(page, 0.4)
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.4, 2)
  expect(await page.locator('.reading-content').evaluate((e) => e.scrollTop)).toBeCloseTo(scroll, 0)
  expect((await page.locator('.library-sidebar').boundingBox()).width).toBe(sidebar.width)
  await expect(page.getByRole('checkbox', { name: '选择 首页笔记.md', exact: true })).toBeChecked()
  await page.screenshot({ path: join(results, 'library-split-light.png') })
  await page.getByRole('tab', { name: '标签', exact: true }).click()
  await page.getByLabel('文档备注').fill('尚未保存的备注')
  await dragTo(page, 0.7)
  await expect(page.getByLabel('文档备注')).toHaveValue('尚未保存的备注')
  expect(
    (await page.evaluate((id) => window.localDocs.readDocument(id), md.id)).document.notes,
  ).toBe('')
  await page.getByRole('tab', { name: '预览', exact: true }).click()
  await dragTo(page, 0.01)
  expect((await sizes(page)).left).toBeGreaterThanOrEqual(329)
  await expect(page.locator('.file-table th').nth(2)).toBeHidden()
  await expect(page.locator('.file-table th').nth(3)).toBeHidden()
  await expect(page.getByRole('button', { name: '批量收藏', exact: true })).toBeVisible()
  await page.getByRole('button', { name: '网格视图', exact: true }).click()
  await expect(page.locator('.file-card')).toHaveCount(3)
  await page.getByRole('button', { name: '列表视图', exact: true }).click()
  await dragTo(page, 0.99)
  expect((await sizes(page)).right).toBeGreaterThanOrEqual(329)
  await actionsFit(page)
  await expect(page.getByRole('button', { name: '复制路径', exact: true })).toBeVisible()
  await page.getByRole('row', { name: /首页Word.docx/ }).click()
  await expect(page.locator('.preview-panel .word-layout')).toBeVisible()
  await actionsFit(page)
  await page.getByRole('row', { name: /首页Excel.xlsx/ }).click()
  await expect(page.getByRole('button', { name: 'A2 核心交换机', exact: true })).toBeVisible()
  await actionsFit(page)
  await app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()[0].setContentSize(1000, 680),
  )
  await expect.poll(async () => (await sizes(page)).overflow).toBe(false)
  await expect.poll(async () => (await sizes(page)).right).toBeGreaterThanOrEqual(329)
  await divider(page).focus()
  await page.keyboard.press('Home')
  expect((await sizes(page)).left).toBeGreaterThanOrEqual(329)
  await page.keyboard.press('End')
  expect((await sizes(page)).right).toBeGreaterThanOrEqual(329)
  await page.getByRole('button', { name: '切换主题', exact: true }).click()
  await page.getByRole('menuitemradio', { name: '深色', exact: true }).click()
  await page.screenshot({ path: join(results, 'library-split-dark-small.png') })
  await actionsFit(page)
  await app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()[0].setContentSize(1536, 1000),
  )
  await dragTo(page, 0.55)
  await previewToggle(page, '收起预览')
  await expect(divider(page)).toHaveCount(0)
  await expect(page.locator('.preview-panel')).toHaveCount(0)
  expect((await sizes(page)).ratio).toBe(1)
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.55, 2)
  await page.getByRole('row', { name: /首页笔记.md/ }).dblclick()
  const editorDivider = page.getByRole('separator', { name: '调整编辑区与预览区宽度' })
  await expect(editorDivider).toHaveAttribute('aria-valuenow', '50')
  await editorDivider.focus()
  await page.keyboard.press('Shift+ArrowRight')
  await expect(editorDivider).toHaveAttribute('aria-valuenow', '60')
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.55, 2)
  await divider(page).focus()
  await page.keyboard.press('ArrowLeft')
  await expect(divider(page)).toHaveAttribute('aria-valuenow', '53')
  await app.close()
  app = undefined
  page = await launch()
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.53, 2)
  await page.getByRole('row', { name: /首页笔记.md/ }).dblclick()
  await expect(page.getByRole('separator', { name: '调整编辑区与预览区宽度' })).toHaveAttribute(
    'aria-valuenow',
    '60',
  )
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await divider(page).dblclick()
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.616, 2)
  expect(await page.evaluate((id) => window.localDocs.versions(id), md.id)).toHaveLength(2)
  expect(errors).toEqual([])
  console.log(
    `Library split passed: drag/reset/keyboard, collapse/reopen, narrow controls, all formats, unchanged selection/notes/scroll, separate editor ratio and restart. ${userData}`,
  )
} catch (error) {
  const page = app?.windows()[0]
  if (page)
    await page.screenshot({ path: join(results, 'library-split-failure.png') }).catch(() => {})
  throw error
} finally {
  if (app) await app.close()
}
