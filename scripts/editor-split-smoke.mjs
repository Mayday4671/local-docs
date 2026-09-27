import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp } from 'node:fs/promises'
import { join, resolve } from 'node:path'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const userData = await mkdtemp(join(results, 'editor-split-'))
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
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].showInactive())
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  return page
}
const divider = (page) => page.getByRole('separator', { name: '调整编辑区与预览区宽度' })
async function sizes(page) {
  return page.locator('.editor-split').evaluate((root) => {
    const left = root.querySelector('.editor-source').getBoundingClientRect()
    const right = root.querySelector('.editor-preview').getBoundingClientRect()
    return {
      left: left.width,
      right: right.width,
      ratio: left.width / (left.width + right.width),
      overflow: document.documentElement.scrollWidth > innerWidth + 1,
    }
  })
}
async function dragTo(page, ratio) {
  const handle = await divider(page).boundingBox(),
    root = await page.locator('.editor-split').boundingBox()
  await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2)
  await page.mouse.down()
  await page.mouse.move(
    root.x + (root.width - handle.width) * ratio + handle.width / 2,
    handle.y + handle.height / 2,
    { steps: 15 },
  )
  await page.mouse.up()
}
try {
  let page = await launch()
  const sample =
    '# 宽度调整检查\n\n| 项目 | 状态 |\n| --- | --- |\n| 本地保存 | 正常 |\n\n' +
    Array.from({ length: 80 }, (_, i) => `段落 ${i + 1}：编辑内容保持不变。\n`).join('\n')
  const doc = await page.evaluate(async (text) => {
    const doc = await window.localDocs.createMarkdown('宽度检查', null)
    await window.localDocs.saveMarkdown(doc.id, text, 1)
    await window.localDocs.setTheme('light')
    return doc
  }, sample)
  await page.reload()
  await page.getByRole('row', { name: /宽度检查.md/ }).dblclick()
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.5, 2)
  const source = page.getByLabel('Markdown 内容')
  const draft = sample + '\n未保存内容也应保留。'
  await source.fill(draft)
  await source.evaluate((e) => {
    e.scrollTop = 350
  })
  const scrollBefore = await source.evaluate((e) => e.scrollTop)
  await dragTo(page, 0.7)
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.7, 2)
  await expect(source).toHaveValue(draft)
  expect(await source.evaluate((e) => e.scrollTop)).toBeCloseTo(scrollBefore, 0)
  expect(await page.evaluate((id) => window.localDocs.versions(id), doc.id)).toHaveLength(2)
  await page.screenshot({ path: join(results, 'editor-split-light.png') })
  // Dragging beyond either edge clamps panes instead of hiding content.
  await dragTo(page, 0.99)
  expect((await sizes(page)).right).toBeGreaterThanOrEqual(279)
  await dragTo(page, 0.01)
  expect((await sizes(page)).left).toBeGreaterThanOrEqual(279)
  await divider(page).dblclick()
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.5, 2)
  await divider(page).focus()
  await page.keyboard.press('ArrowRight')
  await expect(divider(page)).toHaveAttribute('aria-valuenow', '52')
  await page.keyboard.press('Shift+ArrowRight')
  await expect(divider(page)).toHaveAttribute('aria-valuenow', '62')
  // Cancel an in-progress drag with Escape.
  const handle = await divider(page).boundingBox()
  await page.mouse.move(handle.x + 4, handle.y + 100)
  await page.mouse.down()
  await page.mouse.move(handle.x - 100, handle.y + 100, { steps: 5 })
  await page.keyboard.press('Escape')
  await page.mouse.up()
  await expect(divider(page)).toHaveAttribute('aria-valuenow', '62')
  await page.getByRole('button', { name: '历史版本', exact: true }).click()
  await app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()[0].setContentSize(1000, 680),
  )
  await expect.poll(async () => (await sizes(page)).overflow).toBe(false)
  await expect.poll(async () => (await sizes(page)).right).toBeGreaterThanOrEqual(279)
  await divider(page).focus()
  await page.keyboard.press('Home')
  expect((await sizes(page)).left).toBeGreaterThanOrEqual(279)
  await page.keyboard.press('End')
  expect((await sizes(page)).right).toBeGreaterThanOrEqual(279)
  await page.keyboard.press('Enter')
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.5, 2)
  await page.getByRole('button', { name: '切换主题', exact: true }).click()
  await page.getByRole('menuitemradio', { name: '深色', exact: true }).click()
  await expect(source).toHaveValue(draft)
  await page.screenshot({ path: join(results, 'editor-split-dark-history.png') })
  await page.getByRole('button', { name: '关闭历史版本', exact: true }).click()
  await app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()[0].setContentSize(1536, 1000),
  )
  await dragTo(page, 0.65)
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('已保存在本地', { exact: true })).toBeVisible()
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('row', { name: /宽度检查.md/ }).dblclick()
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.65, 2)
  await app.close()
  app = undefined
  page = await launch()
  await page.getByRole('row', { name: /宽度检查.md/ }).dblclick()
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.65, 2)
  await expect(page.getByLabel('Markdown 内容')).toHaveValue(draft)
  // Invalid UI preferences cannot break the editor.
  await page.evaluate(() => localStorage.setItem('local-docs:markdown-editor-ratio', 'NaN'))
  await page.reload()
  await page.getByRole('row', { name: /宽度检查.md/ }).dblclick()
  await expect.poll(async () => (await sizes(page)).ratio).toBeCloseTo(0.5, 2)
  expect(errors).toEqual([])
  console.log(
    `Editor split passed: drag/clamp, keyboard/reset/cancel, small window/history, themes, unchanged draft/versions, reopen/restart and invalid preference fallback. ${userData}`,
  )
} catch (error) {
  const page = app?.windows()[0]
  if (page)
    await page.screenshot({ path: join(results, 'editor-split-failure.png') }).catch(() => {})
  throw error
} finally {
  if (app) await app.close()
}
