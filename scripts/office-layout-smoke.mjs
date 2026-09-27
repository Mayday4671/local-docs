import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile, readFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { wordFixture, sheetFixture } from '../tests/fixtures/office.ts'
const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'office-layout-')),
  userData = join(root, 'data')
const word = Buffer.from(wordFixture()),
  sheet = Buffer.from(sheetFixture())
await writeFile(join(root, '段落测试.docx'), word)
await writeFile(join(root, '单元格测试.xlsx'), sheet)
const errors = [],
  network = []
let app
async function launch() {
  const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: userData }
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
  page.on('request', (r) => {
    if (/^https?:/.test(r.url())) network.push(r.url())
  })
  await app.evaluate(({ BrowserWindow, session }) => {
    BrowserWindow.getAllWindows()[0].showInactive()
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ['http://*/*', 'https://*/*'] },
      (_, cb) => cb({ cancel: true }),
    )
  })
  await page.emulateMedia({ colorScheme: null })
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  return page
}
async function choose(path) {
  await app.evaluate(({ dialog }, path) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [path] })
  }, path)
}

try {
  let page = await launch()
  await page.evaluate(() => window.localDocs.setTheme('light'))
  await choose(join(root, '段落测试.docx'))
  await page.evaluate(() => window.localDocs.importFiles(null))
  await choose(join(root, '单元格测试.xlsx'))
  await page.evaluate(() => window.localDocs.importFiles(null))
  await page.reload()
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1500, 950))
  await page.getByRole('row', { name: /段落测试.docx/ }).dblclick()
  await page
    .locator('.office-edit-toolbar')
    .getByRole('button', { name: '编辑', exact: true })
    .click()
  const layout = page.getByLabel('Word 页面编辑', { exact: true })
  const title = page.getByLabel('编辑 正文 · 段落 1', { exact: true })
  await expect(title).toBeVisible()
  expect(await title.evaluate((e) => e.tagName)).toBe('P')
  await expect(layout.locator('table')).toHaveCount(1)
  await expect(layout.locator('img')).toHaveCount(1)
  await expect(layout.locator('header')).toContainText('网络运维资料')
  expect(
    await layout.locator('p').filter({ hasText: '脚本示例' }).getAttribute('contenteditable'),
  ).toBeNull()
  const tableCell = page.getByLabel('编辑 正文 · 段落 5', { exact: true })
  expect(await tableCell.evaluate((e) => !!e.closest('td'))).toBe(true)
  await tableCell.fill('表格内直接编辑')
  await title.fill('页面中的中文标题')
  await title.press('Enter')
  expect(await title.textContent()).toBe('页面中的中文标题')
  await page.getByLabel('查找编辑位置', { exact: true }).fill('表格内')
  await page.getByRole('button', { name: '查找', exact: true }).click()
  await expect(tableCell).toHaveAttribute('data-found', '')
  const mixed = page.getByLabel('编辑 正文 · 段落 2', { exact: true })
  await mixed.fill('请检查端口配置与连接方式。并确认状态。')
  await title.click() // blur restores unchanged rich-text runs
  expect(
    await mixed
      .locator('span')
      .evaluateAll((spans) =>
        spans.some((s) => s.textContent === '端口' && getComputedStyle(s).fontWeight === '700'),
      ),
  ).toBe(true)
  await page.getByLabel('文档缩放', { exact: true }).selectOption('90')
  await expect(page.getByText('恢复草稿已自动保存', { exact: false }).first()).toBeVisible()
  await page.screenshot({ path: join(results, 'office-layout-word-light.png') })
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
  await expect(title).toHaveText('页面中的中文标题')
  await expect(tableCell).toHaveText('表格内直接编辑')
  // An empty paragraph remains editable after saving and reopening.
  await title.fill('')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
  await title.fill('重新输入标题')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
  await expect(title).toHaveText('重新输入标题')
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('row', { name: /单元格测试.xlsx/ }).dblclick()
  await page
    .locator('.office-edit-toolbar')
    .getByRole('button', { name: '编辑', exact: true })
    .click()
  const cell = (address) => page.getByLabel(`编辑 设备清单 · ${address}`, { exact: true })
  await expect(cell('A1')).toBeVisible()
  expect(await cell('A1').evaluate((e) => e.closest('td').colSpan)).toBe(2)
  await expect(cell('B1')).toHaveCount(0)
  await cell('A2').fill('网格中的新名称')
  await cell('A2').press('Tab')
  await expect(cell('B2')).toBeFocused()
  await page.keyboard.insertText('新空白格')
  await cell('B2').press('Enter')
  await expect(cell('B3')).toBeFocused()
  await page.getByLabel('单元格内容', { exact: true }).fill('输入栏创建的值')
  await expect(cell('B3')).toHaveValue('输入栏创建的值')
  await page.getByLabel('查找编辑位置', { exact: true }).fill('AA120')
  await page.getByRole('button', { name: '定位', exact: true }).click()
  await expect(cell('AA120')).toBeFocused()
  await cell('AA120').fill('远端编辑已定位')
  await page.getByRole('tab', { name: '预算（隐藏）', exact: true }).click()
  const formula = page.getByLabel('编辑 预算 · D2', { exact: true })
  await expect(formula).toHaveAttribute('readonly', '')
  await formula.click()
  await expect(page.getByLabel('单元格内容', { exact: true })).toHaveValue('=SUM(B2:C2)')
  await expect(page.getByLabel('单元格内容', { exact: true })).toHaveAttribute('readonly', '')
  await page.getByLabel('编辑 预算 · B2', { exact: true }).fill('150')
  await page.getByRole('tab', { name: '设备清单', exact: true }).click()
  await expect(cell('B2')).toHaveValue('新空白格')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
  await expect(cell('B2')).toHaveValue('新空白格')
  await expect(cell('B3')).toHaveValue('输入栏创建的值')
  await page.screenshot({ path: join(results, 'office-layout-excel-light.png') })
  // Draft survives mode switch and an actual application restart.
  await cell('D4').fill('重启后恢复的格子')
  await expect(page.getByText('恢复草稿已自动保存', { exact: false }).first()).toBeVisible()
  await page.getByRole('button', { name: '阅读与标记', exact: true }).click()
  await expect(page.getByText(/当前阅读的是已保存版本/)).toBeVisible()
  await page.getByRole('button', { name: '返回编辑', exact: true }).click()
  await expect(cell('D4')).toHaveValue('重启后恢复的格子')
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('button', { name: '保留草稿并返回', exact: true }).click()
  await app.close()
  app = undefined
  page = await launch()
  await page.getByRole('row', { name: /单元格测试.xlsx/ }).dblclick()
  await page.getByRole('button', { name: '恢复草稿', exact: true }).click()
  await page
    .locator('.office-edit-toolbar')
    .getByRole('button', { name: '编辑', exact: true })
    .click()
  await expect(cell('D4')).toHaveValue('重启后恢复的格子')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
  await page.getByRole('button', { name: '切换主题', exact: true }).click()
  await page.getByRole('menuitemradio', { name: '深色', exact: true }).click()
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1000, 680))
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true)
  await expect(page.getByRole('button', { name: '下一组列', exact: true })).toBeVisible()
  await page.screenshot({ path: join(results, 'office-layout-excel-dark-small.png') })
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('row', { name: /段落测试.docx/ }).dblclick()
  await page
    .locator('.office-edit-toolbar')
    .getByRole('button', { name: '编辑', exact: true })
    .click()
  await expect(page.getByLabel('编辑 正文 · 段落 1', { exact: true })).toHaveText('重新输入标题')
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true)
  await page.screenshot({ path: join(results, 'office-layout-word-dark-small.png') })
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  const snapshot = await page.evaluate(() => window.localDocs.snapshot())
  for (const name of ['段落测试.docx', '单元格测试.xlsx']) {
    const record = snapshot.documents.find((d) => d.name === name)
    await app.evaluate(
      ({ dialog }, path) => {
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
      },
      join(root, `编辑后-${name}`),
    )
    expect(await page.evaluate((id) => window.localDocs.exportDocument(id), record.id)).toBe(true)
  }
  expect(await readFile(join(root, '段落测试.docx'))).toEqual(word)
  expect(await readFile(join(root, '单元格测试.xlsx'))).toEqual(sheet)
  expect(errors).toEqual([])
  expect(network).toEqual([])
  console.log(
    `Office page/grid editing, merged and empty cells, keyboard, draft recovery, export and themes passed: ${root}`,
  )
} catch (error) {
  const page = app?.windows()[0]
  if (page)
    await page.screenshot({ path: join(results, 'office-layout-failure.png') }).catch(() => {})
  throw error
} finally {
  if (app) await app.close()
}
