import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile, readFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { sheetFixture } from '../tests/fixtures/office.ts'
import { tabStopWordFixture } from '../tests/fixtures/word-editing.ts'
const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'word-edit-fix-')),
  userData = join(root, 'data')
const word = Buffer.from(tabStopWordFixture()),
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
  const page = await launch()
  await choose(join(root, '段落测试.docx'))
  await page.evaluate(() => window.localDocs.importFiles(null))
  await page.evaluate(() => window.localDocs.setTheme('light'))
  await page.reload()
  await page.getByRole('row', { name: /段落测试.docx/ }).dblclick()
  await page
    .locator('.office-edit-toolbar')
    .getByRole('button', { name: '编辑', exact: true })
    .click()
  const title = page.getByLabel('编辑 正文 · 段落 1', { exact: true })
  const blank = page.getByLabel('编辑 正文 · 段落 7', { exact: true })
  await expect(title).toBeFocused()
  await expect(title).toHaveAttribute('contenteditable', 'plaintext-only')
  await title.click()
  await page.keyboard.press('Control+A')
  await page.keyboard.insertText('标题已可正常输入')
  await blank.click()
  await page.keyboard.insertText('空白单元格也能填写')
  await expect(blank).toHaveText('空白单元格也能填写')
  await page.clock.install()
  await page.clock.pauseAt(new Date())
  const save = async () => {
    await page.getByRole('button', { name: '保存', exact: true }).click()
    await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
    await expect(title).toHaveText('标题已可正常输入')
  }
  await save()
  const toast = page.locator('.message')
  await expect(toast).toContainText('已保存到本地文档库')
  await page.clock.runFor(2000)
  await expect(toast).toBeVisible()
  // Same wording on a consecutive save restarts the timer.
  await blank.fill('第二次保存')
  await save()
  await page.clock.runFor(1500)
  await expect(toast).toBeVisible()
  await page.clock.runFor(1600)
  await expect(toast).toHaveCount(0)
  await expect(title).toHaveText('标题已可正常输入')
  await expect(blank).toHaveText('第二次保存')
  await page.screenshot({ path: join(results, 'word-edit-fix-no-toast.png') })
  // A manually dismissed notice leaves no timer that can erase the next one.
  await blank.fill('手动关闭提示')
  await save()
  await page.getByRole('button', { name: '关闭提示', exact: true }).click()
  await expect(toast).toHaveCount(0)
  await blank.fill('保留的最终备注')
  await save()
  await expect(toast).toContainText('已保存到本地文档库')
  await page.clock.runFor(3100)
  await expect(toast).toHaveCount(0)
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  const snapshot = await page.evaluate(() => window.localDocs.snapshot())
  const record = snapshot.documents.find((d) => d.name === '段落测试.docx')
  await app.evaluate(
    ({ dialog }, path) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
    },
    join(root, '修复后.docx'),
  )
  expect(await page.evaluate((id) => window.localDocs.exportDocument(id), record.id)).toBe(true)
  expect(await readFile(join(root, '段落测试.docx'))).toEqual(word)
  // Progress must stay visible while a native operation is still pending.
  await page.getByRole('row', { name: /段落测试.docx/ }).dblclick()
  await app.evaluate(({ dialog }) => {
    dialog.showSaveDialog = () =>
      new Promise((resolve) => {
        globalThis.__finishSaveDialog = resolve
      })
  })
  await page.getByRole('button', { name: '导出', exact: true }).click()
  await expect(toast).toContainText('正在处理…')
  await page.clock.runFor(6000)
  await expect(toast).toContainText('正在处理…')
  await app.evaluate(() => globalThis.__finishSaveDialog({ canceled: true }))
  await expect(toast).toHaveCount(0)
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  // Validation errors remain visible until explicitly dismissed.
  await page.getByRole('button', { name: '列表选项', exact: true }).click()
  await page.getByRole('button', { name: '新建 Word', exact: true }).click()
  await page.getByLabel('文件名称', { exact: true }).fill('/')
  await page.getByRole('button', { name: '创建', exact: true }).click()
  await expect(page.locator('.message.error')).toBeVisible()
  await page.clock.runFor(6000)
  await expect(page.locator('.message.error')).toBeVisible()
  await page.getByRole('button', { name: '关闭提示', exact: true }).click()
  await expect(page.locator('.message.error')).toHaveCount(0)
  expect(errors).toEqual([])
  expect(network).toEqual([])
  console.log(
    `Word tab-stop title and empty cells, real keyboard editing, save/reopen, repeated toast reset, manual close and persistent errors passed: ${root}`,
  )
} catch (error) {
  const page = app?.windows()[0]
  if (page)
    await page.screenshot({ path: join(results, 'word-edit-fix-failure.png') }).catch(() => {})
  throw error
} finally {
  if (app) await app.close()
}
