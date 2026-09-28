import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile, readFile, readdir } from 'node:fs/promises'
import { resolve, join } from 'node:path'
import { wordFixture, sheetFixture } from '../tests/fixtures/office.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const directory = await mkdtemp(join(results, 'transfer-desktop-'))
const userData = join(directory, 'data')
const output = join(directory, 'output')
await mkdir(output)
const backup = join(output, '完整备份.localdocs-backup')
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
    cwd: executablePath ? directory : undefined,
    env,
    timeout: 30_000,
  })
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  const page = await app.firstWindow()
  await app.evaluate(({ BrowserWindow, session }) => {
    BrowserWindow.getAllWindows()[0].showInactive()
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ['http://*/*', 'https://*/*'] },
      (_details, callback) => callback({ cancel: true }),
    )
  })
  page.setDefaultTimeout(20_000)
  page.on('pageerror', (error) => errors.push(String(error)))
  page.on('request', (request) => {
    if (/^https?:/.test(request.url())) network.push(request.url())
  })
  await page.emulateMedia({ colorScheme: null })
  await expect(page.getByRole('heading', { name: '全部文件' })).toBeVisible()
  return page
}
async function chooseOpen(paths) {
  await app.evaluate(({ dialog }, paths) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
  }, paths)
}
async function chooseSave(path) {
  await app.evaluate(({ dialog }, path) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
  }, path)
}
try {
  let page = await launch()
  const seed = await page.evaluate(async () => {
    const api = window.localDocs
    const work = await api.createCategory('工作资料')
    const design = await api.createCategory('项目设计', work.id)
    await api.createCategory('空分类', design.id)
    let doc = await api.createMarkdown('项目记录', design.id)
    doc = await api.saveMarkdown(doc.id, '# 初稿\n\n最早的内容', doc.revision)
    doc = await api.saveMarkdown(
      doc.id,
      '# 项目记录\n\n| 项目 | 状态 |\n| --- | --- |\n| 本地备份 | 已验收 |',
      doc.revision,
    )
    await api.updateDocument(doc.id, {
      favorite: true,
      tags: ['离线', '验收'],
      notes: '完整保留分类和历史',
    })
    await api.openDocument(doc.id)
    const trash = await api.createMarkdown('暂时不用的资料', null)
    await api.trashDocument(trash.id)
    await api.setTheme('dark')
    return { work, design, doc, trash }
  })
  const wordPath = join(directory, '说明.docx'),
    sheetPath = join(directory, '数据.xlsx')
  const originalWord = Buffer.from(wordFixture())
  const originalSheet = Buffer.from(sheetFixture())
  await writeFile(wordPath, originalWord)
  await writeFile(sheetPath, originalSheet)
  await chooseOpen([wordPath, sheetPath])
  await page.evaluate((id) => window.localDocs.importFiles(id), seed.design.id)
  await page.reload()
  await expect(page.getByRole('row', { name: /项目记录.md/ })).toBeVisible()
  const before = await page.evaluate(() => window.localDocs.snapshot())
  const history = await page.evaluate((id) => window.localDocs.versions(id), seed.doc.id)
  await page.getByRole('button', { name: '设置', exact: true }).click()
  await chooseSave(backup)
  await page.getByRole('button', { name: '创建备份', exact: true }).click()
  await expect(page.getByText('完整备份已保存', { exact: true })).toBeVisible()
  expect((await readFile(backup)).length).toBeGreaterThan(1000)
  await page.screenshot({ path: join(results, 'transfer-settings-dark.png') })

  await chooseOpen([backup])
  await page.getByRole('button', { name: '选择备份', exact: true }).click()
  await expect(page.getByRole('heading', { name: '确认恢复文档库' })).toBeVisible()
  await expect(page.locator('.restore-summary')).toContainText('3 / 1 份')
  await page.getByRole('button', { name: '取消', exact: true }).click()
  await expect(page.getByRole('heading', { name: '确认恢复文档库' })).toHaveCount(0)
  await expect(page.locator('.transfer-backdrop')).toHaveCount(0)
  expect(await page.evaluate(() => window.localDocs.snapshot())).toEqual(before)
  await page.getByRole('button', { name: '关闭对话框', exact: true }).click()

  await page.evaluate(async (id) => {
    await window.localDocs.createMarkdown('备份后新增', null)
    await window.localDocs.trashDocument(id)
    await window.localDocs.setTheme('light')
  }, seed.doc.id)
  await page.reload()
  await page.getByRole('button', { name: '设置', exact: true }).click()
  await page.getByRole('button', { name: '选择备份', exact: true }).click()
  await expect(page.getByRole('heading', { name: '确认恢复文档库' })).toBeVisible()
  await page.screenshot({ path: join(results, 'transfer-restore-preview.png') })
  await page.getByRole('button', { name: '备份当前资料并恢复', exact: true }).click()
  await expect(page.getByText('恢复完成 · 已保留恢复前的备份', { exact: true })).toBeVisible()
  expect(await page.evaluate(() => window.localDocs.snapshot())).toEqual(before)
  expect(await page.evaluate((id) => window.localDocs.versions(id), seed.doc.id)).toEqual(history)
  expect(await page.evaluate(() => window.localDocs.getTheme())).toBe('dark')
  const recovery = await readdir(join(before.storagePath, 'recovery-backups'))
  expect(recovery).toHaveLength(1)
  // Rebuilding derived Office data is exercised after the atomic restore.
  expect((await page.evaluate(() => window.localDocs.searchResults('端口'))).length).toBe(2)
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1000, 650))
  await expect(page.getByRole('button', { name: '导出整个文档库', exact: true })).toBeVisible()
  await page.screenshot({ path: join(results, 'transfer-settings-compact.png') })
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1536, 1024))
  await page.getByRole('button', { name: '关闭对话框', exact: true }).click()

  await page
    .locator('.category-list')
    .getByRole('button', { name: '工作资料', exact: true })
    .click()
  await page.getByRole('button', { name: '列表选项', exact: true }).click()
  await chooseOpen([output])
  await page.getByRole('button', { name: '导出此分类（含子分类）', exact: true }).click()
  await expect(page.getByText('已按分类导出 3 份文件', { exact: true })).toBeVisible()
  const exportFolder = (await readdir(output)).find((name) => name.startsWith('工作资料-'))
  expect(await readFile(join(output, exportFolder, '工作资料', '项目设计', '说明.docx'))).toEqual(
    originalWord,
  )
  expect(await readFile(join(output, exportFolder, '工作资料', '项目设计', '数据.xlsx'))).toEqual(
    originalSheet,
  )
  expect(await readFile(wordPath)).toEqual(originalWord)
  expect(await readFile(sheetPath)).toEqual(originalSheet)
  const exported = JSON.parse(await readFile(join(output, exportFolder, '导出清单.json'), 'utf8'))
  expect(exported.files).toHaveLength(3)
  expect(exported.files.find((f) => f.name === '项目记录.md').tags).toEqual(['离线', '验收'])
  await page.screenshot({ path: join(results, 'transfer-export-result.png') })

  const damaged = join(output, '损坏备份.localdocs-backup')
  await writeFile(damaged, 'invalid archive')
  await chooseOpen([damaged])
  await page.getByRole('button', { name: '选择备份', exact: true }).click()
  await expect(page.getByRole('alert')).toBeVisible()
  expect(await page.evaluate(() => window.localDocs.snapshot())).toEqual(before)
  await app.close()
  page = await launch()
  expect(await page.evaluate(() => window.localDocs.snapshot())).toEqual(before)
  expect(await page.evaluate(() => window.localDocs.getTheme())).toBe('dark')
  expect(errors).toEqual([])
  expect(network).toEqual([])
  console.log(
    `Backup, preview/cancel, restore, safety backup, category export, failure and restart checks passed: ${directory}`,
  )
} finally {
  if (app) await app.close()
}
