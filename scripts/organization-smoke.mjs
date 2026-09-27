import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile, readFile } from 'node:fs/promises'
import { resolve, join } from 'node:path'
import { wordFixture, sheetFixture } from '../tests/fixtures/office.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const directory = await mkdtemp(join(results, 'organization-desktop-'))
const userData = join(directory, 'data'),
  source = join(directory, '项目资料')
await mkdir(join(source, '开发', '空目录'), { recursive: true })
const originalWord = Buffer.from(wordFixture()),
  originalSheet = Buffer.from(sheetFixture())
await writeFile(join(source, '设计.docx'), originalWord)
await writeFile(join(source, '开发', '清单.xlsx'), originalSheet)
await writeFile(
  join(source, '开发', '记录.md'),
  '| 项目 | 状态 |\n| --- | --- |\n| 文件夹导入 | 完成 |',
)
await writeFile(join(source, '忽略.txt'), '不支持的合成测试资料')
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
  })
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  const page = await app.firstWindow()
  page.setDefaultTimeout(20_000)
  page.on('pageerror', (e) => errors.push(String(e)))
  page.on('request', (r) => {
    if (/^https?:/.test(r.url())) network.push(r.url())
  })
  await app.evaluate(({ BrowserWindow, session }) => {
    BrowserWindow.getAllWindows()[0].showInactive()
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ['http://*/*', 'https://*/*'] },
      (_, callback) => callback({ cancel: true }),
    )
  })
  await page.emulateMedia({ colorScheme: null })
  await expect(page.getByRole('heading', { name: '全部文件' })).toBeVisible()
  return page
}
async function chooseFolder(page, folder = source) {
  await app.evaluate(({ dialog }, folder) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [folder] })
  }, folder)
  await page.getByRole('button', { name: '列表选项', exact: true }).click()
  await page.getByRole('button', { name: '导入文件夹', exact: true }).click()
  await expect(page.getByRole('dialog', { name: '导入文件夹', exact: true })).toBeVisible()
}
try {
  let page = await launch()
  const seed = await page.evaluate(async () => {
    const api = window.localDocs
    const parent = await api.createCategory('待整理'),
      child = await api.createCategory('旧资料', parent.id)
    const archive = await api.createCategory('归档')
    const first = await api.createMarkdown('项目笔记', child.id),
      second = await api.createMarkdown('待归档', null)
    await api.saveMarkdown(first.id, '# 已保存的正文', 1)
    await api.updateDocument(first.id, {
      tags: ['保留'],
      notes: '分类变化不丢失内容',
      favorite: true,
    })
    const trash = await api.createMarkdown('旧版笔记', parent.id)
    await api.trashDocument(trash.id)
    await api.setTheme('light')
    return { parent, child, archive, first, second, trash }
  })
  const original = await page.evaluate((id) => window.localDocs.readDocument(id), seed.first.id)
  const versions = await page.evaluate((id) => window.localDocs.versions(id), seed.first.id)
  await page.reload()
  await page.getByRole('button', { name: '管理分类 待整理', exact: true }).click()
  await expect(page.getByLabel('移动到上级分类').locator('option')).toHaveCount(2)
  await page.getByLabel('分类名称', { exact: true }).fill('项目归档')
  await page.getByLabel('移动到上级分类').selectOption(seed.archive.id)
  await page.screenshot({ path: join(results, 'organization-category-light.png') })
  await page.getByRole('button', { name: '保存分类', exact: true }).click()
  await expect(page.getByRole('dialog', { name: '管理分类', exact: true })).toHaveCount(0)
  expect(await page.evaluate((id) => window.localDocs.readDocument(id), seed.first.id)).toEqual(
    original,
  )
  await page.getByRole('button', { name: '管理分类 项目归档', exact: true }).click()
  await page.getByRole('button', { name: '删除分类…', exact: true }).click()
  await expect(page.getByRole('dialog')).toContainText('另有 1 份回收站文件')
  await expect(page.getByLabel('文件转移到').locator('option')).toHaveCount(2)
  await page.getByRole('button', { name: '删除分类，保留文件', exact: true }).click()
  await expect(page.getByRole('dialog')).toHaveCount(0)
  const afterDelete = await page.evaluate(() => window.localDocs.snapshot())
  expect(afterDelete.documents.find((d) => d.id === seed.trash.id)).toMatchObject({
    categoryId: null,
  })
  expect(afterDelete.documents.find((d) => d.id === seed.trash.id).deletedAt).toBeTruthy()
  expect(await page.evaluate((id) => window.localDocs.versions(id), seed.first.id)).toEqual(
    versions,
  )
  await page.getByRole('checkbox', { name: '选择全部文件', exact: true }).check()
  await page.getByRole('button', { name: '移动到分类', exact: true }).click()
  await page.getByLabel('目标分类').selectOption(seed.archive.id)
  await page.getByRole('button', { name: '确认移动', exact: true }).click()
  await expect(page.getByRole('heading', { name: '归档', exact: true })).toBeVisible()
  await expect(page.locator('.selection-toolbar')).toHaveCount(0)
  expect(
    (await page.evaluate(() => window.localDocs.snapshot())).documents
      .filter((d) => !d.deletedAt)
      .every((d) => d.categoryId === seed.archive.id),
  ).toBe(true)

  await page.getByRole('button', { name: '切换主题', exact: true }).click()
  await page.getByRole('menuitemradio', { name: '深色', exact: true }).click()
  await chooseFolder(page)
  const before = await page.evaluate(() => window.localDocs.snapshot())
  await page.getByRole('button', { name: '取消', exact: true }).click()
  await expect(page.getByRole('dialog')).toHaveCount(0)
  expect(await page.evaluate(() => window.localDocs.snapshot())).toEqual(before)
  await chooseFolder(page)
  await expect(page.locator('.folder-stats')).toContainText('3 份文件')
  await page.getByText('查看目录预览（最多显示 100 项）', { exact: true }).click()
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1000, 650))
  const start = page.getByRole('button', { name: '开始导入', exact: true })
  await start.scrollIntoViewIfNeeded()
  expect((await start.boundingBox()).y).toBeLessThan(650)
  await page.screenshot({ path: join(results, 'organization-import-compact-dark.png') })
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1536, 1024))
  await page.screenshot({ path: join(results, 'organization-import-dark.png') })
  await start.click()
  await expect(page.getByRole('heading', { name: '文件夹导入完成', exact: true })).toBeVisible()
  await expect(page.locator('.folder-stats')).toContainText('3 份已导入')
  await expect(page.locator('.folder-stats')).toContainText('0 项失败')
  await page
    .getByRole('dialog', { name: '文件夹导入完成', exact: true })
    .getByRole('button', { name: '查看本批文件', exact: true })
    .click()
  await expect(page.locator('.file-table tbody tr')).toHaveCount(3)
  await page.getByRole('row', { name: /记录.md/ }).click()
  await expect(
    page.getByRole('region', { name: 'Markdown 表格', exact: true }).getByRole('table'),
  ).toBeVisible()
  await page.screenshot({ path: join(results, 'organization-import-result-dark.png') })
  await page.getByRole('row', { name: /设计.docx/ }).click()
  await expect(page.locator('.preview-panel .word-layout')).toBeVisible()
  await page.getByRole('row', { name: /清单.xlsx/ }).click()
  await expect(page.locator('.preview-panel .sheet-grid')).toBeVisible()
  await expect(readFile(join(source, '设计.docx'))).resolves.toEqual(originalWord)
  await expect(readFile(join(source, '开发', '清单.xlsx'))).resolves.toEqual(originalSheet)
  // Reuse the same destination; importing from "all" defaults to root, so choose the archive explicitly.
  await chooseFolder(page)
  await page.getByLabel('导入到', { exact: true }).selectOption(seed.archive.id)
  await page.getByRole('button', { name: '开始导入', exact: true }).click()
  await expect(page.getByRole('heading', { name: '文件夹导入完成', exact: true })).toBeVisible()
  await expect(page.locator('.folder-stats')).toContainText('0 份已导入')
  await expect(page.locator('.folder-stats')).toContainText('4 项跳过')
  await page.getByRole('button', { name: '关闭', exact: true }).click()
  await expect(page.locator('.file-table tbody tr')).toHaveCount(5)
  await chooseFolder(page)
  await page.getByLabel('导入到', { exact: true }).selectOption(seed.archive.id)
  await writeFile(join(source, '开发', '记录.md'), '预览后修改的内容')
  await page.getByRole('button', { name: '开始导入', exact: true }).click()
  await expect(page.getByRole('heading', { name: '文件夹导入完成', exact: true })).toBeVisible()
  await expect(page.locator('.folder-stats')).toContainText('1 项失败')
  await expect(page.locator('.folder-details')).toContainText('预览后已修改')
  await page.getByRole('button', { name: '关闭', exact: true }).click()

  // Exercise the real progress/cancel control with a bounded synthetic batch.
  const bulk = join(directory, '取消导入样本')
  await mkdir(bulk)
  for (let i = 0; i < 500; i++)
    await writeFile(join(bulk, `测试${i}.md`), `测试 ${i}\n${'资料内容\n'.repeat(4000)}`)
  await chooseFolder(page, bulk)
  await page.getByRole('button', { name: '开始导入', exact: true }).click()
  await page.getByRole('button', { name: '取消操作', exact: true }).click()
  await expect(page.getByRole('heading', { name: '导入已取消', exact: true })).toBeVisible()
  await page.getByRole('button', { name: '关闭', exact: true }).click()
  const saved = await page.evaluate(() => window.localDocs.snapshot())
  const savedHistory = await page.evaluate((id) => window.localDocs.versions(id), seed.first.id)
  await app.close()
  page = await launch()
  expect(await page.evaluate(() => window.localDocs.snapshot())).toEqual(saved)
  expect(await page.evaluate((id) => window.localDocs.versions(id), seed.first.id)).toEqual(
    savedHistory,
  )
  expect(await page.evaluate(() => window.localDocs.getTheme())).toBe('dark')
  expect(errors).toEqual([])
  expect(network).toEqual([])
  console.log(
    `Category management, batch move, folder import/skip/failure/cancel, previews and restart passed: ${directory}`,
  )
} finally {
  if (app) await app.close()
}
