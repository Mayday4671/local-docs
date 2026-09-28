import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { pdfFixture } from '../tests/fixtures/pdf.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'context-menu-'))
const source = join(root, '阅读.pdf'),
  original = Buffer.from(pdfFixture())
await writeFile(source, original)
const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: join(root, 'data') }
delete env.ELECTRON_RUN_AS_NODE
const executablePath = process.env.LOCAL_DOCS_TEST_EXECUTABLE
const app = await electron.launch({ executablePath, args: executablePath ? [] : ['.'], env })
const page = await app.firstWindow()
page.setDefaultTimeout(15000)
await page.emulateMedia({ colorScheme: null })
const errors = []
page.on('pageerror', (error) => errors.push(String(error)))
const menu = page.getByRole('menu', { name: '文件操作', exact: true })
const row = (name) => page.getByRole('row').filter({ has: page.getByText(name, { exact: true }) })
const card = (name) =>
  page.locator('.file-card').filter({ has: page.getByText(name, { exact: true }) })
const check = (name) => page.getByRole('checkbox', { name: `选择 ${name}`, exact: true })
const action = async (name) => {
  await menu.getByRole('menuitem', { name, exact: true }).click()
  await expect(menu).toHaveCount(0)
}
const nav = async (name) => {
  await page
    .getByRole('navigation', { name: '文档导航' })
    .getByRole('button', { name: new RegExp(name) })
    .click()
  await expect(page.getByRole('heading', { name, exact: true })).toBeVisible()
}
const snapshot = () => page.evaluate(() => window.localDocs.snapshot())
try {
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  await app.evaluate(({ BrowserWindow, dialog }, source) => {
    const win = BrowserWindow.getAllWindows()[0]
    win.setContentSize(1280, 800)
    win.showInactive()
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [source] })
  }, source)
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  const seed = await page.evaluate(async () => {
    const api = window.localDocs
    const category = await api.createCategory('归档')
    const docs = []
    for (const name of ['甲', '乙', '丙']) {
      const doc = await api.createMarkdown(name, category.id)
      docs.push(await api.saveMarkdown(doc.id, `# ${name}的内容`, doc.revision))
    }
    await api.importFiles(null)
    await api.setTheme('dark')
    return { docs, category }
  })
  await page.reload()
  await expect
    .poll(() => page.evaluate(() => matchMedia('(prefers-color-scheme: dark)').matches))
    .toBe(true)
  await row('甲.md').click()
  await row('乙.md').click({ button: 'right' })
  await expect(menu).toContainText('乙.md')
  await expect(page.locator('.preview-info h2')).toHaveText('乙.md')
  await action('重命名')
  await page.getByRole('dialog').getByRole('textbox').fill('乙改名.md')
  await page.getByRole('button', { name: '保存名称', exact: true }).click()
  await expect(row('乙改名.md')).toBeVisible()
  await expect(row('甲.md')).toBeVisible()
  await row('乙改名.md').click({ button: 'right' })
  await action('编辑')
  await expect(page.locator('.editor-page')).toBeVisible()
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await row('乙改名.md').click({ button: 'right' })
  await action('收藏')
  await expect
    .poll(
      async () => (await snapshot()).documents.find((doc) => doc.id === seed.docs[1].id).favorite,
    )
    .toBe(true)
  await row('乙改名.md').click({ button: 'right' })
  await action('复制路径')
  const copied = await app.evaluate(({ clipboard }) => clipboard.readText())
  expect(copied.startsWith(root)).toBe(true)
  expect(await readFile(copied, 'utf8')).toBe('# 乙的内容')
  const exported = join(root, '导出乙.md')
  await app.evaluate(({ dialog }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath })
  }, exported)
  await row('乙改名.md').click({ button: 'right' })
  await action('导出')
  await expect
    .poll(async () => {
      try {
        return await readFile(exported, 'utf8')
      } catch {
        return ''
      }
    })
    .toBe('# 乙的内容')
  await row('阅读.pdf').click({ button: 'right' })
  await expect(menu.getByRole('menuitem', { name: '编辑', exact: true })).toHaveCount(0)
  await action('打开')
  await expect(page.locator('.editor-page .pdf-reader canvas')).toBeVisible()
  await page.getByRole('button', { name: '返回列表', exact: true }).click()

  // Native keyboard invocation, roving focus, Escape return, outside click and edge placement.
  await row('甲.md').focus()
  await page.keyboard.press('Shift+F10')
  await expect(menu).toBeVisible()
  await page.keyboard.press('End')
  await expect(menu.getByRole('menuitem', { name: '移入回收站' })).toBeFocused()
  await page.keyboard.press('Home')
  await expect(menu.getByRole('menuitem', { name: '打开', exact: true })).toBeFocused()
  await page.keyboard.press('Escape')
  await expect(row('甲.md')).toBeFocused()
  await row('甲.md').dispatchEvent('contextmenu', { clientX: 1278, clientY: 798, button: 2 })
  await expect(menu).toBeVisible()
  const bounds = await menu.boundingBox()
  expect(bounds.x).toBeGreaterThanOrEqual(8)
  expect(bounds.y).toBeGreaterThanOrEqual(8)
  expect(bounds.x + bounds.width).toBeLessThanOrEqual(1273)
  expect(bounds.y + bounds.height).toBeLessThanOrEqual(793)
  await page.screenshot({ path: join(results, 'context-menu-dark.png') })
  await page.getByRole('heading', { name: '全部文件', exact: true }).click()
  await expect(menu).toHaveCount(0)

  await check('甲.md').check()
  await check('乙改名.md').check()
  await row('甲.md').click({ button: 'right' })
  await expect(menu).toContainText('已选择 2 份文件')
  await expect(menu.getByRole('menuitem', { name: '重命名' })).toHaveCount(0)
  await action('移动到分类')
  await page.getByRole('combobox', { name: '目标分类', exact: true }).selectOption('')
  await page.getByRole('button', { name: '确认移动', exact: true }).click()
  await expect(page.getByRole('dialog')).toHaveCount(0)
  for (const id of seed.docs.slice(0, 2).map((doc) => doc.id))
    expect((await snapshot()).documents.find((doc) => doc.id === id).categoryId).toBeNull()

  await check('甲.md').check()
  await check('乙改名.md').check()
  await row('丙.md').click({ button: 'right' })
  await expect(menu).toContainText('丙.md')
  await expect(check('甲.md')).not.toBeChecked()
  await action('移入回收站')
  await expect(row('丙.md')).toHaveCount(0)
  expect((await snapshot()).documents.filter((doc) => doc.deletedAt).map((doc) => doc.id)).toEqual([
    seed.docs[2].id,
  ])
  await nav('回收站')
  await row('丙.md').click({ button: 'right' })
  await expect(menu.getByRole('menuitem', { name: '移入回收站' })).toHaveCount(0)
  await expect(menu.getByRole('menuitem', { name: '编辑', exact: true })).toHaveCount(0)
  await action('恢复文档')
  await expect(row('丙.md')).toHaveCount(0)
  await nav('全部文件')
  await check('甲.md').check()
  await check('乙改名.md').check()
  await row('甲.md').click({ button: 'right' })
  await action('移入回收站')
  await expect(row('甲.md')).toHaveCount(0)
  await nav('回收站')
  await check('甲.md').check()
  await check('乙改名.md').check()
  await row('乙改名.md').click({ button: 'right' })
  await action('恢复所选文件')
  await expect(page.locator('tbody tr')).toHaveCount(0)
  await nav('全部文件')

  await page.getByRole('button', { name: '网格视图', exact: true }).click()
  await check('甲.md').check()
  await check('乙改名.md').check()
  await card('甲.md').click({ button: 'right' })
  await expect(menu).toContainText('已选择 2 份文件')
  await action('收藏')
  await expect
    .poll(async () => (await snapshot()).documents.filter((doc) => doc.favorite).length)
    .toBe(2)
  await card('丙.md').click({ button: 'right' })
  await action('移入回收站')
  await expect(card('丙.md')).toHaveCount(0)
  await nav('回收站')
  await page.getByRole('button', { name: '切换主题', exact: true }).click()
  await page.getByRole('menuitemradio', { name: '浅色', exact: true }).click()
  await card('丙.md').click({ button: 'right' })
  await page.screenshot({ path: join(results, 'context-menu-trash-light.png') })
  await action('彻底删除')
  const confirm = page.getByRole('dialog', { name: '彻底删除文件', exact: true })
  await expect(confirm).toContainText('丙.md')
  await expect(confirm).toContainText('无法从回收站恢复')
  await confirm.getByRole('button', { name: '取消', exact: true }).click()
  expect((await snapshot()).documents.some((doc) => doc.id === seed.docs[2].id)).toBe(true)
  await card('丙.md').click({ button: 'right' })
  await action('彻底删除')
  await page.screenshot({ path: join(results, 'context-menu-purge-confirm.png') })
  await confirm.getByRole('button', { name: '确认彻底删除', exact: true }).click()
  await expect(confirm).toHaveCount(0)
  await expect(card('丙.md')).toHaveCount(0)
  await page.reload()
  expect((await snapshot()).documents.map((doc) => doc.id)).not.toContain(seed.docs[2].id)
  expect((await snapshot()).documents).toHaveLength(3)
  await page.getByRole('button', { name: '网格视图', exact: true }).click()
  await check('甲.md').check()
  await check('乙改名.md').check()
  await card('甲.md').click({ button: 'right' })
  await action('移入回收站')
  await expect(card('甲.md')).toHaveCount(0)
  await nav('回收站')
  await check('甲.md').check()
  await check('乙改名.md').check()
  await card('甲.md').click({ button: 'right' })
  await action('彻底删除')
  await expect(confirm).toContainText('2 份文件')
  await expect(confirm).toContainText('甲.md')
  await expect(confirm).toContainText('乙改名.md')
  await confirm.getByRole('button', { name: '确认彻底删除', exact: true }).click()
  await expect(confirm).toHaveCount(0)
  await expect(page.locator('.file-card')).toHaveCount(0)
  expect((await snapshot()).documents.map((doc) => doc.name)).toEqual(['阅读.pdf'])
  expect(await readFile(source)).toEqual(original)
  expect(errors).toEqual([])
  console.log(
    `Context menu passed: list/grid, exact target, keyboard/edges/dismissal, edit/read/rename/export/copy, batch move/favorite/trash/restore, purge confirmation/cancel/persistence and source protection. ${root}`,
  )
} catch (error) {
  await page.screenshot({ path: join(results, 'context-menu-failure.png') }).catch(() => {})
  throw error
} finally {
  await app.close()
}
