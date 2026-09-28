import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { wordFixture } from '../tests/fixtures/office.ts'
import { pdfFixture } from '../tests/fixtures/pdf.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'dropdown-'))
const source = join(root, 'import')
await mkdir(source)
await writeFile(join(source, '示例.txt'), 'isolated dropdown fixture')
await writeFile(join(root, '预览.docx'), wordFixture())
await writeFile(join(root, '预览.pdf'), pdfFixture())
const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: join(root, 'data') }
delete env.ELECTRON_RUN_AS_NODE
const executablePath = process.env.LOCAL_DOCS_TEST_EXECUTABLE
const app = await electron.launch({ executablePath, args: executablePath ? [] : ['.'], env })
const errors = []
const page = await app.firstWindow()
page.setDefaultTimeout(15000)
page.on('pageerror', (error) => errors.push(String(error)))
await page.emulateMedia({ colorScheme: null })
const combo = (name) => page.getByRole('combobox', { name, exact: true })

async function openPicker(name) {
  const select = combo(name)
  await select.click()
  await expect(select).toHaveCSS('appearance', 'base-select')
  await expect.poll(() => select.evaluate((el) => el.matches(':open'))).toBe(true)
  const appearance = await select.evaluate(
    (el) => getComputedStyle(el, '::picker(select)').appearance,
  )
  expect(appearance).toBe('base-select')
  const options = select.locator('option')
  await expect(options.first()).toBeVisible()
  expect(
    await options.first().evaluate((el) => el.getBoundingClientRect().height),
  ).toBeGreaterThanOrEqual(34)
  const bounds = await select.evaluate((el) => {
    const style = getComputedStyle(el, '::picker(select)')
    return { height: parseFloat(style.height), maxHeight: parseFloat(style.maxHeight) }
  })
  expect(bounds.height).toBeLessThanOrEqual(bounds.maxHeight)
  return select
}
async function dismissPicker(name) {
  await page.keyboard.press('Escape')
  await expect.poll(() => combo(name).evaluate((el) => el.matches(':open'))).toBe(false)
  await expect(combo(name)).toBeFocused()
}
async function pick(name, label) {
  const select = await openPicker(name)
  await select.getByRole('option', { name: label, exact: true }).click()
  await expect(page.locator('select:open')).toHaveCount(0)
}
async function fits() {
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true)
  for (const select of await page.getByRole('combobox').all()) {
    const box = await select.boundingBox()
    if (box)
      expect(box.x + box.width).toBeLessThanOrEqual(
        (await page.viewportSize())?.width ?? (await page.evaluate(() => innerWidth)),
      )
  }
}
try {
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  await app.evaluate(({ BrowserWindow }) => {
    const win = BrowserWindow.getAllWindows()[0]
    win.setContentSize(1280, 850)
    win.showInactive()
  })
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  expect(await page.evaluate(() => CSS.supports('appearance', 'base-select'))).toBe(true)
  const seed = await page.evaluate(async () => {
    const api = window.localDocs
    const java = await api.createCategory('Java')
    const doc = await api.createMarkdown('下拉框检查', java.id)
    await api.saveMarkdown(doc.id, '# 标记检查\n\n选择这段测试文字。', doc.revision)
    await api.setTheme('dark')
    return { java, doc }
  })
  await app.evaluate(
    ({ dialog }, paths) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
    },
    [join(root, '预览.docx'), join(root, '预览.pdf')],
  )
  await page.evaluate(() => window.localDocs.importFiles(null))
  await page.reload()

  await page.getByRole('button', { name: '新建文件夹', exact: true }).click()
  await pick('上级分类', 'Java')
  await openPicker('上级分类')
  expect((await combo('上级分类').locator('option').first().boundingBox()).y).toBeGreaterThan(
    (await combo('上级分类').boundingBox()).y,
  )
  await page.screenshot({ path: join(results, 'dropdown-category-dark-short.png') })
  await dismissPicker('上级分类')
  await combo('上级分类').press('Space')
  await page.keyboard.press('Tab')
  await expect(page.locator('select:open')).toHaveCount(0)
  await expect(page.getByRole('dialog', { name: '新建分类' })).toBeVisible()
  await page.getByRole('button', { name: '取消', exact: true }).click()
  const long = await page.evaluate(async () => {
    const api = window.localDocs
    const category = await api.createCategory('很长的分类名称用于验证显示范围与完整阅读'.repeat(4))
    for (let i = 1; i <= 35; i++) await api.createCategory(`测试分类 ${String(i).padStart(2, '0')}`)
    return category
  })
  await page.reload()

  // Same dialog as the reported screenshot; mouse, keyboard, long labels and scrolling.
  await page.getByRole('button', { name: '新建文件夹', exact: true }).click()
  await page.getByLabel('分类名称', { exact: true }).fill('键盘创建')
  let select = await openPicker('上级分类')
  await page.screenshot({ path: join(results, 'dropdown-category-dark.png') })
  await dismissPicker('上级分类')
  await combo('上级分类').press('Space')
  await expect.poll(() => combo('上级分类').evaluate((el) => el.matches(':open'))).toBe(true)
  await page.keyboard.press('End')
  await page.keyboard.press('Enter')
  await expect(combo('上级分类')).toHaveValue(
    await combo('上级分类').locator('option').last().getAttribute('value'),
  )
  await expect(page.getByRole('dialog', { name: '新建分类' })).toBeVisible()
  await pick('上级分类', long.name)
  await fits()
  await page.screenshot({ path: join(results, 'dropdown-long-selected.png') })
  await pick('上级分类', 'Java')
  await page.getByRole('button', { name: '创建', exact: true }).click()
  await expect(page.getByRole('dialog', { name: '新建分类' })).toHaveCount(0)
  expect(
    (await page.evaluate(() => window.localDocs.snapshot())).categories.find(
      (c) => c.name === '键盘创建',
    ).parentId,
  ).toBe(seed.java.id)

  await page.getByRole('button', { name: '管理分类 Java', exact: true }).click()
  await openPicker('移动到上级分类')
  await dismissPicker('移动到上级分类')
  await expect(page.getByRole('dialog', { name: '管理分类' })).toBeVisible()
  await page.getByRole('button', { name: '删除分类…', exact: true }).click()
  await openPicker('文件转移到')
  await dismissPicker('文件转移到')
  await page.getByRole('button', { name: '关闭整理对话框', exact: true }).click()

  await page
    .locator('.library-sidebar')
    .getByRole('button', { name: /全部文件/ })
    .click()
  await page.getByRole('button', { name: '列表选项', exact: true }).click()
  await pick('文件类型', 'Markdown')
  await expect(page.locator('.file-table tbody tr')).toHaveCount(1)
  await page.getByRole('button', { name: '列表选项', exact: true }).click()
  await openPicker('排序方式')
  await dismissPicker('排序方式')
  await expect(combo('文件类型')).toBeVisible()
  await pick('排序方式', '名称')
  await pick('文件类型', '全部类型')

  await page.getByRole('checkbox', { name: '选择全部文件', exact: true }).check()
  await page.getByRole('button', { name: '移动到分类', exact: true }).click()
  await openPicker('目标分类')
  await dismissPicker('目标分类')
  await page.getByRole('button', { name: '关闭整理对话框', exact: true }).click()
  await page.getByRole('checkbox', { name: '选择全部文件', exact: true }).uncheck()

  await app.evaluate(({ dialog }, folder) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [folder] })
  }, source)
  await page.getByRole('button', { name: '列表选项', exact: true }).click()
  await page.getByRole('button', { name: '导入文件夹', exact: true }).click()
  await openPicker('导入到')
  await dismissPicker('导入到')
  const duplicateName = await page
    .getByRole('dialog')
    .locator('select')
    .last()
    .getAttribute('aria-label')
  await pick(duplicateName, '全部保留为新副本')
  await page.getByRole('button', { name: '取消', exact: true }).click()
  await page.evaluate((id) => window.localDocs.trashDocument(id), seed.doc.id)
  await page.reload()
  await page.getByRole('button', { name: /回收站/ }).click()
  await page.getByRole('row', { name: /下拉框检查.md/ }).click()
  await page.getByRole('tab', { name: '信息', exact: true }).click()
  await expect(combo('所在分类')).toBeDisabled()
  await expect(page.locator('select:open')).toHaveCount(0)
  await page.getByRole('button', { name: '恢复文档', exact: true }).click()
  await expect(page.getByRole('row', { name: /下拉框检查.md/ })).toHaveCount(0)
  await page
    .locator('.library-sidebar')
    .getByRole('button', { name: /全部文件/ })
    .click()

  await page.getByRole('row', { name: /下拉框检查.md/ }).click()
  await page.getByRole('tab', { name: '信息', exact: true }).click()
  await pick('所在分类', '不指定分类')
  await expect
    .poll(
      async () =>
        (await page.evaluate((id) => window.localDocs.readDocument(id), seed.doc.id)).document
          .categoryId,
    )
    .toBe(null)
  await page.getByRole('tab', { name: '预览', exact: true }).click()
  await page.locator('.markdown-body p').evaluate((p) => {
    const range = document.createRange()
    range.selectNodeContents(p)
    const selection = window.getSelection()
    selection.removeAllRanges()
    selection.addRange(range)
    p.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }))
  })
  await page.getByRole('button', { name: '标记所选内容', exact: true }).click()
  await pick('标记颜色', '绿色')
  await page.getByRole('button', { name: '保存标记', exact: true }).click()
  await expect(page.locator('.mark-editor')).toHaveCount(0)

  await page.getByRole('button', { name: '编辑', exact: true }).click()
  await page.getByRole('button', { name: '历史版本', exact: true }).click()
  await page
    .locator('.history-item')
    .first()
    .getByRole('button', { name: /查看差异|查看内容/ })
    .click()
  for (const name of ['对比前版本', '对比后版本']) {
    await openPicker(name)
    await dismissPicker(name)
    await expect(page.getByRole('dialog', { name: '历史版本对比', exact: true })).toBeVisible()
  }
  await page.getByRole('button', { name: '关闭版本对比' }).click()
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('row', { name: /预览.docx/ }).click()
  await page.getByRole('button', { name: '编辑', exact: true }).click()
  await pick('文档缩放', '125%')
  await expect(combo('文档缩放')).toHaveValue('125')
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('row', { name: /预览.pdf/ }).click()
  await pick('PDF 缩放', '125%')

  // Light appearance and viewport-constrained menu, including outside-click dismissal.
  await page.getByRole('button', { name: '切换主题', exact: true }).click()
  await page.getByRole('menuitemradio', { name: '浅色', exact: true }).click()
  await app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()[0].setContentSize(1000, 680),
  )
  await page.getByRole('button', { name: '新建文件夹', exact: true }).click()
  select = await openPicker('上级分类')
  await fits()
  await page.screenshot({ path: join(results, 'dropdown-category-light-small.png') })
  await page.getByText('先建一个分类，之后随时可以把文件移进来。', { exact: true }).click()
  await expect.poll(() => select.evaluate((el) => el.matches(':open'))).toBe(false)
  await expect(page.getByRole('dialog', { name: '新建分类' })).toBeVisible()
  await page.getByRole('button', { name: '取消', exact: true }).click()
  expect(errors).toEqual([])
  console.log(
    `Dropdown checks passed: category dialogs, move/import, filters/sort, metadata, marks, history, Word/PDF zoom; mouse/keyboard/Escape/outside click, themes, long labels and small window. ${root}`,
  )
} catch (error) {
  await page.screenshot({ path: join(results, 'dropdown-failure.png') }).catch(() => {})
  throw error
} finally {
  await app.close()
}
