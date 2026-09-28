import { _electron as electron, expect } from '@playwright/test'
import { mkdtemp, mkdir, writeFile, readFile } from 'node:fs/promises'
import { resolve, join } from 'node:path'
import {
  markdownFixture,
  formattedSheetFixture,
  wideWordFixture,
} from '../tests/fixtures/formatting.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const userData = await mkdtemp(join(results, 'formatting-desktop-'))
const fixtures = [
  ['格式样本.MD', Buffer.from(markdownFixture)],
  ['扩展格式.markdown', Buffer.from('\ufeff' + markdownFixture)],
  ['宽表说明.DOCX', Buffer.from(wideWordFixture())],
  ['格式工作簿.XLSX', Buffer.from(formattedSheetFixture())],
  ['损坏工作簿.xlsx', Buffer.from('Not an Office archive')],
]
for (const [name, bytes] of fixtures) await writeFile(join(userData, name), bytes)
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
    cwd: executablePath ? userData : undefined,
    env,
  })
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  const page = await app.firstWindow()
  // Clear Playwright's default light emulation to observe Electron's native appearance.
  await page.emulateMedia({ colorScheme: null })
  page.setDefaultTimeout(15000)
  page.on('pageerror', (error) => errors.push(String(error)))
  page.on('request', (request) => {
    if (/^https?:/.test(request.url())) network.push(request.url())
  })
  await app.evaluate(({ BrowserWindow, session }) => {
    BrowserWindow.getAllWindows()[0].showInactive()
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ['http://*/*', 'https://*/*'] },
      (_details, callback) => callback({ cancel: true }),
    )
  })
  await expect(page.getByRole('heading', { name: '全部文件' })).toBeVisible()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  return page
}
async function setTheme(page, theme, label) {
  await page
    .getByRole('button', { name: '切换主题', exact: true })
    .filter({ visible: true })
    .click()
  const option = page.getByRole('menuitemradio', { name: label, exact: true })
  await option.click()
  await expect.poll(() => page.evaluate(() => window.localDocs.getTheme())).toBe(theme)
  expect(await app.evaluate(({ nativeTheme }) => nativeTheme.themeSource)).toBe(theme)
  if (theme !== 'system') {
    await expect
      .poll(() =>
        app.evaluate(({ BrowserWindow }) =>
          BrowserWindow.getAllWindows()[0].getBackgroundColor().toLowerCase(),
        ),
      )
      .toBe(theme === 'dark' ? '#141c2b' : '#f5f8fc')
  }
  await page
    .getByRole('button', { name: '切换主题', exact: true })
    .filter({ visible: true })
    .click()
  await expect(page.getByRole('menuitemradio', { name: label, exact: true })).toHaveAttribute(
    'aria-checked',
    'true',
  )
  await page.screenshot({ path: join(results, `theme-menu-${theme}.png`) })
  await page.keyboard.press('Escape')
}
async function checkMarkdown(scope) {
  const tables = scope.locator('.markdown-body table')
  await expect(tables).toHaveCount(2)
  await expect(tables.first().locator('th')).toHaveCount(3)
  await expect(tables.first().locator('tbody tr')).toHaveCount(3)
  await expect(tables.first().locator('tbody tr').first().locator('td').last()).toHaveText('')
  await expect(tables.first().getByText('左|右', { exact: true })).toBeVisible()
  await expect(scope.locator('.markdown-body del')).toHaveText('旧说明')
  await expect(scope.locator('.task-list-item input:checked')).toHaveCount(1)
  await expect(scope.locator('pre code')).toHaveText('A    B\n  保留缩进\n')
  const wide = scope.locator('.markdown-table-scroll').last()
  expect(
    await wide.evaluate((element) => {
      const overflow = element.scrollWidth > element.clientWidth
      element.scrollLeft = element.scrollWidth
      return overflow && element.scrollLeft > 0
    }),
  ).toBe(true)
  await expect(scope.locator('.markdown-body img')).toHaveCount(0)
}
async function noPageOverflow(page) {
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1)).toBe(
    true,
  )
}
try {
  let page = await launch()
  await app.evaluate(
    ({ dialog }, paths) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
    },
    fixtures.map(([name]) => join(userData, name)),
  )
  await page.getByRole('button', { name: '添加文件', exact: true }).click()
  await expect(page.getByRole('row', { name: /宽表说明.DOCX/ })).toBeVisible()

  for (const [theme, label] of [
    ['light', '浅色'],
    ['dark', '深色'],
  ]) {
    await setTheme(page, theme, label)
    await expect
      .poll(() => page.evaluate(() => matchMedia('(prefers-color-scheme: dark)').matches))
      .toBe(theme === 'dark')
    await expect(page.locator('.library-shell')).toHaveCSS(
      'background-color',
      theme === 'dark' ? 'rgb(20, 28, 43)' : 'rgb(245, 248, 252)',
    )
    for (const name of ['格式样本.MD', '扩展格式.markdown']) {
      const row = page.getByRole('row', { name: new RegExp(name.replace('.', '\\.')) })
      await row.click()
      await checkMarkdown(page.locator('.preview-body'))
      await noPageOverflow(page)
      if (name.endsWith('.MD'))
        await page.screenshot({ path: join(results, `formatting-${theme}.png`) })
      await row.dblclick()
      await checkMarkdown(page.locator('.editor-preview'))
      expect(await page.getByLabel('Markdown 内容').inputValue()).toContain('| 项目 | 值 | |')
      await page.screenshot({ path: join(results, `formatting-editor-${theme}.png`) })
      if (theme === 'light' && name.endsWith('.MD')) {
        const draft = `${await page.getByLabel('Markdown 内容').inputValue()}\n未保存的主题切换检查`
        await page.getByLabel('Markdown 内容').fill(draft)
        await setTheme(page, 'dark', '深色')
        await expect(page.getByLabel('Markdown 内容')).toHaveValue(draft)
        await expect(page.getByText('有未保存的修改', { exact: true })).toBeVisible()
        await setTheme(page, 'light', '浅色')
        // Keyboard and outside-click behavior of the compact menu.
        await page.getByRole('button', { name: '切换主题' }).filter({ visible: true }).click()
        await page.keyboard.press('End')
        await expect(page.getByRole('menuitemradio', { name: '跟随系统' })).toBeFocused()
        await page.keyboard.press('Home')
        await expect(page.getByRole('menuitemradio', { name: '浅色', exact: true })).toBeFocused()
        await page.keyboard.press('Escape')
        await expect(
          page.getByRole('button', { name: '切换主题' }).filter({ visible: true }),
        ).toBeFocused()
        await page.getByRole('button', { name: '切换主题' }).filter({ visible: true }).click()
        await page.getByLabel('Markdown 内容').click()
        await expect(page.getByRole('menu', { name: '外观主题' })).toHaveCount(0)
        await page.getByRole('button', { name: '返回列表' }).click()
        await page.getByRole('button', { name: '放弃修改并返回' }).click()
        continue
      }
      await page.getByRole('button', { name: '返回列表' }).click()
    }
    await page.getByRole('row', { name: /宽表说明.DOCX/ }).click()
    const word = page.locator('.preview-body .word-layout')
    await expect(word.getByText('合并标题', { exact: true })).toBeVisible()
    await expect(word.locator('td[colspan="8"]')).toHaveCount(1)
    await expect(word.locator('img')).toHaveCount(1)
    const wide = word.getByRole('region', { name: 'Word 表格' }).last()
    expect(
      await wide.evaluate((el) => {
        el.scrollLeft = el.scrollWidth
        return el.scrollLeft > 0
      }),
    ).toBe(true)
    await expect(word.getByText('第 8 列', { exact: false })).toBeVisible()
    await expect(word.locator('section.word-page')).toHaveCSS(
      'background-color',
      'rgb(255, 255, 255)',
    )
    await page.screenshot({ path: join(results, `formatting-word-${theme}.png`) })
    await noPageOverflow(page)
    await page.getByRole('row', { name: /宽表说明.DOCX/ }).dblclick()
    await expect(page.locator('.editor-workspace .word-layout').getByText('合并标题')).toBeVisible()
    await page.getByRole('button', { name: '正文', exact: true }).click()
    await expect(page.locator('.word-text')).toContainText('第二行中文')
    await page.getByRole('button', { name: '返回列表' }).click()

    await page.getByRole('row', { name: /格式工作簿.XLSX/ }).click()
    const sheet = page.locator('.preview-body .office-reader')
    await sheet.getByRole('tab', { name: '预算（隐藏表）' }).click()
    await expect(sheet.getByRole('table', { name: '预算 工作表' })).toBeVisible()
    await expect(sheet.getByRole('button', { name: 'B2 2024-01-01', exact: true })).toBeVisible()
    await expect(sheet.getByRole('button', { name: 'B3 12.50%', exact: true })).toBeVisible()
    await expect(sheet.getByRole('button', { name: 'B4 12,345.60', exact: true })).toBeVisible()
    await sheet.getByRole('button', { name: 'B6 25.00%', exact: true }).click()
    await expect(sheet.locator('.formula-bar')).toContainText('=1/4')
    await page.screenshot({ path: join(results, `formatting-excel-${theme}.png`) })
    await page.getByRole('row', { name: /格式工作簿.XLSX/ }).dblclick()
    const fullSheet = page.locator('.editor-workspace .office-reader')
    await fullSheet.getByRole('tab', { name: '预算（隐藏表）' }).click()
    await expect(fullSheet.getByRole('button', { name: 'B5 00042', exact: true })).toBeVisible()
    await expect(
      fullSheet.getByRole('button', { name: 'B7 第一行 第二行', exact: true }),
    ).toBeVisible()
    await expect(fullSheet.getByRole('button', { name: 'B9 #DIV/0!', exact: true })).toBeVisible()
    await page.getByRole('button', { name: '返回列表' }).click()
    await page.getByRole('row', { name: /损坏工作簿.xlsx/ }).click()
    await expect(page.getByText('暂时无法预览', { exact: true })).toBeVisible()
    await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1000, 682))
    await page.getByRole('row', { name: /格式样本.MD/ }).click()
    await checkMarkdown(page.locator('.preview-body'))
    await noPageOverflow(page)
    const themeButton = page.getByRole('button', { name: '切换主题' }).filter({ visible: true })
    const bounds = await themeButton.boundingBox()
    expect(bounds.x + bounds.width).toBeLessThanOrEqual(await page.evaluate(() => innerWidth - 138))
    await themeButton.click()
    await expect(page.getByRole('menuitemradio', { name: '跟随系统' })).toBeVisible()
    await page.keyboard.press('Escape')
    await page.screenshot({ path: join(results, `formatting-narrow-${theme}.png`) })
    await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1536, 1024))
  }
  await app.close()
  app = undefined
  page = await launch()
  expect(await page.evaluate(() => window.localDocs.getTheme())).toBe('dark')
  await expect
    .poll(() => page.evaluate(() => matchMedia('(prefers-color-scheme: dark)').matches))
    .toBe(true)
  await setTheme(page, 'system', '跟随系统')
  expect(await page.evaluate(() => matchMedia('(prefers-color-scheme: dark)').matches)).toBe(
    await app.evaluate(({ nativeTheme }) => nativeTheme.shouldUseDarkColors),
  )
  // Simulate native appearance updates only inside the test app; never change Windows settings.
  for (const value of ['dark', 'light']) {
    await app.evaluate(({ nativeTheme }, value) => {
      nativeTheme.themeSource = value
    }, value)
    await expect
      .poll(() => page.evaluate(() => matchMedia('(prefers-color-scheme: dark)').matches))
      .toBe(value === 'dark')
    expect(await page.evaluate(() => window.localDocs.getTheme())).toBe('system')
  }
  await app.evaluate(({ nativeTheme }) => {
    nativeTheme.themeSource = 'system'
  })
  await app.close()
  app = undefined
  page = await launch()
  expect(await page.evaluate(() => window.localDocs.getTheme())).toBe('system')
  expect(await app.evaluate(({ nativeTheme }) => nativeTheme.themeSource)).toBe('system')
  const documents = await page.evaluate(async () => (await window.localDocs.snapshot()).documents)
  for (const [name, bytes] of fixtures) {
    const doc = documents.find((doc) => doc.name === name)
    expect(doc).toBeTruthy()
    const path = join(userData, 'export-' + name)
    await app.evaluate(({ dialog }, path) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
    }, path)
    expect(await page.evaluate((id) => window.localDocs.exportDocument(id), doc.id)).toBe(true)
    expect(await readFile(path)).toEqual(bytes)
    expect(await readFile(join(userData, name))).toEqual(bytes)
  }
  expect(await page.evaluate(() => window.__unsafe)).toBeUndefined()
  expect(errors).toEqual([])
  expect(network).toEqual([])
  console.log(
    'Formatting desktop passed: MD/markdown tables and editor, DOCX wide/merged tables and encoded text, XLSX formatted cells in preview and reader, light/dark/system and restart, 1000px layout, source/export bytes and offline rendering.',
  )
} catch (error) {
  const page = app?.windows()[0]
  if (page) await page.screenshot({ path: join(results, 'formatting-failure.png') }).catch(() => {})
  throw error
} finally {
  if (app) await app.close()
}
