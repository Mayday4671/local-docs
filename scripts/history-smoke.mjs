import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile, readFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { wordFixture, sheetFixture } from '../tests/fixtures/office.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'history-'))
const word = Buffer.from(wordFixture()),
  sheet = Buffer.from(sheetFixture())
await writeFile(join(root, '文字对比.docx'), word)
await writeFile(join(root, '表格对比.xlsx'), sheet)
await writeFile(join(root, '损坏样本.docx'), 'invalid synthetic office file')
const errors = [],
  network = []
let app
try {
  const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: join(root, 'data') }
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
    BrowserWindow.getAllWindows()[0].setSize(1500, 950)
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ['http://*/*', 'https://*/*'] },
      (_, cb) => cb({ cancel: true }),
    )
  })
  await page.emulateMedia({ colorScheme: null })
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  await page.evaluate(() => window.localDocs.setTheme('light'))
  await app.evaluate(
    ({ dialog }, paths) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
    },
    ['文字对比.docx', '表格对比.xlsx', '损坏样本.docx'].map((name) => join(root, name)),
  )
  await page.evaluate(() => window.localDocs.importFiles(null))
  const markdown = await page.evaluate(async () => {
    const api = window.localDocs,
      doc = await api.createMarkdown('文本对比', null)
    const first = await api.saveMarkdown(
      doc.id,
      '# 项目说明\n状态：准备中\n需要删除的行\n共同段落',
      doc.revision,
    )
    const second = await api.saveMarkdown(
      doc.id,
      '# 项目说明\n状态：已完成\n共同段落\n新增备注',
      first.revision,
    )
    return second
  })
  await page.reload()
  const edit = async (name) => {
    await page.getByRole('row', { name: new RegExp(name) }).click()
    await page.getByRole('button', { name: '编辑', exact: true }).click()
  }
  const showHistory = async () => {
    await page.getByRole('button', { name: '历史版本', exact: true }).click()
    await page
      .locator('.history-item')
      .first()
      .getByRole('button', { name: /查看差异|查看内容/ })
      .click()
    return page.getByRole('dialog', { name: '历史版本对比', exact: true })
  }
  const closeComparison = async () => {
    await page.getByRole('button', { name: '关闭版本对比' }).click()
  }
  const back = async () => {
    await page.getByRole('button', { name: '返回列表', exact: true }).click()
  }

  await edit('文本对比.md')
  const source = page.getByLabel('Markdown 内容', { exact: true })
  const unsaved = '# 未保存的编辑内容，查看历史后必须保留'
  await source.fill(unsaved)
  await expect(page.getByText('恢复草稿已自动保存', { exact: false }).first()).toBeVisible()
  const baseline = await page.evaluate(
    async (id) => ({
      doc: await window.localDocs.readDocument(id),
      versions: await window.localDocs.versions(id),
      draft: await window.localDocs.draft(id),
    }),
    markdown.id,
  )
  let dialog = await showHistory()
  await expect(dialog.getByRole('status')).toHaveText('新增 1 · 删除 1 · 修改 1')
  await expect(dialog.locator('[data-kind="modified"]')).toContainText('状态：准备中')
  await expect(dialog.locator('[data-kind="modified"]')).toContainText('状态：已完成')
  await page.keyboard.press('Control+s')
  await page.screenshot({ path: join(results, 'history-markdown-light.png') })
  // Both selectors can choose historical versions, including an empty initial version.
  await dialog.getByLabel('对比前版本', { exact: true }).selectOption(baseline.versions[2].id)
  await dialog.getByLabel('对比后版本', { exact: true }).selectOption(baseline.versions[1].id)
  await expect(dialog.getByRole('status')).toHaveText('新增 4 · 删除 0 · 修改 0')
  await dialog.getByLabel('对比前版本', { exact: true }).selectOption(baseline.versions[1].id)
  await expect(dialog.getByText('两个版本的文件内容相同', { exact: true })).toBeVisible()
  await dialog.getByRole('checkbox', { name: '仅看差异' }).uncheck()
  await expect(dialog.getByRole('table', { name: '版本差异' })).toContainText('状态：准备中')
  await page.keyboard.press('Escape')
  await expect(dialog).toHaveCount(0)
  await expect(source).toHaveValue(unsaved)
  expect(
    await page.evaluate(
      async (id) => ({
        doc: await window.localDocs.readDocument(id),
        versions: await window.localDocs.versions(id),
        draft: await window.localDocs.draft(id),
      }),
      markdown.id,
    ),
  ).toEqual(baseline)
  await source.fill(baseline.doc.text)
  await back()

  // Long comparisons stay bounded and changing versions resets pagination.
  await page.evaluate(async () => {
    const api = window.localDocs,
      doc = await api.createMarkdown('分页对比', null)
    const first = await api.saveMarkdown(
      doc.id,
      Array.from({ length: 205 }, (_, i) => `原文${i}`).join('\n'),
      doc.revision,
    )
    await api.saveMarkdown(
      doc.id,
      Array.from({ length: 205 }, (_, i) => `新文${i}`).join('\n'),
      first.revision,
    )
  })
  await page.reload()
  await edit('分页对比.md')
  dialog = await showHistory()
  await expect(dialog.getByRole('status')).toHaveText('新增 0 · 删除 0 · 修改 205')
  await expect(dialog.locator('tbody tr')).toHaveCount(100)
  await dialog.getByRole('button', { name: '下一页', exact: true }).click()
  await dialog.getByRole('button', { name: '下一页', exact: true }).click()
  await expect(dialog.locator('tbody tr')).toHaveCount(5)
  await expect(dialog.getByRole('table')).toContainText('新文204')
  const latest = await dialog.getByLabel('对比后版本', { exact: true }).inputValue()
  await dialog.getByLabel('对比前版本', { exact: true }).selectOption(latest)
  await expect(dialog.getByText('两个版本的文件内容相同', { exact: true })).toBeVisible()
  await dialog.getByRole('checkbox', { name: '仅看差异' }).uncheck()
  await expect(dialog.getByRole('table')).toContainText('新文0')
  await closeComparison()
  await expect(
    page.locator('.history-item').first().getByRole('button', { name: '查看差异' }),
  ).toBeFocused()
  await back()

  // Word explicit edit lands directly on editable paragraphs, no second click.
  await edit('文字对比.docx')
  const title = page.getByLabel('编辑 正文 · 段落 1', { exact: true })
  await expect(title).toBeVisible()
  await expect(page.locator('.office-mode-label')).toHaveText('编辑中')
  await expect(page.locator('.office-edit-toolbar button')).toHaveCount(1)
  await expect(page.getByRole('button', { name: '编辑内容', exact: true })).toHaveCount(0)
  await expect(
    page.locator('.office-edit-toolbar').getByRole('button', { name: '编辑', exact: true }),
  ).toHaveCount(0)
  await title.fill('历史对比专用标题')
  await page.getByRole('button', { name: '阅读与标记', exact: true }).click()
  await expect(page.locator('.office-mode-label')).toHaveText('阅读中')
  await page.getByRole('button', { name: '返回编辑', exact: true }).click()
  await expect(title).toHaveText('历史对比专用标题')
  await page.screenshot({ path: join(results, 'office-mode-word-edit.png') })
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
  dialog = await showHistory()
  await expect(dialog.getByRole('status')).toHaveText('新增 0 · 删除 0 · 修改 1')
  await expect(dialog.getByRole('table')).toContainText('项目网络说明')
  await expect(dialog.getByRole('table')).toContainText('历史对比专用标题')
  await expect(dialog.getByRole('table')).toContainText('正文 · 段落 1')
  await page.screenshot({ path: join(results, 'history-word-light.png') })
  await closeComparison()
  await back()
  await page.getByRole('row', { name: /文字对比.docx/ }).click()
  await page.getByRole('button', { name: '打开', exact: true }).click()
  await expect(page.locator('.office-mode-label')).toHaveText('阅读中')
  await expect(page.locator('.office-edit-toolbar button')).toHaveText('编辑')
  await expect(page.getByLabel('Word 页面编辑', { exact: true })).toHaveCount(0)
  await back()

  // Excel direct editing, address-based history and saved formula content.
  await edit('表格对比.xlsx')
  await expect(page.locator('.office-mode-label')).toHaveText('编辑中')
  await expect(page.locator('.office-edit-toolbar button')).toHaveText('阅读与标记')
  await expect(page.getByRole('button', { name: '编辑内容', exact: true })).toHaveCount(0)
  await page.getByLabel('编辑 设备清单 · A2', { exact: true }).fill('交换机的新名称')
  await page.getByLabel('编辑 设备清单 · B2', { exact: true }).fill('新建单元格')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
  await page.getByRole('button', { name: '切换主题', exact: true }).click()
  await page.getByRole('menuitemradio', { name: '深色', exact: true }).click()
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1000, 680))
  await page.screenshot({ path: join(results, 'office-mode-excel-dark.png') })
  dialog = await showHistory()
  await expect(dialog.getByRole('table')).toContainText('设备清单 · A2')
  await expect(dialog.getByRole('table')).toContainText('核心交换机')
  await expect(dialog.getByRole('table')).toContainText('交换机的新名称')
  await expect(dialog.locator('[data-kind="added"]')).toContainText('设备清单 · B2')
  await dialog.getByRole('checkbox', { name: '仅看差异' }).uncheck()
  await expect(dialog.getByRole('table')).toContainText('公式：=SUM(B2:C2)')
  await expect(dialog.getByRole('table')).toContainText('预算 · 工作表')
  expect(await dialog.evaluate((e) => e.scrollWidth <= e.clientWidth)).toBe(true)
  await page.screenshot({ path: join(results, 'history-excel-dark-small.png') })
  await closeComparison()
  await back()

  // Newly created Office documents also start in edit mode; one-version view works.
  for (const kind of ['Word', 'Excel']) {
    await page.getByRole('button', { name: '列表选项', exact: true }).click()
    await page.getByRole('button', { name: `新建 ${kind}`, exact: true }).click()
    await page.getByLabel('文件名称', { exact: true }).fill(`新建直接编辑${kind}`)
    await page.getByRole('button', { name: '创建', exact: true }).click()
    await expect(page.locator('.office-mode-label')).toHaveText('编辑中')
    await expect(page.locator('.office-edit-toolbar button')).toHaveText('阅读与标记')
    await expect(page.getByRole('button', { name: '编辑内容', exact: true })).toHaveCount(0)
    await expect(
      page.getByLabel(kind === 'Word' ? '编辑 正文 · 段落 1' : '编辑 工作表1 · A1', {
        exact: true,
      }),
    ).toBeVisible()
    dialog = await showHistory()
    await expect(dialog.getByText('两个版本的文件内容相同', { exact: true })).toBeVisible()
    await closeComparison()
    await back()
  }

  // Unreadable historical files show a recoverable error, never a false empty diff.
  await edit('损坏样本.docx')
  dialog = await showHistory()
  await expect(dialog.getByRole('alert')).toContainText('文件已损坏')
  await expect(dialog.getByRole('button', { name: '重新读取', exact: true })).toBeVisible()
  await closeComparison()
  const documents = await page.evaluate(() => window.localDocs.snapshot())
  const wordDoc = documents.documents.find((d) => d.name === '文字对比.docx')
  const sheetDoc = documents.documents.find((d) => d.name === '表格对比.xlsx')
  expect(
    await page.evaluate(
      async ({ wordId, sheetId }) => {
        const versions = await window.localDocs.versions(wordId)
        try {
          await window.localDocs.readVersion(sheetId, versions[0].id)
          return false
        } catch {
          return true
        }
      },
      { wordId: wordDoc.id, sheetId: sheetDoc.id },
    ),
  ).toBe(true)
  expect(
    (await page.evaluate((id) => window.localDocs.readOffice(id), wordDoc.id)).data.blocks[0].text,
  ).toBe('历史对比专用标题')
  expect(await readFile(join(root, '文字对比.docx'))).toEqual(word)
  expect(await readFile(join(root, '表格对比.xlsx'))).toEqual(sheet)
  expect(errors).toEqual([])
  expect(network).toEqual([])
  console.log(`History comparison and direct editing verified: ${root}`)
} finally {
  await app?.close()
}
