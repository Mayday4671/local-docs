import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, writeFile, readFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { wordFixture, sheetFixture } from '../tests/fixtures/office.ts'
const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'content-desktop-')),
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
async function selectText(page, scope, quote) {
  await page.locator(scope).evaluate((element, quote) => {
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT)
    let node
    while ((node = walker.nextNode())) {
      const at = node.textContent.indexOf(quote)
      if (at < 0) continue
      const range = document.createRange()
      range.setStart(node, at)
      range.setEnd(node, at + quote.length)
      const selection = window.getSelection()
      selection.removeAllRanges()
      selection.addRange(range)
      element.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }))
      return
    }
    throw new Error('Selection text not found')
  }, quote)
}
try {
  let page = await launch()
  const md = await page.evaluate(async () => {
    const d = await window.localDocs.createMarkdown('阅读笔记', null)
    await window.localDocs.saveMarkdown(
      d.id,
      '# 阅读笔记\n\n这是需要标记的重点内容。\n\n另一段文字用于查找。',
      1,
    )
    await window.localDocs.setTheme('light')
    return d
  })
  await choose(join(root, '段落测试.docx'))
  await page.evaluate(() => window.localDocs.importFiles(null))
  await choose(join(root, '单元格测试.xlsx'))
  await page.evaluate(() => window.localDocs.importFiles(null))
  await page.reload()
  await page.getByRole('button', { name: '收藏 阅读笔记.md', exact: true }).click()
  await page.getByRole('button', { name: /^收藏 1$/ }).click()
  await expect(page.locator('.file-table tbody tr')).toHaveCount(1)
  await page.getByRole('row', { name: /阅读笔记.md/ }).click()
  await selectText(page, '.preview-panel [data-mark-scope="markdown"]', '重点内容')
  await page.getByRole('button', { name: '标记所选内容', exact: true }).click()
  await page.getByLabel('标记备注', { exact: true }).fill('整理时重点核对')
  await page.getByLabel('标记颜色', { exact: true }).selectOption('green')
  await page.getByRole('button', { name: '保存标记', exact: true }).click()
  await expect(page.getByRole('complementary', { name: '内容标记', exact: true })).toContainText(
    '整理时重点核对',
  )
  expect(await page.evaluate((id) => window.localDocs.annotations(id), md.id)).toHaveLength(1)
  await page.getByLabel('文内查找', { exact: true }).fill('文字')
  await expect(page.getByRole('button', { name: '1/1 下一处', exact: true })).toBeVisible()
  await page.screenshot({ path: join(results, 'content-annotation-light.png') })
  await page.getByRole('button', { name: /^内容标记 1$/ }).click()
  await expect(page.locator('.file-table tbody tr')).toHaveCount(1)
  await page.getByRole('row', { name: /阅读笔记.md/ }).dblclick()
  const source = page.getByLabel('Markdown 内容', { exact: true })
  await source.fill('# 阅读笔记\n\n自动保存的恢复草稿，仍有重点内容。')
  await expect(page.getByText('恢复草稿已自动保存', { exact: true })).toBeVisible()
  expect((await page.evaluate((id) => window.localDocs.versions(id), md.id)).length).toBe(2)
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('button', { name: '保留草稿并返回', exact: true }).click()
  await page.getByRole('row', { name: /阅读笔记.md/ }).dblclick()
  await page.getByRole('button', { name: '恢复草稿', exact: true }).click()
  expect(await source.inputValue()).toContain('恢复草稿')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('已保存在本地', { exact: true })).toBeVisible()
  // Synthetic image uses the actual native image decoder and local attachment protocol.
  const image = join(root, '图片样本.png')
  await writeFile(
    image,
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
      'base64',
    ),
  )
  await choose(image)
  await page.getByRole('button', { name: '插入图片', exact: true }).click()
  await expect(page.locator('.editor-preview .managed-image')).toBeVisible()
  expect(
    await page.locator('.managed-image').evaluate((img) => img.complete && img.naturalWidth > 0),
  ).toBe(true)
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('已保存在本地', { exact: true })).toBeVisible()
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('button', { name: /^全部文件 3$/ }).click()
  await page.getByRole('row', { name: /段落测试.docx/ }).dblclick()
  await page.getByRole('button', { name: '正文', exact: true }).click()
  await selectText(page, '.editor-workspace .word-text', '端口')
  await page.getByRole('button', { name: '标记所选内容', exact: true }).click()
  await page.getByLabel('标记备注', { exact: true }).fill('Word 标记')
  await page.getByRole('button', { name: '保存标记', exact: true }).click()
  await page
    .locator('.office-edit-toolbar')
    .getByRole('button', { name: '编辑', exact: true })
    .click()
  await page.getByLabel('编辑 正文 · 段落 1', { exact: true }).fill('已编辑的项目标题')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
  await page.screenshot({ path: join(results, 'content-word-editor.png') })
  await page.getByRole('button', { name: '阅读与标记', exact: true }).click()
  await expect(
    page.locator('.editor-workspace .word-layout').getByText('已编辑的项目标题', { exact: true }),
  ).toBeVisible()
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('row', { name: /单元格测试.xlsx/ }).dblclick()
  await page.getByRole('button', { name: 'A2 核心交换机', exact: true }).click()
  await page.getByRole('button', { name: '标记所选内容', exact: true }).click()
  await page.getByLabel('标记备注', { exact: true }).fill('Excel 标记')
  await page.getByRole('button', { name: '保存标记', exact: true }).click()
  await page
    .locator('.office-edit-toolbar')
    .getByRole('button', { name: '编辑', exact: true })
    .click()
  await page.getByLabel('编辑 设备清单 · A2', { exact: true }).fill('已编辑的交换机')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
  await page.getByRole('button', { name: '切换主题', exact: true }).click()
  await page.getByRole('menuitemradio', { name: '深色', exact: true }).click()
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1000, 650))
  await page.screenshot({ path: join(results, 'content-excel-compact-dark.png') })
  await page.getByRole('button', { name: '阅读与标记', exact: true }).click()
  await expect(page.getByRole('button', { name: 'A2 已编辑的交换机', exact: true })).toBeVisible()
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  expect(await readFile(join(root, '段落测试.docx'))).toEqual(word)
  expect(await readFile(join(root, '单元格测试.xlsx'))).toEqual(sheet)
  const saved = await page.evaluate(() => window.localDocs.snapshot()),
    marks = await page.evaluate(() => window.localDocs.annotations())
  expect(marks).toHaveLength(3)
  await app.close()
  page = await launch()
  expect(await page.evaluate(() => window.localDocs.snapshot())).toEqual(saved)
  expect(await page.evaluate(() => window.localDocs.annotations())).toEqual(marks)
  await page.getByRole('row', { name: /阅读笔记.md/ }).click()
  await expect(page.locator('.preview-panel .managed-image')).toBeVisible()
  expect(await page.evaluate(() => window.localDocs.searchResults('重点核对'))).toHaveLength(1)
  // Create both Office formats through the actual menus, edit and export them.
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(1500, 950))
  for (const kind of ['Word', 'Excel']) {
    await page.getByRole('button', { name: '列表选项', exact: true }).click()
    await page.getByRole('button', { name: `新建 ${kind}`, exact: true }).click()
    await page.getByLabel('文件名称', { exact: true }).fill(`新建${kind}测试`)
    await page.getByRole('button', { name: '创建', exact: true }).click()
    await expect(page.locator('.office-mode-label')).toHaveText('编辑中')
    await expect(page.getByRole('button', { name: '编辑内容', exact: true })).toHaveCount(0)
    await page
      .getByLabel(kind === 'Word' ? '编辑 正文 · 段落 1' : '编辑 工作表1 · A1', { exact: true })
      .fill('新建文件的正文')
    if (kind === 'Word') {
      await page.getByRole('button', { name: '在末尾新增段落', exact: true }).click()
      await page
        .getByLabel('编辑 正文 · 末尾新增段落', { exact: true })
        .fill('第二段正文\n第三段正文')
    }
    await page.getByRole('button', { name: '保存', exact: true }).click()
    await expect(page.getByText('阅读 / 内容编辑 · 原件保留', { exact: true })).toBeVisible()
    await page.getByRole('button', { name: '返回列表', exact: true }).click()
    const record = await page.evaluate(
      async (name) =>
        (await window.localDocs.snapshot()).documents.find((d) => d.name.startsWith(name)),
      `新建${kind}测试`,
    )
    const output = join(root, `新建${kind}${kind === 'Word' ? '.docx' : '.xlsx'}`)
    await app.evaluate(({ dialog }, path) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
    }, output)
    expect(await page.evaluate((id) => window.localDocs.exportDocument(id), record.id)).toBe(true)
  }
  for (const [name, filename] of [
    ['段落测试.docx', '已编辑.docx'],
    ['单元格测试.xlsx', '已编辑.xlsx'],
  ]) {
    const record = saved.documents.find((d) => d.name === name)
    await app.evaluate(
      ({ dialog }, path) => {
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
      },
      join(root, filename),
    )
    expect(await page.evaluate((id) => window.localDocs.exportDocument(id), record.id)).toBe(true)
  }
  // Bypass normal window closing only in this isolated app to exercise recovery.
  await page.getByRole('row', { name: /阅读笔记.md/ }).dblclick()
  const prior = await page.getByLabel('Markdown 内容', { exact: true }).inputValue()
  await page.getByLabel('Markdown 内容', { exact: true }).fill(prior + '\n稍后撤回的输入')
  await expect(page.getByText('恢复草稿已自动保存', { exact: true })).toBeVisible()
  await page.getByLabel('Markdown 内容', { exact: true }).fill(prior)
  await expect.poll(() => page.evaluate((id) => window.localDocs.draft(id), md.id)).toBeNull()
  await page.getByLabel('Markdown 内容', { exact: true }).fill(prior + '\n突然退出后恢复')
  await expect(page.getByText('恢复草稿已自动保存', { exact: true })).toBeVisible()
  const closed = app.waitForEvent('close')
  await app
    .evaluate(({ app }) => app.exit(0))
    .catch((error) => {
      if (!/closed/i.test(String(error))) throw error
    })
  await closed
  app = undefined
  page = await launch()
  await page.getByRole('row', { name: /阅读笔记.md/ }).dblclick()
  await page.getByRole('button', { name: '恢复草稿', exact: true }).click()
  expect(await page.getByLabel('Markdown 内容', { exact: true }).inputValue()).toContain(
    '突然退出后恢复',
  )
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('已保存在本地', { exact: true })).toBeVisible()
  expect(errors).toEqual([])
  expect(network).toEqual([])
  console.log(
    `Favorites, annotations for three formats, drafts, images, Office editing and restart passed: ${root}`,
  )
} catch (error) {
  const page = app?.windows()[0]
  if (page) await page.screenshot({ path: join(results, 'content-failure.png') }).catch(() => {})
  throw error
} finally {
  if (app) await app.close()
}
