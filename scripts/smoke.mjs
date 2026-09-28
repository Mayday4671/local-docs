import { _electron as electron, expect } from '@playwright/test'
import { mkdtemp, mkdir, writeFile, readFile } from 'node:fs/promises'
import { resolve, join } from 'node:path'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const userData = await mkdtemp(join(results, 'desktop-'))
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
    timeout: 30000,
  })
  if (executablePath) expect(await app.evaluate(({ app }) => app.isPackaged)).toBe(true)
  const page = await app.firstWindow()
  // Windows needs a composited window for packaged-app screenshots; do not steal focus.
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].showInactive())
  page.setDefaultTimeout(15000)
  page.on('pageerror', (error) => errors.push(String(error)))
  await expect(page.getByRole('heading', { name: '全部文件' })).toBeVisible()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  return page
}

try {
  let page = await launch()
  await expect(page.getByText('这里还没有文件')).toBeVisible()
  await page.screenshot({ path: join(results, 'desktop-empty.png') })
  await page.getByRole('button', { name: '新建分类', exact: true }).click()
  await page.getByLabel('分类名称').fill('项目资料')
  await page.getByRole('button', { name: '创建', exact: true }).click()
  await expect(page.getByRole('heading', { name: '项目资料' })).toBeVisible()

  const source = join(userData, '导入记录.md')
  await writeFile(source, '# 原始文档\n原文件不会被修改。', 'utf8')
  await app.evaluate(({ dialog }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] })
  }, source)
  await page.getByRole('button', { name: '添加文件', exact: true }).click()
  await expect(page.getByRole('row', { name: /导入记录.md/ })).toBeVisible()
  await page.getByRole('button', { name: '列表选项', exact: true }).click()
  await page.getByRole('button', { name: '新建 Markdown', exact: true }).click()
  await page.getByLabel('文件名称').fill('开发笔记')
  await page.getByRole('button', { name: '创建', exact: true }).click()
  await page.getByLabel('Markdown 内容').fill('# 项目启动\n\n第一版内容：端口配置。')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('已保存在本地', { exact: true })).toBeVisible()
  await page.getByLabel('Markdown 内容').fill('# 项目启动\n\n第二版内容。')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await expect(page.getByText('已保存在本地', { exact: true })).toBeVisible()
  await page.getByRole('button', { name: '历史版本', exact: true }).click()
  await page.getByRole('button', { name: '恢复此版本', exact: true }).first().click()
  await expect(page.getByLabel('Markdown 内容')).toHaveValue('# 项目启动\n\n第一版内容：端口配置。')
  await page.screenshot({ path: join(results, 'desktop-editor.png') })
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('button', { name: '关闭提示', exact: true }).click()
  await page.getByRole('textbox', { name: '搜索文档' }).fill('端口')
  await expect(page.getByRole('row', { name: /开发笔记.md/ })).toBeVisible()
  await expect(page.getByRole('row', { name: /导入记录.md/ })).toHaveCount(0)
  await page.getByRole('row', { name: /开发笔记.md/ }).dblclick()
  await expect(page.getByLabel('Markdown 内容')).toHaveValue('# 项目启动\n\n第一版内容：端口配置。')
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await expect(page.getByRole('textbox', { name: '搜索文档' })).toHaveValue('端口')
  await page.getByRole('button', { name: '移入回收站', exact: true }).click()
  await page.getByRole('button', { name: /^回收站/ }).click()
  await page.getByRole('row', { name: /开发笔记.md/ }).click()
  await page.getByRole('button', { name: '恢复文档', exact: true }).click()
  await expect(page.getByRole('row', { name: /开发笔记.md/ })).toHaveCount(0)
  await page
    .getByRole('navigation', { name: '文档导航' })
    .getByRole('button', { name: /全部文件/ })
    .click()
  await page.getByRole('row', { name: /开发笔记.md/ }).click()
  const exported = join(userData, '导出副本.md')
  await app.evaluate(({ dialog }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath })
  }, exported)
  await page.getByRole('button', { name: '文档选项', exact: true }).click()
  await page.getByRole('button', { name: '导出', exact: true }).click()
  await expect(page.getByText('已导出文档副本', { exact: true })).toBeVisible()
  expect(await readFile(exported, 'utf8')).toBe('# 项目启动\n\n第一版内容：端口配置。')
  expect(await readFile(source, 'utf8')).toBe('# 原始文档\n原文件不会被修改。')
  await page.getByRole('button', { name: '清空搜索', exact: true }).click()
  await page.getByRole('button', { name: '关闭提示', exact: true }).click()
  await page.screenshot({ path: join(results, 'desktop-library.png') })
  await app.close()
  app = undefined
  page = await launch()
  await page.getByRole('row', { name: /开发笔记.md/ }).dblclick()
  await expect(page.getByLabel('Markdown 内容')).toHaveValue('# 项目启动\n\n第一版内容：端口配置。')
  expect(errors).toEqual([])
  console.log(
    'Desktop smoke passed: import, category, create, save, restore version, search, return context, trash, restore, export and restart persistence.',
  )
  console.log(`Screenshots: ${results}`)
} finally {
  if (app) {
    // This process owns only its temporary test library.
    await app
      .evaluate(({ dialog }) => {
        dialog.showMessageBoxSync = () => 1
      })
      .catch(() => {})
    await app.close()
  }
}
