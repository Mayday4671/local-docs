import { _electron as electron, expect } from '@playwright/test'
import { mkdtemp, mkdir, writeFile, readFile } from 'node:fs/promises'
import { resolve, join } from 'node:path'
import { wordFixture, sheetFixture } from '../tests/fixtures/office.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const userData = await mkdtemp(join(results, 'office-desktop-'))
const wordPath = join(userData, '网络说明.docx')
const sheetPath = join(userData, '设备资料.xlsx')
const badPath = join(userData, '损坏文档.docx')
const originalWord = Buffer.from(wordFixture())
await writeFile(wordPath, originalWord)
await writeFile(sheetPath, sheetFixture())
await writeFile(badPath, 'Not a ZIP file')
const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: userData }
delete env.ELECTRON_RUN_AS_NODE
let app
const errors = []
const network = []
async function launch() {
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
  page.on('request', (request) => {
    if (/^https?:/.test(request.url())) network.push(request.url())
  })
  await app.evaluate(({ session }) => {
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ['http://*/*', 'https://*/*'] },
      (_details, callback) => callback({ cancel: true }),
    )
  })
  await expect(page.getByRole('heading', { name: '全部文件' })).toBeVisible()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  return page
}
try {
  let page = await launch()
  await app.evaluate(
    ({ dialog }, paths) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
    },
    [wordPath, sheetPath, badPath],
  )
  await page.getByRole('button', { name: '添加文件', exact: true }).click()
  await expect(page.getByRole('row', { name: /网络说明.docx/ })).toBeVisible()
  await page.getByRole('row', { name: /网络说明.docx/ }).dblclick()
  const word = page.getByRole('region', { name: 'Word 阅读器' })
  await expect(word.getByText('项目网络说明', { exact: true })).toBeVisible()
  await expect(word.locator('.word-layout img')).toHaveCount(1)
  await expect(word.getByText('网络运维资料', { exact: true })).toBeVisible()
  await page.screenshot({ path: join(results, 'office-word.png') })
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('textbox', { name: '搜索文档' }).fill('端口')
  await expect(page.getByRole('row', { name: /网络说明.docx/ })).toBeVisible()
  await expect(page.getByRole('row', { name: /设备资料.xlsx/ })).toBeVisible()
  await expect(page.getByText('设备清单 · D120', { exact: true }).first()).toBeVisible()
  await page.getByRole('row', { name: /网络说明.docx/ }).dblclick()
  await expect(page.locator('.word-text .located')).toContainText('请检查端口配置与连接方式。')
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await expect(page.getByRole('textbox', { name: '搜索文档' })).toHaveValue('端口')
  await page.getByRole('row', { name: /设备资料.xlsx/ }).dblclick()
  const sheet = page.getByRole('region', { name: 'Excel 阅读器' })
  await expect(sheet.getByRole('tab', { name: '设备清单', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  )
  await expect(sheet.locator('.located')).toContainText('二号机端口')
  await expect(sheet.locator('.formula-bar strong')).toHaveText('D120')
  await page.screenshot({ path: join(results, 'office-excel.png') })
  await sheet.getByRole('tab', { name: '预算（隐藏表）', exact: true }).click()
  await sheet.getByRole('button', { name: 'D2 300', exact: true }).click()
  await expect(sheet.locator('.formula-bar')).toContainText('=SUM(B2:C2)')
  await expect(sheet.getByRole('button', { name: /E2.*无缓存结果/ })).toBeVisible()
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('textbox', { name: '搜索文档' }).fill('费用')
  await page.getByRole('row', { name: /设备资料.xlsx/ }).dblclick()
  await expect(sheet.getByRole('tab', { name: '预算（隐藏表）', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  )
  await expect(sheet.locator('.located')).toContainText('费用汇总')
  const output = join(userData, '导出.xlsx')
  await app.evaluate(({ dialog }, path) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
  }, output)
  // Only the visible reader's export action should be used.
  await page.locator('.editor-header').getByRole('button', { name: '导出', exact: true }).click()
  await expect(page.getByText('已导出文档副本', { exact: true })).toBeVisible()
  expect(await readFile(output)).toEqual(await readFile(sheetPath))
  expect(await readFile(wordPath)).toEqual(originalWord)
  await page.getByRole('button', { name: '返回列表', exact: true }).click()
  await page.getByRole('button', { name: '清空搜索', exact: true }).click()
  await page.getByRole('row', { name: /损坏文档.docx/ }).click()
  await expect(page.getByText('暂时无法预览', { exact: true })).toBeVisible()
  await expect(page.getByText('文件副本已保留，可以原样导出。', { exact: true })).toBeVisible()
  await app.close()
  app = undefined
  page = await launch()
  await page.getByRole('textbox', { name: '搜索文档' }).fill('端口')
  await expect(page.getByRole('row', { name: /设备资料.xlsx/ })).toBeVisible()
  expect(errors).toEqual([])
  expect(network).toEqual([])
  console.log(
    'Office desktop passed: offline DOCX layout, XLSX sheets and formula cache, search snippets and actual cell/paragraph location, source/export bytes, corrupt-file fallback and index persistence.',
  )
} catch (error) {
  const page = app?.windows()[0]
  if (page) await page.screenshot({ path: join(results, 'office-failure.png') }).catch(() => {})
  throw error
} finally {
  if (app) await app.close()
}
