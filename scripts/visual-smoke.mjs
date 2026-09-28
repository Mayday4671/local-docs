import { _electron as electron, expect } from '@playwright/test'
import { mkdtemp, mkdir, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { officeZip, sheetFixture } from '../tests/fixtures/office.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const userData = await mkdtemp(join(results, 'visual-'))
const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: userData }
delete env.ELECTRON_RUN_AS_NODE
const executablePath = process.env.LOCAL_DOCS_TEST_EXECUTABLE
const app = await electron.launch({ executablePath, args: executablePath ? [] : ['.'], env })
const errors = []
const para = (text, style = '') =>
  `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`
const word = officeZip({
  '[Content_Types].xml':
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  '_rels/.rels':
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  'word/_rels/document.xml.rels':
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  'word/styles.xml': `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/><w:color w:val="34435F"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Title"><w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="340"/></w:pPr><w:rPr><w:sz w:val="34"/><w:b/><w:color w:val="101A35"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:spacing w:before="300" w:after="160"/></w:pPr><w:rPr><w:sz w:val="27"/><w:b/><w:color w:val="101A35"/></w:rPr></w:style></w:styles>`,
  'word/document.xml': `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${para('系统功能设计说明', 'Title')}${para('1. 项目概述', 'Heading1')}${para('本系统用于个人文档的统一管理，支持多类型文档的分类、搜索和快速预览，旨在提高文档查找和使用效率。')}${para('2. 主要功能', 'Heading1')}${['文件分类管理', '全文搜索', '支持 Word、Excel、Markdown 等多种格式', '快速预览', '文件增删改查', '本地存储，无需登录'].map((text) => para(`•　${text}`)).join('')}${para('3. 目录结构设计', 'Heading1')}${para('工作 / 项目文档 / 设计')}${para('按项目整理资料，保留文档原件。')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
})
const filenames = [
  '系统功能设计说明.docx',
  '接口清单.xlsx',
  '前端开发笔记.md',
  '原型图说明.md',
  '页面示意说明.md',
  '需求变更记录.docx',
  '数据结构设计.xlsx',
  '部署说明.md',
  '测试用例.docx',
  '架构图说明.md',
  '技术方案.docx',
  '版本迭代记录.md',
]
const paths = filenames.map((name) => join(userData, name))
for (let index = 0; index < paths.length; index++)
  await writeFile(
    paths[index],
    filenames[index].endsWith('.docx')
      ? word
      : filenames[index].endsWith('.xlsx')
        ? sheetFixture()
        : '# 项目资料\n\n这是用于界面对照的独立测试资料。',
  )
try {
  const page = await app.firstWindow()
  page.setDefaultTimeout(15000)
  page.on('pageerror', (error) => errors.push(String(error)))
  await app.evaluate(({ BrowserWindow }) => {
    const win = BrowserWindow.getAllWindows()[0]
    win.setContentSize(1536, 1024)
    win.showInactive()
  })
  await expect(page.getByRole('heading', { name: '全部文件' })).toBeVisible()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  const designId = await page.evaluate(async () => {
    const api = window.localDocs
    const work = await api.createCategory('工作')
    const project = await api.createCategory('项目文档', work.id)
    await api.createCategory('需求', project.id)
    const design = await api.createCategory('设计', project.id)
    await api.createCategory('开发', project.id)
    await api.createCategory('会议记录', work.id)
    await api.createCategory('产品资料', work.id)
    const tech = await api.createCategory('技术学习')
    for (const name of ['Java', '前端', '数据库']) await api.createCategory(name, tech.id)
    const life = await api.createCategory('生活')
    for (const name of ['旅行', '其他']) await api.createCategory(name, life.id)
    return design.id
  })
  await page.reload()
  await page.getByRole('button', { name: '展开预览', exact: true }).click()
  await page.getByRole('button', { name: '设计', exact: true }).click()
  await app.evaluate(
    ({ dialog }, paths) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
    },
    [...paths].reverse(),
  )
  await page.getByRole('button', { name: '添加文件', exact: true }).click()
  await expect(page.locator('.file-table tbody tr')).toHaveCount(12)
  await page.getByRole('button', { name: '关闭导入提示' }).click()
  await page.getByRole('button', { name: '关闭提示' }).click()
  // Stable order makes comparison with the supplied mock meaningful; use the real name sort UI.
  await page.getByRole('row', { name: /系统功能设计说明.docx/ }).click()
  await expect(
    page.locator('.preview-panel .word-layout').getByText('系统功能设计说明', { exact: true }),
  ).toBeVisible()
  await page.screenshot({
    path: join(results, 'design-desktop.png'),
    scale: 'css',
    clip: { x: 0, y: 0, width: 1536, height: 1024 },
  })
  const icon = await page.evaluate(async () => {
    const img = new Image()
    img.src = './icons/app.svg'
    await img.decode()
    const canvas = document.createElement('canvas')
    canvas.width = canvas.height = 256
    canvas.getContext('2d').drawImage(img, 0, 0, 256, 256)
    return canvas.toDataURL('image/png').split(',')[1]
  })
  await writeFile(join(results, 'app-icon.png'), Buffer.from(icon, 'base64'))
  const dims = await page.evaluate(() => ({
    width: innerWidth,
    height: innerHeight,
    dpr: devicePixelRatio,
    overflow: document.documentElement.scrollWidth > innerWidth,
  }))
  console.log('Visual viewport:', dims)
  expect(dims.overflow).toBe(false)
  await page.getByRole('tab', { name: '标签', exact: true }).click()
  await page.getByLabel('文档标签').fill('验收标签，待整理')
  await page.getByLabel('文档备注').fill('独立验收资料，不写入用户文档库。')
  await page.getByRole('button', { name: '保存标签与备注' }).click()
  await expect(page.locator('.tag-chips')).toContainText('验收标签')
  await page.getByRole('textbox', { name: '搜索文档' }).fill('验收标签')
  await expect(page.locator('.file-table tbody tr')).toHaveCount(1)
  await page.getByRole('button', { name: '清空搜索' }).click()
  await page.getByRole('tab', { name: '信息', exact: true }).click()
  await expect(page.getByLabel('所在分类')).toHaveValue(designId)
  await page.getByRole('button', { name: '收藏', exact: true }).click()
  await expect(page.locator('.file-favorite[aria-pressed="true"]')).toHaveCount(1)
  await app.evaluate(({ clipboard }) => {
    clipboard.writeText = (value) => {
      globalThis.__copiedPath = value
    }
  })
  await page.getByRole('button', { name: '复制路径', exact: true }).click()
  const clipboardPath = await app.evaluate(() => globalThis.__copiedPath)
  expect(clipboardPath).toContain(join(userData, 'library', 'objects'))
  await page.getByRole('button', { name: '网格视图' }).click()
  await expect(page.locator('.file-card')).toHaveCount(12)
  await page.getByRole('button', { name: '列表视图' }).click()
  await page.getByRole('tab', { name: '预览', exact: true }).click()
  await page.getByRole('button', { name: '关闭提示' }).click()
  await app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()[0].setContentSize(1000, 680),
  )
  await page.screenshot({ path: join(results, 'design-narrow.png'), scale: 'css' })
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true)
  const actionBar = await page.locator('.preview-actions').boundingBox()
  expect(actionBar.x + actionBar.width).toBeLessThanOrEqual(1000)
  await app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()[0].setContentSize(1536, 1024),
  )
  await page.getByRole('button', { name: '新建文件夹', exact: true }).click()
  await expect(page.getByLabel('上级分类')).toHaveValue(designId)
  await page.getByLabel('分类名称').fill('子分类验收')
  await page.getByRole('button', { name: '创建', exact: true }).click()
  await expect(page.getByRole('heading', { name: '子分类验收', exact: true })).toBeVisible()
  await page.getByRole('button', { name: '折叠 工作', exact: true }).click()
  await expect(page.getByRole('button', { name: '子分类验收', exact: true })).toHaveCount(0)
  await page.getByRole('button', { name: '展开 工作', exact: true }).click()
  await expect(page.getByRole('button', { name: '子分类验收', exact: true })).toBeVisible()
  expect(errors).toEqual([])
  console.log(
    'Visual flow passed: nested folders, preview, tags, notes, search, metadata, favorite, copy path, grid/list and minimum window layout.',
  )
} finally {
  await app.close()
}
