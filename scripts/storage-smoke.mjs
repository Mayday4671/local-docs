import { _electron as electron, expect } from '@playwright/test'
import { mkdir, mkdtemp, readFile, readdir, unlink, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { wordFixture, sheetFixture } from '../tests/fixtures/office.ts'

const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'storage-desktop-'))
const installation = join(root, 'installed-app'),
  target = join(root, '新文档库')
let profile = join(root, 'profile'),
  app
await mkdir(installation)
await mkdir(target)
const errors = [],
  network = []
async function launch(legacy = false, installationPath = installation) {
  const env = { ...process.env, LOCAL_DOCS_SMOKE: '1', LOCAL_DOCS_DATA_DIR: profile }
  delete env.ELECTRON_RUN_AS_NODE
  delete env.LOCAL_DOCS_TEST_INSTALL_DIR
  if (!legacy) env.LOCAL_DOCS_TEST_INSTALL_DIR = installationPath
  const executablePath = process.env.LOCAL_DOCS_TEST_EXECUTABLE
  app = await electron.launch({
    executablePath,
    args: executablePath ? [] : ['.'],
    cwd: executablePath ? root : undefined,
    env,
    timeout: 30_000,
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
      (_, cb) => cb({ cancel: true }),
    )
  })
  await expect(page.getByRole('heading', { name: '全部文件', exact: true })).toBeVisible()
  return page
}
async function choose(paths) {
  await app.evaluate(({ dialog }, paths) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: paths })
  }, paths)
}
async function capture(page, id) {
  return page.evaluate(async (id) => {
    const api = window.localDocs
    const snapshot = await api.snapshot()
    const office = []
    for (const doc of snapshot.documents.filter((d) => ['.docx', '.xlsx'].includes(d.extension)))
      office.push(await api.readOffice(doc.id))
    return {
      snapshot,
      versions: await api.versions(id),
      annotations: await api.annotations(id),
      draft: await api.draft(id),
      attachments: await api.attachments(id),
      content: await api.readDocument(id),
      theme: await api.getTheme(),
      office,
    }
  }, id)
}
try {
  let page = await launch()
  const source = join(installation, 'data', 'library')
  expect((await page.evaluate(() => window.localDocs.storageInfo())).path).toBe(source)
  expect(await readdir(profile)).not.toContain('library')
  const doc = await page.evaluate(async () => {
    const api = window.localDocs
    const category = await api.createCategory('工作资料')
    await api.createCategory('空子分类', category.id)
    let doc = await api.createMarkdown('迁移测试', category.id)
    doc = await api.saveMarkdown(doc.id, '# 第一版\n\n合成资料', doc.revision)
    doc = await api.saveMarkdown(doc.id, '# 第二版\n\n合成资料', doc.revision)
    await api.updateDocument(doc.id, { favorite: true, tags: ['迁移'], notes: '保留全部内容' })
    await api.saveAnnotation(
      doc.id,
      {
        quote: '合成资料',
        offset: 7,
        prefix: '',
        suffix: '',
        scope: 'markdown',
        color: 'yellow',
        note: '保留标记',
      },
      doc.revision,
    )
    await api.saveDraft(doc.id, '# 未保存的恢复草稿', doc.revision, 'markdown')
    const trash = await api.createMarkdown('回收站测试', null)
    await api.trashDocument(trash.id)
    await api.setTheme('dark')
    return doc
  })
  const imagePath = join(root, '附件.png')
  await writeFile(
    imagePath,
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
      'base64',
    ),
  )
  await choose([imagePath])
  await page.evaluate((id) => window.localDocs.addImage(id), doc.id)
  const imports = [join(root, '正文.docx'), join(root, '表格.xlsx'), join(root, '查询.sql')]
  await writeFile(imports[0], Buffer.from(wordFixture()))
  await writeFile(imports[1], Buffer.from(sheetFixture()))
  await writeFile(imports[2], 'SELECT 1; -- 迁移后仍可搜索')
  await choose(imports)
  expect((await page.evaluate(() => window.localDocs.importFiles(null))).failures).toEqual([])
  await page.reload()
  const before = await capture(page, doc.id)
  await page.getByRole('button', { name: '设置', exact: true }).click()
  await expect(page.locator('.storage-card')).toContainText(source)
  await choose([target])
  await page.getByRole('button', { name: '修改位置', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '修改文档库位置', exact: true })
  await expect(dialog).toBeVisible()
  await expect(dialog).toContainText('5 份文件')
  await expect(dialog).toContainText(target)
  await page.screenshot({ path: join(results, 'storage-move-preview.png') })
  await dialog.getByRole('button', { name: '取消', exact: true }).click()
  await expect(dialog).toHaveCount(0)
  expect(await capture(page, doc.id)).toEqual(before)
  expect(await readdir(target)).toEqual([])
  await page.getByRole('button', { name: '修改位置', exact: true }).click()
  await dialog.getByRole('button', { name: '迁移并使用此位置', exact: true }).click()
  await expect(page.getByText('文档库已迁移，新位置已生效', { exact: true })).toBeVisible()
  await expect(page.locator('.storage-card')).toContainText(target)
  await expect(page.locator('.storage-notice')).toContainText(source)
  const after = await capture(page, doc.id)
  expect(after).toEqual({ ...before, snapshot: { ...before.snapshot, storagePath: target } })
  expect(JSON.parse(await readFile(join(profile, 'storage-location.json'), 'utf8')).path).toBe(
    target,
  )
  for (const name of await readdir(join(source, 'objects')))
    expect(await readFile(join(source, 'objects', name))).toEqual(
      await readFile(join(target, 'objects', name)),
    )
  await page.screenshot({ path: join(results, 'storage-settings-moved.png') })
  // The new services and protocol must use the new root, not their original closures.
  await page.evaluate((id) => window.localDocs.copyDocumentPath(id), doc.id)
  expect(await app.evaluate(({ clipboard }) => clipboard.readText())).toContain(target)
  expect(await page.evaluate(() => window.localDocs.search('迁移后仍可搜索'))).toHaveLength(1)
  const image = after.attachments[0]
  expect(
    await page.evaluate(
      async ({ id, imageId }) => (await fetch(`app://local/attachment/${id}/${imageId}`)).status,
      { id: doc.id, imageId: image.id },
    ),
  ).toBe(200)
  const backup = join(root, '迁移后备份.localdocs-backup')
  await app.evaluate(({ dialog }, path) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
  }, backup)
  expect(await page.evaluate(() => window.localDocs.createBackup())).not.toBeNull()
  expect((await readFile(backup)).length).toBeGreaterThan(1000)
  // Folder import after the swap also has a fresh service and stores objects in the new library.
  const folder = join(root, '待导入')
  await mkdir(folder)
  await writeFile(join(folder, '迁移后.txt'), '只在新路径中保存')
  await choose([folder])
  const plan = await page.evaluate(() => window.localDocs.previewFolder())
  await page.evaluate((token) => window.localDocs.importFolder(token, null, 'skip'), plan.token)
  const final = await capture(page, doc.id)
  await app.close()
  app = null
  page = await launch()
  expect(await capture(page, doc.id)).toEqual(final)
  expect((await page.evaluate(() => window.localDocs.storageInfo())).path).toBe(target)
  await app.close()
  app = null

  // Simulate a 0.9 profile without a pointer, then start with a new installation directory.
  profile = join(root, 'legacy-profile')
  page = await launch(true)
  const legacyDoc = await page.evaluate(() => window.localDocs.createMarkdown('旧版资料', null))
  const legacyState = await capture(page, legacyDoc.id)
  await app.close()
  app = null
  await unlink(join(profile, 'storage-location.json'))
  // The first scenario's default still exists; use another isolated installation for migration.
  const legacyInstall = join(root, 'upgraded-app')
  await mkdir(legacyInstall)
  page = await launch(false, legacyInstall)
  expect(await capture(page, legacyDoc.id)).toEqual({
    ...legacyState,
    snapshot: { ...legacyState.snapshot, storagePath: join(legacyInstall, 'data', 'library') },
  })
  expect((await page.evaluate(() => window.localDocs.storageInfo())).notice).toContain(
    '已将旧版资料迁移',
  )
  expect((await readFile(join(profile, 'library', 'library.sqlite'))).length).toBeGreaterThan(0)
  expect(errors).toEqual([])
  expect(network).toEqual([])
  console.log(`Default location, live migration, restart and legacy upgrade verified: ${root}`)
} finally {
  await app?.close()
}
