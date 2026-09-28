import { afterEach, beforeEach, expect, test, vi } from 'vitest'
import {
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  realpath,
  rm,
  unlink,
  writeFile,
} from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { tmpdir } from 'node:os'
import { Library } from '../src/main/library'
import { StorageLocation } from '../src/main/storage-location'

let root: string, profile: string, installation: string, location: StorageLocation
const opened: Library[] = []
const keep = (library: Library) => {
  opened.push(library)
  return library
}
const manifest = (library: Library) => {
  const { createdAt, ...value } = library.backupManifest()
  return value
}
beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'local-docs-storage-test-'))
  profile = join(root, 'profile')
  installation = join(root, 'installed-app')
  await mkdir(profile)
  await mkdir(installation)
  location = new StorageLocation(profile, installation)
})
afterEach(async () => {
  vi.restoreAllMocks()
  for (const library of opened.splice(0)) {
    try {
      library.close()
    } catch {}
  }
  const target = await realpath(root)
  if (!target.startsWith(join(await realpath(tmpdir()), 'local-docs-storage-test-')))
    throw new Error('Unsafe test cleanup')
  await rm(target, { recursive: true, force: true })
})
function seed(library: Library) {
  const category = library.createCategory('资料')
  let doc = library.createMarkdown('历史笔记', category.id)
  doc = library.saveMarkdown(doc.id, '# 第一版\n合成资料', doc.revision)
  doc = library.saveMarkdown(doc.id, '# 第二版\n合成资料', doc.revision)
  library.updateDocument(doc.id, { favorite: true, tags: ['测试'], notes: '路径迁移' })
  library.saveAnnotation(
    doc.id,
    {
      quote: '合成资料',
      offset: 6,
      prefix: '',
      suffix: '',
      scope: 'markdown',
      color: 'yellow',
      note: '保留标记',
    },
    doc.revision,
  )
  library.saveDraft(doc.id, '# 尚未正式保存', doc.revision, 'markdown')
  library.addImage(doc.id, Buffer.from([1, 2, 3]), 'image/png')
  const trash = library.createMarkdown('回收站资料', null)
  library.trashDocument(trash.id)
  library.importBytes('查询.sql', Buffer.from('SELECT 1; -- 迁移全文索引'), null, 'keep')
  library.createOffice('表格', '.xlsx', category.id)
  library.setTheme('dark')
  return doc
}
async function target() {
  const path = join(root, 'new-location')
  await mkdir(path)
  return path
}

test('首次使用默认在安装目录，设置只保存地址；重启不创建第二个空库', async () => {
  const library = keep(await location.open())
  expect(library.root).toBe(join(installation, 'data', 'library'))
  seed(library)
  const before = manifest(library)
  library.close()
  opened.pop()
  const reopened = keep(await new StorageLocation(profile, installation).open())
  expect(manifest(reopened)).toEqual(before)
  expect(JSON.parse(await readFile(location.configPath, 'utf8')).path).toBe(reopened.root)
  expect(await readdir(profile)).not.toContain('library')
})

test('旧版资料自动迁入安装目录，保留全部历史、状态和旧副本', async () => {
  const legacy = new Library(join(profile, 'library'))
  seed(legacy)
  const before = manifest(legacy)
  legacy.close()
  const library = keep(await location.open())
  expect(library.root).toBe(location.defaultPath)
  expect(manifest(library)).toEqual(before)
  expect(manifest(keep(new Library(join(profile, 'library'))))).toEqual(before)
  expect(location.info(library).notice).toContain('旧副本保留')
  expect(library.search('迁移全文索引')).toHaveLength(1)
})

test('单份对象缺失不阻止打开其他资料，但阻止迁移产生不完整副本', async () => {
  const original = await location.open()
  const missing = original.importBytes('缺失.txt', Buffer.from('missing object'), null, 'keep')!
  const good = original.importBytes('完好.txt', Buffer.from('intact object'), null, 'keep')!
  const missingPath = original.documentPath(missing.id)
  original.close()
  await unlink(missingPath)
  const reopened = keep(await new StorageLocation(profile, installation).open())
  expect(reopened.snapshot().documents).toHaveLength(2)
  expect(await readFile(reopened.documentPath(good.id), 'utf8')).toBe('intact object')
  await expect(location.prepare(reopened, await target())).rejects.toThrow()
  expect(JSON.parse(await readFile(location.configPath, 'utf8')).path).toBe(reopened.root)
})

test('在线迁移包含 WAL 内容，原路径保留，完整记录与对象字节相同', async () => {
  const original = keep(await location.open())
  seed(original)
  const before = manifest(original),
    path = await target()
  const plan = await location.prepare(original, path)
  const moved = keep(await location.move(original, plan.token))
  expect(manifest(moved)).toEqual(before)
  expect(manifest(original)).toEqual(before)
  for (const object of before.objects)
    expect(await readFile(moved.backupObjectPath(object.hash))).toEqual(
      await readFile(original.backupObjectPath(object.hash)),
    )
  moved.close()
  opened.pop()
  const restarted = keep(await new StorageLocation(profile, installation).open())
  expect(restarted.root).toBe(path)
  expect(manifest(restarted)).toEqual(before)
  restarted.createMarkdown('迁移后新文件', null)
  expect(original.snapshot().documents).toHaveLength(before.documents.length)
})

test('当前库、父子路径、程序资源目录和非空目标都拒绝，不覆盖已有文件', async () => {
  const library = keep(await location.open())
  await expect(location.prepare(library, library.root)).rejects.toThrow('当前文档库')
  const child = join(library.root, 'child')
  await mkdir(child)
  await expect(location.prepare(library, child)).rejects.toThrow('子目录')
  await expect(location.prepare(library, root)).rejects.toThrow('上级目录')
  const resource = join(installation, 'resources')
  await mkdir(resource)
  await expect(location.prepare(library, resource)).rejects.toThrow('程序文件')
  const path = await target()
  await writeFile(join(path, '原有文件.txt'), '不能覆盖')
  await expect(location.prepare(library, path)).rejects.toThrow('非空')
  expect(await readFile(join(path, '原有文件.txt'), 'utf8')).toBe('不能覆盖')
})

test('复制后的再次检查发现目标变化时停止，原库继续可写', async () => {
  const library = keep(await location.open()),
    doc = seed(library),
    path = await target()
  const plan = await location.prepare(library, path)
  const copy = library.copyDatabase.bind(library)
  vi.spyOn(library, 'copyDatabase').mockImplementation(async (destination) => {
    await copy(destination)
    await writeFile(join(path, '外部文件'), '保留')
  })
  await expect(location.move(library, plan.token)).rejects.toThrow('非空')
  expect(JSON.parse(await readFile(location.configPath, 'utf8')).path).toBe(library.root)
  expect(await readFile(join(path, '外部文件'), 'utf8')).toBe('保留')
  expect(library.saveMarkdown(doc.id, '# 仍可编辑', doc.revision).revision).toBe(doc.revision + 1)
  expect((await readdir(root)).some((n) => n.startsWith('.local-docs-move-'))).toBe(false)
})

test('取消迁移、复制失败和无效令牌不切换地址或改动原库', async () => {
  const library = keep(await location.open())
  seed(library)
  const before = manifest(library),
    path = await target()
  await expect(location.move(library, 'invalid')).rejects.toThrow('失效')
  let plan = await location.prepare(library, path)
  vi.spyOn(library, 'copyDatabase').mockImplementation(async () => {
    location.cancel()
  })
  await expect(location.move(library, plan.token)).rejects.toThrow('取消')
  expect(await readdir(path)).toEqual([])
  vi.restoreAllMocks()
  plan = await location.prepare(library, path)
  vi.spyOn(library, 'copyDatabase').mockRejectedValue(new Error('模拟磁盘写入失败'))
  await expect(location.move(library, plan.token)).rejects.toThrow('写入失败')
  expect(manifest(library)).toEqual(before)
  expect(await readdir(path)).toEqual([])
  expect(location.active).toBe(false)
})

test('配置提交失败保留原地址和完整目标副本，不报告成功', async () => {
  const library = keep(await location.open())
  seed(library)
  const before = manifest(library),
    path = await target()
  const plan = await location.prepare(library, path)
  vi.spyOn(location, 'savePath').mockRejectedValue(new Error('配置不可写'))
  await expect(location.move(library, plan.token)).rejects.toThrow('未切换位置')
  expect(JSON.parse(await readFile(location.configPath, 'utf8')).path).toBe(library.root)
  expect(manifest(library)).toEqual(before)
  expect(manifest(keep(new Library(path)))).toEqual(before)
})

test('源对象内容损坏时校验失败，目标不发布', async () => {
  const library = keep(await location.open()),
    doc = seed(library),
    path = await target()
  const plan = await location.prepare(library, path)
  const objectPath = library.documentPath(doc.id),
    bytes = await readFile(objectPath)
  bytes[0] ^= 1
  await writeFile(objectPath, bytes)
  await expect(location.move(library, plan.token)).rejects.toThrow('校验失败')
  expect(await readdir(path)).toEqual([])
  expect(JSON.parse(await readFile(location.configPath, 'utf8')).path).toBe(library.root)
})

test('自定义磁盘缺失时不偷偷创建空库；可重新选定旧副本', async () => {
  const original = keep(await location.open())
  seed(original)
  const before = manifest(original)
  await location.savePath(join(root, 'missing-drive', 'library'))
  await expect(new StorageLocation(profile, installation).open()).rejects.toThrow(
    '没有创建空文档库',
  )
  expect(await readdir(root)).not.toContain('missing-drive')
  const path = await target()
  await expect(location.recoverLocation(path)).rejects.toThrow('不会以空库替换')
  original.close()
  opened.pop()
  const recovered = keep(await location.recoverLocation(location.defaultPath))
  expect(manifest(recovered)).toEqual(before)
})

test('旧库自动迁移失败时继续旧库，并说明未切换', async () => {
  const legacy = new Library(join(profile, 'library'))
  seed(legacy)
  const before = manifest(legacy)
  legacy.close()
  await mkdir(location.defaultPath, { recursive: true })
  await writeFile(join(location.defaultPath, '占用.txt'), '保留')
  const library = keep(await location.open())
  expect(library.root).toBe(join(profile, 'library'))
  expect(manifest(library)).toEqual(before)
  expect(location.info(library).notice).toContain('继续使用旧文档库')
  expect(await readFile(join(location.defaultPath, '占用.txt'), 'utf8')).toBe('保留')
})
