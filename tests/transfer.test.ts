import { afterEach, beforeEach, expect, test } from 'vitest'
import {
  mkdtempSync,
  readFileSync,
  realpathSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { DatabaseSync } from 'node:sqlite'
import { ZipFile } from 'yazl'
import { zipSync, unzipSync } from 'fflate'
import { Library } from '../src/main/library'
import { TransferService } from '../src/main/transfer-service'
import { validateManifest } from '../src/main/backup-format'
import { wordFixture, sheetFixture } from './fixtures/office'

let root: string, library: Library, transfer: TransferService
beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'local-docs-transfer-test-'))
  library = new Library(join(root, 'data', 'library'))
  transfer = new TransferService(library, join(root, 'data'))
})
afterEach(async () => {
  await transfer.discardPreview()
  library.close()
  const target = realpathSync(root)
  if (!target.startsWith(join(realpathSync(tmpdir()), 'local-docs-transfer-test-')))
    throw new Error('Unsafe test cleanup')
  rmSync(target, { recursive: true, force: true })
})

function seed() {
  const work = library.createCategory('工作')
  const child = library.createCategory('设计', work.id)
  library.createCategory('空分类', child.id)
  const doc = library.createMarkdown('说明', child.id)
  const first = library.saveMarkdown(
    doc.id,
    '# 初稿\n\n| 项目 | 值 |\n| --- | --- |\n| 测试 | 本地 |',
    doc.revision,
  )
  library.saveMarkdown(doc.id, '# 定稿\n离线迁移验收', first.revision)
  library.updateDocument(doc.id, { favorite: true, tags: ['验收'], notes: '保留备注' })
  library.openDocument(doc.id)
  const trash = library.createMarkdown('待回收', null)
  library.trashDocument(trash.id)
  library.setTheme('dark')
  return { work, child, doc, trash }
}
async function archive() {
  const path = join(root, 'sample.localdocs-backup')
  await transfer.backup(path)
  return path
}

test('完整备份跨文档库恢复：字节、层级、历史、回收站、收藏、标签与主题', async () => {
  const { doc } = seed()
  const original = library.backupManifest()
  const path = await archive()
  const second = new Library(join(root, 'new-device', 'library'))
  const receiver = new TransferService(second, join(root, 'new-device'))
  try {
    const newDoc = second.createMarkdown('新电脑原有资料', null)
    const preview = await receiver.preview(path)
    expect(preview).toMatchObject({ files: 1, trash: 1, categories: 3, versions: 4 })
    expect(second.snapshot().documents).toHaveLength(1)
    const result = await receiver.restore(preview.token)
    expect(second.backupManifest()).toEqual({ ...original, createdAt: expect.any(String) })
    expect(second.search('离线迁移')).toEqual([doc.id])
    expect(second.getTheme()).toBe('dark')
    const safety = await receiver.preview(result.safetyBackupPath!)
    expect(safety).toMatchObject({ files: 1, trash: 0 })
    await receiver.restore(safety.token)
    expect(second.snapshot().documents.map((d) => d.id)).toEqual([newDoc.id])
    expect(second.getTheme()).toBe('system')
  } finally {
    await receiver.discardPreview()
    second.close()
  }
}, 20_000)

test('Word、Excel 和两种 Markdown 扩展名备份/导出保持字节一致', async () => {
  const payloads = {
    '文档.docx': wordFixture(),
    '表格.xlsx': sheetFixture(),
    '笔记.md': Buffer.from('# 标题\r\n内容'),
    '长扩展名.markdown': Buffer.from('\ufeff| A | B |\r\n|---|---|\r\n| 1 | 2 |'),
  }
  for (const [name, bytes] of Object.entries(payloads)) {
    const path = join(root, name)
    writeFileSync(path, bytes)
    library.importFile(path, null)
  }
  const path = await archive()
  const prepared = await transfer.preview(path)
  await transfer.restore(prepared.token)
  const result = await transfer.export(root, { kind: 'all' })
  for (const [name, bytes] of Object.entries(payloads))
    expect(readFileSync(join(result.path, '未分类', name))).toEqual(Buffer.from(bytes))
})

test('分类导出含子分类和空分类，排除回收站，处理目录与文件的大小写重名', async () => {
  const { work, child } = seed()
  library.createCategory('说明.md', child.id)
  library.createMarkdown('说明', child.id)
  library.createMarkdown('Case', child.id)
  library.createMarkdown('case', child.id)
  library.createMarkdown('不在选中分类', null)
  const result = await transfer.export(root, { kind: 'category', id: work.id })
  const listing = JSON.parse(readFileSync(join(result.path, '导出清单.json'), 'utf8'))
  expect(result.files).toBe(4)
  expect(listing.files.map((f: { path: string }) => f.path)).toEqual(
    expect.arrayContaining(['工作/设计/说明（2）.md', '工作/设计/说明（3）.md']),
  )
  expect(new Set(listing.files.map((f: { path: string }) => f.path.toLowerCase())).size).toBe(4)
  expect(readdirSync(join(result.path, '工作', '设计'))).toContain('空分类')
  expect(library.snapshot().documents).toHaveLength(6)
})

test('备份不会覆盖已有文件，禁止输出到数据目录或其链接', async () => {
  seed()
  const path = await archive()
  const original = readFileSync(path)
  await expect(transfer.backup(path)).rejects.toThrow('同名')
  expect(readFileSync(path)).toEqual(original)
  await expect(transfer.backup(join(root, 'data', 'oops.localdocs-backup'))).rejects.toThrow('之外')
  await expect(transfer.export(join(root, 'data'), { kind: 'all' })).rejects.toThrow('之外')
})

test('损坏、缺失、多余对象、未来版本和路径穿越备份均不修改当前资料', async () => {
  seed()
  const before = library.snapshot()
  const path = await archive()
  const good = unzipSync(readFileSync(path))
  const objectName = Object.keys(good).find((k) => k.startsWith('objects/'))!
  const variants: Record<string, Uint8Array>[] = [
    { ...good, [objectName]: Buffer.from('changed') },
    Object.fromEntries(Object.entries(good).filter(([k]) => k !== objectName)),
    {
      ...good,
      'objects/aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa':
        Buffer.from('extra'),
    },
    {
      ...good,
      'manifest.json': Buffer.from(
        JSON.stringify({
          ...JSON.parse(Buffer.from(good['manifest.json']).toString()),
          version: 999,
        }),
      ),
    },
    { ...good, '../escaped.txt': Buffer.from('no') },
  ]
  for (const [i, contents] of variants.entries()) {
    const invalid = join(root, `bad-${i}.localdocs-backup`)
    writeFileSync(invalid, zipSync(contents))
    await expect(transfer.preview(invalid)).rejects.toThrow()
    expect(library.snapshot()).toEqual(before)
    expect(readdirSync(join(root, 'data', 'transfer-staging'))).toEqual([])
  }
})

test('清单拒绝循环分类、孤立历史、非法名称和错误大小', () => {
  seed()
  const m = library.backupManifest()
  for (const mutate of [
    (v: typeof m) => {
      v.categories[0].parentId = v.categories[1].id
    },
    (v: typeof m) => {
      v.versions[0].documentId = '00000000-0000-0000-0000-000000000000'
    },
    (v: typeof m) => {
      v.documents[0].name = '../oops.md'
    },
    (v: typeof m) => {
      v.documents[0].size++
    },
  ]) {
    const invalid = structuredClone(m)
    mutate(invalid)
    expect(() => validateManifest(invalid)).toThrow()
  }
})

test('清单仍是合法 JSON 但 ZIP 校验值不符时拒绝恢复', async () => {
  seed()
  const path = await archive()
  const entries = unzipSync(readFileSync(path))
  const bytes = Buffer.from(zipSync(entries, { level: 0 }))
  const index = bytes.indexOf(Buffer.from('"theme":"dark"'))
  expect(index).toBeGreaterThan(0)
  // Change a valid ISO timestamp digit without changing JSON length or the ZIP CRC.
  const date = bytes.indexOf(Buffer.from('"createdAt":"2026'))
  expect(date).toBeGreaterThan(0)
  bytes[date + 16] = '5'.charCodeAt(0)
  const corrupt = join(root, 'metadata-crc.localdocs-backup')
  writeFileSync(corrupt, bytes)
  const before = library.snapshot()
  await expect(transfer.preview(corrupt)).rejects.toThrow('校验')
  expect(library.snapshot()).toEqual(before)
})

test('恢复事务失败会回滚全部数据，自动备份可用且重启仍保留原库', async () => {
  seed()
  const path = await archive()
  library.createMarkdown('备份之后新增', null)
  library.setTheme('light')
  const before = library.snapshot()
  const preview = await transfer.preview(path)
  const db = new DatabaseSync(join(library.root, 'library.sqlite'))
  db.exec(
    "CREATE TRIGGER injected_failure BEFORE INSERT ON documents BEGIN SELECT RAISE(ABORT, 'test injected failure'); END;",
  )
  await expect(transfer.restore(preview.token)).rejects.toThrow('test injected failure')
  db.exec('DROP TRIGGER injected_failure')
  db.close()
  expect(library.snapshot()).toEqual(before)
  expect(readdirSync(join(root, 'data', 'recovery-backups'))).toHaveLength(1)
  library.close()
  library = new Library(join(root, 'data', 'library'))
  transfer = new TransferService(library, join(root, 'data'))
  expect(library.snapshot()).toEqual(before)
  expect(library.getTheme()).toBe('light')
})

test('恢复前自动备份失败不改当前库，取消操作清理输出且预览令牌不可复用', async () => {
  seed()
  const path = await archive()
  const preview = await transfer.preview(path)
  const before = library.snapshot()
  writeFileSync(join(root, 'data', 'recovery-backups'), 'blocking file')
  await expect(transfer.restore(preview.token)).rejects.toThrow()
  expect(library.snapshot()).toEqual(before)
  await expect(transfer.restore(preview.token)).rejects.toThrow('失效')
  const task = transfer.backup(join(root, 'cancelled.localdocs-backup'))
  transfer.cancel()
  await expect(task).rejects.toThrow('取消')
  expect(
    readdirSync(root).some(
      (name) => name.endsWith('.partial') || name === 'cancelled.localdocs-backup',
    ),
  ).toBe(false)
  expect(transfer.getStatus()).toBeNull()
})

test('恢复前重新校验暂存内容，篡改后拒绝替换当前资料', async () => {
  seed()
  const path = await archive()
  const preview = await transfer.preview(path)
  const before = library.snapshot()
  const stage = join(
    root,
    'data',
    'transfer-staging',
    readdirSync(join(root, 'data', 'transfer-staging'))[0],
    'objects',
  )
  const name = readdirSync(stage).find((name) => readFileSync(join(stage, name)).length > 0)!
  const bytes = readFileSync(join(stage, name))
  bytes[0] ^= 1
  writeFileSync(join(stage, name), bytes)
  await expect(transfer.restore(preview.token)).rejects.toThrow('校验')
  expect(library.snapshot()).toEqual(before)
})

test('备份源内容损坏会终止写入并清理临时文件', async () => {
  const { doc } = seed()
  const path = library.documentPath(doc.id)
  const bytes = readFileSync(path)
  bytes[0] ^= 1
  writeFileSync(path, bytes)
  await expect(transfer.backup(join(root, 'bad-source.localdocs-backup'))).rejects.toThrow('校验')
  expect(
    readdirSync(root).some(
      (name) => name.endsWith('.partial') || name === 'bad-source.localdocs-backup',
    ),
  ).toBe(false)
})

test('重复 ZIP 条目被拒绝，进行中的流式备份可以取消', async () => {
  seed()
  const zip = new ZipFile()
  const chunks: Buffer[] = []
  const packed = new Promise<Buffer>((resolve, reject) => {
    zip.outputStream.on('data', (data) => chunks.push(data))
    zip.outputStream.on('end', () => resolve(Buffer.concat(chunks)))
    zip.on('error', reject)
  })
  const manifest = Buffer.from(JSON.stringify(library.backupManifest()))
  zip.addBuffer(manifest, 'manifest.json')
  zip.addBuffer(manifest, 'manifest.json')
  zip.end()
  const duplicate = join(root, 'duplicate.localdocs-backup')
  writeFileSync(duplicate, await packed)
  await expect(transfer.preview(duplicate)).rejects.toThrow('重复')
  const large = join(root, 'large.docx')
  writeFileSync(large, Buffer.alloc(32 * 1024 * 1024, 42))
  library.importFile(large, null)
  const result = transfer.backup(join(root, 'cancel-midstream.localdocs-backup'))
  const cancel = setInterval(() => {
    if (transfer.getStatus()?.phase === '正在备份文件和历史版本') transfer.cancel()
  }, 2)
  try {
    await expect(result).rejects.toThrow('取消')
  } finally {
    clearInterval(cancel)
  }
  expect(
    readdirSync(root).some(
      (name) => name.endsWith('.partial') || name === 'cancel-midstream.localdocs-backup',
    ),
  ).toBe(false)
})
