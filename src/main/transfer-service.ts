import { createHash, randomUUID } from 'node:crypto'
import { createReadStream, createWriteStream } from 'node:fs'
import {
  access,
  lstat,
  mkdir,
  mkdtemp,
  readFile,
  realpath,
  rename,
  rm,
  unlink,
  writeFile,
} from 'node:fs/promises'
import { basename, dirname, extname, isAbsolute, join, relative, resolve, sep } from 'node:path'
import { Readable, Transform } from 'node:stream'
import { pipeline } from 'node:stream/promises'
import { setImmediate } from 'node:timers/promises'
import { crc32 } from 'node:zlib'
import { ZipFile as ZipWriter } from 'yazl'
import { openPromise } from 'yauzl'
import { Library } from './library'
import {
  MAX_BACKUP_BYTES,
  MAX_ENTRIES,
  MAX_MANIFEST_BYTES,
  MAX_OBJECT_BYTES,
  validateManifest,
  type BackupManifest,
} from './backup-format'
import type { BackupPreview, ExportScope, OperationStatus, TransferResult } from '../shared/types'

const stamp = () =>
  new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').replace('Z', '')
const inside = (parent: string, child: string) => {
  const rel = relative(resolve(parent), resolve(child))
  return rel === '' || (!rel.startsWith(`..${sep}`) && rel !== '..' && !isAbsolute(rel))
}

function verifier(
  expectedHash: string | null,
  expectedSize: number,
  maxSize = MAX_OBJECT_BYTES,
  expectedCrc?: number,
): Transform {
  const hash = createHash('sha256')
  let size = 0
  let crc = 0
  return new Transform({
    transform(chunk: Buffer, _encoding, callback) {
      size += chunk.length
      if (size > maxSize || size > expectedSize) {
        callback(new Error('备份内容超过允许大小。'))
        return
      }
      hash.update(chunk)
      if (expectedCrc !== undefined) crc = crc32(chunk, crc)
      callback(null, chunk)
    },
    flush(callback) {
      if (
        size !== expectedSize ||
        (expectedHash && hash.digest('hex') !== expectedHash) ||
        (expectedCrc !== undefined && crc !== expectedCrc)
      )
        callback(new Error('备份内容校验失败，文件可能已损坏。'))
      else callback()
    },
  })
}

export class TransferService {
  private status: OperationStatus | null = null
  private controller: AbortController | null = null
  private prepared: { token: string; directory: string; manifest: BackupManifest } | null = null
  private readonly stagingRoot: string
  lastResultPath: string | null = null
  constructor(
    private readonly library: Library,
    private readonly userData: string,
  ) {
    this.stagingRoot = join(userData, 'transfer-staging')
  }
  get active(): boolean {
    return this.status !== null
  }
  getStatus(): OperationStatus | null {
    return this.status ? { ...this.status } : null
  }
  cancel(): void {
    if (this.status?.cancellable) this.controller?.abort()
  }
  private progress(phase: string, completed = 0, total = 0, cancellable = true): void {
    if (this.status) this.status = { ...this.status, phase, completed, total, cancellable }
  }
  private check(): void {
    this.controller?.signal.throwIfAborted()
  }
  private async run<T>(kind: OperationStatus['kind'], action: () => Promise<T>): Promise<T> {
    if (this.active) throw new Error('请等待当前任务结束。')
    this.status = { kind, phase: '准备中', completed: 0, total: 0, cancellable: true }
    this.controller = new AbortController()
    try {
      return await action()
    } catch (error) {
      if (this.controller.signal.aborted) throw new Error('已取消操作，当前文档库未被替换。')
      if ((error as NodeJS.ErrnoException).code === 'EEXIST')
        throw new Error('目标位置已有同名文件，请选择其他名称。')
      throw error
    } finally {
      this.status = null
      this.controller = null
    }
  }
  private async stage(): Promise<string> {
    await mkdir(this.stagingRoot, { recursive: true })
    return mkdtemp(join(this.stagingRoot, 'transfer-'))
  }
  private async cleanStage(path: string): Promise<void> {
    const target = resolve(path)
    if (
      dirname(target) !== resolve(this.stagingRoot) ||
      !/^transfer-[a-zA-Z0-9]+$/.test(basename(target))
    )
      throw new Error('无效的临时目录。')
    await rm(target, { recursive: true, force: true })
  }
  private async outsideStorage(path: string): Promise<void> {
    if (!isAbsolute(path)) throw new Error('请选择绝对路径。')
    const parent = await realpath(dirname(path))
    if (inside(await realpath(this.userData), join(parent, basename(path))))
      throw new Error('请选择文档库数据目录之外的位置。')
  }

  private async writeArchive(destination: string, manifest: BackupManifest): Promise<void> {
    this.check()
    const temp = join(dirname(destination), `.local-docs-${randomUUID()}.partial`)
    const bytes = Buffer.from(JSON.stringify(manifest))
    if (bytes.length > MAX_MANIFEST_BYTES) throw new Error('文档库清单过大，当前版本无法备份。')
    const zip = new ZipWriter()
    const archiveController = new AbortController()
    const signal = AbortSignal.any([this.controller!.signal, archiveController.signal])
    const inputs: Promise<unknown>[] = []
    const output = zip.outputStream as Readable
    zip.on('error', (error: Error) => output.destroy(error))
    // Attach the destination immediately so every stream failure is observed.
    const writing = pipeline(zip.outputStream, createWriteStream(temp, { flags: 'wx' }), {
      signal,
    })
    void writing.catch(() => {})
    zip.addBuffer(bytes, 'manifest.json')
    let completed = 0
    for (const object of manifest.objects) {
      zip.addReadStreamLazy(
        `objects/${object.hash}`,
        { size: object.size, compress: false },
        (callback) => {
          if (signal.aborted) {
            callback(new Error('备份操作已停止。'), Readable.from([]))
            return
          }
          const checked = verifier(object.hash, object.size)
          checked.on('error', (error) => zip.emit('error', error))
          checked.on('end', () => {
            if (this.status) this.status.completed = ++completed
          })
          inputs.push(
            pipeline(createReadStream(this.library.backupObjectPath(object.hash)), checked, {
              signal,
            }).catch((error) => checked.destroy(error)),
          )
          callback(null, checked)
        },
      )
    }
    zip.end()
    let published = false
    let destinationOwned = false
    try {
      await writing
      this.check()
      // Exclusive output creation never replaces an original or an older backup.
      const destinationStream = createWriteStream(destination, { flags: 'wx', flush: true })
      destinationStream.once('open', () => {
        destinationOwned = true
      })
      await pipeline(createReadStream(temp), destinationStream, { signal })
      this.check()
      published = true
    } finally {
      archiveController.abort()
      output.destroy()
      await Promise.allSettled(inputs)
      await unlink(temp).catch(() => {})
      if (destinationOwned && !published) await unlink(destination).catch(() => {})
    }
  }

  async backup(destination: string): Promise<TransferResult> {
    return this.run('backup', async () => {
      await this.outsideStorage(destination)
      const manifest = this.library.backupManifest()
      this.progress('正在备份文件和历史版本', 0, manifest.objects.length)
      await this.writeArchive(destination, manifest)
      this.lastResultPath = destination
      return { path: destination, files: manifest.documents.length }
    })
  }

  async discardPreview(): Promise<void> {
    const prepared = this.prepared
    this.prepared = null
    if (prepared) await this.cleanStage(prepared.directory)
  }
  async preview(path: string): Promise<BackupPreview> {
    return this.run('preview', async () => {
      await this.discardPreview()
      const info = await lstat(path)
      if (
        !info.isFile() ||
        info.isSymbolicLink() ||
        info.size > MAX_BACKUP_BYTES + MAX_MANIFEST_BYTES + 32 * 1024 * 1024
      )
        throw new Error('请选择有效的文档库备份文件（最多 10 GB 内容）。')
      const directory = await this.stage()
      let keep = false
      try {
        const zip = await openPromise(path, {
          lazyEntries: true,
          autoClose: false,
          strictFileNames: true,
          validateEntrySizes: true,
        })
        const names = new Set<string>()
        let total = 0
        try {
          if (zip.entryCount > MAX_ENTRIES + 1) throw new Error('备份内文件数量超出限制。')
          this.progress('正在校验备份', 0, zip.entryCount)
          await mkdir(join(directory, 'objects'))
          for await (const entry of zip.eachEntry()) {
            this.check()
            const name = entry.fileName
            const manifestEntry = name === 'manifest.json'
            if (
              (!manifestEntry && !/^objects\/[a-f0-9]{64}$/.test(name)) ||
              names.has(name) ||
              entry.isEncrypted() ||
              ((entry.externalFileAttributes >>> 16) & 0xf000) === 0xa000
            )
              throw new Error('备份包含重复、加密或不允许的文件。')
            names.add(name)
            const max = manifestEntry ? MAX_MANIFEST_BYTES : MAX_OBJECT_BYTES
            total += entry.uncompressedSize
            if (entry.uncompressedSize > max || total > MAX_BACKUP_BYTES + MAX_MANIFEST_BYTES)
              throw new Error('备份内容超过允许大小。')
            const input = await zip.openReadStreamPromise(entry)
            await pipeline(
              input,
              verifier(
                manifestEntry ? null : name.slice(8),
                entry.uncompressedSize,
                max,
                entry.crc32,
              ),
              createWriteStream(join(directory, name), { flags: 'wx' }),
              { signal: this.controller!.signal },
            )
            this.progress('正在校验备份', names.size, zip.entryCount)
          }
        } finally {
          zip.close()
        }
        if (!names.has('manifest.json')) throw new Error('备份缺少清单。')
        const manifest = validateManifest(
          JSON.parse(await readFile(join(directory, 'manifest.json'), 'utf8')),
        )
        if (names.size !== manifest.objects.length + 1) throw new Error('备份内容与清单不一致。')
        for (const object of manifest.objects) {
          this.check()
          if (
            !names.has(`objects/${object.hash}`) ||
            (await lstat(join(directory, 'objects', object.hash))).size !== object.size
          )
            throw new Error('备份缺少文件或大小不一致。')
        }
        // Validate historical Markdown as well as current content before any replacement.
        const markdownIds = new Set(
          manifest.documents
            .filter((d) => ['.md', '.markdown'].includes(d.extension))
            .map((d) => d.id),
        )
        const textHashes = new Set([
          ...manifest.documents.filter((d) => markdownIds.has(d.id)).map((d) => d.blobHash),
          ...manifest.versions.filter((v) => markdownIds.has(v.documentId)).map((v) => v.blobHash),
        ])
        for (const hash of textHashes) {
          this.check()
          new TextDecoder('utf-8', { fatal: true }).decode(
            await readFile(join(directory, 'objects', hash)),
          )
        }
        this.check()
        const token = randomUUID()
        this.prepared = { token, directory, manifest }
        keep = true
        return {
          token,
          createdAt: manifest.createdAt,
          files: manifest.documents.filter((d) => !d.deletedAt).length,
          trash: manifest.documents.filter((d) => d.deletedAt).length,
          categories: manifest.categories.length,
          versions: manifest.versions.length,
          bytes: manifest.objects.reduce((sum, o) => sum + o.size, 0),
        }
      } finally {
        if (!keep) await this.cleanStage(directory)
      }
    })
  }

  async restore(token: string): Promise<TransferResult> {
    return this.run('restore', async () => {
      const prepared = this.prepared
      if (!prepared || typeof token !== 'string' || prepared.token !== token)
        throw new Error('备份预览已失效，请重新选择备份。')
      let safetyBackupPath: string | undefined
      try {
        const current = this.library.backupManifest()
        const recoveryRoot = join(this.userData, 'recovery-backups')
        await mkdir(recoveryRoot, { recursive: true })
        const destination = join(
          recoveryRoot,
          `恢复前-${stamp()}-${randomUUID().slice(0, 8)}.localdocs-backup`,
        )
        this.progress('正在保留恢复前的完整备份', 0, current.objects.length)
        await this.writeArchive(destination, current)
        safetyBackupPath = destination
        this.progress('正在恢复文件和历史版本', 0, prepared.manifest.objects.length)
        for (const [index, object] of prepared.manifest.objects.entries()) {
          this.check()
          const file = join(prepared.directory, 'objects', object.hash)
          const info = await lstat(file)
          if (!info.isFile() || info.isSymbolicLink() || info.size !== object.size)
            throw new Error('备份临时文件已变化。')
          this.library.installBackupObject(object.hash, await readFile(file))
          this.progress('正在恢复文件和历史版本', index + 1, prepared.manifest.objects.length)
          await setImmediate()
        }
        this.check()
        this.progress('正在完成恢复，请稍候', 0, 0, false)
        this.library.applyBackup(prepared.manifest)
        this.lastResultPath = safetyBackupPath
        return {
          path: this.library.root,
          safetyBackupPath,
          files: prepared.manifest.documents.length,
        }
      } catch (error) {
        if (safetyBackupPath && !this.controller?.signal.aborted)
          throw new Error(
            `${error instanceof Error ? error.message : String(error)}\n恢复前备份：${safetyBackupPath}`,
          )
        throw error
      } finally {
        await this.discardPreview().catch(() => {})
      }
    })
  }

  async export(directory: string, scope: ExportScope): Promise<TransferResult> {
    return this.run('export', async () => {
      if (
        !scope ||
        !['all', 'category'].includes(scope.kind) ||
        (scope.kind === 'category' && typeof scope.id !== 'string')
      )
        throw new Error('无效的导出范围。')
      const manifest = this.library.backupManifest()
      const selected =
        scope.kind === 'category' ? manifest.categories.find((c) => c.id === scope.id) : null
      if (scope.kind === 'category' && !selected) throw new Error('分类不存在。')
      const base = await realpath(directory)
      const final = join(
        base,
        `${selected ? selected.name.slice(0, 48) : '我的文档库'}-${stamp()}-${randomUUID().slice(0, 8)}`,
      )
      await this.outsideStorage(final)
      const staging = await mkdtemp(join(base, '.local-docs-export-'))
      let complete = false
      try {
        const paths = new Map<string, string>()
        const used = new Map<string, Set<string>>([['', new Set(['导出清单.json'])]])
        const allocate = (parent: string, name: string, file: boolean): string => {
          const taken = used.get(parent) ?? new Set<string>()
          used.set(parent, taken)
          const ext = file ? extname(name) : ''
          const stem = (ext ? name.slice(0, -ext.length) : name).slice(0, 100)
          let candidate = `${stem}${ext}`
          let counter = 1
          while (taken.has(candidate.toLowerCase())) candidate = `${stem}（${++counter}）${ext}`
          taken.add(candidate.toLowerCase())
          const result = parent ? `${parent}/${candidate}` : candidate
          if (join(staging, result).length > 240)
            throw new Error('分类层级过深或导出路径过长，请选择较浅的目标目录，或分分类导出。')
          return result
        }
        for (const c of manifest.categories) {
          if (selected && c.id !== selected.id && (!c.parentId || !paths.has(c.parentId))) continue
          const parent = !c.parentId || c.id === selected?.id ? '' : paths.get(c.parentId)!
          const path = allocate(parent, c.name, false)
          paths.set(c.id, path)
          await mkdir(join(staging, path), { recursive: true })
        }
        const docs = manifest.documents.filter(
          (d) => !d.deletedAt && (!selected || (d.categoryId && paths.has(d.categoryId))),
        )
        let uncategorized = ''
        if (!selected && docs.some((d) => !d.categoryId)) {
          uncategorized = allocate('', '未分类', false)
          await mkdir(join(staging, uncategorized))
        }
        const files: {
          id: string
          name: string
          path: string
          tags: string[]
          notes: string
          favorite: boolean
          sha256: string
        }[] = []
        this.progress('正在按分类导出', 0, docs.length)
        for (const [index, doc] of docs.entries()) {
          this.check()
          const path = allocate(
            doc.categoryId ? paths.get(doc.categoryId)! : uncategorized,
            doc.name,
            true,
          )
          await pipeline(
            createReadStream(this.library.backupObjectPath(doc.blobHash)),
            verifier(doc.blobHash, doc.size),
            createWriteStream(join(staging, path), { flags: 'wx' }),
            { signal: this.controller!.signal },
          )
          this.library.exportAttachments(doc.id, dirname(join(staging, path)))
          files.push({
            id: doc.id,
            name: doc.name,
            path,
            tags: doc.tags,
            notes: doc.notes,
            favorite: doc.favorite,
            sha256: doc.blobHash,
          })
          this.progress('正在按分类导出', index + 1, docs.length)
        }
        await writeFile(
          join(staging, '导出清单.json'),
          JSON.stringify(
            {
              format: 'local-docs-export',
              version: 1,
              createdAt: new Date().toISOString(),
              description: '按分类导出的当前文件；不含回收站和历史版本。完整恢复请使用文档库备份。',
              categories: manifest.categories
                .filter((c) => paths.has(c.id))
                .map((c) => ({ ...c, path: paths.get(c.id) })),
              files,
            },
            null,
            2,
          ),
          { flag: 'wx' },
        )
        this.check()
        try {
          await access(final)
          throw new Error('导出目标已存在。')
        } catch (error) {
          if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
        }
        await rename(staging, final)
        complete = true
        this.lastResultPath = final
        return { path: final, files: files.length }
      } finally {
        if (
          !complete &&
          dirname(resolve(staging)) === base &&
          basename(staging).startsWith('.local-docs-export-')
        )
          await rm(staging, { recursive: true, force: true })
      }
    })
  }
}
