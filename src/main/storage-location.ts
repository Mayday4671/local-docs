import { randomUUID, createHash } from 'node:crypto'
import { createReadStream, createWriteStream } from 'node:fs'
import {
  lstat,
  mkdir,
  mkdtemp,
  open,
  readFile,
  readdir,
  realpath,
  rename,
  rm,
  rmdir,
  statfs,
  unlink,
} from 'node:fs/promises'
import { basename, dirname, isAbsolute, join, relative, resolve, sep } from 'node:path'
import { Transform } from 'node:stream'
import { pipeline } from 'node:stream/promises'
import { Library } from './library'
import type { OperationStatus, StorageInfo, StorageMovePreview } from '../shared/types'

const same = (a: string, b: string) => resolve(a).toLowerCase() === resolve(b).toLowerCase()
const inside = (parent: string, child: string) => {
  const rel = relative(resolve(parent).toLowerCase(), resolve(child).toLowerCase())
  return rel === '' || (rel !== '..' && !rel.startsWith(`..${sep}`) && !isAbsolute(rel))
}
async function exists(path: string) {
  try {
    await lstat(path)
    return true
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === 'ENOENT') return false
    throw e
  }
}
async function writable(path: string) {
  const probe = join(path, `.local-docs-write-${randomUUID()}`)
  const handle = await open(probe, 'wx')
  try {
    await handle.sync()
  } finally {
    await handle.close()
    await unlink(probe)
  }
}

export class StorageLocation {
  readonly defaultPath: string
  readonly configPath: string
  private preview: StorageMovePreview | null = null
  private status: OperationStatus | null = null
  private abort: AbortController | null = null
  private notice = ''
  constructor(
    readonly profile: string,
    readonly installation: string,
    defaultOverride?: string,
  ) {
    this.defaultPath = resolve(defaultOverride || join(installation, 'data', 'library'))
    this.configPath = join(profile, 'storage-location.json')
  }
  get active() {
    return this.status !== null
  }
  getStatus() {
    return this.status && { ...this.status }
  }
  cancel() {
    if (this.status?.cancellable) this.abort?.abort()
  }
  info(library: Library): StorageInfo {
    return { path: library.root, defaultPath: this.defaultPath, notice: this.notice }
  }
  private check() {
    this.abort?.signal.throwIfAborted()
  }
  private progress(phase: string, completed: number, total: number, cancellable = true) {
    this.status = { kind: 'storage-move', phase, completed, total, cancellable }
  }

  /** Small bootstrap preference stays outside the library so moving it never loses its address. */
  async savePath(path: string) {
    const temp = `${this.configPath}.${randomUUID()}.tmp`
    await mkdir(this.profile, { recursive: true })
    try {
      const handle = await open(temp, 'wx')
      try {
        await handle.writeFile(JSON.stringify({ version: 1, path: resolve(path) }))
        await handle.sync()
      } finally {
        await handle.close()
      }
      await rename(temp, this.configPath)
    } finally {
      await unlink(temp).catch(() => {})
    }
  }

  private async openExisting(path: string): Promise<Library> {
    if (!(await exists(join(path, 'library.sqlite'))))
      throw new Error(
        `文档库位置不可用：${path}\n请连接原磁盘后重试。为保护原资料，没有创建空文档库。`,
      )
    const info = await lstat(path),
      db = await lstat(join(path, 'library.sqlite'))
    if (!info.isDirectory() || info.isSymbolicLink() || !db.isFile() || db.isSymbolicLink())
      throw new Error('文档库路径不是普通目录或数据库文件。')
    await writable(path)
    const library = new Library(path)
    try {
      library.checkIntegrity()
      // Opening a valid database must not inherit backup size limits or block
      // access to the whole library because one externally removed object is missing.
      // Full manifest and object verification is performed before migration.
      return library
    } catch (e) {
      library.close()
      throw e
    }
  }

  async recoverLocation(path: string): Promise<Library> {
    const info = await lstat(path)
    if (!info.isDirectory() || info.isSymbolicLink()) throw new Error('请选择普通文档库文件夹。')
    path = await realpath(path)
    if (inside(this.installation, path) && !inside(join(this.installation, 'data'), path))
      throw new Error('安装目录内请使用 data 下的文件夹。')
    const existing = await exists(join(path, 'library.sqlite'))
    if (!existing) {
      if (
        (await exists(this.configPath)) ||
        (await exists(join(this.profile, 'library', 'library.sqlite')))
      )
        throw new Error('请连接原磁盘或选择已有文档库文件夹，不会以空库替换原资料。')
      if ((await readdir(path)).length) throw new Error('请选择已有文档库或一个空文件夹。')
      await writable(path)
    }
    const library = existing ? await this.openExisting(path) : new Library(path)
    try {
      await this.savePath(path)
      return library
    } catch (e) {
      library.close()
      throw e
    }
  }

  async open(): Promise<Library> {
    if (await exists(this.configPath)) {
      const config = JSON.parse(await readFile(this.configPath, 'utf8'))
      if (config.version !== 1 || typeof config.path !== 'string' || !isAbsolute(config.path))
        throw new Error('文档库位置配置无效，未自动创建或替换资料。')
      return this.openExisting(config.path)
    }
    if (await exists(join(this.defaultPath, 'library.sqlite'))) {
      const library = await this.openExisting(this.defaultPath)
      try {
        await this.savePath(library.root)
        return library
      } catch (e) {
        library.close()
        throw e
      }
    }
    const legacy = join(this.profile, 'library')
    if (!same(legacy, this.defaultPath) && (await exists(join(legacy, 'library.sqlite')))) {
      const original = await this.openExisting(legacy)
      try {
        await mkdir(this.defaultPath, { recursive: true })
        const plan = await this.prepare(original, this.defaultPath)
        const replacement = await this.move(original, plan.token)
        original.close()
        this.notice = `已将旧版资料迁移到安装目录，旧副本保留在：${legacy}`
        return replacement
      } catch (error) {
        // A read-only installation or full disk must not make an existing library disappear.
        this.notice = `未能迁移到默认位置，继续使用旧文档库。可在设置中选择其他位置。${String(error)}`
        try {
          await this.savePath(legacy)
        } catch (e) {
          original.close()
          throw e
        }
        return original
      }
    }
    await mkdir(this.defaultPath, { recursive: true })
    if ((await readdir(this.defaultPath)).length)
      throw new Error('默认文档库目录非空，未覆盖其中内容。')
    await writable(this.defaultPath)
    const library = new Library(this.defaultPath)
    try {
      await this.savePath(library.root)
      return library
    } catch (e) {
      library.close()
      throw e
    }
  }

  private async validateTarget(library: Library, path: string): Promise<string> {
    if (!isAbsolute(path)) throw new Error('请选择绝对路径。')
    const info = await lstat(path)
    if (!info.isDirectory() || info.isSymbolicLink())
      throw new Error('请选择普通的空文件夹，不支持目录链接。')
    const target = await realpath(path),
      source = await realpath(library.root)
    if (inside(source, target) || inside(target, source))
      throw new Error('新位置不能是当前文档库、其子目录或上级目录。')
    if (inside(this.installation, target) && !inside(join(this.installation, 'data'), target))
      throw new Error('安装目录内请使用 data 下的空文件夹，避免与程序文件混放。')
    if (inside(this.profile, target) && !same(target, this.defaultPath))
      throw new Error('请选择应用配置目录之外的位置。')
    if ((await readdir(target)).length)
      throw new Error('目标文件夹非空，请新建一个空文件夹。不会覆盖已有资料库或其他文件。')
    await writable(target)
    await writable(dirname(target))
    return target
  }

  async prepare(library: Library, path: string): Promise<StorageMovePreview> {
    this.preview = null
    const target = await this.validateTarget(library, path)
    const manifest = library.backupManifest()
    const bytes = manifest.objects.reduce((sum, object) => sum + object.size, 0)
    const disk = await statfs(target)
    const dbSize = (await lstat(join(library.root, 'library.sqlite'))).size
    if (disk.bavail * disk.bsize < bytes + dbSize * 2 + 16 * 1024 ** 2)
      throw new Error('目标磁盘可用空间不足，请换一个位置。')
    this.preview = {
      token: randomUUID(),
      source: library.root,
      target,
      files: manifest.documents.length,
      versions: manifest.versions.length,
      bytes,
    }
    return this.preview
  }
  discardPreview() {
    this.preview = null
  }

  async move(library: Library, token: string): Promise<Library> {
    if (this.active) throw new Error('请等待当前迁移完成。')
    const plan = this.preview
    if (!plan || plan.token !== token || !same(plan.source, library.root))
      throw new Error('存储位置预览已失效，请重新选择。')
    this.preview = null
    this.abort = new AbortController()
    this.progress('正在检查新位置', 0, 0)
    let stage = '',
      published = false,
      replacement: Library | null = null
    try {
      const target = await this.validateTarget(library, plan.target)
      const manifest = library.backupManifest()
      stage = await mkdtemp(join(dirname(target), '.local-docs-move-'))
      await mkdir(join(stage, 'objects'))
      this.progress('正在复制文档、附件与历史版本', 0, manifest.objects.length)
      for (const [i, object] of manifest.objects.entries()) {
        this.check()
        const source = library.backupObjectPath(object.hash)
        const info = await lstat(source)
        if (!info.isFile() || info.isSymbolicLink() || info.size !== object.size)
          throw new Error('源文件已变化或损坏，未切换位置。')
        const hash = createHash('sha256')
        let size = 0
        const verify = new Transform({
          transform(chunk: Buffer, _encoding, next) {
            size += chunk.length
            if (size > object.size) {
              next(new Error('文件大小校验失败。'))
              return
            }
            hash.update(chunk)
            next(null, chunk)
          },
          flush(next) {
            next(
              size === object.size && hash.digest('hex') === object.hash
                ? null
                : new Error('文件内容校验失败。'),
            )
          },
        })
        await pipeline(
          createReadStream(source),
          verify,
          createWriteStream(join(stage, 'objects', object.hash), { flags: 'wx', flush: true }),
          { signal: this.abort.signal },
        )
        this.progress('正在复制文档、附件与历史版本', i + 1, manifest.objects.length)
      }
      this.check()
      this.progress('正在校验数据库与全部记录', 0, 0)
      await library.copyDatabase(join(stage, 'library.sqlite'))
      replacement = new Library(stage)
      replacement.checkIntegrity()
      const { createdAt: _before, ...before } = manifest
      const { createdAt: _after, ...after } = replacement.backupManifest()
      if (JSON.stringify(before) !== JSON.stringify(after))
        throw new Error('迁移前后的资料记录不一致，未切换位置。')
      replacement.close()
      replacement = null
      const db = await open(join(stage, 'library.sqlite'), 'r+')
      try {
        await db.sync()
      } finally {
        await db.close()
      }
      this.check()
      // Recheck after copying; never replace a directory that another operation populated.
      await this.validateTarget(library, target)
      this.check()
      this.progress('正在切换到新文档库', 0, 0, false)
      await rmdir(target) // only succeeds while empty
      try {
        await rename(stage, target)
      } catch (e) {
        await mkdir(target).catch(() => {})
        throw e
      }
      published = true
      replacement = await this.openExisting(target)
      await this.savePath(target)
      this.notice = `已切换到新位置，旧副本保留在：${library.root}`
      return replacement
    } catch (e) {
      replacement?.close()
      if (this.abort.signal.aborted) throw new Error('已取消迁移，继续使用原文档库。')
      throw new Error(
        `${e instanceof Error ? e.message : String(e)}${published ? `\n未切换位置，已复制的副本保留在：${plan.target}` : ''}`,
      )
    } finally {
      if (stage && !published) {
        // Only this operation's uniquely named sibling staging directory can be removed.
        if (
          dirname(resolve(stage)) === dirname(resolve(plan.target)) &&
          /^\.local-docs-move-[\w-]+$/.test(basename(stage))
        )
          await rm(stage, { recursive: true, force: true }).catch(() => {})
      }
      this.status = null
      this.abort = null
    }
  }
}
