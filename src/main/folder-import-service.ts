import { randomUUID } from 'node:crypto'
import { lstat, open, readdir, realpath } from 'node:fs/promises'
import { basename, dirname, isAbsolute, join, relative, resolve, sep } from 'node:path'
import { setImmediate as yieldToEvents } from 'node:timers/promises'
import type { FolderImportPreview, FolderImportResult, OperationStatus } from '../shared/types'
import { Library } from './library'

type Entry = { path: string; size: number; mtime: number; ino: number; dev: number }
type Plan = {
  token: string
  root: string
  name: string
  folders: Entry[]
  files: Entry[]
  skipped: number
}
const limitFor = (path: string) => (/\.(md|markdown)$/i.test(path) ? 5 : 100) * 1024 * 1024
const key = (path: string) => (process.platform === 'win32' ? path.toLowerCase() : path)
function inside(root: string, path: string) {
  const rel = relative(key(root), key(path))
  return rel === '' || (!rel.startsWith(`..${sep}`) && rel !== '..' && !isAbsolute(rel))
}
const reason = (error: unknown) => {
  const code = (error as NodeJS.ErrnoException)?.code
  if (code === 'EACCES' || code === 'EPERM') return '没有读取权限'
  if (code === 'ENOENT') return '文件或目录已不存在，请重新扫描'
  return error instanceof Error ? error.message : '无法读取'
}

/** A preview token refers only to paths selected in the native folder chooser. */
export class FolderImportService {
  private plan: Plan | null = null
  private status: OperationStatus | null = null
  private cancelled = false
  constructor(
    private readonly library: Library,
    private readonly userData: string,
  ) {}
  get active() {
    return this.status !== null
  }
  getStatus(): OperationStatus | null {
    return this.status && { ...this.status }
  }
  cancel(): void {
    this.cancelled = true
  }
  discardPreview(): void {
    this.plan = null
  }
  private begin(kind: OperationStatus['kind'], phase: string, total = 0) {
    if (this.active) throw new Error('请等待当前导入操作完成。')
    this.cancelled = false
    this.status = { kind, phase, completed: 0, total, cancellable: true }
  }

  async preview(source: string): Promise<FolderImportPreview> {
    this.begin('folder-scan', '正在扫描文件夹')
    this.plan = null
    try {
      if (!isAbsolute(source)) throw new Error('请选择本地文件夹。')
      const rootInfo = await lstat(source)
      if (!rootInfo.isDirectory() || rootInfo.isSymbolicLink())
        throw new Error('请选择普通文件夹，暂不支持目录链接。')
      const root = await realpath(source)
      const dataRoot = await realpath(this.userData)
      if (inside(dataRoot, root)) throw new Error('不能将应用自身的存储目录作为导入来源。')
      const name = basename(root)
      if (!name) throw new Error('请选择一个资料文件夹，不要选择整个磁盘。')
      const plan: Plan = { token: randomUUID(), root, name, folders: [], files: [], skipped: 0 }
      const warnings: FolderImportPreview['warnings'] = []
      let visited = 0,
        bytes = 0
      const skip = (path: string, message: string) => {
        plan.skipped++
        if (warnings.length < 100) warnings.push({ name: path || name, reason: message })
      }
      const walk = async (path: string, depth: number): Promise<void> => {
        if (this.cancelled) throw new Error('已取消文件夹扫描。')
        if (++visited > 10000) throw new Error('目录超过 10000 个项目，请分批选择较小的文件夹。')
        this.status!.completed = visited
        const full = join(root, path)
        let info
        try {
          info = await lstat(full)
          if (info.isSymbolicLink()) {
            skip(path, '已跳过文件或目录链接')
            return
          }
          const actual = await realpath(full)
          if (!inside(root, actual) || key(actual) !== key(full) || inside(dataRoot, actual)) {
            skip(path, '已跳过应用存储目录或路径发生变化的项目')
            return
          }
        } catch (error) {
          if (!path) throw error
          skip(path, reason(error))
          return
        }
        const entry = { path, size: info.size, mtime: info.mtimeMs, ino: info.ino, dev: info.dev }
        if (info.isDirectory()) {
          if (depth > 32) {
            skip(path, '目录层级超过 32 层')
            return
          }
          if (plan.folders.length >= 2000) throw new Error('分类将超过 2000 个，请分批导入。')
          let children
          try {
            children = await readdir(full)
          } catch (error) {
            if (!path) throw error
            skip(path, reason(error))
            return
          }
          plan.folders.push(entry)
          for (const child of children.sort((a, b) => a.localeCompare(b, 'zh-CN')))
            await walk(join(path, child), depth + 1)
        } else if (info.isFile()) {
          if (info.size > limitFor(path)) {
            skip(path, /\.(md|markdown)$/i.test(path) ? 'Markdown 超过 5 MB' : '文件超过 100 MB')
            return
          }
          if (plan.files.length >= 5000) throw new Error('支持的文件超过 5000 份，请分批导入。')
          if (bytes + info.size > 2 * 1024 ** 3)
            throw new Error('本次导入内容超过 2 GB，请分批导入。')
          plan.files.push(entry)
          bytes += info.size
        } else skip(path, '已跳过非常规文件')
      }
      await walk('', 1)
      if (this.cancelled) throw new Error('已取消文件夹扫描。')
      if (!plan.folders.length) throw new Error('所选文件夹已变化，请重新选择。')
      this.plan = plan
      return {
        token: plan.token,
        rootName: name,
        sourcePath: root,
        files: plan.files.length,
        folders: plan.folders.length,
        bytes,
        skipped: plan.skipped,
        warnings,
        entries: [
          ...plan.folders.map((e) => ({ path: e.path || name, kind: 'folder' as const })),
          ...plan.files.map((e) => ({ path: e.path, kind: 'file' as const })),
        ].slice(0, 100),
      }
    } finally {
      this.status = null
    }
  }

  private async verify(plan: Plan, entry: Entry, directory: boolean) {
    const full = resolve(plan.root, entry.path)
    if (!inside(plan.root, full) || key(await realpath(full)) !== key(full))
      throw new Error('路径已变化，请重新扫描。')
    const info = await lstat(full)
    if (
      info.isSymbolicLink() ||
      (directory ? !info.isDirectory() : !info.isFile()) ||
      info.ino !== entry.ino ||
      info.dev !== entry.dev
    )
      throw new Error('文件或目录已被替换，请重新扫描。')
    if (!directory && (info.size !== entry.size || info.mtimeMs !== entry.mtime))
      throw new Error('文件在预览后已修改，请重新扫描。')
    return full
  }

  async import(
    token: string,
    parentId: string | null,
    duplicates: 'skip' | 'keep',
  ): Promise<FolderImportResult> {
    if (!this.plan || this.plan.token !== token)
      throw new Error('导入预览已失效，请重新选择文件夹。')
    if (duplicates !== 'skip' && duplicates !== 'keep') throw new Error('无效的重复文件处理方式。')
    const snapshot = this.library.snapshot()
    if (parentId !== null && !snapshot.categories.some((c) => c.id === parentId))
      throw new Error('目标分类不存在。')
    const plan = this.plan
    this.begin('folder-import', '正在导入文件夹', plan.folders.length + plan.files.length)
    this.plan = null
    const result: FolderImportResult = {
      imported: [],
      failures: [],
      skipped: plan.skipped,
      cancelled: false,
      rootCategoryId: null,
    }
    const categories = new Map<string, string>()
    const directoryEntries = new Map(plan.folders.map((entry) => [entry.path, entry]))
    try {
      await this.verify(plan, plan.folders[0], true)
      for (const entry of plan.folders) {
        await yieldToEvents()
        if (this.cancelled) break
        try {
          await this.verify(plan, entry, true)
          const parent = entry.path
            ? categories.get(dirname(entry.path) === '.' ? '' : dirname(entry.path))
            : parentId
          if (parent === undefined) throw new Error('上级目录未能导入')
          const name = entry.path ? basename(entry.path) : plan.name
          const existing = snapshot.categories.find((c) => c.parentId === parent && c.name === name)
          const category = existing ?? this.library.createCategory(name, parent)
          if (!existing) snapshot.categories.push(category)
          categories.set(entry.path, category.id)
          if (!entry.path) result.rootCategoryId = category.id
        } catch (error) {
          result.failures.push({ name: entry.path || plan.name, reason: reason(error) })
        }
        this.status!.completed++
      }
      for (const entry of plan.files) {
        await yieldToEvents()
        if (this.cancelled) break
        try {
          const parent = dirname(entry.path) === '.' ? '' : dirname(entry.path)
          const categoryId = categories.get(parent)
          if (!categoryId) throw new Error('上级目录未能导入')
          // Recheck ancestors in case a folder is replaced between the preview and file copy.
          let ancestor = parent
          while (true) {
            await this.verify(plan, directoryEntries.get(ancestor)!, true)
            if (!ancestor) break
            ancestor = dirname(ancestor) === '.' ? '' : dirname(ancestor)
          }
          const full = await this.verify(plan, entry, false)
          const handle = await open(full, 'r')
          let bytes: Buffer
          try {
            const info = await handle.stat()
            if (
              info.ino !== entry.ino ||
              info.dev !== entry.dev ||
              info.size !== entry.size ||
              info.mtimeMs !== entry.mtime
            )
              throw new Error('文件在预览后已变化，请重新扫描。')
            // Bounded reads also detect a file growing while it is being copied.
            const buffer = Buffer.alloc(entry.size + 1)
            let offset = 0
            while (offset < buffer.length) {
              if (this.cancelled) break
              const { bytesRead } = await handle.read(
                buffer,
                offset,
                Math.min(1024 * 1024, buffer.length - offset),
                offset,
              )
              if (!bytesRead) break
              offset += bytesRead
            }
            if (this.cancelled) break
            const after = await handle.stat()
            if (offset !== entry.size || after.size !== entry.size || after.mtimeMs !== entry.mtime)
              throw new Error('复制期间文件已变化，请重新扫描。')
            await this.verify(plan, entry, false)
            bytes = buffer.subarray(0, offset)
          } finally {
            await handle.close()
          }
          if (this.cancelled) break
          const doc = this.library.importBytes(basename(entry.path), bytes!, categoryId, duplicates)
          if (doc) result.imported.push(doc.id)
          else result.skipped++
        } catch (error) {
          result.failures.push({ name: entry.path, reason: reason(error) })
        }
        this.status!.completed++
      }
      result.cancelled = this.cancelled
      return result
    } finally {
      this.status = null
    }
  }
}
