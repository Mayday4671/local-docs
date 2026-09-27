import { Worker } from 'node:worker_threads'
import { join } from 'node:path'
import { readFile } from 'node:fs/promises'
import type { Library } from './library'
import type { OfficeData, OfficePreview, VersionContent } from '../shared/types'

export class OfficeService {
  private queue: Promise<unknown> = Promise.resolve()
  private running = new Map<string, Promise<void>>()
  private worker: Worker | undefined
  private disposed = false
  constructor(private readonly library: Library) {}

  private parse(path: string, extension: string): Promise<{ data?: OfficeData; error?: string }> {
    return new Promise((resolve) => {
      const worker = new Worker(join(__dirname, 'office-worker.js'), {
        workerData: { path, extension },
        resourceLimits: { maxOldGenerationSizeMb: 192 },
      })
      this.worker = worker
      let settled = false
      const finish = (result: { data?: OfficeData; error?: string }) => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        this.worker = undefined
        void worker.terminate()
        resolve(result)
      }
      const timer = setTimeout(() => finish({ error: '读取文档超时，文件仍可原样导出。' }), 15000)
      worker.once('message', finish)
      worker.once('error', () => finish({ error: '文档过于复杂或内容异常，当前无法读取。' }))
      worker.once('exit', () => finish({ error: '文档读取进程意外结束。' }))
    })
  }

  ensure(id: string): Promise<void> {
    if (this.disposed) return Promise.resolve()
    const source = this.library.officeSource(id)
    if (this.library.officeCache(id, source.hash)) return Promise.resolve()
    const key = `${id}:${source.hash}`
    if (this.running.has(key)) return this.running.get(key)!
    const task = this.queue
      .then(async () => {
        if (this.disposed) return
        const result = await this.parse(source.path, source.extension)
        if (!this.disposed)
          this.library.cacheOffice(id, source.hash, result.data ?? null, result.error ?? null)
      })
      .finally(() => this.running.delete(key))
    this.queue = task.catch(() => {})
    this.running.set(key, task)
    return task
  }

  async indexPending(): Promise<void> {
    for (const id of this.library.pendingOfficeIds()) await this.ensure(id)
  }

  async read(id: string, includeLayout = false): Promise<OfficePreview> {
    await this.ensure(id)
    const source = this.library.officeSource(id)
    const cached = this.library.officeCache(id, source.hash)
    if (!cached) throw new Error('文档发生变化，请重新打开。')
    return {
      ...cached,
      ...(includeLayout && cached.data?.kind === 'docx'
        ? { bytes: new Uint8Array(await readFile(source.path)) }
        : {}),
    }
  }

  async readVersion(id: string, versionId: string): Promise<VersionContent> {
    const { path, ...source } = this.library.versionSource(id, versionId)
    if (source.text !== null) return { ...source, office: null }
    // Historical reads share the bounded worker queue but never replace the current index.
    const task = this.queue.then(async () => {
      if (this.disposed) throw new Error('文档库已关闭。')
      const result = await this.parse(path, source.extension)
      if (!result.data) throw new Error(result.error || '无法读取此历史版本。')
      return { ...source, office: result.data }
    })
    this.queue = task.catch(() => {})
    return task
  }

  dispose(): void {
    this.disposed = true
    void this.worker?.terminate()
  }
}
