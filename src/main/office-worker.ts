import { parentPort, workerData } from 'node:worker_threads'
import { readFileSync, statSync } from 'node:fs'
import { parseOffice } from './office-parser'

try {
  if (statSync(workerData.path).size > 20 * 1024 * 1024)
    throw new Error('当前仅预览 20 MB 以内的 Office 文件；原文件仍可导出。')
  parentPort!.postMessage({
    data: parseOffice(readFileSync(workerData.path), workerData.extension),
  })
} catch (error) {
  parentPort!.postMessage({ error: error instanceof Error ? error.message : '文档读取失败。' })
}
