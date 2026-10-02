import { access, readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { spawnSync } from 'node:child_process'

const { version } = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'))
const executablePath = resolve(process.argv[2] || `release/${version}/win-unpacked/我的文档库.exe`)
await access(executablePath)
for (const script of [
  'scripts/storage-smoke.mjs',
  'scripts/file-import-smoke.mjs',
  'scripts/text-syntax-smoke.mjs',
  'scripts/pdf-smoke.mjs',
  'scripts/smoke.mjs',
  'scripts/office-smoke.mjs',
  'scripts/visual-smoke.mjs',
  'scripts/formatting-smoke.mjs',
  'scripts/transfer-smoke.mjs',
  'scripts/organization-smoke.mjs',
  'scripts/content-smoke.mjs',
  'scripts/office-layout-smoke.mjs',
  'scripts/word-edit-fix-smoke.mjs',
  'scripts/history-smoke.mjs',
  'scripts/editor-split-smoke.mjs',
  'scripts/library-split-smoke.mjs',
]) {
  const result = spawnSync(process.execPath, [script], {
    stdio: 'inherit',
    windowsHide: true,
    env: { ...process.env, LOCAL_DOCS_TEST_EXECUTABLE: executablePath },
  })
  if (result.error) throw result.error
  if (result.status !== 0) process.exit(result.status ?? 1)
}
console.log(`Packaged application verified: ${executablePath}`)
