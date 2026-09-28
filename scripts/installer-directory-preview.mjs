// Build a harmless native directory-page fixture for UI verification. It uses
// the production header, has no install/uninstall payload or registry writes,
// and quits if Install is clicked. Run it with Computer Use to verify Browse.
import { spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdir, mkdtemp, readdir, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import assert from 'node:assert/strict'

const root = await mkdtemp(join(resolve('test-results'), 'installer-directory-'))
const parent = join(root, 'Soft')
await mkdir(join(parent, 'local-docs'), { recursive: true })
await mkdir(join(parent, 'local-docs-backups'))
async function findCompiler(dir, depth = 0) {
  if (depth > 4) return null
  if (existsSync(join(dir, 'makensis.exe'))) return join(dir, 'makensis.exe')
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    if (!entry.isDirectory() || (depth === 0 && !entry.name.startsWith('nsis-'))) continue
    const found = await findCompiler(join(dir, entry.name), depth + 1)
    if (found) return found
  }
  return null
}
const compiler = await findCompiler(join(process.env.LOCALAPPDATA, 'electron-builder', 'Cache'))
if (!compiler) throw new Error('Build a Windows package first to cache NSIS')
const exe = join(root, 'directory-preview.exe')
const script = join(root, 'directory-preview.nsi')
const escape = (text) => text.replaceAll('$', '$$').replaceAll('"', '$\\"')
await writeFile(
  script,
  `Unicode true
Name "Local Docs Installer Path Check"
OutFile "${escape(exe)}"
RequestExecutionLevel user
!include "MUI2.nsh"
!define APP_EXECUTABLE_FILENAME "unused.exe"
!define UNINSTALL_FILENAME "unused-uninstall.exe"
!include "${escape(resolve('build/installer.nsh'))}"
!insertmacro customHeader
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_LANGUAGE "SimpChinese"
Section
  Quit
SectionEnd
`,
  'utf8',
)
const result = spawnSync(compiler, ['/V2', '/INPUTCHARSET', 'UTF8', script], {
  windowsHide: true,
  encoding: 'utf8',
  timeout: 60_000,
})
if (result.error) throw result.error
assert.equal(result.status, 0, result.stdout + result.stderr)
console.log(
  JSON.stringify({
    exe,
    parent,
    appDirectory: join(parent, 'local-docs'),
    similarDirectory: join(parent, 'local-docs-backups'),
  }),
)
