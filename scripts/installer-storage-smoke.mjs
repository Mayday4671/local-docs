// Compile the exact cleanup hook into a minimal uninstaller. This harness has no
// registry/shortcut/install detection logic and never runs the production installer.
import { spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdir, mkdtemp, readFile, readdir, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import assert from 'node:assert/strict'

if (process.platform !== 'win32') throw new Error('Windows NSIS verification only')
const results = resolve('test-results')
await mkdir(results, { recursive: true })
const root = await mkdtemp(join(results, 'installer-storage-'))
const installation = join(root, 'installation')
const cache = join(process.env.LOCALAPPDATA, 'electron-builder', 'Cache')
async function findCompiler(dir, depth = 0) {
  if (depth > 4) return null
  if (existsSync(join(dir, 'makensis.exe'))) return join(dir, 'makensis.exe')
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    if (!entry.isDirectory() || (depth === 0 && !entry.name.startsWith('nsis-'))) continue
    const result = await findCompiler(join(dir, entry.name), depth + 1)
    if (result) return result
  }
  return null
}
const compiler = await findCompiler(cache)
if (!compiler) throw new Error('Build the Windows installer first to populate the NSIS tool cache')
await mkdir(join(installation, 'resources'), { recursive: true })
await mkdir(join(installation, 'locales'))
await mkdir(join(installation, 'data', 'library', 'objects'), { recursive: true })
const retained = [
  join(installation, 'data', 'library', 'library.sqlite'),
  join(installation, 'data', 'library', 'objects', 'synthetic-object'),
  join(installation, '用户保留文件.txt'),
  join(root, 'outside.txt'),
]
for (const path of retained) await writeFile(path, `Must be retained: ${path}`)
const payload = [
  '我的文档库.exe',
  'resources/app.asar',
  'locales/zh-CN.pak',
  'icudtl.dat',
  'ffmpeg.dll',
]
for (const path of payload)
  await writeFile(join(installation, path), 'synthetic application payload')
const escape = (path) => path.replaceAll('$', '$$').replaceAll('"', '$\\"')
const bootstrap = join(root, 'make-test-uninstaller.exe'),
  uninstaller = join(installation, 'test-uninstall.exe')
const script = join(root, 'storage-cleanup.nsi')
await writeFile(
  script,
  `Unicode true
Name "Local Docs isolated cleanup verification"
OutFile "${escape(bootstrap)}"
RequestExecutionLevel user
SilentInstall silent
SilentUnInstall silent
!define APP_EXECUTABLE_FILENAME "我的文档库.exe"
!define UNINSTALL_FILENAME "test-uninstall.exe"
!include "${escape(resolve('build/installer.nsh'))}"
Section
  WriteUninstaller "${escape(uninstaller)}"
SectionEnd
Section "Uninstall"
  StrCpy $INSTDIR "${escape(installation)}"
  !insertmacro customRemoveFiles
SectionEnd
`,
  'utf8',
)
function run(file, args) {
  const result = spawnSync(file, args, { windowsHide: true, encoding: 'utf8', timeout: 60_000 })
  if (result.error) throw result.error
  assert.equal(result.status, 0, result.stdout + result.stderr)
}
run(compiler, ['/V2', '/INPUTCHARSET', 'UTF8', script])
run(bootstrap, ['/S'])
// _?= keeps the test uninstaller in the verified isolated directory, so the
// process we wait on performs cleanup itself rather than spawning a detached copy.
run(uninstaller, ['/S', `_?=${installation}`])
for (const path of retained) assert.equal(await readFile(path, 'utf8'), `Must be retained: ${path}`)
for (const path of payload) assert.equal(existsSync(join(installation, path)), false, path)
assert.equal(existsSync(join(installation, 'resources')), false)
assert.equal(existsSync(join(installation, 'locales')), false)
console.log(`Installer cleanup preserved the library and unrelated files: ${root}`)
