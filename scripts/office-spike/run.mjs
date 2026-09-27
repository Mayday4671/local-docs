import { _electron as electron, expect } from '@playwright/test'
import { readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
const root = resolve('test-results/office-spike')
const env = { ...process.env }
delete env.ELECTRON_RUN_AS_NODE
const app = await electron.launch({ args: [resolve(root, 'electron.cjs')], env })
const report = { version: '2.18.0', engine: '0.17.0', errors: [], network: [], phases: [] }
try {
  const page = await app.firstWindow()
  page.on('pageerror', (e) => report.errors.push(String(e)))
  page.on('console', (msg) => {
    if (msg.type() === 'error') report.errors.push(msg.text())
  })
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].showInactive())
  await page.waitForFunction(() => typeof window.loadDocument === 'function')
  await page.evaluate(
    (data) => window.loadDocument(data),
    (await readFile(resolve(root, 'fixture.docx'))).toString('base64'),
  )
  report.phases.push('opened')
  await page.screenshot({ path: resolve(root, 'opened.png') })
  await writeFile(
    resolve(root, 'unmodified-export.docx'),
    Buffer.from(await page.evaluate(() => window.exportDocument()), 'base64'),
  )
  report.phases.push('unmodified-export')
  await page.evaluate(() => window.spikeEditor.ui.search.find('ORIGINAL_MARKER'))
  await expect
    .poll(() => page.evaluate(() => window.spikeEditor.ui.search.getSnapshot().total))
    .toBe(1)
  const edit = await page.evaluate(() => window.spikeEditor.ui.search.replace('已完成离线修改'))
  report.phases.push('edited')
  const saved = await page.evaluate(() => window.exportDocument())
  await writeFile(resolve(root, 'edited.docx'), Buffer.from(saved, 'base64'))
  report.phases.push('saved')
  await page.evaluate((data) => window.loadDocument(data), saved)
  await expect(page.locator('body')).toContainText('已完成离线修改')
  await expect(page.locator('body')).not.toContainText('ORIGINAL_MARKER')
  report.phases.push('reopened')
  await page.screenshot({ path: resolve(root, 'reopened.png') })
} catch (error) {
  report.errors.push(String(error))
  process.exitCode = 1
  console.log(String(error))
} finally {
  report.network = await app.evaluate(() => global.networkAttempts)
  await writeFile(resolve(root, 'result.json'), JSON.stringify(report, null, 2))
  console.log(JSON.stringify(report))
  await app.close()
  if (report.errors.length || report.network.length) process.exitCode = 1
}
