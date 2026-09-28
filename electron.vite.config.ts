import { defineConfig } from 'electron-vite'
import react from '@vitejs/plugin-react'
import { resolve } from 'node:path'
import { cpSync, mkdirSync } from 'node:fs'

// Generated local assets are copied by Vite in both development and packaged builds.
const pdfAssets = resolve('src/renderer/public/pdf-assets')
mkdirSync(pdfAssets, { recursive: true })
for (const name of ['cmaps', 'standard_fonts', 'wasm', 'iccs', 'LICENSE'])
  cpSync(resolve('node_modules/pdfjs-dist', name), resolve(pdfAssets, name), { recursive: true })

export default defineConfig({
  main: {
    build: {
      rollupOptions: {
        input: {
          index: resolve('src/main/index.ts'),
          'office-worker': resolve('src/main/office-worker.ts'),
        },
      },
    },
  },
  preload: {},
  renderer: {
    plugins: [
      react(),
      {
        name: 'development-csp',
        apply: 'serve',
        transformIndexHtml: (html) =>
          html.replace("script-src 'self'", "script-src 'self' 'unsafe-inline'"),
      },
    ],
    server: { host: '127.0.0.1' },
  },
})
