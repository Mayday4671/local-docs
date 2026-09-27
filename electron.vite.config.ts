import { defineConfig } from 'electron-vite'
import react from '@vitejs/plugin-react'
import { resolve } from 'node:path'

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
