import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// The complete upstream App/Ribbon/editor, with an explicit Web host entry.
export default defineConfig({
  root: resolve(import.meta.dirname, 'src/renderer'),
  base: './',
  plugins: [react()],
  build: {
    target: 'esnext',
    outDir: resolve(import.meta.dirname, 'out/web'),
    emptyOutDir: true,
    rollupOptions: { input: resolve(import.meta.dirname, 'src/renderer/web.html') },
  },
})
