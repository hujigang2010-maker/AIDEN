import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { viteSingleFile } from 'vite-plugin-singlefile'

// SINGLE=1 时产出可直接双击打开的单文件 HTML（内联全部 JS/CSS）
const single = process.env.SINGLE === '1'

export default defineConfig({
  plugins: [react(), ...(single ? [viteSingleFile()] : [])],
  server: {
    host: true,
    port: 5173,
  },
  build: single
    ? {
        outDir: 'release/standalone',
        chunkSizeWarningLimit: 4000,
        assetsInlineLimit: 100000000,
        cssCodeSplit: false,
      }
    : {
        chunkSizeWarningLimit: 900,
        rollupOptions: {
          output: {
            manualChunks: {
              echarts: ['echarts', 'echarts-for-react'],
              react: ['react', 'react-dom'],
            },
          },
        },
      },
})
