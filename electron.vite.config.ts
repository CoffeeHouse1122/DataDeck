import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import vue from '@vitejs/plugin-vue'
import { resolve } from 'node:path'

export default defineConfig({
  main: {
    plugins: [externalizeDepsPlugin()],
    build: {
      rollupOptions: {
        input: resolve('src/electron/main/index.ts')
      }
    },
    resolve: {
      alias: {
        '@shared': resolve('src/shared'),
        '@electron': resolve('src/electron'),
        '@main': resolve('src/electron/main')
      }
    }
  },
  preload: {
    plugins: [externalizeDepsPlugin()],
    build: {
      rollupOptions: {
        input: resolve('src/electron/preload/index.ts')
      }
    },
    resolve: {
      alias: {
        '@shared': resolve('src/shared')
      }
    }
  },
  renderer: {
    resolve: {
      alias: {
        '@renderer': resolve('src/renderer'),
        '@shared': resolve('src/shared')
      }
    },
    plugins: [vue()]
  }
})
