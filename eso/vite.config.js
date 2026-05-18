import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [vue()],
  server: {
    host: '0.0.0.0',
    port: 3030,
    proxy: {
      '/api': {
        target: 'http://127.0.0.1:8030',
        changeOrigin: true,
      },
      '/sql-executor': {
        target: 'http://127.0.0.1:8030',
        changeOrigin: true,
      }
    }
  }
})
