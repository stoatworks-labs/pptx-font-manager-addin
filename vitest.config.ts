import { defineConfig } from 'vitest/config'
import { resolve } from 'node:path'
export default defineConfig({
  resolve: { alias: { '@fm': resolve(import.meta.dirname, 'vendor/pptx-font-manager/src') } },
  test: { environment: 'node', include: ['test/**/*.test.ts'] },
})
