import type { Plugin } from 'vite'
import type { ViteUserConfig } from 'vitest/config'
import { readFile } from 'node:fs/promises'
import process from 'node:process'
import { fileURLToPath, URL } from 'node:url'
import ui from '@nuxt/ui/vite'
import vue from '@vitejs/plugin-vue'
import { getHttpsServerOptions } from 'office-addin-dev-certs'
import { loadEnv } from 'vite'
import { defineConfig } from 'vitest/config'

const DEV_SERVER_PORT = 3000
const DEV_SERVER_URL = `https://localhost:${DEV_SERVER_PORT}/`

function fromRoot(path: string): string {
  return fileURLToPath(new URL(path, import.meta.url))
}

/**
 * Emits `manifest.xml` into the build output with every dev server URL
 * rewritten to the URL the add-in is hosted at.
 */
function officeManifest(publicUrl: string): Plugin {
  return {
    name: 'office-manifest',
    apply: 'build',
    async generateBundle() {
      if (publicUrl === DEV_SERVER_URL)
        this.warn('ADDIN_PUBLIC_URL is not set, dist/manifest.xml still points at the dev server.')

      const manifest = await readFile(fromRoot('./manifest.xml'), 'utf8')
      this.emitFile({
        type: 'asset',
        fileName: 'manifest.xml',
        source: manifest.replaceAll(DEV_SERVER_URL, publicUrl),
      })
    },
  }
}

export default defineConfig(async ({ command, mode }): Promise<ViteUserConfig> => {
  const env = loadEnv(mode, process.cwd(), 'ADDIN_')
  const publicUrl = (env['ADDIN_PUBLIC_URL'] ?? DEV_SERVER_URL).replace(/\/?$/, '/')

  // Outlook only loads add-ins over HTTPS. Reuse the localhost certificate that
  // office-addin-dev-certs installs, but don't touch certificates when running tests.
  const devServer = {
    port: DEV_SERVER_PORT,
    strictPort: true,
    ...(command === 'serve' && mode !== 'test' && { https: await getHttpsServerOptions() }),
  }

  return {
    // Relative asset URLs, so the build can be hosted under any path.
    base: './',
    plugins: [
      vue(),
      ui({
        // The task pane is a single view, so links render as plain anchors.
        router: false,
        ui: {
          colors: { primary: 'blue', neutral: 'slate' },
        },
      }),
      officeManifest(publicUrl),
    ],
    resolve: {
      alias: {
        '@': fromRoot('./src'),
      },
    },
    server: devServer,
    preview: devServer,
    build: {
      rolldownOptions: {
        input: {
          taskpane: fromRoot('./taskpane.html'),
          commands: fromRoot('./commands.html'),
        },
      },
    },
    test: {
      environment: 'jsdom',
      include: ['src/**/*.spec.ts'],
      unstubGlobals: true,
      coverage: {
        provider: 'v8',
        include: ['src/**/*.{ts,vue}'],
        exclude: ['src/test-utils/**'],
        reporter: ['text', 'html', 'cobertura'],
      },
    },
  }
})
