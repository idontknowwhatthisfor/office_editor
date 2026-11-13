import { mkdirSync, writeFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig, type Plugin } from 'vite';

/**
 * Development only: lets a test in the browser save a file it built (an edited document) into
 * ../test-files/out/, so it can be checked with Word, Excel or PowerPoint. POST the bytes to
 * /__dev-save?name=file.ext. Not part of the app build.
 */
const devSave: Plugin = {
  name: 'dev-save',
  apply: 'serve',
  configureServer(server) {
    server.middlewares.use('/__dev-save', (req, res) => {
      const name = basename(new URL(req.url ?? '', 'http://x').searchParams.get('name') ?? 'out.bin');
      const chunks: Buffer[] = [];
      req.on('data', (c: Buffer) => chunks.push(c));
      req.on('end', () => {
        const dir = fileURLToPath(new URL('../test-files/out', import.meta.url));
        mkdirSync(dir, { recursive: true });
        writeFileSync(join(dir, name), Buffer.concat(chunks));
        res.end(join(dir, name));
      });
    });
  },
};

export default defineConfig({
  plugins: [devSave],
  // Relative URLs: the app is served from /assets/web/ inside the Android WebView.
  base: './',
  // Lets the dev server read ../test-files (see the `?file=` option in src/main.ts).
  server: { fs: { allow: ['..'] } },
  build: {
    outDir: '../app/src/main/assets/web',
    emptyOutDir: true,
    target: 'es2022',
  },
});
