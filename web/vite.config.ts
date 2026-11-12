import { defineConfig } from 'vite';

export default defineConfig({
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
