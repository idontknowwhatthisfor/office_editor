// Copies the PDF.js runtime files (worker, character maps, standard fonts, image decoders) into
// public/pdfjs so Vite serves them in dev and ships them in the build. They are loaded by URL at
// run time, so they can't go through the bundler.
import { cpSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const web = join(dirname(fileURLToPath(import.meta.url)), '..');
const src = join(web, 'node_modules', 'pdfjs-dist');
const dest = join(web, 'public', 'pdfjs');

rmSync(dest, { recursive: true, force: true });
mkdirSync(dest, { recursive: true });
for (const dir of ['cmaps', 'standard_fonts', 'wasm', 'iccs']) {
  cpSync(join(src, dir), join(dest, dir), { recursive: true });
}
// Renamed to .js so every WebView version serves it with a JavaScript MIME type.
cpSync(join(src, 'legacy', 'build', 'pdf.worker.min.mjs'), join(dest, 'pdf.worker.min.js'));
cpSync(join(src, 'LICENSE'), join(dest, 'LICENSE'));
