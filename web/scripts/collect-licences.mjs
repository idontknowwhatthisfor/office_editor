// Writes public/licences.json: every bundled third-party component with its licence text, for
// the app's "About & licences" screen (AboutActivity.kt reads it from assets/web/licences.json).
// Texts are read from the installed packages, so the notices always match what ships. Runs
// before `vite build` / `vite` (package.json). Add new runtime dependencies to WEB below.
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const web = join(dirname(fileURLToPath(import.meta.url)), '..');
const modules = join(web, 'node_modules');

/** Web packages bundled into the app, and what each is for. */
const WEB = [
  ['pdfjs-dist', 'PDF rendering, text and search'],
  ['docx-preview', 'Word (DOCX) rendering'],
  ['jszip', 'Reading and writing DOCX, XLSX and PPTX files', 'MIT (dual licensed MIT / GPL-3.0-or-later; used under MIT)'],
  ['pdf-lib', 'Writing edited PDFs'],
  ['@pdf-lib/standard-fonts', 'Writing edited PDFs (used by pdf-lib)'],
  ['@pdf-lib/upng', 'Writing edited PDFs (used by pdf-lib)'],
  ['pako', 'Compression (used by pdf-lib and jszip)'],
  ['tslib', 'Runtime helpers (used by pdf-lib)'],
  ['ssf', 'Spreadsheet number and date formats'],
  ['frac', 'Spreadsheet fraction formats (used by ssf)'],
  ['@fontsource/tinos', 'Tinos font, stand-in for Times New Roman'],
  ['@fontsource/arimo', 'Arimo font, stand-in for Arial'],
  ['@fontsource/carlito', 'Carlito font, stand-in for Calibri'],
  ['lie', 'Promises (used by jszip)'],
  ['immediate', 'Scheduling (used by jszip)'],
  ['setimmediate', 'Scheduling (used by jszip)'],
  ['readable-stream', 'Streams (used by jszip)'],
  ['safe-buffer', 'Buffers (used by jszip)'],
  ['string_decoder', 'Text decoding (used by jszip)'],
  ['core-util-is', 'Utilities (used by jszip)'],
  ['inherits', 'Utilities (used by jszip)'],
  ['isarray', 'Utilities (used by jszip)'],
  ['process-nextick-args', 'Utilities (used by jszip)'],
  ['util-deprecate', 'Utilities (used by jszip)'],
];

/** Files shipped with PDF.js that carry their own licences (copied by copy-pdf-assets.mjs). */
const PDFJS_EXTRA = [
  ['wasm/LICENSE_PDFJS_JBIG2', 'PDF.js JBIG2 image decoder', 'Apache-2.0'],
  ['wasm/LICENSE_JBIG2', 'PDFium JBIG2 decoder (in the PDF.js JBIG2 decoder)', 'BSD-3-Clause'],
  ['wasm/LICENSE_PDFJS_OPENJPEG', 'PDF.js JPEG 2000 image decoder', 'BSD-2-Clause'],
  ['wasm/LICENSE_OPENJPEG', 'OpenJPEG (in the PDF.js JPEG 2000 decoder)', 'BSD-2-Clause'],
  ['wasm/LICENSE_PDFJS_QCMS', 'PDF.js colour management', 'MIT'],
  ['wasm/LICENSE_QCMS', 'qcms colour management', 'MIT'],
  ['wasm/LICENSE_PDFJS_QUICKJS', 'PDF.js form scripting sandbox', 'MIT'],
  ['wasm/LICENSE_QUICKJS', 'QuickJS JavaScript engine (in the form scripting sandbox)', 'MIT'],
  ['standard_fonts/LICENSE_FOXIT', 'Foxit standard PDF fonts', 'BSD-3-Clause'],
  ['standard_fonts/LICENSE_LIBERATION', 'Liberation Sans font (standard PDF fonts)', 'GPL-2.0 with font exception'],
  ['cmaps/LICENSE', 'Adobe character maps for PDF text', 'BSD-3-Clause'],
  ['iccs/LICENSE', 'Colour profiles for PDF images', 'CC0-1.0'],
];

/** The native shell's libraries (Gradle). All are Apache-2.0; PDF.js ships that licence's text. */
const ANDROID = [
  ['AndroidX (Jetpack Compose, Activity, Core, Lifecycle, WebKit, Window and their dependencies)', 'App screens and the document view'],
  ['Kotlin standard library and kotlinx.coroutines', 'App language runtime'],
  ['Guava ListenableFuture', 'Used by AndroidX'],
  ['JetBrains annotations, JSpecify annotations', 'Used by AndroidX and Kotlin'],
];

const licenceFile = (dir) =>
  readdirSync(dir).find((f) => /^(licen[cs]e|copying)(\.(md|markdown|txt))?$/i.test(f));

const mit = (holder) => `MIT License

Copyright (c) ${holder}

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.`;

const read = (path) => readFileSync(path, 'utf8').replace(/\r\n/g, '\n').trim();

const out = [];
for (const [name, usedFor, licenceOverride] of WEB) {
  const dir = [join(modules, name), join(modules, 'pdf-lib', 'node_modules', name)].find((d) =>
    existsSync(join(d, 'package.json')),
  );
  if (!dir) throw new Error(`collect-licences: ${name} is not installed`);
  const pkg = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'));
  const file = licenceFile(dir);
  // A few small packages ship no licence file; their package.json says MIT, so give its text.
  const author = typeof pkg.author === 'string' ? pkg.author : pkg.author?.name ?? 'the authors';
  if (!file && pkg.license !== 'MIT') throw new Error(`collect-licences: ${name} has no licence file`);
  const text = file ? read(join(dir, file)) : mit(author);
  out.push({ name, version: pkg.version, licence: licenceOverride ?? pkg.license, usedFor, text });
}
for (const [path, usedFor, licence] of PDFJS_EXTRA) {
  out.push({ name: usedFor, version: '', licence, usedFor: 'Part of PDF.js', text: read(join(modules, 'pdfjs-dist', path)) });
}
const apache = read(join(modules, 'pdfjs-dist', 'LICENSE'));
for (const [name, usedFor] of ANDROID) {
  out.push({ name, version: '', licence: 'Apache-2.0', usedFor, text: apache });
}

mkdirSync(join(web, 'public'), { recursive: true });
writeFileSync(join(web, 'public', 'licences.json'), JSON.stringify(out));
console.log(`collect-licences: ${out.length} components`);
