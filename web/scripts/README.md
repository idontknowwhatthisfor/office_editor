# web/scripts — build helper scripts

## Purpose

Node scripts that prepare generated files for the web app. Two run automatically before every
`npm run dev` and `npm run build`; one is run by hand.

## Contents

| File | Runs | What it does |
| --- | --- | --- |
| `copy-pdf-assets.mjs` | `predev`, `prebuild` | Copies PDF.js runtime files (worker, cmaps, standard fonts, wasm decoders, ICC profiles, licence) from `node_modules/pdfjs-dist` into `public/pdfjs`. The worker is renamed `.mjs` → `.js` so every WebView serves it as JavaScript. |
| `collect-licences.mjs` | `predev`, `prebuild` | Writes `public/licences.json`: every bundled component with its licence text, read from the installed packages. The About screen shows it. |
| `make-docx-fonts.mjs` | by hand | Regenerates [`../src/docx/fonts.css`](../src/docx/fonts.css): `@font-face` rules that make Tinos, Arimo and Carlito answer to Times New Roman, Arial and Calibri (latin and latin-ext subsets). |

## Dependencies & relationships

- Read from `web/node_modules`, so `npm ci` must have run.
- Outputs `public/pdfjs/` and `public/licences.json` are git-ignored; `fonts.css` is committed.
- `licences.json` is read by `AboutActivity.kt` from `assets/web/licences.json`.

## Usage notes

- Adding a runtime dependency: add it to the `WEB` list in `collect-licences.mjs` (files PDF.js
  ships with their own licence go in `PDFJS_EXTRA`, Gradle libraries in `ANDROID`).
- After changing the font list or updating `@fontsource` packages:

  ```bash
  node scripts/make-docx-fonts.mjs
  ```
