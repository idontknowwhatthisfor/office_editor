# web/src — web viewers and editors (TypeScript)

## Purpose

All the code that runs inside the app's WebView: one entry point that asks the native shell which
document is open, then a viewer/editor per format, plus the pieces they share.

## Contents

| Path | What it does |
| --- | --- |
| [`pdf/`](pdf/README.md) | PDF viewer (PDF.js) and editor (pdf-lib). |
| [`docx/`](docx/README.md) | Word viewer (docx-preview + own pagination) and in-place editor. |
| [`xlsx/`](xlsx/README.md) | Spreadsheet viewer, formula engine and editor. |
| [`pptx/`](pptx/README.md) | Presentation renderer, viewer, Present mode and editor. |
| [`chart/`](chart/README.md) | Chart drawing shared by the three Office viewers. |
| `main.ts` | Entry point. Gets the document from the bridge (or `?file=` / a file picker in dev), fetches its bytes and dynamically imports that format's viewer. |
| `bridge.ts` | Types for `window.OfficeBridge` and `nativeDocument()`. The bridge is absent in a desktop browser. |
| `save.ts` | Saving, shared by the editors: `registerEditor`, `setDirty`, and streaming the built file to the shell in base64 chunks. Offers a download when there is no shell. |
| `print.ts` | `preparePrint()`: lays a Word/Excel/PowerPoint file out for paper ("Export as PDF"). |
| `immersive.ts` | A single tap on a PDF, Word or PowerPoint document (not editing or presenting) slides the bottom bar down and, through `OfficeBridge.setBarsHidden`, the shell's top bar up (and hides the status bar); another tap brings them back. A double tap (zoom) leaves them alone. The top bar floats over these viewers: `--oe-top` / `topInset()` is the room the page keeps for it. |
| `toolbar-mirror.ts` | `mirrorToolbar(row)`: in the app, the shell draws an editor toolbar (the PDF edit row, the Word edit bar) natively from a description of its buttons, sent whenever they change (`OfficeBridge.setToolbar`); taps and holds come back through `window.officeToolbar` and are replayed on the real buttons (press, release, click; a hold is `contextmenu`). The row stays in the page, invisible (`.oe-mirrored`). |
| `pagebar.ts` | The bottom bar of the PDF and Word viewers: page indicator, thumbnails, search, edit button. Also `iconButton`. In the app the shell draws the main and search rows natively (`PageBar.kt`; the page reports with `setPageBar` and gets taps through `window.officePageBar`); the page keeps an invisible spacer row, the thumbnail strip and its own rows (the PDF edit toolbar). In a desktop browser it draws them all. |
| `overlay.ts` | The layer editors' panels, toasts and dialogs live in, counter-scaled so pinch-zoom doesn't magnify them. |
| `ui.ts` | Dialogs, menus, toast, the font size field (`askFontSize`, `fontSizeField`, `parseFontSize`), and `holdable` (tap vs hold on a button). |
| `ooxml.ts` | Zip/XML/relationship helpers for Office files (`Package`, `child`, `children`, `first`, `all`, `relAttr`, `relOfType`). |
| `i18n.ts` | All web strings in English and Indonesian; `t(key, params)`. Indonesian is used when the browser language starts with `id` or `in`. |
| `message.ts` | `showMessage()`: replaces the page with a centred message (loading, errors). |
| `style.css`, `ui.css`, `pagebar.css`, `print.css` | Base styles and the styles of the modules above. |

## How it works

- **Opening**: `main.ts` calls `OfficeBridge.getDocument()` → `{name, format, mime, size, url, edit,
  print}`. Bytes come from `fetch(url)`, not through the bridge. PDFs are read from the URL in
  ranges; other formats are fetched whole. Each format is a separate chunk, so opening a PDF
  never loads spreadsheet code.
- **Saving**: the shell owns the Save button. An editor calls `setDirty(true)`; on Save the shell
  calls `window.officeSave(mode)`; the editor's `produce()` builds the file and re-reads it, and
  `save.ts` streams it back (`beginSave` / `writeChunk` / `finishSave`). The shell answers with
  `window.officeSaveResult(...)`.
- **Printing**: with `print` set, `main.ts` calls `preparePrint()` and reports the page size with
  `OfficeBridge.printReady(w, h)` (or `printFailed()`).

## Dependencies & relationships

- The bridge's native side is
  [`DocumentWebView.kt`](../../app/src/main/java/com/officeeditor/README.md).
- The format folders are not fully independent: `chart/` uses `pptx/` and `xlsx/` helpers, the
  Word and Excel code reads themes with `pptx/deck.ts`, the Word editor uses `pdf/style.ts`, and
  `docx/fonts.css` is imported by all three Office viewers.

## Usage notes

- Every user-visible string goes in `i18n.ts`, in both languages.
- Dev URL options (dev server only): `?file=/@fs/<abs path>`, `&edit`, `&print`.
- No network requests: everything must be served by the app itself.
