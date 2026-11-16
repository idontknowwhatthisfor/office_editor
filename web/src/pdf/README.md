# web/src/pdf — PDF viewer and editor

## Purpose

Shows PDFs with PDF.js and lets the user mark them up, fill forms, sign, and rotate or delete
pages. Loaded as its own chunk by [`../main.ts`](../main.ts) when the open document is a PDF.

## Contents

| File | What it does |
| --- | --- |
| `viewer.ts` | `openPdf()`: renders pages lazily on canvases (capped per page to limit memory), text layer for selection, the bottom bar, and starts `PdfEditor`. Reads the file from its URL in ranges. |
| `search.ts` | `PdfSearch`: finds text across pages and highlights matches in the text layer. |
| `editor.ts` | `PdfEditor`: edit mode. Tools (highlight, underline, pen, text, signature, eraser), form inputs over the page, page rotate/delete, undo/redo, and saving via `write.ts`. |
| `edits.ts` | The edit model (`Edit`, `Style`, `PageSpace`) in PDF page space (points, y up, unrotated), and `editSvg()`, which draws an edit as the on-screen SVG overlay. |
| `write.ts` | `writeEdits()`: writes edits into the file with pdf-lib. Also page operations and form values. Throws `EncryptedPdfError` for encrypted files. |
| `marks.ts` | `readMarks()`: reads the app's own saved annotations back into `Edit`s so they can be erased or changed again. |
| `erase.ts` | The two erasers in PDF space: `eraseWhole` (whole marks; one line of a highlight) and `erasePart` (cuts pen strokes). |
| `textbox.ts` | `TextBoxes`: select a text box the app placed, then move, scale, retype, recolour or delete it. |
| `style.ts` | `openStylePanel()`: a tool's settings panel (colour grid / HSV sliders / eyedropper, opacity, pen thickness, eraser mode and size). Settings are kept in localStorage. Also used by the Word editor. |
| `pdf.css` | Styles for the pages, edit layer and edit bar. |
| `style.css` | Styles for the settings panel. |

## How it works

- **Edits stay in memory** as `Edit` objects in PDF space, drawn over the pages as SVG, until the
  user saves. Undo works on operations; an eraser swipe is one `erase` op, a text-box change one
  `replace` op.
- **Saving** (`write.ts`): highlights, underlines, ink and text boxes become real annotations,
  each with its own appearance stream and named `/NM (OfficeEditor-…)`. Signatures are drawn into
  the page. See-through marks use an inner transparency group with `ca`/`CA` rather than an
  annotation-level `/CA`. Form fields get fresh appearances.
- **Re-editing saved marks** (`marks.ts`): only annotations this app wrote are picked up (by name,
  or by the exact key set older versions wrote). Other apps' annotations are left alone. PDF.js
  is told to leave the picked-up ones off the canvas through its annotation storage.
- **Touch**: while a tool is on, the edit layer has `touch-action: pinch-zoom`: one finger uses
  the tool, two zoom and scroll.
- Panels, toasts and dialogs live in [`../overlay.ts`](../overlay.ts)'s layer so pinch-zoom does
  not magnify them.

## Dependencies & relationships

- Libraries: `pdfjs-dist` (legacy build; its worker, fonts, cmaps and decoders are served from
  `public/pdfjs`, copied by [`../../scripts/copy-pdf-assets.mjs`](../../scripts/copy-pdf-assets.mjs)),
  `pdf-lib`.
- Uses [`../pagebar.ts`](../pagebar.ts), [`../save.ts`](../save.ts), [`../ui.ts`](../ui.ts),
  [`../overlay.ts`](../overlay.ts), [`../i18n.ts`](../i18n.ts), [`../message.ts`](../message.ts).
- Used by: `../main.ts`; `style.ts` is also imported by [`../docx/editor.ts`](../docx/editor.ts).

## Usage notes

- Dev: `http://localhost:5173/?file=/@fs/<repo>/test-files/pdf-form.pdf`.

## Known limits

Carried over from the handoff notes written at release 1.1.0. Items fixed since were dropped; the rest were not each rechecked.

Viewer:
- Memory grows as a large PDF is read: PDF.js keeps every range it has loaded.
- Zoomed pages are capped at about 6 megapixels, so sharpness stops improving past roughly 2×.
- No password prompt (a message is shown) and no link handling.
- Range serving reports sizes through `InputStream.available()`, an `int`: files over 2 GB will
  misbehave.

Editor:
- Not done: merge/split PDFs, changing annotations made by other apps, signature fields
  (signatures are drawn into the page, not digitally signed).
- Text boxes use Helvetica: characters outside Windows-1252 are saved as "?". On a page with
  `/Rotate`, a text box is written along the page's own axis (looks turned).
- PDFs over 150 MB can't be edited (`MAX_EDIT_BYTES`; saving rewrites the whole file in memory).
  Encrypted PDFs can't be saved (a message is shown).
- Saved files were checked with pdf-lib, macOS PDFKit and PDF.js, not in Chrome's or Acrobat's
  viewer.
