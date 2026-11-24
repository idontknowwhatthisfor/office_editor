# web/src/docx — Word viewer and editor

## Purpose

Shows `.docx` files in fixed page layout, as Word does, and edits them in place on the file's own
XML. Loaded as its own chunk by [`../main.ts`](../main.ts).

## Contents

| File | What it does |
| --- | --- |
| `viewer.ts` | `openDocx()`: renders with docx-preview, paginates, fixes line spacing, fills page numbers, scales pages to the screen, adds the bottom bar (drawn by the shell in the app) and a tap to hide the bars (`../immersive.ts`). `printDocx()` lays the pages out for Export as PDF. Also exports `fontsLoaded`, `scaleLineHeights`, `GUTTER` for the editor. |
| `paginate.ts` | `paginate()`: splits docx-preview's output into real pages (paragraphs between lines, tables between rows). |
| `search.ts` | `DocxSearch`: finds text and highlights it with the CSS Custom Highlight API, leaving the markup alone. |
| `charts.ts` | `prepareCharts()`: matches docx-preview's empty chart boxes to the document's XML and draws the charts with [`../chart/`](../chart/README.md). |
| `editor.ts` | `startDocxEditing()`: edit mode. Text, bold/italic/underline, colours, font size, alignment, headings, lists, undo/redo, save. |
| `pictures.ts` | `PictureTool`: insert, select, move, resize, crop and delete pictures. |
| `docx.css` | Overrides for docx-preview's markup, plus the editor's styles. |
| `fonts.css` | **Generated** by [`../../scripts/make-docx-fonts.mjs`](../../scripts/make-docx-fonts.mjs); do not edit. Aliases Times New Roman / Arial / Calibri to Tinos / Arimo / Carlito. |

## How it works

- **Viewing**: docx-preview only breaks pages at explicit page breaks, so `paginate.ts` does the
  rest by measuring layout. It must run with the document attached, fonts loaded and no zoom.
  Several docx-preview quirks are worked around in `viewer.ts`; read its comments before
  upgrading the library.
- **Editing** (`editor.ts`): the edit view is docx-preview's rendering on real-size paper (not
  split into pages). A copy of `document.xml` gets a bookmark before every paragraph and run so
  screen elements map back to XML. Only paragraphs the user changed are rewritten, from their
  original `pPr`/`rPr`; the rest stay byte for byte. Paragraphs holding anything but plain text
  runs are locked. Undo is a snapshot of the view's HTML.
- **Pictures** (`pictures.ts`): each editable picture is wrapped in `span.oe-pic` whose attributes
  hold its state (`data-cx/cy` in EMU, `data-crop`), so HTML-snapshot undo covers it. The
  selection UI lives in the overlay, in screen coordinates. On save only extents, crops and run
  positions change, and new files go into `word/media` with a relationship and content type.

## Dependencies & relationships

- Libraries: `docx-preview`, `jszip`, `@fontsource/{tinos,arimo,carlito}` (through `fonts.css`).
- Uses [`../pagebar.ts`](../pagebar.ts), [`../save.ts`](../save.ts), [`../ui.ts`](../ui.ts),
  [`../overlay.ts`](../overlay.ts), [`../ooxml.ts`](../ooxml.ts), [`../chart/`](../chart/README.md),
  [`../pdf/style.ts`](../pdf/style.ts) (settings panel), `readTheme` from
  [`../pptx/deck.ts`](../pptx/deck.ts).
- Used by: `../main.ts`, `../print.ts`. `fonts.css` is also imported by the XLSX and PPTX viewers.

## Usage notes

- Dev: `?file=/@fs/<repo>/test-files/sample.docx`; add `&edit` for edit mode, `&print` for the
  print layout.
- Check an edited file in Word with [`../../../test-files/word_to_pdf.sh`](../../../test-files/word_to_pdf.sh).
- To add a stand-in font, edit the list in `make-docx-fonts.mjs` and rerun it.

## Known limits

Carried over from the handoff notes written at release 1.1.0. Items fixed since were dropped; the rest were not each rechecked.

Design decision: a Word file must look as it does in Word, with fixed pages, no reflow, and the
whole page zooming. An earlier reflowed "reading mode" was replaced at the owner's request; don't
bring it back unasked.

Page layout:
- On the test essay 17 of 20 page breaks equalled Word's; three pages broke one or two lines
  early. Cause: Word compresses spaces in justified lines to fit another word; CSS can't.
- Not handled, and untested: footnotes at the page bottom, multi-column sections, text boxes and
  floating images, `keepLines`, a paragraph or table row taller than a page.
- Page numbers are a guess: any header/footer run that is only digits is replaced with the page
  index (`numberPages` in `viewer.ts`). A year in a footer would be overwritten.
- Only Times New Roman, Arial and Calibri have stand-in fonts, Latin subsets only. Other fonts
  fall back to the system font and paginate differently.
- Links are inert. Pagination runs synchronously after render; not timed on long documents.
- docx-preview only breaks pages at explicit breaks, writes Word's line-spacing multiples as CSS
  multiples of font size, hyphenates everything, and misreads `<w:tblHeader/>`. Each has a
  workaround in `viewer.ts`.

Editor:
- The edit view is the page at paper size, scaled to the screen, but not split into pages.
- Locked (shown greyed, kept unchanged): paragraphs with links, fields (TOC, page numbers),
  comments, tracked changes, text boxes. Paragraphs can't be joined across table cells. A warning
  lists such content before editing.
- Not done: insert table, links, find & replace, word count. Lists are one level only; the
  heading look in the edit view is a preview.
