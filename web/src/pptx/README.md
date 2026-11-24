# web/src/pptx — presentation viewer and editor

## Purpose

Draws `.pptx` slides with our own renderer (no third-party PPTX library), presents them full
screen, and edits text and shape positions in place on each slide's XML. Loaded as its own chunk
by [`../main.ts`](../main.ts).

## Contents

| File | What it does |
| --- | --- |
| `deck.ts` | `Deck`: loads slide → layout → master → theme. `readTheme()` and `DEFAULT_THEME` are reused by the Word and Excel chart code. `EMU` = 9525 per CSS px. |
| `render.ts` | `renderSlide()`: one slide as HTML. Shapes are absolutely positioned boxes with an SVG outline and an HTML text layer. Follows placeholder inheritance and theme style references. `topShapes()` / `editableShapes()` number the shapes for edit mode. |
| `drawing.ts` | DrawingML colours, fills, lines, borders and shadows (`colour`, `readFill`, `readLine`, `fillPaint`, `strokeAttrs`, …). |
| `geometry.ts` | Shape outlines as SVG path data: preset shapes (`presetOutline`) and custom geometry (`customOutline`). Unknown presets fall back to a rectangle. |
| `text.ts` | Text bodies: merges paragraph/run styles down the inheritance chain and renders them (`renderText`). `fontFamily()` maps theme fonts. |
| `table.ts` | `renderTable()`: tables and their styles. Built-in styles other than the default are drawn as "Medium Style 2 – Accent 1". |
| `viewer.ts` | `openPptx()`: one slide per screen with scroll-snap, thumbnails, own pinch/double-tap zoom, Present mode, a single tap to hide the bars (`../immersive.ts`; a double tap still zooms). `printPptx()` centres slides on landscape A4. |
| `editor.ts` | `PptxEditor`: text editing in text boxes and placeholders, bold/italic/underline, alignment, font size, autofit, undo/redo, save. |
| `shapes.ts` | `ShapeTool`: select a top-level element, move it, resize it with handles. Writes `off`/`ext`. |
| `pptx.css` | Styles for the viewer, slide markup and editor. |

## How it works

- **Rendering** is lazy: the current slide and its neighbours, thumbnails as they scroll in.
  SmartArt without a saved drawing and chart kinds [`../chart/`](../chart/README.md) can't draw
  show a placeholder and a notice.
- **Present**: in the app the page calls `OfficeBridge.setPresenting(true)` and the shell hides
  its bars and turns landscape; Back calls `window.officeEndPresentation()`. In a browser the
  Fullscreen API is used.
- **Text editing**: edit mode renders slides with marks (`data-sp` per editable shape, `data-p`,
  `data-r`). Typing only changes the screen; after a pause a box's paragraphs are rebuilt into the
  slide's working XML (unchanged ones copied). Leaving a box fits its text and redraws the slide.
  Never redraw while a box is being typed in: Android keyboards compose words, and changing the
  DOM mid-composition misplaces the word.
- **Move/resize**: the first tap selects an element, a second tap on its text types. Tables move
  only. A group's `chOff`/`chExt` stay, so its children scale with it.
- **Saving**: only changed slides are written; their root start tag is kept as in the original.
- Locked: paragraphs with fields or links, tables, charts, SmartArt.

## Dependencies & relationships

- Libraries: `jszip` (through `../ooxml.ts`, and directly in `editor.ts`).
- Uses [`../ooxml.ts`](../ooxml.ts), [`../chart/`](../chart/README.md), [`../save.ts`](../save.ts),
  [`../ui.ts`](../ui.ts), [`../pagebar.ts`](../pagebar.ts) (icon button only),
  [`../docx/fonts.css`](../docx/fonts.css).
- Used by: `../main.ts`, `../print.ts`. `drawing.ts`, `text.ts` and `deck.ts` are also imported by
  `../chart/read.ts`, `../docx/charts.ts` and `../xlsx/workbook.ts`.
- `editor.ts` is loaded on demand by `viewer.ts`.

## Usage notes

- Dev: `?file=/@fs/<repo>/test-files/pptx-sample.pptx`.
- Check a saved deck in PowerPoint with
  [`../../../test-files/pptx_resave.applescript`](../../../test-files/pptx_resave.applescript);
  a file PowerPoint repairs or refuses is a failure.
- To add a preset shape, add its outline to `presetOutline` in `geometry.ts`.

## Known limits

Carried over from the handoff notes written at release 1.1.0. Items fixed since were dropped; the rest were not each rechecked.

Viewer:
- Not drawn: SmartArt without a saved drawing, EMF/WMF/TIFF pictures, video/audio (poster frame
  only), 3-D, most effects (only outer shadows), speaker notes, transitions.
- About 50 common preset shapes are drawn exactly; cloud shapes are approximated; others fall
  back to a rectangle. Custom shapes with computed guides draw with them as 0.
- Text is laid out by the browser, so line breaks differ slightly from PowerPoint, and Aptos
  (Office's current default font) falls back to the system sans-serif.
- About 40–70 ms per slide on the phone for the samples. Not tried on a deck with hundreds of
  slides or large photos.

Text editing:
- Slides change from the thumbnails in edit mode (swiping is off so drags select text).
- New paragraphs show a copy of their neighbour's bullet until the box is left. Autofit uses our
  text layout, so a box near its limit may shrink one step differently from PowerPoint (which
  re-fits on open).
- `spAutoFit` grows only shapes with their own position; one placed by its layout keeps its size.
  Rotated/vertical text is not re-fitted. Shift+Enter makes a new paragraph, not a line break.

Move/resize:
- Only the slide's own top-level elements: layout/master shapes and shapes inside a group can't
  be selected. Elements inside `mc:AlternateContent` (e.g. chartex charts) can't be moved.
  Handles near the slide's edge are clipped by it.
- `spLocks`/`picLocks` `noMove`/`noResize` are honoured.

## Planned

Each step is done when PowerPoint opens the results without repair and its PDF matches ours.

- **Slide operations**: duplicate, delete and reorder from the thumbnail strip, with undo. Must
  keep `presentation.xml`'s `sldIdLst`, its rels, `[Content_Types].xml` and each slide's own
  relationships consistent; a duplicated slide gets its own copies of parts that can't be shared
  (notes, charts).
- **Rest of shape editing**: rotate handle; add text box, picture and basic shapes; delete; bring
  forward / send backward; shapes inside groups on their own.
- Not planned: new slides from layouts, themes, editing tables, charts or SmartArt, animations,
  transitions, speaker notes.
