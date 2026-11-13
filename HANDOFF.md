# Handoff — OfficeEditor

Written 12 Nov 2026; last updated 13 Nov 2026 after release 1.1.0 (charts). Read `CLAUDE.md` first (architecture,
build commands, rules from the PRD); this file is the state of the work and what to do next.

## Where things stand

| Milestone (PRD §8) | State |
| --- | --- |
| M0 Skeleton | Done, verified on the phone. |
| M1 PDF viewer | Working on the phone: scroll, pinch-zoom, text selection, search, thumbnails. Tested on generated samples, one Word-exported PDF and a 700 MB / 1,191-page textbook. |
| M2 DOCX viewer | Working on the phone in Word-style page layout. Tested on one real Word file (`test-files/FallSafe_Essay_v5-2.docx`) and two synthetic ones. |
| M3 XLSX viewer | Done; the owner checked it on the phone (session 2). |
| M4 PPTX viewer | Done; the owner checked it (session 2). |
| M5 PDF editing | Working (session 2): highlight/underline selected text, draw, text boxes, signature (kept for reuse), rotate/delete pages, form filling, undo/redo. Saved files checked with pdf-lib, macOS PDFKit and PDF.js; not yet in Chrome/Acrobat. |
| M6 XLSX editing | Working (session 2): edit values and formulas with recalculation, insert/delete rows and columns, undo/redo. Saved files open in Excel without repair and Excel's recalculated values equal ours. |
| M7 DOCX editing | Working (session 2): text, Enter/Backspace across paragraphs, bold/italic/underline, alignment, headings, bullets/numbering, undo/redo. Saved files open in Word; untouched paragraphs stay byte-identical. |
| M8 Polish | Passed: new document/spreadsheet, Export as PDF, tablet and landscape layouts, release signing, About & licences. Released as 1.0.0. |
| Charts (beyond the PRD) | Done: drawn in Word, PowerPoint and Excel files (`web/src/chart/`), incl. Excel chart sheets and printing. Released as 1.1.0. |
| M9 PPTX text editing | Working (session 3): edit text in text boxes and placeholders (empty ones too), Enter/Backspace across paragraphs, bold/italic/underline, alignment, autofit on leaving a box, undo/redo, save. Checked on the phone (Samsung keyboard) and in the desktop browser; saved files open in PowerPoint without repair and its PDF shows the edits. Untouched slides are byte-identical in the zip. Owner hasn't checked it yet. |
| PPTX present mode (owner's request) | Done (session 3): Present button beside the edit pencil; full screen, landscape, tap to advance, Back or ✕ to end. Checked on the phone. |
| Font size, all editors (owner's request) | Done (session 3): Word/PowerPoint selected text or paragraph (PowerPoint: a selected box), Excel selected cells (new font + cell format in styles.xml, rows grow), PDF new text boxes. The size is typed in (1–400 pt, to the half point). Checked in Word, PowerPoint, Excel; PowerPoint and Excel also on the phone. |
| M10 PPTX slide operations | Planned. |
| M11 PPTX shape editing | Started (session 3, owner's request before M10): move and resize shapes, text boxes, pictures, charts, groups (tables: move only), with undo; checked on the phone and in PowerPoint. Not done: rotate handle, add text box/picture/shapes, delete, bring forward/back, shapes inside groups on their own. |

`main` is pushed to `origin` (GitHub). Releases are tagged `v1.0.0` and `v1.1.0` (versionCode 3). To
release: raise `versionCode`/`versionName` in `app/build.gradle.kts` and `web/package.json`
(`npm version x.y.z --no-git-tag-version`), build `assembleRelease` + `bundleRelease`, check the
signer with `apksigner`, test on the phone, tag `vx.y.z` and push. The GitHub release page and
any Play upload (with `mapping.txt`) are done by the owner.

## What the owner asked for, beyond the PRD

- **DOCX must look like Word**: fixed pages, no reflow, and the whole page zooms. An earlier
  reflowed "reading mode" was replaced at their request; don't bring it back unasked.
- Use the PRD as the guide and the phone (Galaxy S21 FE, Android 16, over `adb`) to build and
  debug. They follow along in short messages and unlock the phone when asked.
- Global instruction: run `./gradlew --stop` after Gradle builds (8 GB Mac).

## Next

- Owner checks M9 on the phone; then M10 (slide operations) and M11, one at a time (below).
- P1 items listed under "Open items" as the owner chooses.

## PPTX editing milestones (beyond the PRD)

The PRD keeps PPTX view-only; the owner asked for editing in three steps, one milestone at a
time, each checked on the phone and by opening saved files in PowerPoint
(`pptx_resave.applescript` re-saves and exports PowerPoint's PDF; a file PowerPoint repairs or
refuses is a failure). Follow `docx/editor.ts`: edit the file's own XML in place, rewrite only
what changed, keep everything else byte for byte, and save through `src/save.ts` and the
shell's Save button. New strings in English and Indonesian.

**M9 — Edit text in existing text boxes** (about 1–2 weeks)
- An edit button (as in the Word viewer) switches the current slide to edit mode. Tapping a
  text box or placeholder puts a caret in it; typing, Enter (new paragraph), Backspace across
  paragraphs.
- Bold, italic, underline and paragraph alignment, from a toolbar like the Word editor's.
- Changed paragraphs are rewritten into the slide XML from their original `a:pPr`/`a:rPr`;
  untouched paragraphs, shapes and slides stay byte-identical. Empty placeholders get text.
- Text keeps fitting as PowerPoint would: honour the box's autofit (`normAutofit` shrink,
  `spAutoFit` grow) as the viewer already does when laying text out.
- Locked (greyed, unchanged, with a notice): text with fields (slide number, date), links,
  text in tables, SmartArt, charts, and shapes inside groups if they prove fragile.
- Undo/redo; leaving with changes asks Save / Discard (the shell already does this).
- Done when: edited decks open in PowerPoint without repair, its PDF matches ours for the
  edited slides, and untouched slides are byte-identical in the zip.

**M10 — Slide operations** (a few days, after M9)
- From the thumbnail strip: duplicate, delete, and reorder (drag, or Move up/down in a menu).
- Keeps `presentation.xml`'s `sldIdLst`, `presentation.xml.rels`, `[Content_Types].xml` and
  each slide's own relationships (layout, pictures, charts, notes) consistent. A duplicated
  slide gets its own copies of parts that can't be shared (notes, charts); pictures can be
  shared. Deleting a slide removes parts nothing else refers to.
- Undo/redo for these operations.
- Done when: PowerPoint opens the results without repair, with slides in the new order and
  duplicated slides' charts/notes independent of the originals.

**M11 — Move, resize and add shapes** (several weeks, after M10)
- Select a shape (tap), with handles: move, resize (corner keeps aspect ratio for pictures),
  rotate. Shapes inside groups move within the group's coordinate space (`chOff`/`chExt`).
- Add: text box, picture (from the system picker; embedded as a new media part with its
  relationship and content type) and a few basic shapes (rectangle, oval, arrow, line).
- Delete a shape; bring forward / send backward.
- Gestures: a tap selects, a drag on a selected shape moves it, a drag elsewhere scrolls, and
  pinch still zooms; selection never fights slide swiping.
- Undo/redo across M9–M11 edits in one history.
- Done when: PowerPoint opens the results without repair and its PDF shows shapes where ours
  does; tested on the phone sideways and on a tablet size (`wm size`).

Not planned (would be M12+): new slides from layouts, themes, editing tables, charts (their
embedded workbooks), SmartArt, animations and transitions, speaker notes.

## Saving (all editors)

Shell: Save button appears when the page reports changes; dialog offers Replace original (only
when the app may write the file: picker files, now persisted with write permission) or Save as
copy (system save picker; the copy then opens). Leaving with changes asks Save / Discard.
The original is copied to `cache/backup/original` before being replaced and restored if the
write fails. Web: `src/save.ts`. Reloading the viewer after a save is the one navigation
`DocumentWebView` allows.

## Open items

**PPTX editing** (`web/src/pptx/editor.ts`)
- Slides change from the thumbnails in edit mode (swiping is off so drags select text).
- New paragraphs show a copy of their neighbour's bullet until the box is left (auto-numbers
  are then redrawn). Autofit uses our text layout, which differs slightly from PowerPoint's, so a
  box near its limit may be shrunk one step differently (PowerPoint re-fits on open).
- spAutoFit grows only shapes with their own position; one placed by its layout keeps its size.
  Rotated/vertical text is not re-fitted. Shift+Enter makes a new paragraph, not a line break.
- Tested on pptx-sample, pptx_chart (empty title placeholders) and on the phone; not yet on a
  large real-world deck.

**PPTX move/resize** (`web/src/pptx/shapes.ts`)
- Since move/resize, a first tap selects a text box and a second tap types (as in PowerPoint
  and Google Slides); before, one tap typed.
- Only the slide's own top-level elements; layout/master shapes and shapes inside a group (the
  group moves as a whole) can't be selected. Elements inside `mc:AlternateContent` (e.g.
  chartex charts) can't be moved. Handles near the slide's edge are clipped by it.
- spLocks/picLocks noMove/noResize are honoured.

**DOCX page layout** (`web/src/docx/`)
- On the test essay 17 of 20 page breaks equal Word's; pages 3, 14 and 16 break one or two lines
  early. Cause: Word compresses spaces in justified lines to fit another word; CSS can't.
- Not handled, and untested: footnotes at the page bottom, multi-column sections, text boxes and
  floating images, `keepLines`, a paragraph or table row taller than a page.
- Page numbers are a guess: any header/footer run that is only digits is replaced with the page
  index (`numberPages` in `viewer.ts`). A year in a footer would be overwritten.
- Only Times New Roman, Arial and Calibri have stand-in fonts, Latin subsets only
  (`scripts/make-docx-fonts.mjs`). Other fonts fall back to the system font and will paginate
  differently.
- Links are inert. Pagination runs synchronously after render; not timed on long documents.
- Only one real Word file tested. The PRD's fidelity warning (unsupported content) belongs to M7.

**XLSX viewer** (`web/src/xlsx/`)
- Values are Excel's cached results. A file written by a library that leaves formulas without
  results (some exporters do) shows those cells empty. No formula engine yet (M6).
- Not shown: pictures and shapes (a notice says so; charts are drawn), conditional formatting, data
  validation, comments, rotated text, row/column default styles on empty cells, hyperlinks.
- Text in a frozen column is cut at the frozen edge instead of spilling into the next column.
- Built-in date/time formats follow en-US, not the device locale (Excel follows its locale).
- Parsing the 5,000-row sample takes ~0.5 s on the phone (open-to-grid ~1.3 s); a DOM parse of a
  much larger sheet will be slow and memory-hungry. The PRD's "read-only beyond 50k cells"
  notice belongs to M6.
- The comparison against Excel (`excel_text_dump.applescript`, then the formats in a browser
  console) matched 9,082 of 9,087 cells; the rest are the Mac's Indonesian time separator.

**PDF editing** (`web/src/pdf/editor.ts`, `write.ts`)
- Not done: merge/split PDFs (P1), moving or deleting one annotation (only undo), editing
  annotations that were already in the file, colour/thickness choice, signature fields
  (signatures are drawn into the page, not digitally signed).
- Text boxes use Helvetica: characters outside Windows-1252 are saved as "?". On a page with
  /Rotate, a text box is written along the page's own axis (looks turned).
- PDFs over 150 MB can't be edited (saving rewrites the whole file in memory). Encrypted PDFs
  can't be saved (message shown).
- Saved files not yet opened in Chrome's or Acrobat's viewer (no PDFium/Acrobat here).

**DOCX editing** (`web/src/docx/editor.ts`)
- The edit view is the page at paper size, scaled to the screen, but not split into pages.
- Locked (shown greyed, kept unchanged): paragraphs with pictures, links, fields (TOC, page
  numbers), comments, tracked changes, text boxes. Paragraphs can't be joined across table
  cells. A warning lists such content before editing.
- Not done (P1): font colour, insert image/table, links, find & replace, word count.
  Lists are one level only; heading look in the edit view is a preview.

**XLSX editing** (`web/src/xlsx/edit.ts`, `calc.ts`)
- Functions not implemented (anything outside calc.ts's list: e.g. INDIRECT, financial
  functions other than PMT, statistical beyond MEDIAN/STDEV, array/dynamic functions like
  FILTER, SORT, UNIQUE) keep the value Excel saved; a notice says so and Excel recalculates
  on open (`fullCalcOnLoad`).
- Not done (P1): cell formatting other than font size, sort/filter, add/rename/delete sheets.
  Font size: a rich-text cell's runs keep their own sizes; rows grow for larger text but never
  shrink; a whole row/column only changes its filled cells. New cells get the
  default style; a date typed into a cell not formatted as a date stays text.
- Inserting/deleting rows or columns: blocked on sheets with Excel tables; pictures/charts do
  not move with the cells; conditional formats/validations in *other* sheets that point at
  the changed sheet are not shifted.
- Numbers with a decimal comma are read as such when the phone's language uses one (id, de…).

**PPTX viewer** (`web/src/pptx/`)
- Not drawn: chart kinds `src/chart/` lacks (waterfall, treemap, sunburst, histogram, funnel,
  stock, surface: placeholder + notice), SmartArt without a saved drawing, EMF/WMF/TIFF
  pictures, video/audio (poster frame only), 3-D, most effects (only outer shadows), speaker
  notes (PRD P1), transitions.
- Built-in table styles other than the default ("Medium Style 2 – Accent 1") are drawn as that
  one: files store only their id.
- About 50 common preset shapes are drawn exactly; cloud shapes are approximated; others fall
  back to a rectangle (`geometry.ts`). Custom shapes with computed guides draw with them as 0.
- Text is laid out by the browser, so line breaks differ slightly from PowerPoint, and Aptos
  (Office's current default font) falls back to the system sans-serif.
- Rendering is lazy (current slide ±2, thumbnails as they scroll in); ~40–70 ms per slide on
  the phone for the samples. Not tried on a deck with hundreds of slides or large photos.

**Charts** (`web/src/chart/`)
- 3-D charts are drawn flat with a depth edge; Office's perspective is not reproduced.
- Axis and label numbers use `.` as the decimal mark whatever the file's locale.
- In Excel, inserting rows/columns doesn't move charts (their anchors aren't shifted).
- Excel box & whisker is untested (Excel's scripting can't make one); PowerPoint and Word
  samples cover chartex.

**PDF viewer** (`web/src/pdf/`)
- Memory grows as a large PDF is read: PDF.js keeps every range it has loaded. Not measured to
  the end of a big file.
- Zoomed pages are capped at about 6 megapixels, so sharpness stops improving past roughly 2×.
- No password prompt (shows a message), no link handling, no annotation layer.
- Range serving reports sizes through `InputStream.available()`, an `int`: files over 2 GB will
  misbehave.

**Shell** (`app/`)
- `resolveDocument` runs on the main thread in `ViewerActivity.onCreate`; a slow provider
  (Drive) could stall the UI.
- Share copies the whole file to cache first; slow for very large files.
- No settings screen (PRD §7).

## Things that cost time and will again

- **Play Protect can reject `adb install`** with `INSTALL_FAILED_VERIFICATION_FAILURE` and hide
  its prompt. The owner turned off "Verify apps over USB" on the phone. If it returns, ask them;
  don't change the setting.
- **The phone locks.** `adb shell dumpsys window | grep isKeyguardShowing`; ask the owner to
  unlock. Never enter the PIN.
- **Files pushed with `adb push` are invisible to the picker** until a rescan:
  `adb shell content call --uri content://media --method scan_volume --arg external_primary`.
- **Opening a file from the shell:** find its MediaStore id, then
  `adb shell am start -a android.intent.action.VIEW --grant-read-uri-permission -n com.officeeditor/.ViewerActivity -t <mime> -d content://media/external/file/<id>`.
  `content://com.android.externalstorage...` URIs are refused for the shell.
- **Driving the WebView:** `adb forward tcp:9222 localabstract:webview_devtools_remote_$(adb shell pidof com.officeeditor)`,
  then the DevTools protocol on `127.0.0.1:9222` (`Runtime.evaluate`,
  `Input.synthesizePinchGesture`). `adb shell input` can't pinch; `uiautomator dump` returned
  nothing useful on this phone.
- **WebView range requests:** see the `EndingAt` comment in `DocumentWebView.kt` before touching
  how documents are served.
- **docx-preview** only breaks pages at explicit breaks, writes Word's line-spacing multiples as
  CSS multiples of font size, hyphenates everything, and misreads `<w:tblHeader/>`. Each has a
  workaround in `web/src/docx/viewer.ts`; read its comments before upgrading the library.
- **Desktop browser checks** are much faster than the phone for layout work:
  `cd web && npm run dev`, then `?file=/@fs/<repo>/test-files/<name>`. Use phone width, and still
  confirm on the phone: one fit-to-width bug only showed there.
- zsh does not word-split unquoted variables; use shell functions for repeated `adb` commands.

## Test files

`excel-sample.xlsx` (4 sheets: budget with shared formulas and frozen panes, formats, a hidden
sheet, hidden rows/columns without gridlines) and `excel-5000-rows.xlsx` were made by Mac Excel
via `make_xlsx_samples.applescript` (Excel's locale here is Indonesian, so currency is "Rp").
Both are on the phone in `Download/officeeditor-test/` (MediaStore ids 2130, 2131).
`pptx-sample.pptx` (python-pptx deck re-saved by PowerPoint: bullets, shapes, picture, table,
group, chart, dark background) and `pptx-native.pptx` (made in PowerPoint, default theme) have
PowerPoint's PDFs beside them; on the phone as ids 2132, 2133.

Chart samples, each with Office's PDF as reference: `pptx_chart.pptx` (pie, 3-D column, line,
clustered column, box & whisker), `docx_chart.docx` (smooth scatter, box & whisker) and
`xlsx_chart.xlsx` (five embedded charts, frozen header, a chart sheet; built by
`make_xlsx_charts.applescript`; Excel's PDF leaves out the chart sheet and misplaces line
markers). All three are in the phone's `Download/`.

`test-files/` is git-ignored apart from `make_samples.py` (run it to regenerate the synthetic
PDFs and DOCX). The owner's essay (`FallSafe_Essay_v5-2.docx`) is there; Word's PDF export of it
is at `~/Downloads/FallSafe_Essay_v5-2.pdf` and is the reference for DOCX pagination. Copies of
the samples and both essay files are on the phone in `Download/officeeditor-test/`.
