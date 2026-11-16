# web/src/xlsx — spreadsheet viewer and editor

## Purpose

Shows `.xlsx` workbooks as a scrolling grid and edits cell values and formulas with our own
formula engine. Loaded as its own chunk by [`../main.ts`](../main.ts).

## Contents

| File | What it does |
| --- | --- |
| `workbook.ts` | `Workbook`: reads the zip's XML directly (cached formula results, styles, theme colours, widths, heights, merges, frozen panes). Also column helpers (`encodeCol`, `decodeCol`) and unit helpers. |
| `format.ts` | Number and date formats with `ssf` (`formatNumber`, `formatText`, `isDateFormat`, `shrinkGeneral`). |
| `formula.ts` | `shiftFormula()`: expands a shared formula for a cell by moving its relative references. |
| `drawings.ts` | `readDrawings()`: reads a sheet's drawing part for chart anchors; `chartBox()` turns an anchor into pixels. |
| `viewer.ts` | `openXlsx()`: the grid (only cells near the screen are drawn), sticky headers and frozen panes, sheet tabs, own pinch-zoom, chart overlays, and the editing UI (cell bar, row/column menus). `printXlsx()` lays out each sheet's used range on A4. |
| `calc.ts` | The formula engine: `parse`, `evaluate`, ~120 functions, and `shiftReferences` for row/column insert and delete. Unknown functions throw `Unsupported`. |
| `edit.ts` | `SheetEditor`: holds all sheets in memory, recalculates dependents, inserts/deletes rows and columns, font size, undo/redo, and writes back into the sheet XML. |
| `xlsx.css` | Styles for the grid, tabs and edit bar. The sheet is always light, as in Excel. |

## How it works

- **Viewing** uses the values Excel cached in the file; nothing is recalculated until an edit.
- **Editing**: cells are never changed in place. An edit puts a new `Cell` object in the sheet,
  so the viewer's per-cell caches stay right and undo can put the old one back.
- **Saving**: each worksheet's XML is updated only where the model changed. Values are written as
  inline strings, `fullCalcOnLoad` is set and `calcChain.xml` is removed, so Excel recalculates on
  open. A formula with a function `calc.ts` lacks keeps Excel's saved value.
- Sheets with Excel tables can't have rows or columns inserted or deleted.
- Column widths use Excel's digit width of the default font (`DIGIT_RATIO` in `workbook.ts`).
- Zoom is CSS `zoom` driven by the viewer; page zoom is off for this format.

## Dependencies & relationships

- Libraries: `jszip`, `ssf`.
- Uses [`../ooxml.ts`](../ooxml.ts), [`../chart/`](../chart/README.md), [`../save.ts`](../save.ts),
  [`../ui.ts`](../ui.ts), `readTheme` from [`../pptx/deck.ts`](../pptx/deck.ts),
  [`../docx/fonts.css`](../docx/fonts.css).
- Used by: `../main.ts`, `../print.ts`; `format.ts` is used by [`../chart/read.ts`](../chart/read.ts).
- `edit.ts` is loaded on demand by `viewer.ts` when editing starts.

## Usage notes

- Dev: `?file=/@fs/<repo>/test-files/excel-sample.xlsx`.
- To compare displayed text against Excel, see
  [`../../../test-files/excel_text_dump.applescript`](../../../test-files/excel_text_dump.applescript).
- To add a function, add it to the `FUNCTIONS` table in `calc.ts`.

## Known limits

Carried over from the handoff notes written at release 1.1.0. Items fixed since were dropped; the rest were not each rechecked.

Viewer:
- A file written by a library that leaves formulas without cached results shows those cells
  empty until edited.
- Not shown: pictures and shapes (a notice says so; charts are drawn), conditional formatting,
  data validation, comments, rotated text, row/column default styles on empty cells, hyperlinks.
- Text in a frozen column is cut at the frozen edge instead of spilling into the next column.
- Built-in date/time formats follow en-US, not the device locale.
- Parsing the 5,000-row sample took about 0.5 s on the phone; a DOM parse of a much larger sheet
  will be slow and memory-hungry.
- Compared against Excel with `excel_text_dump.applescript`, 9,082 of 9,087 cells matched; the
  rest were the Mac's Indonesian time separator.

Editor:
- Functions outside `calc.ts`'s list (e.g. INDIRECT, financial functions other than PMT,
  statistical beyond MEDIAN/STDEV, dynamic arrays like FILTER, SORT, UNIQUE) keep the value Excel
  saved; a notice says so and Excel recalculates on open.
- Not done: cell formatting other than font size, sort/filter, add/rename/delete sheets.
- Font size: a rich-text cell's runs keep their own sizes; rows grow for larger text but never
  shrink; a whole row/column only changes its filled cells.
- New cells get the default style; a date typed into a cell not formatted as a date stays text.
- Inserting/deleting rows or columns: pictures and charts do not move with the cells;
  conditional formats and validations in *other* sheets that point at the changed sheet are not
  shifted.
- Numbers with a decimal comma are read as such when the phone's language uses one (id, de…).
