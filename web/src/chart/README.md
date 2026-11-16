# web/src/chart — chart drawing

## Purpose

Draws Office charts as SVG for all three Office viewers (Word, Excel, PowerPoint), from the values
Office caches inside each chart part. Nothing is recalculated.

## Contents

| File | What it does |
| --- | --- |
| `index.ts` | `renderChart(root, theme, w, h)`: the only entry point. Returns SVG markup, or `null` if the chart can't be drawn (the host then shows its placeholder). |
| `read.ts` | `readChart()`: parses a chart part into a plain model (`Chart`, `Group`, `Series`, `Axis`, …). Handles DrawingML charts (bar/column incl. 3-D, line, area, pie, doughnut, scatter, bubble, radar) and chartex (box & whisker only). |
| `draw.ts` | `drawChart()`: lays the model out like Office (title, legend, axis labels, "nice" axis scale) and emits SVG. |

## How it works

`read.ts` → model → `draw.ts`. The SVG has no ids, so copies of it (page thumbnails) stay valid.
3-D charts are drawn flat with a depth edge. Axis numbers always use `.` as the decimal mark.

## Dependencies & relationships

- Uses [`../ooxml.ts`](../ooxml.ts), colours from [`../pptx/drawing.ts`](../pptx/drawing.ts), fonts
  from [`../pptx/text.ts`](../pptx/text.ts), number formats from
  [`../xlsx/format.ts`](../xlsx/format.ts).
- Used by: [`../pptx/render.ts`](../pptx/render.ts) (graphic frames),
  [`../docx/charts.ts`](../docx/charts.ts), [`../xlsx/viewer.ts`](../xlsx/viewer.ts) (overlay
  layers and chart sheets).

## Usage notes

- To support a new chart kind: add it to `GroupType` and the reader in `read.ts`, then a drawing
  branch in `draw.ts`.
- Samples with Office's own PDF as reference: `pptx_chart.pptx`, `docx_chart.docx`,
  `xlsx_chart.xlsx` in `test-files/` (git-ignored; see the [root README](../../../README.md#testing)).

## Known limits

Carried over from the handoff notes written at release 1.1.0. Items fixed since were dropped; the rest were not each rechecked.

- Not drawn (placeholder + notice): waterfall, treemap, sunburst, histogram, funnel, stock,
  surface.
- 3-D charts are drawn flat with a depth edge; Office's perspective is not reproduced.
- Axis and label numbers use `.` as the decimal mark whatever the file's locale.
- In Excel, inserting rows/columns doesn't move charts (their anchors aren't shifted).
- Excel box & whisker is untested (Excel's scripting can't make one); the PowerPoint and Word
  samples cover chartex.
