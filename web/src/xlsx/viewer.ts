import '../docx/fonts.css';
import './xlsx.css';
import { renderChart } from '../chart';
import { t } from '../i18n';
import { showMessage } from '../message';
import { registerEditor, setDirty } from '../save';
import { menu, toast } from '../ui';
import { numberText } from './calc';
import { chartBox } from './drawings';
import type { SheetEditor } from './edit';
import { formatNumber, formatText, shrinkGeneral } from './format';
import { shiftFormula } from './formula';
import { Workbook, encodeCol, type Cell, type CellFormat, type Font, type Range, type Sheet } from './workbook';

/** Height of the column-letter header row, CSS px at 100%. */
const HDR_H = 20;
/** Space between a cell's edge and its text. */
const PAD = 2;
const MIN_ZOOM = 0.25;
const MAX_ZOOM = 4;
/** Text can spill this many columns beyond the rendered window and still be drawn. */
const SPILL_SCAN = 12;

/**
 * Read-only spreadsheet viewer: sheet tabs, frozen rows and columns, Excel's column widths,
 * number formats, fonts, fills, borders and merged cells. Values are the ones Excel computed
 * and saved; tapping a cell shows its formula.
 */
export async function openXlsx(root: HTMLElement, data: ArrayBuffer): Promise<void> {
  let book: Workbook;
  try {
    book = await Workbook.open(data);
  } catch (e) {
    console.error('xlsx open failed', e);
    showMessage(root, t('xlsxInvalid'));
    return;
  }
  // The grid zooms itself, so the header row and frozen panes stay on screen; the page must not.
  document.querySelector('meta[name=viewport]')
    ?.setAttribute('content', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  await new XlsxViewer(root, book).show(book.activeSheet);
}

/** "Export as PDF": the workbook's print layout (see XlsxViewer.printPages). Returns the page size in CSS px. */
export async function printXlsx(root: HTMLElement, data: ArrayBuffer): Promise<{ width: number; height: number } | null> {
  let book: Workbook;
  try {
    book = await Workbook.open(data);
  } catch (e) {
    console.error('xlsx open failed', e);
    return null;
  }
  const viewer = new XlsxViewer(root, book);
  const { pages, width, height } = await viewer.printPages();
  // The interactive view stays in the page, hidden, for its style sheet of cell formats.
  (root.firstElementChild as HTMLElement).style.display = 'none';
  root.append(...pages);
  return { width, height };
}

/**
 * The cells that print: values, fills, borders and merges. Null for an empty sheet.
 */
function usedRange(sheet: Sheet, formats: CellFormat[]): Range | null {
  let r1 = -1;
  let c1 = -1;
  sheet.rows.forEach((row, r) => {
    row?.forEach((cell, c) => {
      const fmt = formats[cell.xf];
      if (cell.t === 'z' && !fmt?.fill && !fmt?.border) return;
      if (r > r1) r1 = r;
      if (c > c1) c1 = c;
    });
  });
  for (const m of sheet.merges) {
    r1 = Math.max(r1, m.r1);
    c1 = Math.max(c1, m.c1);
  }
  return r1 < 0 ? null : { r0: 0, c0: 0, r1, c1 };
}

interface Geometry {
  /** Left edge of each column (and the right edge of the last), sheet px. */
  colX: Float64Array;
  rowY: Float64Array;
  /** Width of the row-number header. */
  hdrW: number;
  /** Frozen rows / columns. */
  fr: number;
  fc: number;
  /** Width of row header + frozen columns; height of column header + frozen rows. */
  lw: number;
  ht: number;
}

/** Part of the grid drawn into one layer: the frozen corner, top, left, or the scrolling body. */
interface Zone {
  rLo: number;
  rHi: number;
  cLo: number;
  cHi: number;
  x(c: number): number;
  y(r: number): number;
  /** Draws the column letters / row numbers for its columns / rows. */
  colHeads: boolean;
  rowHeads: boolean;
}

interface Display {
  text: string;
  /** Text either side of a format's fill (`*`), pushed to the cell's two edges. */
  split?: [string, string];
  /** Colour from the number format ([Red]), or ''. */
  colour: string;
  kind: 'n' | 's' | 'b' | 'e';
}

interface SheetView {
  zoom: number;
  left: number;
  top: number;
  sel: Range | null;
}

class XlsxViewer {
  private bar = el('div', 'xlsx-bar');
  private barRef = el('div', 'xlsx-ref');
  private barText = el('div', 'xlsx-text');
  /** Cell editor (shown in place of barText while typing). */
  private input = el('input', 'xlsx-input');
  private undoButton = el('button', 'xlsx-bar-button');
  private redoButton = el('button', 'xlsx-bar-button');
  private editor: SheetEditor | null = null;
  /** The editor's text when editing began, to tell whether anything was changed. */
  private editStart = '';
  private notice = el('div', 'xlsx-notice');
  private noticeText = el('span');
  private scroller = el('div', 'xlsx-scroll');
  private message = el('div', 'xlsx-message');
  private content = el('div', 'xlsx-content');
  private top = el('div', 'xlsx-top');
  private corner = el('div', 'xlsx-corner');
  private body = el('div', 'xlsx-body');
  private left = el('div', 'xlsx-left');
  private cornerCells = el('div', 'xlsx-cells');
  private topCells = el('div', 'xlsx-cells');
  private leftCells = el('div', 'xlsx-cells');
  private bodyCells = el('div', 'xlsx-cells');
  /** Charts over each layer of the grid (not redrawn while scrolling). */
  private cornerCharts = el('div', 'xlsx-cells xlsx-charts');
  private topCharts = el('div', 'xlsx-cells xlsx-charts');
  private leftCharts = el('div', 'xlsx-cells xlsx-charts');
  private bodyCharts = el('div', 'xlsx-cells xlsx-charts');
  /** A chart sheet's chart, fitted to the screen. */
  private chartView = el('div', 'xlsx-chart-sheet');
  private tabs = el('div', 'xlsx-tabs');
  private styleEl = document.createElement('style');

  private index = -1;
  private sheet: Sheet | null = null;
  private g!: Geometry;
  private zoom = 1;
  /** Cells currently rendered in the scrolling layers. */
  private win = { r0: 0, r1: -1, c0: 0, c1: -1 };
  private sel: Range | null = null;
  private displays = new WeakMap<Cell, Display>();
  private covered = new Set<number>();
  private views = new Map<number, SheetView>();
  private dismissed = new Set<number>();
  private measureCtx = document.createElement('canvas').getContext('2d')!;
  private fontCss = new WeakMap<Partial<Font>, string>();
  private fontsReady: Promise<void>;
  private frame = 0;

  constructor(root: HTMLElement, private book: Workbook) {
    const view = el('div', 'xlsx');
    this.barText.setAttribute('aria-live', 'polite');
    this.bar.append(this.barRef, this.barText, this.input, this.undoButton, this.redoButton);
    this.barText.onclick = () => void this.startEdit();
    this.input.hidden = true;
    this.input.enterKeyHint = 'done';
    this.input.autocapitalize = 'off';
    this.input.setAttribute('autocomplete', 'off');
    this.input.setAttribute('aria-label', t('cellContents'));
    this.input.onkeydown = (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.commit(1);
      } else if (e.key === 'Escape') this.stopEdit();
    };
    this.input.onblur = () => {
      if (!this.input.hidden) this.commit(0);
    };
    for (const [b, icon, label, action] of [
      [this.undoButton, ICON_UNDO, t('undo'), () => this.history('undo')],
      [this.redoButton, ICON_REDO, t('redo'), () => this.history('redo')],
    ] as const) {
      b.innerHTML = icon;
      b.setAttribute('aria-label', label);
      b.title = label;
      b.hidden = true;
      b.addEventListener('pointerdown', (e) => e.preventDefault());
      b.onclick = action;
    }

    const close = el('button', 'xlsx-notice-close');
    close.textContent = '×';
    close.setAttribute('aria-label', t('close'));
    close.onclick = () => {
      this.dismissed.add(this.index);
      this.notice.hidden = true;
    };
    this.notice.append(this.noticeText, close);
    this.notice.hidden = true;

    this.corner.append(this.cornerCells, this.cornerCharts);
    this.top.append(this.corner, this.topCells, this.topCharts);
    this.left.append(this.leftCells, this.leftCharts);
    this.body.append(this.left, this.bodyCells, this.bodyCharts);
    this.chartView.hidden = true;
    this.content.append(this.top, this.body);
    this.scroller.append(this.content);
    this.message.hidden = true;
    this.tabs.setAttribute('role', 'tablist');
    this.tabs.setAttribute('aria-label', t('sheets'));
    view.append(this.styleEl, this.bar, this.notice, this.scroller, this.message, this.chartView, this.tabs);
    root.replaceChildren(view);

    this.styleEl.textContent = formatRules(book.formats, (f) => this.cssFamily(f));
    this.fontsReady = loadFonts(book.formats.map((f) => this.font(f.font)));
    this.buildTabs();
    this.listen();
  }

  async show(index: number): Promise<void> {
    if (this.sheet && this.index >= 0) {
      this.views.set(this.index, {
        zoom: this.zoom, left: this.scroller.scrollLeft, top: this.scroller.scrollTop, sel: this.sel,
      });
    }
    this.index = index;
    for (const tab of this.tabs.children) {
      const on = Number((tab as HTMLElement).dataset.index) === index;
      tab.setAttribute('aria-selected', String(on));
      if (on) tab.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
    this.sheet = null;
    this.sel = null;
    this.showBar();
    this.notice.hidden = true;

    this.chartView.hidden = true;
    const entry = this.book.sheets[index];
    if (entry.kind === 'chart' && (await this.showChartSheet(index))) return;
    if (entry.kind !== 'worksheet') {
      if (this.index === index) this.showMessage(t(entry.kind === 'chart' ? 'chartSheet' : 'otherSheet'));
      return;
    }
    const started = performance.now();
    let sheet: Sheet;
    try {
      [sheet] = await Promise.all([this.book.sheet(index), this.fontsReady]);
    } catch (e) {
      console.error('sheet load failed', e);
      if (this.index === index) this.showMessage(t('sheetInvalid'));
      return;
    }
    if (this.index !== index) return; // another tab was chosen meanwhile
    this.message.hidden = true;
    this.scroller.hidden = false;
    this.sheet = sheet;
    this.displays = new WeakMap();
    this.layout(sheet);

    const saved = this.views.get(index);
    this.zoom = saved?.zoom ?? sheet.zoom;
    this.content.style.zoom = String(this.zoom);
    this.sel = saved?.sel ?? null;
    this.scroller.scrollLeft = saved?.left ?? 0;
    this.scroller.scrollTop = saved?.top ?? 0;
    this.showBar();
    this.renderAll();
    this.drawCharts();
    console.log(`sheet ${index} (${sheet.rows.length} rows) shown in ${Math.round(performance.now() - started)} ms`);

    this.notice.hidden = !sheet.drawings.others || this.dismissed.has(index);
    this.noticeText.textContent = t('drawingsHidden');
  }

  /** A chart sheet: its chart drawn to fill the screen. False if it has none that can be drawn. */
  private async showChartSheet(index: number): Promise<boolean> {
    let svg: string | null = null;
    try {
      const { charts } = await this.book.chartSheet(index);
      const chart = charts[0];
      if (chart) {
        const size = chart.size ?? { w: 960, h: 640 };
        svg = renderChart(chart.root, this.book.chartTheme, size.w, size.h);
      }
    } catch (e) {
      console.error('chart sheet failed', e);
    }
    if (!svg || this.index !== index) return !!svg;
    this.scroller.hidden = true;
    this.message.hidden = true;
    this.chartView.hidden = false;
    // Drawn at the size Excel saved, then scaled (keeping its shape) to the screen.
    this.chartView.innerHTML = svg.replace(/ width="[^"]*" height="[^"]*"/, ' width="100%" height="100%"');
    return true;
  }

  /** The sheet's charts, each in the layer its top-left corner is in. */
  private drawCharts(): void {
    const sheet = this.sheet!;
    const { colX, rowY, fr, fc } = this.g;
    const z = this.zones();
    const html = { corner: '', top: '', left: '', body: '' };
    for (const chart of sheet.drawings.charts) {
      const box = chartBox(chart, colX, rowY);
      if (!box || box.w < 2 || box.h < 2) continue;
      const svg = renderChart(chart.root, this.book.chartTheme, box.w, box.h);
      if (!svg) continue;
      const row = chart.from?.row ?? 0;
      const col = chart.from?.col ?? 0;
      const layer = row < fr ? (col < fc ? 'corner' : 'top') : col < fc ? 'left' : 'body';
      // Zone positions are offset versions of the sheet's own coordinates.
      const x = z[layer].x(0) + box.x;
      const y = z[layer].y(0) + box.y;
      html[layer] += `<div class="xlsx-chart" style="left:${x}px;top:${y}px;width:${box.w}px;height:${box.h}px">${svg}</div>`;
    }
    this.cornerCharts.innerHTML = html.corner;
    this.topCharts.innerHTML = html.top;
    this.leftCharts.innerHTML = html.left;
    this.bodyCharts.innerHTML = html.body;
  }

  private showMessage(text: string): void {
    this.scroller.hidden = true;
    this.message.hidden = false;
    this.message.textContent = text;
  }

  private buildTabs(): void {
    this.book.sheets.forEach((sheet, i) => {
      if (sheet.hidden) return;
      const tab = el('button', 'xlsx-tab');
      tab.textContent = sheet.name;
      tab.dataset.index = String(i);
      tab.setAttribute('role', 'tab');
      tab.onclick = () => {
        if (i !== this.index) void this.show(i);
      };
      this.tabs.append(tab);
    });
  }

  // ---------------------------------------------------------------- geometry

  private layout(sheet: Sheet): void {
    const nCols = sheet.colWidths.length;
    const nRows = sheet.rowHeights.length;
    const colX = new Float64Array(nCols + 1);
    for (let c = 0; c < nCols; c++) colX[c + 1] = colX[c] + sheet.colWidths[c];
    const rowY = new Float64Array(nRows + 1);
    for (let r = 0; r < nRows; r++) rowY[r + 1] = rowY[r] + sheet.rowHeights[r];
    const fr = Math.min(sheet.frozenRows, nRows);
    const fc = Math.min(sheet.frozenCols, nCols);
    const hdrW = Math.max(28, String(nRows).length * 7 + 12);
    const g: Geometry = { colX, rowY, hdrW, fr, fc, lw: hdrW + colX[fc], ht: HDR_H + rowY[fr] };
    this.g = g;

    const width = g.lw + colX[nCols] - colX[fc];
    const bodyHeight = rowY[nRows] - rowY[fr];
    this.content.style.width = `${width}px`;
    this.content.style.height = `${g.ht + bodyHeight}px`;
    this.top.style.height = `${g.ht}px`;
    this.top.style.width = `${width}px`;
    this.corner.style.width = `${g.lw}px`;
    this.corner.style.height = `${g.ht}px`;
    this.body.style.height = `${bodyHeight}px`;
    this.left.style.width = `${g.lw}px`;
    this.left.style.height = `${bodyHeight}px`;
  }

  private zones(): { corner: Zone; top: Zone; left: Zone; body: Zone } {
    const { colX, rowY, hdrW, fr, fc, lw } = this.g;
    const nRows = rowY.length - 1;
    const nCols = colX.length - 1;
    const frozenX = (c: number) => hdrW + colX[c];
    const scrollX = (c: number) => lw + colX[c] - colX[fc];
    const frozenY = (r: number) => HDR_H + rowY[r];
    const scrollY = (r: number) => rowY[r] - rowY[fr];
    return {
      corner: { rLo: 0, rHi: fr - 1, cLo: 0, cHi: fc - 1, x: frozenX, y: frozenY, colHeads: true, rowHeads: true },
      top: { rLo: 0, rHi: fr - 1, cLo: fc, cHi: nCols - 1, x: scrollX, y: frozenY, colHeads: true, rowHeads: false },
      left: { rLo: fr, rHi: nRows - 1, cLo: 0, cHi: fc - 1, x: frozenX, y: scrollY, colHeads: false, rowHeads: true },
      body: { rLo: fr, rHi: nRows - 1, cLo: fc, cHi: nCols - 1, x: scrollX, y: scrollY, colHeads: false, rowHeads: false },
    };
  }

  /** The scrolling part of the sheet that is on screen, as cell indices. */
  private visible(overscan: number): { r0: number; r1: number; c0: number; c1: number } {
    const { colX, rowY, fr, fc, lw, ht } = this.g;
    const z = this.zoom;
    const vw = this.scroller.clientWidth / z;
    const vh = this.scroller.clientHeight / z;
    const x0 = colX[fc] + this.scroller.scrollLeft / z;
    const y0 = rowY[fr] + this.scroller.scrollTop / z;
    const x1 = x0 + Math.max(0, vw - lw);
    const y1 = y0 + Math.max(0, vh - ht);
    const ox = vw * overscan;
    const oy = vh * overscan;
    return {
      c0: Math.max(fc, find(colX, x0 - ox)),
      c1: find(colX, x1 + ox),
      r0: Math.max(fr, find(rowY, y0 - oy)),
      r1: find(rowY, y1 + oy),
    };
  }

  // ---------------------------------------------------------------- printing

  /**
   * "Export as PDF": every visible worksheet's used range on A4 pages, as Excel prints by
   * default (no gridlines or headings, 0.7"/0.75" margins), scaled down to fit the page width
   * rather than splitting columns across pages. Portrait, or landscape when a sheet would have
   * to shrink below 70% in portrait. Pages break between rows.
   */
  async printPages(): Promise<{ pages: HTMLElement[]; width: number; height: number }> {
    await this.fontsReady;
    const sheets: { index: number; sheet: Sheet; used: Range }[] = [];
    /** Chart sheets, in workbook order with the worksheets (index → chart SVG). */
    const chartPages = new Map<number, { svg: string; w: number; h: number }>();
    for (let index = 0; index < this.book.sheets.length; index++) {
      const entry = this.book.sheets[index];
      if (entry.hidden) continue;
      if (entry.kind === 'chart') {
        const chart = (await this.book.chartSheet(index).catch(() => null))?.charts[0];
        const size = chart?.size ?? { w: 960, h: 640 };
        const svg = chart ? renderChart(chart.root, this.book.chartTheme, size.w, size.h) : null;
        if (svg) chartPages.set(index, { svg, ...size });
        continue;
      }
      if (entry.kind !== 'worksheet') continue;
      const sheet = await this.book.sheet(index);
      let used = usedRange(sheet, this.book.formats);
      // Charts print too: take in the cells they cover.
      for (const chart of sheet.drawings.charts) {
        const end = chart.to ?? chart.from;
        if (!end) continue;
        used ??= { r0: 0, c0: 0, r1: 0, c1: 0 };
        used = { ...used, r1: Math.max(used.r1, end.row), c1: Math.max(used.c1, end.col) };
      }
      if (used) sheets.push({ index, sheet, used });
    }
    const usedWidth = (sheet: Sheet, used: Range) => {
      let w = 0;
      for (let c = 0; c <= used.c1; c++) w += sheet.colWidths[c];
      return w;
    };
    const A4 = { w: 794, h: 1123 };
    const MX = 67;
    const MY = 72;
    const landscape = sheets.some(({ sheet, used }) => (A4.w - 2 * MX) / usedWidth(sheet, used) < 0.7);
    const width = landscape ? A4.h : A4.w;
    const height = landscape ? A4.w : A4.h;

    const pages: HTMLElement[] = [];
    const chartPage = (index: number) => {
      const chart = chartPages.get(index);
      if (!chart) return;
      // The chart scaled to fit inside the margins, centred.
      const scale = Math.min((width - 2 * MX) / chart.w, (height - 2 * MY) / chart.h);
      const page = el('div', 'print-page xlsx-print-page');
      page.style.width = `${width}px`;
      page.style.height = `${height}px`;
      const box = el('div', 'xlsx-cells');
      box.style.cssText = `left:${(width - chart.w * scale) / 2 / scale}px;top:${MY / scale}px;width:${chart.w}px;height:${chart.h}px;zoom:${scale}`;
      box.innerHTML = chart.svg;
      page.append(box);
      pages.push(page);
    };
    let nextChart = 0;
    const chartsBefore = (index: number) => {
      for (; nextChart < index; nextChart++) chartPage(nextChart);
    };
    for (const { index, sheet, used } of sheets) {
      chartsBefore(index);
      this.index = index;
      this.sheet = sheet;
      this.sel = null;
      this.displays = new WeakMap();
      this.layout(sheet);
      const { colX, rowY } = this.g;
      this.addSpill(sheet, used);
      const sheetWidth = colX[used.c1 + 1];
      const scale = Math.min(1, (width - 2 * MX) / sheetWidth);
      const band = (height - 2 * MY) / scale;
      const charts = sheet.drawings.charts
        .map((chart) => ({ chart, box: chartBox(chart, colX, rowY) }))
        .filter((c): c is { chart: typeof c.chart; box: NonNullable<typeof c.box> } => !!c.box && c.box.w > 1 && c.box.h > 1);
      for (let r = 0; r <= used.r1;) {
        let end = r;
        while (end < used.r1 && rowY[end + 2] - rowY[r] <= band) end++;
        // Don't cut a chart in two if starting the page before it lets it fit whole.
        for (const { chart, box } of charts) {
          const from = chart.from?.row ?? 0;
          if (from > r && from <= end && box.y + box.h > rowY[end + 1] && box.h <= band) end = Math.min(end, from - 1);
        }
        const top = rowY[r];
        const zone: Zone = {
          rLo: r, rHi: end, cLo: 0, cHi: used.c1,
          x: (c) => colX[c], y: (row) => rowY[row] - top, colHeads: false, rowHeads: false,
        };
        const cells = el('div', 'xlsx-cells');
        // CSS zoom, not a transform: the print engine shrinks the whole document to fit the
        // widest element's unscaled size, and a transform leaves that size as it was.
        cells.style.cssText = `left:${MX / scale}px;top:${MY / scale}px;width:${sheetWidth}px;height:${rowY[end + 1] - top}px;` +
          `zoom:${scale}`;
        cells.innerHTML = this.render(zone, r, end, 0, used.c1, false);
        // Each chart prints on the page its top edge is on.
        for (const { chart, box } of charts) {
          const from = Math.min(chart.from?.row ?? 0, used.r1);
          if (from < r || from > end) continue;
          const svg = renderChart(chart.root, this.book.chartTheme, box.w, box.h);
          if (svg) cells.insertAdjacentHTML('beforeend', `<div class="xlsx-chart" style="left:${box.x}px;top:${box.y - top}px;width:${box.w}px;height:${box.h}px">${svg}</div>`);
        }
        const page = el('div', 'print-page xlsx-print-page');
        page.style.width = `${width}px`;
        page.style.height = `${height}px`;
        page.append(cells);
        pages.push(page);
        r = end + 1;
      }
    }
    chartsBefore(this.book.sheets.length);
    if (!pages.length) {
      const page = el('div', 'print-page');
      page.style.width = `${width}px`;
      page.style.height = `${height}px`;
      pages.push(page);
    }
    return { pages, width, height };
  }

  /** Widens `used` to take in text that runs on past its last column, as Excel prints it. */
  private addSpill(sheet: Sheet, used: Range): void {
    const { colX } = this.g;
    const last = colX.length - 2;
    sheet.rows.forEach((row) => {
      row?.forEach((cell, c) => {
        if (cell.t !== 's') return;
        const fmt = this.book.formats[cell.xf] ?? this.book.formats[0];
        if (fmt.wrap || (fmt.h !== 'general' && fmt.h !== 'left')) return;
        const end = colX[c] + this.measure(this.display(cell, fmt).text, fmt.font) + 2 * PAD + 1;
        while (used.c1 < last && colX[used.c1 + 1] < end) used.c1++;
      });
    });
  }

  // ---------------------------------------------------------------- rendering

  private update(): void {
    if (!this.sheet) return;
    const v = this.visible(0);
    const w = this.win;
    if (v.r0 >= w.r0 && v.r1 <= w.r1 && v.c0 >= w.c0 && v.c1 <= w.c1) return;
    this.renderScrolling();
  }

  private renderAll(): void {
    const z = this.zones();
    this.cornerCells.innerHTML = this.render(z.corner, 0, z.corner.rHi, 0, z.corner.cHi);
    this.renderScrolling();
  }

  private renderScrolling(): void {
    const w = (this.win = this.visible(0.75));
    const z = this.zones();
    this.topCells.innerHTML = this.render(z.top, 0, z.top.rHi, w.c0, w.c1);
    this.leftCells.innerHTML = this.render(z.left, w.r0, w.r1, 0, z.left.cHi);
    this.bodyCells.innerHTML = this.render(z.body, w.r0, w.r1, w.c0, w.c1);
  }

  /** HTML for rows r0..r1 × columns c0..c1 of one zone, with its headers (and gridlines). */
  private render(zone: Zone, r0: number, r1: number, c0: number, c1: number, gridlines = true): string {
    const sheet = this.sheet!;
    const { colX, rowY, hdrW } = this.g;
    const formats = this.book.formats;
    const out: string[] = [];
    const grid: string[] = [];
    const fills: string[] = [];
    const texts: string[] = [];
    const borders: string[] = [];
    const empty = r1 < r0 || c1 < c0;

    // Column letters along the top, row numbers down the left.
    if (zone.colHeads) {
      for (let c = c0; c <= c1; c++) {
        const w = colX[c + 1] - colX[c];
        if (!w) continue;
        const on = this.sel && c >= this.sel.c0 && c <= this.sel.c1 ? ' on' : '';
        out.push(`<div class="hcol${on}" style="left:${zone.x(c)}px;width:${w}px;height:${HDR_H}px">${encodeCol(c)}</div>`);
      }
    }
    if (zone.rowHeads) {
      for (let r = r0; r <= r1; r++) {
        const h = rowY[r + 1] - rowY[r];
        if (!h) continue;
        const on = this.sel && r >= this.sel.r0 && r <= this.sel.r1 ? ' on' : '';
        out.push(`<div class="hrow${on}" style="top:${zone.y(r)}px;width:${hdrW}px;height:${h}px">${r + 1}</div>`);
      }
    }
    if (zone.colHeads && zone.rowHeads) out.push(`<div class="hx" style="width:${hdrW}px;height:${HDR_H}px"></div>`);
    if (empty) return out.join('');

    const xa = zone.x(c0);
    const xb = zone.x(c1 + 1);
    const ya = zone.y(r0);
    const yb = zone.y(r1 + 1);
    if (sheet.showGrid && gridlines) {
      for (let c = c0; c <= c1; c++) {
        if (colX[c + 1] > colX[c]) grid.push(`<i class="gv" style="left:${zone.x(c + 1) - 1}px;top:${ya}px;height:${yb - ya}px"></i>`);
      }
      for (let r = r0; r <= r1; r++) {
        if (rowY[r + 1] > rowY[r]) grid.push(`<i class="gh" style="top:${zone.y(r + 1) - 1}px;left:${xa}px;width:${xb - xa}px"></i>`);
      }
    }

    // Merged cells: drawn once, by the zone holding their first cell.
    this.covered.clear();
    for (const m of sheet.merges) {
      if (m.r1 < r0 || m.r0 > r1 || m.c1 < c0 || m.c0 > c1) continue;
      for (let r = Math.max(m.r0, r0); r <= Math.min(m.r1, r1); r++) {
        for (let c = Math.max(m.c0, c0); c <= Math.min(m.c1, c1); c++) this.covered.add(r * 16384 + c);
      }
      if (m.r0 < zone.rLo || m.r0 > zone.rHi || m.c0 < zone.cLo || m.c0 > zone.cHi) continue;
      const x = zone.x(m.c0);
      const y = zone.y(m.r0);
      const w = zone.x(Math.min(m.c1, zone.cHi) + 1) - x;
      const h = zone.y(Math.min(m.r1, zone.rHi) + 1) - y;
      if (w <= 0 || h <= 0) continue;
      const cell = sheet.rows[m.r0]?.get(m.c0);
      const fmt = formats[cell?.xf ?? 0] ?? formats[0];
      // A merged area has no gridlines inside it.
      fills.push(fmt.fill ? box('f', x - 1, y - 1, w + 1, h + 1, `background:${fmt.fill}`) : box('f m', x, y, w - 1, h - 1, ''));
      if (cell && cell.t !== 'z') texts.push(this.cellHtml(cell, fmt, m.r0, m.c0, x, y, w, h, zone, false));
    }

    for (let r = r0; r <= r1; r++) {
      const row = sheet.rows[r];
      const h = rowY[r + 1] - rowY[r];
      if (!row || !h) continue;
      const y = zone.y(r);
      // Text from cells just outside the window can spill into it.
      const from = Math.max(zone.cLo, c0 - SPILL_SCAN);
      const to = Math.min(zone.cHi, c1 + SPILL_SCAN);
      for (let c = from; c <= to; c++) {
        const cell = row.get(c);
        if (!cell) continue;
        const w = colX[c + 1] - colX[c];
        if (!w) continue;
        const inside = c >= c0 && c <= c1;
        const fmt = formats[cell.xf] ?? formats[0];
        const x = zone.x(c);
        const merged = this.covered.has(r * 16384 + c);
        if (inside && fmt.border) borders.push(borderBox(fmt.border, x, y, w, h));
        if (merged) continue;
        if (inside && fmt.fill) fills.push(box('f', x - 1, y - 1, w + 1, h + 1, `background:${fmt.fill}`));
        if (cell.t !== 'z' && (inside || cell.t === 's')) texts.push(this.cellHtml(cell, fmt, r, c, x, y, w, h, zone, true));
      }
    }

    const sel = this.sel;
    if (sel && sel.r0 >= zone.rLo && sel.r0 <= zone.rHi && sel.c0 >= zone.cLo && sel.c0 <= zone.cHi) {
      const x = zone.x(sel.c0);
      const y = zone.y(sel.r0);
      const w = zone.x(Math.min(sel.c1, zone.cHi) + 1) - x;
      const h = zone.y(Math.min(sel.r1, zone.rHi) + 1) - y;
      borders.push(box('sel', x - 2, y - 2, w + 3, h + 3, ''));
    }
    return grid.join('') + fills.join('') + texts.join('') + borders.join('') + out.join('');
  }

  private cellHtml(
    cell: Cell, fmt: CellFormat, r: number, c: number,
    x: number, y: number, w: number, h: number, zone: Zone, canSpill: boolean,
  ): string {
    const d = this.display(cell, fmt);
    let align = fmt.h;
    if (align === 'general') align = d.kind === 'n' ? 'right' : d.kind === 's' ? 'left' : 'center';
    const a = align === 'right' ? 'ar'
      : align === 'center' || align === 'centerContinuous' ? 'ac'
      : align === 'justify' || align === 'distributed' ? 'aj'
      : 'al';
    const wrap = fmt.wrap || a === 'aj';
    const indent = fmt.indent ? fmt.indent * 3 * this.book.digitWidth : 0;
    let text = d.text;
    let split = d.split;
    let left = x;
    let width = w;
    let spill = false;

    if (d.kind === 'n') {
      const room = w - 2 * PAD - indent - 1;
      if (this.measure(text, fmt.font) > room) {
        const fits = (s: string) => this.measure(s, fmt.font) <= room;
        const shrunk = fmt.numFmt === 'General' ? shrinkGeneral(cell.v as number, fits) : null;
        text = shrunk ?? '#'.repeat(Math.max(1, Math.floor(room / Math.max(1, this.measure('#', fmt.font)))));
        split = undefined;
      }
    } else if (d.kind === 's' && canSpill && !wrap && !text.includes('\n')) {
      // Text wider than its cell runs on over empty neighbours, as in Excel.
      const need = this.measure(text, fmt.font) + 2 * PAD + indent + 1;
      if (need > w) {
        const { colX } = this.g;
        const free = (col: number) => col >= zone.cLo && col <= zone.cHi && this.isEmpty(r, col);
        let lo = c;
        let hi = c;
        if (a === 'ac') {
          while (colX[hi + 1] - colX[lo] < need) {
            const grew = free(lo - 1) && free(hi + 1);
            if (!grew) break;
            lo--;
            hi++;
          }
        } else if (a === 'ar') {
          while (colX[hi + 1] - colX[lo] < need && free(lo - 1)) lo--;
        } else {
          while (colX[hi + 1] - colX[lo] < need && free(hi + 1)) hi++;
        }
        if (lo !== c || hi !== c) {
          left = zone.x(lo);
          width = zone.x(hi + 1) - left;
          spill = true;
        }
      }
    }
    if (!wrap) text = text.replace(/\r?\n/g, ' ');

    let style = `left:${left}px;top:${y}px;width:${width - 1}px;height:${h - 1}px`;
    if (d.colour) style += `;color:${d.colour}`;
    if (indent) style += a === 'ar' ? `;padding-right:${indent + PAD}px` : `;padding-left:${indent + PAD}px`;
    if (spill && !fmt.fill) style += ';background:#fff';
    const v = fmt.v === 'top' ? 'vt' : fmt.v === 'center' ? 'vc' : 'vb';
    const content = cell.rich && d.kind === 's'
      ? cell.rich.map((run) => (run.font ? `<span style="${this.runCss(run.font)}">${escape(run.text)}</span>` : escape(run.text))).join('')
      : split ? `<span>${escape(split[0])}</span><span>${escape(split[1])}</span>`
      : escape(text);
    return `<div class="c x${cell.xf} ${split ? 'as' : a} ${v}${wrap ? ' w' : ''}" style="${style}">${content}</div>`;
  }

  private isEmpty(r: number, c: number): boolean {
    const cell = this.sheet!.rows[r]?.get(c);
    if (cell && cell.t !== 'z') return false;
    if (this.covered.has(r * 16384 + c)) return false;
    return !this.sheet!.merges.some((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1);
  }

  private display(cell: Cell, fmt: CellFormat): Display {
    let d = this.displays.get(cell);
    if (!d) {
      if (cell.t === 'n') {
        const f = formatNumber(cell.v as number, fmt.numFmt, this.book.date1904);
        d = { text: f.text, colour: f.colour, kind: 'n', split: f.split };
      } else if (cell.t === 'b') d = { text: cell.v ? 'TRUE' : 'FALSE', colour: '', kind: 'b' };
      else if (cell.t === 'e') d = { text: String(cell.v), colour: '', kind: 'e' };
      else d = { text: formatText(String(cell.v ?? ''), fmt.numFmt), colour: '', kind: 's' };
      this.displays.set(cell, d);
    }
    return d;
  }

  // ---------------------------------------------------------------- fonts

  private cssFamily(font: Partial<Font>): string {
    const name = (font.name ?? '').replace(/['"\\]/g, '');
    const fallback = /narrow|condensed/i.test(name) ? "'Arial Narrow', sans-serif-condensed, 'Roboto Condensed', sans-serif"
      : font.family === 1 ? 'serif'
      : font.family === 3 ? 'monospace'
      : 'sans-serif';
    return name ? `'${name}', ${fallback}` : fallback;
  }

  /** CSS `font` shorthand for a font (for measuring and preloading). */
  private font(font: Font): string {
    let css = this.fontCss.get(font);
    if (!css) {
      css = `${font.italic ? 'italic ' : ''}${font.bold ? '700 ' : ''}${(font.size * 96) / 72}px ${this.cssFamily(font)}`;
      this.fontCss.set(font, css);
    }
    return css;
  }

  private measure(text: string, font: Font): number {
    this.measureCtx.font = this.font(font);
    return this.measureCtx.measureText(text).width;
  }

  private runCss(font: Partial<Font>): string {
    let css = this.fontCss.get(font);
    if (css === undefined) {
      css = fontRule(font, (f) => this.cssFamily(f));
      this.fontCss.set(font, css);
    }
    return css;
  }

  // ---------------------------------------------------------------- selection

  private showBar(): void {
    const sel = this.sel;
    const cell = sel && this.sheet?.rows[sel.r0]?.get(sel.c0);
    this.barRef.textContent = sel ? `${encodeCol(sel.c0)}${sel.r0 + 1}` : '';
    let text = '';
    if (cell) {
      const formula = this.formula(cell, sel.r0, sel.c0);
      if (formula) text = formula;
      else if (cell.t !== 'z') text = this.display(cell, this.book.formats[cell.xf] ?? this.book.formats[0]).text;
    }
    this.barText.textContent = text;
    this.barText.classList.remove('open');
  }

  private formula(cell: Cell, r: number, c: number): string {
    let f = cell.f;
    if (f === undefined && cell.si !== undefined) {
      const master = this.sheet!.sharedFormulas.get(cell.si);
      if (master) f = shiftFormula(master.f, r - master.r, c - master.c);
    }
    if (f === undefined) return '';
    return cell.array ? `{=${f}}` : `=${f}`;
  }

  private select(clientX: number, clientY: number): void {
    if (!this.sheet) return;
    const { colX, rowY, hdrW, fr, fc, lw, ht } = this.g;
    const rect = this.scroller.getBoundingClientRect();
    const z = this.zoom;
    const vx = (clientX - rect.left) / z;
    const vy = (clientY - rect.top) / z;
    if (vx < hdrW && vy < HDR_H) return;
    if (vx < hdrW || vy < HDR_H) {
      // A row number or column letter: insert or delete that row/column.
      const axis = vx < hdrW ? 'row' : 'col';
      const at = axis === 'row'
        ? (vy < ht ? find(rowY, vy - HDR_H) : find(rowY, rowY[fr] + this.scroller.scrollTop / z + vy - ht))
        : (vx < lw ? find(colX, vx - hdrW) : find(colX, colX[fc] + this.scroller.scrollLeft / z + vx - lw));
      void this.headerMenu(axis, at);
      return;
    }
    const c = vx < lw ? find(colX, vx - hdrW) : find(colX, colX[fc] + this.scroller.scrollLeft / z + vx - lw);
    const r = vy < ht ? find(rowY, vy - HDR_H) : find(rowY, rowY[fr] + this.scroller.scrollTop / z + vy - ht);
    const merge = this.sheet.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1);
    this.sel = merge ?? { r0: r, c0: c, r1: r, c1: c };
    this.showBar();
    this.renderAll();
  }

  // ---------------------------------------------------------------- editing

  private async ensureEditor(): Promise<SheetEditor> {
    if (this.editor) return this.editor;
    const { SheetEditor } = await import('./edit');
    this.editor = await SheetEditor.open(this.book);
    const editor = this.editor;
    registerEditor({
      fileName: 'workbook.xlsx',
      produce: () => editor.write(),
      saved: (result) => {
        if (result.ok && result.mode === 'overwrite') location.reload();
      },
    });
    return editor;
  }

  private async startEdit(): Promise<void> {
    const sel = this.sel;
    if (!sel || !this.sheet) {
      toast(t('tapCellFirst'));
      return;
    }
    const editor = await this.ensureEditor();
    const cell = this.sheet.rows[sel.r0]?.get(sel.c0);
    const shown = cell && cell.t !== 'z' ? this.display(cell, this.book.formats[cell.xf] ?? this.book.formats[0]).text : '';
    this.input.value = this.editStart = editText(cell, editor.formula(this.index, sel.r0, sel.c0), shown);
    this.barText.hidden = true;
    this.input.hidden = false;
    this.input.focus();
    this.input.setSelectionRange(this.input.value.length, this.input.value.length);
  }

  private stopEdit(): void {
    this.input.hidden = true;
    this.barText.hidden = false;
    this.input.blur();
  }

  /** Puts the typed value in the cell; `down` moves the selection on, as Enter does in Excel. */
  private commit(down: number): void {
    const sel = this.sel;
    const editor = this.editor;
    if (!sel || !editor || this.input.hidden) return;
    const typed = this.input.value;
    this.input.hidden = true;
    this.barText.hidden = false;
    if (typed !== this.editStart) {
      try {
        editor.setCell(this.index, sel.r0, sel.c0, typed);
      } catch (e) {
        if ((e as Error).message === 'formula') {
          toast(t('formulaError'));
          this.input.hidden = false;
          this.barText.hidden = true;
          this.input.focus();
          return;
        }
        throw e;
      }
    }
    if (down && this.sheet) {
      const r = Math.min(this.sheet.rowHeights.length - 1, sel.r1 + down);
      this.sel = { r0: r, r1: r, c0: sel.c0, c1: sel.c0 };
    }
    this.afterEdit(false);
  }

  private async headerMenu(axis: 'row' | 'col', at: number): Promise<void> {
    if (!this.sheet) return;
    const nCols = this.sheet.colWidths.length;
    const nRows = this.sheet.rowHeights.length;
    this.sel = axis === 'row' ? { r0: at, r1: at, c0: 0, c1: nCols - 1 } : { r0: 0, r1: nRows - 1, c0: at, c1: at };
    this.showBar();
    this.renderAll();
    const label = axis === 'row' ? t('rowN', { n: at + 1 }) : t('columnN', { n: encodeCol(at) });
    const choice = await menu(label, axis === 'row'
      ? [['before', t('insertAbove')], ['after', t('insertBelow')], ['delete', t('deleteRow')]]
      : [['before', t('insertLeft')], ['after', t('insertRight')], ['delete', t('deleteColumn')]]);
    if (!choice) return;
    const editor = await this.ensureEditor();
    if (!editor.canShift(this.index)) {
      toast(t('tablesNoShift'));
      return;
    }
    if (choice === 'delete') editor.shift(this.index, axis, at, -1);
    else editor.shift(this.index, axis, choice === 'before' ? at : at + 1, 1);
    this.sel = null;
    this.afterEdit(true);
  }

  private history(which: 'undo' | 'redo'): void {
    const editor = this.editor;
    if (!editor) return;
    const op = which === 'undo' ? editor.undo() : editor.redo();
    if (!op) return;
    this.afterEdit(op.kind === 'shift');
  }

  /** Redraws after an edit; `layout` when rows or columns were added or removed. */
  private afterEdit(layout: boolean): void {
    const editor = this.editor;
    if (!editor || !this.sheet) return;
    if (layout) this.layout(this.sheet);
    this.showBar();
    this.renderAll();
    if (layout) this.drawCharts();
    setDirty(editor.changed);
    this.undoButton.hidden = this.redoButton.hidden = false;
    this.undoButton.disabled = !editor.canUndo;
    this.redoButton.disabled = !editor.canRedo;
    if (editor.stale.size) {
      this.noticeText.textContent = t('formulasStale');
      this.notice.hidden = false;
    }
  }

  // ---------------------------------------------------------------- input

  private listen(): void {
    this.scroller.addEventListener('scroll', () => this.schedule(), { passive: true });
    window.addEventListener('resize', () => {
      if (this.sheet) this.renderAll();
    });
    this.scroller.addEventListener('click', (e) => this.select(e.clientX, e.clientY));
    this.scroller.addEventListener('dblclick', () => void this.startEdit());

    // Pinch to zoom, around the point between the fingers.
    let pinch: { dist: number; zoom: number } | null = null;
    const dist = (t: TouchList) => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY);
    this.scroller.addEventListener('touchstart', (e) => {
      if (e.touches.length !== 2) return;
      // Otherwise the browser takes the two-finger gesture and stops sending touch moves.
      if (e.cancelable) e.preventDefault();
      pinch = { dist: dist(e.touches) || 1, zoom: this.zoom };
    }, { passive: false });
    this.scroller.addEventListener('touchmove', (e) => {
      if (!pinch || e.touches.length !== 2) return;
      if (e.cancelable) e.preventDefault();
      const rect = this.scroller.getBoundingClientRect();
      const fx = (e.touches[0].clientX + e.touches[1].clientX) / 2 - rect.left;
      const fy = (e.touches[0].clientY + e.touches[1].clientY) / 2 - rect.top;
      this.setZoom((pinch.zoom * dist(e.touches)) / pinch.dist, fx, fy);
    }, { passive: false });
    const end = (e: TouchEvent) => {
      if (e.touches.length < 2) pinch = null;
    };
    this.scroller.addEventListener('touchend', end, { passive: true });
    this.scroller.addEventListener('touchcancel', end, { passive: true });
    // Desktop: ctrl + wheel (and trackpad pinch).
    this.scroller.addEventListener('wheel', (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      const rect = this.scroller.getBoundingClientRect();
      this.setZoom(this.zoom * Math.exp(-e.deltaY / 200), e.clientX - rect.left, e.clientY - rect.top);
    }, { passive: false });
  }

  private pendingZoom: { z: number; fx: number; fy: number } | null = null;

  private setZoom(z: number, fx: number, fy: number): void {
    this.pendingZoom = { z: Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z)), fx, fy };
    this.schedule();
  }

  private schedule(): void {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      const pz = this.pendingZoom;
      this.pendingZoom = null;
      if (pz && this.sheet && pz.z !== this.zoom) {
        // Keep the sheet point under the fingers where it is.
        const s = this.scroller;
        const px = (s.scrollLeft + pz.fx) / this.zoom;
        const py = (s.scrollTop + pz.fy) / this.zoom;
        this.zoom = pz.z;
        this.content.style.zoom = String(pz.z);
        s.scrollLeft = px * pz.z - pz.fx;
        s.scrollTop = py * pz.z - pz.fy;
        this.renderScrolling();
        return;
      }
      this.update();
    });
  }
}

// ------------------------------------------------------------------ helpers

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className = ''): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  return e;
}

/** Index of the column/row whose span [edges[i], edges[i+1]) contains pos, clamped. */
function find(edges: Float64Array, pos: number): number {
  const n = edges.length - 1;
  let lo = 0;
  let hi = n - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (edges[mid] <= pos) lo = mid;
    else hi = mid - 1;
  }
  return Math.max(0, lo);
}

function box(cls: string, x: number, y: number, w: number, h: number, style: string): string {
  return `<i class="${cls}" style="left:${x}px;top:${y}px;width:${w}px;height:${h}px${style ? `;${style}` : ''}"></i>`;
}

const BORDER: Record<string, string> = {
  thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed',
  dashDotDot: '1px dotted', medium: '2px solid', mediumDashed: '2px dashed', mediumDashDot: '2px dashed',
  mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed', thick: '3px solid', double: '3px double',
};

/** A cell's borders, drawn over the gridlines around it. */
function borderBox(border: NonNullable<CellFormat['border']>, x: number, y: number, w: number, h: number): string {
  let css = '';
  for (const side of ['left', 'right', 'top', 'bottom'] as const) {
    const edge = border[side];
    if (edge) css += `border-${side}:${BORDER[edge.style] ?? '1px solid'} ${edge.colour};`;
  }
  return box('b', x - 1, y - 1, w + 1, h + 1, css);
}

function fontRule(font: Partial<Font>, family: (f: Partial<Font>) => string): string {
  let css = '';
  if (font.name !== undefined) css += `font-family:${family(font)};`;
  if (font.size !== undefined) css += `font-size:${font.size}pt;`;
  if (font.bold !== undefined) css += `font-weight:${font.bold ? 700 : 400};`;
  if (font.italic !== undefined) css += `font-style:${font.italic ? 'italic' : 'normal'};`;
  if (font.colour) css += `color:${font.colour};`;
  if (font.underline !== undefined || font.strike !== undefined) {
    const lines = [font.underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ');
    css += `text-decoration:${lines || 'none'}${font.underline === 'double' ? ' double' : ''};`;
  }
  return css;
}

/** One CSS class per cell format (`.x12`) for its font. */
function formatRules(formats: CellFormat[], family: (f: Partial<Font>) => string): string {
  const byFont = new Map<Font, number[]>();
  formats.forEach((f, i) => {
    const list = byFont.get(f.font);
    if (list) list.push(i);
    else byFont.set(f.font, [i]);
  });
  let css = '';
  for (const [font, ids] of byFont) {
    css += `${ids.map((i) => `.xlsx-cells .x${i}`).join(',')}{${fontRule({ ...font, colour: font.colour || '#000' }, family)}}\n`;
  }
  return css;
}

async function loadFonts(fonts: string[]): Promise<void> {
  const unique = [...new Set(fonts)];
  const loads = Promise.all(unique.map((f) => document.fonts.load(f).catch(() => [])));
  // Text is measured to decide where it overflows; wait (briefly) for the real fonts.
  await Promise.race([loads, new Promise((resolve) => setTimeout(resolve, 1500))]);
}

// Material Symbols paths (Apache-2.0).
const svgIcon = (d: string) => `<svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><path fill="currentColor" d="${d}"/></svg>`;
const ICON_UNDO = svgIcon('M12.5 8c-2.65 0-5.05.99-6.9 2.6L2 7v9h9l-3.62-3.62c1.39-1.16 3.16-1.88 5.12-1.88 3.54 0 6.55 2.31 7.6 5.5l2.37-.78C21.08 11.03 17.15 8 12.5 8z');
const ICON_REDO = svgIcon('M18.4 10.6C16.55 8.99 14.15 8 11.5 8c-4.65 0-8.58 3.03-9.96 7.22L3.9 16c1.05-3.19 4.05-5.5 7.6-5.5 1.95 0 3.73.72 5.12 1.88L13 16h9V7l-3.6 3.6z');

/** The cell's contents as they'd be typed into it. */
function editText(cell: Cell | undefined, formula: string | undefined, formatted: string): string {
  if (formula !== undefined) return `=${formula}`;
  if (!cell || cell.t === 'z') return '';
  if (cell.t === 'n') return formatted && /[/:-]|[A-Za-z]{3}/.test(formatted) ? formatted : numberText(cell.v as number);
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
  return String(cell.v ?? '');
}

function escape(s: string): string {
  return s.replace(/[&<>]/g, (ch) => (ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : '&gt;'));
}
