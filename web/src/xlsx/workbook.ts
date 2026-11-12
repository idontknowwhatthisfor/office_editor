import { Package, all, child, children, first, relAttr } from '../ooxml';
import { formatCode } from './format';

// Reads what a viewer needs from an XLSX file: cell values as Excel last computed them (cached
// in the file next to each formula), formats, column widths, row heights, merges and frozen
// panes. Formulas are not recalculated; Excel always saves their results.

export interface Font {
  name: string;
  /** Points. */
  size: number;
  bold: boolean;
  italic: boolean;
  underline: '' | 'single' | 'double';
  strike: boolean;
  /** CSS colour, or '' for the default (black). */
  colour: string;
  /** OOXML font family class: 1 roman (serif), 2 swiss (sans), 3 modern (mono)… */
  family: number;
}

export interface BorderEdge {
  style: string;
  colour: string;
}

export interface CellFormat {
  numFmt: string;
  font: Font;
  /** CSS background, or ''. */
  fill: string;
  border: { left?: BorderEdge; right?: BorderEdge; top?: BorderEdge; bottom?: BorderEdge } | null;
  /** Excel's horizontal alignment ('general' follows the value type). */
  h: string;
  v: 'top' | 'center' | 'bottom';
  wrap: boolean;
  indent: number;
}

export interface Run {
  text: string;
  /** What the run sets; the rest comes from the cell's font. */
  font?: Partial<Font>;
}

export interface Cell {
  /** n number, s text, b boolean, e error, z no value (formatting only). */
  t: 'n' | 's' | 'b' | 'e' | 'z';
  v: number | string | boolean | null;
  /** Index into Workbook.formats. */
  xf: number;
  /** Rich text runs, for text with mixed formatting. */
  rich?: Run[];
  /** Formula text without the leading '='. */
  f?: string;
  /** Shared-formula group this cell belongs to (its formula is the group's, shifted). */
  si?: number;
  array?: boolean;
}

export interface Range {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export interface Sheet {
  name: string;
  /** Cells by row, then column (both 0-based). */
  rows: (Map<number, Cell> | undefined)[];
  /** CSS px per column / row; 0 when hidden. */
  colWidths: Float64Array;
  rowHeights: Float64Array;
  frozenRows: number;
  frozenCols: number;
  merges: Range[];
  showGrid: boolean;
  /** Sheet zoom as saved (1 = 100%). */
  zoom: number;
  /** The sheet has charts, pictures or shapes (not shown). */
  hasDrawings: boolean;
  /** Master formulas of shared-formula groups. */
  sharedFormulas: Map<number, { f: string; r: number; c: number }>;
}

export interface SheetEntry {
  name: string;
  hidden: boolean;
  /** 'worksheet', or 'chart' / 'other' for sheets that hold no cells. */
  kind: 'worksheet' | 'chart' | 'other';
  path: string;
}

/** Empty rows and columns shown past the last used cell, as a spreadsheet shows more grid. */
const PAD_ROWS = 30;
const PAD_COLS = 8;
/** Enough grid to fill the screen when zoomed out. */
const MIN_ROWS = 200;
const MIN_COLS = 30;

export class Workbook {
  sheets: SheetEntry[] = [];
  activeSheet = 0;
  formats: CellFormat[] = [];
  date1904 = false;
  /** Excel's "maximum digit width" of the default font in px: the unit of column widths. */
  digitWidth = 7;
  defaultFont!: Font;
  private pkg!: Package;
  private strings: (string | Run[])[] = [];
  private loaded = new Map<number, Promise<Sheet>>();

  static async open(data: ArrayBuffer): Promise<Workbook> {
    const book = new Workbook();
    book.pkg = await Package.open(data);
    await book.readWorkbook();
    return book;
  }

  sheet(index: number): Promise<Sheet> {
    let sheet = this.loaded.get(index);
    if (!sheet) {
      sheet = this.readSheet(this.sheets[index]);
      this.loaded.set(index, sheet);
    }
    return sheet;
  }

  private xml(path: string): Promise<Document | null> {
    return this.pkg.xml(path);
  }

  private async rels(path: string) {
    return new Map([...(await this.pkg.rels(path))].filter(([, rel]) => !rel.external));
  }

  private async readWorkbook(): Promise<void> {
    const root = await this.rels('_rels/.rels');
    const bookPath = [...root.values()].find((r) => r.type.endsWith('/officeDocument'))?.target
      ?? 'xl/workbook.xml';
    const book = await this.xml(bookPath);
    if (!book) throw new Error('no workbook part');
    const rels = await this.rels(bookPath);

    const pr = first(book, 'workbookPr');
    this.date1904 = /^(1|true)$/.test(pr?.getAttribute('date1904') ?? '');

    for (const el of all(book, 'sheet')) {
      const rel = rels.get(relAttr(el));
      if (!rel) continue;
      const kind = rel.type.endsWith('/worksheet') ? 'worksheet'
        : rel.type.endsWith('/chartsheet') ? 'chart'
        : 'other';
      this.sheets.push({
        name: el.getAttribute('name') ?? `Sheet${this.sheets.length + 1}`,
        hidden: (el.getAttribute('state') ?? 'visible') !== 'visible',
        kind,
        path: rel.target,
      });
    }
    if (!this.sheets.length) throw new Error('no sheets');
    const active = Number(first(book, 'workbookView')?.getAttribute('activeTab') ?? 0);
    const visible = this.sheets.findIndex((s, i) => !s.hidden && i >= active);
    this.activeSheet = visible >= 0 ? visible : Math.max(0, this.sheets.findIndex((s) => !s.hidden));

    const target = (suffix: string) => [...rels.values()].find((r) => r.type.endsWith(suffix))?.target;
    const themePath = target('/theme');
    this.theme = new Theme(themePath ? await this.xml(themePath) : null);
    const stylesPath = target('/styles');
    this.readStyles(stylesPath ? await this.xml(stylesPath) : null);
    const stringsPath = target('/sharedStrings');
    if (stringsPath) this.readStrings(await this.xml(stringsPath));
  }

  private theme = new Theme(null);
  private fonts: Font[] = [];
  private palette = INDEXED.slice();

  private readStyles(doc: Document | null): void {
    const theme = this.theme;
    const indexed = all(first(doc, 'indexedColors'), 'rgbColor');
    indexed.forEach((el, i) => (this.palette[i] = `#${(el.getAttribute('rgb') ?? '000000').slice(-6)}`));
    const colour = (el: Element | null | undefined) => resolveColour(el, theme, this.palette);

    const numFmts = new Map<number, string>();
    for (const el of all(doc, 'numFmt')) {
      numFmts.set(Number(el.getAttribute('numFmtId')), el.getAttribute('formatCode') ?? 'General');
    }
    this.fonts = all(first(doc, 'fonts'), 'font').map((el) => readFont(el, colour));
    const fills = all(first(doc, 'fills'), 'fill').map((el) => {
      const pattern = first(el, 'patternFill');
      if (pattern) {
        const type = pattern.getAttribute('patternType') ?? (first(pattern, 'fgColor') ? 'solid' : 'none');
        if (type === 'none') return '';
        // Non-solid patterns (hatching) are drawn as their foreground colour.
        return colour(first(pattern, 'fgColor')) || (type === 'solid' ? '' : colour(first(pattern, 'bgColor')));
      }
      return colour(first(first(el, 'gradientFill'), 'color'));
    });
    const borders = all(first(doc, 'borders'), 'border').map((el) => {
      const edges: CellFormat['border'] = {};
      for (const side of ['left', 'right', 'top', 'bottom'] as const) {
        const edge = child(el, side) ?? child(el, side === 'left' ? 'start' : side === 'right' ? 'end' : '');
        const style = edge?.getAttribute('style');
        if (style && style !== 'none') edges[side] = { style, colour: colour(first(edge, 'color')) || '#000000' };
      }
      return Object.keys(edges).length ? edges : null;
    });

    const fallbackFont: Font = {
      name: 'Calibri', size: 11, bold: false, italic: false, underline: '', strike: false, colour: '', family: 2,
    };
    const xfs = all(first(doc, 'cellXfs'), 'xf');
    this.formats = (xfs.length ? xfs : [null]).map((el) => {
      const num = (name: string) => Number(el?.getAttribute(name) ?? 0);
      const align = first(el, 'alignment');
      const v = align?.getAttribute('vertical');
      return {
        numFmt: formatCode(num('numFmtId'), numFmts),
        font: this.fonts[num('fontId')] ?? this.fonts[0] ?? fallbackFont,
        fill: fills[num('fillId')] ?? '',
        border: borders[num('borderId')] ?? null,
        h: align?.getAttribute('horizontal') ?? 'general',
        v: v === 'top' ? 'top' : v === 'center' || v === 'distributed' || v === 'justify' ? 'center' : 'bottom',
        wrap: /^(1|true)$/.test(align?.getAttribute('wrapText') ?? ''),
        indent: Number(align?.getAttribute('indent') ?? 0),
      };
    });
    this.defaultFont = this.formats[0].font;
    this.digitWidth = digitWidth(this.defaultFont);
  }

  private readStrings(doc: Document | null): void {
    const colour = (el: Element | null | undefined) => resolveColour(el, this.theme, this.palette);
    this.strings = all(doc, 'si').map((si) => readRich(si, colour));
  }

  private async readSheet(entry: SheetEntry): Promise<Sheet> {
    const doc = await this.xml(entry.path);
    if (!doc) throw new Error(`missing ${entry.path}`);
    const ws = doc.documentElement;
    const colour = (el: Element | null | undefined) => resolveColour(el, this.theme, this.palette);

    const rows: Sheet['rows'] = [];
    const heights = new Map<number, number>();
    const sharedFormulas: Sheet['sharedFormulas'] = new Map();
    let maxRow = -1;
    let maxCol = -1;
    let rowIndex = -1;
    for (const rowEl of children(child(ws, 'sheetData'), 'row')) {
      rowIndex = rowEl.hasAttribute('r') ? Number(rowEl.getAttribute('r')) - 1 : rowIndex + 1;
      if (rowEl.getAttribute('hidden') === '1' || rowEl.getAttribute('hidden') === 'true') heights.set(rowIndex, 0);
      else if (rowEl.hasAttribute('ht')) heights.set(rowIndex, ptToPx(Number(rowEl.getAttribute('ht'))));
      let colIndex = -1;
      let cells: Map<number, Cell> | undefined;
      for (const c of rowEl.children) {
        if (c.localName !== 'c') continue;
        const ref = c.getAttribute('r');
        colIndex = ref ? decodeCol(ref) : colIndex + 1;
        const cell = this.readCell(c, colour);
        if (cell.t === 'z' && cell.xf === 0 && !cell.f) continue;
        if (cell.f !== undefined && cell.si !== undefined) {
          sharedFormulas.set(cell.si, { f: cell.f, r: rowIndex, c: colIndex });
          delete cell.si;
        }
        (cells ??= new Map()).set(colIndex, cell);
        if (colIndex > maxCol) maxCol = colIndex;
      }
      if (cells) {
        rows[rowIndex] = cells;
        if (rowIndex > maxRow) maxRow = rowIndex;
      }
    }

    const merges: Range[] = [];
    for (const el of all(child(ws, 'mergeCells'), 'mergeCell')) {
      const range = decodeRange(el.getAttribute('ref') ?? '');
      if (!range) continue;
      merges.push(range);
      maxRow = Math.max(maxRow, range.r1);
      maxCol = Math.max(maxCol, range.c1);
    }

    const view = child(child(ws, 'sheetViews'), 'sheetView');
    const pane = child(view, 'pane');
    let frozenRows = 0;
    let frozenCols = 0;
    if (pane && /^frozen/.test(pane.getAttribute('state') ?? '')) {
      frozenRows = Math.max(0, Math.round(Number(pane.getAttribute('ySplit') ?? 0)));
      frozenCols = Math.max(0, Math.round(Number(pane.getAttribute('xSplit') ?? 0)));
    }
    // A frozen pane far down the sheet (or a broken file) shouldn't take the whole screen.
    frozenRows = Math.min(frozenRows, 50);
    frozenCols = Math.min(frozenCols, 20);

    const nRows = Math.max(MIN_ROWS, Math.max(maxRow + 1, frozenRows) + PAD_ROWS);
    const nCols = Math.min(16384, Math.max(MIN_COLS, Math.max(maxCol + 1, frozenCols) + PAD_COLS));

    const format = child(ws, 'sheetFormatPr');
    const defaultHeight = ptToPx(Number(format?.getAttribute('defaultRowHeight') ?? defaultRowPt(this.defaultFont)));
    const rowHeights = new Float64Array(nRows).fill(defaultHeight);
    for (const [r, h] of heights) if (r < nRows) rowHeights[r] = h;

    const mdw = this.digitWidth;
    const defaultColWidth = format?.hasAttribute('defaultColWidth')
      ? widthToPx(Number(format.getAttribute('defaultColWidth')), mdw)
      : Math.ceil((Number(format?.getAttribute('baseColWidth') ?? 8) * mdw + 5) / 8) * 8;
    const colWidths = new Float64Array(nCols).fill(defaultColWidth);
    for (const col of all(child(ws, 'cols'), 'col')) {
      const min = Number(col.getAttribute('min') ?? 1) - 1;
      const max = Math.min(nCols, Number(col.getAttribute('max') ?? min + 1));
      const hidden = /^(1|true)$/.test(col.getAttribute('hidden') ?? '');
      const width = hidden ? 0
        : col.hasAttribute('width') ? widthToPx(Number(col.getAttribute('width')), mdw)
        : defaultColWidth;
      for (let c = Math.max(0, min); c < max; c++) colWidths[c] = width;
    }

    const zoom = Number(view?.getAttribute('zoomScale') ?? 100) / 100;
    return {
      name: entry.name,
      rows,
      colWidths,
      rowHeights,
      frozenRows,
      frozenCols,
      merges,
      showGrid: view?.getAttribute('showGridLines') !== '0' && view?.getAttribute('showGridLines') !== 'false',
      zoom: zoom >= 0.1 && zoom <= 4 ? zoom : 1,
      hasDrawings: !!child(ws, 'drawing'),
      sharedFormulas,
    };
  }

  private readCell(c: Element, colour: (el: Element | null | undefined) => string): Cell {
    const type = c.getAttribute('t') ?? 'n';
    let v: Element | null = null;
    let f: Element | null = null;
    let is: Element | null = null;
    for (const el of c.children) {
      if (el.localName === 'v') v = el;
      else if (el.localName === 'f') f = el;
      else if (el.localName === 'is') is = el;
    }
    const cell: Cell = { t: 'z', v: null, xf: Number(c.getAttribute('s') ?? 0) };
    const raw = v?.textContent ?? null;
    if (type === 'inlineStr' && is) {
      setText(cell, readRich(is, colour));
    } else if (raw !== null) {
      if (type === 's') setText(cell, this.strings[Number(raw)] ?? '');
      else if (type === 'str' || type === 'd') setText(cell, raw);
      else if (type === 'b') {
        cell.t = 'b';
        cell.v = raw === '1' || raw === 'true';
      } else if (type === 'e') {
        cell.t = 'e';
        cell.v = raw;
      } else {
        const n = Number(raw);
        if (Number.isFinite(n)) {
          cell.t = 'n';
          cell.v = n;
        } else setText(cell, raw);
      }
    }
    if (f) {
      const kind = f.getAttribute('t');
      const text = f.textContent ?? '';
      if (kind === 'shared' && f.hasAttribute('si')) {
        cell.si = Number(f.getAttribute('si'));
        // The group's first cell carries the formula; the others only the group index.
        if (text) cell.f = text;
      } else {
        cell.f = text;
        if (kind === 'array') cell.array = true;
      }
    }
    return cell;
  }
}

function setText(cell: Cell, value: string | Run[]): void {
  cell.t = 's';
  if (typeof value === 'string') cell.v = value;
  else {
    cell.v = value.map((r) => r.text).join('');
    if (value.some((r) => r.font)) cell.rich = value;
  }
}

// ------------------------------------------------------------------ XML helpers

/** Excel escapes control characters in text as _xHHHH_. */
function unescapeText(s: string): string {
  return s.includes('_x') ? s.replace(/_x([0-9a-fA-F]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16))) : s;
}

/** Text of a shared string or inline string (`<si>` / `<is>`), with runs if it has any. */
function readRich(el: Element, colour: (el: Element | null | undefined) => string): string | Run[] {
  const runs: Run[] = [];
  for (const part of el.children) {
    if (part.localName === 't') runs.push({ text: unescapeText(part.textContent ?? '') });
    else if (part.localName === 'r') {
      const props = child(part, 'rPr');
      runs.push({
        text: unescapeText(child(part, 't')?.textContent ?? ''),
        font: props ? readRunFont(props, colour) : undefined,
      });
    }
    // <rPh> (phonetic guides) and <phoneticPr> are not shown.
  }
  if (runs.length === 1 && !runs[0].font) return runs[0].text;
  return runs.some((r) => r.font) ? runs : runs.map((r) => r.text).join('');
}

/** Reads `<font>` (styles) or `<rPr>` (a rich-text run; uses `rFont` for the name). */
function readFont(el: Element, colour: (el: Element | null | undefined) => string): Font {
  const val = (name: string) => child(el, name)?.getAttribute('val') ?? null;
  const flag = (name: string) => {
    const c = child(el, name);
    return !!c && !/^(0|false)$/.test(c.getAttribute('val') ?? '1');
  };
  const u = child(el, 'u');
  const uVal = u ? u.getAttribute('val') ?? 'single' : 'none';
  return {
    name: val('name') ?? val('rFont') ?? 'Calibri',
    size: Number(val('sz') ?? 11) || 11,
    bold: flag('b'),
    italic: flag('i'),
    underline: uVal === 'none' ? '' : uVal.startsWith('double') ? 'double' : 'single',
    strike: flag('strike'),
    colour: colour(child(el, 'color')),
    family: Number(val('family') ?? 2),
  };
}

/** A rich-text run's `<rPr>`: only the properties it sets. */
function readRunFont(el: Element, colour: (el: Element | null | undefined) => string): Partial<Font> {
  const full = readFont(el, colour);
  const font: Partial<Font> = {};
  const has = (name: string) => !!child(el, name);
  if (has('rFont') || has('name')) font.name = full.name;
  if (has('sz')) font.size = full.size;
  if (has('b')) font.bold = full.bold;
  if (has('i')) font.italic = full.italic;
  if (has('u')) font.underline = full.underline;
  if (has('strike')) font.strike = full.strike;
  if (has('color')) font.colour = full.colour;
  if (has('family')) font.family = full.family;
  return font;
}

// ------------------------------------------------------------------ addresses and units

export function decodeCol(ref: string): number {
  let col = 0;
  for (let i = 0; i < ref.length; i++) {
    const code = ref.charCodeAt(i) & ~32; // upper case
    if (code < 65 || code > 90) break;
    col = col * 26 + (code - 64);
  }
  return col - 1;
}

export function encodeCol(col: number): string {
  let s = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function decodeCell(ref: string): { r: number; c: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
  return m ? { c: decodeCol(m[1]), r: Number(m[2]) - 1 } : null;
}

function decodeRange(ref: string): Range | null {
  const [a, b = a] = ref.split(':');
  const start = decodeCell(a);
  const end = decodeCell(b);
  if (!start || !end) return null;
  return {
    r0: Math.min(start.r, end.r), c0: Math.min(start.c, end.c),
    r1: Math.max(start.r, end.r), c1: Math.max(start.c, end.c),
  };
}

function ptToPx(pt: number): number {
  return Math.round((pt * 96) / 72);
}

/** Column width in characters (as stored) to pixels, per the OOXML spec. */
function widthToPx(width: number, mdw: number): number {
  return Math.max(0, Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw));
}

/** Width of a digit as a fraction of the font size, for fonts Excel files commonly use. */
const DIGIT_RATIO: Record<string, number> = {
  calibri: 0.507, 'calibri light': 0.507, aptos: 0.554, 'aptos narrow': 0.5, 'aptos display': 0.554,
  arial: 0.556, helvetica: 0.556, 'liberation sans': 0.556, 'arial narrow': 0.456,
  'times new roman': 0.5, cambria: 0.556, georgia: 0.614, verdana: 0.636, tahoma: 0.546,
  'segoe ui': 0.559, 'courier new': 0.6, consolas: 0.55,
};

/**
 * Excel measures column widths in digits of the workbook's default font, at the resolution the
 * author's Excel used. Use the known proportions of common fonts rather than measuring the
 * stand-in font here, so columns are as wide as the author saw them.
 */
function digitWidth(font: Font): number {
  const ratio = DIGIT_RATIO[font.name.toLowerCase()] ?? 0.53;
  return Math.max(5, Math.round(((font.size * 96) / 72) * ratio));
}

/** Excel's default row height for a font, when the file doesn't give one. */
function defaultRowPt(font: Font): number {
  return Math.round(font.size * 1.36 * 4) / 4;
}

// ------------------------------------------------------------------ colours

/** The 64 legacy indexed colours, then system foreground (64) and background (65). */
const INDEXED = (
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333 ' +
  '000000 FFFFFF'
).split(' ').map((hex) => `#${hex}`);

class Theme {
  /** Theme colour slots in the order cells refer to them: lt1, dk1, lt2, dk2, accent1–6, hlink, folHlink. */
  colours: string[] = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#0563C1', '#954F72'];

  constructor(doc: Document | null) {
    const scheme = first(doc, 'clrScheme');
    if (!scheme) return;
    const read = (name: string) => {
      const el = child(scheme, name);
      const c = el?.firstElementChild;
      const hex = c?.getAttribute('val') && c.localName === 'srgbClr' ? c.getAttribute('val') : c?.getAttribute('lastClr');
      return hex ? `#${hex}` : null;
    };
    const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    order.forEach((name, i) => (this.colours[i] = read(name) ?? this.colours[i]));
  }
}

function resolveColour(el: Element | null | undefined, theme: Theme, palette: string[]): string {
  if (!el) return '';
  let hex = '';
  const rgb = el.getAttribute('rgb');
  if (rgb) hex = `#${rgb.slice(-6)}`;
  else if (el.hasAttribute('theme')) hex = theme.colours[Number(el.getAttribute('theme'))] ?? '';
  else if (el.hasAttribute('indexed')) {
    const i = Number(el.getAttribute('indexed'));
    hex = i === 64 ? '' : palette[i] ?? '';
  }
  if (!hex) return '';
  const tint = Number(el.getAttribute('tint') ?? 0);
  return tint ? applyTint(hex, tint) : hex.toLowerCase();
}

/** Excel's tint: lightens (tint > 0) or darkens (tint < 0) the colour's HSL lightness. */
function applyTint(hex: string, tint: number): string {
  const n = parseInt(hex.slice(1), 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((x) => x / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let l = (max + min) / 2;
  const d = max - min;
  const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
  let h = 0;
  if (d) {
    if (max === r) h = ((g - b) / d) % 6;
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    if (h < 0) h += 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs((h % 2) - 1));
  const m = l - c / 2;
  const [r1, g1, b1] = h < 1 ? [c, x, 0] : h < 2 ? [x, c, 0] : h < 3 ? [0, c, x] : h < 4 ? [0, x, c] : h < 5 ? [x, 0, c] : [c, 0, x];
  const to = (v: number) => Math.round(Math.min(1, Math.max(0, v + m)) * 255).toString(16).padStart(2, '0');
  return `#${to(r1)}${to(g1)}${to(b1)}`;
}
