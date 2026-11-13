import { child, children, first } from '../ooxml';
import { colour, css, type ColourCtx, type Rgba } from '../pptx/drawing';
import { fontFamily } from '../pptx/text';
import { formatNumber } from '../xlsx/format';

// Reads a chart part — DrawingML charts (`c:chartSpace`, chartN.xml) and the newer "chartex"
// charts (`cx:chartSpace`, chartExN.xml: box & whisker, waterfall…) — into a plain model that
// draw.ts lays out and draws. Office saves every value a chart shows (each series caches its
// numbers and category names), so nothing is recalculated.

/** What the host document lends the chart: its theme (colours, fonts). */
export interface ChartTheme extends ColourCtx {
  theme: { colours: Record<string, string>; majorFont: string; minorFont: string };
  /** Default text size in pt where the chart sets none (PowerPoint's charts run larger). */
  baseSize: number;
  /** Word and Excel frame a chart that sets no outline of its own; PowerPoint doesn't. */
  frame: boolean;
}

export interface TextStyle {
  /** px */
  size: number;
  colour: string;
  bold: boolean;
  italic: boolean;
  font: string;
}

export interface Paint {
  fill: string | null;
  stroke: string | null;
  /** px */
  width: number;
  dash: number[] | null;
}

export interface Marker {
  symbol: string;
  /** px */
  size: number;
  paint: Paint;
}

export interface Labels {
  value: boolean;
  percent: boolean;
  category: boolean;
  series: boolean;
  pos: string;
  style: TextStyle;
  format: string | null;
  separator: string;
}

export interface Series {
  name: string;
  values: (number | null)[];
  /** Scatter and bubble charts: x values. */
  xs: (number | null)[] | null;
  /** Bubble charts: sizes. */
  sizes: (number | null)[] | null;
  categories: string[];
  format: string;
  paint: Paint;
  /** Fills of single points (pies, and bars coloured one by one). */
  points: Map<number, Paint>;
  marker: Marker | null;
  smooth: boolean;
  labels: Labels | null;
  explosion: number;
  pointExplosion: Map<number, number>;
  /** Box & whisker: the raw values of each category, in category order. */
  groups?: { category: string; values: number[] }[];
  box?: { mean: boolean; meanLine: boolean; outliers: boolean; inner: boolean; exclusive: boolean };
}

export type GroupType = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'bubble' | 'radar' | 'box';

export interface Group {
  type: GroupType;
  /** Bars: 'col' (vertical) or 'bar' (horizontal). */
  dir: 'col' | 'bar';
  grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  varyColors: boolean;
  series: Series[];
  /** % of a bar's width */
  gapWidth: number;
  overlap: number;
  /** Doughnuts: % of the radius. */
  holeSize: number;
  /** Degrees clockwise from 12 o'clock. */
  firstSliceAngle: number;
  threeD: boolean;
  axisIds: string[];
  /** Line charts: drop lines, high-low lines… are not drawn. */
  radarStyle: string;
}

export interface Axis {
  id: string;
  kind: 'cat' | 'val' | 'date' | 'ser';
  /** Where the axis sits: b, t, l, r. */
  pos: string;
  deleted: boolean;
  min: number | null;
  max: number | null;
  majorUnit: number | null;
  reverse: boolean;
  log: number | null;
  grid: Paint | null;
  line: Paint | null;
  text: TextStyle;
  format: string | null;
  sourceLinked: boolean;
  crossAx: string;
  crosses: string;
  crossesAt: number | null;
  labelPos: string;
  majorTick: string;
  title: Title | null;
  /** Value axis: categories sit between ticks ('between') or on them ('midCat'). */
  crossBetween: string;
  rotate: number | null;
}

export interface Title {
  lines: string[];
  style: TextStyle;
  rotate: number;
}

export interface Legend {
  pos: string;
  style: TextStyle;
  overlay: boolean;
  /** Indexes of entries the author deleted. */
  hidden: Set<number>;
}

/** A layout the author set by hand, as fractions of the chart. */
export interface ManualLayout {
  x: number;
  y: number;
  w: number | null;
  h: number | null;
  /** The box is the plot area without its axis labels. */
  inner: boolean;
}

export interface Chart {
  title: Title | null;
  legend: Legend | null;
  groups: Group[];
  axes: Map<string, Axis>;
  background: Paint;
  plotBackground: Paint;
  plotLayout: ManualLayout | null;
  /** The chart's base text style (label defaults). */
  text: TextStyle;
  /** Kinds the drawing leaves out (shown as a notice by the host). */
  unsupported: string[];
}

const PT = 4 / 3;
const EMU = 9525;

/** A chart part's root element → model. Returns null for kinds this reader doesn't know. */
export function readChart(root: Element, theme: ChartTheme): Chart | null {
  return root.namespaceURI?.includes('chartex') ? readChartEx(root, theme) : readClassic(root, theme);
}

// ------------------------------------------------------------------ shared pieces

const val = (el: Element | null | undefined, name: string): string | null => child(el, name)?.getAttribute('val') ?? null;
const flag = (el: Element | null | undefined, name: string, dflt: boolean): boolean => {
  const c = child(el, name);
  if (!c) return dflt;
  const v = c.getAttribute('val');
  return v === null ? true : v === '1' || v === 'true';
};
const num = (el: Element | null | undefined, name: string): number | null => {
  const v = val(el, name);
  return v === null || v === '' || Number.isNaN(Number(v)) ? null : Number(v);
};

function rgba(c: Rgba | null): string | null {
  return c && c.a > 0 ? css(c) : null;
}

/** Text style from a `txPr`/`rich` body (its first paragraph's defRPr or first run). */
function textStyle(holder: Element | null, base: TextStyle, theme: ChartTheme, sizeDefault?: number): TextStyle {
  const p = first(holder, 'p');
  const rPr = child(child(p, 'pPr'), 'defRPr') ?? first(p, 'rPr');
  const runPr = first(p, 'rPr');
  const style = { ...base };
  if (sizeDefault !== undefined) style.size = sizeDefault;
  for (const pr of [rPr, runPr]) {
    if (!pr) continue;
    const sz = pr.getAttribute('sz');
    if (sz) style.size = (Number(sz) / 100) * PT;
    const b = pr.getAttribute('b');
    if (b !== null) style.bold = b === '1' || b === 'true';
    const i = pr.getAttribute('i');
    if (i !== null) style.italic = i === '1' || i === 'true';
    const fill = child(pr, 'solidFill');
    const c = fill ? rgba(colour(fill, theme)) : null;
    if (c) style.colour = c;
    const latin = child(pr, 'latin')?.getAttribute('typeface');
    if (latin) style.font = fontFamily(latin, theme);
  }
  return style;
}

const DASH: Record<string, number[]> = {
  dash: [4, 3], sysDash: [3, 1], dot: [1, 3], sysDot: [1, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** Fill and outline from an `spPr`, over defaults. Gradients are drawn as their first colour. */
function paint(spPr: Element | null | undefined, theme: ChartTheme, dflt: Partial<Paint> = {}): Paint {
  const out: Paint = { fill: dflt.fill ?? null, stroke: dflt.stroke ?? null, width: dflt.width ?? 0.75 * PT, dash: dflt.dash ?? null };
  if (!spPr) return out;
  const fillEl = children(spPr).find((c) => /^(noFill|solidFill|gradFill|pattFill)$/.test(c.localName));
  if (fillEl) out.fill = fillColour(fillEl, theme);
  const ln = child(spPr, 'ln');
  if (ln) {
    if (ln.hasAttribute('w')) out.width = Number(ln.getAttribute('w')) / EMU;
    const lnFill = children(ln).find((c) => /^(noFill|solidFill|gradFill|pattFill)$/.test(c.localName));
    if (lnFill) out.stroke = fillColour(lnFill, theme);
    const dash = val(ln, 'prstDash');
    if (dash) out.dash = DASH[dash] ?? null;
  }
  return out;
}

function fillColour(el: Element, theme: ChartTheme): string | null {
  switch (el.localName) {
    case 'solidFill':
      return rgba(colour(el, theme));
    case 'gradFill': {
      const stops = children(child(el, 'gsLst'), 'gs');
      return rgba(colour(stops[Math.floor(stops.length / 2)] ?? stops[0], theme));
    }
    case 'pattFill':
      return rgba(colour(child(el, 'fgClr'), theme));
  }
  return null;
}

/** Office's default series colours: the six accents, then darker and lighter rounds of them. */
export function seriesColour(i: number, theme: ChartTheme): string {
  const accent = theme.theme.colours[`accent${(i % 6) + 1}`] ?? '808080';
  const round = Math.floor(i / 6);
  const n = parseInt(accent, 16);
  let [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  if (round) {
    // Rounds alternate darker (60%) and lighter (60% tint), as Office's colour styles do.
    const k = round % 2 ? 0.6 : 0.4;
    [r, g, b] = round % 2 ? [r, g, b].map((v) => v * k) : [r, g, b].map((v) => v + (255 - v) * (1 - k));
  }
  return `rgb(${Math.round(r)} ${Math.round(g)} ${Math.round(b)})`;
}

/**
 * Text where the chart says nothing: chartex charts default to Office's grey; DrawingML charts
 * that carry no text properties are older ones, whose default is the text colour (black).
 */
function baseText(theme: ChartTheme, legacy = false): TextStyle {
  const colour = legacy ? ink(theme) : 'rgb(89 89 89)';
  return { size: theme.baseSize * PT, colour, bold: false, italic: false, font: fontFamily('+mn-lt', theme) };
}

/** The theme's text colour (tx1). */
function ink(theme: ChartTheme): string {
  const hex = theme.theme.colours[theme.clrMap.tx1 ?? 'dk1'] ?? '000000';
  const n = parseInt(hex, 16);
  return `rgb(${(n >> 16) & 255} ${(n >> 8) & 255} ${n & 255})`;
}

/** Rich text paragraphs (`c:rich` / `cx:rich`), one string per paragraph. */
function richLines(rich: Element | null): string[] {
  return children(rich, 'p').map((p) => all(p, 't').map((t) => t.textContent ?? '').join(''));
}

function all(el: Element | null, name: string): Element[] {
  return el ? [...el.getElementsByTagNameNS('*', name)] : [];
}

function manualLayout(layout: Element | null): ManualLayout | null {
  const m = child(layout, 'manualLayout');
  if (!m) return null;
  // "factor" positions are offsets from where Office would put it; only "edge" ones are absolute.
  if ((val(m, 'xMode') ?? 'factor') !== 'edge' || (val(m, 'yMode') ?? 'factor') !== 'edge') return null;
  const x = num(m, 'x');
  const y = num(m, 'y');
  if (x === null || y === null) return null;
  return { x, y, w: num(m, 'w'), h: num(m, 'h'), inner: val(m, 'layoutTarget') === 'inner' };
}

// ------------------------------------------------------------------ DrawingML charts

const GROUP_TYPES: Record<string, GroupType> = {
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area',
  pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter',
  bubbleChart: 'bubble', radarChart: 'radar',
};

function readClassic(root: Element, theme: ChartTheme): Chart | null {
  const chart = child(root, 'chart');
  const plotArea = child(chart, 'plotArea');
  if (!plotArea) return null;
  const text = textStyle(child(root, 'txPr'), baseText(theme, true), theme);
  const unsupported: string[] = [];

  const groups: Group[] = [];
  let seriesIndex = 0;
  for (const g of children(plotArea)) {
    const type = GROUP_TYPES[g.localName];
    if (!type) {
      if (g.localName.endsWith('Chart')) unsupported.push(g.localName);
      continue;
    }
    const group: Group = {
      type,
      dir: val(g, 'barDir') === 'bar' ? 'bar' : 'col',
      grouping: (val(g, 'grouping') ?? (type === 'bar' ? 'clustered' : 'standard')) as Group['grouping'],
      varyColors: flag(g, 'varyColors', false),
      series: [],
      gapWidth: num(g, 'gapWidth') ?? 150,
      overlap: num(g, 'overlap') ?? 0,
      holeSize: num(g, 'holeSize') ?? 50,
      firstSliceAngle: num(g, 'firstSliceAng') ?? 0,
      threeD: g.localName.includes('3D'),
      axisIds: children(g, 'axId').map((a) => a.getAttribute('val') ?? ''),
      radarStyle: val(g, 'radarStyle') ?? 'marker',
    };
    if (type === 'scatter' && /^(line|smooth)$/.test(val(g, 'scatterStyle') ?? '')) group.radarStyle = 'noMarker';
    const groupLabels = child(g, 'dLbls');
    const groupMarkers = type === 'line' ? flag(g, 'marker', true) : true;
    for (const s of children(g, 'ser').sort((a, b) => (num(a, 'order') ?? 0) - (num(b, 'order') ?? 0))) {
      group.series.push(readSeries(s, group, seriesIndex++, groupLabels, groupMarkers, text, theme));
    }
    // "Vary colours by point" applies to a chart with a single series.
    if (group.varyColors && group.series.length === 1 && (type === 'bar' || type === 'line' || type === 'area')) {
      const s = group.series[0];
      for (let i = 0; i < s.values.length; i++) {
        if (!s.points.has(i)) s.points.set(i, { ...s.paint, fill: type === 'line' ? s.paint.fill : seriesColour(i, theme) });
      }
    }
    if (group.series.length) groups.push(group);
  }
  if (!groups.length) return null;

  const axes = new Map<string, Axis>();
  for (const a of children(plotArea)) {
    if (!/^(catAx|valAx|dateAx|serAx)$/.test(a.localName)) continue;
    const axis = readAxis(a, text, theme);
    axes.set(axis.id, axis);
  }

  let title: Title | null = null;
  const titleEl = child(chart, 'title');
  const allSeries = groups.flatMap((g) => g.series);
  if (titleEl) {
    let lines = richLines(first(child(titleEl, 'tx'), 'rich'));
    // No text of its own: Office shows the series name when there is one series.
    if (!lines.length) lines = [allSeries.length === 1 && allSeries[0].name ? allSeries[0].name : 'Chart Title'];
    const big = { ...text, size: text.size * 1.4 };
    title = { lines, style: textStyle(child(titleEl, 'txPr') ?? first(child(titleEl, 'tx'), 'rich'), big, theme), rotate: 0 };
  } else if (!flag(chart, 'autoTitleDeleted', false) && allSeries.length === 1 && allSeries[0].name) {
    // A chart with one series and no title element gets the series name as an automatic title.
    title = { lines: [allSeries[0].name], style: { ...text, size: text.size * 1.2, bold: true }, rotate: 0 };
  }

  let legend: Legend | null = null;
  const legendEl = child(chart, 'legend');
  if (legendEl) {
    legend = {
      pos: val(legendEl, 'legendPos') ?? 'r',
      style: textStyle(child(legendEl, 'txPr'), text, theme),
      overlay: flag(legendEl, 'overlay', false),
      hidden: new Set(children(legendEl, 'legendEntry').filter((e) => flag(e, 'delete', false)).map((e) => num(e, 'idx') ?? -1)),
    };
  }

  return {
    title,
    legend,
    groups,
    axes,
    // A chart without its own outline is an older one: Office frames those in the text colour.
    background: paint(child(root, 'spPr'), theme, theme.frame ? { ...chartFrame(theme), stroke: ink(theme) } : {}),
    plotBackground: paint(child(plotArea, 'spPr'), theme),
    plotLayout: manualLayout(child(plotArea, 'layout')),
    text,
    unsupported,
  };
}

/** Cached values of a `numRef`/`numLit`/`strRef`/`strLit`/`multiLvlStrRef` holder. */
function cache(holder: Element | null): { values: (string | null)[]; format: string } {
  if (!holder) return { values: [], format: 'General' };
  const src = children(holder)[0];
  if (!src) return { values: [], format: 'General' };
  const c = src.localName.endsWith('Ref') ? children(src).find((x) => /Cache$/.test(x.localName)) : src;
  if (!c) return { values: [], format: 'General' };
  // Multi-level categories: use the innermost level (the one next to the axis).
  const level = c.localName === 'multiLvlStrCache' ? child(c, 'lvl') : c;
  const count = num(c, 'ptCount') ?? 0;
  const values: (string | null)[] = new Array(count).fill(null);
  for (const pt of children(level, 'pt')) {
    const idx = Number(pt.getAttribute('idx') ?? 0);
    values[idx] = child(pt, 'v')?.textContent ?? null;
  }
  return { values, format: child(c, 'formatCode')?.textContent ?? 'General' };
}

function numbers(holder: Element | null): { values: (number | null)[]; format: string } {
  const { values, format } = cache(holder);
  return { values: values.map((v) => (v === null || v === '' || Number.isNaN(Number(v)) ? null : Number(v))), format };
}

function readSeries(s: Element, group: Group, index: number, groupLabels: Element | null, groupMarkers: boolean, text: TextStyle, theme: ChartTheme): Series {
  const idx = num(s, 'idx') ?? index;
  const name = cache(child(s, 'tx')).values[0] ?? (child(s, 'tx')?.textContent?.trim() || `Series ${idx + 1}`);
  const lineLike = group.type === 'line' || group.type === 'scatter' || group.type === 'radar';
  const auto = seriesColour(idx, theme);
  const dflt: Partial<Paint> = lineLike ? { stroke: auto, width: 2.25 * PT } : { fill: auto, stroke: group.type === 'pie' || group.type === 'doughnut' ? '#fff' : null };
  const spPr = child(s, 'spPr');
  const seriesPaint = paint(spPr, theme, dflt);
  // A scatter chart drawn "with markers only" has series lines switched off.
  if (group.type === 'scatter' && group.radarStyle === 'noMarker' && !child(spPr, 'ln')) seriesPaint.stroke = null;

  const values = numbers(child(s, 'val') ?? child(s, 'yVal'));
  const catSource = child(s, 'cat') ?? child(s, 'xVal');
  const cats = cache(catSource);
  const catIsNumber = !!first(catSource, 'numRef') || !!first(catSource, 'numLit');
  const categories = cats.values.map((v) => {
    if (v === null) return '';
    if (catIsNumber && cats.format !== 'General' && !Number.isNaN(Number(v))) return formatLabel(Number(v), cats.format);
    return v;
  });
  const xs = group.type === 'scatter' || group.type === 'bubble' ? numbers(child(s, 'xVal')).values : null;

  const points = new Map<number, Paint>();
  const pointExplosion = new Map<number, number>();
  for (const dPt of children(s, 'dPt')) {
    const i = num(dPt, 'idx') ?? 0;
    const pd = group.varyColors ? { ...dflt, fill: seriesColour(i, theme) } : { ...seriesPaint };
    points.set(i, paint(child(dPt, 'spPr'), theme, pd));
    const ex = num(dPt, 'explosion');
    if (ex !== null) pointExplosion.set(i, ex);
  }
  if (group.varyColors && (group.type === 'pie' || group.type === 'doughnut')) {
    for (let i = 0; i < values.values.length; i++) if (!points.has(i)) points.set(i, paint(spPr, theme, { ...dflt, fill: seriesColour(i, theme) }));
  }

  let marker: Marker | null = null;
  if (group.type === 'line' || group.type === 'scatter' || group.type === 'radar') {
    const m = child(s, 'marker');
    let symbol = val(m, 'symbol') ?? (groupMarkers && group.radarStyle !== 'noMarker' && group.radarStyle !== 'standard' ? 'auto' : 'none');
    if (group.type === 'radar' && !m && group.radarStyle !== 'marker') symbol = 'none';
    if (symbol === 'auto') symbol = AUTO_SYMBOLS[idx % AUTO_SYMBOLS.length];
    if (symbol !== 'none') {
      const colourOf = seriesPaint.stroke ?? auto;
      marker = {
        symbol,
        size: (num(m, 'size') ?? 5) * PT,
        paint: paint(child(m, 'spPr'), theme, { fill: colourOf, stroke: colourOf, width: 0.75 * PT }),
      };
    }
  }

  const labelsEl = child(s, 'dLbls') ?? groupLabels;
  return {
    name,
    values: values.values,
    xs,
    sizes: group.type === 'bubble' ? numbers(child(s, 'bubbleSize')).values : null,
    categories,
    format: values.format,
    paint: seriesPaint,
    points,
    marker,
    smooth: flag(s, 'smooth', false),
    labels: readLabels(labelsEl, group, text, theme),
    explosion: num(s, 'explosion') ?? 0,
    pointExplosion,
  };
}

const AUTO_SYMBOLS = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus', 'dot', 'dash'];

function readLabels(el: Element | null, group: Group, text: TextStyle, theme: ChartTheme): Labels | null {
  if (!el || flag(el, 'delete', false)) return null;
  const labels: Labels = {
    value: flag(el, 'showVal', false),
    percent: flag(el, 'showPercent', false),
    category: flag(el, 'showCatName', false),
    series: flag(el, 'showSerName', false),
    pos: val(el, 'dLblPos') ?? (group.type === 'pie' || group.type === 'doughnut' ? 'bestFit' : group.type === 'bar' ? (group.grouping === 'clustered' ? 'outEnd' : 'ctr') : 'r'),
    style: textStyle(child(el, 'txPr'), { ...text, size: text.size * 0.9 }, theme),
    format: child(el, 'numFmt')?.getAttribute('formatCode') ?? null,
    separator: child(el, 'separator')?.textContent ?? (group.type === 'pie' ? '; ' : ', '),
  };
  return labels.value || labels.percent || labels.category || labels.series ? labels : null;
}

function readAxis(a: Element, text: TextStyle, theme: ChartTheme): Axis {
  const scaling = child(a, 'scaling');
  const numFmt = child(a, 'numFmt');
  const kind = a.localName.slice(0, 3) as Axis['kind'];
  const titleEl = child(a, 'title');
  let title: Title | null = null;
  if (titleEl) {
    const lines = richLines(first(child(titleEl, 'tx'), 'rich'));
    const rot = first(titleEl, 'bodyPr')?.getAttribute('rot');
    const pos = val(a, 'axPos') ?? 'b';
    title = {
      lines: lines.length ? lines : ['Axis Title'],
      style: textStyle(child(titleEl, 'txPr') ?? first(child(titleEl, 'tx'), 'rich'), { ...text, bold: false }, theme),
      rotate: rot !== null && rot !== undefined ? Number(rot) / 60000 : pos === 'l' || pos === 'r' ? -90 : 0,
    };
  }
  const grid = child(a, 'majorGridlines');
  const rot = first(child(a, 'txPr'), 'bodyPr')?.getAttribute('rot');
  return {
    id: val(a, 'axId') ?? '',
    kind,
    pos: val(a, 'axPos') ?? 'b',
    deleted: flag(a, 'delete', false),
    min: num(scaling, 'min'),
    max: num(scaling, 'max'),
    majorUnit: num(a, 'majorUnit'),
    reverse: val(scaling, 'orientation') === 'maxMin',
    log: num(scaling, 'logBase'),
    // Without styling markup (older files) Office draws axes and gridlines in the text colour.
    grid: grid ? paint(child(grid, 'spPr'), theme, { stroke: ink(theme), width: 0.75 }) : null,
    line: paint(child(a, 'spPr'), theme, { stroke: ink(theme), width: 0.75 }),
    text: textStyle(child(a, 'txPr'), { ...text, size: text.size * 0.9 }, theme),
    format: numFmt?.getAttribute('formatCode') ?? null,
    sourceLinked: numFmt ? numFmt.getAttribute('sourceLinked') !== '0' : true,
    crossAx: val(a, 'crossAx') ?? '',
    crosses: val(a, 'crosses') ?? (child(a, 'crossesAt') ? 'at' : 'autoZero'),
    crossesAt: num(a, 'crossesAt'),
    labelPos: val(a, 'tickLblPos') ?? 'nextTo',
    majorTick: val(a, 'majorTickMark') ?? 'out',
    title,
    crossBetween: val(a, 'crossBetween') ?? 'between',
    rotate: rot && rot !== '-60000000' ? Number(rot) / 60000 : null,
  };
}

// ------------------------------------------------------------------ chartex (Office 2016 charts)

function readChartEx(root: Element, theme: ChartTheme): Chart | null {
  const chart = child(root, 'chart');
  const plotArea = child(chart, 'plotArea');
  const region = child(plotArea, 'plotAreaRegion');
  if (!region) return null;
  const text = textStyle(child(root, 'txPr'), baseText(theme), theme);

  // Data sets by id: categories (string dimension) and values (number dimension).
  const data = new Map<string, { cats: string[]; values: (number | null)[]; format: string }>();
  for (const d of children(child(root, 'chartData'), 'data')) {
    const level = (dim: Element | null) => {
      const lvl = children(dim, 'lvl')[children(dim, 'lvl').length - 1] ?? null;
      const count = Number(lvl?.getAttribute('ptCount') ?? 0);
      const out: (string | null)[] = new Array(count).fill(null);
      for (const pt of children(lvl, 'pt')) out[Number(pt.getAttribute('idx') ?? 0)] = pt.textContent;
      return { out, format: lvl?.getAttribute('formatCode') ?? 'General' };
    };
    const cats = level(child(d, 'strDim') ?? children(d, 'numDim').find((n) => n.getAttribute('type') === 'cat') ?? null);
    const vals = level(children(d, 'numDim').find((n) => n.getAttribute('type') !== 'cat') ?? null);
    data.set(d.getAttribute('id') ?? '', {
      cats: cats.out.map((c) => c ?? ''),
      values: vals.out.map((v) => (v === null || Number.isNaN(Number(v)) ? null : Number(v))),
      format: vals.format,
    });
  }

  const unsupported: string[] = [];
  const series: Series[] = [];
  let index = 0;
  for (const s of children(region, 'series')) {
    const layout = s.getAttribute('layoutId') ?? '';
    if (s.getAttribute('hidden') === '1') continue;
    if (layout !== 'boxWhisker') {
      unsupported.push(layout);
      continue;
    }
    const d = data.get(val(s, 'dataId') ?? '');
    if (!d) continue;
    const i = index++;
    const fill = seriesColour(i, theme);
    const p = paint(child(s, 'spPr'), theme, { fill, stroke: darker(fill), width: 0.75 * PT });
    const groupsByCat: { category: string; values: number[] }[] = [];
    d.values.forEach((v, k) => {
      if (v === null) return;
      const cat = d.cats[k] ?? '';
      let g = groupsByCat.find((x) => x.category === cat);
      if (!g) groupsByCat.push((g = { category: cat, values: [] }));
      g.values.push(v);
    });
    const vis = first(s, 'visibility');
    const visible = (name: string, dflt: boolean) => {
      const v = vis?.getAttribute(name);
      return v === null || v === undefined ? dflt : v === '1' || v === 'true';
    };
    series.push({
      name: first(child(s, 'tx'), 'v')?.textContent ?? `Series ${i + 1}`,
      values: d.values,
      xs: null,
      sizes: null,
      categories: d.cats,
      format: d.format,
      paint: p,
      points: new Map(),
      marker: null,
      smooth: false,
      labels: null,
      explosion: 0,
      pointExplosion: new Map(),
      groups: groupsByCat,
      box: {
        mean: visible('meanMarker', true),
        meanLine: visible('meanLine', false),
        outliers: visible('outliers', true),
        inner: visible('nonoutliers', false),
        exclusive: first(s, 'statistics')?.getAttribute('quartileMethod') !== 'inclusive',
      },
    });
  }
  if (!series.length) return null;

  const axisEls = children(plotArea, 'axis');
  const catEl = axisEls.find((a) => child(a, 'catScaling')) ?? null;
  const valEl = axisEls.find((a) => child(a, 'valScaling')) ?? null;
  const gap = Number(child(catEl, 'catScaling')?.getAttribute('gapWidth') ?? 1.5);
  const axes = new Map<string, Axis>();
  const axisOf = (el: Element | null, kind: 'cat' | 'val', id: string, cross: string): Axis => {
    const scaling = child(el, kind === 'cat' ? 'catScaling' : 'valScaling');
    const attr = (n: string) => {
      const v = scaling?.getAttribute(n);
      return v && v !== 'auto' && !Number.isNaN(Number(v)) ? Number(v) : null;
    };
    const grid = child(el, 'majorGridlines');
    return {
      id, kind, pos: kind === 'cat' ? 'b' : 'l', deleted: !el || el.getAttribute('hidden') === '1',
      min: attr('min'), max: attr('max'), majorUnit: attr('majorUnit'), reverse: false, log: null,
      grid: grid ? paint(child(grid, 'spPr'), theme, { stroke: 'rgb(217 217 217)', width: 0.75 * PT }) : null,
      line: paint(child(el, 'spPr'), theme, { stroke: null, width: 0.75 * PT }),
      text: textStyle(child(el, 'txPr'), { ...text, size: text.size * 0.9 }, theme),
      format: null, sourceLinked: true, crossAx: cross, crosses: 'autoZero', crossesAt: null,
      // Category labels sit below the plot, wherever zero is.
      labelPos: child(el, 'tickLabels') ? (kind === 'cat' ? 'low' : 'nextTo') : 'none', majorTick: 'none', title: null,
      crossBetween: 'between', rotate: null,
    };
  };
  axes.set('c', axisOf(catEl, 'cat', 'c', 'v'));
  axes.set('v', axisOf(valEl, 'val', 'v', 'c'));

  let title: Title | null = null;
  const titleEl = child(chart, 'title');
  if (titleEl) {
    const tx = child(titleEl, 'tx');
    let lines = richLines(first(tx, 'rich'));
    if (!lines.length) lines = [first(tx, 'v')?.textContent ?? 'Chart Title'];
    title = { lines, style: textStyle(child(titleEl, 'txPr') ?? first(tx, 'rich'), { ...text, size: text.size * 1.4 }, theme), rotate: 0 };
  }
  let legend: Legend | null = null;
  const legendEl = child(chart, 'legend');
  if (legendEl) {
    legend = {
      pos: legendEl.getAttribute('pos') ?? 'r',
      style: textStyle(child(legendEl, 'txPr'), text, theme),
      overlay: legendEl.getAttribute('overlay') === '1',
      hidden: new Set(),
    };
  }
  return {
    title,
    legend,
    groups: [{
      type: 'box', dir: 'col', grouping: 'clustered', varyColors: false, series, gapWidth: gap * 100, overlap: -10,
      holeSize: 0, firstSliceAngle: 0, threeD: false, axisIds: ['c', 'v'], radarStyle: '',
    }],
    axes,
    background: paint(child(root, 'spPr') ?? child(chart, 'spPr'), theme, chartFrame(theme)),
    plotBackground: paint(child(plotArea, 'spPr'), theme),
    plotLayout: null,
    text,
    unsupported,
  };
}

function chartFrame(theme: ChartTheme): Partial<Paint> {
  return theme.frame ? { fill: '#fff', stroke: 'rgb(217 217 217)', width: 0.75 * PT } : {};
}

/** A colour 25% darker (box outlines, whiskers). */
export function darker(c: string): string {
  const m = /rgb\((\d+) (\d+) (\d+)/.exec(c);
  if (!m) return c;
  return `rgb(${[1, 2, 3].map((i) => Math.round(Number(m[i]) * 0.75)).join(' ')})`;
}

// ------------------------------------------------------------------ number formats


/** A number as an axis or data label shows it, with an Excel number format. */
export function formatLabel(value: number, code: string | null): string {
  if (!code || code === 'General') return general(value);
  try {
    return formatNumber(value, code, false).text;
  } catch {
    return general(value);
  }
}

function general(value: number): string {
  if (!Number.isFinite(value)) return '';
  if (value === 0) return '0';
  const abs = Math.abs(value);
  if (abs >= 1e11 || abs < 1e-9) return value.toExponential(4).replace(/\.?0+e/, 'E').replace('E+', 'E+');
  return String(+value.toPrecision(10));
}
