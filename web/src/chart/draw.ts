import { darker, formatLabel, type Axis, type Chart, type Group, type Labels, type Marker, type Paint, type Series, type TextStyle } from './read';

// Lays out and draws a chart model as SVG markup, approximating Office's automatic layout:
// title on top, legend on its side, axis labels around a plot area, value axes scaled to
// "nice" units. The SVG has no ids, so copies of it (page thumbnails) stay valid.

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface Scale {
  min: number;
  max: number;
  unit: number;
  ticks: number[];
  log: number | null;
  /** 0..1 along the axis (before reversing). */
  frac(v: number): number;
}

/** One axis as placed: its model, scale (value axes) and which way it runs on screen. */
interface Placed {
  axis: Axis;
  vertical: boolean;
  /** Value axes */
  scale: Scale | null;
  /** Category axes: names and whether points sit between ticks or on them. */
  categories: string[];
  between: boolean;
  /** Formats for tick labels. */
  format: string | null;
  /** Category label lines (wrapped) and rotation. */
  wrapped: string[][];
  rotate: number;
}

const canvas = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;

function fontCss(s: TextStyle): string {
  return `${s.italic ? 'italic ' : ''}${s.bold ? 'bold ' : ''}${s.size.toFixed(2)}px ${s.font}`;
}

function measure(text: string, s: TextStyle): number {
  if (!canvas) return text.length * s.size * 0.5;
  canvas.font = fontCss(s);
  return canvas.measureText(text).width;
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const n2 = (v: number) => +v.toFixed(2);

function textSvg(x: number, y: number, text: string, s: TextStyle, anchor = 'middle', rotate = 0): string {
  const font = s.font.replace(/"/g, "'");
  const rot = rotate ? ` transform="rotate(${n2(rotate)} ${n2(x)} ${n2(y)})"` : '';
  return `<text x="${n2(x)}" y="${n2(y)}" font-family="${esc(font)}" font-size="${n2(s.size)}" fill="${s.colour}"${s.bold ? ' font-weight="bold"' : ''}${s.italic ? ' font-style="italic"' : ''} text-anchor="${anchor}" dominant-baseline="central"${rot}>${esc(text)}</text>`;
}

function fillAttrs(p: Paint): string {
  return `fill="${p.fill ?? 'none'}"${strokeAttrs(p)}`;
}

function strokeAttrs(p: Paint): string {
  if (!p.stroke) return ' stroke="none"';
  const dash = p.dash ? ` stroke-dasharray="${p.dash.map((d) => n2(d * p.width)).join(' ')}"` : '';
  return ` stroke="${p.stroke}" stroke-width="${n2(p.width)}"${dash}`;
}

const lineHeight = (s: TextStyle) => s.size * 1.2;

export function drawChart(chart: Chart, w: number, h: number): string {
  const out: string[] = [];
  const bg = chart.background;
  if (bg.fill || bg.stroke) {
    const inset = bg.stroke ? bg.width / 2 : 0;
    out.push(`<rect x="${n2(inset)}" y="${n2(inset)}" width="${n2(Math.max(0, w - 2 * inset))}" height="${n2(Math.max(0, h - 2 * inset))}" ${fillAttrs(bg)}/>`);
  }
  const pad = Math.max(5, Math.min(10, Math.min(w, h) * 0.025));
  let area: Rect = { x: pad, y: pad, w: w - 2 * pad, h: h - 2 * pad };

  if (chart.title) {
    const t = chart.title;
    const lh = lineHeight(t.style);
    t.lines.forEach((line, i) => out.push(textSvg(w / 2, area.y + lh * (i + 0.5), line, t.style)));
    const used = lh * t.lines.length + t.style.size * 0.4;
    area = { ...area, y: area.y + used, h: area.h - used };
  }

  if (chart.legend) area = drawLegend(chart, area, out);
  if (area.w < 10 || area.h < 10) return wrap(out, w, h);

  const main = chart.groups[0];
  if (main.type === 'pie' || main.type === 'doughnut') drawPie(main, area, out);
  else if (main.type === 'radar') drawRadar(chart, main, area, out);
  else drawAxes(chart, area, w, h, out);
  return wrap(out, w, h);
}

function wrap(out: string[], w: number, h: number): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${n2(w)}" height="${n2(h)}" viewBox="0 0 ${n2(w)} ${n2(h)}" style="overflow:visible;display:block">${out.join('')}</svg>`;
}

// ------------------------------------------------------------------ legend

interface Entry {
  text: string;
  paint: Paint;
  line: boolean;
  marker: Marker | null;
}

function legendEntries(chart: Chart): Entry[] {
  const g = chart.groups[0];
  const series = chart.groups.flatMap((x) => x.series);
  const byPoint = (g.type === 'pie' || g.type === 'doughnut' || g.varyColors) && series.length === 1;
  let entries: Entry[];
  if (byPoint) {
    const s = series[0];
    entries = s.categories.map((c, i) => ({ text: c, paint: s.points.get(i) ?? s.paint, line: false, marker: null }));
  } else {
    entries = chart.groups.flatMap((grp) => grp.series.map((s) => {
      const line = grp.type === 'line' || grp.type === 'scatter' || grp.type === 'radar';
      return { text: s.name, paint: s.paint, line: line && !!s.paint.stroke, marker: line ? s.marker : null };
    }));
  }
  const shown = entries.filter((_, i) => !chart.legend!.hidden.has(i));
  // A legend beside a vertical stack lists the series top to bottom, as they are stacked.
  const stacked = !byPoint && chart.groups.length === 1 && (g.grouping === 'stacked' || g.grouping === 'percentStacked')
    && (g.type === 'area' || (g.type === 'bar' && g.dir === 'col'));
  return stacked && (chart.legend!.pos === 'r' || chart.legend!.pos === 'l') ? shown.reverse() : shown;
}

function drawLegend(chart: Chart, area: Rect, out: string[]): Rect {
  const legend = chart.legend!;
  const entries = legendEntries(chart);
  if (!entries.length) return area;
  const st = legend.style;
  const lh = lineHeight(st);
  const keyW = (e: Entry) => (e.line ? st.size * 1.6 : st.size * 0.55);
  const gap = st.size * 0.3;
  const widths = entries.map((e) => keyW(e) + gap + measure(e.text, st));
  const spacing = st.size * 0.9;
  const horizontal = legend.pos === 'b' || legend.pos === 't';
  const items: { e: Entry; x: number; y: number }[] = [];
  let box: Rect;
  if (horizontal) {
    // Rows of entries, each row centred.
    const rows: number[][] = [[]];
    let rowW = 0;
    widths.forEach((ew, i) => {
      if (rows[rows.length - 1].length && rowW + spacing + ew > area.w) {
        rows.push([]);
        rowW = 0;
      }
      rowW += (rows[rows.length - 1].length ? spacing : 0) + ew;
      rows[rows.length - 1].push(i);
    });
    const height = rows.length * lh;
    const top = legend.pos === 't' ? area.y : area.y + area.h - height;
    rows.forEach((row, r) => {
      const total = row.reduce((sum, i) => sum + widths[i], 0) + spacing * (row.length - 1);
      let x = area.x + (area.w - total) / 2;
      for (const i of row) {
        items.push({ e: entries[i], x, y: top + lh * (r + 0.5) });
        x += widths[i] + spacing;
      }
    });
    box = { x: area.x, y: top, w: area.w, h: height };
  } else {
    const width = Math.min(area.w * 0.4, Math.max(...widths));
    const height = entries.length * lh;
    const left = legend.pos === 'l' ? area.x : area.x + area.w - width;
    const top = legend.pos === 'tr' ? area.y : area.y + Math.max(0, (area.h - height) / 2);
    entries.forEach((e, i) => items.push({ e, x: left, y: top + lh * (i + 0.5) }));
    box = { x: left, y: top, w: width, h: height };
  }
  for (const { e, x, y } of items) {
    const kw = keyW(e);
    if (e.line) {
      out.push(`<line x1="${n2(x)}" y1="${n2(y)}" x2="${n2(x + kw)}" y2="${n2(y)}"${strokeAttrs({ ...e.paint, width: Math.min(e.paint.width, 3) })}/>`);
      if (e.marker) out.push(markerSvg(e.marker, x + kw / 2, y, Math.min(e.marker.size, st.size * 0.7)));
    } else if (e.marker) {
      out.push(markerSvg(e.marker, x + kw / 2, y, Math.min(e.marker.size, st.size * 0.7)));
    } else {
      const s = st.size * 0.55;
      const p = e.paint.fill ? e.paint : { ...e.paint, fill: e.paint.stroke };
      out.push(`<rect x="${n2(x)}" y="${n2(y - s / 2)}" width="${n2(s)}" height="${n2(s)}" fill="${p.fill ?? 'none'}"/>`);
    }
    out.push(textSvg(x + kw + gap, y, e.text, st, 'start'));
  }
  if (legend.overlay) return area;
  const sep = st.size * 0.6;
  switch (legend.pos) {
    case 'b':
      return { ...area, h: area.h - box.h - sep };
    case 't':
      return { ...area, y: area.y + box.h + sep, h: area.h - box.h - sep };
    case 'l':
      return { ...area, x: area.x + box.w + sep, w: area.w - box.w - sep };
    default:
      return { ...area, w: area.w - box.w - sep };
  }
}

// ------------------------------------------------------------------ markers

function markerSvg(m: Marker, x: number, y: number, size = m.size): string {
  const r = size / 2;
  const p = m.paint;
  const attrs = `fill="${p.fill ?? 'none'}"${strokeAttrs({ ...p, dash: null })}`;
  const lineAttrs = strokeAttrs({ ...p, stroke: p.stroke ?? p.fill, width: Math.max(p.width, size / 6), dash: null });
  switch (m.symbol) {
    case 'circle':
    case 'dot':
      return `<circle cx="${n2(x)}" cy="${n2(y)}" r="${n2(m.symbol === 'dot' ? r / 2 : r)}" ${attrs}/>`;
    case 'square':
      return `<rect x="${n2(x - r)}" y="${n2(y - r)}" width="${n2(size)}" height="${n2(size)}" ${attrs}/>`;
    case 'diamond':
      return `<path d="M${n2(x)} ${n2(y - r)}L${n2(x + r)} ${n2(y)}L${n2(x)} ${n2(y + r)}L${n2(x - r)} ${n2(y)}Z" ${attrs}/>`;
    case 'triangle':
      return `<path d="M${n2(x)} ${n2(y - r)}L${n2(x + r)} ${n2(y + r)}L${n2(x - r)} ${n2(y + r)}Z" ${attrs}/>`;
    case 'x':
      return `<path d="M${n2(x - r)} ${n2(y - r)}L${n2(x + r)} ${n2(y + r)}M${n2(x + r)} ${n2(y - r)}L${n2(x - r)} ${n2(y + r)}" fill="none"${lineAttrs}/>`;
    case 'plus':
      return `<path d="M${n2(x - r)} ${n2(y)}L${n2(x + r)} ${n2(y)}M${n2(x)} ${n2(y - r)}L${n2(x)} ${n2(y + r)}" fill="none"${lineAttrs}/>`;
    case 'star':
      return `<path d="M${n2(x - r)} ${n2(y - r)}L${n2(x + r)} ${n2(y + r)}M${n2(x + r)} ${n2(y - r)}L${n2(x - r)} ${n2(y + r)}M${n2(x)} ${n2(y - r)}L${n2(x)} ${n2(y + r)}" fill="none"${lineAttrs}/>`;
    case 'dash':
      return `<rect x="${n2(x - r)}" y="${n2(y - size / 8)}" width="${n2(size)}" height="${n2(size / 4)}" ${attrs}/>`;
  }
  return '';
}

// ------------------------------------------------------------------ pies

function drawPie(g: Group, area: Rect, out: string[]): void {
  const s = g.series[0];
  const values = s.values.map((v) => (v !== null && v > 0 ? v : 0));
  const total = values.reduce((a, b) => a + b, 0);
  if (!total) return;
  const labels = s.labels;
  const outside = labels && (labels.pos === 'outEnd');
  const margin = outside ? labels.style.size * 2.2 : 0;
  let r = Math.max(4, Math.min(area.w, area.h) / 2 - margin);
  const maxExplosion = Math.max(s.explosion, ...s.pointExplosion.values());
  if (maxExplosion) r /= 1 + maxExplosion / 100;
  const cx = area.x + area.w / 2;
  const cy = area.y + area.h / 2;
  const hole = g.type === 'doughnut' ? (r * g.holeSize) / 100 : 0;
  let angle = (g.firstSliceAngle * Math.PI) / 180;
  const labelSvg: string[] = [];
  values.forEach((v, i) => {
    if (!v) return;
    const sweep = (v / total) * Math.PI * 2;
    const a0 = angle;
    const a1 = angle + sweep;
    angle = a1;
    const mid = (a0 + a1) / 2;
    const ex = ((s.pointExplosion.get(i) ?? s.explosion) / 100) * r;
    const ox = cx + Math.sin(mid) * ex;
    const oy = cy - Math.cos(mid) * ex;
    const p = s.points.get(i) ?? s.paint;
    const at = (a: number, rad: number) => `${n2(ox + Math.sin(a) * rad)} ${n2(oy - Math.cos(a) * rad)}`;
    const large = sweep > Math.PI ? 1 : 0;
    let d: string;
    if (sweep >= Math.PI * 2 - 1e-6) {
      d = `M${at(0, r)}A${n2(r)} ${n2(r)} 0 1 1 ${at(Math.PI, r)}A${n2(r)} ${n2(r)} 0 1 1 ${at(0, r)}Z`;
      if (hole) d += `M${at(0, hole)}A${n2(hole)} ${n2(hole)} 0 1 0 ${at(Math.PI, hole)}A${n2(hole)} ${n2(hole)} 0 1 0 ${at(0, hole)}Z`;
    } else if (hole) {
      d = `M${at(a0, r)}A${n2(r)} ${n2(r)} 0 ${large} 1 ${at(a1, r)}L${at(a1, hole)}A${n2(hole)} ${n2(hole)} 0 ${large} 0 ${at(a0, hole)}Z`;
    } else {
      d = `M${n2(ox)} ${n2(oy)}L${at(a0, r)}A${n2(r)} ${n2(r)} 0 ${large} 1 ${at(a1, r)}Z`;
    }
    out.push(`<path d="${d}" ${fillAttrs(p)} stroke-linejoin="round" fill-rule="evenodd"/>`);
    if (labels) {
      const text = labelText(labels, s, i, v, v / total);
      const inner = hole ? (hole + r) / 2 : r * (labels.pos === 'inEnd' ? 0.8 : labels.pos === 'ctr' ? 0.5 : 0.65);
      const rad = labels.pos === 'outEnd' ? r + labels.style.size * 0.9 : inner;
      const lx = ox + Math.sin(mid) * rad;
      const ly = oy - Math.cos(mid) * rad;
      const anchor = labels.pos === 'outEnd' ? (Math.sin(mid) > 0.3 ? 'start' : Math.sin(mid) < -0.3 ? 'end' : 'middle') : 'middle';
      labelSvg.push(multiline(lx, ly, text, labels.style, anchor));
    }
  });
  out.push(...labelSvg);
}

function multiline(x: number, y: number, text: string, st: TextStyle, anchor = 'middle'): string {
  const lines = text.split('\n');
  const lh = lineHeight(st);
  return lines.map((l, i) => textSvg(x, y + (i - (lines.length - 1) / 2) * lh, l, st, anchor)).join('');
}

function labelText(l: Labels, s: Series, i: number, v: number, share: number | null): string {
  const parts: string[] = [];
  if (l.series) parts.push(s.name);
  if (l.category) parts.push(s.categories[i] ?? '');
  if (l.value) parts.push(formatLabel(v, l.format ?? s.format));
  if (l.percent && share !== null) parts.push(formatLabel(share, l.format && l.format.includes('%') ? l.format : '0%'));
  return parts.join(l.separator.includes('\n') ? '\n' : l.separator);
}

// ------------------------------------------------------------------ scales

/** Office's automatic axis scale: zero-based unless the data sits far from zero, ~5% headroom. */
function autoScale(dmin: number, dmax: number, length: number, labelSize: number, axis: Axis, vertical: boolean, zeroBased: boolean): Scale {
  if (!Number.isFinite(dmin) || !Number.isFinite(dmax)) [dmin, dmax] = [0, 1];
  if (axis.log && axis.log > 1) return logScale(dmin, dmax, axis);
  if (dmin === dmax) {
    if (dmin === 0) dmax = 1;
    else if (dmin > 0) dmin = 0;
    else dmax = 0;
  }
  let lo = dmin;
  let hi = dmax;
  // Excel starts at zero unless the minimum is above 5/6 of the maximum.
  if (axis.min === null && lo > 0 && (zeroBased || (hi - lo) / hi >= 1 / 6)) lo = 0;
  if (axis.max === null && hi < 0 && (zeroBased || (lo - hi) / lo <= -1 / 6 || true)) hi = 0;
  if (axis.min !== null) lo = axis.min;
  if (axis.max !== null) hi = axis.max;
  const range = hi - lo || 1;
  const targetHi = axis.max === null && hi !== 0 ? hi + 0.05 * range : hi;
  const targetLo = axis.min === null && lo !== 0 ? lo - 0.05 * range : lo;
  const maxIntervals = Math.max(2, Math.floor(length / (labelSize * (vertical ? 1.9 : 5))));
  let unit = axis.majorUnit ?? 0;
  if (!unit) {
    const raw = (targetHi - targetLo) / maxIntervals;
    const pow = 10 ** Math.floor(Math.log10(raw));
    unit = [1, 2, 5, 10].map((m) => m * pow).find((u) => {
      const a = axis.min === null ? Math.floor(targetLo / u + 1e-9) * u : lo;
      const b = axis.max === null ? Math.ceil(targetHi / u - 1e-9) * u : hi;
      return (b - a) / u <= maxIntervals + 1e-9;
    }) ?? 10 * pow;
  }
  const min = axis.min ?? Math.floor(targetLo / unit + 1e-9) * unit;
  const max = axis.max ?? Math.ceil(targetHi / unit - 1e-9) * unit;
  const ticks: number[] = [];
  const decimals = Math.max(0, -Math.floor(Math.log10(unit)) + 2);
  for (let k = 0; min + k * unit <= max + unit * 1e-6 && k < 200; k++) ticks.push(+(min + k * unit).toFixed(decimals));
  const span = max - min || 1;
  return { min, max, unit, ticks, log: null, frac: (v) => (v - min) / span };
}

function logScale(dmin: number, dmax: number, axis: Axis): Scale {
  const base = axis.log!;
  const lg = (v: number) => Math.log(v) / Math.log(base);
  const min = axis.min ?? base ** Math.floor(lg(Math.max(dmin, 1e-12)));
  const max = axis.max ?? base ** Math.ceil(lg(Math.max(dmax, min * base)));
  const ticks: number[] = [];
  for (let v = min; v <= max * 1.000001 && ticks.length < 60; v *= base) ticks.push(v);
  const span = lg(max) - lg(min) || 1;
  return { min, max, unit: base, ticks, log: base, frac: (v) => (lg(Math.max(v, 1e-12)) - lg(min)) / span };
}

// ------------------------------------------------------------------ charts with axes

interface Plot {
  rect: Rect;
  /** Pixel position along an axis for a 0..1 fraction. */
  at(p: Placed, frac: number): number;
}

function catFrac(p: Placed, i: number): number {
  const n = Math.max(1, p.categories.length);
  const f = p.between ? (i + 0.5) / n : n === 1 ? 0.5 : i / (n - 1);
  return p.axis.reverse ? 1 - f : f;
}

function valFrac(p: Placed, v: number): number {
  const f = p.scale!.frac(v);
  return p.axis.reverse ? 1 - f : f;
}

/** Each group's values as drawn: stacked and percent groups add up per category. */
function stackedValues(g: Group): { top: (number | null)[][]; base: number[][] } {
  const n = Math.max(...g.series.map((s) => s.values.length));
  const stacked = g.grouping === 'stacked' || g.grouping === 'percentStacked';
  const totals = new Array(n).fill(0);
  if (g.grouping === 'percentStacked') {
    for (const s of g.series) s.values.forEach((v, i) => (totals[i] += Math.abs(v ?? 0)));
  }
  const pos = new Array(n).fill(0);
  const neg = new Array(n).fill(0);
  const top: (number | null)[][] = [];
  const base: number[][] = [];
  for (const s of g.series) {
    const t: (number | null)[] = [];
    const b: number[] = [];
    for (let i = 0; i < n; i++) {
      let v = s.values[i] ?? null;
      if (v === null) {
        t.push(null);
        b.push(0);
        continue;
      }
      if (g.grouping === 'percentStacked') v = totals[i] ? v / totals[i] : 0;
      if (!stacked) {
        t.push(v);
        b.push(0);
        continue;
      }
      // Area charts stack every series on the previous one; bars stack positives and negatives apart.
      const acc = g.type === 'area' || v >= 0 ? pos : neg;
      b.push(acc[i]);
      acc[i] += v;
      t.push(acc[i]);
    }
    top.push(t);
    base.push(b);
  }
  return { top, base };
}

/** Five-number summary for box & whisker (Excel's quartile methods). */
function boxStats(values: number[], exclusive: boolean) {
  const v = [...values].sort((a, b) => a - b);
  const n = v.length;
  const q = (p: number) => {
    let pos = exclusive ? p * (n + 1) : p * (n - 1) + 1;
    pos = Math.min(Math.max(pos, 1), n);
    const lo = Math.floor(pos);
    const frac = pos - lo;
    return v[lo - 1] + (lo < n ? frac * (v[lo] - v[lo - 1]) : 0);
  };
  const q1 = q(0.25);
  const q3 = q(0.75);
  const median = q(0.5);
  const iqr = q3 - q1;
  const inside = v.filter((x) => x >= q1 - 1.5 * iqr && x <= q3 + 1.5 * iqr);
  return {
    q1, q3, median,
    mean: v.reduce((a, b) => a + b, 0) / n,
    low: inside.length ? inside[0] : q1,
    high: inside.length ? inside[inside.length - 1] : q3,
    outliers: v.filter((x) => x < q1 - 1.5 * iqr || x > q3 + 1.5 * iqr),
    all: v,
  };
}

function drawAxes(chart: Chart, area: Rect, w: number, h: number, out: string[]): void {
  // Which axes each group uses; the first group's axes are the primary pair.
  const axisFor = (id: string) => chart.axes.get(id);
  const groups = chart.groups.filter((g) => g.type !== 'pie' && g.type !== 'doughnut' && g.type !== 'radar');
  const placed = new Map<string, Placed>();

  // Categories: the longest category list among the series using each category axis.
  const usesCat = (g: Group) => g.type !== 'scatter' && g.type !== 'bubble';
  for (const g of groups) {
    const [xId, yId] = g.axisIds;
    const xAxis = axisFor(xId);
    const yAxis = axisFor(yId);
    if (!xAxis || !yAxis) continue;
    if (usesCat(g) && xAxis.kind !== 'val') {
      let categories: string[];
      if (g.type === 'box') {
        categories = [];
        for (const s of g.series) for (const grp of s.groups ?? []) if (!categories.includes(grp.category)) categories.push(grp.category);
      } else {
        const n = Math.max(...g.series.map((s) => Math.max(s.values.length, s.categories.length)));
        const named = g.series.find((s) => s.categories.length)?.categories ?? [];
        categories = Array.from({ length: n }, (_, i) => named[i] ?? String(i + 1));
      }
      const prev = placed.get(xId);
      if (!prev || prev.categories.length < categories.length) {
        placed.set(xId, {
          axis: xAxis, vertical: xAxis.pos === 'l' || xAxis.pos === 'r', scale: null, categories,
          between: (yAxis.crossBetween ?? 'between') === 'between', format: null, wrapped: [], rotate: 0,
        });
      }
    }
  }

  // Value ranges per value axis.
  const ranges = new Map<string, { lo: number; hi: number; zero: boolean; format: string | null; percent: boolean }>();
  const extend = (id: string, vals: (number | null)[], zero: boolean, format: string | null, percent = false) => {
    const r = ranges.get(id) ?? { lo: Infinity, hi: -Infinity, zero: false, format, percent: false };
    for (const v of vals) if (v !== null && Number.isFinite(v)) {
      r.lo = Math.min(r.lo, v);
      r.hi = Math.max(r.hi, v);
    }
    r.zero ||= zero;
    r.percent ||= percent;
    r.format ??= format;
    ranges.set(id, r);
  };
  for (const g of groups) {
    const [xId, yId] = g.axisIds;
    if (g.type === 'scatter' || g.type === 'bubble') {
      for (const s of g.series) {
        extend(xId, s.xs && s.xs.some((x) => x !== null) ? s.xs : s.values.map((_, i) => i + 1), false, null);
        extend(yId, s.values, false, s.format);
      }
    } else if (g.type === 'box') {
      for (const s of g.series) extend(yId, s.groups!.flatMap((x) => x.values), false, s.format);
    } else {
      const { top, base } = stackedValues(g);
      top.forEach((t, k) => extend(yId, [...t, ...base[k]], g.type === 'bar' || g.type === 'area', g.series[k].format, g.grouping === 'percentStacked'));
    }
  }

  // Scale value axes against a first guess of the plot size, then measure their labels.
  for (const [id, r] of ranges) {
    const axis = axisFor(id);
    if (!axis) continue;
    const vertical = axis.pos === 'l' || axis.pos === 'r';
    const length = (vertical ? area.h : area.w) * 0.85;
    const scale = autoScale(r.lo, r.hi, length, axis.text.size, axis, vertical, r.zero);
    let format = axis.format && (!axis.sourceLinked || axis.format !== 'General') ? axis.format : r.format;
    if (axis.sourceLinked && r.format) format = r.format;
    if (r.percent && (!format || format === 'General')) format = '0%';
    placed.set(id, { axis, vertical, scale, categories: [], between: true, format, wrapped: [], rotate: 0 });
  }

  // Reserve room for each visible axis's labels and title.
  let rect: Rect = { ...area };
  const shows = (p: Placed) => !p.axis.deleted && p.axis.labelPos !== 'none';
  const labelsOf = (p: Placed): string[] => p.scale ? p.scale.ticks.map((t) => formatLabel(t, p.format)) : p.categories;
  const reserve = { l: 0, r: 0, t: 0, b: 0 };
  for (const p of placed.values()) {
    if (p.axis.deleted) continue;
    const side = (p.axis.labelPos === 'high' ? opposite(p.axis.pos) : p.axis.pos) as keyof typeof reserve;
    let need = 0;
    if (shows(p)) {
      const st = p.axis.text;
      const texts = labelsOf(p);
      if (p.vertical) {
        need = Math.max(0, ...texts.map((t) => measure(t, st))) + st.size * 0.45;
      } else {
        need = lineHeight(st) + st.size * 0.35;
        if (!p.scale) {
          // Category labels wider than their slot wrap onto more lines, then turn 45°.
          const slot = (area.w * 0.9) / Math.max(1, p.categories.length);
          p.wrapped = texts.map((t) => wrapText(t, slot, st));
          const lines = Math.max(1, ...p.wrapped.map((x) => x.length));
          const tooWide = p.wrapped.some((x) => x.some((line) => measure(line, st) > slot));
          if (p.axis.rotate !== null) p.rotate = p.axis.rotate;
          else if (tooWide || lines > 3) p.rotate = -45;
          if (p.rotate) {
            const longest = Math.max(0, ...texts.map((t) => measure(t, st)));
            need = Math.min(area.h * 0.4, longest * Math.abs(Math.sin((p.rotate * Math.PI) / 180)) + st.size) + st.size * 0.35;
            p.wrapped = texts.map((t) => [t]);
          } else need = lineHeight(st) * lines + st.size * 0.35;
        }
      }
    }
    if (p.axis.title) need += lineHeight(p.axis.title.style) * p.axis.title.lines.length + p.axis.title.style.size * 0.3;
    reserve[side] = Math.max(reserve[side], need);
  }
  // Vertical value labels are centred on their ticks: leave half a line above the top one.
  const vLabel = [...placed.values()].find((p) => p.vertical && p.scale && shows(p));
  const half = vLabel ? lineHeight(vLabel.axis.text) / 2 : 0;
  rect = {
    x: rect.x + reserve.l,
    y: rect.y + Math.max(reserve.t, half),
    w: rect.w - reserve.l - reserve.r,
    h: rect.h - Math.max(reserve.t, half) - Math.max(reserve.b, half),
  };
  // A plot area placed by hand.
  const ml = chart.plotLayout;
  if (ml && ml.w !== null && ml.h !== null) {
    const box = { x: ml.x * w, y: ml.y * h, w: ml.w * w, h: ml.h * h };
    rect = ml.inner ? box : { x: box.x + reserve.l, y: box.y + reserve.t, w: box.w - reserve.l - reserve.r, h: box.h - reserve.t - reserve.b };
  }
  if (rect.w < 4 || rect.h < 4) return;

  const plot: Plot = {
    rect,
    at: (p, f) => (p.vertical ? rect.y + rect.h - f * rect.h : rect.x + f * rect.w),
  };

  const pb = chart.plotBackground;
  if (pb.fill || pb.stroke) out.push(`<rect x="${n2(rect.x)}" y="${n2(rect.y)}" width="${n2(rect.w)}" height="${n2(rect.h)}" ${fillAttrs(pb)}/>`);

  // Gridlines.
  for (const p of placed.values()) {
    if (!p.axis.grid || p.axis.deleted && !p.axis.grid) continue;
    const positions = p.scale ? p.scale.ticks.map((t) => valFrac(p, t))
      : Array.from({ length: p.categories.length + (p.between ? 1 : 0) }, (_, i) => (p.between ? i / p.categories.length : catFrac(p, i)));
    for (const f of positions) {
      const v = plot.at(p, f);
      out.push(p.vertical
        ? `<line x1="${n2(rect.x)}" y1="${n2(v)}" x2="${n2(rect.x + rect.w)}" y2="${n2(v)}"${strokeAttrs(p.axis.grid)}/>`
        : `<line x1="${n2(v)}" y1="${n2(rect.y)}" x2="${n2(v)}" y2="${n2(rect.y + rect.h)}"${strokeAttrs(p.axis.grid)}/>`);
    }
  }

  // Series, group by group.
  const labelSvg: string[] = [];
  for (const g of groups) {
    const x = placed.get(g.axisIds[0]);
    const y = placed.get(g.axisIds[1]);
    if (!x || !y || !y.scale) continue;
    if (g.type === 'bar') drawBars(g, x, y, plot, out, labelSvg);
    else if (g.type === 'line' || g.type === 'area') drawLines(g, x, y, plot, out, labelSvg);
    else if (g.type === 'scatter' || g.type === 'bubble') drawScatter(g, x, y, plot, out, labelSvg);
    else if (g.type === 'box') drawBoxes(g, x, y, plot, out);
  }

  // Axis lines, crossing their partner axis where it says (zero by default).
  for (const p of placed.values()) {
    if (p.axis.deleted) continue;
    const partner = placed.get(p.axis.crossAx);
    const pos = crossPosition(p, partner, plot);
    if (p.axis.line?.stroke) {
      out.push(p.vertical
        ? `<line x1="${n2(pos)}" y1="${n2(rect.y)}" x2="${n2(pos)}" y2="${n2(rect.y + rect.h)}"${strokeAttrs(p.axis.line)}/>`
        : `<line x1="${n2(rect.x)}" y1="${n2(pos)}" x2="${n2(rect.x + rect.w)}" y2="${n2(pos)}"${strokeAttrs(p.axis.line)}/>`);
      if (p.axis.majorTick === 'out' || p.axis.majorTick === 'cross') {
        const fr = p.scale ? p.scale.ticks.map((t) => valFrac(p, t))
          : Array.from({ length: p.categories.length + 1 }, (_, i) => (p.between ? i / p.categories.length : catFrac(p, i)));
        const len = 4;
        const dir = p.axis.pos === 'l' || p.axis.pos === 't' ? -1 : 1;
        for (const f of fr) {
          const v = plot.at(p, f);
          out.push(p.vertical
            ? `<line x1="${n2(pos)}" y1="${n2(v)}" x2="${n2(pos + dir * len)}" y2="${n2(v)}"${strokeAttrs(p.axis.line)}/>`
            : `<line x1="${n2(v)}" y1="${n2(pos)}" x2="${n2(v)}" y2="${n2(pos + dir * len)}"${strokeAttrs(p.axis.line)}/>`);
        }
      }
    }
    if (!shows(p)) continue;
    drawAxisLabels(p, plot, pos, out);
  }
  for (const p of placed.values()) if (!p.axis.deleted && p.axis.title) drawAxisTitle(p, plot, area, out);
  out.push(...labelSvg);
}

function opposite(pos: string): string {
  return ({ l: 'r', r: 'l', t: 'b', b: 't' } as Record<string, string>)[pos] ?? pos;
}

/** Where an axis line sits across the plot: at its partner's zero, minimum or maximum. */
function crossPosition(p: Placed, partner: Placed | undefined, plot: Plot): number {
  const r = plot.rect;
  const edge = p.vertical ? (p.axis.pos === 'r' ? r.x + r.w : r.x) : (p.axis.pos === 't' ? r.y : r.y + r.h);
  if (!partner) return edge;
  let f: number | null = null;
  const crosses = p.axis.crosses;
  if (partner.scale) {
    const s = partner.scale;
    const v = crosses === 'max' ? s.max : crosses === 'min' ? s.min : crosses === 'at' && p.axis.crossesAt !== null ? p.axis.crossesAt : s.log ? s.min : Math.min(Math.max(0, s.min), s.max);
    f = valFrac(partner, v);
  } else {
    if (crosses === 'max') f = partner.axis.reverse ? 0 : 1;
    else return edge;
  }
  return plot.at(partner, f);
}

function wrapText(text: string, width: number, st: TextStyle): string[] {
  if (measure(text, st) <= width) return [text];
  const words = text.split(/\s+/);
  const lines: string[] = [];
  let line = '';
  for (const word of words) {
    const next = line ? `${line} ${word}` : word;
    if (line && measure(next, st) > width) {
      lines.push(line);
      line = word;
    } else line = next;
  }
  if (line) lines.push(line);
  return lines;
}

function drawAxisLabels(p: Placed, plot: Plot, pos: number, out: string[]): void {
  const st = p.axis.text;
  const r = plot.rect;
  const low = p.axis.labelPos === 'low';
  const high = p.axis.labelPos === 'high';
  const side = high ? opposite(p.axis.pos) : p.axis.pos;
  let at = low || high ? (side === 'l' ? r.x : side === 'r' ? r.x + r.w : side === 't' ? r.y : r.y + r.h) : pos;
  if (p.vertical) {
    const left = side !== 'r';
    at += left ? -st.size * 0.4 : st.size * 0.4;
    const items = p.scale ? p.scale.ticks.map((t) => ({ f: valFrac(p, t), text: [formatLabel(t, p.format)] }))
      : p.categories.map((c, i) => ({ f: catFrac(p, i), text: [c] }));
    for (const it of items) out.push(textSvg(at, plot.at(p, it.f), it.text[0], st, left ? 'end' : 'start'));
    return;
  }
  const below = side !== 't';
  at += below ? st.size * 0.35 : -st.size * 0.35;
  if (p.scale) {
    for (const t of p.scale.ticks) {
      out.push(textSvg(plot.at(p, valFrac(p, t)), at + (below ? lineHeight(st) / 2 : -lineHeight(st) / 2), formatLabel(t, p.format), st));
    }
    return;
  }
  const lh = lineHeight(st);
  // Too many categories for their labels: show every nth one, as Office does.
  const slot = r.w / Math.max(1, p.categories.length);
  const step = p.rotate ? Math.max(1, Math.ceil((lh * 1.1) / slot)) : 1;
  p.categories.forEach((_, i) => {
    if (i % step) return;
    const x = plot.at(p, catFrac(p, i));
    const lines = p.wrapped[i] ?? [p.categories[i]];
    if (p.rotate) {
      const y = at + st.size * 0.3;
      out.push(textSvg(x, y, lines.join(' '), st, below ? 'end' : 'start', p.rotate));
      return;
    }
    lines.forEach((line, k) => out.push(textSvg(x, below ? at + lh * (k + 0.5) : at - lh * (lines.length - k - 0.5), line, st)));
  });
}

function drawAxisTitle(p: Placed, plot: Plot, area: Rect, out: string[]): void {
  const t = p.axis.title!;
  const r = plot.rect;
  const lh = lineHeight(t.style);
  const text = t.lines.join(' ');
  if (p.vertical) {
    const x = p.axis.pos === 'r' ? area.x + area.w - lh / 2 : area.x + lh / 2;
    out.push(textSvg(x, r.y + r.h / 2, text, t.style, 'middle', p.axis.pos === 'r' ? 90 : -90));
  } else {
    const y = p.axis.pos === 't' ? area.y + lh / 2 : area.y + area.h - lh / 2;
    out.push(textSvg(r.x + r.w / 2, y, text, t.style));
  }
}

// ------------------------------------------------------------------ bars

function drawBars(g: Group, x: Placed, y: Placed, plot: Plot, out: string[], labels: string[]): void {
  const { top, base } = stackedValues(g);
  const n = Math.max(1, x.categories.length);
  const band = (x.vertical ? plot.rect.h : plot.rect.w) / n;
  const stacked = g.grouping === 'stacked' || g.grouping === 'percentStacked';
  const count = stacked ? 1 : g.series.length;
  const overlap = stacked ? 100 : Math.max(-100, Math.min(100, g.overlap));
  const width = band / (count - ((count - 1) * overlap) / 100 + g.gapWidth / 100);
  const step = width * (1 - overlap / 100);
  const depth = g.threeD ? Math.min(width * 0.35, 12) : 0;
  const zero = valFrac(y, Math.min(Math.max(0, y.scale!.min), y.scale!.max));
  g.series.forEach((s, k) => {
    const slot = stacked ? 0 : k;
    for (let i = 0; i < n; i++) {
      const v = top[k][i];
      if (v === null || v === undefined) continue;
      const centre = plot.at(x, catFrac(x, i));
      // Along the category axis, series sit left to right (or bottom to top for horizontal bars).
      const dir = x.vertical ? -1 : 1;
      const start = centre - (dir * (width + step * (count - 1))) / 2 + dir * slot * step;
      const end = start + dir * width;
      const b = stacked ? valFrac(y, clamp(base[k][i], y.scale!)) : zero;
      const f = valFrac(y, clamp(v, y.scale!));
      const p0 = plot.at(y, b);
      const p1 = plot.at(y, f);
      const paint = s.points.get(i) ?? (g.varyColors && g.series.length === 1 ? s.points.get(i) ?? s.paint : s.paint);
      const rx = x.vertical ? Math.min(p0, p1) : Math.min(start, end);
      const ry = x.vertical ? Math.min(start, end) : Math.min(p0, p1);
      const rw = x.vertical ? Math.abs(p1 - p0) : width;
      const rh = x.vertical ? width : Math.abs(p1 - p0);
      if (depth) {
        const side = darker(paint.fill ?? '');
        const topFace = shade(paint.fill ?? '', 0.85);
        out.push(`<path d="M${n2(rx + rw)} ${n2(ry)}l${n2(depth)} ${n2(-depth)}v${n2(rh)}l${n2(-depth)} ${n2(depth)}Z" fill="${side}"/>`);
        out.push(`<path d="M${n2(rx)} ${n2(ry)}l${n2(depth)} ${n2(-depth)}h${n2(rw)}l${n2(-depth)} ${n2(depth)}Z" fill="${topFace}"/>`);
      }
      out.push(`<rect x="${n2(rx)}" y="${n2(ry)}" width="${n2(rw)}" height="${n2(rh)}" ${fillAttrs(paint)}/>`);
      if (s.labels) {
        const l = s.labels;
        const text = labelText(l, s, i, s.values[i] ?? 0, null);
        const along = l.pos === 'ctr' ? (p0 + p1) / 2 : l.pos === 'inBase' ? p0 + Math.sign(p1 - p0) * l.style.size * 0.8
          : l.pos === 'inEnd' ? p1 - Math.sign(p1 - p0 || 1) * l.style.size * 0.8 : p1 + Math.sign(p1 - p0 || 1) * (x.vertical ? measure(text, l.style) / 2 + 4 : l.style.size * 0.8);
        const across = (start + end) / 2;
        labels.push(multiline(x.vertical ? along : across, x.vertical ? across : along, text, l.style));
      }
    }
  });
}

function clamp(v: number, s: Scale): number {
  return Math.min(Math.max(v, s.min), s.max);
}

function shade(c: string, k: number): string {
  const m = /rgb\((\d+) (\d+) (\d+)/.exec(c);
  if (!m) return c;
  return `rgb(${[1, 2, 3].map((i) => Math.round(Number(m[i]) * k)).join(' ')})`;
}

// ------------------------------------------------------------------ lines and areas

function smoothPath(pts: [number, number][]): string {
  if (pts.length < 3) return pts.map((p, i) => `${i ? 'L' : 'M'}${n2(p[0])} ${n2(p[1])}`).join('');
  let d = `M${n2(pts[0][0])} ${n2(pts[0][1])}`;
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[i - 1] ?? pts[i];
    const p1 = pts[i];
    const p2 = pts[i + 1];
    const p3 = pts[i + 2] ?? p2;
    // Catmull-Rom through the points, as cubic Béziers.
    const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6];
    const c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
    d += `C${n2(c1[0])} ${n2(c1[1])} ${n2(c2[0])} ${n2(c2[1])} ${n2(p2[0])} ${n2(p2[1])}`;
  }
  return d;
}

function linePath(pts: [number, number][], smooth: boolean): string {
  return smooth ? smoothPath(pts) : pts.map((p, i) => `${i ? 'L' : 'M'}${n2(p[0])} ${n2(p[1])}`).join('');
}

function point(x: Placed, y: Placed, plot: Plot, xf: number, yf: number): [number, number] {
  const a = plot.at(x, xf);
  const b = plot.at(y, yf);
  return x.vertical ? [b, a] : [a, b];
}

function drawLines(g: Group, x: Placed, y: Placed, plot: Plot, out: string[], labels: string[]): void {
  const { top, base } = stackedValues(g);
  const markers: string[] = [];
  g.series.forEach((s, k) => {
    // Runs of points without gaps (blank cells break the line).
    const runs: { i: number; p: [number, number]; b: [number, number] }[][] = [[]];
    top[k].forEach((v, i) => {
      if (v === null) {
        if (runs[runs.length - 1].length) runs.push([]);
        return;
      }
      runs[runs.length - 1].push({
        i,
        p: point(x, y, plot, catFrac(x, i), valFrac(y, clamp(v, y.scale!))),
        b: point(x, y, plot, catFrac(x, i), valFrac(y, clamp(base[k][i], y.scale!))),
      });
    });
    for (const run of runs) {
      if (!run.length) continue;
      if (g.type === 'area') {
        const upper = run.map((r) => r.p);
        const lower = run.map((r) => r.b).reverse();
        const d = `${linePath(upper, false)}${lower.map((p) => `L${n2(p[0])} ${n2(p[1])}`).join('')}Z`;
        out.push(`<path d="${d}" ${fillAttrs(s.paint)} stroke-linejoin="round"/>`);
      } else if (s.paint.stroke) {
        out.push(`<path d="${linePath(run.map((r) => r.p), s.smooth)}" fill="none"${strokeAttrs(s.paint)} stroke-linejoin="round" stroke-linecap="round"/>`);
      }
      for (const r of run) {
        if (s.marker) markers.push(markerSvg(s.marker, r.p[0], r.p[1]));
        if (s.labels) labels.push(pointLabel(s.labels, s, r.i, s.values[r.i] ?? 0, r.p, s.marker));
      }
    }
  });
  out.push(...markers);
}

function pointLabel(l: Labels, s: Series, i: number, v: number, p: [number, number], m: Marker | null): string {
  const text = labelText(l, s, i, v, null);
  const off = (m ? m.size / 2 : 2) + l.style.size * 0.3;
  switch (l.pos) {
    case 'ctr':
      return multiline(p[0], p[1], text, l.style);
    case 't':
      return multiline(p[0], p[1] - off - l.style.size / 2, text, l.style);
    case 'b':
      return multiline(p[0], p[1] + off + l.style.size / 2, text, l.style);
    case 'l':
      return multiline(p[0] - off, p[1], text, l.style, 'end');
    default:
      return multiline(p[0] + off, p[1], text, l.style, 'start');
  }
}

// ------------------------------------------------------------------ scatter and bubbles

function drawScatter(g: Group, x: Placed, y: Placed, plot: Plot, out: string[], labels: string[]): void {
  const markers: string[] = [];
  const sizes = g.type === 'bubble' ? g.series.flatMap((s) => s.sizes ?? []).filter((v): v is number => v !== null && v > 0) : [];
  const maxSize = Math.max(1, ...sizes);
  const maxR = Math.min(plot.rect.w, plot.rect.h) * 0.125;
  g.series.forEach((s) => {
    const pts: { i: number; p: [number, number] }[] = [];
    s.values.forEach((v, i) => {
      const xv = s.xs?.[i] ?? (s.xs && s.xs.some((q) => q !== null) ? null : i + 1);
      if (v === null || xv === null) return;
      pts.push({ i, p: point(x, y, plot, valFrac(x, clamp(xv, x.scale!)), valFrac(y, clamp(v, y.scale!))) });
    });
    if (g.type === 'bubble') {
      for (const { i, p } of pts) {
        const size = s.sizes?.[i] ?? 0;
        if (size <= 0) continue;
        const r = Math.sqrt(size / maxSize) * maxR;
        const paint = s.points.get(i) ?? s.paint;
        out.push(`<circle cx="${n2(p[0])}" cy="${n2(p[1])}" r="${n2(r)}" ${fillAttrs({ ...paint, fill: paint.fill ?? paint.stroke })} fill-opacity="0.75"/>`);
      }
      return;
    }
    if (s.paint.stroke && pts.length > 1) {
      out.push(`<path d="${linePath(pts.map((q) => q.p), s.smooth)}" fill="none"${strokeAttrs(s.paint)} stroke-linejoin="round" stroke-linecap="round"/>`);
    }
    for (const { i, p } of pts) {
      if (s.marker) markers.push(markerSvg(s.marker, p[0], p[1]));
      if (s.labels) labels.push(pointLabel(s.labels, s, i, s.values[i] ?? 0, p, s.marker));
    }
  });
  out.push(...markers);
}

// ------------------------------------------------------------------ box & whisker

function drawBoxes(g: Group, x: Placed, y: Placed, plot: Plot, out: string[]): void {
  const n = Math.max(1, x.categories.length);
  const band = plot.rect.w / n;
  const count = g.series.length;
  const width = band / (count + (count - 1) * 0.1 + g.gapWidth / 100);
  const step = width * 1.1;
  const Y = (v: number) => plot.at(y, valFrac(y, clamp(v, y.scale!)));
  g.series.forEach((s, k) => {
    const line = s.paint.stroke ?? darker(s.paint.fill ?? '#888');
    const lw = Math.max(0.75, s.paint.width);
    for (const grp of s.groups ?? []) {
      const i = x.categories.indexOf(grp.category);
      if (i < 0 || !grp.values.length) continue;
      const st = boxStats(grp.values, s.box?.exclusive ?? true);
      const centre = plot.at(x, catFrac(x, i));
      const left = centre - (width + step * (count - 1)) / 2 + k * step;
      const mid = left + width / 2;
      const stroke = `stroke="${line}" stroke-width="${n2(lw)}"`;
      const cap = width / 4;
      // Whiskers with end caps.
      out.push(`<path d="M${n2(mid)} ${n2(Y(st.q3))}V${n2(Y(st.high))}M${n2(mid - cap)} ${n2(Y(st.high))}H${n2(mid + cap)}M${n2(mid)} ${n2(Y(st.q1))}V${n2(Y(st.low))}M${n2(mid - cap)} ${n2(Y(st.low))}H${n2(mid + cap)}" fill="none" ${stroke}/>`);
      const top = Y(st.q3);
      const bottom = Y(st.q1);
      out.push(`<rect x="${n2(left)}" y="${n2(top)}" width="${n2(width)}" height="${n2(Math.max(0, bottom - top))}" fill="${s.paint.fill ?? 'none'}" ${stroke}/>`);
      out.push(`<line x1="${n2(left)}" y1="${n2(Y(st.median))}" x2="${n2(left + width)}" y2="${n2(Y(st.median))}" ${stroke}/>`);
      if (s.box?.mean) {
        const r = Math.min(width / 4, 5);
        const my = Y(st.mean);
        // On a collapsed box the cross is drawn bolder in the series colour, as Office does.
        const flat = bottom - top < r * 2;
        const markStroke = flat ? `stroke="${s.paint.fill ?? line}" stroke-width="${n2(Math.max(2.5, r * 0.6))}" stroke-linecap="round"` : stroke;
        out.push(`<path d="M${n2(mid - r)} ${n2(my - r)}L${n2(mid + r)} ${n2(my + r)}M${n2(mid + r)} ${n2(my - r)}L${n2(mid - r)} ${n2(my + r)}" fill="none" ${markStroke}/>`);
      }
      if (s.box?.outliers) {
        for (const o of st.outliers) out.push(`<circle cx="${n2(mid)}" cy="${n2(Y(o))}" r="${n2(Math.min(width / 10, 2.5))}" fill="${line}"/>`);
      }
      if (s.box?.inner) {
        for (const v of st.all) if (!st.outliers.includes(v)) out.push(`<circle cx="${n2(mid)}" cy="${n2(Y(v))}" r="${n2(Math.min(width / 10, 2.5))}" fill="none" stroke="${line}"/>`);
      }
    }
  });
}

// ------------------------------------------------------------------ radar

function drawRadar(chart: Chart, g: Group, area: Rect, out: string[]): void {
  const n = Math.max(...g.series.map((s) => s.values.length));
  if (n < 3) return;
  const valAxis = chart.axes.get(g.axisIds[1]);
  const catAxis = chart.axes.get(g.axisIds[0]);
  const st = catAxis?.text ?? chart.text;
  const cats = g.series[0].categories;
  const labelRoom = Math.max(0, ...cats.map((c) => measure(c, st))) + st.size;
  const r = Math.max(4, Math.min(area.w / 2 - labelRoom, area.h / 2 - lineHeight(st) * 1.5));
  const cx = area.x + area.w / 2;
  const cy = area.y + area.h / 2;
  const vals = g.series.flatMap((s) => s.values).filter((v): v is number => v !== null);
  const axis = valAxis ?? ({ min: null, max: null, majorUnit: null, log: null } as Axis);
  const scale = autoScale(Math.min(0, ...vals), Math.max(...vals), r, axis.text?.size ?? st.size, axis, true, true);
  const at = (i: number, f: number): [number, number] => {
    const a = (i / n) * Math.PI * 2;
    return [cx + Math.sin(a) * r * f, cy - Math.cos(a) * r * f];
  };
  const grid = valAxis?.grid ?? { stroke: 'rgb(217 217 217)', width: 1, fill: null, dash: null };
  for (const t of scale.ticks) {
    const f = scale.frac(t);
    const d = Array.from({ length: n }, (_, i) => at(i, f)).map((p, i) => `${i ? 'L' : 'M'}${n2(p[0])} ${n2(p[1])}`).join('') + 'Z';
    out.push(`<path d="${d}" fill="none"${strokeAttrs(grid)}/>`);
  }
  for (let i = 0; i < n; i++) {
    const [ex, ey] = at(i, 1);
    out.push(`<line x1="${n2(cx)}" y1="${n2(cy)}" x2="${n2(ex)}" y2="${n2(ey)}"${strokeAttrs(grid)}/>`);
    const [lx, ly] = at(i, 1 + (st.size * 0.8) / r);
    const sin = Math.sin((i / n) * Math.PI * 2);
    out.push(textSvg(lx, ly, cats[i] ?? String(i + 1), st, sin > 0.2 ? 'start' : sin < -0.2 ? 'end' : 'middle'));
  }
  const fmt = g.series[0].format;
  if (valAxis && !valAxis.deleted) for (const t of scale.ticks) out.push(textSvg(cx - 4, cy - scale.frac(t) * r, formatLabel(t, fmt), valAxis.text, 'end'));
  for (const s of g.series) {
    const pts = s.values.map((v, i) => at(i, scale.frac(v ?? 0)));
    const d = pts.map((p, i) => `${i ? 'L' : 'M'}${n2(p[0])} ${n2(p[1])}`).join('') + 'Z';
    if (g.radarStyle === 'filled') out.push(`<path d="${d}" fill="${s.paint.fill ?? s.paint.stroke}" fill-opacity="0.8"/>`);
    else out.push(`<path d="${d}" fill="none"${strokeAttrs(s.paint)} stroke-linejoin="round"/>`);
    if (s.marker) pts.forEach((p) => out.push(markerSvg(s.marker!, p[0], p[1])));
  }
}
