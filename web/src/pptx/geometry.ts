import { child, children } from '../ooxml';

// Outlines of DrawingML shapes as SVG path data, in the shape's own pixel box (w × h).
// PowerPoint defines ~180 preset shapes; the common ones are drawn here, a few rarer ones are
// approximated, and the rest fall back to a rectangle.

export interface Outline {
  d: string;
  /** Open paths (lines, braces) are stroked but never filled. */
  open?: boolean;
}

type Adj = (name: string, fallback: number) => number;

const f = (n: number) => +n.toFixed(2);
const poly = (pts: [number, number][]) => `M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join('L')}Z`;

function ellipse(cx: number, cy: number, rx: number, ry: number): string {
  return `M${f(cx - rx)} ${f(cy)}A${f(rx)} ${f(ry)} 0 1 0 ${f(cx + rx)} ${f(cy)}A${f(rx)} ${f(ry)} 0 1 0 ${f(cx - rx)} ${f(cy)}Z`;
}

function roundRect(x: number, y: number, w: number, h: number, r: number): string {
  r = Math.min(r, w / 2, h / 2);
  return `M${f(x + r)} ${f(y)}H${f(x + w - r)}A${f(r)} ${f(r)} 0 0 1 ${f(x + w)} ${f(y + r)}V${f(y + h - r)}A${f(r)} ${f(r)} 0 0 1 ${f(x + w - r)} ${f(y + h)}H${f(x + r)}A${f(r)} ${f(r)} 0 0 1 ${f(x)} ${f(y + h - r)}V${f(y + r)}A${f(r)} ${f(r)} 0 0 1 ${f(x + r)} ${f(y)}Z`;
}

function star(w: number, h: number, points: number, inner: number): string {
  const pts: [number, number][] = [];
  for (let i = 0; i < points * 2; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / points;
    const k = i % 2 ? inner : 1;
    pts.push([w / 2 + (Math.cos(a) * k * w) / 2, h / 2 + (Math.sin(a) * k * h) / 2]);
  }
  return poly(pts);
}

function regular(w: number, h: number, sides: number): string {
  const pts: [number, number][] = [];
  for (let i = 0; i < sides; i++) {
    const a = -Math.PI / 2 + (i * 2 * Math.PI) / sides;
    pts.push([w / 2 + (Math.cos(a) * w) / 2, h / 2 + (Math.sin(a) * h) / 2]);
  }
  return poly(pts);
}

/** A bumpy outline around an ellipse, standing in for the cloud shapes. */
function cloud(w: number, h: number): string {
  const n = 10;
  let d = '';
  for (let i = 0; i <= n; i++) {
    const a = (i * 2 * Math.PI) / n;
    const x = w / 2 + Math.cos(a) * w * 0.42;
    const y = h / 2 + Math.sin(a) * h * 0.4;
    if (!i) d = `M${f(x)} ${f(y)}`;
    else d += `A${f(w * 0.14)} ${f(h * 0.14)} 0 0 1 ${f(x)} ${f(y)}`;
  }
  return `${d}Z`;
}

function arrow(w: number, h: number, adj: Adj): [number, number][] {
  const ss = Math.min(w, h);
  const a1 = adj('adj1', 50000) / 100000;
  const a2 = adj('adj2', 50000) / 100000;
  const head = Math.min(w, ss * a2);
  const y1 = (h * (1 - a1)) / 2;
  return [[0, y1], [w - head, y1], [w - head, 0], [w, h / 2], [w - head, h], [w - head, h - y1], [0, h - y1]];
}

const turn = (pts: [number, number][], w: number, h: number, how: 'left' | 'up' | 'down'): [number, number][] =>
  pts.map(([x, y]) => (how === 'left' ? [w - x, y] : how === 'up' ? [y * (w / h), h - x * (h / w)] : [y * (w / h), x * (h / w)]));

export function presetOutline(name: string, w: number, h: number, adj: Adj): Outline {
  const ss = Math.min(w, h);
  const a = (n: string, d: number) => adj(n, d) / 100000;
  switch (name) {
    case 'ellipse': case 'flowChartConnector': case 'smileyFace': case 'flowChartOr': case 'flowChartSummingJunction': case 'chord': case 'pie':
      return { d: ellipse(w / 2, h / 2, w / 2, h / 2) };
    case 'roundRect': case 'flowChartAlternateProcess':
      return { d: roundRect(0, 0, w, h, ss * (name === 'roundRect' ? a('adj', 16667) : 1 / 6)) };
    case 'round2SameRect': {
      const r = ss * a('adj1', 16667);
      return { d: `M${f(r)} 0H${f(w - r)}A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)}V${f(h)}H0V${f(r)}A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0Z` };
    }
    case 'round1Rect': {
      const r = ss * a('adj', 16667);
      return { d: `M0 0H${f(w - r)}A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)}V${f(h)}H0Z` };
    }
    case 'snip1Rect': {
      const s = ss * a('adj', 16667);
      return { d: poly([[0, 0], [w - s, 0], [w, s], [w, h], [0, h]]) };
    }
    case 'snip2SameRect': {
      const s = ss * a('adj1', 16667);
      return { d: poly([[s, 0], [w - s, 0], [w, s], [w, h], [0, h], [0, s]]) };
    }
    case 'triangle': case 'flowChartExtract':
      return { d: poly([[w * (name === 'triangle' ? a('adj', 50000) : 0.5), 0], [w, h], [0, h]]) };
    case 'flowChartMerge':
      return { d: poly([[0, 0], [w, 0], [w / 2, h]]) };
    case 'rtTriangle':
      return { d: poly([[0, 0], [w, h], [0, h]]) };
    case 'diamond': case 'flowChartDecision':
      return { d: poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]) };
    case 'parallelogram': case 'flowChartInputOutput': {
      const o = name === 'parallelogram' ? Math.min(w, ss * a('adj', 25000)) : w / 5;
      return { d: poly([[o, 0], [w, 0], [w - o, h], [0, h]]) };
    }
    case 'trapezoid': case 'flowChartManualOperation': {
      const o = name === 'trapezoid' ? Math.min(w / 2, ss * a('adj', 25000)) : w / 5;
      return name === 'trapezoid' ? { d: poly([[o, 0], [w - o, 0], [w, h], [0, h]]) } : { d: poly([[0, 0], [w, 0], [w - o, h], [o, h]]) };
    }
    case 'pentagon':
      return { d: regular(w, h, 5) };
    case 'hexagon': {
      const o = Math.min(w / 2, ss * a('adj', 25000));
      return { d: poly([[o, 0], [w - o, 0], [w, h / 2], [w - o, h], [o, h], [0, h / 2]]) };
    }
    case 'heptagon':
      return { d: regular(w, h, 7) };
    case 'octagon': {
      const o = ss * a('adj', 29289);
      return { d: poly([[o, 0], [w - o, 0], [w, o], [w, h - o], [w - o, h], [o, h], [0, h - o], [0, o]]) };
    }
    case 'decagon':
      return { d: regular(w, h, 10) };
    case 'dodecagon':
      return { d: regular(w, h, 12) };
    case 'homePlate': case 'flowChartOffpageConnector': {
      const x = w - Math.min(w, ss * a('adj', 50000));
      return name === 'homePlate'
        ? { d: poly([[0, 0], [x, 0], [w, h / 2], [x, h], [0, h]]) }
        : { d: poly([[0, 0], [w, 0], [w, h * 0.8], [w / 2, h], [0, h * 0.8]]) };
    }
    case 'chevron': {
      const x = Math.min(w, ss * a('adj', 50000));
      return { d: poly([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h], [x, h / 2]]) };
    }
    case 'plus': case 'flowChartPredefinedProcess': {
      if (name === 'flowChartPredefinedProcess') {
        return { d: `M0 0H${f(w)}V${f(h)}H0ZM${f(w / 8)} 0V${f(h)}M${f((w * 7) / 8)} 0V${f(h)}` };
      }
      const o = ss * a('adj', 25000);
      return { d: poly([[o, 0], [w - o, 0], [w - o, o], [w, o], [w, h - o], [w - o, h - o], [w - o, h], [o, h], [o, h - o], [0, h - o], [0, o], [o, o]]) };
    }
    case 'star4': return { d: star(w, h, 4, 0.38) };
    case 'star5': return { d: star(w, h, 5, 0.38) };
    case 'star6': return { d: star(w, h, 6, 0.5) };
    case 'star7': return { d: star(w, h, 7, 0.55) };
    case 'star8': return { d: star(w, h, 8, 0.6) };
    case 'star10': return { d: star(w, h, 10, 0.65) };
    case 'star12': return { d: star(w, h, 12, 0.7) };
    case 'star16': case 'star24': case 'star32': case 'irregularSeal1': case 'irregularSeal2':
      return { d: star(w, h, 16, 0.8) };
    case 'rightArrow': case 'notchedRightArrow': case 'stripedRightArrow': case 'curvedRightArrow':
      return { d: poly(arrow(w, h, adj)) };
    case 'leftArrow': case 'curvedLeftArrow':
      return { d: poly(turn(arrow(w, h, adj), w, h, 'left')) };
    case 'upArrow': case 'curvedUpArrow':
      return { d: poly(turn(arrow(h, w, adj), w, h, 'up')) };
    case 'downArrow': case 'curvedDownArrow':
      return { d: poly(turn(arrow(h, w, adj), w, h, 'down')) };
    case 'leftRightArrow': {
      const head = Math.min(w / 2, ss * a('adj2', 50000));
      const y1 = (h * (1 - a('adj1', 50000))) / 2;
      return { d: poly([[0, h / 2], [head, 0], [head, y1], [w - head, y1], [w - head, 0], [w, h / 2], [w - head, h], [w - head, h - y1], [head, h - y1], [head, h]]) };
    }
    case 'upDownArrow': {
      const head = Math.min(h / 2, ss * a('adj2', 50000));
      const x1 = (w * (1 - a('adj1', 50000))) / 2;
      return { d: poly([[w / 2, 0], [w, head], [w - x1, head], [w - x1, h - head], [w, h - head], [w / 2, h], [0, h - head], [x1, h - head], [x1, head], [0, head]]) };
    }
    case 'line': case 'straightConnector1': case 'curvedConnector2': case 'curvedConnector3': case 'curvedConnector4': case 'curvedConnector5': case 'arc':
      return { d: `M0 0L${f(w)} ${f(h)}`, open: true };
    case 'bentConnector2':
      return { d: `M0 0H${f(w)}V${f(h)}`, open: true };
    case 'bentConnector3': case 'bentConnector4': case 'bentConnector5': {
      const x = w * a('adj1', 50000);
      return { d: `M0 0H${f(x)}V${f(h)}H${f(w)}`, open: true };
    }
    case 'flowChartTerminator':
      return { d: roundRect(0, 0, w, h, Math.min(w, h) / 2) };
    case 'flowChartDocument':
      return { d: `M0 0H${f(w)}V${f(h * 0.83)}C${f(w * 0.72)} ${f(h * 0.62)} ${f(w * 0.36)} ${f(h * 1.08)} 0 ${f(h * 0.9)}Z` };
    case 'flowChartManualInput':
      return { d: poly([[0, h / 5], [w, 0], [w, h], [0, h]]) };
    case 'flowChartPunchedCard':
      return { d: poly([[w / 5, 0], [w, 0], [w, h], [0, h], [0, h / 5]]) };
    case 'flowChartDelay':
      return { d: `M0 0H${f(w / 2)}A${f(w / 2)} ${f(h / 2)} 0 0 1 ${f(w / 2)} ${f(h)}H0Z` };
    case 'wedgeRectCallout': case 'wedgeRoundRectCallout': case 'wedgeEllipseCallout': case 'cloudCallout': {
      const tx = w / 2 + w * a('adj1', -20833);
      const ty = h / 2 + h * a('adj2', 62500);
      const tail = poly([[w * 0.4, h / 2], [tx, ty], [w * 0.6, h / 2]]);
      const body = name === 'wedgeRectCallout' ? `M0 0H${f(w)}V${f(h)}H0Z`
        : name === 'wedgeRoundRectCallout' ? roundRect(0, 0, w, h, ss / 6)
        : name === 'cloudCallout' ? cloud(w, h)
        : ellipse(w / 2, h / 2, w / 2, h / 2);
      return { d: name === 'cloudCallout' ? `${body}${ellipse(tx, ty, ss * 0.04, ss * 0.04)}${ellipse((tx + w / 2) / 2, (ty + h / 2) / 2 + h * 0.15, ss * 0.06, ss * 0.06)}` : `${body}${tail}` };
    }
    case 'cloud':
      return { d: cloud(w, h) };
    case 'donut': {
      const t = ss * a('adj', 25000);
      return { d: ellipse(w / 2, h / 2, w / 2, h / 2) + ellipse(w / 2, h / 2, w / 2 - t, h / 2 - t) };
    }
    case 'frame': {
      const t = ss * a('adj1', 12500);
      return { d: `M0 0H${f(w)}V${f(h)}H0ZM${f(t)} ${f(t)}V${f(h - t)}H${f(w - t)}V${f(t)}Z` };
    }
    case 'heart':
      return { d: `M${f(w / 2)} ${f(h / 4)}C${f(w / 2)} 0 0 0 0 ${f(h / 4)}C0 ${f(h / 2)} ${f(w / 2)} ${f(h * 0.75)} ${f(w / 2)} ${f(h)}C${f(w / 2)} ${f(h * 0.75)} ${f(w)} ${f(h / 2)} ${f(w)} ${f(h / 4)}C${f(w)} 0 ${f(w / 2)} 0 ${f(w / 2)} ${f(h / 4)}Z` };
    case 'can': case 'flowChartMagneticDisk': {
      const ry = Math.min(h / 4, ss * a('adj', 25000) / 2);
      return { d: `M0 ${f(ry)}A${f(w / 2)} ${f(ry)} 0 0 1 ${f(w)} ${f(ry)}V${f(h - ry)}A${f(w / 2)} ${f(ry)} 0 0 1 0 ${f(h - ry)}Z${ellipse(w / 2, ry, w / 2, ry)}` };
    }
    case 'cube': {
      const o = ss * a('adj', 25000);
      return { d: `${poly([[0, o], [w - o, o], [w - o, h], [0, h]])}${poly([[0, o], [o, 0], [w, 0], [w - o, o]])}${poly([[w - o, o], [w, 0], [w, h - o], [w - o, h]])}` };
    }
    case 'teardrop':
      return { d: `M0 ${f(h / 2)}A${f(w / 2)} ${f(h / 2)} 0 1 0 ${f(w / 2)} 0H${f(w)}V${f(h / 2)}A${f(w / 2)} ${f(h / 2)} 0 0 1 0 ${f(h / 2)}Z` };
    case 'lightningBolt':
      return { d: poly([[0.39, 0], [0.62, 0.3], [0.53, 0.33], [0.84, 0.64], [0.75, 0.68], [1, 1], [0.48, 0.73], [0.58, 0.69], [0.21, 0.43], [0.31, 0.38], [0, 0.18]].map(([x, y]) => [x * w, y * h])) };
    case 'leftBracket':
      return { d: `M${f(w)} 0Q0 0 0 ${f(h / 10)}V${f((h * 9) / 10)}Q0 ${f(h)} ${f(w)} ${f(h)}`, open: true };
    case 'rightBracket':
      return { d: `M0 0Q${f(w)} 0 ${f(w)} ${f(h / 10)}V${f((h * 9) / 10)}Q${f(w)} ${f(h)} 0 ${f(h)}`, open: true };
    case 'leftBrace':
      return { d: `M${f(w)} 0Q${f(w / 2)} 0 ${f(w / 2)} ${f(h / 8)}V${f((h * 3) / 8)}Q${f(w / 2)} ${f(h / 2)} 0 ${f(h / 2)}Q${f(w / 2)} ${f(h / 2)} ${f(w / 2)} ${f((h * 5) / 8)}V${f((h * 7) / 8)}Q${f(w / 2)} ${f(h)} ${f(w)} ${f(h)}`, open: true };
    case 'rightBrace':
      return { d: `M0 0Q${f(w / 2)} 0 ${f(w / 2)} ${f(h / 8)}V${f((h * 3) / 8)}Q${f(w / 2)} ${f(h / 2)} ${f(w)} ${f(h / 2)}Q${f(w / 2)} ${f(h / 2)} ${f(w / 2)} ${f((h * 5) / 8)}V${f((h * 7) / 8)}Q${f(w / 2)} ${f(h)} 0 ${f(h)}`, open: true };
    default:
      return { d: `M0 0H${f(w)}V${f(h)}H0Z` };
  }
}

/** Reads `<a:avLst>` adjust values (`<a:gd name="adj" fmla="val 25000"/>`). */
export function adjustments(geom: Element | null): Adj {
  const values = new Map<string, number>();
  for (const gd of children(child(geom, 'avLst'), 'gd')) {
    const m = /^val\s+(-?\d+)/.exec(gd.getAttribute('fmla') ?? '');
    if (m) values.set(gd.getAttribute('name') ?? '', Number(m[1]));
  }
  return (name, fallback) => values.get(name) ?? (name === 'adj1' ? values.get('adj') : undefined) ?? fallback;
}

/**
 * A custom shape (`<a:custGeom>`): each path has its own coordinate space (w × h) scaled to the
 * shape. Guide names in coordinates are looked up in the shape's guide list when they are plain
 * values; computed guides are not evaluated.
 */
export function customOutline(geom: Element, w: number, h: number): Outline[] {
  const guides = new Map<string, number>();
  for (const gd of [...children(child(geom, 'avLst'), 'gd'), ...children(child(geom, 'gdLst'), 'gd')]) {
    const m = /^val\s+(-?\d+)/.exec(gd.getAttribute('fmla') ?? '');
    if (m) guides.set(gd.getAttribute('name') ?? '', Number(m[1]));
  }
  const num = (v: string | null) => {
    if (v === null) return 0;
    const n = Number(v);
    return Number.isFinite(n) ? n : guides.get(v) ?? 0;
  };
  const out: Outline[] = [];
  for (const path of children(child(geom, 'pathLst'), 'path')) {
    const pw = num(path.getAttribute('w')) || w;
    const ph = num(path.getAttribute('h')) || h;
    const sx = w / pw;
    const sy = h / ph;
    const pt = (el: Element | null) => [num(el?.getAttribute('x') ?? null) * sx, num(el?.getAttribute('y') ?? null) * sy];
    let d = '';
    let cx = 0;
    let cy = 0;
    for (const cmd of path.children) {
      const pts = children(cmd, 'pt').map(pt);
      switch (cmd.localName) {
        case 'moveTo':
          [cx, cy] = pts[0];
          d += `M${f(cx)} ${f(cy)}`;
          break;
        case 'lnTo':
          [cx, cy] = pts[0];
          d += `L${f(cx)} ${f(cy)}`;
          break;
        case 'cubicBezTo':
          d += `C${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join(' ')}`;
          [cx, cy] = pts[2];
          break;
        case 'quadBezTo':
          d += `Q${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join(' ')}`;
          [cx, cy] = pts[1];
          break;
        case 'arcTo': {
          const rx = num(cmd.getAttribute('wR')) * sx;
          const ry = num(cmd.getAttribute('hR')) * sy;
          const st = (num(cmd.getAttribute('stAng')) / 60000) * (Math.PI / 180);
          const sw = (num(cmd.getAttribute('swAng')) / 60000) * (Math.PI / 180);
          // The arc starts at the current point, which lies on the ellipse at angle st.
          const ox = cx - rx * Math.cos(st);
          const oy = cy - ry * Math.sin(st);
          cx = ox + rx * Math.cos(st + sw);
          cy = oy + ry * Math.sin(st + sw);
          d += `A${f(rx)} ${f(ry)} 0 ${Math.abs(sw) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${f(cx)} ${f(cy)}`;
          break;
        }
        case 'close':
          d += 'Z';
          break;
      }
    }
    if (d) out.push({ d, open: path.getAttribute('fill') === 'none' });
  }
  return out;
}
