import { child, children, first, relAttr } from '../ooxml';
import { EMU, type Deck, type Part, type Theme } from './deck';

// DrawingML colours, fills and lines.

export interface Ctx {
  deck: Deck;
  theme: Theme;
  clrMap: Record<string, string>;
  /** The part whose relationships resolve pictures for the shapes being drawn. */
  part: Part;
  /** Colour standing in for `phClr` in a theme style (the style reference's colour). */
  phClr?: Rgba;
  slideNumber: number;
  /** Kinds of content on the slide that are not drawn (shown as a notice). */
  missing: Set<string>;
  /** Unique ids for SVG gradients and markers. */
  ids: { n: number };
}

export interface Rgba {
  r: number;
  g: number;
  b: number;
  a: number;
}

const COLOUR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);

const PRESET: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00',
  gray: '808080', grey: '808080', darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500',
  purple: '800080', navy: '000080', silver: 'C0C0C0', cyan: '00FFFF', magenta: 'FF00FF',
};

function hex(h: string): Rgba {
  const n = parseInt(h.padStart(6, '0').slice(-6), 16) || 0;
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: 1 };
}

/** The colour inside `el` (a solidFill, gs, buClr…), or `el` itself if it is a colour element. */
export function colour(el: Element | null | undefined, ctx: Ctx): Rgba | null {
  if (!el) return null;
  const c = COLOUR_TAGS.has(el.localName) ? el : children(el).find((x) => COLOUR_TAGS.has(x.localName));
  if (!c) return null;
  const val = c.getAttribute('val') ?? '';
  let base: Rgba;
  switch (c.localName) {
    case 'srgbClr':
      base = hex(val);
      break;
    case 'schemeClr':
      if (val === 'phClr') {
        if (!ctx.phClr) return null;
        base = { ...ctx.phClr };
      } else base = hex(ctx.theme.colours[ctx.clrMap[val] ?? val] ?? '000000');
      break;
    case 'sysClr':
      base = hex(c.getAttribute('lastClr') ?? (val === 'window' ? 'FFFFFF' : '000000'));
      break;
    case 'prstClr':
      base = hex(PRESET[val] ?? '000000');
      break;
    case 'scrgbClr': {
      const lin = (name: string) => toSrgb(Number(c.getAttribute(name) ?? 0) / 100000);
      base = { r: lin('r'), g: lin('g'), b: lin('b'), a: 1 };
      break;
    }
    default: {
      const [r, g, b] = hslToRgb(
        Number(c.getAttribute('hue') ?? 0) / 60000 / 360,
        Number(c.getAttribute('sat') ?? 0) / 100000,
        Number(c.getAttribute('lum') ?? 0) / 100000,
      );
      base = { r, g, b, a: 1 };
    }
  }
  return modify(base, c);
}

function toLinear(v: number): number {
  const s = v / 255;
  return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
}

function toSrgb(lin: number): number {
  const v = Math.min(1, Math.max(0, lin));
  return Math.round((v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055) * 255);
}

/** Applies DrawingML colour transforms (lumMod, tint, shade, alpha…) in document order. */
function modify(c: Rgba, el: Element): Rgba {
  let { r, g, b, a } = c;
  for (const m of el.children) {
    const v = Number(m.getAttribute('val') ?? 0) / 100000;
    switch (m.localName) {
      case 'alpha':
        a = v;
        break;
      case 'alphaMod':
        a *= v;
        break;
      case 'tint': // towards white, in linear light
        [r, g, b] = [r, g, b].map((x) => toSrgb(toLinear(x) * v + (1 - v)));
        break;
      case 'shade': // towards black, in linear light
        [r, g, b] = [r, g, b].map((x) => toSrgb(toLinear(x) * v));
        break;
      case 'lumMod':
      case 'lumOff':
      case 'satMod':
      case 'hueOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const out = m.localName === 'lumMod' ? hslToRgb(h, s, l * v)
          : m.localName === 'lumOff' ? hslToRgb(h, s, l + v)
          : m.localName === 'satMod' ? hslToRgb(h, Math.min(1, s * v), l)
          : hslToRgb((h + Number(m.getAttribute('val') ?? 0) / 60000 / 360) % 1, s, l);
        [r, g, b] = out;
        break;
      }
      case 'inv':
        [r, g, b] = [255 - r, 255 - g, 255 - b];
        break;
      case 'gray': {
        const y = Math.round(0.299 * r + 0.587 * g + 0.114 * b);
        [r, g, b] = [y, y, y];
        break;
      }
    }
  }
  return { r, g, b, a };
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const [R, G, B] = [r / 255, g / 255, b / 255];
  const max = Math.max(R, G, B);
  const min = Math.min(R, G, B);
  const l = (max + min) / 2;
  const d = max - min;
  if (!d) return [0, 0, l];
  const s = d / (1 - Math.abs(2 * l - 1));
  let h = max === R ? ((G - B) / d) % 6 : max === G ? (B - R) / d + 2 : (R - G) / d + 4;
  if (h < 0) h += 6;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  l = Math.min(1, Math.max(0, l));
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const hp = h * 6;
  const x = c * (1 - Math.abs((hp % 2) - 1));
  const [r, g, b] = hp < 1 ? [c, x, 0] : hp < 2 ? [x, c, 0] : hp < 3 ? [0, c, x] : hp < 4 ? [0, x, c] : hp < 5 ? [x, 0, c] : [c, 0, x];
  const m = l - c / 2;
  return [r, g, b].map((v) => Math.round(Math.min(1, Math.max(0, v + m)) * 255)) as [number, number, number];
}

export function css(c: Rgba | null): string {
  if (!c) return 'transparent';
  return c.a >= 1 ? `rgb(${c.r} ${c.g} ${c.b})` : `rgb(${c.r} ${c.g} ${c.b} / ${+c.a.toFixed(3)})`;
}

// ------------------------------------------------------------------ fills

export type Fill =
  | { kind: 'none' }
  | { kind: 'solid'; colour: Rgba }
  | { kind: 'gradient'; stops: { pos: number; colour: Rgba }[]; angle: number; radial: boolean }
  | { kind: 'image'; url: string | null; stretch: boolean };

const FILL_TAGS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);

/** The fill element among `el`'s children, if any. */
export function fillElement(el: Element | null | undefined): Element | null {
  return el ? children(el).find((c) => FILL_TAGS.has(c.localName)) ?? null : null;
}

export async function readFill(f: Element | null, ctx: Ctx, groupFill?: Fill): Promise<Fill | null> {
  if (!f) return null;
  switch (f.localName) {
    case 'noFill':
      return { kind: 'none' };
    case 'solidFill': {
      const c = colour(f, ctx);
      return c ? { kind: 'solid', colour: c } : { kind: 'none' };
    }
    case 'gradFill': {
      const stops = children(child(f, 'gsLst'), 'gs')
        .map((gs) => ({ pos: Number(gs.getAttribute('pos') ?? 0) / 100000, colour: colour(gs, ctx) }))
        .filter((s): s is { pos: number; colour: Rgba } => !!s.colour)
        .sort((a, b) => a.pos - b.pos);
      if (!stops.length) return { kind: 'none' };
      const lin = child(f, 'lin');
      return {
        kind: 'gradient',
        stops,
        angle: lin ? Number(lin.getAttribute('ang') ?? 0) / 60000 : 90,
        radial: !lin && !!child(f, 'path'),
      };
    }
    case 'pattFill': {
      // Hatching is drawn as its foreground colour.
      const c = colour(child(f, 'fgClr'), ctx) ?? colour(child(f, 'bgClr'), ctx);
      return c ? { kind: 'solid', colour: c } : { kind: 'none' };
    }
    case 'blipFill': {
      const target = await imageTarget(child(f, 'blip'), ctx);
      if (!target.url) ctx.missing.add('picture');
      return { kind: 'image', url: target.url, stretch: !child(f, 'tile') };
    }
    case 'grpFill':
      return groupFill ?? { kind: 'none' };
  }
  return null;
}

export async function imageTarget(blip: Element | null, ctx: Ctx): Promise<{ url: string | null }> {
  const rel = ctx.part.rels.get(relAttr(blip, 'embed'));
  if (!rel || rel.external) return { url: null };
  return { url: await ctx.deck.image(rel.target) };
}

/** A theme fill/line style for a `<a:fillRef>`/`<a:lnRef>`, with its colour as phClr. */
export function styleRef(ref: Element | null, list: Element[], bgList: Element[], ctx: Ctx): { el: Element | null; ctx: Ctx } {
  if (!ref) return { el: null, ctx };
  const idx = Number(ref.getAttribute('idx') ?? 0);
  const el = idx >= 1001 ? bgList[idx - 1001] ?? null : idx >= 1 ? list[idx - 1] ?? null : null;
  return { el, ctx: { ...ctx, phClr: colour(ref, ctx) ?? undefined } };
}

/** A fill as a CSS background (for slide backgrounds and table cells). */
export function fillCss(fill: Fill | null): string {
  if (!fill) return '';
  switch (fill.kind) {
    case 'none':
      return 'background:transparent';
    case 'solid':
      return `background:${css(fill.colour)}`;
    case 'gradient': {
      const stops = fill.stops.map((s) => `${css(s.colour)} ${+(s.pos * 100).toFixed(2)}%`).join(',');
      return fill.radial ? `background:radial-gradient(circle,${stops})` : `background:linear-gradient(${fill.angle + 90}deg,${stops})`;
    }
    case 'image':
      return fill.url ? `background:url("${fill.url}") center/${fill.stretch ? '100% 100%' : 'auto'} no-repeat` : '';
  }
}

/** An SVG paint for a fill; gradients and pictures add a definition to `defs`. */
export function fillPaint(fill: Fill | null, defs: string[], ctx: Ctx, w: number, h: number): string {
  if (!fill || fill.kind === 'none') return 'none';
  if (fill.kind === 'solid') return css(fill.colour);
  const id = `d${++ctx.ids.n}`;
  if (fill.kind === 'gradient') {
    const stops = fill.stops.map((s) => `<stop offset="${s.pos}" stop-color="${css({ ...s.colour, a: 1 })}" stop-opacity="${s.colour.a}"/>`).join('');
    if (fill.radial) defs.push(`<radialGradient id="${id}">${stops}</radialGradient>`);
    else {
      // DrawingML angles run clockwise from the +x axis; express as a vector across the box.
      const rad = (fill.angle * Math.PI) / 180;
      const dx = Math.cos(rad) / 2;
      const dy = Math.sin(rad) / 2;
      defs.push(`<linearGradient id="${id}" x1="${0.5 - dx}" y1="${0.5 - dy}" x2="${0.5 + dx}" y2="${0.5 + dy}">${stops}</linearGradient>`);
    }
    return `url(#${id})`;
  }
  if (!fill.url) return 'none';
  defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="${w}" height="${h}"><image href="${fill.url}" width="${w}" height="${h}" preserveAspectRatio="${fill.stretch ? 'none' : 'xMidYMid slice'}"/></pattern>`);
  return `url(#${id})`;
}

// ------------------------------------------------------------------ lines

export interface Line {
  width: number;
  fill: Fill | null;
  dash: string;
  head: string;
  tail: string;
}

/** Merges a theme line style (lnRef) with the shape's own `<a:ln>`, the latter winning. */
export async function readLine(own: Element | null, styleEl: Element | null, styleCtx: Ctx, ctx: Ctx): Promise<Line | null> {
  const width = (el: Element | null) => (el?.hasAttribute('w') ? Number(el.getAttribute('w')) / EMU : undefined);
  const ownFill = await readFill(fillElement(own), ctx);
  const styleFill = await readFill(fillElement(styleEl), styleCtx);
  const fill = ownFill ?? styleFill;
  if (!fill || fill.kind === 'none') return null;
  const dash = child(own, 'prstDash')?.getAttribute('val') ?? child(styleEl, 'prstDash')?.getAttribute('val') ?? 'solid';
  const end = (name: string) => child(own, name)?.getAttribute('type') ?? child(styleEl, name)?.getAttribute('type') ?? 'none';
  return {
    width: Math.max(0.5, width(own) ?? width(styleEl) ?? 0.75),
    fill,
    dash,
    head: end('headEnd'),
    tail: end('tailEnd'),
  };
}

const DASH: Record<string, number[]> = {
  dash: [4, 3], sysDash: [3, 1], dot: [1, 3], sysDot: [1, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

export function strokeAttrs(line: Line | null, defs: string[], ctx: Ctx, w: number, h: number): string {
  if (!line) return 'stroke="none"';
  let attrs = `stroke="${fillPaint(line.fill, defs, ctx, w, h)}" stroke-width="${line.width}"`;
  const dash = DASH[line.dash];
  if (dash) attrs += ` stroke-dasharray="${dash.map((d) => d * line.width).join(' ')}"`;
  const colourOf = line.fill?.kind === 'solid' ? css(line.fill.colour) : line.fill?.kind === 'gradient' ? css(line.fill.stops[0].colour) : '#000';
  for (const [end, attr] of [[line.head, 'marker-start'], [line.tail, 'marker-end']] as const) {
    if (end === 'none') continue;
    const id = `m${++ctx.ids.n}`;
    const shape = end === 'oval' ? '<circle cx="5" cy="5" r="4"/>'
      : end === 'diamond' ? '<path d="M0 5L5 0L10 5L5 10Z"/>'
      : end === 'arrow' ? '<path d="M0 0L10 5L0 10" fill="none" stroke="currentColor" stroke-width="1.5"/>'
      : end === 'stealth' ? '<path d="M0 0L10 5L0 10L3 5Z"/>'
      : '<path d="M0 0L10 5L0 10Z"/>';
    defs.push(`<marker id="${id}" viewBox="0 0 10 10" refX="${end === 'oval' || end === 'diamond' ? 5 : 9}" refY="5" markerWidth="3" markerHeight="3" orient="auto-start-reverse" style="color:${colourOf}" fill="${colourOf}">${shape}</marker>`);
    attrs += ` ${attr}="url(#${id})"`;
  }
  return attrs;
}

/** CSS for a table border edge `<a:lnL>` etc. */
export async function borderCss(ln: Element | null, ctx: Ctx): Promise<string | null> {
  if (!ln) return null;
  const fill = await readFill(fillElement(ln), ctx);
  if (!fill || fill.kind === 'none') return 'none';
  const w = Number(ln.getAttribute('w') ?? 12700) / EMU;
  const c = fill.kind === 'solid' ? css(fill.colour) : fill.kind === 'gradient' ? css(fill.stops[0].colour) : '#000';
  const dash = child(ln, 'prstDash')?.getAttribute('val') ?? 'solid';
  return `${Math.max(0.5, w)}px ${dash === 'solid' ? 'solid' : dash.toLowerCase().includes('dot') ? 'dotted' : 'dashed'} ${c}`;
}

/** An outer shadow from `<a:effectLst>`, as a CSS drop-shadow filter. */
export function shadowCss(effects: Element | null, ctx: Ctx): string {
  const s = first(effects, 'outerShdw');
  if (!s) return '';
  const c = colour(s, ctx);
  if (!c) return '';
  const dist = Number(s.getAttribute('dist') ?? 0) / EMU;
  const dir = ((Number(s.getAttribute('dir') ?? 0) / 60000) * Math.PI) / 180;
  const blur = Number(s.getAttribute('blurRad') ?? 0) / EMU;
  return `drop-shadow(${(Math.cos(dir) * dist).toFixed(1)}px ${(Math.sin(dir) * dist).toFixed(1)}px ${(blur / 2).toFixed(1)}px ${css(c)})`;
}
