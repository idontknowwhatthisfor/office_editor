import { child, children } from '../ooxml';
import { EMU } from './deck';
import { colour, css, type Ctx, type Rgba } from './drawing';

// Text bodies. A paragraph's look is merged from several sources, lowest first: the
// presentation's default text style, the master's title/body/other style, the master's and the
// layout's placeholder, the shape's own list style, then the paragraph and run properties.

export interface RunStyle {
  /** Points. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: '' | 'single' | 'double';
  strike?: boolean;
  colour?: Rgba | null;
  font?: string;
  /** Superscript (> 0) / subscript (< 0), as a fraction of the size. */
  baseline?: number;
  caps?: '' | 'all' | 'small';
  highlight?: Rgba | null;
  link?: boolean;
}

export interface ParaStyle {
  algn?: string;
  marL?: number;
  indent?: number;
  /** Line spacing as a multiple of single spacing, or exact in px. */
  lineMul?: number;
  linePx?: number;
  beforePx?: number;
  beforeMul?: number;
  afterPx?: number;
  afterMul?: number;
  bullet?: 'none' | 'char' | 'auto';
  buChar?: string;
  buFont?: string;
  buColour?: Rgba | null;
  /** Bullet size as a fraction of the text size, or in points. */
  buSize?: number;
  buSizePt?: number;
  buAuto?: string;
  buStart?: number;
  run: RunStyle;
}

const pt = (hundredths: string | null) => (hundredths === null ? undefined : Number(hundredths) / 100);
const on = (v: string | null) => (v === null ? undefined : v === '1' || v === 'true');

function defined<T extends object>(o: T): Partial<T> {
  const out: Partial<T> = {};
  for (const k in o) if (o[k] !== undefined) out[k] = o[k];
  return out;
}

export function mergePara(a: ParaStyle, b: ParaStyle): ParaStyle {
  return { ...a, ...defined({ ...b, run: undefined }), run: { ...a.run, ...defined(b.run) } };
}

export function parseRPr(el: Element | null | undefined, ctx: Ctx): RunStyle {
  if (!el) return {};
  const u = el.getAttribute('u');
  const strike = el.getAttribute('strike');
  const cap = el.getAttribute('cap');
  const base = el.getAttribute('baseline');
  const fill = children(el).find((c) => c.localName === 'solidFill' || c.localName === 'noFill' || c.localName === 'gradFill');
  let col: Rgba | null | undefined;
  if (fill?.localName === 'noFill') col = { r: 0, g: 0, b: 0, a: 0 };
  else if (fill?.localName === 'solidFill') col = colour(fill, ctx);
  else if (fill) col = colour(child(child(fill, 'gsLst'), 'gs'), ctx);
  const hl = child(el, 'highlight');
  return defined({
    size: pt(el.getAttribute('sz')),
    bold: on(el.getAttribute('b')),
    italic: on(el.getAttribute('i')),
    underline: u === null ? undefined : u === 'none' ? '' : u.startsWith('dbl') ? 'double' : 'single',
    strike: strike === null ? undefined : strike !== 'noStrike',
    colour: col,
    font: child(el, 'latin')?.getAttribute('typeface') ?? undefined,
    baseline: base === null ? undefined : Number(base) / 100000,
    caps: cap === null ? undefined : cap === 'none' ? '' : cap,
    highlight: hl ? colour(hl, ctx) : undefined,
    link: child(el, 'hlinkClick') ? true : undefined,
  }) as RunStyle;
}

function spacing(el: Element | null): { px?: number; mul?: number } {
  const pct = child(el, 'spcPct');
  if (pct) return { mul: Number(pct.getAttribute('val') ?? 100000) / 100000 };
  const pts = child(el, 'spcPts');
  if (pts) return { px: ((Number(pts.getAttribute('val') ?? 0) / 100) * 96) / 72 };
  return {};
}

export function parsePPr(el: Element | null | undefined, ctx: Ctx): ParaStyle {
  if (!el) return { run: {} };
  const emu = (name: string) => (el.hasAttribute(name) ? Number(el.getAttribute(name)) / EMU : undefined);
  const line = spacing(child(el, 'lnSpc'));
  const before = spacing(child(el, 'spcBef'));
  const after = spacing(child(el, 'spcAft'));
  const style: ParaStyle = {
    algn: el.getAttribute('algn') ?? undefined,
    marL: emu('marL'),
    indent: emu('indent'),
    lineMul: line.mul,
    linePx: line.px,
    beforeMul: before.mul,
    beforePx: before.px,
    afterMul: after.mul,
    afterPx: after.px,
    run: parseRPr(child(el, 'defRPr'), ctx),
  };
  if (line.mul !== undefined) style.linePx = undefined;
  if (line.px !== undefined) style.lineMul = undefined;
  if (child(el, 'buNone')) style.bullet = 'none';
  const buChar = child(el, 'buChar');
  if (buChar) {
    style.bullet = 'char';
    style.buChar = buChar.getAttribute('char') ?? '•';
  }
  const auto = child(el, 'buAutoNum');
  if (auto) {
    style.bullet = 'auto';
    style.buAuto = auto.getAttribute('type') ?? 'arabicPeriod';
    style.buStart = Number(auto.getAttribute('startAt') ?? 1);
  }
  const buFont = child(el, 'buFont');
  if (buFont) style.buFont = buFont.getAttribute('typeface') ?? undefined;
  if (child(el, 'buFontTx')) style.buFont = '';
  const buClr = child(el, 'buClr');
  if (buClr) style.buColour = colour(buClr, ctx);
  if (child(el, 'buClrTx')) style.buColour = null;
  const buSzPct = child(el, 'buSzPct');
  if (buSzPct) style.buSize = Number(buSzPct.getAttribute('val') ?? 100000) / 100000;
  const buSzPts = child(el, 'buSzPts');
  if (buSzPts) style.buSizePt = Number(buSzPts.getAttribute('val') ?? 0) / 100;
  if (child(el, 'buSzTx')) {
    style.buSize = 1;
    style.buSizePt = undefined;
  }
  return defined(style) as ParaStyle;
}

/**
 * Per-level paragraph styles (9 levels) merged from list-style elements (`lstStyle`,
 * `titleStyle`, `defaultTextStyle`…), lowest priority first. `base` goes under all of them.
 */
/** Each level's style with a list-style element's level properties merged over it. */
export function overlay(levels: ParaStyle[], src: Element | null | undefined, ctx: Ctx): ParaStyle[] {
  if (!src) return levels;
  return levels.map((style, i) =>
    mergePara(mergePara(style, parsePPr(child(src, 'defPPr'), ctx)), parsePPr(child(src, `lvl${i + 1}pPr`), ctx)));
}

export function levelStyles(sources: (Element | null | undefined)[], ctx: Ctx, base: ParaStyle = { run: {} }): ParaStyle[] {
  const levels: ParaStyle[] = [];
  for (let n = 1; n <= 9; n++) {
    let style = base;
    for (const src of sources) {
      if (!src) continue;
      style = mergePara(style, parsePPr(child(src, 'defPPr'), ctx));
      style = mergePara(style, parsePPr(child(src, `lvl${n}pPr`), ctx));
    }
    levels.push(style);
  }
  return levels;
}

// ------------------------------------------------------------------ body properties

export interface BodyProps {
  insets: [number, number, number, number];
  anchor: string;
  anchorCtr: boolean;
  wrap: boolean;
  vert: string;
  fontScale: number;
  lineReduction: number;
}

/** Body properties from `<a:bodyPr>` elements, lowest priority first. */
export function bodyProps(chain: (Element | null | undefined)[]): BodyProps {
  const attr = (name: string) => {
    let v: string | null = null;
    for (const el of chain) if (el?.hasAttribute(name)) v = el.getAttribute(name);
    return v;
  };
  const ins = (name: string, d: number) => Number(attr(name) ?? d) / EMU;
  let fontScale = 1;
  let lineReduction = 0;
  for (const el of chain) {
    const fit = el && children(el).find((c) => /Autofit$|^spAutoFit$/.test(c.localName));
    if (!fit) continue;
    fontScale = fit.localName === 'normAutofit' ? Number(fit.getAttribute('fontScale') ?? 100000) / 100000 : 1;
    lineReduction = fit.localName === 'normAutofit' ? Number(fit.getAttribute('lnSpcReduction') ?? 0) / 100000 : 0;
  }
  return {
    insets: [ins('tIns', 45720), ins('rIns', 91440), ins('bIns', 45720), ins('lIns', 91440)],
    anchor: attr('anchor') ?? 't',
    anchorCtr: attr('anchorCtr') === '1',
    wrap: attr('wrap') !== 'none',
    vert: attr('vert') ?? 'horz',
    fontScale,
    lineReduction,
  };
}

// ------------------------------------------------------------------ rendering

const SERIF = /times|georgia|garamond|cambria|book antiqua|palatino|serif|minion|baskerville/i;
const MONO = /courier|consolas|mono|menlo/i;

export function fontFamily(name: string | undefined, ctx: Ctx): string {
  let face = name ?? '+mn-lt';
  if (face.startsWith('+mj')) face = ctx.theme.majorFont;
  else if (face.startsWith('+mn')) face = ctx.theme.minorFont;
  const clean = face.replace(/['"\\;{}]/g, '');
  const generic = SERIF.test(clean) ? 'serif' : MONO.test(clean) ? 'monospace' : 'sans-serif';
  // Calibri Light has no stand-in of its own; Calibri's (Carlito) is the closest.
  const extra = /^calibri light$/i.test(clean) ? ", 'Calibri'" : '';
  return clean ? `'${clean}'${extra}, ${generic}` : generic;
}

/** Characters of symbol fonts used for bullets, as the Unicode characters they look like. */
const SYMBOL_BULLETS: Record<string, string> = {
  '§': '▪', 'Ø': '➢', 'ü': '✓', 'q': '❑', 'n': '■', 'l': '●', 'v': '❖', 'Ü': '➢', 'ð': '⇨',
  'Þ': '➔', 'à': '➔', 'o': '□', 'p': '◆', 'u': '◆', '·': '•',
  // The same characters as private-use code points (Symbol-encoded fonts in some files).
  '\uF0A7': '▪', '\uF0D8': '➢', '\uF0FC': '✓', '\uF06E': '■', '\uF06C': '●', '\uF076': '❖', '\uF0B7': '•',
};

function bulletChar(ch: string, font: string | undefined): string {
  if ((font && /wingdings|symbol|webdings/i.test(font)) || /[\uF000-\uF0FF]/.test(ch)) return SYMBOL_BULLETS[ch] ?? '•';
  return ch;
}

function roman(n: number): string {
  const table: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let s = '';
  for (const [v, r] of table) while (n >= v) { s += r; n -= v; }
  return s;
}

function autoNumber(type: string, n: number): string {
  const alpha = (k: number) => {
    let s = '';
    for (let x = k; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(97 + ((x - 1) % 26)) + s;
    return s;
  };
  let core = type.startsWith('alpha') ? alpha(n) : type.startsWith('roman') ? roman(n) : String(n);
  if (/Uc/.test(type)) core = core.toUpperCase();
  if (type.endsWith('ParenBoth')) return `(${core})`;
  if (type.endsWith('ParenR')) return `${core})`;
  if (type.endsWith('Plain')) return core;
  return `${core}.`;
}

function escape(s: string): string {
  return s.replace(/[&<>"]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;'));
}

const px = (points: number) => (points * 96) / 72;

function runCss(s: RunStyle, ctx: Ctx, scale: number): string {
  let out = `font-size:${+(px(s.size ?? 18) * scale).toFixed(2)}px;font-family:${fontFamily(s.font, ctx)}`;
  if (s.bold) out += ';font-weight:700';
  if (s.italic) out += ';font-style:italic';
  const lines = [s.underline || s.link ? 'underline' : '', s.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (lines) out += `;text-decoration:${lines}${s.underline === 'double' ? ' double' : ''}`;
  out += `;color:${s.link ? `#${ctx.theme.colours.hlink}` : s.colour ? css(s.colour) : '#000'}`;
  if (s.baseline) out += `;vertical-align:${+(s.baseline * 100).toFixed(1)}%;font-size:${+(px(s.size ?? 18) * scale * 0.7).toFixed(2)}px`;
  if (s.caps === 'all') out += ';text-transform:uppercase';
  else if (s.caps === 'small') out += ';font-variant:small-caps';
  if (s.highlight) out += `;background:${css(s.highlight)}`;
  return out;
}

/**
 * HTML for the paragraphs of a text body, laid out in a box of the shape's size. Text that
 * overflows the box stays visible, as in PowerPoint.
 */
export function renderText(
  txBody: Element, levels: ParaStyle[], body: BodyProps, ctx: Ctx, box: { w: number; h: number },
): string {
  const scale = body.fontScale;
  const counters: number[] = new Array(9).fill(0);
  let html = '';
  for (const p of children(txBody, 'p')) {
    const pPr = child(p, 'pPr');
    const lvl = Math.min(8, Math.max(0, Number(pPr?.getAttribute('lvl') ?? 0)));
    const style = mergePara(levels[lvl], parsePPr(pPr, ctx));
    let runs = '';
    let firstRun: RunStyle | null = null;
    for (const r of p.children) {
      if (r.localName === 'r' || r.localName === 'fld') {
        const rs = { ...style.run, ...parseRPr(child(r, 'rPr'), ctx) };
        let text = child(r, 't')?.textContent ?? '';
        if (r.localName === 'fld' && r.getAttribute('type') === 'slidenum') text = String(ctx.slideNumber);
        if (!text) continue;
        firstRun ??= rs;
        runs += `<span style="${runCss(rs, ctx, scale)}">${escape(text)}</span>`;
      } else if (r.localName === 'br') {
        const rs = { ...style.run, ...parseRPr(child(r, 'rPr'), ctx) };
        firstRun ??= rs;
        runs += '<br>';
      }
    }
    const end = { ...style.run, ...parseRPr(child(p, 'endParaRPr'), ctx) };
    const lead = firstRun ?? end;
    const size = px(lead.size ?? 18) * scale;

    // Numbering continues while paragraphs stay at this level; a shallower level restarts deeper ones.
    for (let k = lvl + 1; k < 9; k++) counters[k] = 0;
    let bullet = '';
    if (firstRun && style.bullet === 'char' && style.buChar) {
      bullet = escape(bulletChar(style.buChar, style.buFont));
    } else if (firstRun && style.bullet === 'auto') {
      counters[lvl] = counters[lvl] ? counters[lvl] + 1 : style.buStart ?? 1;
      bullet = autoNumber(style.buAuto ?? 'arabicPeriod', counters[lvl]);
    }
    if (style.bullet !== 'auto') counters[lvl] = 0;

    let pcss = `text-align:${ALIGN[style.algn ?? 'l'] ?? 'left'};font-size:${+size.toFixed(2)}px`;
    const marL = style.marL ?? 0;
    const indent = style.indent ?? 0;
    if (marL) pcss += `;padding-left:${marL}px`;
    if (indent) pcss += `;text-indent:${indent}px`;
    const mul = style.lineMul !== undefined ? Math.max(0.1, style.lineMul - body.lineReduction) : style.linePx === undefined ? 1 - body.lineReduction : undefined;
    pcss += mul !== undefined ? `;line-height:${+(mul * 1.2).toFixed(3)}` : `;line-height:${+(style.linePx! * scale).toFixed(2)}px`;
    const before = style.beforePx ?? (style.beforeMul !== undefined ? style.beforeMul * size * 1.2 : 0);
    const after = style.afterPx ?? (style.afterMul !== undefined ? style.afterMul * size * 1.2 : 0);
    if (before) pcss += `;margin-top:${+(before * (style.beforePx !== undefined ? scale : 1)).toFixed(2)}px`;
    if (after) pcss += `;margin-bottom:${+(after * (style.afterPx !== undefined ? scale : 1)).toFixed(2)}px`;

    if (bullet) {
      const bsize = style.buSizePt ? px(style.buSizePt) * scale : size * (style.buSize ?? 1);
      const bcol = style.buColour ?? lead.colour;
      const family = style.buFont && !/wingdings|symbol|webdings/i.test(style.buFont) ? fontFamily(style.buFont, ctx) : fontFamily(lead.font, ctx);
      const width = indent < 0 ? -indent : 0;
      bullet = `<span class="bu" style="min-width:${width}px;font-size:${+bsize.toFixed(2)}px;font-family:${family};color:${bcol ? css(bcol) : '#000'}">${bullet}</span>`;
    }
    const content = runs || `<span style="${runCss(end, ctx, scale)}">&#8203;</span>`;
    html += `<p style="${pcss}">${bullet}${content}</p>`;
  }

  const [t, r, b, l] = body.insets;
  const justify = body.anchor === 'ctr' ? 'center' : body.anchor === 'b' ? 'flex-end' : 'flex-start';
  let style = `width:${box.w}px;height:${box.h}px;padding:${t}px ${r}px ${b}px ${l}px;justify-content:${justify}`;
  if (!body.wrap) style += ';white-space:pre';
  if (body.anchorCtr) style += ';align-items:center';
  if (body.vert === 'vert' || body.vert === 'eaVert' || body.vert === 'vert270') {
    // Rotate a box of swapped size so the text runs top to bottom (or bottom to top).
    const deg = body.vert === 'vert270' ? -90 : 90;
    style = `width:${box.h}px;height:${box.w}px;padding:${t}px ${r}px ${b}px ${l}px;justify-content:${justify};` +
      `transform:translate(${(box.w - box.h) / 2}px,${(box.h - box.w) / 2}px) rotate(${deg}deg)`;
  }
  return `<div class="tx" style="${style}">${html}</div>`;
}

const ALIGN: Record<string, string> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify', justLow: 'justify', thaiDist: 'justify' };
