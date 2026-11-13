import { child, children, first, relAttr, relOfType } from '../ooxml';
import { EMU, type Deck, type Layout, type Master, type Part, type SlidePart } from './deck';
import {
  fillCss, fillElement, fillPaint, imageTarget, readFill, readLine, shadowCss, strokeAttrs, styleRef,
  type Ctx, type Fill,
} from './drawing';
import { adjustments, customOutline, presetOutline, type Outline } from './geometry';
import { renderChart } from '../chart';
import { renderTable } from './table';
import { bodyProps, levelStyles, mergePara, overlay, renderText, type ParaStyle, type TextLayout } from './text';

// Draws one slide as HTML: a box of the slide's size in CSS px (1 px = 9525 EMU) holding the
// master's and layout's own shapes, then the slide's. Shapes are absolutely positioned boxes
// with an SVG outline and an HTML text layer.

export interface RenderedSlide {
  el: HTMLElement;
  /** Kinds of content not drawn: 'chart', 'diagram', 'object', 'picture', 'media'. */
  missing: Set<string>;
  /** Edit mode: the text of each editable shape, by its number in `editableShapes()`. */
  texts: Map<number, TextLayout>;
}

const PML = 'http://schemas.openxmlformats.org/presentationml/2006/main';
/** Placeholders that hold pictures, tables, charts… rather than text. */
const NON_TEXT_PH = new Set(['pic', 'chart', 'tbl', 'dgm', 'media', 'clipArt', 'sldImg']);

const SHAPES = new Set(['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame', 'contentPart', 'AlternateContent']);

/** The drawing elements directly in a shape tree, in order (edit mode's `data-el` numbers). */
export function topShapes(spTree: Element | null): Element[] {
  return children(spTree).filter((c) => SHAPES.has(c.localName));
}

/**
 * The slide's shapes whose text can be edited, in document order (their numbers are what edit
 * mode writes into `data-sp`). Shapes inside alternate content are left out: changing one
 * choice would leave the other stale.
 */
export function editableShapes(slideRoot: Element): Element[] {
  const out: Element[] = [];
  for (const sp of slideRoot.getElementsByTagNameNS(PML, 'sp')) {
    let inAlt = false;
    for (let a = sp.parentElement; a && !inAlt; a = a.parentElement) inAlt = a.localName === 'AlternateContent';
    if (!inAlt) out.push(sp);
  }
  return out;
}

/** Maps a shape tree's coordinates (EMU) to px in the element the shapes are drawn into. */
interface Space {
  ox: number;
  oy: number;
  kx: number;
  ky: number;
}

const ROOT: Space = { ox: 0, oy: 0, kx: 1 / EMU, ky: 1 / EMU };

interface Inherited {
  slide: SlidePart;
  layout: Layout;
  master: Master;
}

export async function renderSlide(deck: Deck, index: number, options: { edit?: boolean } = {}): Promise<RenderedSlide> {
  const slide = await deck.slide(index);
  const { layout } = slide;
  const { master } = layout;
  const missing = new Set<string>();
  const texts = new Map<number, TextLayout>();
  const base: Ctx = {
    deck, theme: master.theme, clrMap: slide.clrMap, part: slide, slideNumber: index + 1, missing, ids: { n: 0 },
  };
  if (options.edit) base.edit = { shapes: new Map(editableShapes(slide.root).map((sp, n) => [sp, n])), texts };
  const at = (part: Part): Ctx => ({ ...base, part, ids: base.ids });
  const inh: Inherited = { slide, layout, master };

  const el = document.createElement('div');
  el.className = 'slide';
  el.style.width = `${deck.width}px`;
  el.style.height = `${deck.height}px`;

  // Background: the slide's, else the layout's, else the master's.
  for (const part of [slide, layout, master] as Part[]) {
    const bg = child(child(part.root, 'cSld'), 'bg');
    if (!bg) continue;
    const ctx = at(part);
    const bgPr = child(bg, 'bgPr');
    let fill: Fill | null = null;
    if (bgPr) fill = await readFill(fillElement(bgPr), ctx);
    else {
      const ref = styleRef(child(bg, 'bgRef'), ctx.theme.fillStyles, ctx.theme.bgFillStyles, ctx);
      fill = await readFill(fillElement(ref.el ? wrap(ref.el) : null), ref.ctx);
    }
    if (fill) el.setAttribute('style', `${el.getAttribute('style')};${fillCss(fill)}`);
    break;
  }

  let html = '';
  const showMaster = (root: Element) => root.getAttribute('showMasterSp') !== '0';
  if (showMaster(slide.root) && showMaster(layout.root)) html += await renderTree(tree(master), ROOT, at(master), inh, false);
  if (showMaster(slide.root)) html += await renderTree(tree(layout), ROOT, at(layout), inh, false);
  if (options.edit) {
    // Edit mode: each of the slide's own top-level elements is marked with its number in
    // topShapes(), so it can be selected, moved and resized.
    for (const [n, c] of topShapes(tree(slide)).entries()) {
      const one = await renderElement(c, ROOT, at(slide), inh, true);
      html += c.localName === 'AlternateContent' ? one : one.replace(/^<(\w+)/, (_, tag: string) => `<${tag} data-el="${n}"`);
    }
  } else {
    html += await renderTree(tree(slide), ROOT, at(slide), inh, true);
  }
  el.insertAdjacentHTML('beforeend', html);
  return { el, missing, texts };
}

/** A fill style element wrapped so fillElement() finds it among children. */
function wrap(el: Element): Element {
  const holder = el.ownerDocument.createElementNS(el.namespaceURI, 'holder');
  holder.append(el.cloneNode(true));
  return holder;
}

function tree(part: Part): Element | null {
  return child(child(part.root, 'cSld'), 'spTree');
}

async function renderTree(root: Element | null, space: Space, ctx: Ctx, inh: Inherited, isSlide: boolean, groupFill?: Fill): Promise<string> {
  let html = '';
  for (const el of children(root)) html += await renderElement(el, space, ctx, inh, isSlide, groupFill);
  return html;
}

async function renderElement(el: Element, space: Space, ctx: Ctx, inh: Inherited, isSlide: boolean, groupFill?: Fill): Promise<string> {
  switch (el.localName) {
    case 'sp':
    case 'cxnSp':
      return renderShape(el, space, ctx, inh, isSlide, groupFill);
    case 'pic':
      return renderPicture(el, space, ctx, inh, isSlide);
    case 'grpSp':
      return renderGroup(el, space, ctx, inh, isSlide, groupFill);
    case 'graphicFrame':
      return renderFrame(el, space, ctx, inh, isSlide);
    case 'AlternateContent': {
      // Prefer the fallback: it holds content every reader understands (often a picture). But
      // newer charts (chartex) are drawn from their data rather than the fallback's picture.
      const choice = child(el, 'Choice');
      const chartEx = first(choice, 'graphicData')?.getAttribute('uri')?.includes('chartex');
      const pick = chartEx ? choice : child(el, 'Fallback') ?? choice;
      let html = '';
      for (const c of children(pick)) html += await renderElement(c, space, ctx, inh, isSlide, groupFill);
      return html;
    }
  }
  return '';
}

// ------------------------------------------------------------------ placeholders

interface Placeholder {
  type: string;
  idx: string | null;
}

function placeholderOf(el: Element): Placeholder | null {
  const nv = children(el).find((c) => c.localName.startsWith('nv'));
  const ph = child(child(nv, 'nvPr'), 'ph');
  if (!ph) return null;
  return { type: ph.getAttribute('type') ?? 'body', idx: ph.getAttribute('idx') };
}

const MASTER_TYPE: Record<string, string> = { ctrTitle: 'title', subTitle: 'body', obj: 'body', body: 'body' };

function findPlaceholder(part: Part, ph: Placeholder, byIdx: boolean): Element | null {
  const shapes = children(tree(part)).filter((c) => c.localName === 'sp' || c.localName === 'pic');
  const typed = shapes.map((s) => ({ s, p: placeholderOf(s) })).filter((x) => x.p);
  if (byIdx && ph.idx !== null) {
    const hit = typed.find((x) => x.p!.idx === ph.idx);
    if (hit) return hit.s;
  }
  const want = byIdx ? ph.type : MASTER_TYPE[ph.type] ?? ph.type;
  const same = (t: string) => (byIdx ? t === want || (MASTER_TYPE[t] ?? t) === (MASTER_TYPE[want] ?? want) : (MASTER_TYPE[t] ?? t) === want);
  return typed.find((x) => same(x.p!.type))?.s ?? null;
}

/** The layout's and master's placeholder a slide placeholder inherits from. */
function inheritance(el: Element, inh: Inherited, isSlide: boolean): { layoutPh: Element | null; masterPh: Element | null; ph: Placeholder | null } {
  const ph = placeholderOf(el);
  if (!ph || !isSlide) return { layoutPh: null, masterPh: null, ph };
  const layoutPh = findPlaceholder(inh.layout, ph, true);
  const masterPh = findPlaceholder(inh.master, layoutPh ? placeholderOf(layoutPh) ?? ph : ph, false);
  return { layoutPh, masterPh, ph };
}

// ------------------------------------------------------------------ geometry and boxes

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
  rot: number;
  flipH: boolean;
  flipV: boolean;
}

function xfrmOf(spPr: Element | null): Element | null {
  return child(spPr, 'xfrm');
}

function boxOf(xfrm: Element | null, space: Space): Box | null {
  const off = child(xfrm, 'off');
  const ext = child(xfrm, 'ext');
  if (!off || !ext) return null;
  return {
    x: (Number(off.getAttribute('x') ?? 0) - space.ox) * space.kx,
    y: (Number(off.getAttribute('y') ?? 0) - space.oy) * space.ky,
    w: Number(ext.getAttribute('cx') ?? 0) * space.kx,
    h: Number(ext.getAttribute('cy') ?? 0) * space.ky,
    rot: Number(xfrm!.getAttribute('rot') ?? 0) / 60000,
    flipH: xfrm!.getAttribute('flipH') === '1',
    flipV: xfrm!.getAttribute('flipV') === '1',
  };
}

const n2 = (v: number) => +v.toFixed(2);

function boxStyle(box: Box, extra = ''): string {
  let style = `left:${n2(box.x)}px;top:${n2(box.y)}px;width:${n2(box.w)}px;height:${n2(box.h)}px`;
  if (box.rot) style += `;transform:rotate(${n2(box.rot)}deg)`;
  return style + extra;
}

function outlines(spPr: Element | null, box: Box): Outline[] {
  const cust = child(spPr, 'custGeom');
  if (cust) return customOutline(cust, box.w, box.h);
  const prst = child(spPr, 'prstGeom');
  return [presetOutline(prst?.getAttribute('prst') ?? 'rect', box.w, box.h, adjustments(prst))];
}

// ------------------------------------------------------------------ shapes

async function renderShape(el: Element, space: Space, ctx: Ctx, inh: Inherited, isSlide: boolean, groupFill?: Fill): Promise<string> {
  const { layoutPh, masterPh, ph } = inheritance(el, inh, isSlide);
  // Placeholders on layouts and masters are prompts for the slide, never drawn themselves.
  if (ph && !isSlide) return '';
  const spPr = child(el, 'spPr');
  const chain = [masterPh, layoutPh, el];
  const xfrm = xfrmOf(spPr) ?? xfrmOf(child(layoutPh, 'spPr')) ?? xfrmOf(child(masterPh, 'spPr'));
  const box = boxOf(xfrm, space);
  if (!box) return '';

  // Fill and line: the shape's own, else the placeholder's it inherits, else its theme style.
  const styleEl = child(el, 'style') ?? child(layoutPh, 'style') ?? child(masterPh, 'style');
  let fill: Fill | null = null;
  for (const src of [el, layoutPh, masterPh]) {
    fill = await readFill(fillElement(child(src, 'spPr')), ctx, groupFill);
    if (fill) break;
  }
  if (!fill && styleEl) {
    const ref = styleRef(child(styleEl, 'fillRef'), ctx.theme.fillStyles, ctx.theme.bgFillStyles, ctx);
    if (ref.el) fill = await readFill(fillElement(wrap(ref.el)), ref.ctx);
  }
  const ownLn = child(spPr, 'ln') ?? child(child(layoutPh, 'spPr'), 'ln') ?? child(child(masterPh, 'spPr'), 'ln');
  const lnRef = styleRef(child(styleEl, 'lnRef'), ctx.theme.lineStyles, [], ctx);
  const line = await readLine(ownLn, lnRef.el, lnRef.ctx, ctx);
  const effects = child(spPr, 'effectLst') ?? (() => {
    const ref = styleRef(child(styleEl, 'effectRef'), ctx.theme.effectStyles, [], ctx);
    return ref.el ? child(ref.el, 'effectLst') : null;
  })();
  const shadow = shadowCss(effects, ctx);

  let svg = '';
  if ((fill && fill.kind !== 'none') || line) {
    const defs: string[] = [];
    const paint = fillPaint(fill, defs, ctx, box.w, box.h);
    const stroke = strokeAttrs(line, defs, ctx, box.w, box.h);
    const paths = outlines(spPr ?? child(layoutPh, 'spPr'), box)
      .map((o) => `<path d="${o.d}" fill="${o.open ? 'none' : paint}" fill-rule="evenodd" ${stroke}/>`)
      .join('');
    const flip = box.flipH || box.flipV ? ` style="transform:scale(${box.flipH ? -1 : 1},${box.flipV ? -1 : 1})"` : '';
    svg = `<svg width="${n2(Math.max(box.w, 1))}" height="${n2(Math.max(box.h, 1))}"${flip}>${defs.length ? `<defs>${defs.join('')}</defs>` : ''}${paths}</svg>`;
  }

  let text = '';
  const txBody = child(el, 'txBody');
  // Edit mode: a slide's shape with a text body, or a text placeholder (even an empty one).
  const editNo = isSlide && el.localName === 'sp' ? ctx.edit?.shapes.get(el) : undefined;
  const editable = editNo !== undefined && (!!txBody || (!!ph && !NON_TEXT_PH.has(ph.type)));
  if ((txBody && hasText(txBody)) || editable) {
    let levels: ParaStyle[];
    if (ph) {
      const master = inh.master;
      const typeStyle = ph.type === 'title' || ph.type === 'ctrTitle' ? master.titleStyle
        : ['body', 'subTitle', 'obj'].includes(ph.type) || (ph.idx !== null && !['dt', 'ftr', 'sldNum', 'hdr'].includes(ph.type)) ? master.bodyStyle
        : master.otherStyle;
      levels = levelStyles([typeStyle, lstStyle(masterPh), lstStyle(layoutPh), child(txBody, 'lstStyle')], ctx);
    } else {
      levels = levelStyles([ctx.deck.defaultTextStyle], ctx);
      const fontRef = child(styleEl, 'fontRef');
      if (fontRef) {
        const ref = styleRef(fontRef, [], [], ctx);
        const run: ParaStyle['run'] = { font: fontRef.getAttribute('idx') === 'major' ? '+mj-lt' : '+mn-lt' };
        if (ref.ctx.phClr) run.colour = ref.ctx.phClr;
        levels = levels.map((l) => mergePara(l, { run }));
      }
      levels = overlay(levels, child(txBody, 'lstStyle'), ctx);
    }
    const body = bodyProps(chain.map((s) => child(child(s, 'txBody'), 'bodyPr')));
    // SmartArt drawings give the text its own rectangle inside the shape.
    const txBox = boxOf(child(el, 'txXfrm'), space);
    if (editable) {
      const layout: TextLayout = {
        body, w: box.w, h: box.h, ky: space.ky, inherited: bodyProps(chain.slice(0, 2).map((s) => child(child(s, 'txBody'), 'bodyPr'))).autofit,
        // The text body is looked up again each time: the editor may have added or changed it.
        html: (over = {}) => renderText(
          child(el, 'txBody') ?? emptyBody(), levels, { ...body, ...over }, ctx, { w: box.w, h: over.h ?? box.h }, true),
      };
      ctx.edit!.texts.set(editNo!, layout);
      text = layout.html();
    } else {
      const html = renderText(txBody!, levels, body, ctx, { w: (txBox ?? box).w, h: (txBox ?? box).h });
      text = txBox ? `<div class="sp" style="${boxStyle({ ...txBox, x: txBox.x - box.x, y: txBox.y - box.y, rot: 0 })}">${html}</div>` : html;
    }
  }
  if (!svg && !text) return '';
  const mark = editable ? ` data-sp="${editNo}"` : '';
  return `<div class="sp"${mark} style="${boxStyle(box, shadow ? `;filter:${shadow}` : '')}">${svg}${text}</div>`;
}

let empty: Element | null = null;

/** A text body with one empty paragraph, drawn for a placeholder that has none yet. */
function emptyBody(): Element {
  empty ??= new DOMParser().parseFromString(
    '<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>',
    'application/xml').documentElement;
  return empty;
}

function lstStyle(sp: Element | null): Element | null {
  return child(child(sp, 'txBody'), 'lstStyle');
}

function hasText(txBody: Element): boolean {
  for (const p of children(txBody, 'p')) {
    for (const r of children(p)) if ((r.localName === 'r' || r.localName === 'fld') && child(r, 't')?.textContent) return true;
  }
  return false;
}

// ------------------------------------------------------------------ pictures

async function renderPicture(el: Element, space: Space, ctx: Ctx, inh: Inherited, isSlide: boolean): Promise<string> {
  const { layoutPh, masterPh, ph } = inheritance(el, inh, isSlide);
  if (ph && !isSlide) return '';
  const spPr = child(el, 'spPr');
  const box = boxOf(xfrmOf(spPr) ?? xfrmOf(child(layoutPh, 'spPr')) ?? xfrmOf(child(masterPh, 'spPr')), space);
  if (!box) return '';
  const nvPr = child(child(el, 'nvPicPr'), 'nvPr');
  if (child(nvPr, 'videoFile') || child(nvPr, 'audioFile') || first(nvPr, 'media')) ctx.missing.add('media');
  const blipFill = child(el, 'blipFill');
  const blip = child(blipFill, 'blip');
  const { url } = await imageTarget(blip, ctx);
  if (!url) {
    ctx.missing.add('picture');
    return placeholderBox(box, 'picture');
  }
  // Cropping: srcRect gives the share cut from each side, in 1/1000 %.
  const src = child(blipFill, 'srcRect');
  const cut = (side: string) => Number(src?.getAttribute(side) ?? 0) / 100000;
  const [l, t, r, b] = [cut('l'), cut('t'), cut('r'), cut('b')];
  const iw = box.w / Math.max(0.01, 1 - l - r);
  const ih = box.h / Math.max(0.01, 1 - t - b);
  let imgStyle = `left:${n2(-l * iw)}px;top:${n2(-t * ih)}px;width:${n2(iw)}px;height:${n2(ih)}px`;
  const alpha = child(blip, 'alphaModFix');
  if (alpha) imgStyle += `;opacity:${Number(alpha.getAttribute('amt') ?? 100000) / 100000}`;
  if (child(blip, 'grayscl')) imgStyle += ';filter:grayscale(1)';
  if (box.flipH || box.flipV) imgStyle += `;transform:scale(${box.flipH ? -1 : 1},${box.flipV ? -1 : 1})`;

  const shapes = outlines(spPr, box);
  const prst = child(spPr, 'prstGeom')?.getAttribute('prst');
  const clip = prst && prst !== 'rect' || child(spPr, 'custGeom') ? `;clip-path:path('${shapes.map((o) => o.d).join('')}')` : '';
  const line = await readLine(child(spPr, 'ln'), null, ctx, ctx);
  let outline = '';
  if (line) {
    const defs: string[] = [];
    const stroke = strokeAttrs(line, defs, ctx, box.w, box.h);
    outline = `<svg width="${n2(box.w)}" height="${n2(box.h)}">${defs.length ? `<defs>${defs.join('')}</defs>` : ''}${shapes.map((o) => `<path d="${o.d}" fill="none" ${stroke}/>`).join('')}</svg>`;
  }
  const shadow = shadowCss(child(spPr, 'effectLst'), ctx);
  return `<div class="sp" style="${boxStyle(box, shadow ? `;filter:${shadow}` : '')}"><div class="pic" style="${clip.slice(1)}"><img src="${url}" alt="" style="${imgStyle}"></div>${outline}</div>`;
}

const LABELS: Record<string, string> = { chart: '📊', diagram: '◫', object: '▣', picture: '🖼', media: '▶' };

function placeholderBox(box: Box, kind: string): string {
  return `<div class="sp missing" style="${boxStyle(box)}" data-kind="${kind}"><span>${LABELS[kind] ?? ''}</span></div>`;
}

// ------------------------------------------------------------------ groups and frames

async function renderGroup(el: Element, space: Space, ctx: Ctx, inh: Inherited, isSlide: boolean, groupFill?: Fill): Promise<string> {
  const grpSpPr = child(el, 'grpSpPr');
  const xfrm = child(grpSpPr, 'xfrm');
  const box = boxOf(xfrm, space);
  if (!box) return renderTree(el, space, ctx, inh, isSlide, groupFill);
  const chOff = child(xfrm, 'chOff');
  const chExt = child(xfrm, 'chExt');
  const cw = Number(chExt?.getAttribute('cx') ?? 0);
  const ch = Number(chExt?.getAttribute('cy') ?? 0);
  const inner: Space = {
    ox: Number(chOff?.getAttribute('x') ?? 0),
    oy: Number(chOff?.getAttribute('y') ?? 0),
    kx: cw ? box.w / cw : space.kx,
    ky: ch ? box.h / ch : space.ky,
  };
  const fill = (await readFill(fillElement(grpSpPr), ctx, groupFill)) ?? groupFill;
  let html = '';
  for (const c of children(el)) html += await renderElement(c, inner, ctx, inh, isSlide, fill);
  const flip = box.flipH || box.flipV ? `;transform:${box.rot ? `rotate(${box.rot}deg) ` : ''}scale(${box.flipH ? -1 : 1},${box.flipV ? -1 : 1})` : '';
  return `<div class="grp" style="${boxStyle(box)}${flip}">${html}</div>`;
}

async function renderFrame(el: Element, space: Space, ctx: Ctx, inh: Inherited, isSlide: boolean): Promise<string> {
  const { ph } = inheritance(el, inh, isSlide);
  if (ph && !isSlide) return '';
  const box = boxOf(child(el, 'xfrm'), space);
  const data = first(el, 'graphicData');
  if (!box || !data) return '';
  const uri = data.getAttribute('uri') ?? '';
  if (uri.endsWith('/table')) {
    const tbl = child(data, 'tbl');
    return tbl ? renderTable(tbl, box, ctx) : '';
  }
  if (uri.endsWith('/diagram')) {
    const drawing = await diagramDrawing(data, ctx);
    if (drawing) {
      // The drawing's shapes are positioned relative to the frame.
      const dctx: Ctx = { ...ctx, part: drawing };
      const inner: Space = { ox: 0, oy: 0, kx: space.kx, ky: space.ky };
      const html = await renderTree(first(drawing.root, 'spTree'), inner, dctx, inh, false);
      return `<div class="grp" style="${boxStyle(box)}">${html}</div>`;
    }
    ctx.missing.add('diagram');
    return placeholderBox(box, 'diagram');
  }
  // Embedded objects (OLE) usually carry a picture of themselves.
  const pic = first(data, 'pic');
  if (pic) {
    return renderPicture(withBox(pic, el), space, ctx, inh, false);
  }
  const kind = uri.endsWith('/chart') || uri.includes('chartex') ? 'chart' : 'object';
  if (kind === 'chart') {
    const rel = ctx.part.rels.get(relAttr(first(data, 'chart')));
    const part = rel && !rel.external ? await ctx.deck.part(rel.target) : null;
    const theme = { theme: ctx.theme, clrMap: ctx.clrMap, baseSize: 13.3, frame: false };
    const svg = part ? renderChart(part.root, theme, box.w, box.h) : null;
    if (svg) return `<div class="sp chart" style="${boxStyle(box)}">${svg}</div>`;
  }
  ctx.missing.add(kind);
  return placeholderBox(box, kind);
}

/** An OLE object's picture takes the frame's position when it has none of its own. */
function withBox(pic: Element, frame: Element): Element {
  const spPr = child(pic, 'spPr');
  if (child(spPr, 'xfrm') || !spPr) return pic;
  const copy = pic.cloneNode(true) as Element;
  const xfrm = child(frame, 'xfrm')!.cloneNode(true) as Element;
  child(copy, 'spPr')!.prepend(xfrm);
  return copy;
}

/** The shapes PowerPoint saved for a SmartArt diagram (its "drawing" part), if present. */
async function diagramDrawing(data: Element, ctx: Ctx): Promise<Part | null> {
  const relIds = child(data, 'relIds');
  const dm = ctx.part.rels.get(relAttr(relIds, 'dm'));
  if (dm) {
    const model = await ctx.deck.part(dm.target);
    const ext = first(model?.root, 'dataModelExt');
    const rel = ext ? ctx.part.rels.get(ext.getAttribute('relId') ?? '') : undefined;
    if (rel) return ctx.deck.part(rel.target);
  }
  const fallback = relOfType(ctx.part.rels, '/diagramDrawing');
  return fallback ? ctx.deck.part(fallback.target) : null;
}
