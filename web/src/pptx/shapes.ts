import { t } from '../i18n';
import { child, children } from '../ooxml';
import { toast } from '../ui';
import { EMU } from './deck';
import { topShapes } from './render';

// Moving and resizing a slide's shapes (edit mode). A tap selects one of the slide's own
// top-level elements (shape, text box, picture, table, chart, group): a frame with eight handles
// is drawn over it. A drag on the selected element moves it, a drag on a handle resizes it
// (corners keep a picture's proportions); a drag on anything else still scrolls, and a tap on
// text still puts the caret in it. On release the element's xfrm is written (off/ext, in EMU;
// a group's children scale with it, as its chOff/chExt stay) and the slide is redrawn.

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';
const DIRS: Record<Handle, [number, number]> = {
  nw: [-1, -1], n: [0, -1], ne: [1, -1], e: [1, 0], se: [1, 1], s: [0, 1], sw: [-1, 1], w: [-1, 0],
};
/** Smallest size a drag leaves, in slide px. */
const MIN = 8;
/** Finger travel (screen px) before a press on the selected shape becomes a move. */
const SLOP = 6;

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
  /** Degrees, clockwise. */
  rot: number;
}

export interface ShapeHost {
  pager: HTMLElement;
  /** Slide `i`'s working XML. */
  doc(i: number): Document | null;
  /** Writes pending text edits before a shape changes. */
  flush(): void;
  /** The slide's XML changed: record an undo step and redraw it. */
  changed(i: number): Promise<void>;
}

interface Drag {
  kind: 'move' | Handle;
  pointer: number;
  x0: number;
  y0: number;
  /** Screen px per slide px. */
  scale: number;
  box: Box;
  started: boolean;
  el: HTMLElement;
  keepAspect: boolean;
  /** The box being dragged to, in slide px. */
  now?: Box;
}

export class ShapeTool {
  private sel: { slide: number; n: number } | null = null;
  private frame = document.createElement('div');
  private drag: Drag | null = null;
  /** A drag just ended: the click that follows it is not a tap. */
  private swallowClick = false;

  constructor(private host: ShapeHost) {
    this.frame.className = 'pptx-sel';
    for (const h of Object.keys(DIRS) as Handle[]) {
      const handle = document.createElement('div');
      handle.className = 'pptx-handle';
      handle.dataset.h = h;
      this.frame.append(handle);
    }
    const pager = host.pager;
    pager.addEventListener('pointerdown', (e) => this.down(e));
    pager.addEventListener('dragstart', (e) => e.preventDefault());
    window.addEventListener('pointermove', (e) => this.move(e));
    window.addEventListener('pointerup', (e) => void this.up(e));
    window.addEventListener('pointercancel', () => this.cancel());
    // Capture: runs before the text editor's click handling.
    pager.addEventListener('click', (e) => this.click(e), true);
  }

  /** A slide was drawn again: put the frame back on the selected element. */
  drawn(i: number, slideEl: HTMLElement): void {
    if (this.sel?.slide !== i) return;
    const el = slideEl.querySelector<HTMLElement>(`[data-el="${this.sel.n}"]`);
    if (el) this.show(el);
    else this.clear();
  }

  clear(): void {
    this.sel = null;
    this.frame.remove();
    for (const old of this.host.pager.querySelectorAll('.oe-selected')) old.classList.remove('oe-selected');
  }

  // ---------------------------------------------------------------- selection

  private click(e: MouseEvent): void {
    if (this.swallowClick) {
      this.swallowClick = false;
      e.stopPropagation();
      e.preventDefault();
      return;
    }
    const target = e.target as Element;
    if (target.closest('.pptx-sel')) return;
    const el = target.closest<HTMLElement>('.slide.editing [data-el]');
    if (!el) {
      this.clear();
      return;
    }
    const slide = Number(el.closest<HTMLElement>('.slide')!.dataset.slide);
    const n = Number(el.dataset.el);
    if (this.sel?.slide === slide && this.sel.n === n) return;
    // Selecting a shape ends typing elsewhere, and doesn't put a caret in it yet.
    e.stopPropagation();
    (document.activeElement as HTMLElement | null)?.blur?.();
    if (!this.xml(slide, n) || !canMove(this.xml(slide, n)!)) {
      this.clear();
      // What can't be moved (or selected) still takes text, as before.
      el.querySelector<HTMLElement>('p[contenteditable="true"]')?.focus();
      return;
    }
    this.sel = { slide, n };
    this.show(el);
  }

  private show(el: HTMLElement): void {
    for (const old of this.host.pager.querySelectorAll('.oe-selected')) if (old !== el) old.classList.remove('oe-selected');
    el.classList.add('oe-selected');
    const xml = this.sel && this.xml(this.sel.slide, this.sel.n);
    this.frame.classList.toggle('fixed-size', !xml || !canResize(xml));
    el.closest('.slide')!.append(this.frame);
    this.place(this.boxOf(el));
  }

  private place(box: Box): void {
    const s = this.frame.style;
    s.left = `${box.x}px`;
    s.top = `${box.y}px`;
    s.width = `${box.w}px`;
    s.height = `${box.h}px`;
    s.rotate = box.rot ? `${box.rot}deg` : '';
  }

  /** The element's box in slide px (offsets ignore its rotation, as xfrm does). */
  private boxOf(el: HTMLElement): Box {
    const xml = this.sel ? this.xml(this.sel.slide, this.sel.n) : null;
    const rot = Number(xfrmOf(xml)?.getAttribute('rot') ?? 0) / 60000;
    // The renderer's own values (to 1/100 px) where it wrote them; offsets are whole pixels.
    const px = (v: string, fallback: number) => (v.endsWith('px') ? parseFloat(v) : fallback);
    const st = el.style;
    return { x: px(st.left, el.offsetLeft), y: px(st.top, el.offsetTop), w: px(st.width, el.offsetWidth), h: px(st.height, el.offsetHeight), rot };
  }

  private xml(slide: number, n: number): Element | null {
    const doc = this.host.doc(slide);
    const tree = doc && child(child(doc.documentElement, 'cSld'), 'spTree');
    return topShapes(tree)[n] ?? null;
  }

  // ---------------------------------------------------------------- dragging

  private down(e: PointerEvent): void {
    if (!e.isPrimary || this.drag) return;
    const target = e.target as Element;
    // The first tap on a shape only selects it (click), without a caret or the keyboard, which
    // would move the slide away from the finger; a second tap on its text starts typing.
    const pressed = target.closest<HTMLElement>('.slide.editing [data-el]');
    if (pressed && !pressed.classList.contains('oe-selected')) {
      e.preventDefault();
      return;
    }
    if (!this.sel) return;
    const handle = target.closest<HTMLElement>('.pptx-handle');
    const el = this.host.pager.querySelector<HTMLElement>(`.slide[data-slide="${this.sel.slide}"] [data-el="${this.sel.n}"]`);
    if (!el) return;
    if (!handle && !el.contains(target)) return;
    const slideEl = el.closest<HTMLElement>('.slide')!;
    const xml = this.xml(this.sel.slide, this.sel.n)!;
    const kind = (handle?.dataset.h as Handle | undefined) ?? 'move';
    this.drag = {
      kind,
      pointer: e.pointerId,
      x0: e.clientX,
      y0: e.clientY,
      scale: slideEl.getBoundingClientRect().width / slideEl.offsetWidth || 1,
      box: this.boxOf(el),
      started: false,
      el,
      keepAspect: xml.localName === 'pic' && kind.length === 2,
    };
    // Outside text, a press must not start the browser's own picture drag or a text selection.
    if (!handle && !target.closest('[contenteditable="true"]')) e.preventDefault();
    if (handle) {
      // A handle drag starts at once and must not move the caret or scroll.
      e.preventDefault();
      handle.setPointerCapture(e.pointerId);
      this.start();
    }
  }

  private start(): void {
    const drag = this.drag!;
    drag.started = true;
    this.host.flush();
    // Moving a text box drags it, not a text selection.
    if (drag.kind === 'move') {
      (document.activeElement as HTMLElement | null)?.blur?.();
      getSelection()?.removeAllRanges();
    }
    this.frame.classList.add('dragging');
  }

  private move(e: PointerEvent): void {
    const drag = this.drag;
    if (!drag || e.pointerId !== drag.pointer) return;
    const sx = e.clientX - drag.x0;
    const sy = e.clientY - drag.y0;
    if (!drag.started) {
      if (Math.hypot(sx, sy) < SLOP) return;
      this.start();
    }
    e.preventDefault();
    getSelection()?.removeAllRanges();
    const dx = sx / drag.scale;
    const dy = sy / drag.scale;
    if (drag.kind === 'move') {
      drag.now = { ...drag.box, x: drag.box.x + dx, y: drag.box.y + dy };
      drag.el.style.translate = `${dx}px ${dy}px`;
    } else {
      drag.now = resized(drag.box, drag.kind, dx, dy, drag.keepAspect);
    }
    this.place(drag.now);
  }

  private async up(e: PointerEvent): Promise<void> {
    const drag = this.drag;
    if (!drag || e.pointerId !== drag.pointer) return;
    this.drag = null;
    this.frame.classList.remove('dragging');
    if (!drag.started) return;
    this.swallowClick = true;
    setTimeout(() => (this.swallowClick = false), 400);
    const sel = this.sel!;
    const xml = this.xml(sel.slide, sel.n);
    if (!drag.now || !xml) {
      drag.el.style.translate = '';
      return;
    }
    const xfrm = xfrmOf(xml, true);
    if (!xfrm) {
      drag.el.style.translate = '';
      this.place(drag.box);
      return toast(t('cantMove'));
    }
    write(xfrm, drag.box, drag.now, drag.kind === 'move');
    await this.host.changed(sel.slide);
  }

  private cancel(): void {
    const drag = this.drag;
    if (!drag) return;
    this.drag = null;
    this.frame.classList.remove('dragging');
    drag.el.style.translate = '';
    this.place(drag.box);
  }
}

// ------------------------------------------------------------------ geometry

/**
 * The box after dragging handle `h` by (dx, dy) slide px: the opposite corner or side stays
 * where it is on the slide, also for a rotated shape.
 */
function resized(b: Box, h: Handle, dx: number, dy: number, keepAspect: boolean): Box {
  const [sx, sy] = DIRS[h];
  const th = (b.rot * Math.PI) / 180;
  const cos = Math.cos(th);
  const sin = Math.sin(th);
  // The drag in the shape's own (unrotated) axes.
  const lx = dx * cos + dy * sin;
  const ly = -dx * sin + dy * cos;
  // A line has no height (or width); it keeps none.
  let w = b.w ? Math.max(MIN, b.w + sx * lx) : 0;
  let h2 = b.h ? Math.max(MIN, b.h + sy * ly) : 0;
  if (keepAspect && sx && sy && b.w && b.h) {
    const k = Math.max(w / b.w, h2 / b.h);
    w = b.w * k;
    h2 = b.h * k;
  }
  const cx = b.x + b.w / 2;
  const cy = b.y + b.h / 2;
  // The fixed point, relative to the centre, before and after.
  const ax = (-sx * b.w) / 2;
  const ay = (-sy * b.h) / 2;
  const fx = cx + ax * cos - ay * sin;
  const fy = cy + ax * sin + ay * cos;
  const bx = (-sx * w) / 2;
  const by = (-sy * h2) / 2;
  const ncx = fx - (bx * cos - by * sin);
  const ncy = fy - (bx * sin + by * cos);
  return { x: ncx - w / 2, y: ncy - h2 / 2, w, h: h2, rot: b.rot };
}

// ------------------------------------------------------------------ XML

/** Where an element keeps its position and size; with `create`, a placeholder that inherits its own gets one. */
function xfrmOf(el: Element | null, create = false): Element | null {
  if (!el) return null;
  switch (el.localName) {
    case 'graphicFrame':
      return child(el, 'xfrm');
    case 'grpSp':
      return child(child(el, 'grpSpPr'), 'xfrm');
    case 'sp':
    case 'cxnSp':
    case 'pic': {
      const spPr = child(el, 'spPr');
      let xfrm = child(spPr, 'xfrm');
      if (!xfrm && create && spPr) {
        const doc = el.ownerDocument;
        const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
        const prefix = doc.documentElement.lookupPrefix(a) ?? 'a';
        xfrm = doc.createElementNS(a, `${prefix}:xfrm`);
        xfrm.append(doc.createElementNS(a, `${prefix}:off`), doc.createElementNS(a, `${prefix}:ext`));
        spPr.prepend(xfrm);
      }
      return xfrm;
    }
  }
  return null;
}

/** Writes the new box. A move only shifts the offset, so the size stays exactly as it was. */
function write(xfrm: Element, from: Box, to: Box, moveOnly: boolean): void {
  const off = child(xfrm, 'off')!;
  const ext = child(xfrm, 'ext')!;
  const emu = (px: number) => String(Math.round(px * EMU));
  if (moveOnly && off.hasAttribute('x') && off.hasAttribute('y')) {
    off.setAttribute('x', String(Number(off.getAttribute('x')) + Math.round((to.x - from.x) * EMU)));
    off.setAttribute('y', String(Number(off.getAttribute('y')) + Math.round((to.y - from.y) * EMU)));
    return;
  }
  off.setAttribute('x', emu(to.x));
  off.setAttribute('y', emu(to.y));
  if (!moveOnly || !ext.hasAttribute('cx')) {
    ext.setAttribute('cx', emu(to.w));
    ext.setAttribute('cy', emu(to.h));
  }
}

/** The element's lock flags (spLocks, picLocks, grpSpLocks…), from its non-visual properties. */
function locked(el: Element, flag: string): boolean {
  const nv = el.firstElementChild;
  for (const props of children(nv)) {
    for (const c of children(props)) if (c.localName.endsWith('Locks') && c.getAttribute(flag) === '1') return true;
  }
  return false;
}

function canMove(el: Element): boolean {
  return el.localName !== 'contentPart' && el.localName !== 'AlternateContent' && !locked(el, 'noMove');
}

/** Tables keep the size of their rows and columns (resizing them would mean rewriting both). */
function canResize(el: Element): boolean {
  return !locked(el, 'noResize') && !(el.localName === 'graphicFrame' && el.getElementsByTagNameNS('*', 'tbl').length);
}
