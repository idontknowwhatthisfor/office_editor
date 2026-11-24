import JSZip from 'jszip';
import { t } from '../i18n';
import { overlay } from '../overlay';
import { topInset } from '../immersive';
import { toast } from '../ui';

// Pictures in the Word editor: insert, select, move, resize, crop and delete.
//
// What the editor can change is a picture in its own run (w:r > w:drawing > wp:inline or
// wp:anchor > … > pic:pic with an embedded blip). Its picture on screen is wrapped in a
// span.oe-pic that holds its state as attributes, so the editor's undo (snapshots of the view's
// HTML) covers it: data-pic (the run, "paragraph_run") or data-new (an inserted picture), data-cx
// / data-cy (shown size, EMU) and data-crop (srcRect l,t,r,b in 1/100000). Moving puts the
// picture in a paragraph of its own between the document's top-level paragraphs and tables: a
// picture alone in a top-level paragraph moves with its paragraph (data-moved), one sharing a
// paragraph with text goes into a new one (data-picpara). A top-level paragraph left empty is
// hidden (data-gone) and removed on save. On save (write) only what changed is touched: the
// sizes and crop in the drawing XML, the run moved to its new paragraph, new pictures added to
// word/media with a relationship and a content type.
//
// The selection (frame, handles, the picture's own small toolbar, the crop view and the drop
// line) lives in the editors' overlay (../overlay.ts), outside the view (so it is never saved):
// it stays at normal size on a pinch-zoomed page and isn't cut off at the paper's edge. It is
// placed in screen coordinates and placed again when the page scrolls or zooms. Corners resize
// keeping the picture's proportions; the side handles stretch it one way.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const KEEP = { createFolders: false };

/** EMU per CSS px (96 per inch) and per point. */
const EMU_PX = 9525;
const EMU_PT = 12700;
/** srcRect units: 100000 = the whole picture. */
const WHOLE = 100000;
/** Smallest picture a resize leaves, in CSS px of the page. */
const MIN_PX = 16;
/** Least of the picture a crop keeps, as a fraction. */
const MIN_KEEP = 0.05;
/** Finger travel (screen px) before a press on the selected picture becomes a move. */
const SLOP = 8;

type Corner = 'nw' | 'ne' | 'se' | 'sw';
type Edge = Corner | 'n' | 'e' | 's' | 'w';
const DIRS: Record<Edge, [number, number]> = {
  nw: [-1, -1], n: [0, -1], ne: [1, -1], e: [1, 0], se: [1, 1], s: [0, 1], sw: [-1, 1], w: [-1, 0],
};

/** A picture added in this session: its file, kept until the editor closes. */
interface NewImage {
  bytes: Uint8Array;
  ext: string;
  mime: string;
  name: string;
  url: string;
}

/** A gap between two top-level blocks of a section, where a moved picture can go. */
interface Gap {
  article: HTMLElement;
  before: Element | null;
  y: number;
}

type Crop = [number, number, number, number];

export interface PictureHost {
  view: HTMLElement;
  frame: HTMLElement;
  /** The document as it was opened (indices of data-pic refer to its w:p and w:r). */
  doc: Document;
  /** The view changed: record an undo step and mark the document unsaved. */
  changed(): void;
}

export class PictureTool {
  private layer = el('div', 'oe-pic-layer');
  private sel = el('div', 'oe-pic-sel');
  private bar = el('div', 'oe-pic-bar');
  private dropLine = el('div', 'oe-pic-drop');
  private selected: HTMLElement | null = null;
  private images = new Map<string, NewImage>();
  private counter = 0;
  /** Each editable picture's size and crop as opened, by data-pic. */
  private original = new Map<string, { cx: number; cy: number; crop: Crop }>();
  private drag: {
    kind: 'move' | Edge;
    pointer: number;
    x0: number;
    y0: number;
    started: boolean;
    /** Shown size (EMU) and box (screen px) when the drag began. */
    cx: number;
    cy: number;
    rect: DOMRect;
    ghost?: HTMLElement;
    gap?: Gap | null;
    lastX?: number;
    lastY?: number;
  } | null = null;
  private crop: { box: HTMLElement; crop: Crop; window: HTMLElement; bright: HTMLImageElement } | null = null;
  private swallowClick = false;
  private scrollTimer = 0;
  private input = el('input', '');

  constructor(private host: PictureHost) {
    this.layer.append(this.dropLine);
    this.dropLine.hidden = true;
    for (const h of Object.keys(DIRS) as Edge[]) {
      const handle = el('div', 'oe-pic-handle');
      handle.dataset.h = h;
      this.sel.append(handle);
    }
    this.bar.append(
      barButton(ICON_CROP, t('cropPicture'), () => this.startCrop()),
      barButton(ICON_DELETE, t('deletePicture'), () => this.remove()),
    );
    this.input.type = 'file';
    this.input.accept = 'image/*';
    this.input.hidden = true;
    this.input.onchange = () => {
      const file = this.input.files?.[0];
      this.input.value = '';
      if (file) void this.insertFile(file);
    };

    const view = host.view;
    view.addEventListener('dragstart', (e) => e.preventDefault());
    view.addEventListener('pointerdown', (e) => this.down(e));
    this.layer.addEventListener('pointerdown', (e) => this.down(e));
    window.addEventListener('pointermove', (e) => this.move(e));
    window.addEventListener('pointerup', (e) => this.up(e));
    window.addEventListener('pointercancel', () => this.cancelDrag());
    view.addEventListener('click', (e) => this.click(e), true);
    // The controls are on the screen, not the page: follow the picture when it scrolls or zooms.
    const again = () => this.reposition();
    window.addEventListener('resize', again);
    window.addEventListener('scroll', again, { passive: true });
    window.visualViewport?.addEventListener('scroll', again);
    window.visualViewport?.addEventListener('resize', again);
  }

  /** Wraps every picture the editor can change; call once the document is rendered. */
  prepare(): void {
    const paras = this.host.doc.getElementsByTagNameNS(W, 'p');
    for (const run of this.host.view.querySelectorAll<HTMLElement>('span[data-r]')) {
      const imgs = run.querySelectorAll('img');
      if (imgs.length !== 1) continue;
      const [i, k] = run.dataset.r!.split('_').map(Number);
      const xmlRun = childrenNS(paras[i], W, 'r')[k];
      const drawing = editableDrawing(xmlRun);
      if (!drawing) continue;
      const extent = childNS(drawing, WP, 'extent');
      const cx = Number(extent?.getAttribute('cx')) || 0;
      const cy = Number(extent?.getAttribute('cy')) || 0;
      if (!cx || !cy) continue;
      const img = imgs[0];
      const box = el('span', 'oe-pic');
      box.dataset.pic = run.dataset.r;
      box.dataset.cx = String(cx);
      box.dataset.cy = String(cy);
      const crop = cropOf(drawing);
      box.dataset.crop = crop.join(',');
      const rot = Number(first(drawing, A, 'xfrm')?.getAttribute('rot') ?? 0) / 60000;
      if (rot) box.dataset.rot = String(rot);
      img.replaceWith(box);
      box.append(img);
      img.draggable = false;
      this.original.set(run.dataset.r!, { cx, cy, crop });
      render(box);
      const p = box.closest('p');
      if (p && !p.textContent?.trim()) p.classList.add('oe-piconly');
    }
    overlay().append(this.layer);
    document.body.append(this.input);
  }

  /** The view was replaced (undo, redo): forget the selection. */
  clear(): void {
    this.endCrop(false);
    this.selected?.classList.remove('oe-selected');
    this.selected = null;
    this.sel.remove();
    this.bar.remove();
  }

  // ---------------------------------------------------------------- insert

  /** The picture button: the system's picker, then the picture goes after the current paragraph. */
  pick(): void {
    this.input.click();
  }

  private async insertFile(file: File): Promise<void> {
    let image: { bytes: Uint8Array; ext: string; mime: string; width: number; height: number };
    try {
      image = await readImage(file);
    } catch (e) {
      console.error('picture unreadable', e);
      toast(t('pictureUnreadable'));
      return;
    }
    const id = `n${++this.counter}`;
    const url = URL.createObjectURL(new Blob([image.bytes as BlobPart], { type: image.mime }));
    this.images.set(id, { bytes: image.bytes, ext: image.ext, mime: image.mime, name: file.name || 'Picture', url });

    const { article, before } = this.insertionPoint();
    // At most the text width, and not taller than it is wide plus a fifth.
    const maxW = Math.max(MIN_PX, article.clientWidth || 600);
    let w = Math.min(image.width, maxW);
    let h = (w * image.height) / image.width;
    if (h > maxW * 1.2) {
      h = maxW * 1.2;
      w = (h * image.width) / image.height;
    }
    const box = el('span', 'oe-pic');
    box.dataset.new = id;
    box.dataset.cx = String(Math.round(w * EMU_PX));
    box.dataset.cy = String(Math.round(h * EMU_PX));
    box.dataset.crop = '0,0,0,0';
    const img = el('img', '');
    img.src = url;
    img.alt = '';
    img.draggable = false;
    box.append(img);
    const p = pictureParagraph(box);
    article.insertBefore(p, before);
    render(box);
    this.host.changed();
    this.select(box);
    showOnScreen(box);
  }

  /** After the top-level block holding the caret; with no caret, at the end of the document. */
  private insertionPoint(): { article: HTMLElement; before: Element | null } {
    const node = getSelection()?.anchorNode;
    const at = (node instanceof Element ? node : node?.parentElement)?.closest('p');
    const view = this.host.view;
    if (at && view.contains(at)) {
      const block = topLevel(at);
      if (block) return { article: block.parentElement!, before: block.nextElementSibling };
    }
    const articles = view.querySelectorAll<HTMLElement>('section.docx > article');
    return { article: articles[articles.length - 1], before: null };
  }

  // ---------------------------------------------------------------- selection

  private click(e: MouseEvent): void {
    if (this.swallowClick) {
      this.swallowClick = false;
      e.stopPropagation();
      e.preventDefault();
      return;
    }
    // A tap outside the picture being cropped ends cropping.
    if (this.crop) {
      this.endCrop(true);
      return;
    }
    const box = (e.target as Element).closest<HTMLElement>('.oe-pic');
    if (box && this.host.view.contains(box)) {
      e.stopPropagation();
      if (box !== this.selected) this.select(box);
    } else if (this.selected) {
      this.clear();
    }
  }

  private select(box: HTMLElement): void {
    this.clear();
    // The keyboard and the text caret go away, as for a shape in PowerPoint.
    (document.activeElement as HTMLElement | null)?.blur?.();
    getSelection()?.removeAllRanges();
    this.selected = box;
    box.classList.add('oe-selected');
    this.layer.append(this.sel, this.bar);
    this.place();
  }

  /** Puts the frame and the toolbar on the selected picture. */
  private place(): void {
    const box = this.selected;
    if (!box || !box.isConnected) return;
    const r = this.local(box.getBoundingClientRect());
    setRect(this.sel, r);
    this.placeBar(r);
  }

  /** The toolbar centred above `r` (layer coordinates), or below it when there is no room above. */
  private placeBar(r: DOMRect): void {
    // The layer's own size (it is not scaled with a pinch-zoom, so offsetWidth is on-screen size).
    const bar = { width: this.bar.offsetWidth, height: this.bar.offsetHeight };
    const width = document.documentElement.clientWidth;
    const left = Math.min(Math.max(4, r.x + r.width / 2 - bar.width / 2), width - bar.width - 4);
    const above = r.y - bar.height - 14;
    this.bar.style.left = `${left}px`;
    this.bar.style.top = `${above >= topInset() + 8 ? above : r.y + r.height + 14}px`;
  }

  /**
   * A rectangle from getBoundingClientRect (or pointer coordinates), which are in the page's
   * layout viewport, in the overlay's coordinates: the visible area, at normal size.
   */
  private local(r: DOMRect): DOMRect {
    const v = window.visualViewport;
    const s = v?.scale ?? 1;
    const x = v?.offsetLeft ?? 0;
    const y = v?.offsetTop ?? 0;
    return new DOMRect((r.left - x) * s, (r.top - y) * s, r.width * s, r.height * s);
  }

  /** The page scrolled or zoomed: the controls follow the picture. */
  private reposition(): void {
    if (this.crop) this.placeCrop();
    else if (!this.drag?.started) this.place();
  }

  // ---------------------------------------------------------------- move and resize

  private down(e: PointerEvent): void {
    if (!e.isPrimary || this.drag) return;
    const target = e.target as Element;
    if (this.crop) {
      const handle = target.closest<HTMLElement>('.oe-crop-handle');
      if (handle) this.cropDrag(e, handle.dataset.h as Edge);
      return;
    }
    const box = target.closest<HTMLElement>('.oe-pic');
    // An unselected picture: no browser picture drag or text selection; the click selects it.
    if (box && box !== this.selected) {
      e.preventDefault();
      return;
    }
    const handle = target.closest<HTMLElement>('.oe-pic-handle');
    if (!this.selected || (!handle && box !== this.selected)) return;
    e.preventDefault();
    const sel = this.selected;
    this.drag = {
      kind: (handle?.dataset.h as Edge | undefined) ?? 'move',
      pointer: e.pointerId,
      x0: e.clientX,
      y0: e.clientY,
      started: !!handle,
      cx: Number(sel.dataset.cx),
      cy: Number(sel.dataset.cy),
      rect: sel.getBoundingClientRect(),
    };
    if (handle) {
      handle.setPointerCapture(e.pointerId);
      this.sel.classList.add('dragging');
      this.bar.remove();
    }
  }

  private move(e: PointerEvent): void {
    const d = this.drag;
    if (!d || e.pointerId !== d.pointer) return;
    const dx = e.clientX - d.x0;
    const dy = e.clientY - d.y0;
    if (!d.started) {
      if (Math.hypot(dx, dy) < SLOP) return;
      d.started = true;
      this.bar.remove();
      this.sel.classList.add('dragging');
    }
    e.preventDefault();
    const box = this.selected!;
    if (d.kind === 'move') {
      d.lastX = e.clientX;
      d.lastY = e.clientY;
      this.showMove(e.clientX, e.clientY);
      this.autoScroll();
      return;
    }
    const [sx, sy] = DIRS[d.kind];
    const maxW = (box.closest('article')?.clientWidth ?? Infinity) * EMU_PX;
    const min = MIN_PX * EMU_PX;
    if (sx && sy) {
      // A corner: the picture keeps its proportions.
      const f = Math.max((d.rect.width + sx * dx) / d.rect.width, (d.rect.height + sy * dy) / d.rect.height);
      const k = Math.min(Math.max(f, min / Math.min(d.cx, d.cy)), maxW / d.cx);
      box.dataset.cx = String(Math.round(d.cx * k));
      box.dataset.cy = String(Math.round(d.cy * k));
    } else {
      // A side: stretched (or squeezed) that way only.
      const fx = sx ? (d.rect.width + sx * dx) / d.rect.width : 1;
      const fy = sy ? (d.rect.height + sy * dy) / d.rect.height : 1;
      box.dataset.cx = String(Math.round(clamp(d.cx * fx, min, Math.max(min, maxW))));
      box.dataset.cy = String(Math.round(Math.max(d.cy * fy, min)));
    }
    render(box);
    // The frame follows the picture, which may reflow with the text around it.
    setRect(this.sel, this.local(box.getBoundingClientRect()));
  }

  private up(e: PointerEvent): void {
    const d = this.drag;
    if (!d || e.pointerId !== d.pointer) return;
    this.drag = null;
    clearInterval(this.scrollTimer);
    this.sel.classList.remove('dragging');
    d.ghost?.remove();
    this.dropLine.hidden = true;
    if (!d.started) return;
    this.swallowClick = true;
    setTimeout(() => (this.swallowClick = false), 400);
    const box = this.selected!;
    if (d.kind === 'move') {
      if (d.gap && this.moveTo(box, d.gap)) this.host.changed();
    } else if (box.dataset.cx !== String(d.cx) || box.dataset.cy !== String(d.cy)) {
      this.host.changed();
    }
    this.layer.append(this.bar);
    this.place();
  }

  private cancelDrag(): void {
    const d = this.drag;
    if (!d) return;
    this.drag = null;
    clearInterval(this.scrollTimer);
    d.ghost?.remove();
    this.dropLine.hidden = true;
    this.sel.classList.remove('dragging');
    if (d.kind !== 'move' && this.selected) {
      this.selected.dataset.cx = String(d.cx);
      this.selected.dataset.cy = String(d.cy);
      render(this.selected);
    }
    this.layer.append(this.bar);
    this.place();
  }

  /** While moving: the picture follows the finger, and a line shows where it will go. */
  private showMove(x: number, y: number): void {
    const d = this.drag!;
    const box = this.selected!;
    if (!d.ghost) {
      d.ghost = el('div', 'oe-pic-ghost');
      const img = box.querySelector('img')!.cloneNode() as HTMLImageElement;
      img.removeAttribute('style');
      d.ghost.append(img);
      this.layer.append(d.ghost);
    }
    // Half size, under the finger.
    const w = d.rect.width / 2;
    const h = d.rect.height / 2;
    const at = this.local(new DOMRect(x - w / 2, y - h / 2, w, h));
    setRect(d.ghost, at);
    d.gap = this.gapAt(y);
    if (!d.gap) {
      this.dropLine.hidden = true;
      return;
    }
    const a = this.local(d.gap.article.getBoundingClientRect());
    const line = this.local(new DOMRect(0, d.gap.y, 0, 0));
    this.dropLine.hidden = false;
    setRect(this.dropLine, new DOMRect(a.x, line.y - 1.5, a.width, 3));
  }

  /** The gap between top-level blocks nearest to screen y. */
  private gapAt(y: number): Gap | null {
    let best: Gap | null = null;
    let bestDist = Infinity;
    for (const article of this.host.view.querySelectorAll<HTMLElement>('section.docx > article')) {
      const kids = [...article.children].filter((c) => !(c as HTMLElement).hidden && (c as HTMLElement).dataset.gone === undefined);
      const rects = kids.map((c) => c.getBoundingClientRect());
      const consider = (before: Element | null, gy: number) => {
        const dist = Math.abs(gy - y);
        if (dist < bestDist) {
          bestDist = dist;
          best = { article, before, y: gy };
        }
      };
      if (!kids.length) {
        consider(null, article.getBoundingClientRect().top);
        continue;
      }
      consider(kids[0], rects[0].top);
      for (let j = 1; j < kids.length; j++) consider(kids[j], (rects[j - 1].bottom + rects[j].top) / 2);
      consider(null, rects[rects.length - 1].bottom);
    }
    return best;
  }

  /** Moves the picture to a gap; false if that is where it already is. */
  private moveTo(box: HTMLElement, gap: Gap): boolean {
    const p = box.closest('p') as HTMLElement | null;
    const container = holderOf(box);
    const alone = !!p && topLevel(p) === p && onlyPicture(p, box);
    if (alone) {
      // The whole paragraph goes, keeping its alignment and spacing.
      if (gap.before === p || gap.before === p!.nextElementSibling) return false;
      gap.article.insertBefore(p!, gap.before);
      if (p!.dataset.p !== undefined) p!.dataset.moved = '1';
    } else {
      const np = pictureParagraph(container);
      gap.article.insertBefore(np, gap.before);
      if (p) this.leftEmpty(p);
    }
    render(box);
    showOnScreen(box);
    return true;
  }

  /** A top-level paragraph that no longer holds anything is hidden (and removed on save). */
  private leftEmpty(p: HTMLElement): void {
    if (topLevel(p) !== p || p.textContent?.trim() || p.querySelector('.oe-pic, img, svg, canvas')) return;
    if (p.dataset.picpara !== undefined) p.remove();
    else {
      p.dataset.gone = '1';
      p.hidden = true;
    }
  }

  /** Near the top or bottom of the screen while moving, the page scrolls. */
  private autoScroll(): void {
    clearInterval(this.scrollTimer);
    const d = this.drag;
    if (!d?.lastY) return;
    const edge = 72;
    const h = window.innerHeight;
    const speed = d.lastY < edge ? -(edge - d.lastY) / 3 : d.lastY > h - edge - 56 ? (d.lastY - (h - edge - 56)) / 3 : 0;
    if (!speed) return;
    this.scrollTimer = window.setInterval(() => {
      window.scrollBy(0, speed);
      if (this.drag?.lastY !== undefined) this.showMove(this.drag.lastX ?? this.drag.x0, this.drag.lastY);
    }, 16);
  }

  // ---------------------------------------------------------------- delete

  private remove(): void {
    const box = this.selected;
    if (!box) return;
    const p = box.closest('p') as HTMLElement | null;
    this.clear();
    holderOf(box).remove();
    if (p) this.leftEmpty(p);
    this.host.changed();
  }

  // ---------------------------------------------------------------- crop

  private startCrop(): void {
    const box = this.selected;
    if (!box) return;
    if (box.dataset.rot) {
      toast(t('cropRotated'));
      return;
    }
    const crop = cropData(box);
    const src = box.querySelector('img')!.src;
    const area = el('div', 'oe-crop');
    const dim = el('img', 'oe-crop-dim');
    dim.src = src;
    const bright = el('img', 'oe-crop-bright');
    bright.src = src;
    const win = el('div', 'oe-crop-window');
    for (const h of Object.keys(DIRS) as Edge[]) {
      const handle = el('div', 'oe-crop-handle');
      handle.dataset.h = h;
      win.append(handle);
    }
    area.append(dim, bright, win);
    // The crop view shows the picture instead (the dimmed part must not show through).
    box.classList.add('oe-cropping');
    this.sel.remove();
    this.bar.replaceChildren(barButton(ICON_DONE, t('done'), () => this.endCrop(true)));
    this.layer.append(area, this.bar);
    this.crop = { box, crop: [...crop] as Crop, window: area, bright };
    this.drawCrop();
    this.placeCrop();
  }

  /** The whole picture (of which the box shows the part not cropped), in page coordinates. */
  private fullRect(box: HTMLElement): DOMRect {
    const crop = cropData(box);
    const r = box.getBoundingClientRect();
    const fw = r.width / (1 - (crop[0] + crop[2]) / WHOLE);
    const fh = r.height / (1 - (crop[1] + crop[3]) / WHOLE);
    return new DOMRect(r.left - (crop[0] / WHOLE) * fw, r.top - (crop[1] / WHOLE) * fh, fw, fh);
  }

  /** The crop view over the whole picture, with the toolbar above it. */
  private placeCrop(): void {
    const c = this.crop!;
    const full = this.local(this.fullRect(c.box));
    setRect(c.window, full);
    this.placeBar(full);
  }

  private drawCrop(): void {
    const c = this.crop!;
    const [l, tp, r, b] = c.crop.map((v) => (v / WHOLE) * 100);
    c.bright.style.clipPath = `inset(${tp}% ${r}% ${b}% ${l}%)`;
    const win = c.window.querySelector<HTMLElement>('.oe-crop-window')!;
    win.style.left = `${l}%`;
    win.style.top = `${tp}%`;
    win.style.right = `${r}%`;
    win.style.bottom = `${b}%`;
  }

  private cropDrag(e: PointerEvent, edge: Edge): void {
    e.preventDefault();
    const c = this.crop!;
    const handle = e.target as HTMLElement;
    handle.setPointerCapture(e.pointerId);
    const start = [...c.crop] as Crop;
    const full = this.fullRect(c.box);
    const x0 = e.clientX;
    const y0 = e.clientY;
    const [sx, sy] = DIRS[edge];
    const onMove = (m: PointerEvent) => {
      if (m.pointerId !== e.pointerId) return;
      const fx = ((m.clientX - x0) / full.width) * WHOLE;
      const fy = ((m.clientY - y0) / full.height) * WHOLE;
      const next = [...start] as Crop;
      const most = WHOLE * (1 - MIN_KEEP);
      // A left/top edge moves its own crop in; a right/bottom edge moves the opposite way.
      if (sx < 0) next[0] = clamp(start[0] + fx, 0, most - start[2]);
      if (sx > 0) next[2] = clamp(start[2] - fx, 0, most - start[0]);
      if (sy < 0) next[1] = clamp(start[1] + fy, 0, most - start[3]);
      if (sy > 0) next[3] = clamp(start[3] - fy, 0, most - start[1]);
      c.crop = next.map(Math.round) as Crop;
      this.drawCrop();
    };
    const onUp = (u: PointerEvent) => {
      if (u.pointerId !== e.pointerId) return;
      handle.removeEventListener('pointermove', onMove);
      handle.removeEventListener('pointerup', onUp);
      handle.removeEventListener('pointercancel', onUp);
    };
    handle.addEventListener('pointermove', onMove);
    handle.addEventListener('pointerup', onUp);
    handle.addEventListener('pointercancel', onUp);
  }

  /** Leaves crop mode; `apply`: the box takes the new crop (its size changes with it). */
  private endCrop(apply: boolean): void {
    const c = this.crop;
    if (!c) return;
    this.crop = null;
    c.window.remove();
    c.box.classList.remove('oe-cropping');
    this.bar.replaceChildren(
      barButton(ICON_CROP, t('cropPicture'), () => this.startCrop()),
      barButton(ICON_DELETE, t('deletePicture'), () => this.remove()),
    );
    if (!apply) return;
    const box = c.box;
    const old = cropData(box);
    if (old.join() !== c.crop.join()) {
      const fullCx = Number(box.dataset.cx) / (1 - (old[0] + old[2]) / WHOLE);
      const fullCy = Number(box.dataset.cy) / (1 - (old[1] + old[3]) / WHOLE);
      box.dataset.cx = String(Math.round(fullCx * (1 - (c.crop[0] + c.crop[2]) / WHOLE)));
      box.dataset.cy = String(Math.round(fullCy * (1 - (c.crop[1] + c.crop[3]) / WHOLE)));
      box.dataset.crop = c.crop.join(',');
      render(box);
      this.host.changed();
    }
    this.layer.append(this.sel);
    this.place();
  }

  // ---------------------------------------------------------------- saving

  /**
   * Writes the pictures' changes into `doc` (the copy being saved, whose w:p are `paras` as
   * opened) and adds new pictures' files to `zip`. `written` maps each view paragraph with
   * data-p to its XML paragraph in `doc`.
   */
  async write(doc: Document, zip: JSZip, docPath: string, paras: Element[], written: Map<HTMLElement, Element>): Promise<void> {
    this.endCrop(true);
    const view = this.host.view;
    const runOf = (id: string) => {
      const [i, k] = id.split('_').map(Number);
      return childrenNS(paras[i], W, 'r')[k] ?? null;
    };

    // Sizes and crops of the pictures that are still there; runs of deleted ones go.
    const present = new Set<string>();
    for (const box of view.querySelectorAll<HTMLElement>('.oe-pic[data-pic]')) {
      const id = box.dataset.pic!;
      present.add(id);
      const was = this.original.get(id);
      const crop = cropData(box);
      const cx = Number(box.dataset.cx);
      const cy = Number(box.dataset.cy);
      if (was && was.cx === cx && was.cy === cy && was.crop.join() === crop.join()) continue;
      const drawing = editableDrawing(runOf(id));
      if (drawing) writeSize(doc, drawing, cx, cy, crop);
    }
    for (const id of this.original.keys()) if (!present.has(id)) runOf(id)?.remove();

    // New pictures' files.
    const media = await this.addMedia(zip, docPath);
    let docPrId = Math.max(0, ...[...doc.getElementsByTagNameNS(WP, 'docPr')].map((e) => Number(e.getAttribute('id')) || 0));

    // Moved and new paragraphs, in the order they are on screen, each after the block before it.
    const body = doc.getElementsByTagNameNS(W, 'body')[0];
    for (const v of view.querySelectorAll<HTMLElement>('p[data-picpara], p[data-moved]')) {
      if (v.dataset.gone !== undefined) continue;
      let x: Element;
      if (v.dataset.picpara !== undefined) {
        x = doc.createElementNS(W, 'w:p');
        for (const box of v.querySelectorAll<HTMLElement>('.oe-pic')) {
          if (box.dataset.pic) {
            const run = runOf(box.dataset.pic);
            if (run) x.append(run);
          } else if (box.dataset.new && media.has(box.dataset.new)) {
            x.append(newRun(doc, media.get(box.dataset.new)!, Number(box.dataset.cx), Number(box.dataset.cy), cropData(box), ++docPrId));
          }
        }
        if (!x.childNodes.length) continue;
      } else {
        x = written.get(v) ?? paras[Number(v.dataset.p)];
      }
      written.set(v, x);
      const prev = previousBlock(v);
      const prevXml = prev && xmlOfBlock(prev, written, paras);
      if (prevXml) {
        prevXml.after(x);
        continue;
      }
      const next = nextBlock(v);
      const nextXml = next && xmlOfBlock(next, written, paras);
      if (nextXml) nextXml.before(x);
      else body.insertBefore(x, childNS(body, W, 'sectPr'));
    }

    // Paragraphs left empty by a move or a delete (not ones that end a section).
    for (const v of view.querySelectorAll<HTMLElement>('p[data-gone][data-p]')) {
      const x = paras[Number(v.dataset.p)];
      if (!x || first(x, W, 'sectPr') || x.parentElement?.localName === 'tc') continue;
      if ([...x.getElementsByTagNameNS(W, 't')].some((t) => t.textContent) || x.getElementsByTagNameNS(W, 'drawing').length) continue;
      x.remove();
    }
  }

  /** Adds the new pictures still in the document to word/media; their relationship ids. */
  private async addMedia(zip: JSZip, docPath: string): Promise<Map<string, { rId: string; name: string }>> {
    const out = new Map<string, { rId: string; name: string }>();
    const ids = [...this.host.view.querySelectorAll<HTMLElement>('.oe-pic[data-new]')].map((b) => b.dataset.new!);
    if (!ids.length) return out;
    const dir = docPath.includes('/') ? docPath.slice(0, docPath.lastIndexOf('/')) : '';
    const base = docPath.slice(docPath.lastIndexOf('/') + 1);
    const relsPath = `${dir}/_rels/${base}.rels`;
    const relsXml = (await zip.file(relsPath)?.async('string'))
      ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';
    const rels = new DOMParser().parseFromString(relsXml, 'application/xml');
    const usedIds = new Set([...rels.documentElement.children].map((r) => r.getAttribute('Id')));
    let types = (await zip.file('[Content_Types].xml')!.async('string'));
    let n = 0;
    for (const id of ids) {
      const image = this.images.get(id);
      if (!image || out.has(id)) continue;
      let path: string;
      do path = `media/oe-image${++n}.${image.ext}`; while (zip.file(`${dir}/${path}`));
      zip.file(`${dir}/${path}`, image.bytes, KEEP);
      let rId: string;
      do rId = `rIdOeImg${n++}`; while (usedIds.has(rId));
      usedIds.add(rId);
      const rel = rels.createElementNS(rels.documentElement.namespaceURI, 'Relationship');
      rel.setAttribute('Id', rId);
      rel.setAttribute('Type', IMAGE_TYPE);
      rel.setAttribute('Target', path);
      rels.documentElement.append(rel);
      if (!new RegExp(`<Default [^>]*Extension="${image.ext}"`, 'i').test(types)) {
        types = types.replace('<Default ', `<Default Extension="${image.ext}" ContentType="${image.mime}"/><Default `);
      }
      out.set(id, { rId, name: image.name });
    }
    zip.file(relsPath, new XMLSerializer().serializeToString(rels), KEEP);
    zip.file('[Content_Types].xml', types, KEEP);
    return out;
  }
}

// ------------------------------------------------------------------ the view

/** Shows a picture box at its size and crop. */
function render(box: HTMLElement): void {
  const img = box.querySelector('img')!;
  const w = Number(box.dataset.cx) / EMU_PT;
  const h = Number(box.dataset.cy) / EMU_PT;
  const [l, t, r, b] = cropData(box).map((v) => v / WHOLE);
  const fw = w / (1 - l - r);
  const fh = h / (1 - t - b);
  box.style.width = `${w}pt`;
  box.style.height = `${h}pt`;
  img.style.position = 'absolute';
  img.style.left = `${-l * fw}pt`;
  img.style.top = `${-t * fh}pt`;
  img.style.width = `${fw}pt`;
  img.style.height = `${fh}pt`;
  img.style.maxWidth = 'none';
  img.style.clipPath = '';
  img.style.transform = box.dataset.rot ? `rotate(${box.dataset.rot}deg)` : '';
  // docx-preview's inline holder has the picture's size too.
  const holder = box.parentElement;
  if (holder && holder.tagName === 'DIV' && holder.style.display === 'inline-block') {
    holder.style.width = `${w}pt`;
    holder.style.height = `${h}pt`;
  }
}

/** The element a picture sits in within its paragraph (docx-preview's holder, or the box). */
function holderOf(box: HTMLElement): HTMLElement {
  const holder = box.parentElement;
  return holder && holder.tagName === 'DIV' && holder.childElementCount === 1 ? holder : box;
}

function pictureParagraph(content: HTMLElement): HTMLElement {
  const p = el('p', 'oe-picpara oe-piconly');
  p.dataset.picpara = '';
  const holder = content.classList.contains('oe-pic') ? wrapHolder(content) : content;
  p.append(holder);
  return p;
}

function wrapHolder(box: HTMLElement): HTMLElement {
  const holder = el('div', '');
  holder.style.display = 'inline-block';
  holder.style.position = 'relative';
  holder.style.textIndent = '0';
  holder.append(box);
  return holder;
}

/** The paragraph's block among its section's top-level blocks, if it is one. */
function topLevel(p: Element): HTMLElement | null {
  let e: Element | null = p;
  while (e && e.parentElement && !(e.parentElement.tagName === 'ARTICLE' && e.parentElement.parentElement?.matches('section.docx'))) e = e.parentElement;
  return e?.parentElement ? (e as HTMLElement) : null;
}

function onlyPicture(p: HTMLElement, box: HTMLElement): boolean {
  return !p.textContent?.trim() && [...p.querySelectorAll('.oe-pic')].every((b) => b === box);
}

function previousBlock(v: HTMLElement): HTMLElement | null {
  let e = v.previousElementSibling as HTMLElement | null;
  while (e && e.dataset.gone !== undefined) e = e.previousElementSibling as HTMLElement | null;
  return e;
}

function nextBlock(v: HTMLElement): HTMLElement | null {
  let e = v.nextElementSibling as HTMLElement | null;
  while (e && e.dataset.gone !== undefined) e = e.nextElementSibling as HTMLElement | null;
  return e;
}

/** The XML element of a top-level block on screen: its paragraph, or the table it is in. */
function xmlOfBlock(block: HTMLElement, written: Map<HTMLElement, Element>, paras: Element[]): Element | null {
  if (block.tagName === 'P') {
    return written.get(block) ?? (block.dataset.p !== undefined ? paras[Number(block.dataset.p)] : null) ?? null;
  }
  const inner = block.querySelector<HTMLElement>('p[data-p]');
  let x = inner ? written.get(inner) ?? paras[Number(inner.dataset.p)] : null;
  // Up to the element that sits in the body (the table, or a content control around it).
  while (x && x.parentElement && x.parentElement.localName !== 'body') x = x.parentElement;
  return x ?? null;
}

function cropData(box: HTMLElement): Crop {
  const c = (box.dataset.crop ?? '').split(',').map(Number);
  return [0, 1, 2, 3].map((i) => c[i] || 0) as Crop;
}

// ------------------------------------------------------------------ the XML

/** The run's wp:inline / wp:anchor, if it holds one embedded picture and nothing else drawn. */
function editableDrawing(run: Element | null | undefined): Element | null {
  if (!run) return null;
  const drawings = childrenNS(run, W, 'drawing');
  if (drawings.length !== 1) return null;
  const wrapper = [...drawings[0].children].find((c) => c.namespaceURI === WP && (c.localName === 'inline' || c.localName === 'anchor'));
  if (!wrapper) return null;
  const pic = first(wrapper, PIC, 'pic');
  const blip = pic && first(pic, A, 'blip');
  return blip?.getAttributeNS(R, 'embed') ? wrapper : null;
}

function cropOf(drawing: Element): Crop {
  const rect = first(drawing, A, 'srcRect');
  // Negative values (the picture padded out) are left as they are: shown as no crop.
  return ['l', 't', 'r', 'b'].map((k) => Math.max(0, Number(rect?.getAttribute(k)) || 0)) as Crop;
}

function writeSize(doc: Document, drawing: Element, cx: number, cy: number, crop: Crop): void {
  const extent = childNS(drawing, WP, 'extent');
  extent?.setAttribute('cx', String(cx));
  extent?.setAttribute('cy', String(cy));
  const pic = first(drawing, PIC, 'pic')!;
  const ext = first(first(pic, A, 'xfrm'), A, 'ext');
  ext?.setAttribute('cx', String(cx));
  ext?.setAttribute('cy', String(cy));
  const blipFill = first(pic, PIC, 'blipFill');
  if (!blipFill) return;
  let rect = childNS(blipFill, A, 'srcRect');
  if (crop.every((v) => !v)) {
    // Keep negative (padding) values a file had; otherwise no crop.
    if (rect && !['l', 't', 'r', 'b'].some((k) => Number(rect!.getAttribute(k)) < 0)) rect.remove();
    return;
  }
  if (!rect) {
    rect = doc.createElementNS(A, 'a:srcRect');
    const blip = childNS(blipFill, A, 'blip');
    blipFill.insertBefore(rect, blip ? blip.nextSibling : blipFill.firstChild);
  }
  ['l', 't', 'r', 'b'].forEach((k, i) => {
    if (crop[i]) rect!.setAttribute(k, String(crop[i]));
    else rect!.removeAttribute(k);
  });
}

/** A run holding a new inline picture, as Word writes one. */
function newRun(doc: Document, media: { rId: string; name: string }, cx: number, cy: number, crop: Crop, id: number): Element {
  const name = escapeXml(media.name);
  const srcRect = crop.some((v) => v) ? `<a:srcRect${['l', 't', 'r', 'b'].map((k, i) => (crop[i] ? ` ${k}="${crop[i]}"` : '')).join('')}/>` : '';
  const xml = `<w:r xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:r="${R}"><w:drawing>`
    + `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
    + `<wp:docPr id="${id}" name="Picture ${id}" descr="${name}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`
    + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>`
    + `<pic:nvPicPr><pic:cNvPr id="0" name="${name}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${media.rId}"/>${srcRect}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
    + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const parsed = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.importNode(parsed.documentElement, true);
}

const escapeXml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function childNS(el: Element | null | undefined, ns: string, name: string): Element | null {
  if (!el) return null;
  for (const c of el.children) if (c.localName === name && c.namespaceURI === ns) return c;
  return null;
}

function childrenNS(el: Element | null | undefined, ns: string, name: string): Element[] {
  return el ? [...el.children].filter((c) => c.localName === name && c.namespaceURI === ns) : [];
}

function first(el: Element | null | undefined, ns: string, name: string): Element | null {
  return el?.getElementsByTagNameNS(ns, name)[0] ?? null;
}

// ------------------------------------------------------------------ reading a picked image

/**
 * A picked image as Word can hold it: PNG, GIF and JPEG files are kept as they are (a JPEG
 * turned by its EXIF orientation is drawn upright first, since not every reader turns it);
 * anything else the browser can read becomes a PNG.
 */
async function readImage(file: File): Promise<{ bytes: Uint8Array; ext: string; mime: string; width: number; height: number }> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const bitmap = await createImageBitmap(new Blob([bytes as BlobPart], { type: file.type }));
  const { width, height } = bitmap;
  try {
    const type = sniff(bytes);
    if (type === 'png') return { bytes, ext: 'png', mime: 'image/png', width, height };
    if (type === 'gif') return { bytes, ext: 'gif', mime: 'image/gif', width, height };
    if (type === 'jpeg' && exifOrientation(bytes) <= 1) return { bytes, ext: 'jpeg', mime: 'image/jpeg', width, height };
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    canvas.getContext('2d')!.drawImage(bitmap, 0, 0);
    const photo = type === 'jpeg' || /hei[cf]/i.test(file.type);
    const mime = photo ? 'image/jpeg' : 'image/png';
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, mime, 0.92));
    if (!blob) throw new Error('could not encode the picture');
    return { bytes: new Uint8Array(await blob.arrayBuffer()), ext: photo ? 'jpeg' : 'png', mime, width, height };
  } finally {
    bitmap.close();
  }
}

function sniff(b: Uint8Array): 'png' | 'gif' | 'jpeg' | null {
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png';
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif';
  if (b[0] === 0xff && b[1] === 0xd8) return 'jpeg';
  return null;
}

/** A JPEG's EXIF orientation (1 = upright, also when there is none). */
function exifOrientation(b: Uint8Array): number {
  const view = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let at = 2;
  while (at + 4 < b.length && b[at] === 0xff) {
    const marker = b[at + 1];
    const size = view.getUint16(at + 2);
    if (marker === 0xe1 && view.getUint32(at + 4) === 0x45786966) {
      const tiff = at + 10;
      const little = view.getUint16(tiff) === 0x4949;
      const ifd = tiff + view.getUint32(tiff + 4, little);
      const count = view.getUint16(ifd, little);
      for (let i = 0; i < count; i++) {
        const entry = ifd + 2 + i * 12;
        if (entry + 10 > b.length) break;
        if (view.getUint16(entry, little) === 0x0112) return view.getUint16(entry + 8, little);
      }
      return 1;
    }
    if (marker === 0xda) break;
    at += 2 + size;
  }
  return 1;
}

// ------------------------------------------------------------------ small things

/** Scrolls the page (not the paper's frame) so the picture is on screen, above the toolbar. */
function showOnScreen(e: HTMLElement): void {
  const r = e.getBoundingClientRect();
  const bottom = window.innerHeight - 72;
  const top = topInset() + 64;
  if (r.top >= top && r.bottom <= bottom) return;
  window.scrollBy({ top: r.top - Math.max(top, (bottom - r.height) / 2), behavior: 'smooth' });
}

function setRect(e: HTMLElement, r: DOMRect): void {
  e.style.left = `${r.x}px`;
  e.style.top = `${r.y}px`;
  e.style.width = `${r.width}px`;
  e.style.height = `${r.height}px`;
}

const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v));

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  return e;
}

function barButton(icon: string, label: string, action: () => void): HTMLButtonElement {
  const b = el('button', 'oe-pic-tool');
  b.innerHTML = `<svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><path fill="currentColor" d="${icon}"/></svg><span>${label}</span>`;
  b.setAttribute('aria-label', label);
  b.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    e.stopPropagation();
  });
  b.onclick = (e) => {
    e.stopPropagation();
    action();
  };
  return b;
}

// Material Symbols paths (Apache-2.0): crop, delete, check.
const ICON_CROP = 'M17 15h2V7c0-1.1-.9-2-2-2H9v2h8v8zM7 17V1H5v4H1v2h4v10c0 1.1.9 2 2 2h10v4h2v-4h4v-2H7z';
const ICON_DELETE = 'M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z';
const ICON_DONE = 'M9 16.17 4.83 12l-1.42 1.42L9 19 21 7l-1.41-1.41z';
