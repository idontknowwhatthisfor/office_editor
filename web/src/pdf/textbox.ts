import { t } from '../i18n';
import { overlay } from '../overlay';
import { DEFAULT_STYLE, type Edit, type PageSpace, type Rgb } from './edits';
import { openStylePanel } from './style';

// Text boxes the app placed, new or saved earlier (marks.ts reads those back): a tap selects one.
// The selected box gets a frame in the overlay layer (screen coordinates, not zoomed with the
// page): dragging the frame moves the box, its corner handle scales the text (the font size),
// and a bar above it edits the text, changes its colour or deletes it. Tapping the selected box
// again edits its text. Every change is one step (the editor's `replace` operation), drawn live
// while the finger moves.

export type TextEdit = Extract<Edit, { kind: 'text' }>;

export interface TextBoxHost {
  page(index: number): { el: HTMLElement; overlay: SVGSVGElement; space: PageSpace } | undefined;
  /** A page's edits in drawing order; each is one child of that page's overlay. */
  edits(page: number): Edit[];
  /** Draws the page with `from` shown as `to` (nothing is recorded). */
  preview(page: number, from: Edit, to: Edit): void;
  /** Records the change as one undo step; `to` null deletes the box. */
  commit(from: TextEdit, to: TextEdit | null): void;
  /** Asks for new text and size; empty text means delete. */
  editText(edit: TextEdit): Promise<{ text: string; size: number } | null>;
  /** Closes a tool's settings panel, if one is open. */
  closePanels(): void;
}

const MIN_SIZE = 1;
const MAX_SIZE = 400;
/** How far (CSS px) a finger may move and still be a tap. */
const SLOP = 8;
/** Room around the text, in screen px: the frame and the area a tap selects it from. */
const PAD = 6;

export class TextBoxes {
  private layer = el('div', 'pdf-tb-layer');
  private frame = el('div', 'pdf-tb-frame');
  private handle = el('div', 'pdf-tb-handle');
  private bar = el('div', 'pdf-tb-bar');
  private selected: TextEdit | null = null;
  private drag: {
    kind: 'move' | 'size';
    pointer: number;
    x0: number;
    y0: number;
    started: boolean;
    from: TextEdit;
    now: TextEdit;
  } | null = null;
  /** While the colour panel is open, the box it is changing. */
  private colouring: TextEdit | null = null;

  constructor(private host: TextBoxHost) {
    this.frame.append(this.handle);
    this.layer.append(this.frame, this.bar);
    this.bar.append(
      barButton(ICON_EDIT, t('editText'), () => void this.edit()),
      barButton(ICON_COLOUR, t('boxColour'), () => this.colour()),
      barButton(ICON_DELETE, t('deleteText'), () => this.remove()),
    );
    this.frame.addEventListener('pointerdown', (e) => this.down(e, 'move'));
    this.handle.addEventListener('pointerdown', (e) => this.down(e, 'size'));
    for (const target of [this.frame, this.handle]) {
      target.addEventListener('pointermove', (e) => this.move(e));
      target.addEventListener('pointerup', (e) => this.up(e, false));
      target.addEventListener('pointercancel', (e) => this.up(e, true));
    }
    // The frame is on the screen, not the page: follow the box when the page scrolls or zooms.
    const again = () => this.drag?.started || this.place();
    window.addEventListener('scroll', again, { passive: true });
    window.addEventListener('resize', again);
    window.visualViewport?.addEventListener('scroll', again);
    window.visualViewport?.addEventListener('resize', again);
  }

  get active(): boolean {
    return !!this.selected;
  }

  /** The topmost text box under a point on a page (client coordinates), if any. */
  hit(page: number, x: number, y: number): TextEdit | null {
    const state = this.host.page(page);
    if (!state) return null;
    const edits = this.host.edits(page);
    const scale = window.visualViewport?.scale ?? 1;
    const pad = PAD / scale;
    for (let i = edits.length - 1; i >= 0; i--) {
      const edit = edits[i];
      if (edit.kind !== 'text') continue;
      const r = state.overlay.children[i]?.getBoundingClientRect();
      if (r && x > r.left - pad && x < r.right + pad && y > r.top - pad && y < r.bottom + pad) return edit;
    }
    return null;
  }

  select(edit: TextEdit): void {
    this.selected = edit;
    (document.activeElement as HTMLElement | null)?.blur?.();
    getSelection()?.removeAllRanges();
    if (!this.layer.isConnected) overlay().append(this.layer);
    this.place();
  }

  clear(): void {
    this.selected = null;
    this.drag = null;
    this.frame.classList.remove('dragging');
    this.layer.remove();
  }

  /** Undo or redo turned the selected box `from` into `to`: the selection goes with it. */
  follow(from: Edit, to: Edit): void {
    if (this.selected === from && to.kind === 'text') this.selected = to;
  }

  /** The edits changed (a step, undo, redo): keep the frame on the box, or drop it if it went. */
  refresh(): void {
    const sel = this.selected;
    if (!sel || this.colouring) return;
    if (!this.host.edits(sel.page).includes(sel)) this.clear();
    else this.place();
  }

  // ---------------------------------------------------------------- placing the controls

  /** The point the text hangs from, (x, y) of the edit, in client coordinates. */
  private anchor(edit: TextEdit): [number, number] | null {
    const state = this.host.page(edit.page);
    if (!state) return null;
    const box = state.overlay.getBoundingClientRect();
    const [vx, vy] = state.space.toView(edit.x, edit.y);
    const k = box.width / state.space.width;
    return [box.left + vx * k, box.top + vy * k];
  }

  private place(): void {
    const sel = this.drag?.now ?? this.selected;
    if (!sel) return;
    // During a drag the page shows the box as it is being changed, at the same place in the list.
    const state = this.host.page(sel.page);
    const list = this.host.edits(sel.page);
    const i = list.indexOf(this.drag?.from ?? this.colouring ?? sel);
    const node = state?.overlay.children[i];
    const anchor = this.anchor(sel);
    if (!node || !anchor) {
      this.layer.hidden = true;
      return;
    }
    this.layer.hidden = false;
    const c = node.getBoundingClientRect();
    const r = local(new DOMRect(c.left, c.top, c.width, c.height));
    const frame = new DOMRect(r.x - PAD, r.y - PAD, r.width + PAD * 2, r.height + PAD * 2);
    setRect(this.frame, frame);
    // The handle goes on the corner farthest from where the text hangs, which stays put.
    const a =local(new DOMRect(anchor[0], anchor[1], 0, 0));
    const right = Math.abs(frame.right - a.x) >= Math.abs(frame.left - a.x);
    const bottom = Math.abs(frame.bottom - a.y) >= Math.abs(frame.top - a.y);
    this.handle.style.left = right ? '100%' : '0';
    this.handle.style.top = bottom ? '100%' : '0';
    this.placeBar(frame);
  }

  /** The bar centred above the frame, or below it when there is no room above. */
  private placeBar(r: DOMRect): void {
    const width = document.documentElement.clientWidth;
    const w = this.bar.offsetWidth;
    const h = this.bar.offsetHeight;
    const left = Math.min(Math.max(4, r.x + r.width / 2 - w / 2), width - w - 4);
    const above = r.y - h - 12;
    this.bar.style.left = `${left}px`;
    this.bar.style.top = `${above >= 8 ? above : r.bottom + 12}px`;
  }

  // ---------------------------------------------------------------- move and scale

  private down(e: PointerEvent, kind: 'move' | 'size'): void {
    const sel = this.selected;
    if (!sel || !e.isPrimary || this.drag) return;
    e.preventDefault();
    e.stopPropagation();
    try {
      (e.currentTarget as Element).setPointerCapture(e.pointerId);
    } catch {
      // Synthetic event: the drag still works while over the frame.
    }
    this.drag = { kind, pointer: e.pointerId, x0: e.clientX, y0: e.clientY, started: false, from: sel, now: sel };
  }

  private move(e: PointerEvent): void {
    const d = this.drag;
    if (!d || e.pointerId !== d.pointer) return;
    if (!d.started) {
      if (Math.hypot(e.clientX - d.x0, e.clientY - d.y0) < SLOP) return;
      d.started = true;
      this.frame.classList.add('dragging');
      this.bar.hidden = true;
    }
    const from = d.from;
    const state = this.host.page(from.page);
    if (!state) return;
    if (d.kind === 'move') {
      const k = state.space.width / state.overlay.getBoundingClientRect().width;
      const [vx, vy] = state.space.toView(from.x, from.y);
      const [x, y] = state.space.toPdf(vx + (e.clientX - d.x0) * k, vy + (e.clientY - d.y0) * k);
      d.now = { ...from, x, y };
    } else {
      const a = this.anchor(from);
      if (!a) return;
      const ratio = Math.hypot(e.clientX - a[0], e.clientY - a[1]) / Math.max(1, Math.hypot(d.x0 - a[0], d.y0 - a[1]));
      const size = Math.min(MAX_SIZE, Math.max(MIN_SIZE, Math.round(from.size * ratio * 2) / 2));
      d.now = { ...from, size };
    }
    this.host.preview(from.page, from, d.now);
    this.place();
  }

  private up(e: PointerEvent, cancelled: boolean): void {
    const d = this.drag;
    if (!d || e.pointerId !== d.pointer) return;
    this.drag = null;
    this.frame.classList.remove('dragging');
    this.bar.hidden = false;
    if (!d.started) {
      // A tap on the selected box: edit its text.
      if (!cancelled && d.kind === 'move') void this.edit();
      return;
    }
    const changed = d.now.x !== d.from.x || d.now.y !== d.from.y || d.now.size !== d.from.size;
    if (cancelled || !changed) {
      this.host.preview(d.from.page, d.from, d.from);
      this.place();
      return;
    }
    this.selected = d.now;
    this.host.commit(d.from, d.now);
  }

  // ---------------------------------------------------------------- the bar

  private async edit(): Promise<void> {
    const sel = this.selected;
    if (!sel) return;
    this.host.closePanels();
    const result = await this.host.editText(sel);
    if (!result || this.selected !== sel) return;
    if (!result.text) {
      this.remove();
      return;
    }
    if (result.text === sel.text && result.size === sel.size) return;
    const to = { ...sel, text: result.text, size: result.size };
    this.selected = to;
    this.host.commit(sel, to);
  }

  /** The colour panel; the box shows each colour at once, and the change is one step. */
  private colour(): void {
    const from = this.selected;
    if (!from) return;
    this.host.closePanels();
    const style = structuredClone(DEFAULT_STYLE);
    style.text.colour = [...from.colour] as Rgb;
    this.colouring = from;
    const shown = () => ({ ...from, colour: [...style.text.colour] as Rgb });
    openStylePanel({
      style,
      target: 'text',
      title: t('boxColour'),
      changed: () => {
        this.host.preview(from.page, from, shown());
        this.place();
      },
      closed: () => {
        this.colouring = null;
        const to = shown();
        if (to.colour.join() === from.colour.join() || this.selected !== from) {
          this.host.preview(from.page, from, from);
          this.refresh();
          return;
        }
        this.selected = to;
        this.host.commit(from, to);
      },
    });
  }

  private remove(): void {
    const sel = this.selected;
    if (!sel) return;
    this.clear();
    this.host.commit(sel, null);
  }
}

/** Client coordinates (the layout viewport) → the overlay's: the visible area at normal size. */
function local(r: DOMRect): DOMRect {
  const v = window.visualViewport;
  const s = v?.scale ?? 1;
  const x = v?.offsetLeft ?? 0;
  const y = v?.offsetTop ?? 0;
  return new DOMRect((r.left - x) * s, (r.top - y) * s, r.width * s, r.height * s);
}

function setRect(e: HTMLElement, r: DOMRect): void {
  e.style.left = `${r.x}px`;
  e.style.top = `${r.y}px`;
  e.style.width = `${r.width}px`;
  e.style.height = `${r.height}px`;
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  e.className = className;
  return e;
}

function barButton(icon: string, label: string, action: () => void): HTMLButtonElement {
  const b = el('button', 'pdf-tb-tool');
  b.innerHTML = `<svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><path fill="currentColor" d="${icon}"/></svg><span>${label}</span>`;
  b.setAttribute('aria-label', label);
  // Keeps the selection (and doesn't start a drag or a page tap).
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

// Material Symbols paths (Apache-2.0).
const ICON_EDIT = 'M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04a.996.996 0 0 0 0-1.41l-2.34-2.34a.996.996 0 0 0-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z';
const ICON_COLOUR = 'M12 3a9 9 0 0 0 0 18c.83 0 1.5-.67 1.5-1.5 0-.39-.15-.74-.39-1.01-.23-.26-.38-.61-.38-.99 0-.83.67-1.5 1.5-1.5H16c2.76 0 5-2.24 5-5 0-4.42-4.03-8-9-8zm-5.5 9c-.83 0-1.5-.67-1.5-1.5S5.67 9 6.5 9 8 9.67 8 10.5 7.33 12 6.5 12zm3-4C8.67 8 8 7.33 8 6.5S8.67 5 9.5 5s1.5.67 1.5 1.5S10.33 8 9.5 8zm5 0c-.83 0-1.5-.67-1.5-1.5S13.67 5 14.5 5s1.5.67 1.5 1.5S15.33 8 14.5 8zm3 4c-.83 0-1.5-.67-1.5-1.5S16.67 9 17.5 9s1.5.67 1.5 1.5-.67 1.5-1.5 1.5z';
const ICON_DELETE = 'M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z';
