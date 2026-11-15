import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { t } from '../i18n';
import { isDirty, registerEditor, setDirty, type SaveResult } from '../save';
import { fontSizeField, holdable, parseFontSize } from '../ui';
import { css, editSvg, type ColourTarget, type Edit, type FormValue, type PageSpace, type Rgb } from './edits';
import { overlay } from './overlay';
import { along, erasePart, eraseWhole, type Point } from './erase';
import { openStylePanel, readStyle, saveStyle, type StyleTarget } from './style';

// PDF editing (PRD M5): highlight and underline selected text, draw, add text, sign, rotate or
// delete pages, and fill in forms. Edits stay in memory, drawn over the pages, until the user
// saves; write.ts then puts them into the file.

type Tool = 'none' | 'ink' | 'text' | 'place-signature' | 'erase' | 'pick-colour';

interface PageState {
  el: HTMLElement;
  space: PageSpace;
  overlay: SVGSVGElement;
  badge?: HTMLElement;
}

type Op =
  | { kind: 'edit'; edit: Edit }
  /** One swipe of an eraser: each edit it touched, and what that edit became. */
  | { kind: 'erase'; page: number; changes: Map<Edit, Edit[]> }
  | { kind: 'rotate'; page: number }
  | { kind: 'delete'; page: number };

export interface EditorHost {
  pdf: PDFDocumentProxy;
  fileName: string;
  /** Size of the file in bytes, or -1. */
  size: number;
  currentPage(): number;
  /** Shows the saved file in place of the one on screen. */
  reload(): void;
  /** Leaves these annotations (PDF.js ids, "12R") off the rendered pages: the editor draws them. */
  hideAnnotations(list: { page: number; id: string }[]): void;
}

const SIGNATURE_KEY = 'officeeditor.signature';
/** Larger files are not edited: saving rewrites the whole file in memory. */
const MAX_EDIT_BYTES = 150 * 1024 * 1024;
const SIGNATURE_WIDTH = 150;
/** A mark begun less than this before a second finger lands is the start of a pinch. */
const PINCH_GRACE_MS = 300;
/** How far (CSS px) a finger may move and still be a tap. */
const TAP_SLOP = 12;

export class PdfEditor {
  readonly row = el('div', 'pagebar-row pdf-edit-row');
  private pages = new Map<number, PageState>();
  /**
   * The marks the page already had when the edits in `ops` began: the app's own annotations read
   * from the file (marks.ts), and after a save that didn't reload, the marks then saved.
   */
  private base: Edit[] = [];
  /** Base marks that are annotations in the file on screen, by their object reference. */
  private savedRefs = new Map<Edit, { num: number; gen: number }>();
  private marksRead = false;
  private ops: Op[] = [];
  private undone: Op[] = [];
  private forms = new Map<string, FormValue>();
  private formInputs = new Map<string, HTMLElement[]>();
  private tool: Tool = 'none';
  private toolButtons = new Map<Tool, HTMLButtonElement>();
  /** How new marks look (colours, opacity, pen thickness, eraser); kept on the device. */
  private style = readStyle();
  private eraserButton: HTMLButtonElement;
  private markButtons = new Map<'highlight' | 'underline', HTMLButtonElement>();
  /** The open settings panel: whose it is, and how to close it. */
  private panel: { target: StyleTarget; close: () => void } | null = null;
  /** While the eyedropper waits for a tap: whose colour it sets, and the tool to go back to. */
  private picking: { target: ColourTarget; tool: Tool } | null = null;
  /** Fingers on the pages while a tool is on; with two or more, the tools stand aside. */
  private fingers = new Set<number>();
  private twoFingers = false;
  /** The stroke or eraser swipe in progress, so a second finger (on any page) can end it. */
  private gesture: { page: number; started: number; end: (keep: boolean) => void } | null = null;
  private signature: { png: Uint8Array; url: string; ratio: number } | null = null;
  private undoButton: HTMLButtonElement;
  private redoButton: HTMLButtonElement;
  private toast = el('div', 'pdf-toast');
  private toastTimer = 0;
  active = false;

  constructor(private host: EditorHost, private hasForms: boolean, onDone: () => void) {
    /** A button that first closes a tool's settings panel, if one is open. */
    const button = (icon: string, label: string, action: () => void) => {
      const b = iconButton(icon, label);
      b.onclick = () => {
        this.panel?.close();
        action();
      };
      this.row.append(b);
      return b;
    };
    // A tap uses the tool (or turns it on and off); holding the button opens that tool's own
    // settings: colour, opacity, thickness, eraser mode and size.
    const markButton = (kind: 'highlight' | 'underline', icon: string) => {
      const b = iconButton(icon, t(kind));
      holdable(b, () => this.markTap(kind), () => this.hold(kind));
      this.row.append(b);
      this.markButtons.set(kind, b);
    };
    markButton('highlight', ICON_HIGHLIGHT);
    markButton('underline', ICON_UNDERLINE);
    const toolButton = (tool: 'ink' | 'text' | 'erase', icon: string, label: string) => {
      const b = iconButton(icon, label);
      holdable(b, () => this.toolTap(tool), () => this.hold(tool === 'erase' ? 'eraser' : tool));
      this.row.append(b);
      return b;
    };
    this.toolButtons.set('ink', toolButton('ink', ICON_PEN, t('pen')));
    this.toolButtons.set('text', toolButton('text', ICON_TEXT, t('addText')));
    this.eraserButton = toolButton('erase', '', t('eraser'));
    this.drawEraserButton();
    this.toolButtons.set('place-signature', button(ICON_SIGN, t('sign'), () => void this.chooseSignature()));
    button(ICON_PAGE, t('pageTools'), () => this.pageMenu());
    this.undoButton = button(ICON_UNDO, t('undo'), () => this.undo());
    this.redoButton = button(ICON_REDO, t('redo'), () => this.redo());
    button(ICON_DONE, t('done'), () => {
      this.setTool('none');
      this.active = false;
      onDone();
    });
    this.row.hidden = true;
    // A finger may lift away from the page it touched (only the first one is captured).
    for (const type of ['pointerup', 'pointercancel'] as const) {
      window.addEventListener(type, (e) => {
        this.fingers.delete(e.pointerId);
        if (this.fingers.size === 0) this.twoFingers = false;
      }, true);
    }
    overlay().append(this.toast);
    this.updateButtons();

    registerEditor({
      fileName: host.fileName,
      produce: () => this.produce(),
      saved: (result) => this.saved(result),
    });
  }

  /** Enters edit mode; false (with a message) if this file can't be edited. */
  start(): boolean {
    if (this.host.size > MAX_EDIT_BYTES) {
      this.say(t('pdfTooLarge'));
      return false;
    }
    this.active = true;
    this.say(t('editHint'));
    void this.readSavedMarks();
    return true;
  }

  /**
   * Makes the marks this app saved earlier editable again: they join the base list, drawn by the
   * editor, and PDF.js stops drawing them into the page.
   */
  private async readSavedMarks(): Promise<void> {
    if (this.marksRead) return;
    this.marksRead = true;
    try {
      const { readMarks, refId } = await import('./marks');
      const marks = await readMarks(await this.host.pdf.getData());
      if (!marks.length) return;
      for (const { edit, ref } of marks) {
        this.base.push(edit);
        this.savedRefs.set(edit, ref);
      }
      this.host.hideAnnotations(marks.map(({ edit, ref }) => ({ page: edit.page, id: refId(ref) })));
      for (const page of new Set(marks.map((m) => m.edit.page))) this.drawPage(page);
    } catch (e) {
      console.warn('saved marks not read', e);
    }
  }

  // ---------------------------------------------------------------- pages

  /** Called by the viewer once a page has been drawn the first time. */
  async pageReady(index: number, pageEl: HTMLElement, proxy: PDFPageProxy): Promise<void> {
    if (this.pages.has(index)) return;
    const viewport = proxy.getViewport({ scale: 1 });
    const space: PageSpace = {
      width: viewport.width,
      height: viewport.height,
      rotation: viewport.rotation,
      toView: (x, y) => viewport.convertToViewportPoint(x, y) as [number, number],
      toPdf: (x, y) => viewport.convertToPdfPoint(x, y) as [number, number],
    };
    const overlay = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    overlay.classList.add('pdf-edit-layer');
    overlay.setAttribute('viewBox', `0 0 ${viewport.width} ${viewport.height}`);
    overlay.setAttribute('preserveAspectRatio', 'none');
    pageEl.append(overlay);
    const state: PageState = { el: pageEl, space, overlay };
    this.pages.set(index, state);
    this.listenDrawing(index, state);
    this.drawPage(index);
    this.drawBadge(index);
    if (this.hasForms) await this.addFormFields(state, proxy);
  }

  /** The edits as they now stand (added, less what was erased), in drawing order. */
  private edits(page?: number): Edit[] {
    let list: Edit[] = [...this.base];
    for (const op of this.ops) {
      if (op.kind === 'edit') list.push(op.edit);
      else if (op.kind === 'erase') list = list.flatMap((e) => op.changes.get(e) ?? [e]);
    }
    return page === undefined ? list : list.filter((e) => e.page === page);
  }

  private drawPage(index: number, edits = this.edits(index)): void {
    const state = this.pages.get(index);
    if (!state) return;
    state.overlay.innerHTML = edits.map((edit) => editSvg(edit, state.space)).join('');
  }

  /** Rotations and deletions are applied when saving; until then a label says so. */
  private drawBadge(index: number): void {
    const state = this.pages.get(index);
    if (!state) return;
    const { rotate, deleted } = this.pageOps();
    state.el.classList.toggle('pdf-page-deleted', deleted.has(index));
    const turn = rotate.get(index) ?? 0;
    const text = deleted.has(index) ? t('willDelete') : turn ? t('willRotate', { deg: turn }) : '';
    if (!text) {
      state.badge?.remove();
      state.badge = undefined;
      return;
    }
    state.badge ??= el('div', 'pdf-page-badge');
    state.badge.textContent = text;
    state.el.append(state.badge);
  }

  private pageOps(): { rotate: Map<number, number>; deleted: Set<number> } {
    const rotate = new Map<number, number>();
    const deleted = new Set<number>();
    for (const op of this.ops) {
      if (op.kind === 'rotate') rotate.set(op.page, ((rotate.get(op.page) ?? 0) + 90) % 360);
      else if (op.kind === 'delete') deleted.add(op.page);
    }
    for (const [page, turn] of rotate) if (!turn) rotate.delete(page);
    return { rotate, deleted };
  }

  // ---------------------------------------------------------------- operations and undo

  private push(op: Op): void {
    this.ops.push(op);
    this.undone = [];
    this.changed(op);
  }

  private undo(): void {
    const op = this.ops.pop();
    if (!op) return;
    this.undone.push(op);
    this.changed(op);
  }

  private redo(): void {
    const op = this.undone.pop();
    if (!op) return;
    this.ops.push(op);
    this.changed(op);
  }

  private changed(op: Op): void {
    const page = op.kind === 'edit' ? op.edit.page : op.page;
    this.drawPage(page);
    this.drawBadge(page);
    this.updateButtons();
    setDirty(this.ops.length > 0 || this.forms.size > 0);
  }

  private updateButtons(): void {
    this.undoButton.disabled = !this.ops.length;
    this.redoButton.disabled = !this.undone.length;
  }

  // ---------------------------------------------------------------- tools

  private setTool(tool: Tool, hint = true): void {
    if (this.tool === 'pick-colour' && tool !== 'pick-colour') this.picking = null;
    this.tool = tool;
    for (const [name, b] of this.toolButtons) b.classList.toggle('active', name === tool);
    this.drawEraserButton();
    // Drawing and placing take single-finger touches on the page instead of scrolling.
    document.body.classList.toggle('pdf-drawing', tool !== 'none');
    if (!hint) return;
    if (tool === 'ink') this.say(t('penHint'));
    else if (tool === 'text') this.say(t('textHint'));
    else if (tool === 'place-signature') this.say(t('placeHint'));
    else if (tool === 'erase') this.say(this.style.eraser.mode === 'part' ? t('erasePartHint') : t('eraseWholeHint'));
    else if (tool === 'pick-colour') this.say(t('pickHint'));
  }

  /** Highlight / Underline: marks the selected text. */
  private markTap(kind: 'highlight' | 'underline'): void {
    this.panel?.close();
    this.markSelection(kind);
  }

  /** Draw / Add text / eraser: on, or off when on. */
  private toolTap(tool: 'ink' | 'text' | 'erase'): void {
    this.panel?.close();
    this.setTool(this.tool === tool ? 'none' : tool);
  }

  /** A held tool button: that tool's settings; a drawing tool is turned on, ready to use them. */
  private hold(target: StyleTarget): void {
    this.panel?.close();
    const tool: Partial<Record<StyleTarget, Tool>> = { ink: 'ink', text: 'text', eraser: 'erase' };
    const on = tool[target];
    if (on && this.tool !== on) {
      // Without the hint toast, which the panel would cover.
      this.setTool(on, false);
    }
    this.showStyle(target);
  }

  private markSelection(kind: 'highlight' | 'underline'): void {
    const selection = getSelection();
    if (!selection || selection.isCollapsed) {
      this.say(t('selectTextFirst'));
      return;
    }
    const byPage = new Map<number, number[][]>();
    for (let r = 0; r < selection.rangeCount; r++) {
      for (const rect of lineRects(selection.getRangeAt(r))) {
        for (const [index, state] of this.pages) {
          const box = state.el.getBoundingClientRect();
          const cx = rect.left + rect.width / 2;
          const cy = rect.top + rect.height / 2;
          if (cx < box.left || cx > box.right || cy < box.top || cy > box.bottom) continue;
          const k = state.space.width / box.width;
          const x0 = (rect.left - box.left) * k;
          const x1 = (rect.right - box.left) * k;
          const y0 = (rect.top - box.top) * k;
          const y1 = (rect.bottom - box.top) * k;
          // Top-left, top-right, bottom-left, bottom-right, as screen corners, in PDF space.
          const q = [[x0, y0], [x1, y0], [x0, y1], [x1, y1]].flatMap(([x, y]) => state.space.toPdf(x, y));
          (byPage.get(index) ?? byPage.set(index, []).get(index)!).push(q);
        }
      }
    }
    if (!byPage.size) {
      this.say(t('selectTextFirst'));
      return;
    }
    for (const [page, quads] of byPage) this.push({ kind: 'edit', edit: { kind, page, quads, colour: this.style[kind].colour, opacity: this.style[kind].opacity } });
    selection.removeAllRanges();
  }

  private drawEraserButton(): void {
    const part = this.style.eraser.mode === 'part';
    const b = this.eraserButton;
    b.classList.toggle('active', this.tool === 'erase');
    const label = part ? t('eraserPart') : t('eraser');
    b.setAttribute('aria-label', label);
    b.title = label;
    // The partial eraser: a smaller eraser leaving a dotted trail.
    b.innerHTML = `<svg viewBox="0 0 24 24" width="24" height="24" aria-hidden="true">${part
      ? `<g transform="translate(6 0) scale(0.75)"><path fill="currentColor" d="${ICON_ERASER}"/></g><path fill="currentColor" d="${ICON_DOTS}"/>`
      : `<path fill="currentColor" d="${ICON_ERASER}"/>`}</svg>`;
  }

  private listenDrawing(index: number, state: PageState): void {
    let erasing: ReturnType<PdfEditor['startErasing']> | null = null;
    let path: number[] | null = null;
    let live: SVGPolylineElement | null = null;
    /** A tap tool (text, signature, eyedropper) acts when the finger lifts, if it was a tap. */
    let tap: { id: number; x: number; y: number; vx: number; vy: number } | null = null;
    const local = (e: PointerEvent): [number, number] => {
      const box = state.overlay.getBoundingClientRect();
      return [((e.clientX - box.left) / box.width) * state.space.width, ((e.clientY - box.top) / box.height) * state.space.height];
    };
    /** Ends this page's stroke or eraser swipe: kept, or (keep = false) dropped as if never made. */
    const end = (keep: boolean) => {
      tap = null;
      if (erasing) {
        if (keep) erasing.end();
        else erasing.cancel();
        erasing = null;
      }
      if (path) {
        const done = path;
        path = null;
        live = null;
        if (keep) {
          this.push({ kind: 'edit', edit: { kind: 'ink', page: index, paths: [done], colour: this.style.ink.colour, width: this.style.ink.width, opacity: this.style.ink.opacity } });
        } else this.drawPage(index);
      }
      if (this.gesture?.page === index) this.gesture = null;
    };
    state.overlay.addEventListener('pointerdown', (e) => {
      if (this.tool === 'none') return;
      this.fingers.add(e.pointerId);
      if (this.fingers.size > 1) {
        // A second finger: zooming or scrolling the page, not using the tool. A mark the first
        // finger only just began is the start of that pinch; an older one is kept as drawn.
        const g = this.gesture;
        if (g) g.end(performance.now() - g.started > PINCH_GRACE_MS);
        this.twoFingers = true;
        return;
      }
      if (this.twoFingers) return;
      e.preventDefault();
      const [vx, vy] = local(e);
      if (this.tool === 'text' || this.tool === 'place-signature' || this.tool === 'pick-colour') {
        tap = { id: e.pointerId, x: e.clientX, y: e.clientY, vx, vy };
        return;
      }
      capture(state.overlay, e.pointerId);
      this.gesture = { page: index, started: performance.now(), end };
      if (this.tool === 'erase') {
        erasing = this.startErasing(index, state, this.style.eraser.mode === 'part');
        erasing.move(vx, vy);
        return;
      }
      path = [...state.space.toPdf(vx, vy)];
      live = document.createElementNS('http://www.w3.org/2000/svg', 'polyline');
      live.setAttribute('fill', 'none');
      live.setAttribute('stroke', css(this.style.ink.colour));
      live.setAttribute('stroke-width', String(this.style.ink.width));
      live.setAttribute('stroke-linejoin', 'round');
      if (this.style.ink.opacity < 1) live.setAttribute('opacity', String(this.style.ink.opacity));
      live.setAttribute('stroke-linecap', 'round');
      live.setAttribute('points', `${vx},${vy}`);
      state.overlay.append(live);
    });
    state.overlay.addEventListener('pointermove', (e) => {
      if (this.twoFingers) return;
      if (erasing) {
        const coalesced = e.getCoalescedEvents?.() ?? [];
        for (const ev of coalesced.length ? coalesced : [e]) erasing.move(...local(ev));
        return;
      }
      if (!path || !live) return;
      const coalesced = e.getCoalescedEvents?.() ?? [];
      for (const ev of coalesced.length ? coalesced : [e]) {
        const [vx, vy] = local(ev);
        path.push(...state.space.toPdf(vx, vy));
        live.setAttribute('points', `${live.getAttribute('points')} ${vx},${vy}`);
      }
    });
    const lifted = (e: PointerEvent, cancelled: boolean) => {
      this.fingers.delete(e.pointerId);
      if (this.fingers.size === 0) this.twoFingers = false;
      const t = tap;
      if (t && t.id === e.pointerId && !cancelled && Math.hypot(e.clientX - t.x, e.clientY - t.y) < TAP_SLOP) {
        tap = null;
        if (this.tool === 'text') void this.addText(index, state, t.vx, t.vy);
        else if (this.tool === 'place-signature') this.placeSignature(index, state, t.vx, t.vy);
        else if (this.tool === 'pick-colour') this.pickColour(state, e);
        return;
      }
      // The browser cancels the pointer when it takes over (a pinch): drop what it was drawing.
      end(!cancelled);
    };
    state.overlay.addEventListener('pointerup', (e) => lifted(e, false));
    state.overlay.addEventListener('pointercancel', (e) => lifted(e, true));
  }

  /**
   * One swipe of an eraser on a page. The page shows the result as it goes; the swipe becomes
   * one operation (one undo step) when the finger lifts.
   */
  private startErasing(index: number, state: PageState, part: boolean) {
    const base = this.edits(index);
    const changes = new Map<Edit, Edit[]>();
    /** For an edit the swipe produced, the original it came from. */
    const origin = new Map<Edit, Edit>();
    let last: Point | null = null;
    // Radius in page units, so it is the same size on screen at any zoom.
    const r = this.style.eraser.size * (state.space.width / Math.max(1, state.overlay.getBoundingClientRect().width));
    const current = () => base.flatMap((e) => changes.get(e) ?? [e]);
    const cursor = (vx: number, vy: number) =>
      `<circle cx="${vx.toFixed(2)}" cy="${vy.toFixed(2)}" r="${r.toFixed(2)}" fill="rgb(128 128 128 / 0.15)" stroke="rgb(90 90 90)" stroke-width="${(r / 12).toFixed(2)}"/>`;
    return {
      move: (vx: number, vy: number) => {
        const p = state.space.toPdf(vx, vy);
        const points = along(last ?? p, p, r / 2);
        last = p;
        for (const edit of current()) {
          const result = part ? erasePart(edit, points, r) : eraseWhole(edit, points, r, state.space.rotation);
          if (!result) continue;
          const o = origin.get(edit) ?? edit;
          const now = changes.get(o) ?? [o];
          now.splice(now.indexOf(edit), 1, ...result);
          changes.set(o, now);
          for (const e of result) origin.set(e, o);
        }
        this.drawPage(index, current());
        state.overlay.insertAdjacentHTML('beforeend', cursor(vx, vy));
      },
      end: () => {
        if (changes.size) this.push({ kind: 'erase', page: index, changes });
        else this.drawPage(index);
      },
      /** Drops the swipe: nothing it touched is erased. */
      cancel: () => this.drawPage(index),
    };
  }

  private async addText(index: number, state: PageState, vx: number, vy: number): Promise<void> {
    const result = await textDialog();
    if (!result) return;
    const [x, y] = state.space.toPdf(vx, vy);
    this.push({ kind: 'edit', edit: { kind: 'text', page: index, x, y, text: result.text, size: result.size, colour: this.style.text.colour } });
    this.setTool('none');
  }

  private showStyle(target: StyleTarget): void {
    this.markButtons.forEach((b, kind) => b.classList.toggle('active', kind === target));
    const close = openStylePanel({
      style: this.style,
      target,
      changed: () => {
        saveStyle(this.style);
        this.drawEraserButton();
      },
      pick: (colourTarget) => {
        this.picking = { target: colourTarget, tool: this.tool };
        this.setTool('pick-colour');
      },
      closed: () => {
        this.panel = null;
        this.markButtons.forEach((b) => b.classList.remove('active'));
      },
    });
    this.panel = { target, close };
  }

  /** The eyedropper's tap: the colour of the page under the finger, then back to the settings. */
  private pickColour(state: PageState, e: PointerEvent): void {
    const picking = this.picking;
    const canvas = state.el.querySelector('canvas');
    let colour: Rgb | null = null;
    if (canvas) {
      const box = canvas.getBoundingClientRect();
      const x = Math.floor(((e.clientX - box.left) / box.width) * canvas.width);
      const y = Math.floor(((e.clientY - box.top) / box.height) * canvas.height);
      try {
        const d = canvas.getContext('2d')?.getImageData(x, y, 1, 1).data;
        // Transparent canvas pixels are white paper.
        if (d) colour = d[3] ? [d[0] / 255, d[1] / 255, d[2] / 255] : [1, 1, 1];
      } catch {
        // Unreadable canvas: nothing picked.
      }
    }
    this.setTool(picking?.tool ?? 'none');
    if (!picking) return;
    if (colour) {
      this.style[picking.target].colour = colour.map((v) => Math.round(v * 255) / 255) as Rgb;
      saveStyle(this.style);
    }
    this.showStyle(picking.target);
  }

  private async chooseSignature(): Promise<void> {
    if (this.tool === 'place-signature') {
      this.setTool('none');
      return;
    }
    const url = await signatureDialog(readStored());
    if (!url) return;
    try {
      localStorage.setItem(SIGNATURE_KEY, url);
    } catch {
      // Not kept for next time; still usable now.
    }
    const png = Uint8Array.from(atob(url.split(',')[1]), (c) => c.charCodeAt(0));
    const img = new Image();
    img.src = url;
    await img.decode().catch(() => undefined);
    this.signature = { png, url, ratio: img.naturalHeight / Math.max(1, img.naturalWidth) || 0.4 };
    this.setTool('place-signature');
  }

  private placeSignature(index: number, state: PageState, vx: number, vy: number): void {
    const sig = this.signature;
    if (!sig) return;
    const w = Math.min(SIGNATURE_WIDTH, state.space.width * 0.6);
    const h = w * sig.ratio;
    // Centred on the tap; store the PDF-space corner nearest the origin.
    const a = state.space.toPdf(vx - w / 2, vy - h / 2);
    const b = state.space.toPdf(vx + w / 2, vy + h / 2);
    this.push({
      kind: 'edit',
      edit: { kind: 'image', page: index, x: Math.min(a[0], b[0]), y: Math.min(a[1], b[1]), w, h, png: sig.png, url: sig.url },
    });
    this.setTool('none');
  }

  private pageMenu(): void {
    const page = this.host.currentPage();
    const { deleted } = this.pageOps();
    void menuDialog(t('pageN', { n: page + 1 }), [
      { label: t('rotatePage'), action: () => this.push({ kind: 'rotate', page }) },
      deleted.has(page)
        ? { label: t('keepPage'), action: () => this.dropDelete(page) }
        : {
          label: t('deletePage'),
          action: () => {
            if (deleted.size + 1 >= this.host.pdf.numPages) this.say(t('lastPage'));
            else this.push({ kind: 'delete', page });
          },
        },
    ]);
  }

  private dropDelete(page: number): void {
    this.ops = this.ops.filter((op) => !(op.kind === 'delete' && op.page === page));
    this.changed({ kind: 'delete', page });
  }

  // ---------------------------------------------------------------- forms

  private async addFormFields(state: PageState, proxy: PDFPageProxy): Promise<void> {
    const annotations = await proxy.getAnnotations({ intent: 'display' });
    const layer = el('div', 'pdf-form-layer');
    const viewport = proxy.getViewport({ scale: 1 });
    for (const a of annotations as FieldAnnotation[]) {
      if (a.subtype !== 'Widget' || !a.fieldName || a.hidden) continue;
      const input = this.fieldInput(a);
      if (!input) continue;
      const [x1, y1] = viewport.convertToViewportPoint(a.rect[0], a.rect[1]);
      const [x2, y2] = viewport.convertToViewportPoint(a.rect[2], a.rect[3]);
      const left = Math.min(x1, x2);
      const top = Math.min(y1, y2);
      const w = Math.abs(x2 - x1);
      const h = Math.abs(y2 - y1);
      // On a turned page the field turns with it: size it as the field is, then rotate.
      const turned = viewport.rotation % 180 !== 0;
      const fw = turned ? h : w;
      const fh = turned ? w : h;
      input.style.left = `${((left + w / 2 - fw / 2) / viewport.width) * 100}%`;
      input.style.top = `${((top + h / 2 - fh / 2) / viewport.height) * 100}%`;
      input.style.width = `${(fw / viewport.width) * 100}%`;
      input.style.height = `${(fh / viewport.height) * 100}%`;
      if (viewport.rotation) input.style.transform = `rotate(${viewport.rotation}deg)`;
      // 0 means "fit to the box"; some files give sizes no line could fit, so check.
      let size = a.defaultAppearanceData?.fontSize ?? 0;
      if (!size || size > fh * 0.9 || (a.multiLine && size > 14)) size = Math.min(12, fh * 0.7);
      input.style.fontSize = `calc(var(--total-scale-factor) * ${size}px)`;
      input.title = a.alternativeText || a.fieldName;
      if (a.readOnly) input.setAttribute('disabled', '');
      layer.append(input);
      const list = this.formInputs.get(a.fieldName) ?? [];
      list.push(input);
      this.formInputs.set(a.fieldName, list);
    }
    if (layer.childElementCount) state.el.append(layer);
  }

  private fieldInput(a: FieldAnnotation): HTMLElement | null {
    const name = a.fieldName!;
    const known = this.forms.get(name);
    const set = (value: FormValue) => {
      this.forms.set(name, value);
      setDirty(true);
      // Other widgets of the same field show the same value.
      for (const other of this.formInputs.get(name) ?? []) {
        if (other instanceof HTMLInputElement && other.type === 'checkbox') other.checked = value === true;
        else if (other instanceof HTMLInputElement && other.type === 'radio') other.checked = other.value === value;
        else if (other instanceof HTMLInputElement || other instanceof HTMLTextAreaElement) {
          if (other.value !== value) other.value = String(value);
        }
      }
    };
    if (a.fieldType === 'Tx') {
      const input = a.multiLine ? el('textarea', 'pdf-field') : el('input', 'pdf-field');
      input.value = String(known ?? a.fieldValue ?? '');
      if (a.maxLen) input.maxLength = a.maxLen;
      input.oninput = () => set(input.value);
      return input;
    }
    if (a.fieldType === 'Btn' && (a.checkBox || a.radioButton)) {
      const input = el('input', 'pdf-field pdf-check');
      if (a.checkBox) {
        input.type = 'checkbox';
        input.checked = known !== undefined ? known === true : !!a.fieldValue && a.fieldValue !== 'Off';
        input.onchange = () => set(input.checked);
      } else {
        input.type = 'radio';
        input.name = `field:${name}`;
        input.value = a.buttonValue ?? '';
        input.checked = (known ?? a.fieldValue) === a.buttonValue;
        input.onchange = () => input.checked && set(input.value);
      }
      return input;
    }
    if (a.fieldType === 'Ch') {
      const select = el('select', 'pdf-field');
      select.multiple = !a.combo && !!a.multiSelect;
      const current = known ?? a.fieldValue;
      const values = Array.isArray(current) ? current : [current];
      for (const option of a.options ?? []) {
        const o = el('option', '');
        o.value = option.exportValue;
        o.textContent = option.displayValue;
        o.selected = values.includes(option.exportValue);
        select.append(o);
      }
      select.onchange = () => set(select.multiple ? [...select.selectedOptions].map((o) => o.value) : select.value);
      return select;
    }
    return null;
  }

  // ---------------------------------------------------------------- saving

  private async produce(): Promise<Uint8Array> {
    const { writeEdits, EncryptedPdfError } = await import('./write');
    const edits = this.edits();
    // Saved marks still as they were stay in the file untouched; erased (or partly erased) ones
    // come out, and what is left of them is written again like a new mark.
    const now = new Set(edits);
    const removed = [...this.savedRefs].filter(([edit]) => !now.has(edit)).map(([, ref]) => ref);
    try {
      const original = await this.host.pdf.getData();
      return await writeEdits(original, edits.filter((e) => !this.savedRefs.has(e)), this.forms, this.pageOps(), removed);
    } catch (e) {
      this.say(e instanceof EncryptedPdfError ? t('pdfProtected') : t('saveFailed'));
      throw e;
    }
  }

  private saved(result: SaveResult): void {
    if (!result.ok) return;
    // The page still shows the file as it was opened, which the next save starts from again.
    this.base = this.edits();
    this.ops = [];
    this.undone = [];
    this.forms.clear();
    // Show the file as saved, edits now part of it.
    if (result.mode === 'overwrite') this.host.reload();
  }

  get hasChanges(): boolean {
    return isDirty();
  }

  private say(text: string): void {
    this.toast.textContent = text;
    this.toast.classList.add('show');
    clearTimeout(this.toastTimer);
    this.toastTimer = window.setTimeout(() => this.toast.classList.remove('show'), 3500);
  }
}

interface FieldAnnotation {
  subtype: string;
  fieldName?: string;
  fieldType?: string;
  fieldValue?: string | string[];
  rect: number[];
  hidden?: boolean;
  readOnly?: boolean;
  multiLine?: boolean;
  maxLen?: number;
  checkBox?: boolean;
  radioButton?: boolean;
  buttonValue?: string;
  combo?: boolean;
  multiSelect?: boolean;
  options?: { exportValue: string; displayValue: string }[];
  alternativeText?: string;
  defaultAppearanceData?: { fontSize?: number };
}

/** The selection's rectangles, one per line of text (the text layer splits lines into spans). */
function lineRects(range: Range): DOMRect[] {
  const rects = [...range.getClientRects()].filter((r) => r.width > 1 && r.height > 1);
  if (!rects.length) return [];
  const heights = rects.map((r) => r.height).sort((a, b) => a - b);
  const typical = heights[Math.floor(heights.length / 2)];
  // Drop boxes of whole elements (a page's text layer, a paragraph) that come with them.
  const lines = rects.filter((r) => r.height < typical * 2.5).sort((a, b) => a.top - b.top || a.left - b.left);
  const out: DOMRect[] = [];
  for (const r of lines) {
    const last = out[out.length - 1];
    if (last && Math.abs(last.top - r.top) < r.height * 0.5 && r.left <= last.right + r.height) {
      out[out.length - 1] = new DOMRect(
        Math.min(last.left, r.left), Math.min(last.top, r.top),
        Math.max(last.right, r.right) - Math.min(last.left, r.left),
        Math.max(last.bottom, r.bottom) - Math.min(last.top, r.top),
      );
    } else out.push(r);
  }
  return out;
}

/** Keeps a drag's events coming to `el` even when the finger leaves it. */
function capture(el: Element, pointerId: number): void {
  try {
    el.setPointerCapture(pointerId);
  } catch {
    // Not a live pointer (synthetic event); the drag still works while over the element.
  }
}

function readStored(): string | null {
  try {
    return localStorage.getItem(SIGNATURE_KEY);
  } catch {
    return null;
  }
}

// ------------------------------------------------------------------ dialogs

function dialog(title: string): { box: HTMLElement; close: () => void; actions: HTMLElement } {
  const scrim = el('div', 'pdf-dialog-scrim');
  const box = el('div', 'pdf-dialog');
  box.setAttribute('role', 'dialog');
  const h = el('h2', '');
  h.textContent = title;
  const actions = el('div', 'pdf-dialog-actions');
  box.append(h, actions);
  scrim.append(box);
  overlay().append(scrim);
  return { box, actions, close: () => scrim.remove() };
}

function textButton(label: string, onClick: () => void, primary = false): HTMLButtonElement {
  const b = el('button', primary ? 'pdf-text-button primary' : 'pdf-text-button');
  b.textContent = label;
  b.onclick = onClick;
  return b;
}

/** The size the last text box was given, for the next one. */
let lastTextSize = 14;

function textDialog(): Promise<{ text: string; size: number } | null> {
  return new Promise((resolve) => {
    const d = dialog(t('addText'));
    const area = el('textarea', 'pdf-dialog-text');
    area.rows = 3;
    // Font size in points, typed in (PDF units are points, so it is the size on the page).
    const { field, input, hint } = fontSizeField(lastTextSize);
    field.classList.add('pdf-size-field');
    const caption = el('span', 'pdf-size-caption');
    caption.textContent = t('fontSize');
    field.prepend(caption);
    // A preview in the box, within reason for a phone screen.
    const preview = () => {
      const n = parseFontSize(input.value);
      if (n) area.style.fontSize = `${Math.min(Math.max(n, 10), 32)}px`;
    };
    input.addEventListener('input', preview);
    preview();
    d.box.insertBefore(area, d.actions);
    d.box.insertBefore(field, d.actions);
    d.box.insertBefore(hint, d.actions);
    d.actions.append(
      textButton(t('cancel'), () => { d.close(); resolve(null); }),
      textButton(t('add'), () => {
        const text = area.value.replace(/\s+$/, '');
        const size = parseFontSize(input.value);
        if (text && size === null) {
          hint.classList.add('error');
          input.focus();
          return;
        }
        d.close();
        if (text && size) lastTextSize = size;
        resolve(text && size ? { text, size } : null);
      }, true),
    );
    area.focus();
  });
}

/** Draw a signature (or reuse the saved one). Resolves with a PNG data URL. */
function signatureDialog(saved: string | null): Promise<string | null> {
  return new Promise((resolve) => {
    const d = dialog(t('signTitle'));
    const canvas = el('canvas', 'pdf-sign-pad');
    const ratio = Math.min(devicePixelRatio, 3);
    const W = 320;
    const H = 160;
    canvas.width = W * ratio;
    canvas.height = H * ratio;
    const ctx = canvas.getContext('2d')!;
    ctx.scale(ratio, ratio);
    ctx.lineWidth = 2.5;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = '#102a6b';
    let drawn = false;
    let last: [number, number] | null = null;
    const at = (e: PointerEvent): [number, number] => {
      const r = canvas.getBoundingClientRect();
      return [((e.clientX - r.left) / r.width) * W, ((e.clientY - r.top) / r.height) * H];
    };
    canvas.onpointerdown = (e) => {
      capture(canvas, e.pointerId);
      last = at(e);
    };
    canvas.onpointermove = (e) => {
      if (!last) return;
      const p = at(e);
      ctx.beginPath();
      ctx.moveTo(...last);
      ctx.lineTo(...p);
      ctx.stroke();
      last = p;
      drawn = true;
    };
    canvas.onpointerup = () => (last = null);
    const hint = el('p', 'pdf-dialog-hint');
    hint.textContent = t('signHint');
    d.box.insertBefore(canvas, d.actions);
    d.box.insertBefore(hint, d.actions);
    if (saved) {
      const img = new Image();
      img.onload = () => {
        const k = Math.min(W / img.width, H / img.height, 1) * 0.9;
        ctx.drawImage(img, (W - img.width * k) / 2, (H - img.height * k) / 2, img.width * k, img.height * k);
        drawn = true;
      };
      img.src = saved;
    }
    d.actions.append(
      textButton(t('clear'), () => {
        ctx.clearRect(0, 0, W, H);
        drawn = false;
      }),
      textButton(t('cancel'), () => { d.close(); resolve(null); }),
      textButton(t('useSignature'), () => {
        d.close();
        resolve(drawn ? cropped(canvas) : null);
      }, true),
    );
  });
}

/** The canvas cut to what was drawn, as a PNG data URL with a transparent background. */
function cropped(canvas: HTMLCanvasElement): string {
  const ctx = canvas.getContext('2d')!;
  const { width, height } = canvas;
  const data = ctx.getImageData(0, 0, width, height).data;
  let x0 = width, y0 = height, x1 = 0, y1 = 0;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (data[(y * width + x) * 4 + 3] > 8) {
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        if (y > y1) y1 = y;
      }
    }
  }
  if (x1 < x0) return canvas.toDataURL('image/png');
  const pad = 4;
  x0 = Math.max(0, x0 - pad);
  y0 = Math.max(0, y0 - pad);
  x1 = Math.min(width - 1, x1 + pad);
  y1 = Math.min(height - 1, y1 + pad);
  const out = document.createElement('canvas');
  out.width = x1 - x0 + 1;
  out.height = y1 - y0 + 1;
  out.getContext('2d')!.drawImage(canvas, x0, y0, out.width, out.height, 0, 0, out.width, out.height);
  return out.toDataURL('image/png');
}

function menuDialog(title: string, items: { label: string; action: () => void }[]): Promise<void> {
  return new Promise((resolve) => {
    const d = dialog(title);
    const list = el('div', 'pdf-dialog-menu');
    for (const item of items) {
      list.append(textButton(item.label, () => {
        d.close();
        item.action();
        resolve();
      }));
    }
    d.box.insertBefore(list, d.actions);
    d.actions.append(textButton(t('cancel'), () => { d.close(); resolve(); }));
  });
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  return node;
}

function iconButton(path: string, label: string): HTMLButtonElement {
  const button = el('button', 'pagebar-icon');
  button.setAttribute('aria-label', label);
  button.title = label;
  button.innerHTML = `<svg viewBox="0 0 24 24" width="24" height="24" aria-hidden="true"><path fill="currentColor" d="${path}"/></svg>`;
  return button;
}

// Material Symbols paths (Apache-2.0).
const ICON_HIGHLIGHT = 'M6 14l3 3v5h6v-5l3-3V9H6v5zm5-12h2v3h-2V2zM3.5 5.88l1.41-1.41 2.12 2.12-1.41 1.41L3.5 5.88zm13.46.71 2.12-2.12 1.41 1.41-2.12 2.12-1.41-1.41z';
const ICON_UNDERLINE = 'M12 17c3.31 0 6-2.69 6-6V3h-2.5v8c0 1.93-1.57 3.5-3.5 3.5S8.5 12.93 8.5 11V3H6v8c0 3.31 2.69 6 6 6zm-7 2v2h14v-2H5z';
const ICON_PEN = 'M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04a.996.996 0 0 0 0-1.41l-2.34-2.34a.996.996 0 0 0-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z';
const ICON_TEXT = 'M2.5 4v3h5v12h3V7h5V4h-13zm19 5h-9v3h3v7h3v-7h3V9z';
const ICON_SIGN = 'M4.59 6.89c.7-.71 1.4-1.35 1.71-1.22.5.2 0 1.03-.3 1.52-.25.42-2.86 3.89-2.86 6.31 0 1.28.48 2.34 1.34 2.98.75.56 1.74.73 2.64.46 1.07-.31 1.95-1.4 3.06-2.77 1.21-1.49 2.83-3.44 4.08-3.44 1.63 0 1.65 1.01 1.76 1.79-3.78.64-5.38 3.67-5.38 5.37 0 1.7 1.44 3.09 3.21 3.09 1.63 0 4.29-1.33 4.69-6.1H21v-2.5h-2.47c-.15-1.65-1.09-4.2-4.03-4.2-2.25 0-4.18 1.91-4.94 2.84-.58.73-2.06 2.48-2.29 2.72-.25.3-.68.84-1.11.84-.45 0-.72-.83-.36-1.92.35-1.09 1.4-2.86 1.85-3.52.78-1.14 1.3-1.92 1.3-3.28C8.95 3.69 7.31 3 6.44 3 5.12 3 3.97 4 3.72 4.25c-.36.36-.66.66-.88.93l1.75 1.71zm9.29 11.66c-.31 0-.74-.26-.74-.72 0-.6.73-2.2 2.87-2.76-.3 2.69-1.43 3.48-2.13 3.48z';
const ICON_PAGE = 'M15.55 5.55 11 1v3.07C7.06 4.56 4 7.92 4 12s3.05 7.44 7 7.93v-2.02c-2.84-.48-5-2.94-5-5.91s2.16-5.43 5-5.91V10l4.55-4.45zM19.93 11a7.906 7.906 0 0 0-1.62-3.89l-1.42 1.42c.54.75.88 1.6 1.02 2.47h2.02zM13 17.9v2.02c1.39-.17 2.74-.71 3.9-1.61l-1.44-1.44c-.75.54-1.59.89-2.46 1.03zm3.89-2.42 1.42 1.41c.9-1.16 1.45-2.5 1.62-3.89h-2.02c-.14.87-.48 1.72-1.02 2.48z';
const ICON_UNDO = 'M12.5 8c-2.65 0-5.05.99-6.9 2.6L2 7v9h9l-3.62-3.62c1.39-1.16 3.16-1.88 5.12-1.88 3.54 0 6.55 2.31 7.6 5.5l2.37-.78C21.08 11.03 17.15 8 12.5 8z';
const ICON_REDO = 'M18.4 10.6C16.55 8.99 14.15 8 11.5 8c-4.65 0-8.58 3.03-9.96 7.22L3.9 16c1.05-3.19 4.05-5.5 7.6-5.5 1.95 0 3.73.72 5.12 1.88L13 16h9V7l-3.6 3.6z';
// Eraser: Material Design Icons "eraser" (Apache-2.0).
const ICON_ERASER = 'M16.24 3.56l4.95 4.94c.78.79.78 2.05 0 2.84L12 20.53a4.008 4.008 0 0 1-5.66 0L2.81 17c-.78-.79-.78-2.05 0-2.84l10.6-10.6c.79-.78 2.05-.78 2.83 0M4.22 15.58l3.54 3.53c.78.79 2.04.79 2.83 0l3.53-3.53l-4.95-4.95l-4.95 4.95z';
const ICON_DOTS = 'M2 21.5a1.5 1.5 0 1 0 3 0 1.5 1.5 0 1 0-3 0zm5 0a1.5 1.5 0 1 0 3 0 1.5 1.5 0 1 0-3 0zm5 0a1.5 1.5 0 1 0 3 0 1.5 1.5 0 1 0-3 0z';
const ICON_DONE = 'M9 16.17 4.83 12l-1.42 1.42L9 19 21 7l-1.41-1.41z';
