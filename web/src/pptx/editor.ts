import JSZip from 'jszip';
import { t } from '../i18n';
import { child, children } from '../ooxml';
import { registerEditor, setDirty, type SaveResult } from '../save';
import { askFontSize, ICON_FONT_SIZE, toast } from '../ui';
import { Deck } from './deck';
import { editableShapes, type RenderedSlide } from './render';
import { ShapeTool } from './shapes';
import type { TextLayout } from './text';

// PowerPoint text editing (M9), done in place on each slide's own XML.
//
// In edit mode slides are drawn by render.ts with marks: every editable shape's box carries
// `data-sp` (its number among the slide's shapes), every paragraph `data-p` and every run
// `data-r` ("paragraph_child"), so text on screen maps back to XML elements. Paragraphs are
// edited in place (each one is its own contenteditable). Typing only changes the screen; after
// a pause the box's paragraphs are rebuilt into the slide's working XML: unchanged paragraphs
// are copies of their elements, changed ones are rebuilt from their own pPr, runs' rPr and
// endParaRPr. Enter, joining paragraphs, autofit and undo redraw the slide from that XML.
// Only slides whose XML changed are written on save; everything else in the file stays byte
// for byte.
//
// Locked (shown, not editable): paragraphs with fields or links, tables, charts, SmartArt.

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const KEEP = { createFolders: false };
const ZWSP = /​/g;

/** What the editor needs from the viewer. */
export interface EditHost {
  deck: Deck;
  fileName: string;
  /** The element slides are drawn in (event delegation). */
  pager: HTMLElement;
  current(): number;
  go(i: number): void;
  /** Draws slide `i` again from its working XML and puts it on screen. */
  redraw(i: number): Promise<void>;
  /** The slide element on screen for slide `i`, if it is attached. */
  slideEl(i: number): HTMLElement | null;
  refreshThumb(i: number): void;
}

interface SlideState {
  doc: Document;
  /** The XML declaration of the original part (XMLSerializer leaves it out). */
  decl: string;
  /** The root element's start tag as written (XMLSerializer reorders its namespace declarations). */
  rootTag: string;
  initial: string;
  last: string;
}

interface Box {
  slide: number;
  sp: Element;
  /** The shape's paragraphs when the box was drawn; data-p / data-r refer to these. */
  src: Element[];
  layout: TextLayout;
  /** Text changed since the box was drawn (autofit is checked when leaving it). */
  edited: boolean;
  /** Paragraphs were added or joined on screen (the slide is redrawn when the box is left). */
  reshaped: boolean;
}

interface Caret {
  slide: number;
  sp: number;
  para: number;
  offset: number;
}

interface Look {
  b: boolean;
  i: boolean;
  u: boolean;
}

/** normAutofit steps, as PowerPoint shrinks text: font scale and line spacing reduction. */
const SHRINK: [number, number][] = [
  [1, 0], [1, 0.1], [0.925, 0.2], [0.85, 0.2], [0.775, 0.2], [0.7, 0.2], [0.625, 0.2], [0.55, 0.2], [0.475, 0.2], [0.4, 0.2], [0.325, 0.2], [0.25, 0.2],
];

export class PptxEditor {
  readonly bar = el('div', 'pptx-tools');
  private states = new Map<number, SlideState>();
  private boxes = new WeakMap<HTMLElement, Box>();
  /** Each paragraph's markup when drawn, to tell which ones the user changed. */
  private pristine = new WeakMap<HTMLElement, string>();
  private pending = new Set<HTMLElement>();
  private timer = 0;
  private undoStack: { slide: number; xml: string }[] = [];
  private redoStack: { slide: number; xml: string }[] = [];
  /** Typing goes into one undo step until a pause or another kind of change. */
  private group: { slide: number; at: number } | null = null;
  private measureHost: HTMLElement | null = null;
  private composing = false;
  /** A dialog of the toolbar's is open. */
  private inDialog = false;
  private afterComposition: (() => void) | null = null;

  private shapes: ShapeTool;

  constructor(private host: EditHost) {
    this.shapes = new ShapeTool({
      pager: host.pager,
      doc: (i) => this.states.get(i)?.doc ?? null,
      flush: () => this.flush(false),
      changed: async (i) => {
        this.group = null;
        this.commit(i, false);
        await this.redraw(i, null);
      },
    });
    this.buildBar();
    this.listen();
    registerEditor({ fileName: host.fileName, produce: () => this.produce(), saved: (r) => this.saved(r) });
    toast(t('pptxEditHint'));
  }

  // ---------------------------------------------------------------- drawing

  /** Called by the viewer for every slide drawn in edit mode, before it goes on screen. */
  async prepare(i: number, slide: RenderedSlide): Promise<void> {
    const doc = await this.host.deck.slideDoc(i);
    if (!doc) return;
    if (!this.states.has(i)) {
      const raw = (await this.host.deck.pkg.zip.file(this.host.deck.slidePaths[i])?.async('string')) ?? '';
      const xml = serialize(doc);
      const decl = /^<\?xml[^>]*\?>\s*/.exec(raw)?.[0] ?? '';
      const rootTag = /^<[^?!][^>]*>/.exec(raw.slice(decl.length))?.[0] ?? '';
      this.states.set(i, { doc, decl, rootTag, initial: xml, last: xml });
    }
    const shapes = editableShapes(doc.documentElement);
    slide.el.classList.add('editing');
    slide.el.dataset.slide = String(i);
    for (const div of slide.el.querySelectorAll<HTMLElement>('.sp[data-sp]')) {
      const n = Number(div.dataset.sp);
      const sp = shapes[n];
      const layout = slide.texts.get(n);
      if (!sp || !layout) continue;
      const src = children(child(sp, 'txBody'), 'p').map((p) => p.cloneNode(true) as Element);
      this.boxes.set(div, { slide: i, sp, src, layout, edited: false, reshaped: false });
      this.prepareParagraphs(div);
    }
    // Text the editor can't change: SmartArt, tables, text in shapes it doesn't own.
    for (const tx of slide.el.querySelectorAll<HTMLElement>('.tx, .tbl')) {
      if (!tx.closest('.sp[data-sp]')) tx.classList.add('oe-locked');
    }
    for (const box of slide.el.querySelectorAll<HTMLElement>('.chart, .missing')) box.classList.add('oe-locked');
  }

  private prepareParagraphs(div: HTMLElement): void {
    const tx = textLayer(div);
    if (!tx) return;
    div.dataset.prompt = t('tapToAddText');
    for (const p of paragraphs(div)) {
      if (p.dataset.lock) {
        p.classList.add('oe-locked');
        continue;
      }
      p.contentEditable = 'true';
      for (const bu of p.querySelectorAll<HTMLElement>('.bu')) bu.contentEditable = 'false';
      // How each span looked when drawn (the empty paragraph's stand-in span too).
      for (const span of p.querySelectorAll<HTMLElement>(':scope > span:not(.bu)')) span.dataset.look = lookFlags(styleLook(span));
      this.pristine.set(p, p.innerHTML);
    }
    this.markEmpty(div);
  }

  /** An empty placeholder shows a prompt, as in PowerPoint (other empty shapes show none). */
  private markEmpty(div: HTMLElement): void {
    const sp = this.boxes.get(div)?.sp;
    const placeholder = !!sp && sp.getElementsByTagNameNS(P, 'ph').length > 0;
    div.classList.toggle('oe-empty', placeholder && isEmpty(div));
  }

  // ---------------------------------------------------------------- toolbar

  private buildBar(): void {
    const add = (label: string, title: string, action: () => void) => {
      const b = el('button', 'pptx-tool');
      b.innerHTML = label;
      b.title = title;
      b.setAttribute('aria-label', title);
      // Keep the text selection: the button must not take focus.
      b.addEventListener('pointerdown', (e) => e.preventDefault());
      b.onclick = action;
      this.bar.append(b);
    };
    add('<b>B</b>', t('bold'), () => this.format('bold'));
    add('<i>I</i>', t('italic'), () => this.format('italic'));
    add('<u>U</u>', t('underlineText'), () => this.format('underline'));
    add(ICON_FONT_SIZE, t('fontSize'), () => void this.fontSize());
    add(ICON_ALIGN, t('alignment'), () => this.cycleAlign());
    add(ICON_UNDO, t('undo'), () => void this.undo());
    add(ICON_REDO, t('redo'), () => void this.redo());
  }

  /** The editable paragraph holding the caret. */
  private currentPara(): HTMLElement | null {
    const node = getSelection()?.anchorNode;
    const p = (node instanceof Element ? node : node?.parentElement)?.closest<HTMLElement>('p[contenteditable="true"]');
    return p && this.host.pager.contains(p) ? p : null;
  }

  private format(command: 'bold' | 'italic' | 'underline'): void {
    const p = this.currentPara();
    if (!p) return toast(t('tapTextFirst'));
    document.execCommand('styleWithCSS', false, 'false');
    document.execCommand(command);
    this.touched(p, false);
    this.flush(false);
  }

  /**
   * Font size: of the selected text; with only a caret, of its paragraph; with a text box
   * selected (its frame, no caret), of all its text. Sizes are points, as in the file (`sz`).
   */
  private async fontSize(): Promise<void> {
    const p = this.currentPara();
    const selected = this.host.pager.querySelector<HTMLElement>('.sp[data-sp].oe-selected');
    const div = p?.closest<HTMLElement>('.sp[data-sp]') ?? selected;
    const box = div && this.boxes.get(div);
    if (!div || !box) return toast(t('tapTextFirst'));
    // The dialog takes the focus: remember the selection first.
    const sel = getSelection();
    const range = p && sel?.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    const at = range ? (range.startContainer instanceof Element ? range.startContainer : range.startContainer.parentElement) : paragraphs(div)[0];
    const scale = box.layout.body.fontScale || 1;
    const current = at ? (parseFloat(getComputedStyle(at).fontSize) * 72) / 96 / scale : undefined;
    // The dialog takes the focus; that isn't leaving the box (which would redraw the slide).
    this.inDialog = true;
    const pt = await askFontSize(current).finally(() => (this.inDialog = false));
    if (!pt) {
      p?.focus();
      return;
    }
    const px = `${+((pt * 96) / 72 * scale).toFixed(2)}px`;
    if (p && range && !range.collapsed && p.contains(range.commonAncestorContainer)) {
      // Just the selected text: in a span of its own, with sizes inside it cleared.
      const piece = range.extractContents();
      for (const bu of piece.querySelectorAll('.bu')) bu.remove();
      for (const e of piece.querySelectorAll<HTMLElement>('[style]')) e.style.fontSize = '';
      const span = document.createElement('span');
      span.dataset.sz = String(pt);
      span.style.fontSize = px;
      span.append(piece);
      range.insertNode(span);
      p.focus();
      const again = document.createRange();
      again.selectNodeContents(span);
      sel!.removeAllRanges();
      sel!.addRange(again);
      this.touched(p, false);
      box.reshaped = true;
      this.flush(false);
      return;
    }
    for (const para of p ? [p] : paragraphs(div).filter((x) => x.isContentEditable)) {
      para.dataset.sz = String(pt);
      for (const e of para.querySelectorAll<HTMLElement>('span:not(.bu)')) e.style.fontSize = px;
      for (const e of para.querySelectorAll<HTMLElement>('[data-sz]')) e.dataset.sz = String(pt);
      this.touched(para, false);
    }
    box.reshaped = true;
    if (p) {
      p.focus();
      if (range) {
        sel!.removeAllRanges();
        sel!.addRange(range);
      }
      this.flush(false);
      return;
    }
    // A whole box, not being typed in: write it, fit it and redraw now.
    this.flush(false);
    box.edited = false;
    if (this.fit(box)) this.commit(box.slide, false);
    await this.redraw(box.slide, null);
  }

  private cycleAlign(): void {
    const p = this.currentPara();
    if (!p) return toast(t('tapTextFirst'));
    const order = ['left', 'center', 'right', 'justify'];
    const now = getComputedStyle(p).textAlign.replace('start', 'left');
    const next = order[(order.indexOf(now) + 1) % order.length];
    p.style.textAlign = next;
    p.dataset.align = next;
    toast(t(`align_${next}` as Parameters<typeof t>[0]));
    this.touched(p, false);
    this.flush(false);
  }

  // ---------------------------------------------------------------- events

  private listen(): void {
    const pager = this.host.pager;
    pager.addEventListener('beforeinput', (e) => {
      const p = this.currentPara();
      if (e.inputType === 'historyUndo') {
        e.preventDefault();
        void this.undo();
      } else if (e.inputType === 'historyRedo') {
        e.preventDefault();
        void this.redo();
      } else if (p && (e.inputType === 'insertParagraph' || e.inputType === 'insertLineBreak')) {
        e.preventDefault();
        // A keyboard may send Enter while a word is still being composed. Changing the text
        // then makes it put the word back at the caret, in the new paragraph: wait for it.
        if (this.composing || e.isComposing) this.afterComposition = () => this.split(this.currentPara() ?? p);
        else this.split(p);
      } else if (p && e.inputType === 'deleteContentBackward' && caretAt(p, 'start')) {
        e.preventDefault();
        this.join(sibling(p, -1), p);
      } else if (p && e.inputType === 'deleteContentForward' && caretAt(p, 'end')) {
        e.preventDefault();
        this.join(p, sibling(p, 1));
      }
    });
    pager.addEventListener('compositionstart', () => {
      this.composing = true;
    });
    pager.addEventListener('compositionend', () => {
      this.composing = false;
      const then = this.afterComposition;
      this.afterComposition = null;
      if (then) setTimeout(then);
    });
    pager.addEventListener('input', (e) => {
      const p = (e.target as Element).closest?.<HTMLElement>('p[contenteditable="true"]');
      if (p) this.touched(p, true);
    });
    // Leaving a box: write it back and fit its text, as PowerPoint does when you click away.
    pager.addEventListener('focusout', (e) => {
      const from = (e.target as Element).closest?.<HTMLElement>('.sp[data-sp]');
      const to = (e.relatedTarget as Element | null)?.closest?.('.sp[data-sp]');
      if (from && from !== to) void this.leave(from);
    });
    pager.addEventListener('click', (e) => {
      const target = e.target as Element;
      if (target.closest('p[contenteditable="true"]')) return;
      const div = target.closest<HTMLElement>('.sp[data-sp]');
      // A tap in a box but not on its text: the caret goes to the end of its last paragraph.
      const last = div && paragraphs(div).filter((p) => p.isContentEditable).pop();
      if (last) {
        placeCaret(last, Infinity);
        return;
      }
      if (target.closest('p.oe-locked')) toast(t('pptxLocked'));
    });
  }

  /** The user changed paragraph `p`: write it back after a pause. */
  private touched(p: HTMLElement, typing: boolean): void {
    const div = p.closest<HTMLElement>('.sp[data-sp]');
    const box = div && this.boxes.get(div);
    if (!div || !box) return;
    box.edited = true;
    this.markEmpty(div);
    this.pending.add(div);
    if (!typing) this.group = null;
    clearTimeout(this.timer);
    this.timer = window.setTimeout(() => this.flush(true), 400);
  }

  /** Writes boxes with pending changes into their slides' XML. */
  flush(typing = true): void {
    clearTimeout(this.timer);
    const slides = new Set<number>();
    for (const div of this.pending) {
      const box = this.boxes.get(div);
      if (box && this.sync(div, box)) slides.add(box.slide);
    }
    this.pending.clear();
    for (const i of slides) this.commit(i, typing);
  }

  private async leave(div: HTMLElement): Promise<void> {
    const box = this.boxes.get(div);
    if (!box) return;
    if (this.inDialog) return;
    // focusout comes before the new focus: wait for it, so the redraw keeps the caret there.
    await new Promise((r) => setTimeout(r));
    if (this.inDialog) return;
    this.flush(true);
    if (!box.edited || !div.isConnected) return;
    box.edited = false;
    const fitted = this.fit(box);
    if (fitted) this.commit(box.slide, false);
    // New paragraphs were drawn as copies of their neighbours: redraw for real bullets and numbers.
    if (fitted || box.reshaped) await this.redraw(box.slide);
  }

  // Enter and joins only change the screen while the box is being typed in (redrawing it then
  // would upset the keyboard); the slide is drawn again from its XML when the box is left.

  /** Enter: the text after the caret becomes a new paragraph like this one. */
  private split(p: HTMLElement): void {
    const sel = getSelection();
    if (!sel?.rangeCount || !p.isConnected) return;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    const tail = document.createRange();
    tail.setStart(range.endContainer, range.endOffset);
    tail.setEnd(p, p.childNodes.length);
    const rest = tail.extractContents();
    for (const bu of rest.querySelectorAll('.bu')) bu.remove();
    const next = p.cloneNode(false) as HTMLElement;
    const bullet = p.querySelector('.bu');
    if (bullet) next.append(bullet.cloneNode(true));
    next.append(rest);
    p.after(next);
    // An empty paragraph keeps a styled stand-in, so text typed into it looks like its neighbours.
    for (const para of [p, next]) if (!plainText(para)) standIn(para, p);
    placeCaret(next, 0);
    this.reshaped(p);
  }

  /** Backspace at the start / Delete at the end: join two paragraphs of one box. */
  private join(into: HTMLElement | null, from: HTMLElement | null): void {
    if (!into || !from) return;
    const offset = textLength(into);
    for (const bu of from.querySelectorAll('.bu')) bu.remove();
    into.append(...from.childNodes);
    from.remove();
    placeCaret(into, offset);
    this.reshaped(into);
  }

  private reshaped(p: HTMLElement): void {
    const div = p.closest<HTMLElement>('.sp[data-sp]');
    const box = div && this.boxes.get(div);
    if (box) box.reshaped = true;
    this.touched(p, false);
    this.flush(false);
  }

  // ---------------------------------------------------------------- XML

  /** Rebuilds a box's paragraphs into the shape's text body. Returns whether the XML changed. */
  private sync(div: HTMLElement, box: Box): boolean {
    const doc = box.sp.ownerDocument;
    const seen = new Set<number>();
    const out: Element[] = [];
    let anyText = false;
    for (const p of paragraphs(div)) {
      const k = Number(p.dataset.p);
      const again = seen.has(k);
      seen.add(k);
      const src = box.src[k] ?? null;
      if (src && !again && (p.classList.contains('oe-locked') || (this.pristine.get(p) === p.innerHTML && !p.dataset.align && !p.dataset.sz))) {
        out.push(src.cloneNode(true) as Element);
        anyText ||= !!src.textContent;
        continue;
      }
      if (p.classList.contains('oe-locked')) continue;
      const built = this.buildParagraph(doc, p, src, box);
      anyText ||= !!built.textContent;
      out.push(built);
    }
    let txBody = child(box.sp, 'txBody');
    if (!txBody) {
      // An empty placeholder: it gets a text body only once it has text.
      if (!anyText) return false;
      txBody = doc.createElementNS(P, `${prefix(doc, P, 'p')}:txBody`);
      txBody.append(doc.createElementNS(A, `${prefix(doc, A, 'a')}:bodyPr`), doc.createElementNS(A, `${prefix(doc, A, 'a')}:lstStyle`));
      box.sp.insertBefore(txBody, child(box.sp, 'extLst'));
    }
    const before = serialize(txBody);
    for (const p of children(txBody, 'p')) p.remove();
    if (!out.length) out.push(doc.createElementNS(A, `${prefix(doc, A, 'a')}:p`));
    txBody.append(...out);
    return serialize(txBody) !== before;
  }

  private buildParagraph(doc: Document, p: HTMLElement, src: Element | null, box: Box): Element {
    const a = prefix(doc, A, 'a');
    const out = doc.createElementNS(A, `${a}:p`);
    let pPr = (child(src, 'pPr')?.cloneNode(true) as Element | undefined) ?? null;
    if (p.dataset.align) {
      pPr ??= doc.createElementNS(A, `${a}:pPr`);
      pPr.setAttribute('algn', ALGN[p.dataset.align] ?? 'l');
    }
    if (pPr) out.append(pPr);

    // The text, in pieces with the same source run and the same bold/italic/underline.
    type Piece = { run: string | null; look: Look; base: Look; text: string; br: boolean; size?: string };
    // A size the user chose (on a span, or for the whole paragraph), in points.
    const sizeOf = (holder: Element) => {
      const e = holder.closest<HTMLElement>('[data-sz]');
      return e && p.contains(e) ? e.dataset.sz : p.dataset.sz;
    };
    const pieces: Piece[] = [];
    const firstRun = src ? [...src.children].findIndex((c) => c.localName === 'r') : -1;
    let lastRun: string | null = firstRun >= 0 ? `${p.dataset.p}_${firstRun}` : null;
    const walk = (node: Node) => {
      if (node instanceof HTMLElement && node.classList.contains('bu')) return;
      if (node.nodeType === Node.TEXT_NODE) {
        const text = (node.textContent ?? '').replace(ZWSP, '').replace(/ /g, ' ');
        if (!text) return;
        const holder = node.parentElement!;
        const runEl = holder.closest<HTMLElement>('[data-r]');
        const run = runEl && p.contains(runEl) ? runEl.dataset.r! : lastRun;
        lastRun = run;
        pieces.push({ run, look: computedLook(holder, p), base: drawnLook(holder, p), text, br: false, size: sizeOf(holder) });
        return;
      }
      if (node instanceof HTMLBRElement) {
        // The <br> an empty paragraph keeps to stay tappable is not a line break.
        if (!node.nextSibling) return;
        pieces.push({ run: lastRun, look: computedLook(node.parentElement!, p), base: drawnLook(node.parentElement!, p), text: '', br: true, size: sizeOf(node.parentElement!) });
        return;
      }
      for (const c of node.childNodes) walk(c);
    };
    walk(p);

    const endPr = child(src, 'endParaRPr');
    let current: { key: string; r: Element } | null = null;
    for (const piece of pieces) {
      const key = `${piece.run}|${lookFlags(piece.look)}|${piece.size ?? ''}`;
      const source = sourceRun(piece.run, box);
      const rPr = (child(source, 'rPr')?.cloneNode(true) as Element | undefined) ?? renamed(doc, endPr, `${a}:rPr`);
      // Spell-check marks no longer apply to the new text.
      rPr.removeAttribute('err');
      const base = piece.base;
      if (piece.look.b !== base.b) rPr.setAttribute('b', piece.look.b ? '1' : '0');
      if (piece.look.i !== base.i) rPr.setAttribute('i', piece.look.i ? '1' : '0');
      if (piece.look.u !== base.u) rPr.setAttribute('u', piece.look.u ? 'sng' : 'none');
      if (piece.size) rPr.setAttribute('sz', String(Math.round(Number(piece.size) * 100)));
      const hasRPr = rPr.attributes.length > 0 || rPr.childNodes.length > 0;
      if (piece.br) {
        const br = doc.createElementNS(A, `${a}:br`);
        if (hasRPr) br.append(rPr);
        out.append(br);
        current = null;
        continue;
      }
      if (!current || current.key !== key) {
        const r = doc.createElementNS(A, `${a}:r`);
        if (hasRPr) r.append(rPr);
        r.append(doc.createElementNS(A, `${a}:t`));
        out.append(r);
        current = { key, r };
      }
      current.r.lastElementChild!.textContent += piece.text;
    }
    if (endPr || p.dataset.sz) {
      const end = (endPr?.cloneNode(true) as Element | undefined) ?? doc.createElementNS(A, `${a}:endParaRPr`);
      if (p.dataset.sz) end.setAttribute('sz', String(Math.round(Number(p.dataset.sz) * 100)));
      out.append(end);
    }
    return out;
  }

  /**
   * Fits a box's text as PowerPoint would: shrinks it (normAutofit) or resizes the shape to it
   * (spAutoFit). Returns whether the XML changed.
   */
  private fit(box: Box): boolean {
    const { layout } = box;
    if (layout.body.vert !== 'horz') return false;
    if (layout.body.autofit === 'norm') {
      let pick = SHRINK[SHRINK.length - 1];
      for (const step of SHRINK) {
        if (this.measure(layout.html({ fontScale: step[0], lineReduction: step[1] })) <= layout.h + 0.5) {
          pick = step;
          break;
        }
      }
      if (pick[0] === layout.body.fontScale && Math.abs(pick[1] - layout.body.lineReduction) < 1e-6) return false;
      const bodyPr = this.bodyPr(box.sp);
      let fit = children(bodyPr).find((c) => /Autofit$|^spAutoFit$/.test(c.localName));
      if (fit?.localName !== 'normAutofit') {
        fit?.remove();
        fit = bodyPr.ownerDocument.createElementNS(A, `${prefix(bodyPr.ownerDocument, A, 'a')}:normAutofit`);
        bodyPr.insertBefore(fit, child(bodyPr, 'prstTxWarp')?.nextSibling ?? bodyPr.firstChild);
      }
      if (pick[0] === 1) fit.removeAttribute('fontScale');
      else fit.setAttribute('fontScale', String(Math.round(pick[0] * 100000)));
      if (pick[1]) fit.setAttribute('lnSpcReduction', String(Math.round(pick[1] * 100000)));
      else fit.removeAttribute('lnSpcReduction');
      // Back to full size: the layout's autofit says the same, so the shape needs none of its own.
      if (!fit.attributes.length && layout.inherited === 'norm') fit.remove();
      return true;
    }
    if (layout.body.autofit === 'shape' && layout.body.wrap) {
      // Only a shape with its own position grows; one placed by its layout keeps the layout's.
      const ext = child(child(child(box.sp, 'spPr'), 'xfrm'), 'ext');
      if (!ext) return false;
      const h = this.measure(layout.html());
      if (Math.abs(h - layout.h) < 1) return false;
      ext.setAttribute('cy', String(Math.round(h / layout.ky)));
      return true;
    }
    return false;
  }

  /** The height text laid out as `html` needs, insets included, in px. */
  private measure(html: string): number {
    if (!this.measureHost) {
      this.measureHost = el('div', 'slide');
      this.measureHost.style.cssText = 'position:fixed;left:-100000px;top:0;visibility:hidden;width:4000px;height:4000px';
      document.body.append(this.measureHost);
    }
    this.measureHost.innerHTML = html;
    const tx = this.measureHost.firstElementChild as HTMLElement;
    const cs = getComputedStyle(tx);
    let h = parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom);
    for (const p of tx.children) {
      const ps = getComputedStyle(p);
      h += (p as HTMLElement).offsetHeight + parseFloat(ps.marginTop) + parseFloat(ps.marginBottom);
    }
    return h;
  }

  private bodyPr(sp: Element): Element {
    const txBody = child(sp, 'txBody')!;
    let bodyPr = child(txBody, 'bodyPr');
    if (!bodyPr) {
      bodyPr = sp.ownerDocument.createElementNS(A, `${prefix(sp.ownerDocument, A, 'a')}:bodyPr`);
      txBody.prepend(bodyPr);
    }
    return bodyPr;
  }

  // ---------------------------------------------------------------- history

  /** Records slide `i`'s XML after a change. */
  private commit(i: number, typing: boolean): void {
    const state = this.states.get(i);
    if (!state) return;
    const xml = serialize(state.doc);
    if (xml === state.last) return;
    const now = Date.now();
    const continues = typing && this.group?.slide === i && now - this.group.at < 1500;
    if (!continues) this.undoStack.push({ slide: i, xml: state.last });
    if (this.undoStack.length > 100) this.undoStack.shift();
    this.group = typing ? { slide: i, at: now } : null;
    state.last = xml;
    this.redoStack = [];
    this.changed(i);
  }

  private changed(i: number): void {
    setDirty([...this.states.values()].some((s) => s.last !== s.initial));
    this.host.refreshThumb(i);
  }

  private async undo(): Promise<void> {
    this.flush(true);
    const step = this.undoStack.pop();
    if (!step) return;
    this.redoStack.push({ slide: step.slide, xml: this.states.get(step.slide)!.last });
    await this.restore(step.slide, step.xml);
  }

  private async redo(): Promise<void> {
    this.flush(true);
    const step = this.redoStack.pop();
    if (!step) return;
    this.undoStack.push({ slide: step.slide, xml: this.states.get(step.slide)!.last });
    await this.restore(step.slide, step.xml);
  }

  private async restore(i: number, xml: string): Promise<void> {
    const state = this.states.get(i)!;
    state.doc = new DOMParser().parseFromString(xml, 'application/xml');
    state.last = xml;
    this.host.deck.setSlideDoc(i, state.doc);
    this.group = null;
    if (this.host.current() !== i) this.host.go(i);
    await this.redraw(i, null);
    this.changed(i);
  }

  /** Redraws a slide, keeping the caret where it was (or putting it at `caret`). */
  private async redraw(i: number, caret: Caret | null = this.caret()): Promise<void> {
    await this.host.redraw(i);
    const slideEl = this.host.slideEl(i);
    if (slideEl) this.shapes.drawn(i, slideEl);
    // Keys pressed while it was drawn went into the old boxes: write them back and draw again.
    for (let n = 0; n < 3 && [...this.pending].some((div) => this.boxes.get(div)?.slide === i); n++) {
      caret = this.caret() ?? caret;
      this.flush(false);
      await this.host.redraw(i);
    }
    if (!caret || caret.slide !== i) return;
    const div = this.host.slideEl(i)?.querySelector<HTMLElement>(`.sp[data-sp="${caret.sp}"]`);
    const ps = div ? paragraphs(div) : [];
    const p = ps[Math.min(caret.para, ps.length - 1)];
    if (p?.isContentEditable) placeCaret(p, caret.offset);
  }

  private caret(): Caret | null {
    const p = this.currentPara();
    const sel = getSelection();
    if (!p || !sel?.rangeCount) return null;
    const r = document.createRange();
    r.setStart(p, 0);
    r.setEnd(sel.anchorNode!, sel.anchorOffset);
    return caretIn(p, plainText(r.cloneContents()).length);
  }

  // ---------------------------------------------------------------- saving

  private async produce(): Promise<Uint8Array> {
    this.flush(false);
    const deck = this.host.deck;
    const out = await JSZip.loadAsync(await deck.pkg.zip.generateAsync({ type: 'arraybuffer' }));
    const changed: number[] = [];
    for (const [i, state] of this.states) {
      if (state.last === state.initial) continue;
      out.file(deck.slidePaths[i], state.decl + withRootTag(state.last, state.rootTag), KEEP);
      changed.push(i);
    }
    const bytes = await out.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    // Read it back as the viewer would before it may replace the original.
    const check = await Deck.open(bytes.slice().buffer as ArrayBuffer);
    for (const i of changed) await check.slide(i);
    return bytes;
  }

  private saved(result: SaveResult): void {
    if (result.ok && result.mode === 'overwrite') location.reload();
  }
}

// ------------------------------------------------------------------ helpers

const ALGN: Record<string, string> = { left: 'l', center: 'ctr', right: 'r', justify: 'just' };

/** The original start tag in place of the serializer's, when it declares the same things. */
function withRootTag(xml: string, original: string): string {
  const tag = /^<[^>]*>/.exec(xml)?.[0] ?? '';
  const attrs = (t: string) => [...t.matchAll(/\s([^\s=]+)="([^"]*)"/g)].map((m) => `${m[1]}=${m[2]}`).sort().join(' ');
  return original && tag && attrs(tag) === attrs(original) ? original + xml.slice(tag.length) : xml;
}

function serialize(node: Node): string {
  return new XMLSerializer().serializeToString(node).replace(/^<\?xml[^>]*\?>\s*/, '');
}

/** The document's prefix for a namespace (slides normally use p: and a:). */
function prefix(doc: Document, ns: string, fallback: string): string {
  return doc.documentElement.lookupPrefix(ns) ?? fallback;
}

/** A copy of `el` (endParaRPr) under another name (rPr), or an empty element if there is none. */
function renamed(doc: Document, el: Element | null, name: string): Element {
  const out = doc.createElementNS(A, name);
  if (!el) return out;
  for (const attr of el.attributes) out.setAttribute(attr.name, attr.value);
  for (const c of el.childNodes) out.append(c.cloneNode(true));
  return out;
}

/** The XML run a span came from ("paragraph_child"). */
function sourceRun(id: string | null, box: Box): Element | null {
  if (!id) return null;
  const [pi, k] = id.split('_').map(Number);
  const el = box.src[pi]?.children[k] ?? null;
  return el?.localName === 'r' || el?.localName === 'br' ? el : null;
}

/** How the span holding some text looked when the slide was drawn. */
function drawnLook(holder: Element, p: HTMLElement): Look {
  const span = holder.closest<HTMLElement>('[data-look]');
  const flags = span && p.contains(span) ? span.dataset.look! : '000';
  return { b: flags[0] === '1', i: flags[1] === '1', u: flags[2] === '1' };
}

/** Gives an empty paragraph a zero-width character in a copy of `like`'s last text span. */
function standIn(p: HTMLElement, like: HTMLElement): void {
  const spans = [...like.querySelectorAll<HTMLElement>(':scope > span:not(.bu)')];
  const model = spans.pop();
  for (const c of [...p.childNodes]) if (!(c instanceof HTMLElement && c.classList.contains('bu'))) c.remove();
  const span = (model?.cloneNode(false) as HTMLElement | undefined) ?? document.createElement('span');
  span.textContent = '\u200B';
  p.append(span);
}

function textLayer(div: HTMLElement): HTMLElement | null {
  return div.querySelector<HTMLElement>(':scope > .tx');
}

function paragraphs(div: HTMLElement): HTMLElement[] {
  return [...(textLayer(div)?.children ?? [])].filter((c): c is HTMLElement => c.tagName === 'P');
}

function isEmpty(div: HTMLElement): boolean {
  return !paragraphs(div).some((p) => plainText(p).trim());
}

/** A paragraph's text as the XML will hold it: no bullet, no stand-in characters. */
function plainText(node: Node): string {
  const copy = node.cloneNode(true) as Element | DocumentFragment;
  for (const bu of copy.querySelectorAll('.bu')) bu.remove();
  return (copy.textContent ?? '').replace(ZWSP, '');
}

function textLength(p: HTMLElement): number {
  return plainText(p).length;
}

function caretIn(p: HTMLElement, offset: number): Caret | null {
  const div = p.closest<HTMLElement>('.sp[data-sp]');
  const slide = p.closest<HTMLElement>('.slide[data-slide]');
  if (!div || !slide) return null;
  return { slide: Number(slide.dataset.slide), sp: Number(div.dataset.sp), para: paragraphs(div).indexOf(p), offset };
}

/** Puts the caret `offset` characters into a paragraph's text (Infinity: at the end). */
function placeCaret(p: HTMLElement, offset: number): void {
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => (n.parentElement?.closest('.bu') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  const range = document.createRange();
  let placed = false;
  let left = offset;
  let last: Text | null = null;
  for (let n = walker.nextNode() as Text | null; n; n = walker.nextNode() as Text | null) {
    last = n;
    const text = n.data;
    let i = 0;
    for (; i < text.length && left > 0; i++) if (text[i] !== '​') left--;
    if (left === 0 && offset !== Infinity) {
      range.setStart(n, i);
      placed = true;
      break;
    }
  }
  if (!placed) {
    if (last) range.setStart(last, last.data.length);
    else range.setStart(p, p.childNodes.length);
  }
  range.collapse(true);
  p.focus({ preventScroll: true });
  const sel = getSelection()!;
  sel.removeAllRanges();
  sel.addRange(range);
}

function caretAt(p: HTMLElement, where: 'start' | 'end'): boolean {
  const sel = getSelection();
  if (!sel?.isCollapsed || !sel.rangeCount) return false;
  const r = document.createRange();
  if (where === 'start') {
    r.setStart(p, 0);
    r.setEnd(sel.anchorNode!, sel.anchorOffset);
  } else {
    r.setStart(sel.anchorNode!, sel.anchorOffset);
    r.setEnd(p, p.childNodes.length);
  }
  return plainText(r.cloneContents()).length === 0;
}

/** The editable paragraph just before / after in the same box. */
function sibling(p: HTMLElement, dir: 1 | -1): HTMLElement | null {
  const s = (dir < 0 ? p.previousElementSibling : p.nextElementSibling) as HTMLElement | null;
  return s?.tagName === 'P' && s.isContentEditable ? s : null;
}

/** Bold / italic / underline from a run span's own style, as drawn. */
function styleLook(span: HTMLElement): Look {
  return {
    b: Number(span.style.fontWeight) >= 600,
    i: span.style.fontStyle === 'italic',
    u: span.style.textDecorationLine.includes('underline') || span.style.textDecoration.includes('underline'),
  };
}

/** Bold / italic / underline as shown now (underline can come from ancestors up to `stop`). */
function computedLook(elm: Element, stop: Element): Look {
  const cs = getComputedStyle(elm);
  let u = false;
  for (let e: Element | null = elm; e; e = e.parentElement) {
    if (getComputedStyle(e).textDecorationLine.includes('underline')) u = true;
    if (e === stop) break;
  }
  return { b: Number(cs.fontWeight) >= 600, i: cs.fontStyle === 'italic', u };
}

function lookFlags(look: Look): string {
  return `${+look.b}${+look.i}${+look.u}`;
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className = ''): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  return e;
}

// Material Symbols paths (Apache-2.0).
const svg = (d: string) => `<svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><path fill="currentColor" d="${d}"/></svg>`;
const ICON_ALIGN = svg('M3 21h18v-2H3v2zm0-4h12v-2H3v2zm0-4h18v-2H3v2zm0-4h12V7H3v2zm0-6v2h18V3H3z');
const ICON_UNDO = svg('M12.5 8c-2.65 0-5.05.99-6.9 2.6L2 7v9h9l-3.62-3.62c1.39-1.16 3.16-1.88 5.12-1.88 3.54 0 6.55 2.31 7.6 5.5l2.37-.78C21.08 11.03 17.15 8 12.5 8z');
const ICON_REDO = svg('M18.4 10.6C16.55 8.99 14.15 8 11.5 8c-4.65 0-8.58 3.03-9.96 7.22L3.9 16c1.05-3.19 4.05-5.5 7.6-5.5 1.95 0 3.73.72 5.12 1.88L13 16h9V7l-3.6 3.6z');
