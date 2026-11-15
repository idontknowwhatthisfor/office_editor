import { parseAsync, renderDocument } from 'docx-preview';
import { prepareCharts } from './charts';
import JSZip from 'jszip';
import { t } from '../i18n';
import { registerEditor, setDirty, type SaveResult } from '../save';
import { fontSizeField, holdable, HOLD_MS, parseFontSize } from '../ui';
import { openStylePanel, readStyle, saveStyle } from '../pdf/style';
import { overlay } from '../overlay';
import { PictureTool } from './pictures';
import { fontsLoaded, GUTTER, scaleLineHeights } from './viewer';

// Word editing (PRD M7), done in place on the file's own XML.
//
// The document is shown for editing as docx-preview renders it, on paper of the document's own
// size, scaled as a whole to the screen width so it looks as it does in the viewer (the pages
// aren't split, though: each section is one long sheet). Before rendering, a copy of document.xml gets a bookmark at the start of every
// paragraph and before every run; docx-preview turns bookmarks into <span id>, so each paragraph
// and run on screen can be traced back to its XML element. On save only the paragraphs the user
// changed are rewritten, from the original elements: their properties, and the properties of
// the runs the text came from, are kept. Everything else in the file stays byte for byte.
//
// A paragraph can be edited only if it holds nothing but plain text runs (no pictures, fields,
// links, tracked changes…); others are shown but locked.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NUMBERING_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
/** Replace files without adding folder entries the original zip didn't have. */
const KEEP = { createFolders: false };
const NUMBERING_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';

/** Run children that are plain text (or harmless markers). */
const PLAIN_RUN = new Set(['rPr', 't', 'tab', 'br', 'lastRenderedPageBreak', 'softHyphen', 'noBreakHyphen']);
/** Paragraph children that don't stop it being edited. */
const PLAIN_PARA = new Set(['pPr', 'r', 'bookmarkStart', 'bookmarkEnd', 'proofErr', 'permStart', 'permEnd']);

type List = 'bullet' | 'number' | 'none';

interface Look {
  b: boolean;
  i: boolean;
  u: boolean;
}

export async function startDocxEditing(root: HTMLElement, data: ArrayBuffer, fileName: string): Promise<boolean> {
  const zip = await JSZip.loadAsync(data);
  const docPath = await mainPart(zip);
  const xml = await zip.file(docPath)!.async('string');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const problems = unsupported(doc, zip);
  if (problems.length && !(await confirmDialog(t('docxFidelityTitle'), t('docxFidelityBody', { list: problems.join(', ') }), t('editAnyway')))) {
    return false;
  }
  const editor = new DocxEditor(root, zip, docPath, doc, fileName);
  await editor.show(data);
  return true;
}

async function mainPart(zip: JSZip): Promise<string> {
  const rels = await zip.file('_rels/.rels')?.async('string');
  const m = rels && /Type="[^"]*\/officeDocument"[^>]*Target="\/?([^"]+)"|Target="\/?([^"]+)"[^>]*Type="[^"]*\/officeDocument"/.exec(rels);
  return (m && (m[1] ?? m[2])) || 'word/document.xml';
}

/** Content the editor can't change, as names for the warning. */
function unsupported(doc: Document, zip: JSZip): string[] {
  const has = (name: string, ns = W) => doc.getElementsByTagNameNS(ns, name).length > 0;
  const out: string[] = [];
  if (has('ins') || has('del') || has('moveFrom') || has('moveTo')) out.push(t('trackedChanges'));
  if (has('txbxContent')) out.push(t('textBoxes'));
  if (doc.getElementsByTagNameNS('*', 'relIds').length) out.push('SmartArt');
  if (has('object') || has('pict')) out.push(t('embeddedObjects'));
  if (doc.getElementsByTagNameNS('*', 'oMath').length) out.push(t('equations'));
  if (zip.file('word/comments.xml')) out.push(t('comments'));
  return out;
}

/** The Word editor's own underline and text colours, kept apart from the PDF tools'. */
const STYLE_KEY = 'officeeditor.docxStyle';

/** The kept colours; until one is chosen, black (Word's automatic colour), not the PDF tools' red. */
function wordStyle(): ReturnType<typeof readStyle> {
  const style = readStyle(STYLE_KEY);
  let kept = false;
  try {
    kept = localStorage.getItem(STYLE_KEY) !== null;
  } catch {
    // Not kept.
  }
  if (!kept) style.underline.colour = style.text.colour = [0, 0, 0];
  return style;
}

class DocxEditor {
  private paras: Element[];
  private editable: boolean[];
  private view = el('div', 'docx-view docx-edit');
  /** Holds the view, which is scaled with a transform, and takes its scaled size. */
  private frame = el('div', 'docx-edit-frame');
  private bar = el('div', 'docx-edit-bar');
  /** Each paragraph's markup when editing began, by paragraph number. */
  private original = new Map<string, string>();
  private history: string[] = [];
  private future: string[] = [];
  private snapshotTimer = 0;
  private styles: Document | null = null;
  /** Lists started in this session; each gets its own numbering definition on save. */
  private pendingLists: ('bullet' | 'number')[] = [];
  /** Inserting, moving, resizing and cropping pictures (made once the document is shown). */
  private pictures!: PictureTool;
  /** The settings panel that is open, and the button it belongs to. */
  private panel: { button: HTMLElement; close: () => void } | null = null;
  /** Underline and text colours from the settings panels. */
  private style = wordStyle();

  constructor(
    private root: HTMLElement, private zip: JSZip, private docPath: string, private doc: Document, private fileName: string,
  ) {
    this.paras = [...doc.getElementsByTagNameNS(W, 'p')];
    this.editable = this.paras.map((p) => isPlain(p));
  }

  async show(data: ArrayBuffer): Promise<void> {
    // A copy of the document with a bookmark marking every paragraph and run.
    const marked = this.doc.cloneNode(true) as Document;
    [...marked.getElementsByTagNameNS(W, 'p')].forEach((p, i) => {
      const pPr = child(p, 'pPr');
      p.insertBefore(bookmark(marked, `oep${i}`), pPr ? pPr.nextSibling : p.firstChild);
      children(p, 'r').forEach((r, k) => p.insertBefore(bookmark(marked, `oer${i}_${k}`), r));
    });
    const copy = await JSZip.loadAsync(data);
    copy.file(this.docPath, new XMLSerializer().serializeToString(marked), KEEP);
    const bytes = await copy.generateAsync({ type: 'arraybuffer' });
    const options = {
      className: 'docx', inWrapper: true, breakPages: false, ignoreLastRenderedPageBreak: true,
      renderHeaders: false, renderFooters: false, renderFootnotes: false, renderEndnotes: false, useBase64URL: true,
    };
    const parsed = await parseAsync(bytes, options);
    const drawCharts = await prepareCharts(parsed, bytes);
    this.view.append(...(await renderDocument(parsed, options)));
    drawCharts(this.view);
    const lineHeights = scaleLineHeights(this.view);
    for (const link of this.view.querySelectorAll('a[href]')) link.removeAttribute('href');

    for (const marker of this.view.querySelectorAll<HTMLElement>('span[id^="oep"]')) {
      const p = marker.closest('p');
      const i = Number(marker.id.slice(3));
      marker.remove();
      if (!p || p.dataset.p !== undefined) continue;
      p.dataset.p = String(i);
      if (this.editable[i] && !p.closest('.oe-locked')) p.contentEditable = 'true';
      else p.classList.add('oe-locked');
    }
    for (const marker of this.view.querySelectorAll<HTMLElement>('span[id^="oer"]')) {
      const run = marker.nextElementSibling as HTMLElement | null;
      if (run?.tagName === 'SPAN') {
        run.dataset.r = marker.id.slice(3);
        const look = lookOf(run);
        run.dataset.look = `${+look.b}${+look.i}${+look.u}`;
        // Tabs are drawn as an em space in their own span; keep them as tabs.
        for (const s of run.querySelectorAll<HTMLElement>(':scope > span')) {
          if (s.textContent === ' ') {
            s.dataset.tab = '1';
            s.contentEditable = 'false';
          }
        }
      }
      marker.remove();
    }

    this.buildBar();
    this.frame.append(this.view);
    this.root.replaceChildren(this.frame);
    // The toolbar (like the toast, dialogs and panels) stays at normal size on a pinch-zoomed page.
    overlay().append(this.bar);
    await fontsLoaded(this.view);
    lineHeights.measure();
    for (const p of this.view.querySelectorAll<HTMLElement>('p[data-p]')) this.original.set(p.dataset.p!, p.outerHTML);
    this.fitToScreen();
    this.pictures = new PictureTool({ view: this.view, frame: this.frame, doc: this.doc, changed: () => this.changed(true) });
    this.pictures.prepare();
    this.listen();
    this.history = [this.view.innerHTML];
    registerEditor({ fileName: this.fileName, produce: () => this.produce(), saved: (r) => this.saved(r) });
    toast(t(this.view.querySelector('.oe-locked') ? 'docxEditHint' : 'docxEditHintPlain'));
  }

  /**
   * Scales the paper to the screen width, as the viewer does. A transform only scales the painted
   * result, so text wraps exactly as on the page; the frame is kept at the scaled size.
   */
  private fitToScreen(): void {
    const sections = [...this.view.querySelectorAll<HTMLElement>('.docx-wrapper > section')];
    const pageWidth = Math.max(1, ...sections.map((s) => s.offsetWidth));
    this.view.style.width = `${pageWidth}px`;
    let scale = 1;
    const size = () => {
      this.frame.style.width = `${pageWidth * scale}px`;
      this.frame.style.height = `${this.view.offsetHeight * scale}px`;
    };
    const fit = () => {
      scale = Math.min(1, (document.documentElement.clientWidth - 2 * GUTTER) / pageWidth);
      this.view.style.transform = `scale(${scale})`;
      size();
    };
    fit();
    // Typing changes the height.
    new ResizeObserver(size).observe(this.view);
    let width = document.documentElement.clientWidth;
    window.addEventListener('resize', () => {
      if (document.documentElement.clientWidth !== width) {
        width = document.documentElement.clientWidth;
        fit();
      }
    });
  }

  // ---------------------------------------------------------------- toolbar

  private buildBar(): void {
    /**
     * A bar button. Any button first closes an open settings panel (so only one is ever open);
     * the button whose panel is open just closes it. `hold`: holding opens the button's settings,
     * as in the PDF editor. `tapOnly`: a long press does nothing (no click on release).
     */
    const add = (label: string, title: string, action: () => void, cls = '', hold?: (b: HTMLButtonElement) => void, tapOnly = false) => {
      const b = el('button', `docx-tool ${cls}`);
      b.innerHTML = label;
      b.title = title;
      b.setAttribute('aria-label', title);
      // Keep the text selection: don't let the button take focus.
      b.addEventListener('pointerdown', (e) => e.preventDefault());
      const tap = () => {
        const open = this.panel;
        open?.close();
        if (open?.button !== b) action();
      };
      if (hold) {
        holdable(b, tap, () => {
          if (this.panel?.button === b) return;
          this.panel?.close();
          hold(b);
        });
      } else if (tapOnly) {
        let down = 0;
        b.addEventListener('pointerdown', () => (down = Date.now()));
        b.addEventListener('contextmenu', (e) => e.preventDefault());
        b.addEventListener('click', () => {
          if (Date.now() - down < HOLD_MS) tap();
        });
      } else {
        b.onclick = tap;
      }
      this.bar.append(b);
      return b;
    };
    add('<b>B</b>', t('bold'), () => this.format('bold'));
    add('<i>I</i>', t('italic'), () => this.format('italic'));
    add(ICON_UNDERLINE, t('underlineText'), () => this.format('underline'), '', (b) => this.showStyle(b, 'underline'));
    // The text button's only action is its settings: a tap opens them (or closes them when open).
    const text = add(ICON_TEXT, t('textSettings'), () => this.showStyle(text, 'text'), '', undefined, true);
    add(ICON_ALIGN, t('alignment'), () => this.cycleAlign());
    add('H', t('headingStyle'), () => void this.headingMenu(), 'docx-tool-h');
    add(ICON_LIST, t('listStyle'), () => this.cycleList(), '', () => void this.listMenu());
    add(ICON_PICTURE, t('insertPicture'), () => this.pictures.pick());
    add(ICON_UNDO, t('undo'), () => this.undo());
    add(ICON_REDO, t('redo'), () => this.redo());
  }

  /**
   * The settings panel of the underline or text button, as in the PDF editor. A colour chosen
   * there is applied at once to the selected text (a text colour with only a caret: the whole
   * paragraph; selected text is underlined in an underline colour), and the underline colour is
   * the underline button's from then on. The text panel also has the font size, applied the same
   * way as the text colour as soon as a valid size is typed.
   */
  private showStyle(button: HTMLElement, target: 'underline' | 'text'): void {
    const p = this.current();
    const sel = getSelection();
    const range = p && sel?.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    // The colour and size go on this span (made on the first change, changed after), or on p.
    let span: HTMLElement | null = null;
    const selected = !!(p && range && !range.collapsed && p.contains(range.commonAncestorContainer));
    const holder = (): HTMLElement => {
      if (!span) {
        span = wrapRange(range!);
        range!.selectNodeContents(span);
      }
      return span;
    };
    let extra: HTMLElement[] | undefined;
    if (target === 'text') {
      const at = range?.startContainer instanceof Element ? range.startContainer : range?.startContainer.parentElement ?? p;
      const { field, input, hint } = fontSizeField(at ? (parseFloat(getComputedStyle(at).fontSize) * 72) / 96 : undefined);
      field.classList.add('pdf-style-field');
      const caption = el('span', 'pdf-style-caption');
      caption.textContent = t('fontSize');
      field.prepend(caption);
      input.addEventListener('input', () => {
        const pt = parseFontSize(input.value);
        hint.classList.toggle('error', pt === null && input.value.trim() !== '');
        if (!p || pt === null) return;
        if (selected) setSpanSize(holder(), pt);
        else setParagraphSize(p, pt);
        // One undo step per pause, as for typing in the document ("20" passes through 2 pt).
        this.changed(false);
      });
      // Done on the keyboard: back to the document. The key must not reach the document too,
      // where it would replace the restored selection with a new paragraph.
      input.addEventListener('keydown', (e) => {
        if (e.key !== 'Enter') return;
        e.preventDefault();
        closePanel();
      });
      extra = [field, hint];
    }
    const apply = () => {
      if (!p) return;
      const colour = hexOf(this.style[target].colour);
      if (selected) {
        paint(holder(), target, colour);
      } else if (target === 'text') {
        paint(p, target, colour);
      } else {
        // Underlining the whole paragraph for a colour would surprise: the colour is only kept
        // for the underline button.
        return;
      }
      this.changed(true);
    };
    button.classList.add('active');
    // Put the keyboard (and the system's selection menu) away while the panel is open: on a phone
    // they leave it a sliver of the screen. The selection is kept in `range` and restored on close.
    if (p && document.activeElement instanceof HTMLElement && p.contains(document.activeElement)) {
      document.activeElement.blur();
      sel?.removeAllRanges();
    }
    const closePanel: () => void = openStylePanel({
      style: this.style,
      target,
      plain: true,
      extra,
      title: t(target === 'underline' ? 'underlineColour' : 'textSettings'),
      changed: () => {
        saveStyle(this.style, STYLE_KEY);
        apply();
      },
      closed: () => {
        button.classList.remove('active');
        if (this.panel?.button === button) this.panel = null;
        if (!p) return;
        // Back to the text, with the coloured part still selected.
        p.focus();
        if (range) {
          sel?.removeAllRanges();
          sel?.addRange(range);
        }
      },
    });
    this.panel = { button, close: closePanel };
  }

  private current(): HTMLElement | null {
    const node = getSelection()?.anchorNode;
    const p = (node instanceof Element ? node : node?.parentElement)?.closest<HTMLElement>('p[contenteditable="true"]');
    return p && this.view.contains(p) ? p : null;
  }

  private format(command: 'bold' | 'italic' | 'underline'): void {
    const p = this.current();
    if (!p) return toast(t('tapTextFirst'));
    document.execCommand('styleWithCSS', false, 'false');
    document.execCommand(command);
    // A new underline is drawn in the underline colour chosen in its settings.
    const colour = hexOf(this.style.underline.colour);
    if (command === 'underline' && colour !== '000000') {
      for (const u of p.querySelectorAll<HTMLElement>('u:not([data-ucolor])')) paint(u, 'underline', colour);
    }
    this.changed(true);
  }

  private cycleAlign(): void {
    const p = this.current();
    if (!p) return toast(t('tapTextFirst'));
    const order = ['left', 'center', 'right', 'justify'];
    const now = getComputedStyle(p).textAlign.replace('start', 'left');
    const next = order[(order.indexOf(now) + 1) % order.length];
    p.style.textAlign = next;
    p.dataset.align = next;
    toast(t(`align_${next}` as Parameters<typeof t>[0]));
    this.changed(true);
  }

  private async headingMenu(): Promise<void> {
    const p = this.current();
    if (!p) return toast(t('tapTextFirst'));
    const choice = await menu(t('headingStyle'), [
      ['normal', t('styleNormal')], ['h1', t('heading1')], ['h2', t('heading2')], ['h3', t('heading3')],
    ]);
    if (!choice) return;
    // docx.css previews the heading from this; the document's heading style decides the real look.
    p.dataset.heading = choice;
    this.changed(true);
  }

  /** The list button: none → bullets → numbering → none. */
  private cycleList(): void {
    const p = this.current();
    if (!p) return toast(t('tapTextFirst'));
    const now = p.dataset.list ?? (p.className.includes('docx-num-') ? 'existing' : 'none');
    const next: List = now === 'none' ? 'bullet' : now === 'bullet' ? 'number' : 'none';
    this.setList(p, next);
  }

  /** Holding the list button: choose the kind. */
  private async listMenu(): Promise<void> {
    const p = this.current();
    if (!p) return toast(t('tapTextFirst'));
    const choice = (await menu(t('listStyle'), [['bullet', t('bullets')], ['number', t('numbering')], ['none', t('noList')]])) as List | null;
    p.focus();
    if (choice) this.setList(p, choice);
  }

  private setList(p: HTMLElement, next: List): void {
    p.dataset.list = next;
    p.classList.remove('oe-bullet', 'oe-number');
    for (const c of [...p.classList]) if (c.startsWith('docx-num-')) p.classList.remove(c);
    if (next !== 'none') p.classList.add(next === 'bullet' ? 'oe-bullet' : 'oe-number');
    toast(t(next === 'bullet' ? 'bullets' : next === 'number' ? 'numbering' : 'noList'));
    this.changed(true);
  }

  // ---------------------------------------------------------------- typing, paragraphs, undo

  private listen(): void {
    this.view.addEventListener('beforeinput', (e) => {
      const p = this.current();
      if (e.inputType === 'historyUndo') {
        e.preventDefault();
        this.undo();
      } else if (e.inputType === 'historyRedo') {
        e.preventDefault();
        this.redo();
      } else if (p && (e.inputType === 'insertParagraph' || e.inputType === 'insertLineBreak')) {
        e.preventDefault();
        this.split(p);
      } else if (p && e.inputType === 'deleteContentBackward' && caretAtStart(p)) {
        e.preventDefault();
        this.merge(previousEditable(p), p);
      } else if (p && e.inputType === 'deleteContentForward' && caretAtEnd(p)) {
        e.preventDefault();
        this.merge(p, nextEditable(p));
      }
    });
    this.view.addEventListener('input', () => this.changed(false));
    // A tap on the paper below the text (all of the page, in a new document) goes to the end of
    // the last paragraph that can be edited, as in Word.
    this.view.addEventListener('click', (e) => {
      const target = e.target as Element;
      if (target.closest('p, table, img') || !target.closest('section.docx')) return;
      const paras = [...this.view.querySelectorAll<HTMLElement>('p[contenteditable="true"]')];
      const last = paras.filter((p) => p.getBoundingClientRect().top <= e.clientY).pop() ?? paras[0];
      if (!last) return;
      last.focus();
      const caret = document.createRange();
      caret.selectNodeContents(last);
      caret.collapse(false);
      const sel = getSelection()!;
      sel.removeAllRanges();
      sel.addRange(caret);
    });
  }

  /** Enter: the text after the caret becomes a new paragraph like this one. */
  private split(p: HTMLElement): void {
    const sel = getSelection()!;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    const tail = document.createRange();
    tail.setStart(range.endContainer, range.endOffset);
    tail.setEnd(p, p.childNodes.length);
    const rest = tail.extractContents();
    const next = p.cloneNode(false) as HTMLElement;
    // Enter at the end of a heading starts normal text, as in Word.
    if (!rest.textContent && p.dataset.heading?.startsWith('h')) next.dataset.heading = 'normal';
    next.append(rest);
    for (const empty of [p, next]) if (!empty.textContent) empty.append(document.createElement('br'));
    p.after(next);
    const caret = document.createRange();
    caret.setStart(next, 0);
    sel.removeAllRanges();
    sel.addRange(caret);
    next.focus();
    this.changed(true);
  }

  /** Backspace at the start / Delete at the end: join two paragraphs. */
  private merge(into: HTMLElement | null, from: HTMLElement | null): void {
    if (!into || !from) return;
    const caret = document.createRange();
    caret.selectNodeContents(into);
    caret.collapse(false);
    into.querySelectorAll(':scope > br:last-child').forEach((br) => br.remove());
    into.append(...from.childNodes);
    from.remove();
    into.focus();
    const sel = getSelection()!;
    sel.removeAllRanges();
    sel.addRange(caret);
    this.changed(true);
  }

  private changed(now: boolean): void {
    setDirty(true);
    this.future = [];
    clearTimeout(this.snapshotTimer);
    // Typing is grouped into one undo step per pause.
    if (now) this.snapshot();
    else this.snapshotTimer = window.setTimeout(() => this.snapshot(), 600);
  }

  private snapshot(): void {
    const html = this.view.innerHTML;
    if (html !== this.history[this.history.length - 1]) this.history.push(html);
    if (this.history.length > 100) this.history.shift();
  }

  private undo(): void {
    clearTimeout(this.snapshotTimer);
    this.snapshot();
    if (this.history.length < 2) return;
    this.future.push(this.history.pop()!);
    this.restore(this.history[this.history.length - 1]);
  }

  private redo(): void {
    const html = this.future.pop();
    if (!html) return;
    this.history.push(html);
    this.restore(html);
  }

  private restore(html: string): void {
    this.pictures.clear();
    this.view.innerHTML = html;
    setDirty(this.history.length > 1);
  }

  // ---------------------------------------------------------------- saving

  private async produce(): Promise<Uint8Array> {
    clearTimeout(this.snapshotTimer);
    this.pendingLists = [];
    const doc = this.doc.cloneNode(true) as Document;
    const paras = [...doc.getElementsByTagNameNS(W, 'p')];
    const seen = new Set<number>();
    let lastWritten: Element | null = null;
    let lastIndex = -1;
    const lists = { bullet: 0, number: 0 };
    // Each paragraph on screen and its paragraph in the file being written (for pictures).
    const written = new Map<HTMLElement, Element>();

    for (const p of this.view.querySelectorAll<HTMLElement>('p[data-p]')) {
      const i = Number(p.dataset.p);
      const source = paras[i];
      if (!source) continue;
      const isNew = seen.has(i);
      seen.add(i);
      written.set(p, source);
      if (!this.editable[i]) {
        lastWritten = source;
        lastIndex = i;
        continue;
      }
      if (!isNew && this.original.get(String(i)) === p.outerHTML) {
        lastWritten = source;
        lastIndex = i;
        continue;
      }
      const built = await this.buildParagraph(doc, p, i, lists, isNew);
      if (isNew) {
        // A paragraph made by Enter goes after the one it was split from.
        const after = lastIndex === i && lastWritten ? lastWritten : source;
        after.parentNode!.insertBefore(built, after.nextSibling);
      } else {
        source.parentNode!.replaceChild(built, source);
        paras[i] = built;
      }
      written.set(p, built);
      lastWritten = built;
      lastIndex = i;
    }
    // Editable paragraphs no longer on screen were joined into another one: remove them.
    this.editable.forEach((ok, i) => {
      if (ok && !seen.has(i) && paras[i].parentNode) {
        const parent = paras[i].parentNode!;
        parent.removeChild(paras[i]);
        // A table cell must keep at least one paragraph.
        if ((parent as Element).localName === 'tc' && !children(parent as Element, 'p').length) parent.appendChild(doc.createElementNS(W, 'w:p'));
      }
    });

    const out = await JSZip.loadAsync(await this.zip.generateAsync({ type: 'arraybuffer' }));
    await this.pictures.write(doc, out, this.docPath, paras, written);
    out.file(this.docPath, new XMLSerializer().serializeToString(doc), KEEP);
    if (this.styles) out.file('word/styles.xml', new XMLSerializer().serializeToString(this.styles), KEEP);
    if (this.pendingLists.length) await this.writeNumbering(out);
    const bytes = await out.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    // Read it back as a viewer would before it may replace the original.
    await parseAsync(bytes.buffer as ArrayBuffer, { className: 'docx' });
    const check = new DOMParser().parseFromString(await (await JSZip.loadAsync(bytes)).file(this.docPath)!.async('string'), 'application/xml');
    if (check.getElementsByTagName('parsererror').length) throw new Error('document.xml did not survive');
    return bytes;
  }

  private async buildParagraph(
    doc: Document, p: HTMLElement, i: number, lists: { bullet: number; number: number }, isNew: boolean,
  ): Promise<Element> {
    const source = this.doc.getElementsByTagNameNS(W, 'p')[i];
    const runs = children(source, 'r');
    const out = doc.createElementNS(W, 'w:p');
    for (const attr of source.attributes) if (!attr.name.startsWith('w14:')) out.setAttribute(attr.name, attr.value);
    const pPr = child(source, 'pPr');
    const newPPr = (pPr ? pPr.cloneNode(true) : doc.createElementNS(W, 'w:pPr')) as Element;
    // Bookmarks stay at the start of the paragraph (unless this is a new one).
    const keep = !isNew ? children(source).filter((c) => c.localName === 'bookmarkStart' || c.localName === 'bookmarkEnd') : [];

    if (p.dataset.align) setVal(doc, newPPr, 'jc', p.dataset.align === 'justify' ? 'both' : p.dataset.align);
    const heading = p.dataset.heading;
    if (heading) {
      const id = heading === 'normal' ? null : await this.headingStyle(Number(heading.slice(1)));
      const existing = child(newPPr, 'pStyle');
      if (!id) existing?.remove();
      else setVal(doc, newPPr, 'pStyle', id);
    }
    if (p.dataset.list) {
      child(newPPr, 'numPr')?.remove();
      if (p.dataset.list !== 'none') {
        const kind = p.dataset.list as 'bullet' | 'number';
        lists[kind] ||= this.listNum(kind);
        const numPr = doc.createElementNS(W, 'w:numPr');
        const ilvl = doc.createElementNS(W, 'w:ilvl');
        ilvl.setAttributeNS(W, 'w:val', '0');
        const numId = doc.createElementNS(W, 'w:numId');
        numId.setAttributeNS(W, 'w:val', String(lists[kind]));
        numPr.append(ilvl, numId);
        insertPPr(newPPr, numPr);
      } else if (styleHasList(child(newPPr, 'pStyle'))) {
        // Turn off numbering that comes from the paragraph's style.
        const numPr = doc.createElementNS(W, 'w:numPr');
        const numId = doc.createElementNS(W, 'w:numId');
        numId.setAttributeNS(W, 'w:val', '0');
        numPr.append(numId);
        insertPPr(newPPr, numPr);
      }
    }
    if (p.dataset.sz) {
      // The paragraph mark too (it sets the height of an empty paragraph).
      let markPr = child(newPPr, 'rPr');
      if (!markPr) {
        markPr = doc.createElementNS(W, 'w:rPr');
        insertPPr(newPPr, markPr);
      }
      const half = String(Math.round(Number(p.dataset.sz) * 2));
      setVal(doc, markPr, 'sz', half);
      setVal(doc, markPr, 'szCs', half);
    }
    if (newPPr.childNodes.length) out.append(newPPr);
    for (const b of keep) out.append(b.cloneNode(true));

    // The text, in pieces with the same source run and the same bold/italic/underline.
    type Piece = { run: string | null; look: Look; text: string; kind: 'text' | 'tab' | 'br'; size?: string; color?: string; ucolor?: string };
    // A size the user chose (on a span, or for the whole paragraph), in points.
    const sizeOf = (holder: Element) => {
      const e = holder.closest<HTMLElement>('[data-sz]');
      return e && p.contains(e) && e !== p ? e.dataset.sz : p.dataset.sz;
    };
    // Colours the user chose (hex, as Word writes them), the nearest one inside the paragraph.
    const colourOf = (holder: Element, key: 'color' | 'ucolor') => {
      const e = holder.closest<HTMLElement>(`[data-${key}]`);
      return e && p.contains(e) ? e.dataset[key] : undefined;
    };
    const coloursOf = (holder: Element) => ({ color: colourOf(holder, 'color'), ucolor: colourOf(holder, 'ucolor') });
    const pieces: Piece[] = [];
    let lastRun: string | null = runs.length ? `${i}_0` : null;
    const walk = (node: Node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent ?? '';
        if (!text) return;
        const holder = node.parentElement!;
        const runEl = holder.closest<HTMLElement>('[data-r]');
        const run = runEl && p.contains(runEl) ? runEl.dataset.r! : lastRun;
        lastRun = run;
        const tab = holder.closest<HTMLElement>('[data-tab]');
        pieces.push({ run, look: lookOf(holder, p), text: tab ? '\t' : text.replace(/ /g, ' '), kind: tab ? 'tab' : 'text', size: sizeOf(holder), ...coloursOf(holder) });
        return;
      }
      if (node instanceof HTMLBRElement) {
        // The <br> an empty paragraph needs to be tappable is not a line break.
        if (node === p.lastChild || !node.nextSibling) return;
        pieces.push({ run: lastRun, look: lookOf(node.parentElement!, p), text: '', kind: 'br', size: sizeOf(node.parentElement!), ...coloursOf(node.parentElement!) });
        return;
      }
      for (const c of node.childNodes) walk(c);
    };
    // The heading preview's bold is the style's, not the text's: read the text without it.
    if (heading) delete p.dataset.heading;
    try {
      walk(p);
    } finally {
      if (heading) p.dataset.heading = heading;
    }

    const runById = (id: string | null): { el: Element | null; look: Look } => {
      if (!id) return { el: null, look: { b: false, i: false, u: false } };
      const [pi, k] = id.split('_').map(Number);
      const el = children(this.doc.getElementsByTagNameNS(W, 'p')[pi], 'r')[k] ?? null;
      const span = this.view.querySelector<HTMLElement>(`[data-r="${id}"]`);
      const flags = span?.dataset.look ?? '000';
      return { el, look: { b: flags[0] === '1', i: flags[1] === '1', u: flags[2] === '1' } };
    };

    let current: { key: string; r: Element } | null = null;
    for (const piece of pieces) {
      const key = `${piece.run}|${+piece.look.b}${+piece.look.i}${+piece.look.u}|${piece.size ?? ''}|${piece.color ?? ''}|${piece.ucolor ?? ''}`;
      if (!current || current.key !== key) {
        const src = runById(piece.run);
        const r = doc.createElementNS(W, 'w:r');
        const rPr = (child(src.el, 'rPr')?.cloneNode(true) as Element | undefined) ?? doc.createElementNS(W, 'w:rPr');
        if (piece.look.b !== src.look.b) setToggle(doc, rPr, 'b', piece.look.b);
        if (piece.look.i !== src.look.i) setToggle(doc, rPr, 'i', piece.look.i);
        if (piece.look.u !== src.look.u) setVal(doc, rPr, 'u', piece.look.u ? 'single' : 'none');
        if (piece.color) setColour(doc, rPr, 'color', piece.color);
        if (piece.ucolor && piece.look.u) {
          // Keeps the run's own underline kind (double, dotted…) if it has one.
          if (child(rPr, 'u')?.getAttributeNS(W, 'val') === 'none' || !child(rPr, 'u')) setVal(doc, rPr, 'u', 'single');
          setColour(doc, rPr, 'u', piece.ucolor);
        }
        if (piece.size) {
          // Half-points; szCs is the size of complex-script text (Arabic, Thai…).
          const half = String(Math.round(Number(piece.size) * 2));
          setVal(doc, rPr, 'sz', half);
          setVal(doc, rPr, 'szCs', half);
        }
        if (rPr.childNodes.length) r.append(rPr);
        out.append(r);
        current = { key, r };
      }
      if (piece.kind === 'tab') current.r.append(doc.createElementNS(W, 'w:tab'));
      else if (piece.kind === 'br') current.r.append(doc.createElementNS(W, 'w:br'));
      else {
        const last = current.r.lastElementChild;
        if (last?.localName === 't') last.textContent += piece.text;
        else {
          const tEl = doc.createElementNS(W, 'w:t');
          tEl.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
          tEl.textContent = piece.text;
          current.r.append(tEl);
        }
      }
    }
    return out;
  }

  /** The id of the document's "heading N" style, adding the style if the document lacks it. */
  private async headingStyle(level: number): Promise<string> {
    this.styles ??= new DOMParser().parseFromString((await this.zip.file('word/styles.xml')?.async('string')) ?? `<w:styles xmlns:w="${W}"/>`, 'application/xml');
    const styles = [...this.styles.getElementsByTagNameNS(W, 'style')];
    const name = (s: Element) => child(s, 'name')?.getAttributeNS(W, 'val')?.toLowerCase();
    const found = styles.find((s) => name(s) === `heading ${level}`);
    if (found) return found.getAttributeNS(W, 'styleId') ?? `Heading${level}`;
    const normal = styles.find((s) => s.getAttributeNS(W, 'type') === 'paragraph' && s.getAttributeNS(W, 'default') === '1');
    const id = `Heading${level}`;
    const size = [32, 26, 24][level - 1] ?? 24;
    const xml = `<w:style xmlns:w="${W}" w:type="paragraph" w:styleId="${id}"><w:name w:val="heading ${level}"/>` +
      (normal ? `<w:basedOn w:val="${normal.getAttributeNS(W, 'styleId')}"/>` : '') +
      `<w:next w:val="${normal?.getAttributeNS(W, 'styleId') ?? 'Normal'}"/><w:uiPriority w:val="9"/><w:qFormat/>` +
      `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
      `<w:rPr><w:b/><w:bCs/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;
    const style = new DOMParser().parseFromString(xml, 'application/xml').documentElement;
    this.styles.documentElement.append(this.styles.importNode(style, true));
    return id;
  }

  /** A stand-in list id, replaced by a real one in writeNumbering (which reads numbering.xml). */
  private listNum(kind: 'bullet' | 'number'): number {
    this.pendingLists.push(kind);
    return 1_000_000 + this.pendingLists.length;
  }

  /** Adds the bullet / number list definitions used by new lists to numbering.xml. */
  private async writeNumbering(out: JSZip): Promise<void> {
    const pending = this.pendingLists;
    let xml = await out.file('word/numbering.xml')?.async('string');
    const created = !xml;
    xml ??= `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"/>`;
    const numbering = new DOMParser().parseFromString(xml, 'application/xml');
    const rootEl = numbering.documentElement;
    const maxOf = (tag: string, attr: string) =>
      Math.max(0, ...[...numbering.getElementsByTagNameNS(W, tag)].map((e) => Number(e.getAttributeNS(W, attr)) || 0));
    let abstractId = maxOf('abstractNum', 'abstractNumId');
    let numId = maxOf('num', 'numId');
    const replace = new Map<number, number>();
    pending.forEach((kind, n) => {
      abstractId++;
      numId++;
      const lvl = kind === 'bullet'
        ? '<w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>'
        : '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>';
      const abstract = new DOMParser().parseFromString(
        `<w:abstractNum xmlns:w="${W}" w:abstractNumId="${abstractId}"><w:multiLevelType w:val="hybridMultilevel"/>` +
        `<w:lvl w:ilvl="0"><w:start w:val="1"/>${lvl}<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>`,
        'application/xml').documentElement;
      const num = new DOMParser().parseFromString(
        `<w:num xmlns:w="${W}" w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/></w:num>`, 'application/xml').documentElement;
      // Schema order: every abstractNum before the first num.
      const firstNum = rootEl.getElementsByTagNameNS(W, 'num')[0] ?? null;
      rootEl.insertBefore(numbering.importNode(abstract, true), firstNum);
      rootEl.append(numbering.importNode(num, true));
      replace.set(1_000_000 + n + 1, numId);
    });
    out.file('word/numbering.xml', new XMLSerializer().serializeToString(numbering), KEEP);
    // Point the new paragraphs at the real ids.
    let docXml = await out.file(this.docPath)!.async('string');
    for (const [placeholderId, real] of replace) docXml = docXml.replaceAll(`w:val="${placeholderId}"`, `w:val="${real}"`);
    out.file(this.docPath, docXml, KEEP);
    if (created) {
      const types = await out.file('[Content_Types].xml')!.async('string');
      if (!types.includes('/word/numbering.xml')) {
        out.file('[Content_Types].xml', types.replace('</Types>', `<Override PartName="/word/numbering.xml" ContentType="${NUMBERING_MIME}"/></Types>`), KEEP);
      }
      const relsPath = 'word/_rels/document.xml.rels';
      const rels = await out.file(relsPath)?.async('string');
      if (rels && !rels.includes(NUMBERING_TYPE)) {
        out.file(relsPath, rels.replace('</Relationships>', `<Relationship Id="rIdOeNumbering" Type="${NUMBERING_TYPE}" Target="numbering.xml"/></Relationships>`), KEEP);
      }
    }
  }

  private saved(result: SaveResult): void {
    if (result.ok && result.mode === 'overwrite') location.reload();
  }
}

// ------------------------------------------------------------------ helpers

function isPlain(p: Element): boolean {
  if (p.closest?.('txbxContent') || hasAncestor(p, 'txbxContent')) return false;
  for (const c of p.children) {
    if (!PLAIN_PARA.has(c.localName)) return false;
    if (c.localName === 'r') for (const rc of c.children) if (!PLAIN_RUN.has(rc.localName)) return false;
  }
  return true;
}

function hasAncestor(el: Element, name: string): boolean {
  for (let a = el.parentElement; a; a = a.parentElement) if (a.localName === name) return true;
  return false;
}

function child(el: Element | null | undefined, name: string): Element | null {
  if (!el) return null;
  for (const c of el.children) if (c.localName === name && c.namespaceURI === W) return c;
  return null;
}

function children(el: Element | null | undefined, name?: string): Element[] {
  return el ? [...el.children].filter((c) => !name || (c.localName === name && c.namespaceURI === W)) : [];
}

function bookmark(doc: Document, name: string): Element {
  const b = doc.createElementNS(W, 'w:bookmarkStart');
  b.setAttributeNS(W, 'w:id', '0');
  b.setAttributeNS(W, 'w:name', name);
  return b;
}

/** `<w:NAME w:val="value"/>` in `parent`, replacing an existing one. */
function setVal(doc: Document, parent: Element, name: string, value: string): void {
  let e = child(parent, name);
  if (!e) {
    e = doc.createElementNS(W, `w:${name}`);
    if (parent.localName === 'pPr') insertPPr(parent, e);
    else if (parent.localName === 'rPr') insertRPr(parent, e);
    else parent.append(e);
  }
  e.setAttributeNS(W, 'w:val', value);
}

/** A colour (hex) on w:color's val or w:u's color, replacing any theme colour there. */
function setColour(doc: Document, rPr: Element, name: 'color' | 'u', hex: string): void {
  let e = child(rPr, name);
  if (!e) {
    e = doc.createElementNS(W, `w:${name}`);
    insertRPr(rPr, e);
  }
  e.setAttributeNS(W, name === 'u' ? 'w:color' : 'w:val', hex);
  for (const a of ['themeColor', 'themeTint', 'themeShade']) e.removeAttributeNS(W, a);
}

/** A font size (points) on a span the user's selection was wrapped in; inner sizes give way. */
function setSpanSize(span: HTMLElement, pt: number): void {
  for (const e of span.querySelectorAll<HTMLElement>('[style]')) e.style.fontSize = '';
  for (const e of span.querySelectorAll<HTMLElement>('[data-sz]')) delete e.dataset.sz;
  span.dataset.sz = String(pt);
  span.style.fontSize = `${pt}pt`;
}

/** A font size (points) for a whole paragraph, and the sizes chosen inside it. */
function setParagraphSize(p: HTMLElement, pt: number): void {
  p.dataset.sz = String(pt);
  for (const e of p.querySelectorAll<HTMLElement>('span')) e.style.fontSize = `${pt}pt`;
  for (const e of p.querySelectorAll<HTMLElement>('[data-sz]')) e.dataset.sz = String(pt);
}

/** Wraps a range of a paragraph in a new span, which it returns. */
function wrapRange(range: Range): HTMLElement {
  const span = el('span');
  span.append(range.extractContents());
  range.insertNode(span);
  return span;
}

/**
 * Shows a chosen text or underline colour on an element and everything in it, and marks it for
 * saving (data-color / data-ucolor, read back in buildParagraph). Inner marks of the same kind give
 * way; docx-preview puts the document's own colours on the inner spans.
 */
function paint(e: HTMLElement, target: 'text' | 'underline', hex: string): void {
  const key = target === 'text' ? 'color' : 'ucolor';
  const css = `#${hex}`;
  for (const inner of e.querySelectorAll<HTMLElement>(`[data-${key}]`)) delete inner.dataset[key];
  e.dataset[key] = hex;
  for (const x of [e, ...e.querySelectorAll<HTMLElement>('*')]) {
    if (target === 'text') x.style.color = css;
    else x.style.textDecorationColor = css;
  }
  if (target === 'underline') e.style.textDecorationLine = 'underline';
}

/** An RGB colour (0–1) as Word writes it: RRGGBB. */
function hexOf(c: [number, number, number]): string {
  return c.map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function setToggle(doc: Document, rPr: Element, name: string, on: boolean): void {
  let e = child(rPr, name);
  if (!e) {
    e = doc.createElementNS(W, `w:${name}`);
    insertRPr(rPr, e);
  }
  if (on) e.removeAttributeNS(W, 'val');
  else e.setAttributeNS(W, 'w:val', '0');
}

/** pPr children have a fixed order in the schema; put the common ones where Word expects them. */
const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers',
  'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
  'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
  'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];

/** rPr children in schema order (Word may refuse a file with them out of order). */
const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow',
  'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs',
  'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];

function insertRPr(rPr: Element, el: Element): void {
  const rank = RPR_ORDER.indexOf(el.localName);
  const before = [...rPr.children].find((c) => RPR_ORDER.indexOf(c.localName) > rank);
  rPr.insertBefore(el, before ?? null);
}

function insertPPr(pPr: Element, el: Element): void {
  const rank = PPR_ORDER.indexOf(el.localName);
  const before = [...pPr.children].find((c) => PPR_ORDER.indexOf(c.localName) > rank);
  pPr.insertBefore(el, before ?? null);
}

function styleHasList(pStyle: Element | null): boolean {
  return !!pStyle && /list/i.test(pStyle.getAttributeNS(W, 'val') ?? '');
}

/** Bold / italic / underline as shown, for a text's element (underline can come from ancestors). */
function lookOf(elm: Element, stop?: Element): Look {
  const cs = getComputedStyle(elm);
  let u = false;
  for (let e: Element | null = elm; e && e !== stop?.parentElement; e = e.parentElement) {
    if (getComputedStyle(e).textDecorationLine.includes('underline')) u = true;
    if (e === stop) break;
  }
  return { b: Number(cs.fontWeight) >= 600, i: cs.fontStyle === 'italic', u };
}

function caretAtStart(p: HTMLElement): boolean {
  const sel = getSelection();
  if (!sel?.isCollapsed || !sel.rangeCount) return false;
  const r = document.createRange();
  r.setStart(p, 0);
  r.setEnd(sel.anchorNode!, sel.anchorOffset);
  return r.toString().length === 0;
}

function caretAtEnd(p: HTMLElement): boolean {
  const sel = getSelection();
  if (!sel?.isCollapsed || !sel.rangeCount) return false;
  const r = document.createRange();
  r.setStart(sel.anchorNode!, sel.anchorOffset);
  r.setEnd(p, p.childNodes.length);
  return r.toString().length === 0;
}

/** The editable paragraph just before / after, in the same container (no joins across cells). */
function previousEditable(p: HTMLElement): HTMLElement | null {
  const prev = p.previousElementSibling as HTMLElement | null;
  return prev?.tagName === 'P' && prev.isContentEditable ? prev : null;
}

function nextEditable(p: HTMLElement): HTMLElement | null {
  const next = p.nextElementSibling as HTMLElement | null;
  return next?.tagName === 'P' && next.isContentEditable ? next : null;
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className = ''): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  return e;
}

let toastEl: HTMLElement | null = null;
let toastTimer = 0;

function toast(text: string): void {
  toastEl ??= overlay().appendChild(el('div', 'docx-toast'));
  toastEl.textContent = text;
  toastEl.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => toastEl?.classList.remove('show'), 3500);
}

function dialogBox(title: string, body?: string): { box: HTMLElement; actions: HTMLElement; close: () => void } {
  const scrim = el('div', 'docx-dialog-scrim');
  const box = el('div', 'docx-dialog');
  const h = el('h2');
  h.textContent = title;
  box.append(h);
  if (body) {
    const p = el('p');
    p.textContent = body;
    box.append(p);
  }
  const actions = el('div', 'docx-dialog-actions');
  box.append(actions);
  scrim.append(box);
  overlay().append(scrim);
  return { box, actions, close: () => scrim.remove() };
}

function button(label: string, onClick: () => void, primary = false): HTMLButtonElement {
  const b = el('button', primary ? 'docx-text-button primary' : 'docx-text-button');
  b.textContent = label;
  b.onclick = onClick;
  return b;
}

function confirmDialog(title: string, body: string, ok: string): Promise<boolean> {
  return new Promise((resolve) => {
    const d = dialogBox(title, body);
    d.actions.append(
      button(t('cancel'), () => { d.close(); resolve(false); }),
      button(ok, () => { d.close(); resolve(true); }, true),
    );
  });
}

function menu(title: string, items: [string, string][]): Promise<string | null> {
  return new Promise((resolve) => {
    const d = dialogBox(title);
    const list = el('div', 'docx-dialog-menu');
    for (const [value, label] of items) list.append(button(label, () => { d.close(); resolve(value); }));
    d.box.insertBefore(list, d.actions);
    d.actions.append(button(t('cancel'), () => { d.close(); resolve(null); }));
  });
}

// Material Symbols paths (Apache-2.0).
const svg = (d: string) => `<svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><path fill="currentColor" d="${d}"/></svg>`;
const ICON_ALIGN = svg('M3 21h18v-2H3v2zm0-4h12v-2H3v2zm0-4h18v-2H3v2zm0-4h12V7H3v2zm0-6v2h18V3H3z');
// Bullets and numbers in one: Material Symbols "list" (Apache-2.0).
const ICON_LIST = svg('M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z');
const ICON_UNDERLINE = svg('M12 17c3.31 0 6-2.69 6-6V3h-2.5v8c0 1.93-1.57 3.5-3.5 3.5S8.5 12.93 8.5 11V3H6v8c0 3.31 2.69 6 6 6zm-7 2v2h14v-2H5z');
const ICON_TEXT = svg('M2.5 4v3h5v12h3V7h5V4h-13zm19 5h-9v3h3v7h3v-7h3V9z');
// Material Symbols "add_photo_alternate".
const ICON_PICTURE = svg('M19 7v2.99s-1.99.01-2 0V7h-3s.01-1.99 0-2h3V2h2v3h3v2h-3zm-3 4V8h-3V5H5c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2v-8h-3zM5 19l3-4 2 3 3-4 4 5H5z');
const ICON_UNDO = svg('M12.5 8c-2.65 0-5.05.99-6.9 2.6L2 7v9h9l-3.62-3.62c1.39-1.16 3.16-1.88 5.12-1.88 3.54 0 6.55 2.31 7.6 5.5l2.37-.78C21.08 11.03 17.15 8 12.5 8z');
const ICON_REDO = svg('M18.4 10.6C16.55 8.99 14.15 8 11.5 8c-4.65 0-8.58 3.03-9.96 7.22L3.9 16c1.05-3.19 4.05-5.5 7.6-5.5 1.95 0 3.73.72 5.12 1.88L13 16h9V7l-3.6 3.6z');
