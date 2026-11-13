import { parseAsync, renderDocument } from 'docx-preview';
import JSZip from 'jszip';
import { t } from '../i18n';
import { registerEditor, setDirty, type SaveResult } from '../save';
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
    this.view.append(...(await renderDocument(await parseAsync(bytes, options), options)));
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
    this.root.replaceChildren(this.frame, this.bar);
    await fontsLoaded(this.view);
    lineHeights.measure();
    for (const p of this.view.querySelectorAll<HTMLElement>('p[data-p]')) this.original.set(p.dataset.p!, p.outerHTML);
    this.fitToScreen();
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
    const add = (label: string, title: string, action: () => void, cls = '') => {
      const b = el('button', `docx-tool ${cls}`);
      b.innerHTML = label;
      b.title = title;
      b.setAttribute('aria-label', title);
      // Keep the text selection: don't let the button take focus.
      b.addEventListener('pointerdown', (e) => e.preventDefault());
      b.onclick = action;
      this.bar.append(b);
      return b;
    };
    add('<b>B</b>', t('bold'), () => this.format('bold'));
    add('<i>I</i>', t('italic'), () => this.format('italic'));
    add('<u>U</u>', t('underlineText'), () => this.format('underline'));
    add(ICON_ALIGN, t('alignment'), () => this.cycleAlign());
    add('H', t('headingStyle'), () => void this.headingMenu(), 'docx-tool-h');
    add(ICON_BULLETS, t('bullets'), () => this.setList('bullet'));
    add(ICON_NUMBERS, t('numbering'), () => this.setList('number'));
    add(ICON_UNDO, t('undo'), () => this.undo());
    add(ICON_REDO, t('redo'), () => this.redo());
  }

  private current(): HTMLElement | null {
    const node = getSelection()?.anchorNode;
    const p = (node instanceof Element ? node : node?.parentElement)?.closest<HTMLElement>('p[contenteditable="true"]');
    return p && this.view.contains(p) ? p : null;
  }

  private format(command: 'bold' | 'italic' | 'underline'): void {
    if (!this.current()) return toast(t('tapTextFirst'));
    document.execCommand('styleWithCSS', false, 'false');
    document.execCommand(command);
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

  private setList(kind: List): void {
    const p = this.current();
    if (!p) return toast(t('tapTextFirst'));
    const now = p.dataset.list ?? (p.className.includes('docx-num-') ? 'existing' : 'none');
    const next: List = now === kind ? 'none' : kind;
    p.dataset.list = next;
    p.classList.remove('oe-bullet', 'oe-number');
    for (const c of [...p.classList]) if (c.startsWith('docx-num-')) p.classList.remove(c);
    if (next !== 'none') p.classList.add(next === 'bullet' ? 'oe-bullet' : 'oe-number');
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

    for (const p of this.view.querySelectorAll<HTMLElement>('p[data-p]')) {
      const i = Number(p.dataset.p);
      const source = paras[i];
      if (!source) continue;
      const isNew = seen.has(i);
      seen.add(i);
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
    if (newPPr.childNodes.length) out.append(newPPr);
    for (const b of keep) out.append(b.cloneNode(true));

    // The text, in pieces with the same source run and the same bold/italic/underline.
    type Piece = { run: string | null; look: Look; text: string; kind: 'text' | 'tab' | 'br' };
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
        pieces.push({ run, look: lookOf(holder, p), text: tab ? '\t' : text.replace(/ /g, ' '), kind: tab ? 'tab' : 'text' });
        return;
      }
      if (node instanceof HTMLBRElement) {
        // The <br> an empty paragraph needs to be tappable is not a line break.
        if (node === p.lastChild || !node.nextSibling) return;
        pieces.push({ run: lastRun, look: lookOf(node.parentElement!, p), text: '', kind: 'br' });
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
      const key = `${piece.run}|${+piece.look.b}${+piece.look.i}${+piece.look.u}`;
      if (!current || current.key !== key) {
        const src = runById(piece.run);
        const r = doc.createElementNS(W, 'w:r');
        const rPr = (child(src.el, 'rPr')?.cloneNode(true) as Element | undefined) ?? doc.createElementNS(W, 'w:rPr');
        if (piece.look.b !== src.look.b) setToggle(doc, rPr, 'b', piece.look.b);
        if (piece.look.i !== src.look.i) setToggle(doc, rPr, 'i', piece.look.i);
        if (piece.look.u !== src.look.u) setVal(doc, rPr, 'u', piece.look.u ? 'single' : 'none');
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
    else parent.append(e);
  }
  e.setAttributeNS(W, 'w:val', value);
}

function setToggle(doc: Document, rPr: Element, name: string, on: boolean): void {
  let e = child(rPr, name);
  if (!e) {
    e = doc.createElementNS(W, `w:${name}`);
    rPr.append(e);
  }
  if (on) e.removeAttributeNS(W, 'val');
  else e.setAttributeNS(W, 'w:val', '0');
}

/** pPr children have a fixed order in the schema; put the common ones where Word expects them. */
const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers',
  'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
  'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
  'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];

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
  toastEl ??= document.body.appendChild(el('div', 'docx-toast'));
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
  document.body.append(scrim);
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
const ICON_BULLETS = svg('M4 10.5c-.83 0-1.5.67-1.5 1.5s.67 1.5 1.5 1.5 1.5-.67 1.5-1.5-.67-1.5-1.5-1.5zm0-6c-.83 0-1.5.67-1.5 1.5S3.17 7.5 4 7.5 5.5 6.83 5.5 6 4.83 4.5 4 4.5zm0 12c-.83 0-1.5.68-1.5 1.5s.68 1.5 1.5 1.5 1.5-.68 1.5-1.5-.67-1.5-1.5-1.5zM7 19h14v-2H7v2zm0-6h14v-2H7v2zm0-8v2h14V5H7z');
const ICON_NUMBERS = svg('M2 17h2v.5H3v1h1v.5H2v1h3v-4H2v1zm1-9h1V4H2v1h1v3zm-1 3h1.8L2 13.1v.9h3v-1H3.2L5 10.9V10H2v1zm5-6v2h14V5H7zm0 14h14v-2H7v2zm0-6h14v-2H7v2z');
const ICON_UNDO = svg('M12.5 8c-2.65 0-5.05.99-6.9 2.6L2 7v9h9l-3.62-3.62c1.39-1.16 3.16-1.88 5.12-1.88 3.54 0 6.55 2.31 7.6 5.5l2.37-.78C21.08 11.03 17.15 8 12.5 8z');
const ICON_REDO = svg('M18.4 10.6C16.55 8.99 14.15 8 11.5 8c-4.65 0-8.58 3.03-9.96 7.22L3.9 16c1.05-3.19 4.05-5.5 7.6-5.5 1.95 0 3.73.72 5.12 1.88L13 16h9V7l-3.6 3.6z');
