import './pdf.css';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import type { PDFDocumentProxy, PDFPageProxy, RenderTask } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { t } from '../i18n';
import { showMessage } from '../message';
import { PdfEditor } from './editor';
import { PdfSearch, type SearchState, type TextContent } from './search';

// Runtime files copied to public/pdfjs by scripts/copy-pdf-assets.mjs.
const assets = new URL('pdfjs/', document.baseURI).href;
pdfjs.GlobalWorkerOptions.workerSrc = `${assets}pdf.worker.min.js`;

/** Upper bound on one page's canvas, in pixels, to keep memory in check when zoomed in. */
const MAX_CANVAS_PIXELS = 6_000_000;
const PAGE_MARGIN = 8;
const MAX_PAGE_WIDTH = 1100;
const THUMB_WIDTH = 72;

interface PageView {
  index: number;
  el: HTMLDivElement;
  /** Size in PDF points (CSS px at scale 1), rotation applied. */
  width: number;
  height: number;
  proxy?: Promise<PDFPageProxy>;
  text?: Promise<TextContent>;
  canvas?: HTMLCanvasElement;
  /** Canvas pixels per PDF point of the current canvas; 0 when there is none. */
  canvasScale: number;
  textEl?: HTMLDivElement;
  textDivs?: HTMLElement[];
  task?: RenderTask;
  rendering?: Promise<void>;
  failed?: boolean;
}

/** Pages above this count are measured only as they come near the screen (see measureAll). */
const MEASURE_ALL_LIMIT = 300;

/**
 * Opens a PDF from a same-origin URL that supports HTTP range requests, or from bytes already in
 * memory. From a URL only the parts of the file that are needed are read (256 KB at a time), so
 * a 700 MB book costs little memory until it is read through.
 */
export async function openPdf(
  root: HTMLElement, source: { url: string } | { data: ArrayBuffer }, file = { name: 'document.pdf', size: -1 },
): Promise<void> {
  let pdf: PDFDocumentProxy;
  let first: PDFPageProxy;
  try {
    pdf = await pdfjs.getDocument({
      ...source,
      rangeChunkSize: 262144,
      disableAutoFetch: true,
      disableStream: true,
      cMapUrl: `${assets}cmaps/`,
      standardFontDataUrl: `${assets}standard_fonts/`,
      wasmUrl: `${assets}wasm/`,
      iccUrl: `${assets}iccs/`,
    }).promise;
    first = await pdf.getPage(1);
  } catch (e) {
    console.error('pdf open failed:', String(e), e);
    const key = e instanceof pdfjs.PasswordException ? 'pdfPassword'
      : e instanceof pdfjs.InvalidPDFException ? 'pdfInvalid'
      : 'loadFailed';
    showMessage(root, t(key));
    return;
  }
  // Forms are filled in with HTML inputs over the page (editor.ts), so the canvas leaves them out.
  let hasForms = false;
  try {
    // (getFieldObjects() comes back empty for some forms; the document info flag doesn't.)
    const { info } = await pdf.getMetadata();
    hasForms = !!(info as { IsAcroFormPresent?: boolean }).IsAcroFormPresent;
  } catch {
    // No usable form.
  }
  new PdfViewer(root, pdf, first, hasForms, file);
}

/**
 * Continuous-scroll PDF viewer. Only pages near the viewport hold a canvas and a text layer;
 * pinch-zoom is the WebView's own, and visible pages are re-rendered sharper once a zoom settles.
 */
class PdfViewer {
  private pages: PageView[] = [];
  private pagesEl = el('div', 'pdf-pages');
  private pageWidth = 0;
  private layoutWidth = 0;
  private current = 0;
  /** Pages close enough to the viewport to be kept rendered. */
  private wanted = new Set<number>();
  private pumping = false;
  private search: PdfSearch;

  private indicator = el('button', 'pdf-indicator');
  private thumbsEl = el('div', 'pdf-thumbs');
  private thumbs: HTMLButtonElement[] = [];
  private thumbQueue: Promise<void> = Promise.resolve();
  private searchInput = el('input', 'pdf-search-input');
  private searchCount = el('span', 'pdf-search-count');

  private editor: PdfEditor;

  constructor(
    root: HTMLElement,
    private pdf: PDFDocumentProxy,
    first: PDFPageProxy,
    private hasForms: boolean,
    file: { name: string; size: number },
  ) {
    // Until each page is measured (measureAll), assume they're all the size of the first.
    const size = first.getViewport({ scale: 1 });
    for (let index = 0; index < pdf.numPages; index++) {
      const pageEl = el('div', 'pdf-page');
      pageEl.style.setProperty('--scale-round-x', '1px');
      pageEl.style.setProperty('--scale-round-y', '1px');
      this.pages.push({ index, el: pageEl, width: size.width, height: size.height, canvasScale: 0 });
    }
    this.pages[0].proxy = Promise.resolve(first);
    this.pagesEl.append(...this.pages.map((p) => p.el));

    this.search = new PdfSearch(
      {
        pageCount: pdf.numPages,
        currentPage: () => this.current,
        pageText: (page) => this.text(this.pages[page]),
        textDivs: (page) => this.pages[page].textDivs ?? null,
        reveal: (page) => this.reveal(page),
      },
      (state) => this.showSearchState(state),
    );

    this.editor = new PdfEditor(
      {
        pdf,
        fileName: file.name,
        size: file.size,
        currentPage: () => this.current,
        reload: () => location.reload(),
      },
      hasForms,
      () => this.showMainRow(),
    );
    const bar = this.buildBar();
    root.replaceChildren(this.pagesEl, bar);
    this.layout();

    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          const page = this.pages[Number((entry.target as HTMLElement).dataset.index)];
          if (entry.isIntersecting) this.wanted.add(page.index);
          else this.release(page);
        }
        void this.pump();
      },
      { rootMargin: '100% 0px' },
    );
    this.pages.forEach((p) => {
      p.el.dataset.index = String(p.index);
      observer.observe(p.el);
    });

    let frame = 0;
    window.addEventListener('scroll', () => {
      if (!frame) frame = requestAnimationFrame(() => {
        frame = 0;
        this.updateCurrent();
      });
    }, { passive: true });

    window.addEventListener('resize', () => {
      // The keyboard only changes the height; pages are laid out by width.
      if (document.documentElement.clientWidth !== this.layoutWidth) {
        this.keepPosition(() => this.layout());
        void this.pump();
      }
    });

    const viewport = window.visualViewport;
    if (viewport) {
      // A fixed element belongs to the layout viewport, so when the page is pinch-zoomed the bar
      // would be magnified and pushed off screen. Counter-transform it to stay put on the glass.
      const pinBar = () => {
        const below = document.documentElement.clientHeight - viewport.offsetTop - viewport.height;
        bar.style.transform =
          `translate(${viewport.offsetLeft}px, ${-below}px) scale(${1 / viewport.scale})`;
      };
      viewport.addEventListener('scroll', pinBar);
      // A pinch-zoom fires a stream of these; re-render sharper once it stops.
      let zoomTimer = 0;
      viewport.addEventListener('resize', () => {
        pinBar();
        clearTimeout(zoomTimer);
        zoomTimer = window.setTimeout(() => void this.pump(), 200);
      });
    }

    this.updateCurrent();
    void this.measureAll();
  }

  // ---- layout ----

  private layout(): void {
    this.layoutWidth = document.documentElement.clientWidth;
    this.pageWidth = Math.min(this.layoutWidth - 2 * PAGE_MARGIN, MAX_PAGE_WIDTH);
    for (const page of this.pages) this.sizePage(page);
  }

  private sizePage(page: PageView): void {
    const scale = this.pageWidth / page.width;
    page.el.style.width = `${this.pageWidth}px`;
    page.el.style.height = `${page.height * scale}px`;
    // The text layer positions and sizes its text from this.
    page.el.style.setProperty('--total-scale-factor', String(scale));
  }

  /** Runs a layout change while keeping the same spot of the current page at the top. */
  private keepPosition(change: () => void): void {
    const pageEl = this.pages[this.current].el;
    const before = pageEl.getBoundingClientRect();
    const fraction = -before.top / before.height;
    change();
    const after = pageEl.getBoundingClientRect();
    window.scrollBy(0, after.top + fraction * after.height);
  }

  /**
   * Pages can differ in size; correct the first-page assumption in the background. Not for long
   * documents: looking at every page would read much of the file. Their pages are measured when
   * they are about to be rendered instead (measure, from doRender).
   */
  private async measureAll(): Promise<void> {
    if (this.pages.length > MEASURE_ALL_LIMIT) return;
    for (const page of this.pages.slice(1)) this.measure(page, await this.proxy(page));
  }

  private measure(page: PageView, proxy: PDFPageProxy): void {
    const size = proxy.getViewport({ scale: 1 });
    if (Math.abs(size.width - page.width) < 0.5 && Math.abs(size.height - page.height) < 0.5) return;
    this.keepPosition(() => {
      page.width = size.width;
      page.height = size.height;
      this.sizePage(page);
    });
    if (page.canvas) {
      // Rendered with the wrong shape; do it again.
      this.release(page);
      this.wanted.add(page.index);
      void this.pump();
    }
  }

  private updateCurrent(): void {
    // The middle of what's on screen, which is the visual viewport when zoomed in.
    const viewport = window.visualViewport;
    const middle = viewport ? viewport.offsetTop + viewport.height / 2 : window.innerHeight / 2;
    let lo = 0;
    let hi = this.pages.length - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.pages[mid].el.getBoundingClientRect().bottom < middle) lo = mid + 1;
      else hi = mid;
    }
    this.thumbs[this.current]?.classList.remove('current');
    this.current = lo;
    this.thumbs[lo]?.classList.add('current');
    this.indicator.textContent = `${lo + 1} / ${this.pages.length}`;
  }

  private goToPage(index: number): void {
    const top = this.pages[index].el.getBoundingClientRect().top;
    window.scrollBy(0, top - (window.visualViewport?.offsetTop ?? 0) - PAGE_MARGIN);
  }

  // ---- rendering ----

  private proxy(page: PageView): Promise<PDFPageProxy> {
    return (page.proxy ??= this.pdf.getPage(page.index + 1));
  }

  private text(page: PageView): Promise<TextContent> {
    return (page.text ??= this.proxy(page).then((p) => p.getTextContent()));
  }

  /** Canvas pixels per PDF point the page should be rendered at right now. */
  private targetScale(page: PageView): number {
    const viewport = window.visualViewport;
    let scale = (this.pageWidth / page.width) * devicePixelRatio;
    // Only pages actually on screen are worth the memory of a zoomed-in render.
    if (viewport && viewport.scale > 1) {
      const rect = page.el.getBoundingClientRect();
      const onScreen = rect.bottom > viewport.offsetTop && rect.top < viewport.offsetTop + viewport.height;
      if (onScreen) scale *= viewport.scale;
    }
    const pixels = page.width * page.height * scale * scale;
    if (pixels > MAX_CANVAS_PIXELS) scale *= Math.sqrt(MAX_CANVAS_PIXELS / pixels);
    return scale;
  }

  private needsRender(page: PageView): boolean {
    return !page.failed && (!page.canvas || this.targetScale(page) > page.canvasScale * 1.1);
  }

  /** Renders wanted pages one at a time, nearest to the current page first. */
  private async pump(): Promise<void> {
    if (this.pumping) return;
    this.pumping = true;
    try {
      for (;;) {
        let next: PageView | undefined;
        for (const index of this.wanted) {
          const page = this.pages[index];
          if (!this.needsRender(page)) continue;
          if (!next || Math.abs(index - this.current) < Math.abs(next.index - this.current)) next = page;
        }
        if (!next) break;
        await this.renderPage(next);
      }
    } finally {
      this.pumping = false;
    }
  }

  private renderPage(page: PageView): Promise<void> {
    return (page.rendering ??= this.doRender(page).finally(() => (page.rendering = undefined)));
  }

  private async doRender(page: PageView): Promise<void> {
    try {
      const proxy = await this.proxy(page);
      this.measure(page, proxy);
      const scale = this.targetScale(page);
      const viewport = proxy.getViewport({ scale });
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(viewport.width);
      canvas.height = Math.round(viewport.height);
      page.task = proxy.render({
        canvas,
        viewport,
        annotationMode: this.hasForms ? pdfjs.AnnotationMode.ENABLE_FORMS : pdfjs.AnnotationMode.ENABLE,
      });
      await page.task.promise;
      page.task = undefined;
      if (!this.wanted.has(page.index)) return;

      // Swap only once the new canvas is complete, so a zoom re-render never flashes blank.
      page.canvas?.remove();
      page.canvas = canvas;
      page.canvasScale = scale;
      page.el.prepend(canvas);

      if (!page.textEl) {
        const textEl = el('div', 'textLayer');
        const layer = new pdfjs.TextLayer({
          textContentSource: await this.text(page),
          container: textEl,
          viewport: proxy.getViewport({ scale: 1 }),
        });
        await layer.render();
        if (!this.wanted.has(page.index) || page.textEl) return;
        page.textEl = textEl;
        page.textDivs = layer.textDivs;
        page.el.append(textEl);
        this.search.pageRendered(page.index);
      }
      void this.editor.pageReady(page.index, page.el, proxy);
    } catch (e) {
      page.task = undefined;
      if (e instanceof pdfjs.RenderingCancelledException) return;
      console.error(`page ${page.index + 1} failed to render`, e);
      page.failed = true;
    }
  }

  /** Frees a page that scrolled far away. */
  private release(page: PageView): void {
    this.wanted.delete(page.index);
    page.task?.cancel();
    if (page.canvas) {
      page.canvas.remove();
      page.canvas.width = page.canvas.height = 0; // release the bitmap now, not at GC
      page.canvas = undefined;
      page.canvasScale = 0;
    }
    page.textEl?.remove();
    page.textEl = undefined;
    page.textDivs = undefined;
    void page.proxy?.then((p) => p.cleanup());
  }

  private async reveal(index: number): Promise<void> {
    const page = this.pages[index];
    const rect = page.el.getBoundingClientRect();
    if (rect.bottom < 0 || rect.top > window.innerHeight) this.goToPage(index);
    this.wanted.add(index);
    if (!page.textEl) await this.renderPage(page);
    void this.pump();
  }

  // ---- bottom bar, thumbnails, search ----

  private buildBar(): HTMLElement {
    const bar = el('div', 'pdf-bar');

    const main = el('div', 'pdf-bar-row');
    const pagesButton = iconButton(ICON_PAGES, t('pages'));
    const searchButton = iconButton(ICON_SEARCH, t('search'));
    const editButton = iconButton(ICON_EDIT, t('edit'));
    editButton.onclick = () => {
      if (!this.editor.start()) return;
      this.thumbsEl.classList.remove('open');
      main.hidden = true;
      this.editor.row.hidden = false;
    };
    this.showMainRow = () => {
      this.editor.row.hidden = true;
      main.hidden = false;
    };
    const toggleThumbs = () => this.toggleThumbs();
    pagesButton.onclick = toggleThumbs;
    this.indicator.onclick = toggleThumbs;
    this.indicator.setAttribute('aria-label', t('pages'));
    main.append(pagesButton, this.indicator, searchButton, editButton);

    const searchRow = el('div', 'pdf-bar-row');
    searchRow.hidden = true;
    this.searchInput.type = 'search';
    this.searchInput.placeholder = t('searchHint');
    this.searchInput.enterKeyHint = 'search';
    this.searchInput.autocapitalize = 'off';
    this.searchInput.setAttribute('aria-label', t('searchHint'));
    const previous = iconButton(ICON_UP, t('previous'));
    const next = iconButton(ICON_DOWN, t('next'));
    const close = iconButton(ICON_CLOSE, t('close'));
    searchRow.append(this.searchInput, this.searchCount, previous, next, close);

    let timer = 0;
    this.searchInput.oninput = () => {
      clearTimeout(timer);
      timer = window.setTimeout(() => void this.search.find(this.searchInput.value), 250);
    };
    this.searchInput.onkeydown = (e) => {
      if (e.key !== 'Enter') return;
      this.searchInput.blur(); // hide the keyboard so the result is visible
      this.search.step(1);
    };
    previous.onclick = () => this.search.step(-1);
    next.onclick = () => this.search.step(1);
    searchButton.onclick = () => {
      this.thumbsEl.classList.remove('open');
      main.hidden = true;
      searchRow.hidden = false;
      this.searchInput.focus();
    };
    close.onclick = () => {
      clearTimeout(timer);
      this.searchInput.value = '';
      this.search.clear();
      searchRow.hidden = true;
      main.hidden = false;
    };

    bar.append(this.thumbsEl, main, searchRow, this.editor.row);
    return bar;
  }

  private showMainRow: () => void = () => undefined;

  private showSearchState(state: SearchState): void {
    this.searchCount.textContent =
      state.total > 0 ? `${state.current + 1}/${state.total}`
      : state.query === '' ? ''
      : state.searching ? '…'
      : t('noResults');
  }

  private toggleThumbs(): void {
    if (this.thumbsEl.classList.toggle('open')) {
      if (this.thumbs.length === 0) this.buildThumbs();
      this.thumbs[this.current].scrollIntoView({ inline: 'center', block: 'nearest' });
    }
  }

  private buildThumbs(): void {
    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          observer.unobserve(entry.target);
          const index = this.thumbs.indexOf(entry.target as HTMLButtonElement);
          this.thumbQueue = this.thumbQueue.then(() => this.renderThumb(index));
        }
      },
      { root: this.thumbsEl, rootMargin: '0px 200px' },
    );
    this.pages.forEach((page) => {
      const button = el('button', 'pdf-thumb');
      button.setAttribute('aria-label', t('page', { n: page.index + 1 }));
      const frame = el('div', 'pdf-thumb-frame');
      frame.style.aspectRatio = `${page.width} / ${page.height}`;
      const number = el('span', '');
      number.textContent = String(page.index + 1);
      button.append(frame, number);
      button.onclick = () => {
        this.goToPage(page.index);
        this.thumbsEl.classList.remove('open');
      };
      this.thumbs.push(button);
      observer.observe(button);
    });
    this.thumbsEl.append(...this.thumbs);
    this.thumbs[this.current].classList.add('current');
  }

  private async renderThumb(index: number): Promise<void> {
    try {
      const page = this.pages[index];
      const proxy = await this.proxy(page);
      const viewport = proxy.getViewport({
        scale: (THUMB_WIDTH * Math.min(devicePixelRatio, 2)) / page.width,
      });
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(viewport.width);
      canvas.height = Math.round(viewport.height);
      await proxy.render({ canvas, viewport }).promise;
      const frame = this.thumbs[index].firstElementChild as HTMLElement;
      frame.style.aspectRatio = `${page.width} / ${page.height}`;
      frame.append(canvas);
    } catch (e) {
      console.error(`thumbnail ${index + 1} failed`, e);
    }
  }
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  return node;
}

function iconButton(path: string, label: string): HTMLButtonElement {
  const button = el('button', 'pdf-icon-button');
  button.setAttribute('aria-label', label);
  button.innerHTML = `<svg viewBox="0 0 24 24" width="24" height="24" aria-hidden="true"><path fill="currentColor" d="${path}"/></svg>`;
  return button;
}

// Material Symbols paths (Apache-2.0).
const ICON_SEARCH = 'M15.5 14h-.79l-.28-.27A6.47 6.47 0 0 0 16 9.5 6.5 6.5 0 1 0 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z';
const ICON_PAGES = 'M4 8h4V4H4v4zm6 12h4v-4h-4v4zm-6 0h4v-4H4v4zm0-6h4v-4H4v4zm6 0h4v-4h-4v4zm6-10v4h4V4h-4zm-6 4h4V4h-4v4zm6 6h4v-4h-4v4zm0 6h4v-4h-4v4z';
const ICON_UP = 'M7.41 15.41 12 10.83l4.59 4.58L18 14l-6-6-6 6z';
const ICON_DOWN = 'M7.41 8.59 12 13.17l4.59-4.58L18 10l-6 6-6-6z';
const ICON_EDIT = 'M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04a.996.996 0 0 0 0-1.41l-2.34-2.34a.996.996 0 0 0-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z';
const ICON_CLOSE = 'M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z';
