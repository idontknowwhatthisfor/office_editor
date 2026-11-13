import './pdf.css';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import type { PDFDocumentProxy, PDFPageProxy, RenderTask } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { t } from '../i18n';
import { showMessage } from '../message';
import { PageBar, THUMB_WIDTH } from '../pagebar';
import { PdfEditor } from './editor';
import { PdfSearch, type TextContent } from './search';

// Runtime files copied to public/pdfjs by scripts/copy-pdf-assets.mjs.
const assets = new URL('pdfjs/', document.baseURI).href;
pdfjs.GlobalWorkerOptions.workerSrc = `${assets}pdf.worker.min.js`;

/** Upper bound on one page's canvas, in pixels, to keep memory in check when zoomed in. */
const MAX_CANVAS_PIXELS = 6_000_000;
const PAGE_MARGIN = 8;
const MAX_PAGE_WIDTH = 1100;

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

  private bar: PageBar;

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
      (state) => this.bar.showSearchState(state),
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
      () => this.bar.showRow(this.bar.main),
    );
    this.bar = new PageBar(
      {
        pageCount: pdf.numPages,
        aspect: (i) => this.pages[i].width / this.pages[i].height,
        renderThumb: (i, frame) => this.renderThumb(i, frame),
        goToPage: (i) => this.goToPage(i),
        find: (query) => void this.search.find(query),
        step: (direction) => this.search.step(direction),
        clearSearch: () => this.search.clear(),
        edit: () => {
          if (this.editor.start()) this.bar.showRow(this.editor.row);
        },
      },
      [this.editor.row],
    );
    root.replaceChildren(this.pagesEl, this.bar.el);
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
      // A pinch-zoom fires a stream of these; re-render sharper once it stops.
      let zoomTimer = 0;
      viewport.addEventListener('resize', () => {
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
    this.current = lo;
    this.bar.setCurrent(lo);
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

  // ---- thumbnails ----

  private async renderThumb(index: number, frame: HTMLElement): Promise<void> {
    const page = this.pages[index];
    const proxy = await this.proxy(page);
    const viewport = proxy.getViewport({ scale: (THUMB_WIDTH * Math.min(devicePixelRatio, 2)) / page.width });
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(viewport.width);
    canvas.height = Math.round(viewport.height);
    await proxy.render({ canvas, viewport }).promise;
    frame.append(canvas);
  }
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  return node;
}
