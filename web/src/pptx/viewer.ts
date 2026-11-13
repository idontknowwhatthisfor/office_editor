import '../docx/fonts.css';
import './pptx.css';
import { t } from '../i18n';
import { showMessage } from '../message';
import { ICON_EDIT, iconButton } from '../pagebar';
import { Deck } from './deck';
import type { PptxEditor } from './editor';
import { renderSlide, type RenderedSlide } from './render';

/**
 * "Export as PDF": one slide per page. Android's PDF printer only offers standard paper, so each
 * slide is scaled to fit landscape A4 and centred on it (as PowerPoint prints slides on paper).
 * Returns the page size in CSS px.
 */
export async function printPptx(root: HTMLElement, data: ArrayBuffer): Promise<{ width: number; height: number } | null> {
  let deck: Deck;
  try {
    deck = await Deck.open(data);
  } catch (e) {
    console.error('pptx open failed', e);
    return null;
  }
  const width = 1123;
  const height = 794;
  const scale = Math.min(width / deck.width, height / deck.height);
  const pages: HTMLElement[] = [];
  for (let i = 0; i < deck.slidePaths.length; i++) {
    const page = el('div', 'print-page');
    page.style.width = `${width}px`;
    page.style.height = `${height}px`;
    const holder = el('div');
    holder.style.cssText = `position:absolute;width:${deck.width}px;height:${deck.height}px;` +
      `left:${(width - deck.width * scale) / 2}px;top:${(height - deck.height * scale) / 2}px;` +
      `transform:scale(${scale});transform-origin:0 0`;
    try {
      holder.append((await renderSlide(deck, i)).el);
    } catch (e) {
      console.error(`slide ${i + 1} failed`, e);
    }
    page.append(holder);
    pages.push(page);
  }
  root.replaceChildren(...pages);
  return { width, height };
}

const GAP = 8;
const THUMB_H = 54;
const MAX_ZOOM = 5;
/** Slides kept attached on each side of the current one. */
const NEAR = 2;

/**
 * Read-only presentation viewer: one slide per screen, swipe to the next, thumbnails along the
 * bottom, pinch to zoom into a slide. Slides are drawn as HTML (render.ts), each only when it
 * comes near the screen.
 */
export async function openPptx(root: HTMLElement, data: ArrayBuffer, fileName = 'presentation.pptx'): Promise<void> {
  let deck: Deck;
  try {
    deck = await Deck.open(data);
    if (!deck.slidePaths.length) throw new Error('no slides');
  } catch (e) {
    console.error('pptx open failed', e);
    showMessage(root, t('pptxInvalid'));
    return;
  }
  // Slides zoom themselves, so the thumbnails stay usable; the page must not zoom.
  document.querySelector('meta[name=viewport]')
    ?.setAttribute('content', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  new PptxViewer(root, deck, fileName);
}

interface Page {
  el: HTMLDivElement;
  sizer: HTMLDivElement;
  slide?: RenderedSlide;
  zoom: number;
}

class PptxViewer {
  private view = el('div', 'pptx');
  private pager = el('div', 'pptx-pager');
  private editButton = iconButton(ICON_EDIT, t('edit'));
  private presentButton = iconButton(ICON_PRESENT, t('present'));
  private endButton = el('button', 'pptx-end');
  private presenting = false;
  private editor: PptxEditor | null = null;
  private startingEdit = false;
  private strip = el('div', 'pptx-strip');
  private count = el('div', 'pptx-count');
  private notice = el('div', 'pptx-notice');
  private noticeText = el('span');
  private pages: Page[] = [];
  private thumbs: HTMLButtonElement[] = [];
  private rendered = new Map<number, Promise<RenderedSlide | null>>();
  private current = -1;
  private noticeDismissed = false;
  private fit = 1;

  constructor(root: HTMLElement, private deck: Deck, private fileName: string) {
    const view = this.view;
    const close = el('button', 'pptx-notice-close');
    close.textContent = '×';
    close.setAttribute('aria-label', t('close'));
    close.onclick = () => {
      this.noticeDismissed = true;
      this.notice.hidden = true;
    };
    this.notice.append(this.noticeText, close);
    this.notice.hidden = true;
    this.strip.setAttribute('aria-label', t('slides'));

    const thumbW = Math.round((THUMB_H * deck.width) / deck.height);
    deck.slidePaths.forEach((_, i) => {
      const page = el('div', 'pptx-page');
      const sizer = el('div', 'pptx-sizer');
      page.append(sizer);
      this.pager.append(page);
      this.pages.push({ el: page, sizer, zoom: 1 });

      const thumb = el('button', 'pptx-thumb');
      thumb.style.width = `${thumbW}px`;
      thumb.style.height = `${THUMB_H}px`;
      thumb.setAttribute('aria-label', t('slideN', { n: i + 1 }));
      thumb.dataset.n = String(i + 1);
      thumb.onclick = () => this.go(i);
      this.strip.append(thumb);
      this.thumbs.push(thumb);
    });
    const bottom = el('div', 'pptx-bottom');
    this.editButton.classList.add('pptx-edit');
    this.editButton.onclick = () => void this.startEditing();
    this.presentButton.classList.add('pptx-edit');
    this.presentButton.onclick = () => this.startPresenting();
    this.endButton.innerHTML = `<svg viewBox="0 0 24 24" width="24" height="24" aria-hidden="true"><path fill="currentColor" d="${ICON_CLOSE}"/></svg>`;
    this.endButton.setAttribute('aria-label', t('endPresentation'));
    this.endButton.onclick = () => this.stopPresenting();
    bottom.append(this.strip, this.presentButton, this.editButton);
    view.append(this.notice, this.pager, this.count, bottom, this.endButton);
    root.replaceChildren(view);

    this.layout();
    this.listen();
    this.observeThumbs(thumbW);
    this.setCurrent(0);
  }

  private render(i: number): Promise<RenderedSlide | null> {
    let r = this.rendered.get(i);
    if (!r) {
      const started = performance.now();
      const editor = this.editor;
      r = renderSlide(this.deck, i, { edit: !!editor }).then(
        async (slide) => {
          console.log(`slide ${i + 1} drawn in ${Math.round(performance.now() - started)} ms`);
          if (editor) await editor.prepare(i, slide);
          return slide;
        },
        (e) => {
          console.error(`slide ${i + 1} failed`, e);
          return null;
        },
      );
      this.rendered.set(i, r);
    }
    return r;
  }

  /** Attaches slides near the current one, detaches the rest. */
  private async mountNear(): Promise<void> {
    const at = this.current;
    for (let i = 0; i < this.pages.length; i++) {
      const page = this.pages[i];
      if (Math.abs(i - at) > NEAR + 2 && page.slide) {
        page.sizer.replaceChildren();
        page.slide = undefined;
      }
    }
    for (const i of [at, at + 1, at - 1, at + 2, at - 2]) {
      if (i < 0 || i >= this.pages.length || Math.abs(i - at) > NEAR) continue;
      const page = this.pages[i];
      if (page.slide) continue;
      const slide = await this.render(i);
      if (this.current !== at && Math.abs(i - this.current) > NEAR) continue;
      if (!slide) {
        page.sizer.replaceChildren(failedBox());
        continue;
      }
      page.slide = slide;
      page.sizer.replaceChildren(slide.el);
      this.applyZoom(page);
      if (i === this.current) this.showNotice(slide);
    }
  }

  private setCurrent(i: number): void {
    if (i === this.current) return;
    this.editor?.flush();
    const old = this.pages[this.current];
    if (old && old.zoom !== 1) {
      old.zoom = 1;
      this.applyZoom(old);
    }
    this.current = i;
    this.count.textContent = t('slideOf', { n: i + 1, total: this.pages.length });
    this.thumbs.forEach((thumb, k) => thumb.setAttribute('aria-current', String(k === i)));
    this.thumbs[i].scrollIntoView({ block: 'nearest', inline: 'center', behavior: 'smooth' });
    this.notice.hidden = true;
    const page = this.pages[i];
    if (page.slide) this.showNotice(page.slide);
    void this.mountNear();
  }

  private showNotice(slide: RenderedSlide): void {
    if (this.noticeDismissed || !slide.missing.size) return;
    const names = [...slide.missing].map((kind) => t(MISSING[kind] ?? 'missingObject'));
    this.noticeText.textContent = t('notShown', { list: names.join(', ') });
    this.notice.hidden = false;
  }

  private go(i: number): void {
    // A pager that can't be swiped (editing) doesn't scroll smoothly: jump.
    this.pager.scrollTo({ left: i * this.pager.clientWidth, behavior: this.editor ? 'instant' : 'smooth' });
    // (and reports no scroll event)
    if (this.editor) this.setCurrent(i);
  }

  // ---------------------------------------------------------------- editing

  /** Edit mode: slides are drawn again with their text editable; swiping stops (thumbnails switch slides). */
  private async startEditing(): Promise<void> {
    if (this.editor || this.startingEdit) return;
    this.startingEdit = true;
    try {
      const { PptxEditor } = await import('./editor');
      this.deck.keepSlides = true;
      this.editor = new PptxEditor({
        deck: this.deck,
        fileName: this.fileName,
        pager: this.pager,
        current: () => this.current,
        go: (i) => this.go(i),
        redraw: (i) => this.redraw(i),
        slideEl: (i) => this.pages[i]?.slide?.el ?? null,
        refreshThumb: (i) => this.refreshThumb(i),
      });
    } catch (e) {
      console.error('could not start editing', e);
      return;
    } finally {
      this.startingEdit = false;
    }
    this.editButton.hidden = true;
    this.presentButton.hidden = true;
    this.view.classList.add('editing');
    this.view.append(this.editor.bar);
    this.rendered.clear();
    for (const page of this.pages) {
      page.sizer.replaceChildren();
      page.slide = undefined;
    }
    await this.mountNear();
  }

  private async redraw(i: number): Promise<void> {
    this.rendered.delete(i);
    const slide = await this.render(i);
    const page = this.pages[i];
    if (slide && (page.slide || Math.abs(i - this.current) <= NEAR)) {
      page.slide = slide;
      page.sizer.replaceChildren(slide.el);
      this.applyZoom(page);
    }
    this.refreshThumb(i);
  }

  /** Draws a thumbnail again from the slide on screen (or its drawing), if it was drawn before. */
  private refreshThumb(i: number): void {
    const thumb = this.thumbs[i];
    const old = thumb.querySelector('.slide');
    if (!old) return;
    const source = this.pages[i].slide?.el;
    const apply = (el: HTMLElement) => {
      const copy = el.cloneNode(true) as HTMLElement;
      copy.classList.remove('editing');
      for (const e of copy.querySelectorAll('[contenteditable]')) e.removeAttribute('contenteditable');
      copy.style.transform = old.getAttribute('style')?.match(/scale\([^)]*\)/)?.[0] ?? '';
      copy.setAttribute('aria-hidden', 'true');
      thumb.querySelector('.slide')?.replaceWith(copy);
    };
    if (source) apply(source);
    else void this.render(i).then((slide) => slide && apply(slide.el));
  }

  // ---------------------------------------------------------------- presenting

  /**
   * Full screen, one slide at a time on black: a tap on the right two thirds goes on, on the left
   * third goes back; swipes and arrow keys work as usual. In the app the shell hides its bars and
   * turns the screen landscape; Back ends it. In a browser it is the Fullscreen API (Esc ends it).
   */
  private startPresenting(): void {
    if (this.presenting || this.editor) return;
    this.presenting = true;
    const page = this.pages[this.current];
    if (page && page.zoom !== 1) {
      page.zoom = 1;
      this.applyZoom(page);
    }
    this.view.classList.add('presenting');
    window.officeEndPresentation = () => this.stopPresenting();
    if (window.OfficeBridge?.setPresenting) window.OfficeBridge.setPresenting(true);
    else document.documentElement.requestFullscreen?.().catch(() => undefined);
    this.layout();
  }

  private stopPresenting(): void {
    if (!this.presenting) return;
    this.presenting = false;
    this.view.classList.remove('presenting');
    if (window.OfficeBridge?.setPresenting) window.OfficeBridge.setPresenting(false);
    else if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
    this.layout();
    this.thumbs[this.current]?.scrollIntoView({ block: 'nearest', inline: 'center' });
  }

  // ---------------------------------------------------------------- sizing and zoom

  private layout(): void {
    const gap = this.presenting ? 0 : GAP;
    const w = this.pager.clientWidth - 2 * gap;
    const h = this.pager.clientHeight - 2 * gap;
    this.fit = Math.max(0.05, Math.min(w / this.deck.width, h / this.deck.height));
    for (const page of this.pages) this.applyZoom(page);
    if (this.current >= 0) this.pager.scrollLeft = this.current * this.pager.clientWidth;
  }

  private applyZoom(page: Page): void {
    const s = this.fit * page.zoom;
    page.sizer.style.width = `${this.deck.width * s}px`;
    page.sizer.style.height = `${this.deck.height * s}px`;
    // Edit mode draws outlines and prompts at screen size whatever the scale.
    page.sizer.style.setProperty('--k', String(1 / s));
    const slide = page.sizer.firstElementChild as HTMLElement | null;
    if (slide) slide.style.transform = `scale(${s})`;
    const zoomed = page.zoom > 1.001;
    page.el.classList.toggle('zoomed', zoomed);
    if (page === this.pages[this.current]) this.pager.classList.toggle('locked', zoomed);
  }

  /** Zooms the current slide to `z`, keeping the point (fx, fy) of the page where it is. */
  private zoomTo(z: number, fx: number, fy: number): void {
    const page = this.pages[this.current];
    if (!page) return;
    z = Math.min(MAX_ZOOM, Math.max(1, z));
    const el = page.el;
    const before = page.zoom;
    const px = (el.scrollLeft + fx - page.sizer.offsetLeft) / before;
    const py = (el.scrollTop + fy - page.sizer.offsetTop) / before;
    page.zoom = z < 1.02 ? 1 : z;
    this.applyZoom(page);
    el.scrollLeft = px * page.zoom + page.sizer.offsetLeft - fx;
    el.scrollTop = py * page.zoom + page.sizer.offsetTop - fy;
  }

  private listen(): void {
    let timer = 0;
    this.pager.addEventListener('scroll', () => {
      clearTimeout(timer);
      timer = window.setTimeout(() => {
        const i = Math.round(this.pager.scrollLeft / Math.max(1, this.pager.clientWidth));
        this.setCurrent(Math.min(this.pages.length - 1, Math.max(0, i)));
      }, 60);
    }, { passive: true });

    let width = 0;
    let height = 0;
    new ResizeObserver(() => {
      if (this.pager.clientWidth === width && this.pager.clientHeight === height) return;
      width = this.pager.clientWidth;
      height = this.pager.clientHeight;
      this.layout();
    }).observe(this.pager);

    window.addEventListener('keydown', (e) => {
      if ((e.target as Element).closest?.('[contenteditable="true"]')) return;
      if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') this.go(Math.min(this.pages.length - 1, this.current + 1));
      else if (e.key === 'ArrowLeft' || e.key === 'PageUp') this.go(Math.max(0, this.current - 1));
      else if (e.key === 'Escape' && this.presenting) this.stopPresenting();
      else return;
      e.preventDefault();
    });

    // Pinch: zoom the current slide around the point between the fingers.
    let pinch: { dist: number; zoom: number } | null = null;
    const dist = (t: TouchList) => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY);
    const focus = (t: TouchList) => {
      const rect = this.pages[this.current].el.getBoundingClientRect();
      return [(t[0].clientX + t[1].clientX) / 2 - rect.left, (t[0].clientY + t[1].clientY) / 2 - rect.top];
    };
    this.pager.addEventListener('touchstart', (e) => {
      if (e.touches.length !== 2) return;
      if (e.cancelable) e.preventDefault();
      pinch = { dist: dist(e.touches) || 1, zoom: this.pages[this.current].zoom };
    }, { passive: false });
    this.pager.addEventListener('touchmove', (e) => {
      if (!pinch || e.touches.length !== 2) return;
      if (e.cancelable) e.preventDefault();
      const [fx, fy] = focus(e.touches);
      this.zoomTo((pinch.zoom * dist(e.touches)) / pinch.dist, fx, fy);
    }, { passive: false });
    const end = (e: TouchEvent) => {
      if (e.touches.length < 2) pinch = null;
    };
    this.pager.addEventListener('touchend', end, { passive: true });
    this.pager.addEventListener('touchcancel', end, { passive: true });

    // Double tap (or double click): zoom in on that point, or back out.
    // Presenting: a tap moves through the slides.
    this.pager.addEventListener('click', (e) => {
      if (!this.presenting || this.pages[this.current]?.zoom > 1) return;
      const back = e.clientX < this.pager.clientWidth / 3;
      this.go(Math.min(this.pages.length - 1, Math.max(0, this.current + (back ? -1 : 1))));
    });
    document.addEventListener('fullscreenchange', () => {
      if (!document.fullscreenElement && this.presenting && !window.OfficeBridge?.setPresenting) this.stopPresenting();
    });
    this.pager.addEventListener('dblclick', (e) => {
      if (this.presenting) return;
      // Editing: a double tap selects a word.
      if ((e.target as Element).closest('[contenteditable="true"]')) return;
      const page = this.pages[this.current];
      const rect = page.el.getBoundingClientRect();
      this.zoomTo(page.zoom > 1 ? 1 : 2.5, e.clientX - rect.left, e.clientY - rect.top);
    });
    // Desktop: ctrl + wheel (and trackpad pinch).
    // While typing (keyboard up) the thumbnails make room for the slide.
    this.pager.addEventListener('focusin', (e) => {
      if ((e.target as Element).closest('[contenteditable="true"]')) this.view.classList.add('typing');
    });
    this.pager.addEventListener('focusout', (e) => {
      if (!(e.relatedTarget as Element | null)?.closest?.('[contenteditable="true"]')) this.view.classList.remove('typing');
    });
    this.pager.addEventListener('wheel', (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      const page = this.pages[this.current];
      const rect = page.el.getBoundingClientRect();
      this.zoomTo(page.zoom * Math.exp(-e.deltaY / 200), e.clientX - rect.left, e.clientY - rect.top);
    }, { passive: false });
  }

  // ---------------------------------------------------------------- thumbnails

  private observeThumbs(thumbW: number): void {
    const scale = thumbW / this.deck.width;
    const io = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const thumb = entry.target as HTMLButtonElement;
        io.unobserve(thumb);
        const i = Number(thumb.dataset.n) - 1;
        void this.render(i).then((slide) => {
          if (!slide) return;
          const copy = slide.el.cloneNode(true) as HTMLElement;
          copy.style.transform = `scale(${scale})`;
          copy.setAttribute('aria-hidden', 'true');
          thumb.append(copy);
        });
      }
    }, { root: this.strip, rootMargin: '0px 200px' });
    for (const thumb of this.thumbs) io.observe(thumb);
  }
}

// Material Symbols paths (Apache-2.0): slideshow, close.
const ICON_PRESENT = 'M10 8v8l5-4-5-4zm9-5H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14z';
const ICON_CLOSE = 'M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z';

declare global {
  interface Window {
    /** Called by the shell when Back is pressed during a presentation. */
    officeEndPresentation?: () => void;
  }
}

const MISSING: Record<string, Parameters<typeof t>[0]> = {
  chart: 'missingChart', diagram: 'missingDiagram', object: 'missingObject', picture: 'missingPicture', media: 'missingMedia',
};

function failedBox(): HTMLElement {
  const box = el('div', 'pptx-failed');
  box.textContent = t('slideInvalid');
  return box;
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className = ''): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  return e;
}
