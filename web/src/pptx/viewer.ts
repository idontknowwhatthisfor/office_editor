import '../docx/fonts.css';
import './pptx.css';
import { t } from '../i18n';
import { showMessage } from '../message';
import { Deck } from './deck';
import { renderSlide, type RenderedSlide } from './render';

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
export async function openPptx(root: HTMLElement, data: ArrayBuffer): Promise<void> {
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
  new PptxViewer(root, deck);
}

interface Page {
  el: HTMLDivElement;
  sizer: HTMLDivElement;
  slide?: RenderedSlide;
  zoom: number;
}

class PptxViewer {
  private pager = el('div', 'pptx-pager');
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

  constructor(root: HTMLElement, private deck: Deck) {
    const view = el('div', 'pptx');
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
    view.append(this.notice, this.pager, this.count, this.strip);
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
      r = renderSlide(this.deck, i).then(
        (slide) => {
          console.log(`slide ${i + 1} drawn in ${Math.round(performance.now() - started)} ms`);
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
    this.pager.scrollTo({ left: i * this.pager.clientWidth, behavior: 'smooth' });
  }

  // ---------------------------------------------------------------- sizing and zoom

  private layout(): void {
    const w = this.pager.clientWidth - 2 * GAP;
    const h = this.pager.clientHeight - 2 * GAP;
    this.fit = Math.max(0.05, Math.min(w / this.deck.width, h / this.deck.height));
    for (const page of this.pages) this.applyZoom(page);
    if (this.current >= 0) this.pager.scrollLeft = this.current * this.pager.clientWidth;
  }

  private applyZoom(page: Page): void {
    const s = this.fit * page.zoom;
    page.sizer.style.width = `${this.deck.width * s}px`;
    page.sizer.style.height = `${this.deck.height * s}px`;
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
      if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') this.go(Math.min(this.pages.length - 1, this.current + 1));
      else if (e.key === 'ArrowLeft' || e.key === 'PageUp') this.go(Math.max(0, this.current - 1));
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
    this.pager.addEventListener('dblclick', (e) => {
      const page = this.pages[this.current];
      const rect = page.el.getBoundingClientRect();
      this.zoomTo(page.zoom > 1 ? 1 : 2.5, e.clientX - rect.left, e.clientY - rect.top);
    });
    // Desktop: ctrl + wheel (and trackpad pinch).
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
