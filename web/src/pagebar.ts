import './pagebar.css';
import { t } from './i18n';

/** What the bar needs from a paged viewer (PDF, Word). Page indexes are 0-based. */
export interface PageBarHost {
  pageCount: number;
  /** A page's width / height, for its thumbnail frame. */
  aspect(page: number): number;
  /** Draws page `page` into its thumbnail frame (THUMB_WIDTH CSS px wide). */
  renderThumb(page: number, frame: HTMLElement): Promise<void>;
  goToPage(page: number): void;
  /** Starts a search; the host reports progress with `showSearchState`. Empty clears it. */
  find(query: string): void;
  step(direction: 1 | -1): void;
  clearSearch(): void;
  edit(): void;
}

export interface SearchState {
  query: string;
  /** Index of the selected match, -1 if none. */
  current: number;
  total: number;
  searching: boolean;
}

export const THUMB_WIDTH = 72;

/**
 * The bottom bar of the paged viewers: page indicator and thumbnails, search, edit. It stays on
 * the glass while the page is pinch-zoomed.
 */
export class PageBar {
  readonly el = el('div', 'pagebar');
  readonly main = el('div', 'pagebar-row');
  private indicator = el('button', 'pagebar-indicator');
  private thumbsEl = el('div', 'pagebar-thumbs');
  private thumbs: HTMLButtonElement[] = [];
  private thumbQueue: Promise<void> = Promise.resolve();
  private searchInput = el('input', 'pagebar-search-input');
  private searchCount = el('span', 'pagebar-search-count');
  private current = 0;

  constructor(private host: PageBarHost, extraRows: HTMLElement[] = []) {
    const pagesButton = iconButton(ICON_PAGES, t('pages'));
    const searchButton = iconButton(ICON_SEARCH, t('search'));
    const editButton = iconButton(ICON_EDIT, t('edit'));
    editButton.onclick = () => host.edit();
    const toggleThumbs = () => this.toggleThumbs();
    pagesButton.onclick = toggleThumbs;
    this.indicator.onclick = toggleThumbs;
    this.indicator.setAttribute('aria-label', t('pages'));
    this.main.append(pagesButton, this.indicator, searchButton, editButton);

    const searchRow = el('div', 'pagebar-row');
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
      timer = window.setTimeout(() => host.find(this.searchInput.value), 250);
    };
    this.searchInput.onkeydown = (e) => {
      if (e.key !== 'Enter') return;
      this.searchInput.blur(); // hide the keyboard so the result is visible
      host.step(1);
    };
    previous.onclick = () => host.step(-1);
    next.onclick = () => host.step(1);
    searchButton.onclick = () => {
      this.showRow(searchRow);
      this.searchInput.focus();
    };
    close.onclick = () => {
      clearTimeout(timer);
      this.searchInput.value = '';
      host.clearSearch();
      this.showRow(this.main);
    };

    this.el.append(this.thumbsEl, this.main, searchRow, ...extraRows);
    this.pinToViewport();
  }

  /** Shows one row of the bar (the main row, or a host's own such as an edit toolbar). */
  showRow(row: HTMLElement): void {
    this.thumbsEl.classList.remove('open');
    for (const r of this.el.querySelectorAll<HTMLElement>(':scope > .pagebar-row')) r.hidden = r !== row;
  }

  setCurrent(page: number): void {
    this.thumbs[this.current]?.classList.remove('current');
    this.current = page;
    this.thumbs[page]?.classList.add('current');
    this.indicator.textContent = `${page + 1} / ${this.host.pageCount}`;
  }

  showSearchState(state: SearchState): void {
    this.searchCount.textContent =
      state.total > 0 ? `${state.current + 1}/${state.total}`
      : state.query === '' ? ''
      : state.searching ? '…'
      : t('noResults');
  }

  /**
   * A fixed element belongs to the layout viewport, so when the page is pinch-zoomed the bar
   * would be magnified and pushed off screen. Counter-transform it to stay put on the glass.
   */
  private pinToViewport(): void {
    const viewport = window.visualViewport;
    if (!viewport) return;
    const pin = () => {
      const below = document.documentElement.clientHeight - viewport.offsetTop - viewport.height;
      this.el.style.transform = `translate(${viewport.offsetLeft}px, ${-below}px) scale(${1 / viewport.scale})`;
    };
    viewport.addEventListener('scroll', pin);
    viewport.addEventListener('resize', pin);
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
          const frame = this.thumbs[index].firstElementChild as HTMLElement;
          this.thumbQueue = this.thumbQueue.then(async () => {
            try {
              await this.host.renderThumb(index, frame);
              frame.style.aspectRatio = String(this.host.aspect(index));
            } catch (e) {
              console.error(`thumbnail ${index + 1} failed`, e);
            }
          });
        }
      },
      { root: this.thumbsEl, rootMargin: '0px 200px' },
    );
    for (let index = 0; index < this.host.pageCount; index++) {
      const button = el('button', 'pagebar-thumb');
      button.setAttribute('aria-label', t('page', { n: index + 1 }));
      const frame = el('div', 'pagebar-thumb-frame');
      frame.style.aspectRatio = String(this.host.aspect(index));
      const number = el('span', '');
      number.textContent = String(index + 1);
      button.append(frame, number);
      button.onclick = () => {
        this.host.goToPage(index);
        this.thumbsEl.classList.remove('open');
      };
      this.thumbs.push(button);
      observer.observe(button);
    }
    this.thumbsEl.append(...this.thumbs);
    this.thumbs[this.current].classList.add('current');
  }
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  return node;
}

export function iconButton(path: string, label: string): HTMLButtonElement {
  const button = el('button', 'pagebar-icon');
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
