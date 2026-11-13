import type { SearchState } from '../pagebar';

interface Match {
  page: number;
  range: Range;
}

/** Elements whose text is searched on its own, so a match never runs from one into the next. */
const BLOCK = 'p, li, td, th, h1, h2, h3, h4, h5, h6';

/**
 * Finds text in the Word viewer's pages and highlights it with the CSS Custom Highlight API, so
 * the document's markup is left alone. A match may span runs, but not paragraphs.
 */
export class DocxSearch {
  private matches: Match[] = [];
  private current = -1;
  private query = '';
  private all = typeof Highlight === 'function' ? new Highlight() : null;
  private selected = typeof Highlight === 'function' ? new Highlight() : null;

  constructor(
    private pages: HTMLElement[],
    private currentPage: () => number,
    private onChange: (state: SearchState) => void,
  ) {
    if (this.all && this.selected) {
      CSS.highlights.set('docx-match', this.all);
      CSS.highlights.set('docx-match-selected', this.selected);
    }
  }

  find(query: string): void {
    this.clear();
    this.query = query.trim();
    if (this.query) {
      const pattern = new RegExp(
        this.query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+'), 'giu',
      );
      this.pages.forEach((page, index) => {
        for (const block of blocks(page)) this.findIn(block, index, pattern);
      });
      for (const m of this.matches) this.all?.add(m.range);
      // Start from the page on screen, as the PDF viewer does.
      const here = this.currentPage();
      const first = this.matches.findIndex((m) => m.page >= here);
      if (this.matches.length) this.select(first < 0 ? 0 : first);
    }
    this.report();
  }

  step(direction: 1 | -1): void {
    if (!this.matches.length) return;
    this.select((this.current + direction + this.matches.length) % this.matches.length);
    this.report();
  }

  clear(): void {
    this.matches = [];
    this.current = -1;
    this.query = '';
    this.all?.clear();
    this.selected?.clear();
    this.report();
  }

  private findIn(block: Element, page: number, pattern: RegExp): void {
    const nodes: Text[] = [];
    const starts: number[] = [];
    let text = '';
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      // Nested blocks (a table inside a cell) are searched on their own.
      if (node.parentElement?.closest(BLOCK) !== block) continue;
      nodes.push(node as Text);
      starts.push(text.length);
      text += (node as Text).data;
    }
    const at = (offset: number, end: boolean): [Text, number] => {
      let i = nodes.length - 1;
      while (i > 0 && (end ? starts[i] >= offset : starts[i] > offset)) i--;
      return [nodes[i], offset - starts[i]];
    };
    for (const m of text.matchAll(pattern)) {
      if (!m[0]) continue;
      const range = document.createRange();
      range.setStart(...at(m.index, false));
      range.setEnd(...at(m.index + m[0].length, true));
      this.matches.push({ page, range });
    }
  }

  private select(index: number): void {
    this.current = index;
    const range = this.matches[index].range;
    this.selected?.clear();
    this.selected?.add(range);
    // Bring the match to the middle of the visible area (the visual viewport when zoomed).
    const rect = range.getBoundingClientRect();
    const viewport = window.visualViewport;
    const top = viewport?.offsetTop ?? 0;
    const height = viewport?.height ?? window.innerHeight;
    if (rect.top < top + 48 || rect.bottom > top + height - 120) {
      window.scrollBy(0, rect.top - top - height / 3);
    }
  }

  private report(): void {
    this.onChange({ query: this.query, current: this.current, total: this.matches.length, searching: false });
  }
}

function blocks(page: HTMLElement): Element[] {
  const found = [...page.querySelectorAll(BLOCK)];
  // Text that sits in no block at all (rare) is searched as part of the page.
  return found.length ? found : [page];
}
