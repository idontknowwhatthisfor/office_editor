import type { PDFPageProxy } from 'pdfjs-dist/legacy/build/pdf.mjs';
import type { SearchState } from '../pagebar';

export type TextContent = Awaited<ReturnType<PDFPageProxy['getTextContent']>>;

/** What the search needs from the viewer. Page indexes are 0-based. */
export interface SearchHost {
  pageCount: number;
  currentPage(): number;
  pageText(page: number): Promise<TextContent>;
  /** The page's text-layer elements, one per text item, or null while it isn't rendered. */
  textDivs(page: number): HTMLElement[] | null;
  /** Scrolls to the page and resolves once its text layer is rendered. */
  reveal(page: number): Promise<void>;
}


interface Match {
  page: number;
  /** Offsets into the page's searchable text. */
  start: number;
  end: number;
}

/** A page's text as one string, with where each text item starts in it. */
interface PageIndex {
  text: string;
  strs: string[];
  starts: number[];
}

/**
 * Finds text across the document and highlights matches by wrapping them in <mark> inside the
 * text layer. Matches may span several text items (a PDF splits lines, and often words, into
 * separate items).
 */
export class PdfSearch {
  private index = new Map<number, PageIndex>();
  private matches: Match[] = [];
  private current = -1;
  private query = '';
  private searching = false;
  private run = 0;
  /** Text divs whose content was replaced with highlights, per page, so they can be restored. */
  private dirty = new Map<number, Set<number>>();

  constructor(
    private host: SearchHost,
    private onChange: (state: SearchState) => void,
  ) {}

  async find(query: string): Promise<void> {
    const run = ++this.run;
    this.query = query.trim();
    this.matches = [];
    this.current = -1;
    this.searching = this.query !== '';
    this.redecorateAll();
    this.notify();
    if (!this.searching) return;

    const escaped = this.query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const pattern = new RegExp(escaped, 'giu');
    // Search from the page being read to the end, then wrap to the start: in a long document
    // the nearest result shows up at once instead of after everything before it was read.
    const count = this.host.pageCount;
    const startPage = this.host.currentPage();
    let wrapped = 0; // matches found on pages before startPage; they sort ahead of the others
    for (let step = 0; step < count; step++) {
      const page = (startPage + step) % count;
      const { text } = await this.pageIndex(page);
      if (run !== this.run) return;
      const found: Match[] = [];
      for (const m of text.matchAll(pattern)) {
        if (m[0].length > 0) found.push({ page, start: m.index, end: m.index + m[0].length });
      }
      if (found.length === 0) continue;
      if (page >= startPage) {
        this.matches.push(...found);
      } else {
        this.matches.splice(wrapped, 0, ...found);
        wrapped += found.length;
        if (this.current >= 0) this.current += found.length;
      }
      if (this.current === -1) {
        void this.select(this.matches.indexOf(found[0]));
      } else {
        this.decorate(page);
        this.notify();
      }
    }
    this.searching = false;
    this.notify();
  }

  /** Moves to the next (1) or previous (-1) match, wrapping around. */
  step(direction: 1 | -1): void {
    const n = this.matches.length;
    if (n > 0) void this.select((this.current + direction + n) % n);
  }

  clear(): void {
    void this.find('');
  }

  /** Call when a page's text layer was (re)created: its divs are new and unhighlighted. */
  pageRendered(page: number): void {
    this.dirty.delete(page);
    this.decorate(page);
  }

  private async select(index: number): Promise<void> {
    const run = this.run;
    const previous = this.matches[this.current];
    this.current = index;
    const match = this.matches[index];
    this.notify();
    if (previous && previous.page !== match.page) this.decorate(previous.page);
    await this.host.reveal(match.page);
    // Compare the match, not the index: results found meanwhile can shift the indexes.
    if (run !== this.run || this.matches[this.current] !== match) return;
    this.decorate(match.page);
    const div = this.host.textDivs(match.page)?.find((d) => d.querySelector('mark.sel'));
    div?.querySelector('mark.sel')?.scrollIntoView({ block: 'center', inline: 'nearest' });
  }

  private notify(): void {
    this.onChange({
      query: this.query,
      current: this.current,
      total: this.matches.length,
      searching: this.searching,
    });
  }

  private async pageIndex(page: number): Promise<PageIndex> {
    let index = this.index.get(page);
    if (!index) {
      const content = await this.host.pageText(page);
      index = { text: '', strs: [], starts: [] };
      for (const item of content.items) {
        if (!('str' in item)) continue;
        index.starts.push(index.text.length);
        index.strs.push(item.str);
        index.text += item.str;
        // Line ends count as whitespace, so a phrase can match across a line break.
        if (item.hasEOL) index.text += '\n';
      }
      this.index.set(page, index);
    }
    return index;
  }

  private redecorateAll(): void {
    for (const page of [...this.dirty.keys()]) this.decorate(page);
  }

  /** Re-applies highlights to one page, if its text layer is rendered. */
  private decorate(page: number): void {
    const divs = this.host.textDivs(page);
    const index = this.index.get(page);
    if (!divs || !index) {
      this.dirty.delete(page);
      return;
    }
    for (const i of this.dirty.get(page) ?? []) {
      if (divs[i]) divs[i].textContent = index.strs[i];
    }

    // Split each match into the part of every text item it covers.
    const ranges = new Map<number, { from: number; to: number; selected: boolean }[]>();
    this.matches.forEach((m, matchIndex) => {
      if (m.page !== page) return;
      let i = itemAt(index.starts, m.start);
      for (; i < index.starts.length && index.starts[i] < m.end; i++) {
        const from = Math.max(m.start - index.starts[i], 0);
        const to = Math.min(m.end - index.starts[i], index.strs[i].length);
        if (to <= from) continue;
        let list = ranges.get(i);
        if (!list) ranges.set(i, (list = []));
        list.push({ from, to, selected: matchIndex === this.current });
      }
    });

    for (const [i, list] of ranges) {
      const div = divs[i];
      if (!div) continue;
      const str = index.strs[i];
      const parts: (string | Node)[] = [];
      let pos = 0;
      for (const r of list) {
        if (r.from < pos) continue; // overlapping matches can't happen, but never duplicate text
        parts.push(str.slice(pos, r.from));
        const mark = document.createElement('mark');
        mark.className = r.selected ? 'hl sel' : 'hl';
        mark.textContent = str.slice(r.from, r.to);
        parts.push(mark);
        pos = r.to;
      }
      parts.push(str.slice(pos));
      div.replaceChildren(...parts);
    }
    if (ranges.size > 0) this.dirty.set(page, new Set(ranges.keys()));
    else this.dirty.delete(page);
  }
}

/** Index of the last item starting at or before `offset`. */
function itemAt(starts: number[], offset: number): number {
  let lo = 0;
  let hi = starts.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (starts[mid] <= offset) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}
