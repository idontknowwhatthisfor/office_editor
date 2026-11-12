/**
 * Splits docx-preview's output into real pages.
 *
 * docx-preview only starts a new page at explicit page breaks, so a chapter comes out as one very
 * tall "page". Word decides the other breaks by layout, which is what this does: fill each page
 * to its bottom margin, then carry the rest over to a copy of the page (same size, margins,
 * header and footer). Paragraphs are split between lines and tables between rows.
 *
 * Must run with the document attached, fonts loaded and no zoom applied: it measures layout.
 */
export function paginate(wrapper: HTMLElement): void {
  for (const section of wrapper.querySelectorAll<HTMLElement>(':scope > section')) {
    paginateSection(section);
  }
}

function paginateSection(first: HTMLElement): void {
  const style = getComputedStyle(first);
  const pageHeight = parseFloat(style.minHeight);
  const marginBottom = parseFloat(style.paddingBottom);
  if (!(pageHeight > 0)) return;
  first.style.height = style.minHeight;
  // Word drops a paragraph's "space before" at the top of a page (kept on the document's first).
  const opening = first.querySelector<HTMLElement>(':scope > article > :first-child');
  if (opening && first.previousElementSibling?.tagName === 'SECTION') opening.style.marginTop = '0';

  let page = first;
  for (let guard = 0; guard < 10_000; guard++) {
    const article = page.querySelector<HTMLElement>(':scope > article');
    if (!article) return;
    const limit = page.getBoundingClientRect().top + pageHeight - marginBottom;
    const blocks = [...article.children] as HTMLElement[];
    const index = blocks.findIndex((b) => b.getBoundingClientRect().bottom > limit + 0.5);
    if (index === -1) return;

    // What moves to the next page: the tail of a split block, or else the whole block.
    let moveFrom: Element | null = splitBlock(blocks[index], limit);
    if (!moveFrom) {
      if (index === 0) {
        moveFrom = blocks[1] ?? null; // taller than a page and unsplittable: leave it alone here
      } else {
        // Blocks marked "keep with next" (headings) go along with the block they precede.
        let from = index;
        while (from > 1 && blocks[from - 1].style.breakAfter === 'avoid') from--;
        moveFrom = blocks[from];
      }
    }
    if (!moveFrom) return;

    const next = page.cloneNode(false) as HTMLElement;
    for (const child of page.children) next.append(child.cloneNode(child !== article));
    const target = next.querySelector(':scope > article')!;
    (moveFrom as HTMLElement).style.marginTop = '0';
    while (moveFrom) {
      const following: Element | null = moveFrom.nextElementSibling;
      target.append(moveFrom);
      moveFrom = following;
    }
    page.after(next);
    page = next;
  }
}

/** Splits a block at `limit` (a viewport y). Returns the new element holding the rest, if split. */
function splitBlock(block: HTMLElement, limit: number): HTMLElement | null {
  if (block.tagName === 'P') return splitParagraph(block, limit);
  if (block.tagName === 'TABLE') return splitTable(block as HTMLTableElement, limit);
  return null;
}

function splitTable(table: HTMLTableElement, limit: number): HTMLElement | null {
  const rows = [...table.rows];
  const index = rows.findIndex((r) => r.getBoundingClientRect().bottom > limit + 0.5);
  // Header rows (marked break-after: avoid by the viewer) lead the table and repeat on each page.
  let headers = 0;
  while (headers < rows.length && rows[headers].style.breakAfter === 'avoid') headers++;
  if (index <= headers) return null; // no body row fits: the whole table moves
  const rest = table.cloneNode(false) as HTMLTableElement;
  const columns = table.querySelector(':scope > colgroup');
  if (columns) rest.append(columns.cloneNode(true));
  const body = document.createElement('tbody');
  body.append(...rows.slice(0, headers).map((row) => row.cloneNode(true)), ...rows.slice(index));
  rest.append(body);
  table.after(rest);
  return rest;
}

/**
 * Splits a paragraph between two lines. Like Word's default widow/orphan control, it never leaves
 * a single line behind or carries a single line over; then the paragraph moves whole instead.
 */
function splitParagraph(p: HTMLElement, limit: number): HTMLElement | null {
  const range = document.createRange();
  range.selectNodeContents(p);
  const lines = lineBoxes(range.getClientRects());
  let keep = lines.findIndex((line) => line.bottom > limit + 0.5);
  if (keep === -1) return null;
  keep = Math.min(keep, lines.length - 2);
  if (keep < 2) return null;
  const splitTop = lines[keep].top;

  // Find the first character on the first line that moves, by binary search over all characters.
  const nodes: Text[] = [];
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) if (walker.currentNode.nodeValue) nodes.push(walker.currentNode as Text);
  const total = nodes.reduce((n, t) => n + t.length, 0);
  const locate = (index: number): [Text, number] => {
    for (const node of nodes) {
      if (index < node.length) return [node, index];
      index -= node.length;
    }
    return [nodes[nodes.length - 1], nodes[nodes.length - 1].length];
  };
  /** Top of the line the character is on; characters without a box (collapsed) count as above. */
  const topOf = (index: number): number => {
    const [node, offset] = locate(index);
    range.setStart(node, offset);
    range.setEnd(node, offset + 1);
    const rect = range.getClientRects()[0];
    return rect ? rect.top : -Infinity;
  };
  let lo = 0;
  let hi = total; // first index whose line starts at or below splitTop
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (topOf(mid) >= splitTop - 1) hi = mid;
    else lo = mid + 1;
  }
  if (lo <= 0 || lo >= total) return null;

  const [node, offset] = locate(lo);
  range.setStart(node, offset);
  range.setEndAfter(p.lastChild!);
  const rest = p.cloneNode(false) as HTMLElement;
  rest.removeAttribute('id');
  rest.append(range.extractContents());
  // The second half is a continuation: no first-line indent, list number or space before it,
  // and the first half's last line is a full line, not a paragraph end.
  rest.classList.add('docx-continued');
  rest.style.textIndent = '0';
  rest.style.marginTop = '0';
  p.style.marginBottom = '0';
  if (getComputedStyle(p).textAlign === 'justify') p.style.textAlignLast = 'justify';
  p.after(rest);
  return rest;
}

/** Groups the boxes of a paragraph's inline content into lines, top to bottom. */
function lineBoxes(rects: DOMRectList): { top: number; bottom: number }[] {
  const lines: { top: number; bottom: number }[] = [];
  for (const rect of [...rects].sort((a, b) => a.top - b.top)) {
    if (rect.width === 0 && rect.height === 0) continue;
    const last = lines[lines.length - 1];
    // Boxes on one line overlap vertically (mixed font sizes); a new line starts below the last.
    if (last && rect.top < last.bottom - 1) last.bottom = Math.max(last.bottom, rect.bottom);
    else lines.push({ top: rect.top, bottom: rect.bottom });
  }
  return lines;
}
