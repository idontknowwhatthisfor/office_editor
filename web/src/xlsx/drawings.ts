import { Package, child, children, first, relAttr, relOfType } from '../ooxml';

// A sheet's drawing part (xl/drawings/drawingN.xml): where each chart sits. Charts are anchored
// to cells — from a cell (plus an offset) to a cell, from a cell with a size, or at an absolute
// position (chart sheets) — so they move with the grid's column widths and row heights.

const EMU = 9525;

/** A cell corner: column/row index plus an offset into it, in CSS px. */
export interface Marker {
  col: number;
  row: number;
  dx: number;
  dy: number;
}

export interface SheetChart {
  /** The chart part's root element (DrawingML chart or chartex). */
  root: Element;
  from: Marker | null;
  to: Marker | null;
  /** Size in px (one-cell and absolute anchors). */
  size: { w: number; h: number } | null;
  /** Position in px (absolute anchors). */
  pos: { x: number; y: number } | null;
}

export interface Drawings {
  charts: SheetChart[];
  /** Pictures, shapes and charts that can't be read: not drawn. */
  others: number;
}

/** The drawing a sheet (or chart sheet) at `sheetPath` refers to with `<drawing r:id>`. */
export async function readDrawings(pkg: Package, sheetPath: string, drawingId: string): Promise<Drawings> {
  const out: Drawings = { charts: [], others: 0 };
  const sheetRels = await pkg.rels(sheetPath);
  const rel = sheetRels.get(drawingId) ?? relOfType(sheetRels, '/drawing');
  if (!rel || rel.external) return out;
  const doc = await pkg.xml(rel.target);
  if (!doc) return out;
  const rels = await pkg.rels(rel.target);
  for (const anchor of children(doc.documentElement)) {
    if (!/^(twoCellAnchor|oneCellAnchor|absoluteAnchor)$/.test(anchor.localName)) continue;
    const frame = graphicFrame(anchor);
    const ref = frame && !hidden(frame) ? first(frame, 'chart') : null;
    const chartRel = ref ? rels.get(relAttr(ref)) : undefined;
    const part = chartRel && !chartRel.external ? await pkg.xml(chartRel.target) : null;
    if (!part) {
      if (!anchorHidden(anchor)) out.others++;
      continue;
    }
    const ext = child(anchor, 'ext');
    const pos = child(anchor, 'pos');
    out.charts.push({
      root: part.documentElement,
      from: marker(child(anchor, 'from')),
      to: anchor.localName === 'twoCellAnchor' ? marker(child(anchor, 'to')) : null,
      size: ext ? { w: Number(ext.getAttribute('cx') ?? 0) / EMU, h: Number(ext.getAttribute('cy') ?? 0) / EMU } : null,
      pos: pos ? { x: Number(pos.getAttribute('x') ?? 0) / EMU, y: Number(pos.getAttribute('y') ?? 0) / EMU } : null,
    });
  }
  return out;
}

/** The anchor's chart frame; newer charts (chartex) sit in an mc:Choice with a fallback shape. */
function graphicFrame(anchor: Element): Element | null {
  const direct = child(anchor, 'graphicFrame');
  if (direct) return direct;
  const alt = child(anchor, 'AlternateContent');
  for (const choice of children(alt, 'Choice')) {
    const frame = child(choice, 'graphicFrame');
    if (frame) return frame;
  }
  return null;
}

function hidden(frame: Element): boolean {
  const v = first(frame, 'cNvPr')?.getAttribute('hidden');
  return v === '1' || v === 'true';
}

function anchorHidden(anchor: Element): boolean {
  const v = first(anchor, 'cNvPr')?.getAttribute('hidden');
  return v === '1' || v === 'true';
}

function marker(el: Element | null): Marker | null {
  if (!el) return null;
  const n = (name: string) => Number(child(el, name)?.textContent ?? 0) || 0;
  return { col: n('col'), row: n('row'), dx: n('colOff') / EMU, dy: n('rowOff') / EMU };
}

/** The last column and row a chart reaches (so the grid extends that far). */
export function chartExtent(c: SheetChart): { col: number; row: number } {
  const end = c.to ?? c.from;
  return { col: end?.col ?? 0, row: end?.row ?? 0 };
}

/**
 * A chart's box in sheet px, given the grid's column and row edges (`colX[c]` = left of column
 * c). Offsets are clamped to their cell, as Excel does.
 */
export function chartBox(c: SheetChart, colX: Float64Array, rowY: Float64Array): { x: number; y: number; w: number; h: number } | null {
  const at = (m: Marker) => {
    const col = Math.min(m.col, colX.length - 2);
    const row = Math.min(m.row, rowY.length - 2);
    return {
      x: colX[col] + Math.min(m.dx, colX[col + 1] - colX[col]),
      y: rowY[row] + Math.min(m.dy, rowY[row + 1] - rowY[row]),
    };
  };
  if (c.pos && c.size) return { ...c.pos, ...c.size };
  if (!c.from) return null;
  const a = at(c.from);
  if (c.to) {
    const b = at(c.to);
    return { x: a.x, y: a.y, w: b.x - a.x, h: b.y - a.y };
  }
  return c.size ? { ...a, ...c.size } : null;
}
