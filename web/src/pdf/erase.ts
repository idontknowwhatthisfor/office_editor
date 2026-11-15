import type { Edit } from './edits';

// The two erasers, in PDF space. Both take the eraser's path as points (the finger's positions,
// close enough together that nothing between them is missed) and a radius, and return what an
// edit becomes: null if the eraser didn't touch it, otherwise its replacements (none = gone).
//
// - Whole marks ("per line"): a touched pen stroke, text box or signature goes; a highlight or
//   underline loses only the touched lines of text.
// - Partial ("per pixel"): pen strokes lose the parts the eraser passes over and are split
//   there. Other marks aren't drawn freehand, so this eraser leaves them alone.

export type Point = [number, number];
type Ink = Extract<Edit, { kind: 'ink' }>;

/** Points along a → b, at most `step` apart, including both ends. */
export function along(a: Point, b: Point, step: number): Point[] {
  const n = Math.max(1, Math.ceil(Math.hypot(b[0] - a[0], b[1] - a[1]) / step));
  const out: Point[] = [];
  for (let i = 0; i <= n; i++) out.push([a[0] + ((b[0] - a[0]) * i) / n, a[1] + ((b[1] - a[1]) * i) / n]);
  return out;
}

export function eraseWhole(edit: Edit, eraser: Point[], r: number, rotation: number): Edit[] | null {
  switch (edit.kind) {
    case 'ink': {
      const reach = r + edit.width / 2;
      const hit = edit.paths.some((p) => {
        if (p.length === 2) return eraser.some((e) => Math.hypot(e[0] - p[0], e[1] - p[1]) < reach);
        for (let i = 2; i < p.length; i += 2) {
          const a: Point = [p[i - 2], p[i - 1]];
          const b: Point = [p[i], p[i + 1]];
          if (eraser.some((e) => toSegment(e, a, b) < reach)) return true;
        }
        return false;
      });
      return hit ? [] : null;
    }
    case 'highlight':
    case 'underline': {
      // Lines of text sit closer together than the eraser is wide: only the line the finger is
      // on goes.
      const quads = edit.quads.filter((q) => !eraser.some((e) => inBox(e, box(q), r / 6)));
      if (quads.length === edit.quads.length) return null;
      return quads.length ? [{ ...edit, quads }] : [];
    }
    case 'text': {
      // As write.ts lays it out: (x, y) is the top-left corner; Helvetica averages about half
      // an em per character.
      const lines = edit.text.split('\n');
      const w = Math.max(...lines.map((l) => l.length), 1) * edit.size * 0.55;
      const h = lines.length * edit.size * 1.2;
      const b = { x0: edit.x, y0: edit.y - h, x1: edit.x + w, y1: edit.y };
      return eraser.some((e) => inBox(e, b, r)) ? [] : null;
    }
    case 'image': {
      const vertical = rotation % 180 !== 0;
      const b = { x0: edit.x, y0: edit.y, x1: edit.x + (vertical ? edit.h : edit.w), y1: edit.y + (vertical ? edit.w : edit.h) };
      return eraser.some((e) => inBox(e, b, r)) ? [] : null;
    }
  }
}

export function erasePart(edit: Edit, eraser: Point[], r: number): Edit[] | null {
  if (edit.kind !== 'ink') return null;
  const reach = r + edit.width / 2;
  const gone = (x: number, y: number) => eraser.some((e) => Math.hypot(e[0] - x, e[1] - y) < reach);
  let touched = false;
  const paths: number[][] = [];
  for (const p of edit.paths) {
    if (p.length === 2) {
      // A dot.
      if (gone(p[0], p[1])) touched = true;
      else paths.push(p);
      continue;
    }
    // Fill in long segments first, so the cut lands where the eraser crossed.
    const dense = densify(p, Math.max(0.5, r / 3));
    const runs: number[][] = [];
    let run: number[] = [];
    let cut = false;
    const flush = () => {
      if (run.length >= 4) runs.push(run);
      run = [];
    };
    for (let i = 0; i < dense.length; i += 2) {
      if (gone(dense[i], dense[i + 1])) {
        cut = true;
        flush();
      } else run.push(dense[i], dense[i + 1]);
    }
    flush();
    // An untouched stroke stays exactly as it was.
    if (cut) {
      touched = true;
      paths.push(...runs);
    } else paths.push(p);
  }
  if (!touched) return null;
  return paths.length ? [{ ...edit, paths } satisfies Ink] : [];
}

function densify(p: number[], step: number): number[] {
  const out = [p[0], p[1]];
  for (let i = 2; i < p.length; i += 2) {
    const pts = along([p[i - 2], p[i - 1]], [p[i], p[i + 1]], step);
    for (let k = 1; k < pts.length; k++) out.push(pts[k][0], pts[k][1]);
  }
  return out;
}

function toSegment(p: Point, a: Point, b: Point): number {
  const dx = b[0] - a[0];
  const dy = b[1] - a[1];
  const len = dx * dx + dy * dy;
  const t = len ? Math.max(0, Math.min(1, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / len)) : 0;
  return Math.hypot(p[0] - (a[0] + t * dx), p[1] - (a[1] + t * dy));
}

interface Box { x0: number; y0: number; x1: number; y1: number }

function box(q: number[]): Box {
  const xs = [q[0], q[2], q[4], q[6]];
  const ys = [q[1], q[3], q[5], q[7]];
  return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) };
}

function inBox(p: Point, b: Box, r: number): boolean {
  return p[0] > b.x0 - r && p[0] < b.x1 + r && p[1] > b.y0 - r && p[1] < b.y1 + r;
}
