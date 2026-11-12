// Edits made in the PDF viewer, kept in PDF page space (points, y up, unrotated) until saved.
// While editing they are drawn as an SVG overlay on each page.

export type Rgb = [number, number, number];

export type Edit =
  | { kind: 'highlight' | 'underline'; page: number; quads: number[][]; colour: Rgb }
  | { kind: 'ink'; page: number; paths: number[][]; colour: Rgb; width: number }
  | { kind: 'text'; page: number; x: number; y: number; text: string; size: number; colour: Rgb }
  | { kind: 'image'; page: number; x: number; y: number; w: number; h: number; png: Uint8Array; url: string };

export type FormValue = string | boolean | string[];

export const COLOURS = {
  highlight: [1, 0.89, 0.2] as Rgb,
  underline: [0.86, 0.15, 0.15] as Rgb,
  ink: [0.1, 0.3, 0.85] as Rgb,
  text: [0, 0, 0] as Rgb,
};

/** Maps PDF space to the page's viewport at scale 1 (CSS px of an unzoomed page). */
export interface PageSpace {
  width: number;
  height: number;
  toView(x: number, y: number): [number, number];
  toPdf(x: number, y: number): [number, number];
  /** Page rotation in degrees (0, 90, 180, 270). */
  rotation: number;
}

const css = (c: Rgb) => `rgb(${c.map((v) => Math.round(v * 255)).join(' ')})`;

/** SVG markup for one edit, in viewport coordinates. */
export function editSvg(edit: Edit, space: PageSpace): string {
  const pt = (x: number, y: number) => space.toView(x, y).map((v) => v.toFixed(2)).join(',');
  switch (edit.kind) {
    case 'highlight':
      return edit.quads.map((q) =>
        `<polygon points="${pt(q[0], q[1])} ${pt(q[2], q[3])} ${pt(q[6], q[7])} ${pt(q[4], q[5])}" fill="${css(edit.colour)}"/>`).join('');
    case 'underline':
      return edit.quads.map((q) => {
        const t = 1.5 / 14;
        const a = pt(q[4] + (q[0] - q[4]) * t, q[5] + (q[1] - q[5]) * t);
        const b = pt(q[6] + (q[2] - q[6]) * t, q[7] + (q[3] - q[7]) * t);
        const h = Math.hypot(q[0] - q[4], q[1] - q[5]);
        return `<polyline points="${a} ${b}" stroke="${css(edit.colour)}" stroke-width="${Math.max(0.5, h / 14)}" fill="none"/>`;
      }).join('');
    case 'ink':
      return edit.paths.map((p) => {
        const pts: string[] = [];
        for (let i = 0; i < p.length; i += 2) pts.push(pt(p[i], p[i + 1]));
        if (pts.length === 1) pts.push(pts[0]);
        return `<polyline points="${pts.join(' ')}" stroke="${css(edit.colour)}" stroke-width="${edit.width}" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`;
      }).join('');
    case 'text': {
      const [x, y] = space.toView(edit.x, edit.y);
      const lines = edit.text.split('\n');
      const tspans = lines.map((l, i) => `<tspan x="${x.toFixed(2)}" dy="${i ? edit.size * 1.2 : edit.size * 0.9}">${escape(l) || ' '}</tspan>`).join('');
      return `<text x="${x.toFixed(2)}" y="${y.toFixed(2)}" font-family="Helvetica, Arial, sans-serif" font-size="${edit.size}" fill="${css(edit.colour)}"${rotateAttr(space, x, y)}>${tspans}</text>`;
    }
    case 'image': {
      // (x, y) is the corner of the picture's box nearest the PDF origin; w × h is its size as
      // seen on screen, which on a quarter-turned page is h × w in PDF space.
      const vertical = space.rotation % 180 !== 0;
      const ex = vertical ? edit.h : edit.w;
      const ey = vertical ? edit.w : edit.h;
      const corners = [space.toView(edit.x, edit.y), space.toView(edit.x + ex, edit.y + ey)];
      const left = Math.min(corners[0][0], corners[1][0]);
      const top = Math.min(corners[0][1], corners[1][1]);
      return `<image href="${edit.url}" x="${left.toFixed(2)}" y="${top.toFixed(2)}" width="${edit.w.toFixed(2)}" height="${edit.h.toFixed(2)}" preserveAspectRatio="none"/>`;
    }
  }
}

/** Text placed on a rotated page is written along the page's x axis, which is turned on screen. */
function rotateAttr(space: PageSpace, x: number, y: number): string {
  return space.rotation ? ` transform="rotate(${space.rotation} ${x.toFixed(2)} ${y.toFixed(2)})"` : '';
}

function escape(s: string): string {
  return s.replace(/[&<>]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;'));
}
