import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNumber, PDFRef, PDFStream, PDFString } from 'pdf-lib';
import type { Edit, Rgb } from './edits';

// Reads back the marks this app saved into a PDF (write.ts), so that they can be edited again:
// highlights, underlines, pen strokes and text boxes. Annotations made by other apps are left
// alone; they may carry things this editor can't keep (authors, comments, replies).
//
// write.ts names each annotation it writes (`/NM (OfficeEditor-…)`). Files saved before it did
// are recognised by the exact set of keys the app has always written, which other apps don't.

export const MARK_NAME_PREFIX = 'OfficeEditor-';

export interface SavedMark {
  edit: Edit;
  ref: { num: number; gen: number };
}

const MARKUP_KEYS = ['Type', 'F', 'P', 'M', 'Subtype', 'Rect', 'QuadPoints', 'C', 'AP'];
const KEYS: Record<string, string[]> = {
  Highlight: MARKUP_KEYS,
  Underline: MARKUP_KEYS,
  Ink: ['Type', 'F', 'P', 'M', 'Subtype', 'Rect', 'InkList', 'C', 'BS', 'AP'],
  FreeText: ['Type', 'F', 'P', 'M', 'Subtype', 'Rect', 'Contents', 'DA', 'BS', 'AP'],
};

/** PDF.js's id for an object reference: "12R", or "12R3" for generation 3. */
export const refId = (ref: { num: number; gen: number }) => `${ref.num}R${ref.gen ? ref.gen : ''}`;

export async function readMarks(bytes: Uint8Array): Promise<SavedMark[]> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false });
  const marks: SavedMark[] = [];
  doc.getPages().forEach((page, index) => {
    const annots = page.node.Annots();
    if (!annots) return;
    for (const entry of annots.asArray()) {
      if (!(entry instanceof PDFRef)) continue;
      const dict = doc.context.lookupMaybe(entry, PDFDict);
      if (!dict) continue;
      try {
        const edit = ours(dict) ? parse(doc, dict, index) : null;
        if (edit) marks.push({ edit, ref: { num: entry.objectNumber, gen: entry.generationNumber } });
      } catch {
        // Not shaped as the app writes it: leave it as it is.
      }
    }
  });
  return marks;
}

function ours(dict: PDFDict): boolean {
  const subtype = name(dict.get(PDFName.of('Subtype')));
  const keys = KEYS[subtype ?? ''];
  if (!keys || number(dict, 'F') !== 4) return false;
  const nm = dict.lookup(PDFName.of('NM'));
  if ((nm instanceof PDFString || nm instanceof PDFHexString) && nm.decodeText().startsWith(MARK_NAME_PREFIX)) return true;
  const present = dict.keys().map((k) => k.decodeText());
  return present.length === keys.length && keys.every((k) => present.includes(k));
}

function parse(doc: PDFDocument, dict: PDFDict, page: number): Edit | null {
  switch (name(dict.get(PDFName.of('Subtype')))) {
    case 'Highlight':
    case 'Underline': {
      const flat = numbers(dict.lookup(PDFName.of('QuadPoints')));
      if (!flat.length || flat.length % 8) return null;
      const quads: number[][] = [];
      for (let i = 0; i < flat.length; i += 8) quads.push(flat.slice(i, i + 8));
      const kind = name(dict.get(PDFName.of('Subtype'))) === 'Highlight' ? 'highlight' : 'underline';
      return { kind, page, quads, colour: colour(dict), opacity: opacity(doc, dict) };
    }
    case 'Ink': {
      const list = dict.lookup(PDFName.of('InkList'));
      if (!(list instanceof PDFArray)) return null;
      const paths = list.asArray().map((p) => numbers(doc.context.lookup(p))).filter((p) => p.length >= 2 && p.length % 2 === 0);
      const bs = dict.lookup(PDFName.of('BS'));
      const width = bs instanceof PDFDict ? number(bs, 'W') : undefined;
      if (!paths.length || !width) return null;
      return { kind: 'ink', page, paths, colour: colour(dict), width, opacity: opacity(doc, dict) };
    }
    case 'FreeText': {
      const contents = dict.lookup(PDFName.of('Contents'));
      const da = dict.lookup(PDFName.of('DA'));
      if (!(contents instanceof PDFString || contents instanceof PDFHexString)) return null;
      if (!(da instanceof PDFString || da instanceof PDFHexString)) return null;
      // As write.ts writes it: "/Helv 14 Tf 0 0 0 rg".
      const m = /\/Helv ([\d.]+) Tf ([\d.]+) ([\d.]+) ([\d.]+) rg/.exec(da.decodeText());
      const rect = numbers(dict.lookup(PDFName.of('Rect')));
      if (!m || rect.length !== 4) return null;
      // write.ts pads the text's box by 2 points; (x, y) is the text's top-left corner.
      return {
        kind: 'text', page, x: rect[0] + 2, y: rect[3] - 2, text: contents.decodeText(),
        size: Number(m[1]), colour: [Number(m[2]), Number(m[3]), Number(m[4])],
      };
    }
  }
  return null;
}

function name(obj: unknown): string | undefined {
  return obj instanceof PDFName ? obj.decodeText() : undefined;
}

function number(dict: PDFDict, key: string): number | undefined {
  const n = dict.lookup(PDFName.of(key));
  return n instanceof PDFNumber ? n.asNumber() : undefined;
}

function numbers(obj: unknown): number[] {
  if (!(obj instanceof PDFArray)) return [];
  const out = obj.asArray().map((n) => (n instanceof PDFNumber ? n.asNumber() : NaN));
  return out.some((n) => !Number.isFinite(n)) ? [] : out;
}

function colour(dict: PDFDict): Rgb {
  const c = numbers(dict.lookup(PDFName.of('C')));
  return c.length === 3 ? (c as Rgb) : [0, 0, 0];
}

/** See-through marks are painted through `/GS0 gs` with `ca` in their appearance (write.ts). */
function opacity(doc: PDFDocument, dict: PDFDict): number {
  const ap = dict.lookup(PDFName.of('AP'));
  const n = ap instanceof PDFDict ? doc.context.lookup(ap.get(PDFName.of('N'))) : undefined;
  if (!(n instanceof PDFStream)) return 1;
  const res = n.dict.lookup(PDFName.of('Resources'));
  const gs = res instanceof PDFDict ? res.lookup(PDFName.of('ExtGState')) : undefined;
  const gs0 = gs instanceof PDFDict ? gs.lookup(PDFName.of('GS0')) : undefined;
  const ca = gs0 instanceof PDFDict ? number(gs0, 'ca') ?? number(gs0, 'CA') : undefined;
  return ca !== undefined && ca > 0 && ca <= 1 ? ca : 1;
}
