import {
  PDFDocument, PDFName, PDFNumber, PDFString, PDFHexString, StandardFonts, degrees,
  PDFCheckBox, PDFDropdown, PDFOptionList, PDFRadioGroup, PDFTextField,
  PDFDict, PDFRef, PDFStream,
  type PDFFont, type PDFPage,
} from 'pdf-lib';
import type { Edit, FormValue, Rgb } from './edits';
import { MARK_NAME_PREFIX } from './marks';

// Writes the edits made in the viewer into the PDF with pdf-lib.
//
// Highlights, underlines, ink and text boxes become real PDF annotations, each with its own
// appearance stream, so every reader (Acrobat, Chrome, Preview, PDF.js) draws them the same and
// they stay editable in other apps. A signature is drawn into the page itself. Form fields get
// their values and fresh appearances. Each annotation is named (`/NM`), so that marks.ts can
// read the app's own marks back for editing; `removed` are such marks that were erased since.

export class EncryptedPdfError extends Error {}

export async function writeEdits(
  original: Uint8Array, edits: Edit[], forms: Map<string, FormValue>, pageOps: PageOps,
  removed: { num: number; gen: number }[] = [],
): Promise<Uint8Array> {
  let doc: PDFDocument;
  try {
    doc = await PDFDocument.load(original, { updateMetadata: false });
  } catch (e) {
    if (String(e).includes('encrypted')) throw new EncryptedPdfError('encrypted');
    throw e;
  }
  dropOldGenerations(doc);
  const pages = doc.getPages();
  let helvetica: PDFFont | null = null;
  const font = async () => (helvetica ??= await doc.embedFont(StandardFonts.Helvetica));

  if (removed.length) removeAnnots(doc, removed.map((r) => PDFRef.of(r.num, r.gen)));
  if (forms.size) await fillForms(doc, forms, await font());

  for (const edit of edits) {
    const page = pages[edit.page];
    if (!page) continue;
    switch (edit.kind) {
      case 'highlight':
      case 'underline':
        markup(doc, page, edit.kind, edit.quads, edit.colour, edit.opacity ?? 1);
        break;
      case 'ink':
        ink(doc, page, edit.paths, edit.colour, edit.width, edit.opacity ?? 1);
        break;
      case 'text':
        freeText(doc, page, edit, await font());
        break;
      case 'image': {
        const image = await doc.embedPng(edit.png);
        drawUpright(page, image, edit);
        break;
      }
    }
  }

  // Page tools last, so the edits above still refer to the original page numbers.
  pages.forEach((page, i) => {
    const turn = pageOps.rotate.get(i);
    if (turn) page.setRotation(degrees((page.getRotation().angle + turn) % 360));
  });
  const deleted = [...pageOps.deleted].sort((a, b) => b - a);
  if (deleted.length >= pages.length) throw new Error('cannot delete every page');
  for (const i of deleted) doc.removePage(i);

  const bytes = await doc.save({ updateFieldAppearances: false });
  // Read the result back before it goes anywhere near the original file.
  const check = await PDFDocument.load(bytes, { updateMetadata: false });
  if (check.getPageCount() !== pages.length - deleted.length) throw new Error('page count changed unexpectedly');
  // pdf-lib reads a file by scanning its objects, not through the cross-reference table, so a
  // broken table passes the check above. Two objects with one number is what breaks it.
  const numbers = check.context.enumerateIndirectObjects().map(([ref]) => ref.objectNumber);
  if (new Set(numbers).size !== numbers.length) throw new Error('object number written twice');
  return bytes;
}

/**
 * An incremental update may give an object number a new generation (e.g. a re-issued catalog,
 * `3184 1 obj` after `3184 0 obj`). pdf-lib reads every revision and keeps both; written out,
 * one number gets two cross-reference entries, every later entry shifts by one, and other readers
 * can no longer find the pages. Only the newest generation is current.
 */
function dropOldGenerations(doc: PDFDocument) {
  const objects = doc.context.enumerateIndirectObjects();
  const newest = new Map<number, PDFRef>();
  for (const [ref] of objects) {
    const seen = newest.get(ref.objectNumber);
    if (!seen || ref.generationNumber > seen.generationNumber) newest.set(ref.objectNumber, ref);
  }
  for (const [ref] of objects) if (newest.get(ref.objectNumber) !== ref) doc.context.delete(ref);
}

export interface PageOps {
  /** Page index → quarter turns clockwise, in degrees (90, 180, 270). */
  rotate: Map<number, number>;
  deleted: Set<number>;
}

// ------------------------------------------------------------------ annotations

const num = (n: number) => PDFNumber.of(+n.toFixed(3));
const arr = (doc: PDFDocument, values: number[]) => doc.context.obj(values.map(num));

function bounds(points: number[], pad: number): number[] {
  const xs = points.filter((_, i) => i % 2 === 0);
  const ys = points.filter((_, i) => i % 2 === 1);
  return [Math.min(...xs) - pad, Math.min(...ys) - pad, Math.max(...xs) + pad, Math.max(...ys) + pad];
}

/** An appearance stream (form XObject) in page space, clipped to `rect`. */
function appearance(doc: PDFDocument, rect: number[], ops: string, resources?: Record<string, unknown>, group = false): PDFRef {
  const stream = doc.context.stream(ops, {
    Type: 'XObject',
    Subtype: 'Form',
    BBox: rect,
    ...(group ? { Group: { S: 'Transparency' } } : {}),
    // pdf-lib turns nested plain objects into PDF dictionaries.
    Resources: (resources ?? {}) as never,
  });
  return doc.context.register(stream);
}

function addAnnot(doc: PDFDocument, page: PDFPage, dict: Record<string, unknown>): void {
  const annot = doc.context.obj({
    Type: 'Annot',
    F: 4, // print
    P: page.ref,
    M: PDFString.fromDate(new Date()),
    NM: PDFString.of(MARK_NAME_PREFIX + Math.random().toString(36).slice(2, 12)),
    ...dict,
  });
  page.node.addAnnot(doc.context.register(annot));
}

/** Takes the given annotations (the app's own) off their pages, with their appearances. */
function removeAnnots(doc: PDFDocument, refs: PDFRef[]): void {
  const gone = new Set(refs);
  for (const page of doc.getPages()) {
    const annots = page.node.Annots();
    if (!annots) continue;
    for (let i = annots.size() - 1; i >= 0; i--) {
      const ref = annots.get(i);
      if (ref instanceof PDFRef && gone.has(ref)) annots.remove(i);
    }
  }
  // Their objects too, or they would stay in the file unseen. write.ts made them, so nothing else
  // shares them: the appearance and, for a see-through one, its inner group.
  for (const ref of gone) {
    const dict = doc.context.lookupMaybe(ref, PDFDict);
    const ap = dict?.lookupMaybe(PDFName.of('AP'), PDFDict);
    const n = ap?.get(PDFName.of('N'));
    if (n instanceof PDFRef) {
      const stream = doc.context.lookup(n);
      const xobjects = stream instanceof PDFStream
        ? stream.dict.lookupMaybe(PDFName.of('Resources'), PDFDict)?.lookupMaybe(PDFName.of('XObject'), PDFDict)
        : undefined;
      const inner = xobjects?.get(PDFName.of('Fm0'));
      if (inner instanceof PDFRef) doc.context.delete(inner);
      doc.context.delete(n);
    }
    doc.context.delete(ref);
  }
}

const rgb = (c: Rgb) => c.map((v) => +v.toFixed(3)).join(' ');

/**
 * The appearance of a mark that is multiplied onto the page (a highlight) and/or see-through:
 * the mark is drawn opaque in an inner group, which is then painted with the blend mode and
 * opacity, so where its parts overlap they don't add up.
 */
function blended(doc: PDFDocument, rect: number[], ops: string, opacity: number, multiply: boolean): PDFRef {
  if (!multiply && opacity >= 1) return appearance(doc, rect, ops);
  const o = +opacity.toFixed(3);
  const inner = appearance(doc, rect, ops, undefined, true);
  return appearance(doc, rect, '/GS0 gs /Fm0 Do', {
    ExtGState: { GS0: doc.context.obj({ ...(multiply ? { BM: 'Multiply' } : {}), CA: o, ca: o }) },
    XObject: { Fm0: inner },
  });
}

/** Highlight / underline over text. Each quad: x1 y1 x2 y2 x3 y3 x4 y4 (TL, TR, BL, BR). */
function markup(doc: PDFDocument, page: PDFPage, kind: 'highlight' | 'underline', quads: number[][], colour: Rgb, opacity: number): void {
  const flat = quads.flat();
  const rect = bounds(flat, 1);
  let ops = '';
  if (kind === 'highlight') {
    ops += `${rgb(colour)} rg `;
    for (const q of quads) ops += `${q[4]} ${q[5]} m ${q[6]} ${q[7]} l ${q[2]} ${q[3]} l ${q[0]} ${q[1]} l h f `;
  } else {
    ops += `${rgb(colour)} RG `;
    for (const q of quads) {
      const h = Math.hypot(q[0] - q[4], q[1] - q[5]);
      const w = Math.max(0.5, h / 14);
      // A line a little above the bottom edge of the quad.
      const t = 1.5 / 14;
      const ax = q[4] + (q[0] - q[4]) * t;
      const ay = q[5] + (q[1] - q[5]) * t;
      const bx = q[6] + (q[2] - q[6]) * t;
      const by = q[7] + (q[3] - q[7]) * t;
      ops += `${w.toFixed(2)} w ${ax} ${ay} m ${bx} ${by} l S `;
    }
  }
  addAnnot(doc, page, {
    Subtype: kind === 'highlight' ? 'Highlight' : 'Underline',
    Rect: arr(doc, rect),
    QuadPoints: arr(doc, flat),
    C: arr(doc, colour),
    // Multiply a highlight, so the text under the colour stays readable.
    AP: doc.context.obj({ N: blended(doc, rect, ops, opacity, kind === 'highlight') }),
  });
}

function ink(doc: PDFDocument, page: PDFPage, paths: number[][], colour: Rgb, width: number, opacity: number): void {
  const rect = bounds(paths.flat(), width);
  let ops = `${rgb(colour)} RG ${width} w 1 J 1 j `;
  for (const p of paths) {
    if (p.length < 2) continue;
    ops += `${p[0]} ${p[1]} m `;
    for (let i = 2; i < p.length; i += 2) ops += `${p[i]} ${p[i + 1]} l `;
    if (p.length === 2) ops += `${p[0] + 0.01} ${p[1]} l `; // a dot
    ops += 'S ';
  }
  addAnnot(doc, page, {
    Subtype: 'Ink',
    Rect: arr(doc, rect),
    InkList: doc.context.obj(paths.map((p) => p.map(num))),
    C: arr(doc, colour),
    BS: doc.context.obj({ W: width, S: 'S' }),
    AP: doc.context.obj({ N: blended(doc, rect, ops, opacity, false) }),
  });
}

/** Characters Helvetica (WinAnsi) can't encode become '?'. */
function encodable(font: PDFFont, text: string): string {
  return [...text].map((ch) => {
    try {
      font.encodeText(ch);
      return ch;
    } catch {
      return '?';
    }
  }).join('');
}

function freeText(doc: PDFDocument, page: PDFPage, edit: Extract<Edit, { kind: 'text' }>, font: PDFFont): void {
  const lines = edit.text.split('\n').map((l) => encodable(font, l));
  const size = edit.size;
  const lead = size * 1.2;
  const width = Math.max(...lines.map((l) => font.widthOfTextAtSize(l, size)), size);
  const height = lines.length * lead;
  // (x, y) is the top-left corner of the text, in PDF space (y up).
  const rect = [edit.x - 2, edit.y - height - 2, edit.x + width + 2, edit.y + 2];
  let ops = `BT /Helv ${size} Tf ${rgb(edit.colour)} rg ${lead} TL ${edit.x} ${edit.y - size} Td `;
  lines.forEach((line, i) => {
    if (i) ops += 'T* ';
    ops += `${font.encodeText(line).toString()} Tj `;
  });
  ops += 'ET';
  addAnnot(doc, page, {
    Subtype: 'FreeText',
    Rect: arr(doc, rect),
    Contents: PDFHexString.fromText(edit.text),
    DA: PDFString.of(`/Helv ${size} Tf ${rgb(edit.colour)} rg`),
    BS: doc.context.obj({ W: 0 }),
    AP: doc.context.obj({ N: appearance(doc, rect, ops, { Font: { Helv: font.ref } }) }),
  });
}

/**
 * Draws a picture (signature) into the page so that it looks upright as the page is shown: on a
 * page with /Rotate, the page's own coordinates are turned relative to the screen.
 */
function drawUpright(page: PDFPage, image: Awaited<ReturnType<PDFDocument['embedPng']>>, box: { x: number; y: number; w: number; h: number }): void {
  const rot = ((page.getRotation().angle % 360) + 360) % 360;
  // box: x, y = bottom-left in PDF space; w, h = size as seen on screen.
  const opts = { width: box.w, height: box.h, rotate: degrees(rot) };
  if (rot === 90) page.drawImage(image, { ...opts, x: box.x + box.h, y: box.y });
  else if (rot === 180) page.drawImage(image, { ...opts, x: box.x + box.w, y: box.y + box.h });
  else if (rot === 270) page.drawImage(image, { ...opts, x: box.x, y: box.y + box.w });
  else page.drawImage(image, { ...opts, x: box.x, y: box.y });
}

// ------------------------------------------------------------------ forms

async function fillForms(doc: PDFDocument, values: Map<string, FormValue>, font: PDFFont): Promise<void> {
  const form = doc.getForm();
  for (const [name, value] of values) {
    try {
      const field = form.getField(name);
      if (field instanceof PDFTextField) field.setText(String(value ?? ''));
      else if (field instanceof PDFCheckBox) {
        if (value) field.check();
        else field.uncheck();
      } else if (field instanceof PDFRadioGroup) {
        if (typeof value === 'string' && value) field.select(value);
        else field.clear();
      } else if (field instanceof PDFDropdown || field instanceof PDFOptionList) {
        const list = Array.isArray(value) ? value : [String(value ?? '')];
        if (list.length && list[0]) field.select(list.length > 1 ? list : list[0]);
        else field.clear();
      }
    } catch (e) {
      console.warn(`form field ${name} not saved`, e);
    }
  }
  try {
    form.updateFieldAppearances(font);
  } catch (e) {
    console.warn('could not redraw form fields', e);
  }
  // Readers should not regenerate what pdf-lib just drew.
  form.acroForm.dict.delete(PDFName.of('NeedAppearances'));
}
