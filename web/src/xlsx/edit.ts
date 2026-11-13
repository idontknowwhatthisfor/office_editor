import JSZip from 'jszip';
import { Unsupported, XError, evaluate, numberText, parse, shiftReferences, toNumber, toSerial, type Book, type RangeRef, type Shift, type Value } from './calc';
import { isDateFormat } from './format';
import { shiftFormula } from './formula';
import { decodeCol, defaultRowPt, encodeCol, ptToPx, Workbook, type Cell, type Range, type Sheet } from './workbook';

// Spreadsheet editing (PRD M6): cell values and formulas (recalculated, with everything that
// depends on them), inserting and deleting rows and columns, undo/redo, and saving back into
// the file's own XML.
//
// Cells are never changed in place: an edit puts a new Cell object in the sheet, so the viewer's
// per-cell caches stay right and undo can put the old object back. On save, each worksheet's XML
// is updated where the model changed; everything else in the file is kept as it was.

type Key = string;
const key = (s: number, r: number, c: number): Key => `${s}:${r}:${c}`;

interface Snapshot {
  sheets: (Sheet | null)[];
  names: { name: string; sheet: number | null; formula: string }[];
}

type Op =
  | { kind: 'cell'; sheet: number; r: number; c: number; before: Cell | undefined; after: Cell | undefined; recalc: Map<Key, [Cell | undefined, Cell | undefined]> }
  | { kind: 'shift'; shift: { sheet: number; axis: 'row' | 'col'; at: number; count: number }; snapshot: Snapshot; after: Snapshot }
  | { kind: 'format'; snapshot: Snapshot; after: Snapshot };

interface Deps {
  /** Formula cells by the single cells they read. */
  single: Map<Key, Set<Key>>;
  /** Formula cells reading ranges. */
  ranges: { ref: RangeRef; cell: Key }[];
  /** Formulas to recompute after any change (TODAY, OFFSET, INDIRECT…). */
  always: Set<Key>;
}

export class EditError extends Error {}

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const KEEP = { createFolders: false };

export class SheetEditor {
  ops: Op[] = [];
  private undone: Op[] = [];
  /** Formula cells this app could not recalculate (kept at Excel's saved value). */
  stale = new Set<Key>();
  /** Cells to write into the file on save. */
  private touched = new WeakSet<Cell>();
  private deps: Deps | null = null;
  private unshared = false;
  names: { name: string; sheet: number | null; formula: string }[];
  /** Cell formats added this session: book.formats[book.xfCount + i] is `base` with another font size. */
  private added: { base: number; size: number }[] = [];
  private sized = new Map<string, number>();
  /** Rows made taller for larger text, by sheet. */
  private grown = new Map<number, Set<number>>();
  /** Each sheet's row heights as loaded (to tell which rows to write). */
  private loadedHeights: (Float64Array | null)[];

  private constructor(readonly book: Workbook, readonly sheets: (Sheet | null)[]) {
    this.names = book.definedNames.map((n) => ({ ...n }));
    this.loadedHeights = sheets.map((sh) => sh && sh.rowHeights.slice());
  }

  /** Loads every worksheet (formulas can reach across sheets). */
  static async open(book: Workbook): Promise<SheetEditor> {
    const sheets = await Promise.all(book.sheets.map((s, i) => (s.kind === 'worksheet' ? book.sheet(i) : Promise.resolve(null))));
    return new SheetEditor(book, sheets);
  }

  get changed(): boolean {
    return this.ops.length > 0;
  }

  get canUndo(): boolean {
    return this.ops.length > 0;
  }

  get canRedo(): boolean {
    return this.undone.length > 0;
  }

  /** A cell's formula as text (without '='), resolving shared formulas. */
  formula(s: number, r: number, c: number, cell = this.sheets[s]?.rows[r]?.get(c)): string | undefined {
    if (!cell) return undefined;
    if (cell.f !== undefined) return cell.f;
    if (cell.si === undefined) return undefined;
    const master = this.sheets[s]!.sharedFormulas.get(cell.si);
    return master ? shiftFormula(master.f, r - master.r, c - master.c) : undefined;
  }

  // ---------------------------------------------------------------- reading for the evaluator

  private value(s: number, r: number, c: number): Value {
    const cell = this.sheets[s]?.rows[r]?.get(c);
    if (!cell || cell.t === 'z') return null;
    if (cell.t === 'e') return new XError(String(cell.v));
    return cell.v as Value;
  }

  private bookFor(lookup: (s: number, r: number, c: number) => Value): Book {
    return {
      value: lookup,
      sheetIndex: (name) => this.book.sheets.findIndex((x) => x.name.toLowerCase() === name.toLowerCase()),
      extent: (s) => ({ rows: this.sheets[s]?.rows.length ?? 0, cols: this.sheets[s]?.colWidths.length ?? 0 }),
      name: (name, sheet) => this.nameRef(name, sheet),
    };
  }

  private nameRef(name: string, sheet: number): RangeRef | null {
    const lower = name.toLowerCase();
    const def = this.names.find((n) => n.name.toLowerCase() === lower && n.sheet === sheet)
      ?? this.names.find((n) => n.name.toLowerCase() === lower && n.sheet === null);
    if (!def) return null;
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(def.formula.trim());
    if (!m) return null;
    const s = this.bookFor(() => null).sheetIndex((m[1] ?? m[2]).replace(/''/g, "'"));
    if (s < 0) return null;
    const r0 = Number(m[4]) - 1;
    const c0 = decodeCol(m[3]);
    return { sheet: s, r0, c0, r1: m[6] ? Number(m[6]) - 1 : r0, c1: m[5] ? decodeCol(m[5]) : c0 };
  }

  // ---------------------------------------------------------------- dependencies

  private dependencies(): Deps {
    if (this.deps) return this.deps;
    const deps: Deps = { single: new Map(), ranges: [], always: new Set() };
    const book = this.bookFor(() => null);
    this.sheets.forEach((sheet, s) => sheet?.rows.forEach((row, r) => row?.forEach((cell, c) => {
      const f = this.formula(s, r, c, cell);
      if (f === undefined) return;
      const k = key(s, r, c);
      if (/\b(TODAY|NOW|RAND|RANDBETWEEN|OFFSET|INDIRECT|INFO|CELL)\s*\(/i.test(f)) deps.always.add(k);
      let tree: ReturnType<typeof parse>;
      try {
        tree = parse(f);
      } catch {
        return;
      }
      const visit = (n: ReturnType<typeof parse>): void => {
        if (n.k === 'ref') {
          const sheetIndex = n.sheet === null ? s : book.sheetIndex(n.sheet);
          if (sheetIndex < 0) return;
          if (n.single) {
            const target = key(sheetIndex, n.r0, n.c0);
            let set = deps.single.get(target);
            if (!set) deps.single.set(target, (set = new Set()));
            set.add(k);
          } else deps.ranges.push({ ref: { sheet: sheetIndex, r0: n.r0, c0: n.c0, r1: n.r1, c1: n.c1 }, cell: k });
        } else if (n.k === 'name') {
          const ref = this.nameRef(n.v, s);
          if (ref) deps.ranges.push({ ref, cell: k });
          else deps.always.add(k);
        } else if (n.k === 'fn') n.args.forEach(visit);
        else if (n.k === 'un' || n.k === 'pct') visit(n.a);
        else if (n.k === 'bin') {
          visit(n.a);
          visit(n.b);
        } else if (n.k === 'arr') n.rows.flat().forEach(visit);
      };
      visit(tree);
    })));
    this.deps = deps;
    return deps;
  }

  private dependents(k: Key): Key[] {
    const deps = this.dependencies();
    const [s, r, c] = k.split(':').map(Number);
    const out = [...(deps.single.get(k) ?? [])];
    for (const { ref, cell } of deps.ranges) {
      if (ref.sheet === s && r >= ref.r0 && r <= ref.r1 && c >= ref.c0 && c <= ref.c1) out.push(cell);
    }
    return out;
  }

  /**
   * Recomputes the formulas that depend on `changed` (and those cells themselves if they are
   * formulas). Returns each recomputed cell's [old, new] object.
   */
  private recalc(changed: Key[], everything = false): Map<Key, [Cell | undefined, Cell | undefined]> {
    const deps = this.dependencies();
    const dirty = new Set<Key>();
    const queue = [...changed];
    if (everything) {
      this.sheets.forEach((sheet, s) => sheet?.rows.forEach((row, r) => row?.forEach((cell, c) => {
        if (this.formula(s, r, c, cell) !== undefined) queue.push(key(s, r, c));
      })));
    }
    for (const k of deps.always) queue.push(k);
    while (queue.length) {
      const k = queue.pop()!;
      if (dirty.has(k)) continue;
      dirty.add(k);
      for (const d of this.dependents(k)) if (!dirty.has(d)) queue.push(d);
    }

    const done = new Set<Key>();
    const busy = new Set<Key>();
    const result = new Map<Key, [Cell | undefined, Cell | undefined]>();
    const lookup = (s: number, r: number, c: number): Value => {
      const k = key(s, r, c);
      if (dirty.has(k) && !done.has(k) && !busy.has(k)) compute(k);
      return this.value(s, r, c);
    };
    const book = this.bookFor(lookup);
    const compute = (k: Key) => {
      const [s, r, c] = k.split(':').map(Number);
      const cell = this.sheets[s]?.rows[r]?.get(c);
      const f = cell && this.formula(s, r, c, cell);
      if (!cell || f === undefined) {
        done.add(k);
        return;
      }
      busy.add(k);
      try {
        let v = evaluate(f, s, r, c, book);
        if (v === null) v = 0;
        const next: Cell = { ...cell, ...asCell(v) };
        if (next.t !== cell.t || next.v !== cell.v) {
          this.put(s, r, c, next);
          result.set(k, [cell, next]);
        }
        this.stale.delete(k);
      } catch (e) {
        if (!(e instanceof Unsupported)) console.warn('formula failed', f, e);
        this.stale.add(k);
      } finally {
        busy.delete(k);
        done.add(k);
      }
    };
    for (const k of dirty) if (!done.has(k)) compute(k);
    return result;
  }

  private put(s: number, r: number, c: number, cell: Cell | undefined): void {
    const sheet = this.sheets[s]!;
    let row = sheet.rows[r];
    if (!cell) {
      row?.delete(c);
      return;
    }
    if (!row) sheet.rows[r] = row = new Map();
    row.set(c, cell);
    this.touched.add(cell);
  }

  // ---------------------------------------------------------------- editing cells

  /** What the user typed into a cell: a value, or a formula starting with '='. */
  setCell(s: number, r: number, c: number, input: string): Map<Key, [Cell | undefined, Cell | undefined]> {
    const sheet = this.sheets[s];
    if (!sheet) throw new EditError('not a worksheet');
    let before = sheet.rows[r]?.get(c);
    // Changing a cell of a shared formula: give each cell of the group its own formula first.
    if (before && (before.si !== undefined || [...sheet.sharedFormulas.values()].some((m) => m.r === r && m.c === c))) {
      this.unshareAll();
      before = sheet.rows[r]?.get(c);
    }
    const xf = before?.xf ?? 0;
    let after: Cell | undefined;
    if (input.startsWith('=') && input.length > 1) {
      const f = input.slice(1);
      try {
        parse(f);
      } catch {
        throw new EditError('formula');
      }
      after = { t: 'n', v: 0, xf, f };
    } else if (input === '') {
      after = xf ? { t: 'z', v: null, xf } : undefined;
    } else {
      after = { ...literal(input, this.book.formats[xf]?.numFmt ?? 'General'), xf };
    }
    const formulaChanged = before?.f !== undefined || before?.si !== undefined || after?.f !== undefined;
    this.put(s, r, c, after);
    if (after) this.touched.add(after);
    if (formulaChanged) this.deps = null;
    const recalc = this.recalc([key(s, r, c)]);
    const op: Op = { kind: 'cell', sheet: s, r, c, before, after: sheet.rows[r]?.get(c), recalc };
    this.ops.push(op);
    this.undone = [];
    return recalc;
  }

  // ---------------------------------------------------------------- formatting

  /** Formatting needs the file's styles part, read in full. */
  get canFormat(): boolean {
    return !!this.book.stylesPath && this.book.xfCount > 0 && this.book.xfCount + this.added.length === this.book.formats.length;
  }

  /**
   * Sets the font size (points) of the cells in `range`. Empty cells get the size too when the
   * range is small (typing there later uses it); a whole row or column only changes its filled
   * cells. Rows grow to fit, as Excel's automatic row height does (they never shrink here).
   */
  setFontSize(s: number, range: Range, pt: number): void {
    const sheet = this.sheets[s];
    if (!sheet) throw new EditError('not a worksheet');
    if (!this.canFormat) throw new EditError('styles');
    const snapshot = this.snapshot();
    const fillEmpty = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1) <= 500;
    const rows = new Set<number>();
    for (let r = range.r0; r <= range.r1; r++) {
      const row = sheet.rows[r];
      if (!row && !fillEmpty) continue;
      for (let c = range.c0; c <= range.c1; c++) {
        const cell = row?.get(c);
        if (!cell && !fillEmpty) continue;
        const xf = this.withSize(cell?.xf ?? 0, pt);
        if (cell?.xf === xf) continue;
        this.put(s, r, c, cell ? { ...cell, xf } : { t: 'z', v: null, xf });
        rows.add(r);
      }
    }
    if (!rows.size) return;
    // Excel's automatic height for the font (default rows are 15 pt for 11 pt text).
    const need = ptToPx(defaultRowPt({ ...this.book.formats[0].font, size: pt }));
    for (const r of rows) {
      const h = sheet.rowHeights[r];
      if (h > 0 && h < need) {
        sheet.rowHeights[r] = need;
        if (!this.grown.has(s)) this.grown.set(s, new Set());
        this.grown.get(s)!.add(r);
      }
    }
    this.ops.push({ kind: 'format', snapshot, after: this.snapshot() });
    this.undone = [];
  }

  /** The cell format like `xf` with font size `pt`, added if the file doesn't have it yet. */
  private withSize(xf: number, pt: number): number {
    const base = xf >= this.book.xfCount ? this.added[xf - this.book.xfCount].base : xf;
    const fmt = this.book.formats[base] ?? this.book.formats[0];
    if (fmt.font.size === pt) return base;
    const k = `${base}:${pt}`;
    let index = this.sized.get(k);
    if (index === undefined) {
      index = this.book.formats.length;
      this.book.formats.push({ ...fmt, font: { ...fmt.font, size: pt } });
      this.added.push({ base, size: pt });
      this.sized.set(k, index);
    }
    return index;
  }

  // ---------------------------------------------------------------- rows and columns

  canShift(s: number): boolean {
    return !!this.sheets[s] && !this.sheets[s]!.hasTables;
  }

  /** Inserts (count > 0) or deletes (count < 0) rows or columns at `at` on sheet `s`. */
  shift(s: number, axis: 'row' | 'col', at: number, count: number): void {
    if (!this.canShift(s)) throw new EditError('tables');
    const snapshot = this.snapshot();
    this.unshareAll();
    const name = this.book.sheets[s].name;
    const shift: Shift = { sheet: name, axis, at, count };
    const sheet = this.sheets[s]!;
    moveCells(sheet, axis, at, count);
    // Every formula in the workbook that points at the moved cells.
    this.sheets.forEach((other, os) => other?.rows.forEach((row, r) => row?.forEach((cell, c) => {
      if (cell.f === undefined) return;
      const f = shiftReferences(cell.f, this.book.sheets[os].name, shift);
      if (f !== cell.f) this.put(os, r, c, { ...cell, f });
    })));
    this.names = this.names.map((n) => ({ ...n, formula: shiftReferences(n.formula, '\u0000', shift) }));
    this.deps = null;
    this.recalc([], true);
    this.ops.push({ kind: 'shift', shift: { sheet: s, axis, at, count }, snapshot, after: this.snapshot() });
    this.undone = [];
  }

  /** Gives every cell of a shared-formula group its own formula text (shifting needs that). */
  private unshareAll(): void {
    if (this.unshared) return;
    this.sheets.forEach((sheet, s) => {
      if (!sheet) return;
      sheet.rows.forEach((row, r) => row?.forEach((cell, c) => {
        if (cell.si === undefined && cell.f === undefined) return;
        const f = this.formula(s, r, c, cell);
        const next: Cell = { ...cell, f };
        delete next.si;
        row.set(c, next);
        this.touched.add(next);
      }));
      sheet.sharedFormulas = new Map();
    });
    this.unshared = true;
  }

  private snapshot(): Snapshot {
    return {
      sheets: this.sheets.map((sheet) => sheet && {
        ...sheet,
        rows: sheet.rows.map((row) => row && new Map(row)),
        colWidths: sheet.colWidths.slice(),
        rowHeights: sheet.rowHeights.slice(),
        merges: sheet.merges.map((m) => ({ ...m })),
        sharedFormulas: new Map(sheet.sharedFormulas),
      }),
      names: this.names.map((n) => ({ ...n })),
    };
  }

  private restore(snap: Snapshot): void {
    snap.sheets.forEach((saved, i) => {
      const sheet = this.sheets[i];
      if (!sheet || !saved) return;
      // Keep the same Sheet objects (the viewer holds them); copy the saved state in.
      Object.assign(sheet, {
        ...saved,
        rows: saved.rows.map((row) => row && new Map(row)),
        colWidths: saved.colWidths.slice(),
        rowHeights: saved.rowHeights.slice(),
        merges: saved.merges.map((m) => ({ ...m })),
        sharedFormulas: new Map(saved.sharedFormulas),
      });
    });
    this.names = snap.names.map((n) => ({ ...n }));
    this.unshared = snap.sheets.every((s) => !s || s.sharedFormulas.size === 0);
    this.deps = null;
  }

  // ---------------------------------------------------------------- undo

  undo(): Op | null {
    const op = this.ops.pop();
    if (!op) return null;
    this.undone.push(op);
    if (op.kind === 'cell') {
      for (const [k, [old]] of op.recalc) {
        const [s, r, c] = k.split(':').map(Number);
        this.put(s, r, c, old);
      }
      this.put(op.sheet, op.r, op.c, op.before);
      if (op.before?.f !== undefined || op.after?.f !== undefined || op.before?.si !== undefined) this.deps = null;
    } else {
      this.restore(op.snapshot);
    }
    return op;
  }

  redo(): Op | null {
    const op = this.undone.pop();
    if (!op) return null;
    this.ops.push(op);
    if (op.kind === 'cell') {
      this.put(op.sheet, op.r, op.c, op.after);
      for (const [k, [, next]] of op.recalc) {
        const [s, r, c] = k.split(':').map(Number);
        this.put(s, r, c, next);
      }
      if (op.before?.f !== undefined || op.after?.f !== undefined) this.deps = null;
    } else {
      this.restore(op.after);
    }
    return op;
  }

  // ---------------------------------------------------------------- saving

  async write(): Promise<Uint8Array> {
    const src = this.book.pkg.zip;
    const out = await JSZip.loadAsync(await src.generateAsync({ type: 'arraybuffer' }));
    const shifts = this.ops.filter((op): op is Extract<Op, { kind: 'shift' }> => op.kind === 'shift').map((op) => op.shift);

    for (let s = 0; s < this.sheets.length; s++) {
      const sheet = this.sheets[s];
      if (!sheet) continue;
      const path = this.book.sheets[s].path;
      const xml = await src.file(path)?.async('string');
      if (!xml) continue;
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      let changed = false;
      for (const shift of shifts) {
        if (shift.sheet === s) {
          shiftSheetXml(doc, this.book.sheets[s].name, shift);
          changed = true;
        }
        // Formulas in conditional formats and validations of other sheets can point here too.
      }
      // Once formulas are unshared, every formula cell is written with its own formula.
      if (this.writeCells(doc, s, shifts.length > 0 || this.unshared)) changed = true;
      if (!changed) continue;
      out.file(path, new XMLSerializer().serializeToString(doc), KEEP);
    }

    await this.writeWorkbook(out, shifts.length > 0);
    if (this.added.length) await this.writeStyles(out);
    const bytes = await out.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    // Read it back as the viewer would before it may replace the original.
    const check = await Workbook.open(bytes.buffer as ArrayBuffer);
    for (let i = 0; i < check.sheets.length; i++) if (check.sheets[i].kind === 'worksheet') await check.sheet(i);
    return bytes;
  }

  /** Puts changed cells into a sheet's XML. Returns whether anything changed. */
  private writeCells(doc: Document, s: number, rewriteFormulas: boolean): boolean {
    const sheet = this.sheets[s]!;
    const root = doc.documentElement;
    const sheetData = childNS(root, 'sheetData') ?? root.appendChild(el(doc, 'sheetData'));
    let changed = false;
    const seen = new Set<string>();

    for (const rowEl of [...childrenNS(sheetData, 'row')]) {
      const r = Number(rowEl.getAttribute('r')) - 1;
      for (const cEl of [...childrenNS(rowEl, 'c')]) {
        const ref = cEl.getAttribute('r');
        if (!ref) continue;
        const c = decodeCol(ref);
        seen.add(`${r}:${c}`);
        const cell = sheet.rows[r]?.get(c);
        if (!cell) {
          // Cleared: keep a styled empty cell, drop an unstyled one.
          if (childNS(cEl, 'v') || childNS(cEl, 'f') || childNS(cEl, 'is')) {
            const s0 = cEl.getAttribute('s');
            if (!s0 || s0 === '0') rowEl.removeChild(cEl);
            else clearCell(cEl);
            changed = true;
          }
          continue;
        }
        if (this.touched.has(cell) || (rewriteFormulas && cell.f !== undefined)) {
          // A recalculated cell whose formula is unchanged keeps its <f> (it may be part of a
          // shared formula); only its value is new.
          const f = childNS(cEl, 'f');
          const same = !rewriteFormulas && f && (cell.si !== undefined || (cell.f !== undefined && f.textContent === cell.f));
          if (same) {
            writeValue(doc, cEl, cell);
            if (cell.xf) cEl.setAttribute('s', String(cell.xf));
            else cEl.removeAttribute('s');
          } else writeCell(doc, cEl, cell);
          changed = true;
        }
      }
    }
    // Cells that are new in this session.
    sheet.rows.forEach((row, r) => row?.forEach((cell, c) => {
      if (seen.has(`${r}:${c}`) || !this.touched.has(cell)) return;
      const rowEl = findOrAddRow(doc, sheetData, r);
      const cEl = el(doc, 'c');
      cEl.setAttribute('r', `${encodeCol(c)}${r + 1}`);
      const next = childrenNS(rowEl, 'c').find((x) => decodeCol(x.getAttribute('r') ?? 'A') > c) ?? null;
      rowEl.insertBefore(cEl, next);
      writeCell(doc, cEl, cell);
      changed = true;
    }));
    // Rows made taller for larger text (unless undone since).
    const loaded = this.loadedHeights[s];
    const shifted = this.ops.some((op) => op.kind === 'shift');
    for (const r of this.grown.get(s) ?? []) {
      const h = sheet.rowHeights[r];
      if (h === undefined || (!shifted && loaded && loaded[r] === h)) continue;
      findOrAddRow(doc, sheetData, r).setAttribute('ht', String((h * 72) / 96));
      changed = true;
    }
    if (changed) updateDimension(doc, sheet);
    return changed;
  }

  /** Adds a font and a cell format to styles.xml for each format added this session, in order. */
  private async writeStyles(out: JSZip): Promise<void> {
    const path = this.book.stylesPath;
    const doc = new DOMParser().parseFromString(await out.file(path)!.async('string'), 'application/xml');
    const root = doc.documentElement;
    const fonts = childNS(root, 'fonts')!;
    const cellXfs = childNS(root, 'cellXfs')!;
    const xfs = childrenNS(cellXfs, 'xf');
    for (const { base, size } of this.added) {
      const baseXf = xfs[base] ?? xfs[0];
      const fontEls = childrenNS(fonts, 'font');
      const font = (fontEls[Number(baseXf.getAttribute('fontId') ?? 0)] ?? fontEls[0]).cloneNode(true) as Element;
      let sz = childNS(font, 'sz');
      if (!sz) {
        sz = el(doc, 'sz');
        // Schema order: b i strike condense extend outline shadow u vertAlign sz color name family charset scheme.
        const after = ['color', 'name', 'family', 'charset', 'scheme'].map((n) => childNS(font, n)).find(Boolean) ?? null;
        font.insertBefore(sz, after);
      }
      sz.setAttribute('val', String(size));
      fonts.append(font);
      fonts.setAttribute('count', String(childrenNS(fonts, 'font').length));
      const xf = baseXf.cloneNode(true) as Element;
      xf.setAttribute('fontId', String(childrenNS(fonts, 'font').length - 1));
      xf.setAttribute('applyFont', '1');
      cellXfs.append(xf);
      cellXfs.setAttribute('count', String(childrenNS(cellXfs, 'xf').length));
    }
    out.file(path, new XMLSerializer().serializeToString(doc), KEEP);
  }

  private async writeWorkbook(out: JSZip, shifted: boolean): Promise<void> {
    const path = this.book.bookPath;
    const doc = new DOMParser().parseFromString(await out.file(path)!.async('string'), 'application/xml');
    const root = doc.documentElement;
    if (shifted) {
      const defs = [...root.getElementsByTagNameNS('*', 'definedName')];
      defs.forEach((d, i) => {
        if (this.names[i]) d.textContent = this.names[i].formula;
      });
    }
    // Excel (and LibreOffice) recompute everything on open, including what this app could not.
    let calcPr = childNS(root, 'calcPr');
    if (!calcPr) {
      calcPr = el(doc, 'calcPr');
      const after = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst']
        .map((n) => childNS(root, n)).find(Boolean) ?? null;
      root.insertBefore(calcPr, after);
    }
    calcPr.setAttribute('fullCalcOnLoad', '1');
    out.file(path, new XMLSerializer().serializeToString(doc), KEEP);

    // The calculation chain lists formula cells by position; it is now out of date. Excel
    // rebuilds it when missing (a wrong one makes it "repair" the file).
    const relsPath = path.replace(/[^/]+$/, (f) => `_rels/${f}.rels`);
    const rels = await out.file(relsPath)?.async('string');
    const chainRel = rels && /<Relationship [^>]*Type="[^"]*\/calcChain"[^>]*\/>/.exec(rels);
    if (chainRel) {
      const target = /Target="([^"]+)"/.exec(chainRel[0])?.[1] ?? 'calcChain.xml';
      out.file(relsPath, rels!.replace(chainRel[0], ''), KEEP);
      const chainPath = target.startsWith('/') ? target.slice(1) : path.replace(/[^/]+$/, '') + target;
      out.remove(chainPath);
      const types = await out.file('[Content_Types].xml')?.async('string');
      if (types) out.file('[Content_Types].xml', types.replace(new RegExp(`<Override [^>]*PartName="/${chainPath.replace(/[.]/g, '\\.')}"[^>]*/>`), ''), KEEP);
    }
  }
}

// ------------------------------------------------------------------ values typed by the user

/** A typed value: number (with the user's decimal comma if their language uses one), %,
 * TRUE/FALSE, a date in a date-formatted cell, or text ("'" forces text). */
function literal(input: string, numFmt: string): Omit<Cell, 'xf'> {
  if (input.startsWith("'")) return { t: 's', v: input.slice(1) };
  const up = input.trim().toUpperCase();
  if (up === 'TRUE' || up === 'FALSE') return { t: 'b', v: up === 'TRUE' };
  const commaDecimal = /^(id|in|de|fr|es|it|nl|pt|ru|tr)\b/i.test(navigator.language);
  let t = input.trim();
  if (commaDecimal && /^-?[\d.]*,\d+%?$|^-?\d{1,3}(\.\d{3})+(,\d+)?%?$/.test(t)) t = t.replace(/\./g, '').replace(',', '.');
  const n = toNumber(t);
  if (n !== null) return { t: 'n', v: n };
  if (isDateFormat(numFmt)) {
    const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
    const dmy = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/.exec(t);
    if (iso) return { t: 'n', v: toSerial(+iso[1], +iso[2], +iso[3]) };
    if (dmy) {
      const monthFirst = /^en-US\b/i.test(navigator.language);
      const [d, m] = monthFirst ? [+dmy[2], +dmy[1]] : [+dmy[1], +dmy[2]];
      return { t: 'n', v: toSerial(+dmy[3], m, d) };
    }
  }
  return { t: 's', v: input };
}

function asCell(v: Value): Pick<Cell, 't' | 'v'> {
  if (v instanceof XError) return { t: 'e', v: v.code };
  if (typeof v === 'number') return { t: 'n', v };
  if (typeof v === 'boolean') return { t: 'b', v };
  return { t: 's', v: v ?? '' };
}

// ------------------------------------------------------------------ moving cells in the model

function moveCells(sheet: Sheet, axis: 'row' | 'col', at: number, count: number): void {
  if (axis === 'row') {
    if (count > 0) {
      sheet.rows.splice(at, 0, ...new Array<undefined>(count));
      const h = new Float64Array(sheet.rowHeights.length + count);
      h.set(sheet.rowHeights.subarray(0, at));
      h.fill(sheet.rowHeights[Math.max(0, at - 1)] || 20, at, at + count);
      h.set(sheet.rowHeights.subarray(at), at + count);
      sheet.rowHeights = h;
    } else {
      sheet.rows.splice(at, -count);
      const h = new Float64Array(sheet.rowHeights.length + count);
      h.set(sheet.rowHeights.subarray(0, at));
      h.set(sheet.rowHeights.subarray(at - count), at);
      sheet.rowHeights = h;
    }
  } else {
    sheet.rows = sheet.rows.map((row) => {
      if (!row) return row;
      const next = new Map<number, Cell>();
      for (const [c, cell] of row) {
        if (c < at) next.set(c, cell);
        else if (count > 0) next.set(c + count, cell);
        else if (c >= at - count) next.set(c + count, cell);
      }
      return next;
    });
    const w = new Float64Array(sheet.colWidths.length + count);
    if (count > 0) {
      w.set(sheet.colWidths.subarray(0, at));
      w.fill(sheet.colWidths[Math.max(0, at - 1)] || 64, at, at + count);
      w.set(sheet.colWidths.subarray(at), at + count);
    } else {
      w.set(sheet.colWidths.subarray(0, at));
      w.set(sheet.colWidths.subarray(at - count), at);
    }
    sheet.colWidths = w;
  }
  sheet.merges = sheet.merges.flatMap((m) => {
    const moved = moveRange([m.r0, m.c0, m.r1, m.c1], axis, at, count);
    return moved ? [{ r0: moved[0], c0: moved[1], r1: moved[2], c1: moved[3] }] : [];
  });
}

/** A range [r0, c0, r1, c1] after a shift, or null if it was deleted entirely. */
function moveRange(range: number[], axis: 'row' | 'col', at: number, count: number): number[] | null {
  const [i0, i1] = axis === 'row' ? [0, 2] : [1, 3];
  const out = range.slice();
  if (count > 0) {
    if (out[i0] >= at) out[i0] += count;
    if (out[i1] >= at) out[i1] += count;
    return out;
  }
  const g0 = at;
  const g1 = at - count - 1;
  if (out[i0] >= g0 && out[i1] <= g1) return null;
  out[i0] = out[i0] > g1 ? out[i0] + count : out[i0] >= g0 ? g0 : out[i0];
  out[i1] = out[i1] > g1 ? out[i1] + count : out[i1] >= g0 ? g0 - 1 : out[i1];
  return out;
}

// ------------------------------------------------------------------ sheet XML

function el(doc: Document, name: string): Element {
  const prefix = doc.documentElement.prefix;
  return doc.createElementNS(doc.documentElement.namespaceURI ?? MAIN, prefix ? `${prefix}:${name}` : name);
}

function childNS(parent: Element | null, name: string): Element | null {
  if (!parent) return null;
  for (const c of parent.children) if (c.localName === name) return c;
  return null;
}

function childrenNS(parent: Element | null, name: string): Element[] {
  return parent ? [...parent.children].filter((c) => c.localName === name) : [];
}

function clearCell(cEl: Element): void {
  for (const c of [...cEl.children]) if (['f', 'v', 'is'].includes(c.localName)) cEl.removeChild(c);
  cEl.removeAttribute('t');
}

/** Replaces a cell's value, keeping its formula element. */
function writeValue(doc: Document, cEl: Element, cell: Cell): void {
  for (const c of [...cEl.children]) if (c.localName === 'v' || c.localName === 'is') cEl.removeChild(c);
  cEl.removeAttribute('t');
  const v = el(doc, 'v');
  if (cell.t === 'n') v.textContent = numberText(cell.v as number);
  else if (cell.t === 'b') {
    cEl.setAttribute('t', 'b');
    v.textContent = cell.v ? '1' : '0';
  } else if (cell.t === 'e') {
    cEl.setAttribute('t', 'e');
    v.textContent = String(cell.v);
  } else if (cell.t === 's') {
    cEl.setAttribute('t', 'str');
    v.textContent = String(cell.v ?? '');
  } else return;
  cEl.append(v);
}

function writeCell(doc: Document, cEl: Element, cell: Cell): void {
  clearCell(cEl);
  if (cell.xf) cEl.setAttribute('s', String(cell.xf));
  else cEl.removeAttribute('s');
  const add = (name: string, text: string) => {
    const e = el(doc, name);
    e.textContent = text;
    cEl.append(e);
    return e;
  };
  if (cell.f !== undefined) {
    add('f', cell.f);
    // The value as computed (or as Excel last saved it, if this app couldn't).
    if (cell.t === 'n') add('v', numberText(cell.v as number));
    else if (cell.t === 's') {
      cEl.setAttribute('t', 'str');
      add('v', String(cell.v ?? ''));
    } else if (cell.t === 'b') {
      cEl.setAttribute('t', 'b');
      add('v', cell.v ? '1' : '0');
    } else if (cell.t === 'e') {
      cEl.setAttribute('t', 'e');
      add('v', String(cell.v));
    }
    return;
  }
  switch (cell.t) {
    case 'n':
      add('v', numberText(cell.v as number));
      break;
    case 'b':
      cEl.setAttribute('t', 'b');
      add('v', cell.v ? '1' : '0');
      break;
    case 'e':
      cEl.setAttribute('t', 'e');
      add('v', String(cell.v));
      break;
    case 's': {
      // An inline string: no need to touch the shared string table.
      cEl.setAttribute('t', 'inlineStr');
      const is = el(doc, 'is');
      const t = el(doc, 't');
      t.textContent = String(cell.v ?? '');
      if (/^\s|\s$|\n/.test(t.textContent)) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      is.append(t);
      cEl.append(is);
      break;
    }
  }
}

function findOrAddRow(doc: Document, sheetData: Element, r: number): Element {
  const rows = childrenNS(sheetData, 'row');
  const existing = rows.find((x) => Number(x.getAttribute('r')) === r + 1);
  if (existing) return existing;
  const row = el(doc, 'row');
  row.setAttribute('r', String(r + 1));
  sheetData.insertBefore(row, rows.find((x) => Number(x.getAttribute('r')) > r + 1) ?? null);
  return row;
}

function updateDimension(doc: Document, sheet: Sheet): void {
  const dim = childNS(doc.documentElement, 'dimension');
  if (!dim) return;
  let maxR = 0;
  let maxC = 0;
  sheet.rows.forEach((row, r) => row?.forEach((cell, c) => {
    if (cell.t === 'z' && !cell.xf) return;
    maxR = Math.max(maxR, r);
    maxC = Math.max(maxC, c);
  }));
  dim.setAttribute('ref', maxR || maxC ? `A1:${encodeCol(maxC)}${maxR + 1}` : 'A1');
}

/** Moves rows/columns in a sheet's XML: cells, merges, column widths and every range attribute. */
function shiftSheetXml(doc: Document, sheetName: string, shift: { axis: 'row' | 'col'; at: number; count: number }): void {
  const root = doc.documentElement;
  const sheetData = childNS(root, 'sheetData');
  const { axis, at, count } = shift;
  const ref = (text: string) => shiftReferences(text, sheetName, { axis, at, count, sheet: sheetName });

  for (const rowEl of childrenNS(sheetData, 'row')) {
    const r = Number(rowEl.getAttribute('r')) - 1;
    if (axis === 'row') {
      if (count < 0 && r >= at && r < at - count) {
        sheetData!.removeChild(rowEl);
        continue;
      }
      const nr = r >= at ? r + count : r;
      rowEl.setAttribute('r', String(nr + 1));
      for (const cEl of childrenNS(rowEl, 'c')) {
        const c = decodeCol(cEl.getAttribute('r') ?? 'A');
        cEl.setAttribute('r', `${encodeCol(c)}${nr + 1}`);
      }
    } else {
      for (const cEl of childrenNS(rowEl, 'c')) {
        const c = decodeCol(cEl.getAttribute('r') ?? 'A');
        if (count < 0 && c >= at && c < at - count) {
          rowEl.removeChild(cEl);
          continue;
        }
        cEl.setAttribute('r', `${encodeCol(c >= at ? c + count : c)}${r + 1}`);
      }
      rowEl.removeAttribute('spans');
    }
  }
  if (axis === 'row' && count > 0) {
    // Inserted rows start out empty; keep them out of the XML.
  }
  if (axis === 'col') {
    const cols = childNS(root, 'cols');
    for (const col of childrenNS(cols, 'col')) {
      const moved = moveRange([0, Number(col.getAttribute('min')) - 1, 0, Number(col.getAttribute('max')) - 1], 'col', at, count);
      if (!moved) cols!.removeChild(col);
      else {
        col.setAttribute('min', String(moved[1] + 1));
        col.setAttribute('max', String(moved[3] + 1));
      }
    }
  }
  // Every attribute holding cell ranges, and formulas in conditional formats and validations.
  const sqref = (value: string) => value.split(/\s+/).map(ref).filter((x) => x && !x.includes('#REF!')).join(' ');
  for (const e of root.getElementsByTagNameNS('*', 'mergeCell')) {
    const v = ref(e.getAttribute('ref') ?? '');
    if (v.includes('#REF!')) e.parentNode!.removeChild(e);
    else e.setAttribute('ref', v);
  }
  const merges = childNS(root, 'mergeCells');
  if (merges) {
    const n = childrenNS(merges, 'mergeCell').length;
    if (n) merges.setAttribute('count', String(n));
    else root.removeChild(merges);
  }
  for (const name of ['conditionalFormatting', 'dataValidation', 'selection', 'protectedRange', 'ignoredError']) {
    for (const e of [...root.getElementsByTagNameNS('*', name)]) {
      const v = sqref(e.getAttribute('sqref') ?? '');
      if (!v) e.parentNode!.removeChild(e);
      else e.setAttribute('sqref', v);
    }
  }
  for (const name of ['hyperlink', 'autoFilter']) {
    for (const e of [...root.getElementsByTagNameNS('*', name)]) {
      const v = ref(e.getAttribute('ref') ?? '');
      if (v.includes('#REF!')) e.parentNode!.removeChild(e);
      else e.setAttribute('ref', v);
    }
  }
  for (const e of root.getElementsByTagNameNS('*', 'selection')) {
    const a = e.getAttribute('activeCell');
    if (a) e.setAttribute('activeCell', ref(a).replace('#REF!', 'A1'));
  }
  for (const e of [...root.getElementsByTagNameNS('*', 'formula'), ...root.getElementsByTagNameNS('*', 'formula1'), ...root.getElementsByTagNameNS('*', 'formula2')]) {
    e.textContent = ref(e.textContent ?? '');
  }
}
