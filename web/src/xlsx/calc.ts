import { format as ssfFormat } from 'ssf';
import { decodeCol, encodeCol } from './workbook';

// A small Excel formula engine: parser, evaluator and the commonly used functions, so a cell
// edited on the phone updates the formulas that depend on it. A formula using a function not
// implemented here raises Unsupported; its cell then keeps the value Excel saved, and the saved
// file asks Excel to recalculate everything when it is next opened.

export class XError {
  constructor(readonly code: string) {}
  toString(): string {
    return this.code;
  }
}

export class Unsupported extends Error {}

export type Value = number | string | boolean | XError | null;

export interface RangeRef {
  sheet: number;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** What the evaluator reads cells through. */
export interface Book {
  /** Value of a cell (formulas already evaluated). */
  value(sheet: number, r: number, c: number): Value;
  sheetIndex(name: string): number;
  /** Rows / columns in use on a sheet, to bound whole-column references. */
  extent(sheet: number): { rows: number; cols: number };
  /** A workbook-level name (`TaxRate`) as a reference, if it is one. */
  name(name: string, sheet: number): RangeRef | null;
}

const MAX_ROW = 1048575;
const MAX_COL = 16383;
const ERR = {
  div0: new XError('#DIV/0!'), na: new XError('#N/A'), name: new XError('#NAME?'), num: new XError('#NUM!'),
  ref: new XError('#REF!'), value: new XError('#VALUE!'),
};

// ------------------------------------------------------------------ parser

type Node =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  | { k: 'err'; v: XError }
  | { k: 'ref'; sheet: string | null; r0: number; c0: number; r1: number; c1: number; single: boolean }
  | { k: 'name'; v: string }
  | { k: 'fn'; name: string; args: Node[] }
  | { k: 'un'; op: string; a: Node }
  | { k: 'pct'; a: Node }
  | { k: 'bin'; op: string; a: Node; b: Node }
  | { k: 'arr'; rows: Node[][] };

type Tok = { t: string; v: string };

const CELL = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;

function tokenize(f: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (ch === '"') {
      let s = '';
      i++;
      for (;;) {
        if (i >= f.length) throw new SyntaxError('unclosed string');
        if (f[i] === '"') {
          if (f[i + 1] === '"') { s += '"'; i += 2; continue; }
          i++;
          break;
        }
        s += f[i++];
      }
      out.push({ t: 'str', v: s });
      continue;
    }
    if (ch === "'") {
      // Quoted sheet name, then !reference.
      let s = '';
      i++;
      for (;;) {
        if (i >= f.length) throw new SyntaxError('unclosed sheet name');
        if (f[i] === "'") {
          if (f[i + 1] === "'") { s += "'"; i += 2; continue; }
          i++;
          break;
        }
        s += f[i++];
      }
      if (f[i] !== '!') throw new SyntaxError('sheet name without reference');
      i++;
      out.push({ t: 'sheet', v: s });
      continue;
    }
    if (ch === '#') {
      const m = /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|GETTING_DATA|SPILL!|CALC!)/i.exec(f.slice(i));
      if (!m) throw new SyntaxError('bad error literal');
      out.push({ t: 'err', v: m[0].toUpperCase() });
      i += m[0].length;
      continue;
    }
    const num = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/.exec(f.slice(i));
    if (num && !/^\d+:\$?\d/.test(f.slice(i)) && !/^\d+[A-Za-z_]/.test(f.slice(i))) {
      out.push({ t: 'num', v: num[0] });
      i += num[0].length;
      continue;
    }
    const word = /^[A-Za-z_\\$][A-Za-z0-9_.$\\]*(?:!)?|^\d+:\$?\d+/.exec(f.slice(i));
    if (word) {
      let w = word[0];
      if (w.endsWith('!')) {
        out.push({ t: 'sheet', v: w.slice(0, -1) });
        i += w.length;
        continue;
      }
      // A:B, $A:$C, 3:5 and A1:B2 are read as one token.
      const range = /^(\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![A-Za-z0-9_(])/.exec(f.slice(i));
      if (range) w = range[0];
      i += w.length;
      if (f[i] === '(' && !range) out.push({ t: 'fn', v: w });
      else out.push({ t: 'word', v: w });
      continue;
    }
    const two = f.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') {
      out.push({ t: 'op', v: two });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%(),:;{}'.includes(ch)) {
      out.push({ t: 'op', v: ch });
      i++;
      continue;
    }
    throw new SyntaxError(`unexpected ${ch}`);
  }
  return out;
}

function refNode(text: string, sheet: string | null): Node | null {
  const [a, b] = text.split(':');
  const ca = CELL.exec(a);
  if (ca && (!b || CELL.exec(b))) {
    const cb = b ? CELL.exec(b)! : ca;
    const r0 = Number(ca[2]) - 1;
    const r1 = Number(cb[2]) - 1;
    const c0 = decodeCol(ca[1]);
    const c1 = decodeCol(cb[1]);
    return { k: 'ref', sheet, r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1), single: !b };
  }
  if (!b) return null;
  if (/^\$?[A-Za-z]{1,3}$/.test(a) && /^\$?[A-Za-z]{1,3}$/.test(b)) {
    const c0 = decodeCol(a.replace('$', ''));
    const c1 = decodeCol(b.replace('$', ''));
    return { k: 'ref', sheet, r0: 0, r1: MAX_ROW, c0: Math.min(c0, c1), c1: Math.max(c0, c1), single: false };
  }
  if (/^\$?\d+$/.test(a) && /^\$?\d+$/.test(b)) {
    const r0 = Number(a.replace('$', '')) - 1;
    const r1 = Number(b.replace('$', '')) - 1;
    return { k: 'ref', sheet, r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: 0, c1: MAX_COL, single: false };
  }
  return null;
}

const PRECEDENCE: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

export function parse(formula: string): Node {
  const toks = tokenize(formula.replace(/^=/, ''));
  let i = 0;
  const peek = () => toks[i];
  const next = () => toks[i++];
  const expect = (v: string) => {
    const t = next();
    if (!t || t.v !== v) throw new SyntaxError(`expected ${v}`);
  };

  const primary = (): Node => {
    const t = next();
    if (!t) throw new SyntaxError('unexpected end');
    if (t.t === 'num') return { k: 'num', v: Number(t.v) };
    if (t.t === 'str') return { k: 'str', v: t.v };
    if (t.t === 'err') return { k: 'err', v: new XError(t.v) };
    if (t.t === 'op' && t.v === '(') {
      const e = expr(0);
      expect(')');
      return e;
    }
    if (t.t === 'op' && (t.v === '-' || t.v === '+')) return { k: 'un', op: t.v, a: unary() };
    if (t.t === 'op' && t.v === '{') {
      const rows: Node[][] = [[]];
      for (;;) {
        rows[rows.length - 1].push(unary());
        const sep = next();
        if (sep?.v === '}') break;
        if (sep?.v === ';') rows.push([]);
        else if (sep?.v !== ',') throw new SyntaxError('bad array');
      }
      return { k: 'arr', rows };
    }
    if (t.t === 'sheet') {
      const r = next();
      const node = r && refNode(r.v, t.v);
      if (!node) throw new SyntaxError('bad sheet reference');
      return rangeTail(node);
    }
    if (t.t === 'fn') {
      expect('(');
      const args: Node[] = [];
      if (peek()?.v !== ')') {
        for (;;) {
          // An empty argument (f(a,,b)) is a missing value.
          if (peek()?.v === ',' || peek()?.v === ')') args.push({ k: 'str', v: '' });
          else args.push(expr(0));
          const sep = next();
          if (sep?.v === ')') break;
          if (sep?.v !== ',') throw new SyntaxError('expected , or )');
        }
      } else next();
      return { k: 'fn', name: t.v.toUpperCase().replace(/^_XLFN\.|^_XLWS\./, ''), args };
    }
    if (t.t === 'word') {
      const up = t.v.toUpperCase();
      if (up === 'TRUE' || up === 'FALSE') return { k: 'bool', v: up === 'TRUE' };
      const ref = refNode(t.v, null);
      if (ref) return rangeTail(ref);
      return { k: 'name', v: t.v };
    }
    throw new SyntaxError(`unexpected ${t.v}`);
  };

  // A1:B2 where the two ends came as separate tokens (e.g. A1:INDEX(...)) is not supported;
  // ranges arrive as one token from the tokenizer.
  const rangeTail = (node: Node): Node => node;

  const postfix = (): Node => {
    let n = primary();
    while (peek()?.v === '%') {
      next();
      n = { k: 'pct', a: n };
    }
    return n;
  };
  const unary = (): Node => {
    const t = peek();
    if (t?.t === 'op' && (t.v === '-' || t.v === '+')) {
      next();
      return { k: 'un', op: t.v, a: unary() };
    }
    return postfix();
  };
  const expr = (min: number): Node => {
    let left = unary();
    for (;;) {
      const t = peek();
      const p = t?.t === 'op' ? PRECEDENCE[t.v] : undefined;
      if (p === undefined || p < min) return left;
      next();
      // ^ is left-associative in Excel too.
      const right = expr(p + 1);
      left = { k: 'bin', op: t!.v, a: left, b: right };
    }
  };
  const tree = expr(0);
  if (i < toks.length) throw new SyntaxError('trailing input');
  return tree;
}

// ------------------------------------------------------------------ values

interface Range {
  k: 'range';
  ref: RangeRef;
}

type Val = Value | Range | Value[][];

const isRange = (v: Val): v is Range => typeof v === 'object' && v !== null && (v as Range).k === 'range';
const isArr = (v: Val): v is Value[][] => Array.isArray(v);

export function evaluate(formula: string, sheet: number, row: number, col: number, book: Book): Value {
  let tree: Node;
  try {
    tree = parse(formula);
  } catch {
    throw new Unsupported(`cannot read ${formula}`);
  }
  const ev = new Evaluator(book, sheet, row, col);
  const v = ev.scalar(ev.eval(tree));
  return typeof v === 'number' && !Number.isFinite(v) ? ERR.num : v;
}

class Evaluator {
  constructor(private book: Book, private sheet: number, private row: number, private col: number) {}

  eval(n: Node): Val {
    switch (n.k) {
      case 'num': return n.v;
      case 'str': return n.v;
      case 'bool': return n.v;
      case 'err': return n.v;
      case 'arr': return n.rows.map((r) => r.map((c) => this.scalar(this.eval(c))));
      case 'ref': {
        const sheet = n.sheet === null ? this.sheet : this.book.sheetIndex(n.sheet);
        if (sheet < 0) return ERR.ref;
        return { k: 'range', ref: { sheet, r0: n.r0, c0: n.c0, r1: n.r1, c1: n.c1 } };
      }
      case 'name': {
        const ref = this.book.name(n.v, this.sheet);
        if (!ref) throw new Unsupported(`name ${n.v}`);
        return { k: 'range', ref };
      }
      case 'un': {
        const v = this.num(this.eval(n.a));
        return v instanceof XError ? v : n.op === '-' ? -v : v;
      }
      case 'pct': {
        const v = this.num(this.eval(n.a));
        return v instanceof XError ? v : v / 100;
      }
      case 'bin': return this.binary(n.op, this.eval(n.a), this.eval(n.b));
      case 'fn': return this.call(n.name, n.args);
    }
  }

  /** A single value: a one-cell range's value, or the cell of a range in this row/column. */
  scalar(v: Val): Value {
    if (isArr(v)) return v[0]?.[0] ?? null;
    if (!isRange(v)) return v;
    const { ref } = v;
    if (ref.r0 === ref.r1 && ref.c0 === ref.c1) return this.book.value(ref.sheet, ref.r0, ref.c0);
    // Implicit intersection with the formula's own row or column.
    if (ref.c0 === ref.c1 && this.row >= ref.r0 && this.row <= ref.r1) return this.book.value(ref.sheet, this.row, ref.c0);
    if (ref.r0 === ref.r1 && this.col >= ref.c0 && this.col <= ref.c1) return this.book.value(ref.sheet, ref.r0, this.col);
    return ERR.value;
  }

  /** A range or array as rows of values (whole columns cut to the used area). */
  grid(v: Val): Value[][] {
    if (isArr(v)) return v;
    if (!isRange(v)) return [[v as Value]];
    const { ref } = v;
    const ext = this.book.extent(ref.sheet);
    const r1 = Math.min(ref.r1, ext.rows - 1);
    const c1 = Math.min(ref.c1, ext.cols - 1);
    const out: Value[][] = [];
    for (let r = ref.r0; r <= r1; r++) {
      const row: Value[] = [];
      for (let c = ref.c0; c <= c1; c++) row.push(this.book.value(ref.sheet, r, c));
      out.push(row);
    }
    return out;
  }

  flat(v: Val): Value[] {
    return this.grid(v).flat();
  }

  num(v: Val): number | XError {
    const s = this.scalar(v);
    if (s instanceof XError) return s;
    if (s === null || s === '') return 0;
    if (typeof s === 'boolean') return s ? 1 : 0;
    if (typeof s === 'number') return s;
    const n = toNumber(s);
    return n === null ? ERR.value : n;
  }

  str(v: Val): string | XError {
    const s = this.scalar(v);
    if (s instanceof XError) return s;
    return text(s);
  }

  bool(v: Val): boolean | XError {
    const s = this.scalar(v);
    if (s instanceof XError) return s;
    if (typeof s === 'boolean') return s;
    if (typeof s === 'number') return s !== 0;
    if (s === null || s === '') return false;
    const up = s.toUpperCase();
    if (up === 'TRUE') return true;
    if (up === 'FALSE') return false;
    return ERR.value;
  }

  binary(op: string, av: Val, bv: Val): Value {
    if (op === '&') {
      const a = this.str(av);
      const b = this.str(bv);
      if (a instanceof XError) return a;
      if (b instanceof XError) return b;
      return a + b;
    }
    if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
      const a = this.scalar(av);
      const b = this.scalar(bv);
      if (a instanceof XError) return a;
      if (b instanceof XError) return b;
      const c = compare(a, b);
      return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    }
    const a = this.num(av);
    const b = this.num(bv);
    if (a instanceof XError) return a;
    if (b instanceof XError) return b;
    switch (op) {
      case '+': return a + b;
      case '-': return a - b;
      case '*': return a * b;
      case '/': return b === 0 ? ERR.div0 : a / b;
      case '^': return a === 0 && b < 0 ? ERR.div0 : Math.pow(a, b);
    }
    throw new Unsupported(op);
  }

  // ---------------------------------------------------------------- functions

  call(name: string, argNodes: Node[]): Val {
    // Functions that must not evaluate every argument first.
    switch (name) {
      case 'IF': {
        const c = this.bool(this.eval(argNodes[0]));
        if (c instanceof XError) return c;
        if (c) return argNodes[1] ? this.eval(argNodes[1]) : true;
        return argNodes[2] ? this.eval(argNodes[2]) : false;
      }
      case 'IFS': {
        for (let i = 0; i + 1 < argNodes.length; i += 2) {
          const c = this.bool(this.eval(argNodes[i]));
          if (c instanceof XError) return c;
          if (c) return this.eval(argNodes[i + 1]);
        }
        return ERR.na;
      }
      case 'IFERROR':
      case 'IFNA': {
        const v = this.eval(argNodes[0]);
        const s = isRange(v) ? this.scalar(v) : v;
        const bad = s instanceof XError && (name === 'IFERROR' || s.code === '#N/A');
        return bad ? this.eval(argNodes[1]) : v;
      }
      case 'CHOOSE': {
        const k = this.num(this.eval(argNodes[0]));
        if (k instanceof XError) return k;
        const pick = argNodes[Math.trunc(k)];
        return k >= 1 && pick ? this.eval(pick) : ERR.value;
      }
      case 'ROW':
      case 'COLUMN': {
        if (!argNodes.length) return (name === 'ROW' ? this.row : this.col) + 1;
        const v = this.eval(argNodes[0]);
        if (!isRange(v)) return ERR.value;
        return (name === 'ROW' ? v.ref.r0 : v.ref.c0) + 1;
      }
      case 'ROWS':
      case 'COLUMNS': {
        const v = this.eval(argNodes[0]);
        if (isRange(v)) return name === 'ROWS' ? v.ref.r1 - v.ref.r0 + 1 : v.ref.c1 - v.ref.c0 + 1;
        const g = this.grid(v);
        return name === 'ROWS' ? g.length : g[0]?.length ?? 0;
      }
      case 'INDEX': return this.index(argNodes);
      case 'OFFSET': return this.offset(argNodes);
    }
    const fn = FUNCTIONS[name];
    if (!fn) throw new Unsupported(name);
    return fn(argNodes.map((a) => this.eval(a)), this);
  }

  private index(argNodes: Node[]): Val {
    const area = this.eval(argNodes[0]);
    const r = argNodes[1] ? this.num(this.eval(argNodes[1])) : 0;
    const c = argNodes[2] ? this.num(this.eval(argNodes[2])) : 0;
    if (r instanceof XError) return r;
    if (c instanceof XError) return c;
    if (isRange(area)) {
      const { ref } = area;
      const rows = ref.r1 - ref.r0 + 1;
      const cols = ref.c1 - ref.c0 + 1;
      let ri = r;
      let ci = c;
      if (rows === 1 && !argNodes[2]) { ci = r; ri = 1; }
      if (ri > rows || ci > cols || ri < 0 || ci < 0) return ERR.ref;
      const rr0 = ri ? ref.r0 + ri - 1 : ref.r0;
      const rr1 = ri ? rr0 : ref.r1;
      const cc0 = ci ? ref.c0 + ci - 1 : ref.c0;
      const cc1 = ci ? cc0 : ref.c1;
      return { k: 'range', ref: { sheet: ref.sheet, r0: rr0, r1: rr1, c0: cc0, c1: cc1 } };
    }
    const g = this.grid(area);
    const row = g[(r || 1) - 1];
    return row?.[(c || 1) - 1] ?? ERR.ref;
  }

  private offset(argNodes: Node[]): Val {
    const base = this.eval(argNodes[0]);
    if (!isRange(base)) return ERR.value;
    const n = (i: number, d: number) => {
      if (!argNodes[i] || (argNodes[i].k === 'str' && (argNodes[i] as { v: string }).v === '')) return d;
      const v = this.num(this.eval(argNodes[i]));
      return v instanceof XError ? NaN : Math.trunc(v);
    };
    const { ref } = base;
    const r0 = ref.r0 + n(1, 0);
    const c0 = ref.c0 + n(2, 0);
    const h = n(3, ref.r1 - ref.r0 + 1);
    const w = n(4, ref.c1 - ref.c0 + 1);
    if ([r0, c0, h, w].some(Number.isNaN) || r0 < 0 || c0 < 0 || h < 1 || w < 1) return ERR.ref;
    return { k: 'range', ref: { sheet: ref.sheet, r0, c0, r1: r0 + h - 1, c1: c0 + w - 1 } };
  }
}

// ------------------------------------------------------------------ helpers

export function toNumber(s: string): number | null {
  const t = s.trim();
  if (!t) return null;
  const pct = t.endsWith('%');
  const n = Number((pct ? t.slice(0, -1) : t).replace(/,(?=\d{3}(\D|$))/g, ''));
  if (!Number.isFinite(n)) return null;
  return pct ? n / 100 : n;
}

/** A number as Excel's General format writes it into text (up to 15 significant digits). */
export function numberText(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  const s = String(Number(n.toPrecision(15)));
  return s.replace(/e\+?/, 'E+').replace('E+-', 'E-');
}

export function text(v: Value): string {
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numberText(v);
  return String(v);
}

/** Excel's ordering: numbers < text < booleans; text compares without case. */
function compare(a: Value, b: Value): number {
  const rank = (v: Value) => (v === null ? -1 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'string') {
    const x = a.toLowerCase();
    const y = (b as string).toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  const x = Number(a);
  const y = Number(b);
  return x < y ? -1 : x > y ? 1 : 0;
}

/** A COUNTIF/SUMIF criterion as a test. */
function criterion(c: Value): (v: Value) => boolean {
  if (typeof c === 'number' || typeof c === 'boolean') return (v) => v !== null && compare(v, c) === 0 && typeof v === typeof c;
  const s = c === null ? '' : String(c);
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(s)!;
  const op = m[1] ?? '=';
  const rest = m[2];
  const n = toNumber(rest);
  if (n !== null && rest !== '') {
    return (v) => {
      const x = typeof v === 'number' ? v : typeof v === 'string' ? toNumber(v) : null;
      if (x === null) return op === '<>';
      return op === '=' ? x === n : op === '<>' ? x !== n : op === '<' ? x < n : op === '>' ? x > n : op === '<=' ? x <= n : x >= n;
    };
  }
  if (op === '=' || op === '<>') {
    const re = wildcard(rest);
    const eq = (v: Value) => (rest === '' ? v === null || v === '' : typeof v === 'string' && re.test(v));
    return op === '=' ? eq : (v) => !eq(v);
  }
  return (v) => typeof v === 'string' && (op === '<' ? compare(v, rest) < 0 : op === '>' ? compare(v, rest) > 0 : op === '<=' ? compare(v, rest) <= 0 : compare(v, rest) >= 0);
}

function wildcard(p: string): RegExp {
  let re = '';
  for (let i = 0; i < p.length; i++) {
    const ch = p[i];
    if (ch === '~' && i + 1 < p.length) re += p[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (ch === '*') re += '.*';
    else if (ch === '?') re += '.';
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`, 'is');
}

const firstError = (vals: Value[]): XError | null => (vals.find((v) => v instanceof XError) as XError) ?? null;

/** Numbers among values from ranges (text and booleans in ranges are skipped, as Excel does). */
function numbers(args: Val[], ev: Evaluator): number[] | XError {
  const out: number[] = [];
  for (const a of args) {
    if (isRange(a) || isArr(a)) {
      for (const v of ev.flat(a)) {
        if (v instanceof XError) return v;
        if (typeof v === 'number') out.push(v);
      }
    } else {
      const n = ev.num(a);
      if (n instanceof XError) return n;
      out.push(n);
    }
  }
  return out;
}

const DAY_MS = 86400000;
/** Serial number ↔ date, 1900 date system (with Excel's 29 Feb 1900). */
export function toSerial(y: number, m: number, d: number): number {
  const ms = Date.UTC(y, m - 1, d);
  const serial = ms / DAY_MS + 25569;
  return serial < 61 ? serial - 1 : serial;
}

function fromSerial(n: number): Date {
  const s = n < 61 ? n + 1 : n;
  return new Date(Math.round((s - 25569) * DAY_MS));
}

function lookupMatch(needle: Value, list: Value[], exact: boolean): number {
  if (exact) {
    const test = typeof needle === 'string' ? criterion(`=${needle}`) : (v: Value) => v !== null && compare(v, needle) === 0;
    return list.findIndex((v) => test(v));
  }
  // Approximate: the last value <= needle, in a list sorted ascending.
  let found = -1;
  for (let i = 0; i < list.length; i++) {
    const v = list[i];
    if (v === null) continue;
    if (compare(v, needle) <= 0 && typeof v === typeof needle) found = i;
    else if (compare(v, needle) > 0) break;
  }
  return found;
}

type Fn = (args: Val[], ev: Evaluator) => Val;

const math = (f: (x: number) => number): Fn => (a, ev) => {
  const x = ev.num(a[0]);
  return x instanceof XError ? x : f(x);
};

const roundTo = (mode: 'round' | 'up' | 'down'): Fn => (a, ev) => {
  const x = ev.num(a[0]);
  const d = a[1] === undefined ? 0 : ev.num(a[1]);
  if (x instanceof XError) return x;
  if (d instanceof XError) return d;
  const k = Math.pow(10, Math.trunc(d));
  const v = Math.abs(x) * k;
  const r = mode === 'round' ? Math.round(Number(v.toPrecision(15))) : mode === 'up' ? Math.ceil(Number(v.toPrecision(15))) : Math.floor(Number(v.toPrecision(15)));
  return (Math.sign(x) * r) / k;
};

/** COUNTIF(S)(range, crit, …) and SUMIF / AVERAGEIF(range, crit, [sum_range]). */
const conditional = (kind: 'count' | 'sum' | 'avg'): Fn => (a, ev) => {
  const pairs: [Value[], (v: Value) => boolean][] = [];
  let target: Value[] | null = null;
  if (kind === 'count') {
    for (let i = 0; i + 1 < a.length; i += 2) pairs.push([ev.flat(a[i]), criterion(ev.scalar(a[i + 1]))]);
  } else {
    pairs.push([ev.flat(a[0]), criterion(ev.scalar(a[1]))]);
    target = a[2] ? ev.flat(a[2]) : pairs[0][0];
  }
  let count = 0;
  let sum = 0;
  for (let i = 0; i < pairs[0][0].length; i++) {
    if (!pairs.every(([vals, test]) => test(vals[i] ?? null))) continue;
    if (!target) {
      count++;
      continue;
    }
    const v = target[i];
    if (typeof v === 'number') {
      sum += v;
      count++;
    }
  }
  return kind === 'count' ? count : kind === 'sum' ? sum : count ? sum / count : ERR.div0;
};

const conditionalMulti = (kind: 'sum' | 'avg'): Fn => (a, ev) => {
  // SUMIFS(sum_range, crit_range1, crit1, …)
  const target = ev.flat(a[0]);
  const pairs: [Value[], (v: Value) => boolean][] = [];
  for (let i = 1; i + 1 < a.length; i += 2) pairs.push([ev.flat(a[i]), criterion(ev.scalar(a[i + 1]))]);
  let count = 0;
  let sum = 0;
  for (let i = 0; i < target.length; i++) {
    if (!pairs.every(([vals, test]) => test(vals[i] ?? null))) continue;
    const v = target[i];
    if (typeof v === 'number') {
      sum += v;
      count++;
    }
  }
  return kind === 'sum' ? sum : count ? sum / count : ERR.div0;
};

const FUNCTIONS: Record<string, Fn> = {
  SUM: (a, ev) => {
    const n = numbers(a, ev);
    return n instanceof XError ? n : n.reduce((s, x) => s + x, 0);
  },
  PRODUCT: (a, ev) => {
    const n = numbers(a, ev);
    return n instanceof XError ? n : n.reduce((s, x) => s * x, 1);
  },
  AVERAGE: (a, ev) => {
    const n = numbers(a, ev);
    if (n instanceof XError) return n;
    return n.length ? n.reduce((s, x) => s + x, 0) / n.length : ERR.div0;
  },
  MIN: (a, ev) => {
    const n = numbers(a, ev);
    return n instanceof XError ? n : n.length ? Math.min(...n) : 0;
  },
  MAX: (a, ev) => {
    const n = numbers(a, ev);
    return n instanceof XError ? n : n.length ? Math.max(...n) : 0;
  },
  MEDIAN: (a, ev) => {
    const n = numbers(a, ev);
    if (n instanceof XError) return n;
    if (!n.length) return ERR.num;
    const s = [...n].sort((x, y) => x - y);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  },
  LARGE: (a, ev) => {
    const n = numbers([a[0]], ev);
    const k = ev.num(a[1]);
    if (n instanceof XError) return n;
    if (k instanceof XError) return k;
    const s = [...n].sort((x, y) => y - x);
    return s[Math.ceil(k) - 1] ?? ERR.num;
  },
  SMALL: (a, ev) => {
    const n = numbers([a[0]], ev);
    const k = ev.num(a[1]);
    if (n instanceof XError) return n;
    if (k instanceof XError) return k;
    const s = [...n].sort((x, y) => x - y);
    return s[Math.ceil(k) - 1] ?? ERR.num;
  },
  'STDEV.S': (a, ev) => stdev(a, ev, 1),
  STDEV: (a, ev) => stdev(a, ev, 1),
  'STDEV.P': (a, ev) => stdev(a, ev, 0),
  COUNT: (a, ev) => a.reduce<number>((s, x) => s + (isRange(x) || isArr(x) ? ev.flat(x).filter((v) => typeof v === 'number').length : typeof ev.scalar(x) === 'number' ? 1 : 0), 0),
  COUNTA: (a, ev) => a.reduce<number>((s, x) => s + (isRange(x) || isArr(x) ? ev.flat(x).filter((v) => v !== null && v !== '').length : ev.scalar(x) !== null ? 1 : 0), 0),
  COUNTBLANK: (a, ev) => ev.flat(a[0]).filter((v) => v === null || v === '').length,
  COUNTIF: conditional('count'),
  COUNTIFS: conditional('count'),
  SUMIF: conditional('sum'),
  AVERAGEIF: conditional('avg'),
  SUMIFS: conditionalMulti('sum'),
  AVERAGEIFS: conditionalMulti('avg'),
  SUMPRODUCT: (a, ev) => {
    const grids = a.map((x) => ev.flat(x));
    const n = grids[0]?.length ?? 0;
    if (grids.some((g) => g.length !== n)) return ERR.value;
    let s = 0;
    for (let i = 0; i < n; i++) s += grids.reduce((p, g) => p * (typeof g[i] === 'number' ? (g[i] as number) : 0), 1);
    return s;
  },
  ABS: math(Math.abs),
  SQRT: (a, ev) => {
    const x = ev.num(a[0]);
    return x instanceof XError ? x : x < 0 ? ERR.num : Math.sqrt(x);
  },
  INT: math(Math.floor),
  TRUNC: (a, ev) => roundTo('down')(a.length > 1 ? a : [a[0], 0], ev),
  SIGN: math(Math.sign),
  EXP: math(Math.exp),
  LN: math(Math.log),
  LOG10: math(Math.log10),
  LOG: (a, ev) => {
    const x = ev.num(a[0]);
    const b = a[1] === undefined ? 10 : ev.num(a[1]);
    if (x instanceof XError) return x;
    if (b instanceof XError) return b;
    return Math.log(x) / Math.log(b);
  },
  PI: () => Math.PI,
  POWER: (a, ev) => ev.binary('^', a[0], a[1]),
  MOD: (a, ev) => {
    const x = ev.num(a[0]);
    const d = ev.num(a[1]);
    if (x instanceof XError) return x;
    if (d instanceof XError) return d;
    if (d === 0) return ERR.div0;
    return x - d * Math.floor(x / d);
  },
  ROUND: roundTo('round'),
  ROUNDUP: roundTo('up'),
  ROUNDDOWN: roundTo('down'),
  CEILING: (a, ev) => {
    const x = ev.num(a[0]);
    const s = a[1] === undefined ? 1 : ev.num(a[1]);
    if (x instanceof XError) return x;
    if (s instanceof XError) return s;
    return s === 0 ? 0 : Math.ceil(x / s) * s;
  },
  FLOOR: (a, ev) => {
    const x = ev.num(a[0]);
    const s = a[1] === undefined ? 1 : ev.num(a[1]);
    if (x instanceof XError) return x;
    if (s instanceof XError) return s;
    return s === 0 ? ERR.div0 : Math.floor(x / s) * s;
  },
  AND: (a, ev) => logical(a, ev, (v) => v.every(Boolean)),
  OR: (a, ev) => logical(a, ev, (v) => v.some(Boolean)),
  XOR: (a, ev) => logical(a, ev, (v) => v.filter(Boolean).length % 2 === 1),
  NOT: (a, ev) => {
    const b = ev.bool(a[0]);
    return b instanceof XError ? b : !b;
  },
  TRUE: () => true,
  FALSE: () => false,
  NA: () => ERR.na,
  ISBLANK: (a, ev) => ev.scalar(a[0]) === null,
  ISNUMBER: (a, ev) => typeof ev.scalar(a[0]) === 'number',
  ISTEXT: (a, ev) => typeof ev.scalar(a[0]) === 'string',
  ISLOGICAL: (a, ev) => typeof ev.scalar(a[0]) === 'boolean',
  ISERROR: (a, ev) => ev.scalar(a[0]) instanceof XError,
  ISERR: (a, ev) => {
    const v = ev.scalar(a[0]);
    return v instanceof XError && v.code !== '#N/A';
  },
  ISNA: (a, ev) => {
    const v = ev.scalar(a[0]);
    return v instanceof XError && v.code === '#N/A';
  },
  N: (a, ev) => {
    const v = ev.scalar(a[0]);
    return typeof v === 'number' ? v : typeof v === 'boolean' ? +v : 0;
  },
  T: (a, ev) => {
    const v = ev.scalar(a[0]);
    return typeof v === 'string' ? v : '';
  },
  // Text
  CONCATENATE: (a, ev) => joinText(a.map((x) => ev.str(x))),
  CONCAT: (a, ev) => joinText(a.flatMap((x) => (isRange(x) || isArr(x) ? ev.flat(x).map(text) : [ev.str(x)]))),
  TEXTJOIN: (a, ev) => {
    const sep = ev.str(a[0]);
    const skip = ev.bool(a[1]);
    if (sep instanceof XError) return sep;
    if (skip instanceof XError) return skip;
    const parts = a.slice(2).flatMap((x) => (isRange(x) || isArr(x) ? ev.flat(x) : [ev.scalar(x)]));
    const err = firstError(parts);
    if (err) return err;
    return parts.map(text).filter((s) => !skip || s !== '').join(sep);
  },
  LEN: (a, ev) => {
    const s = ev.str(a[0]);
    return s instanceof XError ? s : s.length;
  },
  LEFT: (a, ev) => textCut(a, ev, (s, n) => s.slice(0, n)),
  RIGHT: (a, ev) => textCut(a, ev, (s, n) => (n ? s.slice(-n) : '')),
  MID: (a, ev) => {
    const s = ev.str(a[0]);
    const start = ev.num(a[1]);
    const n = ev.num(a[2]);
    if (s instanceof XError) return s;
    if (start instanceof XError) return start;
    if (n instanceof XError) return n;
    if (start < 1 || n < 0) return ERR.value;
    return s.substr(start - 1, n);
  },
  UPPER: (a, ev) => mapText(a, ev, (s) => s.toUpperCase()),
  LOWER: (a, ev) => mapText(a, ev, (s) => s.toLowerCase()),
  PROPER: (a, ev) => mapText(a, ev, (s) => s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, p, c) => p + c.toUpperCase())),
  TRIM: (a, ev) => mapText(a, ev, (s) => s.trim().replace(/ +/g, ' ')),
  REPT: (a, ev) => {
    const s = ev.str(a[0]);
    const n = ev.num(a[1]);
    if (s instanceof XError) return s;
    if (n instanceof XError) return n;
    return n < 0 ? ERR.value : s.repeat(Math.trunc(n));
  },
  SUBSTITUTE: (a, ev) => {
    const [s, from, to] = [ev.str(a[0]), ev.str(a[1]), ev.str(a[2])];
    for (const x of [s, from, to]) if (x instanceof XError) return x;
    const str = s as string;
    if (!from) return str;
    if (a[3] === undefined) return str.split(from as string).join(to as string);
    const nth = ev.num(a[3]);
    if (nth instanceof XError) return nth;
    let idx = -1;
    for (let k = 0; k < nth; k++) {
      idx = str.indexOf(from as string, idx + 1);
      if (idx < 0) return str;
    }
    return str.slice(0, idx) + to + str.slice(idx + (from as string).length);
  },
  FIND: (a, ev) => findText(a, ev, true),
  SEARCH: (a, ev) => findText(a, ev, false),
  CHAR: (a, ev) => {
    const n = ev.num(a[0]);
    return n instanceof XError ? n : n < 1 || n > 255 ? ERR.value : String.fromCharCode(Math.trunc(n));
  },
  UNICHAR: (a, ev) => {
    const n = ev.num(a[0]);
    return n instanceof XError ? n : n < 1 ? ERR.value : String.fromCodePoint(Math.trunc(n));
  },
  CODE: (a, ev) => {
    const s = ev.str(a[0]);
    return s instanceof XError ? s : s ? s.charCodeAt(0) : ERR.value;
  },
  UNICODE: (a, ev) => {
    const s = ev.str(a[0]);
    return s instanceof XError ? s : s ? s.codePointAt(0)! : ERR.value;
  },
  EXACT: (a, ev) => {
    const x = ev.str(a[0]);
    const y = ev.str(a[1]);
    if (x instanceof XError) return x;
    if (y instanceof XError) return y;
    return x === y;
  },
  VALUE: (a, ev) => {
    const v = ev.scalar(a[0]);
    if (typeof v === 'number') return v;
    const n = toNumber(text(v));
    return n === null ? ERR.value : n;
  },
  TEXT: (a, ev) => {
    const v = ev.scalar(a[0]);
    const f = ev.str(a[1]);
    if (v instanceof XError) return v;
    if (f instanceof XError) return f;
    try {
      return ssfFormat(f, typeof v === 'string' ? toNumber(v) ?? v : v ?? 0);
    } catch {
      return ERR.value;
    }
  },
  // Lookup
  VLOOKUP: (a, ev) => tableLookup(a, ev, true),
  HLOOKUP: (a, ev) => tableLookup(a, ev, false),
  MATCH: (a, ev) => {
    const needle = ev.scalar(a[0]);
    if (needle instanceof XError) return needle;
    const list = ev.flat(a[1]);
    const type = a[2] === undefined ? 1 : ev.num(a[2]);
    if (type instanceof XError) return type;
    let i: number;
    if (type === 0) i = lookupMatch(needle, list, true);
    else if (type > 0) i = lookupMatch(needle, list, false);
    else {
      i = -1;
      for (let k = 0; k < list.length; k++) {
        if (list[k] !== null && compare(list[k], needle) >= 0) i = k;
        else break;
      }
    }
    return i < 0 ? ERR.na : i + 1;
  },
  XLOOKUP: (a, ev) => {
    const needle = ev.scalar(a[0]);
    if (needle instanceof XError) return needle;
    const keys = ev.flat(a[1]);
    const values = ev.grid(a[2]);
    const mode = a[4] === undefined ? 0 : ev.num(a[4]);
    if (mode instanceof XError) return mode;
    let i = -1;
    if (mode === 0) i = lookupMatch(needle, keys, true);
    else {
      // -1: exact or next smaller; 1: exact or next larger.
      let best: number | null = null;
      keys.forEach((k, idx) => {
        if (k === null || typeof k !== typeof needle) return;
        const c = compare(k, needle);
        if (c === 0 && i < 0) i = idx;
        if ((mode < 0 && c < 0) || (mode > 0 && c > 0)) {
          if (best === null || (mode < 0 ? compare(k, keys[best]) > 0 : compare(k, keys[best]) < 0)) best = idx;
        }
      });
      if (i < 0 && best !== null) i = best;
    }
    if (i < 0) return a[3] !== undefined && !(isRange(a[3]) === false && ev.scalar(a[3]) === '') ? a[3] : ERR.na;
    // A row of results when looking up in a column, and the other way round.
    return values.length === keys.length ? values[i][0] : values[0][i] ?? ERR.na;
  },
  // Dates
  DATE: (a, ev) => {
    const [y, m, d] = a.map((x) => ev.num(x));
    for (const x of [y, m, d]) if (x instanceof XError) return x;
    let yy = y as number;
    if (yy < 1900) yy += 1900;
    return toSerial(yy, m as number, d as number);
  },
  TIME: (a, ev) => {
    const [h, m, s] = a.map((x) => ev.num(x));
    for (const x of [h, m, s]) if (x instanceof XError) return x;
    return (((h as number) * 3600 + (m as number) * 60 + (s as number)) / 86400) % 1;
  },
  TODAY: () => {
    const now = new Date();
    return toSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
  },
  NOW: () => {
    const now = new Date();
    return toSerial(now.getFullYear(), now.getMonth() + 1, now.getDate()) +
      (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) / 86400;
  },
  YEAR: datePart((d) => d.getUTCFullYear()),
  MONTH: datePart((d) => d.getUTCMonth() + 1),
  DAY: datePart((d) => d.getUTCDate()),
  WEEKDAY: (a, ev) => {
    const n = ev.num(a[0]);
    const type = a[1] === undefined ? 1 : ev.num(a[1]);
    if (n instanceof XError) return n;
    if (type instanceof XError) return type;
    const dow = fromSerial(Math.floor(n)).getUTCDay(); // 0 = Sunday
    return type === 2 ? ((dow + 6) % 7) + 1 : type === 3 ? (dow + 6) % 7 : dow + 1;
  },
  HOUR: (a, ev) => timePart(a, ev, 3600, 24),
  MINUTE: (a, ev) => timePart(a, ev, 60, 60),
  SECOND: (a, ev) => timePart(a, ev, 1, 60),
  DAYS: (a, ev) => ev.binary('-', a[0], a[1]),
  EDATE: (a, ev) => {
    const n = ev.num(a[0]);
    const m = ev.num(a[1]);
    if (n instanceof XError) return n;
    if (m instanceof XError) return m;
    const d = fromSerial(Math.floor(n));
    const target = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + Math.trunc(m), 1));
    const last = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();
    return toSerial(target.getUTCFullYear(), target.getUTCMonth() + 1, Math.min(d.getUTCDate(), last));
  },
  EOMONTH: (a, ev) => {
    const n = ev.num(a[0]);
    const m = ev.num(a[1]);
    if (n instanceof XError) return n;
    if (m instanceof XError) return m;
    const d = fromSerial(Math.floor(n));
    const end = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + Math.trunc(m) + 1, 0));
    return toSerial(end.getUTCFullYear(), end.getUTCMonth() + 1, end.getUTCDate());
  },
  // Finance (the one people use on phones: loan payments)
  PMT: (a, ev) => {
    const [r, n, pv, fv = 0, type = 0] = a.map((x) => ev.num(x));
    for (const x of [r, n, pv, fv, type]) if (x instanceof XError) return x;
    const R = r as number;
    const N = n as number;
    if (R === 0) return -((pv as number) + (fv as number)) / N;
    const p = Math.pow(1 + R, N);
    return (-R * ((pv as number) * p + (fv as number))) / ((1 + R * (type as number)) * (p - 1));
  },
  ADDRESS: (a, ev) => {
    const r = ev.num(a[0]);
    const c = ev.num(a[1]);
    const abs = a[2] === undefined ? 1 : ev.num(a[2]);
    if (r instanceof XError) return r;
    if (c instanceof XError) return c;
    if (abs instanceof XError) return abs;
    const col = encodeCol(c - 1);
    const d = (on: boolean) => (on ? '$' : '');
    return `${d(abs === 1 || abs === 3)}${col}${d(abs === 1 || abs === 2)}${r}`;
  },
};

function stdev(a: Val[], ev: Evaluator, ddof: number): Value {
  const n = numbers(a, ev);
  if (n instanceof XError) return n;
  if (n.length <= ddof) return ERR.div0;
  const mean = n.reduce((s, x) => s + x, 0) / n.length;
  return Math.sqrt(n.reduce((s, x) => s + (x - mean) ** 2, 0) / (n.length - ddof));
}

function logical(a: Val[], ev: Evaluator, f: (v: boolean[]) => boolean): Value {
  const vals: boolean[] = [];
  for (const x of a) {
    if (isRange(x) || isArr(x)) {
      for (const v of ev.flat(x)) {
        if (v instanceof XError) return v;
        if (typeof v === 'boolean') vals.push(v);
        else if (typeof v === 'number') vals.push(v !== 0);
      }
    } else {
      const b = ev.bool(x);
      if (b instanceof XError) return b;
      vals.push(b);
    }
  }
  return vals.length ? f(vals) : ERR.value;
}

function joinText(parts: (string | XError)[]): Value {
  for (const p of parts) if (p instanceof XError) return p;
  return parts.join('');
}

function mapText(a: Val[], ev: Evaluator, f: (s: string) => string): Value {
  const s = ev.str(a[0]);
  return s instanceof XError ? s : f(s);
}

function textCut(a: Val[], ev: Evaluator, f: (s: string, n: number) => string): Value {
  const s = ev.str(a[0]);
  const n = a[1] === undefined ? 1 : ev.num(a[1]);
  if (s instanceof XError) return s;
  if (n instanceof XError) return n;
  return n < 0 ? ERR.value : f(s, Math.trunc(n));
}

function findText(a: Val[], ev: Evaluator, caseSensitive: boolean): Value {
  const needle = ev.str(a[0]);
  const hay = ev.str(a[1]);
  const start = a[2] === undefined ? 1 : ev.num(a[2]);
  if (needle instanceof XError) return needle;
  if (hay instanceof XError) return hay;
  if (start instanceof XError) return start;
  let i: number;
  if (caseSensitive) i = hay.indexOf(needle, start - 1);
  else {
    const m = wildcard(`*${needle}*`.replace(/^\*|\*$/g, '')).source.slice(1, -1);
    const re = new RegExp(m, 'i');
    const found = re.exec(hay.slice(start - 1));
    i = found ? found.index + start - 1 : -1;
  }
  return i < 0 ? ERR.value : i + 1;
}

function datePart(f: (d: Date) => number): Fn {
  return (a, ev) => {
    const n = ev.num(a[0]);
    return n instanceof XError ? n : f(fromSerial(Math.floor(n)));
  };
}

function timePart(a: Val[], ev: Evaluator, unit: number, mod: number): Value {
  const n = ev.num(a[0]);
  if (n instanceof XError) return n;
  const secs = Math.round((n - Math.floor(n)) * 86400);
  return Math.floor(secs / unit) % mod;
}

function tableLookup(a: Val[], ev: Evaluator, vertical: boolean): Value {
  const needle = ev.scalar(a[0]);
  if (needle instanceof XError) return needle;
  const table = ev.grid(a[1]);
  const k = ev.num(a[2]);
  if (k instanceof XError) return k;
  const approx = a[3] === undefined ? true : ev.bool(a[3]);
  if (approx instanceof XError) return approx;
  const keys = vertical ? table.map((r) => r[0] ?? null) : table[0] ?? [];
  const i = lookupMatch(needle, keys, !approx);
  if (i < 0) return ERR.na;
  const idx = Math.trunc(k) - 1;
  const v = vertical ? table[i]?.[idx] : table[idx]?.[i];
  return v === undefined ? ERR.ref : v;
}

// ------------------------------------------------------------------ moving references

export interface Shift {
  /** The sheet whose rows or columns move. */
  sheet: string;
  axis: 'row' | 'col';
  /** First row/column (0-based) that moves. */
  at: number;
  /** Positive: inserted; negative: deleted (at … at-count-1 are gone). */
  count: number;
}

/**
 * Rewrites the references in a formula after rows or columns were inserted or deleted.
 * `formulaSheet` is the name of the sheet the formula lives on (unqualified references point
 * there). References to deleted cells become #REF!.
 */
export function shiftReferences(formula: string, formulaSheet: string, shift: Shift): string {
  // Split out string literals, so text in quotes is never changed.
  return formula.replace(/("(?:[^"]|"")*")|([^"]+)/g, (_, quoted: string, plain: string) =>
    quoted ?? plain.replace(
      /(^|[^A-Za-z0-9_.$!'])((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![A-Za-z0-9_(!])/g,
      (match, pre: string, prefix: string | undefined, ref: string) => {
        const sheet = prefix ? prefix.slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'") : formulaSheet;
        if (sheet.toLowerCase() !== shift.sheet.toLowerCase()) return match;
        const moved = moveRef(ref, shift);
        return `${pre}${prefix ?? ''}${moved}`;
      },
    ));
}

function moveRef(ref: string, shift: Shift): string {
  const parts = ref.split(':');
  type End = { abs: [string, string]; col: number | null; row: number | null };
  const read = (p: string): End => {
    const m = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/.exec(p)!;
    // In a whole-row reference ($3:$5) the only $ belongs to the row.
    const rowAbs = m[2] ? m[3] : m[1] || m[3];
    return { abs: [m[2] ? m[1] : '', rowAbs], col: m[2] ? decodeCol(m[2]) : null, row: m[4] ? Number(m[4]) - 1 : null };
  };
  const ends = parts.map(read);
  const key = shift.axis === 'row' ? 'row' : 'col';
  const vals = ends.map((e) => e[key]);
  if (vals.every((v) => v === null)) return ref; // A:B when moving rows, 3:5 when moving columns
  const { at, count } = shift;
  if (count > 0) {
    for (const e of ends) if (e[key] !== null && e[key]! >= at) e[key]! += count;
  } else {
    const gone0 = at;
    const gone1 = at - count - 1;
    const [a, b] = [ends[0][key]!, ends[ends.length - 1][key]!];
    if (ends.length === 1 || a === null) {
      const v = ends[0][key]!;
      if (v >= gone0 && v <= gone1) return '#REF!';
      if (v > gone1) ends[0][key] = v + count;
    } else {
      if (a >= gone0 && b <= gone1) return '#REF!';
      let na = a;
      let nb = b;
      if (a > gone1) na = a + count;
      else if (a >= gone0) na = gone0;
      if (b > gone1) nb = b + count;
      else if (b >= gone0) nb = gone0 - 1;
      ends[0][key] = na;
      ends[ends.length - 1][key] = nb;
    }
  }
  return ends.map((e) => `${e.col !== null ? e.abs[0] + encodeCol(e.col) : ''}${e.row !== null ? e.abs[1] + (e.row + 1) : ''}`).join(':');
}
