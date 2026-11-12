import { format as ssfFormat, get_table, is_date } from 'ssf';

// SheetJS's table writes the built-in date formats with two-digit years; Excel (en-US) shows four.
const builtIn: Record<number, string> = { ...get_table() };
builtIn[14] = 'm/d/yyyy';
builtIn[22] = 'm/d/yyyy h:mm';

/** The format code for a number format id: the file's own, else Excel's built-in one. */
export function formatCode(id: number, custom: Map<number, string>): string {
  return custom.get(id) ?? builtIn[id] ?? 'General';
}

const COLOURS: Record<string, string> = {
  black: '#000000', blue: '#0000ff', cyan: '#00ffff', green: '#00ff00',
  magenta: '#ff00ff', red: '#ff0000', white: '#ffffff', yellow: '#ffff00',
};

/** Splits a format code into its `;` sections, ignoring semicolons in quotes or after `\`. */
function sections(code: string): string[] {
  const out: string[] = [];
  let start = 0;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '\\') i++;
    else if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      i = end < 0 ? code.length : end;
    }
    else if (ch === ';') {
      out.push(code.slice(start, i));
      start = i + 1;
    }
  }
  out.push(code.slice(start));
  return out;
}

/** The colour a format gives a value, e.g. `[Red]` in the negative section; '' for none. */
function sectionColour(code: string, value: number): string {
  if (!code.includes('[')) return '';
  const parts = sections(code);
  // Conditions like [>100] are rare; pick the section by sign as Excel does without them.
  const section = value < 0 && parts.length > 1 ? parts[1]
    : value === 0 && parts.length > 2 ? parts[2]
    : parts[0];
  const name = /\[(black|blue|cyan|green|magenta|red|white|yellow)\]/i.exec(section)?.[1];
  return name ? COLOURS[name.toLowerCase()] : '';
}

export interface Formatted {
  text: string;
  /** Font colour required by the format (`[Red]`), or ''. */
  colour: string;
  /**
   * Where a `*` fill repeats to fill the cell: text before and after it. Accounting formats use
   * it to put the currency symbol at the left edge and the number at the right.
   */
  split?: [string, string];
}

const FILL = '\u0001';

/** Replaces the fill (`*` + character) with a marker character, or returns null if none. */
function markFill(code: string): string | null {
  if (!code.includes('*')) return null;
  let out = '';
  let found = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '\\' || ch === '_') {
      out += code.slice(i, i + 2);
      i++;
    } else if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const stop = end < 0 ? code.length : end + 1;
      out += code.slice(i, stop);
      i = stop - 1;
    } else if (ch === '*') {
      out += `"${FILL}"`;
      i++;
      found = true;
    } else out += ch;
  }
  return found ? out : null;
}

const marked = new Map<string, string | null>();

export function formatNumber(value: number, code: string, date1904: boolean): Formatted {
  let fill = marked.get(code);
  if (fill === undefined) marked.set(code, (fill = markFill(code)));
  let text: string;
  try {
    text = ssfFormat(fill ?? code, value, { date1904 });
  } catch {
    text = String(value);
  }
  const colour = sectionColour(code, value);
  const at = text.indexOf(FILL);
  if (at < 0) return { text, colour };
  const split: [string, string] = [text.slice(0, at), text.slice(at + 1).replaceAll(FILL, '')];
  return { text: split[0] + split[1], colour, split };
}

/** Text in a cell whose format has a text section (e.g. `0;-0;0;"Name: "@`). */
export function formatText(value: string, code: string): string {
  if (code === 'General' || !code.includes('@')) return value;
  try {
    return ssfFormat(code, value);
  } catch {
    return value;
  }
}

export function isDateFormat(code: string): boolean {
  try {
    return is_date(code);
  } catch {
    return false;
  }
}

/**
 * Excel's General format in a column too narrow for the number: fewer decimals, then
 * scientific notation, as Excel does. `fits` measures a candidate.
 */
export function shrinkGeneral(value: number, fits: (text: string) => boolean): string | null {
  for (let digits = 10; digits >= 1; digits--) {
    const fixed = trimZeros(value.toPrecision(digits));
    if (!/e/i.test(fixed) && fits(fixed)) return fixed;
  }
  for (let digits = 5; digits >= 0; digits--) {
    const [mantissa, exp] = value.toExponential(digits).split('e');
    const sign = exp.startsWith('-') ? '-' : '+';
    const text = `${trimZeros(mantissa)}E${sign}${exp.replace(/^[+-]/, '').padStart(2, '0')}`;
    if (fits(text)) return text;
  }
  return null;
}

function trimZeros(s: string): string {
  return s.includes('.') && !/e/i.test(s) ? s.replace(/\.?0+$/, '') : s;
}
