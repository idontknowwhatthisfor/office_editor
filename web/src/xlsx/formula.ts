import { decodeCol, encodeCol } from './workbook';

const MAX_ROW = 1048576;
const MAX_COL = 16384;

/**
 * The formula of a cell in a shared-formula group: Excel stores it once, in the group's first
 * cell, and the others are that formula with relative references moved by the cell's offset.
 * Text in quotes and quoted sheet names are left alone.
 */
export function shiftFormula(formula: string, dRow: number, dCol: number): string {
  if (!dRow && !dCol) return formula;
  // Split into quoted ("text", 'Sheet name') and unquoted pieces; shift only the unquoted ones.
  return formula.replace(/("(?:[^"]|"")*"|'(?:[^']|'')*')|([^"']+)/g, (_, quoted: string, plain: string) =>
    quoted ?? plain.replace(
      /(^|[^A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(!])/g,
      (match, pre: string, colAbs: string, col: string, rowAbs: string, row: string) => {
        let c = decodeCol(col);
        let r = Number(row) - 1;
        if (!colAbs) c += dCol;
        if (!rowAbs) r += dRow;
        if (c < 0 || r < 0 || c >= MAX_COL || r >= MAX_ROW) return match;
        return `${pre}${colAbs}${encodeCol(c)}${rowAbs}${r + 1}`;
      },
    ),
  );
}
