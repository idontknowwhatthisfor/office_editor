import { child, children } from '../ooxml';
import { EMU } from './deck';
import { borderCss, colour, fillCss, fillElement, readFill, styleRef, type Ctx } from './drawing';
import { bodyProps, levelStyles, mergePara, overlay, renderText, type ParaStyle } from './text';

// Tables (`<a:tbl>` in a graphic frame). Their look comes from a table style: one stored in
// ppt/tableStyles.xml, or one of PowerPoint's built-in styles, which files refer to by id only.

/**
 * PowerPoint's default table style, "Medium Style 2 – Accent 1" (built in, not stored in files).
 * Also used for other built-in styles, which are only known by id.
 */
const MEDIUM_STYLE_2 = (accent: string) => {
  const ln = (w: number) => `<a:ln w="${w}"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>`;
  const fill = (mod: string) => `<a:fill><a:solidFill><a:schemeClr val="${accent}">${mod}</a:schemeClr></a:solidFill></a:fill>`;
  const strong = `<a:tcTxStyle b="on"><a:fontRef idx="minor"/><a:schemeClr val="lt1"/></a:tcTxStyle>`;
  return `<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"/><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>
<a:left>${ln(12700)}</a:left><a:right>${ln(12700)}</a:right><a:top>${ln(12700)}</a:top><a:bottom>${ln(12700)}</a:bottom>
<a:insideH>${ln(12700)}</a:insideH><a:insideV>${ln(12700)}</a:insideV></a:tcBdr>${fill('<a:tint val="20000"/>')}</a:tcStyle></a:wholeTbl>
<a:band1H><a:tcStyle>${fill('<a:tint val="40000"/>')}</a:tcStyle></a:band1H>
<a:band1V><a:tcStyle>${fill('<a:tint val="40000"/>')}</a:tcStyle></a:band1V>
<a:lastCol>${strong}<a:tcStyle>${fill('')}</a:tcStyle></a:lastCol>
<a:firstCol>${strong}<a:tcStyle>${fill('')}</a:tcStyle></a:firstCol>
<a:lastRow>${strong}<a:tcStyle><a:tcBdr><a:top>${ln(38100)}</a:top></a:tcBdr>${fill('')}</a:tcStyle></a:lastRow>
<a:firstRow>${strong}<a:tcStyle><a:tcBdr><a:bottom>${ln(38100)}</a:bottom></a:tcBdr>${fill('')}</a:tcStyle></a:firstRow>
</a:tblStyle>`;
};

let builtIn: Element | null = null;

function tableStyle(tblPr: Element | null, ctx: Ctx): Element | null {
  const id = child(tblPr, 'tableStyleId')?.textContent?.trim() ?? '';
  const own = ctx.deck.tableStyles.get(id);
  if (own) return own;
  if (!id && !child(tblPr, 'tableStyle')) return null;
  builtIn ??= new DOMParser().parseFromString(MEDIUM_STYLE_2('accent1'), 'application/xml').documentElement;
  return builtIn;
}

interface CellLook {
  fill: Element | null;
  fillRef: Element | null;
  borders: Record<string, Element | null>;
  text: ParaStyle['run'];
}

async function partLook(part: Element | null, ctx: Ctx, look: CellLook, edges: { left: string; right: string; top: string; bottom: string }): Promise<void> {
  if (!part) return;
  const tcStyle = child(part, 'tcStyle');
  const fill = child(tcStyle, 'fill');
  if (fill) {
    look.fill = fillElement(fill);
    look.fillRef = null;
  }
  const fillRef = child(tcStyle, 'fillRef');
  if (fillRef) {
    look.fillRef = fillRef;
    look.fill = null;
  }
  const bdr = child(tcStyle, 'tcBdr');
  for (const side of ['left', 'right', 'top', 'bottom'] as const) {
    const edge = child(bdr, edges[side]);
    if (edge) look.borders[side] = child(edge, 'ln');
  }
  const tx = child(part, 'tcTxStyle');
  if (tx) {
    if (tx.hasAttribute('b')) look.text.bold = tx.getAttribute('b') === 'on';
    if (tx.hasAttribute('i')) look.text.italic = tx.getAttribute('i') === 'on';
    const c = colour(tx, ctx);
    if (c) look.text.colour = c;
    const font = child(tx, 'fontRef')?.getAttribute('idx');
    if (font) look.text.font = font === 'major' ? '+mj-lt' : '+mn-lt';
  }
}

export async function renderTable(
  tbl: Element, frame: { x: number; y: number; w: number; h: number }, ctx: Ctx,
): Promise<string> {
  const tblPr = child(tbl, 'tblPr');
  const flag = (name: string) => tblPr?.getAttribute(name) === '1';
  const style = tableStyle(tblPr, ctx);
  const cols = children(child(tbl, 'tblGrid'), 'gridCol').map((c) => Number(c.getAttribute('w') ?? 0) / EMU);
  const rows = children(tbl, 'tr');
  const nRows = rows.length;
  const nCols = cols.length;
  const base = levelStyles([ctx.deck.defaultTextStyle], ctx);

  let html = `<table class="tbl" style="left:${frame.x}px;top:${frame.y}px;width:${cols.reduce((a, b) => a + b, 0)}px"><colgroup>`;
  for (const w of cols) html += `<col style="width:${w}px">`;
  html += '</colgroup>';
  for (let r = 0; r < nRows; r++) {
    const tr = rows[r];
    const h = Number(tr.getAttribute('h') ?? 0) / EMU;
    html += `<tr style="height:${h}px">`;
    const cells = children(tr, 'tc');
    for (let c = 0; c < cells.length; c++) {
      const tc = cells[c];
      if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') continue;
      const colSpan = Number(tc.getAttribute('gridSpan') ?? 1);
      const rowSpan = Number(tc.getAttribute('rowSpan') ?? 1);
      const look: CellLook = { fill: null, fillRef: null, borders: {}, text: {} };
      if (style) {
        const edges = {
          left: c === 0 ? 'left' : 'insideV',
          right: c + colSpan >= nCols ? 'right' : 'insideV',
          top: r === 0 ? 'top' : 'insideH',
          bottom: r + rowSpan >= nRows ? 'bottom' : 'insideH',
        };
        const firstRow = flag('firstRow') && r === 0;
        const lastRow = flag('lastRow') && r === nRows - 1;
        await partLook(child(style, 'wholeTbl'), ctx, look, edges);
        if (flag('bandRow') && !firstRow && !lastRow) {
          const band = (r - (flag('firstRow') ? 1 : 0)) % 2 === 0 ? 'band1H' : 'band2H';
          await partLook(child(style, band), ctx, look, edges);
        }
        if (flag('bandCol')) {
          const band = (c - (flag('firstCol') ? 1 : 0)) % 2 === 0 ? 'band1V' : 'band2V';
          await partLook(child(style, band), ctx, look, edges);
        }
        if (flag('firstCol') && c === 0) await partLook(child(style, 'firstCol'), ctx, look, edges);
        if (flag('lastCol') && c + colSpan >= nCols) await partLook(child(style, 'lastCol'), ctx, look, edges);
        if (lastRow) await partLook(child(style, 'lastRow'), ctx, look, edges);
        if (firstRow) await partLook(child(style, 'firstRow'), ctx, look, edges);
      }

      const tcPr = child(tc, 'tcPr');
      let cellCss = '';
      const ownFill = fillElement(tcPr);
      if (ownFill) cellCss += fillCss(await readFill(ownFill, ctx));
      else if (look.fill) cellCss += fillCss(await readFill(look.fill, ctx));
      else if (look.fillRef) {
        const ref = styleRef(look.fillRef, ctx.theme.fillStyles, ctx.theme.bgFillStyles, ctx);
        cellCss += fillCss(await readFill(ref.el, ref.ctx));
      }
      for (const [side, own] of [['left', 'lnL'], ['right', 'lnR'], ['top', 'lnT'], ['bottom', 'lnB']] as const) {
        const border = await borderCss(child(tcPr, own) ?? look.borders[side] ?? null, ctx);
        if (border) cellCss += `;border-${side}:${border}`;
      }
      const ins = (name: string, d: number) => Number(tcPr?.getAttribute(name) ?? d) / EMU;
      const anchor = tcPr?.getAttribute('anchor') ?? 't';
      cellCss += `;vertical-align:${anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top'}`;

      let text = '';
      const txBody = child(tc, 'txBody');
      if (txBody) {
        const levels = overlay(base.map((l) => mergePara(l, { run: look.text })), child(txBody, 'lstStyle'), ctx);
        const body = bodyProps([child(txBody, 'bodyPr')]);
        body.insets = [ins('marT', 45720), ins('marR', 91440), ins('marB', 45720), ins('marL', 91440)];
        body.anchor = 't';
        const w = cols.slice(c, c + colSpan).reduce((a, b) => a + b, 0);
        text = renderText(txBody, levels, body, ctx, { w, h: 0 }).replace('class="tx"', 'class="tx cell"');
      }
      const span = `${colSpan > 1 ? ` colspan="${colSpan}"` : ''}${rowSpan > 1 ? ` rowspan="${rowSpan}"` : ''}`;
      html += `<td${span} style="${cellCss.replace(/^;/, '')}">${text}</td>`;
    }
    html += '</tr>';
  }
  return `${html}</table>`;
}
