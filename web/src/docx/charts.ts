import { renderChart, type ChartTheme } from '../chart';
import { Package, all, child, first, relAttr, relOfType, type Rel } from '../ooxml';
import { DEFAULT_THEME, readTheme } from '../pptx/deck';

// Charts in a Word document. docx-preview reads a chart's drawing as an empty box of the right
// size (and, for the newer "chartex" charts, shows the picture Word saved as a fallback). Here
// the parsed drawings are matched with the document's XML, in order, and the charts among them
// marked; once the document is rendered, the chart is drawn into its box as SVG.

const EMU = 9525;

interface DocChart {
  root: Element;
  /** CSS px */
  w: number;
  h: number;
}

/**
 * Marks the charts in docx-preview's parsed `doc` (no public types) and returns a function that
 * draws them into the rendered view. Does nothing if the drawings don't line up one-to-one.
 */
export async function prepareCharts(doc: any, data: ArrayBuffer): Promise<(view: HTMLElement) => void> {
  const none = () => {};
  const pkg = await Package.open(data);
  const docPath = relOfType(await pkg.rels('_rels/.rels'), '/officeDocument')?.target ?? 'word/document.xml';
  const xml = await pkg.xml(docPath);
  const body = first(xml, 'body');
  if (!body || !all(body, 'chart').length) return none;
  const rels = await pkg.rels(docPath);

  // The drawings as docx-preview meets them: it always takes an mc:Fallback over its mc:Choice,
  // and never looks inside text boxes (drawings within drawings).
  const drawings = all(body, 'drawing').filter((d) => !within(d, 'Choice') && !within(d, 'drawing'));
  const parsed: any[] = [];
  const visit = (el: any): void => {
    if (el.type === 'drawing') parsed.push(el);
    for (const c of el.children ?? []) visit(c);
  };
  if (doc.documentPart?.body) visit(doc.documentPart.body);
  if (parsed.length !== drawings.length) return none;

  const themePath = relOfType(rels, '/theme')?.target;
  const themeDoc = themePath ? await pkg.xml(themePath) : null;
  const theme: ChartTheme = {
    theme: themeDoc ? readTheme(themeDoc) : DEFAULT_THEME,
    clrMap: await colourMap(pkg, rels),
    baseSize: 10,
    frame: true,
  };

  const charts: DocChart[] = [];
  for (let i = 0; i < drawings.length; i++) {
    const d = drawings[i];
    // A chartex chart's drawing is in the mc:Choice next to the fallback docx-preview reads.
    const fallback = within(d, 'Fallback');
    const choice = fallback ? child(fallback.parentElement, 'Choice') : null;
    const ref = first(d, 'chart') ?? first(choice, 'chart');
    const rel = ref ? rels.get(relAttr(ref)) : undefined;
    if (!rel || rel.external) continue;
    const part = await pkg.xml(rel.target);
    const extent = first(d, 'extent');
    const w = Number(extent?.getAttribute('cx') ?? 0) / EMU;
    const h = Number(extent?.getAttribute('cy') ?? 0) / EMU;
    if (!part || !w || !h) continue;
    parsed[i].className = 'docx-chart';
    parsed[i].children = [];
    charts.push({ root: part.documentElement, w, h });
  }
  if (!charts.length) return none;

  return (view) => {
    view.querySelectorAll<HTMLElement>('.docx-chart').forEach((box, i) => {
      const chart = charts[i];
      const svg = chart && renderChart(chart.root, theme, chart.w, chart.h);
      if (svg) box.innerHTML = svg;
    });
  };
}

/** The nearest ancestor of `el` with this local name. */
function within(el: Element, name: string): Element | null {
  for (let p = el.parentElement; p; p = p.parentElement) if (p.localName === name) return p;
  return null;
}

/** Word's mapping of bg1/tx1… to theme colours (settings.xml), or the usual one. */
async function colourMap(pkg: Package, rels: Map<string, Rel>): Promise<Record<string, string>> {
  const map: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
  const path = relOfType(rels, '/settings')?.target;
  const settings = path ? await pkg.xml(path) : null;
  const mapping = first(settings, 'clrSchemeMapping');
  for (const attr of mapping?.attributes ?? []) {
    const v = attr.value;
    const name = attr.localName;
    const key = name === 't1' ? 'tx1' : name === 't2' ? 'tx2' : name;
    map[key] = v === 'light1' ? 'lt1' : v === 'light2' ? 'lt2' : v === 'dark1' ? 'dk1' : v === 'dark2' ? 'dk2' : v;
  }
  return map;
}
