import { Package, all, child, children, first, relAttr, relOfType, type Rel } from '../ooxml';

// The parts of a presentation a slide needs to be drawn: the slide, its layout, the layout's
// master and the master's theme. Shapes inherit position, text styles and fills down that chain.

/** EMU (English Metric Units) per CSS pixel. */
export const EMU = 9525;

export interface Theme {
  /** Scheme colours by name (dk1, lt1, dk2, lt2, accent1–6, hlink, folHlink), as hex. */
  colours: Record<string, string>;
  majorFont: string;
  minorFont: string;
  fillStyles: Element[];
  lineStyles: Element[];
  effectStyles: Element[];
  bgFillStyles: Element[];
}

export interface Part {
  path: string;
  root: Element;
  rels: Map<string, Rel>;
}

export interface Master extends Part {
  theme: Theme;
  /** bg1/tx1/bg2/tx2 → theme colour names. */
  clrMap: Record<string, string>;
  titleStyle: Element | null;
  bodyStyle: Element | null;
  otherStyle: Element | null;
}

export interface Layout extends Part {
  master: Master;
}

export interface SlidePart extends Part {
  layout: Layout;
  /** Colour map for this slide (the master's, unless the layout or slide overrides it). */
  clrMap: Record<string, string>;
}

export const DEFAULT_THEME: Theme = {
  colours: {
    dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31',
    accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72',
  },
  majorFont: 'Calibri Light',
  minorFont: 'Calibri',
  fillStyles: [],
  lineStyles: [],
  effectStyles: [],
  bgFillStyles: [],
};

export class Deck {
  /** Slide size in CSS px. */
  width = 960;
  height = 540;
  slidePaths: string[] = [];
  defaultTextStyle: Element | null = null;
  /** Table styles defined in the file, by id. */
  tableStyles = new Map<string, Element>();
  private parts = new Map<string, Promise<Part | Layout | Master>>();
  private images = new Map<string, Promise<string | null>>();
  /**
   * Edit mode: each slide's XML is parsed once and kept, so the editor and the renderer share
   * one working copy.
   */
  keepSlides = false;
  private slideDocs = new Map<string, Promise<Document | null>>();

  private constructor(readonly pkg: Package) {}

  static async open(data: ArrayBuffer): Promise<Deck> {
    const deck = new Deck(await Package.open(data));
    await deck.read();
    return deck;
  }

  private async read(): Promise<void> {
    const root = await this.pkg.rels('_rels/.rels');
    const path = relOfType(root, '/officeDocument')?.target ?? 'ppt/presentation.xml';
    const doc = await this.pkg.xml(path);
    if (!doc) throw new Error('no presentation part');
    const rels = await this.pkg.rels(path);
    const size = first(doc, 'sldSz');
    if (size) {
      this.width = Number(size.getAttribute('cx')) / EMU || this.width;
      this.height = Number(size.getAttribute('cy')) / EMU || this.height;
    }
    for (const id of all(first(doc, 'sldIdLst'), 'sldId')) {
      const rel = rels.get(relAttr(id));
      if (rel && !rel.external && this.pkg.has(rel.target)) this.slidePaths.push(rel.target);
    }
    this.defaultTextStyle = first(doc, 'defaultTextStyle');
    const stylesPath = relOfType(rels, '/tableStyles')?.target;
    const styles = stylesPath ? await this.pkg.xml(stylesPath) : null;
    for (const style of all(styles, 'tblStyle')) this.tableStyles.set(style.getAttribute('styleId') ?? '', style);
  }

  /** A slide's XML (in edit mode, the working copy). */
  slideDoc(index: number): Promise<Document | null> {
    const path = this.slidePaths[index];
    let doc = this.slideDocs.get(path);
    if (!doc) {
      doc = this.pkg.xml(path);
      if (this.keepSlides) this.slideDocs.set(path, doc);
    }
    return doc;
  }

  /** Replaces a slide's working XML (undo). */
  setSlideDoc(index: number, doc: Document): void {
    this.slideDocs.set(this.slidePaths[index], Promise.resolve(doc));
  }

  async slide(index: number): Promise<SlidePart> {
    const path = this.slidePaths[index];
    const doc = await this.slideDoc(index);
    if (!doc) throw new Error(`missing ${path}`);
    const rels = await this.pkg.rels(path);
    const layoutPath = relOfType(rels, '/slideLayout')?.target;
    const layout = layoutPath ? await this.layout(layoutPath) : await this.fallbackLayout();
    const clrMap = overrideMap(doc.documentElement, overrideMap(layout.root, layout.master.clrMap));
    return { path, root: doc.documentElement, rels, layout, clrMap };
  }

  private layout(path: string): Promise<Layout> {
    return this.cached(path, async () => {
      const doc = await this.pkg.xml(path);
      if (!doc) throw new Error(`missing ${path}`);
      const rels = await this.pkg.rels(path);
      const masterPath = relOfType(rels, '/slideMaster')?.target;
      if (!masterPath) throw new Error(`${path} has no master`);
      return { path, root: doc.documentElement, rels, master: await this.master(masterPath) };
    }) as Promise<Layout>;
  }

  /** A slide without a layout (broken file): use an empty one on the first master. */
  private async fallbackLayout(): Promise<Layout> {
    const rels = await this.pkg.rels('ppt/presentation.xml');
    const masterPath = relOfType(rels, '/slideMaster')?.target;
    if (!masterPath) throw new Error('no slide master');
    const empty = new DOMParser().parseFromString('<sldLayout/>', 'application/xml').documentElement;
    return { path: '', root: empty, rels: new Map(), master: await this.master(masterPath) };
  }

  private master(path: string): Promise<Master> {
    return this.cached(path, async () => {
      const doc = await this.pkg.xml(path);
      if (!doc) throw new Error(`missing ${path}`);
      const rels = await this.pkg.rels(path);
      const themePath = relOfType(rels, '/theme')?.target;
      const theme = readTheme(themePath ? await this.pkg.xml(themePath) : null);
      const map: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
      const clrMap = child(doc.documentElement, 'clrMap');
      for (const attr of clrMap?.attributes ?? []) map[attr.localName] = attr.value;
      const txStyles = child(doc.documentElement, 'txStyles');
      return {
        path,
        root: doc.documentElement,
        rels,
        theme,
        clrMap: map,
        titleStyle: child(txStyles, 'titleStyle'),
        bodyStyle: child(txStyles, 'bodyStyle'),
        otherStyle: child(txStyles, 'otherStyle'),
      };
    }) as Promise<Master>;
  }

  private cached(path: string, load: () => Promise<Part>): Promise<Part> {
    let part = this.parts.get(path);
    if (!part) this.parts.set(path, (part = load()));
    return part;
  }

  /** Any XML part with its relationships (diagram drawings, chart parts…). */
  async part(path: string): Promise<Part | null> {
    const doc = await this.pkg.xml(path);
    return doc ? { path, root: doc.documentElement, rels: await this.pkg.rels(path) } : null;
  }

  /** A URL for a picture in the package, or null if it isn't a format the browser can show. */
  image(path: string): Promise<string | null> {
    let url = this.images.get(path);
    if (!url) {
      url = (async () => {
        const type = IMAGE_TYPES[path.split('.').pop()!.toLowerCase()];
        if (!type) return null;
        const bytes = await this.pkg.bytes(path);
        return bytes ? URL.createObjectURL(new Blob([bytes as BlobPart], { type })) : null;
      })();
      this.images.set(path, url);
    }
    return url;
  }
}

const IMAGE_TYPES: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml', ico: 'image/x-icon',
};

/** A layout or slide's `<p:clrMapOvr>`, applied over the inherited map. */
function overrideMap(root: Element, base: Record<string, string>): Record<string, string> {
  const over = first(child(root, 'clrMapOvr'), 'overrideClrMapping');
  if (!over) return base;
  const map = { ...base };
  for (const attr of over.attributes) map[attr.localName] = attr.value;
  return map;
}

export function readTheme(doc: Document | null): Theme {
  if (!doc) return DEFAULT_THEME;
  const colours = { ...DEFAULT_THEME.colours };
  for (const el of children(first(doc, 'clrScheme'))) {
    const c = el.firstElementChild;
    const hex = c?.localName === 'sysClr' ? c.getAttribute('lastClr') : c?.getAttribute('val');
    if (hex) colours[el.localName] = hex;
  }
  const font = (name: string, fallback: string) =>
    child(first(doc, name), 'latin')?.getAttribute('typeface') || fallback;
  return {
    colours,
    majorFont: font('majorFont', DEFAULT_THEME.majorFont),
    minorFont: font('minorFont', DEFAULT_THEME.minorFont),
    fillStyles: children(first(doc, 'fillStyleLst')),
    lineStyles: children(first(doc, 'lnStyleLst')),
    effectStyles: children(first(doc, 'effectStyleLst')),
    bgFillStyles: children(first(doc, 'bgFillStyleLst')),
  };
}
