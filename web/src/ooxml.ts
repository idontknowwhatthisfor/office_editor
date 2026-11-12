import JSZip from 'jszip';

// Helpers shared by the XLSX and PPTX readers: an Office file is a zip of XML parts linked by
// relationship files. Elements are matched by local name so any namespace prefix works.

export interface Rel {
  target: string;
  type: string;
  external: boolean;
}

export class Package {
  private constructor(private zip: JSZip) {}

  static async open(data: ArrayBuffer): Promise<Package> {
    return new Package(await JSZip.loadAsync(data));
  }

  has(path: string): boolean {
    return !!this.zip.file(path);
  }

  async xml(path: string): Promise<Document | null> {
    const file = this.zip.file(path);
    if (!file) return null;
    const doc = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error(`bad XML in ${path}`);
    return doc;
  }

  async bytes(path: string): Promise<Uint8Array | null> {
    const file = this.zip.file(path);
    return file ? file.async('uint8array') : null;
  }

  /** Relationship id → target, for the part at `path`. */
  async rels(path: string): Promise<Map<string, Rel>> {
    const slash = path.lastIndexOf('/');
    const dir = path.slice(0, slash + 1);
    const doc = await this.xml(`${dir}_rels/${path.slice(slash + 1)}.rels`);
    const out = new Map<string, Rel>();
    for (const rel of all(doc, 'Relationship')) {
      const external = rel.getAttribute('TargetMode') === 'External';
      const target = rel.getAttribute('Target') ?? '';
      out.set(rel.getAttribute('Id') ?? '', {
        target: external ? target : resolvePath(dir, target),
        type: rel.getAttribute('Type') ?? '',
        external,
      });
    }
    return out;
  }
}

/** First relationship of a type (matched by the end of its type URI, e.g. '/theme'). */
export function relOfType(rels: Map<string, Rel>, suffix: string): Rel | undefined {
  for (const rel of rels.values()) if (rel.type.endsWith(suffix)) return rel;
  return undefined;
}

export function all(root: Document | Element | null | undefined, name: string): Element[] {
  return root ? [...root.getElementsByTagNameNS('*', name)] : [];
}

export function first(root: Document | Element | null | undefined, name: string): Element | null {
  return root ? root.getElementsByTagNameNS('*', name)[0] ?? null : null;
}

export function child(el: Element | null | undefined, name: string): Element | null {
  if (!el || !name) return null;
  for (const c of el.children) if (c.localName === name) return c;
  return null;
}

export function children(el: Element | null | undefined, name?: string): Element[] {
  return el ? [...el.children].filter((c) => !name || c.localName === name) : [];
}

/** An attribute in the relationships namespace (`r:id`, `r:embed`…), whatever the prefix. */
export function relAttr(el: Element | null | undefined, name = 'id'): string {
  if (!el) return '';
  for (const attr of el.attributes) {
    if (attr.localName === name && attr.namespaceURI?.includes('relationships')) return attr.value;
  }
  return '';
}

export function resolvePath(dir: string, target: string): string {
  const parts = (target.startsWith('/') ? target.slice(1) : dir + target).split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}
