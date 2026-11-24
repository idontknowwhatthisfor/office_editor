// The editors' bottom toolbars (the PDF edit row, the Word edit bar), drawn by the app instead.
//
// A bar pinned to the visible area while the page is pinch-zoomed can only follow the view a
// frame late (see pagebar.ts), so in the app the shell draws the toolbar natively
// (`MirroredToolbar.kt`) from a description of its buttons, and replays taps and holds on the
// real buttons here (`window.officeToolbar`). The editors keep building, updating and listening
// to their own buttons as before; the row stays in the page, invisible, so whatever is placed
// relative to it (panels, toasts) stays where it was.
//
// A button is described by its label, its active/disabled state, and either its icon's paths
// (each with the transform of the groups around it, in the icon's 24-unit box) or its text.

interface PathDesc {
  d: string;
  /** Affine transform [a, b, c, d, e, f], when a group moves or scales the path. */
  m?: number[];
}

interface ButtonDesc {
  label: string;
  active: boolean;
  disabled: boolean;
  /** Icon size in CSS px (the box the 24-unit viewBox is drawn in). */
  size: number;
  paths: PathDesc[];
  text: string;
  bold: boolean;
  italic: boolean;
}

const rows: HTMLElement[] = [];
/** The row the shell is drawing now, and its buttons in the order sent. */
let shown: { row: HTMLElement; buttons: HTMLButtonElement[] } | null = null;
let lastSent: string | null = null;
let frame = 0;

const native = () => !!window.OfficeBridge?.setToolbar;

/** Lets the shell draw `row` (a toolbar of buttons) while it is visible. No-op in a browser. */
export function mirrorToolbar(row: HTMLElement): void {
  if (!native() || rows.includes(row)) return;
  rows.push(row);
  row.classList.add('oe-mirrored');
  new MutationObserver(schedule).observe(row, {
    subtree: true, childList: true, attributes: true, characterData: true,
  });
  schedule();
}

/** Changes come in bursts (a tool switch touches several buttons): send once after them. */
function schedule(): void {
  if (!frame) frame = window.setTimeout(update, 16);
}

function visible(row: HTMLElement): boolean {
  return row.isConnected && !row.closest('[hidden]') && getComputedStyle(row).display !== 'none';
}

function update(): void {
  frame = 0;
  const row = rows.find(visible) ?? null;
  const buttons = row
    ? [...row.querySelectorAll<HTMLButtonElement>(':scope > button')].filter((b) => !b.hidden && getComputedStyle(b).display !== 'none')
    : [];
  shown = row ? { row, buttons } : null;
  const json = row ? JSON.stringify({ height: row.offsetHeight, buttons: buttons.map(describe) }) : '';
  if (json === lastSent) return;
  lastSent = json;
  window.OfficeBridge?.setToolbar?.(json);
}

function describe(b: HTMLButtonElement): ButtonDesc {
  const svg = b.querySelector('svg');
  const style = getComputedStyle(b);
  return {
    label: b.getAttribute('aria-label') ?? b.title ?? '',
    active: b.classList.contains('active'),
    disabled: b.disabled,
    size: svg ? Number(svg.getAttribute('width')) || 24 : 24,
    paths: svg ? [...svg.querySelectorAll('path')].map((p) => describePath(p, svg)) : [],
    text: svg ? '' : (b.textContent ?? '').trim(),
    bold: !!b.querySelector('b, strong') || Number(style.fontWeight) >= 600,
    italic: !!b.querySelector('i, em') || style.fontStyle === 'italic',
  };
}

function describePath(p: SVGPathElement, svg: SVGSVGElement): PathDesc {
  let m = new DOMMatrix();
  for (let e: Element | null = p; e && e !== svg; e = e.parentElement) {
    const list = (e as SVGGraphicsElement).transform?.baseVal;
    const t = list?.numberOfItems ? list.consolidate()?.matrix : null;
    if (t) m = new DOMMatrix([t.a, t.b, t.c, t.d, t.e, t.f]).multiply(m);
  }
  const d = p.getAttribute('d') ?? '';
  return m.isIdentity ? { d } : { d, m: [m.a, m.b, m.c, m.d, m.e, m.f] };
}

/**
 * The shell's toolbar was used: replay it on the real button. A tap is a press, a release and a
 * click (hold-aware buttons time the press); a hold is the long-press `contextmenu`, which
 * buttons without a hold action ignore, so then it counts as a tap (as a long press there did).
 */
window.officeToolbar = (index, kind) => {
  const b = shown?.buttons[index];
  if (!b || b.disabled || !b.isConnected) return;
  const pointer = { bubbles: true, cancelable: true, pointerType: 'touch', isPrimary: true };
  const tap = () => {
    b.dispatchEvent(new PointerEvent('pointerdown', pointer));
    b.dispatchEvent(new PointerEvent('pointerup', pointer));
    b.click();
  };
  if (kind === 'hold') {
    const handled = !b.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    if (!handled) tap();
  } else {
    tap();
  }
};

// A page (re)loaded: whatever toolbar the shell drew for the previous one is gone.
if (native()) window.OfficeBridge?.setToolbar?.('');

declare global {
  interface Window {
    /** The shell's toolbar button `index` was tapped or held. */
    officeToolbar?: (index: number, kind: 'tap' | 'hold') => void;
  }
}
