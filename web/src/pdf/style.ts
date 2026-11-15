import { t } from '../i18n';
import { overlay } from './overlay';
import { css, DEFAULT_STYLE, ERASER_SIZE, PEN_WIDTH, type ColourTarget, type Rgb, type Style } from './edits';

// The settings panel of a PDF editing tool: a popover above the edit bar with that tool's
// options only (opened from the tool's own button).
// Pen: a preview stroke, thickness and opacity. Highlight / underline: opacity. Pen, highlight,
// underline and text: a colour, from a grid of swatches or from hue / saturation / brightness
// sliders, or picked from the page. Eraser: pixel (partial) or stroke (whole marks), and its size.
// Every change applies at once and is kept on the device.

export type StyleTarget = ColourTarget | 'eraser';

const STYLE_KEY = 'officeeditor.pdfStyle';
/** Colours only, from version 1.2.2; read once if there is no newer style. */
const OLD_COLOURS_KEY = 'officeeditor.pdfColours';

export function readStyle(): Style {
  const style = structuredClone(DEFAULT_STYLE);
  try {
    const raw = localStorage.getItem(STYLE_KEY);
    if (raw) {
      const s = JSON.parse(raw) as Partial<Style>;
      for (const key of Object.keys(style) as (keyof Style)[]) Object.assign(style[key], valid(s[key]));
    } else {
      const old = JSON.parse(localStorage.getItem(OLD_COLOURS_KEY) ?? '{}') as Record<string, unknown>;
      for (const key of ['highlight', 'underline', 'ink', 'text'] as const) {
        if (isRgb(old[key])) style[key].colour = old[key];
      }
    }
  } catch {
    // Defaults.
  }
  return style;
}

export function saveStyle(style: Style): void {
  try {
    localStorage.setItem(STYLE_KEY, JSON.stringify(style));
  } catch {
    // Not kept for next time; still used now.
  }
}

/** Only the fields that hold sensible values. */
function valid(part: unknown): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (!part || typeof part !== 'object') return out;
  const p = part as Record<string, unknown>;
  if (isRgb(p.colour)) out.colour = p.colour;
  if (typeof p.opacity === 'number' && p.opacity >= 0.05 && p.opacity <= 1) out.opacity = p.opacity;
  if (typeof p.width === 'number') out.width = clamp(p.width, PEN_WIDTH.min, PEN_WIDTH.max);
  if (typeof p.size === 'number') out.size = clamp(p.size, ERASER_SIZE.min, ERASER_SIZE.max);
  if (p.mode === 'whole' || p.mode === 'part') out.mode = p.mode;
  return out;
}

const isRgb = (c: unknown): c is Rgb =>
  Array.isArray(c) && c.length === 3 && c.every((v) => typeof v === 'number' && v >= 0 && v <= 1);
const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v));

// ------------------------------------------------------------------ colours

/** The grid: twelve hues from light to dark, then a grey column from white to black. */
const HUES = [0, 18, 32, 44, 56, 75, 110, 168, 200, 220, 265, 322];
const LIGHTNESS = [0.94, 0.84, 0.73, 0.6, 0.46, 0.32, 0.2];
const GREYS = [1, 0.92, 0.78, 0.6, 0.4, 0.22, 0];

function hsl(h: number, s: number, l: number): Rgb {
  const k = (n: number) => (n + h / 30) % 12;
  const a = s * Math.min(l, 1 - l);
  const f = (n: number) => l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1)));
  return [f(0), f(8), f(4)].map((v) => Math.round(v * 255) / 255) as Rgb;
}

/** HSV (h 0–360, s and v 0–1) → RGB, for the sliders. */
function hsvToRgb(h: number, s: number, v: number): Rgb {
  const f = (n: number) => {
    const k = (n + h / 60) % 6;
    return v - v * s * Math.max(0, Math.min(k, 4 - k, 1));
  };
  return [f(5), f(3), f(1)].map((x) => Math.round(x * 255) / 255) as Rgb;
}

function rgbToHsv([r, g, b]: Rgb): [number, number, number] {
  const max = Math.max(r, g, b);
  const d = max - Math.min(r, g, b);
  let h = 0;
  if (d) {
    if (max === r) h = ((g - b) / d) % 6;
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
  }
  return [(h * 60 + 360) % 360, max ? d / max : 0, max];
}

export const sameRgb = (a: Rgb, b: Rgb) => a.every((v, i) => Math.abs(v - b[i]) < 0.003);
const hex = (c: Rgb) => `#${c.map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('')}`.toUpperCase();

// ------------------------------------------------------------------ the panel

export interface PanelOptions {
  style: Style;
  target: StyleTarget;
  /** Called after every change (the style object has been changed in place). */
  changed: () => void;
  /** The eyedropper: the panel has closed; the editor lets the user tap the page. */
  pick: (target: ColourTarget) => void;
  closed: () => void;
}

/** Which colour tab was open last (Grid or Sliders). */
let colourTab: 'grid' | 'sliders' = 'grid';

export function openStylePanel(o: PanelOptions): () => void {
  const catcher = el('div', 'pdf-style-catcher');
  // The frame spans the room between the top of the screen and the edit bar; the panel sits at
  // its bottom and is at most as tall. (Not sized with vh: in the app's WebView 100vh can be far
  // smaller than the screen, which cut the panel down to a sliver.)
  const frame = el('div', 'pdf-style-frame');
  const panel = el('div', 'pdf-style-panel');
  frame.append(panel);
  panel.setAttribute('role', 'dialog');
  const target = o.target;

  const close = () => {
    catcher.remove();
    frame.remove();
    document.removeEventListener('keydown', onKey);
    o.closed();
  };
  const onKey = (e: KeyboardEvent) => e.key === 'Escape' && close();
  catcher.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    close();
  });
  document.addEventListener('keydown', onKey);

  const labels: Record<StyleTarget, string> = {
    highlight: t('highlight'), underline: t('underline'), ink: t('pen'), text: t('addText'), eraser: t('eraser'),
  };
  panel.setAttribute('aria-label', labels[target]);
  const title = el('div', 'pdf-style-title');
  title.textContent = labels[target];
  panel.append(title, ...(target === 'eraser' ? eraserSection(o) : markSection(o, target, close)));
  overlay().append(catcher, frame);
  return close;
}

function markSection(o: PanelOptions, target: ColourTarget, close: () => void): HTMLElement[] {
  const style = o.style;
  const out: HTMLElement[] = [];
  const updates: (() => void)[] = [];
  const update = () => {
    for (const u of updates) u();
    o.changed();
  };

  if (target === 'ink') {
    const preview = el('div', 'pdf-style-preview');
    out.push(preview);
    updates.push(() => {
      const s = style.ink;
      // The thickness as it will look on a page at 100 %.
      preview.innerHTML = `<svg viewBox="0 0 240 60" aria-hidden="true"><path d="M20 34 C 70 64, 100 6, 150 26 S 210 40, 220 24" fill="none" stroke="${css(s.colour)}" stroke-width="${s.width * 1.333}" stroke-linecap="round" opacity="${s.opacity}"/></svg>`;
    });
    out.push(slider({
      label: t('thickness'), min: PEN_WIDTH.min, max: PEN_WIDTH.max, step: 0.5,
      value: () => style.ink.width,
      text: (v) => String(v),
      set: (v) => { style.ink.width = v; update(); },
    }));
  }

  if (target !== 'text') {
    const s = style[target];
    const row = slider({
      label: t('opacity'), min: 5, max: 100, step: 1,
      value: () => Math.round(s.opacity * 100),
      text: (v) => `${v}%`,
      set: (v) => { s.opacity = v / 100; update(); },
    });
    row.classList.add('pdf-opacity');
    updates.push(() => row.style.setProperty('--colour', css(s.colour)));
    out.push(row);
  }

  // Colour: Grid | Sliders.
  const seg = segmented(
    [['grid', t('colourGrid')], ['sliders', t('colourSliders')]],
    () => colourTab,
    (v) => {
      colourTab = v;
      drawColours();
      seg.refresh();
    },
  );
  const colours = el('div', 'pdf-colours');
  const setColour = (c: Rgb) => {
    style[target].colour = c;
    update();
  };
  const drawColours = () => {
    updates.length = baseUpdates;
    colours.replaceChildren(colourTab === 'grid' ? grid(() => style[target].colour, setColour, updates) : hsvSliders(() => style[target].colour, setColour, updates));
    for (const u of updates.slice(baseUpdates)) u();
  };

  const footer = el('div', 'pdf-colour-footer');
  const dropper = el('button', 'pagebar-icon');
  dropper.setAttribute('aria-label', t('pickColour'));
  dropper.title = t('pickColour');
  dropper.innerHTML = `<svg viewBox="0 0 24 24" width="24" height="24" aria-hidden="true"><path fill="currentColor" d="${ICON_DROPPER}"/></svg>`;
  dropper.onclick = () => {
    close();
    o.pick(target);
  };
  const current = el('div', 'pdf-colour-current');
  const currentHex = el('span', 'pdf-colour-hex');
  updates.push(() => {
    current.style.background = css(style[target].colour);
    currentHex.textContent = hex(style[target].colour);
  });
  footer.append(dropper, currentHex, current);

  const baseUpdates = updates.length;
  out.push(seg.el, colours, footer);
  if (target === 'highlight' || target === 'underline') {
    const hint = el('p', 'pdf-dialog-hint');
    hint.textContent = t('markHint', { tool: target === 'highlight' ? t('highlight') : t('underline') });
    out.push(hint);
  }
  drawColours();
  for (const u of updates) u();
  return out;
}

function eraserSection(o: PanelOptions): HTMLElement[] {
  const e = o.style.eraser;
  const hint = el('p', 'pdf-dialog-hint');
  const seg = segmented(
    [['part', t('eraserPixel')], ['whole', t('eraserStroke')]],
    () => e.mode,
    (v) => {
      e.mode = v;
      seg.refresh();
      hint.textContent = e.mode === 'part' ? t('eraserPixelInfo') : t('eraserStrokeInfo');
      o.changed();
    },
  );
  hint.textContent = e.mode === 'part' ? t('eraserPixelInfo') : t('eraserStrokeInfo');
  const size = slider({
    label: t('eraserSize'), min: ERASER_SIZE.min, max: ERASER_SIZE.max, step: 1,
    value: () => e.size,
    text: (v) => String(v),
    set: (v) => {
      e.size = v;
      dot.style.setProperty('--size', `${v * 2}px`);
      o.changed();
    },
  });
  // The eraser at its size on screen.
  const preview = el('div', 'pdf-style-preview pdf-eraser-preview');
  const dot = el('div', 'pdf-eraser-dot');
  dot.style.setProperty('--size', `${e.size * 2}px`);
  preview.append(dot);
  return [seg.el, size, preview, hint];
}

function grid(get: () => Rgb, set: (c: Rgb) => void, updates: (() => void)[]): HTMLElement {
  const g = el('div', 'pdf-colour-grid');
  g.setAttribute('role', 'radiogroup');
  const cells: [HTMLButtonElement, Rgb][] = [];
  for (let row = 0; row < LIGHTNESS.length; row++) {
    const colours = [...HUES.map((h) => hsl(h, row === 0 ? 1 : 0.88, LIGHTNESS[row])), [GREYS[row], GREYS[row], GREYS[row]] as Rgb];
    for (const c of colours) {
      const b = el('button', 'pdf-colour-cell');
      b.setAttribute('role', 'radio');
      b.setAttribute('aria-label', hex(c));
      b.style.background = css(c);
      b.onclick = () => set(c);
      g.append(b);
      cells.push([b, c]);
    }
  }
  updates.push(() => {
    const now = get();
    for (const [b, c] of cells) {
      const on = sameRgb(c, now);
      b.setAttribute('aria-checked', String(on));
      // The ring is drawn in a colour that shows on this cell.
      if (on) b.style.setProperty('--ring', rgbToHsv(c)[2] > 0.6 && rgbToHsv(c)[1] < 0.5 ? '#000' : '#fff');
    }
  });
  return g;
}

function hsvSliders(get: () => Rgb, set: (c: Rgb) => void, updates: (() => void)[]): HTMLElement {
  const box = el('div', 'pdf-hsv');
  // Kept here rather than read back from the colour, so hue survives grey and black.
  let [h, s, v] = rgbToHsv(get());
  const apply = () => set(hsvToRgb(h, s, v));
  const hue = slider({ label: t('hue'), min: 0, max: 360, step: 1, value: () => Math.round(h), text: (x) => `${x}°`, set: (x) => { h = x; apply(); } });
  const sat = slider({ label: t('saturation'), min: 0, max: 100, step: 1, value: () => Math.round(s * 100), text: (x) => `${x}%`, set: (x) => { s = x / 100; apply(); } });
  const val = slider({ label: t('brightness'), min: 0, max: 100, step: 1, value: () => Math.round(v * 100), text: (x) => `${x}%`, set: (x) => { v = x / 100; apply(); } });
  hue.classList.add('pdf-hue');
  updates.push(() => {
    sat.style.setProperty('--from', css(hsvToRgb(h, 0, v)));
    sat.style.setProperty('--to', css(hsvToRgb(h, 1, v)));
    val.style.setProperty('--to', css(hsvToRgb(h, s, 1)));
  });
  sat.classList.add('pdf-gradient');
  val.classList.add('pdf-gradient');
  val.style.setProperty('--from', '#000');
  box.append(hue, sat, val);
  return box;
}

interface SliderOptions {
  label: string;
  min: number;
  max: number;
  step: number;
  value: () => number;
  text: (v: number) => string;
  set: (v: number) => void;
}

function slider(o: SliderOptions): HTMLElement {
  const row = el('label', 'pdf-slider');
  const head = el('span', 'pdf-slider-head');
  const name = el('span', '');
  name.textContent = o.label;
  const value = el('span', 'pdf-slider-value');
  value.textContent = o.text(o.value());
  head.append(name, value);
  const input = el('input', '');
  input.type = 'range';
  input.min = String(o.min);
  input.max = String(o.max);
  input.step = String(o.step);
  input.value = String(o.value());
  input.oninput = () => {
    const v = Number(input.value);
    value.textContent = o.text(v);
    o.set(v);
  };
  row.append(head, input);
  return row;
}

function segmented<T extends string>(items: [T, string][], get: () => T, set: (v: T) => void): { el: HTMLElement; refresh: () => void } {
  const box = el('div', 'pdf-segmented');
  box.setAttribute('role', 'tablist');
  const buttons = items.map(([value, label]) => {
    const b = el('button', '');
    b.setAttribute('role', 'tab');
    b.textContent = label;
    b.onclick = () => set(value);
    box.append(b);
    return [value, b] as const;
  });
  const refresh = () => {
    for (const [value, b] of buttons) b.setAttribute('aria-selected', String(value === get()));
  };
  refresh();
  return { el: box, refresh };
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  return node;
}

// Material Symbols "colorize" (Apache-2.0).
const ICON_DROPPER = 'M20.71 5.63l-2.34-2.34a.996.996 0 0 0-1.41 0l-3.12 3.12-1.93-1.91-1.41 1.41 1.42 1.42L3 16.25V21h4.75l8.92-8.92 1.42 1.42 1.41-1.41-1.92-1.92 3.12-3.12c.4-.4.4-1.03.01-1.42zM6.92 19 5 17.08l8.06-8.06 1.92 1.92L6.92 19z';
