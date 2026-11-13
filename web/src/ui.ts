import './ui.css';
import { t } from './i18n';

// Small dialogs, menus and a toast for the editors. (The PDF and Word editors still have their
// own copies of these; they can move here.)

let toastEl: HTMLElement | null = null;
let toastTimer = 0;

export function toast(text: string): void {
  toastEl ??= document.body.appendChild(Object.assign(document.createElement('div'), { className: 'ui-toast' }));
  toastEl.textContent = text;
  toastEl.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => toastEl?.classList.remove('show'), 3500);
}

function box(title: string, body?: string): { panel: HTMLElement; actions: HTMLElement; close: () => void } {
  const scrim = document.createElement('div');
  scrim.className = 'ui-scrim';
  const panel = document.createElement('div');
  panel.className = 'ui-dialog';
  panel.setAttribute('role', 'dialog');
  const h = document.createElement('h2');
  h.textContent = title;
  panel.append(h);
  if (body) {
    const p = document.createElement('p');
    p.textContent = body;
    panel.append(p);
  }
  const actions = document.createElement('div');
  actions.className = 'ui-actions';
  panel.append(actions);
  scrim.append(panel);
  document.body.append(scrim);
  return { panel, actions, close: () => scrim.remove() };
}

function button(label: string, onClick: () => void, primary = false): HTMLButtonElement {
  const b = document.createElement('button');
  b.className = primary ? 'ui-button primary' : 'ui-button';
  b.textContent = label;
  b.onclick = onClick;
  return b;
}

/** A list of choices; resolves with the chosen value, or null if cancelled. */
export function menu<T extends string>(title: string, items: [T, string][]): Promise<T | null> {
  return new Promise((resolve) => {
    const d = box(title);
    const list = document.createElement('div');
    list.className = 'ui-menu';
    for (const [value, label] of items) list.append(button(label, () => { d.close(); resolve(value); }));
    d.panel.insertBefore(list, d.actions);
    d.actions.append(button(t('cancel'), () => { d.close(); resolve(null); }));
  });
}

export function confirm(title: string, body: string, ok: string): Promise<boolean> {
  return new Promise((resolve) => {
    const d = box(title, body);
    d.actions.append(
      button(t('cancel'), () => { d.close(); resolve(false); }),
      button(ok, () => { d.close(); resolve(true); }, true),
    );
  });
}

/** Font sizes the editors accept, in points (Excel's limit is 409, Word's 1638). */
export const FONT_SIZE_MIN = 1;
export const FONT_SIZE_MAX = 400;

/** A typed font size in points (a decimal comma is fine), to the half point; null if not one. */
export function parseFontSize(text: string): number | null {
  const n = Number(text.trim().replace(',', '.'));
  if (!text.trim() || !Number.isFinite(n) || n < FONT_SIZE_MIN || n > FONT_SIZE_MAX) return null;
  return Math.round(n * 2) / 2;
}

/** A text field for a font size in points, with a "pt" label and the allowed range below it. */
export function fontSizeField(value?: number): { field: HTMLElement; input: HTMLInputElement; hint: HTMLElement } {
  const field = document.createElement('label');
  field.className = 'ui-field';
  const input = document.createElement('input');
  input.className = 'ui-number';
  input.type = 'text';
  input.inputMode = 'decimal';
  input.enterKeyHint = 'done';
  input.autocomplete = 'off';
  input.setAttribute('aria-label', t('fontSize'));
  if (value) input.value = String(Math.round(value * 2) / 2);
  const unit = document.createElement('span');
  unit.textContent = 'pt';
  field.append(input, unit);
  const hint = document.createElement('p');
  hint.className = 'ui-hint';
  hint.textContent = t('fontSizeRange', { min: FONT_SIZE_MIN, max: FONT_SIZE_MAX });
  input.addEventListener('input', () => hint.classList.remove('error'));
  return { field, input, hint };
}

/** Asks for a font size in points, typed in; `current` (if known) fills the field. */
export function askFontSize(current?: number): Promise<number | null> {
  return new Promise((resolve) => {
    const d = box(t('fontSize'));
    const { field, input, hint } = fontSizeField(current);
    d.panel.insertBefore(field, d.actions);
    d.panel.insertBefore(hint, d.actions);
    const cancel = () => {
      d.close();
      resolve(null);
    };
    const ok = () => {
      const n = parseFontSize(input.value);
      if (n === null) {
        hint.classList.add('error');
        input.focus();
        return;
      }
      d.close();
      resolve(n);
    };
    input.onkeydown = (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        ok();
      } else if (e.key === 'Escape') cancel();
    };
    d.actions.append(button(t('cancel'), cancel), button(t('done'), ok, true));
    // Ready to type over: the current size is selected.
    input.focus();
    input.select();
  });
}

/** Material Symbols "format_size" (Apache-2.0), for a toolbar button. */
export const ICON_FONT_SIZE = '<svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><path fill="currentColor" d="M9 4v3h5v12h3V7h5V4H9zm-6 8h3v7h3v-7h3V9H3v3z"/></svg>';
