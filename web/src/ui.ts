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
export function menu<T extends string>(title: string, items: [T, string][], className = ''): Promise<T | null> {
  return new Promise((resolve) => {
    const d = box(title);
    const list = document.createElement('div');
    list.className = `ui-menu ${className}`.trim();
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

/** Font sizes offered by the editors, in points (as Office's size list). */
export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72];

/** Asks for a font size in points; `current` (if known) is shown in the title. */
export async function fontSizeMenu(current?: number): Promise<number | null> {
  const now = current ? Math.round(current * 2) / 2 : undefined;
  const title = now ? t('fontSizeNow', { n: now }) : t('fontSize');
  const choice = await menu(title, FONT_SIZES.map((n) => [String(n), n === now ? `✓ ${n}` : String(n)] as [string, string]), 'ui-sizes');
  return choice ? Number(choice) : null;
}

/** Material Symbols "format_size" (Apache-2.0), for a toolbar button. */
export const ICON_FONT_SIZE = '<svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><path fill="currentColor" d="M9 4v3h5v12h3V7h5V4H9zm-6 8h3v7h3v-7h3V9H3v3z"/></svg>';
