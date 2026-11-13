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
