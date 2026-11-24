// Reading without bars: a single tap on the document slides the viewer's bottom bar down and the
// shell's top bar up (and hides the status bar); another tap brings them back. Used by the PDF,
// Word and PowerPoint viewers while they are not editing.
//
// In those viewers the shell's top bar floats over the page instead of sitting above it, so
// hiding it never resizes the page (a resize made the document jump). The page keeps room for it
// at the top: `--oe-top`, in CSS px, set from the shell (`getTopInset`, `officeSetTopInset`).

let hidden = false;
let top = 0;

function setTopInset(px: number): void {
  top = px;
  document.documentElement.style.setProperty('--oe-top', `${px}px`);
}

setTopInset(window.OfficeBridge?.getTopInset?.() ?? 0);
window.officeSetTopInset = setTopInset;

/** Height of the shell's top bar over the page, in CSS px (0 if none). */
export function topInset(): number {
  return top;
}

export function barsHidden(): boolean {
  return hidden;
}

export function setBarsHidden(on: boolean): void {
  if (on === hidden) return;
  hidden = on;
  document.documentElement.classList.toggle('bars-hidden', on);
  window.OfficeBridge?.setBarsHidden?.(on);
}

/** A second tap within this many ms makes a double tap (zoom), which leaves the bars alone. */
const DOUBLE_TAP_MS = 300;

/**
 * A single tap on `target` toggles the bars, unless `enabled()` says no (edit mode, presenting),
 * the tap was on a control or link, or it only cleared a text selection.
 */
export function toggleBarsOnTap(target: HTMLElement, enabled: () => boolean): void {
  let timer = 0;
  let hadSelection = false;
  target.addEventListener('pointerdown', () => {
    const selection = window.getSelection();
    hadSelection = !!selection && !selection.isCollapsed;
  }, { passive: true });
  target.addEventListener('click', (e) => {
    if (timer) {
      clearTimeout(timer);
      timer = 0;
      return;
    }
    if (!enabled() || hadSelection) return;
    if ((e.target as Element).closest?.('a, button, input, select, textarea, label, [contenteditable="true"]')) return;
    timer = window.setTimeout(() => {
      timer = 0;
      if (enabled()) setBarsHidden(!hidden);
    }, DOUBLE_TAP_MS);
  });
}

declare global {
  interface Window {
    /** Called by the shell when its floating top bar's height changes (CSS px). */
    officeSetTopInset?: (px: number) => void;
  }
}
