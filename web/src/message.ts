/** Replaces the page with a centred message (loading, errors, not-yet-supported). */
export function showMessage(root: HTMLElement, title: string, body?: string): void {
  const box = document.createElement('div');
  box.className = 'message';
  const h = document.createElement('p');
  h.className = 'message-title';
  h.textContent = title;
  box.append(h);
  if (body) {
    const p = document.createElement('p');
    p.textContent = body;
    box.append(p);
  }
  root.replaceChildren(box);
}
