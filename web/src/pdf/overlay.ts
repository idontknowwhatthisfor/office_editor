// The layer the PDF editor's popovers live in: its settings panels, toast and dialogs.
//
// Pinch-zoom is the WebView's own, and a fixed element belongs to the layout viewport, so it
// would be magnified with the page (and partly pushed off screen). This layer is counter-
// transformed to cover exactly the visible area at normal size, as the bottom bar is (see
// pinToViewport in pagebar.ts); fixed elements inside it are placed within it.

let root: HTMLElement | null = null;

export function overlay(): HTMLElement {
  if (root) return root;
  const layer = document.createElement('div');
  layer.className = 'pdf-overlay';
  document.body.append(layer);
  root = layer;
  const viewport = window.visualViewport;
  if (viewport) {
    const pin = () => {
      const zoomed = viewport.scale !== 1 || viewport.offsetLeft || viewport.offsetTop;
      layer.style.transform = zoomed
        ? `translate(${viewport.offsetLeft}px, ${viewport.offsetTop}px) scale(${1 / viewport.scale})`
        : '';
    };
    viewport.addEventListener('scroll', pin);
    viewport.addEventListener('resize', pin);
    pin();
  }
  return layer;
}
