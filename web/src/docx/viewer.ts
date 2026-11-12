import './fonts.css';
import './docx.css';
import { parseAsync, renderDocument } from 'docx-preview';
import { t } from '../i18n';
import { showMessage } from '../message';
import { paginate } from './paginate';

/** Space around the pages, in CSS px. */
const GUTTER = 8;

/**
 * Read-only Word viewer in page layout: pages have the document's paper size, margins, fonts and
 * sizes, as in Word, and the whole page is scaled to the screen width. Pinch-zoom (the WebView's
 * own) then magnifies the page as a whole; nothing reflows.
 *
 * docx-preview turns the document's XML into styled HTML; paginate.ts then breaks it into pages.
 */
export async function openDocx(root: HTMLElement, data: ArrayBuffer): Promise<void> {
  const view = document.createElement('div');
  view.className = 'docx-view';
  try {
    const options = { className: 'docx', inWrapper: true, breakPages: true, useBase64URL: true };
    const parsed = await parseAsync(data, options);
    markKeepWithNext(parsed);
    view.append(...(await renderDocument(parsed, options)));
  } catch (e) {
    console.error('docx open failed', e);
    showMessage(root, t('docxInvalid'));
    return;
  }

  // The native shell blocks all navigation, but don't leave script or file URLs in the page.
  for (const link of view.querySelectorAll('a[href]')) {
    const href = link.getAttribute('href')!;
    if (!/^(https?:|mailto:|#)/i.test(href)) link.removeAttribute('href');
  }

  const lineHeights = scaleLineHeights(view);

  root.replaceChildren(view);
  const wrapper = view.querySelector<HTMLElement>('.docx-wrapper');
  if (!wrapper) return;

  // Page breaks depend on line breaks, which depend on the fonts: wait for the ones in use.
  void view.offsetHeight;
  do await document.fonts.ready;
  while (document.fonts.status === 'loading');
  lineHeights.measure();
  paginate(wrapper);
  numberPages(wrapper);

  // Scale the pages to the screen width. A transform only scales the painted result, so the
  // layout stays exactly as paginated (CSS zoom would re-lay text out at the smaller size, where
  // pixel rounding shifts line heights). The view is given the scaled size to scroll over.
  const pageWidth = Math.max(...[...wrapper.children].map((page) => (page as HTMLElement).offsetWidth));
  wrapper.style.width = `${pageWidth}px`;
  const pagesHeight = wrapper.offsetHeight;
  wrapper.style.transformOrigin = '0 0';
  const fit = () => {
    const available = document.documentElement.clientWidth - 2 * GUTTER;
    const scale = Math.min(1, available / pageWidth);
    wrapper.style.transform = `scale(${scale})`;
    view.style.width = `${pageWidth * scale}px`;
    view.style.height = `${pagesHeight * scale}px`;
  };
  fit();
  let width = document.documentElement.clientWidth;
  window.addEventListener('resize', () => {
    if (document.documentElement.clientWidth !== width) {
      width = document.documentElement.clientWidth;
      fit();
    }
  });
}

/**
 * Word's "multiple" line spacing (1.0, 1.15, 1.5…) multiplies the font's own line height, which
 * is about 1.15 × the font size for Times New Roman or Arial. docx-preview writes it as a CSS
 * `line-height: 1.15`, which multiplies the font size, so lines come out too tight and too many
 * fit on a page. This rewrites every such line-height to be multiplied by --docx-line, and
 * `measure()` (call once fonts are loaded) sets --docx-line per paragraph from its font.
 */
function scaleLineHeights(view: HTMLElement): { measure(): void } {
  // Only unitless values are multiples; "line-height: 14pt" is Word's exact spacing.
  const rewrite = (css: string) =>
    css.replace(/line-height:\s*([\d.]+)\s*(?=;|}|$)/g, 'line-height: calc($1 * var(--docx-line, 1.15))');
  for (const sheet of view.querySelectorAll('style')) sheet.textContent = rewrite(sheet.textContent ?? '');
  for (const styled of view.querySelectorAll<HTMLElement>('[style*="line-height"]')) {
    styled.setAttribute('style', rewrite(styled.getAttribute('style')!));
  }

  return {
    measure() {
      const probe = document.createElement('span');
      probe.textContent = 'x';
      probe.style.cssText = 'position:absolute;visibility:hidden;font-size:1000px;line-height:normal';
      view.append(probe);
      const ratios = new Map<string, string>();
      for (const p of view.querySelectorAll<HTMLElement>('p')) {
        // Fonts are usually set on the runs (spans), not the paragraph.
        const family = getComputedStyle(p.querySelector('span') ?? p).fontFamily;
        // A line is as tall as the paragraph's own font and its runs' fonts need together; give
        // the paragraph its runs' font so two different fonts don't add stray fractions per line.
        p.style.fontFamily = family;
        let ratio = ratios.get(family);
        if (!ratio) {
          probe.style.fontFamily = family;
          ratios.set(family, (ratio = String(probe.getBoundingClientRect().height / 1000)));
        }
        p.style.setProperty('--docx-line', ratio);
      }
      probe.remove();
    },
  };
}

/**
 * Headers and footers are copied to every page with the page number Word last saved in them.
 * A header/footer run that is just a number is taken to be the page-number field.
 */
function numberPages(wrapper: HTMLElement): void {
  [...wrapper.children].forEach((page, index) => {
    for (const run of page.querySelectorAll(':scope > header span, :scope > footer span')) {
      if (run.childElementCount === 0 && /^\d+$/.test(run.textContent ?? '')) {
        run.textContent = String(index + 1);
      }
    }
  });
}

/**
 * Word's "keep with next" (set on headings) is parsed by docx-preview but not rendered. Carry it
 * into the HTML as `break-after: avoid`, which paginate.ts honours. `doc` is docx-preview's
 * parsed model, which has no public types.
 */
function markKeepWithNext(doc: any): void {
  const styles = new Map<string, any>((doc.stylesPart?.styles ?? []).map((style: any) => [style.id, style]));
  const styleKeeps = (id: string | undefined): boolean => {
    for (let hops = 0; id && hops < 20; hops++) {
      const style = styles.get(id);
      if (style?.paragraphProps?.keepNext !== undefined) return style.paragraphProps.keepNext;
      id = style?.basedOn;
    }
    return false;
  };
  const visit = (element: any): void => {
    if (element.type === 'paragraph' && (element.keepNext ?? styleKeeps(element.styleName))) {
      element.cssStyle = { ...element.cssStyle, 'break-after': 'avoid' };
    }
    for (const child of element.children ?? []) visit(child);
  };
  if (doc.documentPart?.body) visit(doc.documentPart.body);
}
