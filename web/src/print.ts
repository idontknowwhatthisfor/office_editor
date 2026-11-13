import './print.css';
import type { DocFormat } from './bridge';

// "Export as PDF". The shell loads the document a second time, in a hidden WebView, with
// `print` set; the page lays it out for paper and reports the page size, and the shell then
// prints that WebView through Android's print framework, whose "Save as PDF" writes the file.
// The PDF is made by the browser engine, so text stays text.

/** Lays the document out for printing; resolves to the page size in CSS px, or null. */
export async function preparePrint(root: HTMLElement, format: DocFormat, data: ArrayBuffer): Promise<{ width: number; height: number } | null> {
  document.documentElement.classList.add('printing');
  let size: { width: number; height: number } | null = null;
  if (format === 'docx') size = await (await import('./docx/viewer')).printDocx(root, data);
  else if (format === 'xlsx') size = await (await import('./xlsx/viewer')).printXlsx(root, data);
  else if (format === 'pptx') size = await (await import('./pptx/viewer')).printPptx(root, data);
  if (!size) return null;
  // The default page, for formats that don't name their own.
  const style = document.createElement('style');
  style.textContent = `@page { size: ${size.width}px ${size.height}px; margin: 0; }`;
  document.head.append(style);
  // Pictures must be decoded, and fonts loaded, before the page is printed.
  await Promise.all([...root.querySelectorAll('img')].map((img) => img.decode().catch(() => undefined)));
  await document.fonts.ready;
  return size;
}
