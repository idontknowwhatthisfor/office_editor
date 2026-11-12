import './style.css';
import { nativeDocument, type DocFormat, type DocInfo } from './bridge';
import { t } from './i18n';
import { showMessage } from './message';

const app = document.getElementById('app')!;

async function open(info: DocInfo, load: () => Promise<ArrayBuffer>): Promise<void> {
  showMessage(app, t('opening'));
  // Each format's viewer is its own chunk, so opening a PDF never loads spreadsheet code.
  if (info.format === 'pdf' && info.url) {
    // PDFs can be huge; the viewer reads the file from its URL in pieces, as pages are needed.
    console.log(`opening pdf document from ${info.url}, ${info.size} bytes`);
    const { openPdf } = await import('./pdf/viewer');
    await openPdf(app, { url: info.url });
    return;
  }
  let data: ArrayBuffer;
  try {
    data = await load();
  } catch (e) {
    console.error('load failed', e);
    showMessage(app, t('loadFailed'));
    return;
  }
  console.log(`opened ${info.format} document, ${data.byteLength} bytes`);
  if (info.format === 'pdf') {
    const { openPdf } = await import('./pdf/viewer');
    await openPdf(app, { data });
  } else if (info.format === 'docx') {
    const { openDocx } = await import('./docx/viewer');
    await openDocx(app, data);
  } else if (info.format === 'xlsx') {
    const { openXlsx } = await import('./xlsx/viewer');
    await openXlsx(app, data);
  } else {
    const { showNotYetSupported } = await import('./placeholder');
    showNotYetSupported(app, info, data);
  }
}

async function fetchBytes(url: string): Promise<ArrayBuffer> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  return response.arrayBuffer();
}

const doc = nativeDocument();
// Dev server only: `?file=/@fs/abs/path/to/test-files/sample.docx` opens a file without the picker.
const devFile = import.meta.env.DEV ? new URLSearchParams(location.search).get('file') : null;
if (doc) {
  void open(doc, () => fetchBytes(doc.url));
} else if (devFile) {
  const name = devFile.split('/').pop()!;
  const format = name.split('.').pop()!.toLowerCase() as DocFormat;
  void open({ name, format, mime: '', size: -1, url: devFile }, () => fetchBytes(devFile));
} else {
  // Desktop browser during development: pick a local file.
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.pdf,.docx,.xlsx,.pptx';
  input.setAttribute('aria-label', t('chooseFile'));
  input.className = 'dev-picker';
  input.onchange = () => {
    const file = input.files?.[0];
    if (!file) return;
    const format = file.name.split('.').pop()!.toLowerCase() as DocFormat;
    void open(
      { name: file.name, format, mime: file.type, size: file.size, url: '' },
      () => file.arrayBuffer(),
    );
  };
  app.append(input);
}
