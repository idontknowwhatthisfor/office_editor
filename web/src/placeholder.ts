import type { DocInfo } from './bridge';
import { t } from './i18n';
import { showMessage } from './message';

const typeName = { xlsx: 'excel', pptx: 'powerpoint' } as const;

/** Stand-in for the XLSX/PPTX viewers (PRD milestones M3–M4). */
export function showNotYetSupported(root: HTMLElement, info: DocInfo, data: ArrayBuffer): void {
  const type = t(typeName[info.format as keyof typeof typeName]);
  // Office files are ZIP archives, which start with "PK".
  const head = new Uint8Array(data, 0, Math.min(2, data.byteLength));
  if (head[0] !== 0x50 || head[1] !== 0x4b) {
    showMessage(root, t('notZip', { type }));
    return;
  }
  showMessage(root, t('notYetTitle'), t('notYetBody', { type }));
}
