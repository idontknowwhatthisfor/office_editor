// The native shell (Kotlin) exposes `window.OfficeBridge` inside the Android WebView.
// In a desktop browser (`npm run dev`) it is absent and main.ts shows a file picker instead.

export type DocFormat = 'pdf' | 'docx' | 'xlsx' | 'pptx';

export interface DocInfo {
  name: string;
  format: DocFormat;
  mime: string;
  /** Bytes, or -1 if the provider didn't say. */
  size: number;
  /** Same-origin URL that streams the file's bytes. */
  url: string;
}

declare global {
  interface Window {
    OfficeBridge?: { getDocument(): string };
  }
}

export function nativeDocument(): DocInfo | null {
  return window.OfficeBridge ? (JSON.parse(window.OfficeBridge.getDocument()) as DocInfo) : null;
}
