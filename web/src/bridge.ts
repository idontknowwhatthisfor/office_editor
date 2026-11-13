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
  /** Open in edit mode (a new blank document). */
  edit?: boolean;
}

declare global {
  interface Window {
    OfficeBridge?: {
      getDocument(): string;
      /** Editing (see save.ts). */
      setDirty?(dirty: boolean): void;
      beginSave?(): number;
      writeChunk?(token: number, base64: string): void;
      finishSave?(token: number, mode: string): void;
    };
  }
}

export function nativeDocument(): DocInfo | null {
  return window.OfficeBridge ? (JSON.parse(window.OfficeBridge.getDocument()) as DocInfo) : null;
}
