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
  /** Lay the document out for "Export as PDF" (print.ts) instead of showing it. */
  print?: boolean;
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
      /** Export as PDF: the print layout is ready (page size in CSS px), or failed. */
      printReady?(width: number, height: number): void;
      printFailed?(): void;
      /** A presentation is shown full screen (the shell hides its bars and turns landscape). */
      setPresenting?(on: boolean): void;
      /** A tap hid the viewer's bars (immersive.ts): the shell hides its top bar too, or shows it again. */
      setBarsHidden?(on: boolean): void;
      /** Height of the shell's top bar floating over the page, in CSS px. */
      getTopInset?(): number;
      /** What the shell's native bottom bar shows: a NativeBarState as JSON (pagebar.ts). */
      setPageBar?(json: string): void;
    };
  }
}

export function nativeDocument(): DocInfo | null {
  return window.OfficeBridge ? (JSON.parse(window.OfficeBridge.getDocument()) as DocInfo) : null;
}
