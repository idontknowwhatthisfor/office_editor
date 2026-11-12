// Saving, shared by the editors. The native shell owns the Save button and dialog; when the user
// saves, it calls window.officeSave(mode). The editor then builds the new file, re-reads it to
// make sure it is valid, and streams it to the shell in base64 chunks (the JS bridge passes
// strings only). The shell writes it to a temporary file before touching the original.

declare global {
  interface Window {
    officeSave?: (mode: 'overwrite' | 'copy') => Promise<void>;
    officeSaveResult?: (result: SaveResult) => void;
  }
}

export interface SaveResult {
  ok: boolean;
  mode: string;
  cancelled: boolean;
}

export interface Editor {
  /** Builds the edited file and checks it can be read back; throws if not. */
  produce(): Promise<Uint8Array>;
  /** Called after the shell has written the file. */
  saved?(result: SaveResult): void;
  /** File name for the desktop-browser fallback (no shell). */
  fileName: string;
}

let current: Editor | null = null;
let dirty = false;

export function registerEditor(editor: Editor): void {
  current = editor;
  window.officeSave = save;
  window.officeSaveResult = (result) => {
    if (result.ok) setDirty(false);
    current?.saved?.(result);
  };
  if (!window.OfficeBridge) devSaveButton();
}

export function setDirty(value: boolean): void {
  if (value === dirty) return;
  dirty = value;
  window.OfficeBridge?.setDirty?.(value);
  devButton?.toggleAttribute('hidden', !value);
}

export function isDirty(): boolean {
  return dirty;
}

const CHUNK = 3 * 256 * 1024; // a multiple of 3, so chunks encode without padding in between

async function save(mode: 'overwrite' | 'copy'): Promise<void> {
  const bridge = window.OfficeBridge;
  let bytes: Uint8Array | null = null;
  try {
    bytes = await current!.produce();
  } catch (e) {
    console.error('could not build the edited file', e);
  }
  if (!bridge?.beginSave) {
    if (bytes) download(bytes, current!.fileName);
    return;
  }
  const token = bridge.beginSave();
  // A failed build sends nothing: the shell then finds an empty file and reports the failure.
  if (bytes) {
    for (let at = 0; at < bytes.length; at += CHUNK) bridge.writeChunk!(token, base64(bytes.subarray(at, at + CHUNK)));
  }
  bridge.finishSave!(token, mode);
}

function base64(bytes: Uint8Array): string {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

// ---- desktop browser (development): no shell, so saving downloads the file ----

let devButton: HTMLButtonElement | null = null;

function devSaveButton(): void {
  if (devButton) return;
  devButton = document.createElement('button');
  devButton.className = 'dev-save';
  devButton.textContent = 'Download edited file';
  devButton.hidden = true;
  devButton.onclick = () => void save('copy');
  document.body.append(devButton);
}

function download(bytes: Uint8Array, name: string): void {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart]));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}
