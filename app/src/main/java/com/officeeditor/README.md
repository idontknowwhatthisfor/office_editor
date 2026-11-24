# com.officeeditor — the native shell (Kotlin, Jetpack Compose)

## Purpose

The Android side of the app: the home screen, the screen hosting one document in a WebView, the
JS bridge, file saving, printing and the recent-files list. Document rendering and editing happen
in the web app ([`web/src/`](../../../../../../web/src/README.md)).

## Contents

| File | What it does |
| --- | --- |
| `MainActivity.kt` | Home: recent files, "Open file" (system picker), "+" → new document / spreadsheet (the save picker creates the file, it is filled from a blank template in `res/raw` and opened with `EXTRA_START_EDITING`). |
| `ViewerActivity.kt` | One document. Also the target of "Open with" (`ACTION_VIEW`) and share-to (`ACTION_SEND`). Native top bar (Save, share, ⋮ menu; floats over PDF, Word and PowerPoint pages and slides away on a tap), the PDF and Word bottom bar (`PageBar.kt`), save dialogs, Export as PDF, Present mode, the photo picker. Implements `EditorHost`. |
| `DocumentWebView.kt` | `createDocumentWebView()`: WebView setup, serving of app assets and the document's bytes, the `OfficeBridge` JS interface, and the `EditorHost` interface the bridge calls back into. |
| `Saving.kt` | `SaveReceiver` (receives an edited file in base64 chunks into a temp file and checks its signature), `canOverwrite`, `overwrite` (backs up the original, restores it on failure), `writeCopy`, `copyName`. |
| `PageBar.kt` | The PDF and Word viewers' bottom bar (thumbnails, page indicator, search, edit), drawn natively over the bottom of the WebView so it stays put while the page is pinch-zoomed. The page sends what it shows (`setPageBar`, `PageBarState`); taps go back as `window.officePageBar(action, arg)`. Hidden while the page shows its own toolbar (edit mode). |
| `Printing.kt` | `printWebView()`: opens Android's print dialog for a WebView, picking the standard paper size nearest the page. |
| `Documents.kt` | `DocFormat` (PDF/DOCX/XLSX/PPTX, detection by MIME type then extension), `DocumentInfo`, `resolveDocument()`. |
| `RecentFiles.kt` | The recent list (SharedPreferences, max 30). Keeps a persisted URI permission where possible, otherwise a private copy in `files/recent/` (files up to 50 MB). |
| `AboutActivity.kt` | About & licences: version, privacy note, and each bundled component's licence from `assets/web/licences.json`. |
| `Theme.kt` | `OfficeEditorTheme` (dynamic colours on Android 12+), format badge colours, `appBarHeight()` (48 dp on short screens). |

## How it works

- **Serving**: the WebView loads `https://appassets.androidplatform.net/assets/web/index.html`
  through `WebViewAssetLoader`. The document is served at `/doc/current` and honours HTTP `Range`
  when the file is seekable. Every other URL gets a 404, and navigation away from the viewer page
  is blocked.
- **`EndingAt`** (in `DocumentWebView.kt`): for a Range request the WebView skips to the range
  start itself but reads to the end of the stream it is given, so the stream is cut off at the
  range end. Read that comment before changing how documents are served.
- **Bridge** (`window.OfficeBridge`): `getDocument()`, `setDirty()`, `beginSave()` /
  `writeChunk()` / `finishSave()`, `printReady()` / `printFailed()`, `setPresenting()`, `setBarsHidden()` (a tap
  on the document slides the top bar away and hides the status bar), `getTopInset()` (the
  height of the top bar, which floats over PDF, Word and PowerPoint pages). Bridge
  methods run on a WebView thread, not the UI thread.
- **Shell → page**: `window.officeSave(mode)`, `window.officeSaveResult(...)`,
  `window.officeEndPresentation()`, via `evaluateJavascript`.
- **Saving**: nothing touches the original until the whole new file is on disk and has the right
  signature (`%PDF-` or a zip header). Files lent read-only ("Open with") can only be saved as a
  copy.
- **Export as PDF**: the saved file is loaded a second time in a hidden WebView with `print` set;
  when the page reports `printReady`, `printWebView()` runs.

## Dependencies & relationships

- AndroidX: Compose (Material 3), Activity, WebKit (`WebViewAssetLoader`); `FileProvider` for
  sharing (authority `${applicationId}.share`).
- Reads `../../../res/` (strings, blank templates; described in the
  [app README](../../../../../README.md#resources-srcmainres)) and the built web app in
  `assets/web`.
- Web counterparts: `web/src/bridge.ts`, `save.ts`, `print.ts`.

## Usage notes

- Web console output appears in logcat under the tag `OfficeEditorWeb`; save errors under
  `OfficeEditorSave`.
- A new `@JavascriptInterface` method needs no ProGuard change (`proguard-rules.pro` keeps them
  all), but add its type to `web/src/bridge.ts`.
- New strings go in both `res/values/strings.xml` and `res/values-in/strings.xml`.

## Known limits

Carried over from the handoff notes written at release 1.1.0. Items fixed since were dropped; the rest were not each rechecked.

- `resolveDocument` runs on the main thread in `ViewerActivity.onCreate`; a slow provider (e.g.
  Drive) could stall the UI.
- Share copies the whole file to cache first; slow for very large files.
- No settings screen.
