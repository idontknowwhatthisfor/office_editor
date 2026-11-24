# OfficeEditor

A free, offline Android viewer and editor for PDF, Word (`.docx`), Excel (`.xlsx`) and PowerPoint
(`.pptx`) files. No account, no ads, no analytics, and no `INTERNET` permission: files never leave
the device.

Current version: 1.3.1 (version code 10). Application id: `com.officeeditor`.

## What it does

| Format | View | Edit |
| --- | --- | --- |
| PDF | Pages, pinch-zoom, text selection, search, thumbnails | Highlight, underline, pen, text boxes, signature (move, resize), eraser, form filling, rotate/delete pages |
| Word | Fixed page layout as in Word, search, thumbnails, charts | Text, bold/italic/underline, colours, font size, alignment, headings, lists, pictures |
| Excel | Grid with styles, frozen panes, sheet tabs, charts | Cell values and formulas (recalculated), insert/delete rows and columns, font size |
| PowerPoint | Slides, thumbnails, Present mode, charts | Text in boxes and placeholders, move and resize shapes |

Also: a tap on a PDF, Word or PowerPoint page slides the bars away for reading; new blank
document or spreadsheet, "Open with" and share-to from other apps, recent files, Export as PDF
(Word, Excel, PowerPoint), English and Indonesian.

## Tech stack

- **Native shell**: Kotlin + Jetpack Compose (Material 3), AndroidX WebKit. `minSdk` 26,
  `targetSdk` 36, `compileSdk` 37. Gradle 9.3.1, AGP 9.1.1, Kotlin 2.3.21.
- **Viewers and editors**: a Vite 8 + TypeScript 7 web app bundled into the APK's assets and run
  in a WebView. Libraries: PDF.js, pdf-lib, docx-preview, jszip, ssf. The PowerPoint renderer, the
  spreadsheet reader and formula engine, and the chart drawing are our own code.

## Architecture in brief

```
MainActivity (home) ──► ViewerActivity ──► WebView
                          │  top bar, PDF/   │  https://appassets.androidplatform.net/
                          │  Word bottom bar │    assets/web/index.html   (the web app)
                          │  and toolbars,   │    doc/current             (the document's bytes)
                          │  Save dialog,    │
                          │  print, share    │
                          └── OfficeBridge ◄─┘  metadata, dirty flag, save chunks, print, present,
                                                bar and toolbar state, hide bars
```

- The page asks the bridge which document is open, then `fetch()`es its bytes from a URL the app
  serves itself (with HTTP `Range`, so large PDFs are read in pieces).
- Each format's viewer is a separate chunk. Editors change the file's own XML/PDF in place and
  leave untouched parts as they are.
- On Save the editor builds the file, re-reads it, and streams it to the shell, which writes a
  temporary file, checks it, and only then replaces the original (with a backup) or saves a copy.
- The PDF and Word bottom bar and editing toolbars are drawn by the shell over the WebView, from
  state the page sends: a page element can only follow a pinch-zoomed view a frame late, so it
  lagged behind fast pans. The top bar floats over PDF, Word and PowerPoint pages, so hiding
  the bars never resizes the page.

## Directory map

| Path | What it is |
| --- | --- |
| [`app/`](app/README.md) | Android application module: manifest, Gradle build, resources. |
| [`app/src/main/java/com/officeeditor/`](app/src/main/java/com/officeeditor/README.md) | The Kotlin shell: activities, WebView + bridge, saving, printing, recent files. |
| [`web/`](web/README.md) | The Vite + TypeScript web app. |
| [`web/src/`](web/src/README.md) | Entry point and modules shared by all formats. |
| [`web/src/pdf/`](web/src/pdf/README.md) | PDF viewer and editor. |
| [`web/src/docx/`](web/src/docx/README.md) | Word viewer and editor. |
| [`web/src/xlsx/`](web/src/xlsx/README.md) | Spreadsheet viewer, formula engine and editor. |
| [`web/src/pptx/`](web/src/pptx/README.md) | Presentation renderer, viewer and editor. |
| [`web/src/chart/`](web/src/chart/README.md) | Chart drawing for Word, Excel and PowerPoint files. |
| [`web/scripts/`](web/scripts/README.md) | Pre-build scripts: PDF.js assets, licence list, font CSS. |
| `test-files/` | Sample-generating and Office-checking scripts; see [Testing](#testing). |
| `gradle/` | `libs.versions.toml` (version catalog) and the Gradle wrapper. |
| `.claude/launch.json` | Dev-server launch entry (`web-viewers`, port 5183). |

Root files:

| File | What it is |
| --- | --- |
| `CLAUDE.md` | Detailed architecture notes and project rules. The most complete description of how the code works. |
| `build.gradle.kts`, `settings.gradle.kts`, `gradle.properties` | Root Gradle setup; the only module is `:app`. |
| `gradlew`, `gradlew.bat` | Gradle wrapper. |
| `.gitignore` | Ignores build output, signing files, the generated web assets, and everything in `test-files/` except the scripts it names. |
| `keystore.properties.example` | Template for release signing; copy to `keystore.properties` (git-ignored). |

## Build and run

Requirements: the Android SDK, a JDK that Gradle 9.3.1 supports (Android Studio's bundled one
works), and Node with `npm` on `PATH`.

Install the web dependencies once:

```bash
cd web && npm ci
```

Build and install the debug app. On this Mac the system JDK is too new for Gradle, so point
`JAVA_HOME` at Android Studio's:

```bash
export JAVA_HOME="/Applications/Android Studio.app/Contents/jbr/Contents/Home"
```

```bash
./gradlew :app:installDebug
```

Gradle's `preBuild` runs `npm run build` in `web/`, so the web app is rebuilt as needed.

Web-only development in a desktop browser (a file picker replaces the bridge):

```bash
cd web && npm run dev
```

Open a file directly with `http://localhost:5173/?file=/@fs/<repo>/test-files/sample.docx`
(`&edit` for edit mode, `&print` for the print layout).

### Release build

```bash
./gradlew :app:assembleRelease
```

```bash
./gradlew :app:bundleRelease
```

Outputs: `app/build/outputs/apk/release/` and an `.aab` for Google Play. Code shrinking is on, so
test the release build on a phone before publishing. Signing reads `keystore.properties` (copy
`keystore.properties.example`); without it the build is signed with the debug key and must not be
published. Check the signer with `apksigner verify --print-certs`.

### Debugging on a device

```bash
adb logcat -s OfficeEditorWeb
```

Debug builds' WebView is inspectable from `chrome://inspect`.

Things that cost time before:

- **Play Protect can reject `adb install`** with `INSTALL_FAILED_VERIFICATION_FAILURE` and hide
  its prompt. The owner allows it on the phone (Developer options → "Verify apps over USB");
  don't change that setting from the computer.
- **The phone locks.** Check with `adb shell dumpsys window | grep isKeyguardShowing` and ask the
  owner to unlock. Never enter the PIN.
- **Files pushed with `adb push` are invisible to the picker** until a rescan:
  `adb shell content call --uri content://media --method scan_volume --arg external_primary`.
- **Opening a file from the shell:** find its MediaStore id, then
  `adb shell am start -a android.intent.action.VIEW --grant-read-uri-permission -n com.officeeditor/.ViewerActivity -t <mime> -d content://media/external/file/<id>`.
  `content://com.android.externalstorage...` URIs are refused for the shell.
- **Driving the WebView:**
  `adb forward tcp:9222 localabstract:webview_devtools_remote_$(adb shell pidof com.officeeditor)`,
  then the DevTools protocol on `127.0.0.1:9222` (`Runtime.evaluate`,
  `Input.synthesizePinchGesture`). `adb shell input` can't pinch.
- **Desktop browser checks** are much faster than the phone for layout work. Use phone width,
  and still confirm on the phone: one fit-to-width bug only showed there.
- zsh does not word-split unquoted variables; use shell functions for repeated `adb` commands.

### Releasing

1. Raise `versionCode`/`versionName` in `app/build.gradle.kts` and the version in
   `web/package.json` (`npm version x.y.z --no-git-tag-version`).
2. Build `assembleRelease` and `bundleRelease`; check the signer with `apksigner`.
3. Test the release APK on the phone.
4. Tag `vx.y.z` and push.

The GitHub release page and any Play upload (with `mapping.txt`) are done by the owner.

Known limits of each part are listed in that part's README.

## Testing

There are no automated tests. Changes are checked in the desktop browser, on a phone, and by
opening saved files in Word, Excel and PowerPoint.

`test-files/` holds the sample documents and the scripts for this. Only the scripts are tracked;
the documents are git-ignored (personal, large or copyrighted) and are regenerated or supplied
locally. `.gitignore` lists the tracked scripts by name, so a new one needs its own
`!test-files/<name>` line.

| Script | Needs | What it does |
| --- | --- | --- |
| `make_samples.py` | Python 3 (standard library) | Writes `sample-5-pages.pdf`, `sample-200-pages.pdf`, `sample.docx` and the invalid `not-a-pdf.pdf`. |
| `make_pdf_form.mjs` | Node, `web/node_modules` | Writes `pdf-form.pdf`: a fillable form plus a page with `/Rotate 90`. |
| `make_xlsx_samples.applescript` | Microsoft Excel | Builds `excel-sample.xlsx` and `excel-5000-rows.xlsx` in Excel itself. |
| `make_xlsx_charts.applescript` | Microsoft Excel | Builds `xlsx_chart.xlsx` (embedded charts, a chart sheet) and Excel's PDF of it. |
| `excel_text_dump.applescript` | Microsoft Excel | Prints the text Excel displays per cell as JSON lines, to compare with the viewer. |
| `make_pptx_sample.py` | Python 3, `python-pptx` | Writes `pptx-sample-raw.pptx`. |
| `pptx_resave.applescript` | Microsoft PowerPoint | Re-saves a deck in PowerPoint and exports its PDF, the reference rendering. |
| `make_pptx_native.applescript` | Microsoft PowerPoint | Builds `pptx-native.pptx` with PowerPoint's default theme, plus its PDF. |
| `word_to_pdf.sh` | Microsoft Word | Opens a `.docx` in Word and saves Word's PDF beside it, to check edited files. |

Each script's header gives its exact command. The Office apps are sandboxed and write to
`~/Downloads`; the header also gives the `mv` that follows. `test-files/out/` receives files posted
to the dev server's `/__dev-save`.

## Project rules

- No `INTERNET` permission, no network requests, no analytics. The WebView refuses every URL the
  app doesn't serve itself.
- No AGPL or commercial-only dependencies; record each new library's licence (and add it to
  `web/scripts/collect-licences.mjs`).
- Every user-visible string exists in English and Indonesian: `res/values` + `res/values-in` for
  native, `web/src/i18n.ts` for web.
- Never overwrite an original file until the new one is written and re-parsed.

## Licences

Third-party components bundled in the app (full texts are shown in the app's About & licences
screen):

| Library | Licence | Used for |
| --- | --- | --- |
| pdfjs-dist | Apache-2.0 (bundled decoders carry their own BSD/MIT licences) | PDF rendering, text, search |
| pdf-lib (+ @pdf-lib/standard-fonts, @pdf-lib/upng, pako, tslib) | MIT (pako: MIT and Zlib) | Writing edited PDFs |
| docx-preview | Apache-2.0 | DOCX rendering |
| jszip and its dependencies | MIT (dual MIT / GPL-3.0; used under MIT), MIT / ISC | Reading and writing Office zips |
| ssf + frac | Apache-2.0 | Number and date formats |
| @fontsource/tinos, arimo, carlito | OFL-1.1 | Stand-ins for Times New Roman, Arial, Calibri |
| PDF.js runtime files (decoders, colour management, fonts, character maps) | Apache-2.0, BSD-2/3-Clause, MIT, CC0; Liberation Sans: GPL-2.0 with font exception | PDF rendering |
| AndroidX, Kotlin, kotlinx.coroutines | Apache-2.0 | Native shell |

The repository has no licence file for the project's own code.
