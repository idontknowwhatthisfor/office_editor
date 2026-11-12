# Handoff — OfficeEditor

Written 12 Nov 2026 at the end of the first working session. Read `CLAUDE.md` first (architecture,
build commands, rules from the PRD); this file is the state of the work and what to do next.

## Where things stand

| Milestone (PRD §8) | State |
| --- | --- |
| M0 Skeleton | Done, verified on the phone. |
| M1 PDF viewer | Working on the phone: scroll, pinch-zoom, text selection, search, thumbnails. Tested on generated samples, one Word-exported PDF and a 700 MB / 1,191-page textbook. |
| M2 DOCX viewer | Working on the phone in Word-style page layout. Tested on one real Word file (`test-files/FallSafe_Essay_v5-2.docx`) and two synthetic ones. |
| M3 XLSX viewer | **Next.** Not started; `.xlsx` opens to a "can't be viewed yet" message (`web/src/placeholder.ts`). |
| M4 PPTX viewer | Not started; same placeholder. |
| M5–M8 | Not started. |

Two commits on `main`, working tree clean, nothing pushed (no remote).

## What the owner asked for, beyond the PRD

- **DOCX must look like Word**: fixed pages, no reflow, and the whole page zooms. An earlier
  reflowed "reading mode" was replaced at their request; don't bring it back unasked.
- Use the PRD as the guide and the phone (Galaxy S21 FE, Android 16, over `adb`) to build and
  debug. They follow along in short messages and unlock the phone when asked.
- Global instruction: run `./gradlew --stop` after Gradle builds (8 GB Mac).

## Next: M3, the XLSX viewer

PRD scope: multiple sheets with tabs, frozen header rows, zoom, computed formula values. Exit
criterion: a corpus of real Excel files shows correct values.

- Add `web/src/xlsx/viewer.ts` and a branch in `web/src/main.ts` (same pattern as `pdf/` and
  `docx/`); remove `xlsx` from `placeholder.ts`.
- The PRD suggests SheetJS CE or ExcelJS for parsing and Univer or Fortune-sheet for the grid.
  Check licence and bundle size before choosing; a full grid library may be more than a viewer
  needs. Record the choice and licence in `CLAUDE.md`.
- There are no `.xlsx` test files yet. `test-files/make_samples.py` can be extended (it writes
  DOCX by hand with `zipfile`), but ask the owner for real Excel files too: every bug of
  substance so far was found with a real file, not a generated one.
- XLSX is fetched whole (`fetchBytes`); only PDF reads by range.

## Open items

**DOCX page layout** (`web/src/docx/`)
- On the test essay 17 of 20 page breaks equal Word's; pages 3, 14 and 16 break one or two lines
  early. Cause: Word compresses spaces in justified lines to fit another word; CSS can't.
- Not handled, and untested: footnotes at the page bottom, multi-column sections, text boxes and
  floating images, `keepLines`, a paragraph or table row taller than a page.
- Page numbers are a guess: any header/footer run that is only digits is replaced with the page
  index (`numberPages` in `viewer.ts`). A year in a footer would be overwritten.
- Only Times New Roman, Arial and Calibri have stand-in fonts, Latin subsets only
  (`scripts/make-docx-fonts.mjs`). Other fonts fall back to the system font and will paginate
  differently.
- Links are inert. Pagination runs synchronously after render; not timed on long documents.
- Only one real Word file tested. The PRD's fidelity warning (unsupported content) belongs to M7.

**PDF viewer** (`web/src/pdf/`)
- Memory grows as a large PDF is read: PDF.js keeps every range it has loaded. Not measured to
  the end of a big file.
- Zoomed pages are capped at about 6 megapixels, so sharpness stops improving past roughly 2×.
- No password prompt (shows a message), no link handling, no annotation layer.
- Range serving reports sizes through `InputStream.available()`, an `int`: files over 2 GB will
  misbehave.

**Shell** (`app/`)
- `resolveDocument` runs on the main thread in `ViewerActivity.onCreate`; a slow provider
  (Drive) could stall the UI.
- Share copies the whole file to cache first; slow for very large files.
- No "New document / spreadsheet", save, settings or licences screen yet (PRD §4.1, §7).
- Rotation and tablet layouts have never been tried on either viewer.

## Things that cost time and will again

- **Play Protect can reject `adb install`** with `INSTALL_FAILED_VERIFICATION_FAILURE` and hide
  its prompt. The owner turned off "Verify apps over USB" on the phone. If it returns, ask them;
  don't change the setting.
- **The phone locks.** `adb shell dumpsys window | grep isKeyguardShowing`; ask the owner to
  unlock. Never enter the PIN.
- **Files pushed with `adb push` are invisible to the picker** until a rescan:
  `adb shell content call --uri content://media --method scan_volume --arg external_primary`.
- **Opening a file from the shell:** find its MediaStore id, then
  `adb shell am start -a android.intent.action.VIEW --grant-read-uri-permission -n com.officeeditor/.ViewerActivity -t <mime> -d content://media/external/file/<id>`.
  `content://com.android.externalstorage...` URIs are refused for the shell.
- **Driving the WebView:** `adb forward tcp:9222 localabstract:webview_devtools_remote_$(adb shell pidof com.officeeditor)`,
  then the DevTools protocol on `127.0.0.1:9222` (`Runtime.evaluate`,
  `Input.synthesizePinchGesture`). `adb shell input` can't pinch; `uiautomator dump` returned
  nothing useful on this phone.
- **WebView range requests:** see the `EndingAt` comment in `DocumentWebView.kt` before touching
  how documents are served.
- **docx-preview** only breaks pages at explicit breaks, writes Word's line-spacing multiples as
  CSS multiples of font size, hyphenates everything, and misreads `<w:tblHeader/>`. Each has a
  workaround in `web/src/docx/viewer.ts`; read its comments before upgrading the library.
- **Desktop browser checks** are much faster than the phone for layout work:
  `cd web && npm run dev`, then `?file=/@fs/<repo>/test-files/<name>`. Use phone width, and still
  confirm on the phone: one fit-to-width bug only showed there.
- zsh does not word-split unquoted variables; use shell functions for repeated `adb` commands.

## Test files

`test-files/` is git-ignored apart from `make_samples.py` (run it to regenerate the synthetic
PDFs and DOCX). The owner's essay (`FallSafe_Essay_v5-2.docx`) is there; Word's PDF export of it
is at `~/Downloads/FallSafe_Essay_v5-2.pdf` and is the reference for DOCX pagination. Copies of
the samples and both essay files are on the phone in `Download/officeeditor-test/`.
