# Handoff — OfficeEditor

Written 12 Nov 2026; updated at the end of the second session (M3). Read `CLAUDE.md` first (architecture,
build commands, rules from the PRD); this file is the state of the work and what to do next.

## Where things stand

| Milestone (PRD §8) | State |
| --- | --- |
| M0 Skeleton | Done, verified on the phone. |
| M1 PDF viewer | Working on the phone: scroll, pinch-zoom, text selection, search, thumbnails. Tested on generated samples, one Word-exported PDF and a 700 MB / 1,191-page textbook. |
| M2 DOCX viewer | Working on the phone in Word-style page layout. Tested on one real Word file (`test-files/FallSafe_Essay_v5-2.docx`) and two synthetic ones. |
| M3 XLSX viewer | Done; the owner checked it on the phone (session 2). |
| M4 PPTX viewer | Working on the phone (session 2): text, shapes, pictures, tables, groups, backgrounds; swipe, thumbnails, double-tap/pinch zoom. Compared slide by slide with PowerPoint's own PDF of two decks. **Needs the owner's check and real decks.** |
| M5–M8 | Not started. |

Commits on `main`, nothing pushed (no remote).

## What the owner asked for, beyond the PRD

- **DOCX must look like Word**: fixed pages, no reflow, and the whole page zooms. An earlier
  reflowed "reading mode" was replaced at their request; don't bring it back unasked.
- Use the PRD as the guide and the phone (Galaxy S21 FE, Android 16, over `adb`) to build and
  debug. They follow along in short messages and unlock the phone when asked.
- Global instruction: run `./gradlew --stop` after Gradle builds (8 GB Mac).

## Next

- Owner checks M4 on the phone (pinch-zoom by finger, real decks). Every viewer is in place, so
  the PRD's "first usable viewer build" is reached once that passes.
- Then M5 (PDF annotate, fill and sign) per PRD §8.

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

**XLSX viewer** (`web/src/xlsx/`)
- Values are Excel's cached results. A file written by a library that leaves formulas without
  results (some exporters do) shows those cells empty. No formula engine yet (M6).
- Not shown: charts, pictures, shapes (a notice says so), conditional formatting, data
  validation, comments, rotated text, row/column default styles on empty cells, hyperlinks.
- Text in a frozen column is cut at the frozen edge instead of spilling into the next column.
- Built-in date/time formats follow en-US, not the device locale (Excel follows its locale).
- Parsing the 5,000-row sample takes ~0.5 s on the phone (open-to-grid ~1.3 s); a DOM parse of a
  much larger sheet will be slow and memory-hungry. The PRD's "read-only beyond 50k cells"
  notice belongs to M6.
- The comparison against Excel (`excel_text_dump.applescript`, then the formats in a browser
  console) matched 9,082 of 9,087 cells; the rest are the Mac's Indonesian time separator.

**PPTX viewer** (`web/src/pptx/`)
- Not drawn: charts (placeholder + notice; their cached values are in `ppt/charts/*.xml`, so a
  simple bar/line/pie renderer is feasible), SmartArt without a saved drawing, EMF/WMF/TIFF
  pictures, video/audio (poster frame only), 3-D, most effects (only outer shadows), speaker
  notes (PRD P1), transitions.
- Built-in table styles other than the default ("Medium Style 2 – Accent 1") are drawn as that
  one: files store only their id.
- About 50 common preset shapes are drawn exactly; cloud shapes are approximated; others fall
  back to a rectangle (`geometry.ts`). Custom shapes with computed guides draw with them as 0.
- Text is laid out by the browser, so line breaks differ slightly from PowerPoint, and Aptos
  (Office's current default font) falls back to the system sans-serif.
- Rendering is lazy (current slide ±2, thumbnails as they scroll in); ~40–70 ms per slide on
  the phone for the samples. Not tried on a deck with hundreds of slides or large photos.

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

`excel-sample.xlsx` (4 sheets: budget with shared formulas and frozen panes, formats, a hidden
sheet, hidden rows/columns without gridlines) and `excel-5000-rows.xlsx` were made by Mac Excel
via `make_xlsx_samples.applescript` (Excel's locale here is Indonesian, so currency is "Rp").
Both are on the phone in `Download/officeeditor-test/` (MediaStore ids 2130, 2131).
`pptx-sample.pptx` (python-pptx deck re-saved by PowerPoint: bullets, shapes, picture, table,
group, chart, dark background) and `pptx-native.pptx` (made in PowerPoint, default theme) have
PowerPoint's PDFs beside them; on the phone as ids 2132, 2133.

`test-files/` is git-ignored apart from `make_samples.py` (run it to regenerate the synthetic
PDFs and DOCX). The owner's essay (`FallSafe_Essay_v5-2.docx`) is there; Word's PDF export of it
is at `~/Downloads/FallSafe_Essay_v5-2.pdf` and is the reference for DOCX pagination. Copies of
the samples and both essay files are on the phone in `Download/officeeditor-test/`.
