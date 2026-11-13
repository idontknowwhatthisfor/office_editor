# PRD — Android Office Editor (DOCX, PDF, XLSX)

Nov 12, 2026 · @kp

## 1. Overview

**Working name:** OfficeEditor

A free, lightweight Android app that opens and edits the three most common office formats — **.docx**, **.xlsx** and **.pdf** — offline, without ads, accounts or subscriptions.

**Problem.** Mainstream mobile office apps either need a paid subscription for editing, push sign-ins and cloud storage, or are heavy and slow on mid-range phones. Most mobile edits are small: fix a typo, update a few cells, fill in or sign a PDF form.

**Vision.** A fast "pocket editor" for quick, reliable edits — not a full office suite. Its value is: opens fast, saves without breaking the file, works offline, and costs nothing.

## 2. Goals and non-goals

### Goals

- Open, view and make common edits to DOCX, XLSX and PDF files on Android.
- Save back to the same format without corrupting files made in Word, Excel or other tools (round-trip safety).
- Work fully offline; no account, no ads, no tracking.
- Small install size (target < 40 MB) and cold-start-to-document under 2 s for typical files.
- Integrate with Android: "Open with", share sheet, Files app, Google Drive via the system file picker.

### Non-goals (v1)

- Full feature parity with Word/Excel (macros, pivot tables, tracked changes, complex layouts).
- Editing existing text *inside* a PDF's original content stream.
- Real-time collaboration or a cloud service of our own.
- Legacy binary formats (.doc, .xls, .ppt) and PPTX editing (PPTX is view-only) — candidates for later versions.
- iOS / desktop builds.

## 3. Target users and use cases

**Primary user:** students and young professionals who receive office files on their phone (WhatsApp, email, Drive) and need to make quick changes.

| Persona | Typical job | Format |
| --- | --- | --- |
| Student | Fix a typo in an assignment, add a line to group-project notes | DOCX |
| Organizer / treasurer | Update attendance or a budget list, add a row, fix a formula | XLSX |
| Anyone dealing with admin | Fill in and sign a form, highlight a handout, merge scans into one PDF | PDF |

### Key user stories

- As a user, I tap a .docx in WhatsApp → "Open with" → the app, fix a sentence, and share it back.
- As a user, I open a spreadsheet, change values, and formulas recalculate correctly.
- As a user, I fill a PDF form, add my signature, and save a copy.
- As a user, I add notes and highlights to a lecture PDF.
- As a user, I create a new blank document or spreadsheet from the home screen.

## 4. Feature requirements

Priority: **P0** = must ship in v1, **P1** = should ship, **P2** = later.

### 4.1 File management (all formats)

- P0 — Home screen with recent files and "New document / New spreadsheet".
- P0 — Open via Storage Access Framework picker (local storage, Downloads, Google Drive, other providers).
- P0 — Register as handler for .docx / .xlsx / .pdf MIME types ("Open with", share-to).
- P0 — Save (overwrite, with confirmation for files from other apps) and Save as copy.
- P0 — Share / export via the Android share sheet.
- P1 — Autosave draft + crash recovery.
- P1 — Export DOCX/XLSX to PDF.
- P2 — Search inside the document; file thumbnails.

### 4.2 DOCX (Word documents)

- P0 — View with headings, paragraphs, bold/italic/underline, lists, tables, images.
- P0 — Edit text; basic formatting (bold, italic, underline, heading levels, bullet/numbered lists, alignment).
- P0 — Undo / redo.
- P1 — Insert image, insert/edit simple tables, links, font size and colour.
- P1 — Find and replace; word count.
- P2 — Headers/footers, page breaks, comments (read-only first).
- **Fidelity rule:** if the file contains content the editor can't represent (text boxes, SmartArt, tracked changes), show a warning before editing and offer "Save as copy" so the original is never damaged.

### 4.3 XLSX (Excel spreadsheets)

- P0 — View multiple sheets with tabs; frozen header rows; zoom.
- P0 — Edit cell values; common formulas (SUM, AVERAGE, IF, COUNT, VLOOKUP/XLOOKUP, etc.) recalculate.
- P0 — Insert/delete rows and columns; undo/redo.
- P1 — Basic formatting (bold, fill colour, number formats, column width); sort and filter.
- P1 — Add/rename/delete sheets.
- P2 — Charts (view-only first), conditional formatting, CSV import/export.
- Preserve unsupported features (charts, pivot tables) on save where the library allows; otherwise warn.

### 4.4 PDF

- P0 — Fast, smooth viewing with pinch-zoom, page thumbnails, text search, text selection/copy.
- P0 — Annotate: highlight, underline, freehand ink, text notes/boxes.
- P0 — Fill interactive (AcroForm) forms.
- P0 — Add a signature (draw once, reuse).
- P1 — Page tools: reorder, rotate, delete, merge PDFs, split/extract pages.
- P1 — Add images (e.g. stamp, photo).
- P2 — Images → PDF (scan from camera), compress.
- Out of scope: rewriting the original text of a PDF.

### 4.5 PPTX (PowerPoint) — view only

- P0 — Open .pptx and show slides with text, shapes, images and basic layouts.
- P0 — Swipe between slides; slide thumbnail strip; pinch-zoom.
- P1 — Speaker notes view; full-screen presentation mode.
- P2 — Basic transitions/animations (otherwise shown as static slides).
- No editing. Candidate libraries: a JS PPTX renderer in the WebView (e.g. pptx-preview or PPTXjs — check licence and fidelity), or convert slides to images; evaluate during M4.

## 5. Technical architecture

### 5.1 Approach: native shell + WebView editors (hybrid)

The hard part of an office app is the format handling. The best free, permissively licensed format libraries are JavaScript, so the editors run inside a bundled, offline WebView, while a native Kotlin shell handles everything Android-specific.

- **Native shell (Kotlin + Jetpack Compose):** home screen, recent files, file picker (SAF), intent filters, share sheet, permissions, settings, saving bytes back to the content URI.
- **Editor layer (WebView, local assets only):** one small web app per format, built with Vite and shipped inside the APK (`assets/`), loaded via `WebViewAssetLoader`.
- **Bridge:** Kotlin ↔ JS via `addJavascriptInterface` / `WebMessagePort`. Native passes file bytes (base64 or chunked) to JS; JS returns the edited bytes for native to write.

### 5.2 Library stack

| Area | Library | Licence (verify before release) | Role |
| --- | --- | --- | --- |
| XLSX read/write | SheetJS CE or ExcelJS | Apache-2.0 / MIT | Parse and write workbooks |
| XLSX editor UI | Univer (or Fortune-sheet) | Apache-2.0 / MIT | Spreadsheet grid, formulas |
| DOCX → editor | mammoth.js | BSD-2 | Convert docx to HTML for editing |
| DOCX editor UI | TipTap (ProseMirror) | MIT | Rich-text editing |
| Editor → DOCX | `docx` (npm) | MIT | Generate .docx on save |
| PDF viewing | Android `PdfRenderer` (native) or PDF.js | Platform / Apache-2.0 | Fast page rendering |
| PDF editing | pdf-lib (JS) or PdfBox-Android | MIT / Apache-2.0 | Annotations, forms, pages, signatures |

### 5.3 Alternatives considered

- **Fully native (Apache POI + PdfBox-Android):** no WebView, but POI is large and has no ready-made editing UI — every editor would be built from scratch.
- **All-web (Capacitor or Tauri 2 mobile):** simplest code sharing, but weaker Android integration and file handling.
- **LibreOffice / Collabora Android core:** best fidelity, but very large (100+ MB), complex native build, MPL/LGPL obligations. Kept as a possible future "high-fidelity mode".

### 5.4 Platform

- minSdk 26 (Android 8.0), targetSdk = latest required by Play.
- No `INTERNET` permission in v1 — a strong, verifiable privacy promise.
- Kotlin, Gradle, Jetpack Compose; editors in TypeScript.

## 6. Non-functional requirements

- **Performance:** open a 20-page DOCX or a 5,000-row XLSX in < 2 s on a mid-range phone; PDF pages render progressively, scrolling at 60 fps.
- **Large files:** PDFs up to 200 pages; spreadsheets up to \~50k cells editable — beyond that, open read-only with a notice.
- **Reliability:** never overwrite the original until the new file is fully written and re-parsed successfully (write to temp → validate → replace).
- **Offline & privacy:** no network access, no analytics, no accounts. Files never leave the device unless the user shares them.
- **Size:** APK/AAB under 40 MB.
- **Localization:** English first; Bahasa Indonesia as the second language.

## 7. UX and screens

1. **Home** — recent files list (icon, name, date), big "+" with New Document / New Spreadsheet, "Open file" button.
2. **Document editor (DOCX)** — reading mode by default; tap "Edit" to show a compact formatting toolbar above the keyboard.
3. **Spreadsheet editor (XLSX)** — grid with sheet tabs at the bottom, formula bar at the top, cell editor above the keyboard.
4. **PDF viewer** — full-screen pages, bottom bar with Annotate / Fill & Sign / Pages tools; thumbnail drawer.
5. **Save sheet** — Save, Save as copy, Export as PDF, Share.
6. **Settings** — theme, default save behaviour, signature management, about/licences.

**Principles:** view-first (edit mode on demand, so viewing is fast and accidental edits are rare), one-hand reachable toolbars, and clear warnings before any save that could lose formatting.

## 8. Milestones

| Milestone | Scope | Exit criteria |
| --- | --- | --- |
| M0 — Skeleton | Compose shell, SAF open, intent filters, WebView + bridge passing bytes | A file opens from "Open with" and reaches the viewer |
| M1 — PDF viewer | View only: pages, zoom, thumbnails, search, text copy | 200-page PDF scrolls smoothly; search works |
| M2 — DOCX viewer | View only: text, formatting, lists, tables, images | Test corpus of real Word docs renders readably |
| M3 — XLSX viewer | View only: sheets, tabs, frozen rows, computed formula values | Test corpus of real Excel files displays correct values |
| M4 — PPTX viewer | View only: slides with text, shapes, images; swipe between slides; slide thumbnails | Test corpus of real decks is readable; first usable "viewer" build |
| M5 — PDF editing | Annotations, form fill, signature, page tools | Edited PDFs open correctly in Acrobat / Chrome |
| M6 — XLSX editing | Edit cells, formulas, rows/cols, sheets; save | Round-trip corpus passes; formulas match Excel |
| M7 — DOCX editing | Text and basic formatting edits, fidelity warning; save | Corpus survives open → edit → save |
| M8 — Polish | Tablet/landscape layouts, export to PDF, Indonesian, licence notices | No data-loss bugs on own devices — **passed** (13 Nov 2026) |

**Order rationale:** viewers first. Read-only viewing of all four formats is lower-risk (nothing can be corrupted) and gives a usable app by M4. Editing comes after, starting with the formats whose libraries are most mature (PDF, XLSX); DOCX round-tripping is hardest and goes last. PPTX stays view-only.

## 9. Risks, licensing and open questions

### Risks

| Risk | Impact | Mitigation |
| --- | --- | --- |
| DOCX round-trip loses formatting (mammoth → HTML → docx is lossy) | High | Fidelity warning + Save as copy; later, edit the document XML directly instead of regenerating it |
| WebView memory limits on big files | Medium | Size thresholds → read-only mode; chunked byte transfer |
| Formula engine differences vs Excel | Medium | Test corpus with expected values; show unsupported functions clearly |
| Library licence changes | Medium | Pin versions; keep a licence table; avoid AGPL dependencies |
| Scope creep toward "full office suite" | High | Stick to non-goals; P2 items only after M5 |

### Licensing

- Avoid AGPL / commercial-only components (e.g. some full DOCX/PDF editor SDKs) so the app can stay free and closed or open as you choose.
- Include an in-app "Open-source licences" screen.
- Decide the app's own licence (e.g. GPL-3.0 or Apache-2.0) before publishing on GitHub/F-Droid.

### Open questions

- Personal-use APK only, or public release on Play Store / F-Droid? personal use
- Add optional AI features (summarize, rewrite, generate formulas) via the Claude API? no
- Tablet / landscape layouts in v1 or later? v1
- Support PowerPoint (.pptx) viewing in a later version? Yes — view-only, milestone M4.

## 10. Building it with Claude

1. **Keep this PRD in the repo** (e.g. `docs/PRD.md`) and point Claude Code at it at the start of each session, together with a short `CLAUDE.md` describing the architecture and conventions.
2. **One milestone per session / branch.** Ask Claude Code to plan the milestone first, then implement in small steps.
3. **Tests first for formats.** Build a `test-files/` corpus of real DOCX/XLSX/PDF files and have Claude write round-trip tests (open → save → reopen → compare) before writing editor features.
4. **Separate the two worlds.** Develop the web editors in a browser with hot reload (fast iteration), then wire them into the Android WebView.
5. **Device testing.** Run on a real phone via `adb` for every milestone; check memory and speed on large files.
6. **Review licences** of every dependency Claude adds before merging.
