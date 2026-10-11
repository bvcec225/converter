# Vasanth PDF to Word/Excel Converter

A local Windows application that converts PDFs and document images into Excel workbooks or Word documents, using embedded text when available and English/Telugu OCR when needed.

**Current version:** 3.3.1  
**Platform:** Windows x64  
**Repository:** https://github.com/bvcec225/gittest1  
**Portable package:** `Vasanth PDF to Word and Excel Converter 3.3.1.zip`

The application includes its Python and Node.js runtimes and local OCR models. Conversion does not require an internet connection, API key or subscription. OCR and reconstructed layouts can contain errors; the application provides review tools and a source-appearance Word mode to support verification.

## Quick start

1. Download the portable application ZIP from the repository's Releases section when published.
2. Extract the complete ZIP to a local folder. Do not run the application inside the ZIP.
3. Double-click **Start Vasanth PDF to Word and Excel Converter.cmd**.
4. Select **Choose file...** for one document, or **Choose folder...** for a batch.
5. Select a separate output folder and choose **Excel** or **Word**.
6. Leave extraction on **Auto** and OCR languages on **engbest+telbest** for the initial run.
7. Click **Convert / Resume**.
8. Check the result, any review flags and the **Conversion time** column.

Keep the complete extracted application folder together. Moving or removing its `runtime`, `models` or `node_modules` directories prevents normal operation. The title bar should show **Vasanth** and version **3.3.1**.

The output directory receives converted XLSX/DOCX files. Processing caches, logs and OCR alternatives are stored separately. Original input files are unchanged.

## Supported inputs and conversion modes

| Input or scenario | Processing approach | Practical limitation |
|---|---|---|
| PDF with selectable text | Read embedded text and table geometry | Corrupt font mappings can require Force OCR |
| Scanned PDF | Render pages and recognize printed text locally | Recognition depends on source quality |
| PDF containing native and scanned pages | Choose processing per page | Mixed-content regions can still require review |
| Outlined text PDF | Recognize rendered glyphs when selectable text is absent | Very small text and watermarks can remain difficult |
| JPEG, PNG, BMP and TIFF | Treat images as document pages | Multi-frame TIFF is supported |
| Photographed documents | Attempt page cropping and perspective correction | Curved pages, shadows and unclear edges remain difficult |
| Rotated or slightly skewed scans | Orientation detection and optional deskew | Inspect unusual layouts and failed orientation results |
| Password-protected PDF | Use the password supplied in the interface | The password is not saved |

Table handling includes variable column counts, ruled and borderless layouts, merged cells, wrapped text, repeated headers and multi-page continuations. Support describes the implemented processing paths, rather than a guarantee that every possible PDF layout will be reconstructed correctly.

## Excel conversion

Excel export creates table sheets with compact row heights, wrapped text, styled headers and detected merged cells. Document information is recorded on a separate **Report details** sheet. Titles or header bands that form part of a detected source grid may remain in that table; for example, the three-page validation document has three title/header rows in addition to its 73 data records.

Features include:

- Joining matching continuation tables and optionally removing repeated headers.
- Retaining original Telugu text.
- Preserving leading zeros, long identifiers and formula-like strings as text by default.
- Optional conversion of simple numbers when appropriate for the document.
- Blank-cell preservation and removal of detected table perimeter rules before OCR.
- Batched recognition of cell images by column, with bounded retries for uncertain cells.
- Verification of saved workbook values against the extracted records.
- Splitting oversized content at Excel limits; exceptionally long cell text can use a **Long cell text** sheet.

Saved-value verification checks that the exporter retained the extracted values. It does not prove that an OCR reading matches the original source.

### Review and Telugu name transliteration

Double-click an Excel result to open its review session. **Review cells** displays selected values, source crops and reasons for review. Save corrections and use **Export reviewed Excel** to generate the corrected workbook.

**Page tables** supports inspection of all pages, rotation, retries, manual table boundaries and confirmation of pages with no table.

**English columns...** adds phonetic English transliterations beside selected Telugu source columns while retaining the original text. For the original Adangal workflow, select source name columns **N and O**. Transliteration applies only to the selected columns; it is not a translation of the entire workbook. Preferred English spellings of personal names may need manual correction.

## Word conversion

Select the appropriate **Word layout** before conversion:

| Layout | Result | Best suited to |
|---|---|---|
| Editable text (forms may be images) | Reconstructed text, paragraphs, tables and preserved graphics; difficult regions or pages can remain images | Documents that need text editing |
| Source appearance (page images) | Source-page images embedded in DOCX, bypassing OCR | Preserving the visible wording and layout |

Native PDFs retain much of their font information, alignment, tables, columns and images. Scanned paragraphs use source positions, first-line indentation, line spacing and measured Telugu text widths. Complex-script font size and bold are set explicitly in the Word document.

A scan does not contain reliable font metadata. Font family is therefore approximated, and bold is inferred from image features. OCR can still misread names, amounts, dates and characters. Positioned editable paragraphs also do not reflow through the document like a newly authored Word report. Pagination can differ from the PDF.

Complex forms or uncertain handwriting/stamp regions may be preserved as images. **Source appearance** preserves visible content but does not produce editable text. The Excel cell-correction editor does not edit DOCX text; Word corrections should be made in Word after checking the source.

## Conversion settings and timing

| Setting | Purpose |
|---|---|
| Auto | Use embedded text where usable and OCR where necessary |
| Embedded only | Process available native text without OCR |
| Force OCR | Recognize rendered pages even if native text exists |
| OCR languages | Choose locally installed models; English and Telugu are bundled |
| Resolution | Change render scale; higher values increase work and memory |
| Correct small scan skew | Deskew scanned pages |
| Read individual scan cells | Enable table-cell recognition and bounded retries |
| Enhanced OCR | Compare additional readings for uncertain content |
| Join matching continuations | Combine tables with compatible layouts |
| Include subfolders | Include supported files in nested input directories |

Each file's elapsed conversion time appears in the results table after processing. Failed and stopped files also receive timing information when their processing ends. The final screen displays total batch time. File timings include extraction, OCR and saving, rather than OCR alone.

Use **Stop** to interrupt work, then run **Convert / Resume** with the same input and settings to reuse completed pages. Changed sources or relevant settings can require fresh processing. Version 3.3 invalidated the older table-OCR cache so prior erroneous readings would not be silently reused.

## Development history

| Stage | Requirement or problem | Resulting development |
|---|---|---|
| Initial Adangal workflow | Extract every page, retain Telugu and separate document information | Excel table export, report details and optional name transliteration |
| General converter | Work with local input/output folders and different PDF table layouts | Native/OCR processing, variable columns, batch conversion and review tools |
| Test-pack validation | Handle merged cells, wrapped text, blanks, financial formats, rotation and scans | Scenario-based comparisons and extraction improvements |
| Word extension | Add a front-end conversion choice and support English/Telugu documents | Editable DOCX conversion and source-appearance mode |
| Version 3.1 | Large Excel files stalled during formatting; conversions produced unwanted sidecar files | Removed repeated worksheet scans, reused styles, added bounded worker calls and kept processing files outside the output folder |
| Version 3.2 | Scanned Word paragraphs had inconsistent sizes, bold and wrapping | Positioned paragraphs, source indentation/spacing and Windows Uniscribe measurements for shaped Telugu text |
| Version 3.2.1 | Users needed either a single input file or an entire folder | Separate file/folder selection buttons |
| Version 3.3 | Three-page Excel OCR was slow and produced border artifacts and damaged numbers | Batched cell OCR, conservative reading selection, blank detection, continuation reuse and higher-resolution outlined-text cells |
| Version 3.3 | Users needed a conversion-time summary | Per-file elapsed time and total batch time in the interface |
| Version 3.3.1 | Application naming and appearance | Vasanth branding, navy/teal theme, clearer controls and highlighted Convert button |

Version 3.3.1 changes presentation and branding. Its conversion behavior is unchanged from version 3.3.

## Validation results

Validation references were used only for comparison after conversion. The converter does not receive ground-truth CSV/DOCX content to guess or replace source values.

### Three-page Excel document: version 3.3

- Fresh conversion time: **123.36 seconds**, approximately **2 minutes 3 seconds**, on the development computer.
- Retained coverage: **3 pages, 73 data records and 16 columns**.
- Joined worksheet: **76 physical rows**, including three title/header rows.
- Independently transcribed page-1 sample: **185 of 189 numeric/identifier cells matched**, compared with **137 of 189** in the previous workbook.
- Blank remarks in that sample: **21 of 21 preserved**, compared with zero previously.

The sampled comparison excludes Telugu text, headers and page-2/3 values. It is not a whole-document accuracy percentage. Remaining sampled differences include two survey identifiers and two cells with stray border tokens.

### Fourteen-file Excel regression pack: version 3.3

All 14 files exported with no missing or extra data rows. Nine digital fixtures matched all compared data and header values. Scanned fixtures still contained recognition errors.

| Test scenario | Matched data cells / compared cells | Header mismatches |
|---|---:|---:|
| Clean digital table | 150 / 150 | 0 |
| High-quality scan | 90 / 90 | 0 |
| Skewed/noisy scan | 74 / 75 | 0 |
| Multi-page repeated headers | 360 / 360 | 0 |
| Grouped headers and merged cells | 105 / 105 | 0 |
| Multiline cells | 48 / 48 | 0 |
| Financial/date formats | 84 / 84 | 0 |
| Borderless native table | 75 / 75 | 0 |
| Low-resolution scan | 73 / 80 | 0 |
| Rotated native page | 72 / 72 | 0 |
| Confusable-character scan | 40 / 72 | 0 |
| Blanks, subtotals and totals | 72 / 72 | 0 |
| Telugu land-record scan | 29 / 40 | 1 |
| Long 150-row table | 1500 / 1500 | 0 |

These comparisons normalize Unicode NFC and whitespace, with repeated headers excluded from data scoring. They do not measure exact visual layout or original whitespace preservation.

### Historical large-file and Word checks

Version 3.1 processed a fresh **881-page digital PDF in 964.1 seconds**, approximately **16.07 minutes**, with zero failed pages. It extracted 39,214 source table rows and exported a joined sheet of 38,334 rows after omitting 880 repeated headers. Comparison against the previous extraction found zero non-whitespace differences across 548,996 cells. This measures preservation against extracted records, not source-ground-truth accuracy.

The historical cached export-only run took 69.4 seconds. These are version-3.1 measurements, not fresh version-3.3.1 benchmarks. There is no verified 880-page scanned-document timing guarantee.

Earlier Word validation converted and rendered 38 test fixtures and eight real inputs. Normalized body-text agreement against supplied reference DOCX files ranged from 96.99% to 100%; that metric excluded layout, images, headers/footers and list numbering. Twelve representative Word fixtures were subsequently checked for layout regressions in version 3.2. Real Telugu legal documents and photos had no verified transcription and received no claimed OCR accuracy percentage.

Version 3.3 checked native and rotated-scan Word exports and timing. Version 3.3.1 checked GUI startup, layout, timing-column presence, ZIP integrity and startup from the extracted portable package. The previous legal-document OCR errors are not claimed to be resolved by the interface refresh.

## Performance and known limitations

- OCR confidence is a model score, not verified accuracy.
- Missing, blurred or damaged glyphs cannot reliably be recovered automatically.
- Telugu names, legal wording, amounts, dates and similar O/0, I/1, B/8 characters require source comparison.
- Watermarks, faint print, bleed-through, compression, handwriting and page curvature can reduce recognition quality.
- A detected table may have incorrect boundaries, merged cells or reading order; review tools allow correction.
- There is no fixed application page/column cap, but RAM, storage and Word/Excel format limits apply. DOCX construction holds the document in memory.
- Large scanned files can take much longer than digital PDFs. A universal 15–20 minute conversion target is not established.
- Failed pages block a complete export rather than silently producing a complete-looking partial document.

Faster processing in version 3.3 comes from batching cells, limiting retries, reusing matching continuation orientation and language choices, bypassing expensive native parsing on textless pages, and rendering only the cell source at higher resolution when needed. These changes reduce redundant work; they do not guarantee perfect OCR.

## Cache, privacy and troubleshooting

The default internal cache root remains:

```text
%LOCALAPPDATA%\PDF to Excel and Word Converter\cache
```

The legacy cache-directory name is intentional; the Vasanth branding does not move existing sessions. Cache subdirectories depend on the output folder. Set the environment variable `PDF_CONVERTER_CACHE` to choose a different cache root.

Caches contain source images, extracted values, alternatives and logs. They stay local and support resume/review. Removing them loses that saved processing work. Passwords are not stored, and the conversion pipeline sends no document contents to a server.

| Symptom | Action |
|---|---|
| Application does not start | Extract the whole ZIP and run `Diagnose.cmd` |
| Runtime or model missing | Restore the complete application folder |
| Wrong version opens | Launch the CMD file inside the new Vasanth folder |
| Wrong OCR text or empty cells | Compare source crops, review flagged cells and inspect table boundaries |
| Word appearance differs | Compare Editable text with Source appearance; image mode is not editable |
| Long conversion | Check render scale, source quality, enabled OCR passes and the current activity log |
| Need to interrupt | Use Stop, then resume using the same input and settings |

## Command-line use

Open Command Prompt in the extracted application folder:

```bat
set "PYTHONHOME=%CD%\runtime\python"
set "PYTHONUTF8=1"
set "PYTHONNOUSERSITE=1"

runtime\python\python.exe cli.py "C:\Input" "C:\Excel Output" --format excel
runtime\python\python.exe cli.py "C:\Input\sample.pdf" "C:\Word Output" --format word
```

CLI options include `--recursive`, `--languages engbest+telbest`, `--mode "Force OCR"`, `--mode "Embedded only"`, `--force-export` and `--keep-repeated-headers` for Excel. Use `cli.py --help` for the supported arguments.

## Application structure

| Component | Responsibility |
|---|---|
| `app.py` | Tkinter interface, file/folder selection, review screens and timing display |
| `branding.py` | Vasanth name, version and visual theme |
| `converter.py` | PDF page extraction, processing choices, continuation handling and batch conversion |
| `detection.py`, `header_groups.py`, `borderless_headers.py` | Table geometry and header handling |
| `engine.py`, `ocr.cjs`, `ocr-worker.cjs` | Local OCR worker, cancellation, caching and word coordinates |
| `cell_ocr.py`, `enhanced_ocr.py` | Batched table-cell recognition and alternative readings |
| `exporter.py`, `transliterate.py` | XLSX export, value verification and selected English-name columns |
| `word_converter.py`, `word_structure.py`, `word_graphics.py` | DOCX construction, tables, document furniture and graphics |
| `scan_layout.py`, `font_metrics.py`, `paragraph_refinement.py` | Scanned paragraph placement, text measurements and layout refinement |
| `native_refinements.py` | Native Word conversion refinements |
| `cli.py`, `diagnose.py` | Command-line conversion and startup checks |
| `runtime/`, `models/`, `node_modules/` | Bundled runtimes, OCR models and dependencies |
| `TESTED.md`, `THIRD-PARTY-NOTICES.txt` | Validation history and dependency notices |

The portable package includes application source modules. Development validation used separate fixtures and comparison/rendering tools; those test datasets are not part of the portable application.

## GitHub distribution

The portable ZIP is **197,615,312 bytes**, approximately **188.46 MiB**. It exceeds GitHub's 100 MiB ordinary-file limit, so distribute it as a **Release asset**, rather than committing the ZIP as a normal Git file. See [GitHub large-file guidance](https://docs.github.com/en/repositories/working-with-files/managing-large-files/about-large-files-on-github) and [release documentation](https://docs.github.com/en/repositories/releasing-projects-on-github/about-releases).

Suggested publication:

1. Add this README to the repository, or to a dedicated application subdirectory if the repository also contains unrelated projects.
2. Open the repository's Releases section and create a release with tag **v3.3.1**.
3. Use the title **Vasanth PDF to Word/Excel Converter 3.3.1**.
4. Attach `Vasanth PDF to Word and Excel Converter 3.3.1.zip` as a release asset.
5. Include installation instructions and the known OCR limitations in the release description.
6. Publish the release and verify its download link.

Suggested release notes:

> Portable Windows converter for PDF and image inputs, with Excel and Word output and local English/Telugu OCR. Version 3.3.1 introduces Vasanth branding and an updated interface. Includes single-file/folder selection, per-file timing, batch totals and the version-3.3 batched Excel OCR pipeline. Extract the full ZIP and run the Vasanth launcher. OCR-derived text and reconstructed layouts still require review; Source appearance Word preserves pages as non-editable images.

SHA-256 of the final version-3.3.1 portable ZIP:

```text
cce435761d22888907c105a3685e734e3471c963b63902119ea103fa5444048b
```

As of preparation of this README, the ZIP has been created and verified locally. GitHub publication has not been completed; repository access/sign-in is still required. Remove or update this status statement after publication.

## Dependencies and licensing

The bundle includes Python 3.12.14, Node.js 24.19.0, Tesseract.js 7.0.0, English/Telugu models, OpenCV, PDFium/pypdfium2, pdfplumber, openpyxl, python-docx, pdf2docx and PyMuPDF, among other components. Exact notices and original dependency license files are retained in the package.

Read **THIRD-PARTY-NOTICES.txt** and the bundled licenses before redistribution. PyMuPDF is listed under AGPL-3.0 in the supplied notices. This README does not assign an overall license to the application's own source code; that project-level licensing decision remains to be specified.

## Future improvement priorities

These are proposed next steps, not completed features:

1. Improve Telugu recognition and test additional local recognizers against independent references.
2. Reduce remaining identifier and border-token errors without guessing missing content.
3. Expand low-resolution and photographed-document validation.
4. Improve editable Word reconstruction for complex forms and multi-column scans.
5. Benchmark large scanned workloads and resource use on different computers.
6. Add reproducible automated validation and a documented release process.

For reproducible issues, record the application version, settings, elapsed time and affected page/cell, and provide a representative source/output pair through an appropriate private channel if the document contains personal information.
