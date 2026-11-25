# CatPDF v4.3 — development build

- Adds ten client-side conversion tools; 44 unique cards and matching panels.
- New conversions limit input to 20 MiB, PDF input to 50 pages, Office archive expansion to 60 MB (25 MB per entry). Archive size check uses JSZip internal metadata; not a complete hostile-file sandbox.
- Office conversion is basic: text-only Word/Excel/PowerPoint to PDF (rasterized), approximate PDF table extraction, image-based PDF to PowerPoint. HTML input is parsed in an inert template and never inserted into the live page. HTML output is escaped. No macros or formulas are executed. Legacy Office formats are unsupported. Scanned PDF requires OCR separately for text extraction.
- No API keys added. Existing PDF.js eval and XFA protections retained. PDF.js 3.11.174 remains outdated; dependency vulnerability audit and upgrade pending. Existing CDN dependencies remain trusted; offline and security are not certified.
- Deferred: NLLB model download latency, translation accuracy, RAM benchmarks, API/proxy integration, cancellation behavior, complete browser regression tests. NLLB licensing must be reviewed before commercial use.
- Validation: Node syntax checks for inline scripts and JS files; 44 unique card-to-panel mappings; ZIP integrity. No browser or Microsoft Office compatibility tests performed. PPTX is a minimal OOXML package and may require compatibility fixes. XLSX extraction does not guarantee column alignment. Original 34 tools retained in source, not regression-tested.
- Deploy all files except this optional notes file to the existing GitHub Pages root. Keep previous release for rollback.
