# CatPDF 4.3.1 — Stability & Bug Fix
Based on the supplied 4.3 ZIP. All 44 unique tool cards retain matching panels; legacy hidden watermark panel retained. No API translation added.
- PDF.js moved from 3.11.174 to 4.10.38 ESM with matching worker, eval and XFA disabled, image size limit and failed loading-task cleanup. CDN unavailable in build environment; integration and current dependency security not verified.
- Versioned scripts and service worker; local network-first, offline fallback, core-only install. Clear Cache removes CatPDF service worker and app caches, preserves preferences and AI model cache. Confirm before reload.
- Diagnostic log records bounded session-only generic events, tool identifiers and timestamps, no filenames or document text.
- Translation worker terminated after AI tasks to release memory; model may need to reload from cache. Advanced compression canvas cleanup on failure. Preview cap 500 pages; PPTX image budget 60 MB; Office archive entry limit 3000.
- Excel import preserves gaps between referenced columns (max 256). Office conversions remain text/image based. PPTX slide master/layout compatibility remains unresolved; no Office runtime verification.
- Static audit: 44 unique cards and matching panels; all inline and external JS syntax checked. No end-to-end conversion, mobile, offline, CDN or memory benchmark performed. Translation speed and API integration deferred. Full security audit not complete.

## v4.3.2 experimental Office rendering
- PPTX to PDF now renders slide canvases in presentation relationship order, including hidden slides, with original page dimensions. Supports basic shapes, text runs, raster pictures, tables and grouped transforms. Output is rasterized, not selectable text.
- SmartArt, charts, complex shapes, theme inheritance, fonts, and master/layout fidelity remain incomplete. Word and Excel conversion are unchanged. API translation remains deferred.
- Chromium benchmark using supplied 45-slide deck and mocked PDFLib output: 45 page calls, 720 x 540 pt, with warnings for missing geometry, SmartArt/chart and approximate shapes. Actual PDF encoding and visual fidelity not verified; CDN unavailable. No benchmark user files included in release.
