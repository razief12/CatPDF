# CatPDF 4.3.1 — Stability & Bug Fix
Based on the supplied 4.3 ZIP. All 44 unique tool cards retain matching panels; legacy hidden watermark panel retained. No API translation added.
- PDF.js moved from 3.11.174 to 4.10.38 ESM with matching worker, eval and XFA disabled, image size limit and failed loading-task cleanup. CDN unavailable in build environment; integration and current dependency security not verified.
- Versioned scripts and service worker; local network-first, offline fallback, core-only install. Clear Cache removes CatPDF service worker and app caches, preserves preferences and AI model cache. Confirm before reload.
- Diagnostic log records bounded session-only generic events, tool identifiers and timestamps, no filenames or document text.
- Translation worker terminated after AI tasks to release memory; model may need to reload from cache. Advanced compression canvas cleanup on failure. Preview cap 500 pages; PPTX image budget 60 MB; Office archive entry limit 3000.
- Excel import preserves gaps between referenced columns (max 256). Office conversions remain text/image based. PPTX slide master/layout compatibility remains unresolved; no Office runtime verification.
- Static audit: 44 unique cards and matching panels; all inline and external JS syntax checked. No end-to-end conversion, mobile, offline, CDN or memory benchmark performed. Translation speed and API integration deferred. Full security audit not complete.
