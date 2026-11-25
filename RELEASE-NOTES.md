# CatPDF v4.4.1
- Replace manually assembled PPTX with locally bundled PptxGenJS 4.0.0.
- First-page aspect ratio; mixed page sizes fitted without stretching.
- Sequential rendering, bounded canvas resolution and image payload; 20 MB input / 50 pages retained.
- Validate core XML parts, slide count and slide image/layout relationships before download.
- Word/Excel to PDF remain text-only; PowerPoint to PDF remains experimental; PDF to Word/Excel remain approximate.
- Existing unified interface and functions retained; cache version bumped.
- Benchmark: seven-page supplied PDF rendered with PyMuPDF and packaged using PptxGenJS in Node; python-pptx and LibreOffice opened it. Browser integration and Microsoft PowerPoint not tested. No full dependency vulnerability audit. GitHub inaccessible from development environment.
