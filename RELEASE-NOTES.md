# CatPDF v4.5.6 RC4.1 (4.5.6.4.1)
- Restore RC3 Word conversion: no paragraph compaction or blank-page removal.
- Automatic Excel layouts: improve recognized structural schedule and quantity table widths, wrapping, repeated headers and summary pagination.
- Optional RM column formatting (unchecked by default); existing explicit print settings and non-General number formats retained.
- Office and HTML frontend checks backend version before conversion. Cache version updated.
- Local tests: Word sample 7 pages including blank page 2; Excel v4 9 pages, minimum font 8.21 pt, summary together, optional total 698,355.62; Excel v5 9 pages, minimum 7.91 pt, summary together. Workbook cell values and formulas unchanged in tested samples. No deployed browser test or full regression.
