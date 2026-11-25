# CatPDF 4.5.3.1 Navigation Hotfix

- Centralized tool navigation after layout updates.
- Group selector moves into the active panel so it remains accessible.
- Removed duplicate Quick Tools scroll.
- Updated PWA cache version. Backend unchanged.

# CatPDF 4.5.3 preview
Backend: strict line table detection with fallback, approximate cell styles and widths, horizontal merges, numeric types in recognized quantity columns, structured text extraction, SVG-based PDF HTML and legacy HTML positioning repair. Frontend: automatic image page orientation; version/cache update. Existing backend URL retained. SVG HTML prioritizes appearance, not semantic editable tables. Legacy HTML cannot restore missing original backgrounds. Office rendering unchanged. No browser, Docker, security or deployment validation. Smoke tests are not fidelity tests. Requires backend redeployment and frontend update. Do not deploy publicly without isolation and rate limiting.
