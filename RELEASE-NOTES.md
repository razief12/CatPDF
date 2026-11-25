# 4.5.1 validation build
Changes: Dockerfile COPY paths now match Render repository-root build context; bounded PDF raster dimensions; image decompression warnings rejected; mixed image-only PDF pages rejected for editable reconstruction; Word page dimensions and compact table text; cancellation message; cache version updated.
16 backend smoke tests rerun. These are not 95% fidelity tests. Existing limitations in README remain. No Docker or browser deployment validation. Do not deploy publicly without isolation and rate limiting.
