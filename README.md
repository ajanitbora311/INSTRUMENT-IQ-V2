# Instrumentation Estimator

The fixed catalog was built from the supplied **Rate Analysis of Inst Misc SOR Jobs.xls** and five **Material PO Report** exports from 2022–2026. It contains 48 rate-analysis entries and 6,108 material PO items. The landing page accepts one Excel take-off and provides a dashboard and Excel download after generation. There is no user-facing source picker.

## Run this package locally

Extract all files together. From the extracted folder run:

```sh
python -m http.server 8000 --bind 127.0.0.1
```

Open http://localhost:8000. Use **Download blank template** for the five input columns. The app reads `catalog.json` from the same folder and processes the estimate in the browser.

## Deploy the GitHub package

**This package includes NRL rate data.** Upload it only to an approved private repository and an access-controlled deployment. A private GitHub repository alone does not guarantee a private GitHub Pages site. If private Pages is unavailable for your account or organisation, use an approved internal host or run locally. Anyone with access to the website can retrieve the published catalog.

1. Extract the package and place its contents in the same directory as your existing Pages `index.html`. Commit the files to the branch configured under **Settings → Pages**.
2. Remove the previous `catalog-builder.html`, `catalog-builder.js`, and `catalog-builder.css` from the published folder if you previously uploaded them.
3. After Pages redeploys, open **Visit site**, upload your take-off, and select **Generate estimate**. Use **View dashboard** to see the rate, amount, source filename, sheet, row, reference, match method and any conversion. **Download Excel** exports the same estimate.

## Matching and source rules

The matcher first checks exact descriptions and then known equivalent instrumentation terms such as “JB” and “junction box.” It preserves specifications such as sizes, ratings and materials. Related records are ranked using weighted concept similarity and shown separately for review. Similarity alone does not assign a rate. A unique defensible match and compatible units are required for pricing.

The supplied rate-analysis workbook is the only pricing source in this package. The five material PO reports provide supporting review evidence because they contain material prices and lack service-line references. The rate analysis appears under the other source column in Excel and retains the workbook serial and source row. Review rows remain unpriced and outside totals.
