# cleaning_inspection_app

Mobile-friendly inspection form for the Saclay cleaning verification sheet, with an experimental form generator for structured `.docx` templates and fillable `.pdf` files.

## What It Does

- Recreates the verification form in a phone-friendly web app
- Exports a Word-compatible `.docx` document based on the original template
- Saves the current draft automatically in the browser on the same device
- Adds photos to generated inspection reports
- Imports previous `.docx` inspection sheets into a local history
- Stores archived output files locally in the browser
- Exports archived files individually or as a `.zip`
- Builds statistics from archived inspections, including per-section percentages, building lists, and grouped remarks
- Generates a dynamic form app from another structured `.docx` template
- Generates a dynamic form app from a fillable `.pdf` form and exports a filled `.pdf`
- Archives generated forms and summarizes their detected fields in statistics

## Architecture

See [`docs/architecture.md`](docs/architecture.md) for the current module layout and data flow.

High-level structure:

- `src/App.jsx`: React screens, fixed Saclay workflow, generic generator tab, archive/statistics UI
- `src/constants.js`: shared constants, fixed form definitions, template and storage keys
- `src/genericDocx.js`: generic `.docx` template inspection and filled `.docx` export
- `src/genericPdf.js`: fillable `.pdf` field inspection and filled `.pdf` export
- `src/archiveStorage.js`: IndexedDB persistence for archived output files
- `src/statistics.js`: Saclay and generated-form archive aggregation/statistics snapshots
- `src/App.css` and `src/index.css`: app, document, print, history, statistics, photo, and generator styles
- `public/`: browser-served image assets

## Generator Support

The `Generateur` tab accepts:

- `.docx` files that use tables, empty cells for answers, and Word checkbox controls
- fillable `.pdf` files that contain real PDF form fields
- flat text-based `.pdf` files, with inferred text fields and a filled response page appended to the PDF

Current limits:

- Scanned image-only PDFs still require OCR/layout detection before fields can be inferred.
- Flat PDF support is heuristic. It works best when labels are stored as selectable PDF text.
- Generic `.docx` support is heuristic. It works best on structured table-based forms.
- The Saclay-specific inspection flow remains available as the main `Inspection` tab.

## Run Locally

Install dependencies:

```bash
npm install
```

Start the dev server:

```bash
npm run dev
```

If you want to open it from your iPhone while your Mac is nearby:

```bash
npm run dev -- --host
```

Then open the `Network` URL shown by Vite in Safari on your iPhone.

## Deploy With GitHub And Vercel

### 1. Push to GitHub

If this project is not in a Git repo yet:

```bash
git init
git add .
git commit -m "Prepare cleaning inspection app for deployment"
```

Create a new empty GitHub repository, then connect and push:

```bash
git remote add origin <your-github-repo-url>
git branch -M main
git push -u origin main
```

### 2. Deploy on Vercel

1. Go to `https://vercel.com`
2. Sign in with your GitHub account
3. Click `Add New...` then `Project`
4. Import this repository
5. Keep the default Vite settings

Use these values if Vercel asks:

- Build command: `npm run build`
- Output directory: `dist`

6. Click `Deploy`

After deployment, Vercel will give you a public URL that works on your iPhone anywhere.

## iPhone Usage Notes

- The form draft is saved in the browser on that same phone
- Archived reports and imported files are also stored locally on that device
- If you clear Safari website data, the saved draft will be removed
- Export the document before leaving the site if the data is important
- You can use `Add to Home Screen` in Safari to make it feel more like an app

## Statistics

The `Statistiques` tab summarizes archived inspections:

- Section `0`: buildings where ATALIAN workers were present
- Sections `1` to `3`: percentages for `OUI`, `NON`, and `Sans objet` where available
- Section `4`: item presence percentage and material condition breakdown
- Section `5`: quality distribution percentages
- Section `6`: improvement-point percentages
- Remarks grouped by similar meaning, even when the wording differs

The statistics view can also export a `.docx` report with the same consolidated information.

Generated forms are summarized separately by source template:

- number of archived generated files
- detected sections and fields
- completion percentage per field
- option distribution for choices and checkboxes
- most common values for text fields

## Dependencies

The app uses:

- `react` and `react-dom` for the UI
- `vite` for local development and production builds
- `jszip` for `.docx` archive/template manipulation
- `docx` for generated statistics reports
- `pdf-lib` for fillable PDF parsing and export

After dependency changes, run:

```bash
npm install
```

This refreshes `package-lock.json`.

## Production Recommendation

This app is ready for simple phone use, but the saved draft is only local browser storage.

If you want safer real-world usage later, the next step should be one of:

- sync forms to a database
- email each completed report automatically
- upload exported documents to cloud storage
