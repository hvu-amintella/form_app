# Dev History

Date: 2026-04-30

## Scope

This update documents the architecture after adding the generic form generator for structured `.docx` templates and fillable `.pdf` files, then generalizing statistics for generated forms.

## Completed Work

### 1. Documentation architecture refresh

- Updated `README.md` with the current project shape.
- Added `docs/architecture.md` with repo-relative module paths and data flows.
- Documented the split between the fixed Saclay inspection workflow and the generic generator workflow.
- Documented current generator limits for heuristic `.docx` parsing and fillable-only `.pdf` parsing.
- Documented generated-form archive and statistics flow.

### 2. Path portability

- Kept documentation paths relative to the repository root.
- Avoided local machine paths in the architecture docs.

### 3. Generalized statistics

- Added generated-form archive records after DOCX/PDF generator export.
- Stored generated schema and entered values in archive metadata.
- Included generated-form output files in local archive storage.
- Extended statistics snapshots with per-template generated-form statistics.
- Kept Saclay-specific statistics isolated to Saclay archive records.

Date: 2026-04-24

## Scope

This update captures the statistics refactor, clearer percentage displays, smarter remarks grouping, and the extraction of shared logic out of `src/App.jsx`.

## Completed Work

### 1. App structure cleanup

- Split shared constants into `src/constants.js`.
- Moved IndexedDB archive file persistence into `src/archiveStorage.js`.
- Moved archive statistics aggregation into `src/statistics.js`.
- Removed unused helpers that were causing lint failures.

### 2. Statistics percentage expansion

- Added visible percentage chips for sections `1` to `3`.
- Section `1` percentages are calculated only for archives where `Intervenant(s) ATALIAN present` is `Oui`.
- Added percentages for `OUI`, `NON`, and `Sans objet` where the section supports those values.
- Added section `4` item presence percentages and material-condition breakdowns:
  - `Neuf`
  - `Etat d usage`
  - `Vetuste`
- Added section `5` quality percentages.
- Added section `6` improvement-point percentages.
- Added a section before section `1` that lists buildings where ATALIAN workers were present.

### 3. Statistics readability

- Reworked percentage display from long inline text into metric chips.
- Kept count and denominator visible next to the percentage.
- Updated section `5` bars to use real percentages.
- Removed misleading minimum bar width for zero-percent values.

### 4. Remarks grouping

- Reworked remarks consolidation so different formulations can be grouped when they describe the same problem.
- Added theme-based grouping for common issues such as:
  - sanitary/WC problems
  - bins and waste
  - floors and traces
  - dust and surfaces
  - glass cleaning
  - corridors
  - cafeteria area
  - consumables
  - odors
  - scale removal/calcium deposits
  - cleaning equipment
- Added token-similarity grouping for remarks that share the same meaning but do not match a predefined theme.
- Displayed example formulations in the UI and statistics export so grouped remarks remain understandable.

### 5. Statistics export alignment

- Updated the generated statistics `.docx` report to include the same new percentages and grouped remarks shown in the app.
- Kept building lists in the export for traceability.

### 6. Documentation

- Updated `README.md` with current export, archive, statistics, and local-storage behavior.
- Added this development history entry.

### 7. Validation

- Ran `npm run lint` successfully.
- Ran `npm run build` successfully.

Date: 2026-04-22

## Scope

This update captures the latest archive-download improvements and the simplification of the statistics experience.

## Completed Work

### 1. Archive ZIP download

- Added a new action in `Historique` to download all archived `.docx` files as one `.zip`.
- Reused the locally stored archive files instead of regenerating documents.
- Kept filenames collision-safe inside the generated ZIP.

### 2. Statistics page simplification

- Simplified the `Statistiques` page so it focuses on archive building lists instead of the previous dashboard-style overview.
- Removed the denser ranking and risk panels from the main statistics experience.
- Kept the archive summary counters:
  - total archives
  - imported `.docx`
  - app-created records
  - covered buildings versus target

### 3. Item-by-item archive building lists

- Reworked the statistics snapshot logic to group archived buildings directly by item.
- Added the following archive views:
  - Sections `1` to `3`: buildings where the answer is `NON`
  - Section `4`: buildings where the equipment state is `Etat d usage` or `Vetuste`
  - Section `5`: buildings grouped by quality level
  - Section `6`: buildings grouped by selected improvement point

### 4. Section 5 quality visualization

- Added a clearer visualization for section `5` quality results:
  - `Satisfaisant`
  - `Acceptable`
  - `Remise en etat a prevoir`
- Displayed the number of buildings for each quality level.
- Kept the matching building list visible for each quality level.
- Added a dedicated details list for `Remise en etat a prevoir` that also shows archived remarks for each building.

### 5. Remarks consolidation

- Added a new remarks section in the statistics page.
- Combined archived remark sources from:
  - quality comments
  - improvement comments
  - other remarks
- Grouped buildings by similar normalized remarks so repeated observations are easier to review together.

### 6. Statistics export alignment

- Updated the statistics `.docx` export to match the new simplified statistics structure.
- Export now includes:
  - section-by-section building lists
  - section `5` quality counts and building lists
  - `Remise en etat a prevoir` details with remarks
  - grouped remarks with matching buildings

### 7. Validation

- Ran `npm run build` successfully after the latest changes.
- Confirmed the existing lint issues are still limited to older unused helpers already present in `src/App.jsx`.

### 8. Historique import merge

- Extended the `Historique` import flow so one selection can include:
  - one verification `.docx`
  - one or more `.jpg` / `.png` photos
- Added a merged import path that:
  - parses the imported verification file
  - rebuilds the document with the same photo appendix format as the app-generated export
  - stores the merged result as a single archived `.docx`
- Kept the original plain `.docx` import behavior when no photos are included.
- Added validation so photo-assisted import requires exactly one `.docx` in the selected batch.

Date: 2026-04-20

## Scope

This file tracks the main development work completed for the cleaning inspection app during the current collaboration session.

## Completed Work

### 1. Statistics dashboard expansion

- Added statistics by section for sections `1` to `5` in the statistics view.
- Kept the new section breakdown aligned with the current inspection form structure:
  - Section `1`: Formations et habilitations
  - Section `2`: Equipements des intervenants
  - Section `3`: Produits de nettoyage
  - Section `4`: Materiels et documents
  - Section `5`: Qualite de la prestation

### 2. New section-level metrics

- Added per-section calculation logic in `src/App.jsx`.
- Each section now shows:
  - completion percentage
  - answered items count
  - `OUI` count
  - `NON` count
  - `Sans objet` count
- Section `5` also shows:
  - selected quality level
  - number of checked improvement points

### 3. Statistics export update

- Extended the statistics `.docx` export so the exported report also includes the new section-by-section summary.
- Kept the exported content consistent with what is visible in the app UI.

### 4. UI readability improvements

- Added dedicated card styles for the new section statistics in `src/App.css`.
- Adjusted the layout so the new statistics remain readable on smaller screens.

### 5. Validation

- Ran a production build with `npm run build`.
- Build completed successfully after the changes.

### 6. Statistics MVP implementation

- Reorganized the statistics tab into two clear groups:
  - current inspection
  - consolidated archive insights
- Kept the recommended MVP statistics visible in the UI:
  - current progress
  - current `OUI` / `NON` / `Sans objet`
  - current quality
  - current section `1` to `5` breakdown
  - total archives
  - buildings covered versus target
  - quality distribution
  - top improvement points
  - top remarks
  - top buildings

### 7. Must-have statistics implementation

- Added archive-wide section performance statistics for sections `1` to `5`.
- Added a ranked list of the points most frequently marked `NON` across archived inspections.
- Extended the statistics `.docx` export so these new must-have insights are included in the exported report as well.

### 8. Statistics readability pass

- Made the statistics page easier to scan by separating:
  - current inspection indicators
  - archived trends
- Added top-level overview cards so the user can immediately see:
  - the main issue to look at first
  - the highest-risk section
  - overall archive coverage
- Renamed section labels to make the page flow more intuitive for a non-technical reader.

### 9. Dashboard-style statistics pass

- Simplified the statistics page to rely more on visual blocks and less on long text.
- Added:
  - stacked bars for response distribution
  - stacked bars for archive source distribution
  - progress bars for current section completion
  - risk bars for archived section performance
  - bar charts for top `NON` items
  - bar charts for most frequent buildings
- Kept the statistics export available while making the in-app dashboard more visual and faster to read.

## Main Files Updated

- `src/App.jsx`
- `src/App.css`
- `dev-history.md`

## Current Local State

The following files currently show local modifications in the project:

- `package.json`
- `package-lock.json`
- `src/App.jsx`
- `src/App.css`

## Statistics Features Added

### Current Inspection Dashboard

- Overview of the current inspection draft.
- Progress percentage for the current form.
- Photo count for the current form.
- Current quality value.
- Current response distribution:
  - `OUI`
  - `NON`
  - `Sans objet`
- Stacked bar chart for current response distribution.
- Section-by-section completion view for sections `1` to `5`.
- Progress bars for current section completion.
- Section detail line showing:
  - answered count
  - total count
  - `NON` count for sections `1` to `4`
  - selected quality for section `5`

### Archive Dashboard

- Total archived inspections.
- Count of imported `.docx` inspections.
- Count of app-created inspections.
- Buildings covered versus total target.
- Overview card for:
  - top failure item
  - highest-risk section
  - archive coverage
- Stacked bar chart for archive source distribution:
  - app
  - imported Word files
- Archive-wide risk view by section `1` to `5`.
- Risk bars for archived section performance.
- Archive section detail line showing:
  - `NON` count
  - answered count
  - archive `NON` rate
- Quality distribution chart for archived inspections.
- Top improvement points chart.
- Top `NON` items chart.
- Top buildings chart.
- Top remarks list.

### Export

- Statistics `.docx` export from the statistics tab.
- Export includes:
  - current section summary
  - archive section summary
  - top `NON` items
  - quality summary
  - improvement summary
  - remarks summary
  - buildings summary

## What Can Be Improved

- Too many panels are still visible at once; the page could be more focused with clearer grouping or progressive disclosure.
- The page still mixes dashboard blocks and report-style blocks; some sections remain too text-heavy.
- Colors currently help, but severity thresholds are not yet explicit.
- Archive trends are static snapshots only; there is no time axis yet.
- Building insights exist, but there is no real building comparison dashboard yet.
- Section `5` is less visual than sections `1` to `4` because it is based on quality and improvement choices rather than `OUI` / `NON`.
- Top remarks are still displayed as text pills instead of a more visual frequency chart.
- There is no filtering yet by:
  - date range
  - building
  - inspector
  - source type
- There is no drill-down from a chart into the matching archived inspections.
- Mobile readability is improved, but some chart-heavy areas may still feel dense on small screens.

## Todo

### High Priority

- Add date filters and trend charts by week and month.
- Add building-by-building comparison with clearer visual ranking.
- Add severity thresholds for archive risk bars and top `NON` items.
- Add a simpler top-level dashboard mode with only the most important charts.
- Reduce repeated explanatory text and rely more on labels and legends.

### Medium Priority

- Add filters for building, inspector, source type, and date range.
- Add clickable drill-down from charts to matching archived inspections.
- Turn top remarks into a visual frequency chart.
- Improve section `5` visualization so quality and improvement points are easier to compare over time.
- Add a more visual summary of archive coverage by building.

### Nice To Have

- Add mini trend sparklines in overview cards.
- Add benchmarking or target thresholds such as acceptable `NON` rate.
- Add export options beyond `.docx`, such as PDF or image snapshot.
- Add dashboard presets:
  - supervisor view
  - quality manager view
  - audit view
