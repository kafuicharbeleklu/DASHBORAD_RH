# Repository Guidelines

## Project Structure & Module Organization
This repository is a Power BI Project (`.pbip`) centered on `PRESENTATION.pbip`. Report layout lives under `PRESENTATION.Report`, especially `definition/pages/*/page.json` and `visuals/*/visual.json`. The semantic model lives under `PRESENTATION.SemanticModel`, with shared model metadata in `definition/model.tmdl`, relationships in `definition/relationships.tmdl`, and one table per `definition/tables/*.tmdl`. Source inputs currently include the Excel workbook and PDF at the repository root.

## Build, Test, and Development Commands
There is no scripted build or automated test runner in this snapshot. Use the project through Power BI Desktop:

- `start PRESENTATION.pbip` opens the report and linked semantic model locally.
- Refresh data in Power BI Desktop after model changes to validate measures and visuals.
- `rg --files PRESENTATION.SemanticModel\\definition` lists TMDL objects quickly during review.
- `rg "displayName|measure" PRESENTATION.Report PRESENTATION.SemanticModel` helps trace renamed pages, visuals, or measures.

## Coding Style & Naming Conventions
Preserve the existing file formats: JSON for report metadata and TMDL for the semantic model. Use 2-space indentation in JSON and tabs/Power BI-generated layout in `.tmdl` files; avoid reformatting unrelated blocks. Keep business-facing names readable and consistent with current patterns such as `card_*`, `chart_*`, `nav_*` for visuals and title-case measure names like `Absence Rate Latest`. Do not rename generated `LocalDateTable_*` objects or change `lineageTag` values unless the change is intentional and validated.

## Testing Guidelines
Validation is manual. After each change, open the PBIP, refresh the dataset, and check affected pages such as `HR Performance` and TCDP views for broken visuals, filter behavior, and measure totals. If you add or edit DAX, verify both the card visual and at least one chart using that measure.

## Commit & Pull Request Guidelines
No Git history is bundled in this workspace, so no established commit convention can be inferred here. Use short, imperative commit messages, ideally scoped by area, for example: `model: adjust absence measures` or `report: rename HR page cards`. PRs should summarize affected pages/tables, note any source file updates, and include screenshots for visual changes.

## Security & Configuration Tips
Do not commit local Power BI cache artifacts; `.gitignore` already excludes `.pbi/localSettings.json` and `.pbi/cache.abf`. Treat the root Excel and PDF files as source data artifacts and replace them only when the data refresh baseline is meant to change.
