# GEMINI.md

This project is a Power BI development repository (`.pbip`) focused on HR data analytics.

## Directory Overview
This directory contains the source code, semantic model definitions (TMDL), and report layouts for the NEEMBA HR Dashboard.

*   `PRESENTATION.pbip`: Main Power BI Project file.
*   `PRESENTATION.Report/`: Contains report pages, visual definitions, and static resources.
*   `PRESENTATION.SemanticModel/`: Contains the semantic model, including tables (TMDL), relationships, and measures.
*   `simulation_data/`: Contains Excel files used as data sources for the simulation.
*   `backup/`: Contains historical backups of the project state.

## Key Files
*   `PRESENTATION.SemanticModel/definition/model.tmdl`: Main semantic model file.
*   `PRESENTATION.SemanticModel/definition/tables/*.tmdl`: Individual table definitions.
*   `PRESENTATION.Report/definition/pages/*/pages.json`: Individual report page layouts.

## Usage
This project is intended to be opened and managed using **Power BI Desktop**.
- **Opening:** Execute `start PRESENTATION.pbip` to open the full project in Power BI Desktop.
- **Data Refresh:** Use the "Refresh" function in Power BI Desktop to update data from the source Excel files.
- **Development:** Edit semantic model definitions in `.tmdl` files and report layouts in JSON/Power BI Desktop.
- **Validation:** Changes must be manually validated within Power BI Desktop by verifying report visuals and data integrity after updates.

## Guidelines
- **Project Structure:** Adhere to the established organization in `AGENTS.md`.
- **Formatting:** Use 2-space indentation for JSON and preserve the existing structure of `.tmdl` files.
- **Naming:** Follow existing naming conventions (`card_*`, `chart_*`, etc.).
- **Security:** Do not commit local configuration or cache files (`.pbi/localSettings.json`, `.pbi/cache.abf`).
