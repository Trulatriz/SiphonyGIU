# PressTech – Siphony GUI Help

This document explains how to work with the application: folder layout, required file formats, and how to use every module from the main dashboard. Follow it to avoid missing data or validation errors.

## 1. Workflow Overview
- Everything is driven from the GUI; you do not need to edit JSON files manually.
- Each paper has a clean tree: per-module **Input** and **Output** folders.
- Typical order of operations:
  1. **Data Extraction** (collect and consolidate raw measurements).
  2. **Data Analysis** (publication plots and exploratory heatmaps).
  3. **Organization** (keep papers and foam types tidy).

The main window is split accordingly:
- **DATA EXTRACTION**
  - ⚡ **SMART COMBINE** – merges DoE, density, PDR, DSC, SEM, and OC outputs into `All_Results.xlsx`.
  - 🔬 **FOAM-SPECIFIC ANALYSIS** – quick access to polymer-specific tools (PDR, OC/Pycnometry, DSC, SEM editor).
- **DATA ANALYSIS**
  - 📈 **SCATTER PLOTS** – publication plots with constancy-rule filters, groups, and error bars.
  - 🔥 **HEATMAPS** – Spearman/Pearson/distance-correlation matrices with column multi-select.
- **ORGANIZATION**
  - 📁 **MANAGE PAPERS** – add/remove papers, change root directories, relocate paths.
  - 🧶 **MANAGE FOAMS** – add/delete foam types globally or per paper.

## 2. Paper Folder Structure
When you create a paper via **Manage Papers > New Paper**, the application builds:
```
Paper/
├─ DoE.xlsx
├─ Density.xlsx
├─ Results/
└─ <FoamType>/
   ├─ PDR/ Input, Output
   ├─ Open-cell content/ Input, Output
   ├─ DSC/ Input, Output
   └─ SEM/ Input, Output
```
Guidelines:
- Place raw inputs in the corresponding `Input` folder; the GUI writes outputs beside them.
- Do not rename subfolders created by the wizard; modules assume these names.

## 3. Managing Papers & Foams
- **New paper**: File → *Select Paper…* → *New Paper*. Choose the name and associated foams.
- **Switch paper**: File → *Select Paper…*.
- **Switch foam type**: File → *Select Foam Type…*.
- **Update foams for a paper**: use *Manage Foams* from the selector.

## 4. Column & File Requirements
Respect headings exactly (including case, units, spaces). Key expectations:

### DoE.xlsx (paper level)
- One workbook with one sheet per foam type (sheet name = foam type).
- Required columns: `Label`, `m(g)`, `Water (g)`, `T (°C)`, `P CO2 (bar)`, `t (min)`.

### Density.xlsx (paper level)
- One sheet per foam type.
- Recommended columns (used by Combine/OC modules):
  - `Label`, `Av Exp ρ foam (g/cm3)`, `Desvest Exp ρ foam (g/cm3)`, `%DER Exp ρ foam (g/cm3)`, `ρr`, `X`, `Porosity (%)`.
  - For OC, ensure `Density (g/cm3)` is present.

### PDR CSV (foam level)
- Columns: `Time`, `T1 (°C)`, `T2 (°C)`, `P (bar)`.
- Output workbook: `PDR/Output/registros_promedios.xlsx` with sheet `Registros`.

### OC / Picnometry (foam level)
- Original Excel files live in `Open-cell content/Input`.
- Module output (recommended): `Open-cell content/Output/<Foam>_OC.xlsx` containing `%OC` column for Combine.

### DSC TXT (foam level)
- Place `.txt` files in `DSC/Input`.
- Output: `DSC/Output/DSC_<Foam>.xlsx` with Tg, Tm, Xc columns.

### SEM (foam level)
- Images or histogram summary Excel go into `SEM/Input`.
- Use the SEM Image Editor for annotations or histogram combiner as needed.

### All_Results.xlsx
- Produced by Smart Combine in `Results/`.
- Contains all canonical columns (`m(g)`, `Water (g)`, `T (°C)`, … `DSC Tg (°C)`). This file feeds scatter plots and heatmaps.

## 5. Modules

### 5.1 Smart Combine (⚡)
1. Open **SMART COMBINE**.
2. Select the paper base folder; missing paths are hinted from template structure.
3. Review detected files (DoE, Density, PDR, DSC, SEM, OC). Fill gaps if needed.
4. Run combine. Output: `Results/All_Results_YYYYMMDD.xlsx` plus logs.

### 5.2 Foam-Specific Tools (🔬)
- **PDR**: computes Pi, Pf, and PDR per experiment. Appends rows only for new CSVs.
- **OC / Picnometry**: parses comments, lets you override ball counts, outputs `%OC` for Combine.
- **DSC**: extracts Tg/Tm/Xc based on polymer settings.
- **SEM**: image editor and optional histogram combiner.

### 5.3 Scatter Plots (📈)
1. Load the desired `All_Results` workbook (any sheet containing the required columns).
2. Select sheet, X/Y axes, grouping column (optional), **Separate by** (facet panels), **Color by** (series color), and constancy filters. Error bars auto-enable when matching deviation columns exist.
3. Render scatter to visualize trends with the constancy rule enforced.
4. Actions:
   - *Render Plot* refreshes the scatter tab.
   - *Save Scatter* / *Copy Scatter* export the current figure.
   - *Export Data* writes filtered data + JSON config for reproduction.

### 5.4 Heatmaps (🔥) – separate module
1. Launch **HEATMAPS** (button or Tools menu).
2. Default path uses the last heatmap file or the last Combine output.
3. Load workbook → choose sheet. Independent and dependent variables appear in dedicated lists.
4. Select columns (multi-select with Ctrl/Shift). *Select all* and *Clear selection* apply to both lists.
5. Choose correlation method:
   - **Spearman** (default): robust to monotonic but nonlinear relationships.
   - **Pearson**: classic linear correlation; use when relationships are known to be linear.
   - **Distance (dCor)**: captures general dependence patterns (recommended when nonlinear effects dominate). Slightly heavier computationally.
6. Render heatmap → copy or save figure. The status bar summarizes sheet and variables used.

## 6. Frequent Issues & Tips
- **Column mismatch**: errors always list missing columns; adjust Excel headers accordingly.
- **Constancy rule (scatter)**: you must pin all independent variables except X and Group; the app guides you via combobox filters.
- **Heatmap selection**: requires ≥2 numeric columns after removing constant/empty ones; the module reports if selection is insufficient.
- **Clipboard copy on Windows**: requires `pywin32`. Install via `pip install pywin32` if copy buttons show errors.
- **State persistence**: scatter plots remember settings per sheet & file; if a workbook moves, reset selections as needed.
- **Git-friendly exports**: outputs are stored under `Results/` or module-specific `Output/` folders so you can version control only what matters.

---
With these guidelines you can manage papers, keep foams organized, extract raw metrics, and generate both scatter plots and correlation heatmaps ready for publication. Document unusual workflows in `Results/` notes so future combines remain reproducible.
