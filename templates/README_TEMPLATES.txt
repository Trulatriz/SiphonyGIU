PressTech Templates
===================

Place per-foam-type template files here. Suggested structure:

templates/
  PDR/
    registros_promedios_TEMPLATE.xlsx   # Empty Registros with headers/format
    sample_PDR_input.csv                # Example CSV columns: Time,T1 (ºC),T2 (ºC),P (bar)
  OC/
    Density_TEMPLATE.xlsx               # Required columns and sample rows
  DSC/
    DSC_SAMPLE.xlsx                     # Example file format used by your pipeline
  SEM/
    SEM_SAMPLE_SETTINGS.txt             # Example config for SEM processing
  Analysis/
    All_Results_TEMPLATE.xlsx
  Combine/
    DoE_TEMPLATE.xlsx

How to add a new foam type:
1) Duplicate the needed template(s) from the folders above.
2) Rename as appropriate for your foam type, e.g. registros_promedios_HDPE.xlsx.
3) Select your foam type at app start. Use these files when browsing input/output.

Notes:
- PDR: CSV must have 4 columns: Time, T1 (ºC), T2 (ºC), P (bar). Decimal separator in P can be comma.
- PDR Registros: Sheet name must be 'Registros'. Headers: Code | Pi (MPa) | Pf (MPa) | PDR (MPa/s) | Graph.
- OC: Density.xlsx must contain at least: Sample, Density (g/cm3).
- Combine/Analysis: Keep the exact column names expected by your existing modules.
