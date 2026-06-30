# SDEP v4.2 CSV Import Architecture

This package separates the equipment databases from the main workbook so GitHub can track data changes cleanly.

## Files

- `Excel/Current/SDEP_v4_2_csv_import.xlsx` — main workbook.
- `Data/Panels.csv` — panel database.
- `Data/Inverters.csv` — inverter database.
- `Data/Batteries.csv` — battery database.
- `Data/Lists.csv` — dropdown/list database.
- `Scripts/refresh_sdep_from_csv.py` — external refresh utility.
- `VBA/ImportCsvOnOpen.bas` — optional macro-based import module.

## Recommended workflow for now

1. Edit CSV files in `Data/`.
2. Run `python Scripts/refresh_sdep_from_csv.py`.
3. Open the refreshed workbook from `Excel/Current/`.
4. Verify the Build and Dashboard sheets.
5. Commit the CSV files and refreshed workbook.

## On-open startup import

A plain `.xlsx` cannot execute startup code. For true on-open import, save the workbook as `.xlsm`, import `VBA/ImportCsvOnOpen.bas`, and add this to `ThisWorkbook`:

```vb
Private Sub Workbook_Open()
    ImportAllSdepCsv
End Sub
```

Then place the `Data` folder next to the workbook.
