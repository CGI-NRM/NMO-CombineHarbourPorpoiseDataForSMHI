# NMO - Combine Harbour Porpoise Data For SMHI
This script is for combining multiple CPOD export files into a file ready to be delivered to SMHI.

## Options
### These are placed at the top of the CombineDataForSMHI.R file
> path_to_folder <- "./data/cpod_exports/"

> meta_data_path <- "./data/metadata.xlsx"

> target_positions_path <- "./data/target_positions.xlsx"

> smhi_template_path <- "./data/export_template.xlsx"

> export_xlsx <- TRUE

> export_csv <- TRUE

Since excel can only handle 1M rows the export will be split into multiple xlsx files each containing `rows_in_each_xlsx` rows of data if necessary.
> rows_in_each_xlsx <- 500000

## Folder structure
### Recommended folder structure: 
```bash
NMO-CombineHarbourPorpoiseDataForSMHI
├───CombineDataForSMHI.R
└───data
    ├───cpod_exports
    │   └───exported_harbour_porpoise.txt
    ├───metadata.xlsx
    ├───target_positions.xlsx
    └───export_template.xlsx
```

## Expected data
### Expected columns in the metadata
### Expected columns in the target positions data
### Expected columns in the export template

## Using the script
Run the script by using:
```bash
Rscript CombineDataForSMHI.R
```
Depending on `export_xlsx` and `export_csv` options above the script will produce files called `smhi_export_yyyy_mm_dd_index.xlsx`
```bash
NMO-CombineHarbourPorpoiseDataForSMHI
├───CombineDataForSMHI.R
├───data/
├───smhi_export_yyyy_mm_dd_index.csv
├───smhi_export_yyyy_mm_dd_index_part1.xlsx
└───smhi_export_yyyy_mm_dd_index_part2.xlsx
```