# NMO - Combine Harbour Porpoise Data For SMHI
This script is for combining multiple CPOD export files into a file ready to be delivered to SMHI.

## Options
### These are placed at the top of the CombineDataForSMHI.R file
> path_to_folder <- "./data/cpod_exports/"

> meta_data_path <- "./data/metadata.xlsx" (or metadata.accdb)

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
    ├───metadata.accdb
    ├───target_positions.xlsx
    └───export_template.xlsx
```

## Expected data
### Expected columns in the metadata
| # | Station name | Project | Station number | C-POD no | SD-card no | Entered by | Date for C-POD setup | Time for C-POD setup | Battery brand and model | Silica gel | Func test | Setup procedure | Setup operator | Setup comments | C-POD deployment date | C-POD deployment time | C-POD deployment GPS LAT | C-POD deployment GPS LONG | C-POD deployment LAT | C-POD deployment LONG | Xtra anchor deployment GPS LAT | Xtra anchor deployment GPS LONG | Xtra anchor deployment LAT | Xtra anchor deployment LONG | Water depth | Depth type | C-POD depth | Mooring type  | Releaser | Sound recording equipment | Bottom line length | Line weights | Line floats | Wave height | Wind speed | Deployment operator | Deployment vessel | Deployment crew | Deployment comments | Diff lat | Diff long | Diff C-POD-anchor | Date diff start-deploy | Retrieval date | Retrieval time | Retrieval GPS LAT | Retrieval GPS LONG | Retrieval LAT  | Retrieval LONG | Retrieval depth | Retrieval type | # retreival attempts | If unplanned or lost, cause | Estimated time detached | Wave height | Wind speed | Retrieval operator | Retrieval vessel | Retrieval crew | Retrieval comments | Blinking after recovery | Stop date | Stop time | Battery left stack | Battery right stack | Stop operator | Stop comments | Date diff retrieval-stop | File name | Full archive done | Download procedure | Download operator | Download comments | Deployment duration | Crop start date correction | First full day | File end | Last full day | Logging duration | Missing days | Missing days | File cropped | Cropped file duration cpod.exe | KERNO cp3 created | Hel1 cp3 created | Comment | Hel1 DPM | DPM/ yr | Visually validated | False positive DPMs | cpod.exe data exported | Excel file prepared | SMHI form filled in | Date sent to SMHI | Comment | SLIM |
| - | ------------ | ------- | -------------- | -------- | ---------- | ---------- | -------------------- | -------------------- | ----------------------- | ---------- | --------- | --------------- | -------------- | -------------- | --------------------- | --------------------- | ------------------------ | ------------------------- | -------------------- | --------------------- | ------------------------------ | ------------------------------- | -------------------------- | --------------------------- | ----------- | ---------- | ----------- | ------------  | -------- | ------------------------- | ------------------ | ------------ | ----------- | ----------- | ---------- | ------------------- | ----------------- | --------------- | ------------------- | -------- | --------- | ----------------- | ---------------------- | -------------- | -------------- | ----------------- | ------------------ | -------------  | -------------- | --------------- | -------------- | -------------------- | --------------------------- | ----------------------- | ----------- | ---------- | ------------------ | ---------------- | -------------- | ------------------ | ----------------------- | --------- | --------- | ------------------ | ------------------- | ------------- | ------------- | ------------------------ | --------- | ----------------- | ------------------ | ----------------- | ----------------- | ------------------- | -------------------------- | -------------- | -------- | ------------- | ---------------- | ------------ | ------------ | ------------ | ------------------------------ | ----------------- | ---------------- | ------- | -------- | ------- | ------------------ | ------------------- | ---------------------- | ------------------- | ------------------- | ----------------- | ------- | ---- |

All columns containing dates should be set to the date-time type in access (or formated as dates in excel). The script expects the following columns to contain dates: 
| Date for C-POD setup | Time for C-POD setup | C-POD deployment date | C-POD deployment time | Retrieval date | Retrieval time | Stop date | Stop time | First full day | File end | Last full day |
| ------------------- | -------------------- | --------------------- | --------------------- | -------------- | -------------- | --------- | --------- | -------------- | -------- | ------------- |

### Expected columns in the target positions data
| Station | LAT | LON | LAT DEG	| LAT MIN | LON DEG | LON MIN |
| ------- | --- | --- | ------- | ------- | ------- | ------- |

### Expected columns in the export template
| MYEAR | STATN | LATIT | LONGI  | PROJ | ORDERER | SDATE | STIME | EDATE | ETIME | POSYS | PURPM | MPROG | COMNT_VISIT | SLABO | ACKR_SMP | SMTYP | LATIT | LONGI  | METDC | COMNT_SAMP | LATNM | DPM | ODATE | OTIME | ALABO | ACKR_SMP | RAW | COMNT_VAR | SMPDEPTH | WADEPTH |
| ----- | ----- | ----- | -----  | ---- | ------- | ----- | ----- | ----- | ----- | ----- | ----- | ----- | ----------- | ----- | -------- | ----- | ----- | -----  | ----- | ---------- | ----- | --- | ----- | ----- | ----- | -------- | --- | --------- | -------- | ------- |

### Metadata sheet/table name
**Excel** The data is assumed to be in the second sheet.

**MS Access** The data is assumed to be in a table named "Deployment metadata". 

They can easily be changed in the script in the load_meta_data function if the format of the input data were to change. 

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
