###################################################
#                                                 #
#  /-------------------------------------------\  #
#  | The Swedish Museum of Natural History     |  #
#  |                                           |  #
#  | Combine CPOD exports into a file ready    |  #
#  | to be exported to SMHI                    |  #
#  | Also export the data for export to HELCOM |  #
#  |                                           |  #
#  | Version: 0.2.2                            |  #
#  | Written by: Elias Lundell                 |  #
#  \-------------------------------------------/  #
#                                                 #
###################################################

library(readxl)
library(dplyr)
library(lubridate)
library(zeallot)
library(openxlsx)
library(RODBC)
library(magrittr)
library(tibble)

############################
# ------------------------ #
#     PATHS & OPTIONS      #
# ------------------------ #
############################

# The path to the folder containing all the .xlsx or .txt files from CPOD
path_to_folder <- "./data/cpod_exports/"

# The path to the meta-data file, either .xlsx or .accdb
meta_data_path <-
  "./data/metatest.accdb"

# The path to the file containing the target positions
target_positions_path <- "./data/Target positions NMO.xlsx"

# The path to an excel document which will be used as an template for exporting the data
smhi_template_path <-
  "./data/Format Harbour Porpoise 2021-06-21.xlsx"

# The path to an excel document which will bu used as an template for exporting the data to helcom
helcom_template_path <-
  "./data/Template helcom.xlsx"

# Whether to export the data to excel (split into XXXXXX rows in needed) and csv
export_xlsx <- TRUE
export_csv <- TRUE
export_helcom <- TRUE

# The numbers of rows in each excel document, if the data adds upp to more than this
# it will be split into multiple files
rows_in_each_xlsx <- 500000

# HELCOM 'Detection confirmed by visual inspection' replacement map
detection_confirmed_by_visual_inspection_hashmap <-
  new.env(hash = TRUE)
detection_confirmed_by_visual_inspection_hashmap[["y"]] <- "yes"
detection_confirmed_by_visual_inspection_hashmap[["no need"]] <-
  "no"
#detection_confirmed_by_visual_inspection_hashmap[["n.a."]] <- "no"

############################
# ------------------------ #
#       FUNCITONS          #
# ------------------------ #
############################

# Cache results from 'get_first_parts' function since it is slow (operating on strings)
# and it will be called many times with the same input
get_first_parts_dp <- new.env(hash = TRUE)

# Split a string on spaces, join (paste) the 6 first parts and return it
# Since the station name can vary in length we do this instead of taking the first
# 30 letters (used for point 12 in the word document)
get_first_parts <- function(x) {
  # If this input already has been calculated, return that
  dp <- get_first_parts_dp[[x]]

  if (!is.null(dp)) {
    return(dp)
  }

  # Split by spaces, take the first 6 parts and paste together
  res <- x %>% strsplit(" ") %>% unlist %>% head(n = 6) %>%
    paste0(sep = "", collapse = " ")
  # Save the result for future reuse
  get_first_parts_dp[[x]] <- res
  # Return the 6 first pasted parts
  res
}

# Load the meta data file and format the date-columns to be dates
load_meta_data <- function(meta_data_path) {
  # Save which columns contains dates to either convert to datetime (from excel) or forcing UTC on the
  # datetime if the metadata is imported from access without clutering the code later.
  # Could be made inline :)
  date_columns <-
    c(
      "Date for C-POD setup",
      "Time for C-POD setup",
      "C-POD deployment date",
      "C-POD deployment time",
      "Retrieval date",
      "Retrieval time",
      "Stop date",
      "Stop time",
      "First full day",
      "File end",
      "Last full day"
    )

  # Load the file, skip first line and remove the third line (contains descriptions of the columns)
  if (grepl("\\.xlsx$", meta_data_path)) {
    meta_data <-
      readxl::read_excel(
        path = meta_data_path,
        sheet = 2,
        skip = 1,
        # Give it names for now to remove the 'new names: `` -> ...1 `` -> ...2' message
        col_names = as.character(1:103),
        na = c("n.a.", "NA", "n.a", "unk")
      )[-2, ]

    # Using the built-in 'col_names = TRUE' produces a ton of errors, manually setting the
    # colnames based on the first row
    colnames(meta_data) <- meta_data[1, ]
    # The names are set, delete the first row that contains the names
    meta_data <- meta_data[-1, ]

    # Convert all columns containing dates to datetime-objects. Since its excel the dates are saved as
    # days since 30/12 - 1899.
    meta_data[, date_columns] <- meta_data[, date_columns] %>%
      lapply(function(x) {
        # Get a list of TRUE/FALSE whether the value only holds digits and dots and is not already NA
        not_number <-
          is.na(lubridate::as_datetime(
            as.numeric(x) * (24 * 60 * 60),
            origin = "1899-12-30",
            tz = "UTC"
          )) & !is.na(x)
        # If we find any that doesn't only contain digits and dots print it to the user so they can
        # change the data in the xlsx file
        if (any(not_number)) {
          print(paste0(
            "Some values cannot be turned into dates: '",
            paste0(x[not_number], collapse = "' '"),
            "'"
          ))
          # Make them NA here so that we will not get an NAs introduced by coercion error twice, the functions
          # will cause the error and we want the number of errors to corespond to the number of faults in the data
          x[not_number] <- NA
        }

        lubridate::as_datetime(as.numeric(x) * (24 * 60 * 60),
                               origin = "1899-12-30",
                               tz = "UTC")
      })
  } else if (grepl("\\.accdb$", meta_data_path)) {
    channel <- odbcConnectAccess2007(meta_data_path)
    meta_data <- tibble::tibble(sqlFetch(channel, "Deployment metadata"))
    odbcClose(channel)

    # Force all dates to be in UTC
    meta_data[, date_columns] <- meta_data[, date_columns] %>%
      lapply(function (x) {
        force_tz(x, tzone = "UTC")
      })
  } else {
    print("Unknown file format of meta-data, .xlsx and .accdb are supported.")
  }

  # Return the meta-data, with datetime-object in the columns holding dates
  meta_data
}

# Load the file containing the target coordinates
load_target_positions <- function(target_positions_path) {
  target_positions <-
    readxl::read_excel(
      path = target_positions_path,
      sheet = 1,
      skip = 1,
      col_names = TRUE
    )

  # The file is a little weird formated where most columns-names are in the
  # second row (hence the skip = 1). But station and comment are in the first (throw away) raw
  # and so we add them back instead of ...1 and ...8
  colnames(target_positions)[c(1, 8)] <- c("Station", "Comment")
  # Return only the necessary columns to reduce memory usage
  target_positions[, c("Station", "LAT", "LON")]
}

# Format one CPOD-expot file so they can be bound (concatenated) later
prepare_only_dpm_minutes <-
  function(filepath, meta_data, target_positions) {
    print(paste0("Loading file: ", filepath))
    # 4. Copy from row 8 and downwards into a new sheet called "All minutes"
    # ----------------------------------------------------------------------

    # Use the correct function depeding on if the file is a .txt or .xlsx file
    if (grepl('\\.txt$', filepath)) {
      # We only use the first 3 columns (File, ChunkEnd and DPM) so we throw away other columns
      all_minutes <-
        tibble::tibble(read.table(
          file = filepath,
          header = TRUE,
          sep = "\t",
          skip = 7,
          stringsAsFactors = FALSE
        ))[, c("File", "ChunkEnd", "DPM")]


      # We must format the columns ourselves, make the DPM column numbers (1:s and 0:s)
      all_minutes[, "DPM"] <-
        all_minutes[, "DPM"] %>% lapply(as.numeric)

      # In the txt file the date is saved in a string format, convert it into datetime objects
      all_minutes[, "ChunkEnd"] <-
        all_minutes[, "ChunkEnd"] %>% lapply(lubridate::as_datetime, format = "%d/%m/%Y %H:%M")

      print(all_minutes)

    } else if (grepl('\\.xlsx$', filepath)) {
      # We only use the first 3 columns (File, ChunkEnd and DPM) so we throw away other columns
      # The readxl package can automatically load a columns as dates or numeric so there is no
      # need for us to do the convertion manually. Once the file is saved from excel the dates are
      # converted from strings to number of dates since 30/12 - 1899. But read_excel gives us the
      # column as datetime:s
      all_minutes <-
        readxl::read_excel(
          path = unlist(filepath),
          sheet = 1,
          skip = 7,
          col_names = TRUE,
          col_types = c("text", "date", "numeric", "guess", "guess")
        )[, c("File", "ChunkEnd", "DPM")]
    }


    # 5. Check that the start and end times are correct in the meta data
    # ------------------------------------------------------------------

    # Figure out which row in the meta-data this CPOD-file relates to. The 'File' (from CPOD)
    # and 'File name' (metadata) are not an exact match, the metadata doesn't save the last part
    matching_row <- meta_data[, "File name"] %>% unlist %>%
      sapply(grepl, all_minutes[1, "File"], ignore.case = TRUE)

    # If an row is NA in the meta-data we do not wish to match against it
    matching_row[is.na(matching_row)] <- 0
    # Convert the vector from 1:s and 0:s to TRUE:s and FALSE:s (one TRUE with the row matching)
    matching_row <- matching_row == 1

    # If we could not find any row, alert the user
    if (!any(matching_row)) {
      print("    Could not find any matching row in the meta data.")
    }

    # Get the difference between the first row of data in the CPOD export and what has been
    # saved in the meta-data
    first_diff <-
      meta_data[matching_row, "First full day"] - head(all_minutes[, "ChunkEnd"], n = 1)
    # If there is not a perfect match alert the user ansfk how big the difference isuser
    if (first_diff != minutes(0)) {
      print(
        paste0(
          "    First full day does not match first row in file, the last row is early by: ",
          format(first_diff)
        )
      )
    }

    # Get the difference between the last data-point in the CPOD-export and what has been saved in the
    # meta-data
    last_diff = (meta_data[matching_row, "Last full day"] + (23 * 60 + 59) * 60) - tail(all_minutes[, "ChunkEnd"], n = 1)
    # If there is not a perfect match alert the user about how big the difference is, often this is one minute
    if (last_diff != minutes(0)) {
      print(
        paste0(
          "    Last full day does not match last row in file, the last row is early by: ",
          format(last_diff)
        )
      )
    }

    # 6. 7. Copy time data into separate time and date columns
    # --------------------------------------------------------

    # Create the odate and otime columns, one containing the date and one the time. Formatted to strings
    all_minutes[, "ODATE"] <-
      sapply(all_minutes[, "ChunkEnd"], format, format = "%Y-%m-%d")
    all_minutes[, "OTIME"] <-
      sapply(all_minutes[, "ChunkEnd"], format, format = "%H:%M:%S")

    # 8. 9. Create a separate year column
    # -----------------------------------

    # Create the year column containing the string of which year the date is
    all_minutes[, "MYEAR"] <-
      sapply(all_minutes[, "ChunkEnd"], format, format = "%Y")

    # 10. Copy the station number from the meta-data into a new column
    # ----------------------------------------------------------------

    all_minutes[, "STATN"] <-
      meta_data[matching_row, "Station number"]

    # 11. Start and end dates and times are needed in separate columns. Copy this
    # information from the first and last rows into columns named 'SDATE', 'STIME',
    # 'EDATE' and 'ETIME'.
    # -----------------------------------------------------------------------------

    all_minutes[, "SDATE"] <- head(all_minutes[, "ODATE"], n = 1)
    all_minutes[, "STIME"] <- head(all_minutes[, "OTIME"], n = 1)
    all_minutes[, "EDATE"] <- tail(all_minutes[, "ODATE"], n = 1)
    all_minutes[, "ETIME"] <- tail(all_minutes[, "OTIME"], n = 1)

    # 12. 13. Create a column containing the file name in the format
    # '[station number] [deployment date] [C-POD number] [file01]'
    # The 'get_first_parts' function splits by a space and takes the first
    # 6 parts (the date is YEAR MONTH DAY, the rest have no space in them)
    # --------------------------------------------------------------------
    all_minutes[, "RAW"] <-
      all_minutes[, "File"] %>% unlist %>% sapply(get_first_parts)

    # 14. Create a column containing the longitude and latitude of the pod
    # --------------------------------------------------------------------
    all_minutes[, "LONGI"] <-
      meta_data[matching_row, "C-POD deployment LONG"]
    all_minutes[, "LATIT"] <-
      meta_data[matching_row, "C-POD deployment LAT"]

    # 14.5. If we know the target position of the pod (only some projects) add that to a new column
    # ---------------------------------------------------------------------------------------------

    # Get the row in the target_positions by finding the station number from the meta data and
    # matching that to the station in the target_position data
    target_position_row <-
      target_positions[, "Station"] == unlist(meta_data[matching_row, "Station number"])

    # Initialise the variables as empty, since we might now have found a row in the targetpositions data
    target_long <- ""
    target_lat <- ""

    if (any(target_position_row)) {
      # If we matched a row use the longitude and latitude from the data
      target_long <- target_positions[target_position_row, "LON"]
      target_lat <- target_positions[target_position_row, "LAT"]
    } else {
      # If we do not find a row inform the user
      print(paste0("    No target position was found for ", meta_data[matching_row, "Station number"]))
    }

    # We add _prov since the colnames needs to be unique and there are already columns called
    # LONGI and LATIT. This will be solved by using the template that ignores the column name
    # when exporting
    all_minutes[, "LONGI_prov"] <- target_long
    all_minutes[, "LATIT_prov"] <- target_lat


    # 14.6. If the saved depth was measured at the time of placement (as opposed to read from
    # the sea chart) add that to our export
    # ---------------------------------------------------------------------------------------

    # Initialize the depth variables since we might not have any data
    water_depth <- ""
    pod_depth <- ""

    if (!is.na(meta_data[matching_row, "Depth type"]) && meta_data[matching_row, "Depth type"] == "measured") {
      # If it was measured, override the empty variables with the data
      water_depth <- meta_data[matching_row, "Water depth"]
      pod_depth <- meta_data[matching_row, "C-POD depth"]
    }

    # Add the depth to our export (might be empty variables)
    all_minutes[, "WADEPTH"] <- water_depth
    all_minutes[, "SMPDEPTH"] <- pod_depth

    # 14.7. Fill some columns that always contain the same thing
    # ----------------------------------------------------------

    # It will always be NMHS (National Museum of natural History Sweden), otherwise it is easy fix manually
    # in the exported file.
    all_minutes[, "SLABO"] <- "NMHS"
    # Some tags meaning this is research (R) and some other things ;)
    all_minutes[, "PURPM"] <- "B,R,S,T"

    # If this CPOD-export is for the NM? (nationell milj? ?vervakning) projekt the code for
    # type of survelliance is NATL otherwise it is research(?)
    if (!is.na(meta_data[matching_row, "Project"]) && meta_data[matching_row, "Project"] == "NMO") {
      all_minutes[, "MPROG"] <- "NATL"
    } else {
      all_minutes[, "MPROG"] <- "PROJ"
    }

    if (export_helcom) {
      # Copy 'Visually validated' for HELCOM export
      all_minutes[, "Visually validated"] <-
        meta_data[matching_row, "Visually validated"]
    }

    # 15. The all minutes sheet is now complete
    # -----------------------------------------

    # 16. SMHI only wants the row with harbour purpoise activity, create a new tibble
    # -------------------------------------------------------------------------------

    # Using the tibble function copies the data in memory intsead of just copying the pointer
    only_dpm_minutes <- tibble::tibble(all_minutes)

    # 17. Keep only the rows with activity, unless no rows have activity where we save a random
    # row to save that there has been survellience at that point during that time
    # -----------------------------------------------------------------------------------------

    # Keep only the rows with acitvity
    only_dpm_minutes <-
      only_dpm_minutes[only_dpm_minutes[, "DPM"] == 1, ]

    # If we have throws away to much or little alert the user. This should generally never happen
    # since the computer cannot make an error. In the word file this step is (probably) to catch
    # human error.
    if (sum(all_minutes[, "DPM"], na.rm = TRUE) != nrow(only_dpm_minutes)) {
      print("    The sum of DPM was not kept when removing the minutes without a harbour porpoise.")
    }

    # If no rows had activity we keep the first row, because of this we need the copy (of atleast)
    # the first row
    if (nrow(only_dpm_minutes) == 0) {
      # Alert the user that this happened
      print("    There was no row containing a harbour porpoise, keeping one row to save the .")
      # Copy the first row (with no activity, since no rows had any activity)
      only_dpm_minutes[1,] <- all_minutes[1,]
    }

    # 18. Sort the rows based on oldest (up) to newest (down) if the rows have been shuffled
    # --------------------------------------------------------------------------------------

    # Arrange based on the two variables, primarily on the date and secondary on the time.
    only_dpm_minutes <- only_dpm_minutes %>% arrange(ODATE, OTIME)

    # 19. Replace all DPM = 1 by 'Y' and if there is none, replace the zero by 'N'.
    # -----------------------------------------------------------------------------

    # First convert the entire column of numerics (1:s and 0:s) to characters since tibble:s
    # do not allow one column to mix types
    only_dpm_minutes[, "DPM"] <- only_dpm_minutes[, "DPM"] %>%
      sapply(as.character)

    only_dpm_minutes[only_dpm_minutes[, "DPM"] == 1, "DPM"] <- "Y"
    only_dpm_minutes[only_dpm_minutes[, "DPM"] == 0, "DPM"] <- "N"

    # Return the data with a lot of extra columns, only keeping the rows with activity (or one row)
    only_dpm_minutes
  }

# Export data to excel and split it into multiple files if necessary (larger than rows_in_each_xlsx)
export_to_xlsx <-
  function(data_to_export,
           template_path,
           export_file_name,
           sheet_name,
           startCol,
           startRow) {
    # The index of the row_segment to be saved this loop
    part_ind <- 1
    # Cache how many rows of data we have to not have to do this many times
    data_rows <- nrow(data_to_export)

    print(paste0(
      "Writing data to excel template... (",
      ceiling(data_rows / rows_in_each_xlsx),
      " file(s))"
    ))

    # While there are more rows in total than we have exported
    while (data_rows > (part_ind - 1) * rows_in_each_xlsx) {
      # The segment of the data to export, 1:5 -> 6:10 -> 11:15 etc
      part_segment <-
        ((part_ind - 1) * rows_in_each_xlsx + 1):(min(part_ind * rows_in_each_xlsx, data_rows))

      # An ending to add to the file, will reset if only one file is needed
      file_part_number <- paste0("_part", part_ind)

      # If this is the only file we need, add nothing special to the file name
      if (data_rows <= rows_in_each_xlsx) {
        file_part_number <- ""
      }

      # Open the template
      xlsx_export <- openxlsx::loadWorkbook(template_path)

      # Write our data (which is correctly spaced out). As of me writing this (2021-06-23) I could not find any
      # way to disable the header row. colNames chooses between using the names and using Col1 Col2 etc.
      # Maybe another package could do it, this one was nice because it didn't use a JVM (which limits us the 1gb
      # which is way to small).
      openxlsx::writeDataTable(
        xlsx_export,
        sheet_name,
        data_to_export[part_segment,],
        startCol = startCol,
        startRow = startRow,
        headerStyle = NULL,
        colNames = TRUE,
        withFilter = FALSE
      )

      print(paste0(
        "Saving export file: ",
        paste0(export_file_name, index, file_part_number, ".xlsx")
      ))
      # Save the file
      openxlsx::saveWorkbook(xlsx_export,
                             paste0(export_file_name, index, file_part_number, ".xlsx"))

      # This segment of the export is done, now export the next part (or finish the while if we have
      # exported all of the rows of data)
      part_ind <- part_ind + 1
    }
  }

############################
# ------------------------ #
#     PERFORMING CODE      #
# ------------------------ #
############################

# Create a list of full paths (from workingdirectory) to the files in the folder
files_to_load <- paste0(path_to_folder, list.files(path_to_folder))
# Keep only the .txt and .xlsx files
files_to_load <-
  files_to_load[grepl('(\\.xlsx)|(\\.txt)$', files_to_load)] %>% as.list

# Load the metadata
meta_data <-
  load_meta_data(meta_data_path = meta_data_path)

# Load the target_positions
target_positions <-
  load_target_positions(target_positions_path = target_positions_path)

# 20. Join the data together if there is more than 1 file
# -------------------------------------------------------

# Make the combined variable null so we can first replace it with data and then bind more rows to the data
combined_dpm_minutes <- NULL

# Go through all the files in the folder, we use a for loop and do them one at a time instead of lapply
# to reduce our memory usage, using lapply we need to keep them all in memory and they are quite large.
# This way we only hold the combined (which only have the active rows) and all the rows of the current file.
for (path in files_to_load) {
  # If the combined variable is null this is the first file we are processing, replace the variable with data
  if (is.null(combined_dpm_minutes)) {
    combined_dpm_minutes <-
      prepare_only_dpm_minutes(path, meta_data = meta_data, target_positions = target_positions)
  } else {
    # This is not the first file we are processing, instread of replacing the variblae bind it to the end
    combined_dpm_minutes <-
      combined_dpm_minutes %>% rbind(
        prepare_only_dpm_minutes(path, meta_data = meta_data, target_positions = target_positions)
      )
  }
}

# If the combined data is still null the for-loop above did not run, alert the user
if (is.null(combined_dpm_minutes)) {
  print(paste0("There were no files in the folder ", path_to_folder))
}

print("Done combining all files, ordering the columns...")

# 21. 22. Sort the columns in the following way and discard the rest
# ------------------------------------------------------------------

# Now that we start to sort and filter the SMHI export, we copy the visually validated column to a
# separate dataframe to keep it
helcom_save_df <-
  combined_dpm_minutes[, "Visually validated"]

# Choose wether to fit the data to the template (hard coded) (= FALSE) or just order the columns we have
# created/added data to (= TRUE)
if (FALSE) {
  # These are the columns that we have created and filled with data, in the correct order. Sorting the
  # the columns in this way
  combined_dpm_minutes <-
    combined_dpm_minutes[, c(
      "MYEAR",
      "STATN",
      "LATIT",
      "LONGI",
      "SDATE",
      "STIME",
      "EDATE",
      "ETIME",
      "PURPM",
      "MPROG",
      "SLABO",
      "LATIT_prov",
      "LONGI_prov",
      "DPM",
      "ODATE",
      "OTIME",
      "RAW",
      "SMPDEPTH",
      "WADEPTH"
    )]
} else {
  # To fit the data to the template we must add some empty columns between our columns that we already have
  # (and are filled with data)

  # Create a dataframe containing 12 empty columns
  empty_cols <-
    data.frame(as.list(rep(NA, 12)))

  # Name the columns after the columns from the template, this is not really necessary but it makes it easier
  # later on when we must sort all columns
  colnames(empty_cols) <- c(
    "PROJ",
    "ORDERER",
    "POSYS",
    "COMNT_VISIT",
    "ACKR_SMP_prov",
    "SMTYP",
    "METDC",
    "COMNT_SAMP",
    "LATNM",
    "ALABO",
    "ACKR_SMP",
    "COMNT_VAR"
  )

  # Join the empty columns to our data and sort all the columns in the order found in the template to
  # line it up with the descriptions/header in the template. Here it is easier if the empty columns have
  # correct names since we then can check against the template. We could hovewer also
  # name them "A", "B", "C" and so and just use them to create gaps in our data-columns
  combined_dpm_minutes <-
    cbind(combined_dpm_minutes, empty_cols)[, c(
      "MYEAR",
      "STATN",
      "LATIT",
      "LONGI",
      "PROJ",
      "ORDERER",
      "SDATE",
      "STIME",
      "EDATE",
      "ETIME",
      "POSYS",
      "PURPM",
      "MPROG",
      "COMNT_VISIT",
      "SLABO",
      "ACKR_SMP_prov",
      "SMTYP",
      "LATIT_prov",
      "LONGI_prov",
      "METDC",
      "COMNT_SAMP",
      "LATNM",
      "DPM",
      "ODATE",
      "OTIME",
      "ALABO",
      "ACKR_SMP",
      "RAW",
      "COMNT_VAR",
      "SMPDEPTH",
      "WADEPTH"
    )]
}

# Convert the rows containing strings into numbers to stop excel from giving an error, it is not really
# necessary but makes it easier to work with in excel (we can do calculation on them, which is not possible
# to do with strings)
combined_dpm_minutes[, c("MYEAR", "LATIT", "LONGI", "LATIT_prov", "LONGI_prov")] <-
  combined_dpm_minutes[, c("MYEAR", "LATIT", "LONGI", "LATIT_prov", "LONGI_prov")] %>%
  apply(c(1, 2), as.numeric)

# The columns should not be named _prov so we can choose to remove that, but we use the names from the template
# and it is easier here to separate the column names
if (FALSE) {
  colnames(combined_dpm_minutes)[colnames(combined_dpm_minutes) == "LATIT_prov"] <-
    "LATIT"
  colnames(combined_dpm_minutes)[colnames(combined_dpm_minutes) == "LONGI_prov"] <-
    "LONGI"
}

# 23. Fill in the SMHI template by coping the data
# ------------------------------------------------

# Find an available filename to not override our old exports
index <- 1
smhi_export_file_name <- paste0("./smhi_export_", Sys.Date(), "_")
helcom_export_file_name <-
  paste0("./HELCOM_export_", Sys.Date(), "_")

# While a file exists with that name, add one to our index. Each export (the same day) will produce:
# smhi_export_yyyy_mm_dd_1.xlsx
# smhi_export_yyyy_mm_dd_2.xlsx
# smhi_export_yyyy_mm_dd_3.xlsx
while (file.exists(paste0(smhi_export_file_name, index, ".xlsx")) ||
       file.exists(paste0(smhi_export_file_name, index, "_part1.xlsx")) ||
       file.exists(paste0(smhi_export_file_name, index, ".csv")) ||
       file.exists(paste0(helcom_export_file_name, index, ".xlsx")) ||
       file.exists(paste0(helcom_export_file_name, index, "_part1.xlsx"))) {
  index <- index + 1
}

if (export_xlsx) {
  # Export data and split it into multiple files if necessary
  export_to_xlsx(
    data_to_export = combined_dpm_minutes,
    template_path = smhi_template_path,
    export_file_name = smhi_export_file_name,
    sheet_name = "Kolumner",
    startCol = 1,
    startRow = 5
  )
}

if (export_csv) {
  print("Writing data to a csv file...")

  write.csv(
    combined_dpm_minutes,
    file = paste0(smhi_export_file_name, index, ".csv"),
    na = "",
    fileEncoding = "utf8",
    row.names = FALSE
  )
}

if (export_helcom) {
  # Create the empty tibble which we can add data to
  helcom <- tibble::tibble()

  # Set the country to Sweden on as many rows as we will use. Since the tibble is empty we have to use
  # the seq code, instead of it just automatically repeating Sweden across all rows (since there are none)
  # right now
  helcom[seq(1, nrow(combined_dpm_minutes)), "Country"] <-
    rep("Sweden", nrow(combined_dpm_minutes))

  # Copy over the dpm to detection, and replace "Y" with 1 and "N" with 0
  helcom[, "Detection"] <- combined_dpm_minutes[, "DPM"] %>%
    lapply(function (x) {
      if (x == "Y") {
        return(1)
      }
      return(0)
    }) %>% unlist

  # Now we can set all the rows in a column to be the same, adding metadata that is consistant
  helcom[, "Unit"] <- "Detection positive minutes"
  helcom[, "Type of recording device"] <- "C-POD"
  helcom[, "Data collector"] <- "MNHS"
  helcom[, "Data holder"] <- "SMHI"
  helcom[, "Filter used"] <- "Hel1"
  helcom[, "Information_withheld"] <- "no"
  helcom[, "Restriction_yes_no"] <- "no"

  # Extract all the detection dates, since HELCOM wants to have year, month and day in separate columns
  # we first split it by - (2020-04-15) becomes c(2020, 04, 15).
  odates <- combined_dpm_minutes[, "ODATE"] %>% lapply(function(x) {
    unlist(strsplit(x, "-"))
  })
  # We then convert it to a dataframe since we need to rotate the columns and rows. (To set entire cols
  # later). We extract the first, second and third element in each list and make those separate columns
  odates_df <-
    data.frame(
      year = lapply(odates, `[`, 1) %>% unlist %>% as.numeric,
      month = lapply(odates, `[`, 2) %>% unlist %>% as.numeric,
      day = lapply(odates, `[`, 3) %>% unlist %>% as.numeric
    )

  # Now copy the data over from the dataframe, this could be made inline, but this way its a little bit
  # easier to read (I think)
  helcom[, "Day of UTC detection"] <- odates_df[, "day"]
  helcom[, "Month of UTC detection"] <- odates_df[, "month"]
  helcom[, "Year of UTC detection"] <- odates_df[, "year"]
  # Then we copy over the times, they need not be changed
  helcom[, "Time of UTC detection"] <-
    combined_dpm_minutes[, "OTIME"]

  # Copy the water depth, this is only taken if the depth is known previously (measured).
  helcom[, "Position CPOD in the water column"] <-
    combined_dpm_minutes[, "SMPDEPTH"]

  # Now we extract the same dates from the start date and end date
  # Split the date on - to separate year, monht and day
  sdates <- combined_dpm_minutes[, "SDATE"] %>% lapply(function(x) {
    unlist(strsplit(x, "-"))
  })
  # Rotate the list of vectors into a dataframe with a year, month and day columns
  sdates_df <-
    data.frame(
      year = lapply(sdates, `[`, 1) %>% unlist %>% as.numeric,
      month = lapply(sdates, `[`, 2) %>% unlist %>% as.numeric,
      day = lapply(sdates, `[`, 3) %>% unlist %>% as.numeric
    )

  # Split the end dates on - to get separate year, month and day
  edates <- combined_dpm_minutes[, "EDATE"] %>% lapply(function(x) {
    unlist(strsplit(x, "-"))
  })
  # Rotate the list of vectors to get a dataframe with a year, month and day column
  edates_df <-
    data.frame(
      year = lapply(edates, `[`, 1) %>% unlist %>% as.numeric,
      month = lapply(edates, `[`, 2) %>% unlist %>% as.numeric,
      day = lapply(edates, `[`, 3) %>% unlist %>% as.numeric
    )

  # Copy the start and end dates data over from their repsective dataframes to the helcom tibble
  helcom[, "Day device start"] <- sdates_df[, "day"]
  helcom[, "Month device start"] <- sdates_df[, "month"]
  helcom[, "Year device start"] <- sdates_df[, "year"]

  helcom[, "Day device end"] <- edates_df[, "day"]
  helcom[, "Month device end"] <- edates_df[, "month"]
  helcom[, "Year device end"] <- edates_df[, "year"]

  # Copy the measured position over. The target positions are stored in LATIT_prov and LONGI_prov should
  # they be preferable in the future
  helcom[, "Latitude deployment of the device"] <-
    combined_dpm_minutes[, "LATIT"]
  helcom[, "Longitude deployment of the device"] <-
    combined_dpm_minutes[, "LONGI"]

  # Copy visually validated from the dataframe to the helcom
  helcom[, "Detection confirmed by visual inspection"] <-
    helcom_save_df[,"Visually validated"] %>% apply(1, function(x) {
      if (is.na(x)) {
        return(NA)
      }

      res <- detection_confirmed_by_visual_inspection_hashmap[[x]]

      if (is.null(res)) {
        print(
          paste0(
            "The value of visually validated: ",
            x,
            ". Was not found in the replacement map."
          )
        )
        return(NA)
      } else {
        return(res)
      }
    }) %>% unlist

  # Copy Station number to helcom export
  helcom[, "Station Identification"] <- combined_dpm_minutes[, "STATN"]

  # Create a dataframe containing empty columns so that we can space it out correctly for the export
  helcom_empty_cols <- data.frame(as.list(rep(NA, 12)))

  # Name the columns which contain no data
  colnames(helcom_empty_cols) <- c(
    "Detection confirmed by visual inspection",
    "Restriction_description",
    "Citation",
    "HELCOM id",
    "Species id",
    "Collection code",
    "Subunit",
    "Field number",
    "Point_id",
    "Upload_date"
  )

  # Combine the data with all the empty columns and sort them based on the order in the template
  helcom <-
    cbind(helcom, helcom_empty_cols)[, c(
      "Country",
      "Type of recording device",
      "Day of UTC detection",
      "Month of UTC detection",
      "Year of UTC detection",
      "Time of UTC detection",
      "Detection",
      "Unit",
      "Filter used",
      "Detection confirmed by visual inspection",
      "Position CPOD in the water column",
      "Data collector",
      "Data holder",
      "Day device start",
      "Month device start",
      "Year device start",
      "Day device end",
      "Month device end",
      "Year device end",
      "Station Identification",
      "Latitude deployment of the device",
      "Longitude deployment of the device",
      "Information_withheld",
      "Restriction_yes_no",
      "Restriction_description",
      "Citation",
      "HELCOM id",
      "Species id",
      "Collection code",
      "Subunit",
      "Field number",
      "Point_id",
      "Upload_date"
    )]


  # Export to excel file, the function splits it into multiple files if the data is too big
  export_to_xlsx(
    data_to_export = helcom,
    template_path = helcom_template_path,
    export_file_name = helcom_export_file_name,
    sheet_name = 1,
    startCol = 1,
    startRow = 2
  )
}
