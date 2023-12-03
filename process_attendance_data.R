# Session Info =================================================================
# R version 4.3.2 (2023-10-31 ucrt)
# Platform: x86_64-w64-mingw32/x64 (64-bit)
# Running under: Windows 10 x64 (build 19045)

# RStudio 2023.09.1+494 "Desert Sunflower" Release (cd7011dce393115d3a7c3db799dda4b1c7e88711, 2023-10-16) for windows


# Documentation =================================================================

## Purpose --------------------------------- 
# - This script extracts information from attendance reports and merges them into a clean, tabular format.

## IMPORTANT NOTE --------------------------
# - Please open "UpworkAttend.Rproj" first, and then run this script.  


## Dates ------------------------------------
# Creation Date: November 30th, 2023
# Last Modified: December 3rd, 2023

## Steps ------------------------------------
# 1) Load libraries
#   - readxl (v1.4.3) - Reads excel files
#   - openxlsx (v4.2.5.2) - Saves excel files
#   - dplyr (v1.1.4) - For data manipulation
#   - stringr (v1.5.1) - For string manipulation
# 2) Define global constants
#   - Create a vector of column names to keep from attendance reports
# 3) Define a function for reading in and processing attendance report data
# 4) Call function to read in and process attendance report data
# 5) Save data


# Reproducible Environment ====================================================================================

# install.packages("renv") 

# If this line fails, un-comment and run the line above to install renv
renv::restore() 

# Note about renv::restore():
# Running this line of code may prompt you (in the R console) to install additional packages and dependencies.

# Load libraries =============================================================================================
library(readxl) # Version 1.4.3
library(dplyr) # Version 1.1.4
library(stringr) # Version 1.5.1
library(openxlsx) # Version 4.2.5.2

# Global constants ===========================================================================================

# Create a vector of column names to keep from attendance reports
COL_NAMES <- c("ac-no", "timetable", "normal", "actual", "absent", "late", "early", "ot", "afl", 
               "bleave", "weekend_ot")


# Define function for reading in and processing attendance report data =======================================

readAndProcess <- function(fileName) {
  
  print(paste0("Reading in and processing ", fileName))
  
  # Extract date from fileName; Convert to Date object
  date <- as.Date( gsub("\\D", "", fileName), "%Y%m%d" )

  ## Read data ---------------------------------------
  dat <- read_excel(fileName, .name_repair = "unique_quiet") %>% 
    mutate(across(everything(), ~str_to_lower(.x)), # Convert all values to lowercase
           across(everything(), ~gsub("[\r\n]", "", .x))) # Remove any tags that may exist
  
  ## Split data into separate tables -----------------
  
  # Find row indices where each table begins
  rowIndex <- which(
    apply(dat, 1, function(r) any(r %in% COL_NAMES))
  )
  
  # Split into tables
  dats <- split(dat, cumsum(1:nrow(dat) %in% rowIndex))
  
  
  ## Clean each table ---------------------------------
  datsBind <- lapply(names(dats)[-1], function(x) {
    
    tDat <- dats[[x]] # For easier reference
    
    names(tDat) <- as.character( tDat[1, ] ) # Make first row column names
    
    tDat <- tDat[-1, ] # Remove first row
    
    # Remove blank or NA column names
    missing_colNames <- which(names(tDat) %in% c(NA, ""))
    tDat <- tDat[, -missing_colNames]
    
    # Remove rows where all values are missing; Clean and finalize data
    tDat <- tDat %>%
      select(-name) %>% # Remove name column
      filter(!if_all(everything(), ~is.na(.x))) %>% # Removes rows with all NA values
      mutate(source_file = fileName, # Add column with file name
             date = as.character(date),  # ADd column with date
             ) %>% 
      select(source_file, date, any_of(COL_NAMES)) %>% # Select desired columns for final output
      mutate(across(c(`ac-no`, normal:ot), ~as.numeric(.x)), # Convert certain columns to numeric 
             across(afl:weekend_ot, ~as.logical(.x)) # Convert certain columns to logical (TRUE, FALSE)
             ) %>% 
      rename(ac_no = `ac-no`) # Rename column
    
    return(tDat) # Return clean table
    
  }) %>% 
    bind_rows()

}

# Call function to read in and process attendance report data ==============================================

## List files in directory ------------------------------- 
files <- list.files(pattern = "[A-Z]\\.xlsx$") # Lists all files in directory that end with "[UPPERCASE LETTER].xlsx"

## Loop through each file and call function -------------
dataAttendance <- lapply(files, readAndProcess)


# Append and save data ====================================================================================
dataAttendance_final <- dataAttendance %>% 
  bind_rows() %>% 
  arrange(ac_no, timetable)

write.xlsx(dataAttendance_final, 
           file = paste0(Sys.Date(), "_Processed_Attendance_Data.xlsx")
           )



# For Rakoon - Comparing our output to "Test_Merged.xlsx ========================================================

## Notes--------------------------------------------------

# There are minor differences between our output and "Test_Merged.xlsx":

# Difference #1:
# - Our output includes two columns named "timetable" and "weekend_ot"
# - The names for those columns in Test_Merged.xlsx are slightly different: "Timetable" and "weekendot"

# Difference #2: 
# - Test_Merged.xlsx contains a few cell values with tags '\r' and '\n'.
# - These tags can be inconvenient to handle; Hence this script removed such tags and are not present in the final output.

# Different #3:
# - The "Timetable" column in Test_Merged.xlsx has values that are either all uppercase or in proper title format
# - In our output, all character values are lowercase, except in the "source_file" column

## Checking if "Test_Merged.xlsx" and out output are identical ------------

# Read in test_merged.xlsx and handle the minor differences noted above

test_merged <- read_excel("Test_Merged.xlsx") %>%
  mutate(across(where(is.character), ~gsub("[\r\n]", "", .x)), # Remove any tags that may exist
         Timetable = str_to_lower(Timetable)) %>% # Convert all values in "Timetable" column to lowercase  
  rename(timetable = Timetable, # Rename column
         weekend_ot = weekendot) %>% # Rename column
  arrange(ac_no, timetable) 

# Subset our output to only include the data in Test_Merged.xlsx for comparison
dataAttendance_final_subset <- dataAttendance_final %>% 
  filter(source_file %in% unique(test_merged$source_file)) %>% 
  arrange(ac_no, timetable)

# Compare structures
str(test_merged)
str(dataAttendance_final_subset)

# Check if identical
identical(test_merged, dataAttendance_final_subset) # If TRUE, our output and "Test_Merged.xlsx" are identical
all.equal(test_merged, dataAttendance_final_subset) # If TRUE, the values in our output matches with the values in "Test_Merged.xlsx"
