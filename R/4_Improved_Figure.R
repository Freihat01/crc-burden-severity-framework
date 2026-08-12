# ==============================================================================
# R_03_3_Country_Burden_Severity_Framework.R
#
# RESULTS SECTION 3.3
# Country-level colorectal cancer burden–severity framework in 2023
#
# PRIMARY TWO-DIMENSIONAL FRAMEWORK
#
#   X-axis:
#     Age-standardized DALY rate per 100,000 population
#
#   Y-axis:
#     DALYs per incident case
#
# IMPORTANT
#   Mortality-to-incidence ratio (MIR) is NOT used in this framework.
#
# RAW INPUT FOLDERS
#   Incidence/
#   DALYs/
#
# REQUIRED RAW ESTIMATES
#   Incidence:
#     Year = 2023
#     Sex = Both
#     Age = All ages
#     Metric = Number
#
#   DALYs:
#     Year = 2023
#     Sex = Both
#     Age = All ages
#     Metric = Number
#
#   DALY rate:
#     Year = 2023
#     Sex = Both
#     Age = Age-standardized
#     Metric = Rate
#
# OUTPUTS
#
# Publication/Tables/
#   Table_3_Country_burden_severity_quadrants_2023.xlsx
#
# Publication/Figures/
#   Figure_3_Country_burden_severity_framework_2023.tiff
#
# Publication/Supplementary/
#   Table_S1_Complete_country_burden_severity_2023.xlsx
#   Table_S1_Complete_country_burden_severity_2023.csv
#
# Publication/Results_text/
#   Results_3_3_Country_burden_severity_framework_2023.txt
#
# FIGURE
#   One TIFF file at 600 dpi
#
# QUADRANTS
#
#   Q1: High burden–high severity
#   Q2: Low burden–high severity
#   Q3: Low burden–low severity
#   Q4: High burden–low severity
#
# THRESHOLDS
#   Median age-standardized DALY rate
#   Median DALYs per case
#
# The old Clean_Data folder is not used.
# ==============================================================================


# ==============================================================================
# 1. CLEAR THE R ENVIRONMENT
# ==============================================================================

rm(list = ls())

graphics.off()

cat("\014")

options(
  stringsAsFactors = FALSE,
  scipen = 999,
  dplyr.summarise.inform = FALSE
)


# ==============================================================================
# 2. INSTALL AND LOAD REQUIRED PACKAGES
# ==============================================================================

required_packages <- c(
  "tidyverse",
  "janitor",
  "openxlsx",
  "scales",
  "ggrepel",
  "glue"
)

installed_package_names <- rownames(
  installed.packages()
)

missing_packages <- setdiff(
  required_packages,
  installed_package_names
)

if (length(missing_packages) > 0) {
  
  install.packages(
    missing_packages,
    dependencies = TRUE
  )
  
}

suppressPackageStartupMessages({
  
  library(tidyverse)
  library(janitor)
  library(openxlsx)
  library(scales)
  library(ggrepel)
  library(glue)
  
})


# ==============================================================================
# 3. IDENTIFY THE MAIN PROJECT FOLDER
#
# This script works when the R project is located either in:
#
#   Colorectal cancer/
#
# or:
#
#   Colorectal cancer/R/
# ==============================================================================

current_folder <- normalizePath(
  getwd(),
  winslash = "/",
  mustWork = TRUE
)

required_folder_names <- c(
  "Incidence",
  "DALYs"
)

current_folder_has_data <- all(
  dir.exists(
    file.path(
      current_folder,
      required_folder_names
    )
  )
)

parent_folder <- normalizePath(
  file.path(
    current_folder,
    ".."
  ),
  winslash = "/",
  mustWork = TRUE
)

parent_folder_has_data <- all(
  dir.exists(
    file.path(
      parent_folder,
      required_folder_names
    )
  )
)

if (current_folder_has_data) {
  
  project_root <- current_folder
  
} else if (parent_folder_has_data) {
  
  project_root <- parent_folder
  
} else {
  
  stop(
    paste0(
      "\nThe required raw-data folders could not be found.\n\n",
      "Required folders:\n",
      paste(
        required_folder_names,
        collapse = "\n"
      ),
      "\n\nCurrent working directory:\n",
      current_folder
    ),
    call. = FALSE
  )
  
}

message("")
message("Project root:")
message(project_root)


# ==============================================================================
# 4. DEFINE INPUT AND OUTPUT FOLDERS
# ==============================================================================

incidence_folder <- file.path(
  project_root,
  "Incidence"
)

dalys_folder <- file.path(
  project_root,
  "DALYs"
)

publication_folder <- file.path(
  project_root,
  "Publication"
)

tables_folder <- file.path(
  publication_folder,
  "Tables"
)

figures_folder <- file.path(
  publication_folder,
  "Figures"
)

supplementary_folder <- file.path(
  publication_folder,
  "Supplementary"
)

results_text_folder <- file.path(
  publication_folder,
  "Results_text"
)

dir.create(
  tables_folder,
  recursive = TRUE,
  showWarnings = FALSE
)

dir.create(
  figures_folder,
  recursive = TRUE,
  showWarnings = FALSE
)

dir.create(
  supplementary_folder,
  recursive = TRUE,
  showWarnings = FALSE
)

dir.create(
  results_text_folder,
  recursive = TRUE,
  showWarnings = FALSE
)


# ==============================================================================
# 5. ANALYTICAL CONSTANTS
# ==============================================================================

analysis_year <- 2023

analysis_sex <- "Both"

expected_country_count <- 204

quadrant_levels <- c(
  "Q1: High burden–high severity",
  "Q2: Low burden–high severity",
  "Q3: Low burden–low severity",
  "Q4: High burden–low severity"
)


# ==============================================================================
# 6. GENERAL HELPER FUNCTIONS
# ==============================================================================

clean_text <- function(x) {
  
  x |>
    as.character() |>
    stringr::str_replace_all(
      "\u00A0",
      " "
    ) |>
    stringr::str_squish()
  
}


to_numeric_safely <- function(x) {
  
  readr::parse_number(
    as.character(x),
    na = c(
      "",
      "NA",
      "N/A",
      "NaN",
      "NULL"
    )
  )
  
}


find_column <- function(
    data,
    accepted_names,
    variable_name,
    required = TRUE
) {
  
  detected_columns <- intersect(
    accepted_names,
    names(data)
  )
  
  if (length(detected_columns) == 0) {
    
    if (required) {
      
      stop(
        paste0(
          "\nRequired column not found: ",
          variable_name,
          "\n\nAccepted names:\n",
          paste(
            accepted_names,
            collapse = ", "
          ),
          "\n\nDetected columns:\n",
          paste(
            names(data),
            collapse = ", "
          )
        ),
        call. = FALSE
      )
      
    }
    
    return(NULL)
    
  }
  
  detected_columns[[1]]
  
}


standardise_sex <- function(x) {
  
  value <- stringr::str_to_lower(
    clean_text(x)
  )
  
  dplyr::case_when(
    
    value %in% c(
      "both",
      "both sex",
      "both sexes"
    ) ~ "Both",
    
    value %in% c(
      "male",
      "males"
    ) ~ "Male",
    
    value %in% c(
      "female",
      "females"
    ) ~ "Female",
    
    TRUE ~ clean_text(x)
    
  )
  
}


is_number_metric <- function(x) {
  
  stringr::str_to_lower(
    clean_text(x)
  ) %in% c(
    "number",
    "count",
    "counts"
  )
  
}


is_rate_metric <- function(x) {
  
  stringr::str_to_lower(
    clean_text(x)
  ) %in% c(
    "rate",
    "age-standardized rate",
    "age-standardised rate"
  )
  
}


is_all_ages <- function(x) {
  
  stringr::str_detect(
    stringr::str_to_lower(
      clean_text(x)
    ),
    "^all ages?$"
  )
  
}


is_age_standardised <- function(x) {
  
  stringr::str_detect(
    stringr::str_to_lower(
      clean_text(x)
    ),
    "age.?standard|standardized|standardised"
  )
  
}


is_colorectal_cancer <- function(x) {
  
  stringr::str_detect(
    stringr::str_to_lower(
      clean_text(x)
    ),
    "colon and rectum cancer|colorectal cancer|colon.*rectum"
  )
  
}


format_number <- function(
    value,
    digits = 0
) {
  
  formatC(
    value,
    format = "f",
    digits = digits,
    big.mark = ","
  )
  
}


format_percent <- function(
    value,
    digits = 1
) {
  
  paste0(
    formatC(
      value,
      format = "f",
      digits = digits
    ),
    "%"
  )
  
}


# ==============================================================================
# 7. IMPORT ALL CSV FILES FROM A RAW-DATA FOLDER
# ==============================================================================

read_raw_folder <- function(
    folder_path,
    source_label
) {
  
  source_files <- list.files(
    path = folder_path,
    pattern = "\\.csv$",
    full.names = TRUE,
    recursive = TRUE,
    ignore.case = TRUE
  )
  
  # Exclude the separate 2023 age-group files from this all-ages analysis
  source_files <- source_files[
    !stringr::str_detect(
      basename(source_files),
      regex(
        "AgeGroups_2023",
        ignore_case = TRUE
      )
    )
  ]
  
  if (length(source_files) == 0) {
    
    stop(
      paste0(
        "\nNo CSV files were found in:\n",
        folder_path
      ),
      call. = FALSE
    )
    
  }
  
  message("")
  message(
    source_label,
    " files detected: ",
    length(source_files)
  )
  
  imported_files <- lapply(
    
    source_files,
    
    function(file_path) {
      
      readr::read_csv(
        file = file_path,
        show_col_types = FALSE,
        progress = FALSE,
        guess_max = 100000
      ) |>
        janitor::clean_names() |>
        mutate(
          source_file = basename(file_path)
        )
      
    }
    
  )
  
  dplyr::bind_rows(
    imported_files
  )
  
}


# ==============================================================================
# 8. STANDARDISE A RAW MEASURE DATASET
# ==============================================================================

standardise_raw_measure <- function(
    raw_data,
    measure_label
) {
  
  column_location <- find_column(
    raw_data,
    c(
      "location",
      "location_name"
    ),
    "location"
  )
  
  column_location_id <- find_column(
    raw_data,
    c(
      "location_id"
    ),
    "location ID",
    required = FALSE
  )
  
  column_year <- find_column(
    raw_data,
    c(
      "year",
      "year_id"
    ),
    "year"
  )
  
  column_sex <- find_column(
    raw_data,
    c(
      "sex",
      "sex_name"
    ),
    "sex"
  )
  
  column_age <- find_column(
    raw_data,
    c(
      "age",
      "age_name"
    ),
    "age"
  )
  
  column_metric <- find_column(
    raw_data,
    c(
      "metric",
      "metric_name"
    ),
    "metric"
  )
  
  column_value <- find_column(
    raw_data,
    c(
      "val",
      "value",
      "mean",
      "estimate"
    ),
    "central estimate"
  )
  
  column_lower <- find_column(
    raw_data,
    c(
      "lower",
      "lower_bound",
      "lower_ci",
      "lower_ui"
    ),
    "lower uncertainty bound",
    required = FALSE
  )
  
  column_upper <- find_column(
    raw_data,
    c(
      "upper",
      "upper_bound",
      "upper_ci",
      "upper_ui"
    ),
    "upper uncertainty bound",
    required = FALSE
  )
  
  column_cause <- find_column(
    raw_data,
    c(
      "cause",
      "cause_name"
    ),
    "cause",
    required = FALSE
  )
  
  standardised_data <- raw_data |>
    
    transmute(
      
      location = clean_text(
        .data[[column_location]]
      ),
      
      location_id = if (!is.null(column_location_id)) {
        
        as.integer(
          .data[[column_location_id]]
        )
        
      } else {
        
        NA_integer_
        
      },
      
      year = as.integer(
        .data[[column_year]]
      ),
      
      sex = standardise_sex(
        .data[[column_sex]]
      ),
      
      age = clean_text(
        .data[[column_age]]
      ),
      
      metric = clean_text(
        .data[[column_metric]]
      ),
      
      cause = if (!is.null(column_cause)) {
        
        clean_text(
          .data[[column_cause]]
        )
        
      } else {
        
        "Colorectal cancer"
        
      },
      
      value = to_numeric_safely(
        .data[[column_value]]
      ),
      
      lower = if (!is.null(column_lower)) {
        
        to_numeric_safely(
          .data[[column_lower]]
        )
        
      } else {
        
        NA_real_
        
      },
      
      upper = if (!is.null(column_upper)) {
        
        to_numeric_safely(
          .data[[column_upper]]
        )
        
      } else {
        
        NA_real_
        
      },
      
      source_file,
      
      measure = measure_label
      
    )
  
  if (!is.null(column_cause)) {
    
    standardised_data <- standardised_data |>
      
      filter(
        is_colorectal_cancer(cause)
      )
    
  }
  
  standardised_data
  
}


# ==============================================================================
# 9. IMPORT INCIDENCE AND DALY DATA
# ==============================================================================

incidence_raw <- read_raw_folder(
  folder_path = incidence_folder,
  source_label = "Incidence"
)

dalys_raw <- read_raw_folder(
  folder_path = dalys_folder,
  source_label = "DALYs"
)


incidence_standardised <- standardise_raw_measure(
  raw_data = incidence_raw,
  measure_label = "Incidence"
)

dalys_standardised <- standardise_raw_measure(
  raw_data = dalys_raw,
  measure_label = "DALYs"
)


# ==============================================================================
# 10. EXTRACT 2023 BOTH-SEX INCIDENCE NUMBERS
# ==============================================================================

incidence_2023 <- incidence_standardised |>
  
  filter(
    
    year == analysis_year,
    
    sex == analysis_sex,
    
    is_all_ages(age),
    
    is_number_metric(metric),
    
    !is.na(location),
    
    !is.na(value),
    
    value > 0
    
  ) |>
  
  transmute(
    
    location,
    
    location_id,
    
    incident_cases = value,
    
    incident_cases_lower = lower,
    
    incident_cases_upper = upper,
    
    incidence_source_file = source_file
    
  )


# ==============================================================================
# 11. EXTRACT 2023 BOTH-SEX DALY NUMBERS
# ==============================================================================

dalys_number_2023 <- dalys_standardised |>
  
  filter(
    
    year == analysis_year,
    
    sex == analysis_sex,
    
    is_all_ages(age),
    
    is_number_metric(metric),
    
    !is.na(location),
    
    !is.na(value),
    
    value >= 0
    
  ) |>
  
  transmute(
    
    location,
    
    location_id,
    
    dalys = value,
    
    dalys_lower = lower,
    
    dalys_upper = upper,
    
    dalys_number_source_file = source_file
    
  )


# ==============================================================================
# 12. EXTRACT 2023 BOTH-SEX AGE-STANDARDIZED DALY RATES
# ==============================================================================

dalys_rate_2023 <- dalys_standardised |>
  
  filter(
    
    year == analysis_year,
    
    sex == analysis_sex,
    
    is_age_standardised(age),
    
    is_rate_metric(metric),
    
    !is.na(location),
    
    !is.na(value),
    
    value >= 0
    
  ) |>
  
  transmute(
    
    location,
    
    location_id,
    
    daly_rate = value,
    
    daly_rate_lower = lower,
    
    daly_rate_upper = upper,
    
    dalys_rate_source_file = source_file
    
  )


# ==============================================================================
# 13. REMOVE NON-COUNTRY AGGREGATE LOCATIONS WHEN PRESENT
#
# These names are excluded only if they appear in the raw files.
# ==============================================================================

aggregate_location_names <- c(
  
  "Global",
  "World",
  "Worldwide",
  
  "High SDI",
  "High-middle SDI",
  "Middle SDI",
  "Low-middle SDI",
  "Low SDI",
  
  "Andean Latin America",
  "Australasia",
  "Caribbean",
  "Central Asia",
  "Central Europe",
  "Central Latin America",
  "Central Sub-Saharan Africa",
  "East Asia",
  "Eastern Europe",
  "Eastern Sub-Saharan Africa",
  "High-income Asia Pacific",
  "High-income North America",
  "North Africa and Middle East",
  "Oceania",
  "South Asia",
  "Southeast Asia",
  "Southern Latin America",
  "Southern Sub-Saharan Africa",
  "Tropical Latin America",
  "Western Europe",
  "Western Sub-Saharan Africa",
  
  "Africa",
  "Asia",
  "Europe",
  "Latin America and Caribbean",
  "North America"
)


incidence_2023 <- incidence_2023 |>
  
  filter(
    !location %in% aggregate_location_names
  )


dalys_number_2023 <- dalys_number_2023 |>
  
  filter(
    !location %in% aggregate_location_names
  )


dalys_rate_2023 <- dalys_rate_2023 |>
  
  filter(
    !location %in% aggregate_location_names
  )


# ==============================================================================
# 14. CHECK FOR DUPLICATES WITHIN EACH DATASET
# ==============================================================================

incidence_duplicates <- incidence_2023 |>
  
  count(
    location,
    name = "number_of_rows"
  ) |>
  
  filter(
    number_of_rows > 1
  )


dalys_number_duplicates <- dalys_number_2023 |>
  
  count(
    location,
    name = "number_of_rows"
  ) |>
  
  filter(
    number_of_rows > 1
  )


dalys_rate_duplicates <- dalys_rate_2023 |>
  
  count(
    location,
    name = "number_of_rows"
  ) |>
  
  filter(
    number_of_rows > 1
  )


if (nrow(incidence_duplicates) > 0) {
  
  print(incidence_duplicates)
  
  stop(
    "\nDuplicate country incidence estimates were detected.",
    call. = FALSE
  )
  
}


if (nrow(dalys_number_duplicates) > 0) {
  
  print(dalys_number_duplicates)
  
  stop(
    "\nDuplicate country DALY-number estimates were detected.",
    call. = FALSE
  )
  
}


if (nrow(dalys_rate_duplicates) > 0) {
  
  print(dalys_rate_duplicates)
  
  stop(
    "\nDuplicate country DALY-rate estimates were detected.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 15. MERGE INCIDENCE, DALY NUMBERS, AND DALY RATES
# ==============================================================================

country_data <- incidence_2023 |>
  
  full_join(
    dalys_number_2023,
    by = "location",
    suffix = c(
      "_incidence",
      "_dalys"
    )
  ) |>
  
  full_join(
    dalys_rate_2023,
    by = "location"
  )


# ==============================================================================
# 16. CONSOLIDATE LOCATION IDs
# ==============================================================================

country_data <- country_data |>
  
  mutate(
    
    location_id = dplyr::coalesce(
      
      location_id_incidence,
      
      location_id_dalys,
      
      location_id
      
    )
    
  ) |>
  
  select(
    
    location,
    
    location_id,
    
    incident_cases,
    
    incident_cases_lower,
    
    incident_cases_upper,
    
    dalys,
    
    dalys_lower,
    
    dalys_upper,
    
    daly_rate,
    
    daly_rate_lower,
    
    daly_rate_upper,
    
    incidence_source_file,
    
    dalys_number_source_file,
    
    dalys_rate_source_file
    
  )


# ==============================================================================
# 17. IDENTIFY INCOMPLETE COUNTRIES
# ==============================================================================

incomplete_countries <- country_data |>
  
  filter(
    
    is.na(incident_cases) |
      
      is.na(dalys) |
      
      is.na(daly_rate)
    
  )


if (nrow(incomplete_countries) > 0) {
  
  message("")
  message("Incomplete locations detected:")
  
  print(
    incomplete_countries |>
      select(
        location,
        incident_cases,
        dalys,
        daly_rate
      )
  )
  
  stop(
    paste0(
      "\nSome locations are missing incidence, DALYs, or DALY rates.",
      "\nThe framework cannot be calculated until coverage is complete."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 18. CONFIRM COUNTRY COVERAGE
# ==============================================================================

country_count <- n_distinct(
  country_data$location
)

message("")
message(
  "Countries and territories available: ",
  country_count
)

if (country_count != expected_country_count) {
  
  stop(
    paste0(
      "\nExpected ",
      expected_country_count,
      " countries and territories, but ",
      country_count,
      " were found.\n\n",
      "Do not continue until the location coverage is confirmed."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 19. CALCULATE DALYs PER INCIDENT CASE
# ==============================================================================

country_data <- country_data |>
  
  mutate(
    
    dalys_per_case = dalys / incident_cases
    
  )


invalid_ratios <- country_data |>
  
  filter(
    
    is.na(dalys_per_case) |
      
      !is.finite(dalys_per_case) |
      
      dalys_per_case < 0
    
  )


if (nrow(invalid_ratios) > 0) {
  
  print(
    invalid_ratios
  )
  
  stop(
    "\nInvalid DALYs-per-case estimates were detected.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 20. CALCULATE MEDIAN FRAMEWORK THRESHOLDS
# ==============================================================================

median_daly_rate <- median(
  country_data$daly_rate,
  na.rm = TRUE
)

median_dalys_per_case <- median(
  country_data$dalys_per_case,
  na.rm = TRUE
)

message("")
message(
  "Median age-standardized DALY rate: ",
  format_number(
    median_daly_rate,
    2
  )
)

message(
  "Median DALYs per case: ",
  format_number(
    median_dalys_per_case,
    2
  )
)


# ==============================================================================
# 21. ASSIGN THE FOUR QUADRANTS
#
# Countries equal to a median threshold are assigned to the high category.
# ==============================================================================

country_data <- country_data |>
  
  mutate(
    
    burden_category = if_else(
      daly_rate >= median_daly_rate,
      "High burden",
      "Low burden"
    ),
    
    severity_category = if_else(
      dalys_per_case >= median_dalys_per_case,
      "High severity",
      "Low severity"
    ),
    
    quadrant = case_when(
      
      burden_category == "High burden" &
        severity_category == "High severity" ~
        "Q1: High burden–high severity",
      
      burden_category == "Low burden" &
        severity_category == "High severity" ~
        "Q2: Low burden–high severity",
      
      burden_category == "Low burden" &
        severity_category == "Low severity" ~
        "Q3: Low burden–low severity",
      
      burden_category == "High burden" &
        severity_category == "Low severity" ~
        "Q4: High burden–low severity"
      
    ),
    
    quadrant = factor(
      quadrant,
      levels = quadrant_levels
    )
    
  )


# ==============================================================================
# 22. CREATE QUADRANT SUMMARY
# ==============================================================================

quadrant_summary <- country_data |>
  
  group_by(
    quadrant
  ) |>
  
  summarise(
    
    Number_of_countries = n(),
    
    Percentage_of_countries =
      100 * n() / nrow(country_data),
    
    Median_DALY_rate = median(
      daly_rate,
      na.rm = TRUE
    ),
    
    IQR_DALY_rate_lower = as.numeric(
      quantile(
        daly_rate,
        probs = 0.25,
        na.rm = TRUE,
        names = FALSE
      )
    ),
    
    IQR_DALY_rate_upper = as.numeric(
      quantile(
        daly_rate,
        probs = 0.75,
        na.rm = TRUE,
        names = FALSE
      )
    ),
    
    Median_DALYs_per_case = median(
      dalys_per_case,
      na.rm = TRUE
    ),
    
    IQR_DALYs_per_case_lower = as.numeric(
      quantile(
        dalys_per_case,
        probs = 0.25,
        na.rm = TRUE,
        names = FALSE
      )
    ),
    
    IQR_DALYs_per_case_upper = as.numeric(
      quantile(
        dalys_per_case,
        probs = 0.75,
        na.rm = TRUE,
        names = FALSE
      )
    ),
    
    .groups = "drop"
    
  ) |>
  
  arrange(
    quadrant
  )

# ==============================================================================
# 23. IDENTIFY REPRESENTATIVE COUNTRIES
#
# Representative countries are those closest to the within-quadrant medians.
# ==============================================================================

representative_countries <- country_data |>
  
  group_by(
    quadrant
  ) |>
  
  mutate(
    
    quadrant_median_rate = median(
      daly_rate,
      na.rm = TRUE
    ),
    
    quadrant_median_severity = median(
      dalys_per_case,
      na.rm = TRUE
    ),
    
    standardised_rate_distance =
      abs(daly_rate - quadrant_median_rate) /
      stats::sd(
        daly_rate,
        na.rm = TRUE
      ),
    
    standardised_severity_distance =
      abs(dalys_per_case - quadrant_median_severity) /
      stats::sd(
        dalys_per_case,
        na.rm = TRUE
      ),
    
    representative_distance =
      standardised_rate_distance +
      standardised_severity_distance
    
  ) |>
  
  arrange(
    quadrant,
    representative_distance
  ) |>
  
  slice_head(
    n = 3
  ) |>
  
  summarise(
    
    Representative_countries = paste(
      location,
      collapse = "; "
    ),
    
    .groups = "drop"
    
  )


quadrant_summary <- quadrant_summary |>
  
  left_join(
    representative_countries,
    by = "quadrant"
  )


# ==============================================================================
# 24. CREATE TABLE 3
# ==============================================================================

table_3 <- quadrant_summary |>
  
  transmute(
    
    Quadrant = as.character(
      quadrant
    ),
    
    `Countries and territories, n` =
      Number_of_countries,
    
    `Countries and territories, %` =
      round(
        Percentage_of_countries,
        1
      ),
    
    `DALY rate, median (IQR)` = paste0(
      
      format_number(
        Median_DALY_rate,
        2
      ),
      
      " (",
      
      format_number(
        IQR_DALY_rate_lower,
        2
      ),
      
      "–",
      
      format_number(
        IQR_DALY_rate_upper,
        2
      ),
      
      ")"
      
    ),
    
    `DALYs per case, median (IQR)` = paste0(
      
      format_number(
        Median_DALYs_per_case,
        2
      ),
      
      " (",
      
      format_number(
        IQR_DALYs_per_case_lower,
        2
      ),
      
      "–",
      
      format_number(
        IQR_DALYs_per_case_upper,
        2
      ),
      
      ")"
      
    ),
    
    `Representative countries` =
      Representative_countries
    
  )


# ==============================================================================
# 25. CREATE RANKING TABLES
# ==============================================================================

highest_severity <- country_data |>
  
  arrange(
    desc(dalys_per_case)
  ) |>
  
  slice_head(
    n = 10
  ) |>
  
  transmute(
    
    Rank = row_number(),
    
    Country = location,
    
    `DALYs per case` = round(
      dalys_per_case,
      2
    ),
    
    `Age-standardized DALY rate` = round(
      daly_rate,
      2
    ),
    
    Quadrant = as.character(
      quadrant
    )
    
  )


lowest_severity <- country_data |>
  
  arrange(
    dalys_per_case
  ) |>
  
  slice_head(
    n = 10
  ) |>
  
  transmute(
    
    Rank = row_number(),
    
    Country = location,
    
    `DALYs per case` = round(
      dalys_per_case,
      2
    ),
    
    `Age-standardized DALY rate` = round(
      daly_rate,
      2
    ),
    
    Quadrant = as.character(
      quadrant
    )
    
  )


highest_burden <- country_data |>
  
  arrange(
    desc(daly_rate)
  ) |>
  
  slice_head(
    n = 10
  ) |>
  
  transmute(
    
    Rank = row_number(),
    
    Country = location,
    
    `Age-standardized DALY rate` = round(
      daly_rate,
      2
    ),
    
    `DALYs per case` = round(
      dalys_per_case,
      2
    ),
    
    Quadrant = as.character(
      quadrant
    )
    
  )


lowest_burden <- country_data |>
  
  arrange(
    daly_rate
  ) |>
  
  slice_head(
    n = 10
  ) |>
  
  transmute(
    
    Rank = row_number(),
    
    Country = location,
    
    `Age-standardized DALY rate` = round(
      daly_rate,
      2
    ),
    
    `DALYs per case` = round(
      dalys_per_case,
      2
    ),
    
    Quadrant = as.character(
      quadrant
    )
    
  )


# ==============================================================================
# 26. SAVE TABLE 3
# ==============================================================================

table_3_file <- file.path(
  tables_folder,
  "Table_3_Country_burden_severity_quadrants_2023.xlsx"
)

table_workbook <- createWorkbook()

addWorksheet(
  table_workbook,
  "Table 3"
)

writeData(
  table_workbook,
  sheet = "Table 3",
  x = paste(
    "Table 3. Country distribution within the colorectal cancer",
    "burden–severity framework in 2023"
  ),
  startRow = 1,
  startCol = 1
)

writeData(
  table_workbook,
  sheet = "Table 3",
  x = table_3,
  startRow = 3,
  startCol = 1
)

title_style <- createStyle(
  fontSize = 12,
  textDecoration = "bold",
  halign = "left"
)

header_style <- createStyle(
  fontSize = 10,
  textDecoration = "bold",
  halign = "center",
  valign = "center",
  wrapText = TRUE,
  border = "Bottom"
)

body_style <- createStyle(
  fontSize = 10,
  halign = "center",
  valign = "center",
  wrapText = TRUE
)

addStyle(
  table_workbook,
  sheet = "Table 3",
  style = title_style,
  rows = 1,
  cols = 1:ncol(table_3),
  gridExpand = TRUE
)

addStyle(
  table_workbook,
  sheet = "Table 3",
  style = header_style,
  rows = 3,
  cols = 1:ncol(table_3),
  gridExpand = TRUE
)

addStyle(
  table_workbook,
  sheet = "Table 3",
  style = body_style,
  rows = 4:(nrow(table_3) + 3),
  cols = 1:ncol(table_3),
  gridExpand = TRUE
)

setColWidths(
  table_workbook,
  sheet = "Table 3",
  cols = 1,
  widths = 31
)

setColWidths(
  table_workbook,
  sheet = "Table 3",
  cols = 2:5,
  widths = 24
)

setColWidths(
  table_workbook,
  sheet = "Table 3",
  cols = 6,
  widths = 42
)

setRowHeights(
  table_workbook,
  sheet = "Table 3",
  rows = 3,
  heights = 55
)

freezePane(
  table_workbook,
  sheet = "Table 3",
  firstActiveRow = 4,
  firstActiveCol = 2
)

table_note <- paste0(
  
  "Note: The framework includes ",
  country_count,
  " countries and territories. ",
  
  "The median thresholds were ",
  format_number(
    median_daly_rate,
    2
  ),
  " DALYs per 100,000 population for age-standardized DALY rate and ",
  format_number(
    median_dalys_per_case,
    2
  ),
  " DALYs per incident case. ",
  
  "Countries equal to a threshold were assigned to the corresponding ",
  "high-burden or high-severity category."
  
)

writeData(
  table_workbook,
  sheet = "Table 3",
  x = table_note,
  startRow = nrow(table_3) + 6,
  startCol = 1,
  colNames = FALSE
)

saveWorkbook(
  table_workbook,
  table_3_file,
  overwrite = TRUE
)


# ==============================================================================
# 27. CREATE COMPLETE SUPPLEMENTARY TABLE S1
# ==============================================================================

supplementary_s1 <- country_data |>
  
  arrange(
    location
  ) |>
  
  transmute(
    
    Country = location,
    
    Location_ID = location_id,
    
    Incident_cases = incident_cases,
    
    Incident_cases_lower_95_UI =
      incident_cases_lower,
    
    Incident_cases_upper_95_UI =
      incident_cases_upper,
    
    DALYs = dalys,
    
    DALYs_lower_95_UI =
      dalys_lower,
    
    DALYs_upper_95_UI =
      dalys_upper,
    
    Age_standardized_DALY_rate =
      daly_rate,
    
    DALY_rate_lower_95_UI =
      daly_rate_lower,
    
    DALY_rate_upper_95_UI =
      daly_rate_upper,
    
    DALYs_per_case =
      dalys_per_case,
    
    Burden_category =
      burden_category,
    
    Severity_category =
      severity_category,
    
    Quadrant =
      as.character(
        quadrant
      )
    
  )


supplementary_s1_xlsx <- file.path(
  supplementary_folder,
  "Table_S1_Complete_country_burden_severity_2023.xlsx"
)

supplementary_s1_csv <- file.path(
  supplementary_folder,
  "Table_S1_Complete_country_burden_severity_2023.csv"
)


supplementary_workbook <- createWorkbook()

addWorksheet(
  supplementary_workbook,
  "Complete country dataset"
)

writeData(
  supplementary_workbook,
  sheet = "Complete country dataset",
  x = supplementary_s1
)

addWorksheet(
  supplementary_workbook,
  "Highest severity"
)

writeData(
  supplementary_workbook,
  sheet = "Highest severity",
  x = highest_severity
)

addWorksheet(
  supplementary_workbook,
  "Lowest severity"
)

writeData(
  supplementary_workbook,
  sheet = "Lowest severity",
  x = lowest_severity
)

addWorksheet(
  supplementary_workbook,
  "Highest burden"
)

writeData(
  supplementary_workbook,
  sheet = "Highest burden",
  x = highest_burden
)

addWorksheet(
  supplementary_workbook,
  "Lowest burden"
)

writeData(
  supplementary_workbook,
  sheet = "Lowest burden",
  x = lowest_burden
)


supplementary_header_style <- createStyle(
  fontSize = 10,
  textDecoration = "bold",
  halign = "center",
  valign = "center",
  wrapText = TRUE,
  border = "Bottom"
)


for (
  sheet_name in names(
    supplementary_workbook
  )
) {
  
  sheet_data <- switch(
    
    sheet_name,
    
    "Complete country dataset" =
      supplementary_s1,
    
    "Highest severity" =
      highest_severity,
    
    "Lowest severity" =
      lowest_severity,
    
    "Highest burden" =
      highest_burden,
    
    "Lowest burden" =
      lowest_burden
    
  )
  
  addStyle(
    supplementary_workbook,
    sheet = sheet_name,
    style = supplementary_header_style,
    rows = 1,
    cols = 1:ncol(sheet_data),
    gridExpand = TRUE
  )
  
  setColWidths(
    supplementary_workbook,
    sheet = sheet_name,
    cols = 1:ncol(sheet_data),
    widths = "auto"
  )
  
  freezePane(
    supplementary_workbook,
    sheet = sheet_name,
    firstRow = TRUE
  )
  
}


saveWorkbook(
  supplementary_workbook,
  supplementary_s1_xlsx,
  overwrite = TRUE
)


readr::write_csv(
  supplementary_s1,
  supplementary_s1_csv
)


# ==============================================================================
# 28. SELECT OBJECTIVE COUNTRY LABELS FOR FIGURE 3
#
# Three countries are labelled from each quadrant:
#   - one with the greatest severity
#   - one with the greatest burden
#   - one farthest from the two median thresholds
# ==============================================================================

label_candidates_extremes <- country_data |>
  
  group_by(
    quadrant
  ) |>
  
  summarise(
    
    highest_severity_country =
      location[
        which.max(
          dalys_per_case
        )
      ],
    
    highest_burden_country =
      location[
        which.max(
          daly_rate
        )
      ],
    
    .groups = "drop"
    
  ) |>
  
  pivot_longer(
    
    cols = c(
      highest_severity_country,
      highest_burden_country
    ),
    
    names_to = "label_reason",
    
    values_to = "location"
    
  ) |>
  
  distinct(
    location,
    .keep_all = TRUE
  )


country_data <- country_data |>
  
  mutate(
    
    standardised_burden_distance =
      abs(
        daly_rate -
          median_daly_rate
      ) /
      stats::sd(
        daly_rate,
        na.rm = TRUE
      ),
    
    standardised_severity_distance =
      abs(
        dalys_per_case -
          median_dalys_per_case
      ) /
      stats::sd(
        dalys_per_case,
        na.rm = TRUE
      ),
    
    framework_distance =
      sqrt(
        standardised_burden_distance^2 +
          standardised_severity_distance^2
      )
    
  )


label_candidates_distance <- country_data |>
  
  group_by(
    quadrant
  ) |>
  
  slice_max(
    order_by = framework_distance,
    n = 1,
    with_ties = FALSE
  ) |>
  
  ungroup() |>
  
  transmute(
    location,
    label_reason = "greatest_framework_distance"
  )


label_locations <- bind_rows(
  
  label_candidates_extremes |>
    select(
      location,
      label_reason
    ),
  
  label_candidates_distance
  
) |>
  
  distinct(
    location,
    .keep_all = TRUE
  ) |>
  
  pull(
    location
  )


figure_label_data <- country_data |>
  
  filter(
    location %in% label_locations
  )


# ==============================================================================
# 29. CREATE QUADRANT BACKGROUND AREAS
# ==============================================================================

maximum_daly_rate <- max(
  country_data$daly_rate,
  na.rm = TRUE
)

maximum_dalys_per_case <- max(
  country_data$dalys_per_case,
  na.rm = TRUE
)

minimum_daly_rate <- min(
  country_data$daly_rate,
  na.rm = TRUE
)

minimum_dalys_per_case <- min(
  country_data$dalys_per_case,
  na.rm = TRUE
)


quadrant_background <- tibble(
  
  xmin = c(
    median_daly_rate,
    minimum_daly_rate,
    minimum_daly_rate,
    median_daly_rate
  ),
  
  xmax = c(
    maximum_daly_rate,
    median_daly_rate,
    median_daly_rate,
    maximum_daly_rate
  ),
  
  ymin = c(
    median_dalys_per_case,
    median_dalys_per_case,
    minimum_dalys_per_case,
    minimum_dalys_per_case
  ),
  
  ymax = c(
    maximum_dalys_per_case,
    maximum_dalys_per_case,
    median_dalys_per_case,
    median_dalys_per_case
  ),
  
  quadrant = factor(
    quadrant_levels,
    levels = quadrant_levels
  )
  
)


# ==============================================================================
# 30. CREATE FIGURE 3
#
# VISUALIZATION ONLY
#   - Uses the existing analytical objects without changing the analysis.
#   - Limits plotted country labels to two per quadrant for readability.
#   - Uses compact quadrant identifiers inside the plot; full descriptions
#     remain in the legend.
# ==============================================================================

figure_label_plot_data <- figure_label_data |>
  
  group_by(
    quadrant
  ) |>
  
  arrange(
    desc(framework_distance),
    desc(dalys_per_case),
    .by_group = TRUE
  ) |>
  
  slice_head(
    n = 2
  ) |>
  
  ungroup()


x_range_figure <- maximum_daly_rate - minimum_daly_rate

y_range_figure <- maximum_dalys_per_case - minimum_dalys_per_case


quadrant_label_data <- tibble(
  
  x = c(
    maximum_daly_rate - 0.045 * x_range_figure,
    minimum_daly_rate + 0.035 * x_range_figure,
    minimum_daly_rate + 0.035 * x_range_figure,
    maximum_daly_rate - 0.045 * x_range_figure
  ),
  
  y = c(
    maximum_dalys_per_case - 0.055 * y_range_figure,
    maximum_dalys_per_case - 0.055 * y_range_figure,
    minimum_dalys_per_case + 0.055 * y_range_figure,
    minimum_dalys_per_case + 0.055 * y_range_figure
  ),
  
  label = c(
    "Q1",
    "Q2",
    "Q3",
    "Q4"
  ),
  
  hjust = c(
    1,
    0,
    0,
    1
  ),
  
  vjust = c(
    1,
    1,
    0,
    0
  )
  
)


figure_3 <- ggplot() +
  
  geom_rect(
    
    data = quadrant_background,
    
    aes(
      xmin = xmin,
      xmax = xmax,
      ymin = ymin,
      ymax = ymax,
      fill = quadrant
    ),
    
    alpha = 0.065,
    
    colour = NA,
    
    inherit.aes = FALSE
    
  ) +
  
  geom_vline(
    
    xintercept = median_daly_rate,
    
    linetype = "dashed",
    
    linewidth = 0.75,
    
    colour = "grey30"
    
  ) +
  
  geom_hline(
    
    yintercept = median_dalys_per_case,
    
    linetype = "dashed",
    
    linewidth = 0.75,
    
    colour = "grey30"
    
  ) +
  
  geom_point(
    
    data = country_data,
    
    aes(
      x = daly_rate,
      y = dalys_per_case,
      colour = quadrant
    ),
    
    size = 3.0,
    
    alpha = 0.78
    
  ) +
  
  ggrepel::geom_label_repel(
    
    data = figure_label_plot_data,
    
    aes(
      x = daly_rate,
      y = dalys_per_case,
      label = location,
      colour = quadrant
    ),
    
    size = 3.45,
    
    fontface = "bold",
    
    fill = scales::alpha(
      "white",
      0.96
    ),
    
    label.size = 0.18,
    
    label.padding = grid::unit(
      0.14,
      "lines"
    ),
    
    box.padding = 0.80,
    
    point.padding = 0.40,
    
    force = 4.0,
    
    force_pull = 0.10,
    
    min.segment.length = 0,
    
    segment.colour = "grey50",
    
    segment.linewidth = 0.38,
    
    max.overlaps = Inf,
    
    seed = 2023,
    
    show.legend = FALSE
    
  ) +
  
  geom_text(
    
    data = quadrant_label_data,
    
    aes(
      x = x,
      y = y,
      label = label,
      hjust = hjust,
      vjust = vjust
    ),
    
    inherit.aes = FALSE,
    
    size = 7.2,
    
    fontface = "bold",
    
    colour = "grey35",
    
    alpha = 0.75
    
  ) +
  
  scale_color_manual(
    
    values = c(
      "Q1: High burden–high severity" = "#C0392B",
      "Q2: Low burden–high severity" = "#E67E22",
      "Q3: Low burden–low severity" = "#2E86C1",
      "Q4: High burden–low severity" = "#239B56"
    ),
    
    drop = FALSE
    
  ) +
  
  scale_fill_manual(
    
    values = c(
      "Q1: High burden–high severity" = "#C0392B",
      "Q2: Low burden–high severity" = "#E67E22",
      "Q3: Low burden–low severity" = "#2E86C1",
      "Q4: High burden–low severity" = "#239B56"
    ),
    
    drop = FALSE
    
  ) +
  
  scale_x_continuous(
    
    name = "Age-standardized DALY rate per 100,000 population",
    
    labels = scales::label_number(
      accuracy = 1
    ),
    
    breaks = scales::breaks_pretty(
      n = 6
    ),
    
    expand = expansion(
      mult = c(
        0.055,
        0.13
      )
    )
    
  ) +
  
  scale_y_continuous(
    
    name = "DALYs per incident case",
    
    labels = scales::label_number(
      accuracy = 0.1
    ),
    
    breaks = scales::breaks_pretty(
      n = 6
    ),
    
    expand = expansion(
      mult = c(
        0.075,
        0.15
      )
    )
    
  ) +
  
  labs(
    
    title = "Country-level colorectal cancer burden–severity framework in 2023",
    
    subtitle = paste0(
      
      "Median thresholds: DALY rate = ",
      
      format_number(
        median_daly_rate,
        2
      ),
      
      " per 100,000; DALYs per case = ",
      
      format_number(
        median_dalys_per_case,
        2
      )
      
    ),
    
    colour = NULL,
    
    fill = NULL,
    
    caption = paste0(
      
      "Each point represents one of ",
      
      country_count,
      
      " countries or territories. ",
      
      "Population burden is shown on the x-axis and average health loss ",
      
      "per incident case on the y-axis."
      
    )
    
  ) +
  
  guides(
    
    fill = "none",
    
    colour = guide_legend(
      
      nrow = 2,
      
      byrow = TRUE,
      
      override.aes = list(
        size = 4.2,
        alpha = 1
      )
      
    )
    
  ) +
  
  coord_cartesian(
    clip = "off"
  ) +
  
  theme_classic(
    base_size = 13
  ) +
  
  theme(
    
    plot.title = element_text(
      face = "bold",
      size = 18,
      hjust = 0.5,
      margin = margin(
        b = 5
      )
    ),
    
    plot.subtitle = element_text(
      size = 11.5,
      hjust = 0.5,
      margin = margin(
        b = 12
      )
    ),
    
    axis.title.x = element_text(
      face = "bold",
      size = 12.5,
      margin = margin(
        t = 10
      )
    ),
    
    axis.title.y = element_text(
      face = "bold",
      size = 12.5,
      margin = margin(
        r = 10
      )
    ),
    
    axis.text = element_text(
      size = 11,
      colour = "grey15"
    ),
    
    axis.line = element_line(
      linewidth = 0.7,
      colour = "grey25"
    ),
    
    axis.ticks = element_line(
      linewidth = 0.55,
      colour = "grey25"
    ),
    
    panel.grid.major = element_line(
      linewidth = 0.28,
      linetype = "dotted",
      colour = "grey75"
    ),
    
    panel.grid.minor = element_blank(),
    
    legend.position = "bottom",
    
    legend.justification = "center",
    
    legend.text = element_text(
      size = 10.3
    ),
    
    legend.key.width = grid::unit(
      0.9,
      "cm"
    ),
    
    legend.spacing.x = grid::unit(
      0.25,
      "cm"
    ),
    
    plot.caption = element_text(
      size = 9.4,
      hjust = 0,
      colour = "grey30",
      margin = margin(
        t = 10
      )
    ),
    
    plot.margin = margin(
      18,
      38,
      16,
      20
    )
    
  )


# ==============================================================================
# 31. SAVE FIGURE 3 AS ONE TIFF AT 600 DPI
# ==============================================================================

figure_3_file <- file.path(
  figures_folder,
  "Figure_3_Country_burden_severity_framework_2023.tiff"
)

ggsave(
  filename = figure_3_file,
  plot = figure_3,
  device = "tiff",
  width = 14,
  height = 10.5,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white",
  limitsize = FALSE
)


# ==============================================================================
# 32. EXTRACT QUADRANT COUNTS FOR RESULTS TEXT
# ==============================================================================

q1_results <- quadrant_summary |>
  
  filter(
    quadrant ==
      "Q1: High burden–high severity"
  )


q2_results <- quadrant_summary |>
  
  filter(
    quadrant ==
      "Q2: Low burden–high severity"
  )


q3_results <- quadrant_summary |>
  
  filter(
    quadrant ==
      "Q3: Low burden–low severity"
  )


q4_results <- quadrant_summary |>
  
  filter(
    quadrant ==
      "Q4: High burden–low severity"
  )


# ==============================================================================
# 33. CREATE COUNTRY LISTS FOR RESULTS TEXT
# ==============================================================================

q1_examples <- country_data |>
  
  filter(
    quadrant ==
      "Q1: High burden–high severity"
  ) |>
  
  arrange(
    desc(
      framework_distance
    )
  ) |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    location
  ) |>
  
  paste(
    collapse = ", "
  )


q2_examples <- country_data |>
  
  filter(
    quadrant ==
      "Q2: Low burden–high severity"
  ) |>
  
  arrange(
    desc(
      framework_distance
    )
  ) |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    location
  ) |>
  
  paste(
    collapse = ", "
  )


q3_examples <- country_data |>
  
  filter(
    quadrant ==
      "Q3: Low burden–low severity"
  ) |>
  
  arrange(
    desc(
      framework_distance
    )
  ) |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    location
  ) |>
  
  paste(
    collapse = ", "
  )


q4_examples <- country_data |>
  
  filter(
    quadrant ==
      "Q4: High burden–low severity"
  ) |>
  
  arrange(
    desc(
      framework_distance
    )
  ) |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    location
  ) |>
  
  paste(
    collapse = ", "
  )


highest_severity_names <- highest_severity |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    Country
  ) |>
  
  paste(
    collapse = ", "
  )


lowest_severity_names <- lowest_severity |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    Country
  ) |>
  
  paste(
    collapse = ", "
  )



# ==============================================================================
# 33. CREATE COUNTRY LISTS FOR RESULTS TEXT
# ==============================================================================

q1_examples <- country_data |>
  
  filter(
    quadrant ==
      "Q1: High burden–high severity"
  ) |>
  
  arrange(
    desc(
      framework_distance
    )
  ) |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    location
  ) |>
  
  paste(
    collapse = ", "
  )


q2_examples <- country_data |>
  
  filter(
    quadrant ==
      "Q2: Low burden–high severity"
  ) |>
  
  arrange(
    desc(
      framework_distance
    )
  ) |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    location
  ) |>
  
  paste(
    collapse = ", "
  )


q3_examples <- country_data |>
  
  filter(
    quadrant ==
      "Q3: Low burden–low severity"
  ) |>
  
  arrange(
    desc(
      framework_distance
    )
  ) |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    location
  ) |>
  
  paste(
    collapse = ", "
  )


q4_examples <- country_data |>
  
  filter(
    quadrant ==
      "Q4: High burden–low severity"
  ) |>
  
  arrange(
    desc(
      framework_distance
    )
  ) |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    location
  ) |>
  
  paste(
    collapse = ", "
  )


highest_severity_names <- highest_severity |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    Country
  ) |>
  
  paste(
    collapse = ", "
  )


lowest_severity_names <- lowest_severity |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    Country
  ) |>
  
  paste(
    collapse = ", "
  )


# ==============================================================================
# 34. GENERATE RESULTS TEXT
# ==============================================================================

results_text <- glue::glue(
  "3.3 Country-level burden–severity framework

The country-level burden–severity framework included {country_count} countries and territories. The median age-standardized colorectal cancer DALY rate was {format_number(median_daly_rate, 2)} per 100,000 population, while the median DALYs-per-case value was {format_number(median_dalys_per_case, 2)}. These thresholds classified countries according to population burden, represented by the age-standardized DALY rate, and average severity, represented by DALYs per incident case (Figure 3).

The high-burden–high-severity quadrant (Q1) contained {q1_results$Number_of_countries} countries and territories ({format_number(q1_results$Percentage_of_countries, 1)}%). Countries with prominent Q1 profiles included {q1_examples}. The low-burden–high-severity quadrant (Q2) included {q2_results$Number_of_countries} countries and territories ({format_number(q2_results$Percentage_of_countries, 1)}%), including {q2_examples}. Although these countries had comparatively low population-level DALY rates, their high DALYs-per-case values indicated substantial average health loss among individuals who developed colorectal cancer.

The low-burden–low-severity quadrant (Q3) comprised {q3_results$Number_of_countries} countries and territories ({format_number(q3_results$Percentage_of_countries, 1)}%), with prominent examples including {q3_examples}. The high-burden–low-severity quadrant (Q4) contained {q4_results$Number_of_countries} countries and territories ({format_number(q4_results$Percentage_of_countries, 1)}%), including {q4_examples}. These countries experienced comparatively high population-level DALY rates despite lower average health loss per incident case (Table 3; Figure 3).

The countries with the highest DALYs-per-case values included {highest_severity_names}, whereas the lowest values were observed in {lowest_severity_names}. Complete country-level incidence estimates, DALYs, age-standardized DALY rates, DALYs per case, and quadrant assignments are provided in Supplementary Table S1.

The marked dispersion of countries across the four quadrants demonstrates that population-level burden and case-level severity represent related but distinct dimensions of colorectal cancer burden. Countries with similar age-standardized DALY rates could experience substantially different DALYs per case, while countries with comparable severity could differ markedly in population burden."
)


# ==============================================================================
# 35. SAVE RESULTS TEXT
# ==============================================================================

results_text_file <- file.path(
  results_text_folder,
  "Results_3_3_Country_burden_severity_framework_2023.txt"
)

writeLines(
  text = results_text,
  con = results_text_file,
  useBytes = TRUE
)


# ==============================================================================
# 36. DISPLAY RESULTS
# ==============================================================================

message("")
message("Table 3:")

print(
  table_3
)

message("")
message("Highest DALYs per case:")

print(
  highest_severity
)

message("")
message("Lowest DALYs per case:")

print(
  lowest_severity
)

message("")
message("Results text:")

cat(
  results_text
)


# ==============================================================================
# 37. COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("RESULTS SECTION 3.3 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message("Framework:")
message(
  "X-axis = age-standardized DALY rate"
)
message(
  "Y-axis = DALYs per incident case"
)
message(
  "MIR was not used."
)

message("")
message(
  "Countries and territories included: ",
  country_count
)

message("")
message("Median DALY-rate threshold:")
message(
  format_number(
    median_daly_rate,
    2
  )
)

message("")
message("Median DALYs-per-case threshold:")
message(
  format_number(
    median_dalys_per_case,
    2
  )
)

message("")
message("Table 3:")
message(table_3_file)

message("")
message("Figure 3 — one TIFF at 600 dpi:")
message(figure_3_file)

message("")
message("Supplementary Table S1:")
message(supplementary_s1_xlsx)

message("")
message("Results text:")
message(results_text_file)

message("")
message("The old Clean_Data folder was not used.")
message("==============================================================")