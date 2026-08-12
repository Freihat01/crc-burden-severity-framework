# ==============================================================================
# R_03_4_Country_Severity_Decomposition.R
#
# RESULTS SECTION 3.4
# Mortality-driven versus disability-driven colorectal cancer severity in 2023
#
# PURPOSE
#   Explain why countries with similar DALYs per case may have different
#   underlying severity profiles.
#
# COUNTRY-LEVEL INDICATORS
#   DALYs per case
#   YLLs per case
#   YLDs per case
#   YLL:YLD ratio
#   YLL contribution (%)
#   YLD contribution (%)
#
# INPUT FOLDERS
#   Incidence/
#   DALYs/
#   YLLs/
#   YLDs/
#
# OUTPUTS
#
# Publication/Tables/
#   Table_4_Country_CRC_severity_components_2023.xlsx
#
# Publication/Figures/
#   Figure_4_Country_CRC_severity_components_2023.tiff
#
# Publication/Supplementary/
#   Table_S3_Complete_country_severity_decomposition_2023.xlsx
#   Table_S3_Complete_country_severity_decomposition_2023.csv
#
# Publication/Results_text/
#   Results_3_4_Country_severity_decomposition_2023.txt
#
# FIGURE
#   One TIFF at 600 dpi
#
# CLASSIFICATION
#   Countries are classified relative to the international distribution of
#   YLD contribution:
#
#   Mortality-dominant:
#       YLD contribution below the first tertile
#
#   Mixed severity:
#       YLD contribution between the first and second tertiles
#
#   Disability-enriched:
#       YLD contribution above the second tertile
#
# IMPORTANT
#   These are relative country-level categories. A country classified as
#   disability-enriched may still have most of its DALYs attributable to YLLs.
#
# The old Clean_Data folder is not used.
# ==============================================================================


# ==============================================================================
# 1. CLEAR R ENVIRONMENT
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
# 3. IDENTIFY PROJECT ROOT
#
# This works whether the R project is located in:
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
  "DALYs",
  "YLLs",
  "YLDs"
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

ylls_folder <- file.path(
  project_root,
  "YLLs"
)

ylds_folder <- file.path(
  project_root,
  "YLDs"
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

severity_class_levels <- c(
  "Mortality-dominant",
  "Mixed severity",
  "Disability-enriched"
)


# ==============================================================================
# 6. HELPER FUNCTIONS
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


is_all_ages <- function(x) {
  
  stringr::str_detect(
    stringr::str_to_lower(
      clean_text(x)
    ),
    "^all ages?$"
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
    digits = 2
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
# 7. IMPORT ALL CSV FILES FROM ONE FOLDER
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
# 8. STANDARDISE AND EXTRACT ONE MEASURE
# ==============================================================================

extract_country_measure <- function(
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
  
  extracted_data <- raw_data |>
    
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
      
      estimate = to_numeric_safely(
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
      
      source_file
      
    )
  
  if (!is.null(column_cause)) {
    
    extracted_data <- extracted_data |>
      
      filter(
        is_colorectal_cancer(cause)
      )
    
  }
  
  extracted_data <- extracted_data |>
    
    filter(
      
      year == analysis_year,
      
      sex == analysis_sex,
      
      is_all_ages(age),
      
      is_number_metric(metric),
      
      !is.na(location),
      
      !is.na(estimate),
      
      estimate >= 0
      
    ) |>
    
    filter(
      
      !stringr::str_to_lower(location) %in% c(
        "global",
        "world",
        "worldwide"
      )
      
    ) |>
    
    transmute(
      
      location,
      
      location_id,
      
      measure = measure_label,
      
      estimate,
      
      lower,
      
      upper,
      
      source_file
      
    )
  
  duplicate_rows <- extracted_data |>
    
    count(
      location,
      name = "number_of_rows"
    ) |>
    
    filter(
      number_of_rows > 1
    )
  
  if (nrow(duplicate_rows) > 0) {
    
    print(
      duplicate_rows
    )
    
    stop(
      paste0(
        "\nDuplicate country estimates were detected for ",
        measure_label,
        "."
      ),
      call. = FALSE
    )
    
  }
  
  extracted_data
  
}


# ==============================================================================
# 9. IMPORT THE FOUR REQUIRED MEASURES
# ==============================================================================

incidence_raw <- read_raw_folder(
  folder_path = incidence_folder,
  source_label = "Incidence"
)

dalys_raw <- read_raw_folder(
  folder_path = dalys_folder,
  source_label = "DALYs"
)

ylls_raw <- read_raw_folder(
  folder_path = ylls_folder,
  source_label = "YLLs"
)

ylds_raw <- read_raw_folder(
  folder_path = ylds_folder,
  source_label = "YLDs"
)


incidence_data <- extract_country_measure(
  raw_data = incidence_raw,
  measure_label = "Incidence"
)

dalys_data <- extract_country_measure(
  raw_data = dalys_raw,
  measure_label = "DALYs"
)

ylls_data <- extract_country_measure(
  raw_data = ylls_raw,
  measure_label = "YLLs"
)

ylds_data <- extract_country_measure(
  raw_data = ylds_raw,
  measure_label = "YLDs"
)


# ==============================================================================
# 10. IDENTIFY COMMON COUNTRY COVERAGE
# ==============================================================================

common_locations <- Reduce(
  
  intersect,
  
  list(
    
    unique(
      incidence_data$location
    ),
    
    unique(
      dalys_data$location
    ),
    
    unique(
      ylls_data$location
    ),
    
    unique(
      ylds_data$location
    )
    
  )
  
)

country_count <- length(
  common_locations
)

message("")
message(
  "Common countries and territories: ",
  country_count
)

if (country_count != expected_country_count) {
  
  stop(
    paste0(
      "\nExpected ",
      expected_country_count,
      " countries and territories, but ",
      country_count,
      " were found consistently across all four measures."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 11. RETAIN THE COMMON COUNTRY SET
# ==============================================================================

incidence_data <- incidence_data |>
  
  filter(
    location %in% common_locations
  )


dalys_data <- dalys_data |>
  
  filter(
    location %in% common_locations
  )


ylls_data <- ylls_data |>
  
  filter(
    location %in% common_locations
  )


ylds_data <- ylds_data |>
  
  filter(
    location %in% common_locations
  )


# ==============================================================================
# 12. PREPARE EACH MEASURE FOR MERGING
# ==============================================================================

incidence_wide <- incidence_data |>
  
  transmute(
    
    location,
    
    location_id_incidence = location_id,
    
    incident_cases = estimate,
    
    incident_cases_lower = lower,
    
    incident_cases_upper = upper
    
  )


dalys_wide <- dalys_data |>
  
  transmute(
    
    location,
    
    location_id_dalys = location_id,
    
    dalys = estimate,
    
    dalys_lower = lower,
    
    dalys_upper = upper
    
  )


ylls_wide <- ylls_data |>
  
  transmute(
    
    location,
    
    location_id_ylls = location_id,
    
    ylls = estimate,
    
    ylls_lower = lower,
    
    ylls_upper = upper
    
  )


ylds_wide <- ylds_data |>
  
  transmute(
    
    location,
    
    location_id_ylds = location_id,
    
    ylds = estimate,
    
    ylds_lower = lower,
    
    ylds_upper = upper
    
  )


# ==============================================================================
# 13. MERGE ALL FOUR MEASURES
# ==============================================================================

country_severity <- incidence_wide |>
  
  inner_join(
    dalys_wide,
    by = "location"
  ) |>
  
  inner_join(
    ylls_wide,
    by = "location"
  ) |>
  
  inner_join(
    ylds_wide,
    by = "location"
  ) |>
  
  mutate(
    
    location_id = dplyr::coalesce(
      
      location_id_incidence,
      
      location_id_dalys,
      
      location_id_ylls,
      
      location_id_ylds
      
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
    
    ylls,
    
    ylls_lower,
    
    ylls_upper,
    
    ylds,
    
    ylds_lower,
    
    ylds_upper
    
  )


# ==============================================================================
# 14. CALCULATE COUNTRY-LEVEL SEVERITY COMPONENTS
# ==============================================================================

country_severity <- country_severity |>
  
  mutate(
    
    dalys_per_case =
      dalys / incident_cases,
    
    ylls_per_case =
      ylls / incident_cases,
    
    ylds_per_case =
      ylds / incident_cases,
    
    yll_yld_ratio =
      if_else(
        ylds > 0,
        ylls / ylds,
        NA_real_
      ),
    
    yll_percent =
      100 * ylls / dalys,
    
    yld_percent =
      100 * ylds / dalys,
    
    component_sum =
      ylls + ylds,
    
    daly_component_difference =
      dalys - component_sum,
    
    daly_component_relative_difference =
      100 *
      abs(
        daly_component_difference
      ) /
      dalys
    
  )


# ==============================================================================
# 15. CHECK INVALID RESULTS
# ==============================================================================

invalid_results <- country_severity |>
  
  filter(
    
    incident_cases <= 0 |
      
      is.na(dalys_per_case) |
      
      is.na(ylls_per_case) |
      
      is.na(ylds_per_case) |
      
      !is.finite(dalys_per_case) |
      
      !is.finite(ylls_per_case) |
      
      !is.finite(ylds_per_case) |
      
      yll_percent < 0 |
      
      yld_percent < 0
    
  )


if (nrow(invalid_results) > 0) {
  
  print(
    invalid_results
  )
  
  stop(
    "\nInvalid severity-component estimates were detected.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 16. CHECK DALY DECOMPOSITION CONSISTENCY
#
# DALYs should approximately equal YLLs + YLDs.
# Small differences may occur because estimates were generated independently.
# ==============================================================================

maximum_component_difference <- max(
  country_severity$daly_component_relative_difference,
  na.rm = TRUE
)

median_component_difference <- median(
  country_severity$daly_component_relative_difference,
  na.rm = TRUE
)

message("")
message(
  "Median relative DALY decomposition difference: ",
  format_number(
    median_component_difference,
    4
  ),
  "%"
)

message(
  "Maximum relative DALY decomposition difference: ",
  format_number(
    maximum_component_difference,
    4
  ),
  "%"
)

if (maximum_component_difference > 2) {
  
  warning(
    paste0(
      "At least one country has a DALY versus YLL + YLD difference ",
      "greater than 2%. Review the source files before final publication."
    )
  )
  
}


# ==============================================================================
# 17. DETERMINE RELATIVE CLASSIFICATION THRESHOLDS
#
# Tertiles of country-level YLD contribution are used.
# ==============================================================================

yld_first_tertile <- as.numeric(
  quantile(
    country_severity$yld_percent,
    probs = 1 / 3,
    na.rm = TRUE,
    names = FALSE
  )
)

yld_second_tertile <- as.numeric(
  quantile(
    country_severity$yld_percent,
    probs = 2 / 3,
    na.rm = TRUE,
    names = FALSE
  )
)

message("")
message(
  "First tertile of YLD contribution: ",
  format_number(
    yld_first_tertile,
    2
  ),
  "%"
)

message(
  "Second tertile of YLD contribution: ",
  format_number(
    yld_second_tertile,
    2
  ),
  "%"
)


# ==============================================================================
# 18. CLASSIFY COUNTRY SEVERITY PROFILES
# ==============================================================================

country_severity <- country_severity |>
  
  mutate(
    
    severity_profile = case_when(
      
      yld_percent < yld_first_tertile ~
        "Mortality-dominant",
      
      yld_percent >= yld_first_tertile &
        yld_percent < yld_second_tertile ~
        "Mixed severity",
      
      yld_percent >= yld_second_tertile ~
        "Disability-enriched"
      
    ),
    
    severity_profile = factor(
      severity_profile,
      levels = severity_class_levels
    )
    
  )


# ==============================================================================
# 19. CREATE CLASSIFICATION SUMMARY
# ==============================================================================

profile_summary <- country_severity |>
  
  group_by(
    severity_profile
  ) |>
  
  summarise(
    
    Number_of_countries = n(),
    
    Percentage_of_countries =
      100 * n() / nrow(country_severity),
    
    Median_DALYs_per_case = median(
      dalys_per_case,
      na.rm = TRUE
    ),
    
    Median_YLLs_per_case = median(
      ylls_per_case,
      na.rm = TRUE
    ),
    
    Median_YLDs_per_case = median(
      ylds_per_case,
      na.rm = TRUE
    ),
    
    Median_YLL_percent = median(
      yll_percent,
      na.rm = TRUE
    ),
    
    Median_YLD_percent = median(
      yld_percent,
      na.rm = TRUE
    ),
    
    Median_YLL_YLD_ratio = median(
      yll_yld_ratio,
      na.rm = TRUE
    ),
    
    .groups = "drop"
    
  ) |>
  
  arrange(
    severity_profile
  )


# ==============================================================================
# 20. CREATE COUNTRY RANKINGS
# ==============================================================================

highest_ylls_per_case <- country_severity |>
  
  arrange(
    desc(
      ylls_per_case
    )
  ) |>
  
  slice_head(
    n = 10
  ) |>
  
  transmute(
    
    Rank = row_number(),
    
    Country = location,
    
    `YLLs per case` = round(
      ylls_per_case,
      2
    ),
    
    `YLDs per case` = round(
      ylds_per_case,
      2
    ),
    
    `DALYs per case` = round(
      dalys_per_case,
      2
    ),
    
    `YLL contribution (%)` = round(
      yll_percent,
      2
    ),
    
    `YLD contribution (%)` = round(
      yld_percent,
      2
    ),
    
    `Severity profile` =
      as.character(
        severity_profile
      )
    
  )


highest_ylds_per_case <- country_severity |>
  
  arrange(
    desc(
      ylds_per_case
    )
  ) |>
  
  slice_head(
    n = 10
  ) |>
  
  transmute(
    
    Rank = row_number(),
    
    Country = location,
    
    `YLDs per case` = round(
      ylds_per_case,
      2
    ),
    
    `YLLs per case` = round(
      ylls_per_case,
      2
    ),
    
    `DALYs per case` = round(
      dalys_per_case,
      2
    ),
    
    `YLD contribution (%)` = round(
      yld_percent,
      2
    ),
    
    `YLL contribution (%)` = round(
      yll_percent,
      2
    ),
    
    `Severity profile` =
      as.character(
        severity_profile
      )
    
  )


highest_yld_contribution <- country_severity |>
  
  arrange(
    desc(
      yld_percent
    )
  ) |>
  
  slice_head(
    n = 10
  ) |>
  
  transmute(
    
    Rank = row_number(),
    
    Country = location,
    
    `YLD contribution (%)` = round(
      yld_percent,
      2
    ),
    
    `YLL contribution (%)` = round(
      yll_percent,
      2
    ),
    
    `YLDs per case` = round(
      ylds_per_case,
      2
    ),
    
    `YLLs per case` = round(
      ylls_per_case,
      2
    ),
    
    `Severity profile` =
      as.character(
        severity_profile
      )
    
  )


# ==============================================================================
# 21. CREATE TABLE 4
#
# The workbook contains:
#   1. Profile summary
#   2. Highest YLLs per case
#   3. Highest YLDs per case
# ==============================================================================

table_4_file <- file.path(
  tables_folder,
  "Table_4_Country_CRC_severity_components_2023.xlsx"
)

table_workbook <- createWorkbook()


# ------------------------------------------------------------------------------
# Sheet 1: profile summary
# ------------------------------------------------------------------------------

addWorksheet(
  table_workbook,
  "Profile summary"
)

table_4_summary <- profile_summary |>
  
  transmute(
    
    `Relative severity profile` =
      as.character(
        severity_profile
      ),
    
    `Countries and territories, n` =
      Number_of_countries,
    
    `Countries and territories, %` =
      round(
        Percentage_of_countries,
        1
      ),
    
    `Median DALYs per case` =
      round(
        Median_DALYs_per_case,
        2
      ),
    
    `Median YLLs per case` =
      round(
        Median_YLLs_per_case,
        2
      ),
    
    `Median YLDs per case` =
      round(
        Median_YLDs_per_case,
        2
      ),
    
    `Median YLL contribution, %` =
      round(
        Median_YLL_percent,
        2
      ),
    
    `Median YLD contribution, %` =
      round(
        Median_YLD_percent,
        2
      ),
    
    `Median YLL:YLD ratio` =
      round(
        Median_YLL_YLD_ratio,
        2
      )
    
  )

writeData(
  table_workbook,
  sheet = "Profile summary",
  x = paste(
    "Table 4. Country-level colorectal cancer severity",
    "components and relative profiles in 2023"
  ),
  startRow = 1,
  startCol = 1
)

writeData(
  table_workbook,
  sheet = "Profile summary",
  x = table_4_summary,
  startRow = 3,
  startCol = 1
)


# ------------------------------------------------------------------------------
# Sheet 2: highest YLLs per case
# ------------------------------------------------------------------------------

addWorksheet(
  table_workbook,
  "Highest YLLs per case"
)

writeData(
  table_workbook,
  sheet = "Highest YLLs per case",
  x = highest_ylls_per_case,
  startRow = 1,
  startCol = 1
)


# ------------------------------------------------------------------------------
# Sheet 3: highest YLDs per case
# ------------------------------------------------------------------------------

addWorksheet(
  table_workbook,
  "Highest YLDs per case"
)

writeData(
  table_workbook,
  sheet = "Highest YLDs per case",
  x = highest_ylds_per_case,
  startRow = 1,
  startCol = 1
)


# ------------------------------------------------------------------------------
# Formatting
# ------------------------------------------------------------------------------

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
  sheet = "Profile summary",
  style = title_style,
  rows = 1,
  cols = 1:ncol(table_4_summary),
  gridExpand = TRUE
)

addStyle(
  table_workbook,
  sheet = "Profile summary",
  style = header_style,
  rows = 3,
  cols = 1:ncol(table_4_summary),
  gridExpand = TRUE
)

addStyle(
  table_workbook,
  sheet = "Profile summary",
  style = body_style,
  rows = 4:(nrow(table_4_summary) + 3),
  cols = 1:ncol(table_4_summary),
  gridExpand = TRUE
)

setColWidths(
  table_workbook,
  sheet = "Profile summary",
  cols = 1:ncol(table_4_summary),
  widths = 23
)

setRowHeights(
  table_workbook,
  sheet = "Profile summary",
  rows = 3,
  heights = 52
)

freezePane(
  table_workbook,
  sheet = "Profile summary",
  firstActiveRow = 4,
  firstActiveCol = 2
)

for (
  sheet_name in c(
    "Highest YLLs per case",
    "Highest YLDs per case"
  )
) {
  
  sheet_data <- if (
    sheet_name == "Highest YLLs per case"
  ) {
    
    highest_ylls_per_case
    
  } else {
    
    highest_ylds_per_case
    
  }
  
  addStyle(
    table_workbook,
    sheet = sheet_name,
    style = header_style,
    rows = 1,
    cols = 1:ncol(sheet_data),
    gridExpand = TRUE
  )
  
  addStyle(
    table_workbook,
    sheet = sheet_name,
    style = body_style,
    rows = 2:(nrow(sheet_data) + 1),
    cols = 1:ncol(sheet_data),
    gridExpand = TRUE
  )
  
  setColWidths(
    table_workbook,
    sheet = sheet_name,
    cols = 1:ncol(sheet_data),
    widths = "auto"
  )
  
  freezePane(
    table_workbook,
    sheet = sheet_name,
    firstRow = TRUE
  )
  
}

table_note <- paste0(
  
  "Note: Relative severity profiles were defined using tertiles of the ",
  
  "country-level YLD contribution distribution. Mortality-dominant: YLD% < ",
  
  format_number(
    yld_first_tertile,
    2
  ),
  
  "%; mixed severity: YLD% from ",
  
  format_number(
    yld_first_tertile,
    2
  ),
  
  "% to < ",
  
  format_number(
    yld_second_tertile,
    2
  ),
  
  "%; disability-enriched: YLD% >= ",
  
  format_number(
    yld_second_tertile,
    2
  ),
  
  "%. These categories describe relative variation between countries and do ",
  
  "not imply equal mortality and disability contributions."
  
)

writeData(
  table_workbook,
  sheet = "Profile summary",
  x = table_note,
  startRow = nrow(table_4_summary) + 6,
  startCol = 1,
  colNames = FALSE
)

saveWorkbook(
  table_workbook,
  table_4_file,
  overwrite = TRUE
)


# ==============================================================================
# 22. CREATE SUPPLEMENTARY TABLE S3
# ==============================================================================

supplementary_s3 <- country_severity |>
  
  arrange(
    location
  ) |>
  
  transmute(
    
    Country = location,
    
    Location_ID = location_id,
    
    Incident_cases = incident_cases,
    
    DALYs = dalys,
    
    YLLs = ylls,
    
    YLDs = ylds,
    
    DALYs_per_case = dalys_per_case,
    
    YLLs_per_case = ylls_per_case,
    
    YLDs_per_case = ylds_per_case,
    
    YLL_YLD_ratio = yll_yld_ratio,
    
    YLL_contribution_percent = yll_percent,
    
    YLD_contribution_percent = yld_percent,
    
    Relative_severity_profile =
      as.character(
        severity_profile
      ),
    
    DALY_minus_YLL_plus_YLD =
      daly_component_difference,
    
    Relative_decomposition_difference_percent =
      daly_component_relative_difference
    
  )


supplementary_s3_xlsx <- file.path(
  supplementary_folder,
  "Table_S3_Complete_country_severity_decomposition_2023.xlsx"
)

supplementary_s3_csv <- file.path(
  supplementary_folder,
  "Table_S3_Complete_country_severity_decomposition_2023.csv"
)


supplementary_workbook <- createWorkbook()

addWorksheet(
  supplementary_workbook,
  "Complete country dataset"
)

writeData(
  supplementary_workbook,
  sheet = "Complete country dataset",
  x = supplementary_s3
)

addWorksheet(
  supplementary_workbook,
  "Highest YLD contribution"
)

writeData(
  supplementary_workbook,
  sheet = "Highest YLD contribution",
  x = highest_yld_contribution
)

addWorksheet(
  supplementary_workbook,
  "Classification thresholds"
)

classification_thresholds <- tibble(
  
  Item = c(
    "Countries and territories",
    "First tertile of YLD contribution (%)",
    "Second tertile of YLD contribution (%)",
    "Median DALY decomposition difference (%)",
    "Maximum DALY decomposition difference (%)"
  ),
  
  Value = c(
    country_count,
    yld_first_tertile,
    yld_second_tertile,
    median_component_difference,
    maximum_component_difference
  )
  
)

writeData(
  supplementary_workbook,
  sheet = "Classification thresholds",
  x = classification_thresholds
)

for (
  sheet_name in c(
    "Complete country dataset",
    "Highest YLD contribution",
    "Classification thresholds"
  )
) {
  
  sheet_data <- switch(
    
    sheet_name,
    
    "Complete country dataset" =
      supplementary_s3,
    
    "Highest YLD contribution" =
      highest_yld_contribution,
    
    "Classification thresholds" =
      classification_thresholds
    
  )
  
  addStyle(
    supplementary_workbook,
    sheet = sheet_name,
    style = header_style,
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
  supplementary_s3_xlsx,
  overwrite = TRUE
)

readr::write_csv(
  supplementary_s3,
  supplementary_s3_csv
)


# ==============================================================================
# 23. SELECT OBJECTIVE COUNTRY LABELS FOR FIGURE 4
#
# Label:
#   Five highest YLLs per case
#   Five highest YLDs per case
#   Five highest YLD contributions
# ==============================================================================

label_locations <- unique(
  
  c(
    
    highest_ylls_per_case$Country[1:5],
    
    highest_ylds_per_case$Country[1:5],
    
    highest_yld_contribution$Country[1:5]
    
  )
  
)

figure_label_data <- country_severity |>
  
  filter(
    location %in% label_locations
  )


# ==============================================================================
# 24. CREATE FIGURE 4
#
# X-axis:
#   YLLs per case
#
# Y-axis:
#   YLDs per case
#
# Point color:
#   Relative severity profile
# ==============================================================================

median_ylls_per_case <- median(
  country_severity$ylls_per_case,
  na.rm = TRUE
)

median_ylds_per_case <- median(
  country_severity$ylds_per_case,
  na.rm = TRUE
)


figure_4 <- ggplot(
  
  country_severity,
  
  aes(
    x = ylls_per_case,
    y = ylds_per_case,
    color = severity_profile
  )
  
) +
  
  geom_vline(
    xintercept = median_ylls_per_case,
    linetype = "dashed",
    linewidth = 0.6
  ) +
  
  geom_hline(
    yintercept = median_ylds_per_case,
    linetype = "dashed",
    linewidth = 0.6
  ) +
  
  geom_point(
    size = 3,
    alpha = 0.82
  ) +
  
  ggrepel::geom_text_repel(
    
    data = figure_label_data,
    
    aes(
      label = location
    ),
    
    size = 3.2,
    
    fontface = "bold",
    
    box.padding = 0.45,
    
    point.padding = 0.25,
    
    min.segment.length = 0,
    
    max.overlaps = Inf,
    
    seed = 2023,
    
    show.legend = FALSE
    
  ) +
  
  scale_color_manual(
    
    values = c(
      
      "Mortality-dominant" =
        "#B03A2E",
      
      "Mixed severity" =
        "#D68910",
      
      "Disability-enriched" =
        "#2874A6"
      
    ),
    
    drop = FALSE
    
  ) +
  
  scale_x_continuous(
    
    name = "YLLs per incident case",
    
    labels = scales::label_number(
      accuracy = 0.1
    ),
    
    expand = expansion(
      mult = c(
        0.04,
        0.10
      )
    )
    
  ) +
  
  scale_y_continuous(
    
    name = "YLDs per incident case",
    
    labels = scales::label_number(
      accuracy = 0.01
    ),
    
    expand = expansion(
      mult = c(
        0.05,
        0.12
      )
    )
    
  ) +
  
  labs(
    
    title = paste(
      "Country-level decomposition of colorectal cancer",
      "severity in 2023"
    ),
    
    subtitle = paste0(
      
      "Dashed lines indicate country medians: YLLs per case = ",
      
      format_number(
        median_ylls_per_case,
        2
      ),
      
      "; YLDs per case = ",
      
      format_number(
        median_ylds_per_case,
        2
      )
      
    ),
    
    color = "Relative severity profile",
    
    caption = paste0(
      
      "Each point represents one of ",
      
      country_count,
      
      " countries or territories. ",
      
      "Classification is based on tertiles of YLD contribution and therefore ",
      
      "represents relative differences between countries."
      
    )
    
  ) +
  
  theme_classic(
    base_size = 12
  ) +
  
  theme(
    
    plot.title = element_text(
      face = "bold",
      size = 16,
      hjust = 0.5
    ),
    
    plot.subtitle = element_text(
      size = 10.5,
      hjust = 0.5
    ),
    
    axis.title.x = element_text(
      size = 11
    ),
    
    axis.title.y = element_text(
      size = 11
    ),
    
    axis.text = element_text(
      size = 10
    ),
    
    legend.position = "bottom",
    
    legend.title = element_text(
      size = 10,
      face = "bold"
    ),
    
    legend.text = element_text(
      size = 9.5
    ),
    
    panel.grid.major = element_line(
      linewidth = 0.25,
      linetype = "dotted"
    ),
    
    plot.caption = element_text(
      size = 9,
      hjust = 0
    ),
    
    plot.margin = margin(
      14,
      18,
      12,
      14
    )
    
  )


# ==============================================================================
# 25. SAVE FIGURE 4 AS TIFF AT 600 DPI
# ==============================================================================

figure_4_file <- file.path(
  figures_folder,
  "Figure_4_Country_CRC_severity_components_2023.tiff"
)

ggsave(
  filename = figure_4_file,
  plot = figure_4,
  device = "tiff",
  width = 12,
  height = 9,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white"
)


# ==============================================================================
# 26. EXTRACT RESULTS FOR AUTOMATIC TEXT
# ==============================================================================

mortality_profile <- profile_summary |>
  
  filter(
    severity_profile ==
      "Mortality-dominant"
  )


mixed_profile <- profile_summary |>
  
  filter(
    severity_profile ==
      "Mixed severity"
  )


disability_profile <- profile_summary |>
  
  filter(
    severity_profile ==
      "Disability-enriched"
  )


top_yll_names <- highest_ylls_per_case |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    Country
  ) |>
  
  paste(
    collapse = ", "
  )


top_yld_names <- highest_ylds_per_case |>
  
  slice_head(
    n = 5
  ) |>
  
  pull(
    Country
  ) |>
  
  paste(
    collapse = ", "
  )


top_yld_share_names <- highest_yld_contribution |>
  
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
# 27. GENERATE RESULTS TEXT
# ==============================================================================

results_text <- glue::glue(
  "3.4 Mortality-driven versus disability-driven severity

The country-level decomposition included {country_count} countries and territories. Across countries, the median YLLs-per-case value was {format_number(median_ylls_per_case, 2)}, while the median YLDs-per-case value was {format_number(median_ylds_per_case, 2)}. The distribution of countries across these two components demonstrated that similar total DALYs-per-case values could arise from different combinations of premature mortality and disability (Figure 4).

Relative severity profiles were classified using tertiles of the country-level YLD contribution distribution. The mortality-dominant group included {mortality_profile$Number_of_countries} countries and territories ({format_number(mortality_profile$Percentage_of_countries, 1)}%), with median values of {format_number(mortality_profile$Median_YLLs_per_case, 2)} YLLs and {format_number(mortality_profile$Median_YLDs_per_case, 2)} YLDs per case. The median YLL contribution in this group was {format_number(mortality_profile$Median_YLL_percent, 2)}%, compared with a median YLD contribution of {format_number(mortality_profile$Median_YLD_percent, 2)}% (Table 4).

The mixed-severity group comprised {mixed_profile$Number_of_countries} countries and territories ({format_number(mixed_profile$Percentage_of_countries, 1)}%). Its median severity profile included {format_number(mixed_profile$Median_YLLs_per_case, 2)} YLLs and {format_number(mixed_profile$Median_YLDs_per_case, 2)} YLDs per incident case, with median YLL and YLD contributions of {format_number(mixed_profile$Median_YLL_percent, 2)}% and {format_number(mixed_profile$Median_YLD_percent, 2)}%, respectively.

The disability-enriched group contained {disability_profile$Number_of_countries} countries and territories ({format_number(disability_profile$Percentage_of_countries, 1)}%). These countries had median values of {format_number(disability_profile$Median_YLLs_per_case, 2)} YLLs and {format_number(disability_profile$Median_YLDs_per_case, 2)} YLDs per case. Their median YLD contribution was {format_number(disability_profile$Median_YLD_percent, 2)}%, compared with {format_number(disability_profile$Median_YLL_percent, 2)}% attributable to YLLs (Table 4; Figure 4).

The highest YLLs-per-case values were observed in {top_yll_names}, whereas the highest YLDs-per-case values occurred in {top_yld_names}. The largest proportional disability contributions were observed in {top_yld_share_names}. Complete country-level estimates of DALYs, YLLs, YLDs, their corresponding per-case values, component percentages, YLL:YLD ratios, and relative severity classifications are provided in Supplementary Table S3.

These findings show that DALYs per case represents a composite severity measure whose underlying composition varies across countries. Although premature mortality remained the predominant component of colorectal cancer health loss, the relative contribution of disability differed meaningfully between settings."
)


# ==============================================================================
# 28. SAVE RESULTS TEXT
# ==============================================================================

results_text_file <- file.path(
  results_text_folder,
  "Results_3_4_Country_severity_decomposition_2023.txt"
)

writeLines(
  text = results_text,
  con = results_text_file,
  useBytes = TRUE
)


# ==============================================================================
# 29. DISPLAY RESULTS
# ==============================================================================

message("")
message("Table 4 summary:")

print(
  table_4_summary
)

message("")
message("Highest YLLs per case:")

print(
  highest_ylls_per_case
)

message("")
message("Highest YLDs per case:")

print(
  highest_ylds_per_case
)

message("")
message("Results text:")

cat(
  results_text
)


# ==============================================================================
# 30. COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("RESULTS SECTION 3.4 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message(
  "Countries and territories included: ",
  country_count
)

message("")
message("Table 4:")
message(table_4_file)

message("")
message("Figure 4 — one TIFF at 600 dpi:")
message(figure_4_file)

message("")
message("Supplementary Table S3:")
message(supplementary_s3_xlsx)

message("")
message("Results text:")
message(results_text_file)

message("")
message("The old Clean_Data folder was not used.")
message("==============================================================")