# ==============================================================================
# R_03_5_Temporal_Trends_1990_2023.R
#
# RESULTS SECTION 3.5
# Global temporal trends in colorectal cancer burden and severity, 1990–2023
#
# INPUT FOLDER
#   Global/
#
# REQUIRED MEASURES
#   Incidence
#   Mortality
#   DALYs
#
# REQUIRED DATA FOR EACH MEASURE
#   Number, All ages
#   Rate, Age-standardized
#
# REQUIRED SEXES
#   Both
#   Male
#   Female
#
# DERIVED METRICS
#   DALYs per case = DALY number / incidence number
#   MIR            = mortality number / incidence number
#
# OUTPUTS
#
# Publication/Figures/
#   Figure_5_Global_temporal_trends_1990_2023.tiff
#   Figure_5_Global_temporal_trends_1990_2023.png
#
# Publication/Tables/
#   Table_5_Global_temporal_trends_1990_2023.xlsx
#
# Publication/Supplementary/
#   Table_S4_Annual_global_temporal_trends_1990_2023.xlsx
#   Table_S4_Annual_global_temporal_trends_1990_2023.csv
#
# Publication/Results_text/
#   Results_3_5_Global_temporal_trends_1990_2023.txt
#
# IMPORTANT
#   - Uses the official GBD Global location.
#   - Does not use the Clean_Data folder.
#   - Country-level trends will be analysed separately in Section 3.6.
# ==============================================================================


# ==============================================================================
# 1. CLEAR ENVIRONMENT
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
# 2. INSTALL AND LOAD PACKAGES
# ==============================================================================

required_packages <- c(
  "tidyverse",
  "readxl",
  "openxlsx",
  "janitor",
  "patchwork",
  "scales",
  "glue"
)

installed_packages <- rownames(
  installed.packages()
)

missing_packages <- setdiff(
  required_packages,
  installed_packages
)

if (length(missing_packages) > 0) {
  
  install.packages(
    missing_packages,
    dependencies = TRUE
  )
  
}

suppressPackageStartupMessages({
  
  library(tidyverse)
  library(readxl)
  library(openxlsx)
  library(janitor)
  library(patchwork)
  library(scales)
  library(glue)
  
})


# ==============================================================================
# 3. IDENTIFY PROJECT ROOT
# ==============================================================================

current_folder <- normalizePath(
  getwd(),
  winslash = "/",
  mustWork = TRUE
)

parent_folder <- normalizePath(
  file.path(
    current_folder,
    ".."
  ),
  winslash = "/",
  mustWork = TRUE
)

if (
  dir.exists(
    file.path(
      current_folder,
      "Global"
    )
  ) &&
  dir.exists(
    file.path(
      current_folder,
      "Publication"
    )
  )
) {
  
  project_root <- current_folder
  
} else if (
  dir.exists(
    file.path(
      parent_folder,
      "Global"
    )
  ) &&
  dir.exists(
    file.path(
      parent_folder,
      "Publication"
    )
  )
) {
  
  project_root <- parent_folder
  
} else {
  
  stop(
    paste0(
      "\nProject root could not be identified.\n\n",
      "The project root must contain:\n",
      "Global/\n",
      "Publication/\n\n",
      "Current working directory:\n",
      current_folder
    ),
    call. = FALSE
  )
  
}

message("")
message("Project root:")
message(project_root)


# ==============================================================================
# 4. DEFINE FOLDERS
# ==============================================================================

global_folder <- file.path(
  project_root,
  "Global"
)

publication_folder <- file.path(
  project_root,
  "Publication"
)

figures_folder <- file.path(
  publication_folder,
  "Figures"
)

tables_folder <- file.path(
  publication_folder,
  "Tables"
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
  figures_folder,
  recursive = TRUE,
  showWarnings = FALSE
)

dir.create(
  tables_folder,
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
# 5. ANALYSIS SETTINGS
# ==============================================================================

start_year <- 1990

end_year <- 2023

expected_years <- start_year:end_year

sex_order <- c(
  "Both",
  "Male",
  "Female"
)

measure_order <- c(
  "Incidence",
  "Mortality",
  "DALYs"
)

metric_order <- c(
  "Number",
  "ASR"
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
      "NULL",
      "-"
    )
  )
  
}


find_column <- function(
    data,
    accepted_names,
    variable_description,
    required = TRUE
) {
  
  detected <- intersect(
    accepted_names,
    names(data)
  )
  
  if (length(detected) == 0) {
    
    if (required) {
      
      stop(
        paste0(
          "\nRequired column not found: ",
          variable_description,
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
  
  detected[[1]]
  
}


standardise_sex <- function(x) {
  
  value <- stringr::str_to_lower(
    clean_text(x)
  )
  
  dplyr::case_when(
    
    value %in% c(
      "both",
      "both sexes",
      "both sex",
      "total"
    ) ~
      "Both",
    
    value %in% c(
      "male",
      "males",
      "men"
    ) ~
      "Male",
    
    value %in% c(
      "female",
      "females",
      "women"
    ) ~
      "Female",
    
    TRUE ~
      clean_text(x)
    
  )
  
}


standardise_measure <- function(x) {
  
  value <- stringr::str_to_lower(
    clean_text(x)
  )
  
  dplyr::case_when(
    
    stringr::str_detect(
      value,
      "incidence|incident"
    ) ~
      "Incidence",
    
    stringr::str_detect(
      value,
      "mortality|death|deaths"
    ) ~
      "Mortality",
    
    stringr::str_detect(
      value,
      "disability.adjusted|dalys|daly"
    ) ~
      "DALYs",
    
    TRUE ~
      clean_text(x)
    
  )
  
}


detect_measure_from_filename <- function(file_path) {
  
  filename <- stringr::str_to_lower(
    basename(file_path)
  )
  
  dplyr::case_when(
    
    stringr::str_detect(
      filename,
      "incidence|incident"
    ) ~
      "Incidence",
    
    stringr::str_detect(
      filename,
      "mortality|death"
    ) ~
      "Mortality",
    
    stringr::str_detect(
      filename,
      "daly"
    ) ~
      "DALYs",
    
    TRUE ~
      NA_character_
    
  )
  
}


standardise_metric <- function(metric, age) {
  
  metric_value <- stringr::str_to_lower(
    clean_text(metric)
  )
  
  age_value <- stringr::str_to_lower(
    clean_text(age)
  )
  
  number_condition <-
    stringr::str_detect(
      metric_value,
      "number|count"
    ) &
    (
      is.na(age_value) |
        age_value == "" |
        stringr::str_detect(
          age_value,
          "^all ages$|^all age$|all ages combined"
        )
    )
  
  asr_condition <-
    stringr::str_detect(
      metric_value,
      "rate|asr"
    ) &
    (
      stringr::str_detect(
        age_value,
        "age.standard|age standard|standardized|standardised"
      ) |
        stringr::str_detect(
          metric_value,
          "age.standard|age standard|asr"
        )
    )
  
  dplyr::case_when(
    number_condition ~ "Number",
    asr_condition ~ "ASR",
    TRUE ~ NA_character_
  )
}


read_data_file <- function(file_path) {
  
  extension <- stringr::str_to_lower(
    tools::file_ext(
      file_path
    )
  )
  
  message("")
  message("Reading file:")
  message(file_path)
  
  if (extension == "csv") {
    
    output <- readr::read_csv(
      file_path,
      show_col_types = FALSE,
      progress = FALSE,
      guess_max = 500000
    )
    
  } else if (extension %in% c(
    "xlsx",
    "xls"
  )) {
    
    sheet_names <- readxl::excel_sheets(
      file_path
    )
    
    preferred_sheet <- sheet_names[
      stringr::str_detect(
        stringr::str_to_lower(
          sheet_names
        ),
        "data|result|global"
      )
    ][1]
    
    if (is.na(preferred_sheet)) {
      
      preferred_sheet <- sheet_names[1]
      
    }
    
    output <- readxl::read_excel(
      file_path,
      sheet = preferred_sheet
    )
    
  } else {
    
    stop(
      paste0(
        "\nUnsupported file type:\n",
        file_path
      ),
      call. = FALSE
    )
    
  }
  
  output |>
    janitor::clean_names()
  
}


first_non_missing <- function(x) {
  
  values <- x[
    !is.na(x)
  ]
  
  if (length(values) == 0) {
    
    return(
      NA_real_
    )
    
  }
  
  values[1]
  
}


format_number <- function(
    x,
    digits = 2
) {
  
  ifelse(
    is.na(x),
    "NA",
    formatC(
      x,
      format = "f",
      digits = digits,
      big.mark = ","
    )
  )
  
}


format_p_value <- function(p) {
  
  if (
    length(p) == 0 ||
    is.na(p)
  ) {
    
    return("NA")
    
  }
  
  if (p < 0.001) {
    
    return("<0.001")
    
  }
  
  formatC(
    p,
    format = "f",
    digits = 3
  )
  
}


format_change <- function(x) {
  
  dplyr::case_when(
    
    is.na(x) ~
      "NA",
    
    x > 0 ~
      paste0(
        "+",
        formatC(
          x,
          format = "f",
          digits = 1
        ),
        "%"
      ),
    
    TRUE ~
      paste0(
        formatC(
          x,
          format = "f",
          digits = 1
        ),
        "%"
      )
    
  )
  
}


trend_direction <- function(
    estimate,
    p_value
) {
  
  if (
    is.na(estimate) ||
    is.na(p_value)
  ) {
    
    return(
      "could not be estimated"
    )
    
  }
  
  if (p_value >= 0.05) {
    
    return(
      "showed no statistically significant log-linear trend"
    )
    
  }
  
  if (estimate > 0) {
    
    return(
      "increased significantly"
    )
    
  }
  
  if (estimate < 0) {
    
    return(
      "decreased significantly"
    )
    
  }
  
  "remained stable"
  
}


calculate_aapc <- function(data) {
  
  model_data <- data |>
    
    filter(
      !is.na(year),
      !is.na(value),
      value > 0
    ) |>
    
    arrange(
      year
    )
  
  if (nrow(model_data) < 3) {
    
    return(
      tibble(
        aapc = NA_real_,
        lower_95_ci = NA_real_,
        upper_95_ci = NA_real_,
        p_value = NA_real_,
        r_squared = NA_real_,
        observations = nrow(model_data)
      )
    )
    
  }
  
  model <- lm(
    log(value) ~ year,
    data = model_data
  )
  
  model_summary <- summary(
    model
  )
  
  beta <- unname(
    coef(model)[["year"]]
  )
  
  beta_ci <- confint(
    model,
    parm = "year",
    level = 0.95
  )
  
  tibble(
    
    aapc =
      100 * (
        exp(beta) - 1
      ),
    
    lower_95_ci =
      100 * (
        exp(beta_ci[1]) - 1
      ),
    
    upper_95_ci =
      100 * (
        exp(beta_ci[2]) - 1
      ),
    
    p_value =
      model_summary$coefficients[
        "year",
        "Pr(>|t|)"
      ],
    
    r_squared =
      model_summary$r.squared,
    
    observations =
      nrow(model_data)
    
  )
  
}


# ==============================================================================
# 7. LOCATE INPUT FILES
# ==============================================================================

global_files <- list.files(
  path = global_folder,
  pattern = "\\.(csv|xlsx|xls)$",
  full.names = TRUE,
  recursive = TRUE,
  ignore.case = TRUE
)

global_files <- global_files[
  !stringr::str_detect(
    basename(global_files),
    "^~\\$"
  )
]

if (length(global_files) == 0) {
  
  stop(
    paste0(
      "\nNo CSV or Excel files were found in:\n",
      global_folder
    ),
    call. = FALSE
  )
  
}

message("")
message("Global input files detected:")

for (file_path in global_files) {
  
  message(file_path)
  
}


# ==============================================================================
# 8. IMPORT ALL GLOBAL FILES
# ==============================================================================

imported_data_list <- vector(
  mode = "list",
  length = length(global_files)
)

for (i in seq_along(global_files)) {
  
  current_file <- global_files[i]
  
  current_data <- read_data_file(
    current_file
  )
  
  location_column <- find_column(
    current_data,
    c(
      "location",
      "location_name",
      "country",
      "country_name"
    ),
    "location",
    required = FALSE
  )
  
  sex_column <- find_column(
    current_data,
    c(
      "sex",
      "sex_name"
    ),
    "sex"
  )
  
  age_column <- find_column(
    current_data,
    c(
      "age",
      "age_name"
    ),
    "age",
    required = FALSE
  )
  
  cause_column <- find_column(
    current_data,
    c(
      "cause",
      "cause_name"
    ),
    "cause",
    required = FALSE
  )
  
  measure_column <- find_column(
    current_data,
    c(
      "measure",
      "measure_name"
    ),
    "measure",
    required = FALSE
  )
  
  metric_column <- find_column(
    current_data,
    c(
      "metric",
      "metric_name"
    ),
    "metric"
  )
  
  year_column <- find_column(
    current_data,
    c(
      "year",
      "year_id"
    ),
    "year"
  )
  
  value_column <- find_column(
    current_data,
    c(
      "val",
      "value",
      "mean",
      "estimate"
    ),
    "value"
  )
  
  lower_column <- find_column(
    current_data,
    c(
      "lower",
      "lower_ci",
      "lower_bound",
      "lower_value"
    ),
    "lower uncertainty estimate",
    required = FALSE
  )
  
  upper_column <- find_column(
    current_data,
    c(
      "upper",
      "upper_ci",
      "upper_bound",
      "upper_value"
    ),
    "upper uncertainty estimate",
    required = FALSE
  )
  
  filename_measure <- detect_measure_from_filename(
    current_file
  )
  
  standardised_data <- current_data |>
    
    transmute(
      
      source_file =
        basename(
          current_file
        ),
      
      location =
        if (!is.null(location_column)) {
          
          clean_text(
            .data[[location_column]]
          )
          
        } else {
          
          "Global"
          
        },
      
      cause =
        if (!is.null(cause_column)) {
          
          clean_text(
            .data[[cause_column]]
          )
          
        } else {
          
          "Colon and rectum cancer"
          
        },
      
      sex =
        standardise_sex(
          .data[[sex_column]]
        ),
      
      age =
        if (!is.null(age_column)) {
          
          clean_text(
            .data[[age_column]]
          )
          
        } else {
          
          NA_character_
          
        },
      
      measure =
        if (!is.null(measure_column)) {
          
          standardise_measure(
            .data[[measure_column]]
          )
          
        } else {
          
          filename_measure
          
        },
      
      metric_original =
        clean_text(
          .data[[metric_column]]
        ),
      
      year =
        as.integer(
          .data[[year_column]]
        ),
      
      value =
        to_numeric_safely(
          .data[[value_column]]
        ),
      
      lower =
        if (!is.null(lower_column)) {
          
          to_numeric_safely(
            .data[[lower_column]]
          )
          
        } else {
          
          NA_real_
          
        },
      
      upper =
        if (!is.null(upper_column)) {
          
          to_numeric_safely(
            .data[[upper_column]]
          )
          
        } else {
          
          NA_real_
          
        }
      
    ) |>
    
    mutate(
      
      analysis_metric =
        standardise_metric(
          metric_original,
          age
        )
      
    )
  
  imported_data_list[[i]] <- standardised_data
  
}

global_raw <- bind_rows(
  imported_data_list
)


# ==============================================================================
# 9. INSPECT IMPORTED DATA
# ==============================================================================

message("")
message("Detected locations:")

print(
  global_raw |>
    count(
      location,
      sort = TRUE
    )
)

message("")
message("Detected causes:")

print(
  global_raw |>
    count(
      cause,
      sort = TRUE
    )
)

message("")
message("Detected measures:")

print(
  global_raw |>
    count(
      measure,
      sort = TRUE
    )
)

message("")
message("Detected ages and metrics:")

print(
  global_raw |>
    count(
      age,
      metric_original,
      analysis_metric,
      sort = TRUE
    )
)

message("")
message("Detected sexes:")

print(
  global_raw |>
    count(
      sex,
      sort = TRUE
    )
)


# ==============================================================================
# 10. RETAIN COLORECTAL CANCER RECORDS
# ==============================================================================

cause_present <- any(
  stringr::str_detect(
    stringr::str_to_lower(
      global_raw$cause
    ),
    "colon|rectum|colorectal"
  ),
  na.rm = TRUE
)

if (cause_present) {
  
  global_raw <- global_raw |>
    
    filter(
      stringr::str_detect(
        stringr::str_to_lower(
          cause
        ),
        "colon|rectum|colorectal"
      )
    )
  
}


# ==============================================================================
# 11. RETAIN GLOBAL LOCATION
# ==============================================================================

global_location_present <- any(
  stringr::str_to_lower(
    global_raw$location
  ) %in% c(
    "global",
    "world"
  ),
  na.rm = TRUE
)

if (global_location_present) {
  
  global_raw <- global_raw |>
    
    filter(
      stringr::str_to_lower(
        location
      ) %in% c(
        "global",
        "world"
      )
    )
  
} else {
  
  detected_locations <- global_raw |>
    
    distinct(
      location
    ) |>
    
    filter(
      !is.na(location),
      location != ""
    )
  
  if (nrow(detected_locations) > 1) {
    
    stop(
      paste0(
        "\nMore than one location was detected, but no Global location ",
        "was identified.\n\nDetected locations:\n",
        paste(
          detected_locations$location,
          collapse = "\n"
        )
      ),
      call. = FALSE
    )
    
  }
  
}


# ==============================================================================
# 12. FILTER REQUIRED RECORDS
# ==============================================================================

global_filtered <- global_raw |>
  
  filter(
    year >= start_year,
    year <= end_year,
    sex %in% sex_order,
    measure %in% measure_order,
    analysis_metric %in% metric_order
  )

if (nrow(global_filtered) == 0) {
  
  stop(
    "\nNo usable observations remained after filtering.",
    call. = FALSE
  )
  
}

message("")
message("Retained observations:")

print(
  global_filtered |>
    count(
      measure,
      analysis_metric,
      sex,
      sort = TRUE
    )
)


# ==============================================================================
# 13. CHECK REQUIRED COMBINATIONS
# ==============================================================================

required_combinations <- tidyr::expand_grid(
  measure = measure_order,
  analysis_metric = metric_order,
  sex = sex_order
)

available_combinations <- global_filtered |>
  
  distinct(
    measure,
    analysis_metric,
    sex
  )

missing_combinations <- anti_join(
  required_combinations,
  available_combinations,
  by = c(
    "measure",
    "analysis_metric",
    "sex"
  )
)

if (nrow(missing_combinations) > 0) {
  
  message("")
  message("Missing combinations:")
  
  print(
    missing_combinations
  )
  
  stop(
    paste0(
      "\nThe Global folder does not contain all required combinations.\n",
      "The required measures are incidence, mortality, and DALYs only."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 14. RESOLVE DUPLICATES
# ==============================================================================

duplicate_check <- global_filtered |>
  
  count(
    measure,
    analysis_metric,
    sex,
    year,
    name = "records"
  ) |>
  
  filter(
    records > 1
  )

if (nrow(duplicate_check) > 0) {
  
  message("")
  message("Duplicate combinations detected:")
  
  print(
    duplicate_check
  )
  
  message(
    "The first available value will be retained for each combination."
  )
  
}

global_analysis <- global_filtered |>
  
  group_by(
    measure,
    analysis_metric,
    sex,
    year
  ) |>
  
  summarise(
    
    value =
      first_non_missing(
        value
      ),
    
    lower =
      first_non_missing(
        lower
      ),
    
    upper =
      first_non_missing(
        upper
      ),
    
    .groups = "drop"
    
  )


# ==============================================================================
# 15. CREATE NUMBER DATASET
# ==============================================================================

number_data <- global_analysis |>
  
  filter(
    analysis_metric == "Number"
  ) |>
  
  select(
    sex,
    year,
    measure,
    value
  ) |>
  
  pivot_wider(
    names_from = measure,
    values_from = value
  )

missing_number_columns <- setdiff(
  measure_order,
  names(number_data)
)

if (length(missing_number_columns) > 0) {
  
  stop(
    paste0(
      "\nMissing number columns:\n",
      paste(
        missing_number_columns,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
  
}

number_data <- number_data |>
  
  rename(
    incidence_number = Incidence,
    mortality_number = Mortality,
    daly_number = DALYs
  )


# ==============================================================================
# 16. CREATE AGE-STANDARDIZED RATE DATASET
# ==============================================================================

asr_data <- global_analysis |>
  
  filter(
    analysis_metric == "ASR"
  ) |>
  
  select(
    sex,
    year,
    measure,
    value
  ) |>
  
  pivot_wider(
    names_from = measure,
    values_from = value
  )

missing_asr_columns <- setdiff(
  measure_order,
  names(asr_data)
)

if (length(missing_asr_columns) > 0) {
  
  stop(
    paste0(
      "\nMissing age-standardized rate columns:\n",
      paste(
        missing_asr_columns,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
  
}

asr_data <- asr_data |>
  
  rename(
    incidence_asr = Incidence,
    mortality_asr = Mortality,
    daly_asr = DALYs
  )


# ==============================================================================
# 17. CREATE ANNUAL TREND DATASET
# ==============================================================================

annual_trends <- number_data |>
  
  inner_join(
    asr_data,
    by = c(
      "sex",
      "year"
    )
  ) |>
  
  mutate(
    
    dalys_per_case =
      if_else(
        incidence_number > 0,
        daly_number / incidence_number,
        NA_real_
      ),
    
    mir =
      if_else(
        incidence_number > 0,
        mortality_number / incidence_number,
        NA_real_
      ),
    
    sex = factor(
      sex,
      levels = sex_order
    )
    
  ) |>
  
  arrange(
    sex,
    year
  )


# ==============================================================================
# 18. VALIDATE YEARS AND VALUES
# ==============================================================================

expected_sex_years <- tidyr::expand_grid(
  sex = factor(
    sex_order,
    levels = sex_order
  ),
  year = expected_years
)

available_sex_years <- annual_trends |>
  
  distinct(
    sex,
    year
  )

missing_sex_years <- anti_join(
  expected_sex_years,
  available_sex_years,
  by = c(
    "sex",
    "year"
  )
)

if (nrow(missing_sex_years) > 0) {
  
  message("")
  message("Missing sex-year combinations:")
  
  print(
    missing_sex_years
  )
  
  stop(
    "\nThe dataset does not contain all years from 1990 through 2023.",
    call. = FALSE
  )
  
}

invalid_values <- annual_trends |>
  
  filter(
    is.na(incidence_number) |
      incidence_number <= 0 |
      is.na(mortality_number) |
      mortality_number < 0 |
      is.na(daly_number) |
      daly_number < 0 |
      is.na(incidence_asr) |
      incidence_asr < 0 |
      is.na(mortality_asr) |
      mortality_asr < 0 |
      is.na(daly_asr) |
      daly_asr < 0 |
      is.na(dalys_per_case) |
      dalys_per_case < 0 |
      is.na(mir) |
      mir < 0
  )

if (nrow(invalid_values) > 0) {
  
  message("")
  message("Invalid records:")
  
  print(
    invalid_values
  )
  
  stop(
    "\nMissing, negative, or impossible values were detected.",
    call. = FALSE
  )
  
}

expected_rows <- length(
  expected_years
) * length(
  sex_order
)

if (nrow(annual_trends) != expected_rows) {
  
  stop(
    paste0(
      "\nExpected ",
      expected_rows,
      " annual observations, but ",
      nrow(annual_trends),
      " were created."
    ),
    call. = FALSE
  )
  
}

message("")
message("Annual trend data successfully created:")

print(
  annual_trends |>
    select(
      sex,
      year,
      incidence_number,
      mortality_number,
      daly_number,
      incidence_asr,
      mortality_asr,
      daly_asr,
      dalys_per_case,
      mir
    ) |>
    slice_head(
      n = 12
    )
)


# ==============================================================================
# 19. CREATE LONG DATASET FOR TREND MODELS
# ==============================================================================

trend_long <- annual_trends |>
  
  select(
    sex,
    year,
    incidence_number,
    mortality_number,
    daly_number,
    incidence_asr,
    mortality_asr,
    daly_asr,
    dalys_per_case,
    mir
  ) |>
  
  pivot_longer(
    cols = -c(
      sex,
      year
    ),
    names_to = "metric_code",
    values_to = "value"
  ) |>
  
  mutate(
    
    metric = case_when(
      
      metric_code == "incidence_number" ~
        "Incident cases",
      
      metric_code == "mortality_number" ~
        "Deaths",
      
      metric_code == "daly_number" ~
        "Total DALYs",
      
      metric_code == "incidence_asr" ~
        "Incidence ASR",
      
      metric_code == "mortality_asr" ~
        "Mortality ASR",
      
      metric_code == "daly_asr" ~
        "DALY ASR",
      
      metric_code == "dalys_per_case" ~
        "DALYs per case",
      
      metric_code == "mir" ~
        "Mortality-to-incidence ratio",
      
      TRUE ~
        metric_code
      
    )
    
  )


# ==============================================================================
# 20. CALCULATE AVERAGE ANNUAL PERCENTAGE CHANGE
# ==============================================================================

aapc_results <- trend_long |>
  
  group_by(
    sex,
    metric
  ) |>
  
  group_modify(
    ~ calculate_aapc(
      .x
    )
  ) |>
  
  ungroup()


# ==============================================================================
# 21. CALCULATE 1990–2023 CHANGES
# ==============================================================================

period_change <- trend_long |>
  
  filter(
    year %in% c(
      start_year,
      end_year
    )
  ) |>
  
  select(
    sex,
    metric,
    year,
    value
  ) |>
  
  pivot_wider(
    names_from = year,
    values_from = value,
    names_prefix = "year_"
  )

start_value_column <- paste0(
  "year_",
  start_year
)

end_value_column <- paste0(
  "year_",
  end_year
)

period_change <- period_change |>
  
  mutate(
    
    start_value =
      .data[[start_value_column]],
    
    end_value =
      .data[[end_value_column]],
    
    absolute_change =
      end_value -
      start_value,
    
    percentage_change =
      100 *
      (
        end_value -
          start_value
      ) /
      start_value
    
  ) |>
  
  select(
    sex,
    metric,
    start_value,
    end_value,
    absolute_change,
    percentage_change
  )


# ==============================================================================
# 22. COMBINE TREND RESULTS
# ==============================================================================

metric_order_full <- c(
  "Incident cases",
  "Deaths",
  "Total DALYs",
  "Incidence ASR",
  "Mortality ASR",
  "DALY ASR",
  "DALYs per case",
  "Mortality-to-incidence ratio"
)

trend_summary <- aapc_results |>
  
  left_join(
    period_change,
    by = c(
      "sex",
      "metric"
    )
  ) |>
  
  mutate(
    
    sex = factor(
      sex,
      levels = sex_order
    ),
    
    metric = factor(
      metric,
      levels = metric_order_full
    )
    
  ) |>
  
  arrange(
    metric,
    sex
  )


# ==============================================================================
# 23. CREATE TABLE 5
# ==============================================================================

main_table_metrics <- c(
  "Incidence ASR",
  "Mortality ASR",
  "DALY ASR",
  "DALYs per case",
  "Mortality-to-incidence ratio"
)

table_5_numeric <- trend_summary |>
  
  filter(
    as.character(
      metric
    ) %in% main_table_metrics
  ) |>
  
  transmute(
    
    Measure =
      as.character(
        metric
      ),
    
    Sex =
      as.character(
        sex
      ),
    
    Value_1990 =
      start_value,
    
    Value_2023 =
      end_value,
    
    Percentage_change =
      percentage_change,
    
    AAPC_percent =
      aapc,
    
    Lower_95_CI =
      lower_95_ci,
    
    Upper_95_CI =
      upper_95_ci,
    
    P_value =
      p_value,
    
    R_squared =
      r_squared,
    
    Annual_observations =
      observations
    
  )

table_5_formatted <- table_5_numeric |>
  
  transmute(
    
    Measure,
    
    Sex,
    
    `1990 value` =
      format_number(
        Value_1990,
        2
      ),
    
    `2023 value` =
      format_number(
        Value_2023,
        2
      ),
    
    `Change, %` =
      format_change(
        Percentage_change
      ),
    
    `Average annual change, % (95% CI)` =
      paste0(
        format_number(
          AAPC_percent,
          2
        ),
        " (",
        format_number(
          Lower_95_CI,
          2
        ),
        " to ",
        format_number(
          Upper_95_CI,
          2
        ),
        ")"
      ),
    
    `P value` =
      map_chr(
        P_value,
        format_p_value
      )
    
  )


# ==============================================================================
# 24. EXPORT TABLE 5
# ==============================================================================

table_5_file <- file.path(
  tables_folder,
  "Table_5_Global_temporal_trends_1990_2023.xlsx"
)

table_5_workbook <- createWorkbook()

addWorksheet(
  table_5_workbook,
  "Publication table"
)

writeData(
  table_5_workbook,
  sheet = "Publication table",
  x = paste0(
    "Table 5. Global temporal trends in colorectal cancer ",
    "burden and severity, 1990–2023"
  ),
  startRow = 1,
  startCol = 1
)

writeData(
  table_5_workbook,
  sheet = "Publication table",
  x = table_5_formatted,
  startRow = 3,
  startCol = 1
)

table_note <- paste0(
  "Note: Annual percentage changes were estimated using log-linear ",
  "regression of annual values from 1990 through 2023. ",
  "ASR denotes age-standardized rate per 100,000 population. ",
  "DALYs per case and the mortality-to-incidence ratio were calculated ",
  "using annual number estimates."
)

writeData(
  table_5_workbook,
  sheet = "Publication table",
  x = table_note,
  startRow = nrow(table_5_formatted) + 6,
  startCol = 1,
  colNames = FALSE
)

addWorksheet(
  table_5_workbook,
  "Numeric results"
)

writeData(
  table_5_workbook,
  sheet = "Numeric results",
  x = table_5_numeric
)

addWorksheet(
  table_5_workbook,
  "All trend models"
)

writeData(
  table_5_workbook,
  sheet = "All trend models",
  x = trend_summary |>
    mutate(
      sex = as.character(
        sex
      ),
      metric = as.character(
        metric
      )
    )
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
  table_5_workbook,
  sheet = "Publication table",
  style = title_style,
  rows = 1,
  cols = 1:ncol(table_5_formatted),
  gridExpand = TRUE
)

addStyle(
  table_5_workbook,
  sheet = "Publication table",
  style = header_style,
  rows = 3,
  cols = 1:ncol(table_5_formatted),
  gridExpand = TRUE
)

addStyle(
  table_5_workbook,
  sheet = "Publication table",
  style = body_style,
  rows = 4:(nrow(table_5_formatted) + 3),
  cols = 1:ncol(table_5_formatted),
  gridExpand = TRUE
)

setColWidths(
  table_5_workbook,
  sheet = "Publication table",
  cols = 1:ncol(table_5_formatted),
  widths = c(
    31,
    12,
    16,
    16,
    16,
    33,
    13
  )
)

setRowHeights(
  table_5_workbook,
  sheet = "Publication table",
  rows = 3,
  heights = 45
)

freezePane(
  table_5_workbook,
  sheet = "Publication table",
  firstActiveRow = 4,
  firstActiveCol = 3
)

for (
  sheet_name in c(
    "Numeric results",
    "All trend models"
  )
) {
  
  addStyle(
    table_5_workbook,
    sheet = sheet_name,
    style = header_style,
    rows = 1,
    cols = 1:20,
    gridExpand = TRUE
  )
  
  setColWidths(
    table_5_workbook,
    sheet = sheet_name,
    cols = 1:20,
    widths = "auto"
  )
  
  freezePane(
    table_5_workbook,
    sheet = sheet_name,
    firstRow = TRUE
  )
  
}

saveWorkbook(
  table_5_workbook,
  table_5_file,
  overwrite = TRUE
)


# ==============================================================================
# 25. CREATE SUPPLEMENTARY TABLE S4
# ==============================================================================

supplementary_s4 <- annual_trends |>
  
  transmute(
    
    Year =
      year,
    
    Sex =
      as.character(
        sex
      ),
    
    Incidence_number =
      incidence_number,
    
    Mortality_number =
      mortality_number,
    
    DALY_number =
      daly_number,
    
    Incidence_ASR =
      incidence_asr,
    
    Mortality_ASR =
      mortality_asr,
    
    DALY_ASR =
      daly_asr,
    
    DALYs_per_case =
      dalys_per_case,
    
    Mortality_to_incidence_ratio =
      mir
    
  ) |>
  
  arrange(
    Year,
    factor(
      Sex,
      levels = sex_order
    )
  )

s4_excel_file <- file.path(
  supplementary_folder,
  "Table_S4_Annual_global_temporal_trends_1990_2023.xlsx"
)

s4_csv_file <- file.path(
  supplementary_folder,
  "Table_S4_Annual_global_temporal_trends_1990_2023.csv"
)

s4_workbook <- createWorkbook()

addWorksheet(
  s4_workbook,
  "Annual values"
)

writeData(
  s4_workbook,
  sheet = "Annual values",
  x = supplementary_s4
)

addWorksheet(
  s4_workbook,
  "Trend estimates"
)

writeData(
  s4_workbook,
  sheet = "Trend estimates",
  x = trend_summary |>
    mutate(
      sex = as.character(
        sex
      ),
      metric = as.character(
        metric
      )
    )
)

addWorksheet(
  s4_workbook,
  "Publication Table 5"
)

writeData(
  s4_workbook,
  sheet = "Publication Table 5",
  x = table_5_formatted
)

for (
  sheet_name in c(
    "Annual values",
    "Trend estimates",
    "Publication Table 5"
  )
) {
  
  addStyle(
    s4_workbook,
    sheet = sheet_name,
    style = header_style,
    rows = 1,
    cols = 1:20,
    gridExpand = TRUE
  )
  
  setColWidths(
    s4_workbook,
    sheet = sheet_name,
    cols = 1:20,
    widths = "auto"
  )
  
  freezePane(
    s4_workbook,
    sheet = sheet_name,
    firstRow = TRUE
  )
  
}

saveWorkbook(
  s4_workbook,
  s4_excel_file,
  overwrite = TRUE
)

readr::write_csv(
  supplementary_s4,
  s4_csv_file
)


# ==============================================================================
# 26. PREPARE FIGURE DATA
# ==============================================================================

figure_data <- annual_trends |>
  
  select(
    sex,
    year,
    incidence_asr,
    mortality_asr,
    daly_asr,
    dalys_per_case,
    mir
  )

sex_colours <- c(
  "Both" = "#000000",
  "Male" = "#2166AC",
  "Female" = "#B2182B"
)

sex_line_types <- c(
  "Both" = "solid",
  "Male" = "dashed",
  "Female" = "dotted"
)

common_theme <- theme_classic(
  base_size = 12
) +
  
  theme(
    
    plot.title = element_text(
      face = "bold",
      size = 12.5,
      hjust = 0
    ),
    
    axis.title = element_text(
      face = "bold",
      size = 10.5
    ),
    
    axis.text = element_text(
      size = 9.5
    ),
    
    panel.grid.major.y = element_line(
      linewidth = 0.25,
      linetype = "dotted",
      color = "grey75"
    ),
    
    panel.grid.minor = element_blank(),
    
    legend.position = "bottom",
    
    legend.title = element_blank(),
    
    legend.text = element_text(
      size = 10
    ),
    
    plot.margin = margin(
      8,
      10,
      8,
      8
    )
    
  )

year_breaks <- c(
  1990,
  1995,
  2000,
  2005,
  2010,
  2015,
  2020,
  2023
)


# ==============================================================================
# 27. FIGURE 5A — INCIDENCE ASR
# ==============================================================================

figure_5a <- ggplot(
  figure_data,
  aes(
    x = year,
    y = incidence_asr,
    color = sex,
    linetype = sex
  )
) +
  
  geom_line(
    linewidth = 1.05
  ) +
  
  scale_color_manual(
    values = sex_colours
  ) +
  
  scale_linetype_manual(
    values = sex_line_types
  ) +
  
  scale_x_continuous(
    name = NULL,
    breaks = year_breaks,
    limits = c(
      start_year,
      end_year
    ),
    expand = expansion(
      mult = c(
        0.01,
        0.02
      )
    )
  ) +
  
  scale_y_continuous(
    name = "Incidence ASR\nper 100,000",
    labels = scales::label_number(
      accuracy = 0.1
    ),
    expand = expansion(
      mult = c(
        0.05,
        0.08
      )
    )
  ) +
  
  labs(
    title = "A. Incidence"
  ) +
  
  common_theme +
  
  theme(
    legend.position = "none"
  )


# ==============================================================================
# 28. FIGURE 5B — MORTALITY ASR
# ==============================================================================

figure_5b <- ggplot(
  figure_data,
  aes(
    x = year,
    y = mortality_asr,
    color = sex,
    linetype = sex
  )
) +
  
  geom_line(
    linewidth = 1.05
  ) +
  
  scale_color_manual(
    values = sex_colours
  ) +
  
  scale_linetype_manual(
    values = sex_line_types
  ) +
  
  scale_x_continuous(
    name = NULL,
    breaks = year_breaks,
    limits = c(
      start_year,
      end_year
    ),
    expand = expansion(
      mult = c(
        0.01,
        0.02
      )
    )
  ) +
  
  scale_y_continuous(
    name = "Mortality ASR\nper 100,000",
    labels = scales::label_number(
      accuracy = 0.1
    ),
    expand = expansion(
      mult = c(
        0.05,
        0.08
      )
    )
  ) +
  
  labs(
    title = "B. Mortality"
  ) +
  
  common_theme +
  
  theme(
    legend.position = "none"
  )


# ==============================================================================
# 29. FIGURE 5C — DALY ASR
# ==============================================================================

figure_5c <- ggplot(
  figure_data,
  aes(
    x = year,
    y = daly_asr,
    color = sex,
    linetype = sex
  )
) +
  
  geom_line(
    linewidth = 1.05
  ) +
  
  scale_color_manual(
    values = sex_colours
  ) +
  
  scale_linetype_manual(
    values = sex_line_types
  ) +
  
  scale_x_continuous(
    name = "Year",
    breaks = year_breaks,
    limits = c(
      start_year,
      end_year
    ),
    expand = expansion(
      mult = c(
        0.01,
        0.02
      )
    )
  ) +
  
  scale_y_continuous(
    name = "DALY ASR\nper 100,000",
    labels = scales::label_number(
      accuracy = 1
    ),
    expand = expansion(
      mult = c(
        0.05,
        0.08
      )
    )
  ) +
  
  labs(
    title = "C. Disability-adjusted burden"
  ) +
  
  common_theme +
  
  theme(
    legend.position = "none"
  )


# ==============================================================================
# 30. FIGURE 5D — DALYs PER CASE
# ==============================================================================

figure_5d <- ggplot(
  figure_data,
  aes(
    x = year,
    y = dalys_per_case,
    color = sex,
    linetype = sex
  )
) +
  
  geom_line(
    linewidth = 1.05
  ) +
  
  scale_color_manual(
    values = sex_colours
  ) +
  
  scale_linetype_manual(
    values = sex_line_types
  ) +
  
  scale_x_continuous(
    name = "Year",
    breaks = year_breaks,
    limits = c(
      start_year,
      end_year
    ),
    expand = expansion(
      mult = c(
        0.01,
        0.02
      )
    )
  ) +
  
  scale_y_continuous(
    name = "DALYs per incident case",
    labels = scales::label_number(
      accuracy = 0.1
    ),
    expand = expansion(
      mult = c(
        0.05,
        0.08
      )
    )
  ) +
  
  labs(
    title = "D. Per-case severity"
  ) +
  
  common_theme +
  
  theme(
    legend.position = "none"
  )


# ==============================================================================
# 31. FIGURE 5E — MIR
# ==============================================================================

figure_5e <- ggplot(
  figure_data,
  aes(
    x = year,
    y = mir,
    color = sex,
    linetype = sex
  )
) +
  
  geom_line(
    linewidth = 1.05
  ) +
  
  scale_color_manual(
    values = sex_colours
  ) +
  
  scale_linetype_manual(
    values = sex_line_types
  ) +
  
  scale_x_continuous(
    name = "Year",
    breaks = year_breaks,
    limits = c(
      start_year,
      end_year
    ),
    expand = expansion(
      mult = c(
        0.01,
        0.02
      )
    )
  ) +
  
  scale_y_continuous(
    name = "Mortality-to-incidence ratio",
    labels = scales::label_number(
      accuracy = 0.01
    ),
    expand = expansion(
      mult = c(
        0.05,
        0.08
      )
    )
  ) +
  
  labs(
    title = "E. Mortality relative to incidence"
  ) +
  
  common_theme


# ==============================================================================
# 32. COMBINE FIGURE 5
# ==============================================================================

figure_5 <- (
  figure_5a +
    figure_5b
) /
  (
    figure_5c +
      figure_5d
  ) /
  figure_5e +
  
  plot_layout(
    guides = "collect",
    heights = c(
      1,
      1,
      0.90
    )
  ) +
  
  plot_annotation(
    
    title = paste0(
      "Global temporal trends in colorectal cancer burden ",
      "and severity, 1990–2023"
    ),
    
    caption = paste0(
      "ASR denotes age-standardized rate per 100,000 population. ",
      "DALYs per case and the mortality-to-incidence ratio were ",
      "calculated from annual number estimates."
    ),
    
    theme = theme(
      
      plot.title = element_text(
        face = "bold",
        size = 16,
        hjust = 0.5
      ),
      
      plot.caption = element_text(
        size = 9,
        hjust = 0
      ),
      
      plot.margin = margin(
        12,
        16,
        12,
        12
      )
      
    )
    
  ) &
  
  theme(
    legend.position = "bottom"
  )


# ==============================================================================
# 33. SAVE FIGURE 5
# ==============================================================================

figure_5_tiff_file <- file.path(
  figures_folder,
  "Figure_5_Global_temporal_trends_1990_2023.tiff"
)

figure_5_png_file <- file.path(
  figures_folder,
  "Figure_5_Global_temporal_trends_1990_2023.png"
)

ggsave(
  filename = figure_5_tiff_file,
  plot = figure_5,
  device = "tiff",
  width = 13,
  height = 14,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white"
)

ggsave(
  filename = figure_5_png_file,
  plot = figure_5,
  device = "png",
  width = 13,
  height = 14,
  units = "in",
  dpi = 300,
  bg = "white"
)


# ==============================================================================
# 34. EXTRACT RESULTS FOR AUTOMATIC TEXT
# ==============================================================================

get_result <- function(
    metric_name,
    sex_name = "Both"
) {
  
  trend_summary |>
    
    filter(
      as.character(
        metric
      ) == metric_name,
      as.character(
        sex
      ) == sex_name
    ) |>
    
    slice_head(
      n = 1
    )
  
}

both_incidence <- get_result(
  "Incidence ASR"
)

both_mortality <- get_result(
  "Mortality ASR"
)

both_daly_asr <- get_result(
  "DALY ASR"
)

both_dalys_per_case <- get_result(
  "DALYs per case"
)

both_mir <- get_result(
  "Mortality-to-incidence ratio"
)

male_dalys_per_case <- get_result(
  "DALYs per case",
  "Male"
)

female_dalys_per_case <- get_result(
  "DALYs per case",
  "Female"
)

male_mir <- get_result(
  "Mortality-to-incidence ratio",
  "Male"
)

female_mir <- get_result(
  "Mortality-to-incidence ratio",
  "Female"
)


# ==============================================================================
# 35. IDENTIFY EXTREME YEARS
# ==============================================================================

both_annual <- annual_trends |>
  
  filter(
    sex == "Both"
  )

highest_incidence_year <- both_annual |>
  
  slice_max(
    incidence_asr,
    n = 1,
    with_ties = FALSE
  )

lowest_incidence_year <- both_annual |>
  
  slice_min(
    incidence_asr,
    n = 1,
    with_ties = FALSE
  )

highest_severity_year <- both_annual |>
  
  slice_max(
    dalys_per_case,
    n = 1,
    with_ties = FALSE
  )

lowest_severity_year <- both_annual |>
  
  slice_min(
    dalys_per_case,
    n = 1,
    with_ties = FALSE
  )

highest_mir_year <- both_annual |>
  
  slice_max(
    mir,
    n = 1,
    with_ties = FALSE
  )

lowest_mir_year <- both_annual |>
  
  slice_min(
    mir,
    n = 1,
    with_ties = FALSE
  )


# ==============================================================================
# 36. GENERATE RESULTS TEXT
# ==============================================================================

results_text <- glue::glue(
  "3.5 Global temporal trends in burden and severity, 1990–2023

Global colorectal cancer burden and per-case severity followed distinct temporal patterns between 1990 and 2023 (Figure 5; Table 5). Among both sexes, the age-standardized incidence rate changed from {format_number(both_incidence$start_value, 2)} per 100,000 in 1990 to {format_number(both_incidence$end_value, 2)} per 100,000 in 2023, representing a {format_change(both_incidence$percentage_change)} change. Over the complete study period, incidence {trend_direction(both_incidence$aapc, both_incidence$p_value)}, with an average annual change of {format_number(both_incidence$aapc, 2)}% (95% CI {format_number(both_incidence$lower_95_ci, 2)} to {format_number(both_incidence$upper_95_ci, 2)}; p {format_p_value(both_incidence$p_value)}). The highest incidence rate was {format_number(highest_incidence_year$incidence_asr, 2)} per 100,000 in {highest_incidence_year$year}, whereas the lowest was {format_number(lowest_incidence_year$incidence_asr, 2)} per 100,000 in {lowest_incidence_year$year}.

The age-standardized mortality rate changed from {format_number(both_mortality$start_value, 2)} per 100,000 in 1990 to {format_number(both_mortality$end_value, 2)} per 100,000 in 2023 ({format_change(both_mortality$percentage_change)}). Mortality {trend_direction(both_mortality$aapc, both_mortality$p_value)}, with an average annual change of {format_number(both_mortality$aapc, 2)}% (95% CI {format_number(both_mortality$lower_95_ci, 2)} to {format_number(both_mortality$upper_95_ci, 2)}; p {format_p_value(both_mortality$p_value)}). The age-standardized DALY rate changed from {format_number(both_daly_asr$start_value, 2)} to {format_number(both_daly_asr$end_value, 2)} per 100,000, corresponding to a {format_change(both_daly_asr$percentage_change)} change and an average annual change of {format_number(both_daly_asr$aapc, 2)}% (95% CI {format_number(both_daly_asr$lower_95_ci, 2)} to {format_number(both_daly_asr$upper_95_ci, 2)}; p {format_p_value(both_daly_asr$p_value)}).

DALYs per incident case changed from {format_number(both_dalys_per_case$start_value, 2)} in 1990 to {format_number(both_dalys_per_case$end_value, 2)} in 2023, representing a {format_change(both_dalys_per_case$percentage_change)} change. DALYs per case {trend_direction(both_dalys_per_case$aapc, both_dalys_per_case$p_value)}, with an average annual change of {format_number(both_dalys_per_case$aapc, 2)}% (95% CI {format_number(both_dalys_per_case$lower_95_ci, 2)} to {format_number(both_dalys_per_case$upper_95_ci, 2)}; p {format_p_value(both_dalys_per_case$p_value)}). Per-case severity was highest in {highest_severity_year$year} at {format_number(highest_severity_year$dalys_per_case, 2)} DALYs per case and lowest in {lowest_severity_year$year} at {format_number(lowest_severity_year$dalys_per_case, 2)}.

The mortality-to-incidence ratio changed from {format_number(both_mir$start_value, 3)} in 1990 to {format_number(both_mir$end_value, 3)} in 2023, representing a {format_change(both_mir$percentage_change)} change. The ratio {trend_direction(both_mir$aapc, both_mir$p_value)}, with an average annual change of {format_number(both_mir$aapc, 2)}% (95% CI {format_number(both_mir$lower_95_ci, 2)} to {format_number(both_mir$upper_95_ci, 2)}; p {format_p_value(both_mir$p_value)}). The highest ratio was {format_number(highest_mir_year$mir, 3)} in {highest_mir_year$year}, and the lowest was {format_number(lowest_mir_year$mir, 3)} in {lowest_mir_year$year}.

Sex-specific patterns were broadly assessed over the same period. DALYs per case changed from {format_number(male_dalys_per_case$start_value, 2)} to {format_number(male_dalys_per_case$end_value, 2)} among males and from {format_number(female_dalys_per_case$start_value, 2)} to {format_number(female_dalys_per_case$end_value, 2)} among females. The mortality-to-incidence ratio changed from {format_number(male_mir$start_value, 3)} to {format_number(male_mir$end_value, 3)} among males and from {format_number(female_mir$start_value, 3)} to {format_number(female_mir$end_value, 3)} among females. Complete annual estimates are provided in Supplementary Table S4."
)


# ==============================================================================
# 37. SAVE RESULTS TEXT
# ==============================================================================

results_text_file <- file.path(
  results_text_folder,
  "Results_3_5_Global_temporal_trends_1990_2023.txt"
)

writeLines(
  text = results_text,
  con = results_text_file,
  useBytes = TRUE
)


# ==============================================================================
# 38. DISPLAY MAIN RESULTS
# ==============================================================================

message("")
message("==============================================================")
message("TABLE 5")
message("==============================================================")

print(
  table_5_formatted,
  n = Inf
)

message("")
message("==============================================================")
message("RESULTS TEXT")
message("==============================================================")

cat(
  results_text
)


# ==============================================================================
# 39. VERIFY OUTPUT FILES
# ==============================================================================

required_output_files <- c(
  figure_5_tiff_file,
  figure_5_png_file,
  table_5_file,
  s4_excel_file,
  s4_csv_file,
  results_text_file
)

missing_output_files <- required_output_files[
  !file.exists(
    required_output_files
  )
]

if (length(missing_output_files) > 0) {
  
  stop(
    paste0(
      "\nThe following outputs were not created:\n",
      paste(
        missing_output_files,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 40. COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("RESULTS SECTION 3.5 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message(
  "Years analysed: ",
  start_year,
  "–",
  end_year
)

message(
  "Measures used: Incidence, Mortality, DALYs"
)

message(
  "Sexes used: ",
  paste(
    sex_order,
    collapse = ", "
  )
)

message("")
message("Figure 5:")
message(figure_5_tiff_file)

message("")
message("Table 5:")
message(table_5_file)

message("")
message("Supplementary Table S4:")
message(s4_excel_file)

message("")
message("Results text:")
message(results_text_file)

message("")
message("No YLL or YLD data were requested or analysed.")
message("The Clean_Data folder was not used.")
message("==============================================================")