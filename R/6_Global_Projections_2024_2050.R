# ==============================================================================
# R_03_6_Global_Projections_2024_2050.R
#
# RESULTS SECTION 3.6
# Global projections of colorectal cancer burden, 2024–2050
#
# OUTCOMES
#   Incident cases
#   Deaths
#   DALYs
#
# SCENARIOS
#
#   Scenario 1:
#     Population growth only.
#     The observed 2023 burden is carried forward to 2024 and subsequently
#     scaled according to changes in total global population relative to 2024.
#
#   Scenario 2:
#     Population growth plus an attenuated continuation of recent epidemiological trends.
#     Scenario 1 estimates are multiplied by the recent 2010–2023
#     age-standardized-rate trend using progressive trend attenuation.
#
# IMPORTANT METHODOLOGICAL LIMITATION
#   The available UN population file contains total population rather than
#   age-specific population. Consequently, Scenario 1 represents population
#   growth only and must not be described as population growth and ageing.
#
# INPUTS
#
# Global/
#   Global colorectal cancer data containing:
#     Incidence: Number and age-standardized rate
#     Mortality: Number and age-standardized rate
#     DALYs: Number and age-standardized rate
#     Both sexes
#     1990–2023
#
# Population/
#   UN_WPP_Country_Population_2024_2050.csv
#
# REQUIRED POPULATION COLUMNS
#   Country/location
#   Year
#   Population
#
# OUTPUTS
#
# Publication/Figures/
#   Figure_6_Global_CRC_projections_2024_2050.tiff
#
# Publication/Tables/
#   Table_6_Global_CRC_projections_2024_2050.xlsx
#
# Publication/Supplementary/
#   Table_S5_Annual_global_CRC_projections_2024_2050.xlsx
#   Table_S5_Annual_global_CRC_projections_2024_2050.csv
#
# Publication/Results_text/
#   Results_3_6_Global_CRC_projections_2024_2050.txt
#
# NO YLL OR YLD DATA ARE USED.
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

required_project_folders <- c(
  "Global",
  "Population",
  "Publication"
)

current_is_project <- all(
  dir.exists(
    file.path(
      current_folder,
      required_project_folders
    )
  )
)

parent_is_project <- all(
  dir.exists(
    file.path(
      parent_folder,
      required_project_folders
    )
  )
)

if (current_is_project) {
  
  project_root <- current_folder
  
} else if (parent_is_project) {
  
  project_root <- parent_folder
  
} else {
  
  stop(
    paste0(
      "\nThe colorectal cancer project root could not be identified.\n\n",
      "The project root must contain:\n",
      paste(
        required_project_folders,
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

global_folder <- file.path(
  project_root,
  "Global"
)

population_folder <- file.path(
  project_root,
  "Population"
)

population_file <- file.path(
  population_folder,
  "UN_WPP_Country_Population_2024_2050.csv"
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

if (!file.exists(population_file)) {
  
  stop(
    paste0(
      "\nPopulation file was not found:\n",
      population_file
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 5. ANALYTICAL SETTINGS
# ==============================================================================

historical_start_year <- 1990

baseline_burden_year <- 2023

population_baseline_year <- 2024

projection_end_year <- 2050

trend_start_year <- 2010

trend_end_year <- 2023

projection_years <- population_baseline_year:projection_end_year

outcome_order <- c(
  "Incidence",
  "Mortality",
  "DALYs"
)

scenario_order <- c(
  "Scenario 1: Population growth only",
  "Scenario 2: Population growth plus attenuated trend"
)

# Recent epidemiological trends are progressively attenuated during projection
# rather than extrapolated unchanged through 2050.
# An attenuation factor of 0.90 means that 90% of the preceding year's
# log-scale trend contribution is retained in each successive projection year.

trend_attenuation_factor <- 0.90


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
      "males"
    ) ~
      "Male",
    
    value %in% c(
      "female",
      "females"
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


standardise_metric <- function(
    metric,
    age
) {
  
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
  message("Reading:")
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


calculate_annual_change <- function(data) {
  
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
        annual_change = NA_real_,
        lower_95_ci = NA_real_,
        upper_95_ci = NA_real_,
        p_value = NA_real_,
        r_squared = NA_real_
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
    
    annual_change =
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
      model_summary$r.squared
    
  )
  
}


format_number <- function(
    x,
    digits = 0
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


format_percent <- function(
    x,
    digits = 1
) {
  
  ifelse(
    is.na(x),
    "NA",
    paste0(
      ifelse(
        x > 0,
        "+",
        ""
      ),
      formatC(
        x,
        format = "f",
        digits = digits
      ),
      "%"
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


# ==============================================================================
# 7. IMPORT GLOBAL HISTORICAL DATA
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
message("Global input files:")

for (current_file in global_files) {
  
  message(current_file)
  
}

imported_global_list <- vector(
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
        )
      
    ) |>
    
    mutate(
      
      analysis_metric =
        standardise_metric(
          metric_original,
          age
        )
      
    )
  
  imported_global_list[[i]] <- standardised_data
  
}

global_raw <- bind_rows(
  imported_global_list
)


# ==============================================================================
# 8. FILTER GLOBAL BOTH-SEX COLORECTAL CANCER DATA
# ==============================================================================

global_data <- global_raw |>
  
  filter(
    
    stringr::str_detect(
      stringr::str_to_lower(
        cause
      ),
      "colon|rectum|colorectal"
    ),
    
    sex == "Both",
    
    measure %in% outcome_order,
    
    analysis_metric %in% c(
      "Number",
      "ASR"
    ),
    
    year >= historical_start_year,
    
    year <= baseline_burden_year
    
  )

if (
  any(
    stringr::str_to_lower(
      global_data$location
    ) %in% c(
      "global",
      "world"
    ),
    na.rm = TRUE
  )
) {
  
  global_data <- global_data |>
    
    filter(
      stringr::str_to_lower(
        location
      ) %in% c(
        "global",
        "world"
      )
    )
  
}

global_data <- global_data |>
  
  group_by(
    measure,
    analysis_metric,
    year
  ) |>
  
  summarise(
    value = first_non_missing(value),
    .groups = "drop"
  )


# ==============================================================================
# 9. VALIDATE GLOBAL INPUTS
# ==============================================================================

required_global_combinations <- tidyr::expand_grid(
  
  measure =
    outcome_order,
  
  analysis_metric =
    c(
      "Number",
      "ASR"
    )
  
)

available_global_combinations <- global_data |>
  
  distinct(
    measure,
    analysis_metric
  )

missing_global_combinations <- anti_join(
  required_global_combinations,
  available_global_combinations,
  by = c(
    "measure",
    "analysis_metric"
  )
)

if (nrow(missing_global_combinations) > 0) {
  
  print(
    missing_global_combinations
  )
  
  stop(
    paste0(
      "\nGlobal data are missing required incidence, mortality, or DALY ",
      "number/rate combinations."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 10. EXTRACT 2023 BASELINE NUMBERS
# ==============================================================================

baseline_numbers <- global_data |>
  
  filter(
    year == baseline_burden_year,
    analysis_metric == "Number"
  ) |>
  
  select(
    outcome = measure,
    baseline_number = value
  )

if (nrow(baseline_numbers) != length(outcome_order)) {
  
  stop(
    "\nThe 2023 baseline numbers were not available for all three outcomes.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 11. ESTIMATE RECENT 2010–2023 ASR TRENDS
# ==============================================================================

recent_asr_data <- global_data |>
  filter(
    analysis_metric == "ASR",
    year >= trend_start_year,
    year <= trend_end_year
  ) |>
  select(
    outcome = measure,
    year,
    value
  )

recent_trends <- recent_asr_data |>
  group_by(outcome) |>
  group_modify(~ calculate_annual_change(.x)) |>
  ungroup() |>
  mutate(
    annual_log_trend = log(1 + annual_change / 100)
  )

if (nrow(recent_trends) != length(outcome_order)) {
  stop(
    "\nRecent ASR trends could not be estimated for all three outcomes.",
    call. = FALSE
  )
}

if (any(!is.finite(recent_trends$annual_log_trend))) {
  print(recent_trends)
  stop(
    "\nAt least one 2010–2023 annual trend could not be converted to the log scale.",
    call. = FALSE
  )
}

message("")
message("Recent 2010–2023 ASR trends:")
print(recent_trends)


# 12. IMPORT UN POPULATION DATA
# ==============================================================================

population_raw <- readr::read_csv(
  population_file,
  show_col_types = FALSE,
  progress = FALSE,
  guess_max = 500000
) |>
  
  janitor::clean_names()

message("")
message("Population columns detected:")

print(
  names(population_raw)
)

population_location_column <- find_column(
  population_raw,
  c(
    "location",
    "location_name",
    "country",
    "country_name",
    "region_subregion_country_or_area"
  ),
  "population location"
)

population_year_column <- find_column(
  population_raw,
  c(
    "year",
    "year_id"
  ),
  "population year"
)

population_value_column <- find_column(
  population_raw,
  c(
    "population",
    "population_persons",
    "total_population",
    "pop"
  ),
  "population"
)


# ==============================================================================
# 13. PREPARE TOTAL GLOBAL POPULATION
# ==============================================================================

population_data <- population_raw |>
  
  transmute(
    
    location =
      clean_text(
        .data[[population_location_column]]
      ),
    
    year =
      as.integer(
        .data[[population_year_column]]
      ),
    
    population =
      to_numeric_safely(
        .data[[population_value_column]]
      )
    
  ) |>
  
  filter(
    year >= population_baseline_year,
    year <= projection_end_year,
    !is.na(population),
    population > 0
  )


# Exclude aggregate regions if the file contains a standard country-level
# location set plus regional/global rows. The first preference is to use rows
# marked Global/World directly. Otherwise, country rows are summed.

global_population_rows <- population_data |>
  
  filter(
    stringr::str_to_lower(
      location
    ) %in% c(
      "global",
      "world"
    )
  )

if (
  nrow(global_population_rows) ==
  length(projection_years)
) {
  
  global_population <- global_population_rows |>
    
    group_by(
      year
    ) |>
    
    summarise(
      population = first_non_missing(population),
      .groups = "drop"
    )
  
  population_method <-
    "Global population values were taken directly from the UN file."
  
} else {
  
  # Names commonly used for aggregate rows are removed before summing.
  aggregate_patterns <- paste0(
    "world|global|income|region|subregion|continent|",
    "africa|asia|europe|oceania|america|",
    "less developed|more developed|least developed|",
    "landlocked|small island"
  )
  
  country_population <- population_data |>
    
    filter(
      !stringr::str_detect(
        stringr::str_to_lower(
          location
        ),
        aggregate_patterns
      )
    )
  
  global_population <- country_population |>
    
    group_by(
      year
    ) |>
    
    summarise(
      population = sum(
        population,
        na.rm = TRUE
      ),
      .groups = "drop"
    )
  
  population_method <- paste0(
    "Global population was reconstructed by summing the country and ",
    "territory population estimates retained from the UN file."
  )
  
}


# ==============================================================================
# 14. VALIDATE POPULATION SERIES
# ==============================================================================

missing_population_years <- setdiff(
  projection_years,
  global_population$year
)

if (length(missing_population_years) > 0) {
  
  stop(
    paste0(
      "\nPopulation estimates are missing for:\n",
      paste(
        missing_population_years,
        collapse = ", "
      )
    ),
    call. = FALSE
  )
  
}

if (
  any(
    duplicated(
      global_population$year
    )
  )
) {
  
  stop(
    "\nDuplicate global population years were detected.",
    call. = FALSE
  )
  
}

baseline_population <- global_population |>
  
  filter(
    year == population_baseline_year
  ) |>
  
  pull(
    population
  )

if (
  length(baseline_population) != 1 ||
  is.na(baseline_population) ||
  baseline_population <= 0
) {
  
  stop(
    "\nThe 2024 global population baseline could not be identified.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 15. CREATE PROJECTION DATASET
# ==============================================================================

projection_grid <- tidyr::crossing(
  year = projection_years,
  outcome = outcome_order
) |>
  left_join(
    global_population,
    by = "year"
  ) |>
  left_join(
    baseline_numbers,
    by = "outcome"
  ) |>
  left_join(
    recent_trends |>
      select(
        outcome,
        annual_change,
        lower_95_ci,
        upper_95_ci,
        p_value,
        r_squared,
        annual_log_trend
      ),
    by = "outcome"
  ) |>
  mutate(
    population_index = population / baseline_population,
    years_after_2023 = year - baseline_burden_year,
    scenario_1 = baseline_number * population_index,
    
    # Damped epidemiological trend:
    # multiplier = exp{ beta * [1 - a^t] / [1 - a] }
    # beta = annual log-scale ASR trend from 2010–2023
    # a    = trend attenuation factor (0.90)
    # t    = years after 2023
    epidemiological_trend_factor =
      exp(
        annual_log_trend *
          (1 - trend_attenuation_factor ^ years_after_2023) /
          (1 - trend_attenuation_factor)
      ),
    
    scenario_2 = scenario_1 * epidemiological_trend_factor
  )


# 16. VALIDATE PROJECTED VALUES
# ==============================================================================

invalid_projections <- projection_grid |>
  filter(
    is.na(scenario_1) |
      is.na(scenario_2) |
      is.na(epidemiological_trend_factor) |
      !is.finite(scenario_1) |
      !is.finite(scenario_2) |
      !is.finite(epidemiological_trend_factor) |
      scenario_1 < 0 |
      scenario_2 < 0 |
      epidemiological_trend_factor <= 0
  )

if (nrow(invalid_projections) > 0) {
  print(invalid_projections)
  stop(
    "\nInvalid projection estimates were generated.",
    call. = FALSE
  )
}


# 17. ADD OBSERVED 1990–2023 NUMBER SERIES
# ==============================================================================

observed_numbers <- global_data |>
  
  filter(
    analysis_metric == "Number"
  ) |>
  
  transmute(
    
    year,
    
    outcome =
      measure,
    
    value,
    
    series =
      "Observed"
    
  )


projection_long <- projection_grid |>
  
  select(
    year,
    outcome,
    scenario_1,
    scenario_2
  ) |>
  
  pivot_longer(
    
    cols = c(
      scenario_1,
      scenario_2
    ),
    
    names_to = "scenario_code",
    
    values_to = "value"
    
  ) |>
  
  mutate(
    
    series = case_when(
      
      scenario_code ==
        "scenario_1" ~
        "Scenario 1: Population growth only",
      
      scenario_code ==
        "scenario_2" ~
        "Scenario 2: Population growth plus attenuated trend",
      
      TRUE ~
        scenario_code
      
    )
    
  ) |>
  
  select(
    year,
    outcome,
    value,
    series
  )


figure_data <- bind_rows(
  
  observed_numbers,
  
  projection_long
  
) |>
  
  mutate(
    
    outcome = factor(
      outcome,
      levels = outcome_order
    ),
    
    series = factor(
      series,
      levels = c(
        "Observed",
        scenario_order
      )
    )
    
  )


# ==============================================================================
# 18. CREATE TABLE 6 NUMERIC DATA
# ==============================================================================

projection_2050 <- projection_grid |>
  
  filter(
    year == projection_end_year
  ) |>
  
  mutate(
    
    scenario_1_absolute_change =
      scenario_1 -
      baseline_number,
    
    scenario_1_percent_change =
      100 *
      (
        scenario_1 -
          baseline_number
      ) /
      baseline_number,
    
    scenario_2_absolute_change =
      scenario_2 -
      baseline_number,
    
    scenario_2_percent_change =
      100 *
      (
        scenario_2 -
          baseline_number
      ) /
      baseline_number
    
  ) |>
  
  arrange(
    factor(
      outcome,
      levels = outcome_order
    )
  )


table_6_numeric <- projection_2050 |>
  
  transmute(
    
    Outcome =
      outcome,
    
    Observed_2023 =
      baseline_number,
    
    Projected_2050_Scenario_1 =
      scenario_1,
    
    Absolute_change_Scenario_1 =
      scenario_1_absolute_change,
    
    Percent_change_Scenario_1 =
      scenario_1_percent_change,
    
    Projected_2050_Scenario_2 =
      scenario_2,
    
    Absolute_change_Scenario_2 =
      scenario_2_absolute_change,
    
    Percent_change_Scenario_2 =
      scenario_2_percent_change,
    
    Recent_ASR_annual_change_percent =
      annual_change,
    
    Annual_log_trend_beta =
      annual_log_trend,
    
    Trend_attenuation_factor =
      trend_attenuation_factor,
    
    Trend_P_value =
      p_value
    
  )


# ==============================================================================
# 19. CREATE PUBLICATION-FORMATTED TABLE 6
# ==============================================================================

table_6_formatted <- table_6_numeric |>
  
  transmute(
    
    Outcome,
    
    `Observed 2023` =
      format_number(
        Observed_2023,
        0
      ),
    
    `2050 Scenario 1` =
      format_number(
        Projected_2050_Scenario_1,
        0
      ),
    
    `Scenario 1 change, %` =
      format_percent(
        Percent_change_Scenario_1,
        1
      ),
    
    `2050 Scenario 2` =
      format_number(
        Projected_2050_Scenario_2,
        0
      ),
    
    `Scenario 2 change, %` =
      format_percent(
        Percent_change_Scenario_2,
        1
      ),
    
    `Recent annual ASR change, %` =
      format_number(
        Recent_ASR_annual_change_percent,
        2
      ),
    
    `Trend attenuation factor` =
      format_number(
        Trend_attenuation_factor,
        2
      )
    
  )


# ==============================================================================
# 20. EXPORT TABLE 6
# ==============================================================================

table_6_file <- file.path(
  tables_folder,
  "Table_6_Global_CRC_projections_2024_2050.xlsx"
)

table_6_workbook <- createWorkbook()

addWorksheet(
  table_6_workbook,
  "Publication table"
)

writeData(
  table_6_workbook,
  sheet = "Publication table",
  x = paste0(
    "Table 6. Projected global colorectal cancer burden ",
    "under two scenarios, 2024–2050"
  ),
  startRow = 1,
  startCol = 1
)

writeData(
  table_6_workbook,
  sheet = "Publication table",
  x = table_6_formatted,
  startRow = 3,
  startCol = 1
)

table_note <- paste0(
  "Note: Scenario 1 incorporates changes in total global population only. ",
  "Scenario 2 additionally incorporates the ",
  trend_start_year,
  "–",
  trend_end_year,
  " age-standardized-rate trend using a damped-trend projection. ",
  "The log-scale trend contribution is attenuated by 10% in each successive ",
  "projection year (attenuation factor = ",
  trend_attenuation_factor,
  "). Because age-specific population projections were unavailable, ",
  "population ageing was not modelled separately."
)

writeData(
  table_6_workbook,
  sheet = "Publication table",
  x = table_note,
  startRow = nrow(table_6_formatted) + 6,
  startCol = 1,
  colNames = FALSE
)

addWorksheet(
  table_6_workbook,
  "Numeric results"
)

writeData(
  table_6_workbook,
  sheet = "Numeric results",
  x = table_6_numeric
)

addWorksheet(
  table_6_workbook,
  "Recent trend estimates"
)

writeData(
  table_6_workbook,
  sheet = "Recent trend estimates",
  x = recent_trends
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
  table_6_workbook,
  sheet = "Publication table",
  style = title_style,
  rows = 1,
  cols = 1:ncol(table_6_formatted),
  gridExpand = TRUE
)

addStyle(
  table_6_workbook,
  sheet = "Publication table",
  style = header_style,
  rows = 3,
  cols = 1:ncol(table_6_formatted),
  gridExpand = TRUE
)

addStyle(
  table_6_workbook,
  sheet = "Publication table",
  style = body_style,
  rows = 4:(nrow(table_6_formatted) + 3),
  cols = 1:ncol(table_6_formatted),
  gridExpand = TRUE
)

setColWidths(
  table_6_workbook,
  sheet = "Publication table",
  cols = 1:ncol(table_6_formatted),
  widths = c(
    18,
    18,
    20,
    21,
    20,
    21,
    25,
    21
  )
)

setRowHeights(
  table_6_workbook,
  sheet = "Publication table",
  rows = 3,
  heights = 55
)

freezePane(
  table_6_workbook,
  sheet = "Publication table",
  firstActiveRow = 4,
  firstActiveCol = 2
)

for (
  sheet_name in c(
    "Numeric results",
    "Recent trend estimates"
  )
) {
  
  sheet_data <- if (
    sheet_name == "Numeric results"
  ) {
    
    table_6_numeric
    
  } else {
    
    recent_trends
    
  }
  
  addStyle(
    table_6_workbook,
    sheet = sheet_name,
    style = header_style,
    rows = 1,
    cols = 1:ncol(sheet_data),
    gridExpand = TRUE
  )
  
  setColWidths(
    table_6_workbook,
    sheet = sheet_name,
    cols = 1:ncol(sheet_data),
    widths = "auto"
  )
  
  freezePane(
    table_6_workbook,
    sheet = sheet_name,
    firstRow = TRUE
  )
  
}

saveWorkbook(
  table_6_workbook,
  table_6_file,
  overwrite = TRUE
)


# ==============================================================================
# 21. CREATE SUPPLEMENTARY TABLE S5
# ==============================================================================

supplementary_s5 <- projection_grid |>
  
  arrange(
    year,
    factor(
      outcome,
      levels = outcome_order
    )
  ) |>
  
  transmute(
    
    Year =
      year,
    
    Outcome =
      outcome,
    
    Global_population =
      population,
    
    Population_index_relative_to_2024 =
      population_index,
    
    Observed_2023_baseline =
      baseline_number,
    
    Recent_ASR_annual_change_percent =
      annual_change,
    
    Annual_log_trend_beta =
      annual_log_trend,
    
    Trend_attenuation_factor =
      trend_attenuation_factor,
    
    Epidemiological_trend_factor =
      epidemiological_trend_factor,
    
    Scenario_1_population_growth_only =
      scenario_1,
    
    Scenario_2_population_plus_recent_trend =
      scenario_2
    
  )

s5_excel_file <- file.path(
  supplementary_folder,
  "Table_S5_Annual_global_CRC_projections_2024_2050.xlsx"
)

s5_csv_file <- file.path(
  supplementary_folder,
  "Table_S5_Annual_global_CRC_projections_2024_2050.csv"
)

s5_workbook <- createWorkbook()

addWorksheet(
  s5_workbook,
  "Annual projections"
)

writeData(
  s5_workbook,
  sheet = "Annual projections",
  x = supplementary_s5
)

addWorksheet(
  s5_workbook,
  "2050 summary"
)

writeData(
  s5_workbook,
  sheet = "2050 summary",
  x = table_6_numeric
)

addWorksheet(
  s5_workbook,
  "Population series"
)

writeData(
  s5_workbook,
  sheet = "Population series",
  x = global_population
)

addWorksheet(
  s5_workbook,
  "Projection assumptions"
)

projection_assumptions <- tibble(
  Item = c(
    "Observed burden baseline year",
    "Population baseline year",
    "Projection end year",
    "Historical trend period",
    "Trend projection method",
    "Annual trend attenuation factor",
    "Population methodology",
    "Ageing modelled separately",
    "YLL data used",
    "YLD data used"
  ),
  Value = c(
    baseline_burden_year,
    population_baseline_year,
    projection_end_year,
    paste0(trend_start_year, "–", trend_end_year),
    "Damped log-linear trend",
    trend_attenuation_factor,
    population_method,
    "No",
    "No",
    "No"
  )
)

writeData(
  s5_workbook,
  sheet = "Projection assumptions",
  x = projection_assumptions
)

for (
  sheet_name in c(
    "Annual projections",
    "2050 summary",
    "Population series",
    "Projection assumptions"
  )
) {
  
  sheet_data <- switch(
    
    sheet_name,
    
    "Annual projections" =
      supplementary_s5,
    
    "2050 summary" =
      table_6_numeric,
    
    "Population series" =
      global_population,
    
    "Projection assumptions" =
      projection_assumptions
    
  )
  
  addStyle(
    s5_workbook,
    sheet = sheet_name,
    style = header_style,
    rows = 1,
    cols = 1:ncol(sheet_data),
    gridExpand = TRUE
  )
  
  setColWidths(
    s5_workbook,
    sheet = sheet_name,
    cols = 1:ncol(sheet_data),
    widths = "auto"
  )
  
  freezePane(
    s5_workbook,
    sheet = sheet_name,
    firstRow = TRUE
  )
  
}

saveWorkbook(
  s5_workbook,
  s5_excel_file,
  overwrite = TRUE
)

readr::write_csv(
  supplementary_s5,
  s5_csv_file
)


# ==============================================================================
# 22. PREPARE FIGURE SETTINGS
# ==============================================================================

series_colours <- c(
  
  "Observed" =
    "#222222",
  
  "Scenario 1: Population growth only" =
    "#2878B5",
  
  "Scenario 2: Population growth plus attenuated trend" =
    "#C94C4C"
  
)

series_line_types <- c(
  
  "Observed" =
    "solid",
  
  "Scenario 1: Population growth only" =
    "dashed",
  
  "Scenario 2: Population growth plus attenuated trend" =
    "dotdash"
  
)

common_theme <- theme_classic(
  base_size = 12
) +
  
  theme(
    
    plot.title = element_text(
      face = "bold",
      size = 13
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
      size = 9.5
    ),
    
    plot.margin = margin(
      8,
      10,
      8,
      8
    )
    
  )

x_breaks <- c(
  1990,
  2000,
  2010,
  2020,
  2023,
  2030,
  2040,
  2050
)


# ==============================================================================
# 23. CREATE ONE PROJECTION PANEL FUNCTION
# ==============================================================================

create_projection_panel <- function(
    selected_outcome,
    panel_title,
    y_axis_title
) {
  
  panel_data <- figure_data |>
    
    filter(
      outcome == selected_outcome
    )
  
  ggplot(
    panel_data,
    aes(
      x = year,
      y = value,
      color = series,
      linetype = series
    )
  ) +
    
    geom_vline(
      xintercept = 2023.5,
      linewidth = 0.5,
      linetype = "dotted",
      color = "grey35"
    ) +
    
    geom_line(
      linewidth = 1.05,
      na.rm = TRUE
    ) +
    
    scale_color_manual(
      values = series_colours,
      drop = FALSE
    ) +
    
    scale_linetype_manual(
      values = series_line_types,
      drop = FALSE
    ) +
    
    scale_x_continuous(
      name = "Year",
      breaks = x_breaks,
      limits = c(
        historical_start_year,
        projection_end_year
      ),
      expand = expansion(
        mult = c(
          0.01,
          0.02
        )
      )
    ) +
    
    scale_y_continuous(
      name = y_axis_title,
      labels = scales::label_number(
        scale_cut = scales::cut_short_scale(),
        accuracy = 0.1
      ),
      expand = expansion(
        mult = c(
          0.04,
          0.09
        )
      )
    ) +
    
    labs(
      title = panel_title
    ) +
    
    common_theme
  
}


# ==============================================================================
# 24. CREATE FIGURE 6 PANELS
# ==============================================================================

figure_6a <- create_projection_panel(
  
  selected_outcome =
    "Incidence",
  
  panel_title =
    "A. Incident cases",
  
  y_axis_title =
    "Number of incident cases"
  
) +
  
  theme(
    legend.position = "none"
  )


figure_6b <- create_projection_panel(
  
  selected_outcome =
    "Mortality",
  
  panel_title =
    "B. Deaths",
  
  y_axis_title =
    "Number of deaths"
  
) +
  
  theme(
    legend.position = "none"
  )


figure_6c <- create_projection_panel(
  
  selected_outcome =
    "DALYs",
  
  panel_title =
    "C. Disability-adjusted life years",
  
  y_axis_title =
    "Number of DALYs"
  
)


# ==============================================================================
# 25. COMBINE FIGURE 6
# ==============================================================================

figure_6 <- (
  
  figure_6a +
    figure_6b
  
) /
  
  figure_6c +
  
  plot_layout(
    guides = "collect",
    heights = c(
      1,
      1
    )
  ) +
  
  plot_annotation(
    
    title = paste0(
      "Observed and projected global colorectal cancer burden, ",
      "1990–2050"
    ),
    
    caption = paste0(
      
      "The vertical dotted line separates observed estimates from projections. ",
      
      "Scenario 1 represents total population growth only. ",
      
      "Scenario 2 additionally applies the ",
      
      trend_start_year,
      
      "–",
      
      trend_end_year,
      
      " age-standardized-rate trend with progressive 10% annual attenuation. ",
      
      "Population ageing was not modelled separately because age-specific ",
      
      "population projections were unavailable."
      
    ),
    
    theme = theme(
      
      plot.title = element_text(
        face = "bold",
        size = 16,
        hjust = 0.5
      ),
      
      plot.caption = element_text(
        size = 8.8,
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
# 26. SAVE FIGURE 6 AS ONE TIFF AT 600 DPI
# ==============================================================================

figure_6_file <- file.path(
  figures_folder,
  "Figure_6_Global_CRC_projections_2024_2050.tiff"
)

ggsave(
  filename = figure_6_file,
  plot = figure_6,
  device = "tiff",
  width = 13,
  height = 10.5,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white"
)


# ==============================================================================
# 27. EXTRACT VALUES FOR RESULTS TEXT
# ==============================================================================

incidence_2050 <- projection_2050 |>
  
  filter(
    outcome == "Incidence"
  )

mortality_2050 <- projection_2050 |>
  
  filter(
    outcome == "Mortality"
  )

dalys_2050 <- projection_2050 |>
  
  filter(
    outcome == "DALYs"
  )


# ==============================================================================
# 28. GENERATE RESULTS TEXT
# ==============================================================================

results_text <- glue::glue(
  "3.6 Global projections of colorectal cancer burden, 2024–2050

Under Scenario 1, which incorporated changes in total global population while holding the epidemiological profile constant, the global number of incident colorectal cancer cases was projected to increase from {format_number(incidence_2050$baseline_number, 0)} in 2023 to {format_number(incidence_2050$scenario_1, 0)} in 2050, representing a {format_percent(incidence_2050$scenario_1_percent_change, 1)} change. Under Scenario 2, which additionally incorporated the attenuated {trend_start_year}–{trend_end_year} age-standardized incidence trend, the projected number of incident cases in 2050 was {format_number(incidence_2050$scenario_2, 0)}, corresponding to a {format_percent(incidence_2050$scenario_2_percent_change, 1)} change from the 2023 baseline (Figure 6; Table 6).

Global deaths were projected to increase from {format_number(mortality_2050$baseline_number, 0)} in 2023 to {format_number(mortality_2050$scenario_1, 0)} in 2050 under Scenario 1 ({format_percent(mortality_2050$scenario_1_percent_change, 1)}). When the attenuated recent mortality-rate trend was incorporated, the Scenario 2 estimate was {format_number(mortality_2050$scenario_2, 0)} deaths in 2050, representing a {format_percent(mortality_2050$scenario_2_percent_change, 1)} change relative to 2023 (Figure 6; Table 6).

The global number of DALYs was projected to change from {format_number(dalys_2050$baseline_number, 0)} in 2023 to {format_number(dalys_2050$scenario_1, 0)} in 2050 under the population-growth-only scenario ({format_percent(dalys_2050$scenario_1_percent_change, 1)}). Under Scenario 2, the corresponding 2050 estimate was {format_number(dalys_2050$scenario_2, 0)} DALYs, representing a {format_percent(dalys_2050$scenario_2_percent_change, 1)} change from the observed 2023 burden.

Differences between the two scenarios reflect the contribution of attenuated recent epidemiological trends beyond changes in total population size. Annual projection estimates for incidence, mortality, and DALYs from 2024 through 2050, together with the population indices, trend estimates, and projection factors used in each scenario, are provided in Supplementary Table S5."
)


# ==============================================================================
# 29. SAVE RESULTS TEXT
# ==============================================================================

results_text_file <- file.path(
  results_text_folder,
  "Results_3_6_Global_CRC_projections_2024_2050.txt"
)

writeLines(
  text = results_text,
  con = results_text_file,
  useBytes = TRUE
)


# ==============================================================================
# 30. DISPLAY OUTPUTS
# ==============================================================================

message("")
message("==============================================================")
message("TABLE 6")
message("==============================================================")

print(
  table_6_formatted,
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
# 31. FINAL OUTPUT VALIDATION
# ==============================================================================

required_output_files <- c(
  figure_6_file,
  table_6_file,
  s5_excel_file,
  s5_csv_file,
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
      "\nThe following output files were not created:\n",
      paste(
        missing_output_files,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 32. COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("RESULTS SECTION 3.6 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message("Outcomes projected:")
message("Incidence")
message("Mortality")
message("DALYs")

message("")
message("Figure 6 — one TIFF at 600 dpi:")
message(figure_6_file)

message("")
message("Table 6:")
message(table_6_file)

message("")
message("Supplementary Table S5:")
message(s5_excel_file)

message("")
message("Results text:")
message(results_text_file)

message("")
message("Population methodology:")
message(population_method)

message("")
message("No YLL or YLD data were requested or analysed.")
message("Population ageing was not modelled separately.")
message("Scenario 2 uses a damped 2010–2023 ASR trend with 10% annual attenuation.")
message("The Clean_Data folder was not used.")
message("==============================================================")