# ==============================================================================
# R_03_7_Country_Temporal_Trends_AAPC_1990_2023.R
#
# RESULTS SECTION 3.7
# Country-level temporal trends in colorectal cancer burden and severity,
# 1990–2023
#
# PART 1 OF 3
#   Project setup
#   Raw-data import
#   Data standardisation
#   Validation
#   Country-year analytical dataset
#
# COUNTRY-LEVEL OUTCOMES
#
#   1. Incidence ASR
#   2. Mortality ASR
#   3. DALY ASR
#   4. DALYs per incident case
#   5. Mortality-to-incidence ratio
#
# FORMULAS
#
#   DALYs per case =
#       DALY number / incidence number
#
#   MIR =
#       mortality number / incidence number
#
# INPUT FOLDERS
#
#   Incidence/
#   Mortality/
#   DALYs/
#
# EXPECTED RAW DATA
#
#   Years:
#       1990–2023
#
#   Locations:
#       204 countries and territories
#
#   Sex:
#       Both
#
#   Required records:
#
#       Incidence:
#           All ages — Number
#           Age-standardized — Rate
#
#       Mortality:
#           All ages — Number
#           Age-standardized — Rate
#
#       DALYs:
#           All ages — Number
#           Age-standardized — Rate
#
# FINAL OUTPUTS CREATED IN PART 3
#
# Publication/Figures/
#   Figure_7_Country_AAPC_maps_1990_2023.tiff
#
# Publication/Tables/
#   Table_7_Country_temporal_trend_extremes_1990_2023.xlsx
#
# Publication/Supplementary/
#   Table_S6_Complete_country_AAPC_1990_2023.xlsx
#   Table_S6_Complete_country_AAPC_1990_2023.csv
#
# Publication/Results_text/
#   Results_3_7_Country_temporal_trends_1990_2023.txt
#
# IMPORTANT
#
#   - Global data are not used in this section.
#   - YLL and YLD data are not used.
#   - The old Clean_Data folder is not used.
#   - Country projections will be analysed in the next Results section.
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
  dplyr.summarise.inform = FALSE,
  warn = 1
)


# ==============================================================================
# 2. INSTALL AND LOAD REQUIRED PACKAGES
# ==============================================================================

required_packages <- c(
  "tidyverse",
  "readxl",
  "openxlsx",
  "janitor",
  "sf",
  "rnaturalearth",
  "rnaturalearthdata",
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
  library(sf)
  library(rnaturalearth)
  library(rnaturalearthdata)
  library(patchwork)
  library(scales)
  library(glue)
  
})


# ==============================================================================
# 3. IDENTIFY THE PROJECT ROOT
#
# The R project may be open from:
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

parent_folder <- normalizePath(
  file.path(
    current_folder,
    ".."
  ),
  winslash = "/",
  mustWork = TRUE
)

required_input_folders <- c(
  "Incidence",
  "Mortality",
  "DALYs"
)

current_has_inputs <- all(
  dir.exists(
    file.path(
      current_folder,
      required_input_folders
    )
  )
)

parent_has_inputs <- all(
  dir.exists(
    file.path(
      parent_folder,
      required_input_folders
    )
  )
)

if (current_has_inputs) {
  
  project_root <- current_folder
  
} else if (parent_has_inputs) {
  
  project_root <- parent_folder
  
} else {
  
  stop(
    paste0(
      "\nThe colorectal cancer project root could not be identified.\n\n",
      "The project root must contain these folders:\n",
      paste(
        required_input_folders,
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

mortality_folder <- file.path(
  project_root,
  "Mortality"
)

dalys_folder <- file.path(
  project_root,
  "DALYs"
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
# 5. ANALYTICAL SETTINGS
# ==============================================================================

analysis_start_year <- 1990

analysis_end_year <- 2023

analysis_years <- analysis_start_year:analysis_end_year

analysis_sex <- "Both"

expected_country_count <- 204

expected_year_count <- length(
  analysis_years
)

expected_country_year_rows <-
  expected_country_count *
  expected_year_count

required_measures <- c(
  "Incidence",
  "Mortality",
  "DALYs"
)

required_analysis_metrics <- c(
  "Number",
  "ASR"
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
    stringr::str_replace_all(
      "\u2019",
      "'"
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
  
  detected_columns <- intersect(
    accepted_names,
    names(data)
  )
  
  if (length(detected_columns) == 0) {
    
    if (required) {
      
      stop(
        paste0(
          "\nRequired column not found: ",
          variable_description,
          "\n\nAccepted column names:\n",
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


first_non_missing_numeric <- function(x) {
  
  available_values <- x[
    !is.na(x)
  ]
  
  if (length(available_values) == 0) {
    
    return(
      NA_real_
    )
    
  }
  
  available_values[[1]]
  
}


first_non_missing_integer <- function(x) {
  
  available_values <- x[
    !is.na(x)
  ]
  
  if (length(available_values) == 0) {
    
    return(
      NA_integer_
    )
    
  }
  
  as.integer(
    available_values[[1]]
  )
  
}


first_non_missing_character <- function(x) {
  
  available_values <- x[
    !is.na(x) &
      clean_text(x) != ""
  ]
  
  if (length(available_values) == 0) {
    
    return(
      NA_character_
    )
    
  }
  
  clean_text(
    available_values[[1]]
  )
  
}


standardise_sex <- function(x) {
  
  value <- stringr::str_to_lower(
    clean_text(x)
  )
  
  dplyr::case_when(
    
    value %in% c(
      "both",
      "both sex",
      "both sexes",
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
      "disability.?adjusted|dalys|daly"
    ) ~
      "DALYs",
    
    TRUE ~
      clean_text(x)
    
  )
  
}


standardise_analysis_metric <- function(
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
          "^all ages?$|all ages combined"
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
        "age.?standard|standardized|standardised"
      ) |
        stringr::str_detect(
          metric_value,
          "age.?standard|standardized|standardised|asr"
        )
    )
  
  dplyr::case_when(
    
    number_condition ~
      "Number",
    
    asr_condition ~
      "ASR",
    
    TRUE ~
      NA_character_
    
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


format_percent <- function(
    x,
    digits = 1,
    include_plus = FALSE
) {
  
  prefix <- dplyr::if_else(
    !is.na(x) &
      x > 0 &
      include_plus,
    "+",
    ""
  )
  
  dplyr::if_else(
    
    is.na(x),
    
    "NA",
    
    paste0(
      prefix,
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
  
  dplyr::case_when(
    
    is.na(p) ~
      "NA",
    
    p < 0.001 ~
      "<0.001",
    
    TRUE ~
      formatC(
        p,
        format = "f",
        digits = 3
      )
    
  )
  
}


# ==============================================================================
# 7. COUNTRY-NAME STANDARDISATION
#
# The analytical merge is primarily performed using the exact GBD location
# names. This function also harmonises common international naming variants
# for later linkage with the world map.
# ==============================================================================

normalise_country_name <- function(x) {
  
  original_name <- clean_text(x)
  
  dplyr::recode(
    
    original_name,
    
    "Bolivia" =
      "Bolivia (Plurinational State of)",
    
    "Bolivia (Plurinational State of)" =
      "Bolivia (Plurinational State of)",
    
    "Cape Verde" =
      "Cabo Verde",
    
    "Congo, Dem. Rep." =
      "Democratic Republic of the Congo",
    
    "Democratic Republic of Congo" =
      "Democratic Republic of the Congo",
    
    "Democratic Republic of the Congo" =
      "Democratic Republic of the Congo",
    
    "Congo, Democratic Republic of the" =
      "Democratic Republic of the Congo",
    
    "Congo, Rep." =
      "Congo",
    
    "Republic of the Congo" =
      "Congo",
    
    "Czech Republic" =
      "Czechia",
    
    "Cote d'Ivoire" =
      "Côte d'Ivoire",
    
    "Côte d’Ivoire" =
      "Côte d'Ivoire",
    
    "North Korea" =
      "Democratic People's Republic of Korea",
    
    "South Korea" =
      "Republic of Korea",
    
    "Korea, Republic of" =
      "Republic of Korea",
    
    "Iran" =
      "Iran (Islamic Republic of)",
    
    "Iran, Islamic Republic of" =
      "Iran (Islamic Republic of)",
    
    "Laos" =
      "Lao People's Democratic Republic",
    
    "Lao PDR" =
      "Lao People's Democratic Republic",
    
    "Moldova" =
      "Republic of Moldova",
    
    "Moldova, Republic of" =
      "Republic of Moldova",
    
    "Russia" =
      "Russian Federation",
    
    "Syria" =
      "Syrian Arab Republic",
    
    "Tanzania" =
      "United Republic of Tanzania",
    
    "Tanzania, United Republic of" =
      "United Republic of Tanzania",
    
    "Turkey" =
      "Türkiye",
    
    "Turkiye" =
      "Türkiye",
    
    "United States" =
      "United States of America",
    
    "USA" =
      "United States of America",
    
    "Venezuela" =
      "Venezuela (Bolivarian Republic of)",
    
    "Venezuela, Bolivarian Republic of" =
      "Venezuela (Bolivarian Republic of)",
    
    "Vietnam" =
      "Viet Nam",
    
    "West Bank and Gaza" =
      "Palestine",
    
    "State of Palestine" =
      "Palestine",
    
    "Micronesia (country)" =
      "Micronesia (Federated States of)",
    
    "Micronesia, Fed. Sts." =
      "Micronesia (Federated States of)",
    
    "Brunei" =
      "Brunei Darussalam",
    
    "Swaziland" =
      "Eswatini",
    
    .default =
      original_name
    
  )
  
}


# ==============================================================================
# 8. FILE IMPORT FUNCTION
# ==============================================================================

read_one_raw_file <- function(
    file_path
) {
  
  file_extension <- stringr::str_to_lower(
    tools::file_ext(
      file_path
    )
  )
  
  message("")
  message("Reading:")
  message(file_path)
  
  if (file_extension == "csv") {
    
    imported_data <- readr::read_csv(
      file = file_path,
      show_col_types = FALSE,
      progress = FALSE,
      guess_max = 500000
    )
    
  } else if (
    file_extension %in% c(
      "xlsx",
      "xls"
    )
  ) {
    
    workbook_sheets <- readxl::excel_sheets(
      file_path
    )
    
    preferred_sheet <- workbook_sheets[
      stringr::str_detect(
        stringr::str_to_lower(
          workbook_sheets
        ),
        "data|result"
      )
    ][1]
    
    if (is.na(preferred_sheet)) {
      
      preferred_sheet <- workbook_sheets[[1]]
      
    }
    
    imported_data <- readxl::read_excel(
      path = file_path,
      sheet = preferred_sheet
    )
    
  } else {
    
    stop(
      paste0(
        "\nUnsupported input file type:\n",
        file_path
      ),
      call. = FALSE
    )
    
  }
  
  imported_data |>
    janitor::clean_names()
  
}


# ==============================================================================
# 9. IMPORT ALL FILES FROM ONE MEASURE FOLDER
# ==============================================================================

read_measure_folder <- function(
    folder_path,
    expected_measure
) {
  
  source_files <- list.files(
    path = folder_path,
    pattern = "\\.(csv|xlsx|xls)$",
    full.names = TRUE,
    recursive = TRUE,
    ignore.case = TRUE
  )
  
  source_files <- source_files[
    !stringr::str_detect(
      basename(source_files),
      "^~\\$"
    )
  ]
  
  if (length(source_files) == 0) {
    
    stop(
      paste0(
        "\nNo CSV or Excel files were found in:\n",
        folder_path
      ),
      call. = FALSE
    )
    
  }
  
  message("")
  message("==============================================================")
  message(expected_measure)
  message("Files detected: ", length(source_files))
  message("==============================================================")
  
  imported_files <- vector(
    mode = "list",
    length = length(source_files)
  )
  
  for (file_index in seq_along(source_files)) {
    
    current_file <- source_files[[file_index]]
    
    current_data <- read_one_raw_file(
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
      paste0(
        "location in ",
        basename(current_file)
      )
    )
    
    location_id_column <- find_column(
      current_data,
      c(
        "location_id"
      ),
      "location ID",
      required = FALSE
    )
    
    year_column <- find_column(
      current_data,
      c(
        "year",
        "year_id"
      ),
      paste0(
        "year in ",
        basename(current_file)
      )
    )
    
    sex_column <- find_column(
      current_data,
      c(
        "sex",
        "sex_name"
      ),
      paste0(
        "sex in ",
        basename(current_file)
      )
    )
    
    age_column <- find_column(
      current_data,
      c(
        "age",
        "age_name"
      ),
      paste0(
        "age in ",
        basename(current_file)
      )
    )
    
    metric_column <- find_column(
      current_data,
      c(
        "metric",
        "metric_name"
      ),
      paste0(
        "metric in ",
        basename(current_file)
      )
    )
    
    value_column <- find_column(
      current_data,
      c(
        "val",
        "value",
        "mean",
        "estimate"
      ),
      paste0(
        "central estimate in ",
        basename(current_file)
      )
    )
    
    lower_column <- find_column(
      current_data,
      c(
        "lower",
        "lower_ci",
        "lower_ui",
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
        "upper_ui",
        "upper_bound",
        "upper_value"
      ),
      "upper uncertainty estimate",
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
    
    standardised_file <- current_data |>
      
      transmute(
        
        source_file =
          basename(
            current_file
          ),
        
        location =
          clean_text(
            .data[[location_column]]
          ),
        
        location_standardised =
          normalise_country_name(
            .data[[location_column]]
          ),
        
        location_id =
          if (!is.null(location_id_column)) {
            
            as.integer(
              .data[[location_id_column]]
            )
            
          } else {
            
            NA_integer_
            
          },
        
        year =
          as.integer(
            .data[[year_column]]
          ),
        
        sex =
          standardise_sex(
            .data[[sex_column]]
          ),
        
        age =
          clean_text(
            .data[[age_column]]
          ),
        
        metric_original =
          clean_text(
            .data[[metric_column]]
          ),
        
        cause =
          if (!is.null(cause_column)) {
            
            clean_text(
              .data[[cause_column]]
            )
            
          } else {
            
            "Colon and rectum cancer"
            
          },
        
        measure =
          if (!is.null(measure_column)) {
            
            standardise_measure(
              .data[[measure_column]]
            )
            
          } else {
            
            expected_measure
            
          },
        
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
          standardise_analysis_metric(
            metric =
              metric_original,
            age =
              age
          )
        
      )
    
    if (!is.null(cause_column)) {
      
      standardised_file <- standardised_file |>
        
        filter(
          is_colorectal_cancer(
            cause
          )
        )
      
    }
    
    standardised_file <- standardised_file |>
      
      filter(
        measure == expected_measure
      )
    
    imported_files[[file_index]] <-
      standardised_file
    
  }
  
  bind_rows(
    imported_files
  )
  
}


# ==============================================================================
# 10. IMPORT INCIDENCE, MORTALITY, AND DALY DATA
# ==============================================================================

incidence_raw <- read_measure_folder(
  folder_path =
    incidence_folder,
  expected_measure =
    "Incidence"
)

mortality_raw <- read_measure_folder(
  folder_path =
    mortality_folder,
  expected_measure =
    "Mortality"
)

dalys_raw <- read_measure_folder(
  folder_path =
    dalys_folder,
  expected_measure =
    "DALYs"
)


# ==============================================================================
# 11. INSPECT IMPORTED CATEGORIES
# ==============================================================================

message("")
message("==============================================================")
message("IMPORTED DATA SUMMARY")
message("==============================================================")

message("")
message("Incidence records:")
message(format(nrow(incidence_raw), big.mark = ","))

message("Mortality records:")
message(format(nrow(mortality_raw), big.mark = ","))

message("DALY records:")
message(format(nrow(dalys_raw), big.mark = ","))


all_raw_data <- bind_rows(
  incidence_raw,
  mortality_raw,
  dalys_raw
)

message("")
message("Measures detected:")

print(
  all_raw_data |>
    count(
      measure,
      sort = TRUE
    )
)

message("")
message("Sex categories detected:")

print(
  all_raw_data |>
    count(
      sex,
      sort = TRUE
    )
)

message("")
message("Age and metric categories detected:")

print(
  all_raw_data |>
    count(
      age,
      metric_original,
      analysis_metric,
      sort = TRUE
    ) |>
    slice_head(
      n = 30
    )
)

message("")
message("Years detected:")

print(
  all_raw_data |>
    summarise(
      minimum_year =
        min(
          year,
          na.rm = TRUE
        ),
      maximum_year =
        max(
          year,
          na.rm = TRUE
        )
    )
)


# ==============================================================================
# 12. REMOVE AGGREGATE LOCATIONS
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
  "North America",
  
  "Sub-Saharan Africa",
  "Latin America and Caribbean",
  "High-income",
  "Central Europe, Eastern Europe, and Central Asia",
  "Southeast Asia, East Asia, and Oceania",
  "South Asia",
  "North Africa and Middle East"
)


all_raw_data <- all_raw_data |>
  
  filter(
    !location %in%
      aggregate_location_names,
    !location_standardised %in%
      aggregate_location_names
  )


# ==============================================================================
# 13. RETAIN THE REQUIRED ANALYTICAL RECORDS
# ==============================================================================

analysis_records <- all_raw_data |>
  
  filter(
    
    year >=
      analysis_start_year,
    
    year <=
      analysis_end_year,
    
    sex ==
      analysis_sex,
    
    measure %in%
      required_measures,
    
    analysis_metric %in%
      required_analysis_metrics,
    
    !is.na(
      location_standardised
    ),
    
    location_standardised !=
      "",
    
    !is.na(
      value
    ),
    
    value >=
      0
    
  ) |>
  
  mutate(
    
    measure =
      factor(
        measure,
        levels =
          required_measures
      ),
    
    analysis_metric =
      factor(
        analysis_metric,
        levels =
          required_analysis_metrics
      )
    
  )


message("")
message("Records retained for analysis:")

print(
  analysis_records |>
    count(
      measure,
      analysis_metric,
      sort = TRUE
    )
)


# ==============================================================================
# 14. CHECK COUNTRY COVERAGE BY MEASURE AND METRIC
# ==============================================================================

coverage_summary <- analysis_records |>
  
  group_by(
    measure,
    analysis_metric
  ) |>
  
  summarise(
    
    countries =
      n_distinct(
        location_standardised
      ),
    
    years =
      n_distinct(
        year
      ),
    
    observations =
      n(),
    
    .groups =
      "drop"
    
  )

message("")
message("Country and year coverage:")

print(
  coverage_summary
)


incomplete_coverage <- coverage_summary |>
  
  filter(
    
    countries !=
      expected_country_count |
      
      years !=
      expected_year_count
    
  )


if (nrow(incomplete_coverage) > 0) {
  
  print(
    incomplete_coverage
  )
  
  stop(
    paste0(
      "\nCountry or year coverage is incomplete for at least one ",
      "measure-metric combination.\n",
      "Expected ",
      expected_country_count,
      " countries and ",
      expected_year_count,
      " years."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 15. CHECK REQUIRED MEASURE-METRIC COMBINATIONS
# ==============================================================================

required_combinations <- tidyr::expand_grid(
  
  measure =
    factor(
      required_measures,
      levels =
        required_measures
    ),
  
  analysis_metric =
    factor(
      required_analysis_metrics,
      levels =
        required_analysis_metrics
    )
  
)


available_combinations <- analysis_records |>
  
  distinct(
    measure,
    analysis_metric
  )


missing_combinations <- anti_join(
  
  required_combinations,
  
  available_combinations,
  
  by = c(
    "measure",
    "analysis_metric"
  )
  
)


if (nrow(missing_combinations) > 0) {
  
  print(
    missing_combinations
  )
  
  stop(
    "\nOne or more required measure-metric combinations are missing.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 16. IDENTIFY DUPLICATE COUNTRY-YEAR RECORDS
# ==============================================================================

duplicate_records <- analysis_records |>
  
  count(
    
    location_standardised,
    measure,
    analysis_metric,
    year,
    
    name =
      "number_of_records"
    
  ) |>
  
  filter(
    number_of_records >
      1
  )


if (nrow(duplicate_records) > 0) {
  
  message("")
  message("Duplicate country-year records detected:")
  
  print(
    duplicate_records
  )
  
  message("")
  message(
    "Duplicates will be resolved only when their central estimates are equal."
  )
  
}


duplicate_value_conflicts <- analysis_records |>
  
  group_by(
    
    location_standardised,
    measure,
    analysis_metric,
    year
    
  ) |>
  
  summarise(
    
    non_missing_values =
      n_distinct(
        value[
          !is.na(value)
        ]
      ),
    
    .groups =
      "drop"
    
  ) |>
  
  filter(
    non_missing_values >
      1
  )


if (nrow(duplicate_value_conflicts) > 0) {
  
  print(
    duplicate_value_conflicts
  )
  
  stop(
    paste0(
      "\nConflicting duplicate estimates were detected. ",
      "The script will not select one arbitrarily."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 17. COLLAPSE EXACT DUPLICATES
# ==============================================================================

analysis_records_unique <- analysis_records |>
  
  group_by(
    
    location_standardised,
    measure,
    analysis_metric,
    year
    
  ) |>
  
  summarise(
    
    location =
      first_non_missing_character(
        location
      ),
    
    location_id =
      first_non_missing_integer(
        location_id
      ),
    
    value =
      first_non_missing_numeric(
        value
      ),
    
    lower =
      first_non_missing_numeric(
        lower
      ),
    
    upper =
      first_non_missing_numeric(
        upper
      ),
    
    source_file =
      paste(
        sort(
          unique(
            source_file
          )
        ),
        collapse =
          "; "
      ),
    
    .groups =
      "drop"
    
  )


expected_records_after_collapse <-
  
  expected_country_count *
  expected_year_count *
  length(
    required_measures
  ) *
  length(
    required_analysis_metrics
  )


if (
  nrow(
    analysis_records_unique
  ) !=
  expected_records_after_collapse
) {
  
  stop(
    paste0(
      "\nExpected ",
      format(
        expected_records_after_collapse,
        big.mark = ","
      ),
      " unique analytical records, but ",
      format(
        nrow(
          analysis_records_unique
        ),
        big.mark = ","
      ),
      " were created."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 18. CHECK LOCATION SET CONSISTENCY
# ==============================================================================

location_sets <- analysis_records_unique |>
  
  distinct(
    location_standardised,
    measure,
    analysis_metric
  ) |>
  
  count(
    location_standardised,
    name =
      "available_combinations"
  )


incomplete_locations <- location_sets |>
  
  filter(
    
    available_combinations !=
      length(
        required_measures
      ) *
      length(
        required_analysis_metrics
      )
    
  )


if (nrow(incomplete_locations) > 0) {
  
  print(
    incomplete_locations
  )
  
  stop(
    "\nAt least one country is not represented in all six required datasets.",
    call. = FALSE
  )
  
}


country_list <- analysis_records_unique |>
  
  distinct(
    location_standardised
  ) |>
  
  arrange(
    location_standardised
  )


if (
  nrow(
    country_list
  ) !=
  expected_country_count
) {
  
  print(
    country_list,
    n = Inf
  )
  
  stop(
    paste0(
      "\nExpected ",
      expected_country_count,
      " countries and territories, but ",
      nrow(
        country_list
      ),
      " were retained."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 19. CREATE NUMBER DATASET
# ==============================================================================

country_numbers <- analysis_records_unique |>
  
  filter(
    analysis_metric ==
      "Number"
  ) |>
  
  select(
    
    location_standardised,
    location,
    location_id,
    year,
    measure,
    value
    
  ) |>
  
  mutate(
    measure =
      as.character(
        measure
      )
  ) |>
  
  pivot_wider(
    
    names_from =
      measure,
    
    values_from =
      value
    
  )


required_number_columns <- c(
  "Incidence",
  "Mortality",
  "DALYs"
)


missing_number_columns <- setdiff(
  required_number_columns,
  names(
    country_numbers
  )
)


if (
  length(
    missing_number_columns
  ) > 0
) {
  
  stop(
    paste0(
      "\nRequired number columns are missing:\n",
      paste(
        missing_number_columns,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
  
}


country_numbers <- country_numbers |>
  
  rename(
    
    incidence_number =
      Incidence,
    
    mortality_number =
      Mortality,
    
    daly_number =
      DALYs
    
  )


# ==============================================================================
# 20. CREATE AGE-STANDARDIZED RATE DATASET
# ==============================================================================

country_rates <- analysis_records_unique |>
  
  filter(
    analysis_metric ==
      "ASR"
  ) |>
  
  select(
    
    location_standardised,
    location,
    location_id,
    year,
    measure,
    value,
    lower,
    upper
    
  ) |>
  
  mutate(
    measure =
      as.character(
        measure
      )
  ) |>
  
  pivot_wider(
    
    names_from =
      measure,
    
    values_from = c(
      value,
      lower,
      upper
    ),
    
    names_glue =
      "{measure}_{.value}"
    
  )


required_rate_columns <- c(
  "Incidence_value",
  "Mortality_value",
  "DALYs_value"
)


missing_rate_columns <- setdiff(
  required_rate_columns,
  names(
    country_rates
  )
)


if (
  length(
    missing_rate_columns
  ) > 0
) {
  
  stop(
    paste0(
      "\nRequired ASR columns are missing:\n",
      paste(
        missing_rate_columns,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
  
}


country_rates <- country_rates |>
  
  rename(
    
    incidence_asr =
      Incidence_value,
    
    incidence_asr_lower =
      Incidence_lower,
    
    incidence_asr_upper =
      Incidence_upper,
    
    mortality_asr =
      Mortality_value,
    
    mortality_asr_lower =
      Mortality_lower,
    
    mortality_asr_upper =
      Mortality_upper,
    
    daly_asr =
      DALYs_value,
    
    daly_asr_lower =
      DALYs_lower,
    
    daly_asr_upper =
      DALYs_upper
    
  )


# ==============================================================================
# 21. MERGE NUMBER AND RATE DATASETS
# ==============================================================================

country_year_data <- country_numbers |>
  
  select(
    
    location_standardised,
    location_number =
      location,
    location_id_number =
      location_id,
    year,
    
    incidence_number,
    mortality_number,
    daly_number
    
  ) |>
  
  inner_join(
    
    country_rates |>
      
      select(
        
        location_standardised,
        location_rate =
          location,
        location_id_rate =
          location_id,
        year,
        
        incidence_asr,
        incidence_asr_lower,
        incidence_asr_upper,
        
        mortality_asr,
        mortality_asr_lower,
        mortality_asr_upper,
        
        daly_asr,
        daly_asr_lower,
        daly_asr_upper
        
      ),
    
    by = c(
      "location_standardised",
      "year"
    )
    
  ) |>
  
  mutate(
    
    location =
      dplyr::coalesce(
        location_number,
        location_rate,
        location_standardised
      ),
    
    location_id =
      dplyr::coalesce(
        location_id_number,
        location_id_rate
      )
    
  ) |>
  
  select(
    
    location,
    location_standardised,
    location_id,
    year,
    
    incidence_number,
    mortality_number,
    daly_number,
    
    incidence_asr,
    incidence_asr_lower,
    incidence_asr_upper,
    
    mortality_asr,
    mortality_asr_lower,
    mortality_asr_upper,
    
    daly_asr,
    daly_asr_lower,
    daly_asr_upper
    
  ) |>
  
  arrange(
    location_standardised,
    year
  )


# ==============================================================================
# 22. CALCULATE COUNTRY-YEAR SEVERITY INDICATORS
# ==============================================================================

country_year_data <- country_year_data |>
  
  mutate(
    
    dalys_per_case =
      dplyr::if_else(
        
        !is.na(
          incidence_number
        ) &
          incidence_number >
          0,
        
        daly_number /
          incidence_number,
        
        NA_real_
        
      ),
    
    mir =
      dplyr::if_else(
        
        !is.na(
          incidence_number
        ) &
          incidence_number >
          0,
        
        mortality_number /
          incidence_number,
        
        NA_real_
        
      )
    
  )


# ==============================================================================
# 23. VALIDATE THE COUNTRY-YEAR DATASET
# ==============================================================================

if (
  nrow(
    country_year_data
  ) !=
  expected_country_year_rows
) {
  
  stop(
    paste0(
      "\nExpected ",
      format(
        expected_country_year_rows,
        big.mark = ","
      ),
      " country-year observations, but ",
      format(
        nrow(
          country_year_data
        ),
        big.mark = ","
      ),
      " were created."
    ),
    call. = FALSE
  )
  
}


invalid_country_year_values <- country_year_data |>
  
  filter(
    
    is.na(
      incidence_number
    ) |
      incidence_number <=
      0 |
      
      is.na(
        mortality_number
      ) |
      mortality_number <
      0 |
      
      is.na(
        daly_number
      ) |
      daly_number <
      0 |
      
      is.na(
        incidence_asr
      ) |
      incidence_asr <
      0 |
      
      is.na(
        mortality_asr
      ) |
      mortality_asr <
      0 |
      
      is.na(
        daly_asr
      ) |
      daly_asr <
      0 |
      
      is.na(
        dalys_per_case
      ) |
      !is.finite(
        dalys_per_case
      ) |
      dalys_per_case <
      0 |
      
      is.na(
        mir
      ) |
      !is.finite(
        mir
      ) |
      mir <
      0
    
  )


if (
  nrow(
    invalid_country_year_values
  ) > 0
) {
  
  print(
    invalid_country_year_values
  )
  
  stop(
    "\nMissing, negative, or invalid country-year values were detected.",
    call. = FALSE
  )
  
}


mir_above_one <- country_year_data |>
  
  filter(
    mir >
      1
  )


if (
  nrow(
    mir_above_one
  ) > 0
) {
  
  warning(
    paste0(
      "\n",
      nrow(
        mir_above_one
      ),
      " country-year observations have MIR values above 1. ",
      "These values are retained because mortality and incidence are ",
      "independently modelled population estimates, but they should be ",
      "reviewed in the final outputs."
    )
  )
  
}


# ==============================================================================
# 24. CHECK COMPLETE YEAR COVERAGE WITHIN EACH COUNTRY
# ==============================================================================

country_year_coverage <- country_year_data |>
  
  group_by(
    location_standardised
  ) |>
  
  summarise(
    
    first_year =
      min(
        year
      ),
    
    last_year =
      max(
        year
      ),
    
    years_available =
      n_distinct(
        year
      ),
    
    .groups =
      "drop"
    
  )


incomplete_country_year_coverage <- country_year_coverage |>
  
  filter(
    
    first_year !=
      analysis_start_year |
      
      last_year !=
      analysis_end_year |
      
      years_available !=
      expected_year_count
    
  )


if (
  nrow(
    incomplete_country_year_coverage
  ) > 0
) {
  
  print(
    incomplete_country_year_coverage
  )
  
  stop(
    "\nAt least one country does not contain the complete 1990–2023 series.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 25. CREATE LONG ANALYTICAL DATASET FOR TREND MODELLING
# ==============================================================================

country_trend_long <- country_year_data |>
  
  select(
    
    location,
    location_standardised,
    location_id,
    year,
    
    incidence_asr,
    mortality_asr,
    daly_asr,
    dalys_per_case,
    mir
    
  ) |>
  
  pivot_longer(
    
    cols = c(
      incidence_asr,
      mortality_asr,
      daly_asr,
      dalys_per_case,
      mir
    ),
    
    names_to =
      "indicator_code",
    
    values_to =
      "value"
    
  ) |>
  
  mutate(
    
    indicator = case_when(
      
      indicator_code ==
        "incidence_asr" ~
        "Incidence ASR",
      
      indicator_code ==
        "mortality_asr" ~
        "Mortality ASR",
      
      indicator_code ==
        "daly_asr" ~
        "DALY ASR",
      
      indicator_code ==
        "dalys_per_case" ~
        "DALYs per case",
      
      indicator_code ==
        "mir" ~
        "Mortality-to-incidence ratio",
      
      TRUE ~
        indicator_code
      
    ),
    
    indicator = factor(
      
      indicator,
      
      levels = c(
        "Incidence ASR",
        "Mortality ASR",
        "DALY ASR",
        "DALYs per case",
        "Mortality-to-incidence ratio"
      )
      
    )
    
  ) |>
  
  arrange(
    location_standardised,
    indicator,
    year
  )


# ==============================================================================
# 26. FINAL PART 1 QUALITY CHECK
# ==============================================================================

expected_long_rows <-
  
  expected_country_year_rows *
  5


if (
  nrow(
    country_trend_long
  ) !=
  expected_long_rows
) {
  
  stop(
    paste0(
      "\nExpected ",
      format(
        expected_long_rows,
        big.mark = ","
      ),
      " rows in the long trend dataset, but ",
      format(
        nrow(
          country_trend_long
        ),
        big.mark = ","
      ),
      " were created."
    ),
    call. = FALSE
  )
  
}


message("")
message("==============================================================")
message("PART 1 COMPLETED")
message("==============================================================")

message("")
message(
  "Countries and territories: ",
  n_distinct(
    country_year_data$location_standardised
  )
)

message(
  "Years per country: ",
  expected_year_count
)

message(
  "Country-year observations: ",
  format(
    nrow(
      country_year_data
    ),
    big.mark = ","
  )
)

message(
  "Indicator observations for trend modelling: ",
  format(
    nrow(
      country_trend_long
    ),
    big.mark = ","
  )
)

message("")
message("Indicators prepared:")
message("Incidence ASR")
message("Mortality ASR")
message("DALY ASR")
message("DALYs per case")
message("Mortality-to-incidence ratio")

message("")
message("Do not run the script yet.")
message("Append Part 2 immediately below this line.")
message("==============================================================")
# ==============================================================================
# PART 2 OF 3
#
# COUNTRY-SPECIFIC TREND MODELS
# AAPC ESTIMATION
# TREND CLASSIFICATION
# START-END CHANGES
# COUNTRY RANKINGS
# ==============================================================================


# ==============================================================================
# 27. TREND-MODELLING HELPER FUNCTION
#
# AAPC is estimated from:
#
#   log(value) = intercept + beta × year
#
#   AAPC = 100 × [exp(beta) - 1]
#
# Each model uses all annual observations from 1990 through 2023.
# ==============================================================================

calculate_country_aapc <- function(data) {
  
  model_data <- data |>
    
    filter(
      !is.na(year),
      !is.na(value),
      is.finite(value),
      value > 0
    ) |>
    
    arrange(
      year
    )
  
  number_of_observations <- nrow(
    model_data
  )
  
  if (
    number_of_observations <
    expected_year_count
  ) {
    
    return(
      
      tibble(
        
        aapc =
          NA_real_,
        
        lower_95_ci =
          NA_real_,
        
        upper_95_ci =
          NA_real_,
        
        p_value =
          NA_real_,
        
        r_squared =
          NA_real_,
        
        observations =
          number_of_observations,
        
        model_status =
          "Incomplete annual series"
        
      )
      
    )
    
  }
  
  
  model <- tryCatch(
    
    lm(
      log(value) ~ year,
      data = model_data
    ),
    
    error = function(e) {
      
      NULL
      
    }
    
  )
  
  
  if (is.null(model)) {
    
    return(
      
      tibble(
        
        aapc =
          NA_real_,
        
        lower_95_ci =
          NA_real_,
        
        upper_95_ci =
          NA_real_,
        
        p_value =
          NA_real_,
        
        r_squared =
          NA_real_,
        
        observations =
          number_of_observations,
        
        model_status =
          "Model failed"
        
      )
      
    )
    
  }
  
  
  model_summary <- summary(
    model
  )
  
  beta <- unname(
    coef(
      model
    )[["year"]]
  )
  
  beta_confidence_interval <- tryCatch(
    
    confint(
      model,
      parm = "year",
      level = 0.95
    ),
    
    error = function(e) {
      
      matrix(
        c(
          NA_real_,
          NA_real_
        ),
        nrow = 1
      )
      
    }
    
  )
  
  
  aapc_estimate <-
    100 *
    (
      exp(
        beta
      ) -
        1
    )
  
  
  aapc_lower <-
    100 *
    (
      exp(
        beta_confidence_interval[1]
      ) -
        1
    )
  
  
  aapc_upper <-
    100 *
    (
      exp(
        beta_confidence_interval[2]
      ) -
        1
    )
  
  
  model_p_value <-
    model_summary$coefficients[
      "year",
      "Pr(>|t|)"
    ]
  
  
  tibble(
    
    aapc =
      aapc_estimate,
    
    lower_95_ci =
      aapc_lower,
    
    upper_95_ci =
      aapc_upper,
    
    p_value =
      model_p_value,
    
    r_squared =
      model_summary$r.squared,
    
    observations =
      number_of_observations,
    
    model_status =
      "Successful"
    
  )
  
}


# ==============================================================================
# 28. ESTIMATE COUNTRY-SPECIFIC AAPCs
# ==============================================================================

country_aapc_long <- country_trend_long |>
  
  group_by(
    
    location,
    location_standardised,
    location_id,
    indicator
    
  ) |>
  
  group_modify(
    
    ~ calculate_country_aapc(
      .x
    )
    
  ) |>
  
  ungroup() |>
  
  mutate(
    
    indicator =
      factor(
        
        indicator,
        
        levels = c(
          "Incidence ASR",
          "Mortality ASR",
          "DALY ASR",
          "DALYs per case",
          "Mortality-to-incidence ratio"
        )
        
      )
    
  ) |>
  
  arrange(
    indicator,
    location_standardised
  )


# ==============================================================================
# 29. VALIDATE AAPC MODEL COVERAGE
# ==============================================================================

expected_aapc_rows <-
  
  expected_country_count *
  5


if (
  nrow(
    country_aapc_long
  ) !=
  expected_aapc_rows
) {
  
  stop(
    paste0(
      "\nExpected ",
      expected_aapc_rows,
      " country-indicator AAPC models, but ",
      nrow(
        country_aapc_long
      ),
      " were generated."
    ),
    call. = FALSE
  )
  
}


failed_models <- country_aapc_long |>
  
  filter(
    model_status !=
      "Successful"
  )


if (
  nrow(
    failed_models
  ) > 0
) {
  
  print(
    failed_models,
    n = Inf
  )
  
  stop(
    "\nAt least one country-level AAPC model failed or used incomplete data.",
    call. = FALSE
  )
  
}


invalid_aapc_results <- country_aapc_long |>
  
  filter(
    
    is.na(
      aapc
    ) |
      
      is.na(
        lower_95_ci
      ) |
      
      is.na(
        upper_95_ci
      ) |
      
      is.na(
        p_value
      ) |
      
      is.na(
        r_squared
      ) |
      
      !is.finite(
        aapc
      ) |
      
      !is.finite(
        lower_95_ci
      ) |
      
      !is.finite(
        upper_95_ci
      )
    
  )


if (
  nrow(
    invalid_aapc_results
  ) > 0
) {
  
  print(
    invalid_aapc_results,
    n = Inf
  )
  
  stop(
    "\nInvalid AAPC estimates were detected.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 30. CLASSIFY COUNTRY TRENDS
#
# Increasing:
#   AAPC > 0 and p < 0.05
#
# Decreasing:
#   AAPC < 0 and p < 0.05
#
# Stable:
#   p >= 0.05
# ==============================================================================

country_aapc_long <- country_aapc_long |>
  
  mutate(
    
    trend_direction = case_when(
      
      p_value <
        0.05 &
        aapc >
        0 ~
        "Increasing",
      
      p_value <
        0.05 &
        aapc <
        0 ~
        "Decreasing",
      
      TRUE ~
        "Stable"
      
    ),
    
    trend_direction = factor(
      
      trend_direction,
      
      levels = c(
        "Increasing",
        "Stable",
        "Decreasing"
      )
      
    )
    
  )


# ==============================================================================
# 31. EXTRACT 1990 AND 2023 VALUES
# ==============================================================================

country_period_values <- country_trend_long |>
  
  filter(
    year %in%
      c(
        analysis_start_year,
        analysis_end_year
      )
  ) |>
  
  select(
    
    location,
    location_standardised,
    location_id,
    indicator,
    year,
    value
    
  ) |>
  
  pivot_wider(
    
    names_from =
      year,
    
    values_from =
      value,
    
    names_prefix =
      "year_"
    
  )


start_value_column <- paste0(
  "year_",
  analysis_start_year
)

end_value_column <- paste0(
  "year_",
  analysis_end_year
)


country_period_values <- country_period_values |>
  
  mutate(
    
    value_1990 =
      .data[[
        start_value_column
      ]],
    
    value_2023 =
      .data[[
        end_value_column
      ]],
    
    absolute_change =
      value_2023 -
      value_1990,
    
    percentage_change =
      100 *
      (
        value_2023 -
          value_1990
      ) /
      value_1990
    
  ) |>
  
  select(
    
    location,
    location_standardised,
    location_id,
    indicator,
    
    value_1990,
    value_2023,
    absolute_change,
    percentage_change
    
  )


# ==============================================================================
# 32. MERGE AAPC AND PERIOD-CHANGE RESULTS
# ==============================================================================

country_trend_results_long <- country_aapc_long |>
  
  left_join(
    
    country_period_values,
    
    by = c(
      "location",
      "location_standardised",
      "location_id",
      "indicator"
    )
    
  ) |>
  
  arrange(
    indicator,
    location_standardised
  )


if (
  nrow(
    country_trend_results_long
  ) !=
  expected_aapc_rows
) {
  
  stop(
    "\nAAPC and period-change results did not merge correctly.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 33. CHECK START-END VALUES
# ==============================================================================

invalid_period_values <- country_trend_results_long |>
  
  filter(
    
    is.na(
      value_1990
    ) |
      
      is.na(
        value_2023
      ) |
      
      is.na(
        percentage_change
      ) |
      
      !is.finite(
        percentage_change
      )
    
  )


if (
  nrow(
    invalid_period_values
  ) > 0
) {
  
  print(
    invalid_period_values,
    n = Inf
  )
  
  stop(
    "\nInvalid 1990–2023 country changes were detected.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 34. CREATE WIDE COUNTRY AAPC DATASET
#
# One row per country, containing all five indicators.
# ==============================================================================

country_aapc_wide <- country_trend_results_long |>
  
  mutate(
    
    indicator_code = case_when(
      
      indicator ==
        "Incidence ASR" ~
        "incidence_asr",
      
      indicator ==
        "Mortality ASR" ~
        "mortality_asr",
      
      indicator ==
        "DALY ASR" ~
        "daly_asr",
      
      indicator ==
        "DALYs per case" ~
        "dalys_per_case",
      
      indicator ==
        "Mortality-to-incidence ratio" ~
        "mir",
      
      TRUE ~
        NA_character_
      
    )
    
  ) |>
  
  select(
    
    location,
    location_standardised,
    location_id,
    indicator_code,
    
    aapc,
    lower_95_ci,
    upper_95_ci,
    p_value,
    r_squared,
    trend_direction,
    
    value_1990,
    value_2023,
    absolute_change,
    percentage_change
    
  ) |>
  
  pivot_wider(
    
    names_from =
      indicator_code,
    
    values_from = c(
      
      aapc,
      lower_95_ci,
      upper_95_ci,
      p_value,
      r_squared,
      trend_direction,
      
      value_1990,
      value_2023,
      absolute_change,
      percentage_change
      
    ),
    
    names_glue =
      "{indicator_code}_{.value}"
    
  ) |>
  
  arrange(
    location_standardised
  )


if (
  nrow(
    country_aapc_wide
  ) !=
  expected_country_count
) {
  
  stop(
    paste0(
      "\nExpected ",
      expected_country_count,
      " countries in the wide AAPC dataset, but ",
      nrow(
        country_aapc_wide
      ),
      " were created."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 35. CREATE TREND-DIRECTION SUMMARY
# ==============================================================================

trend_direction_summary <- country_trend_results_long |>
  
  count(
    
    indicator,
    trend_direction,
    
    name =
      "number_of_countries"
    
  ) |>
  
  group_by(
    indicator
  ) |>
  
  mutate(
    
    percentage_of_countries =
      100 *
      number_of_countries /
      sum(
        number_of_countries
      )
    
  ) |>
  
  ungroup() |>
  
  complete(
    
    indicator,
    
    trend_direction = factor(
      
      c(
        "Increasing",
        "Stable",
        "Decreasing"
      ),
      
      levels = c(
        "Increasing",
        "Stable",
        "Decreasing"
      )
      
    ),
    
    fill = list(
      
      number_of_countries =
        0,
      
      percentage_of_countries =
        0
      
    )
    
  ) |>
  
  arrange(
    indicator,
    trend_direction
  )


# ==============================================================================
# 36. VALIDATE TREND-DIRECTION COUNTS
# ==============================================================================

trend_direction_totals <- trend_direction_summary |>
  
  group_by(
    indicator
  ) |>
  
  summarise(
    
    total_countries =
      sum(
        number_of_countries
      ),
    
    .groups =
      "drop"
    
  )


invalid_trend_totals <- trend_direction_totals |>
  
  filter(
    total_countries !=
      expected_country_count
  )


if (
  nrow(
    invalid_trend_totals
  ) > 0
) {
  
  print(
    invalid_trend_totals
  )
  
  stop(
    "\nTrend-direction classifications do not total 204 countries.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 37. CREATE INDICATOR-SPECIFIC SUMMARY STATISTICS
# ==============================================================================

indicator_aapc_summary <- country_trend_results_long |>
  
  group_by(
    indicator
  ) |>
  
  summarise(
    
    countries =
      n(),
    
    median_aapc =
      median(
        aapc,
        na.rm = TRUE
      ),
    
    q1_aapc =
      as.numeric(
        quantile(
          aapc,
          probs = 0.25,
          na.rm = TRUE,
          names = FALSE
        )
      ),
    
    q3_aapc =
      as.numeric(
        quantile(
          aapc,
          probs = 0.75,
          na.rm = TRUE,
          names = FALSE
        )
      ),
    
    minimum_aapc =
      min(
        aapc,
        na.rm = TRUE
      ),
    
    maximum_aapc =
      max(
        aapc,
        na.rm = TRUE
      ),
    
    median_percentage_change =
      median(
        percentage_change,
        na.rm = TRUE
      ),
    
    q1_percentage_change =
      as.numeric(
        quantile(
          percentage_change,
          probs = 0.25,
          na.rm = TRUE,
          names = FALSE
        )
      ),
    
    q3_percentage_change =
      as.numeric(
        quantile(
          percentage_change,
          probs = 0.75,
          na.rm = TRUE,
          names = FALSE
        )
      ),
    
    .groups =
      "drop"
    
  ) |>
  
  arrange(
    indicator
  )


# ==============================================================================
# 38. CREATE EXTREME-COUNTRY FUNCTION
# ==============================================================================

create_extreme_country_table <- function(
    indicator_name,
    number_to_retain = 10
) {
  
  indicator_data <- country_trend_results_long |>
    
    filter(
      indicator ==
        indicator_name
    )
  
  
  largest_increases <- indicator_data |>
    
    arrange(
      desc(
        aapc
      )
    ) |>
    
    slice_head(
      n =
        number_to_retain
    ) |>
    
    mutate(
      
      Rank =
        row_number(),
      
      Direction =
        "Largest increases"
      
    )
  
  
  largest_decreases <- indicator_data |>
    
    arrange(
      aapc
    ) |>
    
    slice_head(
      n =
        number_to_retain
    ) |>
    
    mutate(
      
      Rank =
        row_number(),
      
      Direction =
        "Largest decreases"
      
    )
  
  
  bind_rows(
    
    largest_increases,
    
    largest_decreases
    
  ) |>
    
    transmute(
      
      Indicator =
        as.character(
          indicator
        ),
      
      Direction,
      
      Rank,
      
      Country =
        location_standardised,
      
      AAPC_percent =
        aapc,
      
      Lower_95_CI =
        lower_95_ci,
      
      Upper_95_CI =
        upper_95_ci,
      
      P_value =
        p_value,
      
      Trend_classification =
        as.character(
          trend_direction
        ),
      
      Value_1990 =
        value_1990,
      
      Value_2023 =
        value_2023,
      
      Percentage_change_1990_2023 =
        percentage_change,
      
      R_squared =
        r_squared
      
    )
  
}


# ==============================================================================
# 39. CREATE COUNTRY EXTREME TABLES
# ==============================================================================

extremes_incidence <- create_extreme_country_table(
  indicator_name =
    "Incidence ASR",
  number_to_retain =
    10
)

extremes_mortality <- create_extreme_country_table(
  indicator_name =
    "Mortality ASR",
  number_to_retain =
    10
)

extremes_daly_rate <- create_extreme_country_table(
  indicator_name =
    "DALY ASR",
  number_to_retain =
    10
)

extremes_dalys_per_case <- create_extreme_country_table(
  indicator_name =
    "DALYs per case",
  number_to_retain =
    10
)

extremes_mir <- create_extreme_country_table(
  indicator_name =
    "Mortality-to-incidence ratio",
  number_to_retain =
    10
)


all_country_extremes <- bind_rows(
  
  extremes_incidence,
  extremes_mortality,
  extremes_daly_rate,
  extremes_dalys_per_case,
  extremes_mir
  
)


# ==============================================================================
# 40. IDENTIFY TOP FIVE COUNTRIES FOR AUTOMATIC RESULTS TEXT
# ==============================================================================

extract_top_country_names <- function(
    indicator_name,
    direction = c(
      "increase",
      "decrease"
    ),
    number_to_retain = 5
) {
  
  direction <- match.arg(
    direction
  )
  
  selected_data <- country_trend_results_long |>
    
    filter(
      indicator ==
        indicator_name
    )
  
  
  if (
    direction ==
    "increase"
  ) {
    
    selected_data <- selected_data |>
      
      arrange(
        desc(
          aapc
        )
      )
    
  } else {
    
    selected_data <- selected_data |>
      
      arrange(
        aapc
      )
    
  }
  
  
  selected_data |>
    
    slice_head(
      n =
        number_to_retain
    ) |>
    
    pull(
      location_standardised
    )
  
}


top_incidence_increases <- extract_top_country_names(
  "Incidence ASR",
  "increase",
  5
)

top_incidence_decreases <- extract_top_country_names(
  "Incidence ASR",
  "decrease",
  5
)

top_mortality_increases <- extract_top_country_names(
  "Mortality ASR",
  "increase",
  5
)

top_mortality_decreases <- extract_top_country_names(
  "Mortality ASR",
  "decrease",
  5
)

top_daly_increases <- extract_top_country_names(
  "DALY ASR",
  "increase",
  5
)

top_daly_decreases <- extract_top_country_names(
  "DALY ASR",
  "decrease",
  5
)

top_severity_increases <- extract_top_country_names(
  "DALYs per case",
  "increase",
  5
)

top_severity_decreases <- extract_top_country_names(
  "DALYs per case",
  "decrease",
  5
)

top_mir_increases <- extract_top_country_names(
  "Mortality-to-incidence ratio",
  "increase",
  5
)

top_mir_decreases <- extract_top_country_names(
  "Mortality-to-incidence ratio",
  "decrease",
  5
)


# ==============================================================================
# 41. CREATE TREND CONCORDANCE INDICATORS
#
# This identifies countries in which population burden and case severity
# moved in the same or opposite directions.
# ==============================================================================

trend_concordance <- country_aapc_wide |>
  
  mutate(
    
    burden_severity_pattern = case_when(
      
      daly_asr_aapc <
        0 &
        dalys_per_case_aapc <
        0 ~
        "Burden and severity both decreased",
      
      daly_asr_aapc >
        0 &
        dalys_per_case_aapc >
        0 ~
        "Burden and severity both increased",
      
      daly_asr_aapc >
        0 &
        dalys_per_case_aapc <
        0 ~
        "Burden increased while severity decreased",
      
      daly_asr_aapc <
        0 &
        dalys_per_case_aapc >
        0 ~
        "Burden decreased while severity increased",
      
      TRUE ~
        "Other or near-zero pattern"
      
    ),
    
    burden_severity_significance_pattern = case_when(
      
      daly_asr_trend_direction ==
        "Decreasing" &
        dalys_per_case_trend_direction ==
        "Decreasing" ~
        "Significant improvement in both",
      
      daly_asr_trend_direction ==
        "Increasing" &
        dalys_per_case_trend_direction ==
        "Increasing" ~
        "Significant deterioration in both",
      
      daly_asr_trend_direction ==
        "Increasing" &
        dalys_per_case_trend_direction ==
        "Decreasing" ~
        "Burden worsened, severity improved",
      
      daly_asr_trend_direction ==
        "Decreasing" &
        dalys_per_case_trend_direction ==
        "Increasing" ~
        "Burden improved, severity worsened",
      
      TRUE ~
        "At least one stable trend"
      
    )
    
  )


# ==============================================================================
# 42. SUMMARISE TREND CONCORDANCE
# ==============================================================================

trend_concordance_summary <- trend_concordance |>
  
  count(
    
    burden_severity_significance_pattern,
    
    name =
      "number_of_countries"
    
  ) |>
  
  mutate(
    
    percentage_of_countries =
      100 *
      number_of_countries /
      sum(
        number_of_countries
      )
    
  ) |>
  
  arrange(
    desc(
      number_of_countries
    )
  )


if (
  sum(
    trend_concordance_summary$number_of_countries
  ) !=
  expected_country_count
) {
  
  stop(
    "\nBurden-severity trend concordance does not total 204 countries.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 43. CREATE CORRELATION BETWEEN BURDEN AND SEVERITY AAPCs
# ==============================================================================

burden_severity_aapc_correlation <- cor.test(
  
  trend_concordance$daly_asr_aapc,
  
  trend_concordance$dalys_per_case_aapc,
  
  method =
    "spearman",
  
  exact =
    FALSE
  
)


mir_severity_aapc_correlation <- cor.test(
  
  trend_concordance$mir_aapc,
  
  trend_concordance$dalys_per_case_aapc,
  
  method =
    "spearman",
  
  exact =
    FALSE
  
)


# ==============================================================================
# 44. CREATE CORRELATION SUMMARY TABLE
# ==============================================================================

trend_correlation_summary <- tibble(
  
  Comparison = c(
    
    "DALY ASR AAPC versus DALYs-per-case AAPC",
    
    "MIR AAPC versus DALYs-per-case AAPC"
    
  ),
  
  Spearman_rho = c(
    
    unname(
      burden_severity_aapc_correlation$estimate
    ),
    
    unname(
      mir_severity_aapc_correlation$estimate
    )
    
  ),
  
  P_value = c(
    
    burden_severity_aapc_correlation$p.value,
    
    mir_severity_aapc_correlation$p.value
    
  )
  
) |>
  
  mutate(
    
    Spearman_rho =
      round(
        Spearman_rho,
        3
      ),
    
    P_value_formatted =
      format_p_value(
        P_value
      )
    
  )


# ==============================================================================
# 45. PREPARE MAP-LABEL INDICATOR NAMES
# ==============================================================================

indicator_map_titles <- tibble(
  
  indicator = factor(
    
    c(
      "Incidence ASR",
      "Mortality ASR",
      "DALY ASR",
      "DALYs per case",
      "Mortality-to-incidence ratio"
    ),
    
    levels = c(
      "Incidence ASR",
      "Mortality ASR",
      "DALY ASR",
      "DALYs per case",
      "Mortality-to-incidence ratio"
    )
    
  ),
  
  panel_title = c(
    
    "A. Incidence ASR",
    
    "B. Mortality ASR",
    
    "C. DALY ASR",
    
    "D. DALYs per case",
    
    "E. Mortality-to-incidence ratio"
    
  )
  
)


country_trend_results_long <- country_trend_results_long |>
  
  left_join(
    
    indicator_map_titles,
    
    by =
      "indicator"
    
  )


if (
  any(
    is.na(
      country_trend_results_long$panel_title
    )
  )
) {
  
  stop(
    "\nAt least one indicator did not receive a map-panel title.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 46. DISPLAY PART 2 SUMMARY
# ==============================================================================

message("")
message("==============================================================")
message("PART 2 RESULTS SUMMARY")
message("==============================================================")

message("")
message("AAPC models completed:")
message(
  nrow(
    country_aapc_long
  )
)

message("")
message("Trend-direction summary:")

print(
  trend_direction_summary,
  n = Inf
)

message("")
message("Indicator-level AAPC summary:")

print(
  indicator_aapc_summary,
  n = Inf
)

message("")
message("Burden-severity trend concordance:")

print(
  trend_concordance_summary,
  n = Inf
)

message("")
message("Trend correlations:")

print(
  trend_correlation_summary,
  n = Inf
)


# ==============================================================================
# 47. FINAL PART 2 QUALITY CHECK
# ==============================================================================

part_2_required_objects <- c(
  
  "country_trend_results_long",
  "country_aapc_wide",
  "trend_direction_summary",
  "indicator_aapc_summary",
  "all_country_extremes",
  "trend_concordance",
  "trend_concordance_summary",
  "trend_correlation_summary"
  
)


missing_part_2_objects <- part_2_required_objects[
  !vapply(
    part_2_required_objects,
    exists,
    logical(
      1
    )
  )
]


if (
  length(
    missing_part_2_objects
  ) > 0
) {
  
  stop(
    paste0(
      "\nThe following Part 2 objects were not created:\n",
      paste(
        missing_part_2_objects,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
  
}


message("")
message("==============================================================")
message("PART 2 COMPLETED")
message("==============================================================")

message("")
message(
  "Countries analysed: ",
  expected_country_count
)

message(
  "Indicators modelled: 5"
)

message(
  "Country-indicator trend models: ",
  expected_aapc_rows
)

message("")
message("Do not run the script yet.")
message("Append Part 3 immediately below this line.")
message("==============================================================")
# ==============================================================================
# PART 3 OF 3
#
# PUBLICATION TABLE
# SUPPLEMENTARY TABLE
# COUNTRY AAPC MAPS
# AUTOMATIC RESULTS TEXT
# FINAL QUALITY CONTROL
# ==============================================================================


# ==============================================================================
# 48. DEFINE OUTPUT FILE PATHS
# ==============================================================================

figure_7_file <- file.path(
  figures_folder,
  "Figure_7_Country_AAPC_maps_1990_2023.tiff"
)

table_7_file <- file.path(
  tables_folder,
  "Table_7_Country_temporal_trend_extremes_1990_2023.xlsx"
)

supplementary_s6_excel_file <- file.path(
  supplementary_folder,
  "Table_S6_Complete_country_AAPC_1990_2023.xlsx"
)

supplementary_s6_csv_file <- file.path(
  supplementary_folder,
  "Table_S6_Complete_country_AAPC_1990_2023.csv"
)

results_text_file <- file.path(
  results_text_folder,
  "Results_3_7_Country_temporal_trends_1990_2023.txt"
)


# ==============================================================================
# 49. CREATE EXCEL FORMATTING STYLES
# ==============================================================================

title_style <- openxlsx::createStyle(
  fontSize = 12,
  textDecoration = "bold",
  halign = "left",
  valign = "center"
)

subtitle_style <- openxlsx::createStyle(
  fontSize = 10,
  textDecoration = "italic",
  halign = "left",
  valign = "center",
  wrapText = TRUE
)

header_style <- openxlsx::createStyle(
  fontSize = 10,
  textDecoration = "bold",
  halign = "center",
  valign = "center",
  wrapText = TRUE,
  border = "Bottom"
)

body_style <- openxlsx::createStyle(
  fontSize = 10,
  halign = "center",
  valign = "center",
  wrapText = TRUE
)

left_body_style <- openxlsx::createStyle(
  fontSize = 10,
  halign = "left",
  valign = "center",
  wrapText = TRUE
)

integer_style <- openxlsx::createStyle(
  numFmt = "#,##0"
)

two_decimal_style <- openxlsx::createStyle(
  numFmt = "0.00"
)

three_decimal_style <- openxlsx::createStyle(
  numFmt = "0.000"
)

four_decimal_style <- openxlsx::createStyle(
  numFmt = "0.0000"
)

percentage_style <- openxlsx::createStyle(
  numFmt = "0.00"
)


# ==============================================================================
# 50. CREATE FORMATTED EXTREME-COUNTRY TABLE
# ==============================================================================

table_7_formatted <- all_country_extremes |>
  
  mutate(
    
    `AAPC, % (95% CI)` =
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
      format_p_value(
        P_value
      ),
    
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
    
    `Change, 1990–2023` =
      format_percent(
        Percentage_change_1990_2023,
        digits = 1,
        include_plus = TRUE
      )
    
  ) |>
  
  select(
    
    Indicator,
    
    Direction,
    
    Rank,
    
    Country,
    
    `AAPC, % (95% CI)`,
    
    `P value`,
    
    Trend_classification,
    
    `1990 value`,
    
    `2023 value`,
    
    `Change, 1990–2023`
    
  )


# ==============================================================================
# 51. CREATE TABLE 7 WORKBOOK
# ==============================================================================

table_7_workbook <- openxlsx::createWorkbook()


# ------------------------------------------------------------------------------
# Sheet 1: Publication summary
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(
  table_7_workbook,
  "Publication summary"
)

openxlsx::writeData(
  table_7_workbook,
  sheet = "Publication summary",
  x = paste0(
    "Table 7. Countries with the largest changes in colorectal cancer ",
    "burden and severity, 1990–2023"
  ),
  startRow = 1,
  startCol = 1
)

openxlsx::writeData(
  table_7_workbook,
  sheet = "Publication summary",
  x = table_7_formatted,
  startRow = 3,
  startCol = 1
)

table_7_note <- paste0(
  
  "Note: AAPC was estimated independently for each country using ",
  
  "log-linear regression of annual estimates from 1990 through 2023. ",
  
  "Increasing and decreasing trends were defined by p<0.05 and the sign ",
  
  "of the AAPC; trends with p≥0.05 were classified as stable. ",
  
  "ASR denotes age-standardized rate per 100,000 population. ",
  
  "DALYs per case were calculated as DALY number divided by incidence ",
  
  "number. MIR was calculated as mortality number divided by incidence number."
  
)

openxlsx::writeData(
  table_7_workbook,
  sheet = "Publication summary",
  x = table_7_note,
  startRow = nrow(
    table_7_formatted
  ) + 6,
  startCol = 1,
  colNames = FALSE
)


# ------------------------------------------------------------------------------
# Sheet 2: Numeric extreme-country results
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(
  table_7_workbook,
  "Numeric results"
)

openxlsx::writeData(
  table_7_workbook,
  sheet = "Numeric results",
  x = all_country_extremes
)


# ------------------------------------------------------------------------------
# Sheet 3: Trend-direction summary
# ------------------------------------------------------------------------------

trend_direction_summary_export <- trend_direction_summary |>
  
  transmute(
    
    Indicator =
      as.character(
        indicator
      ),
    
    Trend =
      as.character(
        trend_direction
      ),
    
    Countries =
      number_of_countries,
    
    Percentage =
      percentage_of_countries
    
  )

openxlsx::addWorksheet(
  table_7_workbook,
  "Trend direction summary"
)

openxlsx::writeData(
  table_7_workbook,
  sheet = "Trend direction summary",
  x = trend_direction_summary_export
)


# ------------------------------------------------------------------------------
# Sheet 4: Indicator summary
# ------------------------------------------------------------------------------

indicator_aapc_summary_export <- indicator_aapc_summary |>
  
  transmute(
    
    Indicator =
      as.character(
        indicator
      ),
    
    Countries =
      countries,
    
    Median_AAPC =
      median_aapc,
    
    Q1_AAPC =
      q1_aapc,
    
    Q3_AAPC =
      q3_aapc,
    
    Minimum_AAPC =
      minimum_aapc,
    
    Maximum_AAPC =
      maximum_aapc,
    
    Median_percentage_change =
      median_percentage_change,
    
    Q1_percentage_change =
      q1_percentage_change,
    
    Q3_percentage_change =
      q3_percentage_change
    
  )

openxlsx::addWorksheet(
  table_7_workbook,
  "Indicator summary"
)

openxlsx::writeData(
  table_7_workbook,
  sheet = "Indicator summary",
  x = indicator_aapc_summary_export
)


# ------------------------------------------------------------------------------
# Sheet 5: Burden-severity concordance
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(
  table_7_workbook,
  "Trend concordance"
)

openxlsx::writeData(
  table_7_workbook,
  sheet = "Trend concordance",
  x = trend_concordance_summary
)


# ------------------------------------------------------------------------------
# Sheet 6: Trend correlations
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(
  table_7_workbook,
  "Trend correlations"
)

openxlsx::writeData(
  table_7_workbook,
  sheet = "Trend correlations",
  x = trend_correlation_summary
)


# ==============================================================================
# 52. FORMAT TABLE 7 WORKBOOK
# ==============================================================================

openxlsx::addStyle(
  table_7_workbook,
  sheet = "Publication summary",
  style = title_style,
  rows = 1,
  cols = 1:ncol(
    table_7_formatted
  ),
  gridExpand = TRUE
)

openxlsx::addStyle(
  table_7_workbook,
  sheet = "Publication summary",
  style = header_style,
  rows = 3,
  cols = 1:ncol(
    table_7_formatted
  ),
  gridExpand = TRUE
)

openxlsx::addStyle(
  table_7_workbook,
  sheet = "Publication summary",
  style = body_style,
  rows = 4:(
    nrow(
      table_7_formatted
    ) + 3
  ),
  cols = 1:ncol(
    table_7_formatted
  ),
  gridExpand = TRUE
)

openxlsx::addStyle(
  table_7_workbook,
  sheet = "Publication summary",
  style = left_body_style,
  rows = 4:(
    nrow(
      table_7_formatted
    ) + 3
  ),
  cols = c(
    1,
    2,
    4
  ),
  gridExpand = TRUE
)

openxlsx::setColWidths(
  table_7_workbook,
  sheet = "Publication summary",
  cols = 1:ncol(
    table_7_formatted
  ),
  widths = c(
    31,
    20,
    8,
    27,
    24,
    12,
    19,
    14,
    14,
    20
  )
)

openxlsx::setRowHeights(
  table_7_workbook,
  sheet = "Publication summary",
  rows = 3,
  heights = 48
)

openxlsx::freezePane(
  table_7_workbook,
  sheet = "Publication summary",
  firstActiveRow = 4,
  firstActiveCol = 5
)

openxlsx::addFilter(
  table_7_workbook,
  sheet = "Publication summary",
  rows = 3,
  cols = 1:ncol(
    table_7_formatted
  )
)


table_7_sheet_data <- list(
  
  "Numeric results" =
    all_country_extremes,
  
  "Trend direction summary" =
    trend_direction_summary_export,
  
  "Indicator summary" =
    indicator_aapc_summary_export,
  
  "Trend concordance" =
    trend_concordance_summary,
  
  "Trend correlations" =
    trend_correlation_summary
  
)


for (
  current_sheet in names(
    table_7_sheet_data
  )
) {
  
  current_sheet_data <-
    table_7_sheet_data[[
      current_sheet
    ]]
  
  openxlsx::addStyle(
    table_7_workbook,
    sheet = current_sheet,
    style = header_style,
    rows = 1,
    cols = 1:ncol(
      current_sheet_data
    ),
    gridExpand = TRUE
  )
  
  if (
    nrow(
      current_sheet_data
    ) > 0
  ) {
    
    openxlsx::addStyle(
      table_7_workbook,
      sheet = current_sheet,
      style = body_style,
      rows = 2:(
        nrow(
          current_sheet_data
        ) + 1
      ),
      cols = 1:ncol(
        current_sheet_data
      ),
      gridExpand = TRUE
    )
    
  }
  
  openxlsx::setColWidths(
    table_7_workbook,
    sheet = current_sheet,
    cols = 1:ncol(
      current_sheet_data
    ),
    widths = "auto"
  )
  
  openxlsx::freezePane(
    table_7_workbook,
    sheet = current_sheet,
    firstRow = TRUE
  )
  
  openxlsx::addFilter(
    table_7_workbook,
    sheet = current_sheet,
    rows = 1,
    cols = 1:ncol(
      current_sheet_data
    )
  )
  
}


openxlsx::saveWorkbook(
  table_7_workbook,
  table_7_file,
  overwrite = TRUE
)


# ==============================================================================
# 53. PREPARE COMPLETE SUPPLEMENTARY TABLE S6
# ==============================================================================

supplementary_s6 <- country_trend_results_long |>
  
  transmute(
    
    Country =
      location_standardised,
    
    Location_ID =
      location_id,
    
    Indicator =
      as.character(
        indicator
      ),
    
    Value_1990 =
      value_1990,
    
    Value_2023 =
      value_2023,
    
    Absolute_change =
      absolute_change,
    
    Percentage_change_1990_2023 =
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
    
    Trend_classification =
      as.character(
        trend_direction
      ),
    
    Observations =
      observations,
    
    Model_status =
      model_status
    
  ) |>
  
  arrange(
    Country,
    factor(
      Indicator,
      levels = c(
        "Incidence ASR",
        "Mortality ASR",
        "DALY ASR",
        "DALYs per case",
        "Mortality-to-incidence ratio"
      )
    )
  )


# ==============================================================================
# 54. CREATE WIDE SUPPLEMENTARY COUNTRY TABLE
# ==============================================================================

supplementary_s6_wide <- country_aapc_wide |>
  
  left_join(
    
    trend_concordance |>
      
      select(
        location_standardised,
        burden_severity_pattern,
        burden_severity_significance_pattern
      ),
    
    by =
      "location_standardised"
    
  ) |>
  
  transmute(
    
    Country =
      location_standardised,
    
    Location_ID =
      location_id,
    
    Incidence_ASR_1990 =
      incidence_asr_value_1990,
    
    Incidence_ASR_2023 =
      incidence_asr_value_2023,
    
    Incidence_ASR_percentage_change =
      incidence_asr_percentage_change,
    
    Incidence_ASR_AAPC =
      incidence_asr_aapc,
    
    Incidence_ASR_lower_95_CI =
      incidence_asr_lower_95_ci,
    
    Incidence_ASR_upper_95_CI =
      incidence_asr_upper_95_ci,
    
    Incidence_ASR_P_value =
      incidence_asr_p_value,
    
    Incidence_ASR_trend =
      as.character(
        incidence_asr_trend_direction
      ),
    
    Mortality_ASR_1990 =
      mortality_asr_value_1990,
    
    Mortality_ASR_2023 =
      mortality_asr_value_2023,
    
    Mortality_ASR_percentage_change =
      mortality_asr_percentage_change,
    
    Mortality_ASR_AAPC =
      mortality_asr_aapc,
    
    Mortality_ASR_lower_95_CI =
      mortality_asr_lower_95_ci,
    
    Mortality_ASR_upper_95_CI =
      mortality_asr_upper_95_ci,
    
    Mortality_ASR_P_value =
      mortality_asr_p_value,
    
    Mortality_ASR_trend =
      as.character(
        mortality_asr_trend_direction
      ),
    
    DALY_ASR_1990 =
      daly_asr_value_1990,
    
    DALY_ASR_2023 =
      daly_asr_value_2023,
    
    DALY_ASR_percentage_change =
      daly_asr_percentage_change,
    
    DALY_ASR_AAPC =
      daly_asr_aapc,
    
    DALY_ASR_lower_95_CI =
      daly_asr_lower_95_ci,
    
    DALY_ASR_upper_95_CI =
      daly_asr_upper_95_ci,
    
    DALY_ASR_P_value =
      daly_asr_p_value,
    
    DALY_ASR_trend =
      as.character(
        daly_asr_trend_direction
      ),
    
    DALYs_per_case_1990 =
      dalys_per_case_value_1990,
    
    DALYs_per_case_2023 =
      dalys_per_case_value_2023,
    
    DALYs_per_case_percentage_change =
      dalys_per_case_percentage_change,
    
    DALYs_per_case_AAPC =
      dalys_per_case_aapc,
    
    DALYs_per_case_lower_95_CI =
      dalys_per_case_lower_95_ci,
    
    DALYs_per_case_upper_95_CI =
      dalys_per_case_upper_95_ci,
    
    DALYs_per_case_P_value =
      dalys_per_case_p_value,
    
    DALYs_per_case_trend =
      as.character(
        dalys_per_case_trend_direction
      ),
    
    MIR_1990 =
      mir_value_1990,
    
    MIR_2023 =
      mir_value_2023,
    
    MIR_percentage_change =
      mir_percentage_change,
    
    MIR_AAPC =
      mir_aapc,
    
    MIR_lower_95_CI =
      mir_lower_95_ci,
    
    MIR_upper_95_CI =
      mir_upper_95_ci,
    
    MIR_P_value =
      mir_p_value,
    
    MIR_trend =
      as.character(
        mir_trend_direction
      ),
    
    Burden_severity_directional_pattern =
      burden_severity_pattern,
    
    Burden_severity_significance_pattern =
      burden_severity_significance_pattern
    
  ) |>
  
  arrange(
    Country
  )


if (
  nrow(
    supplementary_s6_wide
  ) !=
  expected_country_count
) {
  
  stop(
    paste0(
      "\nExpected ",
      expected_country_count,
      " countries in Supplementary Table S6, but ",
      nrow(
        supplementary_s6_wide
      ),
      " were generated."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 55. CREATE SUPPLEMENTARY TABLE S6 WORKBOOK
# ==============================================================================

supplementary_s6_workbook <- openxlsx::createWorkbook()


# ------------------------------------------------------------------------------
# Sheet 1: One row per country
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(
  supplementary_s6_workbook,
  "Complete country dataset"
)

openxlsx::writeData(
  supplementary_s6_workbook,
  sheet = "Complete country dataset",
  x = supplementary_s6_wide
)


# ------------------------------------------------------------------------------
# Sheet 2: Long country-indicator dataset
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(
  supplementary_s6_workbook,
  "Country indicator results"
)

openxlsx::writeData(
  supplementary_s6_workbook,
  sheet = "Country indicator results",
  x = supplementary_s6
)


# ------------------------------------------------------------------------------
# Sheet 3: Direction summary
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(
  supplementary_s6_workbook,
  "Trend direction summary"
)

openxlsx::writeData(
  supplementary_s6_workbook,
  sheet = "Trend direction summary",
  x = trend_direction_summary_export
)


# ------------------------------------------------------------------------------
# Sheet 4: Indicator summary
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(
  supplementary_s6_workbook,
  "Indicator summary"
)

openxlsx::writeData(
  supplementary_s6_workbook,
  sheet = "Indicator summary",
  x = indicator_aapc_summary_export
)


# ------------------------------------------------------------------------------
# Sheet 5: Trend concordance
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(
  supplementary_s6_workbook,
  "Burden severity concordance"
)

openxlsx::writeData(
  supplementary_s6_workbook,
  sheet = "Burden severity concordance",
  x = trend_concordance_summary
)


# ------------------------------------------------------------------------------
# Sheet 6: Trend correlations
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(
  supplementary_s6_workbook,
  "Trend correlations"
)

openxlsx::writeData(
  supplementary_s6_workbook,
  sheet = "Trend correlations",
  x = trend_correlation_summary
)


# ------------------------------------------------------------------------------
# Sheet 7: Analytical assumptions
# ------------------------------------------------------------------------------

analysis_assumptions <- tibble(
  
  Item = c(
    
    "Study period",
    
    "Countries and territories",
    
    "Sex",
    
    "Indicators",
    
    "AAPC method",
    
    "Increasing trend",
    
    "Decreasing trend",
    
    "Stable trend",
    
    "DALYs per case formula",
    
    "MIR formula",
    
    "YLL data used",
    
    "YLD data used",
    
    "Global data used",
    
    "Clean_Data folder used"
    
  ),
  
  Value = c(
    
    paste0(
      analysis_start_year,
      "–",
      analysis_end_year
    ),
    
    expected_country_count,
    
    analysis_sex,
    
    paste(
      c(
        "Incidence ASR",
        "Mortality ASR",
        "DALY ASR",
        "DALYs per case",
        "Mortality-to-incidence ratio"
      ),
      collapse = "; "
    ),
    
    paste0(
      "Log-linear regression of annual values; ",
      "AAPC = 100 × [exp(beta) − 1]"
    ),
    
    "AAPC > 0 and p<0.05",
    
    "AAPC < 0 and p<0.05",
    
    "p≥0.05",
    
    "DALY number / incidence number",
    
    "Mortality number / incidence number",
    
    "No",
    
    "No",
    
    "No",
    
    "No"
    
  )
  
)

openxlsx::addWorksheet(
  supplementary_s6_workbook,
  "Analytical assumptions"
)

openxlsx::writeData(
  supplementary_s6_workbook,
  sheet = "Analytical assumptions",
  x = analysis_assumptions
)


# ==============================================================================
# 56. FORMAT SUPPLEMENTARY TABLE S6 WORKBOOK
# ==============================================================================

supplementary_sheet_data <- list(
  
  "Complete country dataset" =
    supplementary_s6_wide,
  
  "Country indicator results" =
    supplementary_s6,
  
  "Trend direction summary" =
    trend_direction_summary_export,
  
  "Indicator summary" =
    indicator_aapc_summary_export,
  
  "Burden severity concordance" =
    trend_concordance_summary,
  
  "Trend correlations" =
    trend_correlation_summary,
  
  "Analytical assumptions" =
    analysis_assumptions
  
)


for (
  current_sheet in names(
    supplementary_sheet_data
  )
) {
  
  current_sheet_data <-
    supplementary_sheet_data[[
      current_sheet
    ]]
  
  openxlsx::addStyle(
    supplementary_s6_workbook,
    sheet = current_sheet,
    style = header_style,
    rows = 1,
    cols = 1:ncol(
      current_sheet_data
    ),
    gridExpand = TRUE
  )
  
  if (
    nrow(
      current_sheet_data
    ) > 0
  ) {
    
    openxlsx::addStyle(
      supplementary_s6_workbook,
      sheet = current_sheet,
      style = body_style,
      rows = 2:(
        nrow(
          current_sheet_data
        ) + 1
      ),
      cols = 1:ncol(
        current_sheet_data
      ),
      gridExpand = TRUE
    )
    
  }
  
  openxlsx::setColWidths(
    supplementary_s6_workbook,
    sheet = current_sheet,
    cols = 1:ncol(
      current_sheet_data
    ),
    widths = "auto"
  )
  
  openxlsx::freezePane(
    supplementary_s6_workbook,
    sheet = current_sheet,
    firstRow = TRUE,
    firstCol = TRUE
  )
  
  openxlsx::addFilter(
    supplementary_s6_workbook,
    sheet = current_sheet,
    rows = 1,
    cols = 1:ncol(
      current_sheet_data
    )
  )
  
}


openxlsx::saveWorkbook(
  supplementary_s6_workbook,
  supplementary_s6_excel_file,
  overwrite = TRUE
)

readr::write_csv(
  supplementary_s6_wide,
  supplementary_s6_csv_file
)


# ==============================================================================
# 57. LOAD WORLD MAP
# ==============================================================================

world_map <- rnaturalearth::ne_countries(
  scale = "medium",
  returnclass = "sf"
)

if (
  !inherits(
    world_map,
    "sf"
  )
) {
  
  stop(
    "\nThe Natural Earth world map could not be loaded as an sf object.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 58. STANDARDISE WORLD-MAP COUNTRY NAMES
# ==============================================================================

world_map <- world_map |>
  
  mutate(
    
    map_country_original =
      dplyr::coalesce(
        admin,
        name_long,
        sovereignt
      ),
    
    map_country_standardised =
      normalise_country_name(
        map_country_original
      )
    
  )


# ==============================================================================
# 59. APPLY ADDITIONAL MAP-NAME CORRECTIONS
# ==============================================================================

world_map <- world_map |>
  
  mutate(
    
    map_country_standardised = dplyr::recode(
      
      map_country_standardised,
      
      "The Bahamas" =
        "Bahamas",
      
      "Gambia" =
        "The Gambia",
      
      "Dem. Rep. Congo" =
        "Democratic Republic of the Congo",
      
      "Republic of Congo" =
        "Congo",
      
      "South Korea" =
        "Republic of Korea",
      
      "North Korea" =
        "Democratic People's Republic of Korea",
      
      "United Republic of Tanzania" =
        "United Republic of Tanzania",
      
      "United States of America" =
        "United States of America",
      
      "Czech Republic" =
        "Czechia",
      
      "Western Sahara" =
        "Western Sahara",
      
      "Somaliland" =
        "Somalia",
      
      "Kosovo" =
        "Kosovo",
      
      "Taiwan" =
        "Taiwan (Province of China)",
      
      "Falkland Islands" =
        "Falkland Islands (Malvinas)",
      
      .default =
        map_country_standardised
      
    )
    
  )


# ==============================================================================
# 60. PREPARE MAP DATA
# ==============================================================================

map_trend_data <- country_trend_results_long |>
  
  select(
    
    location_standardised,
    
    indicator,
    
    panel_title,
    
    aapc,
    
    lower_95_ci,
    
    upper_95_ci,
    
    p_value,
    
    trend_direction
    
  )


map_data <- world_map |>
  
  left_join(
    
    map_trend_data,
    
    by = c(
      "map_country_standardised" =
        "location_standardised"
    )
    
  )


# ==============================================================================
# 61. CHECK MAP LINKAGE
# ==============================================================================

mapped_country_names <- world_map |>
  
  st_drop_geometry() |>
  
  distinct(
    map_country_standardised
  )


countries_not_mapped <- country_list |>
  
  anti_join(
    
    mapped_country_names,
    
    by = c(
      "location_standardised" =
        "map_country_standardised"
    )
    
  ) |>
  
  arrange(
    location_standardised
  )


message("")
message("Countries not represented in the Natural Earth map:")

if (
  nrow(
    countries_not_mapped
  ) == 0
) {
  
  message("None")
  
} else {
  
  print(
    countries_not_mapped,
    n = Inf
  )
  
}


# Small territories may not have polygons in Natural Earth.
# They remain in all numerical analyses and Supplementary Table S6.

mapped_analysis_countries <- country_list |>
  
  semi_join(
    
    mapped_country_names,
    
    by = c(
      "location_standardised" =
        "map_country_standardised"
    )
    
  )


message("")
message(
  "Analytical countries represented on map: ",
  nrow(
    mapped_analysis_countries
  ),
  " of ",
  expected_country_count
)


# ==============================================================================
# 62. DEFINE COMMON SYMMETRIC AAPC SCALE
#
# A single scale is not used across all indicators because their AAPC
# distributions differ. Each panel uses a symmetric indicator-specific
# range based on the 2nd and 98th percentiles, reducing distortion from
# isolated extreme estimates while retaining all countries.
# ==============================================================================

map_scale_limits <- country_trend_results_long |>
  
  group_by(
    indicator
  ) |>
  
  summarise(
    
    lower_quantile =
      as.numeric(
        quantile(
          aapc,
          probs = 0.02,
          na.rm = TRUE,
          names = FALSE
        )
      ),
    
    upper_quantile =
      as.numeric(
        quantile(
          aapc,
          probs = 0.98,
          na.rm = TRUE,
          names = FALSE
        )
      ),
    
    symmetric_limit =
      max(
        abs(
          c(
            lower_quantile,
            upper_quantile
          )
        ),
        na.rm = TRUE
      ),
    
    .groups =
      "drop"
    
  )


if (
  any(
    is.na(
      map_scale_limits$symmetric_limit
    )
  ) |
  any(
    map_scale_limits$symmetric_limit <= 0
  )
) {
  
  stop(
    "\nInvalid map scale limits were calculated.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 63. CREATE AAPC MAP FUNCTION
# ==============================================================================

create_aapc_map <- function(
    indicator_name
) {
  
  current_limit <- map_scale_limits |>
    
    filter(
      indicator ==
        indicator_name
    ) |>
    
    pull(
      symmetric_limit
    )
  
  
  if (
    length(
      current_limit
    ) != 1
  ) {
    
    stop(
      paste0(
        "\nA unique map scale could not be identified for ",
        indicator_name,
        "."
      ),
      call. = FALSE
    )
    
  }
  
  
  current_title <- indicator_map_titles |>
    
    filter(
      indicator ==
        indicator_name
    ) |>
    
    pull(
      panel_title
    )
  
  
  current_map_data <- map_data |>
    
    filter(
      is.na(
        indicator
      ) |
        indicator ==
        indicator_name
    )
  
  
  ggplot(
    current_map_data
  ) +
    
    geom_sf(
      aes(
        fill = aapc
      ),
      color = "grey65",
      linewidth = 0.08
    ) +
    
    scale_fill_gradient2(
      
      low =
        "#2166AC",
      
      mid =
        "white",
      
      high =
        "#B2182B",
      
      midpoint =
        0,
      
      limits = c(
        -current_limit,
        current_limit
      ),
      
      oob =
        scales::squish,
      
      na.value =
        "grey90",
      
      name =
        "AAPC (%)",
      
      labels =
        scales::label_number(
          accuracy = 0.1
        )
      
    ) +
    
    coord_sf(
      xlim = c(
        -180,
        180
      ),
      ylim = c(
        -60,
        90
      ),
      expand = FALSE
    ) +
    
    labs(
      title =
        current_title
    ) +
    
    theme_void(
      base_size = 11
    ) +
    
    theme(
      
      plot.title =
        element_text(
          face = "bold",
          size = 12,
          hjust = 0
        ),
      
      legend.position =
        "bottom",
      
      legend.title =
        element_text(
          size = 9.5,
          face = "bold"
        ),
      
      legend.text =
        element_text(
          size = 8.5
        ),
      
      legend.key.width =
        grid::unit(
          1.35,
          "cm"
        ),
      
      plot.margin =
        margin(
          5,
          5,
          5,
          5
        )
      
    )
  
}


# ==============================================================================
# 64. CREATE FIGURE 7 PANELS
# ==============================================================================

figure_7a <- create_aapc_map(
  "Incidence ASR"
)

figure_7b <- create_aapc_map(
  "Mortality ASR"
)

figure_7c <- create_aapc_map(
  "DALY ASR"
)

figure_7d <- create_aapc_map(
  "DALYs per case"
)

figure_7e <- create_aapc_map(
  "Mortality-to-incidence ratio"
)


# ==============================================================================
# 65. COMBINE FIGURE 7
# ==============================================================================

figure_7 <- (
  
  figure_7a +
    figure_7b
  
) /
  
  (
    
    figure_7c +
      figure_7d
    
  ) /
  
  figure_7e +
  
  patchwork::plot_layout(
    
    guides =
      "keep",
    
    heights = c(
      1,
      1,
      0.95
    )
    
  ) +
  
  patchwork::plot_annotation(
    
    title = paste0(
      "Country-level average annual changes in colorectal cancer ",
      "burden and severity, 1990–2023"
    ),
    
    caption = paste0(
      
      "Blue indicates a declining trend and red indicates an increasing trend. ",
      
      "AAPCs were estimated using log-linear regression of annual country-level ",
      
      "estimates. Indicator-specific symmetric colour limits were based on the ",
      
      "2nd and 98th percentiles and extreme values were squished to the scale limits. ",
      
      "Grey areas indicate locations without a corresponding polygon or estimate."
      
    ),
    
    theme = theme(
      
      plot.title =
        element_text(
          face = "bold",
          size = 16,
          hjust = 0.5
        ),
      
      plot.caption =
        element_text(
          size = 8.5,
          hjust = 0
        ),
      
      plot.margin =
        margin(
          12,
          14,
          12,
          14
        )
      
    )
    
  )


# ==============================================================================
# 66. SAVE FIGURE 7 AS TIFF AT 600 DPI
# ==============================================================================

ggsave(
  filename = figure_7_file,
  plot = figure_7,
  device = "tiff",
  width = 14,
  height = 14,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white"
)


# ==============================================================================
# 67. HELPER FUNCTIONS FOR RESULTS TEXT
# ==============================================================================

get_trend_count <- function(
    indicator_name,
    direction_name
) {
  
  selected_row <- trend_direction_summary |>
    
    filter(
      indicator ==
        indicator_name,
      trend_direction ==
        direction_name
    )
  
  
  if (
    nrow(
      selected_row
    ) == 0
  ) {
    
    return(
      tibble(
        number_of_countries = 0,
        percentage_of_countries = 0
      )
    )
    
  }
  
  
  selected_row |>
    
    select(
      number_of_countries,
      percentage_of_countries
    )
  
}


get_indicator_summary <- function(
    indicator_name
) {
  
  selected_row <- indicator_aapc_summary |>
    
    filter(
      indicator ==
        indicator_name
    )
  
  
  if (
    nrow(
      selected_row
    ) != 1
  ) {
    
    stop(
      paste0(
        "\nA unique summary was not available for ",
        indicator_name,
        "."
      ),
      call. = FALSE
    )
    
  }
  
  selected_row
  
}


format_country_list <- function(
    countries
) {
  
  countries <- as.character(
    countries
  )
  
  countries <- countries[
    !is.na(
      countries
    ) &
      countries != ""
  ]
  
  
  if (
    length(
      countries
    ) == 0
  ) {
    
    return(
      "none"
    )
    
  }
  
  
  if (
    length(
      countries
    ) == 1
  ) {
    
    return(
      countries
    )
    
  }
  
  
  if (
    length(
      countries
    ) == 2
  ) {
    
    return(
      paste(
        countries,
        collapse = " and "
      )
    )
    
  }
  
  
  paste0(
    
    paste(
      countries[
        1:(
          length(
            countries
          ) - 1
        )
      ],
      collapse = ", "
    ),
    
    ", and ",
    
    countries[
      length(
        countries
      )
    ]
    
  )
  
}


get_concordance_count <- function(
    pattern_name
) {
  
  selected_row <- trend_concordance_summary |>
    
    filter(
      burden_severity_significance_pattern ==
        pattern_name
    )
  
  
  if (
    nrow(
      selected_row
    ) == 0
  ) {
    
    return(
      tibble(
        number_of_countries = 0,
        percentage_of_countries = 0
      )
    )
    
  }
  
  
  selected_row |>
    
    select(
      number_of_countries,
      percentage_of_countries
    )
  
}


# ==============================================================================
# 68. EXTRACT TREND COUNTS FOR RESULTS TEXT
# ==============================================================================

incidence_increasing <- get_trend_count(
  "Incidence ASR",
  "Increasing"
)

incidence_stable <- get_trend_count(
  "Incidence ASR",
  "Stable"
)

incidence_decreasing <- get_trend_count(
  "Incidence ASR",
  "Decreasing"
)


mortality_increasing <- get_trend_count(
  "Mortality ASR",
  "Increasing"
)

mortality_stable <- get_trend_count(
  "Mortality ASR",
  "Stable"
)

mortality_decreasing <- get_trend_count(
  "Mortality ASR",
  "Decreasing"
)


daly_increasing <- get_trend_count(
  "DALY ASR",
  "Increasing"
)

daly_stable <- get_trend_count(
  "DALY ASR",
  "Stable"
)

daly_decreasing <- get_trend_count(
  "DALY ASR",
  "Decreasing"
)


severity_increasing <- get_trend_count(
  "DALYs per case",
  "Increasing"
)

severity_stable <- get_trend_count(
  "DALYs per case",
  "Stable"
)

severity_decreasing <- get_trend_count(
  "DALYs per case",
  "Decreasing"
)


mir_increasing <- get_trend_count(
  "Mortality-to-incidence ratio",
  "Increasing"
)

mir_stable <- get_trend_count(
  "Mortality-to-incidence ratio",
  "Stable"
)

mir_decreasing <- get_trend_count(
  "Mortality-to-incidence ratio",
  "Decreasing"
)


# ==============================================================================
# 69. EXTRACT INDICATOR SUMMARIES FOR RESULTS TEXT
# ==============================================================================

incidence_summary <- get_indicator_summary(
  "Incidence ASR"
)

mortality_summary <- get_indicator_summary(
  "Mortality ASR"
)

daly_summary <- get_indicator_summary(
  "DALY ASR"
)

severity_summary <- get_indicator_summary(
  "DALYs per case"
)

mir_summary <- get_indicator_summary(
  "Mortality-to-incidence ratio"
)


# ==============================================================================
# 70. EXTRACT CONCORDANCE RESULTS
# ==============================================================================

both_improved <- get_concordance_count(
  "Significant improvement in both"
)

both_deteriorated <- get_concordance_count(
  "Significant deterioration in both"
)

burden_worse_severity_better <- get_concordance_count(
  "Burden worsened, severity improved"
)

burden_better_severity_worse <- get_concordance_count(
  "Burden improved, severity worsened"
)

at_least_one_stable <- get_concordance_count(
  "At least one stable trend"
)


# ==============================================================================
# 71. CREATE RESULTS TEXT
# ==============================================================================

results_text <- glue::glue(
  "3.7 Country-level temporal trends in burden and severity, 1990–2023

Country-level temporal trends showed substantial heterogeneity across the 204 countries and territories included in the analysis (Figure 7; Table 7). For age-standardized incidence, {incidence_increasing$number_of_countries} countries ({format_number(incidence_increasing$percentage_of_countries, 1)}%) had significant increasing trends, {incidence_decreasing$number_of_countries} ({format_number(incidence_decreasing$percentage_of_countries, 1)}%) had significant decreasing trends, and {incidence_stable$number_of_countries} ({format_number(incidence_stable$percentage_of_countries, 1)}%) were classified as stable. The median country-level incidence AAPC was {format_number(incidence_summary$median_aapc, 2)}% (IQR {format_number(incidence_summary$q1_aapc, 2)} to {format_number(incidence_summary$q3_aapc, 2)}). The largest incidence increases were observed in {format_country_list(top_incidence_increases)}, whereas the largest decreases occurred in {format_country_list(top_incidence_decreases)} (Table 7).

Age-standardized mortality decreased significantly in {mortality_decreasing$number_of_countries} countries and territories ({format_number(mortality_decreasing$percentage_of_countries, 1)}%), increased in {mortality_increasing$number_of_countries} ({format_number(mortality_increasing$percentage_of_countries, 1)}%), and remained stable in {mortality_stable$number_of_countries} ({format_number(mortality_stable$percentage_of_countries, 1)}%). The median mortality AAPC was {format_number(mortality_summary$median_aapc, 2)}% (IQR {format_number(mortality_summary$q1_aapc, 2)} to {format_number(mortality_summary$q3_aapc, 2)}). The greatest mortality increases occurred in {format_country_list(top_mortality_increases)}, while the largest declines were observed in {format_country_list(top_mortality_decreases)} (Figure 7; Table 7).

For the age-standardized DALY rate, {daly_decreasing$number_of_countries} countries ({format_number(daly_decreasing$percentage_of_countries, 1)}%) showed significant declines, {daly_increasing$number_of_countries} ({format_number(daly_increasing$percentage_of_countries, 1)}%) showed significant increases, and {daly_stable$number_of_countries} ({format_number(daly_stable$percentage_of_countries, 1)}%) had stable trends. The median DALY-rate AAPC was {format_number(daly_summary$median_aapc, 2)}% (IQR {format_number(daly_summary$q1_aapc, 2)} to {format_number(daly_summary$q3_aapc, 2)}). The largest increases were observed in {format_country_list(top_daly_increases)}, whereas the largest decreases occurred in {format_country_list(top_daly_decreases)}.

Per-case severity declined significantly in {severity_decreasing$number_of_countries} countries and territories ({format_number(severity_decreasing$percentage_of_countries, 1)}%), increased in {severity_increasing$number_of_countries} ({format_number(severity_increasing$percentage_of_countries, 1)}%), and remained stable in {severity_stable$number_of_countries} ({format_number(severity_stable$percentage_of_countries, 1)}%). The median country-level AAPC in DALYs per case was {format_number(severity_summary$median_aapc, 2)}% (IQR {format_number(severity_summary$q1_aapc, 2)} to {format_number(severity_summary$q3_aapc, 2)}). The most pronounced increases in severity occurred in {format_country_list(top_severity_increases)}, while the largest improvements were observed in {format_country_list(top_severity_decreases)} (Figure 7; Table 7).

The mortality-to-incidence ratio decreased significantly in {mir_decreasing$number_of_countries} countries ({format_number(mir_decreasing$percentage_of_countries, 1)}%), increased in {mir_increasing$number_of_countries} ({format_number(mir_increasing$percentage_of_countries, 1)}%), and was stable in {mir_stable$number_of_countries} ({format_number(mir_stable$percentage_of_countries, 1)}%). Its median AAPC was {format_number(mir_summary$median_aapc, 2)}% (IQR {format_number(mir_summary$q1_aapc, 2)} to {format_number(mir_summary$q3_aapc, 2)}). The largest increases occurred in {format_country_list(top_mir_increases)}, whereas the greatest declines were observed in {format_country_list(top_mir_decreases)}.

Assessment of joint DALY-rate and DALYs-per-case trends identified {both_improved$number_of_countries} countries ({format_number(both_improved$percentage_of_countries, 1)}%) with significant improvements in both population burden and per-case severity, compared with {both_deteriorated$number_of_countries} ({format_number(both_deteriorated$percentage_of_countries, 1)}%) with significant deterioration in both dimensions. In {burden_worse_severity_better$number_of_countries} countries ({format_number(burden_worse_severity_better$percentage_of_countries, 1)}%), the DALY rate increased while DALYs per case declined, whereas {burden_better_severity_worse$number_of_countries} ({format_number(burden_better_severity_worse$percentage_of_countries, 1)}%) experienced a declining DALY rate but increasing per-case severity. At least one of the two trends was stable in {at_least_one_stable$number_of_countries} countries ({format_number(at_least_one_stable$percentage_of_countries, 1)}%).

Country-level AAPCs in the DALY rate and DALYs per case were correlated at Spearman rho={format_number(unname(burden_severity_aapc_correlation$estimate), 3)} (p {format_p_value(burden_severity_aapc_correlation$p.value)}). The correlation between the AAPCs of the mortality-to-incidence ratio and DALYs per case was rho={format_number(unname(mir_severity_aapc_correlation$estimate), 3)} (p {format_p_value(mir_severity_aapc_correlation$p.value)}). Complete country-specific estimates, confidence intervals, p values, trend classifications, 1990 and 2023 values, and percentage changes are provided in Supplementary Table S6."
)


# ==============================================================================
# 72. SAVE RESULTS TEXT
# ==============================================================================

writeLines(
  text = results_text,
  con = results_text_file,
  useBytes = TRUE
)


# ==============================================================================
# 73. DISPLAY KEY RESULTS
# ==============================================================================

message("")
message("==============================================================")
message("TREND-DIRECTION SUMMARY")
message("==============================================================")

print(
  trend_direction_summary_export,
  n = Inf
)

message("")
message("==============================================================")
message("TREND CONCORDANCE")
message("==============================================================")

print(
  trend_concordance_summary,
  n = Inf
)

message("")
message("==============================================================")
message("TREND CORRELATIONS")
message("==============================================================")

print(
  trend_correlation_summary,
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
# 74. VERIFY EXPECTED OUTPUT DIMENSIONS
# ==============================================================================

if (
  nrow(
    supplementary_s6
  ) !=
  expected_aapc_rows
) {
  
  stop(
    paste0(
      "\nSupplementary long table should contain ",
      expected_aapc_rows,
      " rows but contains ",
      nrow(
        supplementary_s6
      ),
      "."
    ),
    call. = FALSE
  )
  
}


if (
  nrow(
    supplementary_s6_wide
  ) !=
  expected_country_count
) {
  
  stop(
    paste0(
      "\nSupplementary wide table should contain ",
      expected_country_count,
      " rows but contains ",
      nrow(
        supplementary_s6_wide
      ),
      "."
    ),
    call. = FALSE
  )
  
}


if (
  nrow(
    table_7_formatted
  ) !=
  5 *
  2 *
  10
) {
  
  warning(
    paste0(
      "\nTable 7 was expected to contain 100 ranked rows but contains ",
      nrow(
        table_7_formatted
      ),
      "."
    )
  )
  
}


# ==============================================================================
# 75. VERIFY OUTPUT FILES
# ==============================================================================

required_output_files <- c(
  
  figure_7_file,
  
  table_7_file,
  
  supplementary_s6_excel_file,
  
  supplementary_s6_csv_file,
  
  results_text_file
  
)


missing_output_files <- required_output_files[
  !file.exists(
    required_output_files
  )
]


if (
  length(
    missing_output_files
  ) > 0
) {
  
  stop(
    paste0(
      "\nThe following required outputs were not created:\n",
      paste(
        missing_output_files,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 76. VERIFY OUTPUT FILE SIZES
# ==============================================================================

output_file_information <- file.info(
  required_output_files
)


empty_output_files <- rownames(
  output_file_information
)[
  is.na(
    output_file_information$size
  ) |
    output_file_information$size <=
    0
]


if (
  length(
    empty_output_files
  ) > 0
) {
  
  stop(
    paste0(
      "\nThe following output files are empty or inaccessible:\n",
      paste(
        empty_output_files,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 77. SAVE ANALYTICAL R OBJECTS
#
# This preserves the principal results for use in later sections, including
# projections, trajectories, reclassification, and sensitivity analyses.
# ==============================================================================

analysis_objects_folder <- file.path(
  publication_folder,
  "Analysis_objects"
)

dir.create(
  analysis_objects_folder,
  recursive = TRUE,
  showWarnings = FALSE
)

analysis_objects_file <- file.path(
  analysis_objects_folder,
  "Section_3_7_Country_temporal_trends_1990_2023.rds"
)

analysis_objects_to_save <- list(
  
  country_year_data =
    country_year_data,
  
  country_trend_long =
    country_trend_long,
  
  country_trend_results_long =
    country_trend_results_long,
  
  country_aapc_wide =
    country_aapc_wide,
  
  trend_direction_summary =
    trend_direction_summary,
  
  indicator_aapc_summary =
    indicator_aapc_summary,
  
  trend_concordance =
    trend_concordance,
  
  trend_concordance_summary =
    trend_concordance_summary,
  
  trend_correlation_summary =
    trend_correlation_summary,
  
  countries_not_mapped =
    countries_not_mapped
  
)

saveRDS(
  object = analysis_objects_to_save,
  file = analysis_objects_file
)


if (
  !file.exists(
    analysis_objects_file
  )
) {
  
  stop(
    "\nThe Section 3.7 analytical R object file was not created.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 78. SAVE SESSION INFORMATION
# ==============================================================================

session_information_file <- file.path(
  analysis_objects_folder,
  "Section_3_7_Session_information.txt"
)

session_information <- capture.output(
  sessionInfo()
)

writeLines(
  text = session_information,
  con = session_information_file,
  useBytes = TRUE
)


# ==============================================================================
# 79. FINAL COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("RESULTS SECTION 3.7 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message(
  "Countries and territories analysed: ",
  expected_country_count
)

message(
  "Years analysed: ",
  analysis_start_year,
  "–",
  analysis_end_year
)

message(
  "Country-indicator models completed: ",
  expected_aapc_rows
)

message("")
message("Indicators:")
message("1. Incidence ASR")
message("2. Mortality ASR")
message("3. DALY ASR")
message("4. DALYs per case")
message("5. Mortality-to-incidence ratio")

message("")
message("Figure 7 — TIFF at 600 dpi:")
message(figure_7_file)

message("")
message("Table 7:")
message(table_7_file)

message("")
message("Supplementary Table S6:")
message(supplementary_s6_excel_file)

message("")
message("Results text:")
message(results_text_file)

message("")
message("Saved analytical R objects:")
message(analysis_objects_file)

message("")
message(
  "Countries represented on the world map: ",
  nrow(
    mapped_analysis_countries
  ),
  " of ",
  expected_country_count
)

if (
  nrow(
    countries_not_mapped
  ) > 0
) {
  
  message("")
  message(
    "Countries or territories without Natural Earth polygons remain ",
    "included in every numerical analysis and supplementary table."
  )
  
}

message("")
message("No YLL or YLD data were used.")
message("Global data were not used.")
message("The Clean_Data folder was not used.")
message("==============================================================")