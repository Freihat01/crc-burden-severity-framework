# ==============================================================================
# R_03_14_Internal_validation_DALYs_per_case_vs_MIR_2023.R
#
# RESULTS SECTION 3.14
# Internal convergent validation of DALYs per incident case using the
# mortality-to-incidence ratio (MIR) in 2023
#
# PRIMARY QUESTION
#   Is DALYs per incident case strongly associated with MIR while still
#   retaining country-level information not explained by MIR alone?
#
# REQUIRED PROJECT STRUCTURE
#   Colorectal cancer/
#     Incidence/
#     Mortality/
#     Publication/
#       Figures/
#       Tables/
#       Supplementary/
#
# INPUTS
#   1. Existing country burden-severity file:
#      Publication/Supplementary/
#      Table_S1_Complete_country_burden_severity_2023.*
#      or
#      Table_S2_Complete_country_burden_severity_2023.*
#
#   2. If MIR is not already present in that file, the script derives MIR from:
#      Incidence/*Incidence*Number*Both*AllAges*1990_2023*
#      Mortality/*Deaths*Number*Both*AllAges*1990_2023*
#
# ANALYSES
#   - Pearson correlation
#   - Spearman correlation
#   - Linear regression
#   - Quadratic regression
#   - Restricted cubic spline regression
#   - AIC-based functional-form comparison
#   - Residual and influence diagnostics
#   - Countries with DALYs per case higher/lower than expected for their MIR
#
# OUTPUTS
#   Publication/Figures/
#     Figure_14_Internal_validation_DALYs_per_case_vs_MIR_2023.tiff
#
#   Publication/Tables/
#     Table_14_Internal_validation_DALYs_per_case_vs_MIR.xlsx
#
#   Publication/Supplementary/
#     Table_S14_Country_DALYs_per_case_MIR_validation_2023.xlsx
#     Table_S14_Country_DALYs_per_case_MIR_validation_2023.csv
#     Figure_14_Source_data.xlsx
# ==============================================================================


# ==============================================================================
# 1. CLEAR SESSION
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
# 2. INSTALL AND LOAD PACKAGES
# ==============================================================================

required_packages <- c(
  "tidyverse",
  "readxl",
  "openxlsx",
  "janitor",
  "ggrepel",
  "patchwork",
  "broom",
  "splines",
  "scales"
)

installed_packages <- rownames(installed.packages())
missing_packages <- setdiff(required_packages, installed_packages)

if (length(missing_packages) > 0) {
  install.packages(missing_packages, dependencies = TRUE)
}

suppressPackageStartupMessages({
  library(tidyverse)
  library(readxl)
  library(openxlsx)
  library(janitor)
  library(ggrepel)
  library(patchwork)
  library(broom)
  library(splines)
  library(scales)
})


# ==============================================================================
# 3. IDENTIFY PROJECT ROOT
# ==============================================================================

find_project_root <- function(start_path = getwd()) {
  
  current_path <- normalizePath(
    start_path,
    winslash = "/",
    mustWork = TRUE
  )
  
  repeat {
    
    required_folders <- c(
      "Incidence",
      "Mortality",
      "Publication"
    )
    
    if (
      all(
        dir.exists(
          file.path(
            current_path,
            required_folders
          )
        )
      )
    ) {
      return(current_path)
    }
    
    parent_path <- dirname(current_path)
    
    if (identical(parent_path, current_path)) {
      break
    }
    
    current_path <- parent_path
  }
  
  stop(
    paste0(
      "\nThe colorectal cancer project root could not be identified.\n",
      "It must contain Incidence, Mortality, and Publication folders.\n\n",
      "Starting directory:\n",
      getwd()
    ),
    call. = FALSE
  )
}

project_root <- find_project_root()

message("")
message("Project root:")
message(project_root)


# ==============================================================================
# 4. DEFINE FOLDERS
# ==============================================================================

incidence_folder <- file.path(project_root, "Incidence")
mortality_folder <- file.path(project_root, "Mortality")

publication_folder <- file.path(project_root, "Publication")
figures_folder <- file.path(publication_folder, "Figures")
tables_folder <- file.path(publication_folder, "Tables")
supplementary_folder <- file.path(publication_folder, "Supplementary")

dir.create(figures_folder, recursive = TRUE, showWarnings = FALSE)
dir.create(tables_folder, recursive = TRUE, showWarnings = FALSE)
dir.create(supplementary_folder, recursive = TRUE, showWarnings = FALSE)


# ==============================================================================
# 5. HELPER FUNCTIONS
# ==============================================================================

clean_text <- function(x) {
  x |>
    as.character() |>
    stringr::str_replace_all("\u00A0", " ") |>
    stringr::str_replace_all("â€“|â€”", "–") |>
    stringr::str_replace_all("TÃ¼rkiye", "Türkiye") |>
    stringr::str_replace_all("CÃ´te d'Ivoire", "Côte d'Ivoire") |>
    stringr::str_squish()
}


standardize_country <- function(x) {
  
  x <- clean_text(x)
  
  dplyr::recode(
    x,
    "United States" = "United States of America",
    "USA" = "United States of America",
    "Russia" = "Russian Federation",
    "Turkey" = "Türkiye",
    "Vietnam" = "Viet Nam",
    "Moldova" = "Republic of Moldova",
    "Iran" = "Iran (Islamic Republic of)",
    "Bolivia" = "Bolivia (Plurinational State of)",
    "Venezuela" = "Venezuela (Bolivarian Republic of)",
    "Syria" = "Syrian Arab Republic",
    "Laos" = "Lao People's Democratic Republic",
    "South Korea" = "Republic of Korea",
    "North Korea" = "Democratic People's Republic of Korea",
    "Tanzania" = "United Republic of Tanzania",
    "Palestinian Territory" = "Palestine",
    "Cape Verde" = "Cabo Verde",
    "Ivory Coast" = "Côte d'Ivoire",
    "Congo, Dem. Rep." = "Democratic Republic of the Congo",
    "Congo, Rep." = "Congo",
    "Micronesia" = "Micronesia (Federated States of)",
    "Brunei" = "Brunei Darussalam",
    "East Timor" = "Timor-Leste",
    "Swaziland" = "Eswatini",
    .default = x
  )
}


read_any_file <- function(file_path) {
  
  extension <- stringr::str_to_lower(
    tools::file_ext(file_path)
  )
  
  if (extension == "csv") {
    return(
      readr::read_csv(
        file_path,
        show_col_types = FALSE,
        progress = FALSE,
        guess_max = 100000,
        locale = readr::locale(encoding = "UTF-8")
      )
    )
  }
  
  if (extension %in% c("xlsx", "xls")) {
    
    sheets <- readxl::excel_sheets(file_path)
    
    preferred_sheets <- c(
      "Complete country dataset",
      "Country data",
      "Country comparison"
    )
    
    selected_sheet <- preferred_sheets[
      preferred_sheets %in% sheets
    ][1]
    
    if (is.na(selected_sheet)) {
      selected_sheet <- sheets[[1]]
    }
    
    message(
      "Reading sheet '",
      selected_sheet,
      "' from ",
      basename(file_path)
    )
    
    return(
      readxl::read_excel(
        file_path,
        sheet = selected_sheet
      )
    )
  }
  
  stop(
    paste0(
      "\nUnsupported file type:\n",
      file_path
    ),
    call. = FALSE
  )
}


find_column <- function(
    data,
    candidates,
    label,
    required = TRUE
) {
  
  matched <- candidates[
    candidates %in% names(data)
  ]
  
  if (length(matched) > 0) {
    return(matched[[1]])
  }
  
  if (required) {
    stop(
      paste0(
        "\nRequired column not found: ",
        label,
        "\n\nAccepted names:\n",
        paste(candidates, collapse = "\n"),
        "\n\nDetected columns:\n",
        paste(names(data), collapse = "\n")
      ),
      call. = FALSE
    )
  }
  
  NA_character_
}


find_file_by_keywords <- function(
    folder,
    required_keywords,
    preferred_keywords = character(),
    label
) {
  
  files <- list.files(
    path = folder,
    full.names = TRUE,
    recursive = TRUE
  )
  
  files <- files[
    !stringr::str_detect(
      basename(files),
      "^~\\$"
    )
  ]
  
  extensions <- stringr::str_to_lower(
    tools::file_ext(files)
  )
  
  files <- files[
    extensions %in% c(
      "csv",
      "xlsx",
      "xls"
    )
  ]
  
  lower_names <- stringr::str_to_lower(
    basename(files)
  )
  
  required_match <- rep(
    TRUE,
    length(files)
  )
  
  for (keyword in required_keywords) {
    required_match <- required_match &
      stringr::str_detect(
        lower_names,
        stringr::fixed(
          stringr::str_to_lower(keyword)
        )
      )
  }
  
  candidates <- files[
    required_match
  ]
  
  if (length(candidates) == 0) {
    stop(
      paste0(
        "\nRequired file not found for ",
        label,
        ".\n\nRequired filename keywords:\n",
        paste(required_keywords, collapse = ", "),
        "\n\nFolder searched:\n",
        folder,
        "\n\nFiles present:\n",
        paste(
          basename(files),
          collapse = "\n"
        )
      ),
      call. = FALSE
    )
  }
  
  candidate_names <- stringr::str_to_lower(
    basename(candidates)
  )
  
  scores <- rep(
    0,
    length(candidates)
  )
  
  for (keyword in preferred_keywords) {
    scores <- scores +
      stringr::str_detect(
        candidate_names,
        stringr::fixed(
          stringr::str_to_lower(keyword)
        )
      )
  }
  
  info <- file.info(candidates)
  
  order_index <- order(
    scores,
    info$mtime,
    decreasing = TRUE
  )
  
  candidates <- candidates[
    order_index
  ]
  
  if (length(candidates) > 1) {
    message("")
    message("Candidate files for ", label, ":")
    print(candidates)
  }
  
  selected <- candidates[[1]]
  
  message("")
  message(label, ":")
  message(selected)
  
  selected
}


p_value_text <- function(p) {
  
  if (is.na(p)) {
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


quadrant_colours <- c(
  "Q1: High burden–high severity" = "#B2182B",
  "Q2: Low burden–high severity" = "#EF8A62",
  "Q3: Low burden–low severity" = "#67A9CF",
  "Q4: High burden–low severity" = "#2166AC",
  "Q1" = "#B2182B",
  "Q2" = "#EF8A62",
  "Q3" = "#67A9CF",
  "Q4" = "#2166AC"
)


# ==============================================================================
# 6. LOCATE COUNTRY BURDEN-SEVERITY FILE
# ==============================================================================

crc_file <- find_file_by_keywords(
  folder = supplementary_folder,
  required_keywords = c(
    "complete_country_burden_severity_2023"
  ),
  preferred_keywords = c(
    "table_s1",
    "xlsx"
  ),
  label = "country burden–severity dataset"
)


# ==============================================================================
# 7. IMPORT COUNTRY BURDEN-SEVERITY DATA
# ==============================================================================

crc_raw <- read_any_file(
  crc_file
) |>
  janitor::clean_names()

country_col <- find_column(
  crc_raw,
  candidates = c(
    "country",
    "location",
    "location_name"
  ),
  label = "country"
)

dalys_per_case_col <- find_column(
  crc_raw,
  candidates = c(
    "dalys_per_case",
    "daly_per_case",
    "dal_ys_per_case"
  ),
  label = "DALYs per case"
)

quadrant_col <- find_column(
  crc_raw,
  candidates = c(
    "quadrant",
    "quadrant_code"
  ),
  label = "quadrant",
  required = FALSE
)

mir_col <- find_column(
  crc_raw,
  candidates = c(
    "mir",
    "mortality_to_incidence_ratio",
    "mortality_incidence_ratio"
  ),
  label = "MIR",
  required = FALSE
)

deaths_col <- find_column(
  crc_raw,
  candidates = c(
    "deaths",
    "death_count",
    "mortality",
    "mortality_number"
  ),
  label = "deaths",
  required = FALSE
)

incidence_col <- find_column(
  crc_raw,
  candidates = c(
    "incident_cases",
    "incidence",
    "incidence_number"
  ),
  label = "incident cases",
  required = FALSE
)

crc_data <- crc_raw |>
  transmute(
    country =
      clean_text(
        .data[[country_col]]
      ),
    
    country_key =
      standardize_country(
        .data[[country_col]]
      ),
    
    dalys_per_case =
      as.numeric(
        .data[[dalys_per_case_col]]
      ),
    
    quadrant =
      if (
        is.na(
          quadrant_col
        )
      ) {
        NA_character_
      } else {
        clean_text(
          .data[[quadrant_col]]
        )
      },
    
    mir_existing =
      if (
        is.na(
          mir_col
        )
      ) {
        NA_real_
      } else {
        as.numeric(
          .data[[mir_col]]
        )
      },
    
    deaths_existing =
      if (
        is.na(
          deaths_col
        )
      ) {
        NA_real_
      } else {
        as.numeric(
          .data[[deaths_col]]
        )
      },
    
    incidence_existing =
      if (
        is.na(
          incidence_col
        )
      ) {
        NA_real_
      } else {
        as.numeric(
          .data[[incidence_col]]
        )
      }
  ) |>
  filter(
    !is.na(
      country_key
    ),
    !is.na(
      dalys_per_case
    ),
    is.finite(
      dalys_per_case
    ),
    dalys_per_case >= 0
  ) |>
  distinct(
    country_key,
    .keep_all = TRUE
  )

message("")
message(
  "Countries in burden–severity dataset: ",
  nrow(
    crc_data
  )
)


# ==============================================================================
# 8. DERIVE MIR IF NOT ALREADY AVAILABLE
# ==============================================================================

mir_available_in_crc <- sum(
  !is.na(
    crc_data$mir_existing
  ) &
    is.finite(
      crc_data$mir_existing
    )
) >=
  0.90 *
  nrow(
    crc_data
  )

counts_available_in_crc <- sum(
  !is.na(
    crc_data$deaths_existing
  ) &
    !is.na(
      crc_data$incidence_existing
    )
) >=
  0.90 *
  nrow(
    crc_data
  )

if (mir_available_in_crc) {
  
  message("")
  message("Using MIR already present in the burden–severity dataset.")
  
  analysis_data <- crc_data |>
    mutate(
      deaths =
        deaths_existing,
      
      incident_cases =
        incidence_existing,
      
      mir =
        mir_existing,
      
      mir_source =
        "Existing burden–severity dataset"
    )
  
} else if (counts_available_in_crc) {
  
  message("")
  message(
    "MIR was not present, but deaths and incident cases were available. ",
    "MIR will be calculated from those columns."
  )
  
  analysis_data <- crc_data |>
    mutate(
      deaths =
        deaths_existing,
      
      incident_cases =
        incidence_existing,
      
      mir =
        if_else(
          incident_cases > 0,
          deaths /
            incident_cases,
          NA_real_
        ),
      
      mir_source =
        "Calculated from existing deaths and incidence columns"
    )
  
} else {
  
  message("")
  message(
    "MIR and/or deaths were not available in the burden–severity file. ",
    "Reading the original incidence and mortality files."
  )
  
  incidence_file <- find_file_by_keywords(
    folder = incidence_folder,
    required_keywords = c(
      "incidence",
      "number",
      "allcountries",
      "1990_2023"
    ),
    preferred_keywords = c(
      "bothsexes",
      "allages",
      "01_"
    ),
    label = "both-sex all-age incidence-number dataset"
  )
  
  deaths_file <- find_file_by_keywords(
    folder = mortality_folder,
    required_keywords = c(
      "deaths",
      "number",
      "allcountries",
      "1990_2023"
    ),
    preferred_keywords = c(
      "bothsexes",
      "allages"
    ),
    label = "both-sex all-age deaths-number dataset"
  )
  
  extract_2023_counts <- function(
    file_path,
    value_name
  ) {
    
    data <- read_any_file(
      file_path
    ) |>
      janitor::clean_names()
    
    location_column <- find_column(
      data,
      c(
        "location",
        "country",
        "location_name"
      ),
      "location"
    )
    
    year_column <- find_column(
      data,
      c(
        "year",
        "year_id"
      ),
      "year"
    )
    
    value_column <- find_column(
      data,
      c(
        "val",
        "value",
        "estimate",
        "mean"
      ),
      "value"
    )
    
    sex_column <- find_column(
      data,
      c(
        "sex",
        "sex_name"
      ),
      "sex",
      required = FALSE
    )
    
    age_column <- find_column(
      data,
      c(
        "age",
        "age_name",
        "age_group"
      ),
      "age",
      required = FALSE
    )
    
    metric_column <- find_column(
      data,
      c(
        "metric",
        "metric_name"
      ),
      "metric",
      required = FALSE
    )
    
    output <- data |>
      transmute(
        country =
          clean_text(
            .data[[location_column]]
          ),
        
        country_key =
          standardize_country(
            .data[[location_column]]
          ),
        
        year =
          as.integer(
            .data[[year_column]]
          ),
        
        sex =
          if (
            is.na(
              sex_column
            )
          ) {
            "Both"
          } else {
            clean_text(
              .data[[sex_column]]
            )
          },
        
        age =
          if (
            is.na(
              age_column
            )
          ) {
            "All ages"
          } else {
            clean_text(
              .data[[age_column]]
            )
          },
        
        metric =
          if (
            is.na(
              metric_column
            )
          ) {
            "Number"
          } else {
            clean_text(
              .data[[metric_column]]
            )
          },
        
        value =
          as.numeric(
            .data[[value_column]]
          )
      ) |>
      filter(
        year == 2023
      )
    
    sex_values <- stringr::str_to_lower(
      unique(
        output$sex
      )
    )
    
    if (
      any(
        sex_values %in%
        c(
          "both",
          "both sexes",
          "both sex"
        )
      )
    ) {
      output <- output |>
        filter(
          stringr::str_to_lower(
            sex
          ) %in%
            c(
              "both",
              "both sexes",
              "both sex"
            )
        )
    }
    
    age_values <- stringr::str_to_lower(
      unique(
        output$age
      )
    )
    
    if (
      any(
        age_values %in%
        c(
          "all ages",
          "all age"
        )
      )
    ) {
      output <- output |>
        filter(
          stringr::str_to_lower(
            age
          ) %in%
            c(
              "all ages",
              "all age"
            )
        )
    }
    
    metric_values <- stringr::str_to_lower(
      unique(
        output$metric
      )
    )
    
    if (
      any(
        metric_values %in%
        c(
          "number",
          "numbers",
          "count",
          "counts"
        )
      )
    ) {
      output <- output |>
        filter(
          stringr::str_to_lower(
            metric
          ) %in%
            c(
              "number",
              "numbers",
              "count",
              "counts"
            )
        )
    }
    
    output |>
      group_by(
        country_key
      ) |>
      summarise(
        country_raw =
          first(
            country
          ),
        
        "{value_name}" :=
          if (
            all(
              is.na(
                value
              )
            )
          ) {
            NA_real_
          } else {
            sum(
              value,
              na.rm = TRUE
            )
          },
        
        .groups =
          "drop"
      )
  }
  
  incidence_2023 <- extract_2023_counts(
    incidence_file,
    "incident_cases"
  )
  
  deaths_2023 <- extract_2023_counts(
    deaths_file,
    "deaths"
  )
  
  count_data <- incidence_2023 |>
    full_join(
      deaths_2023,
      by =
        "country_key",
      suffix = c(
        "_incidence",
        "_deaths"
      )
    )
  
  analysis_data <- crc_data |>
    left_join(
      count_data |>
        select(
          country_key,
          incident_cases,
          deaths
        ),
      by =
        "country_key"
    ) |>
    mutate(
      mir =
        if_else(
          !is.na(
            deaths
          ) &
            !is.na(
              incident_cases
            ) &
            incident_cases > 0,
          deaths /
            incident_cases,
          NA_real_
        ),
      
      mir_source =
        "Calculated from original 2023 incidence and deaths files"
    )
}


# ==============================================================================
# 9. QUALITY CONTROL AND MATCHING AUDIT
# ==============================================================================

analysis_data <- analysis_data |>
  mutate(
    mir =
      as.numeric(
        mir
      )
  )

invalid_rows <- analysis_data |>
  filter(
    is.na(
      mir
    ) |
      !is.finite(
        mir
      ) |
      mir < 0 |
      is.na(
        dalys_per_case
      ) |
      !is.finite(
        dalys_per_case
      ) |
      dalys_per_case < 0
  )

valid_data <- analysis_data |>
  filter(
    !is.na(
      mir
    ),
    is.finite(
      mir
    ),
    mir >= 0,
    !is.na(
      dalys_per_case
    ),
    is.finite(
      dalys_per_case
    ),
    dalys_per_case >= 0
  )

message("")
message(
  "Countries with valid DALYs per case and MIR: ",
  nrow(
    valid_data
  )
)

message(
  "Countries excluded from validation: ",
  nrow(
    invalid_rows
  )
)

if (nrow(valid_data) < 180) {
  stop(
    paste0(
      "\nOnly ",
      nrow(valid_data),
      " countries have valid DALYs per case and MIR.\n",
      "Review the excluded-country audit before analysis."
    ),
    call. = FALSE
  )
}


# ==============================================================================
# 10. DESCRIPTIVE STATISTICS
# ==============================================================================

descriptive_summary <- valid_data |>
  summarise(
    Countries =
      n(),
    
    Mean_MIR =
      mean(
        mir
      ),
    
    SD_MIR =
      sd(
        mir
      ),
    
    Median_MIR =
      median(
        mir
      ),
    
    IQR_MIR_lower =
      as.numeric(
        quantile(
          mir,
          0.25
        )
      ),
    
    IQR_MIR_upper =
      as.numeric(
        quantile(
          mir,
          0.75
        )
      ),
    
    Mean_DALYs_per_case =
      mean(
        dalys_per_case
      ),
    
    SD_DALYs_per_case =
      sd(
        dalys_per_case
      ),
    
    Median_DALYs_per_case =
      median(
        dalys_per_case
      ),
    
    IQR_DALYs_per_case_lower =
      as.numeric(
        quantile(
          dalys_per_case,
          0.25
        )
      ),
    
    IQR_DALYs_per_case_upper =
      as.numeric(
        quantile(
          dalys_per_case,
          0.75
        )
      )
  )


# ==============================================================================
# 11. CORRELATION ANALYSES
# ==============================================================================

pearson_test <- cor.test(
  valid_data$mir,
  valid_data$dalys_per_case,
  method =
    "pearson",
  exact =
    FALSE
)

spearman_test <- cor.test(
  valid_data$mir,
  valid_data$dalys_per_case,
  method =
    "spearman",
  exact =
    FALSE
)

correlation_results <- tibble(
  Method = c(
    "Pearson",
    "Spearman"
  ),
  
  Estimate = c(
    unname(
      pearson_test$estimate
    ),
    unname(
      spearman_test$estimate
    )
  ),
  
  Lower_95_CI = c(
    pearson_test$conf.int[[1]],
    NA_real_
  ),
  
  Upper_95_CI = c(
    pearson_test$conf.int[[2]],
    NA_real_
  ),
  
  P_value = c(
    pearson_test$p.value,
    spearman_test$p.value
  )
)


# ==============================================================================
# 12. FIT LINEAR, QUADRATIC, AND SPLINE MODELS
# ==============================================================================

linear_model <- lm(
  dalys_per_case ~ mir,
  data =
    valid_data
)

quadratic_model <- lm(
  dalys_per_case ~ mir + I(mir^2),
  data =
    valid_data
)

spline_model <- lm(
  dalys_per_case ~ splines::ns(
    mir,
    df = 3
  ),
  data =
    valid_data
)

linear_aic <- AIC(
  linear_model
)

quadratic_aic <- AIC(
  quadratic_model
)

spline_aic <- AIC(
  spline_model
)

best_model_name <- dplyr::case_when(
  spline_aic + 2 <
    min(
      linear_aic,
      quadratic_aic
    ) ~
    "Restricted cubic spline",
  
  quadratic_aic + 2 <
    linear_aic ~
    "Quadratic",
  
  TRUE ~
    "Linear"
)

best_model <- switch(
  best_model_name,
  "Restricted cubic spline" =
    spline_model,
  "Quadratic" =
    quadratic_model,
  linear_model
)

model_comparison <- tibble(
  Model = c(
    "Linear",
    "Quadratic",
    "Restricted cubic spline"
  ),
  
  AIC = c(
    linear_aic,
    quadratic_aic,
    spline_aic
  ),
  
  R_squared = c(
    summary(
      linear_model
    )$r.squared,
    
    summary(
      quadratic_model
    )$r.squared,
    
    summary(
      spline_model
    )$r.squared
  ),
  
  Adjusted_R_squared = c(
    summary(
      linear_model
    )$adj.r.squared,
    
    summary(
      quadratic_model
    )$adj.r.squared,
    
    summary(
      spline_model
    )$adj.r.squared
  ),
  
  Selected =
    Model ==
    best_model_name
)

model_coefficients <- bind_rows(
  broom::tidy(
    linear_model,
    conf.int = TRUE
  ) |>
    mutate(
      Model =
        "Linear"
    ),
  
  broom::tidy(
    quadratic_model,
    conf.int = TRUE
  ) |>
    mutate(
      Model =
        "Quadratic"
    ),
  
  broom::tidy(
    spline_model,
    conf.int = TRUE
  ) |>
    mutate(
      Model =
        "Restricted cubic spline"
    )
)

model_fit <- bind_rows(
  broom::glance(
    linear_model
  ) |>
    mutate(
      Model =
        "Linear"
    ),
  
  broom::glance(
    quadratic_model
  ) |>
    mutate(
      Model =
        "Quadratic"
    ),
  
  broom::glance(
    spline_model
  ) |>
    mutate(
      Model =
        "Restricted cubic spline"
    )
)


# ==============================================================================
# 13. RESIDUAL AND INFLUENCE DIAGNOSTICS
# ==============================================================================

linear_augmented <- broom::augment(
  linear_model
)

diagnostic_data <- valid_data |>
  bind_cols(
    linear_augmented |>
      select(
        .fitted,
        .resid,
        .std.resid,
        .hat,
        .sigma,
        .cooksd
      )
  )

cooks_cutoff <- 4 /
  nrow(
    diagnostic_data
  )

diagnostic_data <- diagnostic_data |>
  mutate(
    influential =
      .cooksd >
      cooks_cutoff,
    
    residual_direction =
      case_when(
        .std.resid >= 2 ~
          "Much higher severity than expected",
        
        .std.resid <= -2 ~
          "Much lower severity than expected",
        
        .std.resid > 0 ~
          "Higher severity than expected",
        
        .std.resid < 0 ~
          "Lower severity than expected",
        
        TRUE ~
          "As expected"
      )
  )

largest_positive_residuals <- diagnostic_data |>
  arrange(
    desc(
      .std.resid
    )
  ) |>
  slice_head(
    n = 15
  )

largest_negative_residuals <- diagnostic_data |>
  arrange(
    .std.resid
  ) |>
  slice_head(
    n = 15
  )

influential_countries <- diagnostic_data |>
  filter(
    influential
  ) |>
  arrange(
    desc(
      .cooksd
    )
  )

residual_shapiro <- shapiro.test(
  residuals(
    linear_model
  )
)

diagnostic_summary <- tibble(
  Diagnostic = c(
    "Countries analysed",
    "Cook's distance cutoff",
    "Influential countries",
    "Residual Shapiro–Wilk statistic",
    "Residual Shapiro–Wilk p value"
  ),
  
  Value = c(
    nrow(
      diagnostic_data
    ),
    
    cooks_cutoff,
    
    nrow(
      influential_countries
    ),
    
    unname(
      residual_shapiro$statistic
    ),
    
    residual_shapiro$p.value
  )
)


# ==============================================================================
# 14. PREDICTION CURVE
# ==============================================================================

prediction_grid <- tibble(
  mir =
    seq(
      min(
        valid_data$mir
      ),
      max(
        valid_data$mir
      ),
      length.out = 300
    )
)

best_prediction <- predict(
  best_model,
  newdata =
    prediction_grid,
  interval =
    "confidence",
  level =
    0.95
)

prediction_data <- bind_cols(
  prediction_grid,
  as.data.frame(
    best_prediction
  )
)


# ==============================================================================
# 15. LABEL SELECTION
# ==============================================================================

label_data <- bind_rows(
  largest_positive_residuals |>
    slice_head(
      n = 6
    ),
  
  largest_negative_residuals |>
    slice_head(
      n = 6
    ),
  
  influential_countries |>
    slice_head(
      n = 4
    )
) |>
  distinct(
    country_key,
    .keep_all = TRUE
  )


# ==============================================================================
# 16. CREATE FIGURE 14
# ==============================================================================

annotation_text <- paste0(
  "Pearson r = ",
  formatC(
    unname(
      pearson_test$estimate
    ),
    format = "f",
    digits = 3
  ),
  ", p ",
  p_value_text(
    pearson_test$p.value
  ),
  "\nSpearman ρ = ",
  formatC(
    unname(
      spearman_test$estimate
    ),
    format = "f",
    digits = 3
  ),
  ", p ",
  p_value_text(
    spearman_test$p.value
  ),
  "\nSelected model: ",
  best_model_name,
  "\nAdjusted R² = ",
  formatC(
    summary(
      best_model
    )$adj.r.squared,
    format = "f",
    digits = 3
  )
)

figure_14a <- ggplot(
  valid_data,
  aes(
    x =
      mir,
    y =
      dalys_per_case,
    fill =
      quadrant
  )
) +
  
  geom_ribbon(
    data =
      prediction_data,
    aes(
      x =
        mir,
      ymin =
        lwr,
      ymax =
        upr
    ),
    inherit.aes =
      FALSE,
    alpha =
      0.17
  ) +
  
  geom_line(
    data =
      prediction_data,
    aes(
      x =
        mir,
      y =
        fit
    ),
    inherit.aes =
      FALSE,
    linewidth =
      1.0
  ) +
  
  geom_point(
    shape = 21,
    size = 2.7,
    alpha = 0.82,
    color = "grey20",
    stroke = 0.25
  ) +
  
  ggrepel::geom_label_repel(
    data =
      label_data,
    aes(
      label =
        country
    ),
    size = 2.8,
    fontface = "bold",
    label.size = 0.18,
    label.padding = unit(
      0.10,
      "lines"
    ),
    fill = scales::alpha(
      "white",
      0.93
    ),
    max.overlaps = Inf,
    box.padding = 0.60,
    point.padding = 0.40,
    force = 2.1,
    force_pull = 0.25,
    min.segment.length = 0,
    segment.size = 0.30,
    segment.color = "grey45",
    seed = 20260814,
    show.legend = FALSE
  ) +
  
  annotate(
    "label",
    x =
      min(
        valid_data$mir
      ) +
      0.03 *
      diff(
        range(
          valid_data$mir
        )
      ),
    y =
      max(
        valid_data$dalys_per_case
      ) *
      0.98,
    label =
      annotation_text,
    hjust = 0,
    vjust = 1,
    size = 3.15,
    label.size = 0.25,
    fill = scales::alpha(
      "white",
      0.93
    )
  ) +
  
  scale_fill_manual(
    values =
      quadrant_colours,
    name =
      "Burden–severity quadrant",
    na.value =
      "grey70"
  ) +
  
  scale_x_continuous(
    labels =
      scales::label_number(
        accuracy = 0.1
      ),
    expand =
      expansion(
        mult = c(
          0.03,
          0.08
        )
      )
  ) +
  
  scale_y_continuous(
    labels =
      scales::label_number(
        accuracy = 0.1
      ),
    expand =
      expansion(
        mult = c(
          0.03,
          0.08
        )
      )
  ) +
  
  labs(
    title =
      "A. Association between MIR and DALYs per incident case",
    
    x =
      "Mortality-to-incidence ratio (MIR)",
    
    y =
      "DALYs per incident case"
  ) +
  
  theme_classic(
    base_size = 11
  ) +
  
  theme(
    plot.title =
      element_text(
        face = "bold",
        size = 12
      ),
    
    axis.title =
      element_text(
        face = "bold"
      ),
    
    legend.position =
      "bottom",
    
    legend.title =
      element_text(
        face = "bold"
      ),
    
    legend.text =
      element_text(
        size = 8
      )
  )


residual_plot_data <- diagnostic_data |>
  mutate(
    residual_category =
      case_when(
        .std.resid >= 2 ~
          "≥2 SD above expected",
        
        .std.resid <= -2 ~
          "≥2 SD below expected",
        
        TRUE ~
          "Within ±2 SD"
      )
  )

residual_colours <- c(
  "≥2 SD above expected" = "#B2182B",
  "Within ±2 SD" = "#737373",
  "≥2 SD below expected" = "#2166AC"
)

figure_14b <- ggplot(
  residual_plot_data,
  aes(
    x =
      .fitted,
    y =
      .std.resid,
    fill =
      residual_category
  )
) +
  
  geom_hline(
    yintercept = 0,
    linewidth = 0.55,
    color = "grey30"
  ) +
  
  geom_hline(
    yintercept = c(
      -2,
      2
    ),
    linetype = "dashed",
    linewidth = 0.55,
    color = "grey40"
  ) +
  
  geom_point(
    shape = 21,
    size = 2.6,
    alpha = 0.82,
    color = "grey20",
    stroke = 0.25
  ) +
  
  ggrepel::geom_label_repel(
    data =
      label_data,
    aes(
      x =
        .fitted,
      y =
        .std.resid,
      label =
        country
    ),
    inherit.aes =
      FALSE,
    size = 2.7,
    fontface = "bold",
    label.size = 0.18,
    label.padding = unit(
      0.10,
      "lines"
    ),
    fill = scales::alpha(
      "white",
      0.93
    ),
    max.overlaps = Inf,
    box.padding = 0.55,
    point.padding = 0.35,
    min.segment.length = 0,
    segment.size = 0.28,
    segment.color = "grey45",
    seed = 20260815
  ) +
  
  scale_fill_manual(
    values =
      residual_colours,
    name =
      NULL
  ) +
  
  labs(
    title =
      "B. Countries deviating from expected DALYs per case",
    
    subtitle =
      "Positive residuals indicate greater health loss than predicted from MIR",
    
    x =
      "DALYs per case predicted by the linear MIR model",
    
    y =
      "Standardized residual"
  ) +
  
  theme_classic(
    base_size = 11
  ) +
  
  theme(
    plot.title =
      element_text(
        face = "bold",
        size = 12
      ),
    
    plot.subtitle =
      element_text(
        size = 9.5,
        color = "grey25"
      ),
    
    axis.title =
      element_text(
        face = "bold"
      ),
    
    legend.position =
      "bottom"
  )


figure_14 <- figure_14a /
  figure_14b +
  
  patchwork::plot_layout(
    heights = c(
      1.28,
      0.90
    )
  ) +
  
  patchwork::plot_annotation(
    title =
      "Internal validation of DALYs per incident case using MIR in 2023",
    
    subtitle =
      "Convergent association and country-level deviations from mortality-based severity",
    
    caption =
      paste0(
        "MIR was calculated as deaths divided by incident cases. ",
        "DALYs per incident case was calculated as total DALYs divided by ",
        "incident cases. The fitted curve represents the functional form ",
        "selected using Akaike's information criterion. Residuals in panel B ",
        "are based on the prespecified linear model to retain straightforward ",
        "country-level interpretation."
      ),
    
    theme = theme(
      plot.title =
        element_text(
          face = "bold",
          size = 16,
          hjust = 0.5
        ),
      
      plot.subtitle =
        element_text(
          size = 10.3,
          hjust = 0.5,
          color = "grey25",
          margin = margin(
            b = 8
          )
        ),
      
      plot.caption =
        element_text(
          size = 8.2,
          hjust = 0,
          color = "grey25"
        ),
      
      plot.margin =
        margin(
          12,
          16,
          12,
          16
        )
    )
  )

figure_14_file <- file.path(
  figures_folder,
  "Figure_14_Internal_validation_DALYs_per_case_vs_MIR_2023.tiff"
)

ggsave(
  filename =
    figure_14_file,
  plot =
    figure_14,
  device =
    "tiff",
  width =
    11.5,
  height =
    12.2,
  units =
    "in",
  dpi =
    600,
  compression =
    "lzw",
  bg =
    "white",
  limitsize =
    FALSE
)


# ==============================================================================
# 17. PREPARE TABLE 14
# ==============================================================================

linear_mir_term <- broom::tidy(
  linear_model,
  conf.int = TRUE
) |>
  filter(
    term ==
      "mir"
  )

table_14_primary_results <- tibble(
  Analysis = c(
    "Countries analysed",
    "Pearson correlation",
    "Spearman correlation",
    "Linear regression coefficient for MIR",
    "Linear model R-squared",
    "Linear model adjusted R-squared",
    "Selected functional form",
    "Selected model adjusted R-squared",
    "Countries ≥2 SD above expected severity",
    "Countries ≥2 SD below expected severity",
    "Influential countries by Cook's distance"
  ),
  
  Estimate = c(
    nrow(
      valid_data
    ),
    
    unname(
      pearson_test$estimate
    ),
    
    unname(
      spearman_test$estimate
    ),
    
    linear_mir_term$estimate,
    
    summary(
      linear_model
    )$r.squared,
    
    summary(
      linear_model
    )$adj.r.squared,
    
    NA_real_,
    
    summary(
      best_model
    )$adj.r.squared,
    
    sum(
      diagnostic_data$.std.resid >= 2
    ),
    
    sum(
      diagnostic_data$.std.resid <= -2
    ),
    
    nrow(
      influential_countries
    )
  ),
  
  Lower_95_CI = c(
    NA_real_,
    pearson_test$conf.int[[1]],
    NA_real_,
    linear_mir_term$conf.low,
    rep(
      NA_real_,
      7
    )
  ),
  
  Upper_95_CI = c(
    NA_real_,
    pearson_test$conf.int[[2]],
    NA_real_,
    linear_mir_term$conf.high,
    rep(
      NA_real_,
      7
    )
  ),
  
  P_value = c(
    NA_real_,
    pearson_test$p.value,
    spearman_test$p.value,
    linear_mir_term$p.value,
    rep(
      NA_real_,
      7
    )
  ),
  
  Text_result = c(
    NA_character_,
    NA_character_,
    NA_character_,
    NA_character_,
    NA_character_,
    NA_character_,
    best_model_name,
    NA_character_,
    NA_character_,
    NA_character_,
    NA_character_
  )
)


# ==============================================================================
# 18. PREPARE SUPPLEMENTARY TABLE S14
# ==============================================================================

supplementary_s14 <- diagnostic_data |>
  transmute(
    Country =
      country,
    
    Incident_cases =
      incident_cases,
    
    Deaths =
      deaths,
    
    MIR =
      mir,
    
    DALYs_per_case =
      dalys_per_case,
    
    Quadrant =
      quadrant,
    
    Linear_predicted_DALYs_per_case =
      .fitted,
    
    Linear_residual =
      .resid,
    
    Standardized_residual =
      .std.resid,
    
    Residual_interpretation =
      residual_direction,
    
    Leverage =
      .hat,
    
    Cooks_distance =
      .cooksd,
    
    Influential =
      influential,
    
    MIR_source =
      mir_source
  ) |>
  arrange(
    desc(
      Standardized_residual
    )
  )


# ==============================================================================
# 19. EXPORT TABLES
# ==============================================================================

header_style <- openxlsx::createStyle(
  fontSize = 10,
  textDecoration = "bold",
  halign = "center",
  valign = "center",
  wrapText = TRUE,
  border = "Bottom"
)

write_workbook <- function(
    sheet_list,
    output_file
) {
  
  workbook <- openxlsx::createWorkbook()
  
  for (sheet_name in names(sheet_list)) {
    
    current_data <- sheet_list[[sheet_name]]
    
    openxlsx::addWorksheet(
      workbook,
      sheet_name
    )
    
    openxlsx::writeData(
      workbook,
      sheet =
        sheet_name,
      x =
        current_data
    )
    
    if (ncol(current_data) > 0) {
      
      openxlsx::addStyle(
        workbook,
        sheet =
          sheet_name,
        style =
          header_style,
        rows =
          1,
        cols =
          seq_len(
            ncol(
              current_data
            )
          ),
        gridExpand =
          TRUE
      )
      
      openxlsx::setColWidths(
        workbook,
        sheet =
          sheet_name,
        cols =
          seq_len(
            ncol(
              current_data
            )
          ),
        widths =
          "auto"
      )
      
      openxlsx::freezePane(
        workbook,
        sheet =
          sheet_name,
        firstRow =
          TRUE
      )
    }
  }
  
  openxlsx::saveWorkbook(
    workbook,
    output_file,
    overwrite = TRUE
  )
}


table_14_file <- file.path(
  tables_folder,
  "Table_14_Internal_validation_DALYs_per_case_vs_MIR.xlsx"
)

write_workbook(
  list(
    "Primary results" =
      table_14_primary_results,
    
    "Descriptive summary" =
      descriptive_summary,
    
    "Correlation results" =
      correlation_results,
    
    "Model comparison" =
      model_comparison,
    
    "Model coefficients" =
      model_coefficients,
    
    "Model fit" =
      model_fit,
    
    "Diagnostic summary" =
      diagnostic_summary,
    
    "Higher than expected" =
      largest_positive_residuals,
    
    "Lower than expected" =
      largest_negative_residuals,
    
    "Influential countries" =
      influential_countries,
    
    "Excluded countries" =
      invalid_rows
  ),
  table_14_file
)


s14_excel_file <- file.path(
  supplementary_folder,
  "Table_S14_Country_DALYs_per_case_MIR_validation_2023.xlsx"
)

s14_csv_file <- file.path(
  supplementary_folder,
  "Table_S14_Country_DALYs_per_case_MIR_validation_2023.csv"
)

write_workbook(
  list(
    "Country validation data" =
      supplementary_s14,
    
    "Prediction curve" =
      prediction_data,
    
    "Excluded countries" =
      invalid_rows,
    
    "Model comparison" =
      model_comparison
  ),
  s14_excel_file
)

readr::write_csv(
  supplementary_s14,
  s14_csv_file
)


# ==============================================================================
# 20. EXPORT FIGURE SOURCE DATA
# ==============================================================================

figure_14_source_file <- file.path(
  supplementary_folder,
  "Figure_14_Source_data.xlsx"
)

write_workbook(
  list(
    "Country scatter data" =
      supplementary_s14,
    
    "Fitted curve" =
      prediction_data,
    
    "Correlation results" =
      correlation_results,
    
    "Model comparison" =
      model_comparison
  ),
  figure_14_source_file
)


# ==============================================================================
# 21. DISPLAY KEY RESULTS
# ==============================================================================

message("")
message("==============================================================")
message("INTERNAL VALIDATION RESULTS")
message("==============================================================")

message("")
message(
  "Countries analysed: ",
  nrow(
    valid_data
  )
)

message(
  "Pearson r: ",
  formatC(
    unname(
      pearson_test$estimate
    ),
    format = "f",
    digits = 3
  ),
  "; p ",
  p_value_text(
    pearson_test$p.value
  )
)

message(
  "Spearman rho: ",
  formatC(
    unname(
      spearman_test$estimate
    ),
    format = "f",
    digits = 3
  ),
  "; p ",
  p_value_text(
    spearman_test$p.value
  )
)

message(
  "Linear MIR coefficient: ",
  formatC(
    linear_mir_term$estimate,
    format = "f",
    digits = 3
  ),
  " (95% CI ",
  formatC(
    linear_mir_term$conf.low,
    format = "f",
    digits = 3
  ),
  " to ",
  formatC(
    linear_mir_term$conf.high,
    format = "f",
    digits = 3
  ),
  ")"
)

message(
  "Selected model: ",
  best_model_name
)

message(
  "Selected model adjusted R-squared: ",
  formatC(
    summary(
      best_model
    )$adj.r.squared,
    format = "f",
    digits = 3
  )
)

message("")
message("Countries with highest positive residuals:")
print(
  largest_positive_residuals |>
    select(
      country,
      mir,
      dalys_per_case,
      .fitted,
      .std.resid
    ),
  n = Inf
)

message("")
message("Countries with lowest negative residuals:")
print(
  largest_negative_residuals |>
    select(
      country,
      mir,
      dalys_per_case,
      .fitted,
      .std.resid
    ),
  n = Inf
)


# ==============================================================================
# 22. VALIDATE OUTPUTS
# ==============================================================================

required_output_files <- c(
  figure_14_file,
  table_14_file,
  s14_excel_file,
  s14_csv_file,
  figure_14_source_file
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

output_information <- file.info(
  required_output_files
)

empty_output_files <- rownames(
  output_information
)[
  is.na(
    output_information$size
  ) |
    output_information$size <= 0
]

if (length(empty_output_files) > 0) {
  stop(
    paste0(
      "\nThe following output files are empty:\n",
      paste(
        empty_output_files,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
}


# ==============================================================================
# 23. COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("SECTION 3.14 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message("Main Figure 14:")
message(figure_14_file)

message("")
message("Main Table 14:")
message(table_14_file)

message("")
message("Supplementary Table S14:")
message(s14_excel_file)

message("")
message("Figure 14 source data:")
message(figure_14_source_file)

message("")
message("All outputs were saved under Publication.")
message("==============================================================")