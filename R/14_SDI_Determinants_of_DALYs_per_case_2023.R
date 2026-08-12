# ==============================================================================
# R_03_13_SDI_Determinants_of_DALYs_per_case_2023.R
#
# RESULTS SECTION 3.13
# Association between socio-demographic development and colorectal cancer
# severity in 2023
#
# PROJECT STRUCTURE
#   Colorectal cancer/
#     Population/
#       SDI_2023.csv
#     Publication/
#       Figures/
#       Tables/
#       Supplementary/
#
# REQUIRED CRC INPUT
#   One of the following existing country-level 2023 files:
#     Publication/Supplementary/Table_S1_Complete_country_burden_severity_2023.*
#     Publication/Supplementary/Table_S2_Complete_country_burden_severity_2023.*
#
# REQUIRED SDI INPUT
#   Population/SDI_2023.csv
#
# PRIMARY OUTCOME
#   DALYs per incident case in 2023
#
# PRIMARY EXPOSURE
#   Socio-demographic Index (SDI) in 2023
#
# ANALYSES
#   1. Country-level merge and audit
#   2. Descriptive statistics
#   3. Pearson correlation
#   4. Spearman correlation
#   5. Linear regression
#   6. Quadratic non-linearity test
#   7. Restricted cubic spline comparison
#   8. SDI quintile comparisons
#   9. Outlier and influence diagnostics
#
# OUTPUTS
#   Publication/Figures/
#     Figure_13_SDI_and_DALYs_per_case_2023.tiff
#
#   Publication/Tables/
#     Table_13_SDI_determinants_of_severity.xlsx
#
#   Publication/Supplementary/
#     Table_S13_Country_SDI_and_DALYs_per_case_2023.xlsx
#     Table_S13_Country_SDI_and_DALYs_per_case_2023.csv
#     Figure_13_Source_data.xlsx
#
# IMPORTANT
#   - SDI is used as a continuous country-level index.
#   - Only 2023, Both-sex SDI values are retained.
#   - Regional and subnational SDI records are removed by matching SDI rows
#     directly to countries in the colorectal cancer dataset.
#   - The primary model is linear unless the non-linear model materially
#     improves fit.
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
  "scales",
  "broom",
  "car",
  "splines"
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
  library(scales)
  library(broom)
  library(car)
  library(splines)
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
      "Population",
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
      "The root folder must contain Population and Publication folders.\n\n",
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

population_folder <- file.path(project_root, "Population")
publication_folder <- file.path(project_root, "Publication")
figures_folder <- file.path(publication_folder, "Figures")
tables_folder <- file.path(publication_folder, "Tables")
supplementary_folder <- file.path(publication_folder, "Supplementary")

dir.create(figures_folder, recursive = TRUE, showWarnings = FALSE)
dir.create(tables_folder, recursive = TRUE, showWarnings = FALSE)
dir.create(supplementary_folder, recursive = TRUE, showWarnings = FALSE)


# ==============================================================================
# 5. LOCATE INPUT FILES
# ==============================================================================

find_file_by_stem <- function(folder, stems, label) {
  
  all_files <- list.files(
    path = folder,
    full.names = TRUE,
    recursive = TRUE
  )
  
  all_files <- all_files[
    !stringr::str_detect(
      basename(all_files),
      "^~\\$"
    )
  ]
  
  valid_extensions <- c(
    "csv",
    "xlsx",
    "xls"
  )
  
  all_files <- all_files[
    stringr::str_to_lower(
      tools::file_ext(all_files)
    ) %in%
      valid_extensions
  ]
  
  matches <- purrr::map(
    stems,
    function(stem) {
      all_files[
        stringr::str_detect(
          stringr::str_to_lower(
            basename(all_files)
          ),
          stringr::fixed(
            stringr::str_to_lower(stem)
          )
        )
      ]
    }
  ) |>
    unlist() |>
    unique()
  
  if (length(matches) == 0) {
    stop(
      paste0(
        "\nRequired file not found for ",
        label,
        ".\n\nAccepted filename stems:\n",
        paste(stems, collapse = "\n"),
        "\n\nFolder searched:\n",
        folder,
        "\n\nFiles currently present:\n",
        paste(
          basename(all_files),
          collapse = "\n"
        )
      ),
      call. = FALSE
    )
  }
  
  if (length(matches) > 1) {
    
    message("")
    message("Multiple candidate files found for ", label, ":")
    print(matches)
    
    info <- file.info(matches)
    
    matches <- matches[
      order(
        info$mtime,
        decreasing = TRUE
      )
    ]
  }
  
  selected <- matches[[1]]
  
  message("")
  message(label, ":")
  message(selected)
  
  selected
}


crc_file <- find_file_by_stem(
  supplementary_folder,
  stems = c(
    "Table_S1_Complete_country_burden_severity_2023",
    "Table_S2_Complete_country_burden_severity_2023"
  ),
  label = "country-level colorectal cancer burden–severity file"
)

sdi_file <- find_file_by_stem(
  population_folder,
  stems = c(
    "SDI_2023"
  ),
  label = "2023 SDI file"
)


# ==============================================================================
# 6. HELPER FUNCTIONS
# ==============================================================================

clean_text <- function(x) {
  x |>
    as.character() |>
    stringr::str_replace_all("\u00A0", " ") |>
    stringr::str_replace_all("â€“|â€”", "–") |>
    stringr::str_replace_all("TÃ¼rkiye", "Türkiye") |>
    stringr::str_replace_all("CÃ´te d'Ivoire", "Côte d'Ivoire") |>
    stringr::str_replace_all("MÃ©xico", "México") |>
    stringr::str_replace_all("SÃ£o", "São") |>
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
    "Viet Nam" = "Viet Nam",
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
        locale = readr::locale(
          encoding = "UTF-8"
        )
      )
    )
  }
  
  if (extension %in% c("xlsx", "xls")) {
    
    sheets <- readxl::excel_sheets(file_path)
    
    selected_sheet <- sheets[[1]]
    
    if (
      "Complete country dataset" %in%
      sheets
    ) {
      selected_sheet <- "Complete country dataset"
    }
    
    if (
      "Table_S1_Complete_country_burden_severity_2023" %in%
      sheets
    ) {
      selected_sheet <-
        "Table_S1_Complete_country_burden_severity_2023"
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
    label
) {
  
  available <- candidates[
    candidates %in%
      names(data)
  ]
  
  if (length(available) == 0) {
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
  
  available[[1]]
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


# ==============================================================================
# 7. IMPORT COUNTRY-LEVEL CRC DATA
# ==============================================================================

crc_raw <- read_any_file(
  crc_file
) |>
  janitor::clean_names()

crc_country_col <- find_column(
  crc_raw,
  candidates = c(
    "country",
    "location",
    "location_name"
  ),
  label = "country"
)

crc_dalys_per_case_col <- find_column(
  crc_raw,
  candidates = c(
    "dalys_per_case",
    "daly_per_case",
    "dal_ys_per_case"
  ),
  label = "DALYs per case"
)

crc_daly_rate_col <- find_column(
  crc_raw,
  candidates = c(
    "age_standardized_daly_rate",
    "age_standardized_daly_rate_per_100_000",
    "daly_rate",
    "age_standardized_daly_rate_"
  ),
  label = "age-standardized DALY rate"
)

crc_quadrant_col <- find_column(
  crc_raw,
  candidates = c(
    "quadrant",
    "quadrant_code"
  ),
  label = "quadrant"
)

crc_data <- crc_raw |>
  transmute(
    country_original =
      clean_text(
        .data[[crc_country_col]]
      ),
    
    country_key =
      standardize_country(
        .data[[crc_country_col]]
      ),
    
    dalys_per_case =
      as.numeric(
        .data[[crc_dalys_per_case_col]]
      ),
    
    age_standardized_daly_rate =
      as.numeric(
        .data[[crc_daly_rate_col]]
      ),
    
    quadrant =
      clean_text(
        .data[[crc_quadrant_col]]
      )
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
  "CRC countries available: ",
  nrow(
    crc_data
  )
)


# ==============================================================================
# 8. IMPORT SDI DATA
# ==============================================================================

sdi_raw <- read_any_file(
  sdi_file
) |>
  janitor::clean_names()

required_sdi_columns <- c(
  "location_name",
  "year_id",
  "sex",
  "mean_value"
)

missing_sdi_columns <- setdiff(
  required_sdi_columns,
  names(sdi_raw)
)

if (length(missing_sdi_columns) > 0) {
  stop(
    paste0(
      "\nThe SDI file is missing required columns:\n",
      paste(
        missing_sdi_columns,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
}

sdi_data_all <- sdi_raw |>
  transmute(
    location_original =
      clean_text(
        location_name
      ),
    
    country_key =
      standardize_country(
        location_name
      ),
    
    year =
      as.integer(
        year_id
      ),
    
    sex =
      clean_text(
        sex
      ),
    
    sdi =
      as.numeric(
        mean_value
      ),
    
    sdi_lower =
      if (
        "lower_value" %in%
        names(sdi_raw)
      ) {
        as.numeric(
          lower_value
        )
      } else {
        NA_real_
      },
    
    sdi_upper =
      if (
        "upper_value" %in%
        names(sdi_raw)
      ) {
        as.numeric(
          upper_value
        )
      } else {
        NA_real_
      }
  ) |>
  filter(
    year == 2023,
    stringr::str_to_lower(
      sex
    ) %in%
      c(
        "both",
        "both sexes"
      ),
    !is.na(
      sdi
    ),
    is.finite(
      sdi
    ),
    sdi >= 0,
    sdi <= 1
  )

# Keep only SDI records that match a CRC country.
# This automatically excludes global, regional, and subnational rows.
sdi_country_data <- sdi_data_all |>
  semi_join(
    crc_data |>
      select(
        country_key
      ),
    by =
      "country_key"
  )

duplicate_sdi_countries <- sdi_country_data |>
  count(
    country_key,
    name =
      "records"
  ) |>
  filter(
    records > 1
  )

if (nrow(duplicate_sdi_countries) > 0) {
  
  duplicate_details <- sdi_country_data |>
    semi_join(
      duplicate_sdi_countries,
      by =
        "country_key"
    ) |>
    arrange(
      country_key,
      location_original
    )
  
  print(
    duplicate_details,
    n = Inf
  )
  
  # Prefer the row whose original location exactly matches the standardized key.
  sdi_country_data <- sdi_country_data |>
    mutate(
      exact_country_name =
        location_original ==
        country_key
    ) |>
    arrange(
      country_key,
      desc(
        exact_country_name
      )
    ) |>
    group_by(
      country_key
    ) |>
    slice_head(
      n = 1
    ) |>
    ungroup() |>
    select(
      -exact_country_name
    )
}


# ==============================================================================
# 9. MERGE CRC AND SDI DATA
# ==============================================================================

analysis_data <- crc_data |>
  left_join(
    sdi_country_data |>
      select(
        country_key,
        location_original,
        sdi,
        sdi_lower,
        sdi_upper
      ),
    by =
      "country_key"
  )

unmatched_crc_countries <- analysis_data |>
  filter(
    is.na(
      sdi
    )
  ) |>
  select(
    country_original,
    country_key
  )

unmatched_sdi_records <- sdi_country_data |>
  anti_join(
    crc_data |>
      select(
        country_key
      ),
    by =
      "country_key"
  )

matched_data <- analysis_data |>
  filter(
    !is.na(
      sdi
    ),
    is.finite(
      sdi
    )
  )

message("")
message(
  "Matched countries: ",
  nrow(
    matched_data
  )
)

message(
  "CRC countries without SDI match: ",
  nrow(
    unmatched_crc_countries
  )
)

if (nrow(matched_data) < 180) {
  stop(
    paste0(
      "\nOnly ",
      nrow(matched_data),
      " countries matched between CRC and SDI data.\n",
      "This is too few for the planned analysis.\n",
      "Review the unmatched-country audit."
    ),
    call. = FALSE
  )
}


# ==============================================================================
# 10. CREATE SDI QUINTILES
# ==============================================================================

matched_data <- matched_data |>
  mutate(
    sdi_quintile_number =
      ntile(
        sdi,
        5
      ),
    
    sdi_quintile =
      factor(
        paste0(
          "Q",
          sdi_quintile_number
        ),
        levels =
          paste0(
            "Q",
            1:5
          ),
        labels = c(
          "Lowest SDI",
          "Low-middle SDI",
          "Middle SDI",
          "High-middle SDI",
          "Highest SDI"
        ),
        ordered =
          TRUE
      )
  )


# ==============================================================================
# 11. DESCRIPTIVE STATISTICS
# ==============================================================================

descriptive_summary <- matched_data |>
  summarise(
    Countries =
      n(),
    
    Mean_SDI =
      mean(
        sdi
      ),
    
    SD_SDI =
      sd(
        sdi
      ),
    
    Median_SDI =
      median(
        sdi
      ),
    
    IQR_SDI_lower =
      as.numeric(
        quantile(
          sdi,
          0.25
        )
      ),
    
    IQR_SDI_upper =
      as.numeric(
        quantile(
          sdi,
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

quintile_summary <- matched_data |>
  group_by(
    sdi_quintile
  ) |>
  summarise(
    Countries =
      n(),
    
    Median_SDI =
      median(
        sdi
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
      ),
    
    .groups =
      "drop"
  )


# ==============================================================================
# 12. CORRELATION ANALYSES
# ==============================================================================

pearson_test <- cor.test(
  matched_data$sdi,
  matched_data$dalys_per_case,
  method =
    "pearson",
  exact =
    FALSE
)

spearman_test <- cor.test(
  matched_data$sdi,
  matched_data$dalys_per_case,
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
# 13. REGRESSION MODELS
# ==============================================================================

linear_model <- lm(
  dalys_per_case ~ sdi,
  data =
    matched_data
)

quadratic_model <- lm(
  dalys_per_case ~ sdi + I(sdi^2),
  data =
    matched_data
)

spline_model <- lm(
  dalys_per_case ~ ns(
    sdi,
    df = 3
  ),
  data =
    matched_data
)

linear_vs_quadratic <- anova(
  linear_model,
  quadratic_model
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

best_model_name <- case_when(
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
  
  Selected_model =
    Model ==
    best_model_name
)

linear_model_results <- broom::tidy(
  linear_model,
  conf.int = TRUE
) |>
  mutate(
    Model =
      "Linear"
  )

quadratic_model_results <- broom::tidy(
  quadratic_model,
  conf.int = TRUE
) |>
  mutate(
    Model =
      "Quadratic"
  )

spline_model_results <- broom::tidy(
  spline_model,
  conf.int = TRUE
) |>
  mutate(
    Model =
      "Restricted cubic spline"
  )

model_coefficients <- bind_rows(
  linear_model_results,
  quadratic_model_results,
  spline_model_results
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
# 14. SDI QUINTILE TESTS
# ==============================================================================

kruskal_test <- kruskal.test(
  dalys_per_case ~ sdi_quintile,
  data =
    matched_data
)

quintile_test_result <- tibble(
  Test =
    "Kruskal–Wallis",
  
  Statistic =
    unname(
      kruskal_test$statistic
    ),
  
  Degrees_of_freedom =
    unname(
      kruskal_test$parameter
    ),
  
  P_value =
    kruskal_test$p.value
)


# ==============================================================================
# 15. MODEL DIAGNOSTICS AND INFLUENCE
# ==============================================================================

linear_diagnostics <- broom::augment(
  linear_model
) |>
  bind_cols(
    matched_data |>
      select(
        country_original,
        country_key,
        quadrant,
        sdi,
        dalys_per_case
      )
  )

influence_cutoff <- 4 /
  nrow(
    matched_data
  )

influential_countries <- linear_diagnostics |>
  filter(
    .cooksd >
      influence_cutoff
  ) |>
  arrange(
    desc(
      .cooksd
    )
  )

largest_positive_residuals <- linear_diagnostics |>
  arrange(
    desc(
      .std.resid
    )
  ) |>
  slice_head(
    n = 10
  )

largest_negative_residuals <- linear_diagnostics |>
  arrange(
    .std.resid
  ) |>
  slice_head(
    n = 10
  )

shapiro_result <- if (
  nrow(
    matched_data
  ) <= 5000
) {
  shapiro.test(
    residuals(
      linear_model
    )
  )
} else {
  NULL
}

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
      matched_data
    ),
    
    influence_cutoff,
    
    nrow(
      influential_countries
    ),
    
    if (
      is.null(
        shapiro_result
      )
    ) {
      NA_real_
    } else {
      unname(
        shapiro_result$statistic
      )
    },
    
    if (
      is.null(
        shapiro_result
      )
    ) {
      NA_real_
    } else {
      shapiro_result$p.value
    }
  )
)


# ==============================================================================
# 16. CREATE PREDICTION DATA
# ==============================================================================

prediction_grid <- tibble(
  sdi =
    seq(
      min(
        matched_data$sdi
      ),
      max(
        matched_data$sdi
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
# 17. SELECT COUNTRY LABELS
# ==============================================================================

label_countries <- bind_rows(
  linear_diagnostics |>
    arrange(
      desc(
        abs(
          .std.resid
        )
      )
    ) |>
    slice_head(
      n = 8
    ),
  
  matched_data |>
    slice_min(
      order_by =
        sdi,
      n = 2
    ),
  
  matched_data |>
    slice_max(
      order_by =
        sdi,
      n = 2
    )
) |>
  distinct(
    country_key,
    .keep_all = TRUE
  )


# ==============================================================================
# 18. CREATE FIGURE 13
# ==============================================================================

correlation_annotation <- paste0(
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

figure_13a <- ggplot(
  matched_data,
  aes(
    x =
      sdi,
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
        sdi,
      ymin =
        lwr,
      ymax =
        upr
    ),
    inherit.aes =
      FALSE,
    alpha =
      0.18
  ) +
  
  geom_line(
    data =
      prediction_data,
    aes(
      x =
        sdi,
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
      label_countries,
    aes(
      label =
        country_original
    ),
    size = 2.8,
    fontface = "bold",
    label.size = 0.18,
    label.padding = unit(0.10, "lines"),
    fill = scales::alpha("white", 0.92),
    max.overlaps = Inf,
    box.padding = 0.55,
    point.padding = 0.35,
    min.segment.length = 0,
    segment.size = 0.30,
    segment.color = "grey45",
    seed = 20260813,
    show.legend = FALSE
  ) +
  
  annotate(
    "label",
    x = 0.97,
    y = 30.8,
    label = correlation_annotation,
    hjust = 1,
    vjust = 1,
    size = 3.2,
    label.size = 0.25,
    fill = scales::alpha("white", 0.95)
  ) +
  
  scale_x_continuous(
    limits = c(
      0,
      1
    ),
    breaks =
      seq(
        0,
        1,
        by = 0.2
      )
  ) +
  
  scale_fill_brewer(
    palette =
      "Set1",
    name =
      "Burden–severity quadrant"
  ) +
  
  labs(
    title =
      "A. Country-level association between SDI and DALYs per case",
    
    x =
      "Socio-demographic Index (SDI)",
    
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


figure_13b <- ggplot(
  matched_data,
  aes(
    x =
      sdi_quintile,
    y =
      dalys_per_case
  )
) +
  
  geom_boxplot(
    width = 0.62,
    outlier.shape = NA,
    linewidth = 0.55
  ) +
  
  geom_jitter(
    width = 0.13,
    height = 0,
    alpha = 0.48,
    size = 1.45
  ) +
  
  stat_summary(
    fun = median,
    geom = "point",
    shape = 23,
    size = 3.0,
    fill = "white"
  ) +
  
  annotate(
    "label",
    x = 4.9,
    y =
      max(
        matched_data$dalys_per_case
      ) *
      0.98,
    label =
      paste0(
        "Kruskal–Wallis p ",
        p_value_text(
          kruskal_test$p.value
        )
      ),
    hjust = 1,
    vjust = 1,
    size = 3.2,
    label.size = 0.25
  ) +
  
  labs(
    title =
      "B. DALYs per case across country SDI quintiles",
    
    x =
      NULL,
    
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
    
    axis.title.y =
      element_text(
        face = "bold"
      ),
    
    axis.text.x =
      element_text(
        angle = 25,
        hjust = 1,
        face = "bold"
      )
  )


figure_13 <- figure_13a /
  figure_13b +
  
  patchwork::plot_layout(
    heights = c(
      1.35,
      0.85
    )
  ) +
  
  patchwork::plot_annotation(
    title =
      "Socio-demographic development and colorectal cancer severity in 2023",
    
    subtitle =
      "Country-level relationship between SDI and DALYs lost per incident colorectal cancer case",
    
    caption =
      paste0(
        "The primary outcome was DALYs per incident case. ",
        "The fitted curve represents the model selected using Akaike's ",
        "information criterion from linear, quadratic, and restricted cubic ",
        "spline specifications. Shading indicates the 95% confidence interval."
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
          size = 10.5,
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

figure_13_file <- file.path(
  figures_folder,
  "Figure_13_SDI_and_DALYs_per_case_2023.tiff"
)

ggsave(
  filename =
    figure_13_file,
  plot =
    figure_13,
  device =
    "tiff",
  width =
    11.5,
  height =
    12,
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
# 19. PREPARE TABLE 13
# ==============================================================================

primary_linear_term <- linear_model_results |>
  filter(
    term ==
      "sdi"
  )

table_13_primary_results <- tibble(
  Analysis = c(
    "Pearson correlation",
    "Spearman correlation",
    "Linear regression coefficient for SDI",
    "Linear model R-squared",
    "Linear model adjusted R-squared",
    "Selected functional form",
    "Selected model adjusted R-squared",
    "Kruskal–Wallis comparison across SDI quintiles"
  ),
  
  Estimate = c(
    unname(
      pearson_test$estimate
    ),
    unname(
      spearman_test$estimate
    ),
    primary_linear_term$estimate,
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
    unname(
      kruskal_test$statistic
    )
  ),
  
  Lower_95_CI = c(
    pearson_test$conf.int[[1]],
    NA_real_,
    primary_linear_term$conf.low,
    NA_real_,
    NA_real_,
    NA_real_,
    NA_real_,
    NA_real_
  ),
  
  Upper_95_CI = c(
    pearson_test$conf.int[[2]],
    NA_real_,
    primary_linear_term$conf.high,
    NA_real_,
    NA_real_,
    NA_real_,
    NA_real_,
    NA_real_
  ),
  
  P_value = c(
    pearson_test$p.value,
    spearman_test$p.value,
    primary_linear_term$p.value,
    NA_real_,
    NA_real_,
    NA_real_,
    NA_real_,
    kruskal_test$p.value
  ),
  
  Text_result = c(
    NA_character_,
    NA_character_,
    NA_character_,
    NA_character_,
    NA_character_,
    best_model_name,
    NA_character_,
    NA_character_
  )
)


# ==============================================================================
# 20. PREPARE SUPPLEMENTARY COUNTRY TABLE
# ==============================================================================

supplementary_s13 <- matched_data |>
  transmute(
    Country =
      country_original,
    
    SDI =
      sdi,
    
    SDI_lower =
      sdi_lower,
    
    SDI_upper =
      sdi_upper,
    
    SDI_quintile =
      as.character(
        sdi_quintile
      ),
    
    DALYs_per_case =
      dalys_per_case,
    
    Age_standardized_DALY_rate =
      age_standardized_daly_rate,
    
    Quadrant =
      quadrant
  ) |>
  left_join(
    linear_diagnostics |>
      transmute(
        Country =
          country_original,
        
        Linear_fitted_DALYs_per_case =
          .fitted,
        
        Linear_residual =
          .resid,
        
        Standardized_residual =
          .std.resid,
        
        Cooks_distance =
          .cooksd,
        
        Influential =
          .cooksd >
          influence_cutoff
      ),
    by =
      "Country"
  ) |>
  arrange(
    SDI
  )


# ==============================================================================
# 21. EXCEL STYLES
# ==============================================================================

header_style <- openxlsx::createStyle(
  fontSize = 10,
  textDecoration = "bold",
  halign = "center",
  valign = "center",
  wrapText = TRUE,
  border = "Bottom"
)


# ==============================================================================
# 22. EXPORT TABLE 13
# ==============================================================================

table_13_file <- file.path(
  tables_folder,
  "Table_13_SDI_determinants_of_severity.xlsx"
)

table_13_workbook <- openxlsx::createWorkbook()

table_13_sheets <- list(
  "Primary results" =
    table_13_primary_results,
  
  "Correlation results" =
    correlation_results,
  
  "Model comparison" =
    model_comparison,
  
  "Model coefficients" =
    model_coefficients,
  
  "Model fit" =
    model_fit,
  
  "SDI quintile summary" =
    quintile_summary,
  
  "SDI quintile test" =
    quintile_test_result,
  
  "Descriptive summary" =
    descriptive_summary,
  
  "Diagnostic summary" =
    diagnostic_summary,
  
  "Influential countries" =
    influential_countries,
  
  "Positive residuals" =
    largest_positive_residuals,
  
  "Negative residuals" =
    largest_negative_residuals,
  
  "Unmatched CRC countries" =
    unmatched_crc_countries,
  
  "Duplicate SDI records" =
    duplicate_sdi_countries
)

for (sheet_name in names(table_13_sheets)) {
  
  current_data <- table_13_sheets[[sheet_name]]
  
  openxlsx::addWorksheet(
    table_13_workbook,
    sheet_name
  )
  
  openxlsx::writeData(
    table_13_workbook,
    sheet =
      sheet_name,
    x =
      current_data
  )
  
  if (ncol(current_data) > 0) {
    
    openxlsx::addStyle(
      table_13_workbook,
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
      table_13_workbook,
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
      table_13_workbook,
      sheet =
        sheet_name,
      firstRow =
        TRUE
    )
  }
}

openxlsx::saveWorkbook(
  table_13_workbook,
  table_13_file,
  overwrite = TRUE
)


# ==============================================================================
# 23. EXPORT SUPPLEMENTARY TABLE S13
# ==============================================================================

s13_excel_file <- file.path(
  supplementary_folder,
  "Table_S13_Country_SDI_and_DALYs_per_case_2023.xlsx"
)

s13_csv_file <- file.path(
  supplementary_folder,
  "Table_S13_Country_SDI_and_DALYs_per_case_2023.csv"
)

s13_workbook <- openxlsx::createWorkbook()

s13_sheets <- list(
  "Country data" =
    supplementary_s13,
  
  "Prediction curve" =
    prediction_data,
  
  "Unmatched CRC countries" =
    unmatched_crc_countries,
  
  "All matched SDI records" =
    sdi_country_data
)

for (sheet_name in names(s13_sheets)) {
  
  current_data <- s13_sheets[[sheet_name]]
  
  openxlsx::addWorksheet(
    s13_workbook,
    sheet_name
  )
  
  openxlsx::writeData(
    s13_workbook,
    sheet =
      sheet_name,
    x =
      current_data
  )
  
  if (ncol(current_data) > 0) {
    openxlsx::addStyle(
      s13_workbook,
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
      s13_workbook,
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
      s13_workbook,
      sheet =
        sheet_name,
      firstRow =
        TRUE
    )
  }
}

openxlsx::saveWorkbook(
  s13_workbook,
  s13_excel_file,
  overwrite = TRUE
)

readr::write_csv(
  supplementary_s13,
  s13_csv_file
)


# ==============================================================================
# 24. EXPORT FIGURE SOURCE DATA
# ==============================================================================

figure_13_source_file <- file.path(
  supplementary_folder,
  "Figure_13_Source_data.xlsx"
)

figure_13_source_workbook <- openxlsx::createWorkbook()

figure_13_source_sheets <- list(
  "Country scatter data" =
    supplementary_s13,
  
  "Fitted curve" =
    prediction_data,
  
  "SDI quintile summary" =
    quintile_summary,
  
  "Correlation results" =
    correlation_results,
  
  "Model comparison" =
    model_comparison
)

for (sheet_name in names(figure_13_source_sheets)) {
  
  openxlsx::addWorksheet(
    figure_13_source_workbook,
    sheet_name
  )
  
  openxlsx::writeData(
    figure_13_source_workbook,
    sheet =
      sheet_name,
    x =
      figure_13_source_sheets[[sheet_name]]
  )
}

openxlsx::saveWorkbook(
  figure_13_source_workbook,
  figure_13_source_file,
  overwrite = TRUE
)


# ==============================================================================
# 25. DISPLAY KEY RESULTS
# ==============================================================================

message("")
message("==============================================================")
message("SDI AND DALYs-PER-CASE RESULTS")
message("==============================================================")

message("")
message(
  "Countries analysed: ",
  nrow(
    matched_data
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
  "Linear coefficient for SDI: ",
  formatC(
    primary_linear_term$estimate,
    format = "f",
    digits = 3
  ),
  " (95% CI ",
  formatC(
    primary_linear_term$conf.low,
    format = "f",
    digits = 3
  ),
  " to ",
  formatC(
    primary_linear_term$conf.high,
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

message(
  "Kruskal-Wallis p value: ",
  p_value_text(
    kruskal_test$p.value
  )
)

message("")
message("SDI quintile summary:")
print(
  quintile_summary,
  n = Inf
)

message("")
message("Unmatched CRC countries:")
print(
  unmatched_crc_countries,
  n = Inf
)


# ==============================================================================
# 26. VALIDATE OUTPUTS
# ==============================================================================

required_output_files <- c(
  figure_13_file,
  table_13_file,
  s13_excel_file,
  s13_csv_file,
  figure_13_source_file
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
      "\nThe following outputs are empty:\n",
      paste(
        empty_output_files,
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
}


# ==============================================================================
# 27. COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("SECTION 3.13 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message("Main Figure 13:")
message(figure_13_file)

message("")
message("Main Table 13:")
message(table_13_file)

message("")
message("Supplementary Table S13:")
message(s13_excel_file)

message("")
message("Figure 13 source data:")
message(figure_13_source_file)

message("")
message("All outputs were saved under Publication.")
message("==============================================================")