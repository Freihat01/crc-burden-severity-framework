# ==============================================================================
# R_03_9_Framework_validation_reclassification_stability.R
#
# RESULTS SECTION 3.9
# Framework validation, reclassification, agreement, threshold sensitivity,
# and bootstrap stability
#
# CORRECT INPUT
#   Table_S1_Complete_country_burden_severity_2023.csv
#   OR
#   Table_S1_Complete_country_burden_severity_2023.xlsx
#
# IMPORTANT
#   - Table S2 is never used.
#   - DALYs are handled correctly even when janitor::clean_names() converts
#     "DALYs" to "dal_ys".
#   - No YLL or YLD data are used.
#   - Outputs are saved only under:
#       Publication/Figures
#       Publication/Tables
#       Publication/Supplementary
#
# OUTPUTS
#   Publication/Figures/
#     Figure_9_Country_reclassification_using_DALYs_per_case.tiff
#
#   Publication/Tables/
#     Table_9_Framework_validation_summary.xlsx
#
#   Publication/Supplementary/
#     Figure_S1_Framework_sensitivity_and_stability.tiff
#     Table_S8_Complete_country_reclassification.xlsx
#     Table_S8_Complete_country_reclassification.csv
#     Table_S9_Framework_stability_analysis.xlsx
#     Table_S9_Framework_stability_analysis.csv
#     Figure_9_Source_data.csv
#     Figure_S1_Source_data.xlsx
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
  "ggalluvial",
  "patchwork",
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
  library(ggalluvial)
  library(patchwork)
  library(scales)
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
  file.path(current_folder, ".."),
  winslash = "/",
  mustWork = TRUE
)

if (dir.exists(file.path(current_folder, "Publication"))) {
  project_root <- current_folder
} else if (dir.exists(file.path(parent_folder, "Publication"))) {
  project_root <- parent_folder
} else {
  stop(
    paste0(
      "\nThe colorectal cancer project root could not be identified.\n",
      "The project root must contain the Publication folder.\n\n",
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
# 4. DEFINE OUTPUT FOLDERS
# ==============================================================================

publication_folder <- file.path(project_root, "Publication")
figures_folder <- file.path(publication_folder, "Figures")
tables_folder <- file.path(publication_folder, "Tables")
supplementary_folder <- file.path(publication_folder, "Supplementary")

dir.create(figures_folder, recursive = TRUE, showWarnings = FALSE)
dir.create(tables_folder, recursive = TRUE, showWarnings = FALSE)
dir.create(supplementary_folder, recursive = TRUE, showWarnings = FALSE)


# ==============================================================================
# 5. LOCATE THE CORRECT TABLE S1 FILE
# ==============================================================================

preferred_files <- c(
  file.path(
    supplementary_folder,
    "Table_S1_Complete_country_burden_severity_2023.csv"
  ),
  file.path(
    supplementary_folder,
    "Table_S1_Complete_country_burden_severity_2023.xlsx"
  )
)

existing_preferred <- preferred_files[file.exists(preferred_files)]

if (length(existing_preferred) > 0) {
  
  # Prefer CSV if both exist.
  input_file <- existing_preferred[1]
  
} else {
  
  exact_candidates <- list.files(
    path = project_root,
    pattern = "^Table_S1_Complete_country_burden_severity_2023\\.(csv|xlsx|xls)$",
    full.names = TRUE,
    recursive = TRUE,
    ignore.case = TRUE
  )
  
  exact_candidates <- exact_candidates[
    !stringr::str_detect(basename(exact_candidates), "^~\\$")
  ]
  
  exact_candidates <- exact_candidates[
    !stringr::str_detect(
      stringr::str_to_lower(basename(exact_candidates)),
      "table_s2"
    )
  ]
  
  if (length(exact_candidates) == 0) {
    stop(
      paste0(
        "\nThe correct Table S1 file was not found.\n\n",
        "Required filename:\n",
        "Table_S1_Complete_country_burden_severity_2023.csv\n",
        "or\n",
        "Table_S1_Complete_country_burden_severity_2023.xlsx\n\n",
        "Place it in:\n",
        supplementary_folder
      ),
      call. = FALSE
    )
  }
  
  if (length(exact_candidates) > 1) {
    message("")
    message("Multiple exact Table S1 files were found:")
    print(exact_candidates)
    
    stop(
      paste0(
        "\nKeep only one definitive Table S1 file, preferably inside:\n",
        supplementary_folder
      ),
      call. = FALSE
    )
  }
  
  input_file <- exact_candidates[[1]]
}

if (
  stringr::str_detect(
    stringr::str_to_lower(basename(input_file)),
    "table_s2"
  )
) {
  stop(
    "\nTable S2 was selected. Section 3.9 requires Table S1.",
    call. = FALSE
  )
}

message("")
message("Correct input selected:")
message(input_file)


# ==============================================================================
# 6. ANALYTICAL SETTINGS
# ==============================================================================

expected_country_count <- 204L

primary_percentile <- 0.50
lower_sensitivity_percentile <- 0.45
upper_sensitivity_percentile <- 0.55

bootstrap_iterations <- 1000L
kappa_bootstrap_iterations <- 2000L
bootstrap_seed <- 20260801L

high_stability_threshold <- 0.90
moderate_stability_threshold <- 0.70

priority_levels <- c(
  "Priority 1",
  "Priority 2",
  "Priority 3",
  "Priority 4"
)

quadrant_code_levels <- c(
  "Q1",
  "Q2",
  "Q3",
  "Q4"
)


# ==============================================================================
# 7. HELPER FUNCTIONS
# ==============================================================================

clean_text <- function(x) {
  x |>
    as.character() |>
    stringr::str_replace_all("\u00A0", " ") |>
    stringr::str_replace_all("â€“|â€”", "–") |>
    stringr::str_replace_all("CÃ´te d'Ivoire", "Côte d'Ivoire") |>
    stringr::str_replace_all("TÃ¼rkiye", "Türkiye") |>
    stringr::str_squish()
}


to_numeric_safely <- function(x) {
  readr::parse_number(
    as.character(x),
    na = c("", "NA", "N/A", "NaN", "NULL", "-")
  )
}


format_number <- function(x, digits = 2) {
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


format_percent <- function(x, digits = 1) {
  ifelse(
    is.na(x),
    "NA",
    paste0(
      formatC(
        x,
        format = "f",
        digits = digits
      ),
      "%"
    )
  )
}


extract_quadrant_code <- function(x) {
  stringr::str_extract(
    clean_text(x),
    "Q[1-4]"
  )
}


find_first_column <- function(data, candidates, label, required = TRUE) {
  
  available <- candidates[candidates %in% names(data)]
  
  if (length(available) > 0) {
    return(available[[1]])
  }
  
  if (required) {
    stop(
      paste0(
        "\nRequired variable not found: ",
        label,
        "\n\nAccepted column names:\n",
        paste(candidates, collapse = "\n"),
        "\n\nDetected columns:\n",
        paste(names(data), collapse = "\n")
      ),
      call. = FALSE
    )
  }
  
  NA_character_
}


create_quadrant <- function(
    daly_rate,
    dalys_per_case,
    burden_threshold,
    severity_threshold
) {
  
  dplyr::case_when(
    daly_rate >= burden_threshold &
      dalys_per_case >= severity_threshold ~
      "Q1: High burden–high severity",
    
    daly_rate < burden_threshold &
      dalys_per_case >= severity_threshold ~
      "Q2: Low burden–high severity",
    
    daly_rate < burden_threshold &
      dalys_per_case < severity_threshold ~
      "Q3: Low burden–low severity",
    
    daly_rate >= burden_threshold &
      dalys_per_case < severity_threshold ~
      "Q4: High burden–low severity",
    
    TRUE ~ NA_character_
  )
}


quadrant_to_priority_number <- function(quadrant_code) {
  dplyr::case_when(
    quadrant_code == "Q1" ~ 1L,
    quadrant_code == "Q2" ~ 2L,
    quadrant_code == "Q4" ~ 3L,
    quadrant_code == "Q3" ~ 4L,
    TRUE ~ NA_integer_
  )
}


priority_number_to_label <- function(x) {
  dplyr::case_when(
    x == 1L ~ "Priority 1",
    x == 2L ~ "Priority 2",
    x == 3L ~ "Priority 3",
    x == 4L ~ "Priority 4",
    TRUE ~ NA_character_
  )
}


create_conventional_priority <- function(daly_rate) {
  
  cutoffs <- quantile(
    daly_rate,
    probs = c(0.25, 0.50, 0.75),
    na.rm = TRUE,
    names = FALSE,
    type = 7
  )
  
  dplyr::case_when(
    daly_rate >= cutoffs[[3]] ~ 1L,
    daly_rate >= cutoffs[[2]] ~ 2L,
    daly_rate >= cutoffs[[1]] ~ 3L,
    TRUE ~ 4L
  )
}


calculate_kappa <- function(
    reference,
    comparison,
    category_levels,
    weighted = FALSE
) {
  
  reference_factor <- factor(
    reference,
    levels = category_levels
  )
  
  comparison_factor <- factor(
    comparison,
    levels = category_levels
  )
  
  complete_rows <- !is.na(reference_factor) & !is.na(comparison_factor)
  
  reference_factor <- reference_factor[complete_rows]
  comparison_factor <- comparison_factor[complete_rows]
  
  n <- length(reference_factor)
  
  if (n == 0) {
    stop(
      "\nNo complete observations were available for kappa calculation.",
      call. = FALSE
    )
  }
  
  confusion_matrix <- table(
    Reference = reference_factor,
    Comparison = comparison_factor
  )
  
  observed_proportions <- confusion_matrix / n
  
  reference_marginal <- rowSums(observed_proportions)
  comparison_marginal <- colSums(observed_proportions)
  
  expected_proportions <- outer(
    reference_marginal,
    comparison_marginal
  )
  
  number_categories <- length(category_levels)
  
  if (weighted) {
    
    category_positions <- seq_len(number_categories)
    
    weight_matrix <- 1 -
      abs(
        outer(
          category_positions,
          category_positions,
          "-"
        )
      ) /
      (number_categories - 1)
    
  } else {
    
    weight_matrix <- diag(number_categories)
  }
  
  observed_agreement <- sum(
    weight_matrix * observed_proportions
  )
  
  expected_agreement <- sum(
    weight_matrix * expected_proportions
  )
  
  kappa <- (
    observed_agreement - expected_agreement
  ) / (
    1 - expected_agreement
  )
  
  exact_agreement <- sum(diag(observed_proportions))
  
  list(
    n = n,
    confusion_matrix = confusion_matrix,
    exact_agreement = exact_agreement,
    observed_agreement = observed_agreement,
    expected_agreement = expected_agreement,
    kappa = kappa
  )
}


bootstrap_kappa_ci <- function(
    data,
    reference_column,
    comparison_column,
    category_levels,
    weighted,
    iterations,
    seed
) {
  
  set.seed(seed)
  n <- nrow(data)
  
  bootstrap_values <- replicate(
    iterations,
    {
      sampled_rows <- sample.int(
        n,
        size = n,
        replace = TRUE
      )
      
      sampled_data <- data[
        sampled_rows,
        ,
        drop = FALSE
      ]
      
      tryCatch(
        calculate_kappa(
          reference = sampled_data[[reference_column]],
          comparison = sampled_data[[comparison_column]],
          category_levels = category_levels,
          weighted = weighted
        )$kappa,
        error = function(e) NA_real_
      )
    }
  )
  
  bootstrap_values <- bootstrap_values[
    is.finite(bootstrap_values)
  ]
  
  if (length(bootstrap_values) < iterations * 0.90) {
    stop(
      "\nMore than 10% of bootstrap kappa iterations failed.",
      call. = FALSE
    )
  }
  
  c(
    lower = as.numeric(
      quantile(
        bootstrap_values,
        probs = 0.025,
        na.rm = TRUE,
        names = FALSE
      )
    ),
    upper = as.numeric(
      quantile(
        bootstrap_values,
        probs = 0.975,
        na.rm = TRUE,
        names = FALSE
      )
    )
  )
}


interpret_kappa <- function(kappa) {
  dplyr::case_when(
    is.na(kappa) ~ "Not estimable",
    kappa < 0 ~ "Less than chance agreement",
    kappa <= 0.20 ~ "Slight agreement",
    kappa <= 0.40 ~ "Fair agreement",
    kappa <= 0.60 ~ "Moderate agreement",
    kappa <= 0.80 ~ "Substantial agreement",
    TRUE ~ "Almost perfect agreement"
  )
}


classify_threshold_change <- function(
    primary_code,
    alternative_code
) {
  
  primary_high_burden <- primary_code %in% c("Q1", "Q4")
  alternative_high_burden <- alternative_code %in% c("Q1", "Q4")
  
  primary_high_severity <- primary_code %in% c("Q1", "Q2")
  alternative_high_severity <- alternative_code %in% c("Q1", "Q2")
  
  burden_changed <- primary_high_burden != alternative_high_burden
  severity_changed <- primary_high_severity != alternative_high_severity
  
  dplyr::case_when(
    !burden_changed & !severity_changed ~
      "Unchanged",
    
    burden_changed & !severity_changed ~
      "Burden status changed only",
    
    !burden_changed & severity_changed ~
      "Severity status changed only",
    
    burden_changed & severity_changed ~
      "Both dimensions changed",
    
    TRUE ~ NA_character_
  )
}


# ==============================================================================
# 8. IMPORT TABLE S1
# ==============================================================================

input_extension <- stringr::str_to_lower(
  tools::file_ext(input_file)
)

if (input_extension == "csv") {
  
  framework_raw <- readr::read_csv(
    file = input_file,
    show_col_types = FALSE,
    progress = FALSE,
    guess_max = 10000,
    locale = readr::locale(
      encoding = "UTF-8"
    )
  )
  
} else if (input_extension %in% c("xlsx", "xls")) {
  
  sheet_names <- readxl::excel_sheets(input_file)
  
  preferred_sheet <- sheet_names[
    stringr::str_detect(
      stringr::str_to_lower(sheet_names),
      "complete country|complete_country|table_s1"
    )
  ]
  
  if (length(preferred_sheet) == 0) {
    preferred_sheet <- sheet_names[[1]]
  } else {
    preferred_sheet <- preferred_sheet[[1]]
  }
  
  framework_raw <- readxl::read_excel(
    path = input_file,
    sheet = preferred_sheet
  )
  
} else {
  
  stop(
    "\nTable S1 must be a CSV or Excel file.",
    call. = FALSE
  )
}

framework_raw <- framework_raw |>
  janitor::clean_names()

message("")
message("Detected columns after clean_names():")
print(names(framework_raw))


# ==============================================================================
# 9. DETECT ACTUAL COLUMN NAMES
#
# janitor::clean_names() may convert:
#   DALYs          -> dal_ys
#   DALYs_per_case -> dal_ys_per_case
# ==============================================================================

country_col <- find_first_column(
  framework_raw,
  c("country", "location", "location_name"),
  "Country"
)

location_id_col <- find_first_column(
  framework_raw,
  c("location_id", "locationid"),
  "Location ID",
  required = FALSE
)

incident_cases_col <- find_first_column(
  framework_raw,
  c("incident_cases", "incidence", "incident_case"),
  "Incident cases"
)

incident_cases_lower_col <- find_first_column(
  framework_raw,
  c(
    "incident_cases_lower_95_ui",
    "incident_cases_lower",
    "incidence_lower_95_ui"
  ),
  "Incident cases lower 95% UI",
  required = FALSE
)

incident_cases_upper_col <- find_first_column(
  framework_raw,
  c(
    "incident_cases_upper_95_ui",
    "incident_cases_upper",
    "incidence_upper_95_ui"
  ),
  "Incident cases upper 95% UI",
  required = FALSE
)

dalys_col <- find_first_column(
  framework_raw,
  c(
    "dalys",
    "dal_ys",
    "daly"
  ),
  "DALYs"
)

dalys_lower_col <- find_first_column(
  framework_raw,
  c(
    "dalys_lower_95_ui",
    "dal_ys_lower_95_ui",
    "daly_lower_95_ui"
  ),
  "DALYs lower 95% UI",
  required = FALSE
)

dalys_upper_col <- find_first_column(
  framework_raw,
  c(
    "dalys_upper_95_ui",
    "dal_ys_upper_95_ui",
    "daly_upper_95_ui"
  ),
  "DALYs upper 95% UI",
  required = FALSE
)

daly_rate_col <- find_first_column(
  framework_raw,
  c(
    "age_standardized_daly_rate",
    "age_standardised_daly_rate",
    "daly_rate",
    "asdr"
  ),
  "Age-standardized DALY rate"
)

daly_rate_lower_col <- find_first_column(
  framework_raw,
  c(
    "daly_rate_lower_95_ui",
    "age_standardized_daly_rate_lower_95_ui",
    "age_standardised_daly_rate_lower_95_ui"
  ),
  "DALY rate lower 95% UI",
  required = FALSE
)

daly_rate_upper_col <- find_first_column(
  framework_raw,
  c(
    "daly_rate_upper_95_ui",
    "age_standardized_daly_rate_upper_95_ui",
    "age_standardised_daly_rate_upper_95_ui"
  ),
  "DALY rate upper 95% UI",
  required = FALSE
)

dalys_per_case_col <- find_first_column(
  framework_raw,
  c(
    "dalys_per_case",
    "dal_ys_per_case",
    "daly_per_case",
    "dalys_per_incident_case",
    "dal_ys_per_incident_case"
  ),
  "DALYs per case"
)

burden_category_col <- find_first_column(
  framework_raw,
  c("burden_category", "burden_group"),
  "Burden category"
)

severity_category_col <- find_first_column(
  framework_raw,
  c("severity_category", "severity_group"),
  "Severity category"
)

quadrant_col <- find_first_column(
  framework_raw,
  c("quadrant", "framework_quadrant"),
  "Quadrant"
)


# ==============================================================================
# 10. PREPARE ANALYTICAL DATA
# ==============================================================================

optional_numeric_column <- function(data, column_name) {
  if (is.na(column_name)) {
    return(rep(NA_real_, nrow(data)))
  }
  
  to_numeric_safely(data[[column_name]])
}

optional_integer_column <- function(data, column_name) {
  if (is.na(column_name)) {
    return(rep(NA_integer_, nrow(data)))
  }
  
  suppressWarnings(
    as.integer(data[[column_name]])
  )
}

framework_data <- tibble(
  country =
    clean_text(
      framework_raw[[country_col]]
    ),
  
  location_id =
    optional_integer_column(
      framework_raw,
      location_id_col
    ),
  
  incident_cases =
    to_numeric_safely(
      framework_raw[[incident_cases_col]]
    ),
  
  incident_cases_lower_95_ui =
    optional_numeric_column(
      framework_raw,
      incident_cases_lower_col
    ),
  
  incident_cases_upper_95_ui =
    optional_numeric_column(
      framework_raw,
      incident_cases_upper_col
    ),
  
  dalys =
    to_numeric_safely(
      framework_raw[[dalys_col]]
    ),
  
  dalys_lower_95_ui =
    optional_numeric_column(
      framework_raw,
      dalys_lower_col
    ),
  
  dalys_upper_95_ui =
    optional_numeric_column(
      framework_raw,
      dalys_upper_col
    ),
  
  daly_rate =
    to_numeric_safely(
      framework_raw[[daly_rate_col]]
    ),
  
  daly_rate_lower_95_ui =
    optional_numeric_column(
      framework_raw,
      daly_rate_lower_col
    ),
  
  daly_rate_upper_95_ui =
    optional_numeric_column(
      framework_raw,
      daly_rate_upper_col
    ),
  
  dalys_per_case =
    to_numeric_safely(
      framework_raw[[dalys_per_case_col]]
    ),
  
  supplied_burden_category =
    clean_text(
      framework_raw[[burden_category_col]]
    ),
  
  supplied_severity_category =
    clean_text(
      framework_raw[[severity_category_col]]
    ),
  
  supplied_quadrant =
    clean_text(
      framework_raw[[quadrant_col]]
    )
) |>
  mutate(
    supplied_quadrant_code =
      extract_quadrant_code(
        supplied_quadrant
      )
  ) |>
  arrange(country)


# ==============================================================================
# 11. VALIDATE COUNTRY DATA
# ==============================================================================

if (nrow(framework_data) != expected_country_count) {
  stop(
    paste0(
      "\nExpected ",
      expected_country_count,
      " countries and territories, but ",
      nrow(framework_data),
      " were imported."
    ),
    call. = FALSE
  )
}

duplicate_countries <- framework_data |>
  count(
    country,
    name = "records"
  ) |>
  filter(records > 1)

if (nrow(duplicate_countries) > 0) {
  print(duplicate_countries, n = Inf)
  
  stop(
    "\nDuplicate country records were detected.",
    call. = FALSE
  )
}

invalid_rows <- framework_data |>
  filter(
    is.na(country) |
      country == "" |
      is.na(daly_rate) |
      !is.finite(daly_rate) |
      daly_rate < 0 |
      is.na(dalys_per_case) |
      !is.finite(dalys_per_case) |
      dalys_per_case < 0 |
      is.na(supplied_quadrant_code)
  )

if (nrow(invalid_rows) > 0) {
  print(invalid_rows, n = Inf)
  
  stop(
    "\nInvalid or missing analytical values were detected.",
    call. = FALSE
  )
}


# ==============================================================================
# 12. RECONSTRUCT AND VERIFY THE PRIMARY FRAMEWORK
# ==============================================================================

primary_burden_threshold <- as.numeric(
  quantile(
    framework_data$daly_rate,
    probs = primary_percentile,
    na.rm = TRUE,
    names = FALSE,
    type = 7
  )
)

primary_severity_threshold <- as.numeric(
  quantile(
    framework_data$dalys_per_case,
    probs = primary_percentile,
    na.rm = TRUE,
    names = FALSE,
    type = 7
  )
)

framework_data <- framework_data |>
  mutate(
    reconstructed_primary_quadrant =
      create_quadrant(
        daly_rate = daly_rate,
        dalys_per_case = dalys_per_case,
        burden_threshold = primary_burden_threshold,
        severity_threshold = primary_severity_threshold
      ),
    
    reconstructed_primary_quadrant_code =
      extract_quadrant_code(
        reconstructed_primary_quadrant
      ),
    
    supplied_classification_matches =
      supplied_quadrant_code ==
      reconstructed_primary_quadrant_code
  )

classification_mismatches <- framework_data |>
  filter(!supplied_classification_matches)

if (nrow(classification_mismatches) > 0) {
  print(
    classification_mismatches |>
      select(
        country,
        daly_rate,
        dalys_per_case,
        supplied_quadrant,
        reconstructed_primary_quadrant
      ),
    n = Inf
  )
  
  stop(
    paste0(
      "\nThe supplied quadrant does not match the reconstructed ",
      "median-based quadrant for ",
      nrow(classification_mismatches),
      " countries. Analysis stopped."
    ),
    call. = FALSE
  )
}

message("")
message(
  "Primary DALY-rate threshold: ",
  format_number(primary_burden_threshold, 2)
)

message(
  "Primary DALYs-per-case threshold: ",
  format_number(primary_severity_threshold, 2)
)


# ==============================================================================
# 13. CREATE CONVENTIONAL AND SEVERITY-AWARE PRIORITIES
# ==============================================================================

framework_data <- framework_data |>
  mutate(
    conventional_priority_number =
      create_conventional_priority(
        daly_rate
      ),
    
    conventional_priority =
      priority_number_to_label(
        conventional_priority_number
      ),
    
    severity_aware_priority_number =
      quadrant_to_priority_number(
        reconstructed_primary_quadrant_code
      ),
    
    severity_aware_priority =
      priority_number_to_label(
        severity_aware_priority_number
      ),
    
    signed_priority_change =
      conventional_priority_number -
      severity_aware_priority_number,
    
    priority_levels_moved =
      abs(
        signed_priority_change
      ),
    
    reclassified =
      signed_priority_change != 0,
    
    reclassification_direction =
      case_when(
        signed_priority_change > 0 ~ "Upward",
        signed_priority_change < 0 ~ "Downward",
        TRUE ~ "Unchanged"
      ),
    
    conventional_priority =
      factor(
        conventional_priority,
        levels = priority_levels,
        ordered = TRUE
      ),
    
    severity_aware_priority =
      factor(
        severity_aware_priority,
        levels = priority_levels,
        ordered = TRUE
      )
  )


# ==============================================================================
# 14. RECLASSIFICATION SUMMARY
# ==============================================================================

total_reclassified <- sum(framework_data$reclassified)
unchanged_countries <- sum(!framework_data$reclassified)

upward_countries <- sum(
  framework_data$reclassification_direction == "Upward"
)

downward_countries <- sum(
  framework_data$reclassification_direction == "Downward"
)

one_level_movements <- sum(
  framework_data$priority_levels_moved == 1
)

two_level_movements <- sum(
  framework_data$priority_levels_moved == 2
)

three_level_movements <- sum(
  framework_data$priority_levels_moved == 3
)

percentage_reclassified <- 100 *
  total_reclassified /
  expected_country_count

reclassification_summary <- framework_data |>
  count(
    reclassification_direction,
    name = "countries"
  ) |>
  complete(
    reclassification_direction = c(
      "Upward",
      "Unchanged",
      "Downward"
    ),
    fill = list(
      countries = 0
    )
  ) |>
  mutate(
    percentage =
      100 *
      countries /
      expected_country_count
  )


# ==============================================================================
# 15. TRANSITION MATRIX
# ==============================================================================

priority_transition_matrix <- table(
  Conventional =
    factor(
      framework_data$conventional_priority,
      levels = priority_levels
    ),
  
  Severity_aware =
    factor(
      framework_data$severity_aware_priority,
      levels = priority_levels
    )
)

priority_transition_long <- as.data.frame(
  priority_transition_matrix
) |>
  rename(
    Conventional_priority = Conventional,
    Severity_aware_priority = Severity_aware,
    Countries = Freq
  ) |>
  mutate(
    Percentage_of_all_countries =
      100 *
      Countries /
      expected_country_count
  )


# ==============================================================================
# 16. AGREEMENT AND KAPPA
# ==============================================================================

unweighted_kappa <- calculate_kappa(
  reference =
    framework_data$conventional_priority,
  
  comparison =
    framework_data$severity_aware_priority,
  
  category_levels =
    priority_levels,
  
  weighted =
    FALSE
)

linear_weighted_kappa <- calculate_kappa(
  reference =
    framework_data$conventional_priority,
  
  comparison =
    framework_data$severity_aware_priority,
  
  category_levels =
    priority_levels,
  
  weighted =
    TRUE
)

unweighted_kappa_ci <- bootstrap_kappa_ci(
  data =
    framework_data,
  
  reference_column =
    "conventional_priority",
  
  comparison_column =
    "severity_aware_priority",
  
  category_levels =
    priority_levels,
  
  weighted =
    FALSE,
  
  iterations =
    kappa_bootstrap_iterations,
  
  seed =
    bootstrap_seed
)

weighted_kappa_ci <- bootstrap_kappa_ci(
  data =
    framework_data,
  
  reference_column =
    "conventional_priority",
  
  comparison_column =
    "severity_aware_priority",
  
  category_levels =
    priority_levels,
  
  weighted =
    TRUE,
  
  iterations =
    kappa_bootstrap_iterations,
  
  seed =
    bootstrap_seed + 1L
)


# ==============================================================================
# 17. PRIORITY 1 TRANSITIONS
# ==============================================================================

framework_data <- framework_data |>
  mutate(
    priority_1_transition =
      case_when(
        conventional_priority_number != 1 &
          severity_aware_priority_number == 1 ~
          "Entered Priority 1",
        
        conventional_priority_number == 1 &
          severity_aware_priority_number != 1 ~
          "Left Priority 1",
        
        conventional_priority_number == 1 &
          severity_aware_priority_number == 1 ~
          "Remained in Priority 1",
        
        TRUE ~ "Outside Priority 1"
      )
  )

entered_priority_1 <- framework_data |>
  filter(priority_1_transition == "Entered Priority 1")

left_priority_1 <- framework_data |>
  filter(priority_1_transition == "Left Priority 1")


# ==============================================================================
# 18. THRESHOLD SENSITIVITY
# ==============================================================================

burden_threshold_45 <- as.numeric(
  quantile(
    framework_data$daly_rate,
    probs = lower_sensitivity_percentile,
    na.rm = TRUE,
    names = FALSE,
    type = 7
  )
)

severity_threshold_45 <- as.numeric(
  quantile(
    framework_data$dalys_per_case,
    probs = lower_sensitivity_percentile,
    na.rm = TRUE,
    names = FALSE,
    type = 7
  )
)

burden_threshold_55 <- as.numeric(
  quantile(
    framework_data$daly_rate,
    probs = upper_sensitivity_percentile,
    na.rm = TRUE,
    names = FALSE,
    type = 7
  )
)

severity_threshold_55 <- as.numeric(
  quantile(
    framework_data$dalys_per_case,
    probs = upper_sensitivity_percentile,
    na.rm = TRUE,
    names = FALSE,
    type = 7
  )
)

framework_data <- framework_data |>
  mutate(
    quadrant_45 =
      create_quadrant(
        daly_rate = daly_rate,
        dalys_per_case = dalys_per_case,
        burden_threshold = burden_threshold_45,
        severity_threshold = severity_threshold_45
      ),
    
    quadrant_55 =
      create_quadrant(
        daly_rate = daly_rate,
        dalys_per_case = dalys_per_case,
        burden_threshold = burden_threshold_55,
        severity_threshold = severity_threshold_55
      ),
    
    quadrant_45_code =
      extract_quadrant_code(
        quadrant_45
      ),
    
    quadrant_55_code =
      extract_quadrant_code(
        quadrant_55
      ),
    
    stable_at_45 =
      quadrant_45_code ==
      reconstructed_primary_quadrant_code,
    
    stable_at_55 =
      quadrant_55_code ==
      reconstructed_primary_quadrant_code,
    
    threshold_change_45 =
      classify_threshold_change(
        reconstructed_primary_quadrant_code,
        quadrant_45_code
      ),
    
    threshold_change_55 =
      classify_threshold_change(
        reconstructed_primary_quadrant_code,
        quadrant_55_code
      )
  )

sensitivity_summary <- bind_rows(
  framework_data |>
    count(
      change_type = threshold_change_45,
      name = "countries"
    ) |>
    mutate(
      threshold = "45th percentile"
    ),
  
  framework_data |>
    count(
      change_type = threshold_change_55,
      name = "countries"
    ) |>
    mutate(
      threshold = "55th percentile"
    )
) |>
  complete(
    threshold = c(
      "45th percentile",
      "55th percentile"
    ),
    
    change_type = c(
      "Unchanged",
      "Burden status changed only",
      "Severity status changed only",
      "Both dimensions changed"
    ),
    
    fill = list(
      countries = 0
    )
  ) |>
  mutate(
    percentage =
      100 *
      countries /
      expected_country_count
  )

kappa_primary_vs_45 <- calculate_kappa(
  reference =
    framework_data$reconstructed_primary_quadrant_code,
  
  comparison =
    framework_data$quadrant_45_code,
  
  category_levels =
    quadrant_code_levels,
  
  weighted =
    FALSE
)

kappa_primary_vs_55 <- calculate_kappa(
  reference =
    framework_data$reconstructed_primary_quadrant_code,
  
  comparison =
    framework_data$quadrant_55_code,
  
  category_levels =
    quadrant_code_levels,
  
  weighted =
    FALSE
)


# ==============================================================================
# 19. BOOTSTRAP STABILITY
# ==============================================================================

set.seed(bootstrap_seed)

bootstrap_quadrants <- matrix(
  NA_character_,
  nrow = expected_country_count,
  ncol = bootstrap_iterations
)

bootstrap_burden_thresholds <- numeric(
  bootstrap_iterations
)

bootstrap_severity_thresholds <- numeric(
  bootstrap_iterations
)

for (iteration in seq_len(bootstrap_iterations)) {
  
  sampled_indices <- sample.int(
    expected_country_count,
    size = expected_country_count,
    replace = TRUE
  )
  
  bootstrap_sample <- framework_data[
    sampled_indices,
    ,
    drop = FALSE
  ]
  
  current_burden_threshold <- median(
    bootstrap_sample$daly_rate,
    na.rm = TRUE
  )
  
  current_severity_threshold <- median(
    bootstrap_sample$dalys_per_case,
    na.rm = TRUE
  )
  
  bootstrap_burden_thresholds[iteration] <-
    current_burden_threshold
  
  bootstrap_severity_thresholds[iteration] <-
    current_severity_threshold
  
  bootstrap_quadrants[, iteration] <-
    extract_quadrant_code(
      create_quadrant(
        daly_rate =
          framework_data$daly_rate,
        
        dalys_per_case =
          framework_data$dalys_per_case,
        
        burden_threshold =
          current_burden_threshold,
        
        severity_threshold =
          current_severity_threshold
      )
    )
}


# ==============================================================================
# 20. COUNTRY-SPECIFIC BOOTSTRAP PROBABILITIES
# ==============================================================================

bootstrap_probability_matrix <- matrix(
  0,
  nrow = expected_country_count,
  ncol = length(quadrant_code_levels)
)

colnames(bootstrap_probability_matrix) <-
  quadrant_code_levels

for (current_quadrant in quadrant_code_levels) {
  
  bootstrap_probability_matrix[, current_quadrant] <-
    rowMeans(
      bootstrap_quadrants == current_quadrant,
      na.rm = TRUE
    )
}

primary_quadrant_stability <- vapply(
  seq_len(expected_country_count),
  function(row_number) {
    
    primary_code <-
      framework_data$reconstructed_primary_quadrant_code[
        row_number
      ]
    
    bootstrap_probability_matrix[
      row_number,
      primary_code
    ]
  },
  numeric(1)
)

most_frequent_bootstrap_quadrant <- apply(
  bootstrap_probability_matrix,
  1,
  function(probabilities) {
    names(
      which.max(probabilities)
    )
  }
)

framework_data <- framework_data |>
  mutate(
    bootstrap_probability_q1 =
      bootstrap_probability_matrix[, "Q1"],
    
    bootstrap_probability_q2 =
      bootstrap_probability_matrix[, "Q2"],
    
    bootstrap_probability_q3 =
      bootstrap_probability_matrix[, "Q3"],
    
    bootstrap_probability_q4 =
      bootstrap_probability_matrix[, "Q4"],
    
    bootstrap_primary_quadrant_stability =
      primary_quadrant_stability,
    
    bootstrap_most_frequent_quadrant =
      most_frequent_bootstrap_quadrant,
    
    bootstrap_probability_high_burden =
      bootstrap_probability_q1 +
      bootstrap_probability_q4,
    
    bootstrap_probability_high_severity =
      bootstrap_probability_q1 +
      bootstrap_probability_q2,
    
    bootstrap_stability_class =
      case_when(
        bootstrap_primary_quadrant_stability >=
          high_stability_threshold ~
          "Highly stable",
        
        bootstrap_primary_quadrant_stability >=
          moderate_stability_threshold ~
          "Moderately stable",
        
        TRUE ~ "Unstable"
      ),
    
    absolute_distance_from_burden_threshold =
      abs(
        daly_rate -
          primary_burden_threshold
      ),
    
    absolute_distance_from_severity_threshold =
      abs(
        dalys_per_case -
          primary_severity_threshold
      ),
    
    relative_distance_from_burden_threshold =
      (
        daly_rate -
          primary_burden_threshold
      ) /
      primary_burden_threshold,
    
    relative_distance_from_severity_threshold =
      (
        dalys_per_case -
          primary_severity_threshold
      ) /
      primary_severity_threshold
  )


# ==============================================================================
# 21. STABILITY SUMMARIES
# ==============================================================================

bootstrap_stability_summary <- framework_data |>
  count(
    bootstrap_stability_class,
    name = "countries"
  ) |>
  complete(
    bootstrap_stability_class = c(
      "Highly stable",
      "Moderately stable",
      "Unstable"
    ),
    fill = list(
      countries = 0
    )
  ) |>
  mutate(
    percentage =
      100 *
      countries /
      expected_country_count
  )

quadrant_stability_summary <- framework_data |>
  group_by(
    primary_quadrant =
      reconstructed_primary_quadrant_code
  ) |>
  summarise(
    countries =
      n(),
    
    median_stability =
      median(
        bootstrap_primary_quadrant_stability
      ),
    
    stability_q1 =
      as.numeric(
        quantile(
          bootstrap_primary_quadrant_stability,
          probs = 0.25,
          names = FALSE
        )
      ),
    
    stability_q3 =
      as.numeric(
        quantile(
          bootstrap_primary_quadrant_stability,
          probs = 0.75,
          names = FALSE
        )
      ),
    
    highly_stable_countries =
      sum(
        bootstrap_stability_class ==
          "Highly stable"
      ),
    
    moderately_stable_countries =
      sum(
        bootstrap_stability_class ==
          "Moderately stable"
      ),
    
    unstable_countries =
      sum(
        bootstrap_stability_class ==
          "Unstable"
      ),
    
    .groups = "drop"
  )

overall_median_stability <- median(
  framework_data$bootstrap_primary_quadrant_stability
)

overall_stability_q1 <- as.numeric(
  quantile(
    framework_data$bootstrap_primary_quadrant_stability,
    probs = 0.25,
    names = FALSE
  )
)

overall_stability_q3 <- as.numeric(
  quantile(
    framework_data$bootstrap_primary_quadrant_stability,
    probs = 0.75,
    names = FALSE
  )
)


# ==============================================================================
# 22. CREATE TABLE 9
# ==============================================================================

table_9_numeric <- tibble(
  Measure = c(
    "Countries and territories analysed",
    "Countries reclassified",
    "Countries reclassified (%)",
    "Upward reclassification",
    "Downward reclassification",
    "Unchanged classification",
    "One-level movements",
    "Two-level movements",
    "Three-level movements",
    "Exact agreement (%)",
    "Expected agreement (%)",
    "Unweighted Cohen's kappa",
    "Unweighted kappa lower 95% CI",
    "Unweighted kappa upper 95% CI",
    "Linear weighted Cohen's kappa",
    "Weighted kappa lower 95% CI",
    "Weighted kappa upper 95% CI",
    "Countries stable at 45th-percentile thresholds",
    "Stability at 45th-percentile thresholds (%)",
    "Kappa: primary versus 45th-percentile framework",
    "Countries stable at 55th-percentile thresholds",
    "Stability at 55th-percentile thresholds (%)",
    "Kappa: primary versus 55th-percentile framework",
    "Median bootstrap stability probability",
    "Bootstrap stability IQR lower",
    "Bootstrap stability IQR upper",
    "Highly stable countries",
    "Moderately stable countries",
    "Unstable countries"
  ),
  
  Value = c(
    expected_country_count,
    total_reclassified,
    percentage_reclassified,
    upward_countries,
    downward_countries,
    unchanged_countries,
    one_level_movements,
    two_level_movements,
    three_level_movements,
    100 * unweighted_kappa$exact_agreement,
    100 * unweighted_kappa$expected_agreement,
    unweighted_kappa$kappa,
    unweighted_kappa_ci[["lower"]],
    unweighted_kappa_ci[["upper"]],
    linear_weighted_kappa$kappa,
    weighted_kappa_ci[["lower"]],
    weighted_kappa_ci[["upper"]],
    sum(framework_data$stable_at_45),
    100 * mean(framework_data$stable_at_45),
    kappa_primary_vs_45$kappa,
    sum(framework_data$stable_at_55),
    100 * mean(framework_data$stable_at_55),
    kappa_primary_vs_55$kappa,
    overall_median_stability,
    overall_stability_q1,
    overall_stability_q3,
    sum(
      framework_data$bootstrap_stability_class ==
        "Highly stable"
    ),
    sum(
      framework_data$bootstrap_stability_class ==
        "Moderately stable"
    ),
    sum(
      framework_data$bootstrap_stability_class ==
        "Unstable"
    )
  )
)

percentage_measures <- c(
  "Countries reclassified (%)",
  "Exact agreement (%)",
  "Expected agreement (%)",
  "Stability at 45th-percentile thresholds (%)",
  "Stability at 55th-percentile thresholds (%)"
)

table_9_formatted <- table_9_numeric |>
  mutate(
    `Reported value` =
      case_when(
        Measure %in% percentage_measures ~
          format_percent(
            Value,
            1
          ),
        
        stringr::str_detect(
          stringr::str_to_lower(Measure),
          "kappa|probability|iqr"
        ) ~
          format_number(
            Value,
            3
          ),
        
        TRUE ~
          format_number(
            Value,
            0
          )
      )
  ) |>
  select(
    Measure,
    `Reported value`
  )


# ==============================================================================
# 23. PREPARE SUPPLEMENTARY TABLE S8
# ==============================================================================

supplementary_s8 <- framework_data |>
  transmute(
    Country = country,
    Location_ID = location_id,
    Incident_cases = incident_cases,
    DALYs = dalys,
    Age_standardized_DALY_rate = daly_rate,
    DALYs_per_case = dalys_per_case,
    
    Conventional_DALY_rate_priority =
      as.character(
        conventional_priority
      ),
    
    Conventional_priority_number =
      conventional_priority_number,
    
    Burden_severity_quadrant =
      reconstructed_primary_quadrant,
    
    Severity_aware_priority =
      as.character(
        severity_aware_priority
      ),
    
    Severity_aware_priority_number =
      severity_aware_priority_number,
    
    Reclassified =
      if_else(
        reclassified,
        "Yes",
        "No"
      ),
    
    Reclassification_direction =
      reclassification_direction,
    
    Signed_priority_change =
      signed_priority_change,
    
    Priority_levels_moved =
      priority_levels_moved,
    
    Priority_1_transition =
      priority_1_transition
  ) |>
  arrange(
    Conventional_priority_number,
    Severity_aware_priority_number,
    Country
  )


# ==============================================================================
# 24. PREPARE SUPPLEMENTARY TABLE S9
# ==============================================================================

supplementary_s9 <- framework_data |>
  transmute(
    Country = country,
    Age_standardized_DALY_rate = daly_rate,
    DALYs_per_case = dalys_per_case,
    Primary_burden_threshold = primary_burden_threshold,
    Primary_severity_threshold = primary_severity_threshold,
    Primary_quadrant = reconstructed_primary_quadrant_code,
    Quadrant_at_45th_percentile = quadrant_45_code,
    
    Stable_at_45th_percentile =
      if_else(
        stable_at_45,
        "Yes",
        "No"
      ),
    
    Change_type_at_45th_percentile =
      threshold_change_45,
    
    Quadrant_at_55th_percentile =
      quadrant_55_code,
    
    Stable_at_55th_percentile =
      if_else(
        stable_at_55,
        "Yes",
        "No"
      ),
    
    Change_type_at_55th_percentile =
      threshold_change_55,
    
    Bootstrap_probability_Q1 =
      bootstrap_probability_q1,
    
    Bootstrap_probability_Q2 =
      bootstrap_probability_q2,
    
    Bootstrap_probability_Q3 =
      bootstrap_probability_q3,
    
    Bootstrap_probability_Q4 =
      bootstrap_probability_q4,
    
    Bootstrap_primary_quadrant_stability =
      bootstrap_primary_quadrant_stability,
    
    Bootstrap_most_frequent_quadrant =
      bootstrap_most_frequent_quadrant,
    
    Bootstrap_probability_high_burden =
      bootstrap_probability_high_burden,
    
    Bootstrap_probability_high_severity =
      bootstrap_probability_high_severity,
    
    Bootstrap_stability_class =
      bootstrap_stability_class,
    
    Absolute_distance_from_burden_threshold =
      absolute_distance_from_burden_threshold,
    
    Absolute_distance_from_severity_threshold =
      absolute_distance_from_severity_threshold,
    
    Relative_distance_from_burden_threshold =
      relative_distance_from_burden_threshold,
    
    Relative_distance_from_severity_threshold =
      relative_distance_from_severity_threshold
  ) |>
  arrange(
    Bootstrap_primary_quadrant_stability,
    Country
  )


# ==============================================================================
# 25. CREATE FIGURE 9
# ==============================================================================

figure_9_data <- framework_data |>
  transmute(
    country = country,
    
    conventional_priority =
      factor(
        conventional_priority,
        levels = priority_levels
      ),
    
    severity_aware_priority =
      factor(
        severity_aware_priority,
        levels = priority_levels
      ),
    
    reclassification_direction =
      factor(
        reclassification_direction,
        levels = c(
          "Upward",
          "Unchanged",
          "Downward"
        )
      )
  )

reclassification_colours <- c(
  "Upward" = "#B2182B",
  "Unchanged" = "#737373",
  "Downward" = "#2166AC"
)

figure_9 <- ggplot(
  figure_9_data,
  aes(
    axis1 = conventional_priority,
    axis2 = severity_aware_priority,
    y = 1
  )
) +
  
  ggalluvial::geom_alluvium(
    aes(
      fill = reclassification_direction
    ),
    width = 0.16,
    alpha = 0.78,
    knot.pos = 0.42,
    color = "grey75",
    linewidth = 0.10
  ) +
  
  ggalluvial::geom_stratum(
    width = 0.18,
    fill = "grey97",
    color = "grey25",
    linewidth = 0.45
  ) +
  
  ggplot2::geom_text(
    stat = "stratum",
    aes(
      label = after_stat(stratum)
    ),
    size = 4.4,
    fontface = "bold"
  ) +
  
  scale_x_discrete(
    limits = c(
      "DALY-rate-only priority",
      "Severity-aware priority"
    ),
    expand = c(
      0.16,
      0.16
    )
  ) +
  
  scale_fill_manual(
    values = reclassification_colours,
    drop = FALSE,
    name = "Priority movement"
  ) +
  
  scale_y_continuous(
    name = "Countries and territories",
    breaks = scales::pretty_breaks(n = 6),
    expand = expansion(
      mult = c(
        0,
        0.03
      )
    )
  ) +
  
  labs(
    title =
      "Country reclassification after incorporating DALYs per case",
    
    subtitle =
      paste0(
        total_reclassified,
        " of ",
        expected_country_count,
        " countries and territories (",
        format_number(
          percentage_reclassified,
          1
        ),
        "%) changed priority level"
      ),
    
    caption =
      paste0(
        "Conventional priority was based on quartiles of the ",
        "age-standardized DALY rate. Severity-aware priority was ordered as ",
        "Priority 1=Q1, Priority 2=Q2, Priority 3=Q4, and Priority 4=Q3. ",
        "Upward movement indicates a higher priority after incorporating ",
        "DALYs per incident case."
      )
  ) +
  
  theme_classic(
    base_size = 12
  ) +
  
  theme(
    plot.title =
      element_text(
        face = "bold",
        size = 16,
        hjust = 0.5
      ),
    
    plot.subtitle =
      element_text(
        size = 11,
        hjust = 0.5,
        margin = margin(
          b = 12
        )
      ),
    
    axis.title.x =
      element_blank(),
    
    axis.text.x =
      element_text(
        face = "bold",
        size = 11,
        margin = margin(
          t = 8
        )
      ),
    
    axis.title.y =
      element_text(
        face = "bold",
        size = 10.5
      ),
    
    axis.text.y =
      element_text(
        size = 9.5
      ),
    
    legend.position =
      "bottom",
    
    legend.title =
      element_text(
        face = "bold"
      ),
    
    legend.text =
      element_text(
        size = 10
      ),
    
    plot.caption =
      element_text(
        size = 8.7,
        hjust = 0,
        margin = margin(
          t = 10
        )
      ),
    
    plot.margin =
      margin(
        12,
        18,
        12,
        18
      )
  )

figure_9_file <- file.path(
  figures_folder,
  "Figure_9_Country_reclassification_using_DALYs_per_case.tiff"
)

ggsave(
  filename = figure_9_file,
  plot = figure_9,
  device = "tiff",
  width = 11.5,
  height = 9,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white",
  limitsize = FALSE
)


# ==============================================================================
# 26. CREATE SUPPLEMENTARY FIGURE S1
# ==============================================================================

sensitivity_plot_data <- sensitivity_summary |>
  mutate(
    threshold =
      factor(
        threshold,
        levels = c(
          "45th percentile",
          "55th percentile"
        )
      ),
    
    change_type =
      factor(
        change_type,
        levels = c(
          "Unchanged",
          "Burden status changed only",
          "Severity status changed only",
          "Both dimensions changed"
        )
      )
  )

figure_s1a <- ggplot(
  sensitivity_plot_data,
  aes(
    x = threshold,
    y = percentage,
    fill = change_type
  )
) +
  
  geom_col(
    width = 0.68
  ) +
  
  geom_text(
    aes(
      label =
        ifelse(
          percentage >= 4,
          paste0(
            formatC(
              percentage,
              format = "f",
              digits = 1
            ),
            "%"
          ),
          ""
        )
    ),
    position = position_stack(vjust = 0.5),
    size = 3.4,
    fontface = "bold"
  ) +
  
  scale_fill_manual(
    values = c(
      "Unchanged" = "#636363",
      "Burden status changed only" = "#67A9CF",
      "Severity status changed only" = "#EF8A62",
      "Both dimensions changed" = "#B2182B"
    ),
    name = NULL
  ) +
  
  scale_y_continuous(
    limits = c(
      0,
      100
    ),
    labels = scales::label_percent(
      scale = 1
    ),
    expand = expansion(
      mult = c(
        0,
        0.02
      )
    )
  ) +
  
  labs(
    title =
      "A. Sensitivity to alternative thresholds",
    
    x = NULL,
    
    y =
      "Countries and territories (%)"
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
        face = "bold"
      ),
    
    legend.position =
      "bottom",
    
    legend.text =
      element_text(
        size = 8.5
      )
  )


figure_s1b <- ggplot(
  framework_data,
  aes(
    x =
      bootstrap_primary_quadrant_stability
  )
) +
  
  geom_histogram(
    binwidth = 0.05,
    boundary = 0,
    fill = "#4D4D4D",
    color = "white",
    linewidth = 0.3
  ) +
  
  geom_vline(
    xintercept = moderate_stability_threshold,
    linetype = "dashed",
    linewidth = 0.7
  ) +
  
  geom_vline(
    xintercept = high_stability_threshold,
    linetype = "dotted",
    linewidth = 0.9
  ) +
  
  scale_x_continuous(
    limits = c(
      0,
      1
    ),
    breaks = seq(
      0,
      1,
      0.1
    ),
    labels = scales::label_percent(
      accuracy = 1
    )
  ) +
  
  labs(
    title =
      "B. Distribution of bootstrap stability",
    
    x =
      "Probability of retaining the primary quadrant",
    
    y =
      "Countries and territories"
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
      )
  )


figure_s1c <- ggplot(
  framework_data,
  aes(
    x =
      factor(
        reconstructed_primary_quadrant_code,
        levels = quadrant_code_levels
      ),
    
    y =
      bootstrap_primary_quadrant_stability
  )
) +
  
  geom_boxplot(
    width = 0.60,
    outlier.shape = NA,
    fill = "grey85",
    color = "grey25"
  ) +
  
  geom_jitter(
    width = 0.15,
    height = 0,
    alpha = 0.60,
    size = 1.5
  ) +
  
  geom_hline(
    yintercept = high_stability_threshold,
    linetype = "dotted",
    linewidth = 0.8
  ) +
  
  scale_y_continuous(
    limits = c(
      0,
      1
    ),
    labels = scales::label_percent(
      accuracy = 1
    )
  ) +
  
  labs(
    title =
      "C. Bootstrap stability by primary quadrant",
    
    x =
      "Primary quadrant",
    
    y =
      "Primary-quadrant stability"
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
      )
  )


lowest_stability_countries <- framework_data |>
  arrange(
    bootstrap_primary_quadrant_stability,
    country
  ) |>
  slice_head(
    n = 20
  ) |>
  mutate(
    country_plot =
      forcats::fct_reorder(
        country,
        bootstrap_primary_quadrant_stability
      )
  )

figure_s1d <- ggplot(
  lowest_stability_countries,
  aes(
    x =
      bootstrap_primary_quadrant_stability,
    
    y =
      country_plot,
    
    fill =
      bootstrap_stability_class
  )
) +
  
  geom_col(
    width = 0.72
  ) +
  
  geom_text(
    aes(
      label =
        scales::percent(
          bootstrap_primary_quadrant_stability,
          accuracy = 1
        )
    ),
    hjust = -0.10,
    size = 3.1,
    fontface = "bold"
  ) +
  
  scale_fill_manual(
    values = c(
      "Highly stable" = "#2166AC",
      "Moderately stable" = "#FDB863",
      "Unstable" = "#B2182B"
    ),
    name = NULL
  ) +
  
  scale_x_continuous(
    limits = c(
      0,
      1.08
    ),
    labels = scales::label_percent(
      accuracy = 1
    )
  ) +
  
  labs(
    title =
      "D. Countries with the lowest bootstrap stability",
    
    x =
      "Primary-quadrant stability",
    
    y = NULL
  ) +
  
  theme_classic(
    base_size = 10.5
  ) +
  
  theme(
    plot.title =
      element_text(
        face = "bold",
        size = 12
      ),
    
    axis.title.x =
      element_text(
        face = "bold"
      ),
    
    axis.text.y =
      element_text(
        size = 8.7
      ),
    
    legend.position =
      "bottom"
  )


figure_s1 <- (
  figure_s1a +
    figure_s1b
) / (
  figure_s1c +
    figure_s1d
) +
  
  patchwork::plot_annotation(
    title =
      "Sensitivity and stability of the burden–severity framework",
    
    caption =
      paste0(
        "Primary classifications used median DALY-rate and DALYs-per-case ",
        "thresholds. Threshold sensitivity used joint 45th- and 55th-percentile ",
        "cut-offs. Bootstrap stability was estimated using ",
        format(
          bootstrap_iterations,
          big.mark = ","
        ),
        " country-resampling iterations."
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

figure_s1_file <- file.path(
  supplementary_folder,
  "Figure_S1_Framework_sensitivity_and_stability.tiff"
)

ggsave(
  filename = figure_s1_file,
  plot = figure_s1,
  device = "tiff",
  width = 14,
  height = 11,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white",
  limitsize = FALSE
)


# ==============================================================================
# 27. EXCEL FORMATTING STYLES
# ==============================================================================

title_style <- openxlsx::createStyle(
  fontSize = 12,
  textDecoration = "bold",
  halign = "left"
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

left_style <- openxlsx::createStyle(
  fontSize = 10,
  halign = "left",
  valign = "center",
  wrapText = TRUE
)


# ==============================================================================
# 28. EXPORT TABLE 9
# ==============================================================================

table_9_file <- file.path(
  tables_folder,
  "Table_9_Framework_validation_summary.xlsx"
)

table_9_workbook <- openxlsx::createWorkbook()

openxlsx::addWorksheet(
  table_9_workbook,
  "Validation summary"
)

openxlsx::writeData(
  table_9_workbook,
  sheet = "Validation summary",
  x =
    "Table 9. Validation, reclassification, and stability of the burden–severity framework",
  startRow = 1,
  startCol = 1
)

openxlsx::writeData(
  table_9_workbook,
  sheet = "Validation summary",
  x = table_9_formatted,
  startRow = 3,
  startCol = 1
)

table_9_note <- paste0(
  "Note: Conventional priority was based on quartiles of the ",
  "age-standardized DALY rate. Severity-aware priority was ordered as ",
  "Priority 1=Q1, Priority 2=Q2, Priority 3=Q4, and Priority 4=Q3. ",
  "Kappa confidence intervals were estimated using ",
  format(
    kappa_bootstrap_iterations,
    big.mark = ","
  ),
  " bootstrap samples. Classification stability was assessed using ",
  "45th- and 55th-percentile thresholds and ",
  format(
    bootstrap_iterations,
    big.mark = ","
  ),
  " bootstrap threshold iterations."
)

openxlsx::writeData(
  table_9_workbook,
  sheet = "Validation summary",
  x = table_9_note,
  startRow = nrow(table_9_formatted) + 6,
  startCol = 1,
  colNames = FALSE
)

additional_table_9_sheets <- list(
  "Numeric results" =
    table_9_numeric,
  
  "Transition matrix" =
    priority_transition_long,
  
  "Reclassification summary" =
    reclassification_summary,
  
  "Sensitivity summary" =
    sensitivity_summary,
  
  "Bootstrap summary" =
    bootstrap_stability_summary,
  
  "Quadrant stability" =
    quadrant_stability_summary
)

for (current_sheet in names(additional_table_9_sheets)) {
  
  openxlsx::addWorksheet(
    table_9_workbook,
    current_sheet
  )
  
  openxlsx::writeData(
    table_9_workbook,
    sheet = current_sheet,
    x = additional_table_9_sheets[[current_sheet]]
  )
}

openxlsx::addStyle(
  table_9_workbook,
  sheet = "Validation summary",
  style = title_style,
  rows = 1,
  cols = 1:2,
  gridExpand = TRUE
)

openxlsx::addStyle(
  table_9_workbook,
  sheet = "Validation summary",
  style = header_style,
  rows = 3,
  cols = 1:2,
  gridExpand = TRUE
)

openxlsx::addStyle(
  table_9_workbook,
  sheet = "Validation summary",
  style = body_style,
  rows = 4:(nrow(table_9_formatted) + 3),
  cols = 1:2,
  gridExpand = TRUE
)

openxlsx::addStyle(
  table_9_workbook,
  sheet = "Validation summary",
  style = left_style,
  rows = 4:(nrow(table_9_formatted) + 3),
  cols = 1,
  gridExpand = TRUE
)

openxlsx::setColWidths(
  table_9_workbook,
  sheet = "Validation summary",
  cols = 1:2,
  widths = c(
    62,
    22
  )
)

openxlsx::freezePane(
  table_9_workbook,
  sheet = "Validation summary",
  firstActiveRow = 4
)

for (current_sheet in names(additional_table_9_sheets)) {
  
  current_data <-
    additional_table_9_sheets[[current_sheet]]
  
  openxlsx::addStyle(
    table_9_workbook,
    sheet = current_sheet,
    style = header_style,
    rows = 1,
    cols = seq_len(ncol(current_data)),
    gridExpand = TRUE
  )
  
  openxlsx::setColWidths(
    table_9_workbook,
    sheet = current_sheet,
    cols = seq_len(ncol(current_data)),
    widths = "auto"
  )
  
  openxlsx::freezePane(
    table_9_workbook,
    sheet = current_sheet,
    firstRow = TRUE
  )
  
  openxlsx::addFilter(
    table_9_workbook,
    sheet = current_sheet,
    rows = 1,
    cols = seq_len(ncol(current_data))
  )
}

openxlsx::saveWorkbook(
  table_9_workbook,
  table_9_file,
  overwrite = TRUE
)


# ==============================================================================
# 29. EXPORT SUPPLEMENTARY TABLE S8
# ==============================================================================

s8_excel_file <- file.path(
  supplementary_folder,
  "Table_S8_Complete_country_reclassification.xlsx"
)

s8_csv_file <- file.path(
  supplementary_folder,
  "Table_S8_Complete_country_reclassification.csv"
)

s8_workbook <- openxlsx::createWorkbook()

s8_sheets <- list(
  "Country reclassification" =
    supplementary_s8,
  
  "Transition matrix" =
    priority_transition_long,
  
  "Entered Priority 1" =
    entered_priority_1,
  
  "Left Priority 1" =
    left_priority_1
)

for (current_sheet in names(s8_sheets)) {
  
  current_data <- s8_sheets[[current_sheet]]
  
  openxlsx::addWorksheet(
    s8_workbook,
    current_sheet
  )
  
  openxlsx::writeData(
    s8_workbook,
    sheet = current_sheet,
    x = current_data
  )
  
  if (ncol(current_data) > 0) {
    
    openxlsx::addStyle(
      s8_workbook,
      sheet = current_sheet,
      style = header_style,
      rows = 1,
      cols = seq_len(ncol(current_data)),
      gridExpand = TRUE
    )
    
    openxlsx::setColWidths(
      s8_workbook,
      sheet = current_sheet,
      cols = seq_len(ncol(current_data)),
      widths = "auto"
    )
    
    openxlsx::freezePane(
      s8_workbook,
      sheet = current_sheet,
      firstRow = TRUE
    )
    
    openxlsx::addFilter(
      s8_workbook,
      sheet = current_sheet,
      rows = 1,
      cols = seq_len(ncol(current_data))
    )
  }
}

openxlsx::saveWorkbook(
  s8_workbook,
  s8_excel_file,
  overwrite = TRUE
)

readr::write_csv(
  supplementary_s8,
  s8_csv_file
)


# ==============================================================================
# 30. EXPORT SUPPLEMENTARY TABLE S9
# ==============================================================================

s9_excel_file <- file.path(
  supplementary_folder,
  "Table_S9_Framework_stability_analysis.xlsx"
)

s9_csv_file <- file.path(
  supplementary_folder,
  "Table_S9_Framework_stability_analysis.csv"
)

bootstrap_threshold_summary <- tibble(
  Threshold = c(
    "Primary burden threshold",
    "Primary severity threshold",
    "45th-percentile burden threshold",
    "45th-percentile severity threshold",
    "55th-percentile burden threshold",
    "55th-percentile severity threshold",
    "Bootstrap burden threshold median",
    "Bootstrap burden threshold lower 95% limit",
    "Bootstrap burden threshold upper 95% limit",
    "Bootstrap severity threshold median",
    "Bootstrap severity threshold lower 95% limit",
    "Bootstrap severity threshold upper 95% limit"
  ),
  
  Value = c(
    primary_burden_threshold,
    primary_severity_threshold,
    burden_threshold_45,
    severity_threshold_45,
    burden_threshold_55,
    severity_threshold_55,
    median(
      bootstrap_burden_thresholds
    ),
    as.numeric(
      quantile(
        bootstrap_burden_thresholds,
        probs = 0.025,
        names = FALSE
      )
    ),
    as.numeric(
      quantile(
        bootstrap_burden_thresholds,
        probs = 0.975,
        names = FALSE
      )
    ),
    median(
      bootstrap_severity_thresholds
    ),
    as.numeric(
      quantile(
        bootstrap_severity_thresholds,
        probs = 0.025,
        names = FALSE
      )
    ),
    as.numeric(
      quantile(
        bootstrap_severity_thresholds,
        probs = 0.975,
        names = FALSE
      )
    )
  )
)

s9_workbook <- openxlsx::createWorkbook()

s9_sheets <- list(
  "Country stability" =
    supplementary_s9,
  
  "Sensitivity summary" =
    sensitivity_summary,
  
  "Bootstrap summary" =
    bootstrap_stability_summary,
  
  "Quadrant stability" =
    quadrant_stability_summary,
  
  "Threshold summary" =
    bootstrap_threshold_summary
)

for (current_sheet in names(s9_sheets)) {
  
  current_data <- s9_sheets[[current_sheet]]
  
  openxlsx::addWorksheet(
    s9_workbook,
    current_sheet
  )
  
  openxlsx::writeData(
    s9_workbook,
    sheet = current_sheet,
    x = current_data
  )
  
  openxlsx::addStyle(
    s9_workbook,
    sheet = current_sheet,
    style = header_style,
    rows = 1,
    cols = seq_len(ncol(current_data)),
    gridExpand = TRUE
  )
  
  openxlsx::setColWidths(
    s9_workbook,
    sheet = current_sheet,
    cols = seq_len(ncol(current_data)),
    widths = "auto"
  )
  
  openxlsx::freezePane(
    s9_workbook,
    sheet = current_sheet,
    firstRow = TRUE
  )
  
  openxlsx::addFilter(
    s9_workbook,
    sheet = current_sheet,
    rows = 1,
    cols = seq_len(ncol(current_data))
  )
}

openxlsx::saveWorkbook(
  s9_workbook,
  s9_excel_file,
  overwrite = TRUE
)

readr::write_csv(
  supplementary_s9,
  s9_csv_file
)


# ==============================================================================
# 31. EXPORT FIGURE SOURCE DATA
# ==============================================================================

figure_9_source_data <- framework_data |>
  transmute(
    Country = country,
    DALY_rate = daly_rate,
    DALYs_per_case = dalys_per_case,
    
    Conventional_priority =
      as.character(
        conventional_priority
      ),
    
    Severity_aware_priority =
      as.character(
        severity_aware_priority
      ),
    
    Burden_severity_quadrant =
      reconstructed_primary_quadrant_code,
    
    Reclassification_direction =
      reclassification_direction,
    
    Priority_levels_moved =
      priority_levels_moved
  )

figure_9_source_file <- file.path(
  supplementary_folder,
  "Figure_9_Source_data.csv"
)

readr::write_csv(
  figure_9_source_data,
  figure_9_source_file
)

figure_s1_source_file <- file.path(
  supplementary_folder,
  "Figure_S1_Source_data.xlsx"
)

figure_s1_source_workbook <- openxlsx::createWorkbook()

figure_s1_sheets <- list(
  "Threshold sensitivity" =
    sensitivity_summary,
  
  "Country bootstrap stability" =
    supplementary_s9,
  
  "Quadrant stability" =
    quadrant_stability_summary,
  
  "Lowest stability countries" =
    lowest_stability_countries,
  
  "Threshold distributions" =
    bootstrap_threshold_summary
)

for (current_sheet in names(figure_s1_sheets)) {
  
  openxlsx::addWorksheet(
    figure_s1_source_workbook,
    current_sheet
  )
  
  openxlsx::writeData(
    figure_s1_source_workbook,
    sheet = current_sheet,
    x = figure_s1_sheets[[current_sheet]]
  )
}

openxlsx::saveWorkbook(
  figure_s1_source_workbook,
  figure_s1_source_file,
  overwrite = TRUE
)


# ==============================================================================
# 32. DISPLAY KEY RESULTS
# ==============================================================================

message("")
message("==============================================================")
message("RECLASSIFICATION SUMMARY")
message("==============================================================")

print(
  reclassification_summary,
  n = Inf
)

message("")
message("Priority transition matrix:")

print(
  priority_transition_matrix
)

message("")
message("==============================================================")
message("AGREEMENT RESULTS")
message("==============================================================")

message(
  "Exact agreement: ",
  format_percent(
    100 *
      unweighted_kappa$exact_agreement,
    1
  )
)

message(
  "Expected agreement: ",
  format_percent(
    100 *
      unweighted_kappa$expected_agreement,
    1
  )
)

message(
  "Unweighted Cohen's kappa: ",
  format_number(
    unweighted_kappa$kappa,
    3
  ),
  " (95% CI ",
  format_number(
    unweighted_kappa_ci[["lower"]],
    3
  ),
  " to ",
  format_number(
    unweighted_kappa_ci[["upper"]],
    3
  ),
  ")"
)

message(
  "Interpretation: ",
  interpret_kappa(
    unweighted_kappa$kappa
  )
)

message(
  "Linear weighted Cohen's kappa: ",
  format_number(
    linear_weighted_kappa$kappa,
    3
  ),
  " (95% CI ",
  format_number(
    weighted_kappa_ci[["lower"]],
    3
  ),
  " to ",
  format_number(
    weighted_kappa_ci[["upper"]],
    3
  ),
  ")"
)

message("")
message("==============================================================")
message("THRESHOLD SENSITIVITY")
message("==============================================================")

print(
  sensitivity_summary,
  n = Inf
)

message("")
message("==============================================================")
message("BOOTSTRAP STABILITY")
message("==============================================================")

print(
  bootstrap_stability_summary,
  n = Inf
)

message("")
message("Quadrant-specific stability:")

print(
  quadrant_stability_summary,
  n = Inf
)


# ==============================================================================
# 33. VALIDATE OUTPUT FILES
# ==============================================================================

required_output_files <- c(
  figure_9_file,
  table_9_file,
  figure_s1_file,
  s8_excel_file,
  s8_csv_file,
  s9_excel_file,
  s9_csv_file,
  figure_9_source_file,
  figure_s1_source_file
)

missing_output_files <- required_output_files[
  !file.exists(required_output_files)
]

if (length(missing_output_files) > 0) {
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

output_information <- file.info(
  required_output_files
)

empty_output_files <- rownames(output_information)[
  is.na(output_information$size) |
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
# 34. COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("SECTION 3.9 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message("Correct input used:")
message(input_file)

message("")
message(
  "Countries and territories analysed: ",
  expected_country_count
)

message(
  "Countries reclassified: ",
  total_reclassified,
  " (",
  format_number(
    percentage_reclassified,
    1
  ),
  "%)"
)

message(
  "Upward reclassification: ",
  upward_countries
)

message(
  "Downward reclassification: ",
  downward_countries
)

message(
  "Unchanged classifications: ",
  unchanged_countries
)

message("")
message("Main Figure 9:")
message(figure_9_file)

message("")
message("Main Table 9:")
message(table_9_file)

message("")
message("Supplementary Figure S1:")
message(figure_s1_file)

message("")
message("Supplementary Table S8:")
message(s8_excel_file)

message("")
message("Supplementary Table S9:")
message(s9_excel_file)

message("")
message("Table S2 was not used.")
message("No YLL or YLD data were used.")
message("All outputs were saved under Publication.")
message("==============================================================")