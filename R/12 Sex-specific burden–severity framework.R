# ==============================================================================
# R_03_11_Sex_specific_burden_severity_framework_2023.R
#
# RESULTS SECTION 3.11
# Sex-specific colorectal cancer burden–severity framework in 2023
#
# REQUIRED INPUT FILES
#
# Male:
#   02_Incidence_Number_AllCountries_Males_AllAges_1990_2023*.csv
#   14_DALYs_Number_AllCountries_Males_AllAges_1990_2023*.csv
#   17_DALYs_Rate_AllCountries_Males_AgeStandarized_1990_2023*.csv
#
# Female:
#   03_Incidence_Number_AllCountries_Females_AllAges_1990_2023*.csv
#   15_DALYs_Number_AllCountries_Females_AllAges_1990_2023*.csv
#   18_DALYs_Rate_AllCountries_Females_AgeStandarized_1990_2023*.csv
#
# ANALYSES
#   1. Calculate sex-specific DALYs per incident case in 2023.
#   2. Calculate sex-specific median thresholds for:
#        - age-standardized DALY rate;
#        - DALYs per incident case.
#   3. Assign male and female framework quadrants.
#   4. Compare male and female quadrant distributions.
#   5. Calculate exact agreement, unweighted Cohen's kappa, and linear
#      weighted Cohen's kappa with bootstrap 95% confidence intervals.
#   6. Identify countries with concordant and discordant sex-specific
#      classifications.
#   7. Quantify male–female differences in burden and severity.
#
# OUTPUTS
#
# Publication/Figures/
#   Figure_11_Sex_specific_burden_severity_framework_2023.tiff
#
# Publication/Tables/
#   Table_11_Sex_specific_framework_summary.xlsx
#
# Publication/Supplementary/
#   Table_S11_Complete_sex_specific_framework_2023.xlsx
#   Table_S11_Complete_sex_specific_framework_2023.csv
#   Figure_11_Source_data.xlsx
#
# IMPORTANT
#   - Only 2023 observations are analysed.
#   - Male and female frameworks use their own sex-specific median thresholds.
#   - The male and female scatterplots use identical axis ranges.
#   - No YLL or YLD data are used.
#   - All outputs are saved under Publication.
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
  "openxlsx",
  "janitor",
  "ggalluvial",
  "patchwork",
  "scales",
  "ggrepel"
)

installed_packages <- rownames(installed.packages())
missing_packages <- setdiff(required_packages, installed_packages)

if (length(missing_packages) > 0) {
  install.packages(
    missing_packages,
    dependencies = TRUE
  )
}

suppressPackageStartupMessages({
  library(tidyverse)
  library(openxlsx)
  library(janitor)
  library(ggalluvial)
  library(patchwork)
  library(scales)
  library(ggrepel)
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
# 5. LOCATE THE SIX REQUIRED INPUT FILES
# ==============================================================================

find_single_input <- function(pattern, label) {
  
  candidates <- list.files(
    path = project_root,
    pattern = pattern,
    full.names = TRUE,
    recursive = TRUE,
    ignore.case = TRUE
  )
  
  candidates <- candidates[
    !stringr::str_detect(
      basename(candidates),
      "^~\\$"
    )
  ]
  
  if (length(candidates) == 0) {
    stop(
      paste0(
        "\nRequired input not found: ",
        label,
        "\n\nExpected filename pattern:\n",
        pattern
      ),
      call. = FALSE
    )
  }
  
  # If multiple copies exist, select the most recently modified one.
  candidate_info <- file.info(candidates)
  
  selected_file <- candidates[
    order(
      candidate_info$mtime,
      decreasing = TRUE
    )
  ][1]
  
  message("")
  message(label, ":")
  message(selected_file)
  
  selected_file
}


male_incidence_file <- find_single_input(
  pattern =
    "^02_Incidence_Number_AllCountries_Males_AllAges_1990_2023.*\\.csv$",
  label =
    "Male incidence-number file"
)

female_incidence_file <- find_single_input(
  pattern =
    "^03_Incidence_Number_AllCountries_Females_AllAges_1990_2023.*\\.csv$",
  label =
    "Female incidence-number file"
)

male_dalys_number_file <- find_single_input(
  pattern =
    "^14_DALYs_Number_AllCountries_Males_AllAges_1990_2023.*\\.csv$",
  label =
    "Male DALY-number file"
)

female_dalys_number_file <- find_single_input(
  pattern =
    "^15_DALYs_Number_AllCountries_Females_AllAges_1990_2023.*\\.csv$",
  label =
    "Female DALY-number file"
)

male_dalys_rate_file <- find_single_input(
  pattern =
    "^17_DALYs_Rate_AllCountries_Males_AgeStandarized_1990_2023.*\\.csv$",
  label =
    "Male age-standardized DALY-rate file"
)

female_dalys_rate_file <- find_single_input(
  pattern =
    "^18_DALYs_Rate_AllCountries_Females_AgeStandarized_1990_2023.*\\.csv$",
  label =
    "Female age-standardized DALY-rate file"
)


# ==============================================================================
# 6. ANALYTICAL SETTINGS
# ==============================================================================

analysis_year <- 2023L
expected_country_count <- 204L

kappa_bootstrap_iterations <- 2000L
bootstrap_seed <- 20260801L

top_difference_labels <- 8L

quadrant_levels <- c(
  "Q1: High burden–high severity",
  "Q2: Low burden–high severity",
  "Q3: Low burden–low severity",
  "Q4: High burden–low severity"
)

quadrant_code_levels <- c(
  "Q1",
  "Q2",
  "Q3",
  "Q4"
)

priority_levels <- c(
  "Priority 1",
  "Priority 2",
  "Priority 3",
  "Priority 4"
)

quadrant_colours <- c(
  "Q1: High burden–high severity" = "#B2182B",
  "Q2: Low burden–high severity" = "#EF8A62",
  "Q3: Low burden–low severity" = "#67A9CF",
  "Q4: High burden–low severity" = "#2166AC"
)

agreement_colours <- c(
  "Same quadrant" = "#737373",
  "Different quadrant" = "#B2182B"
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


extract_quadrant_code <- function(x) {
  stringr::str_extract(
    as.character(x),
    "Q[1-4]"
  )
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


priority_number_to_label <- function(priority_number) {
  dplyr::case_when(
    priority_number == 1L ~ "Priority 1",
    priority_number == 2L ~ "Priority 2",
    priority_number == 3L ~ "Priority 3",
    priority_number == 4L ~ "Priority 4",
    TRUE ~ NA_character_
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
    
    positions <- seq_len(number_categories)
    
    weight_matrix <- 1 -
      abs(
        outer(
          positions,
          positions,
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
    observed_agreement -
      expected_agreement
  ) / (
    1 -
      expected_agreement
  )
  
  exact_agreement <- sum(
    diag(
      observed_proportions
    )
  )
  
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
          reference =
            sampled_data[[reference_column]],
          
          comparison =
            sampled_data[[comparison_column]],
          
          category_levels =
            category_levels,
          
          weighted =
            weighted
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
    lower =
      as.numeric(
        quantile(
          bootstrap_values,
          probs = 0.025,
          na.rm = TRUE,
          names = FALSE
        )
      ),
    
    upper =
      as.numeric(
        quantile(
          bootstrap_values,
          probs = 0.975,
          na.rm = TRUE,
          names = FALSE
        )
      )
  )
}


read_2023_file <- function(
    file_path,
    expected_sex,
    expected_age,
    expected_metric,
    value_name
) {
  
  data <- readr::read_csv(
    file_path,
    show_col_types = FALSE,
    progress = FALSE,
    guess_max = 10000,
    locale = readr::locale(
      encoding = "UTF-8"
    )
  ) |>
    janitor::clean_names()
  
  required_columns <- c(
    "measure",
    "location",
    "sex",
    "age",
    "metric",
    "year",
    "val",
    "upper",
    "lower"
  )
  
  missing_columns <- setdiff(
    required_columns,
    names(data)
  )
  
  if (length(missing_columns) > 0) {
    stop(
      paste0(
        "\nMissing columns in:\n",
        file_path,
        "\n\n",
        paste(
          missing_columns,
          collapse = "\n"
        )
      ),
      call. = FALSE
    )
  }
  
  data <- data |>
    transmute(
      country =
        clean_text(
          location
        ),
      
      sex =
        clean_text(
          sex
        ),
      
      age =
        clean_text(
          age
        ),
      
      metric =
        clean_text(
          metric
        ),
      
      year =
        as.integer(
          year
        ),
      
      "{value_name}" :=
        as.numeric(
          val
        ),
      
      "{paste0(value_name, '_lower')}" :=
        as.numeric(
          lower
        ),
      
      "{paste0(value_name, '_upper')}" :=
        as.numeric(
          upper
        )
    ) |>
    filter(
      sex == expected_sex,
      age == expected_age,
      metric == expected_metric,
      year == analysis_year
    )
  
  if (nrow(data) != expected_country_count) {
    stop(
      paste0(
        "\nExpected ",
        expected_country_count,
        " valid 2023 rows in:\n",
        file_path,
        "\nRows found: ",
        nrow(data)
      ),
      call. = FALSE
    )
  }
  
  duplicate_countries <- data |>
    count(
      country,
      name = "records"
    ) |>
    filter(
      records > 1
    )
  
  if (nrow(duplicate_countries) > 0) {
    print(
      duplicate_countries,
      n = Inf
    )
    
    stop(
      paste0(
        "\nDuplicate 2023 country records were detected in:\n",
        file_path
      ),
      call. = FALSE
    )
  }
  
  data
}


# ==============================================================================
# 8. IMPORT THE SIX DATASETS
# ==============================================================================

male_incidence <- read_2023_file(
  file_path =
    male_incidence_file,
  
  expected_sex =
    "Male",
  
  expected_age =
    "All ages",
  
  expected_metric =
    "Number",
  
  value_name =
    "male_incident_cases"
)

female_incidence <- read_2023_file(
  file_path =
    female_incidence_file,
  
  expected_sex =
    "Female",
  
  expected_age =
    "All ages",
  
  expected_metric =
    "Number",
  
  value_name =
    "female_incident_cases"
)

male_dalys_number <- read_2023_file(
  file_path =
    male_dalys_number_file,
  
  expected_sex =
    "Male",
  
  expected_age =
    "All ages",
  
  expected_metric =
    "Number",
  
  value_name =
    "male_dalys"
)

female_dalys_number <- read_2023_file(
  file_path =
    female_dalys_number_file,
  
  expected_sex =
    "Female",
  
  expected_age =
    "All ages",
  
  expected_metric =
    "Number",
  
  value_name =
    "female_dalys"
)

male_dalys_rate <- read_2023_file(
  file_path =
    male_dalys_rate_file,
  
  expected_sex =
    "Male",
  
  expected_age =
    "Age-standardized",
  
  expected_metric =
    "Rate",
  
  value_name =
    "male_daly_rate"
)

female_dalys_rate <- read_2023_file(
  file_path =
    female_dalys_rate_file,
  
  expected_sex =
    "Female",
  
  expected_age =
    "Age-standardized",
  
  expected_metric =
    "Rate",
  
  value_name =
    "female_daly_rate"
)


# ==============================================================================
# 9. MERGE MALE AND FEMALE DATA
# ==============================================================================

male_framework <- male_incidence |>
  select(
    country,
    male_incident_cases,
    male_incident_cases_lower,
    male_incident_cases_upper
  ) |>
  inner_join(
    male_dalys_number |>
      select(
        country,
        male_dalys,
        male_dalys_lower,
        male_dalys_upper
      ),
    by = "country"
  ) |>
  inner_join(
    male_dalys_rate |>
      select(
        country,
        male_daly_rate,
        male_daly_rate_lower,
        male_daly_rate_upper
      ),
    by = "country"
  ) |>
  mutate(
    male_dalys_per_case =
      male_dalys /
      male_incident_cases
  )


female_framework <- female_incidence |>
  select(
    country,
    female_incident_cases,
    female_incident_cases_lower,
    female_incident_cases_upper
  ) |>
  inner_join(
    female_dalys_number |>
      select(
        country,
        female_dalys,
        female_dalys_lower,
        female_dalys_upper
      ),
    by = "country"
  ) |>
  inner_join(
    female_dalys_rate |>
      select(
        country,
        female_daly_rate,
        female_daly_rate_lower,
        female_daly_rate_upper
      ),
    by = "country"
  ) |>
  mutate(
    female_dalys_per_case =
      female_dalys /
      female_incident_cases
  )


sex_framework <- male_framework |>
  inner_join(
    female_framework,
    by = "country"
  ) |>
  arrange(
    country
  )


# ==============================================================================
# 10. QUALITY CONTROL
# ==============================================================================

if (nrow(sex_framework) != expected_country_count) {
  stop(
    paste0(
      "\nExpected ",
      expected_country_count,
      " matched countries, but ",
      nrow(sex_framework),
      " were available after merging."
    ),
    call. = FALSE
  )
}

invalid_values <- sex_framework |>
  filter(
    is.na(male_incident_cases) |
      !is.finite(male_incident_cases) |
      male_incident_cases <= 0 |
      
      is.na(female_incident_cases) |
      !is.finite(female_incident_cases) |
      female_incident_cases <= 0 |
      
      is.na(male_dalys) |
      !is.finite(male_dalys) |
      male_dalys < 0 |
      
      is.na(female_dalys) |
      !is.finite(female_dalys) |
      female_dalys < 0 |
      
      is.na(male_daly_rate) |
      !is.finite(male_daly_rate) |
      male_daly_rate < 0 |
      
      is.na(female_daly_rate) |
      !is.finite(female_daly_rate) |
      female_daly_rate < 0 |
      
      is.na(male_dalys_per_case) |
      !is.finite(male_dalys_per_case) |
      male_dalys_per_case < 0 |
      
      is.na(female_dalys_per_case) |
      !is.finite(female_dalys_per_case) |
      female_dalys_per_case < 0
  )

if (nrow(invalid_values) > 0) {
  print(
    invalid_values,
    n = Inf
  )
  
  stop(
    "\nInvalid male or female burden–severity values were detected.",
    call. = FALSE
  )
}


# ==============================================================================
# 11. CALCULATE SEX-SPECIFIC MEDIAN THRESHOLDS
# ==============================================================================

male_burden_threshold <- median(
  sex_framework$male_daly_rate,
  na.rm = TRUE
)

male_severity_threshold <- median(
  sex_framework$male_dalys_per_case,
  na.rm = TRUE
)

female_burden_threshold <- median(
  sex_framework$female_daly_rate,
  na.rm = TRUE
)

female_severity_threshold <- median(
  sex_framework$female_dalys_per_case,
  na.rm = TRUE
)

message("")
message(
  "Male DALY-rate threshold: ",
  format_number(
    male_burden_threshold,
    2
  )
)

message(
  "Male DALYs-per-case threshold: ",
  format_number(
    male_severity_threshold,
    2
  )
)

message(
  "Female DALY-rate threshold: ",
  format_number(
    female_burden_threshold,
    2
  )
)

message(
  "Female DALYs-per-case threshold: ",
  format_number(
    female_severity_threshold,
    2
  )
)


# ==============================================================================
# 12. ASSIGN MALE AND FEMALE QUADRANTS
# ==============================================================================

sex_framework <- sex_framework |>
  mutate(
    male_quadrant =
      create_quadrant(
        daly_rate =
          male_daly_rate,
        
        dalys_per_case =
          male_dalys_per_case,
        
        burden_threshold =
          male_burden_threshold,
        
        severity_threshold =
          male_severity_threshold
      ),
    
    female_quadrant =
      create_quadrant(
        daly_rate =
          female_daly_rate,
        
        dalys_per_case =
          female_dalys_per_case,
        
        burden_threshold =
          female_burden_threshold,
        
        severity_threshold =
          female_severity_threshold
      ),
    
    male_quadrant_code =
      extract_quadrant_code(
        male_quadrant
      ),
    
    female_quadrant_code =
      extract_quadrant_code(
        female_quadrant
      ),
    
    male_priority_number =
      quadrant_to_priority_number(
        male_quadrant_code
      ),
    
    female_priority_number =
      quadrant_to_priority_number(
        female_quadrant_code
      ),
    
    male_priority =
      priority_number_to_label(
        male_priority_number
      ),
    
    female_priority =
      priority_number_to_label(
        female_priority_number
      ),
    
    same_quadrant =
      male_quadrant_code ==
      female_quadrant_code,
    
    quadrant_agreement =
      if_else(
        same_quadrant,
        "Same quadrant",
        "Different quadrant"
      ),
    
    priority_difference_female_minus_male =
      female_priority_number -
      male_priority_number,
    
    sex_specific_priority_direction =
      case_when(
        female_priority_number <
          male_priority_number ~
          "Higher priority among females",
        
        female_priority_number >
          male_priority_number ~
          "Higher priority among males",
        
        TRUE ~
          "Same priority"
      ),
    
    daly_rate_difference_male_minus_female =
      male_daly_rate -
      female_daly_rate,
    
    daly_rate_ratio_male_to_female =
      male_daly_rate /
      female_daly_rate,
    
    daly_rate_percent_difference =
      100 *
      (
        male_daly_rate -
          female_daly_rate
      ) /
      female_daly_rate,
    
    dalys_per_case_difference_male_minus_female =
      male_dalys_per_case -
      female_dalys_per_case,
    
    dalys_per_case_ratio_male_to_female =
      male_dalys_per_case /
      female_dalys_per_case,
    
    dalys_per_case_percent_difference =
      100 *
      (
        male_dalys_per_case -
          female_dalys_per_case
      ) /
      female_dalys_per_case
  )


# ==============================================================================
# 13. QUADRANT DISTRIBUTIONS
# ==============================================================================

male_quadrant_distribution <- sex_framework |>
  count(
    quadrant =
      male_quadrant,
    name =
      "Countries"
  ) |>
  complete(
    quadrant =
      quadrant_levels,
    fill =
      list(
        Countries = 0
      )
  ) |>
  mutate(
    Sex =
      "Male",
    
    Percentage =
      100 *
      Countries /
      expected_country_count
  )

female_quadrant_distribution <- sex_framework |>
  count(
    quadrant =
      female_quadrant,
    name =
      "Countries"
  ) |>
  complete(
    quadrant =
      quadrant_levels,
    fill =
      list(
        Countries = 0
      )
  ) |>
  mutate(
    Sex =
      "Female",
    
    Percentage =
      100 *
      Countries /
      expected_country_count
  )

quadrant_distribution <- bind_rows(
  male_quadrant_distribution,
  female_quadrant_distribution
) |>
  mutate(
    quadrant =
      factor(
        quadrant,
        levels =
          quadrant_levels
      )
  ) |>
  arrange(
    Sex,
    quadrant
  )


# ==============================================================================
# 14. AGREEMENT AND TRANSITION MATRIX
# ==============================================================================

quadrant_transition_matrix <- table(
  Male =
    factor(
      sex_framework$male_quadrant_code,
      levels =
        quadrant_code_levels
    ),
  
  Female =
    factor(
      sex_framework$female_quadrant_code,
      levels =
        quadrant_code_levels
    )
)

quadrant_transition_long <- as.data.frame(
  quadrant_transition_matrix
) |>
  rename(
    Male_quadrant =
      Male,
    
    Female_quadrant =
      Female,
    
    Countries =
      Freq
  ) |>
  mutate(
    Percentage =
      100 *
      Countries /
      expected_country_count
  )

same_quadrant_count <- sum(
  sex_framework$same_quadrant
)

different_quadrant_count <- sum(
  !sex_framework$same_quadrant
)

exact_agreement_percentage <- 100 *
  same_quadrant_count /
  expected_country_count


# ==============================================================================
# 15. COHEN'S KAPPA
# ==============================================================================

unweighted_kappa <- calculate_kappa(
  reference =
    sex_framework$male_quadrant_code,
  
  comparison =
    sex_framework$female_quadrant_code,
  
  category_levels =
    quadrant_code_levels,
  
  weighted =
    FALSE
)

sex_framework <- sex_framework |>
  mutate(
    male_priority =
      factor(
        male_priority,
        levels =
          priority_levels,
        ordered =
          TRUE
      ),
    
    female_priority =
      factor(
        female_priority,
        levels =
          priority_levels,
        ordered =
          TRUE
      )
  )

weighted_kappa <- calculate_kappa(
  reference =
    sex_framework$male_priority,
  
  comparison =
    sex_framework$female_priority,
  
  category_levels =
    priority_levels,
  
  weighted =
    TRUE
)

unweighted_kappa_ci <- bootstrap_kappa_ci(
  data =
    sex_framework,
  
  reference_column =
    "male_quadrant_code",
  
  comparison_column =
    "female_quadrant_code",
  
  category_levels =
    quadrant_code_levels,
  
  weighted =
    FALSE,
  
  iterations =
    kappa_bootstrap_iterations,
  
  seed =
    bootstrap_seed
)

weighted_kappa_ci <- bootstrap_kappa_ci(
  data =
    sex_framework,
  
  reference_column =
    "male_priority",
  
  comparison_column =
    "female_priority",
  
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
# 16. IDENTIFY SEX-SPECIFIC Q1 COUNTRIES
# ==============================================================================

q1_comparison <- sex_framework |>
  mutate(
    q1_status =
      case_when(
        male_quadrant_code == "Q1" &
          female_quadrant_code == "Q1" ~
          "Q1 in both sexes",
        
        male_quadrant_code == "Q1" &
          female_quadrant_code != "Q1" ~
          "Q1 among males only",
        
        male_quadrant_code != "Q1" &
          female_quadrant_code == "Q1" ~
          "Q1 among females only",
        
        TRUE ~
          "Not Q1 in either sex"
      )
  )

q1_summary <- q1_comparison |>
  count(
    q1_status,
    name =
      "Countries"
  ) |>
  mutate(
    Percentage =
      100 *
      Countries /
      expected_country_count
  )


# ==============================================================================
# 17. IDENTIFY LARGEST SEX DIFFERENCES
# ==============================================================================

sex_framework <- sex_framework |>
  mutate(
    standardized_daly_rate_difference =
      daly_rate_difference_male_minus_female /
      sd(
        daly_rate_difference_male_minus_female,
        na.rm = TRUE
      ),
    
    standardized_severity_difference =
      dalys_per_case_difference_male_minus_female /
      sd(
        dalys_per_case_difference_male_minus_female,
        na.rm = TRUE
      ),
    
    combined_absolute_sex_difference =
      sqrt(
        standardized_daly_rate_difference^2 +
          standardized_severity_difference^2
      )
  )

largest_sex_differences <- sex_framework |>
  arrange(
    desc(
      combined_absolute_sex_difference
    )
  ) |>
  slice_head(
    n =
      top_difference_labels
  )

largest_male_excess_burden <- sex_framework |>
  arrange(
    desc(
      daly_rate_difference_male_minus_female
    )
  ) |>
  slice_head(
    n = 10
  )

largest_female_excess_burden <- sex_framework |>
  arrange(
    daly_rate_difference_male_minus_female
  ) |>
  slice_head(
    n = 10
  )

largest_male_excess_severity <- sex_framework |>
  arrange(
    desc(
      dalys_per_case_difference_male_minus_female
    )
  ) |>
  slice_head(
    n = 10
  )

largest_female_excess_severity <- sex_framework |>
  arrange(
    dalys_per_case_difference_male_minus_female
  ) |>
  slice_head(
    n = 10
  )


# ==============================================================================
# 18. CREATE TABLE 11 SUMMARY
# ==============================================================================

agreement_summary <- tibble(
  Measure = c(
    "Countries and territories analysed",
    "Countries in the same male and female quadrant",
    "Countries in different male and female quadrants",
    "Exact quadrant agreement (%)",
    "Expected agreement (%)",
    "Unweighted Cohen's kappa",
    "Unweighted kappa lower 95% CI",
    "Unweighted kappa upper 95% CI",
    "Linear weighted Cohen's kappa",
    "Weighted kappa lower 95% CI",
    "Weighted kappa upper 95% CI",
    "Countries in Q1 for both sexes",
    "Countries in Q1 among males only",
    "Countries in Q1 among females only",
    "Male median DALY-rate threshold",
    "Male median DALYs-per-case threshold",
    "Female median DALY-rate threshold",
    "Female median DALYs-per-case threshold"
  ),
  
  Value = c(
    expected_country_count,
    same_quadrant_count,
    different_quadrant_count,
    exact_agreement_percentage,
    100 *
      unweighted_kappa$expected_agreement,
    unweighted_kappa$kappa,
    unweighted_kappa_ci[["lower"]],
    unweighted_kappa_ci[["upper"]],
    weighted_kappa$kappa,
    weighted_kappa_ci[["lower"]],
    weighted_kappa_ci[["upper"]],
    sum(
      q1_comparison$q1_status ==
        "Q1 in both sexes"
    ),
    sum(
      q1_comparison$q1_status ==
        "Q1 among males only"
    ),
    sum(
      q1_comparison$q1_status ==
        "Q1 among females only"
    ),
    male_burden_threshold,
    male_severity_threshold,
    female_burden_threshold,
    female_severity_threshold
  )
)

agreement_summary_formatted <- agreement_summary |>
  mutate(
    `Reported value` =
      case_when(
        Measure %in% c(
          "Exact quadrant agreement (%)",
          "Expected agreement (%)"
        ) ~
          format_percent(
            Value,
            1
          ),
        
        stringr::str_detect(
          stringr::str_to_lower(
            Measure
          ),
          "kappa|threshold"
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
# 19. COMMON AXIS LIMITS FOR MALE AND FEMALE FRAMEWORK PLOTS
# ==============================================================================

common_x_max <- max(
  c(
    sex_framework$male_daly_rate,
    sex_framework$female_daly_rate
  ),
  na.rm = TRUE
)

common_y_max <- max(
  c(
    sex_framework$male_dalys_per_case,
    sex_framework$female_dalys_per_case
  ),
  na.rm = TRUE
)

common_x_limit <- c(
  0,
  common_x_max * 1.08
)

common_y_limit <- c(
  0,
  common_y_max * 1.08
)


# ==============================================================================
# 20. FRAMEWORK SCATTERPLOT FUNCTION
# ==============================================================================

create_framework_plot <- function(
    data,
    daly_rate_column,
    severity_column,
    quadrant_column,
    burden_threshold,
    severity_threshold,
    panel_title
) {
  
  plot_data <- data |>
    transmute(
      country =
        country,
      
      daly_rate =
        .data[[daly_rate_column]],
      
      dalys_per_case =
        .data[[severity_column]],
      
      quadrant =
        factor(
          .data[[quadrant_column]],
          levels =
            quadrant_levels
        )
    )
  
  label_data <- plot_data |>
    mutate(
      distance_from_centre =
        sqrt(
          (
            (
              daly_rate -
                burden_threshold
            ) /
              burden_threshold
          )^2 +
            (
              (
                dalys_per_case -
                  severity_threshold
              ) /
                severity_threshold
            )^2
        )
    ) |>
    group_by(
      quadrant
    ) |>
    slice_max(
      order_by =
        distance_from_centre,
      n =
        2,
      with_ties =
        FALSE
    ) |>
    ungroup()
  
  ggplot(
    plot_data,
    aes(
      x =
        daly_rate,
      y =
        dalys_per_case,
      fill =
        quadrant
    )
  ) +
    
    geom_point(
      shape = 21,
      size = 2.5,
      alpha = 0.82,
      color = "grey20",
      stroke = 0.25
    ) +
    
    geom_vline(
      xintercept =
        burden_threshold,
      linetype =
        "dashed",
      linewidth =
        0.65,
      color =
        "grey25"
    ) +
    
    geom_hline(
      yintercept =
        severity_threshold,
      linetype =
        "dashed",
      linewidth =
        0.65,
      color =
        "grey25"
    ) +
    
    ggrepel::geom_text_repel(
      data =
        label_data,
      aes(
        label =
          country
      ),
      size = 3.0,
      fontface = "bold",
      max.overlaps = Inf,
      box.padding = 0.65,
      point.padding = 0.45,
      force = 2.0,
      force_pull = 0.3,
      min.segment.length = 0,
      segment.size = 0.35,
      segment.color = "grey45",
      seed = 20260801,
      direction = "both",
      show.legend = FALSE
    ) +
    
    annotate(
      "text",
      x =
        burden_threshold * 0.05,
      y =
        severity_threshold * 1.05,
      label =
        paste0(
          "Severity threshold = ",
          formatC(
            severity_threshold,
            format = "f",
            digits = 2
          )
        ),
      hjust = 0,
      vjust = -0.5,
      size = 3.0
    ) +
    
    annotate(
      "text",
      x =
        burden_threshold * 1.03,
      y =
        common_y_limit[[1]] +
        0.02 *
        diff(
          common_y_limit
        ),
      label =
        paste0(
          "Burden threshold = ",
          formatC(
            burden_threshold,
            format = "f",
            digits = 2
          )
        ),
      hjust = 0,
      vjust = 0,
      size = 3.0
    ) +
    
    scale_fill_manual(
      values =
        quadrant_colours,
      drop =
        FALSE,
      name =
        "Quadrant"
    ) +
    
    coord_cartesian(
      xlim =
        common_x_limit,
      ylim =
        common_y_limit,
      clip =
        "off"
    ) +
    
    labs(
      title =
        panel_title,
      
      x =
        "Age-standardized DALY rate per 100,000",
      
      y =
        "DALYs per incident case"
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
          size = 7.8
        ),
      
      plot.margin =
        margin(
          5,
          8,
          5,
          5
        )
    )
}


# ==============================================================================
# 21. CREATE FIGURE 11 PANELS A AND B
# ==============================================================================

figure_11a <- create_framework_plot(
  data =
    sex_framework,
  
  daly_rate_column =
    "male_daly_rate",
  
  severity_column =
    "male_dalys_per_case",
  
  quadrant_column =
    "male_quadrant",
  
  burden_threshold =
    male_burden_threshold,
  
  severity_threshold =
    male_severity_threshold,
  
  panel_title =
    "A. Male burden–severity framework"
)

figure_11b <- create_framework_plot(
  data =
    sex_framework,
  
  daly_rate_column =
    "female_daly_rate",
  
  severity_column =
    "female_dalys_per_case",
  
  quadrant_column =
    "female_quadrant",
  
  burden_threshold =
    female_burden_threshold,
  
  severity_threshold =
    female_severity_threshold,
  
  panel_title =
    "B. Female burden–severity framework"
)


# ==============================================================================
# 22. CREATE FIGURE 11 PANEL C: MALE-TO-FEMALE TRANSITIONS
# ==============================================================================

figure_11c_data <- sex_framework |>
  transmute(
    country =
      country,
    
    male_quadrant =
      factor(
        male_quadrant,
        levels =
          quadrant_levels
      ),
    
    female_quadrant =
      factor(
        female_quadrant,
        levels =
          quadrant_levels
      ),
    
    quadrant_agreement =
      factor(
        quadrant_agreement,
        levels = c(
          "Same quadrant",
          "Different quadrant"
        )
      )
  )

figure_11c <- ggplot(
  figure_11c_data,
  aes(
    axis1 =
      male_quadrant,
    axis2 =
      female_quadrant,
    y =
      1
  )
) +
  
  ggalluvial::geom_alluvium(
    aes(
      fill =
        quadrant_agreement
    ),
    width = 0.16,
    alpha = 0.72,
    knot.pos = 0.40,
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
      label =
        after_stat(
          stringr::str_extract(
            stratum,
            "Q[1-4]"
          )
        )
    ),
    size = 4.0,
    fontface = "bold"
  ) +
  
  scale_x_discrete(
    limits = c(
      "Male",
      "Female"
    ),
    expand = c(
      0.16,
      0.16
    )
  ) +
  
  scale_fill_manual(
    values =
      agreement_colours,
    drop =
      FALSE,
    name =
      NULL
  ) +
  
  scale_y_continuous(
    name =
      "Countries and territories",
    breaks =
      scales::pretty_breaks(
        n = 5
      ),
    expand =
      expansion(
        mult = c(
          0,
          0.03
        )
      )
  ) +
  
  labs(
    title =
      "C. Male-to-female quadrant transitions",
    
    subtitle =
      paste0(
        same_quadrant_count,
        " countries (",
        formatC(
          exact_agreement_percentage,
          format = "f",
          digits = 1
        ),
        "%) retained the same quadrant"
      )
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
    
    plot.subtitle =
      element_text(
        size = 9.5
      ),
    
    axis.title.x =
      element_blank(),
    
    axis.text.x =
      element_text(
        face = "bold",
        size = 10
      ),
    
    axis.title.y =
      element_text(
        face = "bold"
      ),
    
    legend.position =
      "bottom"
  )


# ==============================================================================
# 23. CREATE FIGURE 11 PANEL D: SEX-DIFFERENCE PLOT
# ==============================================================================

figure_11d_label_data <- sex_framework |>
  semi_join(
    largest_sex_differences |>
      select(
        country
      ),
    by =
      "country"
  )

figure_11d <- ggplot(
  sex_framework,
  aes(
    x =
      daly_rate_difference_male_minus_female,
    
    y =
      dalys_per_case_difference_male_minus_female,
    
    fill =
      quadrant_agreement
  )
) +
  
  geom_hline(
    yintercept = 0,
    linetype = "dashed",
    linewidth = 0.65,
    color = "grey35"
  ) +
  
  geom_vline(
    xintercept = 0,
    linetype = "dashed",
    linewidth = 0.65,
    color = "grey35"
  ) +
  
  geom_point(
    shape = 21,
    size = 2.5,
    alpha = 0.80,
    color = "grey20",
    stroke = 0.25
  ) +
  
  ggrepel::geom_text_repel(
    data =
      figure_11d_label_data,
    
    aes(
      label =
        country
    ),
    
    size = 3.0,
    fontface = "bold",
    max.overlaps = Inf,
    box.padding = 0.70,
    point.padding = 0.50,
    force = 2.2,
    force_pull = 0.25,
    min.segment.length = 0,
    segment.size = 0.35,
    segment.color = "grey45",
    seed = 20260802,
    direction = "both",
    show.legend = FALSE
  ) +
  
  scale_fill_manual(
    values =
      agreement_colours,
    drop =
      FALSE,
    name =
      NULL
  ) +
  
  labs(
    title =
      "D. Male–female differences in burden and severity",
    
    subtitle =
      "Positive values indicate higher values among males",
    
    x =
      "DALY-rate difference (male − female)",
    
    y =
      "DALYs-per-case difference (male − female)"
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
    
    plot.subtitle =
      element_text(
        size = 9.5
      ),
    
    axis.title =
      element_text(
        face = "bold"
      ),
    
    legend.position =
      "bottom"
  )


# ==============================================================================
# 24. COMBINE FIGURE 11
# ==============================================================================

figure_11 <- (
  figure_11a +
    figure_11b
) / (
  figure_11c +
    figure_11d
) +
  
  patchwork::plot_layout(
    widths = c(
      1,
      1
    ),
    heights = c(
      1,
      0.90
    ),
    guides = "keep"
  ) +
  
  patchwork::plot_annotation(
    title =
      "Sex-specific colorectal cancer burden–severity framework in 2023",
    
    subtitle =
      paste0(
        "Male thresholds: DALY rate ",
        formatC(
          male_burden_threshold,
          format = "f",
          digits = 2
        ),
        " and DALYs per case ",
        formatC(
          male_severity_threshold,
          format = "f",
          digits = 2
        ),
        "; female thresholds: DALY rate ",
        formatC(
          female_burden_threshold,
          format = "f",
          digits = 2
        ),
        " and DALYs per case ",
        formatC(
          female_severity_threshold,
          format = "f",
          digits = 2
        )
      ),
    
    caption =
      paste0(
        "Male and female classifications used sex-specific median thresholds. ",
        "Panels A and B use identical axis ranges. Panel C shows movement ",
        "between male and female quadrants. Panel D displays absolute ",
        "male-minus-female differences; labelled countries have the largest ",
        "combined standardized sex differences."
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
          size = 10,
          hjust = 0.5,
          margin = margin(
            b = 8
          )
        ),
      
      plot.caption =
        element_text(
          size = 8.2,
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
# 25. SAVE FIGURE 11
# ==============================================================================

figure_11_file <- file.path(
  figures_folder,
  "Figure_11_Sex_specific_burden_severity_framework_2023.tiff"
)

ggsave(
  filename =
    figure_11_file,
  
  plot =
    figure_11,
  
  device =
    "tiff",
  
  width =
    16,
  
  height =
    12.8,
  
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
# 26. PREPARE SUPPLEMENTARY TABLE S11
# ==============================================================================

supplementary_s11 <- sex_framework |>
  transmute(
    Country =
      country,
    
    Male_incident_cases =
      male_incident_cases,
    
    Male_DALYs =
      male_dalys,
    
    Male_age_standardized_DALY_rate =
      male_daly_rate,
    
    Male_DALYs_per_case =
      male_dalys_per_case,
    
    Male_burden_threshold =
      male_burden_threshold,
    
    Male_severity_threshold =
      male_severity_threshold,
    
    Male_quadrant =
      male_quadrant_code,
    
    Male_priority =
      as.character(
        male_priority
      ),
    
    Female_incident_cases =
      female_incident_cases,
    
    Female_DALYs =
      female_dalys,
    
    Female_age_standardized_DALY_rate =
      female_daly_rate,
    
    Female_DALYs_per_case =
      female_dalys_per_case,
    
    Female_burden_threshold =
      female_burden_threshold,
    
    Female_severity_threshold =
      female_severity_threshold,
    
    Female_quadrant =
      female_quadrant_code,
    
    Female_priority =
      as.character(
        female_priority
      ),
    
    Same_quadrant =
      if_else(
        same_quadrant,
        "Yes",
        "No"
      ),
    
    Quadrant_agreement =
      quadrant_agreement,
    
    Sex_specific_priority_direction =
      sex_specific_priority_direction,
    
    DALY_rate_difference_male_minus_female =
      daly_rate_difference_male_minus_female,
    
    DALY_rate_ratio_male_to_female =
      daly_rate_ratio_male_to_female,
    
    DALY_rate_percent_difference =
      daly_rate_percent_difference,
    
    DALYs_per_case_difference_male_minus_female =
      dalys_per_case_difference_male_minus_female,
    
    DALYs_per_case_ratio_male_to_female =
      dalys_per_case_ratio_male_to_female,
    
    DALYs_per_case_percent_difference =
      dalys_per_case_percent_difference,
    
    Combined_absolute_sex_difference =
      combined_absolute_sex_difference
  ) |>
  arrange(
    desc(
      Combined_absolute_sex_difference
    ),
    Country
  )


# ==============================================================================
# 27. EXCEL STYLES
# ==============================================================================

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
# 28. EXPORT TABLE 11
# ==============================================================================

table_11_file <- file.path(
  tables_folder,
  "Table_11_Sex_specific_framework_summary.xlsx"
)

table_11_workbook <- openxlsx::createWorkbook()

table_11_sheets <- list(
  "Framework summary" =
    agreement_summary_formatted,
  
  "Numeric results" =
    agreement_summary,
  
  "Quadrant distribution" =
    quadrant_distribution,
  
  "Transition matrix" =
    quadrant_transition_long,
  
  "Q1 comparison" =
    q1_summary,
  
  "Largest sex differences" =
    largest_sex_differences,
  
  "Male excess burden" =
    largest_male_excess_burden,
  
  "Female excess burden" =
    largest_female_excess_burden,
  
  "Male excess severity" =
    largest_male_excess_severity,
  
  "Female excess severity" =
    largest_female_excess_severity
)

for (current_sheet in names(table_11_sheets)) {
  
  current_data <- table_11_sheets[[current_sheet]]
  
  openxlsx::addWorksheet(
    table_11_workbook,
    current_sheet
  )
  
  openxlsx::writeData(
    table_11_workbook,
    sheet =
      current_sheet,
    x =
      current_data
  )
  
  openxlsx::addStyle(
    table_11_workbook,
    sheet =
      current_sheet,
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
    table_11_workbook,
    sheet =
      current_sheet,
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
    table_11_workbook,
    sheet =
      current_sheet,
    firstRow =
      TRUE
  )
  
  openxlsx::addFilter(
    table_11_workbook,
    sheet =
      current_sheet,
    rows =
      1,
    cols =
      seq_len(
        ncol(
          current_data
        )
      )
  )
}

openxlsx::saveWorkbook(
  table_11_workbook,
  table_11_file,
  overwrite = TRUE
)


# ==============================================================================
# 29. EXPORT SUPPLEMENTARY TABLE S11
# ==============================================================================

s11_excel_file <- file.path(
  supplementary_folder,
  "Table_S11_Complete_sex_specific_framework_2023.xlsx"
)

s11_csv_file <- file.path(
  supplementary_folder,
  "Table_S11_Complete_sex_specific_framework_2023.csv"
)

s11_workbook <- openxlsx::createWorkbook()

s11_sheets <- list(
  "Country comparison" =
    supplementary_s11,
  
  "Quadrant distribution" =
    quadrant_distribution,
  
  "Transition matrix" =
    quadrant_transition_long,
  
  "Q1 comparison" =
    q1_comparison,
  
  "Largest sex differences" =
    largest_sex_differences
)

for (current_sheet in names(s11_sheets)) {
  
  current_data <- s11_sheets[[current_sheet]]
  
  openxlsx::addWorksheet(
    s11_workbook,
    current_sheet
  )
  
  openxlsx::writeData(
    s11_workbook,
    sheet =
      current_sheet,
    x =
      current_data
  )
  
  openxlsx::addStyle(
    s11_workbook,
    sheet =
      current_sheet,
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
    s11_workbook,
    sheet =
      current_sheet,
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
    s11_workbook,
    sheet =
      current_sheet,
    firstRow =
      TRUE
  )
  
  openxlsx::addFilter(
    s11_workbook,
    sheet =
      current_sheet,
    rows =
      1,
    cols =
      seq_len(
        ncol(
          current_data
        )
      )
  )
}

openxlsx::saveWorkbook(
  s11_workbook,
  s11_excel_file,
  overwrite = TRUE
)

readr::write_csv(
  supplementary_s11,
  s11_csv_file
)


# ==============================================================================
# 30. EXPORT FIGURE 11 SOURCE DATA
# ==============================================================================

figure_11_source_file <- file.path(
  supplementary_folder,
  "Figure_11_Source_data.xlsx"
)

figure_11_source_workbook <- openxlsx::createWorkbook()

figure_11_source_sheets <- list(
  "Male framework" =
    sex_framework |>
    select(
      country,
      male_daly_rate,
      male_dalys_per_case,
      male_quadrant,
      male_priority
    ),
  
  "Female framework" =
    sex_framework |>
    select(
      country,
      female_daly_rate,
      female_dalys_per_case,
      female_quadrant,
      female_priority
    ),
  
  "Quadrant transitions" =
    figure_11c_data,
  
  "Sex differences" =
    sex_framework |>
    select(
      country,
      daly_rate_difference_male_minus_female,
      dalys_per_case_difference_male_minus_female,
      quadrant_agreement,
      combined_absolute_sex_difference
    ),
  
  "Thresholds" =
    tibble(
      Sex = c(
        "Male",
        "Female"
      ),
      
      DALY_rate_threshold = c(
        male_burden_threshold,
        female_burden_threshold
      ),
      
      DALYs_per_case_threshold = c(
        male_severity_threshold,
        female_severity_threshold
      )
    )
)

for (current_sheet in names(figure_11_source_sheets)) {
  
  openxlsx::addWorksheet(
    figure_11_source_workbook,
    current_sheet
  )
  
  openxlsx::writeData(
    figure_11_source_workbook,
    sheet =
      current_sheet,
    x =
      figure_11_source_sheets[[current_sheet]]
  )
}

openxlsx::saveWorkbook(
  figure_11_source_workbook,
  figure_11_source_file,
  overwrite = TRUE
)


# ==============================================================================
# 31. DISPLAY KEY RESULTS
# ==============================================================================

message("")
message("==============================================================")
message("SEX-SPECIFIC QUADRANT DISTRIBUTION")
message("==============================================================")

print(
  quadrant_distribution,
  n = Inf
)

message("")
message("==============================================================")
message("MALE–FEMALE AGREEMENT")
message("==============================================================")

message(
  "Same quadrant: ",
  same_quadrant_count,
  " (",
  format_number(
    exact_agreement_percentage,
    1
  ),
  "%)"
)

message(
  "Different quadrant: ",
  different_quadrant_count,
  " (",
  format_number(
    100 -
      exact_agreement_percentage,
    1
  ),
  "%)"
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
  "Linear weighted Cohen's kappa: ",
  format_number(
    weighted_kappa$kappa,
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
message("Male-to-female transition matrix:")

print(
  quadrant_transition_matrix
)

message("")
message("Q1 comparison:")

print(
  q1_summary,
  n = Inf
)

message("")
message("Largest combined sex differences:")

print(
  largest_sex_differences |>
    select(
      country,
      male_quadrant_code,
      female_quadrant_code,
      daly_rate_difference_male_minus_female,
      dalys_per_case_difference_male_minus_female,
      combined_absolute_sex_difference
    ),
  n = Inf
)


# ==============================================================================
# 32. VALIDATE OUTPUTS
# ==============================================================================

required_output_files <- c(
  figure_11_file,
  table_11_file,
  s11_excel_file,
  s11_csv_file,
  figure_11_source_file
)

missing_output_files <- required_output_files[
  !file.exists(
    required_output_files
  )
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
# 33. COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("SECTION 3.11 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message("Inputs used:")
message(male_incidence_file)
message(female_incidence_file)
message(male_dalys_number_file)
message(female_dalys_number_file)
message(male_dalys_rate_file)
message(female_dalys_rate_file)

message("")
message(
  "Countries and territories analysed: ",
  expected_country_count
)

message(
  "Same male and female quadrant: ",
  same_quadrant_count,
  " (",
  format_number(
    exact_agreement_percentage,
    1
  ),
  "%)"
)

message(
  "Different male and female quadrant: ",
  different_quadrant_count
)

message("")
message("Main Figure 11:")
message(figure_11_file)

message("")
message("Main Table 11:")
message(table_11_file)

message("")
message("Supplementary Table S11:")
message(s11_excel_file)

message("")
message("Figure 11 source data:")
message(figure_11_source_file)

message("")
message("No YLL or YLD data were used.")
message("All outputs were saved under Publication.")
message("==============================================================")