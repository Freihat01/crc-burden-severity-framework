# ==============================================================================
# R_03_10_Country_trajectories_1990_2023.R
#
# RESULTS SECTION 3.10
# Country trajectories through the colorectal cancer burden–severity framework,
# 1990–2023
#
# REQUIRED INPUT FILES
#   01_Incidence_Number_AllCountries_BothSexes_AllAges_1990_2023*.csv
#   13_DALYs_Number_AllCountries_BothSexes_AllAges_1990_2023*.csv
#   16_DALYs_Rate_AllCountries_BothSexes_AgeStandarized_1990_2023*.csv
#
# ANALYTICAL FRAMEWORK
#   Burden dimension:
#       age-standardized DALY rate
#
#   Severity dimension:
#       DALYs per incident case
#       = DALY number / incident-case number
#
#   Fixed thresholds:
#       The 2023 medians are calculated once and applied to every year.
#       This preserves absolute movement over time and avoids redefining the
#       framework separately within each year.
#
#   Selected trajectory years:
#       1990, 2000, 2010, and 2023
#
# OUTPUTS
#
# Publication/Figures/
#   Figure_10_Country_trajectories_1990_2023.tiff
#
# Publication/Tables/
#   Table_10_Country_trajectory_summary.xlsx
#
# Publication/Supplementary/
#   Table_S10_Complete_country_trajectories_1990_2023.xlsx
#   Table_S10_Complete_country_trajectories_1990_2023.csv
#   Figure_10_Source_data.xlsx
#
# IMPORTANT
#   - No YLL or YLD files are used.
#   - No global data are used.
#   - The same 2023 thresholds are applied to 1990, 2000, 2010, and 2023.
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
# 5. LOCATE THE THREE REQUIRED INPUT FILES
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
  
  # Prefer the most recently modified exact matching copy.
  candidate_info <- file.info(candidates)
  
  selected <- candidates[
    order(
      candidate_info$mtime,
      decreasing = TRUE
    )
  ][1]
  
  message("")
  message(label, ":")
  message(selected)
  
  selected
}


incidence_file <- find_single_input(
  pattern =
    "^01_Incidence_Number_AllCountries_BothSexes_AllAges_1990_2023.*\\.csv$",
  label =
    "Incidence-number file"
)

dalys_number_file <- find_single_input(
  pattern =
    "^13_DALYs_Number_AllCountries_BothSexes_AllAges_1990_2023.*\\.csv$",
  label =
    "DALY-number file"
)

dalys_rate_file <- find_single_input(
  pattern =
    "^16_DALYs_Rate_AllCountries_BothSexes_AgeStandarized_1990_2023.*\\.csv$",
  label =
    "Age-standardized DALY-rate file"
)


# ==============================================================================
# 6. ANALYTICAL SETTINGS
# ==============================================================================

expected_country_count <- 204L

selected_years <- c(
  1990L,
  2000L,
  2010L,
  2023L
)

baseline_year <- 1990L
final_year <- 2023L

quadrant_levels <- c(
  "Q1: High burden–high severity",
  "Q2: Low burden–high severity",
  "Q3: Low burden–low severity",
  "Q4: High burden–low severity"
)

quadrant_codes <- c(
  "Q1",
  "Q2",
  "Q3",
  "Q4"
)

quadrant_colours <- c(
  "Q1: High burden–high severity" = "#B2182B",
  "Q2: Low burden–high severity" = "#EF8A62",
  "Q3: Low burden–low severity" = "#67A9CF",
  "Q4: High burden–low severity" = "#2166AC"
)

trajectory_colours <- c(
  "Improved in both dimensions" = "#2166AC",
  "Burden improved; severity worsened" = "#67A9CF",
  "Burden worsened; severity improved" = "#EF8A62",
  "Deteriorated in both dimensions" = "#B2182B",
  "No meaningful change" = "#737373"
)

top_n_selected <- 10L


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


priority_number_from_quadrant <- function(quadrant_code) {
  dplyr::case_when(
    quadrant_code == "Q1" ~ 1L,
    quadrant_code == "Q2" ~ 2L,
    quadrant_code == "Q4" ~ 3L,
    quadrant_code == "Q3" ~ 4L,
    TRUE ~ NA_integer_
  )
}


read_and_validate_file <- function(
    file_path,
    expected_measure,
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
      country = clean_text(location),
      year = as.integer(year),
      sex = clean_text(sex),
      age = clean_text(age),
      metric = clean_text(metric),
      measure = clean_text(measure),
      
      "{value_name}" :=
        as.numeric(val),
      
      "{paste0(value_name, '_lower')}" :=
        as.numeric(lower),
      
      "{paste0(value_name, '_upper')}" :=
        as.numeric(upper)
    ) |>
    filter(
      sex == "Both",
      age == expected_age,
      metric == expected_metric,
      year >= 1990,
      year <= 2023
    )
  
  if (nrow(data) == 0) {
    stop(
      paste0(
        "\nNo valid rows remained after filtering:\n",
        file_path,
        "\nExpected age: ",
        expected_age,
        "\nExpected metric: ",
        expected_metric
      ),
      call. = FALSE
    )
  }
  
  duplicate_rows <- data |>
    count(
      country,
      year,
      name = "records"
    ) |>
    filter(records > 1)
  
  if (nrow(duplicate_rows) > 0) {
    print(duplicate_rows, n = Inf)
    
    stop(
      paste0(
        "\nDuplicate country-year observations were detected in:\n",
        file_path
      ),
      call. = FALSE
    )
  }
  
  data
}


# ==============================================================================
# 8. IMPORT THE THREE LONGITUDINAL DATASETS
# ==============================================================================

incidence_data <- read_and_validate_file(
  file_path = incidence_file,
  expected_measure = "Incidence",
  expected_age = "All ages",
  expected_metric = "Number",
  value_name = "incident_cases"
)

dalys_number_data <- read_and_validate_file(
  file_path = dalys_number_file,
  expected_measure = "DALYs (Disability-Adjusted Life Years)",
  expected_age = "All ages",
  expected_metric = "Number",
  value_name = "dalys"
)

dalys_rate_data <- read_and_validate_file(
  file_path = dalys_rate_file,
  expected_measure = "DALYs (Disability-Adjusted Life Years)",
  expected_age = "Age-standardized",
  expected_metric = "Rate",
  value_name = "daly_rate"
)


# ==============================================================================
# 9. MERGE COUNTRY-YEAR DATA
# ==============================================================================

trajectory_annual <- incidence_data |>
  select(
    country,
    year,
    incident_cases,
    incident_cases_lower,
    incident_cases_upper
  ) |>
  inner_join(
    dalys_number_data |>
      select(
        country,
        year,
        dalys,
        dalys_lower,
        dalys_upper
      ),
    by = c(
      "country",
      "year"
    )
  ) |>
  inner_join(
    dalys_rate_data |>
      select(
        country,
        year,
        daly_rate,
        daly_rate_lower,
        daly_rate_upper
      ),
    by = c(
      "country",
      "year"
    )
  ) |>
  mutate(
    dalys_per_case =
      dalys /
      incident_cases
  ) |>
  arrange(
    country,
    year
  )


# ==============================================================================
# 10. QUALITY CONTROL OF MERGED DATA
# ==============================================================================

expected_year_count <- length(
  1990:2023
)

country_year_counts <- trajectory_annual |>
  count(
    country,
    name = "years"
  )

incomplete_countries <- country_year_counts |>
  filter(
    years != expected_year_count
  )

if (nrow(incomplete_countries) > 0) {
  print(incomplete_countries, n = Inf)
  
  stop(
    "\nSome countries do not contain all 34 years from 1990 to 2023.",
    call. = FALSE
  )
}

countries_per_year <- trajectory_annual |>
  count(
    year,
    name = "countries"
  )

invalid_year_counts <- countries_per_year |>
  filter(
    countries != expected_country_count
  )

if (nrow(invalid_year_counts) > 0) {
  print(invalid_year_counts, n = Inf)
  
  stop(
    paste0(
      "\nExpected ",
      expected_country_count,
      " countries in every year."
    ),
    call. = FALSE
  )
}

invalid_values <- trajectory_annual |>
  filter(
    is.na(incident_cases) |
      !is.finite(incident_cases) |
      incident_cases <= 0 |
      is.na(dalys) |
      !is.finite(dalys) |
      dalys < 0 |
      is.na(daly_rate) |
      !is.finite(daly_rate) |
      daly_rate < 0 |
      is.na(dalys_per_case) |
      !is.finite(dalys_per_case) |
      dalys_per_case < 0
  )

if (nrow(invalid_values) > 0) {
  print(invalid_values, n = Inf)
  
  stop(
    "\nInvalid longitudinal burden or severity values were detected.",
    call. = FALSE
  )
}


# ==============================================================================
# 11. CALCULATE FIXED 2023 THRESHOLDS
# ==============================================================================

framework_2023 <- trajectory_annual |>
  filter(
    year == final_year
  )

burden_threshold_2023 <- median(
  framework_2023$daly_rate,
  na.rm = TRUE
)

severity_threshold_2023 <- median(
  framework_2023$dalys_per_case,
  na.rm = TRUE
)

message("")
message(
  "Fixed 2023 DALY-rate threshold: ",
  round(
    burden_threshold_2023,
    2
  )
)

message(
  "Fixed 2023 DALYs-per-case threshold: ",
  round(
    severity_threshold_2023,
    2
  )
)


# ==============================================================================
# 12. ASSIGN QUADRANTS FOR EVERY COUNTRY-YEAR
# ==============================================================================

trajectory_annual <- trajectory_annual |>
  mutate(
    quadrant =
      create_quadrant(
        daly_rate =
          daly_rate,
        dalys_per_case =
          dalys_per_case,
        burden_threshold =
          burden_threshold_2023,
        severity_threshold =
          severity_threshold_2023
      ),
    
    quadrant_code =
      extract_quadrant_code(
        quadrant
      ),
    
    priority_number =
      priority_number_from_quadrant(
        quadrant_code
      ),
    
    high_burden =
      daly_rate >=
      burden_threshold_2023,
    
    high_severity =
      dalys_per_case >=
      severity_threshold_2023
  )


# ==============================================================================
# 13. SELECT TRAJECTORY YEARS
# ==============================================================================

trajectory_selected <- trajectory_annual |>
  filter(
    year %in%
      selected_years
  ) |>
  mutate(
    year_factor =
      factor(
        year,
        levels =
          selected_years
      ),
    
    quadrant =
      factor(
        quadrant,
        levels =
          quadrant_levels
      )
  )


# ==============================================================================
# 14. CREATE COUNTRY-LEVEL 1990–2023 TRAJECTORY SUMMARY
# ==============================================================================

baseline_data <- trajectory_annual |>
  filter(
    year == baseline_year
  ) |>
  transmute(
    country,
    daly_rate_1990 = daly_rate,
    dalys_per_case_1990 = dalys_per_case,
    quadrant_1990 = quadrant,
    quadrant_code_1990 = quadrant_code,
    priority_1990 = priority_number
  )

final_data <- trajectory_annual |>
  filter(
    year == final_year
  ) |>
  transmute(
    country,
    daly_rate_2023 = daly_rate,
    dalys_per_case_2023 = dalys_per_case,
    quadrant_2023 = quadrant,
    quadrant_code_2023 = quadrant_code,
    priority_2023 = priority_number
  )

trajectory_summary_country <- baseline_data |>
  inner_join(
    final_data,
    by = "country"
  ) |>
  mutate(
    change_daly_rate =
      daly_rate_2023 -
      daly_rate_1990,
    
    percent_change_daly_rate =
      100 *
      (
        daly_rate_2023 -
          daly_rate_1990
      ) /
      daly_rate_1990,
    
    change_dalys_per_case =
      dalys_per_case_2023 -
      dalys_per_case_1990,
    
    percent_change_dalys_per_case =
      100 *
      (
        dalys_per_case_2023 -
          dalys_per_case_1990
      ) /
      dalys_per_case_1990,
    
    same_quadrant =
      quadrant_code_1990 ==
      quadrant_code_2023,
    
    priority_change =
      priority_1990 -
      priority_2023,
    
    priority_direction =
      case_when(
        priority_change > 0 ~
          "Higher priority in 2023",
        priority_change < 0 ~
          "Lower priority in 2023",
        TRUE ~
          "Same priority"
      ),
    
    movement_category =
      case_when(
        change_daly_rate < 0 &
          change_dalys_per_case < 0 ~
          "Improved in both dimensions",
        
        change_daly_rate < 0 &
          change_dalys_per_case > 0 ~
          "Burden improved; severity worsened",
        
        change_daly_rate > 0 &
          change_dalys_per_case < 0 ~
          "Burden worsened; severity improved",
        
        change_daly_rate > 0 &
          change_dalys_per_case > 0 ~
          "Deteriorated in both dimensions",
        
        TRUE ~
          "No meaningful change"
      )
  )


# ==============================================================================
# 15. STANDARDIZED TWO-DIMENSIONAL MOVEMENT
# ==============================================================================

daly_rate_sd <- sd(
  trajectory_summary_country$change_daly_rate,
  na.rm = TRUE
)

dalys_per_case_sd <- sd(
  trajectory_summary_country$change_dalys_per_case,
  na.rm = TRUE
)

trajectory_summary_country <- trajectory_summary_country |>
  mutate(
    standardized_change_daly_rate =
      change_daly_rate /
      daly_rate_sd,
    
    standardized_change_dalys_per_case =
      change_dalys_per_case /
      dalys_per_case_sd,
    
    trajectory_distance =
      sqrt(
        standardized_change_daly_rate^2 +
          standardized_change_dalys_per_case^2
      ),
    
    improvement_score =
      -standardized_change_daly_rate -
      standardized_change_dalys_per_case,
    
    deterioration_score =
      standardized_change_daly_rate +
      standardized_change_dalys_per_case
  )


# ==============================================================================
# 16. COUNT QUADRANT CHANGES ACROSS ALL YEARS
# ==============================================================================

annual_transition_counts <- trajectory_annual |>
  group_by(
    country
  ) |>
  arrange(
    year,
    .by_group = TRUE
  ) |>
  summarise(
    number_of_annual_quadrant_changes =
      sum(
        quadrant_code !=
          lag(
            quadrant_code
          ),
        na.rm = TRUE
      ),
    
    ever_in_q1 =
      any(
        quadrant_code == "Q1"
      ),
    
    ever_in_q2 =
      any(
        quadrant_code == "Q2"
      ),
    
    ever_in_q3 =
      any(
        quadrant_code == "Q3"
      ),
    
    ever_in_q4 =
      any(
        quadrant_code == "Q4"
      ),
    
    .groups =
      "drop"
  )

trajectory_summary_country <- trajectory_summary_country |>
  left_join(
    annual_transition_counts,
    by = "country"
  )


# ==============================================================================
# 17. TRANSITION MATRIX: 1990 TO 2023
# ==============================================================================

transition_matrix <- table(
  `1990 quadrant` =
    factor(
      trajectory_summary_country$quadrant_code_1990,
      levels =
        quadrant_codes
    ),
  
  `2023 quadrant` =
    factor(
      trajectory_summary_country$quadrant_code_2023,
      levels =
        quadrant_codes
    )
)

transition_matrix_long <- as.data.frame(
  transition_matrix
) |>
  rename(
    Quadrant_1990 =
      `X1990.quadrant`,
    Quadrant_2023 =
      `X2023.quadrant`,
    Countries =
      Freq
  ) |>
  mutate(
    Percentage =
      100 *
      Countries /
      expected_country_count
  )


# ==============================================================================
# 18. OVERALL TRAJECTORY SUMMARIES
# ==============================================================================

quadrant_change_summary <- tibble(
  Measure = c(
    "Countries remaining in the same quadrant",
    "Countries changing quadrant at least once between 1990 and 2023",
    "Countries in the same quadrant in 1990 and 2023",
    "Countries in a different quadrant in 2023",
    "Countries entering Q1",
    "Countries leaving Q1",
    "Countries moving from Q1 to Q3",
    "Countries moving from Q3 to Q1"
  ),
  
  Countries = c(
    sum(
      trajectory_summary_country$number_of_annual_quadrant_changes == 0
    ),
    
    sum(
      trajectory_summary_country$number_of_annual_quadrant_changes > 0
    ),
    
    sum(
      trajectory_summary_country$same_quadrant
    ),
    
    sum(
      !trajectory_summary_country$same_quadrant
    ),
    
    sum(
      trajectory_summary_country$quadrant_code_1990 != "Q1" &
        trajectory_summary_country$quadrant_code_2023 == "Q1"
    ),
    
    sum(
      trajectory_summary_country$quadrant_code_1990 == "Q1" &
        trajectory_summary_country$quadrant_code_2023 != "Q1"
    ),
    
    sum(
      trajectory_summary_country$quadrant_code_1990 == "Q1" &
        trajectory_summary_country$quadrant_code_2023 == "Q3"
    ),
    
    sum(
      trajectory_summary_country$quadrant_code_1990 == "Q3" &
        trajectory_summary_country$quadrant_code_2023 == "Q1"
    )
  )
) |>
  mutate(
    Percentage =
      100 *
      Countries /
      expected_country_count
  )

movement_category_summary <- trajectory_summary_country |>
  count(
    movement_category,
    name = "Countries"
  ) |>
  mutate(
    Percentage =
      100 *
      Countries /
      expected_country_count
  )


# ==============================================================================
# 19. IDENTIFY LARGEST IMPROVEMENTS AND DETERIORATIONS
# ==============================================================================

largest_improvements <- trajectory_summary_country |>
  filter(
    movement_category ==
      "Improved in both dimensions"
  ) |>
  arrange(
    desc(
      improvement_score
    )
  ) |>
  slice_head(
    n = top_n_selected
  )

largest_deteriorations <- trajectory_summary_country |>
  filter(
    movement_category ==
      "Deteriorated in both dimensions"
  ) |>
  arrange(
    desc(
      deterioration_score
    )
  ) |>
  slice_head(
    n = top_n_selected
  )

selected_trajectory_countries <- bind_rows(
  largest_improvements |>
    mutate(
      selected_group =
        "Largest improvements"
    ),
  
  largest_deteriorations |>
    mutate(
      selected_group =
        "Largest deteriorations"
    )
) |>
  distinct(
    country,
    .keep_all = TRUE
  )


# ==============================================================================
# 20. PREPARE FIGURE 10 PANEL A: ALLUVIAL TRANSITIONS
# ==============================================================================

figure_10a_data <- trajectory_selected |>
  select(
    country,
    year_factor,
    quadrant
  ) |>
  pivot_wider(
    names_from =
      year_factor,
    values_from =
      quadrant,
    names_prefix =
      "year_"
  ) |>
  mutate(
    across(
      starts_with("year_"),
      ~ factor(
        .x,
        levels =
          quadrant_levels
      )
    )
  )

figure_10a <- ggplot(
  figure_10a_data,
  aes(
    axis1 = year_1990,
    axis2 = year_2000,
    axis3 = year_2010,
    axis4 = year_2023,
    y = 1
  )
) +
  
  ggalluvial::geom_alluvium(
    aes(
      fill =
        year_2023
    ),
    width = 0.11,
    alpha = 0.60,
    knot.pos = 0.35,
    color = "grey78",
    linewidth = 0.08
  ) +
  
  ggalluvial::geom_stratum(
    width = 0.13,
    fill = "grey97",
    color = "grey30",
    linewidth = 0.40
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
    size = 3.4,
    fontface = "bold"
  ) +
  
  scale_x_discrete(
    limits = c(
      "1990",
      "2000",
      "2010",
      "2023"
    ),
    expand = c(
      0.10,
      0.10
    )
  ) +
  
  scale_fill_manual(
    values =
      quadrant_colours,
    drop =
      FALSE,
    name =
      "2023 quadrant"
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
      "A. Movement between framework quadrants"
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


# ==============================================================================
# 21. PREPARE FIGURE 10 PANEL B: TRANSITION MATRIX
# ==============================================================================

figure_10b <- ggplot(
  transition_matrix_long,
  aes(
    x =
      Quadrant_2023,
    y =
      Quadrant_1990,
    fill =
      Countries
  )
) +
  
  geom_tile(
    color = "white",
    linewidth = 0.8
  ) +
  
  geom_text(
    aes(
      label =
        Countries
    ),
    size = 4,
    fontface = "bold"
  ) +
  
  scale_fill_viridis_c(
    option = "C",
    direction = -1,
    name = "Countries"
  ) +
  
  scale_x_discrete(
    drop = FALSE
  ) +
  
  scale_y_discrete(
    drop = FALSE,
    limits = rev(
      quadrant_codes
    )
  ) +
  
  coord_equal() +
  
  labs(
    title =
      "B. Quadrant transition matrix, 1990–2023",
    
    x =
      "2023 quadrant",
    
    y =
      "1990 quadrant"
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
    
    axis.text =
      element_text(
        face = "bold"
      ),
    
    legend.position =
      "bottom"
  )


# ==============================================================================
# 22. PREPARE FIGURE 10 PANELS C AND D: SELECTED COUNTRY TRAJECTORIES
# ==============================================================================

selected_trajectory_long <- trajectory_selected |>
  inner_join(
    selected_trajectory_countries |>
      select(
        country,
        selected_group
      ),
    by = "country"
  ) |>
  mutate(
    selected_group =
      factor(
        selected_group,
        levels = c(
          "Largest improvements",
          "Largest deteriorations"
        )
      )
  )

create_selected_trajectory_plot <- function(
    group_name,
    panel_title
) {
  
  plot_data <- selected_trajectory_long |>
    filter(
      selected_group ==
        group_name
    )
  
  final_labels <- plot_data |>
    filter(
      year == final_year
    )
  
  ggplot(
    plot_data,
    aes(
      x =
        daly_rate,
      y =
        dalys_per_case,
      group =
        country,
      color =
        country
    )
  ) +
    
    geom_path(
      linewidth = 0.8,
      alpha = 0.80,
      arrow = arrow(
        type = "closed",
        length = grid::unit(
          0.10,
          "inches"
        )
      )
    ) +
    
    geom_point(
      aes(
        shape =
          year_factor
      ),
      size = 2.4,
      stroke = 0.8
    ) +
    
    ggrepel::geom_text_repel(
      data =
        final_labels,
      aes(
        label =
          country
      ),
      size = 3,
      fontface = "bold",
      show.legend = FALSE,
      max.overlaps = Inf,
      box.padding = 0.30,
      point.padding = 0.20,
      min.segment.length = 0
    ) +
    
    geom_vline(
      xintercept =
        burden_threshold_2023,
      linetype =
        "dashed",
      linewidth =
        0.6,
      color =
        "grey35"
    ) +
    
    geom_hline(
      yintercept =
        severity_threshold_2023,
      linetype =
        "dashed",
      linewidth =
        0.6,
      color =
        "grey35"
    ) +
    
    scale_shape_manual(
      values = c(
        "1990" = 21,
        "2000" = 22,
        "2010" = 23,
        "2023" = 19
      ),
      name = "Year"
    ) +
    
    guides(
      color = "none"
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
        )
    )
}


figure_10c <- create_selected_trajectory_plot(
  group_name =
    "Largest improvements",
  panel_title =
    "C. Largest simultaneous improvements"
)

figure_10d <- create_selected_trajectory_plot(
  group_name =
    "Largest deteriorations",
  panel_title =
    "D. Largest simultaneous deteriorations"
)


# ==============================================================================
# 23. COMBINE FIGURE 10
# ==============================================================================

figure_10 <- (
  figure_10a +
    figure_10b
) / (
  figure_10c +
    figure_10d
) +
  
  patchwork::plot_layout(
    widths = c(
      1.20,
      0.80
    ),
    heights = c(
      1,
      1
    ),
    guides = "keep"
  ) +
  
  patchwork::plot_annotation(
    title =
      "Country trajectories through the colorectal cancer burden–severity framework, 1990–2023",
    
    subtitle =
      paste0(
        "Fixed 2023 thresholds: DALY rate ",
        formatC(
          burden_threshold_2023,
          format = "f",
          digits = 2
        ),
        " per 100,000 and ",
        formatC(
          severity_threshold_2023,
          format = "f",
          digits = 2
        ),
        " DALYs per incident case"
      ),
    
    caption =
      paste0(
        "The 2023 median thresholds were applied to all years to preserve ",
        "absolute temporal movement. Panels C and D show the ten countries ",
        "with the largest standardized simultaneous improvements and ",
        "deteriorations, respectively. Arrows indicate movement from 1990 ",
        "through 2000 and 2010 to 2023."
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
          margin = margin(
            b = 8
          )
        ),
      
      plot.caption =
        element_text(
          size = 8.3,
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
# 24. SAVE FIGURE 10
# ==============================================================================

figure_10_file <- file.path(
  figures_folder,
  "Figure_10_Country_trajectories_1990_2023.tiff"
)

ggsave(
  filename =
    figure_10_file,
  
  plot =
    figure_10,
  
  device =
    "tiff",
  
  width =
    15.5,
  
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
# 25. CREATE TABLE 10
# ==============================================================================

table_10_main <- bind_rows(
  quadrant_change_summary |>
    transmute(
      Category =
        "Quadrant movement",
      Measure,
      Countries,
      Percentage
    ),
  
  movement_category_summary |>
    transmute(
      Category =
        "Two-dimensional change",
      Measure =
        movement_category,
      Countries,
      Percentage
    )
)

table_10_thresholds <- tibble(
  Measure = c(
    "Fixed 2023 median DALY-rate threshold",
    "Fixed 2023 median DALYs-per-case threshold"
  ),
  
  Value = c(
    burden_threshold_2023,
    severity_threshold_2023
  )
)


# ==============================================================================
# 26. PREPARE SUPPLEMENTARY TABLE S10
# ==============================================================================

supplementary_s10 <- trajectory_summary_country |>
  transmute(
    Country =
      country,
    
    DALY_rate_1990 =
      daly_rate_1990,
    
    DALYs_per_case_1990 =
      dalys_per_case_1990,
    
    Quadrant_1990 =
      quadrant_code_1990,
    
    DALY_rate_2023 =
      daly_rate_2023,
    
    DALYs_per_case_2023 =
      dalys_per_case_2023,
    
    Quadrant_2023 =
      quadrant_code_2023,
    
    Same_quadrant_1990_2023 =
      if_else(
        same_quadrant,
        "Yes",
        "No"
      ),
    
    Priority_1990 =
      priority_1990,
    
    Priority_2023 =
      priority_2023,
    
    Priority_change =
      priority_change,
    
    Priority_direction =
      priority_direction,
    
    Change_in_DALY_rate =
      change_daly_rate,
    
    Percent_change_in_DALY_rate =
      percent_change_daly_rate,
    
    Change_in_DALYs_per_case =
      change_dalys_per_case,
    
    Percent_change_in_DALYs_per_case =
      percent_change_dalys_per_case,
    
    Movement_category =
      movement_category,
    
    Standardized_change_DALY_rate =
      standardized_change_daly_rate,
    
    Standardized_change_DALYs_per_case =
      standardized_change_dalys_per_case,
    
    Trajectory_distance =
      trajectory_distance,
    
    Improvement_score =
      improvement_score,
    
    Deterioration_score =
      deterioration_score,
    
    Number_of_annual_quadrant_changes =
      number_of_annual_quadrant_changes,
    
    Ever_in_Q1 =
      ever_in_q1,
    
    Ever_in_Q2 =
      ever_in_q2,
    
    Ever_in_Q3 =
      ever_in_q3,
    
    Ever_in_Q4 =
      ever_in_q4
  ) |>
  arrange(
    desc(
      Trajectory_distance
    ),
    Country
  )


# ==============================================================================
# 27. EXCEL STYLES
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
# 28. EXPORT TABLE 10
# ==============================================================================

table_10_file <- file.path(
  tables_folder,
  "Table_10_Country_trajectory_summary.xlsx"
)

table_10_workbook <- openxlsx::createWorkbook()

table_10_sheets <- list(
  "Trajectory summary" =
    table_10_main,
  
  "Transition matrix" =
    transition_matrix_long,
  
  "Thresholds" =
    table_10_thresholds,
  
  "Largest improvements" =
    largest_improvements,
  
  "Largest deteriorations" =
    largest_deteriorations
)

for (current_sheet in names(table_10_sheets)) {
  
  current_data <- table_10_sheets[[current_sheet]]
  
  openxlsx::addWorksheet(
    table_10_workbook,
    current_sheet
  )
  
  openxlsx::writeData(
    table_10_workbook,
    sheet = current_sheet,
    x = current_data
  )
  
  openxlsx::addStyle(
    table_10_workbook,
    sheet = current_sheet,
    style = header_style,
    rows = 1,
    cols = seq_len(ncol(current_data)),
    gridExpand = TRUE
  )
  
  openxlsx::setColWidths(
    table_10_workbook,
    sheet = current_sheet,
    cols = seq_len(ncol(current_data)),
    widths = "auto"
  )
  
  openxlsx::freezePane(
    table_10_workbook,
    sheet = current_sheet,
    firstRow = TRUE
  )
  
  openxlsx::addFilter(
    table_10_workbook,
    sheet = current_sheet,
    rows = 1,
    cols = seq_len(ncol(current_data))
  )
}

openxlsx::saveWorkbook(
  table_10_workbook,
  table_10_file,
  overwrite = TRUE
)


# ==============================================================================
# 29. EXPORT SUPPLEMENTARY TABLE S10
# ==============================================================================

s10_excel_file <- file.path(
  supplementary_folder,
  "Table_S10_Complete_country_trajectories_1990_2023.xlsx"
)

s10_csv_file <- file.path(
  supplementary_folder,
  "Table_S10_Complete_country_trajectories_1990_2023.csv"
)

s10_workbook <- openxlsx::createWorkbook()

s10_sheets <- list(
  "Country trajectory summary" =
    supplementary_s10,
  
  "Selected years" =
    trajectory_selected,
  
  "Annual framework" =
    trajectory_annual,
  
  "Transition matrix" =
    transition_matrix_long,
  
  "Movement categories" =
    movement_category_summary
)

for (current_sheet in names(s10_sheets)) {
  
  current_data <- s10_sheets[[current_sheet]]
  
  openxlsx::addWorksheet(
    s10_workbook,
    current_sheet
  )
  
  openxlsx::writeData(
    s10_workbook,
    sheet = current_sheet,
    x = current_data
  )
  
  openxlsx::addStyle(
    s10_workbook,
    sheet = current_sheet,
    style = header_style,
    rows = 1,
    cols = seq_len(ncol(current_data)),
    gridExpand = TRUE
  )
  
  openxlsx::setColWidths(
    s10_workbook,
    sheet = current_sheet,
    cols = seq_len(ncol(current_data)),
    widths = "auto"
  )
  
  openxlsx::freezePane(
    s10_workbook,
    sheet = current_sheet,
    firstRow = TRUE
  )
  
  openxlsx::addFilter(
    s10_workbook,
    sheet = current_sheet,
    rows = 1,
    cols = seq_len(ncol(current_data))
  )
}

openxlsx::saveWorkbook(
  s10_workbook,
  s10_excel_file,
  overwrite = TRUE
)

readr::write_csv(
  supplementary_s10,
  s10_csv_file
)


# ==============================================================================
# 30. EXPORT FIGURE 10 SOURCE DATA
# ==============================================================================

figure_10_source_file <- file.path(
  supplementary_folder,
  "Figure_10_Source_data.xlsx"
)

figure_10_source_workbook <- openxlsx::createWorkbook()

figure_10_source_sheets <- list(
  "Alluvial trajectories" =
    figure_10a_data,
  
  "Transition matrix" =
    transition_matrix_long,
  
  "Selected trajectories" =
    selected_trajectory_long,
  
  "Thresholds" =
    table_10_thresholds
)

for (current_sheet in names(figure_10_source_sheets)) {
  
  openxlsx::addWorksheet(
    figure_10_source_workbook,
    current_sheet
  )
  
  openxlsx::writeData(
    figure_10_source_workbook,
    sheet = current_sheet,
    x = figure_10_source_sheets[[current_sheet]]
  )
}

openxlsx::saveWorkbook(
  figure_10_source_workbook,
  figure_10_source_file,
  overwrite = TRUE
)


# ==============================================================================
# 31. DISPLAY KEY RESULTS
# ==============================================================================

message("")
message("==============================================================")
message("COUNTRY TRAJECTORY SUMMARY")
message("==============================================================")

print(
  quadrant_change_summary,
  n = Inf
)

message("")
message("Movement categories:")

print(
  movement_category_summary,
  n = Inf
)

message("")
message("1990-to-2023 transition matrix:")

print(
  transition_matrix
)

message("")
message("Largest improvements:")

print(
  largest_improvements |>
    select(
      country,
      quadrant_code_1990,
      quadrant_code_2023,
      percent_change_daly_rate,
      percent_change_dalys_per_case,
      improvement_score
    ),
  n = Inf
)

message("")
message("Largest deteriorations:")

print(
  largest_deteriorations |>
    select(
      country,
      quadrant_code_1990,
      quadrant_code_2023,
      percent_change_daly_rate,
      percent_change_dalys_per_case,
      deterioration_score
    ),
  n = Inf
)


# ==============================================================================
# 32. VALIDATE OUTPUTS
# ==============================================================================

required_output_files <- c(
  figure_10_file,
  table_10_file,
  s10_excel_file,
  s10_csv_file,
  figure_10_source_file
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
message("SECTION 3.10 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message("Inputs used:")
message(incidence_file)
message(dalys_number_file)
message(dalys_rate_file)

message("")
message(
  "Countries and territories analysed: ",
  expected_country_count
)

message(
  "Fixed 2023 DALY-rate threshold: ",
  round(
    burden_threshold_2023,
    2
  )
)

message(
  "Fixed 2023 DALYs-per-case threshold: ",
  round(
    severity_threshold_2023,
    2
  )
)

message("")
message("Main Figure 10:")
message(figure_10_file)

message("")
message("Main Table 10:")
message(table_10_file)

message("")
message("Supplementary Table S10:")
message(s10_excel_file)

message("")
message("Figure 10 source data:")
message(figure_10_source_file)

message("")
message("No YLL or YLD data were used.")
message("No global data were used.")
message("All outputs were saved under Publication.")
message("==============================================================")