# ==============================================================================
# R_03_1_Global_Burden_2023.R
#
# RESULTS SECTION 3.1
# Global colorectal cancer burden in 2023
#
# INPUT
# Global/Global_CRC_Incidence_Deaths_DALYs_Number_Rate_1990_2023.csv
#
# OUTPUTS
# Publication/Tables/
#   Table_1_Global_CRC_burden_2023.xlsx
#
# Publication/Figures/
#   Figure_1_Global_CRC_burden_2023.tiff
#
# Publication/Supplementary/
#   Table_S1_Global_CRC_detailed_estimates_2023.xlsx
#
# Publication/Results_text/
#   Results_3_1_Global_CRC_burden_2023.txt
#
# FIGURE 1
# One TIFF file at 600 dpi containing six panels:
#   A. Incident cases
#   B. Age-standardized incidence rate
#   C. Deaths
#   D. Age-standardized mortality rate
#   E. DALYs
#   F. Age-standardized DALY rate
#
# DATA INCLUDED
#   Incidence
#   Deaths
#   DALYs
#
# The old clean-data folder is not used.
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
  "patchwork",
  "scales",
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
  library(patchwork)
  library(scales)
  library(glue)
  
})


# ==============================================================================
# 3. IDENTIFY THE MAIN PROJECT FOLDER
#
# The script works when the R project is located either in:
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

if (dir.exists(file.path(current_folder, "Global"))) {
  
  project_root <- current_folder
  
} else if (dir.exists(file.path(current_folder, "..", "Global"))) {
  
  project_root <- normalizePath(
    file.path(current_folder, ".."),
    winslash = "/",
    mustWork = TRUE
  )
  
} else {
  
  stop(
    paste0(
      "\nThe Global folder could not be found.\n\n",
      "Current working directory:\n",
      current_folder,
      "\n\nThe script checked the current folder and its parent."
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
  publication_folder,
  recursive = TRUE,
  showWarnings = FALSE
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
# 5. LOCATE THE RAW GLOBAL DATA FILE
# ==============================================================================

global_files <- list.files(
  path = global_folder,
  pattern = paste0(
    "^Global_CRC_Incidence_Deaths_DALYs_",
    "Number_Rate_1990_2023.*\\.csv$"
  ),
  full.names = TRUE,
  ignore.case = TRUE
)

if (length(global_files) == 0) {
  
  stop(
    paste0(
      "\nThe required global CSV file was not found in:\n",
      global_folder,
      "\n\nExpected filename beginning with:\n",
      "Global_CRC_Incidence_Deaths_DALYs_Number_Rate_1990_2023"
    ),
    call. = FALSE
  )
  
}

if (length(global_files) > 1) {
  
  stop(
    paste0(
      "\nMore than one matching global file was found:\n",
      paste(
        basename(global_files),
        collapse = "\n"
      )
    ),
    call. = FALSE
  )
  
}

global_file <- global_files[[1]]

message("")
message("Raw data file:")
message(global_file)


# ==============================================================================
# 6. IMPORT THE RAW DATA
# ==============================================================================

raw_global <- readr::read_csv(
  file = global_file,
  show_col_types = FALSE,
  progress = FALSE,
  guess_max = 100000
) |>
  janitor::clean_names()


# ==============================================================================
# 7. CONFIRM REQUIRED COLUMNS
# ==============================================================================

required_columns <- c(
  "measure",
  "metric",
  "sex",
  "age",
  "year",
  "val",
  "lower",
  "upper"
)

missing_columns <- setdiff(
  required_columns,
  names(raw_global)
)

if (length(missing_columns) > 0) {
  
  stop(
    paste0(
      "\nThe raw file is missing the following required columns:\n",
      paste(
        missing_columns,
        collapse = ", "
      ),
      "\n\nAvailable columns:\n",
      paste(
        names(raw_global),
        collapse = ", "
      )
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 8. CLEAN AND STANDARDISE VARIABLES
# ==============================================================================

clean_text <- function(x) {
  
  x |>
    as.character() |>
    stringr::str_replace_all("\u00A0", " ") |>
    stringr::str_squish()
  
}


global_data <- raw_global |>
  
  mutate(
    
    measure = clean_text(measure),
    
    metric = clean_text(metric),
    
    sex = clean_text(sex),
    
    age = clean_text(age),
    
    year = as.integer(year),
    
    val = as.numeric(val),
    
    lower = as.numeric(lower),
    
    upper = as.numeric(upper)
    
  ) |>
  
  mutate(
    
    sex = case_when(
      
      str_to_lower(sex) %in% c(
        "both",
        "both sex",
        "both sexes"
      ) ~ "Both",
      
      str_to_lower(sex) %in% c(
        "male",
        "males"
      ) ~ "Male",
      
      str_to_lower(sex) %in% c(
        "female",
        "females"
      ) ~ "Female",
      
      TRUE ~ sex
      
    ),
    
    measure = case_when(
      
      str_detect(
        str_to_lower(measure),
        "^incidence$|incident"
      ) ~ "Incidence",
      
      str_detect(
        str_to_lower(measure),
        "^deaths?$|mortality"
      ) ~ "Deaths",
      
      str_detect(
        str_to_lower(measure),
        "^dalys?$|disability.?adjusted"
      ) ~ "DALYs",
      
      TRUE ~ measure
      
    ),
    
    age_standardised = str_detect(
      str_to_lower(age),
      "age.?standard|standardized|standardised"
    ),
    
    all_ages = str_detect(
      str_to_lower(age),
      "^all ages?$"
    ),
    
    analytical_metric = case_when(
      
      str_to_lower(metric) %in% c(
        "number",
        "count",
        "counts"
      ) &
        all_ages ~ "Number",
      
      str_to_lower(metric) %in% c(
        "number",
        "count",
        "counts"
      ) ~ "Number",
      
      str_to_lower(metric) == "rate" &
        age_standardised ~ "Age-standardized rate",
      
      str_detect(
        str_to_lower(metric),
        "age.?standard"
      ) ~ "Age-standardized rate",
      
      TRUE ~ NA_character_
      
    )
    
  )


# ==============================================================================
# 9. FILTER THE REQUIRED 2023 GLOBAL ESTIMATES
# ==============================================================================

required_sexes <- c(
  "Both",
  "Male",
  "Female"
)

required_measures <- c(
  "Incidence",
  "Deaths",
  "DALYs"
)

required_metrics <- c(
  "Number",
  "Age-standardized rate"
)


global_2023 <- global_data |>
  
  filter(
    
    year == 2023,
    
    sex %in% required_sexes,
    
    measure %in% required_measures,
    
    analytical_metric %in% required_metrics
    
  ) |>
  
  transmute(
    
    sex,
    
    measure,
    
    metric = analytical_metric,
    
    val,
    
    lower,
    
    upper
    
  )


# ==============================================================================
# 10. CHECK REQUIRED 2023 COMBINATIONS
#
# Expected:
#   3 sexes × 3 measures × 2 metrics = 18 rows
# ==============================================================================

required_combinations <- tidyr::expand_grid(
  sex = required_sexes,
  measure = required_measures,
  metric = required_metrics
)

available_combinations <- global_2023 |>
  distinct(
    sex,
    measure,
    metric
  )

missing_combinations <- anti_join(
  required_combinations,
  available_combinations,
  by = c(
    "sex",
    "measure",
    "metric"
  )
)

if (nrow(missing_combinations) > 0) {
  
  print(missing_combinations)
  
  stop(
    "\nSome required 2023 estimates are missing. They are printed above.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 11. CHECK FOR DUPLICATES
# ==============================================================================

duplicate_combinations <- global_2023 |>
  
  count(
    sex,
    measure,
    metric,
    name = "number_of_rows"
  ) |>
  
  filter(
    number_of_rows > 1
  )

if (nrow(duplicate_combinations) > 0) {
  
  print(duplicate_combinations)
  
  stop(
    "\nDuplicate estimates were detected and printed above.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 12. CHECK ESTIMATES AND UNCERTAINTY INTERVALS
# ==============================================================================

invalid_estimates <- global_2023 |>
  
  filter(
    
    is.na(val) |
      
      is.na(lower) |
      
      is.na(upper) |
      
      lower < 0 |
      
      lower > val |
      
      val > upper
    
  )

if (nrow(invalid_estimates) > 0) {
  
  print(invalid_estimates)
  
  stop(
    "\nInvalid or incomplete uncertainty intervals were detected.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 13. ORDER THE FINAL ANALYTICAL DATA
# ==============================================================================

global_2023 <- global_2023 |>
  
  mutate(
    
    sex_order = match(
      sex,
      required_sexes
    ),
    
    measure_order = match(
      measure,
      required_measures
    ),
    
    metric_order = match(
      metric,
      required_metrics
    )
    
  ) |>
  
  arrange(
    sex_order,
    measure_order,
    metric_order
  ) |>
  
  select(
    -sex_order,
    -measure_order,
    -metric_order
  )


# ==============================================================================
# 14. SEPARATE NUMBER AND RATE ESTIMATES
# ==============================================================================

number_data <- global_2023 |>
  
  filter(
    metric == "Number"
  ) |>
  
  select(
    sex,
    measure,
    val,
    lower,
    upper
  )


rate_data <- global_2023 |>
  
  filter(
    metric == "Age-standardized rate"
  ) |>
  
  select(
    sex,
    measure,
    val,
    lower,
    upper
  )


# ==============================================================================
# 15. CREATE THE WIDE RESULTS DATASET
# ==============================================================================

number_wide <- number_data |>
  
  pivot_wider(
    
    names_from = measure,
    
    values_from = c(
      val,
      lower,
      upper
    ),
    
    names_glue = "{measure}_{.value}"
    
  )


rate_wide <- rate_data |>
  
  pivot_wider(
    
    names_from = measure,
    
    values_from = c(
      val,
      lower,
      upper
    ),
    
    names_glue = "{measure}_ASR_{.value}"
    
  )


global_results <- number_wide |>
  
  left_join(
    rate_wide,
    by = "sex"
  ) |>
  
  mutate(
    
    sex_order = match(
      sex,
      required_sexes
    )
    
  ) |>
  
  arrange(
    sex_order
  ) |>
  
  select(
    -sex_order
  )


# ==============================================================================
# 16. FORMATTING FUNCTIONS
# ==============================================================================

format_ui <- function(
    estimate,
    lower,
    upper,
    digits
) {
  
  paste0(
    
    formatC(
      estimate,
      format = "f",
      digits = digits,
      big.mark = ","
    ),
    
    " (",
    
    formatC(
      lower,
      format = "f",
      digits = digits,
      big.mark = ","
    ),
    
    "–",
    
    formatC(
      upper,
      format = "f",
      digits = digits,
      big.mark = ","
    ),
    
    ")"
    
  )
  
}


format_result <- function(
    value,
    digits = 1
) {
  
  formatC(
    value,
    format = "f",
    digits = digits,
    big.mark = ","
  )
  
}


# ==============================================================================
# 17. CREATE TABLE 1
# ==============================================================================

table_1 <- global_results |>
  
  transmute(
    
    Sex = sex,
    
    `Incident cases, n (95% UI)` = format_ui(
      Incidence_val,
      Incidence_lower,
      Incidence_upper,
      digits = 0
    ),
    
    `Age-standardized incidence rate per 100,000 (95% UI)` = format_ui(
      Incidence_ASR_val,
      Incidence_ASR_lower,
      Incidence_ASR_upper,
      digits = 2
    ),
    
    `Deaths, n (95% UI)` = format_ui(
      Deaths_val,
      Deaths_lower,
      Deaths_upper,
      digits = 0
    ),
    
    `Age-standardized mortality rate per 100,000 (95% UI)` = format_ui(
      Deaths_ASR_val,
      Deaths_ASR_lower,
      Deaths_ASR_upper,
      digits = 2
    ),
    
    `DALYs, n (95% UI)` = format_ui(
      DALYs_val,
      DALYs_lower,
      DALYs_upper,
      digits = 0
    ),
    
    `Age-standardized DALY rate per 100,000 (95% UI)` = format_ui(
      DALYs_ASR_val,
      DALYs_ASR_lower,
      DALYs_ASR_upper,
      digits = 2
    )
    
  )


# ==============================================================================
# 18. SAVE TABLE 1
# ==============================================================================

table_1_file <- file.path(
  tables_folder,
  "Table_1_Global_CRC_burden_2023.xlsx"
)

table_workbook <- createWorkbook()

addWorksheet(
  table_workbook,
  "Table 1"
)

writeData(
  table_workbook,
  sheet = "Table 1",
  x = "Table 1. Global colorectal cancer burden by sex in 2023",
  startRow = 1,
  startCol = 1
)

writeData(
  table_workbook,
  sheet = "Table 1",
  x = table_1,
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
  sheet = "Table 1",
  style = title_style,
  rows = 1,
  cols = 1:ncol(table_1),
  gridExpand = TRUE
)

addStyle(
  table_workbook,
  sheet = "Table 1",
  style = header_style,
  rows = 3,
  cols = 1:ncol(table_1),
  gridExpand = TRUE
)

addStyle(
  table_workbook,
  sheet = "Table 1",
  style = body_style,
  rows = 4:(nrow(table_1) + 3),
  cols = 1:ncol(table_1),
  gridExpand = TRUE
)

setColWidths(
  table_workbook,
  sheet = "Table 1",
  cols = 1,
  widths = 14
)

setColWidths(
  table_workbook,
  sheet = "Table 1",
  cols = 2:ncol(table_1),
  widths = 27
)

setRowHeights(
  table_workbook,
  sheet = "Table 1",
  rows = 3,
  heights = 55
)

freezePane(
  table_workbook,
  sheet = "Table 1",
  firstActiveRow = 4,
  firstActiveCol = 2
)

writeData(
  table_workbook,
  sheet = "Table 1",
  x = paste(
    "Note: Values are estimates with 95% uncertainty intervals (UIs).",
    "Age-standardized rates are expressed per 100,000 population."
  ),
  startRow = nrow(table_1) + 6,
  startCol = 1,
  colNames = FALSE
)

saveWorkbook(
  table_workbook,
  table_1_file,
  overwrite = TRUE
)


# ==============================================================================
# 19. CREATE AND SAVE SUPPLEMENTARY TABLE S1
# ==============================================================================

supplementary_table <- global_2023 |>
  
  rename(
    
    Sex = sex,
    
    Measure = measure,
    
    Metric = metric,
    
    Estimate = val,
    
    Lower_95_UI = lower,
    
    Upper_95_UI = upper
    
  )


supplementary_file <- file.path(
  supplementary_folder,
  "Table_S1_Global_CRC_detailed_estimates_2023.xlsx"
)


supplementary_workbook <- createWorkbook()

addWorksheet(
  supplementary_workbook,
  "Detailed estimates"
)

writeData(
  supplementary_workbook,
  sheet = "Detailed estimates",
  x = supplementary_table
)

addStyle(
  supplementary_workbook,
  sheet = "Detailed estimates",
  style = header_style,
  rows = 1,
  cols = 1:ncol(supplementary_table),
  gridExpand = TRUE
)

setColWidths(
  supplementary_workbook,
  sheet = "Detailed estimates",
  cols = 1:ncol(supplementary_table),
  widths = "auto"
)

freezePane(
  supplementary_workbook,
  sheet = "Detailed estimates",
  firstRow = TRUE
)

saveWorkbook(
  supplementary_workbook,
  supplementary_file,
  overwrite = TRUE
)


# ==============================================================================
# 20. PREPARE FIGURE DATA
# ==============================================================================

sex_levels <- c(
  "Both",
  "Male",
  "Female"
)

plot_numbers <- number_data |>
  
  mutate(
    
    sex = factor(
      sex,
      levels = sex_levels
    )
    
  )


plot_rates <- rate_data |>
  
  mutate(
    
    sex = factor(
      sex,
      levels = sex_levels
    )
    
  )


# ==============================================================================
# 21. FUNCTION FOR NUMBER PANELS
# ==============================================================================

make_number_panel <- function(
    selected_measure,
    panel_letter,
    panel_title,
    y_title
) {
  
  panel_data <- plot_numbers |>
    
    filter(
      measure == selected_measure
    )
  
  
  ggplot(
    panel_data,
    aes(
      x = sex,
      y = val,
      fill = sex
    )
  ) +
    
    geom_col(
      width = 0.66,
      alpha = 0.88
    ) +
    
    geom_errorbar(
      aes(
        ymin = lower,
        ymax = upper
      ),
      width = 0.13,
      linewidth = 0.55
    ) +
    
    scale_y_continuous(
      
      labels = scales::label_number(
        scale_cut = scales::cut_short_scale(),
        accuracy = 0.1
      ),
      
      expand = expansion(
        mult = c(
          0,
          0.12
        )
      )
      
    ) +
    
    labs(
      
      title = paste0(
        panel_letter,
        ". ",
        panel_title
      ),
      
      x = NULL,
      
      y = y_title,
      
      fill = NULL
      
    ) +
    
    theme_classic(
      base_size = 11
    ) +
    
    theme(
      
      legend.position = "none",
      
      plot.title = element_text(
        face = "bold",
        size = 11,
        hjust = 0
      ),
      
      axis.title.y = element_text(
        size = 10
      ),
      
      axis.text.x = element_text(
        size = 9
      ),
      
      axis.text.y = element_text(
        size = 9
      ),
      
      panel.grid.major.y = element_line(
        linewidth = 0.25,
        linetype = "dotted"
      ),
      
      plot.margin = margin(
        8,
        8,
        8,
        8
      )
      
    )
  
}


# ==============================================================================
# 22. FUNCTION FOR AGE-STANDARDIZED RATE PANELS
# ==============================================================================

make_rate_panel <- function(
    selected_measure,
    panel_letter,
    panel_title
) {
  
  panel_data <- plot_rates |>
    
    filter(
      measure == selected_measure
    )
  
  
  ggplot(
    panel_data,
    aes(
      x = sex,
      y = val
    )
  ) +
    
    geom_errorbar(
      aes(
        ymin = lower,
        ymax = upper
      ),
      width = 0.12,
      linewidth = 0.75
    ) +
    
    geom_point(
      aes(
        fill = sex
      ),
      shape = 21,
      size = 4.2,
      stroke = 0.8
    ) +
    
    scale_y_continuous(
      
      labels = scales::label_number(
        accuracy = 0.1
      ),
      
      expand = expansion(
        mult = c(
          0.08,
          0.14
        )
      )
      
    ) +
    
    labs(
      
      title = paste0(
        panel_letter,
        ". ",
        panel_title
      ),
      
      x = NULL,
      
      y = "Rate per 100,000",
      
      fill = NULL
      
    ) +
    
    theme_classic(
      base_size = 11
    ) +
    
    theme(
      
      legend.position = "none",
      
      plot.title = element_text(
        face = "bold",
        size = 11,
        hjust = 0
      ),
      
      axis.title.y = element_text(
        size = 10
      ),
      
      axis.text.x = element_text(
        size = 9
      ),
      
      axis.text.y = element_text(
        size = 9
      ),
      
      panel.grid.major.y = element_line(
        linewidth = 0.25,
        linetype = "dotted"
      ),
      
      plot.margin = margin(
        8,
        8,
        8,
        8
      )
      
    )
  
}


# ==============================================================================
# 23. CREATE THE SIX FIGURE PANELS
# ==============================================================================

panel_a <- make_number_panel(
  selected_measure = "Incidence",
  panel_letter = "A",
  panel_title = "Incident cases",
  y_title = "Number of cases"
)

panel_b <- make_rate_panel(
  selected_measure = "Incidence",
  panel_letter = "B",
  panel_title = "Age-standardized incidence rate"
)

panel_c <- make_number_panel(
  selected_measure = "Deaths",
  panel_letter = "C",
  panel_title = "Deaths",
  y_title = "Number of deaths"
)

panel_d <- make_rate_panel(
  selected_measure = "Deaths",
  panel_letter = "D",
  panel_title = "Age-standardized mortality rate"
)

panel_e <- make_number_panel(
  selected_measure = "DALYs",
  panel_letter = "E",
  panel_title = "Disability-adjusted life years",
  y_title = "Number of DALYs"
)

panel_f <- make_rate_panel(
  selected_measure = "DALYs",
  panel_letter = "F",
  panel_title = "Age-standardized DALY rate"
)


# ==============================================================================
# 24. COMBINE THE SIX PANELS
# ==============================================================================

figure_1 <- (
  
  panel_a |
    panel_b |
    panel_c
  
) / (
  
  panel_d |
    panel_e |
    panel_f
  
) +
  
  patchwork::plot_annotation(
    
    title = "Global colorectal cancer burden in 2023",
    
    subtitle = paste(
      "Estimates are shown for both sexes combined, males, and females;",
      "error bars indicate 95% uncertainty intervals"
    ),
    
    caption = paste(
      "Age-standardized rates are expressed per 100,000 population.",
      "Both represents both sexes combined."
    ),
    
    theme = theme(
      
      plot.title = element_text(
        face = "bold",
        size = 16,
        hjust = 0.5
      ),
      
      plot.subtitle = element_text(
        size = 10,
        hjust = 0.5
      ),
      
      plot.caption = element_text(
        size = 9,
        hjust = 0
      )
      
    )
    
  )


# ==============================================================================
# 25. SAVE FIGURE 1 AS ONE TIFF AT 600 DPI
# ==============================================================================

figure_1_file <- file.path(
  figures_folder,
  "Figure_1_Global_CRC_burden_2023.tiff"
)

ggsave(
  filename = figure_1_file,
  plot = figure_1,
  device = "tiff",
  width = 15,
  height = 10,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white"
)


# ==============================================================================
# 26. EXTRACT BOTH-SEX, MALE, AND FEMALE RESULTS
# ==============================================================================

both_results <- global_results |>
  
  filter(
    sex == "Both"
  )


male_results <- global_results |>
  
  filter(
    sex == "Male"
  )


female_results <- global_results |>
  
  filter(
    sex == "Female"
  )


# ==============================================================================
# 27. CALCULATE SEX DIFFERENCES IN AGE-STANDARDIZED RATES
# ==============================================================================

incidence_rate_difference <- 100 * (
  
  male_results$Incidence_ASR_val -
    
    female_results$Incidence_ASR_val
  
) / female_results$Incidence_ASR_val


mortality_rate_difference <- 100 * (
  
  male_results$Deaths_ASR_val -
    
    female_results$Deaths_ASR_val
  
) / female_results$Deaths_ASR_val


daly_rate_difference <- 100 * (
  
  male_results$DALYs_ASR_val -
    
    female_results$DALYs_ASR_val
  
) / female_results$DALYs_ASR_val


# ==============================================================================
# 28. GENERATE THE RESULTS TEXT
# ==============================================================================

results_text <- glue::glue(
  "3.1 Global colorectal cancer burden in 2023

In 2023, colorectal cancer accounted for {format_result(both_results$Incidence_val, 0)} incident cases globally (95% UI {format_result(both_results$Incidence_lower, 0)}–{format_result(both_results$Incidence_upper, 0)}). The corresponding age-standardized incidence rate was {format_result(both_results$Incidence_ASR_val, 2)} per 100,000 population (95% UI {format_result(both_results$Incidence_ASR_lower, 2)}–{format_result(both_results$Incidence_ASR_upper, 2)}).

Colorectal cancer caused {format_result(both_results$Deaths_val, 0)} deaths globally (95% UI {format_result(both_results$Deaths_lower, 0)}–{format_result(both_results$Deaths_upper, 0)}), with an age-standardized mortality rate of {format_result(both_results$Deaths_ASR_val, 2)} per 100,000 population (95% UI {format_result(both_results$Deaths_ASR_lower, 2)}–{format_result(both_results$Deaths_ASR_upper, 2)}).

The disease was responsible for {format_result(both_results$DALYs_val, 0)} DALYs worldwide (95% UI {format_result(both_results$DALYs_lower, 0)}–{format_result(both_results$DALYs_upper, 0)}). The global age-standardized DALY rate was {format_result(both_results$DALYs_ASR_val, 2)} per 100,000 population (95% UI {format_result(both_results$DALYs_ASR_lower, 2)}–{format_result(both_results$DALYs_ASR_upper, 2)}).

Sex differences were consistently observed across incidence, mortality, and DALYs. The age-standardized incidence rate was {format_result(male_results$Incidence_ASR_val, 2)} per 100,000 among males and {format_result(female_results$Incidence_ASR_val, 2)} among females, corresponding to a {format_result(abs(incidence_rate_difference), 1)}% higher rate among males. The age-standardized mortality rate was {format_result(male_results$Deaths_ASR_val, 2)} among males and {format_result(female_results$Deaths_ASR_val, 2)} among females, representing a {format_result(abs(mortality_rate_difference), 1)}% higher rate among males. Similarly, the age-standardized DALY rate was {format_result(male_results$DALYs_ASR_val, 2)} per 100,000 among males compared with {format_result(female_results$DALYs_ASR_val, 2)} among females, a difference of {format_result(abs(daly_rate_difference), 1)}%.

Complete global estimates by sex are presented in Table 1. Figure 1 separately displays the global numbers and age-standardized rates for incidence, mortality, and DALYs."
)


# ==============================================================================
# 29. SAVE THE RESULTS TEXT
# ==============================================================================

results_text_file <- file.path(
  results_text_folder,
  "Results_3_1_Global_CRC_burden_2023.txt"
)

writeLines(
  text = results_text,
  con = results_text_file,
  useBytes = TRUE
)


# ==============================================================================
# 30. DISPLAY RESULTS IN R
# ==============================================================================

message("")
message("Table 1:")

print(
  table_1
)

message("")
message("Results text:")

cat(
  results_text
)


# ==============================================================================
# 31. COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("RESULTS SECTION 3.1 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message("Table 1:")
message(table_1_file)

message("")
message("Figure 1 — six panels, one TIFF at 600 dpi:")
message(figure_1_file)

message("")
message("Supplementary Table S1:")
message(supplementary_file)

message("")
message("Results text:")
message(results_text_file)

message("")
message("Included measures: Incidence, Deaths, and DALYs.")
message("The old clean-data folder was not used.")
message("==============================================================")