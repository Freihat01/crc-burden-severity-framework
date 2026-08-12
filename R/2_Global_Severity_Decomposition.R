# ==============================================================================
# R_03_2_Global_Severity_Decomposition.R
#
# RESULTS SECTION 3.2
# Global severity decomposition of colorectal cancer in 2023
#
# RAW INPUT FOLDERS
#   Incidence/
#   DALYs/
#   YLLs/
#   YLDs/
#
# Each folder contains the country-level raw CSV files.
#
# OUTPUTS
# Publication/Tables/
#   Table_2_Global_CRC_severity_decomposition_2023.xlsx
#
# Publication/Figures/
#   Figure_2_Global_CRC_severity_decomposition_2023.tiff
#
# Publication/Results_text/
#   Results_3_2_Global_CRC_severity_decomposition_2023.txt
#
# ANALYSES
#   DALYs per case
#   YLLs per case
#   YLDs per case
#   YLL contribution (%)
#   YLD contribution (%)
#
# FIGURE
# One 600-dpi TIFF showing the decomposition of DALYs per case into:
#   YLLs per case
#   YLDs per case
#
# The old Clean_Data folder is not used.
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
  library(scales)
  library(glue)
  
})


# ==============================================================================
# 3. IDENTIFY THE MAIN PROJECT FOLDER
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

figures_folder <- file.path(
  publication_folder,
  "Figures"
)

tables_folder <- file.path(
  publication_folder,
  "Tables"
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
  results_text_folder,
  recursive = TRUE,
  showWarnings = FALSE
)


# ==============================================================================
# 5. GENERAL HELPER FUNCTIONS
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


find_column <- function(
    data,
    accepted_names,
    variable_name,
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
  
  detected[[1]]
  
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


is_number_metric <- function(metric_value) {
  
  stringr::str_to_lower(
    clean_text(metric_value)
  ) %in% c(
    "number",
    "count",
    "counts"
  )
  
}


is_all_ages <- function(age_value) {
  
  stringr::str_detect(
    stringr::str_to_lower(
      clean_text(age_value)
    ),
    "^all ages?$"
  )
  
}


format_number <- function(
    value,
    digits = 0
) {
  
  formatC(
    value,
    format = "f",
    digits = digits,
    big.mark = ","
  )
  
}


format_percentage <- function(
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
# 6. FUNCTION TO IMPORT ALL CSV FILES FROM ONE MEASURE FOLDER
# ==============================================================================

read_measure_folder <- function(
    folder_path,
    measure_label
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
    measure_label,
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
  
  raw_measure_data <- bind_rows(
    imported_files
  )
  
  column_year <- find_column(
    raw_measure_data,
    c(
      "year",
      "year_id"
    ),
    "year"
  )
  
  column_sex <- find_column(
    raw_measure_data,
    c(
      "sex",
      "sex_name"
    ),
    "sex"
  )
  
  column_metric <- find_column(
    raw_measure_data,
    c(
      "metric",
      "metric_name"
    ),
    "metric"
  )
  
  column_age <- find_column(
    raw_measure_data,
    c(
      "age",
      "age_name"
    ),
    "age"
  )
  
  column_location <- find_column(
    raw_measure_data,
    c(
      "location",
      "location_name"
    ),
    "location"
  )
  
  column_value <- find_column(
    raw_measure_data,
    c(
      "val",
      "value",
      "mean",
      "estimate"
    ),
    "central estimate"
  )
  
  column_cause <- find_column(
    raw_measure_data,
    c(
      "cause",
      "cause_name"
    ),
    "cause",
    required = FALSE
  )
  
  standardised_data <- raw_measure_data |>
    
    transmute(
      
      location = clean_text(
        .data[[column_location]]
      ),
      
      year = as.integer(
        .data[[column_year]]
      ),
      
      sex = standardise_sex(
        .data[[column_sex]]
      ),
      
      metric = clean_text(
        .data[[column_metric]]
      ),
      
      age = clean_text(
        .data[[column_age]]
      ),
      
      cause = if (!is.null(column_cause)) {
        
        clean_text(
          .data[[column_cause]]
        )
        
      } else {
        
        "Colorectal cancer"
        
      },
      
      value = as.numeric(
        .data[[column_value]]
      ),
      
      source_file = source_file
      
    ) |>
    
    filter(
      
      year == 2023,
      
      sex %in% c(
        "Both",
        "Male",
        "Female"
      ),
      
      is_number_metric(metric),
      
      is_all_ages(age),
      
      !is.na(location),
      
      !is.na(value)
      
    )
  
  if (!is.null(column_cause)) {
    
    standardised_data <- standardised_data |>
      
      filter(
        
        stringr::str_detect(
          
          stringr::str_to_lower(cause),
          
          "colon and rectum cancer|colorectal cancer|colon.*rectum"
          
        )
        
      )
    
  }
  
  standardised_data <- standardised_data |>
    
    filter(
      
      !stringr::str_to_lower(location) %in% c(
        "global",
        "world",
        "worldwide"
      )
      
    ) |>
    
    mutate(
      measure = measure_label
    ) |>
    
    select(
      location,
      sex,
      measure,
      value,
      source_file
    )
  
  duplicate_rows <- standardised_data |>
    
    count(
      location,
      sex,
      measure,
      name = "number_of_rows"
    ) |>
    
    filter(
      number_of_rows > 1
    )
  
  if (nrow(duplicate_rows) > 0) {
    
    print(duplicate_rows)
    
    stop(
      paste0(
        "\nDuplicate country estimates were detected for ",
        measure_label,
        "."
      ),
      call. = FALSE
    )
    
  }
  
  standardised_data
  
}


# ==============================================================================
# 7. IMPORT THE FOUR REQUIRED MEASURES
# ==============================================================================

incidence_data <- read_measure_folder(
  folder_path = incidence_folder,
  measure_label = "Incidence"
)

dalys_data <- read_measure_folder(
  folder_path = dalys_folder,
  measure_label = "DALYs"
)

ylls_data <- read_measure_folder(
  folder_path = ylls_folder,
  measure_label = "YLLs"
)

ylds_data <- read_measure_folder(
  folder_path = ylds_folder,
  measure_label = "YLDs"
)


# ==============================================================================
# 8. IDENTIFY THE COMMON COUNTRY SET
#
# Only locations present in all four measures are retained.
# ==============================================================================

incidence_locations <- unique(
  incidence_data$location
)

dalys_locations <- unique(
  dalys_data$location
)

ylls_locations <- unique(
  ylls_data$location
)

ylds_locations <- unique(
  ylds_data$location
)

common_locations <- Reduce(
  intersect,
  list(
    incidence_locations,
    dalys_locations,
    ylls_locations,
    ylds_locations
  )
)

message("")
message(
  "Locations available in all four measures: ",
  length(common_locations)
)

if (length(common_locations) == 0) {
  
  stop(
    "\nNo common locations were found across the four measures.",
    call. = FALSE
  )
  
}


# ==============================================================================
# 9. RETAIN THE COMMON LOCATIONS
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
# 10. CHECK THAT EACH SEX HAS THE SAME LOCATION COVERAGE
# ==============================================================================

coverage_summary <- bind_rows(
  incidence_data,
  dalys_data,
  ylls_data,
  ylds_data
) |>
  
  distinct(
    location,
    sex,
    measure
  ) |>
  
  count(
    sex,
    measure,
    name = "number_of_locations"
  )

print(
  coverage_summary
)

expected_location_count <- length(
  common_locations
)

incomplete_coverage <- coverage_summary |>
  
  filter(
    number_of_locations != expected_location_count
  )

if (nrow(incomplete_coverage) > 0) {
  
  print(
    incomplete_coverage
  )
  
  stop(
    paste0(
      "\nCountry coverage is inconsistent across sexes or measures."
    ),
    call. = FALSE
  )
  
}


# ==============================================================================
# 11. COMBINE ALL MEASURES
# ==============================================================================

combined_country_data <- bind_rows(
  incidence_data,
  dalys_data,
  ylls_data,
  ylds_data
) |>
  
  select(
    location,
    sex,
    measure,
    value
  )


# ==============================================================================
# 12. CALCULATE GLOBAL TOTALS FROM COUNTRY ESTIMATES
# ==============================================================================

global_totals <- combined_country_data |>
  
  group_by(
    sex,
    measure
  ) |>
  
  summarise(
    
    global_value = sum(
      value,
      na.rm = TRUE
    ),
    
    number_of_locations = n_distinct(
      location
    ),
    
    .groups = "drop"
    
  )


# ==============================================================================
# 13. CONVERT GLOBAL TOTALS TO WIDE FORMAT
# ==============================================================================

global_totals_wide <- global_totals |>
  
  select(
    sex,
    measure,
    global_value
  ) |>
  
  pivot_wider(
    
    names_from = measure,
    
    values_from = global_value
    
  ) |>
  
  mutate(
    
    sex_order = match(
      sex,
      c(
        "Both",
        "Male",
        "Female"
      )
    )
    
  ) |>
  
  arrange(
    sex_order
  ) |>
  
  select(
    -sex_order
  )


# ==============================================================================
# 14. CALCULATE GLOBAL SEVERITY INDICATORS
# ==============================================================================

global_severity <- global_totals_wide |>
  
  mutate(
    
    DALYs_per_case =
      DALYs / Incidence,
    
    YLLs_per_case =
      YLLs / Incidence,
    
    YLDs_per_case =
      YLDs / Incidence,
    
    YLL_percent =
      100 * YLLs / DALYs,
    
    YLD_percent =
      100 * YLDs / DALYs,
    
    component_sum =
      YLLs_per_case + YLDs_per_case,
    
    decomposition_difference =
      DALYs_per_case - component_sum,
    
    decomposition_relative_difference =
      100 *
      abs(decomposition_difference) /
      DALYs_per_case
    
  )


# ==============================================================================
# 15. CHECK THE DALY DECOMPOSITION
#
# DALYs per case should equal:
#
# YLLs per case + YLDs per case
# ==============================================================================

maximum_relative_difference <- max(
  global_severity$decomposition_relative_difference,
  na.rm = TRUE
)

message("")
message(
  "Maximum decomposition difference: ",
  formatC(
    maximum_relative_difference,
    format = "f",
    digits = 4
  ),
  "%"
)

if (maximum_relative_difference > 1) {
  
  warning(
    paste0(
      "DALYs differ from YLLs + YLDs by more than 1%. ",
      "Review source-file consistency."
    )
  )
  
}


# ==============================================================================
# 16. CREATE TABLE 2
# ==============================================================================

table_2 <- global_severity |>
  
  transmute(
    
    Sex = sex,
    
    `DALYs per case` = round(
      DALYs_per_case,
      2
    ),
    
    `YLLs per case` = round(
      YLLs_per_case,
      2
    ),
    
    `YLDs per case` = round(
      YLDs_per_case,
      2
    ),
    
    `YLL contribution (%)` = round(
      YLL_percent,
      2
    ),
    
    `YLD contribution (%)` = round(
      YLD_percent,
      2
    )
    
  )


# ==============================================================================
# 17. SAVE TABLE 2
# ==============================================================================

table_2_file <- file.path(
  tables_folder,
  "Table_2_Global_CRC_severity_decomposition_2023.xlsx"
)

table_workbook <- createWorkbook()

addWorksheet(
  table_workbook,
  "Table 2"
)

writeData(
  table_workbook,
  sheet = "Table 2",
  x = paste(
    "Table 2. Global colorectal cancer severity",
    "decomposition by sex in 2023"
  ),
  startRow = 1,
  startCol = 1
)

writeData(
  table_workbook,
  sheet = "Table 2",
  x = table_2,
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
  valign = "center"
)

addStyle(
  table_workbook,
  sheet = "Table 2",
  style = title_style,
  rows = 1,
  cols = 1:ncol(table_2),
  gridExpand = TRUE
)

addStyle(
  table_workbook,
  sheet = "Table 2",
  style = header_style,
  rows = 3,
  cols = 1:ncol(table_2),
  gridExpand = TRUE
)

addStyle(
  table_workbook,
  sheet = "Table 2",
  style = body_style,
  rows = 4:(nrow(table_2) + 3),
  cols = 1:ncol(table_2),
  gridExpand = TRUE
)

setColWidths(
  table_workbook,
  sheet = "Table 2",
  cols = 1,
  widths = 14
)

setColWidths(
  table_workbook,
  sheet = "Table 2",
  cols = 2:ncol(table_2),
  widths = 22
)

setRowHeights(
  table_workbook,
  sheet = "Table 2",
  rows = 3,
  heights = 40
)

freezePane(
  table_workbook,
  sheet = "Table 2",
  firstActiveRow = 4,
  firstActiveCol = 2
)

writeData(
  table_workbook,
  sheet = "Table 2",
  x = paste(
    "Note: Global estimates were reconstructed by summing",
    length(common_locations),
    "country and territory estimates available consistently across",
    "incidence, DALYs, YLLs, and YLDs.",
    "DALYs per case = DALYs divided by incident cases;",
    "YLLs per case = YLLs divided by incident cases;",
    "YLDs per case = YLDs divided by incident cases."
  ),
  startRow = nrow(table_2) + 6,
  startCol = 1,
  colNames = FALSE
)

saveWorkbook(
  table_workbook,
  table_2_file,
  overwrite = TRUE
)


# ==============================================================================
# 18. PREPARE FIGURE 2 DATA
# ==============================================================================

figure_data <- global_severity |>
  
  select(
    sex,
    DALYs_per_case,
    YLLs_per_case,
    YLDs_per_case,
    YLL_percent,
    YLD_percent
  ) |>
  
  pivot_longer(
    
    cols = c(
      YLLs_per_case,
      YLDs_per_case
    ),
    
    names_to = "component",
    
    values_to = "value"
    
  ) |>
  
  mutate(
    
    Sex = factor(
      sex,
      levels = c(
        "Both",
        "Male",
        "Female"
      )
    ),
    
    Component = case_when(
      
      component == "YLLs_per_case" ~
        "Years of life lost per case",
      
      component == "YLDs_per_case" ~
        "Years lived with disability per case"
      
    ),
    
    Component = factor(
      
      Component,
      
      levels = c(
        "Years of life lost per case",
        "Years lived with disability per case"
      )
      
    ),
    
    percentage = case_when(
      
      component == "YLLs_per_case" ~
        YLL_percent,
      
      component == "YLDs_per_case" ~
        YLD_percent
      
    ),
    
    percentage_label = paste0(
      formatC(
        percentage,
        format = "f",
        digits = 1
      ),
      "%"
    )
    
  )


total_labels <- global_severity |>
  
  transmute(
    
    Sex = factor(
      sex,
      levels = c(
        "Both",
        "Male",
        "Female"
      )
    ),
    
    DALYs_per_case,
    
    label = paste0(
      "Total = ",
      formatC(
        DALYs_per_case,
        format = "f",
        digits = 2
      )
    )
    
  )


# ==============================================================================
# 19. CREATE FIGURE 2
# ==============================================================================

figure_2 <- ggplot(
  
  figure_data,
  
  aes(
    x = Sex,
    y = value,
    fill = Component
  )
  
) +
  
  geom_col(
    width = 0.66
  ) +
  
  geom_text(
    
    aes(
      label = percentage_label
    ),
    
    position = position_stack(
      vjust = 0.5
    ),
    
    size = 4,
    
    fontface = "bold"
    
  ) +
  
  geom_text(
    
    data = total_labels,
    
    aes(
      x = Sex,
      y = DALYs_per_case,
      label = label
    ),
    
    inherit.aes = FALSE,
    
    vjust = -0.7,
    
    size = 4,
    
    fontface = "bold"
    
  ) +
  
  scale_y_continuous(
    
    name = "DALYs per incident case",
    
    expand = expansion(
      mult = c(
        0,
        0.15
      )
    ),
    
    labels = scales::label_number(
      accuracy = 0.1
    )
    
  ) +
  
  labs(
    
    title = paste(
      "Global colorectal cancer severity decomposition",
      "in 2023"
    ),
    
    subtitle = paste(
      "DALYs per case decomposed into premature-mortality",
      "and disability components"
    ),
    
    x = NULL,
    
    fill = NULL,
    
    caption = paste(
      "Percentages represent each component's contribution",
      "to total DALYs per case."
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
      size = 11,
      hjust = 0.5
    ),
    
    axis.title.y = element_text(
      size = 11
    ),
    
    axis.text.x = element_text(
      size = 11
    ),
    
    axis.text.y = element_text(
      size = 10
    ),
    
    legend.position = "bottom",
    
    legend.text = element_text(
      size = 10
    ),
    
    plot.caption = element_text(
      size = 9,
      hjust = 0
    ),
    
    panel.grid.major.y = element_line(
      linewidth = 0.25,
      linetype = "dotted"
    ),
    
    plot.margin = margin(
      12,
      20,
      12,
      12
    )
    
  )


# ==============================================================================
# 20. SAVE FIGURE 2 AS ONE TIFF AT 600 DPI
# ==============================================================================

figure_2_file <- file.path(
  figures_folder,
  "Figure_2_Global_CRC_severity_decomposition_2023.tiff"
)

ggsave(
  filename = figure_2_file,
  plot = figure_2,
  device = "tiff",
  width = 9,
  height = 7,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white"
)


# ==============================================================================
# 21. EXTRACT RESULTS BY SEX
# ==============================================================================

both_results <- global_severity |>
  
  filter(
    sex == "Both"
  )

male_results <- global_severity |>
  
  filter(
    sex == "Male"
  )

female_results <- global_severity |>
  
  filter(
    sex == "Female"
  )


# ==============================================================================
# 22. GENERATE RESULTS TEXT
# ==============================================================================

results_text <- glue::glue(
  "3.2 Global severity decomposition

Globally, each incident colorectal cancer case in 2023 was associated with {format_number(both_results$DALYs_per_case, 2)} DALYs. This total comprised {format_number(both_results$YLLs_per_case, 2)} years of life lost and {format_number(both_results$YLDs_per_case, 2)} years lived with disability per case (Table 2). Premature mortality therefore accounted for {format_number(both_results$YLL_percent, 2)}% of the global health loss attributable to colorectal cancer, whereas disability accounted for {format_number(both_results$YLD_percent, 2)}% (Figure 2).

Among males, colorectal cancer generated {format_number(male_results$DALYs_per_case, 2)} DALYs per incident case, including {format_number(male_results$YLLs_per_case, 2)} YLLs and {format_number(male_results$YLDs_per_case, 2)} YLDs per case. YLLs represented {format_number(male_results$YLL_percent, 2)}% of total DALYs per case, compared with {format_number(male_results$YLD_percent, 2)}% attributable to YLDs.

Among females, each incident case was associated with {format_number(female_results$DALYs_per_case, 2)} DALYs, comprising {format_number(female_results$YLLs_per_case, 2)} YLLs and {format_number(female_results$YLDs_per_case, 2)} YLDs per case. Premature mortality contributed {format_number(female_results$YLL_percent, 2)}% of the total severity estimate, while disability contributed {format_number(female_results$YLD_percent, 2)}% (Table 2; Figure 2).

These findings demonstrate that the global severity of colorectal cancer was predominantly driven by premature mortality, with disability contributing a substantially smaller component of total health loss per incident case."
)


# ==============================================================================
# 23. SAVE RESULTS TEXT
# ==============================================================================

results_text_file <- file.path(
  results_text_folder,
  "Results_3_2_Global_CRC_severity_decomposition_2023.txt"
)

writeLines(
  text = results_text,
  con = results_text_file,
  useBytes = TRUE
)


# ==============================================================================
# 24. DISPLAY RESULTS
# ==============================================================================

message("")
message("Table 2:")

print(
  table_2
)

message("")
message("Results text:")

cat(
  results_text
)


# ==============================================================================
# 25. COMPLETION MESSAGE
# ==============================================================================

message("")
message("")
message("==============================================================")
message("RESULTS SECTION 3.2 COMPLETED SUCCESSFULLY")
message("==============================================================")

message("")
message(
  "Common countries and territories included: ",
  length(common_locations)
)

message("")
message("Table 2:")
message(table_2_file)

message("")
message("Figure 2 — one TIFF at 600 dpi:")
message(figure_2_file)

message("")
message("Results text:")
message(results_text_file)

message("")
message("The old Clean_Data folder was not used.")
message("==============================================================")