# ==============================================================================
# R_03_8C_Hybrid_Figure_8_REVISED.R
#
# FINAL REVISED FIGURE 8
#
# LEFT:
#   Scenario 1 versus Scenario 2 dumbbell plots
#   5 largest positive and 5 largest negative absolute differences
#
# RIGHT:
#   Percentage difference maps:
#   100 * (Scenario 2 - Scenario 1) / Scenario 1
#
# INPUT:
#   Publication/Analysis_objects/
#   Section_3_8_Country_CRC_projections_2024_2050.rds
#
# OUTPUT:
#   Publication/Figures/
#   Figure_8_Country_CRC_projections_2050.tiff
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
  dplyr.summarise.inform = FALSE,
  warn = 1
)


# ==============================================================================
# 2. PACKAGES
# ==============================================================================

required_packages <- c(
  "tidyverse",
  "sf",
  "rnaturalearth",
  "rnaturalearthdata",
  "patchwork",
  "scales"
)

installed_packages <- rownames(installed.packages())

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
  library(sf)
  library(rnaturalearth)
  library(rnaturalearthdata)
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
      "\nThe colorectal cancer project root could not be identified.\n\n",
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
# 4. FILE PATHS
# ==============================================================================

publication_folder <- file.path(
  project_root,
  "Publication"
)

analysis_objects_folder <- file.path(
  publication_folder,
  "Analysis_objects"
)

figures_folder <- file.path(
  publication_folder,
  "Figures"
)

analysis_objects_file <- file.path(
  analysis_objects_folder,
  "Section_3_8_Country_CRC_projections_2024_2050.rds"
)

figure_8_file <- file.path(
  figures_folder,
  "Figure_8_Country_CRC_projections_2050.tiff"
)

figure_8_source_file <- file.path(
  figures_folder,
  "Figure_8_Source_data.csv"
)

dir.create(
  figures_folder,
  recursive = TRUE,
  showWarnings = FALSE
)

if (!file.exists(analysis_objects_file)) {
  
  stop(
    paste0(
      "\nThe Section 3.8 analytical object was not found:\n",
      analysis_objects_file
    ),
    call. = FALSE
  )
}


# ==============================================================================
# 5. LOAD SECTION 3.8 RESULTS
# ==============================================================================

section_3_8_objects <- readRDS(
  analysis_objects_file
)

if (!"projection_2050" %in% names(section_3_8_objects)) {
  
  stop(
    "\nThe saved RDS file does not contain projection_2050.",
    call. = FALSE
  )
}

projection_2050 <- section_3_8_objects$projection_2050


# ==============================================================================
# 6. VALIDATE AND PREPARE DATA
# ==============================================================================

required_columns <- c(
  "country",
  "outcome",
  "scenario_1",
  "scenario_2"
)

missing_columns <- setdiff(
  required_columns,
  names(projection_2050)
)

if (length(missing_columns) > 0) {
  
  stop(
    paste0(
      "\nRequired variables missing:\n",
      paste(missing_columns, collapse = "\n")
    ),
    call. = FALSE
  )
}

projection_2050 <- projection_2050 |>
  transmute(
    country = as.character(country),
    outcome = as.character(outcome),
    scenario_1 = as.numeric(scenario_1),
    scenario_2 = as.numeric(scenario_2),
    absolute_difference = scenario_2 - scenario_1,
    percentage_difference =
      100 * (scenario_2 - scenario_1) / scenario_1
  )

if (nrow(projection_2050) != 204 * 3) {
  
  stop(
    paste0(
      "\nExpected 612 country-outcome observations, but ",
      nrow(projection_2050),
      " were found."
    ),
    call. = FALSE
  )
}

invalid_values <- projection_2050 |>
  filter(
    is.na(scenario_1) |
      is.na(scenario_2) |
      is.na(absolute_difference) |
      is.na(percentage_difference) |
      !is.finite(scenario_1) |
      !is.finite(scenario_2) |
      !is.finite(absolute_difference) |
      !is.finite(percentage_difference) |
      scenario_1 < 0 |
      scenario_2 < 0
  )

if (nrow(invalid_values) > 0) {
  
  print(invalid_values, n = Inf)
  
  stop(
    "\nInvalid projection values were detected.",
    call. = FALSE
  )
}


# ==============================================================================
# 7. FIGURE SETTINGS
# ==============================================================================

number_positive_countries <- 5
number_negative_countries <- 5

difference_map_quantile <- 0.98

scenario_1_colour <- "#2166AC"
scenario_2_colour <- "#B2182B"

map_low_colour <- "#2166AC"
map_mid_colour <- "white"
map_high_colour <- "#B2182B"


# ==============================================================================
# 8. COUNTRY-NAME FUNCTIONS
# ==============================================================================

clean_text <- function(x) {
  
  x |>
    as.character() |>
    stringr::str_replace_all("\u00A0", " ") |>
    stringr::str_replace_all("\u2019", "'") |>
    stringr::str_squish()
}


normalise_country_name <- function(x) {
  
  original_name <- clean_text(x)
  
  dplyr::recode(
    original_name,
    
    "Bolivia" =
      "Bolivia (Plurinational State of)",
    
    "Cape Verde" =
      "Cabo Verde",
    
    "Congo, Dem. Rep." =
      "Democratic Republic of the Congo",
    
    "Democratic Republic of Congo" =
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
    
    "Dem. People's Republic of Korea" =
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
    
    "Micronesia" =
      "Micronesia (Federated States of)",
    
    "Micronesia (Fed. States of)" =
      "Micronesia (Federated States of)",
    
    "Micronesia (country)" =
      "Micronesia (Federated States of)",
    
    "China, Taiwan Province of China" =
      "Taiwan",
    
    "Taiwan (Province of China)" =
      "Taiwan",
    
    "Brunei" =
      "Brunei Darussalam",
    
    "Swaziland" =
      "Eswatini",
    
    .default =
      original_name
  )
}


# ==============================================================================
# 9. LOAD WORLD MAP
# ==============================================================================

world_map <- rnaturalearth::ne_countries(
  scale = "medium",
  returnclass = "sf"
)

if (!inherits(world_map, "sf")) {
  
  stop(
    "\nThe Natural Earth world map could not be loaded.",
    call. = FALSE
  )
}

world_map <- world_map |>
  mutate(
    map_country_original =
      dplyr::coalesce(
        admin,
        name_long,
        sovereignt
      ),
    
    map_country =
      normalise_country_name(
        map_country_original
      ),
    
    map_country =
      dplyr::recode(
        map_country,
        
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
        
        "Taiwan (Province of China)" =
          "Taiwan",
        
        "Falkland Islands" =
          "Falkland Islands (Malvinas)",
        
        "Somaliland" =
          "Somalia",
        
        .default =
          map_country
      )
  )


# ==============================================================================
# 10. MAP LINKAGE CHECK
# ==============================================================================

mapped_country_names <- world_map |>
  st_drop_geometry() |>
  distinct(map_country)

countries_not_mapped <- projection_2050 |>
  distinct(country) |>
  anti_join(
    mapped_country_names,
    by = c(
      "country" =
        "map_country"
    )
  )

message("")
message("Countries without Natural Earth polygons:")

if (nrow(countries_not_mapped) == 0) {
  
  message("None")
  
} else {
  
  print(countries_not_mapped, n = Inf)
}


# ==============================================================================
# 11. COMMON MAP SCALE
# ==============================================================================

difference_limit <- projection_2050 |>
  summarise(
    lower_value =
      as.numeric(
        quantile(
          percentage_difference,
          probs = 1 - difference_map_quantile,
          na.rm = TRUE,
          names = FALSE
        )
      ),
    
    upper_value =
      as.numeric(
        quantile(
          percentage_difference,
          probs = difference_map_quantile,
          na.rm = TRUE,
          names = FALSE
        )
      ),
    
    symmetric_limit =
      max(
        abs(c(lower_value, upper_value)),
        na.rm = TRUE
      )
  ) |>
  pull(symmetric_limit)

if (
  length(difference_limit) != 1 ||
  is.na(difference_limit) ||
  !is.finite(difference_limit) ||
  difference_limit <= 0
) {
  
  stop(
    "\nA valid common map limit could not be calculated.",
    call. = FALSE
  )
}

difference_limit <- ceiling(
  difference_limit / 10
) * 10

message("")
message(
  "Difference map limit: ±",
  difference_limit,
  "%"
)


# ==============================================================================
# 12. SELECT 10 COUNTRIES PER OUTCOME
# ==============================================================================

select_dumbbell_countries <- function(selected_outcome) {
  
  outcome_data <- projection_2050 |>
    filter(
      outcome == selected_outcome
    )
  
  positive_countries <- outcome_data |>
    filter(
      absolute_difference > 0
    ) |>
    arrange(
      desc(absolute_difference)
    ) |>
    slice_head(
      n = number_positive_countries
    ) |>
    mutate(
      selected_group =
        "Scenario 2 higher"
    )
  
  negative_countries <- outcome_data |>
    filter(
      absolute_difference < 0
    ) |>
    arrange(
      absolute_difference
    ) |>
    slice_head(
      n = number_negative_countries
    ) |>
    mutate(
      selected_group =
        "Scenario 2 lower"
    )
  
  selected_data <- bind_rows(
    positive_countries,
    negative_countries
  ) |>
    arrange(
      absolute_difference
    ) |>
    mutate(
      country_plot =
        factor(
          country,
          levels = unique(country)
        )
    )
  
  if (nrow(selected_data) != 10) {
    
    warning(
      paste0(
        "\n",
        selected_outcome,
        " dumbbell plot contains ",
        nrow(selected_data),
        " countries instead of 10."
      )
    )
  }
  
  selected_data
}


incidence_dumbbell_data <- select_dumbbell_countries(
  "Incidence"
)

deaths_dumbbell_data <- select_dumbbell_countries(
  "Deaths"
)

dalys_dumbbell_data <- select_dumbbell_countries(
  "DALYs"
)


# ==============================================================================
# 13. OUTCOME DISPLAY SCALES
# ==============================================================================

get_outcome_scale <- function(selected_outcome) {
  
  if (selected_outcome == "DALYs") {
    
    return(
      list(
        divisor = 1000000,
        axis_label =
          "Projected DALYs in 2050, millions",
        suffix = "M",
        accuracy = 0.1
      )
    )
  }
  
  if (selected_outcome == "Incidence") {
    
    return(
      list(
        divisor = 1000,
        axis_label =
          "Projected incident cases in 2050, thousands",
        suffix = "K",
        accuracy = 1
      )
    )
  }
  
  list(
    divisor = 1000,
    axis_label =
      "Projected deaths in 2050, thousands",
    suffix = "K",
    accuracy = 1
  )
}


# ==============================================================================
# 14. DUMBBELL PLOT FUNCTION
# ==============================================================================

create_dumbbell_chart <- function(
    selected_data,
    selected_outcome,
    panel_title,
    show_legend = FALSE
) {
  
  scale_settings <- get_outcome_scale(
    selected_outcome
  )
  
  chart_data <- selected_data |>
    mutate(
      scenario_1_display =
        scenario_1 / scale_settings$divisor,
      
      scenario_2_display =
        scenario_2 / scale_settings$divisor,
      
      lower_display =
        pmin(
          scenario_1_display,
          scenario_2_display
        ),
      
      upper_display =
        pmax(
          scenario_1_display,
          scenario_2_display
        ),
      
      percentage_label =
        paste0(
          ifelse(
            percentage_difference > 0,
            "+",
            ""
          ),
          formatC(
            percentage_difference,
            format = "f",
            digits = 0
          ),
          "%"
        )
    )
  
  chart_range <- max(
    chart_data$upper_display,
    na.rm = TRUE
  )
  
  chart_data <- chart_data |>
    mutate(
      label_position =
        upper_display +
        chart_range * 0.025,
      
      label_colour =
        ifelse(
          percentage_difference >= 0,
          scenario_2_colour,
          scenario_1_colour
        )
    )
  
  maximum_x <- max(
    chart_data$label_position,
    na.rm = TRUE
  )
  
  ggplot(
    chart_data,
    aes(
      y = country_plot
    )
  ) +
    
    geom_segment(
      aes(
        x = scenario_1_display,
        xend = scenario_2_display,
        yend = country_plot
      ),
      linewidth = 1.15,
      color = "grey50"
    ) +
    
    geom_point(
      aes(
        x = scenario_1_display,
        shape = "Scenario 1"
      ),
      size = 3.7,
      stroke = 1,
      color = scenario_1_colour,
      fill = "white"
    ) +
    
    geom_point(
      aes(
        x = scenario_2_display,
        shape = "Scenario 2"
      ),
      size = 3.8,
      stroke = 0.8,
      color = scenario_2_colour,
      fill = scenario_2_colour
    ) +
    
    geom_text(
      aes(
        x = label_position,
        label = percentage_label,
        color = label_colour
      ),
      hjust = 0,
      size = 3.8,
      fontface = "bold",
      show.legend = FALSE
    ) +
    
    scale_color_identity() +
    
    scale_shape_manual(
      values = c(
        "Scenario 1" = 21,
        "Scenario 2" = 19
      ),
      name = NULL
    ) +
    
    scale_x_continuous(
      labels =
        scales::label_number(
          accuracy =
            scale_settings$accuracy,
          suffix =
            scale_settings$suffix,
          big.mark = ","
        ),
      expand =
        expansion(
          mult = c(
            0.02,
            0.18
          )
        )
    ) +
    
    coord_cartesian(
      xlim = c(
        0,
        maximum_x * 1.08
      ),
      clip = "off"
    ) +
    
    labs(
      title = panel_title,
      x = scale_settings$axis_label,
      y = NULL
    ) +
    
    theme_minimal(
      base_size = 11
    ) +
    
    theme(
      plot.title =
        element_text(
          face = "bold",
          size = 12.5,
          hjust = 0,
          margin = margin(
            b = 5
          )
        ),
      
      axis.text.y =
        element_text(
          size = 10.2,
          color = "black"
        ),
      
      axis.text.x =
        element_text(
          size = 9
        ),
      
      axis.title.x =
        element_text(
          size = 9.5,
          margin = margin(
            t = 6
          )
        ),
      
      panel.grid.major.y =
        element_blank(),
      
      panel.grid.minor =
        element_blank(),
      
      panel.grid.major.x =
        element_line(
          linewidth = 0.3,
          color = "grey84"
        ),
      
      legend.position =
        if (show_legend) {
          "bottom"
        } else {
          "none"
        },
      
      legend.text =
        element_text(
          size = 9
        ),
      
      legend.key.width =
        grid::unit(
          1.2,
          "cm"
        ),
      
      legend.margin =
        margin(
          t = 2,
          b = 0
        ),
      
      plot.margin =
        margin(
          t = 4,
          r = 40,
          b = 2,
          l = 4
        )
    )
}


# ==============================================================================
# 15. DIFFERENCE MAP FUNCTION
# ==============================================================================

create_difference_map <- function(
    selected_outcome,
    panel_title
) {
  
  selected_projection <- projection_2050 |>
    filter(
      outcome == selected_outcome
    ) |>
    select(
      country,
      percentage_difference
    )
  
  selected_map <- world_map |>
    left_join(
      selected_projection,
      by = c(
        "map_country" =
          "country"
      )
    )
  
  ggplot(selected_map) +
    
    geom_sf(
      aes(
        fill = percentage_difference
      ),
      color = "grey65",
      linewidth = 0.08
    ) +
    
    scale_fill_gradient2(
      low = map_low_colour,
      mid = map_mid_colour,
      high = map_high_colour,
      midpoint = 0,
      
      limits = c(
        -difference_limit,
        difference_limit
      ),
      
      oob = scales::squish,
      
      na.value = "grey90",
      
      breaks =
        pretty(
          c(
            -difference_limit,
            difference_limit
          ),
          n = 5
        ),
      
      labels =
        scales::label_number(
          accuracy = 1,
          suffix = "%"
        ),
      
      name =
        "Scenario 2\nversus Scenario 1"
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
      title = panel_title
    ) +
    
    theme_void(
      base_size = 11
    ) +
    
    theme(
      plot.title =
        element_text(
          face = "bold",
          size = 12.5,
          hjust = 0,
          margin = margin(
            b = 5
          )
        ),
      
      legend.position = "bottom",
      
      legend.title =
        element_text(
          face = "bold",
          size = 8.7,
          lineheight = 0.9
        ),
      
      legend.text =
        element_text(
          size = 8.5
        ),
      
      legend.key.width =
        grid::unit(
          1.55,
          "cm"
        ),
      
      legend.key.height =
        grid::unit(
          0.33,
          "cm"
        ),
      
      legend.margin =
        margin(
          t = 0,
          b = 0
        ),
      
      plot.margin =
        margin(
          t = 4,
          r = 3,
          b = 2,
          l = 3
        )
    )
}


# ==============================================================================
# 16. CREATE PANELS
# ==============================================================================

figure_8a <- create_dumbbell_chart(
  selected_data =
    incidence_dumbbell_data,
  selected_outcome =
    "Incidence",
  panel_title =
    "A. Incident cases: Scenario 1 versus Scenario 2",
  show_legend =
    TRUE
)

figure_8b <- create_difference_map(
  selected_outcome =
    "Incidence",
  panel_title =
    "B. Relative epidemiological effect on incident cases"
)

figure_8c <- create_dumbbell_chart(
  selected_data =
    deaths_dumbbell_data,
  selected_outcome =
    "Deaths",
  panel_title =
    "C. Deaths: Scenario 1 versus Scenario 2",
  show_legend =
    FALSE
)

figure_8d <- create_difference_map(
  selected_outcome =
    "Deaths",
  panel_title =
    "D. Relative epidemiological effect on deaths"
)

figure_8e <- create_dumbbell_chart(
  selected_data =
    dalys_dumbbell_data,
  selected_outcome =
    "DALYs",
  panel_title =
    "E. DALYs: Scenario 1 versus Scenario 2",
  show_legend =
    FALSE
)

figure_8f <- create_difference_map(
  selected_outcome =
    "DALYs",
  panel_title =
    "F. Relative epidemiological effect on DALYs"
)


# ==============================================================================
# 17. COMBINE FINAL FIGURE
# ==============================================================================

figure_8 <- (
  
  figure_8a + figure_8b
  
) / (
  
  figure_8c + figure_8d
  
) / (
  
  figure_8e + figure_8f
  
) +
  
  patchwork::plot_layout(
    widths = c(
      1.08,
      0.92
    ),
    heights = c(
      1,
      1,
      1
    ),
    guides = "keep"
  ) +
  
  patchwork::plot_annotation(
    title =
      "Country-level projected colorectal cancer burden in 2050",
    
    subtitle =
      "Scenario comparison for countries with the largest absolute differences and geographic distribution of the relative epidemiological effect",
    
    caption =
      paste0(
        "Scenario 1 incorporates population growth only. Scenario 2 additionally ",
        "incorporates constrained country-specific 2010–2023 age-standardized-rate ",
        "trends. Dumbbell plots show the five largest positive and five largest ",
        "negative absolute Scenario 2 minus Scenario 1 differences for each outcome. ",
        "Percentage labels and maps represent 100 × (Scenario 2 − Scenario 1) / ",
        "Scenario 1. Blue indicates lower burden under Scenario 2; red indicates ",
        "higher burden."
      ),
    
    theme = theme(
      plot.title =
        element_text(
          face = "bold",
          size = 17,
          hjust = 0.5
        ),
      
      plot.subtitle =
        element_text(
          size = 10.5,
          hjust = 0.5,
          margin = margin(
            b = 5
          )
        ),
      
      plot.caption =
        element_text(
          size = 8.3,
          hjust = 0,
          margin = margin(
            t = 4
          )
        ),
      
      plot.margin =
        margin(
          t = 10,
          r = 12,
          b = 8,
          l = 12
        )
    )
  )


# ==============================================================================
# 18. SAVE FINAL FIGURE 8
# ==============================================================================

ggsave(
  filename = figure_8_file,
  plot = figure_8,
  device = "tiff",
  width = 16,
  height = 12.2,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white",
  limitsize = FALSE
)


# ==============================================================================
# 19. EXPORT FIGURE SOURCE DATA
# ==============================================================================

figure_8_source_data <- bind_rows(
  
  incidence_dumbbell_data |>
    mutate(
      panel = "A. Incidence"
    ),
  
  deaths_dumbbell_data |>
    mutate(
      panel = "C. Deaths"
    ),
  
  dalys_dumbbell_data |>
    mutate(
      panel = "E. DALYs"
    )
  
) |>
  select(
    panel,
    country,
    outcome,
    scenario_1,
    scenario_2,
    absolute_difference,
    percentage_difference,
    selected_group
  ) |>
  arrange(
    panel,
    absolute_difference
  )

readr::write_csv(
  figure_8_source_data,
  figure_8_source_file
)


# ==============================================================================
# 20. VERIFY OUTPUT
# ==============================================================================

if (!file.exists(figure_8_file)) {
  
  stop(
    "\nThe revised Figure 8 was not created.",
    call. = FALSE
  )
}

figure_information <- file.info(
  figure_8_file
)

if (
  is.na(figure_information$size) ||
  figure_information$size <= 0
) {
  
  stop(
    "\nThe revised Figure 8 file is empty.",
    call. = FALSE
  )
}


# ==============================================================================
# 21. DISPLAY SOURCE DATA
# ==============================================================================

message("")
message("==============================================================")
message("FIGURE 8 SOURCE DATA")
message("==============================================================")

print(
  figure_8_source_data,
  n = Inf
)


# ==============================================================================
# 22. COMPLETION MESSAGE
# ==============================================================================

message("")
message("==============================================================")
message("REVISED HYBRID FIGURE 8 COMPLETED")
message("==============================================================")

message("")
message("Changes applied:")
message("- 5 largest positive differences per outcome")
message("- 5 largest negative differences per outcome")
message("- Wider dumbbell panels")
message("- Larger country labels")
message("- Larger points and percentage labels")
message("- Reduced vertical white space")
message("- Shorter caption")
message("- Common percentage scale across all three maps")

message("")
message("Figure:")
message(figure_8_file)

message("")
message("Figure source data:")
message(figure_8_source_file)

message("")
message("Projection calculations were not changed.")
message("Table 8 was not changed.")
message("Supplementary Table S7 was not changed.")
message("==============================================================")