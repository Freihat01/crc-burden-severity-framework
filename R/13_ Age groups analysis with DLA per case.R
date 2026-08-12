# ============================================================================
# R_03_12_Age_group_burden_severity_framework_2023.R
# Section 3.12: Age-group burden–severity framework, 2023
# ============================================================================

rm(list = ls()); graphics.off(); cat("\014")
options(stringsAsFactors = FALSE, scipen = 999, dplyr.summarise.inform = FALSE, warn = 1)

pkgs <- c("tidyverse","readxl","openxlsx","janitor","ggalluvial","patchwork","ggrepel","scales")
miss <- setdiff(pkgs, rownames(installed.packages()))
if (length(miss)) install.packages(miss, dependencies = TRUE)
suppressPackageStartupMessages({
  library(tidyverse); library(readxl); library(openxlsx); library(janitor)
  library(ggalluvial); library(patchwork); library(ggrepel); library(scales)
})

# ----------------------------------------------------------------------------
# 1. Project root and exact project folders
# ----------------------------------------------------------------------------
# This script is normally run from:
#   Colorectal cancer/R
# but the function below also works if RStudio starts in the project root or
# another subfolder. It searches upward until it finds the actual project root.

find_project_root <- function(start_path = getwd(), max_levels = 8L) {
  current <- normalizePath(start_path, winslash = "/", mustWork = TRUE)
  
  for (i in 0:max_levels) {
    required_dirs <- c(
      "Incidence",
      "Mortality",
      "DALYs",
      "YLDs",
      "YLLs",
      "Population",
      "Publication"
    )
    
    if (all(dir.exists(file.path(current, required_dirs)))) {
      return(current)
    }
    
    parent <- normalizePath(
      file.path(current, ".."),
      winslash = "/",
      mustWork = TRUE
    )
    
    if (identical(parent, current)) break
    current <- parent
  }
  
  stop(
    paste0(
      "Could not find the 'Colorectal cancer' project root.\n",
      "The correct root must contain these folders:\n",
      "Incidence, Mortality, DALYs, YLDs, YLLs, Population, Publication.\n\n",
      "Current working directory:\n",
      normalizePath(getwd(), winslash = "/", mustWork = TRUE)
    ),
    call. = FALSE
  )
}

project_root <- find_project_root()
message("Project root: ", project_root)

folders <- list(
  Incidence = file.path(project_root, "Incidence"),
  Mortality = file.path(project_root, "Mortality"),
  DALYs = file.path(project_root, "DALYs"),
  YLDs = file.path(project_root, "YLDs"),
  YLLs = file.path(project_root, "YLLs"),
  Population = file.path(project_root, "Population"),
  Figures = file.path(project_root, "Publication", "Figures"),
  Tables = file.path(project_root, "Publication", "Tables"),
  Supplementary = file.path(project_root, "Publication", "Supplementary")
)

walk(
  folders[c("Figures", "Tables", "Supplementary")],
  dir.create,
  recursive = TRUE,
  showWarnings = FALSE
)

# ----------------------------------------------------------------------------
# 2. Use the exact filenames and exact folders supplied by the user
# ----------------------------------------------------------------------------
# No guessing is used for the five epidemiological Excel files.

find_file_by_stem <- function(folder, exact_stem, label) {
  if (!dir.exists(folder)) {
    stop(
      paste0(
        "Required folder not found for ", label, ":\n",
        folder
      ),
      call. = FALSE
    )
  }
  
  # IHME downloads on this computer are commonly saved as .xlsx.csv.
  # Therefore, match the filename stem and accept .csv, .xlsx.csv, .xlsx, or .xls.
  candidates <- list.files(
    path = folder,
    full.names = TRUE,
    recursive = FALSE
  )
  
  candidates <- candidates[
    file.exists(candidates) &
      !dir.exists(candidates) &
      !stringr::str_detect(basename(candidates), "^~\\$")
  ]
  
  candidate_names <- basename(candidates)
  
  exact_matches <- candidates[
    stringr::str_to_lower(candidate_names) %in%
      stringr::str_to_lower(
        c(
          exact_stem,
          paste0(exact_stem, ".csv"),
          paste0(exact_stem, ".xlsx"),
          paste0(exact_stem, ".xls"),
          paste0(exact_stem, ".xlsx.csv"),
          paste0(exact_stem, ".xls.csv")
        )
      )
  ]
  
  if (length(exact_matches) == 0) {
    # Fallback: accept files beginning with the exact stem, including copies such
    # as '(1)' or '(2)', while excluding unrelated all-age files.
    stem_pattern <- paste0(
      "^",
      stringr::str_replace_all(
        exact_stem,
        "([.()+*?^$|{}\\[\\]\\\\])",
        "\\\\\\1"
      )
    )
    
    exact_matches <- candidates[
      stringr::str_detect(
        candidate_names,
        stringr::regex(stem_pattern, ignore_case = TRUE)
      )
    ]
  }
  
  if (length(exact_matches) == 0) {
    stop(
      paste0(
        "Required file not found for ", label, ".\n\n",
        "Expected filename stem:\n",
        exact_stem,
        "\n\nFolder searched:\n",
        folder,
        "\n\nFiles currently present in the expected folder:\n",
        paste(list.files(folder), collapse = "\n")
      ),
      call. = FALSE
    )
  }
  
  if (length(exact_matches) > 1) {
    info <- file.info(exact_matches)
    exact_matches <- exact_matches[order(info$mtime, decreasing = TRUE)]
    message(
      "Multiple files matched for ", label,
      ". The most recently modified file was selected: ",
      basename(exact_matches[[1]])
    )
  }
  
  normalizePath(exact_matches[[1]], winslash = "/", mustWork = TRUE)
}

find_population_file <- function(folder) {
  candidates <- list.files(
    folder,
    pattern = "^IHME-GBD_2023_AgeGroups_2023",
    full.names = TRUE,
    recursive = FALSE,
    ignore.case = TRUE
  )
  
  candidates <- candidates[
    file.exists(candidates) &
      !dir.exists(candidates) &
      !str_detect(basename(candidates), "^~\\$")
  ]
  
  if (length(candidates) == 0) {
    stop(
      paste0(
        "Population file not found in:\n",
        folder,
        "\n\nExpected a file beginning with:\n",
        "IHME-GBD_2023_AgeGroups_2023",
        "\n\nFiles currently present in the Population folder:\n",
        paste(list.files(folder), collapse = "\n")
      ),
      call. = FALSE
    )
  }
  
  if (length(candidates) > 1) {
    info <- file.info(candidates)
    candidates <- candidates[order(info$mtime, decreasing = TRUE)]
    message(
      "Multiple population files matched. The most recently modified file was selected: ",
      basename(candidates[[1]])
    )
  }
  
  normalizePath(candidates[[1]], winslash = "/", mustWork = TRUE)
}

files <- list(
  incidence = find_file_by_stem(
    folders$Incidence,
    "Incidence_Number_AllCountries_BothSexes_AgeGroups_2023",
    "incidence"
  ),
  
  deaths = find_file_by_stem(
    folders$Mortality,
    "Deaths_Number_AllCountries_BothSexes_AgeGroups_2023",
    "deaths"
  ),
  
  dalys = find_file_by_stem(
    folders$DALYs,
    "DALYs_Number_AllCountries_BothSexes_AgeGroups_2023",
    "DALYs"
  ),
  
  ylds = find_file_by_stem(
    folders$YLDs,
    "YLDs_Number_AllCountries_BothSexes_AgeGroups_2023",
    "YLDs"
  ),
  
  ylls = find_file_by_stem(
    folders$YLLs,
    "YLLs_Number_AllCountries_BothSexes_AgeGroups_2023",
    "YLLs"
  ),
  
  population = find_population_file(
    folders$Population
  )
)

message("\nInputs selected:")
walk(files, message)

# ----------------------------------------------------------------------------
# 3. Settings
# ----------------------------------------------------------------------------
analysis_year <- 2023L
expected_country_count <- 204L
age_levels <- c("0-4","5-9","10-14","15-19","20-24","25-29","30-34","35-39","40-44","45-49",
                "50-54","55-59","60-64","65-69","70-74","75-79","80-84","85-89","90-94","95+")
broad_levels <- c("<50 years","50–69 years","≥70 years")
quad_levels <- c("Q1: High burden–high severity","Q2: Low burden–high severity",
                 "Q3: Low burden–low severity","Q4: High burden–low severity")
quad_codes <- c("Q1","Q2","Q3","Q4")
quad_cols <- c(
  "Q1: High burden–high severity"="#B2182B",
  "Q2: Low burden–high severity"="#EF8A62",
  "Q3: Low burden–low severity"="#67A9CF",
  "Q4: High burden–low severity"="#2166AC"
)

# ----------------------------------------------------------------------------
# 4. Helpers
# ----------------------------------------------------------------------------
clean_text <- function(x) x |> as.character() |> str_replace_all("\u00A0", " ") |>
  str_replace_all("â€“|â€”", "–") |> str_replace_all("CÃ´te d'Ivoire", "Côte d'Ivoire") |>
  str_replace_all("TÃ¼rkiye", "Türkiye") |> str_squish()

read_any <- function(path) {
  ext <- tolower(tools::file_ext(path))
  if (ext == "csv") return(read_csv(path, show_col_types = FALSE, progress = FALSE, guess_max = 100000))
  sheets <- excel_sheets(path)
  score <- map_dbl(sheets, ~tryCatch({x <- read_excel(path, sheet = .x, n_max = 20); nrow(x)*max(1,ncol(x))}, error=function(e) 0))
  read_excel(path, sheet = sheets[[which.max(score)]])
}

find_col <- function(x, candidates, label, required = TRUE) {
  hit <- candidates[candidates %in% names(x)]
  if (length(hit)) return(hit[[1]])
  if (required) stop("Missing column for ", label, ". Detected: ", paste(names(x), collapse=", "), call. = FALSE)
  NA_character_
}

num <- function(x) parse_number(as.character(x), na = c("","NA","N/A","NaN","NULL","-"))

norm_age <- function(x) {
  
  z <- x |>
    clean_text() |>
    str_to_lower() |>
    str_replace_all("years|year|yrs|yr", "") |>
    str_replace_all("to", "-") |>
    str_replace_all("–|—", "-") |>
    str_replace_all("\\s+", "")
  
  case_when(
    
    # Standard five-year age groups — exact matches first
    z == "5-9"   ~ "5-9",
    z == "10-14" ~ "10-14",
    z == "15-19" ~ "15-19",
    z == "20-24" ~ "20-24",
    z == "25-29" ~ "25-29",
    z == "30-34" ~ "30-34",
    z == "35-39" ~ "35-39",
    z == "40-44" ~ "40-44",
    z == "45-49" ~ "45-49",
    z == "50-54" ~ "50-54",
    z == "55-59" ~ "55-59",
    z == "60-64" ~ "60-64",
    z == "65-69" ~ "65-69",
    z == "70-74" ~ "70-74",
    z == "75-79" ~ "75-79",
    z == "80-84" ~ "80-84",
    z == "85-89" ~ "85-89",
    z == "90-94" ~ "90-94",
    
    # Oldest group
    z %in% c(
      "95+",
      "95plus",
      "95andover",
      "95orolder"
    ) ~ "95+",
    
    # Under-five categories
    z %in% c(
      "<5",
      "under5",
      "lessthan5",
      "0-4",
      "<1",
      "under1",
      "lessthan1",
      "0-0",
      "1-4",
      "earlyneonatal",
      "lateneonatal",
      "postneonatal",
      "neonatal"
    ) ~ "0-4",
    
    TRUE ~ NA_character_
  )
}

broad_age <- function(a) case_when(
  a %in% age_levels[1:10] ~ "<50 years",
  a %in% age_levels[11:14] ~ "50–69 years",
  a %in% age_levels[15:20] ~ "≥70 years",
  TRUE ~ NA_character_
)

# IHME can omit country-age rows when the estimated count is zero.
# Therefore, missing epidemiological COUNT rows are treated as zero only after
# the original missingness has been recorded. Population is never imputed.
sum_sparse_count <- function(x) {
  if (length(x) == 0) return(0)
  sum(replace_na(as.numeric(x), 0), na.rm = TRUE)
}

sum_population_available <- function(x) {
  x <- as.numeric(x)
  if (length(x) == 0 || all(is.na(x))) return(NA_real_)
  sum(x, na.rm = TRUE)
}

make_quad <- function(b,s,bt,st) case_when(
  b >= bt & s >= st ~ "Q1: High burden–high severity",
  b <  bt & s >= st ~ "Q2: Low burden–high severity",
  b <  bt & s <  st ~ "Q3: Low burden–low severity",
  b >= bt & s <  st ~ "Q4: High burden–low severity",
  TRUE ~ NA_character_
)
qcode <- function(x) str_extract(as.character(x), "Q[1-4]")

kappa_unweighted <- function(a,b) {
  a <- factor(a, levels=quad_codes); b <- factor(b, levels=quad_codes)
  keep <- !is.na(a)&!is.na(b); a <- a[keep]; b <- b[keep]; n <- length(a)
  tab <- table(a,b); p <- tab/n; po <- sum(diag(p)); pe <- sum(rowSums(p)*colSums(p))
  tibble(n=n, observed_agreement=po, expected_agreement=pe,
         kappa=ifelse(abs(1-pe)<1e-12, NA_real_, (po-pe)/(1-pe)))
}

# ----------------------------------------------------------------------------
# 5. Standardize one source
# ----------------------------------------------------------------------------
prepare_source <- function(path, value_name, structural_zero_allowed = FALSE) {
  raw <- read_any(path) |> clean_names()
  ctry <- find_col(raw, c("location","country","location_name","name"), "country")
  agec <- find_col(raw, c("age","age_group","age_name","age_group_name"), "age")
  yearc <- find_col(raw, c("year","year_id"), "year", FALSE)
  sexc <- find_col(raw, c("sex","sex_name"), "sex", FALSE)
  metricc <- find_col(raw, c("metric","metric_name"), "metric", FALSE)
  valuec <- find_col(raw, c("val","value","estimate","mean","population","pop"), "value")
  
  d <- tibble(
    country = clean_text(raw[[ctry]]),
    raw_age = clean_text(raw[[agec]]),
    age_group = norm_age(raw[[agec]]),
    year = if (is.na(yearc)) analysis_year else suppressWarnings(as.integer(raw[[yearc]])),
    sex = if (is.na(sexc)) "Both" else clean_text(raw[[sexc]]),
    metric = if (is.na(metricc)) "Number" else clean_text(raw[[metricc]]),
    value = num(raw[[valuec]])
  )
  
  # Filter only when the corresponding columns genuinely contain those labels.
  if (!is.na(yearc)) {
    d <- d |> filter(year == analysis_year)
  }
  
  sex_values <- unique(str_to_lower(d$sex))
  if (any(sex_values %in% c("both", "both sexes", "both sex"))) {
    d <- d |>
      filter(str_to_lower(sex) %in% c("both", "both sexes", "both sex"))
  }
  
  metric_values <- unique(str_to_lower(d$metric))
  if (any(metric_values %in% c("number", "numbers", "count", "counts"))) {
    d <- d |>
      filter(str_to_lower(metric) %in% c("number", "numbers", "count", "counts"))
  }
  
  unrecognized <- d |> filter(is.na(age_group)) |> distinct(raw_age)
  d <- d |> filter(!is.na(age_group)) |> group_by(country, age_group) |>
    summarise(value = if (all(is.na(value))) NA_real_ else sum(value, na.rm=TRUE), .groups="drop") |>
    rename(!!value_name := value)
  
  present <- sort(unique(d$age_group)); absent <- setdiff(age_levels, present)
  list(data=d, absent=absent, present=present, unrecognized=unrecognized,
       structural_zero_allowed=structural_zero_allowed, value_name=value_name)
}

src <- list(
  incidence = prepare_source(files$incidence, "incidence", TRUE),
  deaths = prepare_source(files$deaths, "deaths", TRUE),
  dalys = prepare_source(files$dalys, "dalys", TRUE),
  ylls = prepare_source(files$ylls, "ylls", TRUE),
  ylds = prepare_source(files$ylds, "ylds", TRUE),
  population = prepare_source(files$population, "population", FALSE)
)

# Report imported row counts before building the framework.
import_audit <- imap_dfr(
  src,
  ~ tibble(
    Source = .y,
    Imported_country_age_rows = nrow(.x$data),
    Countries = n_distinct(.x$data$country),
    Recognized_age_groups = length(.x$present),
    Unrecognized_age_labels = nrow(.x$unrecognized)
  )
)

message("\nImport audit:")
print(import_audit, n = Inf)

age_coverage <- imap_dfr(src, ~tibble(Source=.y, Ages_present=paste(.x$present,collapse=", "),
                                      Ages_absent=paste(.x$absent,collapse=", ")))
print(age_coverage, n=Inf)

# ----------------------------------------------------------------------------
# 6. Country universe and full grid
# ----------------------------------------------------------------------------
country_universe <- Reduce(intersect, list(unique(src$incidence$data$country),
                                           unique(src$dalys$data$country),
                                           unique(src$population$data$country))) |> sort()
if (length(country_universe) != expected_country_count) {
  stop("Expected 204 common countries across incidence, DALYs and population; found ",
       length(country_universe), call. = FALSE)
}

grid <- expand_grid(country=country_universe, age_group=age_levels) |>
  mutate(broad_age_group=broad_age(age_group))

merge_src <- function(base, s) {
  v <- s$value_name
  out <- left_join(base, s$data, by = c("country", "age_group"))
  
  missing_name <- paste0(v, "_missing_original")
  imputed_name <- paste0(v, "_zero_imputed")
  
  out[[missing_name]] <- is.na(out[[v]])
  out[[imputed_name]] <- FALSE
  
  # For sparse IHME epidemiological count files, omitted country-age rows are
  # treated as zero counts. The original omission remains recorded in the audit.
  # Population is never imputed.
  if (s$structural_zero_allowed) {
    idx <- is.na(out[[v]])
    out[[imputed_name]][idx] <- TRUE
    out[[v]][idx] <- 0
  }
  
  out
}

age_detail <- reduce(src, merge_src, .init=grid)

missingness_audit <- tibble(
  Check = c(
    "Incidence rows absent in source",
    "Incidence absent rows treated as zero",
    "Deaths rows absent in source",
    "Deaths absent rows treated as zero",
    "DALY rows absent in source",
    "DALY absent rows treated as zero",
    "YLL rows absent in source",
    "YLL absent rows treated as zero",
    "YLD rows absent in source",
    "YLD absent rows treated as zero",
    "Population rows absent in source"
  ),
  Rows = c(
    sum(age_detail$incidence_missing_original),
    sum(age_detail$incidence_zero_imputed),
    sum(age_detail$deaths_missing_original),
    sum(age_detail$deaths_zero_imputed),
    sum(age_detail$dalys_missing_original),
    sum(age_detail$dalys_zero_imputed),
    sum(age_detail$ylls_missing_original),
    sum(age_detail$ylls_zero_imputed),
    sum(age_detail$ylds_missing_original),
    sum(age_detail$ylds_zero_imputed),
    sum(age_detail$population_missing_original)
  )
)

# ----------------------------------------------------------------------------
# 7. Aggregate to <50, 50–69, ≥70, then derive metrics
# ----------------------------------------------------------------------------
framework <- age_detail |> group_by(country, broad_age_group) |>
  summarise(
    incidence = sum_sparse_count(incidence),
    deaths = sum_sparse_count(deaths),
    dalys = sum_sparse_count(dalys),
    ylls = sum_sparse_count(ylls),
    ylds = sum_sparse_count(ylds),
    population = sum_population_available(population),
    
    incidence_rows_imputed = sum(incidence_zero_imputed),
    deaths_rows_imputed = sum(deaths_zero_imputed),
    dalys_rows_imputed = sum(dalys_zero_imputed),
    ylls_rows_imputed = sum(ylls_zero_imputed),
    ylds_rows_imputed = sum(ylds_zero_imputed),
    population_rows_missing = sum(population_missing_original),
    
    incomplete_incidence = FALSE,
    incomplete_deaths = FALSE,
    incomplete_dalys = FALSE,
    incomplete_ylls = FALSE,
    incomplete_ylds = FALSE,
    incomplete_population = all(is.na(population)),
    .groups = "drop"
  ) |>
  mutate(
    broad_age_group=factor(broad_age_group, levels=broad_levels),
    daly_rate_per_100000=if_else(!is.na(dalys)&!is.na(population)&population>0, 100000*dalys/population, NA_real_),
    dalys_per_case=if_else(!is.na(dalys)&!is.na(incidence)&incidence>0, dalys/incidence, NA_real_),
    mir=if_else(!is.na(deaths)&!is.na(incidence)&incidence>0, deaths/incidence, NA_real_),
    ylls_per_case=if_else(!is.na(ylls)&!is.na(incidence)&incidence>0, ylls/incidence, NA_real_),
    ylds_per_case=if_else(!is.na(ylds)&!is.na(incidence)&incidence>0, ylds/incidence, NA_real_),
    yll_percentage=if_else(!is.na(ylls)&!is.na(dalys)&dalys>0, 100*ylls/dalys, NA_real_),
    yld_percentage=if_else(!is.na(ylds)&!is.na(dalys)&dalys>0, 100*ylds/dalys, NA_real_),
    component_relative_difference=if_else(!is.na(dalys)&dalys>0&!is.na(ylls)&!is.na(ylds),
                                          100*abs(dalys-(ylls+ylds))/dalys, NA_real_)
  )

if (any((framework |> count(broad_age_group))$n != expected_country_count))
  stop("Each broad age group must contain 204 countries.", call. = FALSE)

thresholds <- framework |> group_by(broad_age_group) |>
  summarise(valid_countries=sum(!is.na(daly_rate_per_100000)&!is.na(dalys_per_case)),
            burden_threshold=median(daly_rate_per_100000, na.rm=TRUE),
            severity_threshold=median(dalys_per_case, na.rm=TRUE), .groups="drop")
validity_audit <- thresholds |>
  mutate(
    excluded_countries = expected_country_count - valid_countries,
    valid_percent = 100 * valid_countries / expected_country_count
  )

message("\nValid framework observations by broad age group:")
print(validity_audit, n = Inf)

# A fixed 90% rule is not appropriate here: DALYs per case is undefined when
# a country has zero incident cases in an age group, particularly for small
# populations and early-onset CRC. Such country-age observations are retained
# in the supplementary audit but excluded from threshold calculation and
# quadrant assignment. The analysis proceeds using all valid observations.
if (any(thresholds$valid_countries == 0)) {
  
  diagnostic_file <- file.path(
    folders$Supplementary,
    "DIAGNOSTIC_Age_group_import_and_validity.csv"
  )
  
  diagnostic_rows <- framework |>
    transmute(
      country,
      broad_age_group,
      incidence,
      dalys,
      population,
      daly_rate_per_100000,
      dalys_per_case,
      incidence_rows_imputed,
      dalys_rows_imputed,
      population_rows_missing
    )
  
  write_csv(diagnostic_rows, diagnostic_file)
  
  stop(
    paste0(
      "No valid countries were available in at least one broad age group. ",
      "This now indicates an import or population-format problem rather than ",
      "missing zero-count epidemiological rows. A diagnostic file was saved at:\n",
      diagnostic_file,
      "\n\nReview the printed Import audit and Age coverage tables."
    ),
    call. = FALSE
  )
}

if (any(thresholds$valid_countries < expected_country_count * 0.90)) {
  warning(
    paste0(
      "One or more broad age groups contain fewer than 90% valid countries. ",
      "This is expected when incidence is zero or required age-specific data ",
      "are unavailable. See the validity and missingness audit sheets."
    ),
    call. = FALSE
  )
}

framework <- framework |> left_join(thresholds, by="broad_age_group") |>
  mutate(quadrant=make_quad(daly_rate_per_100000, dalys_per_case, burden_threshold, severity_threshold),
         quadrant_code=qcode(quadrant),
         quadrant=factor(quadrant, levels=quad_levels))

# ----------------------------------------------------------------------------
# 8. Summaries and age-group agreement
# ----------------------------------------------------------------------------
quadrant_distribution <- framework |> count(broad_age_group, quadrant, name="Countries") |>
  complete(broad_age_group=factor(broad_levels,levels=broad_levels),
           quadrant=factor(quad_levels,levels=quad_levels), fill=list(Countries=0)) |>
  group_by(broad_age_group) |> mutate(Percentage=100*Countries/sum(Countries)) |> ungroup()

age_group_summary <- framework |> group_by(broad_age_group) |>
  summarise(Countries=n(), Median_DALY_rate=median(daly_rate_per_100000,na.rm=TRUE),
            IQR_DALY_rate_lower=quantile(daly_rate_per_100000,.25,na.rm=TRUE,names=FALSE),
            IQR_DALY_rate_upper=quantile(daly_rate_per_100000,.75,na.rm=TRUE,names=FALSE),
            Median_DALYs_per_case=median(dalys_per_case,na.rm=TRUE),
            IQR_DALYs_per_case_lower=quantile(dalys_per_case,.25,na.rm=TRUE,names=FALSE),
            IQR_DALYs_per_case_upper=quantile(dalys_per_case,.75,na.rm=TRUE,names=FALSE),
            Median_MIR=median(mir,na.rm=TRUE), Median_YLLs_per_case=median(ylls_per_case,na.rm=TRUE),
            Median_YLDs_per_case=median(ylds_per_case,na.rm=TRUE),
            Median_YLL_percentage=median(yll_percentage,na.rm=TRUE),
            Median_YLD_percentage=median(yld_percentage,na.rm=TRUE), .groups="drop")

wide_codes <- framework |> select(country,broad_age_group,quadrant_code) |>
  mutate(age_key=recode(as.character(broad_age_group),"<50 years"="lt50","50–69 years"="age50_69","≥70 years"="age70plus")) |>
  select(-broad_age_group) |> pivot_wider(names_from=age_key, values_from=quadrant_code)

pairs <- tribble(~Age_group_1,~Age_group_2,~col1,~col2,
                 "<50 years","50–69 years","lt50","age50_69",
                 "<50 years","≥70 years","lt50","age70plus",
                 "50–69 years","≥70 years","age50_69","age70plus")
pairwise_agreement <- pmap_dfr(pairs, function(Age_group_1,Age_group_2,col1,col2){
  r <- kappa_unweighted(wide_codes[[col1]], wide_codes[[col2]])
  tibble(Age_group_1,Age_group_2,Countries_compared=r$n,
         Exact_agreement_percent=100*r$observed_agreement,
         Expected_agreement_percent=100*r$expected_agreement,Cohen_kappa=r$kappa)
})

country_comparison <- framework |> select(country,broad_age_group,daly_rate_per_100000,dalys_per_case,mir,
                                          ylls_per_case,ylds_per_case,yll_percentage,yld_percentage,quadrant_code) |>
  mutate(age_key=recode(as.character(broad_age_group),"<50 years"="lt50","50–69 years"="age50_69","≥70 years"="age70plus")) |>
  select(-broad_age_group) |>
  pivot_wider(names_from=age_key,
              values_from=c(daly_rate_per_100000,dalys_per_case,mir,ylls_per_case,ylds_per_case,
                            yll_percentage,yld_percentage,quadrant_code),
              names_glue="{.value}_{age_key}") |>
  rowwise() |>
  mutate(same_quadrant_all_age_groups = quadrant_code_lt50==quadrant_code_age50_69 & quadrant_code_age50_69==quadrant_code_age70plus,
         early_onset_q1_only = quadrant_code_lt50=="Q1" & quadrant_code_age50_69!="Q1" & quadrant_code_age70plus!="Q1",
         q1_in_any_age_group = any(c_across(c(quadrant_code_lt50,quadrant_code_age50_69,quadrant_code_age70plus))=="Q1"),
         number_of_distinct_quadrants = n_distinct(c_across(c(quadrant_code_lt50,quadrant_code_age50_69,quadrant_code_age70plus)), na.rm=TRUE)) |>
  ungroup()

country_patterns <- tibble(
  Measure=c("Same quadrant across all three age groups","Different quadrant across age groups",
            "Q1 for early-onset CRC only","Q1 in at least one age group",
            "One distinct quadrant","Two distinct quadrants","Three distinct quadrants"),
  Countries=c(sum(country_comparison$same_quadrant_all_age_groups,na.rm=TRUE),
              sum(!country_comparison$same_quadrant_all_age_groups,na.rm=TRUE),
              sum(country_comparison$early_onset_q1_only,na.rm=TRUE),
              sum(country_comparison$q1_in_any_age_group,na.rm=TRUE),
              sum(country_comparison$number_of_distinct_quadrants==1,na.rm=TRUE),
              sum(country_comparison$number_of_distinct_quadrants==2,na.rm=TRUE),
              sum(country_comparison$number_of_distinct_quadrants==3,na.rm=TRUE))
) |> mutate(Percentage=100*Countries/expected_country_count)

reconciliation_summary <- framework |> group_by(broad_age_group) |>
  summarise(Countries_with_complete_components=sum(!is.na(component_relative_difference)),
            Median_relative_difference_percent=median(component_relative_difference,na.rm=TRUE),
            Maximum_relative_difference_percent=max(component_relative_difference,na.rm=TRUE),
            Countries_above_1_percent_difference=sum(component_relative_difference>1,na.rm=TRUE), .groups="drop")

# ----------------------------------------------------------------------------
# 9. Figure 12 — publication-quality revised layout
# ----------------------------------------------------------------------------

# Short labels used only in the single shared legend.
quadrant_legend_labels <- c(
  "Q1: High burden–high severity",
  "Q2: Low burden–high severity",
  "Q3: Low burden–low severity",
  "Q4: High burden–low severity"
)

plot_age <- function(age_label, title_text, seed_value) {
  
  d <- framework |>
    filter(
      broad_age_group == age_label,
      !is.na(daly_rate_per_100000),
      !is.na(dalys_per_case),
      !is.na(quadrant)
    )
  
  bt <- unique(d$burden_threshold)
  st <- unique(d$severity_threshold)
  
  # Label only the two most epidemiologically distinctive countries per quadrant.
  # Labels are placed in white boxes so they do not visually merge with points.
  labs_d <- d |>
    mutate(
      normalized_distance =
        sqrt(
          ((daly_rate_per_100000 - bt) / bt)^2 +
            ((dalys_per_case - st) / st)^2
        )
    ) |>
    group_by(quadrant) |>
    slice_max(
      order_by = normalized_distance,
      n = 2,
      with_ties = FALSE
    ) |>
    ungroup()
  
  x_max <- max(d$daly_rate_per_100000, na.rm = TRUE) * 1.10
  y_min <- max(0, min(d$dalys_per_case, na.rm = TRUE) * 0.92)
  y_max <- max(d$dalys_per_case, na.rm = TRUE) * 1.08
  
  ggplot(
    d,
    aes(
      x = daly_rate_per_100000,
      y = dalys_per_case,
      fill = quadrant
    )
  ) +
    
    # Very light quadrant shading improves interpretation without dominating.
    annotate(
      "rect",
      xmin = 0, xmax = bt,
      ymin = st, ymax = Inf,
      fill = quad_cols[["Q2: Low burden–high severity"]],
      alpha = 0.035
    ) +
    annotate(
      "rect",
      xmin = bt, xmax = Inf,
      ymin = st, ymax = Inf,
      fill = quad_cols[["Q1: High burden–high severity"]],
      alpha = 0.035
    ) +
    annotate(
      "rect",
      xmin = 0, xmax = bt,
      ymin = -Inf, ymax = st,
      fill = quad_cols[["Q3: Low burden–low severity"]],
      alpha = 0.035
    ) +
    annotate(
      "rect",
      xmin = bt, xmax = Inf,
      ymin = -Inf, ymax = st,
      fill = quad_cols[["Q4: High burden–low severity"]],
      alpha = 0.035
    ) +
    
    geom_vline(
      xintercept = bt,
      linetype = "dashed",
      linewidth = 0.58,
      color = "grey30"
    ) +
    
    geom_hline(
      yintercept = st,
      linetype = "dashed",
      linewidth = 0.58,
      color = "grey30"
    ) +
    
    geom_point(
      shape = 21,
      size = 2.35,
      alpha = 0.80,
      color = "grey25",
      stroke = 0.22
    ) +
    
    ggrepel::geom_label_repel(
      data = labs_d,
      aes(label = country),
      size = 2.65,
      fontface = "bold",
      label.size = 0.18,
      label.padding = unit(0.10, "lines"),
      label.r = unit(0.08, "lines"),
      fill = scales::alpha("white", 0.92),
      color = "grey10",
      max.overlaps = Inf,
      box.padding = 0.62,
      point.padding = 0.42,
      force = 2.2,
      force_pull = 0.22,
      min.segment.length = 0,
      segment.size = 0.28,
      segment.color = "grey50",
      seed = seed_value,
      direction = "both",
      show.legend = FALSE
    ) +
    
    scale_fill_manual(
      values = quad_cols,
      breaks = quadrant_legend_labels,
      labels = quadrant_legend_labels,
      drop = FALSE
    ) +
    
    scale_x_continuous(
      labels = scales::label_number(
        accuracy = 1,
        big.mark = ","
      ),
      expand = expansion(mult = c(0.03, 0.08))
    ) +
    
    scale_y_continuous(
      labels = scales::label_number(
        accuracy = 0.1,
        big.mark = ","
      ),
      expand = expansion(mult = c(0.03, 0.07))
    ) +
    
    coord_cartesian(
      xlim = c(0, x_max),
      ylim = c(y_min, y_max),
      clip = "off"
    ) +
    
    labs(
      title = title_text,
      subtitle = paste0(
        "Median thresholds: DALY rate ",
        formatC(bt, format = "f", digits = 2, big.mark = ","),
        "; DALYs per case ",
        formatC(st, format = "f", digits = 2)
      ),
      x = "Age-group DALY rate per 100,000",
      y = "DALYs per incident case"
    ) +
    
    theme_classic(base_size = 10.5) +
    
    theme(
      plot.title = element_text(
        face = "bold",
        size = 11.8,
        margin = margin(b = 2)
      ),
      plot.subtitle = element_text(
        size = 8.7,
        color = "grey25",
        margin = margin(b = 5)
      ),
      axis.title = element_text(
        face = "bold",
        size = 9.8
      ),
      axis.text = element_text(
        size = 8.7,
        color = "grey20"
      ),
      legend.position = "none",
      plot.margin = margin(7, 15, 7, 7)
    )
}

p1 <- plot_age(
  "<50 years",
  "A. Early-onset colorectal cancer (<50 years)",
  20260821
)

p2 <- plot_age(
  "50–69 years",
  "B. Screening-age adults (50–69 years)",
  20260822
)

p3 <- plot_age(
  "≥70 years",
  "C. Older adults (≥70 years)",
  20260823
)

alluv <- framework |>
  select(country, broad_age_group, quadrant) |>
  mutate(
    age_key = recode(
      as.character(broad_age_group),
      "<50 years" = "lt50",
      "50–69 years" = "age50_69",
      "≥70 years" = "age70plus"
    )
  ) |>
  select(-broad_age_group) |>
  pivot_wider(
    names_from = age_key,
    values_from = quadrant
  ) |>
  mutate(
    across(
      c(lt50, age50_69, age70plus),
      ~ factor(.x, levels = quad_levels)
    )
  )

p4 <- ggplot(
  alluv,
  aes(
    axis1 = lt50,
    axis2 = age50_69,
    axis3 = age70plus,
    y = 1
  )
) +
  
  ggalluvial::geom_alluvium(
    aes(fill = lt50),
    width = 0.12,
    alpha = 0.67,
    knot.pos = 0.38,
    color = scales::alpha("white", 0.60),
    linewidth = 0.16
  ) +
  
  ggalluvial::geom_stratum(
    width = 0.145,
    fill = "grey98",
    color = "grey35",
    linewidth = 0.42
  ) +
  
  ggplot2::geom_text(
    stat = "stratum",
    aes(
      label = after_stat(
        stringr::str_extract(stratum, "Q[1-4]")
      )
    ),
    size = 3.45,
    fontface = "bold",
    color = "grey15"
  ) +
  
  scale_x_discrete(
    limits = c(
      "<50 years",
      "50–69 years",
      "≥70 years"
    ),
    expand = c(0.11, 0.11)
  ) +
  
  scale_fill_manual(
    values = quad_cols,
    breaks = quadrant_legend_labels,
    labels = quadrant_legend_labels,
    drop = FALSE,
    name = "Quadrant at <50 years"
  ) +
  
  scale_y_continuous(
    name = "Countries and territories",
    breaks = scales::pretty_breaks(n = 5),
    expand = expansion(mult = c(0, 0.035))
  ) +
  
  labs(
    title = "D. Country movement across age-specific frameworks"
  ) +
  
  theme_classic(base_size = 10.5) +
  
  theme(
    plot.title = element_text(
      face = "bold",
      size = 11.8,
      margin = margin(b = 7)
    ),
    axis.title.x = element_blank(),
    axis.text.x = element_text(
      face = "bold",
      size = 9.2,
      color = "grey15"
    ),
    axis.title.y = element_text(
      face = "bold",
      size = 9.8
    ),
    axis.text.y = element_text(
      size = 8.7
    ),
    legend.position = "bottom",
    legend.direction = "horizontal",
    legend.title = element_text(
      face = "bold",
      size = 8.8
    ),
    legend.text = element_text(
      size = 7.5
    ),
    legend.key.width = unit(0.48, "cm"),
    legend.key.height = unit(0.32, "cm"),
    legend.box.margin = margin(t = 2),
    plot.margin = margin(7, 8, 7, 10)
  )

figure_caption <- stringr::str_wrap(
  paste0(
    "Age-specific DALY rates were calculated using age-specific population, ",
    "and severity was defined as DALYs per incident case. Each age group used ",
    "its own median burden and severity thresholds. Flow colours in panel D ",
    "represent the quadrant assigned at age <50 years."
  ),
  width = 190
)

fig12 <- (
  p1 + p2
) / (
  p3 + p4
) +
  patchwork::plot_layout(
    widths = c(1, 1),
    heights = c(1, 1.03)
  ) +
  patchwork::plot_annotation(
    title = "Age-group colorectal cancer burden–severity framework in 2023",
    subtitle = "Country prioritisation differs across early-onset, screening-age, and older populations",
    caption = figure_caption,
    theme = theme(
      plot.title = element_text(
        face = "bold",
        size = 16,
        hjust = 0.5,
        margin = margin(b = 3)
      ),
      plot.subtitle = element_text(
        size = 10.2,
        hjust = 0.5,
        color = "grey25",
        margin = margin(b = 8)
      ),
      plot.caption = element_text(
        size = 8.0,
        hjust = 0,
        color = "grey25",
        lineheight = 1.05,
        margin = margin(t = 8)
      ),
      plot.margin = margin(12, 18, 14, 18)
    )
  )

figure_file <- file.path(
  folders$Figures,
  "Figure_12_Age_group_burden_severity_framework_2023.tiff"
)

ggsave(
  filename = figure_file,
  plot = fig12,
  device = "tiff",
  width = 16.5,
  height = 13.2,
  units = "in",
  dpi = 600,
  compression = "lzw",
  bg = "white",
  limitsize = FALSE
)

# ----------------------------------------------------------------------------
# 10. Export tables and source data
# ----------------------------------------------------------------------------
header_style <- createStyle(fontSize=10,textDecoration="bold",halign="center",valign="center",wrapText=TRUE,border="Bottom")
write_book <- function(path, sheets) {
  wb <- createWorkbook()
  for (nm in names(sheets)) {
    x <- sheets[[nm]]
    addWorksheet(wb,nm); writeData(wb,nm,x)
    if (ncol(x)>0) {
      addStyle(wb,nm,header_style,rows=1,cols=seq_len(ncol(x)),gridExpand=TRUE)
      setColWidths(wb,nm,cols=seq_len(ncol(x)),widths="auto")
      freezePane(wb,nm,firstRow=TRUE); addFilter(wb,nm,rows=1,cols=seq_len(ncol(x)))
    }
  }
  saveWorkbook(wb,path,overwrite=TRUE)
}

supp_long <- framework |> transmute(
  Country=country,Age_group=as.character(broad_age_group),Incident_cases=incidence,Deaths=deaths,DALYs=dalys,YLLs=ylls,YLDs=ylds,
  Population=population,DALY_rate_per_100000=daly_rate_per_100000,DALYs_per_case=dalys_per_case,MIR=mir,
  YLLs_per_case=ylls_per_case,YLDs_per_case=ylds_per_case,YLL_percentage=yll_percentage,YLD_percentage=yld_percentage,
  Burden_threshold=burden_threshold,Severity_threshold=severity_threshold,Quadrant=quadrant_code,
  Incomplete_incidence=incomplete_incidence,Incomplete_deaths=incomplete_deaths,Incomplete_DALYs=incomplete_dalys,
  Incomplete_YLLs=incomplete_ylls,Incomplete_YLDs=incomplete_ylds,Incomplete_population=incomplete_population,
  DALY_component_relative_difference_percent=component_relative_difference
) |> arrange(Age_group,Country)

table_file <- file.path(folders$Tables,"Table_12_Age_group_framework_summary.xlsx")
write_book(table_file,list(
  "Age-group summary"=age_group_summary,
  "Framework thresholds"=thresholds,
  "Quadrant distribution"=quadrant_distribution,
  "Pairwise agreement"=pairwise_agreement,
  "Country patterns"=country_patterns,
  "DALY reconciliation"=reconciliation_summary,
  "Missingness audit"=missingness_audit,
  "Age coverage"=age_coverage
))

s12_xlsx <- file.path(folders$Supplementary,"Table_S12_Complete_age_group_framework_2023.xlsx")
s12_csv <- file.path(folders$Supplementary,"Table_S12_Complete_age_group_framework_2023.csv")
write_book(s12_xlsx,list(
  "Country-age framework"=supp_long,
  "Country comparison"=country_comparison,
  "Detailed age data"=age_detail,
  "Age coverage"=age_coverage,
  "Missingness audit"=missingness_audit,
  "Unrecognized incidence ages"=src$incidence$unrecognized,
  "Unrecognized deaths ages"=src$deaths$unrecognized,
  "Unrecognized DALY ages"=src$dalys$unrecognized,
  "Unrecognized YLL ages"=src$ylls$unrecognized,
  "Unrecognized YLD ages"=src$ylds$unrecognized,
  "Unrecognized population ages"=src$population$unrecognized
))
write_csv(supp_long,s12_csv)

source_file <- file.path(folders$Supplementary,"Figure_12_Source_data.xlsx")
write_book(source_file,list(
  "Framework data"=supp_long,
  "Alluvial transitions"=alluv,
  "Thresholds"=thresholds,
  "Quadrant distribution"=quadrant_distribution
))

# ----------------------------------------------------------------------------
# 11. Final validation and console output
# ----------------------------------------------------------------------------
outputs <- c(figure_file,table_file,s12_xlsx,s12_csv,source_file)
if (any(!file.exists(outputs))) stop("One or more required outputs were not created.", call. = FALSE)
if (any(file.info(outputs)$size<=0)) stop("One or more output files are empty.", call. = FALSE)

message("\n==============================================================")
message("SECTION 3.12 COMPLETED SUCCESSFULLY")
message("==============================================================")
print(age_group_summary,n=Inf)
print(thresholds,n=Inf)
print(pairwise_agreement,n=Inf)
print(country_patterns,n=Inf)
message("\nFigure: ",figure_file)
message("Table: ",table_file)
message("Supplementary table: ",s12_xlsx)
message("Source data: ",source_file)
message("==============================================================")