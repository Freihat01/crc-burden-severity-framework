# ==============================================================================
# R_03_8_Country_Projections_2024_2050_DAMPED_TREND_FINAL.R
# SECTION 3.8 — COUNTRY-LEVEL PROJECTIONS OF CRC BURDEN, 2024–2050
# Scenario 1: population change only
# Scenario 2: population change + damped 2010–2023 country-specific ASR trend
# No arbitrary AAPC clipping is used.
# ==============================================================================

rm(list = ls()); graphics.off(); cat("\014")
options(stringsAsFactors = FALSE, scipen = 999, dplyr.summarise.inform = FALSE)

required_packages <- c("tidyverse","readxl","openxlsx","janitor","stringi")
missing_packages <- setdiff(required_packages, rownames(installed.packages()))
if (length(missing_packages) > 0) install.packages(missing_packages, dependencies = TRUE)
suppressPackageStartupMessages({
  library(tidyverse); library(readxl); library(openxlsx); library(janitor); library(stringi)
})

EXPECTED_IHME_LOCATIONS <- 204L
AAPC_START_YEAR <- 2010L; AAPC_END_YEAR <- 2023L
BASELINE_YEAR <- 2023L; FIRST_PROJECTION_YEAR <- 2024L; LAST_PROJECTION_YEAR <- 2050L
TREND_ATTENUATION_FACTOR <- 0.90
MAIN_TABLE_TOP_N <- 20L

# ---- project root ----
current_directory <- normalizePath(getwd(), winslash = "/", mustWork = TRUE)
if (tolower(basename(current_directory)) == "r") {
  colorectal_root <- normalizePath(dirname(current_directory), winslash = "/", mustWork = TRUE)
} else if (tolower(basename(current_directory)) == "colorectal cancer") {
  colorectal_root <- current_directory
} else if (dir.exists(file.path(current_directory,"Colorectal cancer"))) {
  colorectal_root <- normalizePath(file.path(current_directory,"Colorectal cancer"), winslash = "/", mustWork = TRUE)
} else stop("Could not identify the Colorectal cancer project root.", call. = FALSE)

incidence_folder <- file.path(colorectal_root,"Incidence")
mortality_folder <- file.path(colorectal_root,"Mortality")
dalys_folder <- file.path(colorectal_root,"DALYs")
population_folder <- file.path(colorectal_root,"Population")
publication_folder <- file.path(colorectal_root,"Publication")
figures_folder <- file.path(publication_folder,"Figures")
tables_folder <- file.path(publication_folder,"Tables")
supplementary_folder <- file.path(publication_folder,"Supplementary")
analysis_objects_folder <- file.path(publication_folder,"Analysis_objects")
for (d in c(figures_folder,tables_folder,supplementary_folder,analysis_objects_folder)) dir.create(d, recursive=TRUE, showWarnings=FALSE)

analysis_objects_file <- file.path(analysis_objects_folder,"Section_3_8_Country_CRC_projections_2024_2050.rds")
table_8_file <- file.path(tables_folder,"Table_8_Country_CRC_Projections_2050.xlsx")
table_s7_file <- file.path(supplementary_folder,"Table_S7_Complete_Country_CRC_Projections_2023_2050.xlsx")
table_s7a_file <- file.path(supplementary_folder,"Table_S7A_Country_AAPC_2010_2023_for_Projections.xlsx")
table_s7b_file <- file.path(supplementary_folder,"Table_S7B_Population_Harmonisation_Diagnostics.xlsx")
summary_file <- file.path(supplementary_folder,"Section_3.8_Analysis_Summary.xlsx")
figure_source_file <- file.path(figures_folder,"Figure_8_Source_data.xlsx")

# ---- helpers ----
read_data_file <- function(file_path){
  ext <- tolower(tools::file_ext(file_path))
  if(ext=="csv") x <- readr::read_csv(file_path,show_col_types=FALSE,progress=FALSE,guess_max=500000)
  else if(ext %in% c("xlsx","xls")) x <- readxl::read_excel(file_path,guess_max=500000)
  else if(ext=="txt") x <- readr::read_delim(file_path,delim="\t",show_col_types=FALSE,progress=FALSE,guess_max=500000)
  else stop(paste("Unsupported file type:",file_path),call.=FALSE)
  janitor::clean_names(x)
}
find_column <- function(data,candidates,description,required=TRUE){
  hit <- intersect(candidates,names(data)); if(length(hit)>0) return(hit[1])
  if(required) stop(paste0("Could not identify ",description," column. Available: ",paste(names(data),collapse=", ")),call.=FALSE)
  NULL
}
safe_numeric <- function(x){ if(is.numeric(x)) return(as.numeric(x)); readr::parse_number(as.character(x),na=c("","NA","N/A","NaN","NULL","-"),locale=locale(grouping_mark=",")) }
normalise_text <- function(x){ as.character(x) |> stringi::stri_trans_general("Latin-ASCII") |> str_to_lower() |> str_replace_all("&"," and ") |> str_replace_all("[’'`]","") |> str_replace_all("[^a-z0-9]+"," ") |> str_squish() }
canonical_country_key <- function(x){
  key <- normalise_text(x)
  case_when(
    key %in% c("bolivia plurinational state of","plurinational state of bolivia") ~ "bolivia",
    key=="brunei darussalam" ~ "brunei", key=="cabo verde" ~ "cape verde",
    key %in% c("congo","republic of congo","congo republic") ~ "republic of the congo",
    key %in% c("democratic republic of the congo","congo democratic republic","democratic republic of congo","congo the democratic republic of the","drc") ~ "democratic republic of the congo",
    key %in% c("cote divoire","cote d ivoire") ~ "cote divoire", key=="czech republic" ~ "czechia",
    key %in% c("democratic peoples republic of korea","dem peoples republic of korea","dem peoples rep of korea","dem peoples rep korea","korea democratic peoples republic of","korea dem peoples rep","north korea") ~ "north korea",
    key %in% c("republic of korea","korea republic of","south korea") ~ "south korea",
    key %in% c("iran islamic republic of","islamic republic of iran") ~ "iran",
    key %in% c("lao peoples democratic republic","lao pdr") ~ "laos",
    key %in% c("micronesia federated states of","micronesia fed states of","micronesia fed states","federated states of micronesia") ~ "federated states of micronesia",
    key %in% c("moldova republic of","republic of moldova") ~ "moldova", key=="russian federation" ~ "russia",
    key=="syrian arab republic" ~ "syria", key %in% c("united republic of tanzania","tanzania united republic of") ~ "tanzania",
    key=="turkiye" ~ "turkey", key %in% c("united kingdom of great britain and northern ireland","great britain") ~ "united kingdom",
    key %in% c("united states of america","usa") ~ "united states",
    key %in% c("venezuela bolivarian republic of","bolivarian republic of venezuela") ~ "venezuela",
    key=="viet nam" ~ "vietnam", key %in% c("state of palestine","palestine state of","west bank and gaza") ~ "palestine",
    key %in% c("taiwan province of china","china taiwan province of china") ~ "taiwan", key=="swaziland" ~ "eswatini",
    TRUE ~ key)
}
standardise_sex <- function(x){ v<-normalise_text(x); case_when(v %in% c("both","both sexes","both sex","total")~"Both",v %in% c("male","males")~"Male",v %in% c("female","females")~"Female",TRUE~as.character(x)) }
standardise_measure <- function(x){ v<-normalise_text(x); case_when(str_detect(v,"incidence|incident")~"Incident cases",str_detect(v,"mortality|death|deaths")~"Deaths",str_detect(v,"daly")~"DALYs",TRUE~as.character(x)) }
standardise_metric <- function(metric,age=NA_character_){ m<-normalise_text(metric); a<-normalise_text(age); case_when(str_detect(m,"number|count")~"Number",str_detect(m,"rate|asr") & (str_detect(a,"age standard")|str_detect(m,"age standard|asr"))~"ASR",TRUE~NA_character_) }
find_matching_file <- function(folder,outcome_pattern,metric_pattern){
  files <- list.files(folder,pattern="\\.(csv|xlsx|xls|txt)$",full.names=TRUE,ignore.case=TRUE)
  files <- files[!str_detect(basename(files),"^~\\$")]; nm <- normalise_text(basename(files))
  hit <- files[str_detect(nm,outcome_pattern)&str_detect(nm,metric_pattern)&str_detect(nm,"1990.*2023")&str_detect(nm,"both")&!str_detect(nm,"age groups|agegroups")]
  if(length(hit)==0) stop(paste("Required file not found in",folder,"for",outcome_pattern,metric_pattern),call.=FALSE)
  nmh <- normalise_text(basename(hit)); pref <- if(metric_pattern=="number") hit[str_detect(nmh,"all ages|allages")] else hit[str_detect(nmh,"age standardized|agestandardized|age standardised")]
  if(length(pref)>0) hit<-pref; hit[1]
}
prepare_ihme_file <- function(file_path,expected_outcome,expected_metric){
  raw<-read_data_file(file_path)
  lc<-find_column(raw,c("location","location_name","country","country_name"),"location")
  yc<-find_column(raw,c("year","year_id"),"year"); sc<-find_column(raw,c("sex","sex_name"),"sex",FALSE)
  ac<-find_column(raw,c("age","age_name"),"age",FALSE); mc<-find_column(raw,c("measure","measure_name"),"measure",FALSE)
  mtc<-find_column(raw,c("metric","metric_name"),"metric",FALSE); vc<-find_column(raw,c("val","value","mean","estimate"),"value")
  raw |> transmute(location=as.character(.data[[lc]]),country_key=canonical_country_key(.data[[lc]]),year=as.integer(.data[[yc]]),sex=if(!is.null(sc)) standardise_sex(.data[[sc]]) else "Both",age=if(!is.null(ac)) as.character(.data[[ac]]) else NA_character_,outcome=if(!is.null(mc)) standardise_measure(.data[[mc]]) else expected_outcome,metric=if(!is.null(mtc)) standardise_metric(.data[[mtc]],age) else expected_metric,value=safe_numeric(.data[[vc]])) |> filter(sex=="Both",outcome==expected_outcome,year>=1990,year<=2023) |> filter((expected_metric=="Number" & (metric=="Number"|is.na(metric))) | (expected_metric=="ASR" & (metric=="ASR"|is.na(metric)))) |> select(location,country_key,year,value) |> distinct(country_key,year,.keep_all=TRUE)
}
calculate_country_aapc <- function(data){
  d<-data |> filter(year>=AAPC_START_YEAR,year<=AAPC_END_YEAR,!is.na(value),value>0) |> arrange(year)
  if(nrow(d)<3) return(tibble(n_years=nrow(d),first_year=NA_integer_,last_year=NA_integer_,aapc=NA_real_,aapc_lower_95=NA_real_,aapc_upper_95=NA_real_,p_value=NA_real_,r_squared=NA_real_))
  mod<-lm(log(value)~year,data=d); b<-unname(coef(mod)[["year"]]); ci<-confint(mod,"year",level=.95); sm<-summary(mod)
  tibble(n_years=nrow(d),first_year=min(d$year),last_year=max(d$year),aapc=100*(exp(b)-1),aapc_lower_95=100*(exp(ci[1])-1),aapc_upper_95=100*(exp(ci[2])-1),p_value=sm$coefficients["year","Pr(>|t|)"],r_squared=sm$r.squared)
}

# ---- locate and import IHME files ----
inc_num_f<-find_matching_file(incidence_folder,"incidence","number"); inc_rate_f<-find_matching_file(incidence_folder,"incidence","rate")
dth_num_f<-find_matching_file(mortality_folder,"death|mortality","number"); dth_rate_f<-find_matching_file(mortality_folder,"death|mortality","rate")
dal_num_f<-find_matching_file(dalys_folder,"daly","number"); dal_rate_f<-find_matching_file(dalys_folder,"daly","rate")
message("Selected IHME files:\n",paste(c(inc_num_f,inc_rate_f,dth_num_f,dth_rate_f,dal_num_f,dal_rate_f),collapse="\n"))

baseline_coverage <- bind_rows(
  prepare_ihme_file(inc_num_f,"Incident cases","Number") |> mutate(outcome="Incident cases"),
  prepare_ihme_file(dth_num_f,"Deaths","Number") |> mutate(outcome="Deaths"),
  prepare_ihme_file(dal_num_f,"DALYs","Number") |> mutate(outcome="DALYs")) |> filter(year==BASELINE_YEAR) |> transmute(outcome,location,country_key,baseline_2023=value) |> distinct(outcome,country_key,.keep_all=TRUE)
bc<-baseline_coverage |> count(outcome,name="countries"); print(bc,n=Inf)
if(nrow(bc)!=3 || any(bc$countries!=EXPECTED_IHME_LOCATIONS)) stop("Baseline coverage is not 204 locations for each outcome.",call.=FALSE)

asr_data <- bind_rows(
  prepare_ihme_file(inc_rate_f,"Incident cases","ASR") |> mutate(outcome="Incident cases"),
  prepare_ihme_file(dth_rate_f,"Deaths","ASR") |> mutate(outcome="Deaths"),
  prepare_ihme_file(dal_rate_f,"DALYs","ASR") |> mutate(outcome="DALYs"))
country_aapc <- asr_data |> group_by(outcome,location,country_key) |> group_modify(~calculate_country_aapc(.x)) |> ungroup() |> mutate(annual_log_trend=log(1+aapc/100))
ac<-country_aapc |> count(outcome,name="countries"); print(ac,n=Inf)
if(nrow(ac)!=3 || any(ac$countries!=EXPECTED_IHME_LOCATIONS) || any(!is.finite(country_aapc$aapc)) || any(!is.finite(country_aapc$annual_log_trend))) stop("AAPC estimation failed for one or more country-outcome combinations.",call.=FALSE)

write.xlsx(list(Country_AAPC_2010_2023=country_aapc,Method=tibble(Item=c("Trend period","Trend model","Projection treatment","Annual attenuation factor"),Description=c("2010–2023","log(age-standardized rate) ~ calendar year","No arbitrary clipping; projected log trend is progressively attenuated","0.90"))),table_s7a_file,overwrite=TRUE)

# ---- population files ----
pop_files<-list.files(population_folder,pattern="\\.(csv|xlsx|xls|txt)$",full.names=TRUE,ignore.case=TRUE); pop_nm<-normalise_text(basename(pop_files))
future_file<-pop_files[str_detect(pop_nm,"2024.*2050|wpp.*2024.*2050")][1]
base_file<-pop_files[str_detect(pop_nm,"2023") & str_detect(pop_nm,"pop")][1]
if(is.na(future_file)||is.na(base_file)) stop("Could not locate 2023 and/or 2024–2050 population files.",call.=FALSE)
prepare_population <- function(file,years){
  raw<-read_data_file(file); lc<-find_column(raw,c("location","location_name","country","country_name","region_subregion_country_or_area"),"population location")
  yc<-find_column(raw,c("year","year_id"),"population year",FALSE); pc<-find_column(raw,c("population","population_persons","total_population","pop"),"population")
  raw |> transmute(source_location=as.character(.data[[lc]]),country_key=canonical_country_key(.data[[lc]]),year=if(!is.null(yc)) as.integer(.data[[yc]]) else if(length(years)==1) years else NA_integer_,population=safe_numeric(.data[[pc]])) |> filter(year %in% years,!is.na(population),population>0) |> group_by(country_key,year) |> summarise(source_location=first(source_location),population=first(population),.groups="drop")
}
pop2023_raw<-prepare_population(base_file,BASELINE_YEAR); popfuture_raw<-prepare_population(future_file,FIRST_PROJECTION_YEAR:LAST_PROJECTION_YEAR)
ihme_locations<-baseline_coverage |> distinct(location,country_key); if(nrow(ihme_locations)!=EXPECTED_IHME_LOCATIONS) stop("IHME location list is not 204.",call.=FALSE)
pop2023 <- ihme_locations |> left_join(pop2023_raw |> select(country_key,population_2023=population),by="country_key") |> left_join(popfuture_raw |> filter(year==FIRST_PROJECTION_YEAR) |> select(country_key,population_2024=population),by="country_key") |> mutate(population_2023_source=case_when(!is.na(population_2023)~"UN 2023 population file",is.na(population_2023)&!is.na(population_2024)~"2024 value used because 2023 value unavailable",TRUE~"Missing"),population_2023=coalesce(population_2023,population_2024)) |> select(location,country_key,population_2023,population_2023_source)
popfuture <- crossing(ihme_locations,year=FIRST_PROJECTION_YEAR:LAST_PROJECTION_YEAR) |> left_join(popfuture_raw |> select(country_key,year,population_future=population),by=c("country_key","year"))
unmatched2023<-pop2023 |> filter(is.na(population_2023)); unmatchedfuture<-popfuture |> filter(is.na(population_future))
write.xlsx(list(IHME_locations=ihme_locations,Matched_2023=pop2023,Unmatched_2023=unmatched2023,Unmatched_future=unmatchedfuture,Extra_UN_2023=anti_join(pop2023_raw,ihme_locations,by="country_key"),Extra_UN_future=anti_join(popfuture_raw,ihme_locations,by="country_key")),table_s7b_file,overwrite=TRUE)
if(nrow(unmatched2023)>0 || nrow(unmatchedfuture)>0) stop(paste0("Population matching incomplete. Open: ",table_s7b_file),call.=FALSE)

# ---- baseline analytical dataset ----
baseline_complete <- baseline_coverage |> left_join(pop2023 |> select(country_key,population_2023,population_2023_source),by="country_key") |> left_join(country_aapc |> select(outcome,country_key,aapc,aapc_lower_95,aapc_upper_95,p_value,r_squared,annual_log_trend),by=c("outcome","country_key"))
if(any(!complete.cases(baseline_complete[,c("baseline_2023","population_2023","aapc","annual_log_trend")]))) stop("Baseline analytical dataset is incomplete.",call.=FALSE)

# ---- projections ----
future_projections <- baseline_complete |> select(outcome,location,country_key,baseline_2023,population_2023,population_2023_source,aapc,aapc_lower_95,aapc_upper_95,p_value,r_squared,annual_log_trend) |> crossing(year=FIRST_PROJECTION_YEAR:LAST_PROJECTION_YEAR) |> left_join(popfuture |> select(country_key,year,population=population_future),by=c("country_key","year")) |> mutate(years_after_2023=year-BASELINE_YEAR,population_ratio=population/population_2023,scenario_1=baseline_2023*population_ratio,epidemiological_trend_factor=exp(annual_log_trend*(1-TREND_ATTENUATION_FACTOR^years_after_2023)/(1-TREND_ATTENUATION_FACTOR)),scenario_2=scenario_1*epidemiological_trend_factor,scenario_1_absolute_change=scenario_1-baseline_2023,scenario_1_percentage_change=100*(scenario_1/baseline_2023-1),scenario_2_absolute_change=scenario_2-baseline_2023,scenario_2_percentage_change=100*(scenario_2/baseline_2023-1),scenario_2_minus_scenario_1=scenario_2-scenario_1,scenario_2_vs_scenario_1_percent=100*(scenario_2/scenario_1-1)) |> group_by(outcome,year) |> mutate(scenario_1_rank=min_rank(desc(scenario_1)),scenario_2_rank=min_rank(desc(scenario_2)),scenario_difference_rank=min_rank(desc(scenario_2_minus_scenario_1))) |> ungroup()
baseline_rows <- baseline_complete |> transmute(outcome,location,country_key,year=BASELINE_YEAR,baseline_2023,population_2023,population_2023_source,aapc,aapc_lower_95,aapc_upper_95,p_value,r_squared,annual_log_trend,population=population_2023,years_after_2023=0L,population_ratio=1,epidemiological_trend_factor=1,scenario_1=baseline_2023,scenario_2=baseline_2023,scenario_1_absolute_change=0,scenario_1_percentage_change=0,scenario_2_absolute_change=0,scenario_2_percentage_change=0,scenario_2_minus_scenario_1=0,scenario_2_vs_scenario_1_percent=0) |> group_by(outcome) |> mutate(scenario_1_rank=min_rank(desc(scenario_1)),scenario_2_rank=min_rank(desc(scenario_2)),scenario_difference_rank=NA_integer_) |> ungroup()
complete_projections <- bind_rows(baseline_rows,future_projections) |> arrange(outcome,location,year)
expected_rows<-EXPECTED_IHME_LOCATIONS*3L*length(BASELINE_YEAR:LAST_PROJECTION_YEAR)
if(nrow(complete_projections)!=expected_rows || any(!is.finite(complete_projections$scenario_1)) || any(!is.finite(complete_projections$scenario_2)) || any(complete_projections$scenario_1<0|complete_projections$scenario_2<0)) stop("Projection validation failed.",call.=FALSE)

# ---- S7 ----
table_s7_complete <- complete_projections |> transmute(Country=location,Outcome=outcome,Year=year,Population=population,Population_2023=population_2023,Population_2023_source=population_2023_source,Population_ratio_vs_2023=population_ratio,Observed_2023_baseline=baseline_2023,Recent_AAPC_percent=aapc,AAPC_lower_95_percent=aapc_lower_95,AAPC_upper_95_percent=aapc_upper_95,AAPC_P_value=p_value,Annual_log_trend_beta=annual_log_trend,Trend_attenuation_factor=TREND_ATTENUATION_FACTOR,Years_after_2023=years_after_2023,Epidemiological_trend_factor=epidemiological_trend_factor,Scenario_1_population_change_only=scenario_1,Scenario_1_absolute_change_from_2023=scenario_1_absolute_change,Scenario_1_percentage_change_from_2023=scenario_1_percentage_change,Scenario_1_country_rank=scenario_1_rank,Scenario_2_population_plus_attenuated_trend=scenario_2,Scenario_2_absolute_change_from_2023=scenario_2_absolute_change,Scenario_2_percentage_change_from_2023=scenario_2_percentage_change,Scenario_2_country_rank=scenario_2_rank,Scenario_2_minus_Scenario_1=scenario_2_minus_scenario_1,Scenario_2_vs_Scenario_1_percent=scenario_2_vs_scenario_1_percent,Scenario_difference_rank=scenario_difference_rank)
write.xlsx(list(Complete_2023_2050=table_s7_complete,Selected_years=table_s7_complete |> filter(Year %in% c(2023,2030,2040,2050)),Methods=tibble(Item=c("Scenario 1","Scenario 2","Baseline year","Projection period","AAPC period","Trend model","Trend attenuation factor","Population interpretation","Population ageing separately modelled"),Description=c("Observed 2023 burden × future total population / 2023 total population",paste0("Scenario 1 × exp{beta × [1 - ",TREND_ATTENUATION_FACTOR,"^t] / [1 - ",TREND_ATTENUATION_FACTOR,"]}; beta=ln(1+AAPC/100)"),"2023","2024–2050","2010–2023 country-specific ASR","Damped log-linear trend; no AAPC clipping",as.character(TREND_ATTENUATION_FACTOR),"Total population change only","No"))),table_s7_file,overwrite=TRUE)

# ---- 2050 dataset and Table 8 ----
projection_2050 <- complete_projections |> filter(year==LAST_PROJECTION_YEAR) |> transmute(country=location,country_key,outcome,baseline_2023,population_2023,population_2050=population,aapc,annual_log_trend,trend_attenuation_factor=TREND_ATTENUATION_FACTOR,epidemiological_trend_factor,scenario_1,scenario_2,scenario_1_absolute_change,scenario_1_percentage_change,scenario_2_absolute_change,scenario_2_percentage_change,scenario_2_minus_scenario_1,scenario_2_vs_scenario_1_percent,scenario_1_rank,scenario_2_rank,scenario_difference_rank)
fmt <- function(d) d |> transmute(Outcome=outcome,Country=country,Observed_2023=baseline_2023,`2050 Scenario 1`=scenario_1,`Scenario 1 change, %`=scenario_1_percentage_change,`2050 Scenario 2`=scenario_2,`Scenario 2 change, %`=scenario_2_percentage_change,`Scenario 2 vs 1, %`=scenario_2_vs_scenario_1_percent,`2010–2023 AAPC, %`=aapc)
table8 <- list(Incident_cases=fmt(projection_2050 |> filter(outcome=="Incident cases") |> arrange(desc(scenario_2_absolute_change)) |> slice_head(n=MAIN_TABLE_TOP_N)),Deaths=fmt(projection_2050 |> filter(outcome=="Deaths") |> arrange(desc(scenario_2_absolute_change)) |> slice_head(n=MAIN_TABLE_TOP_N)),DALYs=fmt(projection_2050 |> filter(outcome=="DALYs") |> arrange(desc(scenario_2_absolute_change)) |> slice_head(n=MAIN_TABLE_TOP_N)),Scenario_difference=fmt(projection_2050 |> arrange(desc(abs(scenario_2_minus_scenario_1))) |> slice_head(n=MAIN_TABLE_TOP_N)))
wb<-createWorkbook(); for(nm in names(table8)){ addWorksheet(wb,nm); writeDataTable(wb,nm,table8[[nm]]); freezePane(wb,nm,firstRow=TRUE); setColWidths(wb,nm,cols=1:ncol(table8[[nm]]),widths="auto") }; saveWorkbook(wb,table_8_file,overwrite=TRUE)

# ---- figure source + RDS ----
figure_8_source_data <- projection_2050 |> select(Country=country,Outcome=outcome,Scenario_1=scenario_1,Scenario_2=scenario_2,Scenario_2_vs_Scenario_1_percent=scenario_2_vs_scenario_1_percent)
write.xlsx(figure_8_source_data,figure_source_file,overwrite=TRUE)
saveRDS(list(projection_2050=projection_2050,complete_projections=complete_projections,country_aapc=country_aapc,population_2023_matched=pop2023,population_future_matched=popfuture),analysis_objects_file)

# ---- summary ----
analysis_summary <- projection_2050 |> group_by(outcome) |> summarise(Countries=n(),Total_observed_2023=sum(baseline_2023),Total_projected_2050_Scenario_1=sum(scenario_1),Total_projected_2050_Scenario_2=sum(scenario_2),Scenario_1_global_percent_change=100*(Total_projected_2050_Scenario_1/Total_observed_2023-1),Scenario_2_global_percent_change=100*(Total_projected_2050_Scenario_2/Total_observed_2023-1),Countries_increasing_Scenario_1=sum(scenario_1_percentage_change>0),Countries_declining_Scenario_1=sum(scenario_1_percentage_change<0),Countries_increasing_Scenario_2=sum(scenario_2_percentage_change>0),Countries_declining_Scenario_2=sum(scenario_2_percentage_change<0),Median_Scenario_2_vs_1_percent=median(scenario_2_vs_scenario_1_percent),IQR_lower_Scenario_2_vs_1_percent=quantile(scenario_2_vs_scenario_1_percent,.25),IQR_upper_Scenario_2_vs_1_percent=quantile(scenario_2_vs_scenario_1_percent,.75),.groups="drop")
write.xlsx(list(Projection_summary=analysis_summary,Projection_validation=tibble(Check=c("IHME countries/territories","Outcomes","Projection years","Complete projection rows","Negative projections","Trend method","Trend attenuation factor","AAPC clipping"),Result=c(as.character(n_distinct(complete_projections$location)),as.character(n_distinct(complete_projections$outcome)),paste0(min(complete_projections$year),"–",max(complete_projections$year)),as.character(nrow(complete_projections)),as.character(sum(complete_projections$scenario_1<0|complete_projections$scenario_2<0)),"Damped log-linear ASR trend",as.character(TREND_ATTENUATION_FACTOR),"None"))),summary_file,overwrite=TRUE)

message("\n==============================================================")
message("SECTION 3.8 COMPLETED SUCCESSFULLY")
message("==============================================================")
print(analysis_summary,n=Inf)
message("\nTrend method: damped 2010–2023 country-specific ASR trend")
message("Annual attenuation factor: ",TREND_ATTENUATION_FACTOR)
message("AAPC clipping: none")
message("\nUpdated RDS: ",analysis_objects_file)
message("Main Table 8: ",table_8_file)
message("Supplementary S7: ",table_s7_file)
message("Supplementary S7A: ",table_s7a_file)
message("Population diagnostics S7B: ",table_s7b_file)
message("Summary: ",summary_file)
message("Figure 8 source: ",figure_source_file)
message("\nNow rerun the separate Figure 8 plotting script.")
message("==============================================================")


# ==============================================================================
# 29. FINAL HYBRID FIGURE 8
#
# This section preserves the user's original hybrid Figure 8 design:
# LEFT: dumbbell plots for the 5 largest positive and 5 largest negative
#       absolute Scenario 2 minus Scenario 1 differences.
# RIGHT: percentage-difference maps.
#
# The only methodological update is that Scenario 2 now uses the corrected
# damped 2010–2023 country-specific ASR trend generated above.
# ==============================================================================

figure_8_file <- file.path(
  figures_folder,
  "Figure_8_Country_CRC_projections_2050.tiff"
)

figure_8_source_file <- file.path(
  figures_folder,
  "Figure_8_Source_data.csv"
)

# ------------------------------------------------------------------------------
# Prepare projection data in the exact structure expected by the ORIGINAL
# hybrid Figure 8 code below.
#
# The Section 3.8 analytical object uses:
#   outcome = "Incident cases", "Deaths", "DALYs"
#   scenario_2_minus_scenario_1
#   scenario_2_vs_scenario_1_percent
#
# The original Figure 8 code expects:
#   outcome = "Incidence", "Deaths", "DALYs"
#   absolute_difference
#   percentage_difference
# ------------------------------------------------------------------------------

projection_2050_figure <- projection_2050 |>
  transmute(
    country = as.character(country),
    outcome = dplyr::recode(
      as.character(outcome),
      "Incident cases" = "Incidence",
      .default = as.character(outcome)
    ),
    scenario_1 = as.numeric(scenario_1),
    scenario_2 = as.numeric(scenario_2),
    absolute_difference = as.numeric(
      scenario_2_minus_scenario_1
    ),
    percentage_difference = as.numeric(
      scenario_2_vs_scenario_1_percent
    )
  )

if (nrow(projection_2050_figure) != 204 * 3) {
  stop(
    paste0(
      "\nExpected 612 country-outcome observations for Figure 8, but ",
      nrow(projection_2050_figure),
      " were found."
    ),
    call. = FALSE
  )
}

invalid_figure_values <- projection_2050_figure |>
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

if (nrow(invalid_figure_values) > 0) {
  print(invalid_figure_values, n = Inf)
  stop(
    "\nInvalid values were detected in the Figure 8 projection dataset.",
    call. = FALSE
  )
}

# From this point onward, the original Figure 8 code can run unchanged.
projection_2050 <- projection_2050_figure

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
    "B. Relative effect of attenuated epidemiological trend on incident cases"
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
    "D. Relative effect of attenuated epidemiological trend on deaths"
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
    "F. Relative effect of attenuated epidemiological trend on DALYs"
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
        "Scenario 1 incorporates population change only. Scenario 2 additionally ",
        "incorporates attenuated country-specific 2010–2023 age-standardized-rate ",
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
message("Projection calculations use the corrected damped-trend Scenario 2.")
message("Table 8 was regenerated from the corrected projections.")
message("Supplementary Table S7 was regenerated from the corrected projections.")
message("==============================================================")