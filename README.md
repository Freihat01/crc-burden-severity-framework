# Global Colorectal Cancer Burden–Severity Framework

## Overview

This repository contains the R code used for the analyses reported in the study:

**Global colorectal cancer burden, severity, trends, and projections: a multidimensional population-level analysis**

The study evaluates colorectal cancer (CRC) epidemiology across 204 countries and territories using a severity-aware framework that integrates conventional population-level disease burden with health loss per incident case.

The analyses include global and country-level burden estimates, temporal trends, future projections, burden–severity classification, country reclassification and trajectories, age- and sex-specific analyses, socio-demographic associations, and internal validation of DALYs per incident case against the mortality-to-incidence ratio (MIR).

## Data sources

The analyses use data from:

- **Global Burden of Disease (GBD) 2023 study**, Institute for Health Metrics and Evaluation (IHME): incidence, mortality, disability-adjusted life years (DALYs), years of life lost (YLLs), and years lived with disability (YLDs).
- **United Nations World Population Prospects 2024**: population estimates and projections used for future burden modelling.
- **Socio-demographic Index (SDI), 2023**: used to examine the association between socio-demographic development and colorectal cancer severity.

The main historical analysis covers **1990–2023**, with projections extending to **2050**.

## Data availability

Raw GBD data are not redistributed in this repository. Users wishing to reproduce the analyses should obtain the required GBD data directly from the Institute for Health Metrics and Evaluation in accordance with its data-use requirements.

United Nations population data should likewise be obtained from the original UN World Population Prospects source.

Derived results supporting the manuscript are provided with the article and its Supplementary Material where applicable.

## Main measures

The principal burden measures are:

- Incidence
- Mortality
- DALYs
- YLLs
- YLDs
- Age-standardized rates

Disease severity is primarily quantified as:

**DALYs per incident case = DALY number / incident case number**

Additional analyses use:

**Mortality-to-incidence ratio (MIR) = mortality number / incident case number**

The two-dimensional burden–severity framework combines:

- **Population burden:** age-standardized DALY rate
- **Per-case severity:** DALYs per incident case

Countries are classified into four groups:

1. High burden–high severity
2. Low burden–high severity
3. Low burden–low severity
4. High burden–low severity

## Repository structure

The `R/` directory contains the analytical scripts used for the study.

```text
R/
├── 01_Global_Burden_2023.R
├── 02_Global_Severity_Decomposition.R
├── 03_Country_Burden_Severity_Framework.R
├── 04_Country_Severity_Decomposition.R
├── 05_Temporal_Trends_1990_2023.R
├── 06_Global_Projections_2024_2050.R
├── 07_Country_Temporal_Trends_AAPC_1990_2023.R
├── 08_Country_Projections_2024_2050.R
├── 09_Country_Projection_Figure.R
├── 10_Framework_Validation_Reclassification_Stability.R
├── 11_Country_Trajectories_1990_2023.R
├── 12_Sex_Specific_Burden_Severity_Framework_2023.R
├── 13_Age_Group_Burden_Severity_Framework_2023.R
├── 14_SDI_Determinants_of_DALYs_per_Case_2023.R
└── 15_Internal_Validation_DALYs_per_Case_vs_MIR_2023.R
