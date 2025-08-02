#!/usr/bin/env Rscript

#' Title: Generate baseline demographic and clinical tables for the IIH manuscript
#'
#' This script reads the raw Excel files contained in the repository and
#' produces two baseline tables: one comparing three cohorts (IIH, Population
#' Control, Migraine Control) and another summarising the Venous Sinus Stenting
#' (VSS) cohort.  The three‑group table includes p‑values for each variable
#' comparing the distributions across groups, while the single‑group table is
#' descriptive only.  The results are saved as machine readable CSV files and
#' rendered as simple HTML tables to the `outputs/` directory.  A description
#' of how to regenerate these tables can be found in the repository README.
#'
#' The script is organised into a set of small functions which encapsulate
#' discrete pieces of logic.  At the bottom of the file the functions are
#' invoked in sequence.  The use of functions helps ensure the code is easy
#' to read, unit test, and extend in future.

## Load required libraries ----------------------------------------------------

library(readxl)      # for reading Excel workbooks
library(dplyr)       # for data manipulation
library(tidyr)       # for reshaping data
library(stringr)     # for string processing
library(purrr)       # for functional programming helpers
library(gtsummary)   # for generating publication quality tables
library(gt)          # for saving gtsummary objects to HTML

## Helper: Determine the cohort from a free‑text string ----------------------

#' Given a character vector containing cohort descriptions extracted from the
#' Excel file, this function returns a factor with levels "IIH", "Population
#' Control", and "Migraine Control" based upon simple pattern matching.  In
#' the source data the cohort name may appear at the end of the string or in
#' the middle; here we detect common substrings such as "Population" and
#' "Migraine".  Any string not matching those keywords is treated as IIH by
#' default.
extract_cohort <- function(x) {
  cohort <- case_when(
    str_detect(x, regex("population", ignore_case = TRUE)) ~ "Population Control",
    str_detect(x, regex("migraine", ignore_case = TRUE)) ~ "Migraine Control",
    TRUE ~ "IIH"
  )
  factor(cohort, levels = c("IIH", "Population Control", "Migraine Control"))
}

## Helper: Identify the variable being summarised ----------------------------

#' The raw three‑group table is encoded in a stacked format where the first
#' column (`group_str`) contains both the category of interest and the cohort
#' name.  This function attempts to infer the high level variable name
#' (e.g. "Age categories", "BMI categories", etc.) from the strings in
#' `group_str` and `cat_col`.  A set of regular expressions is used to search
#' for well known patterns.  When no pattern matches, the function returns
#' NA_character_; such rows will subsequently have their variable name
#' populated from `group_str` itself.  It is deliberately generous in its
#' matching because the file contains many variations on the same concept.
infer_variable <- function(group_str, cat_col) {
  # Normalise both inputs to lower case for easier matching
  g <- str_to_lower(group_str %||% "")
  c <- str_to_lower(cat_col %||% "")
  case_when(
    str_detect(g, "\\d+\\s*[-–]\\s*\\d+") ~ "Age categories",
    str_detect(g, "\\bobese|morbidly obese|overweight|underweight|normal weight") ~ "BMI categories",
    str_detect(g, "nonsmoker|ex-smoker|smoker") ~ "Smoking status",
    str_detect(g, "townsend") ~ "Townsend deprivation quintile",
    str_detect(c, "white|black|african|asian|race|south asian|mixed|chinese|middle eastern") ~ "Ethnicity",
    str_detect(c, "back pain|polycystic|osteoa|epilepsy|fibromyalgia|eating disorder|mental illness|sleep apnea|arthritis") ~ "Comorbidities",
    str_detect(c, "headache|migraine") ~ "Outcomes at baseline",
    str_detect(g, "substance use") ~ "Substance use",
    str_detect(g, "analgesic opioid") ~ "Analgesic opioid use",
    str_detect(g, "nonopioid analgesic") ~ "Non‑opioid analgesic use",
    str_detect(g, "triptan use") ~ "Triptan use",
    str_detect(g, "migraine prophylaxis") ~ "Migraine prophylaxis use",
    str_detect(g, "antidepressant use") ~ "Antidepressant use",
    str_detect(g, "age at iih diagnosis|iih duration|age, mean|bmi, mean") ~ "Continuous variables",
    TRUE ~ NA_character_
  )
}

## Helper: Extract the category from the stacked group string -----------------

#' After the cohort has been parsed off, the remaining portion of the
#' `group_str` describes the category.  For example, "16–30 Population Ctl"
#' becomes "16–30".  This helper removes any cohort keywords from the string
#' leaving just the category.  If the resulting string is empty it falls
#' back on the `cat_col` entry provided in the Excel file.
extract_category <- function(group_str, cohort, cat_col) {
  # Remove cohort specific keywords and extraneous whitespace
  cleaned <- group_str %>%
    str_replace_all(regex("population\\s+control|population\\s+ctl|population\\s+conrtol|migraine\\s+control|migraine\\s+ctl|migraine\\s+control for iih|iih cohort", ignore_case = TRUE), "") %>%
    str_trim()
  if (cleaned == "") {
    return(cat_col %||% NA_character_)
  } else {
    return(cleaned)
  }
}

## Main parsing routine for the three‑group table -----------------------------

#' Read and tidy the "Demo & Clin Char Table .xlsx" file.  The function
#' constructs a tidy data frame with four columns: variable, category, group and
#' count.  Rows with missing counts or categories are discarded.  Continuous
#' variables (means and standard deviations) are treated separately later in
#' the workflow and are dropped from the categorical parsing step.
parse_three_group_table <- function(filepath) {
  raw <- read_excel(filepath, col_names = FALSE)
  names(raw) <- c("group_str", "cat_col", "number")
  raw <- raw %>% filter(!is.na(number))
  # Extract cohort and variable names
  tidy <- raw %>%
    mutate(
      cohort = extract_cohort(group_str),
      variable = infer_variable(group_str, cat_col)
    ) %>%
    mutate(
      # If variable is still missing then treat the entire group_str as the
      # variable name and use cat_col as the category
      variable = if_else(is.na(variable), str_trim(group_str), variable),
      category = map2_chr(group_str, cohort, ~ extract_category(.x, .y, cat_col))
    ) %>%
    mutate(count = as.numeric(number)) %>%
    select(variable, category, cohort, count) %>%
    # Drop rows where category is NA (these represent continuous variables or
    # aggregated totals not relevant for categorical summaries)
    filter(!is.na(category), category != "")
  return(tidy)
}

## Generate the three‑group baseline summary table ---------------------------

#' Given the tidy data frame returned by `parse_three_group_table`, create a
#' gtsummary table comparing the distributions across cohorts.  The counts
#' stored in the input represent aggregated frequencies; to leverage
#' `tbl_summary()` and its built in statistical tests we expand the data using
#' `tidyr::uncount()`.  Each row in the expanded data set corresponds to a
#' single individual belonging to a specific cohort, variable and category.  We
#' then pivot the data wider so that each variable forms a separate column.
#' This structure allows gtsummary to compute summary statistics and p‑values
#' for each variable independently.  Chi‑squared tests are used by default for
#' categorical comparisons; Fisher's exact test is automatically invoked by
#' gtsummary when expected counts are low.
create_three_group_table <- function(tidy_data) {
  # Expand counts into individual rows.  Note: this may generate a very large
  # intermediate data frame if the counts are large.  To avoid memory issues
  # during development you can replace uncount() with slice_sample() on a
  # smaller subset; however, for publication ready results it is important to
  # expand all counts so that the p‑values reflect the true sample sizes.
  expanded <- tidy_data %>%
    uncount(weights = count) %>%
    mutate(value = category) %>%
    select(cohort, variable, value)
  # Pivot to wide format: one column per variable.  Missing values remain NA.
  wide <- expanded %>%
    pivot_wider(names_from = variable, values_from = value)
  # Build gtsummary table
  tbl <- wide %>%
    tbl_summary(by = cohort,
                type = all_categorical() ~ "categorical",
                statistic = all_categorical() ~ "{n} ({p}%)",
                missing = "no") %>%
    add_p() %>%
    modify_header(list(label = "Variable", stat_1 = "IIH", stat_2 = "Population Control", stat_3 = "Migraine Control", p.value = "p‑value"))
  return(tbl)
}

## Write outputs for gtsummary objects ---------------------------------------

#' Save a gtsummary table to both CSV and HTML files.  The CSV contains the
#' machine readable representation of the table (variable name, category and
#' formatted statistics) which is useful for downstream analysis or
#' reproduction.  The HTML file provides a human friendly rendering suitable
#' for inclusion in manuscripts or reports.  The filenames should be passed
#' without directory prefixes; the outputs will be written into the
#' repository's `outputs/` folder.
write_table_outputs <- function(tbl, csv_filename, html_filename) {
  # Convert to data frame for CSV; include variable and p‑value columns
  df <- tbl %>% as_tibble() %>% select(-c(id))
  write.csv(df, file = file.path("outputs", csv_filename), row.names = FALSE)
  # Render to HTML using gt
  gt_tbl <- tbl %>% as_gt()
  html_path <- file.path("outputs", html_filename)
  gtsave(gt_tbl, filename = html_path)
}

## Process the VSS baseline table -------------------------------------------

#' Read and summarise the single‑group VSS baseline table.  The input Excel
#' workbook lists a variable or category in the first column and a count in
#' the second column (with some missing categories).  This function
#' interprets the rows into variables (Gender, Age categories, Ethnicity, BMI
#' categories, Stenosis severity) and categories, computes the percentage of
#' patients in each category relative to the total number of VSS patients and
#' returns a tidy data frame ready for presentation.
parse_vss_table <- function(filepath) {
  raw <- read_excel(filepath, col_names = FALSE)
  names(raw) <- c("cat", "count")
  raw <- raw %>% filter(!is.na(count))
  total <- raw$count[1]
  df <- raw %>% slice(-1) %>%
    mutate(
      cat = str_trim(cat),
      variable = case_when(
        cat %in% c("Female", "Male") ~ "Gender",
        str_detect(cat, "\\d+\\s*[-–]\\s*\\d+|\\d+\\+") ~ "Age categories",
        cat %in% c("American Indian or Alaska Native", "Black or African American", "Asian", "White", "None of the above", "Native Hawaiian or Other Pacific Islander", "Other Race", "Prefer not to answer", "Not reported") ~ "Race",
        str_detect(cat, "less than|% or more") ~ "Stenosis severity",
        cat %in% c("non-obese", "obese", "morbidly-obese") ~ "BMI categories",
        TRUE ~ NA_character_
      ),
      category = cat,
      count = as.numeric(count),
      percent = round((count / total) * 100, 1)
    ) %>%
    filter(!is.na(variable))
  df <- df %>% select(variable, category, count, percent)
  return(df)
}

## Render the VSS baseline table to gtsummary format -------------------------

#' Convert the tidy VSS data frame into a gtsummary object.  Because there
#' is only a single group, no p‑values are computed.  The statistic string
#' displays the count and percentage for each category.
create_vss_table <- function(vss_data) {
  # Generate table using gtsummary.  Setting by = NULL ensures a single
  # column output.  Using the statistic list instructs gtsummary to
  # concatenate count and percent into one cell.  The label column holds the
  # variable and category names.
  tbl <- vss_data %>%
    tbl_summary(by = NULL,
                statistic = all_categorical() ~ "{n} ({round(p,1)}%)",
                type = all_categorical() ~ "categorical",
                missing = "no") %>%
    modify_header(list(label = "Variable", stat_1 = "VSS"))
  return(tbl)
}

## Main execution block ------------------------------------------------------

main <- function() {
  # Ensure outputs directory exists
  if (!dir.exists("outputs")) dir.create("outputs", recursive = TRUE)
  # Parse and summarise the three‑group data
  three_group_path <- file.path("data", "Demo & Clin Char Table .xlsx")
  three_tidy <- parse_three_group_table(three_group_path)
  three_table <- create_three_group_table(three_tidy)
  write_table_outputs(three_table, csv_filename = "baseline_three_group.csv", html_filename = "baseline_three_group.html")
  # Parse and summarise the VSS data
  vss_path <- file.path("data", "VSS Baseline.xlsx")
  vss_tidy <- parse_vss_table(vss_path)
  vss_table <- create_vss_table(vss_tidy)
  write_table_outputs(vss_table, csv_filename = "vss_baseline.csv", html_filename = "vss_baseline.html")
}

if (sys.nframe() == 0L) {
  main()
}