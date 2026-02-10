# _common.R - CSV-ONLY VERSION (no database)
# Loads directly from report_project_list.csv + pssi_form_data.csv

# ===================================================================
# Package Loading with Auto-Install
# ===================================================================

ensure_package <- function(pkg) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    message('Installing missing package: ', pkg)
    install.packages(pkg, repos = 'https://cloud.r-project.org/')
    library(pkg, character.only = TRUE)
  }
}

ensure_package('here')
ensure_package('tidyverse')
ensure_package('knitr')
ensure_package('readr')
ensure_package('flextable')
ensure_package('openxlsx')
ensure_package('bookdown')
ensure_package('officer')
ensure_package('officedown')
ensure_package('yaml')

# ===================================================================
# Define Common Paths
# ===================================================================

template_path <<- here('templates', 'project-template.Rmd')
rawdata_path <<- here('data', 'raw')

# CSV file paths
TRACKING_CSV <- here('data', 'raw', 'report_project_list.csv')
CONTENT_CSV <- here('data', 'processed', 'pssi_form_data.csv')

# ===================================================================
# Source Helper Functions (for tables only)
# ===================================================================

# Load just the table-making function from helper-functions
if (file.exists(here("helper-functions", "helper-functions.R"))) {
  # Source it but ignore the database parts
  suppressMessages({
    source(here("helper-functions", "helper-functions.R"))
  })
}

# ===================================================================
# Load Project Data DIRECTLY from CSVs
# ===================================================================

cat("\n")
cat(rep("=", 80), "\n", sep = "")
cat("LOADING PROJECT DATA FROM CSV FILES ONLY\n")
cat(rep("=", 80), "\n\n", sep = "")

# Step 1: Load tracking/metadata CSV
cat("STEP 1: Loading tracking data\n")
cat(rep("-", 80), "\n", sep = "")

if (!file.exists(TRACKING_CSV)) {
  stop("Tracking CSV not found: ", TRACKING_CSV)
}

tracking_df <- read_csv(TRACKING_CSV, show_col_types = FALSE) %>%
  mutate(project_id = as.character(project_id))

cat("✓ Loaded tracking CSV:", nrow(tracking_df), "projects\n")
cat("  File:", TRACKING_CSV, "\n")
cat("  Columns:", paste(names(tracking_df), collapse = ", "), "\n\n")

# Step 2: Load content CSV (has line breaks)
cat("STEP 2: Loading content data (with line breaks preserved)\n")
cat(rep("-", 80), "\n", sep = "")

if (!file.exists(CONTENT_CSV)) {
  stop("Content CSV not found: ", CONTENT_CSV)
}

content_df <- read_csv(CONTENT_CSV, show_col_types = FALSE) %>%
  mutate(project_id = as.character(project_id)) %>%
  # PRESERVE original project_id as project_number for figure directories
  mutate(project_number = project_id)

cat("✓ Loaded content CSV:", nrow(content_df), "projects\n")
cat("  File:", CONTENT_CSV, "\n")
cat("  Columns:", paste(names(content_df), collapse = ", "), "\n")

# Verify line breaks are preserved
test_bg <- content_df$background[!is.na(content_df$background)][1]
if (!is.na(test_bg)) {
  has_newlines <- grepl("\n", test_bg, fixed = TRUE)
  if (has_newlines) {
    n_lines <- length(strsplit(test_bg, "\n", fixed = TRUE)[[1]])
    cat("  ✓ Line breaks preserved! Sample field has", n_lines, "lines\n")
  } else {
    warning("  ⚠ Line breaks NOT found in content CSV")
  }
}
cat("\n")

# Step 3: Merge tracking + content
cat("STEP 3: Merging tracking and content data\n")
cat(rep("-", 80), "\n", sep = "")

# Define content fields
content_fields <- c('highlights', 'background', 'methods_findings', 
                    'insights', 'next_steps', 'tables_figures', 'references')

# Remove content fields from tracking (if they exist)
tracking_clean <- tracking_df %>%
  select(-any_of(content_fields))

# Merge - keep project_number from content_df
projects_df <- tracking_clean %>%
  left_join(
    content_df %>% select(project_id, source_file, project_number, all_of(content_fields)), 
    by = "project_id"
  )

cat("✓ Merged data:", nrow(projects_df), "projects\n")
cat("  Columns:", ncol(projects_df), "\n")
cat("  ✓ Preserved project_number from content CSV\n\n")

# Step 4: Add DFO_ prefix to DFO Science project IDs
cat("STEP 4: Transforming project IDs (for display)\n")
cat(rep("-", 80), "\n", sep = "")

projects_df <- projects_df %>%
  mutate(
    project_id = if_else(
      source == "DFO Science" & !str_starts(project_id, "DFO_"),
      paste0("DFO_", project_id),
      project_id
    )
  )

n_transformed <- sum(grepl("^DFO_", projects_df$project_id), na.rm = TRUE)
cat("✓ Transformed", n_transformed, "DFO Science project IDs\n")
cat("  Note: project_number (for figures) remains unchanged\n\n")

# Step 5: Summary statistics
cat("STEP 5: Data summary\n")
cat(rep("-", 80), "\n", sep = "")

cat("Total projects:", nrow(projects_df), "\n")
cat("Projects by source:\n")
source_counts <- table(projects_df$source, useNA = "always")
for (src in names(source_counts)) {
  cat("  ", src, ": ", source_counts[src], "\n", sep = "")
}

cat("\nDFO Science projects by include:\n")
dfo_include <- projects_df %>%
  filter(source == "DFO Science") %>%
  count(include, .drop = FALSE)
print(dfo_include)

cat("\nDFO Science projects to render (include='y'):\n")
dfo_to_render <- projects_df %>%
  filter(source == "DFO Science", include %in% c("y", "Y"))
cat("  ", nrow(dfo_to_render), " projects\n", sep = "")

# Check for project 3682
has_3682 <- any(grepl("3682", projects_df$project_id, fixed = TRUE))
if (has_3682) {
  p3682 <- projects_df %>% filter(grepl("3682", project_id, fixed = TRUE))
  cat("\n✓ Project 3682 found:\n")
  cat("    ID:", p3682$project_id, "\n")
  cat("    Number:", p3682$project_number, "(for figures)\n")
  cat("    Source:", p3682$source, "\n")
  cat("    Include:", p3682$include, "\n")
  cat("    Background:", nchar(p3682$background), "chars\n")
  
  if (p3682$source == "DFO Science" && p3682$include %in% c("y", "Y")) {
    cat("    ✓ Will be rendered\n")
    cat("    ✓ Figures at: figures/project_figures/", p3682$project_number, "/\n", sep = "")
  } else {
    cat("    ✗ Will NOT be rendered (check source/include)\n")
  }
} else {
  cat("\n⚠ Project 3682 NOT found in data\n")
}

cat("\n")
cat(rep("=", 80), "\n", sep = "")
cat("DATA LOADED SUCCESSFULLY FROM CSV FILES\n")
cat(rep("=", 80), "\n\n", sep = "")

# ===================================================================
# Set global variables
# ===================================================================

n_projects <<- nrow(projects_df)

# ===================================================================
# Knitr Options
# ===================================================================

knitr::opts_chunk$set(
  echo = FALSE,
  warning = FALSE,
  message = FALSE,
  fig.path = 'figures/generated/'
)

message('=== _common.R loaded successfully (CSV-only mode) ===\n')