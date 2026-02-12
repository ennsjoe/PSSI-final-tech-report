# _common.R
# Sourced at the top of index.Rmd and by build-report.R.
# Loads packages, sets paths, merges project data, and loads icon maps.

# ============================================================================
# 1. Packages
# ============================================================================

ensure_package <- function(pkg) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    message("Installing missing package: ", pkg)
    install.packages(pkg, repos = "https://cloud.r-project.org/")
    library(pkg, character.only = TRUE)
  }
}

ensure_package("here")
ensure_package("tidyverse")
ensure_package("knitr")
ensure_package("readr")
ensure_package("DBI")
ensure_package("RSQLite")
ensure_package("flextable")
ensure_package("openxlsx")
ensure_package("officer")
ensure_package("officedown")
ensure_package("yaml")

# ============================================================================
# 2. Paths
# ============================================================================

template_path  <<- here("templates", "project-template.Rmd")
rawdata_path   <<- here("data", "raw")
OUTPUT_DB      <<- here("data", "projects.db")
TABLE_NAME      <- "projects"

TRACKING_CSV <- here("data", "raw",       "report_project_list.csv")
CONTENT_CSV  <- here("data", "processed", "pssi_form_data.csv")

# ============================================================================
# 3. Helper functions
# ============================================================================

helper_file <- here("helper-functions", "helper-functions.R")

if (file.exists(helper_file)) {
  suppressMessages(source(helper_file))
} else {
  stop("helper-functions.R not found at: ", helper_file)
}

# ============================================================================
# 4. Load and merge project data from CSVs
# ============================================================================

message("\n", strrep("=", 70))
message("LOADING PROJECT DATA")
message(strrep("=", 70))

# -- 4a. Tracking CSV (metadata, section, include flags) ---------------------

if (!file.exists(TRACKING_CSV)) stop("Tracking CSV not found: ", TRACKING_CSV)

tracking_df <- read_csv(TRACKING_CSV, show_col_types = FALSE) %>%
  mutate(project_id = as.character(project_id))

message("✓ Tracking CSV:  ", nrow(tracking_df), " rows  |  ", TRACKING_CSV)

# -- 4b. Content CSV (form extractions: highlights, background, etc.) --------

if (!file.exists(CONTENT_CSV)) stop("Content CSV not found: ", CONTENT_CSV)

content_df <- read_csv(CONTENT_CSV, show_col_types = FALSE) %>%
  mutate(
    project_id     = as.character(project_id),
    project_number = project_id   # preserve bare number for figure directories
  )

message("✓ Content CSV:   ", nrow(content_df), " rows  |  ", CONTENT_CSV)

# Verify line breaks are preserved in content fields
test_bg <- content_df$background[!is.na(content_df$background)][1]
if (!is.na(test_bg)) {
  if (grepl("\n", test_bg, fixed = TRUE)) {
    message("✓ Line breaks preserved in content CSV")
  } else {
    warning("Line breaks NOT detected in content CSV — check extraction output")
  }
}

# -- 4c. Merge ----------------------------------------------------------------

content_fields <- c("highlights", "background", "methods_findings",
                    "insights", "next_steps", "tables_figures", "references")

projects_df <- tracking_df %>%
  select(-any_of(content_fields)) %>%
  left_join(
    content_df %>%
      select(project_id, source_file, project_number, all_of(content_fields)),
    by = "project_id"
  )

message("✓ Merged:        ", nrow(projects_df), " projects, ",
        ncol(projects_df), " columns")

# -- 4d. Add DFO_ prefix to DFO Science project IDs -------------------------
# project_id becomes the display ID and Word bookmark name (e.g. DFO_042).
# project_number retains the bare numeric ID for figure directory lookups.

projects_df <- projects_df %>%
  mutate(
    project_id = if_else(
      source == "DFO Science" & !str_starts(project_id, "PSSI_"),
      paste0("PSSI_", project_id),
      project_id
    )
  )

n_pssi <- sum(grepl("^PSSI_", projects_df$project_id), na.rm = TRUE)
message("✓ PSSI_ prefix:  ", n_pssi, " DFO Science IDs transformed")
message("  (project_number unchanged for figure directory lookups)")

# -- 4e. Summary --------------------------------------------------------------

message("\nProjects by source:")
source_counts <- sort(table(projects_df$source), decreasing = TRUE)
for (nm in names(source_counts)) {
  message("  ", nm, ": ", source_counts[nm])
}

dfo_y <- projects_df %>%
  filter(source == "DFO Science", include %in% c("y", "Y")) %>%
  nrow()
message("\nDFO Science projects flagged include='y': ", dfo_y)

message(strrep("=", 70))
message("PROJECT DATA LOADED\n")

# ============================================================================
# 5. Global variables
# ============================================================================

n_projects <<- nrow(projects_df)

# ============================================================================
# 6. Icon maps for project page banners
# ============================================================================
# Reads assets/icons/icon_map.csv to build SECTION_ICON_MAP, SECTION_COLOUR_MAP,
# and SECTION_ANCHOR_MAP. The CSV and PNG files are managed manually —
# no icon generation script is needed.
# Columns required in icon_map.csv: section, icon_file, colour, anchor

load_icon_maps()

# ============================================================================
# 7. Knitr options
# ============================================================================

knitr::opts_chunk$set(
  echo    = FALSE,
  warning = FALSE,
  message = FALSE,
  fig.path = "figures/generated/"
)

message("=== _common.R loaded successfully ===\n")