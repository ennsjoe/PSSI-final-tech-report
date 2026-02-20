# _common.R
# Sourced at the start of every chapter Rmd and by the build script.

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
ensure_package('DBI')
ensure_package('RSQLite')
ensure_package('readr')
ensure_package('flextable')
ensure_package('openxlsx')
ensure_package('bookdown')
ensure_package('officer')
ensure_package('officedown')
ensure_package('yaml')
ensure_package('magick')     # for project page banner image compositing

# ===================================================================
# Define Common Paths
# ===================================================================

OUTPUT_DB     <<- here('data', 'projects.db')
template_path <<- here('templates', 'project-template.Rmd')
rawdata_path  <<- here('data', 'raw')
TABLE_NAME    <- "projects"

# ===================================================================
# Source Helper Functions
# ===================================================================

source(here("helper-functions", "helper-functions.R"))

# ===================================================================
# Database Initialization/Validation
# ===================================================================

db_status <- check_database(OUTPUT_DB)
if (db_status$exists) {
  message('\u2714 Database found: ', db_status$message)
  
  required_cols <- c('project_id', 'source', 'broad_header', 'section', 'include')
  missing <- setdiff(required_cols, db_status$columns)
  if (length(missing) > 0) {
    warning('Database missing required columns: ', paste(missing, collapse=', '))
    message('Rebuilding database...')
    init_database(overwrite = TRUE)
  }
} else {
  message('Database exists but appears corrupted. Rebuilding...')
  init_database(overwrite = TRUE)
}

# ===================================================================
# Load Project Data into Global Environment
# ===================================================================

if (!exists('projects_df')) {
  tryCatch({
    con <- dbConnect(RSQLite::SQLite(), OUTPUT_DB)
    projects_df <<- dbReadTable(con, 'projects')
    dbDisconnect(con)
    
    message('\u2714 Loaded ', nrow(projects_df), ' projects from database')
    message('  Columns: ', paste(names(projects_df), collapse=', '))
    
  }, error = function(e) {
    stop('Failed to load projects from database: ', e$message)
  })
}


# ===================================================================
# Load Icon Map and Banner Paths
# ===================================================================

ICON_MAP_PATH <- here("data", "raw", "icon_map.csv")

if (file.exists(ICON_MAP_PATH)) {
  
  SECTION_ICON_MAP <<- readr::read_csv(ICON_MAP_PATH, show_col_types = FALSE) %>%
    dplyr::filter(!is.na(theme_section), theme_section != "") %>%
    dplyr::mutate(dplyr::across(dplyr::everything(), as.character))
  
  # Named vector kept for backward compatibility
  SECTION_ANCHOR_MAP <<- stats::setNames(
    SECTION_ICON_MAP$anchor,
    SECTION_ICON_MAP$theme_section
  )
  
  message('\u2714 Icon map loaded: ', nrow(SECTION_ICON_MAP), ' sections')
  
} else {
  warning("icon_map.csv not found at: ", ICON_MAP_PATH,
          "\n  Project banners will be skipped.")
  SECTION_ICON_MAP   <<- NULL
  SECTION_ANCHOR_MAP <<- NULL
}

# Folder containing one background PNG per section (named per icon_map$icon_file)
BANNER_DIR <<- here("assets", "project-banner-backgrounds")

if (!dir.exists(BANNER_DIR)) {
  warning("Banner directory not found: ", BANNER_DIR,
          "\n  Create this folder and add your background PNGs.")
}

# ===================================================================
# Set Global Variable for Number of Projects
# ===================================================================

n_projects <<- nrow(projects_df)

# ===================================================================
# Knitr Options
# ===================================================================

knitr::opts_chunk$set(
  echo    = FALSE,
  warning = FALSE,
  message = FALSE,
  fig.path = 'figures/generated/'
)

message('\n=== _common.R loaded successfully ===\n')