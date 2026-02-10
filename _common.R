# _common.R - Fixed version that properly loads content from CSV

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

# ===================================================================
# Define Common Paths
# ===================================================================

OUTPUT_DB <<- here('data', 'projects.db')
template_path <<- here('templates', 'project-template.Rmd')
rawdata_path <<- here('data', 'raw')
TABLE_NAME <- "projects"

# ===================================================================
# Source Helper Functions
# ===================================================================

source(here("helper-functions", "helper-functions.R"))

# ===================================================================
# Database Initialization/Validation
# ===================================================================

db_status <- check_database(OUTPUT_DB)
if (db_status$exists) {
  message('✓ Database found: ', db_status$message)
  
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
    # Load metadata from database
    con <- dbConnect(RSQLite::SQLite(), OUTPUT_DB)
    projects_meta <- dbReadTable(con, 'projects')
    dbDisconnect(con)
    
    message('✓ Loaded ', nrow(projects_meta), ' projects from database')
    
    # CRITICAL: Remove content fields from metadata
    # (They lost newlines in SQLite storage)
    content_fields <- c('highlights', 'background', 'methods_findings', 
                        'insights', 'next_steps', 'tables_figures', 'references')
    
    projects_meta <- projects_meta %>%
      select(-any_of(content_fields))
    
    # Load content data DIRECTLY from CSV to preserve line breaks
    PROCESSED_CSV <- here('data', 'processed', 'pssi_form_data.csv')
    
    if (file.exists(PROCESSED_CSV)) {
      message('✓ Loading content data from CSV to preserve line breaks...')
      
      # CRITICAL: Remove content fields from database metadata FIRST
      # (SQLite strips newlines from these fields)
      content_fields <- c('highlights', 'background', 'methods_findings', 
                          'insights', 'next_steps', 'tables_figures', 'references')
      
      projects_meta <- projects_meta %>%
        select(-any_of(content_fields))
      
      # Load content from CSV (preserves newlines)
      projects_content <- read_csv(PROCESSED_CSV, show_col_types = FALSE) %>%
        mutate(project_id = as.character(project_id)) %>%
        select(project_id, source_file, all_of(content_fields))
      
      # Merge metadata with content
      # No conflict now - only CSV has content fields
      projects_df <<- projects_meta %>%
        left_join(projects_content, by = "project_id")
      
      # Verify line breaks are preserved
      test_bg <- projects_df$background[!is.na(projects_df$background)][1]
      if (!is.na(test_bg)) {
        has_newlines <- grepl("\n", test_bg, fixed = TRUE)
        if (has_newlines) {
          n_lines <- length(strsplit(test_bg, "\n", fixed = TRUE)[[1]])
          message('  ✓ Line breaks preserved! Test field has ', n_lines, ' lines')
        } else {
          warning('  ✗ Line breaks NOT found - content may display as single paragraphs')
        }
      }
      
    } else {
      warning('Content CSV not found at: ', PROCESSED_CSV)
      message('Using database only (line breaks may be lost)')
      projects_df <<- projects_meta
    }
    
    message('  Total columns: ', ncol(projects_df))
    message('  Column names: ', paste(head(names(projects_df), 15), collapse = ", "), " ...")
    
  }, error = function(e) {
    stop('Failed to load projects: ', e$message)
  })
}

# ===================================================================
# Transform Project IDs
# ===================================================================

tryCatch({
  if (!"source" %in% names(projects_df)) {
    warning("'source' column not found. Skipping project_id transformation.")
  } else if (!"project_id" %in% names(projects_df)) {
    warning("'project_id' column not found.")
  } else {
    projects_df$source <- as.character(projects_df$source)
    projects_df$project_id <- as.character(projects_df$project_id)
    
    projects_df <- projects_df %>%
      mutate(
        project_id = if_else(
          source == "DFO Science" & !str_starts(project_id, "DFO_"),
          paste0("DFO_", project_id),
          project_id
        )
      )
    
    n_transformed <- sum(grepl("^DFO_", projects_df$project_id), na.rm = TRUE)
    message('  Transformed ', n_transformed, ' DFO Science project IDs')
  }
}, error = function(e) {
  warning('Error transforming project_id: ', e$message)
})

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

message('\n=== _common.R loaded successfully ===\n')