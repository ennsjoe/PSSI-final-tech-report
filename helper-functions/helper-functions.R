# helper-functions.R
# Functions for managing project data in SQLite database
# Used by pssi-report for Pacific Salmon Strategy Initiative report
# Updated for csasdown compatibility

library(DBI)
library(RSQLite)
library(readr)
library(dplyr)
library(here)
library(csasdown)

# --- Configuration using here() for project-relative paths ---
DB_PATH <- here("data", "projects.db")
CSV_PATH <- here("data", "raw", "report_project_list.csv")
TABLE_NAME <- "projects"

# Print paths on load so user knows where files are
message("Database path: ", DB_PATH)
message("CSV path: ", CSV_PATH)

# --- Database Setup Functions ---

#' Initialize the SQLite database from CSV
#' @param csv_path Path to the CSV file (default: CSV_PATH)
#' @param db_path Path to the SQLite database (default: DB_PATH)
#' @param overwrite If TRUE, recreate the table even if it exists
#' @return Invisible NULL
init_database <- function(csv_path = CSV_PATH, db_path = DB_PATH, overwrite = FALSE) {
  # Show where we're creating the database
  message("Initializing database at: ", db_path)
  message("Reading CSV from: ", csv_path)
  
  # Create output directory if it doesn't exist
  db_dir <- dirname(db_path)
  if (!dir.exists(db_dir)) {
    dir.create(db_dir, recursive = TRUE)
    message("Created directory: ", db_dir)
  }
  
  # Connect to database (creates file if doesn't exist)
  con <- dbConnect(RSQLite::SQLite(), db_path)
  on.exit(dbDisconnect(con))
  
  # Check if table exists and whether to overwrite
  table_exists <- dbExistsTable(con, TABLE_NAME)
  
  if (table_exists && !overwrite) {
    message("Database table '", TABLE_NAME, "' already exists. Use overwrite=TRUE to recreate.")
    return(invisible(NULL))
  }
  
  # Read CSV file
  projects <- read_csv(csv_path, show_col_types = FALSE)
  
  # Clean column names (remove BOM if present)
  names(projects) <- gsub("^\uFEFF", "", names(projects))
  
  # Write to database (overwrite if table exists)
  if (table_exists) {
    dbRemoveTable(con, TABLE_NAME)
  }
  
  dbWriteTable(con, TABLE_NAME, projects)
  
  # Create indexes for common queries
  dbExecute(con, paste0("CREATE INDEX IF NOT EXISTS idx_section ON ", TABLE_NAME, " (section)"))
  dbExecute(con, paste0("CREATE INDEX IF NOT EXISTS idx_broad_header ON ", TABLE_NAME, " (broad_header)"))
  dbExecute(con, paste0("CREATE INDEX IF NOT EXISTS idx_include ON ", TABLE_NAME, " (include)"))
  dbExecute(con, paste0("CREATE INDEX IF NOT EXISTS idx_source ON ", TABLE_NAME, " (source)"))
  
  n_rows <- nrow(projects)
  message("SUCCESS: Database initialized with ", n_rows, " projects")
  message("Database location: ", db_path)
  
  return(invisible(NULL))
}

#' Get a database connection
#' @param db_path Path to the SQLite database
#' @return Database connection object
get_db_connection <- function(db_path = DB_PATH) {
  if (!file.exists(db_path)) {
    stop("Database not found at: ", db_path, "\nRun init_database() first.")
  }
  dbConnect(RSQLite::SQLite(), db_path)
}

# --- Query Functions ---

#' Get all projects from the database
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only return projects where include = 'y'
#' @return Data frame of all projects
get_all_projects <- function(db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  if (include_only) {
    query <- paste0("SELECT * FROM ", TABLE_NAME, " WHERE include = 'y'")
  } else {
    query <- paste0("SELECT * FROM ", TABLE_NAME)
  }
  
  dbGetQuery(con, query)
}

#' Get projects by section
#' @param section_name Name of the section to filter by
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only return projects where include = 'y'
#' @return Data frame of matching projects
get_projects_by_section <- function(section_name, db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0("SELECT * FROM ", TABLE_NAME, " WHERE section = ?")
  if (include_only) {
    query <- paste0(query, " AND include = 'y'")
  }
  
  dbGetQuery(con, query, params = list(section_name))
}

#' Get projects by broad header
#' @param header_name Name of the broad header to filter by
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only return projects where include = 'y'
#' @return Data frame of matching projects
get_projects_by_header <- function(header_name, db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0("SELECT * FROM ", TABLE_NAME, " WHERE broad_header = ?")
  if (include_only) {
    query <- paste0(query, " AND include = 'y'")
  }
  
  dbGetQuery(con, query, params = list(header_name))
}

#' Get projects by source (BCSRIF or DFO)
#' @param source_name Source to filter by ('BCSRIF' or 'DFO')
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only return projects where include = 'y'
#' @return Data frame of matching projects
get_projects_by_source <- function(source_name, db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0("SELECT * FROM ", TABLE_NAME, " WHERE source = ?")
  if (include_only) {
    query <- paste0(query, " AND include = 'y'")
  }
  
  dbGetQuery(con, query, params = list(source_name))
}

#' Get a single project by project number
#' @param project_num Project number to look up
#' @param db_path Path to the SQLite database
#' @return Data frame with single project (or empty if not found)
get_project <- function(project_num, db_path = DB_PATH) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0("SELECT * FROM ", TABLE_NAME, " WHERE project_number = ?")
  dbGetQuery(con, query, params = list(project_num))
}

#' Search projects by keyword in title or description
#' @param keyword Search term
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only return projects where include = 'y'
#' @return Data frame of matching projects
search_projects <- function(keyword, db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  search_term <- paste0("%", keyword, "%")
  query <- paste0(
    "SELECT * FROM ", TABLE_NAME, 
    " WHERE (project_title LIKE ? OR description LIKE ?)"
  )
  if (include_only) {
    query <- paste0(query, " AND include = 'y'")
  }
  
  dbGetQuery(con, query, params = list(search_term, search_term))
}

#' Get summary counts by section
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only count projects where include = 'y'
#' @return Data frame with section counts
get_section_summary <- function(db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0(
    "SELECT section, COUNT(*) as n_projects FROM ", TABLE_NAME
  )
  if (include_only) {
    query <- paste0(query, " WHERE include = 'y'")
  }
  query <- paste0(query, " GROUP BY section ORDER BY n_projects DESC")
  
  dbGetQuery(con, query)
}

#' Get list of unique sections
#' @param db_path Path to the SQLite database
#' @return Character vector of section names
get_sections <- function(db_path = DB_PATH) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0(
    "SELECT DISTINCT section FROM ", TABLE_NAME, 
    " WHERE section IS NOT NULL AND include = 'y' ORDER BY section"
  )
  result <- dbGetQuery(con, query)
  result$section
}

#' Get list of unique broad headers
#' @param db_path Path to the SQLite database
#' @return Character vector of broad header names
get_broad_headers <- function(db_path = DB_PATH) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0(
    "SELECT DISTINCT broad_header FROM ", TABLE_NAME, 
    " WHERE broad_header IS NOT NULL AND include = 'y' ORDER BY broad_header"
  )
  result <- dbGetQuery(con, query)
  result$broad_header
}

# --- Table Generation Functions for csasdown ---

#' Create a projects table using csasdown::csas_table
#' @param df Data frame with project data
#' @param num_col Column name for project number
#' @param title_col Column name for project title
#' @param source_col Column name for source
#' @param caption Table caption (optional)
#' @return A csas_table object
csas_projects_table <- function(df,
                                num_col = "project_number",
                                title_col = "project_title",
                                source_col = "source",
                                caption = NULL) {
  
  if (nrow(df) == 0) {
    return(NULL)
  }
  
  # Prepare the data
  
  display_df <- df |>
    dplyr::mutate(
      dplyr::across(dplyr::all_of(source_col), ~ stringr::str_squish(as.character(.x))),
      dplyr::across(dplyr::all_of(num_col), ~ as.character(.x)),
      dplyr::across(dplyr::all_of(title_col), ~ stringr::str_squish(as.character(.x)))
    ) |>
    dplyr::select(
      Project = dplyr::all_of(num_col),
      Source = dplyr::all_of(source_col),
      Title = dplyr::all_of(title_col)
    ) |>
    dplyr::arrange(Source, Project)
  
  csasdown::csas_table(
    display_df,
    caption = caption,
    format = "pandoc"
  )
}

#' Create a section table of projects using csasdown::csas_table
#' @param section_name Name of the section
#' @param db_path Path to the SQLite database
#' @param caption Optional table caption (defaults to section name)
#' @return csas_table object
make_section_table <- function(section_name, db_path = DB_PATH, caption = NULL) {
  projects <- get_projects_by_section(section_name, db_path)
  
  if (nrow(projects) == 0) {
    return(NULL)
  }
  
  # Select and rename columns for display
  table_data <- projects |>
    dplyr::select(
      Title = project_title,
      Source = source
    )
  
  table_caption <- if (is.null(caption)) {
    paste0("Associated Projects: ", section_name)
  } else {
    caption
  }
  
  csasdown::csas_table(
    table_data,
    caption = table_caption,
    format = "pandoc"
  )
}

#' Create a detailed table with descriptions using csasdown::csas_table
#' @param projects Data frame of projects
#' @param caption Table caption
#' @return csas_table object
make_detailed_table <- function(projects, caption = "Projects") {
  if (nrow(projects) == 0) {
    return(NULL)
  }
  
  table_data <- projects |>
    dplyr::select(
      `Project #` = project_number,
      Title = project_title,
      Source = source,
      Location = location
    )
  
  csasdown::csas_table(
    table_data,
    caption = caption,
    format = "pandoc"
  )
}

#' Create a summary table by broad header using csasdown::csas_table
#' @param db_path Path to the SQLite database
#' @return csas_table object
make_header_summary_table <- function(db_path = DB_PATH) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0(
    "SELECT broad_header as 'Research Theme', COUNT(*) as 'Number of Projects' ",
    "FROM ", TABLE_NAME, " ",
    "WHERE include = 'y' AND broad_header IS NOT NULL ",
    "GROUP BY broad_header ORDER BY COUNT(*) DESC"
  )
  
  result <- dbGetQuery(con, query)
  
  csasdown::csas_table(
    result,
    caption = "Projects by Research Theme",
    format = "pandoc"
  )
}

# --- Utility Functions ---

#' Create valid Word bookmark name from project ID
#' 
#' Word bookmarks have specific requirements:
#' - Must start with a letter
#' - Can only contain letters, numbers, and underscores
#' - Maximum 40 characters
#' 
#' @param project_id The project ID to convert to a bookmark name
#' @return A valid Word bookmark name string
make_bkm <- function(project_id) {
  # Convert to character if needed
  id <- as.character(project_id)
  
  
  # Replace any non-alphanumeric characters with underscores
  bkm <- gsub("[^A-Za-z0-9]", "_", id)
  
  # Ensure it starts with a letter (prefix with "proj_" if it starts with a number)
  if (grepl("^[0-9]", bkm)) {
    bkm <- paste0("proj_", bkm)
  }
  
  # Truncate to 40 characters (Word bookmark limit)
  if (nchar(bkm) > 40) {
    bkm <- substr(bkm, 1, 40)
  }
  
  return(bkm)
}

#' Check database status
#' @param db_path Path to the SQLite database
#' @return List with database info
check_database <- function(db_path = DB_PATH) {
  if (!file.exists(db_path)) {
    return(list(
      exists = FALSE,
      path = db_path,
      message = paste0("Database does not exist at: ", db_path, "\nRun init_database() to create it.")
    ))
  }
  
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  n_total <- dbGetQuery(con, paste0("SELECT COUNT(*) as n FROM ", TABLE_NAME))$n
  n_included <- dbGetQuery(con, paste0("SELECT COUNT(*) as n FROM ", TABLE_NAME, " WHERE include = 'y'"))$n
  
  list(
    exists = TRUE,
    path = db_path,
    total_projects = n_total,
    included_projects = n_included,
    message = paste0("Database at: ", db_path, "\nContains ", n_total, " projects (", n_included, " included in report)")
  )
}

#' Rebuild the database from CSV
#' @param csv_path Path to the CSV file
#' @param db_path Path to the SQLite database
#' @return Invisible NULL
rebuild_database <- function(csv_path = CSV_PATH, db_path = DB_PATH) {
  init_database(csv_path, db_path, overwrite = TRUE)
}

#' Show current path configuration
#' @return Named list of paths
show_paths <- function() {
  paths <- list(
    project_root = here(),
    database = DB_PATH,
    csv_file = CSV_PATH
  )
  
  message("Project root: ", paths$project_root)
  message("Database: ", paths$database)
  message("CSV file: ", paths$csv_file)
  message("Database exists: ", file.exists(DB_PATH))
  message("CSV exists: ", file.exists(CSV_PATH))
  
  invisible(paths)
}
# --- Appendix Numbering Functions ---

# Counter for appendix letters (initialized on first use)
if (!exists(".appendix_counter", envir = .GlobalEnv)) {
  assign(".appendix_counter", 0, envir = .GlobalEnv)
}

#' Generate next appendix letter
#' 
#' Automatically increments and returns the next appendix letter (A, B, C, etc.)
#' Used in chapter headers to automatically number appendices.
#' 
#' @return The next appendix letter as a character string
#' @examples
#' # In your Rmd file header:
#' # APPENDIX `r new_appendix()`. Title
new_appendix <- function() {
  current <- get(".appendix_counter", envir = .GlobalEnv)
  current <- current + 1
  assign(".appendix_counter", current, envir = .GlobalEnv)
  LETTERS[current]
}

#' Reset appendix counter
#' 
#' Resets the appendix counter back to 0. Useful for debugging or when
#' regenerating the document.
#' 
#' @return Invisible NULL
reset_appendix <- function() {
  assign(".appendix_counter", 0, envir = .GlobalEnv)
  invisible(NULL)
}