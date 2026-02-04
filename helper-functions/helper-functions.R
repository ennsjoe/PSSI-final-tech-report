# helper-functions.R
library(DBI)
library(RSQLite)
library(readr)
library(dplyr)
library(here)

# --- Configuration ---
DB_PATH <- here("data", "projects.db")
TABLE_NAME <- "projects"

# --- MANUAL PATH OVERRIDE (if needed) ---
# If the automatic search doesn't find your CSV files, uncomment and edit these lines:
# RAW_CSV <- here("your", "path", "to", "report_project_list.csv")
# PROCESSED_CSV <- here("your", "path", "to", "pssi_form_data.csv")
# Then comment out or delete the "Find CSV files" section below (lines 20-75)

# --- Helper function to find CSV files ---
find_csv_file <- function(filename, possible_paths = NULL) {
  # Default search paths if none provided
  if (is.null(possible_paths)) {
    possible_paths <- c(
      here(filename),                           # Project root
      here("data", filename),                   # data/
      here("data", "raw", filename),           # data/raw/
      here("data", "processed", filename),     # data/processed/
      here("input", filename),                 # input/
      here("inputs", filename)                 # inputs/
    )
  }
  
  # Check each path
  for (path in possible_paths) {
    if (file.exists(path)) {
      return(path)
    }
  }
  
  # Not found
  return(NULL)
}

# --- Find CSV files ---
if (!exists("RAW_CSV") || !exists("PROCESSED_CSV")) {
  message("\n=== Locating CSV Files ===")
  
  RAW_CSV <- find_csv_file("report_project_list.csv")
  if (is.null(RAW_CSV)) {
    stop("\n❌ Cannot find 'report_project_list.csv'\n",
         "   Searched in:\n",
         "   - ", here("report_project_list.csv"), "\n",
         "   - ", here("data", "report_project_list.csv"), "\n",
         "   - ", here("data", "raw", "report_project_list.csv"), "\n",
         "   - ", here("data", "processed", "report_project_list.csv"), "\n",
         "   - ", here("input", "report_project_list.csv"), "\n",
         "   - ", here("inputs", "report_project_list.csv"), "\n\n",
         "   Solution: Run find-csv-files.R to diagnose, or manually set path at top of helper-functions.R")
  }
  
  PROCESSED_CSV <- find_csv_file("pssi_form_data.csv")
  if (is.null(PROCESSED_CSV)) {
    stop("\n❌ Cannot find 'pssi_form_data.csv'\n",
         "   Searched in:\n",
         "   - ", here("pssi_form_data.csv"), "\n",
         "   - ", here("data", "pssi_form_data.csv"), "\n",
         "   - ", here("data", "raw", "pssi_form_data.csv"), "\n",
         "   - ", here("data", "processed", "pssi_form_data.csv"), "\n",
         "   - ", here("input", "pssi_form_data.csv"), "\n",
         "   - ", here("inputs", "pssi_form_data.csv"), "\n\n",
         "   Solution: Run find-csv-files.R to diagnose, or manually set path at top of helper-functions.R")
  }
  
  message("✓ Found tracking CSV: ", RAW_CSV)
  message("✓ Found content CSV: ", PROCESSED_CSV)
}

# --- Database Setup Functions ---

#' Initialize and Merge the SQLite database
#' Merges tracking data (report_project_list.csv) with content (pssi_form_data.csv)
init_database <- function(overwrite = FALSE) {
  message("\n=== Initializing Projects Database ===")
  
  # Files already located by find_csv_file() above
  message("Using files:")
  message("  Tracking: ", RAW_CSV)
  message("  Content: ", PROCESSED_CSV)
  
  # 1. Load Tracking Data (metadata for filtering and organization)
  message("\nLoading tracking data...")
  tracking_df <- read_csv(RAW_CSV, show_col_types = FALSE) %>%
    mutate(project_id = as.character(project_id))
  message("  ✓ Loaded ", nrow(tracking_df), " rows")
  message("  Columns: ", paste(names(tracking_df), collapse = ", "))
  
  # 2. Load Extracted Content (detailed project information)
  message("\nLoading content data...")
  content_df <- read_csv(PROCESSED_CSV, show_col_types = FALSE) %>%
    mutate(project_id = as.character(project_id))
  message("  ✓ Loaded ", nrow(content_df), " rows")
  message("  Columns: ", paste(names(content_df), collapse = ", "))
  
  # 3. Merge the datasets on project_id
  message("\nMerging datasets on 'project_id'...")
  final_df <- tracking_df %>%
    left_join(content_df, by = "project_id", suffix = c("", ".y"))
  
  # 4. Handle duplicate columns (keep tracking version, drop .y versions)
  duplicate_cols <- grep("\\.y$", names(final_df), value = TRUE)
  if (length(duplicate_cols) > 0) {
    message("  Resolving ", length(duplicate_cols), " duplicate columns (keeping tracking data)")
    final_df <- final_df %>% select(-all_of(duplicate_cols))
  }
  
  # 5. Create project_number as backward compatibility alias
  final_df$project_number <- final_df$project_id
  
  # 6. Ensure 'include' column exists for filtering
  if (!"include" %in% names(final_df)) {
    message("  Adding 'include' column (default: 'y')")
    final_df$include <- "y"
  }
  
  message("  ✓ Final dataset: ", nrow(final_df), " rows, ", ncol(final_df), " columns")
  message("  Final columns: ", paste(names(final_df), collapse = ", "))
  
  # 7. Create database directory if needed
  db_dir <- dirname(DB_PATH)
  if (!dir.exists(db_dir)) {
    message("\nCreating database directory: ", db_dir)
    dir.create(db_dir, recursive = TRUE)
  }
  
  # 8. Write to SQLite database
  message("\nWriting to database: ", DB_PATH)
  tryCatch({
    con <- dbConnect(SQLite(), DB_PATH)
    dbWriteTable(con, TABLE_NAME, final_df, overwrite = overwrite)
    
    # Verify
    row_count <- dbGetQuery(con, paste0("SELECT COUNT(*) as n FROM ", TABLE_NAME))$n
    columns <- dbListFields(con, TABLE_NAME)
    
    dbDisconnect(con)
    
    message("\n✅ Database created successfully!")
    message("   Location: ", DB_PATH)
    message("   Projects: ", row_count)
    message("   Columns: ", length(columns))
    
    # Check for critical columns
    required <- c("project_id", "source", "project_title", "broad_header", 
                  "section", "include")
    missing <- setdiff(required, columns)
    if (length(missing) > 0) {
      warning("⚠ Missing required columns: ", paste(missing, collapse = ", "))
    } else {
      message("   ✓ All required columns present")
    }
    
    return(invisible(final_df))
    
  }, error = function(e) {
    if (exists("con")) dbDisconnect(con)
    stop("Database write error: ", e$message)
  })
}

# --- Database Access Functions ---

#' Get database connection
get_db_con <- function(db_path = DB_PATH) {
  if (!file.exists(db_path)) {
    stop("Database not found at: ", db_path, "\n",
         "Run init_database() to create it.")
  }
  dbConnect(SQLite(), db_path)
}

# Alias for backward compatibility
get_db_connection <- function(...) get_db_con(...)

# --- Database Inspection Functions ---

#' Check database status
check_database <- function(db_path = DB_PATH) {
  if (!file.exists(db_path)) {
    return(list(
      exists = FALSE,
      message = "Database not found"
    ))
  }
  
  tryCatch({
    con <- dbConnect(SQLite(), db_path)
    tables <- dbListTables(con)
    
    if (TABLE_NAME %in% tables) {
      row_count <- dbGetQuery(con, paste0("SELECT COUNT(*) as n FROM ", TABLE_NAME))$n
      columns <- dbListFields(con, TABLE_NAME)
      dbDisconnect(con)
      
      return(list(
        exists = TRUE,
        message = paste0("Database contains ", row_count, " projects"),
        row_count = row_count,
        columns = columns
      ))
    } else {
      dbDisconnect(con)
      return(list(
        exists = FALSE,
        message = paste0("Database exists but '", TABLE_NAME, "' table not found")
      ))
    }
  }, error = function(e) {
    if (exists("con")) dbDisconnect(con)
    return(list(
      exists = FALSE,
      message = paste0("Error accessing database: ", e$message)
    ))
  })
}

# --- Appendix Numbering ---
if (!exists(".appendix_counter", envir = .GlobalEnv)) {
  assign(".appendix_counter", 0, envir = .GlobalEnv)
}

new_appendix <- function() {
  current <- get(".appendix_counter", envir = .GlobalEnv) + 1
  assign(".appendix_counter", current, envir = .GlobalEnv)
  LETTERS[current]
}

# --- Bookmark Helper for Cross-References ---
make_bkm <- function(project_id) {
  paste0("project_", gsub("[^A-Za-z0-9]", "_", project_id))
}