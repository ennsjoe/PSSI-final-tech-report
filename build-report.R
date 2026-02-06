# build-report.R
# Master script to build PSSI Technical Report
# 
# This script:
# 1. Loads _common.R 
# 2. Initializes/updates the project database
# 3. Renders the report with Quarto

library(here)
library(bookdown)

# Ensure we're in the project directory
project_root <- here::here()
cat("Working directory:", getwd(), "\n")

# STEP 1: Load _common.R
source(here("_common.R"))

# STEP 2: Initialize/Update Database
tryCatch({
  # Check if database needs rebuilding
  if (exists("check_database")) {
    db_status <- check_database()
    
    if (!db_status$exists) {
      cat("  Database not found, creating new one...\n")
      init_database(overwrite = TRUE)
    } else {
      cat("  ", db_status$message, "\n")
      
      # Check if CSV files are newer than database
      db_path <- here("data", "projects.db")
      csv1_path <- here("report_project_list.csv")
      csv2_path <- here("pssi_form_data.csv")
      
      if (file.exists(db_path) && file.exists(csv1_path) && file.exists(csv2_path)) {
        db_mtime <- file.info(db_path)$mtime
        csv1_mtime <- file.info(csv1_path)$mtime
        csv2_mtime <- file.info(csv2_path)$mtime
        
        if (csv1_mtime > db_mtime || csv2_mtime > db_mtime) {
          cat("  CSV files updated since last build. Rebuilding database...\n")
          init_database(overwrite = TRUE)
        } else {
          cat("  ✓ Database is up-to-date\n")
        }
      }
    }
  }
}, error = function(e) {
  cat("  ⚠ Database initialization warning:", conditionMessage(e), "\n")
  cat("  Attempting to continue...\n")
})
cat("\n")


# STEP 3: Render Report with bookdown (DOCX)
# ===================================================================

# bookdown compiles the multi-chapter book listed in _bookdown.yml
# using the output format specified in _output.yml.

# Render the book (uses index.Rmd by convention;
bookdown::render_book()





