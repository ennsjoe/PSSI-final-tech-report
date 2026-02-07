# build-report.R
# Master script to build PSSI Technical Report
# 
# This script:
# 1. Loads _common.R 
# 2. Initializes/updates the project database
# 3. Renders the report with Quarto

library(here)
library(officedown)
library(officer)

# Ensure we're in the project directory
project_root <- here::here()
cat("Working directory:", getwd(), "\n")

# STEP 1: Load _common.R
source(here("_common.R"))

# STEP 2: Initialize/Update Database
# ===================================================================
tryCatch({
  
  # Define the paths where your data actually lives
  db_path    <- here("data", "projects.db")
  csv_list   <- here("report_project_list.csv")
  # CHANGE: Point to the 'processed' subfolder where your extraction script saves
  csv_forms  <- here("data", "processed", "pssi_form_data.csv") 
  
  cat("Checking for data updates...\n")
  
  # Check if the extracted CSV actually exists
  if (!file.exists(csv_forms)) {
    stop(paste("Extraction file not found at:", csv_forms, 
               "\nDid you run project_report_extraction.R first?"))
  }
  
  # If the CSV is newer than the database, or the database is missing, rebuild it.
  db_info   <- file.info(db_path)$mtime
  form_info <- file.info(csv_forms)$mtime
  
  if (is.na(db_info) || form_info > db_info) {
    cat("  New data detected in pssi_form_data.csv. Rebuilding database...\n")
    
    # We pass overwrite=TRUE to ensure the 3682 data replaces the old empty record
    init_database(overwrite = TRUE) 
    
    cat("  ✓ Database updated with latest extractions.\n")
  } else {
    cat("  ✓ Database is already up to date.\n")
  }
  
}, error = function(e) {
  cat("  ⚠ Database Update Error: ", conditionMessage(e), "\n")
  cat("  Attempting to proceed with existing data...\n")
})


# STEP 3: Render Report
# ===================================================================
cat("Rendering main report content with officedown...\n")

# Explicitly call the officedown format
bookdown::render_book("index.Rmd", output_format = "officedown::rdocx_document")

cat("✓ Main report rendered\n\n")


# STEP 4: Add CSAS Technical Report Front Matter
# ===================================================================

# 1. Source the function
source(here("add_frontmatter.R"))

# 2. Get Metadata from index.Rmd
index_content <- readLines(here("index_full.Rmd"))
yaml_bounds <- which(index_content == "---")
meta <- yaml::yaml.load(paste(index_content[(yaml_bounds[1]+1):(yaml_bounds[2]-1)], collapse = "\n"))

# 3. Identify the file
# officedown usually respects the book_filename in _bookdown.yml
book_config <- yaml::yaml.load_file(here("_bookdown.yml"))
out_name <- if(!is.null(book_config$book_filename)) book_config$book_filename else "book"
main_docx_path <- here(paste0(out_name, ".docx"))

# 4. Define final path
final_docx_path <- here(paste0(out_name, "_wFrontMatter.docx"))

# 5. IMPORTANT: Pass your custom template to the post-processor
# This ensures the front matter uses the same styles as your report
meta$reference_docx <- here("templates/custom-reference.docx")

if(file.exists(main_docx_path)) {
  add_frontmatter(
    main_docx   = main_docx_path,
    output_docx = final_docx_path,
    meta        = meta
  )
  cat("✓ Final DFO Technical Report created at:", final_docx_path, "\n")
}

