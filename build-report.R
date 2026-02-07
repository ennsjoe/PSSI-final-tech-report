# build-report.R
# Master script to build PSSI Technical Report
# 
# This script:
# 1. Loads _common.R 
# 2. Initializes/updates the project database
# 3. Renders the report with Quarto

library(here)

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

