library(xml2)
library(dplyr)
library(purrr)
library(tibble)
library(stringr)
library(DBI)
library(RSQLite)
library(here)

# ----------------------------------------------------------
# Define standard paths - data/ folder is the single source of truth
# ----------------------------------------------------------
if (!exists("OUTPUT_DB")) {
  OUTPUT_DB <- here('data', 'projects.db')
}
if (!exists("OUTPUT_CSV")) {
  OUTPUT_CSV <- here('data', 'pssi_form_data.csv')
}
if (!exists("INPUT_DIR")) {
  INPUT_DIR <- here('data', 'word_docs')
}

# ENSURE DIRECTORY EXISTS BEFORE CONNECTION
if (!dir.exists(dirname(OUTPUT_DB))) {
  dir.create(dirname(OUTPUT_DB), recursive = TRUE, showWarnings = FALSE)
}

message("Using paths:")
message("  Database: ", OUTPUT_DB)
message("  CSV:      ", OUTPUT_CSV)
message("  Input:    ", INPUT_DIR)
message("")

# Fields to extract - Updated to include 'Project Number' to prevent Rmd errors
FIELD_LABELS <- c(
  "Project ID (number)" = "project_id",
  "Project ID" = "project_id",
  "Project Number" = "project_id", # Maps to project_id for 10_project-reports.Rmd
  "Project Title" = "project_title",
  "Project Leads" = "project_leads",
  "Collaborations and External Partners" = "collaborations",
  "Location (if applicable)" = "location",
  "Location" = "location",
  "Salmon species (if applicable)" = "species",
  "Salmon species" = "species",
  "Waterbodies (if applicable)" = "waterbodies",
  "Waterbodies" = "waterbodies",
  "Life history phases (if applicable)" = "life_history",
  "Life history phases" = "life_history",
  "Region" = "region",
  "Stock (if applicable)" = "stock",
  "Stock" = "stock",
  "Population (if applicable)" = "population",
  "Population" = "population",
  "Conservation Unit (if applicable)" = "cu",
  "Conservation Unit" = "cu",
  "Highlights" = "highlights",
  "Background" = "background",
  "Methods and Findings" = "methods_findings",
  "Insights" = "insights",
  "Next Steps" = "next_steps",
  "Tables and Figures" = "tables_figures",
  "References" = "references"
)

OUTPUT_FIELDS <- unique(FIELD_LABELS)
W_NS <- c(w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

# [Helper functions: fix_encoding, extract_text_with_breaks, extract_content_controls, 
#  extract_from_tables, extract_section_content remain the same as your source]

# ----------------------------------------------------------
# Main Processing - Updated with NULL handling and Connection Safety
# ----------------------------------------------------------
extract_docx_content <- function(docx_path) {
  filename <- basename(docx_path)
  message("Processing: ", filename)
  
  temp_dir <- tempfile(); dir.create(temp_dir)
  on.exit(unlink(temp_dir, recursive = TRUE))
  
  tryCatch({ unzip(docx_path, exdir = temp_dir) }, error = function(e) return(NULL))
  doc_path <- file.path(temp_dir, "word", "document.xml")
  if (!file.exists(doc_path)) return(NULL)
  
  doc_xml <- read_xml(doc_path, encoding = "UTF-8")
  results <- setNames(as.list(rep(NA_character_, length(OUTPUT_FIELDS))), OUTPUT_FIELDS)
  
  cc_data <- extract_content_controls(doc_xml)
  table_data <- extract_from_tables(doc_xml)
  section_data <- extract_section_content(doc_xml)
  
  for (f in OUTPUT_FIELDS) {
    # Layered check: Content Controls > Tables > Section Patterns
    if (!is.null(cc_data[[f]]) && !is.na(cc_data[[f]])) {
      results[[f]] <- cc_data[[f]]
    } 
    if (is.na(results[[f]]) && !is.null(table_data[[f]]) && !is.na(table_data[[f]])) {
      results[[f]] <- table_data[[f]]
    }
    if (is.na(results[[f]]) && !is.null(section_data[[f]]) && !is.na(section_data[[f]])) {
      results[[f]] <- section_data[[f]]
    }
  }
  
  tibble(source_file = filename, extraction_date = as.character(Sys.Date())) %>%
    bind_cols(as_tibble(results))
}

extract_all_forms <- function(input_dir = INPUT_DIR, output_csv = OUTPUT_CSV, output_db = OUTPUT_DB) {
  docx_files <- list.files(input_dir, pattern = "\\.docx$", full.names = TRUE, ignore.case = TRUE)
  docx_files <- docx_files[!grepl("^~\\$", basename(docx_files))]
  
  if (length(docx_files) == 0) stop("No Word documents found in: ", input_dir)
  
  results <- compact(map(docx_files, extract_docx_content))
  if (length(results) == 0) stop("No data extracted.")
  combined <- bind_rows(results)
  
  # Save CSV
  con_csv <- file(output_csv, "wb")
  writeBin(charToRaw('\uFEFF'), con_csv)
  write.table(combined, con_csv, row.names = FALSE, na = "", sep = ",", quote = TRUE, fileEncoding = "UTF-8", col.names = TRUE)
  close(con_csv)
  
  # Save Database with explicit error handling for file locks
  tryCatch({
    con_db <- dbConnect(SQLite(), output_db)
    dbWriteTable(con_db, "projects", combined, overwrite = TRUE)
    dbDisconnect(con_db)
    message("\nSuccess! Database and CSV updated in /data.")
  }, error = function(e) {
    stop("Could not connect to database: ", e$message, 
         "\nCheck if the file 'projects.db' is open in another program.")
  })
  
  return(combined)
}

extract_all_forms()