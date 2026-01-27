# =============================================================================
# Extract PSSI Word Document Forms to CSV
# =============================================================================
# This script reads .docx files containing PSSI project reporting forms
# and extracts the form data into a CSV spreadsheet.
#
# Required packages: here, xml2, dplyr, stringr, purrr
# =============================================================================

# Install packages if needed
required_packages <- c("here", "xml2", "dplyr", "stringr", "purrr", "DBI", "RSQLite")
new_packages <- required_packages[!(required_packages %in% installed.packages()[,"Package"])]
if(length(new_packages)) install.packages(new_packages)

library(here)
library(xml2)
library(dplyr)
library(stringr)
library(purrr)
library(DBI)
library(RSQLite)

# Word XML namespaces
NS <- c(w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

# Placeholder text patterns to treat as NA
PLACEHOLDER_PATTERNS <- c(
  "Click or tap here to enter text",
  "Click or tap here to enter text.",
  "click or tap here to enter text"
)

#' Get all text from an XML node, preserving paragraph breaks
get_text <- function(node) {
  # Find all paragraphs
  paragraphs <- xml_find_all(node, ".//w:p", NS)
  
  if (length(paragraphs) == 0) {
    # Fallback: just get all text nodes
    text_nodes <- xml_find_all(node, ".//w:t", NS)
    return(paste(xml_text(text_nodes), collapse = ""))
  }
  
  # Extract text from each paragraph using a for loop (safer than sapply with XML)
  para_texts <- character(length(paragraphs))
  for (i in seq_along(paragraphs)) {
    text_nodes <- xml_find_all(paragraphs[[i]], ".//w:t", NS)
    para_texts[i] <- paste(xml_text(text_nodes), collapse = "")
  }
  
  # Filter out empty paragraphs and join with double newline (markdown paragraph break)
  para_texts <- para_texts[nchar(trimws(para_texts)) > 0]
  paste(para_texts, collapse = "\n\n")
}

#' Get text without preserving paragraph breaks (for single-value fields)
get_text_simple <- function(node) {
  text_nodes <- xml_find_all(node, ".//w:t", NS)
  paste(xml_text(text_nodes), collapse = "")
}

#' Normalize whitespace in text (preserves paragraph breaks)
normalize <- function(text) {
  # Normalize multiple newlines to double newline (paragraph break)
  text <- gsub("\\r\\n", "\n", text)  # Windows line endings
  text <- gsub("\\n{3,}", "\n\n", text)  # Max 2 newlines
  # Normalize spaces (but not newlines)
  text <- gsub("[ \\t]+", " ", text)
  # Trim leading/trailing whitespace
  trimws(text)
}

#' Clean a value - return NA if blank or placeholder text
#' 
#' @param value The text value to clean
#' @return Cleaned value or NA if blank/placeholder
clean_value <- function(value) {
  if (is.null(value)) return(NA_character_)
  
  # Normalize whitespace (preserving paragraph breaks)
  cleaned <- normalize(value)
  
  # Return NA if empty
  if (nchar(cleaned) == 0) {
    return(NA_character_)
  }
  
  # Check for placeholder text (case-insensitive)
  for (pattern in PLACEHOLDER_PATTERNS) {
    if (grepl(pattern, cleaned, ignore.case = TRUE)) {
      # Remove the placeholder text
      cleaned <- gsub(pattern, "", cleaned, ignore.case = TRUE)
    }
  }
  
  # Normalize again after removal
  cleaned <- normalize(cleaned)
  
  # Check if empty after removal
  if (nchar(cleaned) == 0) {
    return(NA_character_)
  }
  
  return(cleaned)
}

#' Extract form data from a single PSSI Word document
#'
#' @param docx_path Path to the .docx file
#' @return A named list with extracted field values
extract_form_data <- function(docx_path) {
  
  # Create temp directory and unzip
  temp_dir <- tempfile()
  dir.create(temp_dir)
  on.exit(unlink(temp_dir, recursive = TRUE))
  
  unzip(docx_path, exdir = temp_dir)
  
  # Read document.xml
  doc_xml <- read_xml(file.path(temp_dir, "word", "document.xml"))
  
  # Initialize results
  results <- list(
    source_file = basename(docx_path)
  )
  
  # Find all tables
  tables <- xml_find_all(doc_xml, "//w:tbl", NS)
  
  # Process Table 3: General Project Info (usually index 3)
  if (length(tables) >= 3) {
    table <- tables[[3]]
    rows <- xml_find_all(table, ".//w:tr", NS)
    
    # Row 2 (index 2): values for Project ID and Title
    if (length(rows) >= 2) {
      row_text <- normalize(get_text(rows[[2]]))
      # Extract project ID (4 digits at start)
      match <- str_match(row_text, "^(\\d{4})\\s*(.+)$")
      if (!is.na(match[1, 2])) {
        results$project_id <- clean_value(match[1, 2])
        results$project_title <- clean_value(match[1, 3])
      }
    }
    
    # Row 4 (index 4): values for Project Leads and Collaborators
    if (length(rows) >= 4) {
      row_text <- normalize(get_text(rows[[4]]))
      # Split at organization patterns
      dfo_parts <- str_split(row_text, "\\(DFO\\)\\s*")[[1]]
      if (length(dfo_parts) >= 2) {
        leads <- paste(head(dfo_parts, -1), "(DFO)", sep = "")
        results$project_leads <- clean_value(paste(leads, collapse = " "))
        results$collaborators <- clean_value(tail(dfo_parts, 1))
      }
    }
    
    # Row 6 (index 6): Location value
    if (length(rows) >= 6) {
      results$location <- clean_value(get_text(rows[[6]]))
    }
  }
  
  # Process Table 4: Geographic & Stock Info
  if (length(tables) >= 4) {
    table <- tables[[4]]
    rows <- xml_find_all(table, ".//w:tr", NS)
    
    # Row 2: Salmon species and Waterbodies values
    if (length(rows) >= 2) {
      row_text <- normalize(get_text(rows[[2]]))
      if (str_detect(row_text, "Chinook")) {
        results$salmon_species <- clean_value("Chinook")
        waterbodies <- str_remove(row_text, "^Chinook\\s*")
        if (nchar(waterbodies) > 0) results$waterbodies <- clean_value(waterbodies)
      }
    }
    
    # Row 4: Life history and Region values
    if (length(rows) >= 4) {
      row_text <- normalize(get_text(rows[[4]]))
      if (str_detect(row_text, "Okanagan") && str_detect(row_text, "population")) {
        idx <- str_locate(row_text, "Okanagan")[1, "start"]
        idx <- tail(str_locate_all(row_text, "Okanagan")[[1]][, "start"], 1)
        results$life_history_phases <- clean_value(str_sub(row_text, 1, idx - 1))
        results$region <- clean_value(str_sub(row_text, idx))
      }
    }
    
    # Row 6: Stock and Population values
    if (length(rows) >= 6) {
      row_text <- normalize(get_text(rows[[6]]))
      if (str_detect(row_text, "Canadian Okanagan")) {
        results$stock <- clean_value("Canadian Okanagan")
        results$population <- clean_value("Canadian Okanagan")
      }
    }
    
    # Row 8: Conservation Unit value
    if (length(rows) >= 8) {
      results$conservation_unit <- clean_value(get_text(rows[[8]]))
    }
  }
  
  # Process Table 6: Project Outputs (single column format)
  if (length(tables) >= 6) {
    table <- tables[[6]]
    rows <- xml_find_all(table, ".//w:tr", NS)
    
    output_fields <- list(
      publications = "scientific publications",
      datasets = "datasets generated",
      dataset_locations = "dataset locations",
      code_software = "code, programs",
      communications = "communication",
      field_work = "field work",
      samples = "samples collected",
      capital_assets = "capital assets"
    )
    
    for (i in seq_along(rows)) {
      row_text <- tolower(normalize(get_text(rows[[i]])))
      
      for (field_name in names(output_fields)) {
        pattern <- output_fields[[field_name]]
        if (str_detect(row_text, pattern) && i < length(rows)) {
          # Next row is the value
          value <- clean_value(get_text(rows[[i + 1]]))
          if (!is.na(value) && is.null(results[[field_name]])) {
            results[[field_name]] <- str_sub(value, 1, 3000)
          }
        }
      }
    }
  }
  
  # Process narrative tables (Highlights, Background, etc.)
  narrative_fields <- list(
    highlights = 7,
    background = 8,
    methods_findings = 9,
    insights = 10,
    next_steps = 11
  )
  
  for (field_name in names(narrative_fields)) {
    table_idx <- narrative_fields[[field_name]]
    if (table_idx <= length(tables)) {
      table <- tables[[table_idx]]
      text <- clean_value(get_text(table))
      if (!is.na(text) && is.null(results[[field_name]])) {
        results[[field_name]] <- str_sub(text, 1, 3000)
      }
    }
  }
  
  return(results)
}


#' Extract form data from all Word documents in a folder
#'
#' @param folder_path Path to folder containing .docx files
#' @param output_csv Path for output CSV file (default: "pssi_form_data.csv")
#' @param output_db Path for output SQLite database (default: NULL, no database)
#' @param table_name Name of table in database (default: "projects")
#' @return Data frame with extracted data
extract_all_forms <- function(folder_path, 
                              output_csv = "pssi_form_data.csv",
                              output_db = NULL,
                              table_name = "projects") {
  
  # Find all .docx files (excluding temp files starting with ~$)
  docx_files <- list.files(
    folder_path,
    pattern = "\\.docx$",
    full.names = TRUE,
    ignore.case = TRUE
  ) %>%
    .[!str_detect(basename(.), "^~\\$")]
  
  if (length(docx_files) == 0) {
    stop("No .docx files found in: ", folder_path)
  }
  
  message("Found ", length(docx_files), " Word document(s) to process\n")
  
  # Extract data from each file
  all_data <- map(docx_files, function(f) {
    message("Processing: ", basename(f))
    tryCatch({
      data <- extract_form_data(f)
      n_fields <- sum(sapply(data, function(x) !is.null(x) && !is.na(x) && nchar(x) > 0)) - 1
      message("  -> Extracted ", n_fields, " fields")
      data
    }, error = function(e) {
      message("  Error: ", e$message)
      list(source_file = basename(f), error = e$message)
    })
  })
  
  # Convert to data frame
  result_df <- bind_rows(all_data)
  
  # Define preferred column order
  col_order <- c(
    "source_file", "project_id", "project_title", "project_leads", "collaborators",
    "location", "salmon_species", "waterbodies", "life_history_phases", "region",
    "stock", "population", "conservation_unit", "publications", "datasets",
    "dataset_locations", "code_software", "communications", "field_work", "samples",
    "capital_assets", "highlights", "background", "methods_findings", "insights",
    "next_steps", "error"
  )
  
  # Reorder columns (keeping any extras at the end)
  existing_cols <- col_order[col_order %in% names(result_df)]
  extra_cols <- setdiff(names(result_df), col_order)
  result_df <- result_df %>% select(all_of(c(existing_cols, extra_cols)))
  
  # Write to CSV
  write.csv(result_df, output_csv, row.names = FALSE, na = "")
  message("\nCSV saved to: ", output_csv)
  
  # Write to SQLite database if path provided
  if (!is.null(output_db)) {
    # Create output directory if needed
    db_dir <- dirname(output_db)
    if (!dir.exists(db_dir)) {
      dir.create(db_dir, recursive = TRUE)
    }
    
    # Connect to database (creates if doesn't exist)
    con <- dbConnect(SQLite(), output_db)
    on.exit(dbDisconnect(con), add = TRUE)
    
    # Write table (overwrite if exists)
    dbWriteTable(con, table_name, result_df, overwrite = TRUE)
    
    message("Database saved to: ", output_db)
    message("  Table: ", table_name)
    message("  Rows: ", nrow(result_df))
  }
  
  return(result_df)
}


# =============================================================================
# USAGE EXAMPLES
# =============================================================================

# To process a single file:
# result <- extract_form_data(here("data", "word_docs", "your_document.docx"))

# To process all files in a folder:
# all_data <- extract_all_forms(
#   folder_path = here("data", "word_docs"),
#   output_csv = here("data", "processed", "pssi_form_data.csv"),
#   output_db = here("output", "projects.db"),
#   table_name = "projects"
# )

# =============================================================================
# RUN EXTRACTION
# =============================================================================

# Create output directories if they don't exist
if (!dir.exists(here("data", "processed"))) {
  dir.create(here("data", "processed"), recursive = TRUE)
}
if (!dir.exists(here("output"))) {
  dir.create(here("output"), recursive = TRUE)
}

# Extract all forms from word_docs folder
all_data <- extract_all_forms(
  folder_path = here("data", "word_docs"),
  output_csv = here("data", "processed", "pssi_form_data.csv"),
  output_db = here("output", "projects.db"),
  table_name = "projects"
)
