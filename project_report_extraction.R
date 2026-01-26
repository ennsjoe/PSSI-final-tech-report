# ============================================================
#  PSSI DOCX EXTRACTOR — CLEAN, STABLE, WORKING VERSION
# ============================================================

required_packages <- c("here", "xml2", "dplyr", "stringr", "purrr", "DBI", "RSQLite", "tibble")
new_packages <- required_packages[!(required_packages %in% installed.packages()[,"Package"])]
if(length(new_packages)) install.packages(new_packages)

library(here)
library(xml2)
library(dplyr)
library(stringr)
library(purrr)
library(DBI)
library(RSQLite)
library(tibble)

NS <- c(w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

PLACEHOLDER_PATTERNS <- c(
  "Click or tap here to enter text",
  "Click or tap here to enter text.",
  "click or tap here to enter text"
)

# ------------------------------------------------------------
# TEXT HELPERS
# ------------------------------------------------------------

get_text <- function(node) {
  paragraphs <- xml_find_all(node, ".//w:p", NS)
  if (length(paragraphs) == 0) {
    text_nodes <- xml_find_all(node, ".//w:t", NS)
    return(paste(xml_text(text_nodes), collapse = ""))
  }
  para_texts <- character(length(paragraphs))
  for (i in seq_along(paragraphs)) {
    text_nodes <- xml_find_all(paragraphs[[i]], ".//w:t", NS)
    para_texts[i] <- paste(xml_text(text_nodes), collapse = "")
  }
  para_texts <- para_texts[nchar(trimws(para_texts)) > 0]
  paste(para_texts, collapse = "\n\n")
}

normalize <- function(text) {
  text <- gsub("\\r\\n", "\n", text)
  text <- gsub("\\n{3,}", "\n\n", text)
  text <- gsub("[ \\t]+", " ", text)
  trimws(text)
}

clean_value <- function(value) {
  if (is.null(value)) return(NA_character_)
  cleaned <- normalize(value)
  if (nchar(cleaned) == 0) return(NA_character_)
  for (pattern in PLACEHOLDER_PATTERNS) {
    if (grepl(pattern, cleaned, ignore.case = TRUE)) {
      cleaned <- gsub(pattern, "", cleaned, ignore.case = TRUE)
    }
  }
  cleaned <- normalize(cleaned)
  if (nchar(cleaned) == 0) return(NA_character_)
  cleaned
}

# ------------------------------------------------------------
# EXTRACT SINGLE DOCX
# ------------------------------------------------------------

extract_form_data <- function(docx_path) {
  
  temp_dir <- tempfile()
  dir.create(temp_dir)
  on.exit(unlink(temp_dir, recursive = TRUE))
  
  unzip(docx_path, exdir = temp_dir)
  doc_xml <- read_xml(file.path(temp_dir, "word", "document.xml"))
  
  results <- list(source_file = basename(docx_path))
  
  tables <- xml_find_all(doc_xml, "//w:tbl", NS)
  
  # -------------------------
  # TABLE 3 — GENERAL INFO
  # -------------------------
  if (length(tables) >= 3) {
    table <- tables[[3]]
    rows <- xml_find_all(table, ".//w:tr", NS)
    
    if (length(rows) >= 2) {
      row_text <- normalize(get_text(rows[[2]]))
      match <- str_match(row_text, "^(\\d{4})\\s*(.+)$")
      if (!is.na(match[1,2])) {
        results$project_id <- clean_value(match[1,2])
        results$project_title <- clean_value(match[1,3])
      }
    }
    
    if (length(rows) >= 4) {
      row_text <- normalize(get_text(rows[[4]]))
      dfo_parts <- str_split(row_text, "\\(DFO\\)\\s*")[[1]]
      if (length(dfo_parts) >= 2) {
        leads <- paste(head(dfo_parts, -1), "(DFO)", sep = "")
        results$project_leads <- clean_value(paste(leads, collapse = " "))
        results$collaborators <- clean_value(tail(dfo_parts, 1))
      }
    }
    
    if (length(rows) >= 6) {
      results$location <- clean_value(get_text(rows[[6]]))
    }
  }
  
  # -------------------------
  # TABLE 4 — GEOGRAPHIC INFO
  # -------------------------
  if (length(tables) >= 4) {
    table <- tables[[4]]
    rows <- xml_find_all(table, ".//w:tr", NS)
    
    if (length(rows) >= 2) {
      row_text <- normalize(get_text(rows[[2]]))
      if (str_detect(row_text, "Chinook")) {
        results$salmon_species <- "Chinook"
        waterbodies <- str_remove(row_text, "^Chinook\\s*")
        results$waterbodies <- clean_value(waterbodies)
      }
    }
    
    if (length(rows) >= 4) {
      row_text <- normalize(get_text(rows[[4]]))
      if (str_detect(row_text, "Okanagan") && str_detect(row_text, "population")) {
        idx <- tail(str_locate_all(row_text, "Okanagan")[[1]][, "start"], 1)
        results$life_history_phases <- clean_value(str_sub(row_text, 1, idx - 1))
        results$region <- clean_value(str_sub(row_text, idx))
      }
    }
    
    if (length(rows) >= 6) {
      row_text <- normalize(get_text(rows[[6]]))
      if (str_detect(row_text, "Canadian Okanagan")) {
        results$stock <- "Canadian Okanagan"
        results$population <- "Canadian Okanagan"
      }
    }
    
    if (length(rows) >= 8) {
      results$conservation_unit <- clean_value(get_text(rows[[8]]))
    }
  }
  
  # -------------------------
  # TABLE 6 — PROJECT OUTPUTS
  # -------------------------
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
          value <- clean_value(get_text(rows[[i + 1]]))
          if (!is.na(value) && is.null(results[[field_name]])) {
            results[[field_name]] <- str_sub(value, 1, 3000)
          }
        }
      }
    }
  }
  
  # -------------------------
  # NARRATIVE TABLES
  # -------------------------
  narrative_fields <- list(
    highlights = 7,
    background = 8,
    methods_findings = 9,
    insights = 10,
    next_steps = 11
  )
  
  for (field_name in names(narrative_fields)) {
    idx <- narrative_fields[[field_name]]
    if (idx <= length(tables)) {
      text <- clean_value(get_text(tables[[idx]]))
      if (!is.na(text)) {
        results[[field_name]] <- str_sub(text, 1, 3000)
      }
    }
  }
  
  tibble(!!!results)
}

# ------------------------------------------------------------
# PROCESS ALL DOCX IN FOLDER
# ------------------------------------------------------------

process_docx_folder <- function(folder_path, output_csv, output_db = NULL, table_name = "projects") {
  
  docx_files <- list.files(folder_path, pattern = "\\.docx$", full.names = TRUE)
  docx_files <- docx_files[!str_detect(basename(docx_files), "^~\\$")]
  
  if (length(docx_files) == 0) stop("No .docx files found.")
  
  all_data <- map_dfr(docx_files, extract_form_data)
  
  # Ensure proper column order
  col_order <- c(
    "source_file", "project_id", "project_title", "project_leads", "collaborators",
    "location", "salmon_species", "waterbodies", "life_history_phases", "region",
    "stock", "population", "conservation_unit", "publications", "datasets",
    "dataset_locations", "code_software", "communications", "field_work", "samples",
    "capital_assets", "highlights", "background", "methods_findings", "insights",
    "next_steps"
  )
  
  all_data <- all_data %>% select(all_of(col_order))
  
  write.csv(all_data, output_csv, row.names = FALSE, na = "")
  
  if (!is.null(output_db)) {
    con <- dbConnect(SQLite(), output_db)
    dbWriteTable(con, table_name, all_data, overwrite = TRUE)
    dbDisconnect(con)
  }
  
  all_data
}

# ------------------------------------------------------------
# END OF SCRIPT — NOTHING RUNS AUTOMATICALLY
# ------------------------------------------------------------

message("PSSI extractor loaded. Call process_docx_folder() to run.")
