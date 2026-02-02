library(xml2)
library(dplyr)
library(purrr)
library(tibble)
library(stringr)
library(DBI)
library(RSQLite)
library(here)

# --- Configuration ---
INPUT_DIR  <- here("data", "word_docs")
OUTPUT_CSV <- here("data", "processed", "pssi_form_data.csv")
OUTPUT_DB  <- here("output", "projects.db")

# Fields to extract - labels as they appear in the document
# Format: "label_in_document" = "output_field_name"
FIELD_LABELS <- c(
  "Project ID (number)" = "project_id",
  "Project ID" = "project_id",
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

# Output field names (unique values from FIELD_LABELS)
OUTPUT_FIELDS <- unique(FIELD_LABELS)

# Word XML namespace
W_NS <- c(w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

# ----------------------------------------------------------
# Helper: Fix encoding issues (UTF-8 mojibake)
# ----------------------------------------------------------
fix_encoding <- function(text) {
  if (is.na(text)) return(NA_character_)
  
  # Common UTF-8 mojibake patterns when UTF-8 is read as Latin-1
  replacements <- c(
    "Ã©" = "é",
    "Ã¨" = "è",
    "Ã " = "à",
    "Ã¢" = "â",
    "Ã®" = "î",
    "Ã´" = "ô",
    "Ã»" = "û",
    "Ã§" = "ç",
    "Ã‰" = "É",
    "Ã€" = "À",
    "Ã'" = "Ñ",
    "Ã±" = "ñ",
    "Ã¼" = "ü",
    "Ã¶" = "ö",
    "Ã¤" = "ä",
    "â€™" = "'",
    "â€œ" = "'",
    "â€" = "'",
    "â€" = "—",
    "â€" = "–",
    "â€¦" = "…"
  )
  
  result <- text
  for (pattern in names(replacements)) {
    result <- str_replace_all(result, fixed(pattern), replacements[pattern])
  }
  
  return(result)
}

# ----------------------------------------------------------
# Helper: Extract text from XML node preserving line breaks
# ----------------------------------------------------------
extract_text_with_breaks <- function(node, ns = W_NS) {
  if (is.na(node) || is.null(node)) return(NA_character_)
  
  
  # Find all paragraph elements
  paragraphs <- xml_find_all(node, ".//w:p", ns = ns)
  
  if (length(paragraphs) == 0) {
    # No paragraphs, try to get any text
    text <- xml_text(node)
    text <- fix_encoding(text)
    return(if (nchar(trimws(text)) == 0) NA_character_ else trimws(text))
  }
  
  para_texts <- map_chr(paragraphs, function(p) {
    # Get all text runs and breaks within the paragraph
    runs <- xml_find_all(p, ".//w:r", ns = ns)
    
    if (length(runs) == 0) return("")
    
    run_texts <- map_chr(runs, function(r) {
      # Get all children (text and breaks)
      children <- xml_children(r)
      
      parts <- map_chr(children, function(child) {
        child_name <- xml_name(child)
        if (child_name == "t") {
          return(xml_text(child))
        } else if (child_name == "br") {
          return("\n")
        } else if (child_name == "tab") {
          return("\t")
        }
        return("")
      })
      
      paste(parts, collapse = "")
    })
    
    paste(run_texts, collapse = "")
  })
  
  # Join paragraphs with double newline
  result <- paste(para_texts, collapse = "\n\n")
  
  # Clean up excessive whitespace while preserving intentional breaks
  result <- str_replace_all(result, "\n{3,}", "\n\n")
  result <- trimws(result)
  
  # Fix encoding issues
  result <- fix_encoding(result)
  
  if (nchar(result) == 0) NA_character_ else result
}

# ----------------------------------------------------------
# Extract Content Controls (w:sdt elements) from document.xml
# ----------------------------------------------------------
extract_content_controls <- function(doc_xml, ns = W_NS) {
  
  # Find all structured document tags (Content Controls)
  sdt_nodes <- xml_find_all(doc_xml, "//w:sdt", ns = ns)
  
  if (length(sdt_nodes) == 0) {
    message("  No Content Controls found")
    return(list())
  }
  
  message("  Found ", length(sdt_nodes), " Content Control(s)")
  
  results <- list()
  
  for (sdt in sdt_nodes) {
    # Get properties
    sdt_pr <- xml_find_first(sdt, ".//w:sdtPr", ns = ns)
    
    # Try to get tag or alias
    tag_node <- xml_find_first(sdt_pr, ".//w:tag", ns = ns)
    alias_node <- xml_find_first(sdt_pr, ".//w:alias", ns = ns)
    
    tag_val <- if (!is.na(tag_node)) xml_attr(tag_node, "val") else NA
    alias_val <- if (!is.na(alias_node)) xml_attr(alias_node, "val") else NA
    
    # Get content
    sdt_content <- xml_find_first(sdt, ".//w:sdtContent", ns = ns)
    content_text <- extract_text_with_breaks(sdt_content, ns)
    
    # Skip placeholder text
    if (!is.na(content_text) && 
        str_detect(content_text, "^Click or tap here to enter text\\.?$")) {
      content_text <- NA_character_
    }
    
    # Store with available identifiers
    if (!is.na(tag_val) || !is.na(alias_val)) {
      key <- if (!is.na(tag_val)) tag_val else alias_val
      results[[key]] <- content_text
    }
  }
  
  return(results)
}

# ----------------------------------------------------------
# Extract content from tables based on label matching
# Handles both adjacent cell patterns AND header row/value row patterns
# ----------------------------------------------------------
extract_from_tables <- function(doc_xml, ns = W_NS) {
  
  tables <- xml_find_all(doc_xml, "//w:tbl", ns = ns)
  
  if (length(tables) == 0) {
    message("  No tables found")
    return(list())
  }
  
  message("  Found ", length(tables), " table(s)")
  
  results <- list()
  
  for (tbl in tables) {
    rows <- xml_find_all(tbl, ".//w:tr", ns = ns)
    
    # Pattern 1: Header row followed by value row
    # Check consecutive row pairs for header-value pattern
    if (length(rows) >= 2) {
      for (r in 1:(length(rows) - 1)) {
        header_row <- rows[[r]]
        value_row <- rows[[r + 1]]
        
        header_cells <- xml_find_all(header_row, ".//w:tc", ns = ns)
        value_cells <- xml_find_all(value_row, ".//w:tc", ns = ns)
        
        # Check if this looks like a header-value pair
        # Headers typically have same number of cells as values
        if (length(header_cells) >= 1 && length(header_cells) == length(value_cells)) {
          for (i in seq_along(header_cells)) {
            header_text <- extract_text_with_breaks(header_cells[[i]], ns)
            value_text <- extract_text_with_breaks(value_cells[[i]], ns)
            
            if (!is.na(header_text)) {
              # Clean header for matching
              header_clean <- str_trim(header_text)
              header_clean <- str_remove_all(header_clean, "^\\*+|\\*+$")
              header_clean <- str_trim(header_clean)
              
              # Check if this matches a known field label
              if (header_clean %in% names(FIELD_LABELS)) {
                field_name <- FIELD_LABELS[[header_clean]]
                
                # Skip if value is also a known header (this is a header row, not value row)
                value_clean <- str_trim(value_text)
                value_clean <- str_remove_all(value_clean, "^\\*+|\\*+$")
                value_clean <- str_trim(value_clean)
                
                if (!is.na(value_text) && 
                    !str_detect(value_text, "^Click or tap here to enter text\\.?$") &&
                    !(value_clean %in% names(FIELD_LABELS))) {
                  # Only set if not already found (first match wins)
                  if (is.null(results[[field_name]])) {
                    results[[field_name]] <- value_text
                  }
                }
              }
            }
          }
        }
      }
    }
    
    # Pattern 2: Adjacent cells in same row (label | value)
    for (row in rows) {
      cells <- xml_find_all(row, ".//w:tc", ns = ns)
      
      if (length(cells) >= 2) {
        # Check pairs of adjacent cells for label-value patterns
        for (i in 1:(length(cells) - 1)) {
          label_text <- extract_text_with_breaks(cells[[i]], ns)
          value_text <- extract_text_with_breaks(cells[[i + 1]], ns)
          
          if (!is.na(label_text)) {
            # Clean label for matching
            label_clean <- str_trim(label_text)
            label_clean <- str_remove_all(label_clean, "^\\*+|\\*+$")
            label_clean <- str_trim(label_clean)
            
            # Check if this matches a known field label
            if (label_clean %in% names(FIELD_LABELS)) {
              field_name <- FIELD_LABELS[[label_clean]]
              
              # Check that value is not also a header
              value_clean <- str_trim(value_text)
              value_clean <- str_remove_all(value_clean, "^\\*+|\\*+$")
              value_clean <- str_trim(value_clean)
              
              # Skip placeholder text and skip if value is another header
              if (!is.na(value_text) && 
                  !str_detect(value_text, "^Click or tap here to enter text\\.?$") &&
                  !(value_clean %in% names(FIELD_LABELS))) {
                # Only set if not already found
                if (is.null(results[[field_name]])) {
                  results[[field_name]] <- value_text
                }
              }
            }
          }
        }
      }
      
      # Also handle single-cell rows that might be section content
      if (length(cells) == 1) {
        cell_text <- extract_text_with_breaks(cells[[1]], ns)
        if (!is.na(cell_text) && nchar(cell_text) > 100) {
          # This might be a content cell following a header
          # Store temporarily for later matching
          results[[paste0("_content_", length(results))]] <- cell_text
        }
      }
    }
  }
  
  return(results)
}

# ----------------------------------------------------------
# Extract section content based on headings
# ----------------------------------------------------------
extract_section_content <- function(doc_xml, ns = W_NS) {
  
  results <- list()
  
  # Section markers to look for
  sections <- c(
    "Highlights" = "highlights",
    "Background" = "background",
    "Methods and Findings" = "methods_findings",
    "Insights" = "insights",
    "Next Steps" = "next_steps",
    "Tables and Figures" = "tables_figures",
    "References" = "references"
  )
  
  # Get all paragraphs and tables in document order
  body <- xml_find_first(doc_xml, "//w:body", ns = ns)
  if (is.na(body)) return(results)
  
  # Find table cells that contain section content
  # These are typically single-cell tables following section headers
  tables <- xml_find_all(doc_xml, "//w:tbl", ns = ns)
  
  for (tbl in tables) {
    rows <- xml_find_all(tbl, ".//w:tr", ns = ns)
    
    for (row in rows) {
      cells <- xml_find_all(row, ".//w:tc", ns = ns)
      
      # Single cell tables often contain the main content
      if (length(cells) == 1) {
        cell_text <- extract_text_with_breaks(cells[[1]], ns)
        
        if (!is.na(cell_text) && nchar(cell_text) > 50) {
          # Try to identify which section this belongs to
          # by checking content patterns
          
          # Highlights typically start with bullet points about main idea
          if (str_detect(cell_text, regex("main idea|Key findings|implications", ignore_case = TRUE)) &&
              is.null(results[["highlights"]])) {
            results[["highlights"]] <- cell_text
          }
          # Background discusses context and knowledge gaps
          else if (str_detect(cell_text, regex("knowledge gap|context|collaboration", ignore_case = TRUE)) &&
                   is.null(results[["background"]])) {
            results[["background"]] <- cell_text
          }
          # Methods section
          else if (str_detect(cell_text, regex("method|sample|analyz|CTD|core", ignore_case = TRUE)) &&
                   nchar(cell_text) > 500 &&
                   is.null(results[["methods_findings"]])) {
            results[["methods_findings"]] <- cell_text
          }
          # Insights about management and salmon
          else if (str_detect(cell_text, regex("management|salmon|inform|ecosystem", ignore_case = TRUE)) &&
                   !str_detect(cell_text, regex("method|CTD|sample", ignore_case = TRUE)) &&
                   nchar(cell_text) > 100 && nchar(cell_text) < 2000 &&
                   is.null(results[["insights"]])) {
            results[["insights"]] <- cell_text
          }
          # Next steps
          else if (str_detect(cell_text, regex("next step|future|recommend|ongoing", ignore_case = TRUE)) &&
                   is.null(results[["next_steps"]])) {
            results[["next_steps"]] <- cell_text
          }
          # References with citations
          else if (str_detect(cell_text, regex("\\d{4}\\.|doi:|Journal|Can\\.|pp\\.", ignore_case = TRUE)) &&
                   is.null(results[["references"]])) {
            results[["references"]] <- cell_text
          }
        }
      }
    }
  }
  
  return(results)
}

# ----------------------------------------------------------
# Main extraction function for a single document
# ----------------------------------------------------------
extract_docx_content <- function(docx_path) {
  
  filename <- basename(docx_path)
  message("\nProcessing: ", filename)
  
  # Create temp directory and unzip
  temp_dir <- tempfile()
  dir.create(temp_dir)
  on.exit(unlink(temp_dir, recursive = TRUE))
  
  tryCatch({
    unzip(docx_path, exdir = temp_dir)
  }, error = function(e) {
    message("  ERROR: Could not unzip file - ", e$message)
    return(NULL)
  })
  
  # Load document.xml with explicit UTF-8 encoding
  doc_path <- file.path(temp_dir, "word", "document.xml")
  if (!file.exists(doc_path)) {
    message("  ERROR: document.xml not found")
    return(NULL)
  }
  
  doc_xml <- read_xml(doc_path, encoding = "UTF-8")
  
  # Initialize results with all fields as NA
  results <- setNames(
    as.list(rep(NA_character_, length(OUTPUT_FIELDS))),
    OUTPUT_FIELDS
  )
  
  # Method 1: Extract from Content Controls
  cc_data <- extract_content_controls(doc_xml)
  for (name in names(cc_data)) {
    if (name %in% OUTPUT_FIELDS && !is.na(cc_data[[name]])) {
      results[[name]] <- cc_data[[name]]
    }
  }
  
  # Method 2: Extract from tables (label-value pairs)
  table_data <- extract_from_tables(doc_xml)
  for (name in names(table_data)) {
    # Only use table data if Content Control did not have it
    if (name %in% OUTPUT_FIELDS && is.na(results[[name]]) && !is.na(table_data[[name]])) {
      results[[name]] <- table_data[[name]]
    }
  }
  
  # Method 3: Extract section content by pattern matching
  section_data <- extract_section_content(doc_xml)
  for (name in names(section_data)) {
    if (name %in% OUTPUT_FIELDS && is.na(results[[name]]) && !is.na(section_data[[name]])) {
      results[[name]] <- section_data[[name]]
    }
  }
  
  # Count extracted fields
  extracted_count <- sum(!is.na(unlist(results)))
  message("  Extracted ", extracted_count, " of ", length(OUTPUT_FIELDS), " fields")
  
  # Build output tibble
  df <- tibble(
    source_file = filename,
    extraction_date = as.character(Sys.Date())
  ) %>%
    bind_cols(as_tibble(results))
  
  return(df)
}

# ----------------------------------------------------------
# Process all documents in directory
# ----------------------------------------------------------
extract_all_forms <- function(input_dir = INPUT_DIR,
                              output_csv = OUTPUT_CSV,
                              output_db = OUTPUT_DB) {
  
  # Find all Word documents
  docx_files <- list.files(
    input_dir,
    pattern = "\\.docx$",
    full.names = TRUE,
    ignore.case = TRUE
  )
  
  # Exclude temp files (start with ~$)
  docx_files <- docx_files[!grepl("^~\\$", basename(docx_files))]
  
  message("Found ", length(docx_files), " Word document(s) in: ", input_dir)
  
  if (length(docx_files) == 0) {
    stop("No Word documents found in: ", input_dir)
  }
  
  # Process each document
  results <- map(docx_files, extract_docx_content)
  results <- compact(results)  # Remove NULLs
  
  if (length(results) == 0) {
    stop("No data extracted from any documents")
  }
  
  # Combine all results
  combined <- bind_rows(results)
  
  # Save to CSV with UTF-8 BOM for Excel compatibility
  dir.create(dirname(output_csv), recursive = TRUE, showWarnings = FALSE)
  
  # Write with UTF-8 BOM for better Excel compatibility
  con <- file(output_csv, "w", encoding = "UTF-8")
  writeChar("\uFEFF", con, eos = NULL)  # Write BOM
  close(con)
  write.table(combined, output_csv, row.names = FALSE, na = "", 
              sep = ",", quote = TRUE, fileEncoding = "UTF-8", append = TRUE)
  message("\nCSV saved to: ", output_csv)
  
  # Save to SQLite
  dir.create(dirname(output_db), recursive = TRUE, showWarnings = FALSE)
  con <- dbConnect(SQLite(), output_db)
  on.exit(dbDisconnect(con), add = TRUE)
  dbWriteTable(con, "projects", combined, overwrite = TRUE)
  message("Database saved to: ", output_db)
  
  # Summary
  message("\n=== EXTRACTION SUMMARY ===")
  message("Documents processed: ", nrow(combined))
  message("Fields per document: ", ncol(combined) - 2)  # Exclude source_file and extraction_date
  
  # Show field coverage
  field_coverage <- combined %>%
    select(-source_file, -extraction_date) %>%
    summarise(across(everything(), ~sum(!is.na(.) & . != ""))) %>%
    tidyr::pivot_longer(everything(), names_to = "field", values_to = "count") %>%
    mutate(percent = round(count / nrow(combined) * 100, 1))
  
  message("\nField coverage:")
  for (i in 1:nrow(field_coverage)) {
    message("  ", field_coverage$field[i], ": ", 
            field_coverage$count[i], "/", nrow(combined),
            " (", field_coverage$percent[i], "%)")
  }
  
  return(combined)
}

# ----------------------------------------------------------
# Diagnostic function to inspect a single document
# ----------------------------------------------------------
inspect_document <- function(docx_path) {
  
  message("Inspecting: ", basename(docx_path))
  
  temp_dir <- tempfile()
  dir.create(temp_dir)
  on.exit(unlink(temp_dir, recursive = TRUE))
  
  unzip(docx_path, exdir = temp_dir)
  
  doc_path <- file.path(temp_dir, "word", "document.xml")
  doc_xml <- read_xml(doc_path, encoding = "UTF-8")
  
  # Count Content Controls
  sdt_nodes <- xml_find_all(doc_xml, "//w:sdt", ns = W_NS)
  message("\nContent Controls found: ", length(sdt_nodes))
  
  if (length(sdt_nodes) > 0) {
    message("\nContent Control details:")
    for (i in seq_along(sdt_nodes)) {
      sdt_pr <- xml_find_first(sdt_nodes[[i]], ".//w:sdtPr", ns = W_NS)
      tag_node <- xml_find_first(sdt_pr, ".//w:tag", ns = W_NS)
      alias_node <- xml_find_first(sdt_pr, ".//w:alias", ns = W_NS)
      
      tag_val <- if (!is.na(tag_node)) xml_attr(tag_node, "val") else "(none)"
      alias_val <- if (!is.na(alias_node)) xml_attr(alias_node, "val") else "(none)"
      
      sdt_content <- xml_find_first(sdt_nodes[[i]], ".//w:sdtContent", ns = W_NS)
      preview <- str_trunc(xml_text(sdt_content), 50)
      
      message("  [", i, "] tag='", tag_val, "' alias='", alias_val, "' content='", preview, "'")
    }
  }
  
  # Count tables
  tables <- xml_find_all(doc_xml, "//w:tbl", ns = W_NS)
  message("\nTables found: ", length(tables))
  
  # Show table structure
  for (i in seq_along(tables)) {
    rows <- xml_find_all(tables[[i]], ".//w:tr", ns = W_NS)
    cells_per_row <- map_int(rows, ~length(xml_find_all(.x, ".//w:tc", ns = W_NS)))
    message("  Table ", i, ": ", length(rows), " rows, cells per row: ", 
            paste(head(cells_per_row, 5), collapse = ", "),
            if (length(cells_per_row) > 5) "..." else "")
  }
  
  invisible(NULL)
}

# --- Run extraction automatically when sourced ---
extract_all_forms()