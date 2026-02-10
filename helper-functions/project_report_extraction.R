library(xml2)
library(dplyr)
library(purrr)
library(tibble)
library(stringr)
library(here)

# ----------------------------------------------------------
# Define standard paths
# ----------------------------------------------------------
OUTPUT_CSV <- here('data', 'processed', 'pssi_form_data.csv')
INPUT_DIR  <- here('data', 'word_docs')

# Ensure the output directory exists
if (!dir.exists(dirname(OUTPUT_CSV))) {
  dir.create(dirname(OUTPUT_CSV), recursive = TRUE, showWarnings = FALSE)
}

message("Using paths:")
message("  CSV Output: ", OUTPUT_CSV)
message("  Input:      ", INPUT_DIR)
message("")

# ----------------------------------------------------------
# Field Definitions
# ----------------------------------------------------------
FIELD_LABELS <- c(
  "Project ID (number)" = "project_id",
  "project_id" = "project_id",
  "Project Title" = "project_title",
  "project_title" = "project_title",
  "Project Leads" = "project_leads",
  "project_leads" = "project_leads",
  "Collaborations and External Partners" = "collaborations",
  "collaborations" = "collaborations",
  "location" = "location",
  "species" = "species",
  "waterbodies" = "waterbodies",
  "life_history" = "life_history",
  "region" = "region",
  "stock" = "stock",
  "population" = "population",
  "cu" = "cu",
  "highlights" = "highlights",
  "background" = "background",
  "methods_findings" = "methods_findings",
  "insights" = "insights",
  "next_steps" = "next_steps",
  "tables_figures" = "tables_figures",
  "references" = "references"
)

OUTPUT_FIELDS <- unique(FIELD_LABELS)
W_NS <- c(w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
PLACEHOLDER_PATTERNS <- c("^Click or tap here", "^Enter text here", "^Type here")

# Set this to TRUE to see detailed extraction diagnostics
VERBOSE_MODE <- FALSE

# ----------------------------------------------------------
# Text Cleaning Function with Line Break Preservation
# ----------------------------------------------------------

clean_text <- function(text) {
  if (is.null(text) || is.na(text) || length(text) == 0) {
    return(NA_character_)
  }
  
  # Ensure proper UTF-8 encoding
  text <- enc2utf8(text)
  
  # Standardize smart quotes and special characters to prevent CSV breakage
  text <- gsub("[\u201C\u201D]", '"', text)  # Smart double quotes
  text <- gsub("[\u2018\u2019]", "'", text)  # Smart single quotes
  text <- gsub("\u2013", "-", text)          # En dash
  text <- gsub("\u2014", "--", text)         # Em dash
  text <- gsub("\u2026", "...", text)        # Ellipsis
  
  # Normalize different line ending types to \n
  text <- gsub("\r\n", "\n", text)  # Windows line endings
  text <- gsub("\r", "\n", text)    # Old Mac line endings
  
  # Remove control characters EXCEPT newlines (0x0A) and tabs (0x09)
  text <- gsub("[\x01-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text, perl = TRUE)
  
  # Collapse multiple spaces/tabs on the same line (but not newlines)
  text <- gsub("[ \t]+", " ", text)
  
  # Normalize excessive blank lines (4+ consecutive newlines) to 3
  # This preserves intentional paragraph spacing while preventing huge gaps
  text <- gsub("\n{4,}", "\n\n\n", text)
  
  # Trim spaces from each line (but preserve blank lines)
  lines <- strsplit(text, "\n", fixed = TRUE)[[1]]
  lines <- trimws(lines)
  text <- paste(lines, collapse = "\n")
  
  # Remove leading/trailing blank lines
  text <- gsub("^\n+", "", text)
  text <- gsub("\n+$", "", text)
  
  # Return NA for empty strings
  if (nchar(text) == 0) {
    return(NA_character_)
  }
  
  # Check for placeholder text
  for (p in PLACEHOLDER_PATTERNS) {
    if (grepl(p, text, ignore.case = TRUE)) {
      return(NA_character_)
    }
  }
  
  return(text)
}

# ----------------------------------------------------------
# Extract text from XML node preserving paragraph structure
# ----------------------------------------------------------

extract_text_from_node <- function(node, ns) {
  if (is.null(node) || length(node) == 0) {
    return(NA_character_)
  }
  
  # Find all paragraph nodes within this content control
  p_nodes <- xml_find_all(node, ".//w:p", ns)
  
  if (length(p_nodes) == 0) {
    return(NA_character_)
  }
  
  if (VERBOSE_MODE) {
    message("    Found ", length(p_nodes), " paragraph(s)")
  }
  
  # Extract text from each paragraph
  paragraphs <- character()
  
  for (i in seq_along(p_nodes)) {
    p_node <- p_nodes[[i]]
    
    # Find all text runs within this paragraph
    t_nodes <- xml_find_all(p_node, ".//w:t", ns)
    
    if (length(t_nodes) > 0) {
      # Extract and combine text from all runs in this paragraph
      para_texts <- xml_text(t_nodes)
      para_texts <- sapply(para_texts, enc2utf8, USE.NAMES = FALSE)
      para_text <- paste(para_texts, collapse = "")
      
      # Trim spaces from this paragraph
      para_text <- trimws(para_text)
      
      if (nchar(para_text) > 0) {
        # Non-empty paragraph
        paragraphs <- c(paragraphs, para_text)
        if (VERBOSE_MODE) {
          preview <- substr(para_text, 1, 60)
          if (nchar(para_text) > 60) preview <- paste0(preview, "...")
          message("      Para ", i, ": ", preview)
        }
      } else {
        # Empty paragraph = blank line (preserve structure)
        paragraphs <- c(paragraphs, "")
        if (VERBOSE_MODE) {
          message("      Para ", i, ": [blank line]")
        }
      }
    } else {
      # Paragraph with no text nodes = blank line
      paragraphs <- c(paragraphs, "")
      if (VERBOSE_MODE) {
        message("      Para ", i, ": [blank line - no text nodes]")
      }
    }
  }
  
  if (length(paragraphs) == 0) {
    return(NA_character_)
  }
  
  # Join paragraphs with single newlines
  # Blank lines (empty strings) will create the spacing we want
  combined <- paste(paragraphs, collapse = "\n")
  
  if (VERBOSE_MODE) {
    message("    Combined length: ", nchar(combined), " characters")
    message("    Newline count: ", str_count(combined, "\n"))
  }
  
  # Clean and return (this will normalize excessive blank lines)
  result <- clean_text(combined)
  
  if (VERBOSE_MODE && !is.na(result)) {
    message("    After cleaning: ", nchar(result), " characters")
    message("    Final newline count: ", str_count(result, "\n"))
  }
  
  return(result)
}

# ----------------------------------------------------------
# Extract Content Controls (Primary Method)
# ----------------------------------------------------------

extract_content_controls <- function(doc_xml) {
  results <- setNames(as.list(rep(NA_character_, length(OUTPUT_FIELDS))), OUTPUT_FIELDS)
  
  sdt_nodes <- xml_find_all(doc_xml, "//w:sdt", W_NS)
  
  if (VERBOSE_MODE) {
    message("  Found ", length(sdt_nodes), " content control(s)")
  }
  
  for (sdt in sdt_nodes) {
    # Try to get identifier from alias or tag
    alias <- xml_attr(xml_find_first(sdt, ".//w:sdtPr/w:alias", W_NS), "val")
    tag   <- xml_attr(xml_find_first(sdt, ".//w:sdtPr/w:tag", W_NS), "val")
    
    id <- if (!is.na(alias)) alias else if (!is.na(tag)) tag else NULL
    
    if (!is.null(id) && id %in% names(FIELD_LABELS)) {
      col_name <- FIELD_LABELS[[id]]
      
      if (VERBOSE_MODE) {
        message("  Processing content control: ", id, " -> ", col_name)
      }
      
      val <- extract_text_from_node(sdt, W_NS)
      
      if (!is.na(val)) {
        # If this field already has data, append with separator
        if (is.na(results[[col_name]])) {
          results[[col_name]] <- val
        } else {
          results[[col_name]] <- paste(results[[col_name]], val, sep = "; ")
        }
        
        if (VERBOSE_MODE) {
          message("    Extracted: ", substr(val, 1, 100), "...")
        }
      }
    }
  }
  
  return(results)
}

# ----------------------------------------------------------
# Extract from Tables (Fallback Method)
# ----------------------------------------------------------

extract_from_tables <- function(doc_xml) {
  results <- setNames(as.list(rep(NA_character_, length(OUTPUT_FIELDS))), OUTPUT_FIELDS)
  
  tables <- xml_find_all(doc_xml, "//w:tbl", W_NS)
  
  if (VERBOSE_MODE && length(tables) > 0) {
    message("  Found ", length(tables), " table(s)")
  }
  
  for (tbl in tables) {
    rows <- xml_find_all(tbl, ".//w:tr", W_NS)
    
    for (row in rows) {
      cells <- xml_find_all(row, ".//w:tc", W_NS)
      
      if (length(cells) >= 2) {
        # Extract label and value
        label_text <- extract_text_from_node(cells[[1]], W_NS)
        value_text <- extract_text_from_node(cells[[2]], W_NS)
        
        if (!is.na(label_text) && label_text %in% names(FIELD_LABELS)) {
          col_name <- FIELD_LABELS[[label_text]]
          
          if (!is.na(value_text)) {
            if (is.na(results[[col_name]])) {
              results[[col_name]] <- value_text
            } else {
              results[[col_name]] <- paste(results[[col_name]], value_text, sep = "; ")
            }
          }
        }
      }
    }
  }
  
  return(results)
}

# ----------------------------------------------------------
# Extract from Section Patterns (Last Resort)
# ----------------------------------------------------------

extract_section_content <- function(doc_xml) {
  results <- setNames(as.list(rep(NA_character_, length(OUTPUT_FIELDS))), OUTPUT_FIELDS)
  
  # Section patterns to look for
  patterns <- list(
    highlights = "^highlights?$",
    background = "^background$",
    methods_findings = "^methods?( and findings?)?$|^findings?$",
    insights = "^insights?$",
    next_steps = "^next steps?$",
    tables_figures = "^tables?( and figures?)?$|^figures?$",
    references = "^references?$|^citations?$"
  )
  
  paragraphs <- xml_find_all(doc_xml, "//w:p", W_NS)
  current_section <- NULL
  section_content <- list()
  
  for (p in paragraphs) {
    text <- extract_text_from_node(p, W_NS)
    
    if (is.na(text)) next
    
    text_clean <- trimws(tolower(text))
    
    # Check if this is a section heading
    matched_section <- NULL
    for (section_name in names(patterns)) {
      if (grepl(patterns[[section_name]], text_clean, ignore.case = TRUE)) {
        matched_section <- section_name
        break
      }
    }
    
    if (!is.null(matched_section)) {
      # Start new section
      current_section <- matched_section
      section_content[[current_section]] <- character()
    } else if (!is.null(current_section)) {
      # Add to current section
      section_content[[current_section]] <- c(section_content[[current_section]], text)
    }
  }
  
  # Combine section content
  for (section in names(section_content)) {
    if (length(section_content[[section]]) > 0) {
      results[[section]] <- clean_text(paste(section_content[[section]], collapse = "\n"))
    }
  }
  
  return(results)
}

# ----------------------------------------------------------
# Main extraction function for a single document
# ----------------------------------------------------------

extract_docx <- function(path) {
  message("Processing: ", basename(path))
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  
  tryCatch({
    # Unzip the Word document
    suppressMessages(unzip(path, exdir = tmp, overwrite = TRUE))
    
    # Read the document XML
    xml_obj <- read_xml(file.path(tmp, "word", "document.xml"))
    
    # Extract using all three methods
    results <- setNames(as.list(rep(NA_character_, length(OUTPUT_FIELDS))), OUTPUT_FIELDS)
    
    # Try content controls first (most reliable)
    cc_results <- extract_content_controls(xml_obj)
    for (field in OUTPUT_FIELDS) {
      if (!is.na(cc_results[[field]])) {
        results[[field]] <- cc_results[[field]]
      }
    }
    
    # Try tables for any missing fields
    table_results <- extract_from_tables(xml_obj)
    for (field in OUTPUT_FIELDS) {
      if (is.na(results[[field]]) && !is.na(table_results[[field]])) {
        results[[field]] <- table_results[[field]]
      }
    }
    
    # Try section patterns for any still missing
    section_results <- extract_section_content(xml_obj)
    for (field in OUTPUT_FIELDS) {
      if (is.na(results[[field]]) && !is.na(section_results[[field]])) {
        results[[field]] <- section_results[[field]]
      }
    }
    
    # Return as tibble
    return(bind_cols(
      tibble(
        source_file = basename(path),
        extraction_date = as.character(Sys.Date())
      ),
      as_tibble(results)
    ))
    
  }, error = function(e) {
    message("  ! Failed: ", path, " - ", conditionMessage(e))
    return(NULL)
  })
}

# ----------------------------------------------------------
# Main Run
# ----------------------------------------------------------

files <- list.files(INPUT_DIR, pattern = "\\.docx$", full.names = TRUE)
files <- files[!grepl("^~\\$", basename(files))]

if (length(files) == 0) {
  stop("No .docx files found in: ", INPUT_DIR)
}

message("Found ", length(files), " Word document(s)\n")

all_data <- map(files, extract_docx) %>% 
  compact() %>%
  bind_rows()

if (nrow(all_data) == 0) {
  stop("No data extracted from any documents")
}

# Final cleanup and column ordering
final_df <- all_data %>%
  mutate(across(everything(), as.character)) 

# Ensure all expected fields exist
for (f in OUTPUT_FIELDS) {
  if (!f %in% names(final_df)) {
    final_df[[f]] <- NA_character_
  }
}

final_df <- final_df %>% 
  select(source_file, extraction_date, all_of(OUTPUT_FIELDS))

# Save to CSV with UTF-8 BOM for Excel compatibility
write_csv(final_df, OUTPUT_CSV, na = "")

message("\n✓ Success: Extracted ", nrow(final_df), " project(s)")
message("✓ CSV saved to: ", OUTPUT_CSV)
message("\nNext steps:")
message("  1. Rebuild database: source(here('helper-functions/helper-functions.R')); init_database(overwrite=TRUE)")
message("  2. Verify: source(here('verify_after_update.R'))")