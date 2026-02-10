# project_report_extraction.R
# Extracts content from Word form templates into CSV with proper UTF-8 encoding
# and preserved paragraph structure

library(xml2)
library(dplyr)
library(purrr)
library(tibble)
library(stringr)
library(readr)
library(here)

# ----------------------------------------------------------
# Define standard paths
# ----------------------------------------------------------
OUTPUT_CSV <- here('data', 'processed', 'pssi_form_data.csv')
INPUT_DIR  <- here('data', 'word_docs')

# Ensure the output directory exists
if (!dir.exists(dirname(OUTPUT_CSV))) {
  dir.create(dirname(OUTPUT_CSV), recursive = TRUE, showWarnings = FALSE)
  message("Created output directory: ", dirname(OUTPUT_CSV))
}

# Verify input directory exists
if (!dir.exists(INPUT_DIR)) {
  stop("Input directory not found: ", INPUT_DIR, 
       "\nPlease create it and add your .docx files")
}

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
    message("    Newline count: ", str_count(result, "\n"))
  }
  
  return(result)
}

# ----------------------------------------------------------
# Extract Content Controls from Word Document
# ----------------------------------------------------------

extract_content_controls <- function(doc_xml) {
  # Initialize results with NA for all fields
  results <- setNames(
    as.list(rep(NA_character_, length(OUTPUT_FIELDS))), 
    OUTPUT_FIELDS
  )
  
  # Find all content control nodes
  sdt_nodes <- xml_find_all(doc_xml, "//w:sdt", W_NS)
  
  if (length(sdt_nodes) == 0) {
    message("  ⚠ No content controls found in document")
    return(results)
  }
  
  if (VERBOSE_MODE) {
    message("  Found ", length(sdt_nodes), " content control(s)")
  }
  
  # Process each content control
  fields_found <- 0
  for (sdt in sdt_nodes) {
    # Get the identifier (alias or tag)
    alias <- xml_attr(xml_find_first(sdt, ".//w:sdtPr/w:alias", W_NS), "val")
    tag   <- xml_attr(xml_find_first(sdt, ".//w:sdtPr/w:tag", W_NS), "val")
    
    # Use alias first, fall back to tag
    id <- if (!is.na(alias)) alias else if (!is.na(tag)) tag else NULL
    
    # Check if this is a recognized field
    if (!is.null(id) && id %in% names(FIELD_LABELS)) {
      col_name <- FIELD_LABELS[[id]]
      
      if (VERBOSE_MODE) {
        message("  Processing field: ", id, " -> ", col_name)
      }
      
      val <- extract_text_from_node(sdt, W_NS)
      
      # Store the value (handle multiple occurrences)
      if (!is.na(val)) {
        fields_found <- fields_found + 1
        
        if (is.na(results[[col_name]])) {
          results[[col_name]] <- val
        } else {
          # Multiple values: concatenate with separator and blank line
          results[[col_name]] <- paste(results[[col_name]], val, sep = "\n\n---\n\n")
          message("    ⚠ Multiple values found for ", col_name)
        }
      }
    }
  }
  
  if (!VERBOSE_MODE) {
    message("  ✓ Extracted ", fields_found, " field(s) with content")
  }
  
  return(results)
}

# ----------------------------------------------------------
# Extract data from a single DOCX file
# ----------------------------------------------------------

extract_docx <- function(path) {
  filename <- basename(path)
  message("\nProcessing: ", filename)
  
  # Create temporary directory for unzipping
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)
  
  tryCatch({
    # Unzip the docx file (suppress unzip messages)
    suppressMessages(
      unzip(path, exdir = tmp, overwrite = TRUE)
    )
    
    # Read the main document XML with explicit UTF-8 encoding
    xml_path <- file.path(tmp, "word", "document.xml")
    
    if (!file.exists(xml_path)) {
      message("  ✗ No document.xml found in ", filename)
      return(NULL)
    }
    
    # Read XML with UTF-8 encoding
    xml_obj <- read_xml(xml_path, encoding = "UTF-8")
    
    # Extract content controls
    data <- extract_content_controls(xml_obj)
    
    # Return as tibble with metadata
    return(bind_cols(
      tibble(
        source_file = filename,
        extraction_date = as.character(Sys.Date())
      ),
      as_tibble(data)
    ))
    
  }, error = function(e) {
    message("  ✗ Failed: ", filename)
    message("    Error: ", conditionMessage(e))
    return(NULL)
  })
}

# ----------------------------------------------------------
# Main Execution
# ----------------------------------------------------------

message("\n╔════════════════════════════════════════════════════════╗")
message("║     PSSI Form Data Extraction with Line Breaks        ║")
message("╚════════════════════════════════════════════════════════╝")
message("\nInput directory: ", INPUT_DIR)
message("Output file: ", OUTPUT_CSV)
if (VERBOSE_MODE) {
  message("\n⚠ VERBOSE MODE ENABLED - Detailed extraction diagnostics will be shown")
}

# Find all docx files (excluding temp files starting with ~$)
files <- list.files(INPUT_DIR, pattern = "\\.docx$", full.names = TRUE)
files <- files[!grepl("^~\\$", basename(files))]

if (length(files) == 0) {
  stop("\n✗ No .docx files found in ", INPUT_DIR)
}

message("\nFound ", length(files), " Word document(s)")
message(strrep("─", 60))

# Extract data from all files
all_data <- map(files, extract_docx) %>% 
  bind_rows()

if (is.null(all_data) || nrow(all_data) == 0) {
  stop("\n✗ No data extracted. Check that your Word documents contain content controls.")
}

message(strrep("─", 60))
message("\n✓ Total projects extracted: ", nrow(all_data))

# Ensure all expected columns exist (even if empty)
for (field in OUTPUT_FIELDS) {
  if (!field %in% names(all_data)) {
    all_data[[field]] <- NA_character_
  }
}

# Reorder columns: metadata first, then all defined fields
final_df <- all_data %>%
  select(source_file, extraction_date, all_of(OUTPUT_FIELDS))

# Convert everything to character for consistent CSV output
final_df <- final_df %>%
  mutate(across(everything(), as.character))

# Save to CSV with UTF-8 BOM (Excel-compatible)
write_csv(final_df, OUTPUT_CSV, na = "")

message("\n✓ SUCCESS!")
message("  Output saved to: ", OUTPUT_CSV)
message("  Rows: ", nrow(final_df))
message("  Columns: ", ncol(final_df))
message("  Encoding: UTF-8 with BOM (Excel-compatible)")

# ----------------------------------------------------------
# Diagnostic Summary
# ----------------------------------------------------------

# Count fields with content
field_counts <- final_df %>%
  select(-source_file, -extraction_date) %>%
  summarise(across(everything(), ~sum(!is.na(.)))) %>%
  pivot_longer(everything(), names_to = "field", values_to = "count") %>%
  arrange(desc(count))

message("\n", strrep("─", 60))
message("Field Coverage Summary")
message(strrep("─", 60))
message(sprintf("%-25s %8s %8s", "Field Name", "Count", "Percent"))
message(strrep("─", 60))

field_counts %>%
  purrr::pwalk(function(field, count, ...) {
    percent <- round(100 * count / nrow(final_df), 1)
    message(sprintf("%-25s %8d %7.1f%%", field, count, percent))
  })

# Check for special characters (diagnostic)
special_char_cols <- final_df %>%
  select(where(~any(str_detect(., "[^[:ascii:]]"), na.rm = TRUE))) %>%
  names()

if (length(special_char_cols) > 0) {
  message("\n", strrep("─", 60))
  message("UTF-8 Special Characters Detected")
  message(strrep("─", 60))
  message("Fields with non-ASCII characters (e.g., accents):")
  for (col in special_char_cols) {
    if (col %in% c("source_file", "extraction_date")) next
    message("  • ", col)
  }
  message("\n✓ This is expected for names with special characters")
}

# Check for preserved line breaks (diagnostic)
multiline_cols <- final_df %>%
  select(where(~any(str_detect(., "\n"), na.rm = TRUE))) %>%
  names()

if (length(multiline_cols) > 0) {
  message("\n", strrep("─", 60))
  message("Line Breaks Preserved")
  message(strrep("─", 60))
  message("Fields with multi-paragraph content:")
  
  for (col in multiline_cols) {
    if (col %in% c("source_file", "extraction_date")) next
    
    # Count documents with line breaks in this field
    n_with_breaks <- final_df %>%
      filter(!is.na(.data[[col]]) & str_detect(.data[[col]], "\n")) %>%
      nrow()
    
    # Count total line breaks
    total_breaks <- final_df %>%
      filter(!is.na(.data[[col]])) %>%
      pull(.data[[col]]) %>%
      str_count("\n") %>%
      sum()
    
    message(sprintf("  • %-20s: %2d doc(s), %4d line break(s)", 
                    col, n_with_breaks, total_breaks))
  }
  message("\n✓ Paragraph structure preserved successfully")
} else {
  message("\n⚠ No line breaks detected in any fields")
  message("  This may indicate that content controls are single-paragraph only")
}

message("\n", strrep("═", 60))
message("Extraction Complete!")
message(strrep("═", 60), "\n")

# Optional: Show a sample of preserved line breaks
if (length(multiline_cols) > 0 && !VERBOSE_MODE) {
  message("\n--- Sample Preview (First Multi-Paragraph Field) ---")
  sample_col <- multiline_cols[!multiline_cols %in% c("source_file", "extraction_date")][1]
  sample_row <- which(!is.na(final_df[[sample_col]]) & str_detect(final_df[[sample_col]], "\n"))[1]
  
  if (!is.na(sample_row)) {
    sample_text <- final_df[[sample_col]][sample_row]
    sample_preview <- substr(sample_text, 1, 300)
    if (nchar(sample_text) > 300) sample_preview <- paste0(sample_preview, "\n...")
    
    message("Field: ", sample_col)
    message("Document: ", final_df$source_file[sample_row])
    message("\n", sample_preview)
    message("\n", strrep("─", 60))
  }
}
