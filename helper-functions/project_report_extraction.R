library(xml2)
library(dplyr)
library(purrr)
library(tibble)
library(stringr)
library(here)

# ----------------------------------------------------------
# Define standard paths - UPDATED TO data/processed
# ----------------------------------------------------------
OUTPUT_CSV <- here('data', 'processed', 'pssi_form_data.csv')
INPUT_DIR  <- here('data', 'word_docs')

# Ensure the output directory exists
if (!dir.exists(dirname(OUTPUT_CSV))) {
  dir.create(dirname(OUTPUT_CSV), recursive = TRUE, showWarnings = FALSE)
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

# ----------------------------------------------------------
# Cleaning Function (Keeps 3682 formatting intact)
# ----------------------------------------------------------

clean_text <- function(text) {
  if (is.null(text) || is.na(text) || length(text) == 0) return(NA_character_)
  
  # Standardize quotes to prevent CSV breakage
  text <- gsub("[\u201C\u201D]", '"', text)
  text <- gsub("[\u2018\u2019]", "'", text)
  
  # Remove newlines to keep projects on a single row
  text <- gsub("[\r\n]+", " ", text)
  
  # Strip non-printable characters
  text <- gsub("[[:cntrl:]]", "", text)
  
  text <- trimws(text)
  if (nchar(text) == 0) return(NA_character_)
  
  for (p in PLACEHOLDER_PATTERNS) {
    if (grepl(p, text, ignore.case = TRUE)) return(NA_character_)
  }
  
  return(text)
}

extract_text_from_node <- function(node, ns) {
  t_nodes <- xml_find_all(node, ".//w:t", ns)
  combined <- paste(xml_text(t_nodes), collapse = "")
  return(clean_text(combined))
}

# ----------------------------------------------------------
# Extraction Logic
# ----------------------------------------------------------

extract_content_controls <- function(doc_xml) {
  results <- setNames(as.list(rep(NA_character_, length(OUTPUT_FIELDS))), OUTPUT_FIELDS)
  sdt_nodes <- xml_find_all(doc_xml, "//w:sdt", W_NS)
  
  for (sdt in sdt_nodes) {
    alias <- xml_attr(xml_find_first(sdt, ".//w:sdtPr/w:alias", W_NS), "val")
    tag   <- xml_attr(xml_find_first(sdt, ".//w:sdtPr/w:tag", W_NS), "val")
    
    id <- if (!is.na(alias)) alias else if (!is.na(tag)) tag else NULL
    
    if (!is.null(id) && id %in% names(FIELD_LABELS)) {
      col_name <- FIELD_LABELS[[id]]
      val <- extract_text_from_node(sdt, W_NS)
      
      if (!is.na(val)) {
        if (is.na(results[[col_name]])) {
          results[[col_name]] <- val
        } else {
          results[[col_name]] <- paste(results[[col_name]], val, sep = "; ")
        }
      }
    }
  }
  return(results)
}

extract_docx <- function(path) {
  message("Processing: ", basename(path))
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  
  tryCatch({
    unzip(path, exdir = tmp)
    xml_obj <- read_xml(file.path(tmp, "word", "document.xml"))
    data <- extract_content_controls(xml_obj)
    return(bind_cols(tibble(source_file = basename(path), extraction_date = as.character(Sys.Date())), as_tibble(data)))
  }, error = function(e) {
    message("  ! Failed: ", path)
    return(NULL)
  })
}

# ----------------------------------------------------------
# Main Run
# ----------------------------------------------------------

files <- list.files(INPUT_DIR, pattern = "\\.docx$", full.names = TRUE)
files <- files[!grepl("^~\\$", basename(files))]

all_data <- map(files, extract_docx) %>% bind_rows()

# Final cleanup and column ordering
final_df <- all_data %>%
  mutate(across(everything(), as.character)) 

for (f in OUTPUT_FIELDS) {
  if (!f %in% names(final_df)) final_df[[f]] <- NA_character_
}

final_df <- final_df %>% 
  select(source_file, extraction_date, all_of(OUTPUT_FIELDS))

# Save to CSV (Excel-friendly format)
write.csv(final_df, OUTPUT_CSV, row.names = FALSE, fileEncoding = "UTF-8", na = "")

message("\n✓ Success: Processed CSV saved to ", OUTPUT_CSV)