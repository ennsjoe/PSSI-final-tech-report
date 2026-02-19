# helper-functions.R

# helper: define %||% if not present
`%||%` <- function(x, y) if (is.null(x) || length(x) == 0) y else x

# --- Helper function to find CSV files ---
find_csv_file <- function(filename, possible_paths = NULL) {
  if (is.null(possible_paths)) {
    possible_paths <- c(
      here(filename),
      here("data", filename),
      here("data", "raw", filename),
      here("data", "processed", filename)
    )
  }
  for (path in possible_paths) {
    if (file.exists(path)) return(path)
  }
  return(NULL)
}

# --- Find CSV files --------------------------------
RAW_CSV       <- find_csv_file("report_project_list.csv")
PROCESSED_CSV <- find_csv_file("pssi_form_data.csv")
OUTPUT_DB     <- here::here("data", "projects_database.sqlite")
TABLE_NAME    <- "projects"

# --- Database Initialization -------------------------------------------------

init_database <- function(overwrite = FALSE) {
  message("\n=== Initializing Projects Database ===")
  
  # 1. Load Tracking Data (The 'Raw' list)
  tracking_df <- readr::read_csv(RAW_CSV, show_col_types = FALSE) %>%
    dplyr::mutate(project_id = as.character(project_id))
  
  # 2. Load Content Data (The 'Form' data)
  content_df <- readr::read_csv(PROCESSED_CSV, show_col_types = FALSE) %>%
    dplyr::mutate(project_id = as.character(project_id))
  
  # 3. Identify IDs that have full form data
  ids_in_content <- unique(content_df$project_id)
  
  # --- PROCESS SET A: Projects with Full Form Data ---
  # We use the content_df as the base so we don't get the 'description' 
  # or 'location' from the tracking sheet.
  set_a <- content_df %>%
    dplyr::left_join(
      tracking_df %>% dplyr::select(project_id, source, section, broad_header, include), 
      by = "project_id"
    )
  
  # --- PROCESS SET B: Fallback Projects (Science only) ----
  # These projects ARE in the tracking list, are Science/Include, 
  # but are NOT in the form data.
  set_b <- tracking_df %>%
    dplyr::filter(!(project_id %in% ids_in_content)) %>%
    dplyr::filter(source == "DFO Science" & include == "y") %>%
    # MAP THE COLUMNS SO THE RMD TEMPLATE SEES THEM
    # Your Rmd looks for 'background'. The tracking sheet has 'description'.
    dplyr::mutate(
      background = description,
      project_title = project_title,
      # Ensure other fields expected by the template exist as NAs if missing
      methods_findings = NA_character_,
      insights = NA_character_,
      next_steps = NA_character_
    )
  
  # 4. Combine them
  final_df <- dplyr::bind_rows(set_a, set_b)
  
  # 5. Final Housekeeping
  final_df$project_number <- final_df$project_id
  
  # 6. Write to SQLite
  db_dir <- dirname(OUTPUT_DB)
  if (!dir.exists(db_dir)) dir.create(db_dir, recursive = TRUE)
  
  con <- DBI::dbConnect(RSQLite::SQLite(), OUTPUT_DB)
  DBI::dbWriteTable(con, TABLE_NAME, final_df, overwrite = overwrite)
  DBI::dbDisconnect(con)
  
  message("\u2705 Database created. Total projects: ", nrow(final_df))
  message("   - From Form Data: ", nrow(set_a))
  message("   - From Tracking Fallback: ", nrow(set_b))
}

# --- Project Banner Generator ------------------------------------------------

make_project_banner <- function(
    project,
    icon_map,
    banner_dir      = NULL,
    font_size_ratio = 0.11,
    meta_size_ratio = 0.032,
    wrap_width      = 35,
    right_pad       = 60,
    shadow_offset   = 2
) {
  if (!requireNamespace("magick", quietly = TRUE)) return(invisible(NULL))
  
  section  <- as.character(project[["section"]] %||% "")
  icon_row <- icon_map[icon_map$theme_section == section, , drop = FALSE]
  if (nrow(icon_row) == 0) return(invisible(NULL))
  
  if (is.null(banner_dir)) banner_dir <- here::here("assets", "project-banner-backgrounds")
  banner_path <- file.path(banner_dir, as.character(icon_row$icon_file[1]))
  if (!file.exists(banner_path)) return(invisible(NULL))
  
  img    <- magick::image_read(banner_path)
  info   <- magick::image_info(img)
  fsize  <- round(info$height * font_size_ratio)
  msize  <- round(info$height * meta_size_ratio)
  
  title   <- as.character(project[["project_title"]] %||% "Untitled Project")
  wrapped <- paste(strwrap(title, width = wrap_width), collapse = "\n")
  
  # Define the meta fields to display on the banner
  meta_cols <- c("project_leads", "location", "species", "waterbodies", 
                 "life_history", "stock", "population", "cu")
  
  meta_values <- sapply(meta_cols, function(col) {
    val <- as.character(project[[col]] %||% "")
    if (is.na(val) || val == "" || val == "NA" || val == "NULL") return(NULL)
    # Clean SharePoint formatting (e.g., "Location;#123" -> "Location, 123")
    val <- gsub(";#", ", ", val)
    return(val)
  })
  
  meta_string <- paste(unlist(meta_values), collapse = " | ")
  
  # Draw Title
  img <- magick::image_annotate(img, text = wrapped, gravity = "East",
                                location = paste0("+", max(1L, right_pad - shadow_offset), "+", shadow_offset),
                                color = "gray15", size = fsize, font = "Arial", weight = 700)
  img <- magick::image_annotate(img, text = wrapped, gravity = "East",
                                location = paste0("+", right_pad, "+0"),
                                color = "white", size = fsize, font = "Arial", weight = 700)
  
  # Draw Metadata string at the bottom
  if (nchar(meta_string) > 0) {
    img <- magick::image_annotate(img, text = meta_string, gravity = "South",
                                  location = "+0+12", color = "gray15", size = msize, font = "Arial")
    img <- magick::image_annotate(img, text = meta_string, gravity = "South",
                                  location = "+0+10", color = "white", size = msize, font = "Arial")
  }
  
  tmp <- tempfile(fileext = ".png")
  magick::image_write(img, path = tmp, format = "png")
  
  width_in  <- 6.5
  height_in <- width_in * info$height / info$width
  ft <- flextable::flextable(data.frame(x = ""))
  ft <- flextable::delete_part(ft, part = "header")
  ft <- flextable::compose(ft, i = 1, j = 1, value = flextable::as_paragraph(
    flextable::as_image(tmp, width = width_in, height = height_in)))
  ft <- flextable::width(ft, j = 1, width = width_in)
  ft <- flextable::border_remove(ft)
  ft <- flextable::padding(ft, padding = 0, part = "all")
  ft <- flextable::set_table_properties(ft, layout = "fixed")
  return(ft)
}

# --- ID Formatting Helper ---
display_id <- function(x) {
  if (is.null(x) || is.na(x)) return("")
  clean_id <- gsub("^DFO_PSSI_", "", as.character(x), ignore.case = TRUE)
  clean_id <- gsub("^_", "", clean_id)
  paste("PSSI", clean_id)
}
# --- Root Relationships Repair -----------------------------------------------
# Pandoc omits _rels/.rels from its docx output. Officer requires it.
# Injects a standard root .rels if missing; no-ops if already present.

ensure_root_rels <- function(docx_path) {
  stopifnot(requireNamespace("zip", quietly = TRUE))
  
  entries <- utils::unzip(docx_path, list = TRUE)$Name
  if ("_rels/.rels" %in% entries) {
    message("  ensure_root_rels: _rels/.rels already present — skipping")
    return(invisible(docx_path))
  }
  
  message("  ensure_root_rels: _rels/.rels missing — injecting standard root relationships")
  
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  
  utils::unzip(docx_path, exdir = tmp)
  
  rels_dir <- file.path(tmp, "_rels")
  dir.create(rels_dir, showWarnings = FALSE)
  
  writeLines(
    c('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '  <Relationship Id="rId1"',
      '    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"',
      '    Target="word/document.xml"/>',
      '  <Relationship Id="rId2"',
      '    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"',
      '    Target="docProps/core.xml"/>',
      '  <Relationship Id="rId3"',
      '    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"',
      '    Target="docProps/app.xml"/>',
      '</Relationships>'),
    file.path(rels_dir, ".rels")
  )
  
  orig_wd   <- getwd()
  # all.files = TRUE ensures dotfile dirs like _rels/ are included
  all_files <- list.files(tmp, recursive = TRUE, full.names = FALSE, all.files = TRUE)
  setwd(tmp)
  zip::zip(docx_path, files = all_files, mode = "mirror")
  setwd(orig_wd)
  
  message("  ensure_root_rels: _rels/.rels injected — saved to ", docx_path)
  invisible(docx_path)
}

# --- Internal Hyperlink Fixer ------------------------------------------------
# Converts flextable hyperlink_text(url = "#bookmark") external relationships
# into Word-native w:anchor links so Ctrl+Click navigates within the document.
#
# Fixes applied:
#   1. mode = "mirror" (not "cherry-pick") preserves word/ folder structure
#   2. all.files = TRUE ensures _rels/.rels dotfile dir is included in rezip
#   3. local-name() XPath + namespace fallback fixes "converted 0" bug

fix_internal_hyperlinks <- function(docx_path) {
  stopifnot(requireNamespace("xml2", quietly = TRUE))
  stopifnot(requireNamespace("zip",  quietly = TRUE))
  
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  
  utils::unzip(docx_path, exdir = tmp)
  
  doc_xml_path  <- file.path(tmp, "word", "document.xml")
  rels_xml_path <- file.path(tmp, "word", "_rels", "document.xml.rels")
  
  doc_xml  <- xml2::read_xml(doc_xml_path)
  rels_xml <- xml2::read_xml(rels_xml_path)
  
  wns  <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  rns  <- "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  pkns <- "http://schemas.openxmlformats.org/package/2006/relationships"
  
  rels             <- xml2::xml_find_all(rels_xml, "//pkns:Relationship", c(pkns = pkns))
  anchor_ids       <- character(0)
  anchor_map_local <- list()
  
  for (rel in rels) {
    target <- xml2::xml_attr(rel, "Target")
    if (!is.na(target) && startsWith(target, "#")) {
      rid                     <- xml2::xml_attr(rel, "Id")
      anchor                  <- sub("^#", "", target)
      anchor_ids              <- c(anchor_ids, rid)
      anchor_map_local[[rid]] <- anchor
    }
  }
  
  if (length(anchor_ids) == 0) {
    message("  fix_internal_hyperlinks: no #fragment hyperlinks found — nothing to convert")
    return(invisible(docx_path))
  }
  
  message("  fix_internal_hyperlinks: converting ", length(anchor_ids),
          " internal hyperlinks to w:anchor format")
  
  # local-name() XPath avoids namespace-prefix resolution issues in xml2
  hyperlinks <- xml2::xml_find_all(
    doc_xml,
    "//*[local-name()='hyperlink' and @*[local-name()='id']]"
  )
  
  converted <- 0
  for (hl in hyperlinks) {
    rid <- xml2::xml_attr(hl, paste0("{", rns, "}id"))
    if (is.na(rid)) rid <- xml2::xml_attr(hl, "id")
    if (is.na(rid) || !(rid %in% anchor_ids)) next
    
    anchor_val <- anchor_map_local[[rid]]
    xml2::xml_set_attr(hl, paste0("{", wns, "}anchor"), anchor_val)
    xml2::xml_set_attr(hl, paste0("{", rns, "}id"),     NULL)
    converted <- converted + 1
  }
  
  for (rel in rels) {
    rid <- xml2::xml_attr(rel, "Id")
    if (!is.na(rid) && rid %in% anchor_ids) xml2::xml_remove(rel)
  }
  
  xml2::write_xml(doc_xml,  doc_xml_path)
  xml2::write_xml(rels_xml, rels_xml_path)
  
  orig_wd   <- getwd()
  all_files <- list.files(tmp, recursive = TRUE, full.names = FALSE, all.files = TRUE)
  setwd(tmp)
  zip::zip(docx_path, files = all_files, mode = "mirror")
  setwd(orig_wd)
  
  message("  fix_internal_hyperlinks: converted ", converted,
          " hyperlinks — saved to ", docx_path)
  invisible(docx_path)
}

# --- Appendix Table Row Bookmark Injector ------------------------------------
# Injects Word bookmarks at the row level in the BCSRIF appendix table.
#
# Fixes applied:
#   1. mode = "mirror" preserves folder structure on rezip
#   2. all.files = TRUE includes _rels/.rels dotfile dir
#   3. gsub() sanitises project_ids to valid XML attribute name characters

add_table_row_bookmarks <- function(docx_path, project_ids) {
  stopifnot(requireNamespace("xml2", quietly = TRUE))
  stopifnot(requireNamespace("zip",  quietly = TRUE))
  
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  
  utils::unzip(docx_path, exdir = tmp)
  
  doc_xml_path <- file.path(tmp, "word", "document.xml")
  doc_xml      <- xml2::read_xml(doc_xml_path)
  
  wns <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  ns  <- c(w = wns)
  
  rows  <- xml2::xml_find_all(doc_xml, "//w:tr", ns)
  added <- 0
  
  for (row in rows) {
    first_cell_text <- xml2::xml_text(
      xml2::xml_find_first(row, ".//w:tc[1]//w:t", ns)
    )
    
    matched_id <- project_ids[
      sapply(project_ids, function(id) grepl(id, first_cell_text, fixed = TRUE))
    ]
    if (length(matched_id) == 0) next
    
    # Sanitise to valid XML attribute name (letters, digits, _ . - only)
    bookmark_name <- paste0("proj_", gsub("[^A-Za-z0-9_.-]", "_", matched_id[1]))
    bk_id         <- as.character(added + 1000L)
    
    bk_start <- xml2::read_xml(sprintf(
      '<w:bookmarkStart xmlns:w="%s" w:id="%s" w:name="%s"/>',
      wns, bk_id, bookmark_name
    ))
    bk_end <- xml2::read_xml(sprintf(
      '<w:bookmarkEnd xmlns:w="%s" w:id="%s"/>',
      wns, bk_id
    ))
    
    first_child <- xml2::xml_child(row, 1)
    xml2::xml_add_sibling(first_child, bk_start, .where = "before")
    xml2::xml_add_sibling(first_child, bk_end,   .where = "before")
    added <- added + 1
  }
  
  xml2::write_xml(doc_xml, doc_xml_path)
  
  orig_wd   <- getwd()
  all_files <- list.files(tmp, recursive = TRUE, full.names = FALSE, all.files = TRUE)
  setwd(tmp)
  zip::zip(docx_path, files = all_files, mode = "mirror")
  setwd(orig_wd)
  
  message("  add_table_row_bookmarks: added ", added,
          " bookmarks — saved to ", docx_path)
  invisible(docx_path)
}# helper-functions.R

# helper: define %||% if not present
`%||%` <- function(x, y) if (is.null(x) || length(x) == 0) y else x

# --- Helper function to find CSV files ---
find_csv_file <- function(filename, possible_paths = NULL) {
  if (is.null(possible_paths)) {
    possible_paths <- c(
      here(filename),
      here("data", filename),
      here("data", "raw", filename),
      here("data", "processed", filename)
    )
  }
  for (path in possible_paths) {
    if (file.exists(path)) return(path)
  }
  return(NULL)
}

# --- Find CSV files --------------------------------
RAW_CSV       <- find_csv_file("report_project_list.csv")
PROCESSED_CSV <- find_csv_file("pssi_form_data.csv")
OUTPUT_DB     <- here::here("data", "projects_database.sqlite")
TABLE_NAME    <- "projects"

# --- Database Initialization -------------------------------------------------

init_database <- function(overwrite = FALSE) {
  message("\n=== Initializing Projects Database ===")
  
  # 1. Load Tracking Data (The 'Raw' list)
  tracking_df <- readr::read_csv(RAW_CSV, show_col_types = FALSE) %>%
    dplyr::mutate(project_id = as.character(project_id))
  
  # 2. Load Content Data (The 'Form' data)
  content_df <- readr::read_csv(PROCESSED_CSV, show_col_types = FALSE) %>%
    dplyr::mutate(project_id = as.character(project_id))
  
  # 3. Identify IDs that have full form data
  ids_in_content <- unique(content_df$project_id)
  
  # --- PROCESS SET A: Projects with Full Form Data ---
  # We use the content_df as the base so we don't get the 'description' 
  # or 'location' from the tracking sheet.
  set_a <- content_df %>%
    dplyr::left_join(
      tracking_df %>% dplyr::select(project_id, source, section, broad_header, include), 
      by = "project_id"
    )
  
  # --- PROCESS SET B: Fallback Projects (Science only) ----
  # These projects ARE in the tracking list, are Science/Include, 
  # but are NOT in the form data.
  set_b <- tracking_df %>%
    dplyr::filter(!(project_id %in% ids_in_content)) %>%
    dplyr::filter(source == "DFO Science" & include == "y") %>%
    # MAP THE COLUMNS SO THE RMD TEMPLATE SEES THEM
    # Your Rmd looks for 'background'. The tracking sheet has 'description'.
    dplyr::mutate(
      background = description,
      project_title = project_title,
      # Ensure other fields expected by the template exist as NAs if missing
      methods_findings = NA_character_,
      insights = NA_character_,
      next_steps = NA_character_
    )
  
  # 4. Combine them
  final_df <- dplyr::bind_rows(set_a, set_b)
  
  # 5. Final Housekeeping
  final_df$project_number <- final_df$project_id
  
  # 6. Write to SQLite
  db_dir <- dirname(OUTPUT_DB)
  if (!dir.exists(db_dir)) dir.create(db_dir, recursive = TRUE)
  
  con <- DBI::dbConnect(RSQLite::SQLite(), OUTPUT_DB)
  DBI::dbWriteTable(con, TABLE_NAME, final_df, overwrite = overwrite)
  DBI::dbDisconnect(con)
  
  message("\u2705 Database created. Total projects: ", nrow(final_df))
  message("   - From Form Data: ", nrow(set_a))
  message("   - From Tracking Fallback: ", nrow(set_b))
}

# --- Project Banner Generator ------------------------------------------------

make_project_banner <- function(
    project,
    icon_map,
    banner_dir      = NULL,
    font_size_ratio = 0.11,
    meta_size_ratio = 0.032,
    wrap_width      = 35,
    right_pad       = 60,
    shadow_offset   = 2
) {
  if (!requireNamespace("magick", quietly = TRUE)) return(invisible(NULL))
  
  section  <- as.character(project[["section"]] %||% "")
  icon_row <- icon_map[icon_map$theme_section == section, , drop = FALSE]
  if (nrow(icon_row) == 0) return(invisible(NULL))
  
  if (is.null(banner_dir)) banner_dir <- here::here("assets", "project-banner-backgrounds")
  banner_path <- file.path(banner_dir, as.character(icon_row$icon_file[1]))
  if (!file.exists(banner_path)) return(invisible(NULL))
  
  # Skip banner entirely for Set B projects (no form data)
  # Set B projects lack highlights and methods_findings; only tracking-sheet fields are present
  is_set_b <- is.null(project[["highlights"]]) || 
    (is.na(project[["highlights"]]) && is.null(project[["methods_findings"]]))
  if (is_set_b) {
    message("  make_project_banner: skipping banner for Set B project ", project[["project_id"]])
    return(invisible(NULL))
  }
  
  img    <- magick::image_read(banner_path)
  info   <- magick::image_info(img)
  fsize  <- round(info$height * font_size_ratio)
  msize  <- round(info$height * meta_size_ratio)
  
  # Only use title if it came from pssi_form_data.csv (Set A)
  # Fall back to blank rather than pulling from the tracking sheet
  title   <- trimws(as.character(project[["project_title"]] %||% ""))
  if (nchar(title) == 0 || title %in% c("NA", "NULL")) title <- ""
  wrapped <- if (nchar(title) > 0) paste(strwrap(title, width = wrap_width), collapse = "\n") else ""
  
  # Define the meta fields to display on the banner
  meta_cols <- c("project_leads", "location", "species", "waterbodies",
                 "life_history", "stock", "population", "cu")
  
  meta_values <- sapply(meta_cols, function(col) {
    val <- as.character(project[[col]] %||% "")
    if (is.na(val) || val == "" || val == "NA" || val == "NULL") return(NULL)
    # Clean SharePoint formatting (e.g., "Location;#123" -> "Location, 123")
    val <- gsub(";#", ", ", val)
    return(val)
  })
  
  meta_string <- paste(unlist(meta_values), collapse = " | ")
  
  # Draw Title (only if one exists in form data)
  if (nchar(wrapped) > 0) {
    img <- magick::image_annotate(img, text = wrapped, gravity = "East",
                                  location = paste0("+", max(1L, right_pad - shadow_offset), "+", shadow_offset),
                                  color = "gray15", size = fsize, font = "Arial", weight = 700)
    img <- magick::image_annotate(img, text = wrapped, gravity = "East",
                                  location = paste0("+", right_pad, "+0"),
                                  color = "white", size = fsize, font = "Arial", weight = 700)
  }
  
  # Draw Metadata string at the bottom
  if (nchar(meta_string) > 0) {
    img <- magick::image_annotate(img, text = meta_string, gravity = "South",
                                  location = "+0+12", color = "gray15", size = msize, font = "Arial")
    img <- magick::image_annotate(img, text = meta_string, gravity = "South",
                                  location = "+0+10", color = "white", size = msize, font = "Arial")
  }
  
  tmp <- tempfile(fileext = ".png")
  magick::image_write(img, path = tmp, format = "png")
  
  width_in  <- 6.5
  height_in <- width_in * info$height / info$width
  ft <- flextable::flextable(data.frame(x = ""))
  ft <- flextable::delete_part(ft, part = "header")
  ft <- flextable::compose(ft, i = 1, j = 1, value = flextable::as_paragraph(
    flextable::as_image(tmp, width = width_in, height = height_in)))
  ft <- flextable::width(ft, j = 1, width = width_in)
  ft <- flextable::border_remove(ft)
  ft <- flextable::padding(ft, padding = 0, part = "all")
  ft <- flextable::set_table_properties(ft, layout = "fixed")
  return(ft)
}

# --- ID Formatting Helper ---
display_id <- function(x) {
  if (is.null(x) || is.na(x)) return("")
  clean_id <- gsub("^DFO_PSSI_", "", as.character(x), ignore.case = TRUE)
  clean_id <- gsub("^_", "", clean_id)
  paste("PSSI", clean_id)
}
# --- Root Relationships Repair -----------------------------------------------
# Pandoc omits _rels/.rels from its docx output. Officer requires it.
# Injects a standard root .rels if missing; no-ops if already present.

ensure_root_rels <- function(docx_path) {
  stopifnot(requireNamespace("zip", quietly = TRUE))
  
  entries <- utils::unzip(docx_path, list = TRUE)$Name
  if ("_rels/.rels" %in% entries) {
    message("  ensure_root_rels: _rels/.rels already present — skipping")
    return(invisible(docx_path))
  }
  
  message("  ensure_root_rels: _rels/.rels missing — injecting standard root relationships")
  
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  
  utils::unzip(docx_path, exdir = tmp)
  
  rels_dir <- file.path(tmp, "_rels")
  dir.create(rels_dir, showWarnings = FALSE)
  
  writeLines(
    c('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '  <Relationship Id="rId1"',
      '    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"',
      '    Target="word/document.xml"/>',
      '  <Relationship Id="rId2"',
      '    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"',
      '    Target="docProps/core.xml"/>',
      '  <Relationship Id="rId3"',
      '    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"',
      '    Target="docProps/app.xml"/>',
      '</Relationships>'),
    file.path(rels_dir, ".rels")
  )
  
  orig_wd   <- getwd()
  on.exit(setwd(orig_wd), add = TRUE)
  # all.files = TRUE ensures dotfile dirs like _rels/ are included
  all_files <- list.files(tmp, recursive = TRUE, full.names = FALSE, all.files = TRUE)
  setwd(tmp)
  zip::zip(docx_path, files = all_files, mode = "mirror")
  setwd(orig_wd)
  
  message("  ensure_root_rels: _rels/.rels injected — saved to ", docx_path)
  invisible(docx_path)
}

# --- Internal Hyperlink Fixer ------------------------------------------------
# Converts flextable hyperlink_text(url = "#bookmark") external relationships
# into Word-native w:anchor links so Ctrl+Click navigates within the document.
#
# Fixes applied:
#   1. mode = "mirror" (not "cherry-pick") preserves word/ folder structure
#   2. all.files = TRUE ensures _rels/.rels dotfile dir is included in rezip
#   3. local-name() XPath + namespace fallback fixes "converted 0" bug

fix_internal_hyperlinks <- function(docx_path) {
  stopifnot(requireNamespace("xml2", quietly = TRUE))
  stopifnot(requireNamespace("zip",  quietly = TRUE))
  
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  
  utils::unzip(docx_path, exdir = tmp)
  
  doc_xml_path  <- file.path(tmp, "word", "document.xml")
  rels_xml_path <- file.path(tmp, "word", "_rels", "document.xml.rels")
  
  doc_xml  <- xml2::read_xml(doc_xml_path)
  rels_xml <- xml2::read_xml(rels_xml_path)
  
  wns  <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  rns  <- "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  pkns <- "http://schemas.openxmlformats.org/package/2006/relationships"
  
  rels             <- xml2::xml_find_all(rels_xml, "//pkns:Relationship", c(pkns = pkns))
  anchor_ids       <- character(0)
  anchor_map_local <- list()
  
  for (rel in rels) {
    target <- xml2::xml_attr(rel, "Target")
    if (!is.na(target) && startsWith(target, "#")) {
      rid                     <- xml2::xml_attr(rel, "Id")
      anchor                  <- sub("^#", "", target)
      anchor_ids              <- c(anchor_ids, rid)
      anchor_map_local[[rid]] <- anchor
    }
  }
  
  if (length(anchor_ids) == 0) {
    message("  fix_internal_hyperlinks: no #fragment hyperlinks found — nothing to convert")
    return(invisible(docx_path))
  }
  
  message("  fix_internal_hyperlinks: converting ", length(anchor_ids),
          " internal hyperlinks to w:anchor format")
  
  # local-name() XPath avoids namespace-prefix resolution issues in xml2
  hyperlinks <- xml2::xml_find_all(
    doc_xml,
    "//*[local-name()='hyperlink' and @*[local-name()='id']]"
  )
  
  converted <- 0
  for (hl in hyperlinks) {
    rid <- xml2::xml_attr(hl, paste0("{", rns, "}id"))
    if (is.na(rid)) rid <- xml2::xml_attr(hl, "id")
    if (is.na(rid) || !(rid %in% anchor_ids)) next
    
    anchor_val <- anchor_map_local[[rid]]
    xml2::xml_set_attr(hl, paste0("{", wns, "}anchor"), anchor_val)
    xml2::xml_set_attr(hl, paste0("{", rns, "}id"),     NULL)
    converted <- converted + 1
  }
  
  for (rel in rels) {
    rid <- xml2::xml_attr(rel, "Id")
    if (!is.na(rid) && rid %in% anchor_ids) xml2::xml_remove(rel)
  }
  
  xml2::write_xml(doc_xml,  doc_xml_path)
  xml2::write_xml(rels_xml, rels_xml_path)
  
  orig_wd   <- getwd()
  on.exit(setwd(orig_wd), add = TRUE)
  all_files <- list.files(tmp, recursive = TRUE, full.names = FALSE, all.files = TRUE)
  setwd(tmp)
  zip::zip(docx_path, files = all_files, mode = "mirror")
  setwd(orig_wd)
  
  message("  fix_internal_hyperlinks: converted ", converted,
          " hyperlinks — saved to ", docx_path)
  invisible(docx_path)
}

# --- Appendix Table Row Bookmark Injector ------------------------------------
# Injects Word bookmarks at the row level in the BCSRIF appendix table.
#
# Fixes applied:
#   1. mode = "mirror" preserves folder structure on rezip
#   2. all.files = TRUE includes _rels/.rels dotfile dir
#   3. gsub() sanitises project_ids to valid XML attribute name characters

add_table_row_bookmarks <- function(docx_path, project_ids) {
  stopifnot(requireNamespace("xml2", quietly = TRUE))
  stopifnot(requireNamespace("zip",  quietly = TRUE))
  
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  
  utils::unzip(docx_path, exdir = tmp)
  
  doc_xml_path <- file.path(tmp, "word", "document.xml")
  doc_xml      <- xml2::read_xml(doc_xml_path)
  
  wns <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  ns  <- c(w = wns)
  
  rows  <- xml2::xml_find_all(doc_xml, "//w:tr", ns)
  added <- 0
  
  for (row in rows) {
    first_cell_text <- xml2::xml_text(
      xml2::xml_find_first(row, ".//w:tc[1]//w:t", ns)
    )
    
    matched_id <- project_ids[
      sapply(project_ids, function(id) grepl(id, first_cell_text, fixed = TRUE))
    ]
    if (length(matched_id) == 0) next
    
    # Sanitise to valid XML attribute name (letters, digits, _ . - only)
    bookmark_name <- paste0("proj_", gsub("[^A-Za-z0-9_.-]", "_", matched_id[1]))
    bk_id         <- as.character(added + 1000L)
    
    bk_start <- xml2::read_xml(sprintf(
      '<w:bookmarkStart xmlns:w="%s" w:id="%s" w:name="%s"/>',
      wns, bk_id, bookmark_name
    ))
    bk_end <- xml2::read_xml(sprintf(
      '<w:bookmarkEnd xmlns:w="%s" w:id="%s"/>',
      wns, bk_id
    ))
    
    first_child <- xml2::xml_child(row, 1)
    xml2::xml_add_sibling(first_child, bk_start, .where = "before")
    xml2::xml_add_sibling(first_child, bk_end,   .where = "before")
    added <- added + 1
  }
  
  xml2::write_xml(doc_xml, doc_xml_path)
  
  orig_wd   <- getwd()
  on.exit(setwd(orig_wd), add = TRUE)
  all_files <- list.files(tmp, recursive = TRUE, full.names = FALSE, all.files = TRUE)
  setwd(tmp)
  zip::zip(docx_path, files = all_files, mode = "mirror")
  setwd(orig_wd)
  
  message("  add_table_row_bookmarks: added ", added,
          " bookmarks — saved to ", docx_path)
  invisible(docx_path)
}