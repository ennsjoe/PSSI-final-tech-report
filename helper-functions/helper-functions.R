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

# --- Project Heading Bookmark Injector --------------------------------------
# Searches the rendered docx for every project heading paragraph (identified
# by its display_id text) and injects a w:bookmarkStart / w:bookmarkEnd pair
# using the raw project_id as the bookmark name.
#
# This is the authoritative source of project bookmarks.  It runs on the
# rendered docx BEFORE fix_internal_hyperlinks so that when the hyperlink
# fix converts r:id -> w:anchor, the target bookmarks already exist.
#
# The bookmark name used here MUST match the URL fragment used in
# make_section_table(): paste0("#", project_id).

inject_project_bookmarks <- function(docx_path, projects_df) {
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
  
  # Collect all paragraph nodes once (cheaper than repeated XPath calls)
  paras <- xml2::xml_find_all(doc_xml, "//w:p", ns)
  
  # Pre-compute paragraph text (concatenate all w:t runs)
  para_texts <- vapply(paras, function(p) {
    paste(xml2::xml_text(xml2::xml_find_all(p, ".//w:t", ns)), collapse = "")
  }, character(1))
  
  # Determine highest existing bookmark id to avoid collisions
  existing_bk <- xml2::xml_find_all(doc_xml, "//*[local-name()='bookmarkStart']")
  existing_ids <- suppressWarnings(as.integer(
    xml2::xml_attr(existing_bk, paste0("{", wns, "}id"))
  ))
  existing_ids <- existing_ids[!is.na(existing_ids)]
  next_bk_id   <- if (length(existing_ids) > 0) max(existing_ids) + 1L else 2000L
  
  added    <- 0L
  not_found <- character(0)
  
  for (i in seq_len(nrow(projects_df))) {
    pid      <- as.character(projects_df$project_id[i])
    disp_txt <- display_id(pid)
    
    # Find the paragraph whose full text exactly matches the display heading
    hit <- which(trimws(para_texts) == disp_txt)
    if (length(hit) == 0) {
      not_found <- c(not_found, pid)
      next
    }
    matched_para <- paras[[hit[1]]]
    
    # Word bookmark names: letters, digits, underscores only; start with letter/underscore
    bk_name <- gsub("[^A-Za-z0-9_]", "_", pid)
    if (!grepl("^[A-Za-z_]", bk_name)) bk_name <- paste0("bk_", bk_name)
    
    bk_id_str <- as.character(next_bk_id)
    next_bk_id <- next_bk_id + 1L
    
    bk_start <- xml2::read_xml(sprintf(
      '<w:bookmarkStart xmlns:w="%s" w:id="%s" w:name="%s"/>',
      wns, bk_id_str, bk_name
    ))
    bk_end <- xml2::read_xml(sprintf(
      '<w:bookmarkEnd xmlns:w="%s" w:id="%s"/>',
      wns, bk_id_str
    ))
    
    first_child <- xml2::xml_child(matched_para, 1)
    xml2::xml_add_sibling(first_child, bk_start, .where = "before")
    xml2::xml_add_sibling(first_child, bk_end,   .where = "before")
    added <- added + 1L
  }
  
  xml2::write_xml(doc_xml, doc_xml_path)
  
  orig_wd   <- getwd()
  on.exit(setwd(orig_wd), add = TRUE)
  all_files <- list.files(tmp, recursive = TRUE, full.names = FALSE, all.files = TRUE)
  setwd(tmp)
  zip::zip(docx_path, files = all_files, mode = "mirror")
  setwd(orig_wd)
  
  message("  inject_project_bookmarks: added ", added, " bookmarks")
  if (length(not_found) > 0) {
    message("  inject_project_bookmarks: ", length(not_found),
            " project headings not found in docx:")
    for (x in not_found) message("    ", x, " (display: ", display_id(x), ")")
  }
  invisible(docx_path)
}

# --- Internal Hyperlink Fixer ------------------------------------------------
# Converts flextable hyperlink_text(url = "#bookmark") external relationships
# into Word-native w:anchor links so Ctrl+Click navigates within the document.
#
# Strategy: try a fast fixed-string text replacement first.  If that finds 0
# matches (flextable may serialise r:id with single quotes or a different prefix
# depending on its version), fall back to xml2 to locate and convert the nodes,
# then do a post-process text pass to rename any synthetic namespace prefix
# (ns0:anchor, ns1:anchor, ...) that xml2/libxml2 mints for w:anchor back to
# the correct w:anchor and strip the orphaned xmlns:nsN declaration.  This avoids
# the "error parsing attribute name" corruption seen when relying solely on xml2.

fix_internal_hyperlinks <- function(docx_path) {
  stopifnot(requireNamespace("xml2", quietly = TRUE))
  stopifnot(requireNamespace("zip",  quietly = TRUE))
  
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  
  utils::unzip(docx_path, exdir = tmp)
  
  doc_xml_path  <- file.path(tmp, "word", "document.xml")
  rels_xml_path <- file.path(tmp, "word", "_rels", "document.xml.rels")
  
  wns  <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  rns  <- "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  pkns <- "http://schemas.openxmlformats.org/package/2006/relationships"
  
  # --- 1. Build rId -> anchor map from rels ----------------------------------
  rels_xml <- xml2::read_xml(rels_xml_path)
  rels     <- xml2::xml_find_all(rels_xml, "//pkns:Relationship", c(pkns = pkns))
  
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
    message("  fix_internal_hyperlinks: no #fragment hyperlinks found -- nothing to convert")
    return(invisible(docx_path))
  }
  
  message("  fix_internal_hyperlinks: converting ", length(anchor_ids),
          " internal hyperlinks to w:anchor format")
  
  # --- 2. Read document.xml as text ------------------------------------------
  doc_text <- paste(
    readLines(doc_xml_path, encoding = "UTF-8", warn = FALSE),
    collapse = "\n"
  )
  
  # --- 3a. Fast path: fixed-string text replacement -------------------------
  # Handles both double-quoted (r:id="rIdN") and single-quoted (r:id='rIdN')
  # forms.  Each rId is unique in the document so fixed=TRUE is safe.
  converted <- 0L
  for (rid in anchor_ids) {
    anchor_val <- anchor_map_local[[rid]]
    replaced   <- FALSE
    for (q in c('"', "'")) {
      old_attr <- paste0("r:id=", q, rid, q)
      new_attr <- paste0('w:anchor="', anchor_val, '"')
      new_text <- gsub(old_attr, new_attr, doc_text, fixed = TRUE)
      if (!identical(new_text, doc_text)) {
        doc_text  <- new_text
        converted <- converted + 1L
        replaced  <- TRUE
        break
      }
    }
    # If neither quote style matched, the attribute may use a different prefix.
    # We will catch these in the xml2 fallback below.
    if (!replaced) {
      # Diagnostic: check whether the rId string is present at all
      if (!grepl(rid, doc_text, fixed = TRUE)) {
        message("    Note: ", rid, " not found anywhere in document.xml text")
      }
    }
  }
  
  # --- 3b. xml2 fallback (if text replacement missed any) --------------------
  if (converted < length(anchor_ids)) {
    remaining <- anchor_ids[!vapply(anchor_ids, function(rid) {
      any(grepl(paste0('w:anchor="', anchor_map_local[[rid]], '"'),
                doc_text, fixed = TRUE))
    }, logical(1))]
    
    if (length(remaining) > 0) {
      message("  Falling back to xml2 for ", length(remaining), " unconverted links")
      
      doc_xml <- xml2::read_xml(doc_xml_path)
      
      hlinks <- xml2::xml_find_all(
        doc_xml,
        "//*[local-name()='hyperlink' and @*[local-name()='id']]"
      )
      
      xml2_converted <- 0L
      for (hl in hlinks) {
        rid <- xml2::xml_attr(hl, paste0("{", rns, "}id"))
        if (is.na(rid)) rid <- xml2::xml_attr(hl, "id")
        if (is.na(rid) || !(rid %in% remaining)) next
        
        anchor_val <- anchor_map_local[[rid]]
        # xml_set_attr with Clark notation may create ns0:anchor -- fixed below
        xml2::xml_set_attr(hl, paste0("{", wns, "}anchor"), anchor_val)
        xml2::xml_set_attr(hl, paste0("{", rns, "}id"), NULL)
        xml2_converted <- converted + 1L
      }
      
      xml2::write_xml(doc_xml, doc_xml_path)
      
      # Re-read and fix any synthetic namespace prefix libxml2 introduced
      doc_text <- paste(
        readLines(doc_xml_path, encoding = "UTF-8", warn = FALSE),
        collapse = "\n"
      )
      # Rename ns0:anchor (or ns1:anchor, etc.) -> w:anchor
      doc_text <- gsub("\\bns[0-9]+:anchor=", "w:anchor=", doc_text, perl = TRUE)
      # Remove the orphaned synthetic namespace declaration on the same element
      wns_escaped <- gsub("\\.", "\\\\.", wns)
      doc_text <- gsub(
        paste0(' xmlns:ns[0-9]+="', wns_escaped, '"'),
        "", doc_text, perl = TRUE
      )
      
      converted <- converted + xml2_converted
    }
  }
  
  # --- 4. Write document.xml back --------------------------------------------
  con <- file(doc_xml_path, open = "w", encoding = "UTF-8")
  writeLines(doc_text, con = con, useBytes = FALSE)
  close(con)
  
  # --- 5. Remove #fragment entries from rels ---------------------------------
  for (rel in rels) {
    rid <- xml2::xml_attr(rel, "Id")
    if (!is.na(rid) && rid %in% anchor_ids) xml2::xml_remove(rel)
  }
  xml2::write_xml(rels_xml, rels_xml_path)
  
  # --- 6. Repack the docx ----------------------------------------------------
  orig_wd   <- getwd()
  on.exit(setwd(orig_wd), add = TRUE)
  all_files <- list.files(tmp, recursive = TRUE, full.names = FALSE, all.files = TRUE)
  setwd(tmp)
  zip::zip(docx_path, files = all_files, mode = "mirror")
  setwd(orig_wd)
  
  message("  fix_internal_hyperlinks: converted ", converted,
          " hyperlinks -- saved to ", docx_path)
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
}

# --- Hyperlink / Bookmark Audit ----------------------------------------------
# Compares every w:anchor link (written by fix_internal_hyperlinks) against
# every w:bookmarkStart name in document.xml.  Any "unmatched link" means the
# URL used in make_section_table() does not equal the bookmark name set by
# officer::run_bookmark() in project-template.Rmd.
#
# The link URL must be paste0("#", project_id) using the raw project_id —
# not display_id(), not a sanitised form — to match officer::run_bookmark().
#
# Usage: audit_hyperlinks(here::here("PSSI-Technical-Report-2026.docx"))
# Returns invisibly: list(bookmarks, anchors, unmatched_links)

audit_hyperlinks <- function(docx_path) {
  stopifnot(requireNamespace("xml2", quietly = TRUE))
  
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  utils::unzip(docx_path, exdir = tmp)
  
  doc_xml_path  <- file.path(tmp, "word", "document.xml")
  rels_xml_path <- file.path(tmp, "word", "_rels", "document.xml.rels")
  
  doc_xml  <- xml2::read_xml(doc_xml_path)
  rels_xml <- xml2::read_xml(rels_xml_path)
  
  wns  <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  pkns <- "http://schemas.openxmlformats.org/package/2006/relationships"
  
  # 1. All bookmarks in the document
  # Use local-name() to avoid namespace prefix resolution issues
  bk_nodes  <- xml2::xml_find_all(doc_xml,
                                  "//*[local-name()='bookmarkStart']")
  # Try Clark notation first, fall back to bare attribute name
  bookmarks <- xml2::xml_attr(bk_nodes, paste0("{", wns, "}name"))
  if (all(is.na(bookmarks)))
    bookmarks <- xml2::xml_attr(bk_nodes, "name")
  bookmarks <- bookmarks[!is.na(bookmarks) & nchar(bookmarks) > 0]
  
  # 2. w:anchor links already converted by fix_internal_hyperlinks
  hl_nodes <- xml2::xml_find_all(
    doc_xml,
    "//*[local-name()='hyperlink' and @*[local-name()='anchor']]"
  )
  # Try Clark notation first, fall back to bare attribute name
  anchors <- xml2::xml_attr(hl_nodes, paste0("{", wns, "}anchor"))
  if (all(is.na(anchors)))
    anchors <- xml2::xml_attr(hl_nodes, "anchor")
  anchors  <- anchors[!is.na(anchors) & nchar(anchors) > 0]
  
  # 3. Any remaining unconverted #fragment external links
  rels         <- xml2::xml_find_all(rels_xml, "//pkns:Relationship", c(pkns = pkns))
  frag_targets <- character(0)
  for (rel in rels) {
    tgt <- xml2::xml_attr(rel, "Target")
    if (!is.na(tgt) && startsWith(tgt, "#"))
      frag_targets <- c(frag_targets, sub("^#", "", tgt))
  }
  
  # Report
  message("\n--- audit_hyperlinks: ", basename(docx_path), " ---")
  message("  Bookmarks found       : ", length(bookmarks))
  message("  w:anchor links found  : ", length(anchors))
  if (length(frag_targets) > 0)
    message("  !! Unconverted #fragment links still in rels: ",
            length(frag_targets),
            "\n     fix_internal_hyperlinks() may not have run, or found 0 links.")
  
  unmatched_links <- setdiff(anchors, bookmarks)
  unmatched_bks   <- setdiff(bookmarks, anchors)
  
  if (length(anchors) == 0 && length(frag_targets) == 0) {
    message("  No internal links found — has make_section_table() added any?")
  } else if (length(unmatched_links) == 0 && length(anchors) > 0) {
    message("  \u2705 All ", length(anchors),
            " anchor links have a matching bookmark.")
  } else if (length(unmatched_links) > 0) {
    message("  \u274c Links with NO matching bookmark (", length(unmatched_links), "):")
    for (x in unmatched_links) message("       ", x)
    message("")
    message("  Fix: in make_section_table(), the link URL must be")
    message("       paste0('#', project_id)")
    message("       using the same raw project_id value that officer::run_bookmark()")
    message("       receives in project-template.Rmd — not display_id() output,")
    message("       not a sanitised/prefixed form.")
  }
  
  if (length(unmatched_bks) > 0) {
    n_show <- min(10L, length(unmatched_bks))
    message("  Bookmarks with no link (", length(unmatched_bks),
            ") — informational only:")
    for (x in head(unmatched_bks, n_show)) message("       ", x)
    if (length(unmatched_bks) > 10)
      message("       ... and ", length(unmatched_bks) - 10, " more")
  }
  
  message("--- end audit ---\n")
  
  invisible(list(bookmarks      = bookmarks,
                 anchors         = anchors,
                 unmatched_links = unmatched_links))
}