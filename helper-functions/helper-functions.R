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

# --- Find CSV files ---
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
  
  # --- PROCESS SET B: Fallback Projects (Science only) ---
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