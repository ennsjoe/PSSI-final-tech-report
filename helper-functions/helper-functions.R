# helper-functions.R


# helper: define %||% if not present
`%||%` <- function(x, y) if (is.null(x) || length(x) == 0) y else x

# --- MANUAL PATH OVERRIDE (if needed) ---
# If the automatic search doesn't find your CSV files, uncomment and edit these lines:
# RAW_CSV <- here("your", "path", "to", "report_project_list.csv")
# PROCESSED_CSV <- here("your", "path", "to", "pssi_form_data.csv")
# Then comment out or delete the "Find CSV files" section below (lines 20-75)

# --- Helper function to find CSV files ---
find_csv_file <- function(filename, possible_paths = NULL) {
  # Default search paths if none provided
  if (is.null(possible_paths)) {
    possible_paths <- c(
      here(filename),                           # Project root
      here("data", filename),                   # data/
      here("data", "raw", filename),           # data/raw/
      here("data", "processed", filename),     # data/processed/
      here("input", filename),                 # input/
      here("inputs", filename)                 # inputs/
    )
  }
  
  # Check each path
  for (path in possible_paths) {
    if (file.exists(path)) {
      return(path)
    }
  }
  
  # Not found
  return(NULL)
}

# --- Find CSV files ---
if (!exists("RAW_CSV") || !exists("PROCESSED_CSV")) {
  message("\n=== Locating CSV Files ===")
  
  RAW_CSV <- find_csv_file("report_project_list.csv")
  if (is.null(RAW_CSV)) {
    stop("\n❌ Cannot find 'report_project_list.csv'\n",
         "   Searched in:\n",
         "   - ", here("report_project_list.csv"), "\n",
         "   - ", here("data", "report_project_list.csv"), "\n",
         "   - ", here("data", "raw", "report_project_list.csv"), "\n",
         "   - ", here("data", "processed", "report_project_list.csv"), "\n",
         "   - ", here("input", "report_project_list.csv"), "\n",
         "   - ", here("inputs", "report_project_list.csv"), "\n\n",
         "   Solution: Run find-csv-files.R to diagnose, or manually set path at top of helper-functions.R")
  }
  
  PROCESSED_CSV <- find_csv_file("pssi_form_data.csv")
  if (is.null(PROCESSED_CSV)) {
    stop("\n❌ Cannot find 'pssi_form_data.csv'\n",
         "   Searched in:\n",
         "   - ", here("pssi_form_data.csv"), "\n",
         "   - ", here("data", "pssi_form_data.csv"), "\n",
         "   - ", here("data", "raw", "pssi_form_data.csv"), "\n",
         "   - ", here("data", "processed", "pssi_form_data.csv"), "\n",
         "   - ", here("input", "pssi_form_data.csv"), "\n",
         "   - ", here("inputs", "pssi_form_data.csv"), "\n\n",
         "   Solution: Run find-csv-files.R to diagnose, or manually set path at top of helper-functions.R")
  }
  
  message("✓ Found tracking CSV: ", RAW_CSV)
  message("✓ Found content CSV: ", PROCESSED_CSV)
}

# --- Database Setup Functions ---

#' Initialize and Merge the SQLite database
#' Merges tracking data (report_project_list.csv) with content (pssi_form_data.csv)
init_database <- function(overwrite = FALSE) {
  message("\n=== Initializing Projects Database ===")
  
  # Files already located by find_csv_file() above
  message("Using files:")
  message("  Tracking: ", RAW_CSV)
  message("  Content: ", PROCESSED_CSV)
  
  # 1. Load Tracking Data (metadata for filtering and organization)
  message("\nLoading tracking data...")
  tracking_df <- read_csv(RAW_CSV, show_col_types = FALSE) %>%
    mutate(project_id = as.character(project_id))
  message("  ✓ Loaded ", nrow(tracking_df), " rows")
  message("  Columns: ", paste(names(tracking_df), collapse = ", "))
  
  # 2. Load Extracted Content (detailed project information)
  message("\nLoading content data...")
  content_df <- read_csv(PROCESSED_CSV, show_col_types = FALSE) %>%
    mutate(project_id = as.character(project_id))
  message("  ✓ Loaded ", nrow(content_df), " rows")
  message("  Columns: ", paste(names(content_df), collapse = ", "))
  
  # 3. Merge the datasets on project_id
  message("\nMerging datasets on 'project_id'...")
  final_df <- tracking_df %>%
    left_join(content_df, by = "project_id", suffix = c("", ".y"))
  
  # 4. Handle duplicate columns (keep tracking version, drop .y versions)
  duplicate_cols <- grep("\\.y$", names(final_df), value = TRUE)
  if (length(duplicate_cols) > 0) {
    message("  Resolving ", length(duplicate_cols), " duplicate columns (keeping tracking data)")
    final_df <- final_df %>% select(-all_of(duplicate_cols))
  }
  
  # 5. Create project_number as backward compatibility alias
  final_df$project_number <- final_df$project_id
  
  # 6. Ensure 'include' column exists for filtering
  if (!"include" %in% names(final_df)) {
    message("  Adding 'include' column (default: 'y')")
    final_df$include <- "y"
  }
  
  message("  ✓ Final dataset: ", nrow(final_df), " rows, ", ncol(final_df), " columns")
  message("  Final columns: ", paste(names(final_df), collapse = ", "))
  
  # 7. Create database directory if needed
  db_dir <- dirname(OUTPUT_DB)
  if (!dir.exists(db_dir)) {
    message("\nCreating database directory: ", db_dir)
    dir.create(db_dir, recursive = TRUE)
  }
  
  # 8. Write to SQLite database
  message("\nWriting to database: ", OUTPUT_DB)
  tryCatch({
    con <- dbConnect(SQLite(), OUTPUT_DB)
    dbWriteTable(con, TABLE_NAME, final_df, overwrite = overwrite)
    
    # Verify
    row_count <- dbGetQuery(con, paste0("SELECT COUNT(*) as n FROM ", TABLE_NAME))$n
    columns <- dbListFields(con, TABLE_NAME)
    
    dbDisconnect(con)
    
    message("\n✅ Database created successfully!")
    message("   Location: ", OUTPUT_DB)
    message("   Projects: ", row_count)
    message("   Columns: ", length(columns))
    
    # Check for critical columns
    required <- c("project_id", "source", "project_title", "broad_header", 
                  "section", "include")
    missing <- setdiff(required, columns)
    if (length(missing) > 0) {
      warning("⚠ Missing required columns: ", paste(missing, collapse = ", "))
    } else {
      message("   ✓ All required columns present")
    }
    
    return(invisible(final_df))
    
  }, error = function(e) {
    if (exists("con")) dbDisconnect(con)
    stop("Database write error: ", e$message)
  })
}

# --- Database Inspection Functions ---

#' Check database status
check_database <- function(OUTPUT_DB = OUTPUT_DB) {
  if (!file.exists(OUTPUT_DB)) {
    return(list(
      exists = FALSE,
      message = "Database not found"
    ))
  }
  
  tryCatch({
    con <- dbConnect(SQLite(), OUTPUT_DB)
    tables <- dbListTables(con)
    
    if (TABLE_NAME %in% tables) {
      row_count <- dbGetQuery(con, paste0("SELECT COUNT(*) as n FROM ", TABLE_NAME))$n
      columns <- dbListFields(con, TABLE_NAME)
      dbDisconnect(con)
      
      return(list(
        exists = TRUE,
        message = paste0("Database contains ", row_count, " projects"),
        row_count = row_count,
        columns = columns
      ))
    } else {
      dbDisconnect(con)
      return(list(
        exists = FALSE,
        message = paste0("Database exists but '", TABLE_NAME, "' table not found")
      ))
    }
  }, error = function(e) {
    if (exists("con")) dbDisconnect(con)
    return(list(
      exists = FALSE,
      message = paste0("Error accessing database: ", e$message)
    ))
  })
}

# --- Appendix Numbering ---
if (!exists(".appendix_counter", envir = .GlobalEnv)) {
  assign(".appendix_counter", 0, envir = .GlobalEnv)
}

new_appendix <- function() {
  current <- get(".appendix_counter", envir = .GlobalEnv) + 1
  assign(".appendix_counter", current, envir = .GlobalEnv)
  LETTERS[current]
}

# --- Bookmark Helper for Cross-References ---
make_bkm <- function(project_id) {
  paste0("project_", gsub("[^A-Za-z0-9]", "_", project_id))
}
  
# make a table from projects df for selected theme -------------------------------------------------------

# get_section_table <- fuction(df,
#                              section_pick = "Salmon population monitoring")
# {
#   df_sub <- df %>%
#     filter(section == section_pick,
#            include %in% c("y", "Y")) %>%
#     select(project_id, source, project_title)
# }


make_section_table <- function(df,
                              section_pick = "Salmon population monitoring",
                              source_col = "source",
                              num_col = "project_id",
                              title_col = "project_title",
                              num_label = "ID",
                              title_label = "Project title",
                              header_bg = "royalblue",
                              header_fg = "white",
                              group_bg  = "darkgrey",
                              group_fg  = "white",
                              font_size_body = 8,
                              font_size_header = 11,
                              font_size_group = 10,
                              pad = 1,
                              w_num = 1.2,   # inches (Word-friendly)
                              w_title = 5.2  # inches (Word-friendly)
) {
  
  df_sub <- df %>%
    filter(section == section_pick,
           include %in% c("y", "Y")) %>%
    select(project_id, source, project_title)
  
  # 1) Clean + sort 
  df_clean <- df_sub |>
    dplyr::mutate(
      dplyr::across(dplyr::all_of(source_col), ~ stringr::str_squish(as.character(.x))),
      dplyr::across(dplyr::all_of(num_col),    ~ as.character(.x)),
      dplyr::across(dplyr::all_of(title_col),  ~ stringr::str_squish(as.character(.x)))
    ) |>
    dplyr::arrange(.data[[source_col]], .data[[num_col]])
  
  # 2) Insert divider rows using group key (.y) so it's always available
  df_div <- df_clean |>
    dplyr::group_by(.data[[source_col]]) |>
    dplyr::group_modify(function(.x, .y) {
      grp <- as.character(.y[[source_col]][1])
      
      tibble::tibble(
        .is_group = c(TRUE, rep(FALSE, nrow(.x))),
        # divider row has blank project_number, source label in title column
        !!num_col   := c("", .x[[num_col]]),
        !!title_col := c(grp, .x[[title_col]])
      )
    }) |>
    dplyr::ungroup()
  
  group_rows <- which(df_div$.is_group)
  
  # Keep only visible columns
  display_df <- df_div[, c(num_col, title_col), drop = FALSE]
  
  # 3) Build flextable
  ft <- flextable::flextable(display_df)
  
  # rows that are actual projects (not divider rows)
  data_rows <- which(display_df[[num_col]] != "")
  
  # bookmarks == project_ids (as-is)
  bkms <- display_df[[num_col]][data_rows]
  
  ft <- flextable::compose(
    x = ft,
    i = data_rows,
    j = num_col,
    value = flextable::as_paragraph(
      flextable::hyperlink_text(
        x   = bkms,                 # what the user sees (project_id)
        url = paste0("#", bkms)     # internal doc link target
      )
    )
  )
  
  
  # Optional: style like a hyperlink (blue + underlined)
  ft <- flextable::style(
    x    = ft,
    i    = data_rows,
    j    = num_col,
    pr_t = flextable::fp_text_default(color = "#0563C1", underlined = TRUE),
    part = "body"
  )
  
  # 4) Set header labels programmatically (NO :=)
  label_map <- stats::setNames(
    object = c(num_label, title_label),
    nm     = c(num_col, title_col)
  )
  ft <- do.call(flextable::set_header_labels, c(list(x = ft), as.list(label_map)))
  
  # 5) Base styling
  ft <- ft |>
    flextable::fontsize(size = font_size_body, part = "body") |>
    flextable::fontsize(size = font_size_header, part = "header") |>
    flextable::bold(part = "header") |>
    flextable::align(align = "left", part = "all") |>
    flextable::valign(valign = "top", part = "body") |>
    flextable::padding(padding = pad, part = "all") |>
    flextable::width(j = num_col, width = w_num) |>
    flextable::width(j = title_col, width = w_title) |>
    flextable::set_table_properties(layout = "fixed") |>
    flextable::bg(part = "header", bg = header_bg) |>
    flextable::color(part = "header", color = header_fg) 
  
  # 6) Merge + style each divider row (must be one-by-one)
  for (r in group_rows) {
    ft <- ft |>
      flextable::merge_at(i = r, j = c(num_col, title_col), part = "body") |>
      flextable::bg(i = r, bg = group_bg, part = "body") |>
      flextable::color(i = r, color = group_fg, part = "body") |>
      flextable::bold(i = r, part = "body") |>
      flextable::fontsize(i = r, size = font_size_group, part = "body") |>
      flextable::padding(i = r, padding = pad, part = "body")
  }
  
  # 7) Borders (Word-friendly)
  ft <- ft |>
    flextable::border_outer(
      part = "all",
      border = officer::fp_border(color = "#BFBFBF", width = 1)
    ) |>
    flextable::border_inner_h(
      part = "all",
      border = officer::fp_border(color = "#E0E0E0", width = 0.75)
    ) |>
    flextable::border_inner_v(
      part = "all",
      border = officer::fp_border(color = "#E0E0E0", width = 0.75)
    )
  
  ft
}
