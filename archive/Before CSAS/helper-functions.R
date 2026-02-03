# helper-functions.R
# Functions for managing project data in SQLite database
# Used by main-report.Rmd for Pacific Salmon Strategy Initiative report

library(DBI)
library(RSQLite)
library(readr)
library(dplyr)
library(gt)
library(here)

# --- Configuration using here() for project-relative paths ---
DB_PATH <- here("output", "projects.db")
CSV_PATH <- here("data", "raw", "report_project_list.csv")
TABLE_NAME <- "projects"

# Print paths on load so user knows where files are
message("Database path: ", DB_PATH)
message("CSV path: ", CSV_PATH)

# --- Database Setup Functions ---

#' Initialize the SQLite database from CSV
#' @param csv_path Path to the CSV file (default: CSV_PATH)
#' @param db_path Path to the SQLite database (default: DB_PATH)
#' @param overwrite If TRUE, recreate the table even if it exists
#' @return Invisible NULL
init_database <- function(csv_path = CSV_PATH, db_path = DB_PATH, overwrite = FALSE) {
  # Show where we're creating the database
  message("Initializing database at: ", db_path)
  message("Reading CSV from: ", csv_path)
  
  # Create output directory if it doesn't exist
  db_dir <- dirname(db_path)
  if (!dir.exists(db_dir)) {
    dir.create(db_dir, recursive = TRUE)
    message("Created directory: ", db_dir)
  }
  
  # Connect to database (creates file if doesn't exist)
  con <- dbConnect(RSQLite::SQLite(), db_path)
  on.exit(dbDisconnect(con))
  
  # Check if table exists and whether to overwrite
  table_exists <- dbExistsTable(con, TABLE_NAME)
  
  if (table_exists && !overwrite) {
    message("Database table '", TABLE_NAME, "' already exists. Use overwrite=TRUE to recreate.")
    return(invisible(NULL))
  }
  
  # Read CSV file
  projects <- read_csv(csv_path, show_col_types = FALSE)
  
  # Clean column names (remove BOM if present)
  names(projects) <- gsub("^\uFEFF", "", names(projects))
  
  # Write to database (overwrite if table exists)
  if (table_exists) {
    dbRemoveTable(con, TABLE_NAME)
  }
  
  dbWriteTable(con, TABLE_NAME, projects)
  
  # Create indexes for common queries
  dbExecute(con, paste0("CREATE INDEX IF NOT EXISTS idx_section ON ", TABLE_NAME, " (section)"))
  dbExecute(con, paste0("CREATE INDEX IF NOT EXISTS idx_broad_header ON ", TABLE_NAME, " (broad_header)"))
  dbExecute(con, paste0("CREATE INDEX IF NOT EXISTS idx_include ON ", TABLE_NAME, " (include)"))
  dbExecute(con, paste0("CREATE INDEX IF NOT EXISTS idx_source ON ", TABLE_NAME, " (source)"))
  
  n_rows <- nrow(projects)
  message("SUCCESS: Database initialized with ", n_rows, " projects")
  message("Database location: ", db_path)
  
  return(invisible(NULL))
}

#' Get a database connection
#' @param db_path Path to the SQLite database
#' @return Database connection object
get_db_connection <- function(db_path = DB_PATH) {
  if (!file.exists(db_path)) {
    stop("Database not found at: ", db_path, "\nRun init_database() first.")
  }
  dbConnect(RSQLite::SQLite(), db_path)
}

# --- Query Functions ---

#' Get all projects from the database
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only return projects where include = 'y'
#' @return Data frame of all projects
get_all_projects <- function(db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  if (include_only) {
    query <- paste0("SELECT * FROM ", TABLE_NAME, " WHERE include = 'y'")
  } else {
    query <- paste0("SELECT * FROM ", TABLE_NAME)
  }
  
  dbGetQuery(con, query)
}

#' Get projects by section
#' @param section_name Name of the section to filter by
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only return projects where include = 'y'
#' @return Data frame of matching projects
get_projects_by_section <- function(section_name, db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0("SELECT * FROM ", TABLE_NAME, " WHERE section = ?")
  if (include_only) {
    query <- paste0(query, " AND include = 'y'")
  }
  
  dbGetQuery(con, query, params = list(section_name))
}

#' Get projects by broad header
#' @param header_name Name of the broad header to filter by
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only return projects where include = 'y'
#' @return Data frame of matching projects
get_projects_by_header <- function(header_name, db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0("SELECT * FROM ", TABLE_NAME, " WHERE broad_header = ?")
  if (include_only) {
    query <- paste0(query, " AND include = 'y'")
  }
  
  dbGetQuery(con, query, params = list(header_name))
}

#' Get projects by source (BCSRIF or DFO)
#' @param source_name Source to filter by ('BCSRIF' or 'DFO')
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only return projects where include = 'y'
#' @return Data frame of matching projects
get_projects_by_source <- function(source_name, db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0("SELECT * FROM ", TABLE_NAME, " WHERE source = ?")
  if (include_only) {
    query <- paste0(query, " AND include = 'y'")
  }
  
  dbGetQuery(con, query, params = list(source_name))
}

#' Get a single project by project number
#' @param project_num Project number to look up
#' @param db_path Path to the SQLite database
#' @return Data frame with single project (or empty if not found)
get_project <- function(project_num, db_path = DB_PATH) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0("SELECT * FROM ", TABLE_NAME, " WHERE project_number = ?")
  dbGetQuery(con, query, params = list(project_num))
}

#' Search projects by keyword in title or description
#' @param keyword Search term
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only return projects where include = 'y'
#' @return Data frame of matching projects
search_projects <- function(keyword, db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  search_term <- paste0("%", keyword, "%")
  query <- paste0(
    "SELECT * FROM ", TABLE_NAME, 
    " WHERE (project_title LIKE ? OR description LIKE ?)"
  )
  if (include_only) {
    query <- paste0(query, " AND include = 'y'")
  }
  
  dbGetQuery(con, query, params = list(search_term, search_term))
}

#' Get summary counts by section
#' @param db_path Path to the SQLite database
#' @param include_only If TRUE, only count projects where include = 'y'
#' @return Data frame with section counts
get_section_summary <- function(db_path = DB_PATH, include_only = TRUE) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0(
    "SELECT section, COUNT(*) as n_projects FROM ", TABLE_NAME
  )
  if (include_only) {
    query <- paste0(query, " WHERE include = 'y'")
  }
  query <- paste0(query, " GROUP BY section ORDER BY n_projects DESC")
  
  dbGetQuery(con, query)
}

#' Get list of unique sections
#' @param db_path Path to the SQLite database
#' @return Character vector of section names
get_sections <- function(db_path = DB_PATH) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0(
    "SELECT DISTINCT section FROM ", TABLE_NAME, 
    " WHERE section IS NOT NULL AND include = 'y' ORDER BY section"
  )
  result <- dbGetQuery(con, query)
  result$section
}

#' Get list of unique broad headers
#' @param db_path Path to the SQLite database
#' @return Character vector of broad header names
get_broad_headers <- function(db_path = DB_PATH) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0(
    "SELECT DISTINCT broad_header FROM ", TABLE_NAME, 
    " WHERE broad_header IS NOT NULL AND include = 'y' ORDER BY broad_header"
  )
  result <- dbGetQuery(con, query)
  result$broad_header
}

# --- Table Generation Functions for Reports ---

#' Create a gt table of projects for a section
#' @param section_name Name of the section
#' @param db_path Path to the SQLite database
#' @param title Optional table title (defaults to section name)
#' @return gt table object
make_section_table <- function(section_name, db_path = DB_PATH, title = NULL) {
  projects <- get_projects_by_section(section_name, db_path)
  
  if (nrow(projects) == 0) {
    return(NULL)
  }
  
  # Select and rename columns for display
  table_data <- projects %>%
    select(
      Title = project_title,
      Source = source
    )
  
  table_title <- if (is.null(title)) {
    paste0("**Associated Projects: ", section_name, "**")
  } else {
    paste0("**", title, "**")
  }
  
  gt(table_data) %>%
    tab_header(title = md(table_title))
}

#' Create a detailed gt table with descriptions
#' @param projects Data frame of projects
#' @param title Table title
#' @return gt table object
make_detailed_table <- function(projects, title = "Projects") {
  if (nrow(projects) == 0) {
    return(NULL)
  }
  
  table_data <- projects %>%
    select(
      `Project #` = project_number,
      Title = project_title,
      Source = source,
      Location = location
    )
  
  gt(table_data) %>%
    tab_header(title = md(paste0("**", title, "**")))
}

#' Create a summary table by broad header
#' @param db_path Path to the SQLite database
#' @return gt table object
make_header_summary_table <- function(db_path = DB_PATH) {
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  query <- paste0(
    "SELECT broad_header as 'Research Theme', COUNT(*) as 'Number of Projects' ",
    "FROM ", TABLE_NAME, " ",
    "WHERE include = 'y' AND broad_header IS NOT NULL ",
    "GROUP BY broad_header ORDER BY COUNT(*) DESC"
  )
  
  result <- dbGetQuery(con, query)
  
  gt(result) %>%
    tab_header(title = md("**Projects by Research Theme**"))
}

# --- Utility Functions ---

#' Check database status
#' @param db_path Path to the SQLite database
#' @return List with database info
check_database <- function(db_path = DB_PATH) {
  if (!file.exists(db_path)) {
    return(list(
      exists = FALSE,
      path = db_path,
      message = paste0("Database does not exist at: ", db_path, "\nRun init_database() to create it.")
    ))
  }
  
  con <- get_db_connection(db_path)
  on.exit(dbDisconnect(con))
  
  n_total <- dbGetQuery(con, paste0("SELECT COUNT(*) as n FROM ", TABLE_NAME))$n
  n_included <- dbGetQuery(con, paste0("SELECT COUNT(*) as n FROM ", TABLE_NAME, " WHERE include = 'y'"))$n
  
  list(
    exists = TRUE,
    path = db_path,
    total_projects = n_total,
    included_projects = n_included,
    message = paste0("Database at: ", db_path, "\nContains ", n_total, " projects (", n_included, " included in report)")
  )
}

#' Rebuild the database from CSV
#' @param csv_path Path to the CSV file
#' @param db_path Path to the SQLite database
#' @return Invisible NULL
rebuild_database <- function(csv_path = CSV_PATH, db_path = DB_PATH) {
  init_database(csv_path, db_path, overwrite = TRUE)
}

#' Show current path configuration
#' @return Named list of paths
show_paths <- function() {
  paths <- list(
    project_root = here(),
    database = DB_PATH,
    csv_file = CSV_PATH
  )
  
  message("Project root: ", paths$project_root)
  message("Database: ", paths$database)
  message("CSV file: ", paths$csv_file)
  message("Database exists: ", file.exists(DB_PATH))
  message("CSV exists: ", file.exists(CSV_PATH))
  
  invisible(paths)
}


#make bookmark for word hyperlinks
make_bkm <- function(x, prefix = "DFO_") {
  # Word bookmark IDs: only letters/numbers + '-' and '_' are safe
  paste0(prefix, gsub("[^A-Za-z0-9_-]", "_", as.character(x)))
}



# Table functions ---------------------------------------------------------
gt_projects_table <- function(df,
                                  group_col = "source",
                                  num_col   = "project_number",
                                  title_col = "project_title",
                                  num_label = "Project #",
                                  title_label = "Project title",
                                  num_width_px = 130,
                                  title_width_px = 850,
                                  table_font_px = 8,
                                  col_label_font_px = 11,
                                  col_label_bg = "royalblue",
                                  row_group_bg = "darkgrey") {
  
  # packages (use :: so you don't *need* library() calls)
  df2 <- df |>
    dplyr::mutate(
      dplyr::across(dplyr::all_of(group_col), ~ stringr::str_squish(as.character(.x))),
      dplyr::across(dplyr::all_of(num_col),   ~ as.character(.x)),
      dplyr::across(dplyr::all_of(title_col), ~ stringr::str_squish(as.character(.x)))
    ) |>
    dplyr::arrange(.data[[group_col]], .data[[num_col]])
  
  # Build dynamic named vectors for cols_label()
  label_map <- rlang::set_names(
    c(num_label, title_label),
    c(num_col, title_col)
  )
  
  # Build dynamic list of formulas for cols_width()
  width_map <- rlang::set_names(
    list(
      rlang::quo(gt::px(!!num_width_px)),
      rlang::quo(gt::px(!!title_width_px))
    ),
    c(num_col, title_col)
  )
  
  gt::gt(df2, groupname_col = group_col) |>
    gt::cols_hide(columns = gt::all_of(group_col)) |>
    gt::cols_label(!!!label_map) |>
    gt::cols_width(!!!lapply(names(width_map), function(nm) {
      # turn "colname" + width expression into: colname ~ px(...)
      rlang::new_formula(rlang::sym(nm), rlang::eval_tidy(width_map[[nm]]))
    })) |>
    gt::tab_style(
      style = gt::cell_text(whitespace = "normal"),
      locations = gt::cells_body(columns = gt::all_of(title_col))
    ) |>
    gt::tab_options(
      table.width = gt::pct(100),
      table.font.size = gt::px(table_font_px),
      
      column_labels.font.size = gt::px(col_label_font_px),
      column_labels.font.weight = "bold",
      column_labels.background.color = col_label_bg,
      
      data_row.padding = gt::px(2),
      
      row_group.padding = gt::px(2),
      row_group.font.weight = "bold",
      row_group.font.size = gt::px(10),
      row_group.background.color = row_group_bg
    )
}



ft_projects_table <- function(df,
                                  source_col = "source",
                                  num_col = "project_number",
                                  title_col = "project_title",
                                  num_label = "Project #",
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
  
  stopifnot(requireNamespace("dplyr", quietly = TRUE))
  stopifnot(requireNamespace("stringr", quietly = TRUE))
  stopifnot(requireNamespace("tibble", quietly = TRUE))
  stopifnot(requireNamespace("flextable", quietly = TRUE))
  stopifnot(requireNamespace("officer", quietly = TRUE))
  
  # 1) Clean + sort (like your gt pipeline)
  df_clean <- df |>
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
  
  
  data_rows <- which(display_df[[num_col]] != "")
  bkms      <- make_bkm(display_df[[num_col]][data_rows])
  
  bslash <- intToUtf8(92)  # "\" character
  field_code <- sprintf('HYPERLINK %sl "%s"', bslash, bkms)
  
  ft <- flextable::compose(
    x = ft,
    i = data_rows,
    j = num_col,
    value = flextable::as_paragraph(
      flextable::as_word_field(field_code)
    )
  )
  
  
  # Optional: make them *look* like hyperlinks
  ft <- ft |>
    flextable::color(i = data_rows, j = num_col, color = "#0563C1", part = "body")
  
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

