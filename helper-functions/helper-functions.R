# helper-functions.R

# ============================================================================
# 0. Configuration
# ============================================================================

if (!exists("OUTPUT_DB"))  OUTPUT_DB  <- here("data", "projects.db")
if (!exists("TABLE_NAME")) TABLE_NAME <- "projects"

`%||%` <- function(x, y) if (is.null(x) || length(x) == 0) y else x

# ============================================================================
# 1. CSV file discovery
# ============================================================================

find_csv_file <- function(filename, possible_paths = NULL) {
  if (is.null(possible_paths)) {
    possible_paths <- c(
      here(filename),
      here("data", filename),
      here("data", "raw", filename),
      here("data", "processed", filename),
      here("input", filename),
      here("inputs", filename)
    )
  }
  for (path in possible_paths) {
    if (file.exists(path)) return(path)
  }
  return(NULL)
}

if (!exists("RAW_CSV") || !exists("PROCESSED_CSV") || !exists("PRESENTATION_CSV")) {
  message("\n=== Locating CSV Files ===")
  
  RAW_CSV <- find_csv_file("report_project_list.csv")
  if (is.null(RAW_CSV)) {
    stop("\n❌ Cannot find 'report_project_list.csv'\n",
         "   Searched: project root, data/, data/raw/, data/processed/, input/, inputs/\n",
         "   Solution: manually set RAW_CSV at top of helper-functions.R")
  }
  
  PROCESSED_CSV <- find_csv_file("pssi_form_data.csv")
  if (is.null(PROCESSED_CSV)) {
    stop("\n❌ Cannot find 'pssi_form_data.csv'\n",
         "   Searched: project root, data/, data/raw/, data/processed/, input/, inputs/\n",
         "   Solution: manually set PROCESSED_CSV at top of helper-functions.R")
  }
  
  PRESENTATION_CSV <- find_csv_file("presentation_info.csv")
  if (is.null(PRESENTATION_CSV)) {
    warning("\n⚠ Cannot find 'presentation_info.csv' — skipping presentation data merge")
  }
  
  message("✓ Tracking CSV:     ", RAW_CSV)
  message("✓ Content CSV:      ", PROCESSED_CSV)
  if (!is.null(PRESENTATION_CSV)) message("✓ Presentation CSV: ", PRESENTATION_CSV)
}

# ============================================================================
# 2. Database functions
# ============================================================================

#' Initialise / rebuild the SQLite projects database
init_database <- function(overwrite = FALSE) {
  message("\n=== Initializing Projects Database ===")
  message("  Tracking:      ", RAW_CSV)
  message("  Content:       ", PROCESSED_CSV)
  if (!is.null(PRESENTATION_CSV)) message("  Presentations: ", PRESENTATION_CSV)
  
  # Load tracking
  message("\nLoading tracking data...")
  tracking_df <- read_csv(RAW_CSV, show_col_types = FALSE) %>%
    mutate(project_id = as.character(project_id))
  message("  ✓ ", nrow(tracking_df), " rows | columns: ",
          paste(names(tracking_df), collapse = ", "))
  
  # Load content
  message("\nLoading content data...")
  content_df <- read_csv(PROCESSED_CSV, show_col_types = FALSE) %>%
    mutate(project_id = as.character(project_id))
  message("  ✓ ", nrow(content_df), " rows | columns: ",
          paste(names(content_df), collapse = ", "))
  
  # Load presentations (optional — fills gaps for projects without form data)
  presentation_df <- NULL
  if (!is.null(PRESENTATION_CSV) && file.exists(PRESENTATION_CSV)) {
    message("\nLoading presentation data...")
    presentation_df <- read_csv(PRESENTATION_CSV, show_col_types = FALSE) %>%
      mutate(project_id = as.character(project_id)) %>%
      filter(!is.na(abstract) & abstract != "NA" & abstract != "")
    
    if (nrow(presentation_df) > 0) {
      presentation_df <- presentation_df %>%
        select(
          project_id,
          project_title  = pres_title,
          project_leads  = speakers,
          collaborations = collaborators,
          highlights     = abstract,
          source
        ) %>%
        filter(!is.na(project_title) & project_title != "NA" & project_title != "")
      message("  ✓ ", nrow(presentation_df), " rows with abstracts")
    } else {
      message("  ℹ No presentations with abstracts found")
      presentation_df <- NULL
    }
  }
  
  # Merge tracking + content
  message("\nMerging tracking + content...")
  final_df <- tracking_df %>%
    left_join(content_df, by = "project_id", suffix = c("", ".content"))
  
  # Fill gaps from presentation data (never overwrites content data)
  if (!is.null(presentation_df) && nrow(presentation_df) > 0) {
    message("Supplementing with presentation data where content is missing...")
    projects_with_content <- content_df %>% pull(project_id)
    to_add <- presentation_df %>% filter(!(project_id %in% projects_with_content))
    message("  Adding ", nrow(to_add), " presentation-only projects")
    
    if (nrow(to_add) > 0) {
      for (i in seq_len(nrow(to_add))) {
        pres <- to_add[i, ]
        idx  <- which(final_df$project_id == pres$project_id)
        if (length(idx) > 0) {
          final_df$project_title[idx]    <- pres$project_title
          final_df$project_leads[idx]    <- pres$project_leads
          final_df$collaborations[idx]   <- pres$collaborations
          final_df$highlights[idx]       <- pres$highlights
          if (is.na(final_df$source[idx]) || final_df$source[idx] == "")
            final_df$source[idx] <- pres$source
        }
      }
    }
  }
  
  # Drop duplicate columns from join
  dup_cols <- grep("\\.content$|\\.y$", names(final_df), value = TRUE)
  if (length(dup_cols) > 0) {
    message("  Dropping ", length(dup_cols), " duplicate columns")
    final_df <- final_df %>% select(-all_of(dup_cols))
  }
  
  final_df$project_number <- final_df$project_id
  if (!"include" %in% names(final_df)) final_df$include <- "y"
  
  message("  ✓ Final dataset: ", nrow(final_df), " rows, ", ncol(final_df), " columns")
  
  # Write to SQLite
  db_dir <- dirname(OUTPUT_DB)
  if (!dir.exists(db_dir)) dir.create(db_dir, recursive = TRUE)
  
  message("\nWriting to: ", OUTPUT_DB)
  tryCatch({
    con <- dbConnect(SQLite(), OUTPUT_DB)
    dbWriteTable(con, TABLE_NAME, final_df, overwrite = overwrite)
    row_count <- dbGetQuery(con, paste0("SELECT COUNT(*) as n FROM ", TABLE_NAME))$n
    columns   <- dbListFields(con, TABLE_NAME)
    dbDisconnect(con)
    
    message("✅ Database ready — ", row_count, " projects, ", length(columns), " columns")
    
    required <- c("project_id", "source", "project_title", "broad_header", "section", "include")
    missing  <- setdiff(required, columns)
    if (length(missing) > 0) warning("⚠ Missing required columns: ", paste(missing, collapse = ", "))
    
    return(invisible(final_df))
  }, error = function(e) {
    if (exists("con")) dbDisconnect(con)
    stop("Database write error: ", e$message)
  })
}

#' Get a database connection
get_db_con <- function(db_path = OUTPUT_DB) {
  if (!file.exists(db_path))
    stop("Database not found at: ", db_path, "\nRun init_database() to create it.")
  dbConnect(SQLite(), db_path)
}

get_db_connection <- function(...) get_db_con(...)  # backward compatibility alias

#' Check database status
check_database <- function(OUTPUT_DB = OUTPUT_DB) {
  if (!file.exists(OUTPUT_DB))
    return(list(exists = FALSE, message = "Database not found"))
  
  tryCatch({
    con    <- dbConnect(SQLite(), OUTPUT_DB)
    tables <- dbListTables(con)
    
    if (TABLE_NAME %in% tables) {
      row_count <- dbGetQuery(con, paste0("SELECT COUNT(*) as n FROM ", TABLE_NAME))$n
      columns   <- dbListFields(con, TABLE_NAME)
      dbDisconnect(con)
      return(list(exists = TRUE,
                  message = paste0("Database contains ", row_count, " projects"),
                  row_count = row_count, columns = columns))
    } else {
      dbDisconnect(con)
      return(list(exists = FALSE,
                  message = paste0("Database exists but '", TABLE_NAME, "' table not found")))
    }
  }, error = function(e) {
    if (exists("con")) dbDisconnect(con)
    return(list(exists = FALSE, message = paste0("Error: ", e$message)))
  })
}

# ============================================================================
# 3. Appendix counter
# ============================================================================

if (!exists(".appendix_counter", envir = .GlobalEnv))
  assign(".appendix_counter", 0, envir = .GlobalEnv)

new_appendix <- function() {
  current <- get(".appendix_counter", envir = .GlobalEnv) + 1
  assign(".appendix_counter", current, envir = .GlobalEnv)
  LETTERS[current]
}

# ============================================================================
# 4. Bookmark helper
# ============================================================================

#' Sanitise a project_id into a valid Word bookmark name
make_bkm <- function(project_id) {
  paste0("project_", gsub("[^A-Za-z0-9]", "_", project_id))
}

#' Format a project_id for display (replaces underscore with space)
#' "PSSI_2400" -> "PSSI 2400"
#' Bookmark/internal IDs use underscores; display uses spaces.
display_id <- function(project_id) {
  sub("^(PSSI|DFO|BCSRIF)_", "\\1 ", as.character(project_id))
}

# ============================================================================
# 5. Icon maps
# ============================================================================
# Loads SECTION_ICON_MAP and SECTION_COLOUR_MAP from assets/icons/icon_map.csv.
# The CSV must have columns: section, icon_file, colour, anchor
#   section   — must match values in projects_df$section exactly
#   icon_file — filename only (e.g. "freshwater.png"); files live in assets/icons/
#   colour    — hex colour for the banner header (e.g. "#2E8B57")
#   anchor    — Word bookmark name for the matching intro heading
#               (e.g. "freshwater-ecosystems"); used for back-hyperlinks
#
# Call load_icon_maps() once from _common.R after sourcing this file.

ICON_MAP_CSV <- here("data", "raw", "icon_map.csv")
ICON_DIR     <- here("assets", "icons")

# Single banner colour applied to all project pages
BANNER_COLOUR <- "#50B1BA"

load_icon_maps <- function(icon_map_csv = ICON_MAP_CSV,
                           icon_dir     = ICON_DIR) {
  
  if (!file.exists(icon_map_csv)) {
    warning("icon_map.csv not found at: ", icon_map_csv,
            "\nProject banners will show no icons and no back-links.",
            "\nExpected columns: theme_section, icon_file, anchor")
    SECTION_ICON_MAP   <<- character(0)
    SECTION_ANCHOR_MAP <<- character(0)
    return(invisible(NULL))
  }
  
  icon_map <- read_csv(icon_map_csv, show_col_types = FALSE)
  
  # Required columns
  required_cols <- c("theme_section", "icon_file", "anchor")
  missing_cols  <- setdiff(required_cols, names(icon_map))
  if (length(missing_cols) > 0)
    stop("icon_map.csv is missing required columns: ", paste(missing_cols, collapse = ", "))
  
  # Build full icon paths and check they exist
  icon_map <- icon_map %>%
    mutate(
      full_path = file.path(icon_dir, icon_file),
      file_ok   = file.exists(full_path)
    )
  
  missing_files <- icon_map %>% filter(!file_ok) %>% pull(icon_file)
  if (length(missing_files) > 0)
    warning("Icon files not found in ", icon_dir, ":\n  ",
            paste(missing_files, collapse = "\n  "),
            "\nBanners for affected sections will show no icon.")
  
  # NOTE: theme_section values must match the lookup key used in make_project_banner().
  # By default the banner looks up project$section. If your icon_map uses broader
  # theme labels, either:
  #   (a) add a 'theme' column to report_project_list.csv and pass
  #       project$theme to make_project_banner(), or
  #   (b) expand icon_map.csv to one row per exact section value.
  
  SECTION_ICON_MAP   <<- setNames(icon_map$full_path, icon_map$theme_section)
  SECTION_ANCHOR_MAP <<- setNames(icon_map$anchor,    icon_map$theme_section)
  
  n_ok <- sum(icon_map$file_ok)
  message("✓ Icon maps loaded — ", nrow(icon_map), " theme sections (",
          n_ok, " icon files found)")
  
  invisible(icon_map)
}

# ============================================================================
# 6. Section tables
# ============================================================================

#' Build a styled flextable listing projects for a given section.
#' Project ID cells contain Word REF field hyperlinks pointing to the
#' bookmark inserted by project-template.Rmd (= project$project_id).

make_section_table <- function(df,
                               section_pick     = "Salmon population monitoring",
                               source_col       = "source",
                               num_col          = "project_id",
                               title_col        = "project_title",
                               num_label        = "ID",
                               title_label      = "Project title",
                               header_bg        = "royalblue",
                               header_fg        = "white",
                               group_bg         = "darkgrey",
                               group_fg         = "white",
                               font_size_body   = 8,
                               font_size_header = 11,
                               font_size_group  = 10,
                               pad              = 1,
                               w_num            = 1.2,
                               w_title          = 5.2) {
  
  df_sub <- df %>%
    filter(section == section_pick, include %in% c("y", "Y")) %>%
    select(all_of(c(num_col, source_col, title_col)))
  
  # 1. Clean and sort
  df_clean <- df_sub %>%
    mutate(
      across(all_of(source_col), ~ stringr::str_squish(as.character(.x))),
      across(all_of(num_col),    ~ as.character(.x)),
      across(all_of(title_col),  ~ stringr::str_squish(as.character(.x)))
    ) %>%
    arrange(.data[[source_col]], .data[[num_col]])
  
  # 2. Insert source-group divider rows
  df_div <- df_clean %>%
    group_by(.data[[source_col]]) %>%
    group_modify(function(.x, .y) {
      grp <- as.character(.y[[source_col]][1])
      tibble::tibble(
        .is_group        = c(TRUE, rep(FALSE, nrow(.x))),
        !!num_col   := c("", .x[[num_col]]),
        !!title_col := c(grp, .x[[title_col]])
      )
    }) %>%
    ungroup()
  
  # Keep source alongside display columns so we can decide link type per row
  df_div_src  <- df_div %>% mutate(.row_source = .data[[source_col]])
  
  group_rows  <- which(df_div$.is_group)
  display_df  <- df_div[, c(num_col, title_col), drop = FALSE]
  data_rows   <- which(display_df[[num_col]] != "")
  
  # Which data rows are DFO Science (only these have bookmarks in the report)
  dfo_rows   <- data_rows[df_div_src$.row_source[data_rows] == "DFO Science"]
  other_rows <- setdiff(data_rows, dfo_rows)
  
  # 3. Build flextable
  ft <- flextable::flextable(display_df)
  
  # 4a. DFO Science rows — clickable hyperlink to project page bookmark.
  #     hyperlink_text() creates a proper Word hyperlink relationship with
  #     display text (the project ID). The "#bookmark" URL format navigates
  #     within the same document on click (supported by Word 2016+).
  #     This avoids REF fields, which embed bookmark content and cause
  #     nested table corruption.
  if (length(dfo_rows) > 0) {
    dfo_ids  <- display_df[[num_col]][dfo_rows]
    
    ft <- flextable::compose(
      x     = ft,
      i     = dfo_rows,
      j     = num_col,
      value = flextable::as_paragraph(
        flextable::hyperlink_text(
          x   = display_id(dfo_ids),   # "PSSI 2400" displayed
          url = paste0("#", dfo_ids)    # "#PSSI_2400" bookmark target
        )
      )
    )
    
    ft <- flextable::style(
      x    = ft,
      i    = dfo_rows,
      j    = num_col,
      pr_t = flextable::fp_text_default(color = "#0563C1", underlined = TRUE),
      part = "body"
    )
  }
  
  # 4b. Non-DFO rows (BCSRIF, etc.) — plain text, no hyperlink
  #     These projects have no project page, so REF fields would fail.
  if (length(other_rows) > 0) {
    other_ids <- display_df[[num_col]][other_rows]
    ft <- flextable::compose(
      x     = ft,
      i     = other_rows,
      j     = num_col,
      value = flextable::as_paragraph(
        flextable::as_chunk(display_id(other_ids))
      )
    )
  }
  
  # 5. Header labels
  label_map <- stats::setNames(c(num_label, title_label), c(num_col, title_col))
  ft <- do.call(flextable::set_header_labels, c(list(x = ft), as.list(label_map)))
  
  # 6. Base styling
  ft <- ft %>%
    flextable::fontsize(size = font_size_body,   part = "body")   %>%
    flextable::fontsize(size = font_size_header, part = "header") %>%
    flextable::bold(part = "header")                               %>%
    flextable::align(align = "left", part = "all")                 %>%
    flextable::valign(valign = "top", part = "body")               %>%
    flextable::padding(padding = pad, part = "all")                %>%
    flextable::width(j = num_col,   width = w_num)                 %>%
    flextable::width(j = title_col, width = w_title)               %>%
    flextable::set_table_properties(layout = "fixed")              %>%
    flextable::bg(part = "header", bg = header_bg)                 %>%
    flextable::color(part = "header", color = header_fg)
  
  # 7. Style divider rows
  for (r in group_rows) {
    ft <- ft %>%
      flextable::merge_at(i = r, j = c(num_col, title_col), part = "body") %>%
      flextable::bg(i = r, bg = group_bg, part = "body")                    %>%
      flextable::color(i = r, color = group_fg, part = "body")              %>%
      flextable::bold(i = r, part = "body")                                 %>%
      flextable::fontsize(i = r, size = font_size_group, part = "body")    %>%
      flextable::padding(i = r, padding = pad, part = "body")
  }
  
  # 8. Borders
  ft <- ft %>%
    flextable::border_outer(
      part = "all",
      border = officer::fp_border(color = "#BFBFBF", width = 1)
    ) %>%
    flextable::border_inner_h(
      part = "all",
      border = officer::fp_border(color = "#E0E0E0", width = 0.75)
    ) %>%
    flextable::border_inner_v(
      part = "all",
      border = officer::fp_border(color = "#E0E0E0", width = 0.75)
    )
  
  ft
}

# ============================================================================
# 7. Project page banner
# ============================================================================
#
# Renders a two-row flextable banner at the top of each project page:
#
#   Row 1 (theme colour): [icon]  project_id  •  section label  |  project title
#   Row 2 (light grey):   Leads   Species   Region   CU
#
# The icon is a hyperlink back to the matching intro section heading via a
# Word REF field, using the anchor stored in SECTION_ANCHOR_MAP.
#
# Arguments:
#   project      — named list (one row of projects_df via as.list())
#   icon_map     — SECTION_ICON_MAP  (built by load_icon_maps())
#   colour_map   — SECTION_COLOUR_MAP
#   anchor_map   — SECTION_ANCHOR_MAP
#   default_bg   — fallback header colour if section not in colour_map
#   icon_size    — icon height/width in inches inside the cell

make_project_banner <- function(project,
                                icon_map   = SECTION_ICON_MAP,
                                anchor_map = SECTION_ANCHOR_MAP,
                                icon_size  = 0.75) {
  
  # 1. Resolve theme values
  section_val <- as.character(project$section %||% "")
  bg_colour   <- BANNER_COLOUR
  icon_path   <- icon_map[[section_val]] %||% NULL
  anchor      <- anchor_map[[section_val]] %||% NULL
  
  # Validate icon path
  if (!is.null(icon_path) && !file.exists(icon_path)) icon_path <- NULL
  
  # 2. Text values
  proj_id     <- as.character(project$project_id    %||% "")
  proj_title  <- as.character(project$project_title %||% "")
  section_lbl <- section_val
  
  safe <- function(x, n = 999) {
    v <- trimws(as.character(x %||% ""))
    if (is.na(v) || nchar(v) == 0) return("\u2014")   # em-dash for empty
    if (nchar(v) > n) paste0(substr(v, 1, n - 1), "\u2026") else v
  }
  
  leads_val   <- safe(project$project_leads,  80)
  species_val <- safe(project$species,        50)
  region_val  <- safe(project$region,         40)
  cu_val      <- safe(project$cu,             60)
  
  # 3. Two-row data frame (4 columns)
  df <- data.frame(
    col1 = c(paste0(display_id(proj_id), "  \u2022  ", section_lbl),
             paste0("\U0001F464 ", leads_val)),
    col2 = c(proj_title, paste0("\U0001F41F ", species_val)),
    col3 = c("",         paste0("\U0001F5FA ", region_val)),
    col4 = c("",         paste0("CU: ", cu_val)),
    stringsAsFactors = FALSE
  )
  
  ft <- flextable::flextable(df)
  ft <- flextable::delete_part(ft, part = "header")
  
  # Merge title row cols 2-4
  ft <- flextable::merge_at(ft, i = 1, j = 2:4, part = "body")
  
  # 4. Row 1, col 1 — icon + ID text
  if (!is.null(icon_path)) {
    ft <- flextable::compose(
      x     = ft, i = 1, j = 1, part = "body",
      value = flextable::as_paragraph(
        flextable::as_image(src = icon_path, width = icon_size, height = icon_size),
        "  ",
        flextable::as_chunk(
          paste0(display_id(proj_id), "  \u2022  ", section_lbl),
          props = flextable::fp_text_default(color = "white", bold = TRUE, font.size = 9)
        )
      )
    )
  }
  
  # 5. Row 1, col 2 — project title
  ft <- flextable::compose(
    x     = ft, i = 1, j = 2, part = "body",
    value = flextable::as_paragraph(
      flextable::as_chunk(
        proj_title,
        props = flextable::fp_text_default(color = "white", bold = TRUE, font.size = 11)
      )
    )
  )
  
  # 6. Colours
  ft <- flextable::bg(ft,    i = 1, bg = bg_colour,  part = "body")
  ft <- flextable::color(ft, i = 1, color = "white", part = "body")
  ft <- flextable::bg(ft,    i = 2, bg = "#F0F4F8",  part = "body")
  ft <- flextable::color(ft, i = 2, color = "#333333", part = "body")
  ft <- flextable::fontsize(ft, i = 2, size = 8,      part = "body")
  ft <- flextable::italic(ft,   i = 2, italic = TRUE,  part = "body")
  
  # 7. Padding
  ft <- flextable::padding(ft, i = 1,
                           padding.top = 5, padding.bottom = 5,
                           padding.left = 6, padding.right = 6, part = "body")
  ft <- flextable::padding(ft, i = 2,
                           padding.top = 3, padding.bottom = 3,
                           padding.left = 6, padding.right = 6, part = "body")
  
  # 8. Alignment
  ft <- flextable::valign(ft, valign = "center", part = "body")
  ft <- flextable::align(ft,  align  = "left",   part = "body")
  
  # 9. Column widths (total ~6.4" for standard 1" margins on 8.5" page)
  ft <- flextable::width(ft, j = 1, width = 1.6)
  ft <- flextable::width(ft, j = 2, width = 2.8)
  ft <- flextable::width(ft, j = 3, width = 1.0)
  ft <- flextable::width(ft, j = 4, width = 1.0)
  ft <- flextable::set_table_properties(ft, layout = "fixed")
  
  # 10. Borders
  ft <- flextable::border_remove(ft)
  ft <- flextable::border_outer(
    ft, part = "all",
    border = officer::fp_border(color = bg_colour, width = 1.5)
  )
  ft <- flextable::border_inner_h(
    ft, part = "body",
    border = officer::fp_border(color = bg_colour, width = 0.5)
  )
  ft <- flextable::border_inner_v(
    ft, part = "body",
    border = officer::fp_border(color = "#CCCCCC", width = 0.5)
  )
  
  ft
}