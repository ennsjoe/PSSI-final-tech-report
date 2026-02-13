# build-report.R
# Master build script for PSSI Technical Report
#
# Usage: source("build-report.R") from the project root, or
#        Rscript build-report.R from the terminal
#
# Steps:
#   1. Load _common.R  (packages, database helpers, icon maps, banner function)
#   2. Initialise / update the SQLite project database
#   3. Render the report via rmarkdown::render("index.Rmd")
#   4. Post-process: prepend CSAS front matter via add_frontmatter.R

library(here)
library(rmarkdown)
library(officedown)
library(officer)
library(yaml)

cat("\n================================================================\n")
cat(" PSSI Technical Report - Build Script\n")
cat("================================================================\n")
cat("Project root:", here::here(), "\n\n")

# ── STEP 1: Load _common.R ───────────────────────────────────────────────────
# Loads all packages, locates CSVs, sources helper-functions.R,
# initialises projects_df, and loads SECTION_ICON_MAP / SECTION_COLOUR_MAP

cat("STEP 1: Loading _common.R ...\n")

source(here("_common.R"))

cat("  ✓ _common.R loaded\n\n")


# ── STEP 2: Initialise / update database ────────────────────────────────────

cat("STEP 2: Checking project database ...\n")

tryCatch({
  
  csv_forms <- here("data", "processed", "pssi_form_data.csv")
  
  if (!file.exists(csv_forms)) {
    stop(
      "Extraction file not found at: ", csv_forms,
      "\nRun project_report_extraction.R before building the report."
    )
  }
  
  db_mtime   <- file.info(here("data", "projects.db"))$mtime
  form_mtime <- file.info(csv_forms)$mtime
  
  if (is.na(db_mtime) || form_mtime > db_mtime) {
    cat("  New data detected — rebuilding database ...\n")
    init_database(overwrite = TRUE)
    cat("  ✓ Database rebuilt\n\n")
  } else {
    cat("  ✓ Database is up to date\n\n")
  }
  
}, error = function(e) {
  cat("  ⚠ Database error:", conditionMessage(e), "\n")
  cat("  Attempting to proceed with existing database ...\n\n")
})


# ── STEP 3: Render report ────────────────────────────────────────────────────
# Output format and filename are declared in index.Rmd YAML.
# Output: PSSI-Technical-Report-2026.docx in the project root.

cat("STEP 3: Rendering report ...\n")

tryCatch({
  
  rmarkdown::render(
    input       = here("index.Rmd"),
    output_file = "PSSI-Technical-Report-2026.docx",
    output_dir  = here(),
    quiet       = FALSE,
    envir       = new.env(parent = globalenv())
  )
  
  cat("  ✓ Report rendered\n\n")
  
}, error = function(e) {
  stop("Render failed: ", conditionMessage(e))
})


# ── STEP 3.5: Fix internal hyperlinks ────────────────────────────────────────
# flextable hyperlink_text(url = "#bookmark") creates external relationships.
# fix_internal_hyperlinks() converts them to Word-native w:anchor links so
# Ctrl+Click (and regular click in reading view) navigates within the document.

cat("STEP 3.5: Fixing internal hyperlinks and adding BCSRIF bookmarks ...\n")

tryCatch({
  rendered_docx <- here("PSSI-Technical-Report-2026.docx")
  
  if (!file.exists(rendered_docx)) {
    cat("  ⚠ Rendered docx not found — skipping\n\n")
  } else {
    
    # 3.5a: Convert #fragment external hyperlinks to w:anchor internal links
    fix_internal_hyperlinks(rendered_docx)
    cat("  ✓ Internal hyperlinks converted\n")
    
    # 3.5b: Inject row-level bookmarks into the appendix B BCSRIF table
    bcsrif_ids <- projects_df %>%
      filter(source != "DFO Science", include %in% c("y", "Y")) %>%
      pull(project_id) %>%
      as.character() %>%
      unique()
    
    if (length(bcsrif_ids) > 0) {
      add_table_row_bookmarks(rendered_docx, bcsrif_ids)
      cat("  ✓ BCSRIF appendix bookmarks added\n")
    } else {
      cat("  ℹ No BCSRIF projects found — skipping appendix bookmarks\n")
    }
    
    cat("\n")
  }
  
}, error = function(e) {
  cat("  ⚠ Post-processing failed:", conditionMessage(e), "\n")
  cat("  Links may require Ctrl+Click — continuing build\n\n")
})

# ── STEP 4: Post-process — prepend CSAS front matter ────────────────────────

cat("STEP 4: Adding CSAS front matter ...\n")

tryCatch({
  
  front_matter_script <- here("add_frontmatter.R")
  
  if (!file.exists(front_matter_script)) {
    cat("  ⚠ add_frontmatter.R not found at:", front_matter_script, "\n")
    cat("  Skipping front matter step.\n\n")
  } else {
    
    cat("  Found add_frontmatter.R — sourcing ...\n")
    source(front_matter_script)
    
    # Read metadata from index.Rmd YAML, handling Windows CRLF endings
    index_lines <- readLines(here("index.Rmd"), warn = FALSE)
    index_lines <- trimws(index_lines, which = "right")
    yaml_bounds <- which(index_lines == "---")
    
    if (length(yaml_bounds) < 2) {
      stop("Could not locate YAML front matter in index.Rmd. ",
           "Ensure the file starts and ends the YAML block with '---'.")
    }
    
    meta <- yaml::yaml.load(
      paste(index_lines[(yaml_bounds[1] + 1):(yaml_bounds[2] - 1)], collapse = "\n")
    )
    
    cat("  YAML fields found:", paste(names(meta), collapse = ", "), "\n")
    cat("  report_abstract present:", !is.null(meta$report_abstract), "\n")
    
    out_name   <- "PSSI-Technical-Report-2026"
    main_docx  <- here(paste0(out_name, ".docx"))
    final_docx <- here(paste0(out_name, "_wFrontMatter.docx"))
    
    meta$reference_docx <- here("templates", "custom-reference.docx")
    
    if (!file.exists(main_docx)) {
      cat("  ⚠ Rendered docx not found at:", main_docx, "\n")
      cat("  Skipping front matter step.\n\n")
    } else {
      cat("  Building front matter document ...\n")
      add_frontmatter(
        main_docx   = main_docx,
        output_docx = final_docx,
        meta        = meta
      )
      cat("  ✓ Final report saved to:\n")
      cat("    ", final_docx, "\n\n")
    }
  }
  
}, error = function(e) {
  cat("  ✗ Front matter error:", conditionMessage(e), "\n")
  cat("  The rendered report (without front matter) is at:\n")
  cat("   ", here("PSSI-Technical-Report-2026.docx"), "\n\n")
})


# ── Done ─────────────────────────────────────────────────────────────────────

cat("================================================================\n")
cat(" Build complete\n")
cat("================================================================\n\n")