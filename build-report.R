# build-report.R
# Master build script for PSSI Technical Report
#
# Usage: source("build-report.R") from the project root, or
#        Rscript build-report.R from the terminal
#
# Steps:
#   1. Load _common.R  (packages, helper functions, icon maps, banner function)
#   2. Render the report via rmarkdown::render("index.Rmd")
#   3. Post-process: inject bookmarks, fix hyperlinks
#   4. Post-process: prepend CSAS front matter via add_frontmatter.R

library(here)
library(rmarkdown)
library(officedown)
library(officer)
library(yaml)
library(magick)

#clear workspace to create new objects
rm(list = ls())

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



# ── STEP 2: Render report ────────────────────────────────────────────────────
# Output format and filename are declared in index.Rmd YAML.
# Output: PSSI-Technical-Report-2026.docx in the project root.

cat("STEP 2: Rendering report ...\n")

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




#--- STEP 3.2: Inject missing _rels/.rels ----
# Pandoc does not always emit _rels/.rels in its docx output. Officer (used by
# add_frontmatter in Step 4) requires this file to open the document.
# ensure_root_rels() injects a standard one if absent; no-ops if already there.

cat("STEP 3.2: Ensuring root _rels/.rels is present ...
")

tryCatch({
  rendered_docx <- here("PSSI-Technical-Report-2026.docx")
  if (file.exists(rendered_docx)) {
    ensure_root_rels(rendered_docx)
    cat("  ✓ Root relationships verified

")
  } else {
    cat("  ⚠ Rendered docx not found - skipping

")
  }
}, error = function(e) {
  cat("  ⚠ ensure_root_rels failed:", conditionMessage(e), "

")
})

# --- STEP 3.3: Inject project heading bookmarks -------------------------
# Searches the rendered docx for each project's heading paragraph and
# injects a Word bookmark using the raw project_id as the name.
# Must run BEFORE fix_internal_hyperlinks (step 3.5) so the bookmark
# targets exist when the anchor links are written.

cat("STEP 3.3: Injecting project heading bookmarks ...\n")

tryCatch({
  if (file.exists(rendered_docx) && exists("projects_df")) {
    inject_project_bookmarks(rendered_docx, projects_df)
    cat("  \u2713 Project bookmarks injected\n\n")
  } else {
    cat("  \u26a0 Skipping: rendered docx or projects_df not found\n\n")
  }
}, error = function(e) {
  cat("  \u26a0 inject_project_bookmarks failed:", conditionMessage(e), "\n\n")
})

# --- STEP 3.4: Inject BCSRIF table row bookmarks -----------------------
# BCSRIF projects appear as rows in the Appendix B table, not as pages.
# add_table_row_bookmarks() searches every table row's first cell for a
# project_id match and injects a bookmark using the raw project_id as name,
# which matches the #BCSRIF_2022_XXX anchors built by make_section_table().

cat("STEP 3.4: Injecting BCSRIF table row bookmarks ...\n")

tryCatch({
  if (file.exists(rendered_docx) && exists("projects_df")) {
    bcsrif_ids <- projects_df$project_id[
      projects_df$source == "BCSRIF" &
        projects_df$include %in% c("y", "Y")
    ]
    if (length(bcsrif_ids) > 0) {
      add_table_row_bookmarks(rendered_docx, bcsrif_ids)
      cat("  \u2713 BCSRIF table row bookmarks injected\n\n")
    } else {
      cat("  \u26a0 No BCSRIF projects found in projects_df\n\n")
    }
  } else {
    cat("  \u26a0 Skipping: rendered docx or projects_df not found\n\n")
  }
}, error = function(e) {
  cat("  \u26a0 add_table_row_bookmarks failed:", conditionMessage(e), "\n\n")
})

# --- STEP 3.5: Convert #fragment hyperlinks to Word-native w:anchor links ----
# Runs on the pre-frontmatter docx.  Step 4.5 repeats this on the final
# _wFrontMatter.docx in case officer's merge step regenerates hyperlinks.

cat("STEP 3.5: Converting internal hyperlinks to w:anchor format ...\n")

rendered_docx <- here("PSSI-Technical-Report-2026.docx")

tryCatch({
  if (file.exists(rendered_docx)) {
    fix_internal_hyperlinks(rendered_docx)
    cat("  \u2713 Internal hyperlinks converted\n\n")
  } else {
    cat("  \u26a0 Rendered docx not found - skipping\n\n")
  }
}, error = function(e) {
  cat("  \u26a0 fix_internal_hyperlinks failed:", conditionMessage(e), "\n\n")
})


# --- STEP 4: Post-process — prepend CSAS front matter ----

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


# --- STEP 4.5: Re-run hyperlink fix on the final merged document -----------
# officer's add_frontmatter assembles a brand-new docx by copying body
# elements from the rendered report.  It preserves w:anchor attributes, but
# running the fix again is cheap insurance and ensures the audit output
# reflects the actual file the user will open.

cat("STEP 4.5: Fixing hyperlinks in final document and auditing ...\n")

tryCatch({
  final_docx_path <- here("PSSI-Technical-Report-2026_wFrontMatter.docx")
  if (file.exists(final_docx_path)) {
    fix_internal_hyperlinks(final_docx_path)
    cat("  \u2713 Hyperlinks fixed in final document\n")
    audit_hyperlinks(final_docx_path)
  } else {
    # Fall back to auditing the pre-frontmatter docx
    pre_docx_path <- here("PSSI-Technical-Report-2026.docx")
    if (file.exists(pre_docx_path)) {
      audit_hyperlinks(pre_docx_path)
    } else {
      cat("  \u26a0 No output docx found to audit\n\n")
    }
  }
}, error = function(e) {
  cat("  \u26a0 Step 4.5 failed:", conditionMessage(e), "\n\n")
})


# --- Done ----
cat("================================================================\n")
cat(" Build complete\n")
cat("================================================================\n\n")