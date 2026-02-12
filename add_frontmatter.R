# add_frontmatter.R
#
# Prepends CSAS front matter (title page, abstract, TOC) into the rendered
# report by opening the main docx and inserting content at the beginning.
#
# This avoids officer's body_add_docx() / altChunk approach, which requires
# Word to perform a manual document conversion before the content is visible.
#
# Strategy: cursor_begin() + pos = "before" inserts in reverse layout order
# so the final document reads: title page -> abstract -> TOC -> report body.

library(officer)
library(magrittr)

add_frontmatter <- function(main_docx, output_docx = NULL, meta = list()) {
  
  if (is.null(output_docx))
    output_docx <- sub("\\.docx$", "_wFrontMatter.docx", main_docx)
  
  # -- Metadata ---------------------------------------------------------------
  `%||%` <- function(x, y) if (is.null(x) || length(x) == 0) y else x
  
  title    <- as.character(meta$report_title  %||% meta$title %||% "")
  series   <- as.character(meta$doc_series    %||%
                             "Canadian Technical Report of Fisheries and Aquatic Sciences")
  number   <- as.character(meta$series_number %||% "XXXX")
  year     <- as.character(meta$year          %||% format(Sys.Date(), "%Y"))
  authors  <- meta$authors   %||% list()
  citation <- as.character(meta$citation %||% "")
  abstract <- as.character(meta$report_abstract %||% meta$abstract %||% "")
  resume   <- as.character(meta$resume   %||% "")
  
  # -- Text styles ------------------------------------------------------------
  ft_title  <- fp_text(font.size = 16, bold = TRUE,  font.family = "Arial")
  ft_author <- fp_text(font.size = 12,                font.family = "Arial")
  ft_series <- fp_text(font.size = 12,                font.family = "Arial")
  ft_normal <- fp_text(font.size = 11,                font.family = "Arial")
  ft_italic <- fp_text(font.size = 11, italic = TRUE, font.family = "Arial")
  center    <- fp_par(text.align = "center")
  
  # Helper: centred fpar
  cfpar <- function(text, prop = ft_normal)
    fpar(ftext(text, prop = prop), fp_p = center)
  
  # -- Open the rendered main report ------------------------------------------
  message("  Opening main report: ", main_docx)
  doc <- read_docx(main_docx)
  doc <- cursor_begin(doc)
  
  # -- Insert front matter in REVERSE layout order ----------------------------
  # Each pos="before" insert pushes everything else down, so inserting in
  # reverse gives the correct final order:
  #   series/number -> title -> authors -> abstract -> [resume] -> TOC -> body
  
  # 1. Section break - isolates front matter page numbering from report body
  doc <- body_end_section_portrait(doc)
  
  # 2. Page break before TOC
  doc <- body_add_break(doc, pos = "before")
  
  # 3. Table of Contents
  doc <- body_add_toc(doc, level = 3, pos = "before")
  doc <- body_add_par(doc, "Table of Contents", style = "heading 1",
                      pos = "before")
  
  # 4. Resume (optional French abstract)
  if (nchar(trimws(resume)) > 0) {
    doc <- body_add_break(doc, pos = "before")
    doc <- body_add_par(doc, resume, style = "Normal", pos = "before")
    doc <- body_add_par(doc, "", style = "Normal", pos = "before")
    if (nchar(trimws(citation)) > 0)
      doc <- body_add_fpar(doc, cfpar(citation, ft_italic), pos = "before")
    doc <- body_add_par(doc, "RESUME", style = "heading 1", pos = "before")
  }
  
  # 5. Abstract
  if (nchar(trimws(abstract)) > 0) {
    doc <- body_add_break(doc, pos = "before")
    doc <- body_add_par(doc, abstract, style = "Normal", pos = "before")
    doc <- body_add_par(doc, "", style = "Normal", pos = "before")
    if (nchar(trimws(citation)) > 0)
      doc <- body_add_fpar(doc, cfpar(citation, ft_italic), pos = "before")
    doc <- body_add_par(doc, "ABSTRACT", style = "heading 1", pos = "before")
  }
  
  # 6. Title page (inserted last = appears first)
  doc <- body_add_break(doc, pos = "before")
  
  # Bottom: year
  doc <- body_add_fpar(doc, cfpar(year, ft_normal), pos = "before")
  doc <- body_add_par(doc, "", style = "Normal", pos = "before")
  
  # Institution address
  doc <- body_add_fpar(doc, cfpar("Nanaimo, BC V9T 6N7",      ft_normal), pos = "before")
  doc <- body_add_fpar(doc, cfpar("3190 Hammond Bay Rd",       ft_normal), pos = "before")
  doc <- body_add_fpar(doc, cfpar("Pacific Biological Station", ft_normal), pos = "before")
  doc <- body_add_fpar(doc, cfpar("Fisheries and Oceans Canada", ft_normal), pos = "before")
  doc <- body_add_par(doc, "", style = "Normal", pos = "before")
  
  # Spacers to push content toward bottom of page
  for (i in seq_len(6))
    doc <- body_add_par(doc, "", style = "Normal", pos = "before")
  
  # Authors
  if (length(authors) > 0) {
    author_str <- paste(sapply(authors, function(x) x$name), collapse = ", ")
    doc <- body_add_fpar(doc, cfpar(author_str, ft_author), pos = "before")
  }
  
  doc <- body_add_par(doc, "", style = "Normal", pos = "before")
  
  # Title (centred, bold, large)
  doc <- body_add_fpar(doc, cfpar(title, ft_title), pos = "before")
  
  # Spacers to push title toward middle of page
  for (i in seq_len(6))
    doc <- body_add_par(doc, "", style = "Normal", pos = "before")
  
  # Top: series and series number
  doc <- body_add_fpar(doc, cfpar(number, ft_series), pos = "before")
  doc <- body_add_fpar(doc, cfpar(series, ft_series), pos = "before")
  
  # -- Save -------------------------------------------------------------------
  message("  Saving final report to: ", output_docx)
  print(doc, target = output_docx)
  
  n_elements <- nrow(officer::docx_summary(doc))
  message("  Total elements in merged document: ", n_elements)
  
  return(output_docx)
}