# R/add_csas_tech_frontmatter.R
library(officer)
library(magrittr)

add_frontmatter <- function(main_docx, output_docx = NULL, meta = list()) {
  
  if (is.null(output_docx)) output_docx <- sub("\\.docx$", "_final.docx", main_docx)
  
  # Metadata extraction
  # If 'title' is empty (because we hid it), use 'report_title'
  title <- if(!is.null(meta$report_title)) meta$report_title else meta$title
  series     <- if(is.null(meta$doc_series)) "Canadian Technical Report of Fisheries and Aquatic Sciences" else meta$doc_series
  number     <- if(is.null(meta$series_number)) "XXXX" else meta$series_number
  year       <- if(is.null(meta$year)) format(Sys.Date(), "%Y") else meta$year
  authors    <- if(is.null(meta$authors)) list() else meta$authors
  citation   <- if(is.null(meta$citation)) "" else meta$citation
  abstract <- if(!is.null(meta$report_abstract)) meta$report_abstract else meta$abstract
  resume     <- if(is.null(meta$resume)) "" else meta$resume
  
  # Define precise DFO styles
  ft_title  <- fp_text(font.size = 16, bold = TRUE, font.family = "Arial")
  ft_author <- fp_text(font.size = 12, font.family = "Arial")
  ft_series <- fp_text(font.size = 12, font.family = "Arial")
  ft_normal <- fp_text(font.size = 11, font.family = "Arial")
  
  # Paragraph properties for centering
  center_p  <- fp_par(padding.bottom = 0, padding.top = 0)
  # Note: alignment is handled via the fp_p argument in fpar() using text.align
  
  doc <- read_docx(path = meta$reference_docx)
  
  #### 1. TITLE PAGE --------------------------------------------------
  
  # Series and Number at the top
  doc <- doc %>%
    body_add_fpar(fpar(ftext(series, prop = ft_series), 
                       fp_p = fp_par(text.align = "center"))) %>%
    body_add_fpar(fpar(ftext(number, prop = ft_series), 
                       fp_p = fp_par(text.align = "center")))
  
  # Vertical space to middle
  for(i in 1:8) doc <- body_add_par(doc, "", style = "Normal")
  
  # Title (Centered, Bold)
  doc <- doc %>%
    body_add_fpar(fpar(ftext(title, prop = ft_title), 
                       fp_p = fp_par(text.align = "center"))) %>%
    body_add_par("", style = "Normal")
  
  # Authors (Centered)
  if (length(authors) > 0) {
    author_names <- paste(sapply(authors, function(x) x$name), collapse = ", ")
    doc <- doc %>%
      body_add_fpar(fpar(ftext(author_names, prop = ft_author), 
                         fp_p = fp_par(text.align = "center")))
  }
  
  # Vertical space to bottom
  for(i in 1:8) doc <- body_add_par(doc, "", style = "Normal")
  
  # Department and Station Info (Bottom Centered)
  doc <- doc %>%
    body_add_fpar(fpar(ftext("Fisheries and Oceans Canada", prop = ft_normal), 
                       fp_p = fp_par(text.align = "center"))) %>%
    body_add_fpar(fpar(ftext("Pacific Biological Station", prop = ft_normal), 
                       fp_p = fp_par(text.align = "center"))) %>%
    body_add_fpar(fpar(ftext("3190 Hammond Bay Rd", prop = ft_normal), 
                       fp_p = fp_par(text.align = "center"))) %>%
    body_add_fpar(fpar(ftext("Nanaimo, BC V9T 6N7", prop = ft_normal), 
                       fp_p = fp_par(text.align = "center"))) %>%
    body_add_par("", style = "Normal") %>%
    body_add_fpar(fpar(ftext(year, prop = ft_normal), 
                       fp_p = fp_par(text.align = "center")))
  
  doc <- doc %>% body_add_break()
  
  #### 2. ABSTRACT ----------------------------------------------------
  
  if(nchar(abstract) > 0) {
    doc <- doc %>%
      body_add_par("ABSTRACT", style = "heading 1") %>%
      body_add_fpar(fpar(ftext(citation, prop = fp_text(italic = TRUE, font.family = "Arial")))) %>%
      body_add_par("", style = "Normal") %>%
      body_add_par(abstract, style = "Normal") %>%
      body_add_break()
  }
  
  #### 3. RÉSUMÉ ------------------------------------------------------
  
  if(nchar(resume) > 0) {
    doc <- doc %>%
      body_add_par("RÉSUMÉ", style = "heading 1") %>%
      body_add_fpar(fpar(ftext(citation, prop = fp_text(italic = TRUE, font.family = "Arial")))) %>%
      body_add_par("", style = "Normal") %>%
      body_add_par(resume, style = "Normal") %>%
      body_add_break()
  }
  
  # #### 3. TABLE OF CONTENTS -------------------------------------------
  # 
  # doc <- doc %>%
  #   body_add_par("Table of Contents", style = "heading 1") %>%
  #   body_add_toc(level = 3) %>%
  #   body_add_break()
  
  #### 5. LIST OF TABLES ----------------------------------------------
  doc <- doc %>%
    body_add_par("LIST OF TABLES", style = "heading 1") %>%
    body_add_toc(style = "Table Caption") %>%
    body_add_break()
  
  #### 6. LIST OF FIGURES ---------------------------------------------
  doc <- doc %>%
    body_add_par("LIST OF FIGURES", style = "heading 1") %>%
    body_add_toc(style = "Image Caption") %>%
    body_add_break()
  
  #### 4. APPEND MAIN CONTENT -----------------------------------------
  doc <- body_add_docx(doc, src = main_docx)
  
  print(doc, target = output_docx)
  return(output_docx)
}