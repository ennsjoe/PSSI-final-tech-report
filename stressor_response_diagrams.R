# Salmon Project Circle Diagram Generator
# Uses ggplot2 with curved text positioned along arcs

library(ggplot2)
library(readxl)
library(dplyr)
library(png)

# --- Configuration ---
SALMON_COLOR <- "#5d8a7a"
OUTPUT_DIR <- "output"
INPUT_FILE <- "data/raw/Stressor_Response_Data.xlsx"
SALMON_ICON <- "data/raw/salmon_icon.png"

# --- Helper: Create curved text as individual rotated letters ---
curved_text <- function(text, radius, center_angle, letter_spacing = NULL, 
                        above = TRUE, size = 4, color = SALMON_COLOR) {
  chars <- strsplit(toupper(text), "")[[1]]
  n <- length(chars)
  
  # Dynamic spacing: shorter text = wider spacing, longer text = tighter
  if (is.null(letter_spacing)) {
    letter_spacing <- max(2.5, min(6, 60 / n))
  }
  
  # Calculate angle for each character (centered on center_angle)
  total_span <- (n - 1) * letter_spacing
  
  if (above) {
    # Upper arc: letters go left-to-right (counterclockwise from left)
    angles <- center_angle + seq(total_span/2, -total_span/2, length.out = n)
    rotation <- angles - 90
  } else {
    # Lower arc: letters go left-to-right (clockwise from left)
    angles <- center_angle + seq(-total_span/2, total_span/2, length.out = n)
    rotation <- angles + 90
  }
  
  # Convert to radians for position
  angles_rad <- angles * pi / 180
  
  data.frame(
    x = radius * cos(angles_rad),
    y = radius * sin(angles_rad),
    label = chars,
    angle = rotation,
    size = size,
    color = color,
    stringsAsFactors = FALSE
  )
}

# --- Main diagram function ---
create_project_diagram <- function(project_no, title, measured_modelled_variable,
                                   salmon_response, response_type,
                                   stressor_threat_category, DFO_mgmt_decision,
                                   DFO_division) {
  
  # Build text data for all arcs
  text_layers <- rbind(
    # Upper arcs (above center, text curves upward)
    curved_text(response_type, radius = 2.6, center_angle = 90, above = TRUE, size = 5),
    curved_text(salmon_response, radius = 2.2, center_angle = 90, above = TRUE, size = 5),
    curved_text("SALMON RESPONSE", radius = 1.8, center_angle = 90, above = TRUE, size = 4),
    
    # Lower arcs (below center, text curves downward)
    curved_text(stressor_threat_category, radius = 1.8, center_angle = 270, above = FALSE, size = 4),
    curved_text("STRESSOR", radius = 2.2, center_angle = 270, above = FALSE, size = 4),
    curved_text(DFO_mgmt_decision, radius = 2.6, center_angle = 270, above = FALSE, size = 4),
    curved_text(DFO_division, radius = 3.0, center_angle = 270, above = FALSE, size = 4)
  )
  
  # Create center circle coordinates
  circle_data <- data.frame(
    x = 1.2 * cos(seq(0, 2*pi, length.out = 100)),
    y = 1.2 * sin(seq(0, 2*pi, length.out = 100))
  )
  
  # Load salmon icon if exists
  salmon_img <- NULL
  if (file.exists(SALMON_ICON)) {
    salmon_img <- readPNG(SALMON_ICON)
  }
  
  # Build the plot
  p <- ggplot() +
    
    # Center circle
    geom_polygon(data = circle_data, aes(x, y), 
                 fill = SALMON_COLOR, color = SALMON_COLOR)
  
  # Add salmon icon overlay if loaded
  if (!is.null(salmon_img)) {
    p <- p + annotation_raster(salmon_img, xmin = -0.9, xmax = 0.9, ymin = -0.9, ymax = 0.9)
  }
  
  p <- p +
    # Curved text (individual rotated letters)
    geom_text(data = text_layers,
              aes(x = x, y = y, label = label, angle = angle),
              size = text_layers$size, color = SALMON_COLOR, fontface = "bold") +
    
    # Project title and number (right of center)
    annotate("text", x = 1.5, y = 0, 
             label = paste0(project_no, "\n", title),
             size = 3.5, color = SALMON_COLOR, hjust = 0, fontface = "bold",
             lineheight = 0.9) +
    
    # Theme
    coord_fixed(xlim = c(-4, 5), ylim = c(-4, 4)) +
    theme_void() +
    theme(plot.background = element_rect(fill = "white", color = NA))
  
  return(p)
}

# --- Batch processing function ---
generate_all_diagrams <- function(excel_path, output_dir = OUTPUT_DIR) {
  
  if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
  
  projects <- read_excel(excel_path, sheet = "Project variables")
  
  for (i in seq_len(nrow(projects))) {
    row <- projects[i, ]
    
    p <- create_project_diagram(
      project_no = row$project_no,
      title = row$title,
      measured_modelled_variable = row$measured_modelled_variable,
      salmon_response = row$salmon_response,
      response_type = row$response_type,
      stressor_threat_category = row$stressor_threat_category,
      DFO_mgmt_decision = row$DFO_mgmt_decision,
      DFO_division = row$DFO_division
    )
    
    filename <- file.path(output_dir, paste0("project_", row$project_no, ".png"))
    ggsave(filename, p, width = 8, height = 6, dpi = 300, bg = "white")
    message("Created: ", filename)
  }
}

# --- Test with sample data ---
test_diagram <- function() {
  if (!dir.exists(OUTPUT_DIR)) dir.create(OUTPUT_DIR, recursive = TRUE)
  
  p <- create_project_diagram(
    project_no = "PRJ-001",
    title = "Fraser River Study",
    measured_modelled_variable = "Temperature, Flow",
    salmon_response = "Physiology",
    response_type = "Population",
    stressor_threat_category = "Threat Category",
    DFO_mgmt_decision = "DFO Management Decision",
    DFO_division = "DFO Division"
  )
  
  filename <- file.path(OUTPUT_DIR, "test_diagram.png")
  ggsave(filename, p, width = 8, height = 6, dpi = 300, bg = "white")
  message("Test diagram saved to ", filename)
  return(p)
}

# --- Run ---
test_diagram()
# generate_all_diagrams()