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
  if (is.na(text) || text == "") return(data.frame())
  
  chars <- strsplit(toupper(text), "")[[1]]
  n <- length(chars)
  if (n == 0) return(data.frame())
  
  # Dynamic spacing: shorter text = wider spacing, longer text = tighter
  if (is.null(letter_spacing)) {
    letter_spacing <- max(2.5, min(6, 60 / n))
  }
  
  total_span <- (n - 1) * letter_spacing
  
  if (above) {
    angles <- center_angle + seq(total_span/2, -total_span/2, length.out = n)
    rotation <- angles - 90
  } else {
    angles <- center_angle + seq(-total_span/2, total_span/2, length.out = n)
    rotation <- angles + 90
  }
  
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

# --- Helper: Create staggered text for multiple values ---
staggered_text <- function(values, base_radius, center_angle, above, size, color, 
                           angle_offset = 20, radius_offset = 0.12) {
  if (length(values) == 0) return(data.frame())
  
  result <- data.frame()
  n <- length(values)
  
  if (n == 1) {
    # Single value - centered
    result <- curved_text(values[1], radius = base_radius, center_angle = center_angle,
                          above = above, size = size, color = color)
  } else {
    # Multiple values - stagger by angle and slight radius variation
    angle_spread <- min(angle_offset * (n - 1), 50)  # Max spread of 50 degrees
    angles <- seq(-angle_spread/2, angle_spread/2, length.out = n)
    
    for (i in seq_along(values)) {
      # Alternate radius slightly to reduce overlap
      r_adj <- base_radius + ((i %% 2) * radius_offset)
      angle_adj <- center_angle + angles[i]
      
      layer <- curved_text(values[i], radius = r_adj, center_angle = angle_adj,
                           above = above, size = size, color = color)
      result <- rbind(result, layer)
    }
  }
  return(result)
}

# --- Single project diagram (aggregates multiple rows per project) ---
create_project_diagram <- function(project_data) {
  
  # project_data is a dataframe with all rows for one project
  project_no <- project_data$project_no[1]
  title <- project_data$title[1]
  
  # Get unique values helper
  get_unique <- function(col) unique(na.omit(project_data[[col]]))
  
  # Split measured_modelled_variable by variable_category
  response_vars <- unique(na.omit(
    project_data$measured_modelled_variable[project_data$variable_category == "response"]
  ))
  stressor_vars <- unique(na.omit(
    project_data$measured_modelled_variable[project_data$variable_category == "stressor"]
  ))
  
  salmon_response_types <- get_unique("salmon_response_type")
  stressor_categories <- get_unique("stressor_threat_category")
  decisions <- get_unique("DFO_mgmt_decision")
  divisions <- get_unique("DFO_division")
  
  # Build text layers with new structure
  text_layers <- rbind(
    # === ABOVE CIRCLE (closest to farthest) ===
    # 1. Response variables (black) - closest
    staggered_text(response_vars, base_radius = 1.5, center_angle = 90, above = TRUE, 
                   size = 3.5, color = "black"),
    # 2. "Salmon Response" heading (green)
    curved_text("SALMON RESPONSE", radius = 1.9, center_angle = 90, above = TRUE, 
                size = 3, color = SALMON_COLOR),
    # 3. salmon_response_type value (black)
    staggered_text(salmon_response_types, base_radius = 2.2, center_angle = 90, above = TRUE,
                   size = 3.5, color = "black"),
    # 4. "Salmon Response Type" heading (green) - farthest
    curved_text("SALMON RESPONSE TYPE", radius = 2.5, center_angle = 90, above = TRUE,
                size = 3, color = SALMON_COLOR),
    
    # === BELOW CIRCLE (closest to farthest) ===
    # 1. "Stressor" heading (green) - closest
    curved_text("STRESSOR", radius = 1.5, center_angle = 270, above = FALSE,
                size = 3, color = SALMON_COLOR),
    # 2. Stressor variables (black)
    staggered_text(stressor_vars, base_radius = 1.8, center_angle = 270, above = FALSE,
                   size = 3.5, color = "black"),
    # 3. "Human Activity Threat" heading (green)
    curved_text("HUMAN ACTIVITY THREAT", radius = 2.2, center_angle = 270, above = FALSE,
                size = 3, color = SALMON_COLOR),
    # 4. stressor_threat_category value (black)
    staggered_text(stressor_categories, base_radius = 2.5, center_angle = 270, above = FALSE,
                   size = 3.5, color = "black"),
    # 5. "DFO Management Decision" heading (green)
    curved_text("DFO MANAGEMENT DECISION", radius = 2.8, center_angle = 270, above = FALSE,
                size = 3, color = SALMON_COLOR),
    # 6. DFO_mgmt_decision value (black)
    staggered_text(decisions, base_radius = 3.1, center_angle = 270, above = FALSE,
                   size = 3.5, color = "black"),
    # 7. "DFO Division" heading (green)
    curved_text("DFO DIVISION", radius = 3.4, center_angle = 270, above = FALSE,
                size = 3, color = SALMON_COLOR),
    # 8. DFO_division value (black) - farthest
    staggered_text(divisions, base_radius = 3.7, center_angle = 270, above = FALSE,
                   size = 3.5, color = "black")
  )
  
  circle_data <- data.frame(
    x = 1.2 * cos(seq(0, 2*pi, length.out = 100)),
    y = 1.2 * sin(seq(0, 2*pi, length.out = 100))
  )
  
  salmon_img <- NULL
  if (file.exists(SALMON_ICON)) salmon_img <- readPNG(SALMON_ICON)
  
  p <- ggplot() +
    geom_polygon(data = circle_data, aes(x, y), fill = SALMON_COLOR, color = SALMON_COLOR)
  
  if (!is.null(salmon_img)) {
    p <- p + annotation_raster(salmon_img, xmin = -0.9, xmax = 0.9, ymin = -0.9, ymax = 0.9)
  }
  
  if (nrow(text_layers) > 0) {
    p <- p + geom_text(data = text_layers,
                       aes(x = x, y = y, label = label, angle = angle),
                       size = text_layers$size, color = text_layers$color, fontface = "bold")
  }
  
  # Wrap title to max width
  wrap_text <- function(text, width = 30) {
    if (is.na(text) || nchar(text) <= width) return(text)
    words <- strsplit(text, " ")[[1]]
    lines <- character()
    current_line <- ""
    for (word in words) {
      test_line <- if (current_line == "") word else paste(current_line, word)
      if (nchar(test_line) <= width) {
        current_line <- test_line
      } else {
        if (current_line != "") lines <- c(lines, current_line)
        current_line <- word
      }
    }
    if (current_line != "") lines <- c(lines, current_line)
    paste(lines, collapse = "\n")
  }
  
  wrapped_title <- wrap_text(title, width = 35)
  
  p <- p +
    annotate("text", x = 1.5, y = 0, 
             label = paste0(project_no, "\n", wrapped_title),
             size = 3, color = SALMON_COLOR, hjust = 0, fontface = "bold",
             lineheight = 0.9) +
    coord_fixed(xlim = c(-5, 6), ylim = c(-5, 4)) +
    theme_void() +
    theme(plot.background = element_rect(fill = "white", color = NA))
  
  return(p)
}

# --- Category aggregation diagram ---
create_category_diagram <- function(category_name, projects_df) {
  
  # Get unique values helper
  get_unique <- function(col) unique(na.omit(projects_df[[col]]))
  
  # Split measured_modelled_variable by variable_category
  response_vars <- unique(na.omit(
    projects_df$measured_modelled_variable[projects_df$variable_category == "response"]
  ))
  stressor_vars <- unique(na.omit(
    projects_df$measured_modelled_variable[projects_df$variable_category == "stressor"]
  ))
  
  salmon_response_types <- get_unique("salmon_response_type")
  stressor_categories <- get_unique("stressor_threat_category")
  decisions <- get_unique("DFO_mgmt_decision")
  divisions <- get_unique("DFO_division")
  
  # Get list of project numbers
  project_nums <- unique(na.omit(projects_df$project_no))
  n_projects <- length(project_nums)
  
  # Format project list (wrap if many)
  if (n_projects <= 10) {
    project_list <- paste(project_nums, collapse = "\n")
  } else {
    # Show first 8, then "..." and count
    project_list <- paste(c(project_nums[1:8], "...", paste0("(", n_projects, " total)")), collapse = "\n")
  }
  
  center_label <- paste0(category_name, "\n\nProjects:\n", project_list)
  
  # Build text layers
  text_layers <- rbind(
    # === ABOVE CIRCLE ===
    staggered_text(response_vars, base_radius = 1.5, center_angle = 90, above = TRUE, 
                   size = 3, color = "black"),
    curved_text("SALMON RESPONSE", radius = 1.9, center_angle = 90, above = TRUE, 
                size = 2.5, color = SALMON_COLOR),
    staggered_text(salmon_response_types, base_radius = 2.2, center_angle = 90, above = TRUE,
                   size = 3, color = "black"),
    curved_text("SALMON RESPONSE TYPE", radius = 2.5, center_angle = 90, above = TRUE,
                size = 2.5, color = SALMON_COLOR),
    
    # === BELOW CIRCLE ===
    curved_text("STRESSOR", radius = 1.5, center_angle = 270, above = FALSE,
                size = 2.5, color = SALMON_COLOR),
    staggered_text(stressor_vars, base_radius = 1.8, center_angle = 270, above = FALSE,
                   size = 3, color = "black"),
    curved_text("HUMAN ACTIVITY THREAT", radius = 2.2, center_angle = 270, above = FALSE,
                size = 2.5, color = SALMON_COLOR),
    staggered_text(stressor_categories, base_radius = 2.5, center_angle = 270, above = FALSE,
                   size = 3, color = "black"),
    curved_text("DFO MANAGEMENT DECISION", radius = 2.8, center_angle = 270, above = FALSE,
                size = 2.5, color = SALMON_COLOR),
    staggered_text(decisions, base_radius = 3.1, center_angle = 270, above = FALSE,
                   size = 3, color = "black"),
    curved_text("DFO DIVISION", radius = 3.4, center_angle = 270, above = FALSE,
                size = 2.5, color = SALMON_COLOR),
    staggered_text(divisions, base_radius = 3.7, center_angle = 270, above = FALSE,
                   size = 3, color = "black")
  )
  
  circle_data <- data.frame(
    x = 1.2 * cos(seq(0, 2*pi, length.out = 100)),
    y = 1.2 * sin(seq(0, 2*pi, length.out = 100))
  )
  
  salmon_img <- NULL
  if (file.exists(SALMON_ICON)) salmon_img <- readPNG(SALMON_ICON)
  
  p <- ggplot() +
    geom_polygon(data = circle_data, aes(x, y), fill = SALMON_COLOR, color = SALMON_COLOR)
  
  if (!is.null(salmon_img)) {
    p <- p + annotation_raster(salmon_img, xmin = -0.9, xmax = 0.9, ymin = -0.9, ymax = 0.9)
  }
  
  if (nrow(text_layers) > 0) {
    p <- p + geom_text(data = text_layers,
                       aes(x = x, y = y, label = label, angle = angle),
                       size = text_layers$size, color = text_layers$color, fontface = "bold")
  }
  
  p <- p +
    annotate("text", x = 1.5, y = 0, label = center_label,
             size = 2.5, color = SALMON_COLOR, hjust = 0, fontface = "bold",
             lineheight = 0.85, vjust = 0.5) +
    coord_fixed(xlim = c(-5, 7), ylim = c(-5, 4)) +
    theme_void() +
    theme(plot.background = element_rect(fill = "white", color = NA))
  
  return(p)
}

# --- Generate diagrams by category ---
generate_category_diagrams <- function(excel_path = INPUT_FILE, output_dir = OUTPUT_DIR) {
  if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
  
  projects <- read_excel(excel_path, sheet = "Project variables")
  categories <- unique(na.omit(projects$project_category))
  
  for (cat in categories) {
    cat_projects <- projects %>% filter(project_category == cat)
    p <- create_category_diagram(cat, cat_projects)
    
    filename <- file.path(output_dir, paste0("category_", gsub("[^A-Za-z0-9]", "_", cat), ".png"))
    ggsave(filename, p, width = 10, height = 8, dpi = 300, bg = "white")
    message("Created: ", filename)
  }
}

# --- Generate individual project diagrams (aggregating rows by project_no) ---
generate_project_diagrams <- function(excel_path = INPUT_FILE, output_dir = OUTPUT_DIR) {
  if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
  
  projects <- read_excel(excel_path, sheet = "Project variables")
  project_ids <- unique(projects$project_no)
  
  for (pid in project_ids) {
    if (is.na(pid)) next
    
    project_data <- projects %>% filter(project_no == pid)
    p <- create_project_diagram(project_data)
    
    filename <- file.path(output_dir, paste0("project_", pid, ".png"))
    ggsave(filename, p, width = 10, height = 8, dpi = 300, bg = "white")
    message("Created: ", filename)
  }
}

# --- Test diagram with sample multi-row data ---
test_diagram <- function() {
  if (!dir.exists(OUTPUT_DIR)) dir.create(OUTPUT_DIR, recursive = TRUE)
  
  # Simulate a project with multiple rows
  test_data <- data.frame(
    project_no = c("PRJ-001", "PRJ-001", "PRJ-001"),
    title = c("Fraser River Study", "Fraser River Study", "Fraser River Study"),
    measured_modelled_variable = c("Temperature", "Flow Rate", "Sediment Load"),
    variable_category = c("response", "response", "stressor"),
    salmon_response_type = c("population", "population", "population"),
    stressor_threat_category = c(NA, NA, "Habitat Degradation"),
    DFO_mgmt_decision = c("harvest rates", "harvest rates", "harvest rates"),
    DFO_division = c("Stock Assessment", "Stock Assessment", "Stock Assessment"),
    stringsAsFactors = FALSE
  )
  
  p <- create_project_diagram(test_data)
  
  filename <- file.path(OUTPUT_DIR, "test_diagram.png")
  ggsave(filename, p, width = 10, height = 8, dpi = 300, bg = "white")
  message("Test diagram saved to ", filename)
  return(p)
}

# --- Run ---
test_diagram()
generate_project_diagrams()
generate_category_diagrams()