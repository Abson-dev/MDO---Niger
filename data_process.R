# Load required libraries
library(tidyverse)
library(readxl)
library(readr) # For parse_number()

# --- Define File Path and Sheet Name ---
file_path <- "MDO 2024 S43.xls"
#sheet_name <- "Palu confirmé"
sheet_name <- "Palu Susp"

# --- PART 1: Determine Column Names Programmatically & Read Data Robustly ---

# Step 1a: Read header rows to construct column names
# We need rows 3 and 5 to get the full header information.
# Let's read up to row 5 to see the full header structure.
message("--- Debugging Part 1: Reading header rows to construct column names ---")
header_df_raw <- read_excel(
  path = file_path,
  sheet = sheet_name,
  range = cell_rows(1:5), # Read rows 1 to 5 to get all header parts
  col_names = FALSE,     # Do not use any row as names yet
  col_types = "text"     # Read everything as text
)
message("head(header_df_raw) after initial header read:")
print(head(header_df_raw, n = 5))
message("Dimensions of header_df_raw: ", paste(dim(header_df_raw), collapse = "x"))

# Based on your LAST header_df output:
# Row 3 (index 3 in header_df_raw) has primary headers: Region, District, Isocode, Année, Population, Semaine X...
# Row 5 (index 5 in header_df_raw) has secondary headers: cas, décès, Taux d'attaque, Létalité (starting under Semaine X)...

# FIXED COLUMNS: From header_df_raw[3, 1:5]
fixed_cols_from_excel <- as.character(header_df_raw[3, 1:5]) # "Region", "District", "Isocode", "Année", "Population"

weekly_constructed_names <- c()
# CORRECTED: current_semaine_col_in_excel now starts at 6 (Excel column F)
current_semaine_col_in_excel <- 6

# Iterate while we find a "Semaine X" header in row 3 (header_df_raw[3, ...])
# and enough sub-headers in row 5 (header_df_raw[5, ...])
while(current_semaine_col_in_excel <= ncol(header_df_raw) &&
      current_semaine_col_in_excel + 3 <= ncol(header_df_raw)) { # Ensure there's space for 4 sub-cols
  
  semaine_header_val <- as.character(header_df_raw[3, current_semaine_col_in_excel])
  # Sub-headers for this week are also at current_semaine_col_in_excel in ROW 5
  sub_headers_vals <- as.character(header_df_raw[5, current_semaine_col_in_excel:(current_semaine_col_in_excel + 3)])
  
  if (!is.na(semaine_header_val) && semaine_header_val != "" &&
      # Ensure at least one sub-header is not NA/empty to signify a valid block
      !all(is.na(sub_headers_vals) | sub_headers_vals == "")) {
    week_num <- str_extract(semaine_header_val, "\\d+")
    
    # Define the exact sub-header names (lowercase for janitor::clean_names consistency)
    cleaned_sub_headers <- c("cas", "deces", "taux_d_attaque", "letalite")
    
    weekly_constructed_names <- c(weekly_constructed_names,
                                  paste0("Semaine ", week_num, "_", cleaned_sub_headers))
  } else {
    message(paste0("Stopping weekly column construction at Excel column: ",
                   letters[current_semaine_col_in_excel], " (R index ", current_semaine_col_in_excel, ") ",
                   "Semaine header: '", semaine_header_val, "'",
                   ", Sub-headers: ", paste(sub_headers_vals, collapse = ", "),
                   ". Likely reached end of data or invalid section."))
    break # Exit loop if no valid week header or sub-headers
  }
  current_semaine_col_in_excel <- current_semaine_col_in_excel + 4 # Move to next week's starting column
}

all_column_names_constructed <- c(fixed_cols_from_excel, weekly_constructed_names)
# Remove any accidental NAs or empty strings from the final list
all_column_names_constructed <- all_column_names_constructed[!is.na(all_column_names_constructed) & all_column_names_constructed != ""]

message("\nConstructed all_column_names (first 20 and last 10):")
print(head(all_column_names_constructed, 20))
print(tail(all_column_names_constructed, 10))
message("Total number of constructed columns: ", length(all_column_names_constructed))


# Step 1b: Read the actual data block starting from row 6
message("\n--- Debugging Part 2: Reading actual data block (skip=5) ---")
health_data_raw <- read_excel(
  path = file_path,
  sheet = sheet_name,
  skip = 5, # Skip first 5 rows, so data starts from Excel row 6
  col_names = FALSE, # We will assign names later
  col_types = "text" # Read everything as text
)

# Crucial Debugging Point 1: Check dimensions and head of raw data
message("head(health_data_raw) after read (should be actual data):")
print(head(health_data_raw))
message("Dimensions of health_data_raw: ", paste(dim(health_data_raw), collapse = "x"))

# Crucial Debugging Point 2: Compare number of columns
if (ncol(health_data_raw) < length(all_column_names_constructed)) {
  warning("Number of data columns read (", ncol(health_data_raw), ") is LESS than constructed column names (",
          length(all_column_names_constructed), "). Trimming constructed names.")
  all_column_names_final <- all_column_names_constructed[1:ncol(health_data_raw)]
} else if (ncol(health_data_raw) > length(all_column_names_constructed)) {
  # This is the case we are trying to solve: constructed names are too few
  warning("Number of data columns read (", ncol(health_data_raw), ") is MORE than constructed column names (",
          length(all_column_names_constructed), "). Data will have default `...X` names for extra columns if not handled.")
  # Adjust: If we have more raw columns than constructed names, we need to extend constructed names
  # This suggests the loop to build names broke too early or the fixed names are incorrect.
  # For now, let's truncate raw data to match constructed names, but this means we might lose columns.
  # The real fix is to ensure all_column_names_constructed is correct.
  all_column_names_final <- all_column_names_constructed
  # If this still fails, we need to pad all_column_names_final with generic names
  # or re-evaluate the column construction loop even more rigorously.
  if (ncol(health_data_raw) > length(all_column_names_final)) {
    message("Padding constructed names to match raw data columns for `set_names`.")
    num_extra_cols <- ncol(health_data_raw) - length(all_column_names_final)
    padded_names <- c(all_column_names_final, paste0("Unknown_", 1:num_extra_cols))
    all_column_names_final <- padded_names
  }
  
} else {
  message("Number of data columns read matches constructed column names.")
  all_column_names_final <- all_column_names_constructed
}

# Assign the constructed column names to the raw data
health_data <- health_data_raw %>%
  set_names(all_column_names_final)

# Crucial Debugging Point 3: Verify column names and a sample of data
message("\n--- Debugging Part 3: health_data after setting names ---")
print("Column names of health_data:")
print(colnames(health_data))
print("First few rows of health_data (after setting names):")
print(head(health_data))

# --- PART 2: Clean and Reshape Data ---

message("\n--- Part 2: Cleaning and Reshaping Data ---")

# Filter out summary rows (e.g., "Total Agadez") that might be present in the data rows
health_data_filtered <- health_data %>%
  # Filter out rows where 'Region' is NA, "Pays:", "Region", or contains "Total"
  filter(!is.na(Region) & Region != "Pays:" & Region != "Region" & !grepl("Total", Region, ignore.case = TRUE))

# Clean column names using janitor (e.g., "Taux d'attaque" -> "taux_d_attaque")
health_data_clean_names <- health_data_filtered %>%
  janitor::clean_names()

# Crucial Debugging Point 4: Inspect values before numeric conversion
message("\n--- Debugging: Population and first weekly metric values before numeric conversion ---")
print("Sample of population column:")
print(head(health_data_clean_names$population))
print("Sample of first weekly metric (e.g., semaine_1_cas):")
# Dynamically get the first 'semaine_X_cas' column if it exists
first_semaine_cas_col <- colnames(health_data_clean_names) %>%
  str_subset("^semaine_\\d+_cas$") %>%
  head(1)
if (length(first_semaine_cas_col) > 0) {
  print(head(health_data_clean_names[[first_semaine_cas_col]]))
} else {
  message("No 'semaine_X_cas' column found for debugging.")
}

# Convert numeric columns using readr::parse_number for robustness
health_data_typed <- health_data_clean_names %>%
  mutate(
    population = readr::parse_number(population),
    across(starts_with("semaine_") &
             (contains("cas") | contains("deces") |
                contains("taux_d_attaque") | contains("letalite")),
           readr::parse_number)
  )

message("\n--- Debugging: health_data_typed after numeric conversion ---")
print(head(health_data_typed))
print(summary(health_data_typed$population))


# Reshape the data from wide to long format
health_data_long <- health_data_typed %>%
  pivot_longer(
    cols = starts_with("semaine_"),
    names_to = c("Week_Number_Raw", "Metric_Raw"), # Use temp names for pivoting
    names_pattern = "semaine_(\\d+)_(.*)", # Use pattern derived from janitor::clean_names
    values_to = "Value"
  ) %>%
  mutate(
    Week = as.integer(Week_Number_Raw), # Convert to integer
    # Rename metrics for clarity
    Metric = case_when(
      Metric_Raw == "cas" ~ "Cases",
      Metric_Raw == "deces" ~ "Deaths",
      Metric_Raw == "taux_d_attaque" ~ "Attack_Rate",
      Metric_Raw == "letalite" ~ "Fatality_Rate",
      TRUE ~ Metric_Raw # Catch any unexpected metric names
    )
  ) %>%
  select(
    region, district, population,
    week = Week,
    metric = Metric, # Keep metric in long format for now
    value = Value
  )

# Convert back to wide format after cleaning Week and Metric if desired
health_data_final_wide <- health_data_long %>%
  pivot_wider(
    names_from = metric,
    values_from = value,
    values_fn = first # Use first value to preserve original data, handle potential duplicates
  ) %>%
  arrange(region, district, week) # Arrange for consistent order

# Save the cleaned data to a new CSV file
health_data_final_wide <- health_data_final_wide %>% 
  dplyr::select("region",	"district",	"population",	"week",	"Cases",	"Deaths",	"Attack_Rate",	"Fatality_Rate") %>% 
  dplyr::filter(!is.na(week)) %>% 
  dplyr::filter(!region %in% c("REGION")) %>% 
  dplyr::mutate(week = paste0("Week ",week),
                population = round(population,0))
output_file_name <- paste0("health_data_cleaned_final_",sheet_name,".csv")
#output_file_name <- "health_data_cleaned_final_Palu_confirme.csv"
write_csv(health_data_final_wide, output_file_name,na = "")
message(paste0("\nCleaned data saved to: ", output_file_name))

######################################
# --- PART 3: Generate Graphics per District and Save to Excel ---

message("--- Starting Plot Generation and Excel Export ---")

# 1. Get unique districts
unique_districts <- unique(health_data_final_wide$district)

# 2. Create a new Excel workbook
wb <- createWorkbook()

# Define plot dimensions for Excel (in inches for a fixed DPI)
# Using 300 dpi for better quality, so width/height in inches
plot_width_in <- 6 # inches
plot_height_in <- 4 # inches


# Loop through each unique district
for (district_name in unique_districts) {
  message(paste("Processing district:", district_name))
  
  # Filter data for the current district and re-parse 'week' for plotting
  district_data <- health_data_final_wide %>%
    filter(district == district_name) %>%
    mutate(
      # Correctly extract numeric week for plotting, handling potential non-numeric entries
      week_num = as.numeric(str_extract(week, "\\d+")), # Extract digits then convert to numeric
      # Ensure numeric columns are actually numeric after all transformations
      Cases = as.numeric(Cases),
      Deaths = as.numeric(Deaths),
      Attack_Rate = as.numeric(Attack_Rate),
      Fatality_Rate = as.numeric(Fatality_Rate)
    ) %>%
    # Filter out rows where week_num is NA (i.e., couldn't be parsed) or where essential metrics are NA
    filter(!is.na(week_num) & !is.na(Cases) & !is.na(Deaths) & !is.na(Attack_Rate) & !is.na(Fatality_Rate)) %>%
    # Sort by week_num to ensure line plots are correct
    arrange(week_num)
  
  # Skip if there's no data for this district or no valid cases data after filtering and parsing
  if (nrow(district_data) == 0 || nrow(district_data) < 2) { # Need at least 2 points for a line
    warning(paste("Insufficient valid data (less than 2 rows) for plotting for district:", district_name, "- Skipping plot generation."))
    next
  }
  
  # --- Create the plots for the current district ---
  # Plot 1: Cases over time
  p_cases <- ggplot(district_data, aes(x = week_num, y = Cases)) +
    geom_line(color = "steelblue", size = 0.8) +
    geom_point(color = "darkblue", size = 1.8) +
    labs(
      title = paste("Malaria Cases in", district_name),
      x = "Epidemiological Week",
      y = "Number of Cases"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
      axis.title = element_text(size = 12),
      axis.text = element_text(size = 10)
    )
  
  # Plot 2: Deaths over time
  p_deaths <- ggplot(district_data, aes(x = week_num, y = Deaths)) +
    geom_line(color = "firebrick", size = 0.8) +
    geom_point(color = "darkred", size = 1.8) +
    labs(
      title = paste("Malaria Deaths in", district_name),
      x = "Epidemiological Week",
      y = "Number of Deaths"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
      axis.title = element_text(size = 12),
      axis.text = element_text(size = 10)
    )
  
  # Plot 3: Attack Rate over time
  p_attack_rate <- ggplot(district_data, aes(x = week_num, y = Attack_Rate)) +
    geom_line(color = "darkgreen", size = 0.8) +
    geom_point(color = "forestgreen", size = 1.8) +
    labs(
      title = paste("Malaria Attack Rate in", district_name),
      x = "Epidemiological Week",
      y = "Attack Rate"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
      axis.title = element_text(size = 12),
      axis.text = element_text(size = 10)
    )
  
  # Plot 4: Fatality Rate over time
  p_fatality_rate <- ggplot(district_data, aes(x = week_num, y = Fatality_Rate)) +
    geom_line(color = "purple", size = 0.8) +
    geom_point(color = "darkmagenta", size = 1.8) +
    labs(
      title = paste("Malaria Fatality Rate in", district_name),
      x = "Epidemiological Week",
      y = "Fatality Rate (%)"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
      axis.title = element_text(size = 12),
      axis.text = element_text(size = 10)
    )
  
  # --- Add a sheet for the district and insert plots ---
  # Clean the district name for use as a sheet name (Excel sheet names have character limits and invalid chars)
  sheet_safe_name <- str_sub(str_replace_all(district_name, "[^a-zA-Z0-9_]", "_"), 1, 31)
  
  # Add the sheet to the workbook
  addWorksheet(wb, sheetName = sheet_safe_name)
  
  # Insert the plots. IMPORTANT: print() the plot immediately before openxlsx::insertPlot()
  # This makes the plot active on the R graphics device for insertPlot() to capture.
  # The units for width/height will be 'in' and dpi=300
  
  # Plot 1
  print(p_cases)
  openxlsx::insertPlot(wb, sheet = sheet_safe_name, startRow = 1, startCol = 1,
                       width = plot_width_in, height = plot_height_in, units = "in", dpi = 300)
  
  # Plot 2
  print(p_deaths)
  openxlsx::insertPlot(wb, sheet = sheet_safe_name, startRow = 1, startCol = 1 + ceiling(plot_width_in) + 1,
                       width = plot_width_in, height = plot_height_in, units = "in", dpi = 300)
  
  # Plot 3
  print(p_attack_rate)
  openxlsx::insertPlot(wb, sheet = sheet_safe_name, startRow = ceiling(plot_height_in) + 2, startCol = 1,
                       width = plot_width_in, height = plot_height_in, units = "in", dpi = 300)
  
  # Plot 4
  print(p_fatality_rate)
  openxlsx::insertPlot(wb, sheet = sheet_safe_name, startRow = ceiling(plot_height_in) + 2, startCol = 1 + ceiling(plot_width_in) + 1,
                       width = plot_width_in, height = plot_height_in, units = "in", dpi = 300)
  
  # Optional: Add the raw data for the district on the same sheet for reference
  data_start_row <- (ceiling(plot_height_in) * 2) + 4 # 2 plots high, plus 4 row buffer
  writeData(wb, sheet = sheet_safe_name, x = district_data, startRow = data_start_row, startCol = 1,
            rowNames = FALSE, colNames = TRUE)
}

# Save the entire workbook
output_excel_filename <- paste0("Malaria_District_Reports_",sheet_name,"_with_Graphics.xlsx")
saveWorkbook(wb, output_excel_filename, overwrite = TRUE)

message(paste0("--- All district graphics and data saved to: ", output_excel_filename, " ---"))