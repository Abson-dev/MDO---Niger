# Load required libraries
library(tidyverse)
library(readxl)
library(readr) # For parse_number()
library(janitor) # For clean_names()
library(writexl) # For exporting to Excel

# --- Global Data Loading and Cleaning ---

# Define file paths and sheet names for each year
file_configs <- list(
  "2019" = list(file = "MDO_Niger Semaine 52 2019.xlsx", sheets = c("Palu", "Palu conf")),
  "2020" = list(file = "MDO_Niger Semaine 53 2020.xls", sheets = c("Palu", "Palu conf")),
  "2021" = list(file = "MDO_Niger Semaine 52 2021.xlsx", sheets = c("PALU", "PALU CONF")),
  "2022" = list(file = "MDO_NIGER 2022 S52.xlsx", sheets = c("Palu suspect", "Palu confirmé")),
  "2023" = list(file = "MDO 2023 S52.xls", sheets = c("Palu Susp", "Palu confirmé")),
  "2024" = list(file = "MDO 2024 S43.xls", sheets = c("Palu Susp", "Palu confirmé"))
)

# Global list to store errors during initial data loading
all_initial_load_errors_global <- list()

# Function to process a single sheet from a given file and year
process_sheet <- function(file_path, sheet_name, year) {
  # Read header rows with error handling
  header_df_raw <- tryCatch({
    read_excel(
      path = file_path,
      sheet = sheet_name,
      range = cell_rows(1:5),
      col_names = FALSE,
      col_types = "text"
    )
  }, error = function(e) {
    error_msg <- paste("Error reading header:", e$message)
    return(NULL) # Return NULL data frame on error
  })
  
  # --- NEW ERROR HANDLING: Check dimensions of header_df_raw ---
  if (is.null(header_df_raw)) {
    error_msg <- "Failed to read header, returned NULL."
    return(list(data = NULL, error = data.frame(Year = year, File = file_path, Sheet = sheet_name, Error_Message = error_msg)))
  }
  
  if (nrow(header_df_raw) < 3 || ncol(header_df_raw) < 5) {
    error_msg <- paste0("Header data has insufficient dimensions (rows: ", nrow(header_df_raw), ", cols: ", ncol(header_df_raw), "). Expected at least 3 rows and 5 columns.")
    return(list(data = NULL, error = data.frame(Year = year, File = file_path, Sheet = sheet_name, Error_Message = error_msg)))
  }
  # --- END NEW ERROR HANDLING ---
  
  fixed_cols_from_excel <- as.character(header_df_raw[3, 1:5])
  weekly_constructed_names <- c()
  current_semaine_col_in_excel <- 6
  
  while (current_semaine_col_in_excel <= ncol(header_df_raw) &&
         current_semaine_col_in_excel + 3 <= ncol(header_df_raw)) {
    semaine_header_val <- as.character(header_df_raw[3, current_semaine_col_in_excel])
    sub_headers_vals <- as.character(header_df_raw[5, current_semaine_col_in_excel:(current_semaine_col_in_excel + 3)])
    
    if (!is.na(semaine_header_val) && semaine_header_val != "" &&
        !all(is.na(sub_headers_vals) | sub_headers_vals == "")) {
      week_num <- str_extract(semaine_header_val, "\\d+")
      cleaned_sub_headers <- c("cas", "deces", "taux_d_attaque", "letalite")
      weekly_constructed_names <- c(weekly_constructed_names,
                                    paste0("Semaine ", week_num, "_", cleaned_sub_headers))
    } else {
      break
    }
    current_semaine_col_in_excel <- current_semaine_col_in_excel + 4
  }
  
  all_column_names_constructed <- c(fixed_cols_from_excel, weekly_constructed_names)
  all_column_names_constructed <- all_column_names_constructed[!is.na(all_column_names_constructed) & all_column_names_constructed != ""]
  
  # Read data block with error handling
  health_data_raw <- tryCatch({
    read_excel(
      path = file_path,
      sheet = sheet_name,
      skip = 5,
      col_names = FALSE,
      col_types = "text"
    )
  }, error = function(e) {
    error_msg <- paste("Error reading data block:", e$message)
    return(NULL) # Return NULL data frame on error
  })
  if (is.null(health_data_raw)) {
    return(list(data = NULL, error = data.frame(Year = year, File = file_path, Sheet = sheet_name, Error_Message = "Failed to read data block, returned NULL.")))
  }
  
  if (ncol(health_data_raw) < length(all_column_names_constructed)) {
    all_column_names_final <- all_column_names_constructed[1:ncol(health_data_raw)]
  } else if (ncol(health_data_raw) > length(all_column_names_constructed)) {
    num_extra_cols <- ncol(health_data_raw) - length(all_column_names_constructed)
    padded_names <- c(all_column_names_constructed, paste0("Unknown_", 1:num_extra_cols))
    all_column_names_final <- padded_names
  } else {
    all_column_names_final <- all_column_names_constructed
  }
  
  health_data <- health_data_raw %>% set_names(all_column_names_final)
  
  # Check for 'Region' column before filtering
  if (!"Region" %in% colnames(health_data)) {
    error_msg <- "'Region' column not found."
    return(list(data = NULL, error = data.frame(Year = year, File = file_path, Sheet = sheet_name, Error_Message = error_msg)))
  }
  
  # Clean and reshape data - wrap in tryCatch for robustness
  processed_data <- tryCatch({
    health_data_filtered <- health_data %>%
      filter(!is.na(Region) & Region != "Pays:" & Region != "Region" & !grepl("Total", Region, ignore.case = TRUE))
    
    health_data_clean_names <- health_data_filtered %>% janitor::clean_names()
    
    health_data_typed <- health_data_clean_names %>%
      mutate(
        population = readr::parse_number(population),
        across(
          starts_with("semaine_") &
            (contains("cas") | contains("deces") |
               contains("taux_d_attaque") | contains("letalite")),
          readr::parse_number
        )
      )
    
    health_data_long <- health_data_typed %>%
      pivot_longer(
        cols = starts_with("semaine_"),
        names_to = c("Week_Number_Raw", "Metric_Raw"),
        names_pattern = "semaine_(\\d+)_(.*)",
        values_to = "Value"
      ) %>%
      mutate(
        Week = as.integer(Week_Number_Raw),
        Metric = case_when(
          Metric_Raw == "cas" ~ "Cases",
          Metric_Raw == "deces" ~ "Deaths",
          Metric_Raw == "taux_d_attaque" ~ "Attack_Rate",
          Metric_Raw == "letalite" ~ "Fatality_Rate",
          TRUE ~ Metric_Raw
        )
      ) %>%
      select(
        region, district, population,
        week = Week,
        metric = Metric,
        value = Value
      )
    
    health_data_final_wide <- health_data_long %>%
      pivot_wider(
        names_from = metric,
        values_from = value,
        values_fn = first
      ) %>%
      arrange(region, district, week)
    
    health_data_final_wide %>%
      dplyr::select(region, district, population, week, Cases, Deaths, Attack_Rate, Fatality_Rate) %>%
      dplyr::filter(!is.na(week)) %>%
      dplyr::filter(!region %in% c("REGION")) %>%
      dplyr::mutate(population = round(population, 0)) %>%
      mutate(Sheet = sheet_name, Year = as.integer(year)) # Add Year column
  }, error = function(e) {
    error_msg <- paste("Error during data cleaning/reshaping:", e$message)
    return(list(data = NULL, error = data.frame(Year = year, File = file_path, Sheet = sheet_name, Error_Message = error_msg)))
  })
  
  if (is.data.frame(processed_data)) { # Check if processed_data is a data frame (success)
    return(list(data = processed_data, error = NULL))
  } else { # If it's not a data frame, it's an error list from the inner tryCatch
    return(processed_data)
  }
}

# Process all sheets and combine
all_health_data_list <- list()
for (year_str in names(file_configs)) {
  config <- file_configs[[year_str]]
  current_file_path <- config$file
  current_sheet_names <- config$sheets
  
  for (s_name in current_sheet_names) {
    result <- process_sheet(current_file_path, s_name, year_str)
    if (!is.null(result$data)) {
      all_health_data_list[[length(all_health_data_list) + 1]] <- result$data
    }
    if (!is.null(result$error)) {
      all_initial_load_errors_global[[length(all_initial_load_errors_global) + 1]] <- result$error
    }
  }
}
health_data_for_shiny <- bind_rows(all_health_data_list)

# Add a combined week-year column for plotting across years
health_data_for_shiny <- health_data_for_shiny %>%
  mutate(
    Year_Week = as.numeric(paste0(Year, sprintf("%02d", week))) # e.g., 201901, 201902
  )

# --- Combined Indicators Calculation ---

# Get the names for suspected and confirmed sheets
suspected_sheet_names <- c("Palu Susp", "Palu suspect", "Palu", "PALU")
confirmed_sheet_names <- c("Palu conf", "PALU CONF", "Palu confirmé")

# Calculate Combined Indicators with region
combined_indicators <- health_data_for_shiny %>%
  group_by(Year, week, region) %>%  # Group by Year, week, and region
summarise(
  # Sum of Cases from suspected sheets
  Suspected_Cases = sum(Cases[Sheet %in% suspected_sheet_names], na.rm = TRUE),
  # Sum of Cases from confirmed sheets
  Confirmed_Cases = sum(Cases[Sheet %in% confirmed_sheet_names], na.rm = TRUE),
  # Sum of Deaths from confirmed sheets
  Confirmed_Deaths = sum(Deaths[Sheet %in% confirmed_sheet_names], na.rm = TRUE),
  # Sum of Population (use unique population per district/region per week)
  Total_Population_Suspected = sum(unique(population[Sheet %in% suspected_sheet_names]), na.rm = TRUE),
  Total_Population_Confirmed = sum(unique(population[Sheet %in% confirmed_sheet_names]), na.rm = TRUE),
  .groups = 'drop'
) %>%
  mutate(
    # Calculate Positive Rate: (Confirmed Cases / Suspected Cases) * 100
    Positive_Rate = ifelse(Suspected_Cases > 0, (Confirmed_Cases / Suspected_Cases) * 100, 0),
    # Recalculate overall attack rate for confirmed cases
    Confirmed_Attack_Rate = ifelse(Total_Population_Confirmed > 0, (Confirmed_Cases / Total_Population_Confirmed) * 100000, 0),
    # Recalculate overall fatality rate for confirmed cases
    Confirmed_Fatality_Rate = ifelse(Confirmed_Cases > 0, (Confirmed_Deaths / Confirmed_Cases) * 100, 0)
  ) %>%
  arrange(Year, week, region)  # Sort by Year, week, and region

# --- Export to Excel ---
write_xlsx(combined_indicators, path = "Malaria_Combined_Indicators.xlsx")

# Optional: Alternative export using openxlsx
# library(openxlsx)
# write.xlsx(combined_indicators, file = "Malaria_Combined_Indicators.xlsx")

# Display the combined indicators
print(combined_indicators)