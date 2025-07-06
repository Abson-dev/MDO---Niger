# Load required libraries for Shiny App
library(shiny)
library(tidyverse)
library(readxl)
library(readr) # For parse_number()
library(janitor) # For clean_names()
library(DT) # For interactive data tables
library(plotly) # For interactive plots
library(bslib) # For modern design/theming
library(scales) # For comma formatting in plots
library(jsonlite)
library(viridis) # Added for scale_color_viridis_d

# --- Global Data Loading and Cleaning ---

# Define file paths and sheet names for each year
file_configs <- list(
  "2019" = list(file = "MDO_Niger Semaine 52 2019.xlsx", sheets = c("Palu", "Palu conf")),
  "2020" = list(file = "MDO_Niger Semaine 53 2020.xls", sheets = c("Palu", "Palu conf")), # Updated sheet names as per user's latest request
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

# Get unique years, districts, and metrics for dropdowns
unique_years <- sort(unique(health_data_for_shiny$Year))
unique_districts <- sort(unique(health_data_for_shiny$district[!is.na(health_data_for_shiny$district) & health_data_for_shiny$district != ""]))
unique_metrics_overall <- c("Cases", "Deaths") # Metrics for overall trend

# Check if data is empty after processing all files
if (nrow(health_data_for_shiny) == 0) {
  showNotification("No data loaded from any sheet. Please check the Excel files and sheet names.", type = "error")
  # Do not stop here, allow the app to run with an empty dataset and show errors
}

# --- Shiny UI ---
ui <- page_fluid(
  theme = bs_theme(
    version = 5,
    bootswatch = "flatly",
    primary = "#007bff",
    secondary = "#6c757d",
    success = "#28a745",
    info = "#17a2b8",
    warning = "#ffc107",
    danger = "#dc3545",
    base_font = font_google("Inter"),
    heading_font = font_google("Inter")
  ),
  
  title = "Malaria Surveillance Data Dashboard (Niger)", # More general title
  
  navbarPage(
    title = span(icon("chart-line"), "Malaria Dashboard"),
    collapsible = TRUE,
    
    tabPanel(
      title = span(icon("info-circle"), "Overview"),
      h2("Overall Malaria Trends and Summaries in Niger"),
      p("Select a year and a sheet to view overall trends for suspected or confirmed malaria cases."),
      layout_columns(
        col_widths = c(6, 6),
        selectInput("selected_year_overview", "Select Year:", choices = unique_years, selected = max(unique_years)),
        selectInput("sheet_select_overview", "Select Data Sheet:", choices = NULL) # Choices set dynamically
      ),
      hr(),
      layout_columns(
        col_widths = c(6, 6),
        card(
          card_header(h4(icon("globe"), "Key Overall Statistics")),
          layout_columns(
            col_widths = c(6, 6),
            div(h5("Total Cases:"), textOutput("total_cases_summary", inline = TRUE)),
            div(h5("Total Deaths:"), textOutput("total_deaths_summary", inline = TRUE)),
            div(h5("Overall Attack Rate:"), textOutput("overall_attack_rate_summary", inline = TRUE)),
            div(h5("Overall Fatality Rate:"), textOutput("overall_fatality_rate_summary", inline = TRUE))
          )
        ),
        card(
          card_header(h4(icon("chart-area"), "Overall Cases and Deaths Trend (Selected Sheet)")),
          selectInput("overall_metric_select", "Select Metric for Overall Trend:",
                      choices = unique_metrics_overall, selected = "Cases"),
          plotlyOutput("overall_cases_deaths_plot", height = "300px")
        )
      ),
      fluidRow(
        column(12,
               card(
                 card_header(h4(icon("chart-line"), "Confirmed Malaria Trend Across All Years")),
                 plotlyOutput("overall_confirmed_trend_plot", height = "400px")
               )
        )
      ),
      fluidRow( # New row for the Positive Rate plot
        column(12,
               card(
                 card_header(h4(icon("chart-line"), "Positive Rate Trend Across All Years (Confirmed / Suspected)")),
                 plotlyOutput("overall_positive_rate_plot", height = "400px")
               )
        )
      ),
      fluidRow( # New row for the Confirmed Deaths plot
        column(12,
               card(
                 card_header(h4(icon("chart-line"), "Confirmed Deaths Trend Across All Years")),
                 plotlyOutput("overall_confirmed_deaths_plot", height = "400px")
               )
        )
      ),
      fluidRow( # New row for the Suspected Cases plot
        column(12,
               card(
                 card_header(h4(icon("chart-line"), "Suspected Cases Trend Across All Years")),
                 plotlyOutput("overall_suspected_cases_plot", height = "400px")
               )
        )
      )
    ),
    
    tabPanel(
      title = span(icon("map-marked-alt"), "District Details"),
      sidebarLayout(
        sidebarPanel(
          h4(icon("filter"), "Select Sheet, Year, and District"),
          selectInput("selected_year_district", "Select Year:", choices = unique_years, selected = max(unique_years)),
          selectInput("sheet_select_district", "Select Data Sheet:", choices = NULL), # Choices set dynamically
          selectInput(
            inputId = "selected_district",
            label = "Choose a District:",
            choices = unique_districts,
            selected = unique_districts[1]
          ),
          hr(),
          h4(icon("table"), "District Summary"),
          tableOutput("district_summary_table")
        ),
        mainPanel(
          h3(textOutput("district_plots_title")),
          layout_columns(
            col_widths = c(6, 6),
            card(card_header("Cases Trend"), plotlyOutput("cases_plot", height = "300px")),
            card(card_header("Deaths Trend"), plotlyOutput("deaths_plot", height = "300px"))
          ),
          layout_columns(
            col_widths = c(6, 6),
            card(card_header("Attack Rate Trend"), plotlyOutput("attack_rate_plot", height = "300px")),
            card(card_header("Fatality Rate Trend"), plotlyOutput("fatality_rate_plot", height = "300px"))
          )
        )
      )
    ),
    
    tabPanel(
      title = span(icon("table"), "Sheet Comparison"), # Changed icon to table
      h2("Compare Malaria Data Across Years and Weeks"),
      p("Select the indicator you wish to view in the table below, showing values across years by week."),
      selectInput( # Changed from checkboxGroupInput to selectInput
        inputId = "selected_indicator_table_single",
        label = "Select Indicator for Table:",
        choices = c("Suspected Cases", "Confirmed Cases", "Confirmed Deaths"),
        selected = "Confirmed Cases"
      ),
      hr(),
      DT::dataTableOutput("combined_trends_table") # This will now show the pivoted table
    ),
    
    tabPanel(
      title = span(icon("database"), "Full Data Table"),
      h2("Complete Cleaned Malaria Surveillance Data"),
      p("Select a year and a sheet to explore the full dataset."),
      layout_columns(
        col_widths = c(6, 6),
        selectInput("selected_year_table", "Select Year:", choices = unique_years, selected = max(unique_years)),
        selectInput("sheet_select_table", "Select Data Sheet:", choices = NULL) # Choices set dynamically
      ),
      hr(),
      DT::dataTableOutput("full_data_table")
    ),
    
    # NEW TAB: Error Log
    tabPanel(
      title = span(icon("exclamation-triangle"), "Error Log"),
      h2("Data Importation Error Log"),
      p("This tab displays any errors encountered during the loading and initial processing of the Excel data."),
      hr(),
      DT::dataTableOutput("error_log_table")
    ),
    
    tabPanel(
      title = span(icon("info-circle"), "About"),
      fluidRow(
        column(12,
               h2("About This Dashboard"),
               p("This interactive dashboard provides a comprehensive view of suspected and confirmed malaria cases in Niger from 2019 to 2024, based on weekly epidemiological data."),
               h4("Data Source:"),
               p("Data is extracted from multiple Excel files, each corresponding to a specific year and containing both suspected and confirmed malaria case sheets."),
               h4("Methodology:"),
               tags$ul(
                 tags$li("Data from all specified sheets and years is loaded using `readxl`, handling complex multi-row headers programmatically."),
                 tags$li("Column names are standardized using `janitor::clean_names()`."),
                 tags$li("Numeric data is converted using `readr::parse_number()`."),
                 tags$li("Data is reshaped using `pivot_longer()` and `pivot_wider()`."),
                 tags$li("Missing weeks are filled with zero values to ensure continuous lines in plots, as observed in the provided example image."),
                 tags$li("Plots are generated using `ggplot2` and `plotly`."),
                 tags$li("Interactive tables use the `DT` package."),
                 tags$li("The dashboard uses `bslib` for a modern interface.")
               ),
               h4("Contact Information:"),
               p("For questions or feedback, contact Médecins Sans Frontières at jobs@msf.org.")
        )
      )
    )
  )
)

# --- Shiny Server Logic ---
server <- function(input, output, session) {
  
  # Reactive value to store errors encountered during data loading
  app_errors <- reactiveVal(tibble(
    Year = integer(),
    File = character(),
    Sheet = character(),
    Error_Message = character()
  ))
  
  # Populate app_errors with initial load errors when the app starts
  observeEvent(TRUE, { # This runs once on app startup
    if (length(all_initial_load_errors_global) > 0) {
      errors_df <- bind_rows(all_initial_load_errors_global)
      app_errors(errors_df)
    }
  }, once = TRUE) # Ensures this block runs only once
  
  # Custom Plot Theme
  custom_plot_theme <- function() {
    theme_minimal() +
      theme(
        plot.title = element_text(hjust = 0.5, face = "bold", size = 16, family = "Inter"),
        axis.title = element_text(size = 12, family = "Inter"),
        axis.text = element_text(size = 10, family = "Inter"),
        panel.grid.major.x = element_blank(),
        panel.grid.minor.x = element_blank(),
        panel.grid.major.y = element_line(color = "grey90", linetype = "solid"),
        panel.grid.minor.y = element_line(color = "grey95", linetype = "dashed"),
        plot.background = element_rect(fill = "#F8F9FA", color = NA),
        panel.background = element_rect(fill = "#FFFFFF", color = NA),
        legend.position = "none"
      )
  }
  
  # Color Palette
  metric_colors <- list(
    Cases = "#1F77B4", # Dark blue for cases
    Deaths = "#D62728", # Red for deaths
    Attack_Rate = "#2CA02C",
    Fatality_Rate = "#9467BD",
    Suspected = "#1F77B4",
    Confirmed = "#D62728"
  )
  
  # Helper function to determine default sheet selection
  get_default_sheet_selection <- function(available_sheets) {
    if ("Palu Susp" %in% available_sheets) {
      return("Palu Susp")
    } else if ("Palu suspect" %in% available_sheets) {
      return("Palu suspect")
    } else if ("Palu confirmé" %in% available_sheets) {
      return("Palu confirmé")
    } else if ("Palu conf" %in% available_sheets) {
      return("Palu conf")
    } else if ("PALU CONF" %in% available_sheets) {
      return("PALU CONF")
    } else if ("PALU" %in% available_sheets) {
      return("PALU")
    } else if ("Palu" %in% available_sheets) {
      return("Palu")
    } else if (length(available_sheets) > 0) {
      return(available_sheets[1]) # Fallback to the first available sheet
    } else {
      return(NULL)
    }
  }
  
  # Reactive for available sheets in Overview tab based on selected year
  available_sheets_overview <- reactive({
    req(input$selected_year_overview)
    health_data_for_shiny %>%
      filter(Year == input$selected_year_overview) %>%
      pull(Sheet) %>%
      unique() %>%
      sort()
  })
  
  # Update sheet choices when year changes in Overview tab
  observeEvent(input$selected_year_overview, {
    sheets <- available_sheets_overview()
    current_selected_sheet <- input$sheet_select_overview
    
    default_selected <- get_default_sheet_selection(sheets)
    
    # If the currently selected sheet is still in the new choices, keep it
    if (!is.null(current_selected_sheet) && current_selected_sheet %in% sheets) {
      default_selected <- current_selected_sheet
    }
    
    updateSelectInput(session, "sheet_select_overview",
                      choices = sheets,
                      selected = default_selected)
  })
  
  # Reactive for available sheets in District Details tab based on selected year
  available_sheets_district <- reactive({
    req(input$selected_year_district)
    health_data_for_shiny %>%
      filter(Year == input$selected_year_district) %>%
      pull(Sheet) %>%
      unique() %>%
      sort()
  })
  
  # Update sheet choices when year changes in District Details tab
  observeEvent(input$selected_year_district, {
    sheets <- available_sheets_district()
    current_selected_sheet <- input$sheet_select_district
    
    default_selected <- get_default_sheet_selection(sheets)
    
    if (!is.null(current_selected_sheet) && current_selected_sheet %in% sheets) {
      default_selected <- current_selected_sheet
    }
    
    updateSelectInput(session, "sheet_select_district",
                      choices = sheets,
                      selected = default_selected)
  })
  
  # Reactive for available sheets in Full Data Table tab based on selected year
  available_sheets_table <- reactive({
    req(input$selected_year_table)
    health_data_for_shiny %>%
      filter(Year == input$selected_year_table) %>%
      pull(Sheet) %>%
      unique() %>%
      sort()
  })
  
  # Update sheet choices when year changes in Full Data Table tab
  observeEvent(input$selected_year_table, {
    sheets <- available_sheets_table()
    current_selected_sheet <- input$sheet_select_table
    
    default_selected <- get_default_sheet_selection(sheets)
    
    if (!is.null(current_selected_sheet) && current_selected_sheet %in% sheets) {
      default_selected <- current_selected_sheet
    }
    
    updateSelectInput(session, "sheet_select_table",
                      choices = sheets,
                      selected = default_selected)
  })
  
  # Reactive Data for Overview Tab
  overall_summary_data <- reactive({
    # Ensure a sheet is selected before filtering
    req(input$sheet_select_overview)
    health_data_for_shiny %>%
      filter(Year == input$selected_year_overview, Sheet == input$sheet_select_overview) %>%
      mutate(week_num = as.numeric(week)) %>%
      filter(!is.na(week_num)) %>%
      group_by(week_num) %>%
      summarise(
        Total_Cases = sum(Cases, na.rm = TRUE),
        Total_Deaths = sum(Deaths, na.rm = TRUE),
        Total_Population = sum(unique(population), na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      # Fill missing weeks with 0 for continuous lines
      complete(week_num = full_seq(week_num, period = 1),
               fill = list(Total_Cases = 0, Total_Deaths = 0, Total_Population = 0)) %>% 
      mutate(
        Overall_Attack_Rate = ifelse(Total_Population > 0, (Total_Cases / Total_Population) * 100000, 0),
        Overall_Fatality_Rate = ifelse(Total_Cases > 0, (Total_Deaths / Total_Cases) * 100, 0)
      ) %>%
      arrange(week_num)
  })
  
  # Overview Tab - Summary Text Outputs
  output$total_cases_summary <- renderText({
    # Ensure data is available before summing
    if (nrow(overall_summary_data()) == 0) return("N/A")
    paste0(format(sum(overall_summary_data()$Total_Cases, na.rm = TRUE), big.mark = ","))
  })
  output$total_deaths_summary <- renderText({
    # Ensure data is available before summing
    if (nrow(overall_summary_data()) == 0) return("N/A")
    paste0(format(sum(overall_summary_data()$Total_Deaths, na.rm = TRUE), big.mark = ","))
  })
  output$overall_attack_rate_summary <- renderText({
    # Ensure data is available before calculating
    # Use req() for sheet_select_overview to prevent errors if it's NULL initially
    req(input$sheet_select_overview) 
    filtered_data <- health_data_for_shiny %>% filter(Year == input$selected_year_overview, Sheet == input$sheet_select_overview)
    if (nrow(filtered_data) == 0) return("N/A")
    total_pop_all_weeks <- sum(unique(filtered_data$population), na.rm = TRUE)
    total_cases_all_weeks <- sum(filtered_data$Cases, na.rm = TRUE)
    overall_attack_rate <- ifelse(total_pop_all_weeks > 0, (total_cases_all_weeks / total_pop_all_weeks) * 100000, 0)
    paste0(format(round(overall_attack_rate, 2), big.mark = ","), " per 100,000")
  })
  output$overall_fatality_rate_summary <- renderText({
    # Ensure data is available before calculating
    req(input$sheet_select_overview)
    filtered_data <- health_data_for_shiny %>% filter(Year == input$selected_year_overview, Sheet == input$sheet_select_overview)
    if (nrow(filtered_data) == 0) return("N/A")
    total_cases_all_weeks <- sum(filtered_data$Cases, na.rm = TRUE)
    total_deaths_all_weeks <- sum(filtered_data$Deaths, na.rm = TRUE)
    overall_fatality_rate <- ifelse(total_cases_all_weeks > 0, (total_deaths_all_weeks / total_cases_all_weeks) * 100, 0)
    paste0(format(round(overall_fatality_rate, 2), big.mark = ","), "%")
  })
  
  # Overview Tab - Overall Cases and Deaths Plot (for selected sheet)
  output$overall_cases_deaths_plot <- renderPlotly({
    plot_data <- overall_summary_data()
    selected_metric_overall <- input$overall_metric_select
    y_col_name <- paste0("Total_", selected_metric_overall)
    
    # Determine plot color based on selected metric
    plot_color <- if (selected_metric_overall == "Cases") {
      "#1F77B4" # Dark blue for cases, matching the example image
    } else if (selected_metric_overall == "Deaths") {
      "#D62728" # Red for deaths
    } else {
      "#6c757d" # Default color if neither (shouldn't happen with current choices)
    }
    
    # Determine y-axis label based on selected metric
    y_axis_label <- if (selected_metric_overall == "Cases") {
      "Number of Cases"
    } else if (selected_metric_overall == "Deaths") {
      "Number of Deaths"
    } else {
      selected_metric_overall # Fallback
    }
    
    if (nrow(plot_data) < 2 || all(is.na(plot_data[[y_col_name]]))) {
      p <- ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = paste("No overall data available for", selected_metric_overall, "plotting in", input$sheet_select_overview, "for", input$selected_year_overview), 
                 size = 6, color = "red", family = "Inter") +
        theme_void()
      plotly::ggplotly(p)
    } else {
      p <- ggplot(plot_data, aes(x = week_num, y = .data[[y_col_name]],
                                 text = paste("Week:", week_num, "<br>",
                                              selected_metric_overall, ":", round(.data[[y_col_name]], 2)),
                                 group = 1)) + 
        geom_point(color = plot_color, size = 1.0, alpha = 0.9) + # Points first, smaller
        geom_line(color = plot_color, size = 1.5) + # Line second, larger
        labs(
          title = paste("Overall Malaria", selected_metric_overall, "Trend (", input$sheet_select_overview, ", ", input$selected_year_overview, ")"),
          x = "Epidemiological Week",
          y = y_axis_label # Use determined y_axis_label
        ) +
        scale_x_continuous(breaks = seq(min(plot_data$week_num, na.rm = TRUE), 
                                        max(plot_data$week_num, na.rm = TRUE), by = 5),
                           labels = function(x) as.integer(x),
                           limits = c(1, max(plot_data$week_num, na.rm = TRUE))) +
        scale_y_continuous(labels = scales::comma, limits = c(0, max(plot_data[[y_col_name]], na.rm = TRUE) * 1.1)) +
        custom_plot_theme()
      
      plotly::ggplotly(p, tooltip = "text") %>%
        layout(
          hovermode = "x unified",
          showlegend = FALSE,
          margin = list(t = 50, b = 50, l = 50, r = 50),
          dragmode = "zoom",
          xaxis = list(fixedrange = FALSE),
          yaxis = list(fixedrange = FALSE),
          modebar = list(
            bgcolor= "transparent",
            color = "#000000",
            activecolor = plot_color # Use determined plot_color for modebar
          )
        ) %>%
        config(
          displaylogo = FALSE,
          modeBarButtonsToAdd = c("downloadImage", "zoom2d", "pan2d", "select2d", "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"),
          toImageButtonOptions = list(format = "png", filename = paste0("overall_", selected_metric_overall, "_", input$sheet_select_overview, "_", input$selected_year_overview))
        )
    }
  })
  
  # Reactive Data for Confirmed Malaria Trend Across All Years (Separate Lines per Year)
  overall_confirmed_trend_data <- reactive({
    confirmed_sheet_names <- c("Palu conf", "PALU CONF", "Palu confirmé")
    
    health_data_for_shiny %>%
      filter(Sheet %in% confirmed_sheet_names) %>%
      mutate(week_num = as.numeric(week)) %>%
      filter(!is.na(week_num)) %>%
      group_by(Year, week_num) %>% # Group by Year and week_num to keep years separate
      summarise(
        Total_Confirmed_Cases = sum(Cases, na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      # Ensure all weeks (1-53) are present for each year, filling with 0s
      complete(Year, week_num = 1:53, fill = list(Total_Confirmed_Cases = 0)) %>%
      arrange(Year, week_num)
  })
  
  # Plot for Confirmed Malaria Trend Across All Years (Multiple Lines)
  output$overall_confirmed_trend_plot <- renderPlotly({
    plot_data <- overall_confirmed_trend_data()
    
    if (nrow(plot_data) == 0 || all(is.na(plot_data$Total_Confirmed_Cases))) {
      p <- ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = "No confirmed malaria data available for plotting across all years.", 
                 size = 6, color = "red", family = "Inter") +
        theme_void()
      plotly::ggplotly(p)
    } else {
      p <- ggplot(plot_data, aes(x = week_num, y = Total_Confirmed_Cases, color = factor(Year), # Color by Year
                                 text = paste("Year:", Year, "<br>",
                                              "Week:", week_num, "<br>",
                                              "Confirmed Cases:", scales::comma(Total_Confirmed_Cases)))) +
        geom_point(size = 1.0, alpha = 0.9) + # Points first, smaller
        geom_line(size = 1.5) + # Line second, larger
        labs(
          title = "Confirmed Malaria Cases Trend Across All Years (2019-2024)",
          x = "Epidemiological Week",
          y = "Total Confirmed Cases",
          color = "Year" # Legend title
        ) +
        scale_x_continuous(breaks = seq(1, 53, by = 5), labels = function(x) as.integer(x)) +
        scale_y_continuous(labels = scales::comma, limits = c(0, max(plot_data$Total_Confirmed_Cases, na.rm = TRUE) * 1.1)) +
        scale_color_viridis_d(option = "D") + # A nice color palette for discrete years
        custom_plot_theme() +
        theme(legend.position = "right", legend.title = element_text(size = 12, face = "bold")) # Show legend
      
      plotly::ggplotly(p, tooltip = "text") %>%
        layout(
          hovermode = "x unified",
          margin = list(t = 50, b = 50, l = 50, r = 50),
          dragmode = "zoom",
          xaxis = list(fixedrange = FALSE),
          yaxis = list(fixedrange = FALSE),
          modebar = list(
            bgcolor = "transparent",
            color = "#000000",
            activecolor = "#1F77B4"
          )
        ) %>%
        config(
          displaylogo = FALSE,
          modeBarButtonsToAdd = c("downloadImage", "zoom2d", "pan2d", "select2d", "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"),
          toImageButtonOptions = list(format = "png", filename = "overall_confirmed_trend_all_years")
        )
    }
  })
  
  # Reactive Data for Positive Rate Trend Across All Years (Separate Lines per Year)
  overall_positive_rate_data <- reactive({
    suspected_sheet_names <- c("Palu Susp", "Palu suspect", "Palu", "PALU")
    confirmed_sheet_names <- c("Palu conf", "PALU CONF", "Palu confirmé")
    
    health_data_for_shiny %>%
      filter(Metric == "Cases") %>%
      mutate(
        Case_Type = case_when(
          Sheet %in% suspected_sheet_names ~ "Suspected",
          Sheet %in% confirmed_sheet_names ~ "Confirmed",
          TRUE ~ NA_character_
        )
      ) %>%
      filter(!is.na(Case_Type)) %>%
      mutate(week_num = as.numeric(week)) %>%
      filter(!is.na(week_num)) %>%
      group_by(Year, week_num, Case_Type) %>% # Group by Year, week, and case type
      summarise(Total_Cases = sum(value, na.rm = TRUE), .groups = 'drop') %>%
      
      # Pivot wider to get Suspected and Confirmed columns for each week and year
      pivot_wider(
        names_from = Case_Type,
        values_from = Total_Cases,
        values_fill = 0
      ) %>%
      # Complete weeks for continuity per year
      group_by(Year) %>%
      complete(week_num = 1:53, fill = list(Suspected = 0, Confirmed = 0)) %>%
      ungroup() %>%
      
      mutate(
        Positive_Rate = ifelse(Suspected > 0, (Confirmed / Suspected) * 100, NA_real_),
        Positive_Rate = ifelse(is.infinite(Positive_Rate) | is.nan(Positive_Rate), NA_real_, Positive_Rate)
      ) %>%
      arrange(Year, week_num)
  })
  
  # Plot for Positive Rate Trend Across All Years (Multiple Lines)
  output$overall_positive_rate_plot <- renderPlotly({
    plot_data <- overall_positive_rate_data()
    
    if (nrow(plot_data) == 0 || all(is.na(plot_data$Positive_Rate))) {
      p <- ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = "No positive rate data available for plotting across all years.", 
                 size = 6, color = "red", family = "Inter") +
        theme_void()
      plotly::ggplotly(p)
    } else {
      # Determine y-axis upper limit dynamically
      max_positive_rate <- max(plot_data$Positive_Rate, na.rm = TRUE)
      y_upper_limit <- ifelse(max_positive_rate == 0, 1, max_positive_rate * 1.1) # Set to 1 if max is 0, else 1.1 times max
      
      p <- ggplot(plot_data, aes(x = week_num, y = Positive_Rate, color = factor(Year), # Color by Year
                                 text = paste("Year:", Year, "<br>",
                                              "Week:", week_num, "<br>",
                                              "Positive Rate:", round(Positive_Rate, 2), "%"))) +
        geom_point(size = 1.0, alpha = 0.9) + # Points first, smaller
        geom_line(size = 1.5) + # Line second, larger
        labs(
          title = "Positive Rate Trend Across All Years (2019-2024)",
          x = "Epidemiological Week",
          y = "Positive Rate (%)",
          color = "Year" # Legend title
        ) +
        scale_x_continuous(breaks = seq(1, 53, by = 5), labels = function(x) as.integer(x)) +
        scale_y_continuous(labels = function(x) paste0(x, "%"), limits = c(0, y_upper_limit)) +
        scale_color_viridis_d(option = "C") + # Another nice color palette for discrete years
        custom_plot_theme() +
        theme(legend.position = "right", legend.title = element_text(size = 12, face = "bold")) # Show legend
      
      plotly::ggplotly(p, tooltip = "text") %>%
        layout(
          hovermode = "x unified",
          margin = list(t = 50, b = 50, l = 50, r = 50),
          dragmode = "zoom",
          xaxis = list(fixedrange = FALSE),
          yaxis = list(fixedrange = FALSE),
          modebar = list(
            bgcolor = "transparent",
            color = "#000000",
            activecolor = "#2CA02C"
          )
        ) %>%
        config(
          displaylogo = FALSE,
          modeBarButtonsToAdd = c("downloadImage", "zoom2d", "pan2d", "select2d", "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"),
          toImageButtonOptions = list(format = "png", filename = "overall_positive_rate_all_years")
        )
    }
  })
  
  # Reactive Data for Confirmed Deaths Trend Across All Years (Separate Lines per Year)
  overall_confirmed_deaths_trend_data <- reactive({
    confirmed_sheet_names <- c("Palu conf", "PALU CONF", "Palu confirmé")
    
    health_data_for_shiny %>%
      filter(Sheet %in% confirmed_sheet_names) %>%
      mutate(week_num = as.numeric(week)) %>%
      filter(!is.na(week_num)) %>%
      group_by(Year, week_num) %>% # Group by Year and week_num to keep years separate
      summarise(
        Total_Confirmed_Deaths = sum(Deaths, na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      # Ensure all weeks (1-53) are present for each year, filling with 0s
      complete(Year, week_num = 1:53, fill = list(Total_Confirmed_Deaths = 0)) %>%
      arrange(Year, week_num)
  })
  
  # Plot for Confirmed Deaths Trend Across All Years (Multiple Lines)
  output$overall_confirmed_deaths_plot <- renderPlotly({
    plot_data <- overall_confirmed_deaths_trend_data()
    
    if (nrow(plot_data) == 0 || all(is.na(plot_data$Total_Confirmed_Deaths))) {
      p <- ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = "No confirmed deaths data available for plotting across all years.", 
                 size = 6, color = "red", family = "Inter") +
        theme_void()
      plotly::ggplotly(p)
    } else {
      p <- ggplot(plot_data, aes(x = week_num, y = Total_Confirmed_Deaths, color = factor(Year), # Color by Year
                                 text = paste("Year:", Year, "<br>",
                                              "Week:", week_num, "<br>",
                                              "Confirmed Deaths:", scales::comma(Total_Confirmed_Deaths)))) +
        geom_point(size = 1.0, alpha = 0.9) + # Points first, smaller
        geom_line(size = 1.5) + # Line second, larger
        labs(
          title = "Confirmed Deaths Trend Across All Years (2019-2024)",
          x = "Epidemiological Week",
          y = "Total Confirmed Deaths",
          color = "Year" # Legend title
        ) +
        scale_x_continuous(breaks = seq(1, 53, by = 5), labels = function(x) as.integer(x)) +
        scale_y_continuous(labels = scales::comma, limits = c(0, max(plot_data$Total_Confirmed_Deaths, na.rm = TRUE) * 1.1)) +
        scale_color_viridis_d(option = "B") + # Using another distinct viridis option
        custom_plot_theme() +
        theme(legend.position = "right", legend.title = element_text(size = 12, face = "bold")) # Show legend
      
      plotly::ggplotly(p, tooltip = "text") %>%
        layout(
          hovermode = "x unified",
          margin = list(t = 50, b = 50, l = 50, r = 50),
          dragmode = "zoom",
          xaxis = list(fixedrange = FALSE),
          yaxis = list(fixedrange = FALSE),
          modebar = list(
            bgcolor = "transparent",
            color = "#000000",
            activecolor = "#D62728"
          )
        ) %>%
        config(
          displaylogo = FALSE,
          modeBarButtonsToAdd = c("downloadImage", "zoom2d", "pan2d", "select2d", "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"),
          toImageButtonOptions = list(format = "png", filename = "overall_confirmed_deaths_all_years")
        )
    }
  })
  
  # Reactive Data for Suspected Cases Trend Across All Years (Separate Lines per Year)
  overall_suspected_cases_trend_data <- reactive({
    suspected_sheet_names <- c("Palu Susp", "Palu suspect", "Palu", "PALU")
    
    health_data_for_shiny %>%
      filter(Sheet %in% suspected_sheet_names) %>%
      mutate(week_num = as.numeric(week)) %>%
      filter(!is.na(week_num)) %>%
      group_by(Year, week_num) %>% # Group by Year and week_num to keep years separate
      summarise(
        Total_Suspected_Cases = sum(Cases, na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      # Ensure all weeks (1-53) are present for each year, filling with 0s
      complete(Year, week_num = 1:53, fill = list(Total_Suspected_Cases = 0)) %>%
      arrange(Year, week_num)
  })
  
  # Plot for Suspected Cases Trend Across All Years (Multiple Lines)
  output$overall_suspected_cases_plot <- renderPlotly({
    plot_data <- overall_suspected_cases_trend_data()
    
    if (nrow(plot_data) == 0 || all(is.na(plot_data$Total_Suspected_Cases))) {
      p <- ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = "No suspected cases data available for plotting across all years.", 
                 size = 6, color = "red", family = "Inter") +
        theme_void()
      plotly::ggplotly(p)
    } else {
      p <- ggplot(plot_data, aes(x = week_num, y = Total_Suspected_Cases, color = factor(Year), # Color by Year
                                 text = paste("Year:", Year, "<br>",
                                              "Week:", week_num, "<br>",
                                              "Suspected Cases:", scales::comma(Total_Suspected_Cases)))) +
        geom_point(size = 1.0, alpha = 0.9) + # Points first, smaller
        geom_line(size = 1.5) + # Line second, larger
        labs(
          title = "Suspected Cases Trend Across All Years (2019-2024)",
          x = "Epidemiological Week",
          y = "Total Suspected Cases",
          color = "Year" # Legend title
        ) +
        scale_x_continuous(breaks = seq(1, 53, by = 5), labels = function(x) as.integer(x)) +
        scale_y_continuous(labels = scales::comma, limits = c(0, max(plot_data$Total_Suspected_Cases, na.rm = TRUE) * 1.1)) +
        scale_color_viridis_d(option = "A") + # Using another distinct viridis option for suspected cases
        custom_plot_theme() +
        theme(legend.position = "right", legend.title = element_text(size = 12, face = "bold")) # Show legend
      
      plotly::ggplotly(p, tooltip = "text") %>%
        layout(
          hovermode = "x unified",
          margin = list(t = 50, b = 50, l = 50, r = 50),
          dragmode = "zoom",
          xaxis = list(fixedrange = FALSE),
          yaxis = list(fixedrange = FALSE),
          modebar = list(
            bgcolor = "transparent",
            color = "#000000",
            activecolor = "#1F77B4"
          )
        ) %>%
        config(
          displaylogo = FALSE,
          modeBarButtonsToAdd = c("downloadImage", "zoom2d", "pan2d", "select2d", "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"),
          toImageButtonOptions = list(format = "png", filename = "overall_suspected_cases_all_years")
        )
    }
  })
  
  # Reactive Data for the new single indicator table in "Sheet Comparison" tab
  single_indicator_table_data <- reactive({
    req(input$selected_indicator_table_single) # Ensure an indicator is selected
    
    suspected_sheet_names <- c("Palu Susp", "Palu suspect", "Palu", "PALU")
    confirmed_sheet_names <- c("Palu conf", "PALU CONF", "Palu confirmé")
    
    processed_data <- health_data_for_shiny %>%
      filter(Metric %in% c("Cases", "Deaths")) %>%
      mutate(
        Indicator = case_when(
          Metric == "Cases" & Sheet %in% suspected_sheet_names ~ "Suspected Cases",
          Metric == "Cases" & Sheet %in% confirmed_sheet_names ~ "Confirmed Cases",
          Metric == "Deaths" & Sheet %in% confirmed_sheet_names ~ "Confirmed Deaths",
          TRUE ~ NA_character_
        )
      ) %>%
      filter(!is.na(Indicator)) %>%
      filter(Year >= 2021 & Year <= 2024) %>% # Filter for years 2021-2024
      filter(Indicator == input$selected_indicator_table_single) %>% # Filter by selected indicator
      mutate(week_num = as.numeric(week)) %>%
      filter(!is.na(week_num)) %>%
      group_by(Year, week_num) %>%
      summarise(Total_Value = sum(value, na.rm = TRUE), .groups = 'drop') %>%
      
      # Complete all weeks for the specified years
      complete(Year = 2021:2024, week_num = 1:53, fill = list(Total_Value = 0)) %>%
      
      # Pivot years to columns
      pivot_wider(
        names_from = Year,
        values_from = Total_Value,
        names_prefix = "Value in ", # Prefix for year columns
        values_fill = 0
      ) %>%
      arrange(week_num) %>%
      rename(Week = week_num) # Rename week_num to Week
    
    # Explicitly ensure year columns are numeric
    year_cols_prefix <- paste0("Value in ", 2021:2024)
    for (col_name in year_cols_prefix) {
      if (col_name %in% colnames(processed_data)) {
        processed_data[[col_name]] <- as.numeric(processed_data[[col_name]])
      }
    }
    
    return(processed_data)
  })
  
  # Render Combined Trends Table
  output$combined_trends_table <- DT::renderDataTable({
    table_data <- single_indicator_table_data()
    
    # Add a check for empty data
    if (nrow(table_data) == 0) {
      return(DT::datatable(
        tibble(Message = "No data available for the selected indicator in the years 2021-2024. Please check data files or select another indicator."),
        options = list(dom = 't'), # 't' for table only, no search/pagination
        rownames = FALSE
      ))
    }
    
    # Dynamically get the year columns for formatting
    year_cols <- as.character(2021:2024)
    cols_to_format <- intersect(paste0("Value in ", year_cols), colnames(table_data))
    
    dt_table <- DT::datatable(
      table_data,
      options = list(pageLength = 10, scrollX = TRUE,
                     dom = 'lfrtip',
                     language = list(search = "Search Data:")),
      rownames = FALSE,
      caption = htmltools::tags$caption(
        style = 'caption-side: top; text-align: center; color: black; font-size: 18px;',
        paste(input$selected_indicator_table_single, "Across Years (2021-2024) by Week")
      )
    )
    
    if (length(cols_to_format) > 0) {
      dt_table <- dt_table %>%
        formatCurrency(
          columns = cols_to_format,
          currency = "",
          mark = ",",
          before = FALSE,
          digits = 0
        )
    }
    
    return(dt_table)
  })
  
  
  # Reactive Data for District Details Tab
  filtered_district_data <- reactive({
    req(input$sheet_select_district, input$selected_year_district, input$selected_district)
    
    health_data_for_shiny %>%
      filter(Sheet == input$sheet_select_district, Year == input$selected_year_district, district == input$selected_district) %>%
      mutate(
        week_num = as.numeric(week),
        Cases = as.numeric(Cases),
        Deaths = as.numeric(Deaths),
        Attack_Rate = as.numeric(Attack_Rate),
        Fatality_Rate = as.numeric(Fatality_Rate)
      ) %>%
      filter(!is.na(week_num)) %>%
      arrange(week_num) %>%
      # Fill missing weeks with 0 for continuous lines
      group_by(district, region, Year, Sheet) %>% # Group by all identifying columns
      complete(week_num = full_seq(week_num, period = 1),
               fill = list(Cases = 0, Deaths = 0, Attack_Rate = 0, Fatality_Rate = 0)) %>% 
      ungroup() %>%
      # Re-fill population for newly completed rows if it's NA.
      group_by(district, Year) %>% # Fill population within each district-year group
      fill(population, .direction = "downup") %>% 
      ungroup() %>%
      # Recalculate rates after filling with 0s for cases/deaths and ensuring population
      mutate(
        Attack_Rate = ifelse(population > 0, (Cases / population) * 100000, 0),
        Fatality_Rate = ifelse(Cases > 0, (Deaths / Cases) * 100, 0)
      )
  })
  
  # Reactive Caption for District Summary Table
  output$district_summary_caption <- renderText({
    paste("Summary for", input$selected_district, "(", input$sheet_select_district, ", ", input$selected_year_district, ")")
  })
  
  # District Summary Table
  output$district_summary_table <- renderTable({
    summary_data <- filtered_district_data() %>%
      summarise(
        `Total Cases` = sum(Cases, na.rm = TRUE),
        `Total Deaths` = sum(Deaths, na.rm = TRUE),
        `Average Attack Rate` = mean(Attack_Rate, na.rm = TRUE),
        `Average Fatality Rate` = round(`Average Attack Rate`, 2),
        `Average Fatality Rate` = round(`Average Fatality Rate`, 2)
      ) %>%
      mutate(
        `Total Cases` = format(`Total Cases`, big.mark = ","),
        `Total Deaths` = format(`Total Deaths`, big.mark = ","),
        `Average Attack Rate` = round(`Average Attack Rate`, 2),
        `Average Fatality Rate` = round(`Average Fatality Rate`, 2)
      )
    summary_data
  }, caption = reactive({ input$district_summary_caption }), caption.placement = "top")
  
  # Plot title for District Details tab
  output$district_plots_title <- renderText({
    paste("Malaria Trends in", input$selected_district, "(", input$sheet_select_district, ", ", input$selected_year_district, ")")
  })
  
  # Render Plotly Function for District Details
  render_district_plotly <- function(metric_name, y_label, color) {
    renderPlotly({
      plot_data <- filtered_district_data()
      
      if (nrow(plot_data) < 2 || all(is.na(plot_data[[metric_name]]))) {
        p <- ggplot() +
          annotate("text", x = 0.5, y = 0.5, 
                   label = paste("No valid data for", metric_name, "in", input$selected_district, "(", input$sheet_select_district, ", ", input$selected_year_district, ")"), 
                   size = 5, color = "red", family = "Inter") +
          theme_void()
        plotly::ggplotly(p)
      } else {
        p <- ggplot(plot_data, aes(x = week_num, y = .data[[metric_name]],
                                   text = paste("Week:", week_num, "<br>",
                                                metric_name, ":", round(.data[[metric_name]], 2)),
                                   group = 1)) + 
          geom_point(color = color, size = 1.8, alpha = 0.9) + # Points first
          geom_line(color = color, size = 0.8) + # Line second
          labs(
            title = paste("Malaria", metric_name, "in", input$selected_district, "(", input$sheet_select_district, ", ", input$selected_year_district, ")"),
            x = "Epidemiological Week",
            y = y_label
          ) +
          scale_x_continuous(breaks = seq(min(plot_data$week_num, na.rm = TRUE), 
                                          max(plot_data$week_num, na.rm = TRUE), by = 5),
                             labels = function(x) as.integer(x),
                             limits = c(1, max(plot_data$week_num, na.rm = TRUE))) +
          scale_y_continuous(labels = scales::comma, limits = c(0, max(plot_data[[metric_name]], na.rm = TRUE) * 1.1)) +
          custom_plot_theme()
        
        plotly::ggplotly(p, tooltip = "text") %>%
          layout(
            hovermode = "x unified",
            showlegend = FALSE,
            margin = list(t = 50, b = 50, l = 50, r = 50),
            dragmode = "zoom",
            xaxis = list(fixedrange = FALSE),
            yaxis = list(fixedrange = FALSE),
            modebar = list(
              bgcolor = "transparent",
              color = "#000000",
              activecolor = color
            )
          ) %>%
          config(
            displaylogo = FALSE,
            modeBarButtonsToAdd = c("downloadImage", "zoom2d", "pan2d", "select2d", "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"),
            toImageButtonOptions = list(format = "png", filename = paste0(input$selected_district, "_", metric_name, "_", input$sheet_select_district, "_", input$selected_year_district))
          )
      }
    })
  }
  
  # Render Plots for District Details Tab
  output$cases_plot <- render_district_plotly("Cases", "Number of Cases", metric_colors$Cases)
  output$deaths_plot <- render_district_plotly("Deaths", "Number of Deaths", metric_colors$Deaths)
  output$attack_rate_plot <- render_district_plotly("Attack_Rate", "Attack Rate (per 100,000)", metric_colors$Attack_Rate)
  output$fatality_rate_plot <- render_district_plotly("Fatality_Rate", "Fatality Rate (%)", metric_colors$Fatality_Rate)
  
  # Data Table Tab
  output$full_data_table <- DT::renderDataTable({
    DT::datatable(
      health_data_for_shiny %>% filter(Year == input$selected_year_table, Sheet == input$sheet_select_table),
      options = list(pageLength = 10, scrollX = TRUE,
                     dom = 'lfrtip',
                     language = list(search = "Search Data:")),
      rownames = FALSE,
      caption = htmltools::tags$caption(
        style = 'caption-side: top; text-align: center; color: black; font-size: 18px;',
        paste('Complete Malaria Surveillance Data (', input$sheet_select_table, ', ', input$selected_year_table, ')')
      )
    )
  })
  
  # Error Log Table
  output$error_log_table <- DT::renderDataTable({
    DT::datatable(
      app_errors(),
      options = list(pageLength = 10, scrollX = TRUE,
                     dom = 'lfrtip',
                     language = list(search = "Filter Errors:")),
      rownames = FALSE,
      caption = htmltools::tags$caption(
        style = 'caption-side: top; text-align: center; color: black; font-size: 18px;',
        'Detailed Log of Data Importation Errors'
      )
    )
  })
  
}

# Run the Shiny app
shinyApp(ui = ui, server = server)
