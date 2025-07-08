# Load required libraries
library(tidyverse)
library(readxl)
library(readr) # For parse_number()
library(janitor) # For clean_names()
library(writexl) # For exporting to Excel
library(shiny)
library(shinydashboard)
library(shinythemes)
library(plotly)
library(DT)
library(shinyjs) # For UI enhancements
library(sf) # For handling shapefiles
library(leaflet) # For interactive maps
library(ggplot2) # For static map
library(RColorBrewer) # For better color palettes

# --- Global Data Loading and Cleaning ---

# Define file paths and sheet names for each year
# IMPORTANT: Ensure these files (MDO 2023 S52.xls, MDO 2024 S43.xls) are in your working directory,
# and the shapefile is in 'data/shp/' relative to your working directory.
file_configs <- list(
  "2019" = list(file = "MDO_Niger Semaine 52 2019.xlsx", sheets = c("Palu", "Palu conf")),
  "2020" = list(file = "MDO_Niger Semaine 53 2020.xls", sheets = c("Palu", "Palu conf")),
  "2021" = list(file = "MDO_Niger Semaine 52 2021.xlsx", sheets = c("PALU", "PALU CONF")),
  "2022" = list(file = "MDO_NIGER 2022 S52.xlsx", sheets = c("Palu suspect", "Palu confirmé")),
  "2023" = list(file = "MDO 2023 S52.xls", sheets = c("Palu Susp", "Palu confirmé")),
  "2024" = list(file = "MDO 2024 S43.xls", sheets = c("Palu Susp", "Palu confirmé"))
)

# Load Niger region shapefile with error handling
niger_shp <- tryCatch({
  shp <- st_read("data/shp/NER_admbnda_adm1_IGNN_20230720.shp")
  shp %>% mutate(nom = toupper(ADM1_FR)) # Standardize region names to uppercase
}, error = function(e) {
  message("Error loading shapefile: ", e$message)
  NULL
})

# Check if shapefile loaded successfully
if (is.null(niger_shp)) {
  stop("Failed to load data/shp/NER_admbnda_adm1_IGNN_20230720.shp. Please verify the file path and format.")
}

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
  
  # Check dimensions of header_df_raw
  if (is.null(header_df_raw)) {
    error_msg <- "Failed to read header, returned NULL."
    return(list(data = NULL, error = data.frame(Year = year, File = file_path, Sheet = sheet_name, Error_Message = error_msg)))
  }
  
  if (nrow(header_df_raw) < 3 || ncol(header_df_raw) < 5) {
    error_msg <- paste0("Header data has insufficient dimensions (rows: ", nrow(header_df_raw), ", cols: ", ncol(header_df_raw), "). Expected at least 3 rows and 5 columns.")
    return(list(data = NULL, error = data.frame(Year = year, File = file_path, Sheet = sheet_name, Error_Message = error_msg)))
  }
  
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
  
  # Clean and reshape data
  processed_data <- tryCatch({
    health_data_filtered <- health_data %>%
      filter(!is.na(Region) & Region != "Pays:" & Region != "Region" & !grepl("Total", Region, ignore.case = TRUE))
    
    health_data_clean_names <- health_data_filtered %>% janitor::clean_names()
    
    # Relax population filter to include non-numeric or missing values
    health_data_clean_names <- health_data_clean_names %>%
      mutate(population = ifelse(is.na(population) | population %in% c("POP", "NA", ""), NA, population))
    
    health_data_typed <- health_data_clean_names %>%
      mutate(
        population = readr::parse_number(population, na = c("", "NA")),
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
  
  if (is.data.frame(processed_data)) {
    return(list(data = processed_data, error = NULL))
  } else {
    return(processed_data)
  }
}

# Process all sheets and combine
all_health_data_list <- list()
all_initial_load_errors_global <- list() # Ensure this is initialized globally for error collection

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
    Year_Week = as.numeric(paste0(Year, sprintf("%02d", week))) # e.g., 202301, 202302
  )

# Calculate Combined Indicators with region
suspected_sheet_names <- c("Palu Susp", "Palu suspect", "Palu", "PALU")
confirmed_sheet_names <- c("Palu conf", "PALU CONF", "Palu confirmé")

combined_indicators <- health_data_for_shiny %>%
  group_by(Year, week, region) %>%
  summarise(
    Suspected_Cases = sum(Cases[Sheet %in% suspected_sheet_names], na.rm = TRUE),
    Confirmed_Cases = sum(Cases[Sheet %in% confirmed_sheet_names], na.rm = TRUE),
    Confirmed_Deaths = sum(Deaths[Sheet %in% confirmed_sheet_names], na.rm = TRUE),
    # Sum unique population values per region, across suspected/confirmed sheets
    Total_Population = sum(unique(population), na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  mutate(
    Positive_Rate = ifelse(Suspected_Cases > 0, (Confirmed_Cases / Suspected_Cases) * 100, 0),
    Confirmed_Attack_Rate = ifelse(Total_Population > 0, (Confirmed_Cases / Total_Population) * 100000, 0),
    Confirmed_Fatality_Rate = ifelse(Confirmed_Cases > 0, (Confirmed_Deaths / Confirmed_Cases) * 100, 0)
  ) %>%
  arrange(Year, week, region)

# Standardize region names in combined_indicators to match ADM1_FR (shapefile)
combined_indicators <- combined_indicators %>%
  mutate(region = toupper(region)) %>%
  mutate(region = recode(region,
                         "TILLABERI" = "TILLABÉRI",
                         "TAHOA" = "TAHOUA",
                         "DIFFA" = "DIFFA",
                         "DOSSO" = "DOSSO",
                         "NIAMEY" = "NIAMEY",
                         "AGADEZ" = "AGADEZ",
                         "MARADI" = "MARADI",
                         "ZINDER" = "ZINDER"
  ))

# Join combined_indicators with niger_shp to get ADM1_PCODE
combined_indicators <- combined_indicators %>%
  left_join(
    niger_shp %>% select(nom, ADM1_PCODE) %>% st_drop_geometry(), # Select ADM1_PCODE and drop geometry
    by = c("region" = "nom")
  ) %>%
  # Ensure ADM1_PCODE is at the beginning of the dataframe for better visibility
  select(Year, week, region, ADM1_PCODE, everything())

# Save to Excel (This line is for initial data processing, not the app download)
write_xlsx(combined_indicators, path = "Malaria_Combined_Indicators.xlsx")

# --- Shiny Application ---

# Define UI
ui <- dashboardPage(
  dashboardHeader(title = tags$a(href = "https://www.unicef.org/", # Changed link to UNICEF for relevance
                                 tags$img(src = "https://www.unicef.org/sites/default/files/styles/logo_mobile_retina/public/UNICEF_Logo_Blue.png?itok=xM9sQ1sC",
                                          height = "40px"),
                                 "Malaria Data Dashboard"),
                  titleWidth = 350), # Adjust title width for logo
  dashboardSidebar(
    sidebarMenu(
      id = "tabs", # Add an ID to the sidebarMenu for active tab management
      menuItem("Dashboard Overview", tabName = "dashboard", icon = icon("dashboard")),
      menuItem("Data Explorer", tabName = "data_explorer", icon = icon("table")),
      menuItem("How-to Guide", tabName = "how_to_guide", icon = icon("info-circle")), # New tab for how-to guide
      menuItem("Information", tabName = "info", icon = icon("info-circle")), # Existing information tab
      hr(), # Horizontal line for separation
      h4("Filters", style = "padding-left: 15px; color: #fff;"),
      selectInput("year", "Select Year:", choices = c("All", unique(combined_indicators$Year)), selected = "All"),
      selectInput("week", "Select Week:", choices = c("All", unique(combined_indicators$week)), selected = "All"),
      selectInput("region", "Select Region:", choices = c("All", sort(unique(combined_indicators$region))), selected = "All"),
      selectInput("metric", "Select Metric:",
                  choices = c(
                    "Suspected Cases" = "Suspected_Cases",
                    "Confirmed Cases" = "Confirmed_Cases",
                    "Confirmed Deaths" = "Confirmed_Deaths",
                    "Positive Rate (%)" = "Positive_Rate",
                    "Attack Rate (per 100,000)" = "Confirmed_Attack_Rate",
                    "Fatality Rate (%)" = "Fatality_Rate"
                  ),
                  selected = "Confirmed_Cases"),
      hr(),
      h4("Download Data", style = "padding-left: 15px; color: #fff;"),
      div(style = "padding-left: 15px;",
          downloadButton("download_csv", "Download CSV", class = "btn-success"),
          downloadButton("download_excel", "Download Excel", class = "btn-info")
      )
    )
  ),
  dashboardBody(
    useShinyjs(),
    tags$head(
      tags$link(rel = "stylesheet", type = "text/css", href = "https://fonts.googleapis.com/css?family=Open+Sans"),
      tags$style(HTML("
        body { font-family: 'Open Sans', sans-serif; }
        .content-wrapper { background-color: #f0f2f5; } /* Lighter background */
        .box {
          border-radius: 8px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15); /* More pronounced shadow */
          border-top: 3px solid #3c8dbc; /* Blue top border */
          margin-bottom: 20px;
        }
        .box-header .box-title {
          font-size: 18px;
          font-weight: bold;
          color: #333;
        }
        .main-header .logo {
          font-weight: bold;
          font-size: 20px;
          display: flex;
          align-items: center;
          justify-content: center;
        }
        .main-header .logo img {
          margin-right: 10px;
        }
        .sidebar-menu li a {
          font-size: 16px;
        }
        .sidebar-menu .active a {
          background-color: #007bff !important; /* Brighter active tab */
          color: #fff !important;
        }
        .download-btn {
          width: 90%; /* Make download buttons wider */
          margin-bottom: 10px;
        }
        .dataTables_wrapper .dataTables_filter input {
          width: 300px; /* Wider search bar for DT */
        }
        .shiny-notification {
          position: fixed;
          top: 10%;
          left: 50%;
          transform: translate(-50%, -50%);
          width: 400px;
          padding: 20px;
          border-radius: 10px;
          box-shadow: 0 5px 15px rgba(0,0,0,0.2);
          background-color: #dff0d8;
          color: #3c763d;
          border: 1px solid #d6e9c6;
          font-size: 16px;
        }
        .info-box {
          background-color: #fff;
          border-radius: 8px;
          box-shadow: 0 4px 8px rgba(0,0,0,0.1);
          padding: 20px;
          margin-bottom: 20px;
        }
      "))
    ),
    tabItems(
      tabItem(tabName = "dashboard",
              fluidRow(
                column(width = 6,
                       box(title = "Malaria Metrics Over Time", width = NULL, plotlyOutput("cases_plot", height = 400))
                ),
                column(width = 6,
                       box(title = "Malaria Metrics by Region (Static Map)", width = NULL, plotOutput("static_map_plot", height = 400))
                )
              ),
              fluidRow(
                box(title = "Malaria Indicators by Region (Interactive Map)", width = 12, leafletOutput("map_plot", height = 600))
              )
      ),
      tabItem(tabName = "data_explorer",
              fluidRow(
                box(title = "Detailed Malaria Data Table", width = 12, DTOutput("data_table"))
              )
      ),
      # New "How-to Guide" tab
      tabItem(tabName = "how_to_guide",
              fluidRow(
                column(width = 12,
                       div(class = "info-box",
                           h2("How to Use This Malaria Data Dashboard"),
                           p("This guide will walk you through the key features and functionalities of the dashboard, helping you to effectively explore and analyze malaria data in Niger."),
                           
                           h3("Phase 1: Getting Started - Dashboard Overview"),
                           tags$ul(
                             tags$li(tags$strong("Dashboard Overview Tab:"), " This is your starting point. It provides a visual summary of key malaria indicators. You'll see:"),
                             tags$ul(
                               tags$li(tags$strong("Malaria Metrics Over Time (Top Left):"), " A dynamic line plot showing trends for your selected metric across different years and weeks."),
                               tags$li(tags$strong("Malaria Metrics by Region (Static Map, Top Right):"), " A static map visualizing the average selected metric across regions of Niger. Regions with no data are shaded grey."),
                               tags$li(tags$strong("Malaria Indicators by Region (Interactive Map, Bottom):"), " An interactive map allowing you to zoom, pan, and click on regions for detailed metric values.")
                             ),
                             tags$li(tags$strong("Left Sidebar Filters:"), " Use the dropdown menus in the left sidebar to refine the data displayed in the plots and maps:"),
                             tags$ul(
                               tags$li(tags$strong("Select Year:"), " Choose a specific year or 'All' to see trends across all available years."),
                               tags$li(tags$strong("Select Week:"), " Filter data by a particular week within a year or view 'All' weeks."),
                               tags$li(tags$strong("Select Region:"), " Focus on a specific administrative region of Niger or view data for 'All' regions."),
                               tags$li(tags$strong("Select Metric:"), " Switch between different malaria indicators like 'Confirmed Cases', 'Fatality Rate', 'Attack Rate', etc. The plots and maps will update dynamically based on your selection.")
                             )
                           ),
                           
                           h3("Phase 2: Deep Dive - Data Explorer"),
                           tags$ul(
                             tags$li(tags$strong("Data Explorer Tab:"), " Navigate to this tab to view the aggregated malaria data in a detailed, interactive table."),
                             tags$ul(
                               tags$li(tags$strong("Table Functionality:"), " You can:"),
                               tags$ul(
                                 tags$li("Use the search bar at the top right to find specific entries."),
                                 tags$li("Sort columns by clicking on their headers (click once for ascending, twice for descending)."),
                                 tags$li("Filter data for specific values in each column using the input boxes directly below the column headers."),
                                 tags$li("Adjust the number of rows displayed per page using the 'Show entries' dropdown.")
                               )
                             ),
                             tags$li("The table includes all calculated metrics along with `ADM1_PCODE` (Administrative Division 1 P-Code) for each region, providing unique identifiers.")
                           ),
                           
                           h3("Phase 3: Understanding the Visualizations"),
                           tags$ul(
                             tags$li(tags$strong("Time Series Plot:"), " Observe trends over time. When 'All' years are selected, each line represents a different year, allowing for easy comparison of malaria patterns across years. Hover over data points to see specific values."),
                             tags$li(tags$strong("Static Map:"), " Provides a quick, visual summary of metric distribution across regions for the currently selected filters. Darker shades generally indicate higher values of the selected metric."),
                             tags$li(tags$strong("Interactive Map:"), " Enables granular exploration."),
                             tags$ul(
                               tags$li(tags$strong("Zoom/Pan:"), " Use your mouse wheel or the +/- controls to zoom in and out, and click-and-drag to pan the map."),
                               tags$li(tags$strong("Click on Regions:"), " Clicking on any region will open a popup displaying the region's name, its `ADM1_PCODE`, and the value of the currently selected malaria metric for that region. This is particularly useful for identifying specific regional data.")
                             )
                           ),
                           
                           h3("Phase 4: Exporting Your Analysis"),
                           tags$ul(
                             tags$li(tags$strong("Download Data Section (Left Sidebar):"), " Once you have filtered the data to your interest, you can download it for further analysis or reporting."),
                             tags$ul(
                               tags$li(tags$strong("Download CSV:"), " Downloads the currently filtered data as a Comma Separated Values (.csv) file, suitable for spreadsheets and data analysis tools."),
                               tags$li(tags$strong("Download Excel:"), " Downloads the currently filtered data as an Excel (.xlsx) spreadsheet, preserving data types and making it easy to use.")
                             )
                           ),
                           p("We hope this guide helps you make the most of this Malaria Data Dashboard. For any further questions, please refer to the 'Information' tab or contact us.")
                       )
                )
              )
      ),
      tabItem(tabName = "info",
              fluidRow(
                column(width = 12,
                       div(class = "info-box",
                           h3("About This Dashboard"),
                           p("This dashboard provides an interactive visualization of malaria indicators in Niger, leveraging data from MDO 2023 and MDO 2024 reports."),
                           p("Key features include:"),
                           tags$ul(
                             tags$li("Time-series plots for various malaria metrics."),
                             tags$li("Static and interactive maps to visualize regional distribution of malaria indicators."),
                             tags$li("A detailed data table for exploring the raw and aggregated data."),
                             tags$li("Download options for filtered data in CSV and Excel formats."),
                             tags$li("Comprehensive 'How-to Guide' for navigating the application.")
                           ),
                           h3("Data Sources"),
                           p("The data used in this dashboard is derived from the following files:"),
                           tags$ul(
                             tags$li("MDO 2023 S52.xls"),
                             tags$li("MDO 2024 S43.xls")
                           ),
                           p("Geographic data (shapefile) for Niger's regions is from ",
                             tags$a(href = "https://data.humdata.org/dataset/niger-admin-level-1-boundaries", target = "_blank", "Humanitarian Data Exchange (HDX).")),
                           h3("Metrics Explained"),
                           tags$ul(
                             tags$li("Suspected Cases: Total number of individuals suspected of having malaria."),
                             tags$li("Confirmed Cases: Total number of laboratory-confirmed malaria cases."),
                             tags$li("Confirmed Deaths: Total number of deaths due to confirmed malaria cases."),
                             tags$li("Positive Rate (%): (Confirmed Cases / Suspected Cases) * 100. Indicates the proportion of suspected cases that are confirmed."),
                             tags$li("Attack Rate (per 100,000): (Confirmed Cases / Total Population) * 100,000. Represents the incidence of confirmed malaria cases per 100,000 population."),
                             tags$li("Fatality Rate (%): (Confirmed Deaths / Confirmed Cases) * 100. Indicates the proportion of confirmed cases that result in death.")
                           ),
                           h3("Contact"),
                           p("For questions or feedback, please contact [Your Name/Organization Name] at [Your Email Address].")
                       )
                )
              )
      )
    )
  ),
  skin = "blue"
)

# Define Server
server <- function(input, output, session) {
  
  # Define the metric display names once for consistent use
  metric_display_names <- c(
    "Suspected_Cases" = "Suspected Cases",
    "Confirmed_Cases" = "Confirmed Cases",
    "Confirmed_Deaths" = "Confirmed Deaths",
    "Positive_Rate" = "Positive Rate (%)",
    "Confirmed_Attack_Rate" = "Attack Rate (per 100,000)",
    "Fatality_Rate" = "Fatality Rate (%)"
  )
  
  # Update week choices based on selected year
  observeEvent(input$year, {
    if (input$year == "All") {
      updateSelectInput(session, "week", choices = c("All", sort(unique(combined_indicators$week))), selected = "All")
    } else {
      weeks_for_year <- combined_indicators %>%
        filter(Year == as.integer(input$year)) %>%
        pull(week) %>%
        unique() %>%
        sort()
      updateSelectInput(session, "week", choices = c("All", weeks_for_year), selected = "All")
    }
  })
  
  # Reactive filtered data for maps (ignores region filter for broader overview)
  filtered_data_for_maps <- reactive({
    data <- combined_indicators
    if (input$year != "All") {
      data <- data %>% filter(Year == as.integer(input$year))
    }
    if (input$week != "All") {
      data <- data %>% filter(week == as.integer(input$week))
    }
    data
  })
  
  # Reactive filtered data for plot and table (includes region filter)
  filtered_data_for_plot_table <- reactive({
    data <- combined_indicators
    if (input$year != "All") {
      data <- data %>% filter(Year == as.integer(input$year))
    }
    if (input$week != "All") {
      data <- data %>% filter(week == as.integer(input$week))
    }
    if (input$region != "All") {
      data <- data %>% filter(region == input$region)
    }
    data
  })
  
  # Line plot for selected metric over time by week
  output$cases_plot <- renderPlotly({
    data <- filtered_data_for_plot_table()
    
    # Dynamic plot title
    plot_title <- paste("Trend of", metric_display_names[input$metric], "over Time")
    if (input$region != "All") {
      plot_title <- paste(plot_title, "in", input$region)
    }
    if (input$year != "All") {
      plot_title <- paste(plot_title, "for", input$year)
    }
    
    # Ensure consistent ordering for plotting
    plot_data <- data %>% arrange(Year, week)
    
    p <- plot_ly(plot_data, x = ~week, y = ~get(input$metric), type = 'scatter', mode = 'lines',
                 color = ~as.factor(Year), # Color by year
                 colors = RColorBrewer::brewer.pal(max(3, length(unique(plot_data$Year))), "Set1"), # Use Brewer palette
                 hoverinfo = 'text',
                 text = ~paste('Year: ', Year, '<br>Week: ', week, '<br>',
                               metric_display_names[input$metric], ': ', round(get(input$metric), 2))) %>%
      layout(
        title = list(text = plot_title, font = list(size = 16)),
        xaxis = list(title = "Week", type = 'category', # Set x-axis title to "Week"
                     tickvals = unique(plot_data$week)[seq(1, length(unique(plot_data$week)), by = 4)], # Show fewer ticks for readability
                     ticktext = unique(paste0("W", plot_data$week))[seq(1, length(unique(plot_data$week)), by = 4)]),
        yaxis = list(title = metric_display_names[input$metric]),
        legend = list(orientation = "h", x = 0, y = -0.2),
        margin = list(b = 100) # Adjust bottom margin for legend
      )
    p
  })
  
  # Static map for selected metric by region for selected year/week
  output$static_map_plot <- renderPlot({
    data <- filtered_data_for_maps()
    
    # Aggregate data by region based on current filters
    agg_data <- data %>%
      group_by(region) %>%
      summarise(
        Metric_Value = mean(get(input$metric), na.rm = TRUE), # Use mean for aggregation
        .groups = 'drop'
      ) %>%
      mutate(region = toupper(region)) # Ensure region names are uppercase for joining
    
    # Join with shapefile
    map_data <- niger_shp %>%
      left_join(agg_data, by = c("nom" = "region"))
    
    # Dynamic map title
    map_title <- paste("Average", metric_display_names[input$metric], "by Region")
    if (input$year != "All") {
      map_title <- paste(map_title, "in", input$year)
    }
    if (input$week != "All") {
      map_title <- paste(map_title, "Week", input$week)
    }
    
    # Create static map with ggplot2
    ggplot(data = map_data) +
      geom_sf(aes(fill = Metric_Value), color = "white", lwd = 0.5) + # Add white border for regions
      scale_fill_viridis_c( # Use viridis for better colorblind-friendliness
        option = "plasma",
        na.value = "#CCCCCC", # Light grey for no data
        name = metric_display_names[input$metric]
      ) +
      theme_minimal() +
      labs(
        title = map_title,
        caption = "Data source: MDO 2023-2024. Regions with no data are grey."
      ) +
      theme(
        plot.title = element_text(hjust = 0.5, size = 16, face = "bold"),
        legend.position = "right",
        legend.title = element_text(size = 12),
        legend.text = element_text(size = 10),
        panel.grid.major = element_blank(), # Remove grid lines
        panel.grid.minor = element_blank(),
        axis.text = element_blank(), # Remove axis text
        axis.title = element_blank() # Remove axis titles
      )
  }, bg = "transparent") # Set background to transparent for integration with dashboard
  
  # Interactive map plot for selected metric by region
  output$map_plot <- renderLeaflet({
    data <- filtered_data_for_maps() %>%
      group_by(region) %>%
      summarise(
        Metric_Value = mean(get(input$metric), na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      mutate(region = toupper(region))
    
    # Join with shapefile
    map_data <- niger_shp %>%
      left_join(data, by = c("nom" = "region"))
    
    # Ensure valid geometry for leaflet
    map_data <- st_make_valid(map_data)
    
    # Define color palette for the map
    pal <- colorNumeric(palette = "YlGnBu", domain = map_data$Metric_Value, na.color = "#CCCCCC") # Use YlGnBu for a different feel
    
    # Create popup content
    popup_content <- paste0(
      "<b>Region:</b> ", map_data$nom, "<br>",
      "<b>ADM1_PCODE:</b> ", map_data$ADM1_PCODE, "<br>", # Include ADM1_PCODE in popup
      "<b>", metric_display_names[input$metric], ":</b> ",
      ifelse(is.na(map_data$Metric_Value), "No data", round(map_data$Metric_Value, 2))
    )
    
    # Create leaflet map
    leaflet(map_data) %>%
      addProviderTiles(providers$CartoDB.Positron) %>% # Different base map for better aesthetics
      addPolygons(
        fillColor = ~pal(Metric_Value),
        fillOpacity = 0.8,
        color = "#444444",
        weight = 1,
        popup = popup_content,
        highlightOptions = highlightOptions(
          color = "white", weight = 2,
          bringToFront = TRUE
        )
      ) %>%
      addLegend(
        pal = pal,
        values = ~Metric_Value,
        title = metric_display_names[input$metric],
        position = "bottomright",
        labFormat = labelFormat(digits = 2)
      ) %>%
      setView(lng = 8, lat = 17, zoom = 6) # Center map on Niger
  })
  
  # Data table
  output$data_table <- renderDT({
    data <- filtered_data_for_plot_table() %>%
      rename(
        "Year" = Year,
        "Week" = week,
        "Region" = region,
        "ADM1 PCODE" = ADM1_PCODE, # Renamed for display
        "Suspected Cases" = Suspected_Cases,
        "Confirmed Cases" = Confirmed_Cases,
        "Confirmed Deaths" = Confirmed_Deaths,
        "Total Population" = Total_Population, # Renamed for clarity
        "Positive Rate (%)" = Positive_Rate,
        "Attack Rate (per 100,000)" = Confirmed_Attack_Rate,
        "Fatality Rate (%)" = Fatality_Rate
      ) %>%
      # Ensure ADM1 PCODE is explicitly included and ordered as desired
      select(Year, Week, `ADM1 PCODE`, Region, district, `Suspected Cases`, `Confirmed Cases`, `Confirmed Deaths`,
             `Total Population`, `Positive Rate (%)`, `Attack Rate (per 100,000)`, `Fatality Rate (%)`) # Reorder columns
    
    datatable(data,
              options = list(
                pageLength = 10,
                autoWidth = TRUE,
                scrollX = TRUE, # Enable horizontal scrolling for many columns
                dom = 'lfrtip', # Show length, filter, table, info, pagination
                lengthMenu = c(10, 25, 50, 100) # Options for number of rows per page
              ),
              rownames = FALSE,
              filter = 'top' # Add filters on top of each column
    ) %>%
      formatRound(columns = c("Positive Rate (%)", "Attack Rate (per 100,000)", "Fatality Rate (%)"), digits = 2)
  })
  
  # Download CSV
  output$download_csv <- downloadHandler(
    filename = function() {
      # Sanitize inputs for filename to prevent issues with special characters
      sanitized_year <- gsub("[^a-zA-Z0-9.-]", "", input$year)
      sanitized_week <- gsub("[^a-zA-Z0-9.-]", "", input$week)
      sanitized_region <- gsub("[^a-zA-Z0-9.-]", "", input$region)
      paste0("Malaria_Indicators_", sanitized_year, "_W", sanitized_week, "_", sanitized_region, ".csv")
    },
    content = function(file) {
      data_to_download <- filtered_data_for_plot_table() %>%
        select(Year, week, ADM1_PCODE, region, district, Suspected_Cases, Confirmed_Cases, Confirmed_Deaths,
               Total_Population, Positive_Rate, Confirmed_Attack_Rate, Fatality_Rate)
      write_csv(data_to_download, file)
    }
  )
  
  # Download Excel
  output$download_excel <- downloadHandler(
    filename = function() {
      # Sanitize inputs for filename to prevent issues with special characters
      sanitized_year <- gsub("[^a-zA-Z0-9.-]", "", input$year)
      sanitized_week <- gsub("[^a-zA-Z0-9.-]", "", input$week)
      sanitized_region <- gsub("[^a-zA-Z0-9.-]", "", input$region)
      paste0("Malaria_Indicators_", sanitized_year, "_W", sanitized_week, "_", sanitized_region, ".xlsx")
    },
    content = function(file) {
      data_to_download <- filtered_data_for_plot_table() %>%
        select(Year, week, ADM1_PCODE, region, district, Suspected_Cases, Confirmed_Cases, Confirmed_Deaths,
               Total_Population, Positive_Rate, Confirmed_Attack_Rate, Fatality_Rate)
      write_xlsx(data_to_download, file)
    }
  )
  
  # Display initial load errors as a Shiny notification
  observe({
    if (length(all_initial_load_errors_global) > 0) {
      error_messages <- map_chr(all_initial_load_errors_global, ~ paste0("- File: ", .$File, ", Sheet: ", .$Sheet, ", Error: ", .$Error_Message))
      showNotification(
        ui = div(
          h4("Data Loading Warnings/Errors"),
          tags$ul(lapply(error_messages, tags$li))
        ),
        duration = NULL, # Persist until dismissed
        type = "warning",
        closeButton = TRUE
      )
    }
  })
}

# Run the application
shinyApp(ui = ui, server = server)