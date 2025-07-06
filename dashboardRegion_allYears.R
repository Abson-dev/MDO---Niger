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
library(shiny.i18n) # For multilingual support
library(jsonlite) # For JSON handling

# Define translations (English and French)
translations <- list(
  en = list(
    title = "Malaria Data Dashboard",
    year = "Year",
    week = "Week",
    region = "Region",
    metric = "Metric",
    all = "All",
    select_language = "Select Language",
    plot_title_cases = "Malaria Cases and Deaths Over Time",
    plot_title_metrics = "Malaria Metrics by Region",
    table_title = "Malaria Indicators Data",
    download_csv = "Download CSV",
    download_excel = "Download Excel",
    suspected_cases = "Suspected Cases",
    confirmed_cases = "Confirmed Cases",
    confirmed_deaths = "Confirmed Deaths",
    positive_rate = "Positive Rate (%)",
    attack_rate = "Attack Rate (per 100,000)",
    fatality_rate = "Fatality Rate (%)"
  ),
  fr = list(
    title = "Tableau de Bord des Données sur le Paludisme",
    year = "Année",
    week = "Semaine",
    region = "Région",
    metric = "Métrique",
    all = "Tous",
    select_language = "Sélectionner la Langue",
    plot_title_cases = "Cas et Décès de Paludisme au Fil du Temps",
    plot_title_metrics = "Métriques du Paludisme par Région",
    table_title = "Données des Indicateurs de Paludisme",
    download_csv = "Télécharger CSV",
    download_excel = "Télécharger Excel",
    suspected_cases = "Cas Suspectés",
    confirmed_cases = "Cas Confirmés",
    confirmed_deaths = "Décès Confirmés",
    positive_rate = "Taux de Positivité (%)",
    attack_rate = "Taux d'Attaque (par 100 000)",
    fatality_rate = "Taux de Létalité (%)"
  )
)

# Write translations to a temporary JSON file
temp_json_file <- tempfile(fileext = ".json")
write_json(translations, temp_json_file, auto_unbox = TRUE)

# Initialize i18n with the temporary JSON file
i18n <- Translator$new(translation_json_path = temp_json_file)

# --- Global Data Loading and Cleaning ---

# Define file paths and sheet names for each year
file_configs <- list(
  # "2019" = list(file = "MDO_Niger Semaine 52 2019.xlsx", sheets = c("Palu", "Palu conf")),
  # "2020" = list(file = "MDO_Niger Semaine 53 2020.xls", sheets = c("Palu", "Palu conf")),
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
  
  if (is.data.frame(processed_data)) {
    return(list(data = processed_data, error = NULL))
  } else {
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

# Calculate Combined Indicators with region
suspected_sheet_names <- c("Palu Susp", "Palu suspect", "Palu", "PALU")
confirmed_sheet_names <- c("Palu conf", "PALU CONF", "Palu confirmé")

combined_indicators <- health_data_for_shiny %>%
  group_by(Year, week, region) %>%
  summarise(
    Suspected_Cases = sum(Cases[Sheet %in% suspected_sheet_names], na.rm = TRUE),
    Confirmed_Cases = sum(Cases[Sheet %in% confirmed_sheet_names], na.rm = TRUE),
    Confirmed_Deaths = sum(Deaths[Sheet %in% confirmed_sheet_names], na.rm = TRUE),
    Total_Population_Suspected = sum(unique(population[Sheet %in% suspected_sheet_names]), na.rm = TRUE),
    Total_Population_Confirmed = sum(unique(population[Sheet %in% confirmed_sheet_names]), na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  mutate(
    Positive_Rate = ifelse(Suspected_Cases > 0, (Confirmed_Cases / Suspected_Cases) * 100, 0),
    Confirmed_Attack_Rate = ifelse(Total_Population_Confirmed > 0, (Confirmed_Cases / Total_Population_Confirmed) * 100000, 0),
    Confirmed_Fatality_Rate = ifelse(Confirmed_Cases > 0, (Confirmed_Deaths / Confirmed_Cases) * 100, 0)
  ) %>%
  arrange(Year, week, region)

# Save to Excel
write_xlsx(combined_indicators, path = "Malaria_Combined_Indicators.xlsx")

# --- Shiny Application ---

# Define UI
ui <- dashboardPage(
  dashboardHeader(title = uiOutput("app_title")),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Dashboard", tabName = "dashboard", icon = icon("dashboard")),
      selectInput("language", "Select Language / Sélectionner la Langue", choices = c("English" = "en", "French" = "fr")),
      selectInput("year", uiOutput("year_label"), choices = c("All" = "All", unique(combined_indicators$Year)), selected = "All"),
      selectInput("week", uiOutput("week_label"), choices = c("All" = "All", unique(combined_indicators$week)), selected = "All"),
      selectInput("region", uiOutput("region_label"), choices = c("All" = "All", unique(combined_indicators$region)), selected = "All"),
      selectInput("metric", uiOutput("metric_label"), 
                  choices = c(
                    "Suspected Cases / Cas Suspectés" = "Suspected_Cases",
                    "Confirmed Cases / Cas Confirmés" = "Confirmed_Cases",
                    "Confirmed Deaths / Décès Confirmés" = "Confirmed_Deaths",
                    "Positive Rate (%) / Taux de Positivité (%)" = "Positive_Rate",
                    "Attack Rate (per 100,000) / Taux d'Attaque (par 100 000)" = "Confirmed_Attack_Rate",
                    "Fatality Rate (%) / Taux de Létalité (%)" = "Confirmed_Fatality_Rate"
                  ),
                  selected = "Confirmed_Cases"),
      downloadButton("download_csv", uiOutput("download_csv_label")),
      downloadButton("download_excel", uiOutput("download_excel_label"))
    )
  ),
  dashboardBody(
    shinyjs::useShinyjs(),
    tags$head(
      tags$style(HTML("
        .content-wrapper { background-color: #f4f6f9; }
        .box { border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
        .main-header .logo { font-weight: bold; font-size: 20px; }
        .download-btn { margin-top: 10px; }
      "))
    ),
    tabItems(
      tabItem(tabName = "dashboard",
              fluidRow(
                box(title = uiOutput("plot_title_cases"), width = 6, plotlyOutput("cases_plot")),
                box(title = uiOutput("plot_title_metrics"), width = 6, plotlyOutput("metrics_plot"))
              ),
              fluidRow(
                box(title = uiOutput("table_title"), width = 12, DTOutput("data_table"))
              )
      )
    )
  ),
  skin = "blue"
)

# Define Server
server <- function(input, output, session) {
  # Reactive for language
  observeEvent(input$language, {
    i18n$set_translation_language(input$language)
  })
  
  # Dynamic UI labels
  output$app_title <- renderUI({ i18n$t("title") })
  output$year_label <- renderUI({ i18n$t("year") })
  output$week_label <- renderUI({ i18n$t("week") })
  output$region_label <- renderUI({ i18n$t("region") })
  output$metric_label <- renderUI({ i18n$t("metric") })
  output$download_csv_label <- renderUI({ i18n$t("download_csv") })
  output$download_excel_label <- renderUI({ i18n$t("download_excel") })
  output$plot_title_cases <- renderUI({ i18n$t("plot_title_cases") })
  output$plot_title_metrics <- renderUI({ i18n$t("plot_title_metrics") })
  output$table_title <- renderUI({ i18n$t("table_title") })
  
  # Reactive filtered data
  filtered_data <- reactive({
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
  
  # Line plot for cases and deaths over time
  output$cases_plot <- renderPlotly({
    data <- filtered_data() %>%
      mutate(Year_Week = as.numeric(paste0(Year, sprintf("%02d", week))))
    
    p <- plot_ly(data, x = ~Year_Week) %>%
      add_lines(y = ~Suspected_Cases, name = i18n$t("suspected_cases"), line = list(color = "#1f77b4")) %>%
      add_lines(y = ~Confirmed_Cases, name = i18n$t("confirmed_cases"), line = list(color = "#ff7f0e")) %>%
      add_lines(y = ~Confirmed_Deaths, name = i18n$t("confirmed_deaths"), line = list(color = "#d62728")) %>%
      layout(
        xaxis = list(title = "Year-Week / Année-Semaine"),
        yaxis = list(title = "Count / Nombre"),
        legend = list(orientation = "h", x = 0, y = -0.2)
      )
    p
  })
  
  # Bar plot for selected metric by region
  output$metrics_plot <- renderPlotly({
    data <- filtered_data()
    p <- plot_ly(data, x = ~region, y = ~get(input$metric), type = "bar",
                 name = i18n$t(input$metric),
                 marker = list(color = "#2ca02c")) %>%
      layout(
        xaxis = list(title = i18n$t("region"), tickangle = 45),
        yaxis = list(title = i18n$t(input$metric))
      )
    p
  })
  
  # Data table
  output$data_table <- renderDT({
    data <- filtered_data() %>%
      rename(
        "Year / Année" = Year,
        "Week / Semaine" = week,
        "Region / Région" = region,
        "Suspected Cases / Cas Suspectés" = Suspected_Cases,
        "Confirmed Cases / Cas Confirmés" = Confirmed_Cases,
        "Confirmed Deaths / Décès Confirmés" = Confirmed_Deaths,
        "Population (Suspected) / Population (Suspectée)" = Total_Population_Suspected,
        "Population (Confirmed) / Population (Confirmée)" = Total_Population_Confirmed,
        "Positive Rate (%) / Taux de Positivité (%)" = Positive_Rate,
        "Attack Rate (per 100,000) / Taux d'Attaque (par 100 000)" = Confirmed_Attack_Rate,
        "Fatality Rate (%) / Taux de Létalité (%)" = Confirmed_Fatality_Rate
      )
    datatable(data, options = list(pageLength = 10, autoWidth = TRUE), rownames = FALSE)
  })
  
  # Download CSV
  output$download_csv <- downloadHandler(
    filename = function() { "Malaria_Indicators.csv" },
    content = function(file) {
      write_csv(filtered_data(), file)
    }
  )
  
  # Download Excel
  output$download_excel <- downloadHandler(
    filename = function() { "Malaria_Indicators.xlsx" },
    content = function(file) {
      write_xlsx(filtered_data(), file)
    }
  )
}

# Run the application
shinyApp(ui = ui, server = server)