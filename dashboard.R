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
library(jsonlite) # Included in your new code, though not explicitly used for JSON parsing in this version

# --- Global Data Loading and Cleaning ---
file_path <- "MDO 2024 S43.xls"
sheet_names <- c("Palu Susp", "Palu confirmé")

# Function to process a single sheet
process_sheet <- function(file_path, sheet_name) {
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
    warning(paste("Error reading header for sheet", sheet_name, ":", e$message))
    return(NULL)
  })
  if (is.null(header_df_raw)) return(NULL)
  
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
    warning(paste("Error reading data for sheet", sheet_name, ":", e$message))
    return(NULL)
  })
  if (is.null(health_data_raw)) return(NULL)
  
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
  
  # Clean and reshape data
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
    mutate(Sheet = sheet_name)
}

# Process both sheets and combine
health_data_for_shiny <- map_dfr(sheet_names, ~process_sheet(file_path, .x))

# Check if data is empty
if (nrow(health_data_for_shiny) == 0) {
  showNotification("No data loaded from either sheet. Please check the Excel file and sheet names.", type = "error")
  stop("No data loaded from either sheet. Please check the Excel file and sheet names.")
}

# Get unique districts and metrics
unique_districts <- sort(unique(health_data_for_shiny$district[!is.na(health_data_for_shiny$district) & health_data_for_shiny$district != ""]))
unique_metrics_overall <- c("Cases", "Deaths")

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
  
  title = "Malaria Surveillance Data Dashboard (Niger 2024)",
  
  navbarPage(
    title = span(icon("chart-line"), "Malaria Dashboard"),
    collapsible = TRUE,
    
    tabPanel(
      title = span(icon("info-circle"), "Overview"),
      h2("Overall Malaria Trends and Summaries in Niger, 2024"),
      p("Select a sheet to view overall trends for suspected or confirmed malaria cases."),
      selectInput("sheet_select_overview", "Select Data Sheet:", choices = sheet_names, selected = "Palu Susp"),
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
          card_header(h4(icon("chart-area"), "Overall Cases and Deaths Trend")),
          selectInput("overall_metric_select", "Select Metric for Overall Trend:",
                      choices = unique_metrics_overall, selected = "Cases"),
          plotlyOutput("overall_cases_deaths_plot", height = "300px")
        )
      )
    ),
    
    tabPanel(
      title = span(icon("map-marked-alt"), "District Details"),
      sidebarLayout(
        sidebarPanel(
          h4(icon("filter"), "Select Sheet and District"),
          selectInput("sheet_select_district", "Select Data Sheet:", choices = sheet_names, selected = "Palu Susp"),
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
      title = span(icon("chart-bar"), "Sheet Comparison"),
      h2("Compare Suspected vs. Confirmed Malaria Data"),
      p("Select a metric to compare trends between suspected ('Palu Susp') and confirmed ('Palu confirmé') malaria cases."),
      selectInput("comparison_metric_select", "Select Metric for Comparison:",
                  choices = c("Cases", "Deaths", "Attack_Rate", "Fatality_Rate"),
                  selected = "Cases"),
      hr(),
      card(
        card_header(h4(icon("chart-line"), "Comparison Trend")),
        plotlyOutput("comparison_plot", height = "400px")
      )
    ),
    
    tabPanel(
      title = span(icon("database"), "Full Data Table"),
      h2("Complete Cleaned Malaria Surveillance Data"),
      p("Select a sheet to explore the full dataset."),
      selectInput("sheet_select_table", "Select Data Sheet:", choices = sheet_names, selected = "Palu Susp"),
      hr(),
      DT::dataTableOutput("full_data_table")
    ),
    
    tabPanel(
      title = span(icon("info-circle"), "About"),
      fluidRow(
        column(12,
               h2("About This Dashboard"),
               p("This interactive dashboard provides a comprehensive view of suspected and confirmed malaria cases in Niger for 2024, based on weekly epidemiological data from 'Palu Susp' and 'Palu confirmé' sheets."),
               h4("Data Source:"),
               p("Data is extracted from the 'MDO 2024 S43.xls' Excel file, covering both suspected ('Palu Susp') and confirmed ('Palu confirmé') malaria cases."),
               h4("Methodology:"),
               tags$ul(
                 tags$li("Data from both sheets is loaded using `readxl`, handling complex multi-row headers programmatically."),
                 tags$li("Column names are standardized using `janitor::clean_names()`."),
                 tags$li("Numeric data is converted using `readr::parse_number()`."),
                 tags$li("Data is reshaped using `pivot_longer()` and `pivot_wider()`."),
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
    Cases = "#1F77B4",
    Deaths = "#D62728",
    Attack_Rate = "#2CA02C",
    Fatality_Rate = "#9467BD",
    Suspected = "#1F77B4",
    Confirmed = "#D62728"
  )
  
  # Reactive Data for Overview Tab
  overall_summary_data <- reactive({
    health_data_for_shiny %>%
      filter(Sheet == input$sheet_select_overview) %>%
      mutate(week_num = as.numeric(week)) %>%
      filter(!is.na(week_num)) %>%
      group_by(week_num) %>%
      summarise(
        Total_Cases = sum(Cases, na.rm = TRUE),
        Total_Deaths = sum(Deaths, na.rm = TRUE),
        Total_Population = sum(unique(population), na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      # Re-introduced complete() to fill missing weeks with 0 for continuous lines
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
    paste0(format(sum(overall_summary_data()$Total_Cases, na.rm = TRUE), big.mark = ","))
  })
  output$total_deaths_summary <- renderText({
    paste0(format(sum(overall_summary_data()$Total_Deaths, na.rm = TRUE), big.mark = ","))
  })
  output$overall_attack_rate_summary <- renderText({
    total_pop_all_weeks <- sum(unique(health_data_for_shiny$population[health_data_for_shiny$Sheet == input$sheet_select_overview]), na.rm = TRUE)
    total_cases_all_weeks <- sum(health_data_for_shiny$Cases[health_data_for_shiny$Sheet == input$sheet_select_overview], na.rm = TRUE)
    overall_attack_rate <- (total_cases_all_weeks / total_pop_all_weeks) * 100000
    paste0(format(round(overall_attack_rate, 2), big.mark = ","), " per 100,000")
  })
  output$overall_fatality_rate_summary <- renderText({
    total_cases_all_weeks <- sum(health_data_for_shiny$Cases[health_data_for_shiny$Sheet == input$sheet_select_overview], na.rm = TRUE)
    total_deaths_all_weeks <- sum(health_data_for_shiny$Deaths[health_data_for_shiny$Sheet == input$sheet_select_overview], na.rm = TRUE)
    overall_fatality_rate <- ifelse(total_cases_all_weeks > 0, (total_deaths_all_weeks / total_cases_all_weeks) * 100, 0)
    paste0(format(round(overall_fatality_rate, 2), big.mark = ","), "%")
  })
  
  # Overview Tab - Overall Cases and Deaths Plot
  output$overall_cases_deaths_plot <- renderPlotly({
    plot_data <- overall_summary_data()
    selected_metric_overall <- input$overall_metric_select
    y_col_name <- paste0("Total_", selected_metric_overall)
    
    if (nrow(plot_data) < 2 || all(is.na(plot_data[[y_col_name]]))) {
      p <- ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = paste("No overall data available for", selected_metric_overall, "plotting in", input$sheet_select_overview), 
                 size = 6, color = "red", family = "Inter") +
        theme_void()
      plotly::ggplotly(p)
    } else {
      p <- ggplot(plot_data, aes(x = week_num, y = .data[[y_col_name]],
                                 text = paste("Week:", week_num, "<br>",
                                              selected_metric_overall, ":", round(.data[[y_col_name]], 2)),
                                 group = 1)) + 
        geom_line(color = metric_colors[[selected_metric_overall]], size = 0.8) + # Adjusted line size
        geom_point(color = metric_colors[[selected_metric_overall]], size = 1.8, alpha = 0.9) + # Adjusted point size
        labs(
          title = paste("Overall Malaria", selected_metric_overall, "Trend (", input$sheet_select_overview, ") - 2024"),
          x = "Epidemiological Week",
          y = selected_metric_overall
        ) +
        scale_x_continuous(breaks = seq(min(plot_data$week_num, na.rm = TRUE), 
                                        max(plot_data$week_num, na.rm = TRUE), by = 5),
                           labels = function(x) as.integer(x),
                           limits = c(1, max(plot_data$week_num, na.rm = TRUE))) +
        scale_y_continuous(limits = c(0, max(plot_data[[y_col_name]], na.rm = TRUE) * 1.1),
                           labels = scales::comma) +
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
            activecolor = metric_colors[[selected_metric_overall]]
          )
        ) %>%
        config(
          displaylogo = FALSE,
          modeBarButtonsToAdd = c("downloadImage", "zoom2d", "pan2d", "select2d", "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"),
          toImageButtonOptions = list(format = "png", filename = paste0("overall_", selected_metric_overall, "_", input$sheet_select_overview))
        )
    }
  })
  
  # Reactive Data for District Details Tab
  filtered_district_data <- reactive({
    req(input$sheet_select_district, input$selected_district)
    
    health_data_for_shiny %>%
      filter(Sheet == input$sheet_select_district, district == input$selected_district) %>%
      mutate(
        week_num = as.numeric(week),
        Cases = as.numeric(Cases),
        Deaths = as.numeric(Deaths),
        Attack_Rate = as.numeric(Attack_Rate),
        Fatality_Rate = as.numeric(Fatality_Rate)
      ) %>%
      filter(!is.na(week_num)) %>%
      arrange(week_num) %>%
      # Re-introduced complete() to fill missing weeks with 0 for continuous lines
      group_by(district, region) %>% # Group by district and region to preserve them during complete()
      complete(week_num = full_seq(week_num, period = 1),
               fill = list(Cases = 0, Deaths = 0, Attack_Rate = 0, Fatality_Rate = 0)) %>% 
      ungroup() %>%
      # Re-fill population for newly completed rows if it's NA.
      group_by(district) %>%
      fill(population, .direction = "downup") %>% # Fill NA population from existing values in the district
      ungroup() %>%
      # Recalculate rates after filling with 0s for cases/deaths and ensuring population
      mutate(
        Attack_Rate = ifelse(population > 0, (Cases / population) * 100000, 0),
        Fatality_Rate = ifelse(Cases > 0, (Deaths / Cases) * 100, 0)
      )
  })
  
  # Reactive Caption for District Summary Table
  output$district_summary_caption <- renderText({
    paste("Summary for", input$selected_district, "(", input$sheet_select_district, ")")
  })
  
  # District Summary Table
  output$district_summary_table <- renderTable({
    summary_data <- filtered_district_data() %>%
      summarise(
        `Total Cases` = sum(Cases, na.rm = TRUE),
        `Total Deaths` = sum(Deaths, na.rm = TRUE),
        `Average Attack Rate` = mean(Attack_Rate, na.rm = TRUE),
        `Average Fatality Rate` = mean(Fatality_Rate, na.rm = TRUE)
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
    paste("Malaria Trends in", input$selected_district, "(", input$sheet_select_district, ") - 2024")
  })
  
  # Render Plotly Function for District Details
  render_district_plotly <- function(metric_name, y_label, color) {
    renderPlotly({
      plot_data <- filtered_district_data()
      
      if (nrow(plot_data) < 2 || all(is.na(plot_data[[metric_name]]))) {
        p <- ggplot() +
          annotate("text", x = 0.5, y = 0.5, 
                   label = paste("No valid data for", metric_name, "in", input$selected_district, "(", input$sheet_select_district, ")"), 
                   size = 5, color = "red", family = "Inter") +
          theme_void()
        plotly::ggplotly(p)
      } else {
        p <- ggplot(plot_data, aes(x = week_num, y = .data[[metric_name]],
                                   text = paste("Week:", week_num, "<br>",
                                                metric_name, ":", round(.data[[metric_name]], 2)),
                                   group = 1)) + 
          geom_line(color = color, size = 0.8) + # Adjusted line size
          geom_point(color = color, size = 1.8, alpha = 0.9) + # Adjusted point size
          labs(
            title = paste("Malaria", metric_name, "in", input$selected_district, "(", input$sheet_select_district, ") - 2024"),
            x = "Epidemiological Week",
            y = y_label
          ) +
          scale_x_continuous(breaks = seq(min(plot_data$week_num, na.rm = TRUE), 
                                          max(plot_data$week_num, na.rm = TRUE), by = 5),
                             labels = function(x) as.integer(x),
                             limits = c(1, max(plot_data$week_num, na.rm = TRUE))) +
          scale_y_continuous(limits = c(0, max(plot_data[[metric_name]], na.rm = TRUE) * 1.1),
                             labels = scales::comma) +
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
            toImageButtonOptions = list(format = "png", filename = paste0(input$selected_district, "_", metric_name, "_", input$sheet_select_district))
          )
      }
    })
  }
  
  # Render Plots for District Details Tab
  output$cases_plot <- render_district_plotly("Cases", "Number of Cases", metric_colors$Cases)
  output$deaths_plot <- render_district_plotly("Deaths", "Number of Deaths", metric_colors$Deaths)
  output$attack_rate_plot <- render_district_plotly("Attack_Rate", "Attack Rate (per 100,000)", metric_colors$Attack_Rate)
  output$fatality_rate_plot <- render_district_plotly("Fatality_Rate", "Fatality Rate (%)", metric_colors$Fatality_Rate)
  
  # Reactive Data for Comparison Tab
  comparison_data <- reactive({
    health_data_for_shiny %>%
      filter(!is.na(week)) %>%
      mutate(week_num = as.numeric(week)) %>%
      group_by(Sheet, week_num) %>%
      summarise(
        Total_Cases = sum(Cases, na.rm = TRUE),
        Total_Deaths = sum(Deaths, na.rm = TRUE),
        Total_Population = sum(unique(population), na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      mutate(
        Attack_Rate = (Total_Cases / Total_Population) * 100000,
        Fatality_Rate = ifelse(Total_Cases > 0, (Total_Deaths / Total_Cases) * 100, 0)
      ) %>%
      pivot_wider(
        names_from = Sheet,
        values_from = c(Total_Cases, Total_Deaths, Attack_Rate, Fatality_Rate),
        names_glue = "{.value}_{Sheet}"
      ) %>%
      arrange(week_num) %>%
      complete(week_num = seq(min(week_num, na.rm = TRUE), max(week_num, na.rm = TRUE), by = 1),
               fill = list(
                 `Total_Cases_Palu Susp` = 0,
                 `Total_Cases_Palu confirmé` = 0,
                 `Total_Deaths_Palu Susp` = 0,
                 `Total_Deaths_Palu confirmé` = 0,
                 `Attack_Rate_Palu Susp` = 0,
                 `Attack_Rate_Palu confirmé` = 0,
                 `Fatality_Rate_Palu Susp` = 0,
                 `Fatality_Rate_Palu confirmé` = 0
               ))
  })
  
  # Comparison Plot (using Plotly)
  output$comparison_plot <- renderPlotly({
    plot_data <- comparison_data()
    selected_metric <- input$comparison_metric_select
    
    # Dynamically select column names based on the chosen metric
    y_col_susp <- case_when(
      selected_metric == "Cases" ~ "Total_Cases_Palu Susp",
      selected_metric == "Deaths" ~ "Total_Deaths_Palu Susp",
      selected_metric == "Attack_Rate" ~ "Attack_Rate_Palu Susp",
      selected_metric == "Fatality_Rate" ~ "Fatality_Rate_Palu Susp",
      TRUE ~ NA_character_ # Should not happen with defined choices
    )
    y_col_conf <- case_when(
      selected_metric == "Cases" ~ "Total_Cases_Palu confirmé",
      selected_metric == "Deaths" ~ "Total_Deaths_Palu confirmé",
      selected_metric == "Attack_Rate" ~ "Attack_Rate_Palu confirmé",
      selected_metric == "Fatality_Rate" ~ "Fatality_Rate_Palu confirmé",
      TRUE ~ NA_character_ # Should not happen with defined choices
    )
    
    if (nrow(plot_data) < 2 || (all(is.na(plot_data[[y_col_susp]])) && all(is.na(plot_data[[y_col_conf]])))) {
      p <- ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = paste("No valid data for", selected_metric, "comparison."), 
                 size = 5, color = "red", family = "Inter") +
        theme_void()
      return(plotly::ggplotly(p))
    }
    
    p <- ggplot(plot_data) +
      geom_line(aes(x = week_num, y = .data[[y_col_susp]], color = "Suspected"), size = 0.8) + # Adjusted line size
      geom_point(aes(x = week_num, y = .data[[y_col_susp]], color = "Suspected"), size = 1.8) + # Added points, adjusted size
      geom_line(aes(x = week_num, y = .data[[y_col_conf]], color = "Confirmed"), size = 0.8) + # Adjusted line size
      geom_point(aes(x = week_num, y = .data[[y_col_conf]], color = "Confirmed"), size = 1.8) + # Added points, adjusted size
      labs(
        title = paste("Comparison of", selected_metric, "Between Suspected and Confirmed Cases (2024)"),
        x = "Epidemiological Week",
        y = if (selected_metric == "Cases") "Number of Cases"
        else if (selected_metric == "Deaths") "Number of Deaths"
        else if (selected_metric == "Attack_Rate") "Attack Rate (per 100,000)"
        else "Fatality Rate (%)"
      ) +
      scale_color_manual(values = c("Suspected" = metric_colors$Suspected, "Confirmed" = metric_colors$Confirmed)) +
      scale_x_continuous(breaks = seq(min(plot_data$week_num, na.rm = TRUE), 
                                      max(plot_data$week_num, na.rm = TRUE), by = 5),
                         labels = function(x) as.integer(x)) +
      scale_y_continuous(labels = scales::comma) +
      custom_plot_theme() +
      theme(legend.position = "top", legend.title = element_blank())
    
    plotly::ggplotly(p, tooltip = c("x", "y", "color")) %>%
      layout(
        hovermode = "x unified",
        margin = list(t = 50, b = 50, l = 50, r = 50),
        legend = list(orientation = "h", x = 0.5, xanchor = "center", y = 1.1),
        dragmode = "zoom",
        xaxis = list(fixedrange = FALSE),
        yaxis = list(fixedrange = FALSE),
        modebar = list(
          bgcolor = "transparent",
          color = "#000000",
          activecolor = metric_colors[[selected_metric]]
        )
      ) %>%
      config(
        displaylogo = FALSE,
        modeBarButtonsToAdd = c("downloadImage", "zoom2d", "pan2d", "select2d", "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"),
        toImageButtonOptions = list(format = "png", filename = paste0("comparison_", selected_metric))
      )
  })
  
  # Data Table Tab
  output$full_data_table <- DT::renderDataTable({
    DT::datatable(
      health_data_for_shiny %>% filter(Sheet == input$sheet_select_table),
      options = list(pageLength = 10, scrollX = TRUE,
                     dom = 'lfrtip',
                     language = list(search = "Search Data:")),
      rownames = FALSE,
      caption = htmltools::tags$caption(
        style = 'caption-side: top; text-align: center; color: black; font-size: 18px;',
        paste('Complete Malaria Surveillance Data (', input$sheet_select_table, ')')
      )
    )
  })
}

# Run the Shiny app
shinyApp(ui = ui, server = server)
