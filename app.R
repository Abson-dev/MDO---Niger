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

# --- Global Data Loading and Cleaning (Runs once when the app starts) ---
# This section is outside the ui and server functions so it's executed only once.

# Define File Path and Sheet Name
file_path <- "MDO 2024 S43.xls"
#sheet_name <- "Palu Susp"
sheet_name <- "Palu confirmé"
#
# PART 1: Determine Column Names Programmatically & Read Data Robustly
# Step 1a: Read header rows to construct column names
header_df_raw <- read_excel(
  path = file_path,
  sheet = sheet_name,
  range = cell_rows(1:5),
  col_names = FALSE,
  col_types = "text"
)

# Fixed columns based on previous debugging and sheet structure
fixed_cols_from_excel <- as.character(header_df_raw[3, 1:5]) # Region, District, Isocode, Année, Population

weekly_constructed_names <- c()
current_semaine_col_in_excel <- 6 # Excel column F

while(current_semaine_col_in_excel <= ncol(header_df_raw) &&
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

# Step 1b: Read the actual data block starting from row 6
health_data_raw <- read_excel(
  path = file_path,
  sheet = sheet_name,
  skip = 5, # Skip first 5 rows, so data starts from Excel row 6
  col_names = FALSE, # We will assign names later
  col_types = "text" # Read everything as text
)

# Compare number of columns and adjust names if necessary
if (ncol(health_data_raw) < length(all_column_names_constructed)) {
  all_column_names_final <- all_column_names_constructed[1:ncol(health_data_raw)]
} else if (ncol(health_data_raw) > length(all_column_names_constructed)) {
  num_extra_cols <- ncol(health_data_raw) - length(all_column_names_constructed)
  padded_names <- c(all_column_names_final, paste0("Unknown_", 1:num_extra_cols))
  all_column_names_final <- padded_names
} else {
  all_column_names_final <- all_column_names_constructed
}

health_data <- health_data_raw %>%
  set_names(all_column_names_final)

# PART 2: Clean and Reshape Data
health_data_filtered <- health_data %>%
  filter(!is.na(Region) & Region != "Pays:" & Region != "Region" & !grepl("Total", Region, ignore.case = TRUE))

health_data_clean_names <- health_data_filtered %>%
  janitor::clean_names()

health_data_typed <- health_data_clean_names %>%
  mutate(
    population = readr::parse_number(population),
    across(starts_with("semaine_") &
             (contains("cas") | contains("deces") |
                contains("taux_d_attaque") | contains("letalite")),
           readr::parse_number)
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

# Final cleaning/formatting for display in app
health_data_for_shiny <- health_data_final_wide %>%
  dplyr::select(region, district, population, week, Cases, Deaths, Attack_Rate, Fatality_Rate) %>%
  dplyr::filter(!is.na(week)) %>%
  dplyr::filter(!region %in% c("REGION")) %>%
  dplyr::mutate(
    population = round(population, 0)
  )

# Get unique districts for the dropdown menu
unique_districts <- sort(unique(health_data_for_shiny$district))
# Get unique metrics for the dropdown menu (for overall plot)
unique_metrics_overall <- c("Cases", "Deaths")

# --- Shiny UI (User Interface) ---
ui <- page_fluid(
  # Apply a modern Bootstrap theme
  theme = bs_theme(
    version = 5,
    bootswatch = "flatly", # A clean, modern theme
    primary = "#007bff", # Primary color for accents
    secondary = "#6c757d",
    success = "#28a745",
    info = "#17a2b8",
    warning = "#ffc107",
    danger = "#dc3545",
    base_font = font_google("Inter"), # Use Google Fonts for a nice look
    heading_font = font_google("Inter")
  ),
  
  title = paste0("Malaria Surveillance Data Dashboard (",sheet_name," - Niger 2024"),
  
  # Navbar for top-level navigation
  navbarPage(
    title = span(icon("chart-line"), "Malaria Dashboard"), # Title with an icon
    collapsible = TRUE, # Make navbar collapsible on small screens
    
    # 1. Overview Tab
    tabPanel(span(icon("info-circle"), "Overview"),
             h2("Overall Malaria Trends and Summaries in Niger, 2024"),
             p("This section provides a high-level overview of suspected malaria cases and related metrics across all districts."),
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
    
    # 2. District Details Tab
    tabPanel(span(icon("map-marked-alt"), "District Details"),
             sidebarLayout(
               sidebarPanel(
                 h4(icon("filter"), "Select District for Detailed Analysis"),
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
    
    # 3. Data Table Tab
    tabPanel(span(icon("database"), "Full Data Table"),
             h2("Complete Cleaned Malaria Surveillance Data"),
             p("Explore the full dataset used in this dashboard. You can search, sort, and paginate through the data."),
             hr(),
             DT::dataTableOutput("full_data_table")
    ),
    
    # 4. About Tab
    tabPanel(span(icon("info-circle"), "About"),
             fluidRow(
               column(12,
                      h2("About This Dashboard"),
                      p("This interactive dashboard provides a comprehensive view of suspected malaria cases in Niger for the year 2024, based on weekly epidemiological data."),
                      h4("Data Source:"),
                      p(paste0("The data is extracted from the 'MDO 2024 S43.xls' Excel file, specifically from the ",sheet_name," sheet. This raw data contains multi-level headers and requires robust cleaning and reshaping for analysis.")),
                      h4("Methodology:"),
                      tags$ul(
                        tags$li("Data is loaded using `readxl`, handling complex multi-row headers programmatically."),
                        tags$li("Column names are standardized using `janitor::clean_names()`."),
                        tags$li("Numeric data (Population, Cases, Deaths, Rates) is robustly converted using `readr::parse_number()` to handle various Excel formatting nuances."),
                        tags$li("The data is transformed from a wide (weekly columns) to a long (tidy) format using `pivot_longer()`, and then back to a slightly wider format with distinct metric columns using `pivot_wider()`."),
                        tags$li("Summary rows and irrelevant header rows are filtered out to ensure data integrity."),
                        tags$li("Time-series plots are generated using `ggplot2` and made interactive with `plotly` to visualize trends for selected districts and metrics."),
                        tags$li("Interactive tables are provided using the `DT` package for detailed data exploration."),
                        tags$li("The dashboard design is enhanced using `bslib` for a modern and responsive user interface.")
                      ),
                      h4("Contact Information:"),
                      p("For questions or feedback, please contact Médecins Sans Frontières at jobs@msf.org.")
               )
             )
    )
  )
)

# --- Shiny Server Logic ---
server <- function(input, output, session) {
  
  # --- Custom Plot Theme ---
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
  
  # --- Color Palette for Metrics ---
  metric_colors <- list(
    Cases = "#1F77B4",        # Blue
    Deaths = "#D62728",       # Red
    Attack_Rate = "#2CA02C",  # Green
    Fatality_Rate = "#9467BD" # Purple
  )
  
  # --- Reactive Data for Overview Tab ---
  overall_summary_data <- reactive({
    health_data_for_shiny %>%
      mutate(week_num = as.numeric(week)) %>%
      filter(!is.na(week_num)) %>%
      group_by(week_num) %>%
      summarise(
        Total_Cases = sum(Cases, na.rm = TRUE),
        Total_Deaths = sum(Deaths, na.rm = TRUE),
        Total_Population = sum(unique(population), na.rm = TRUE),
        .groups = 'drop'
      ) %>%
      mutate(
        Overall_Attack_Rate = (Total_Cases / Total_Population) * 100000,
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
    total_pop_all_weeks <- sum(unique(health_data_for_shiny$population), na.rm = TRUE)
    total_cases_all_weeks <- sum(health_data_for_shiny$Cases, na.rm = TRUE)
    overall_attack_rate <- (total_cases_all_weeks / total_pop_all_weeks) * 100000
    paste0(format(round(overall_attack_rate, 2), big.mark = ","), " per 100,000")
  })
  output$overall_fatality_rate_summary <- renderText({
    total_cases_all_weeks <- sum(health_data_for_shiny$Cases, na.rm = TRUE)
    total_deaths_all_weeks <- sum(health_data_for_shiny$Deaths, na.rm = TRUE)
    overall_fatality_rate <- ifelse(total_cases_all_weeks > 0, (total_deaths_all_weeks / total_cases_all_weeks) * 100, 0)
    paste0(format(round(overall_fatality_rate, 2), big.mark = ","), "%")
  })
  
  # Overview Tab - Enhanced Overall Cases and Deaths Plot
  output$overall_cases_deaths_plot <- renderPlotly({
    plot_data <- overall_summary_data()
    selected_metric_overall <- input$overall_metric_select
    y_col_name <- paste0("Total_", selected_metric_overall)
    
    if (nrow(plot_data) < 2 || all(is.na(plot_data[[y_col_name]]))) {
      p <- ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = paste("No overall data available for", selected_metric_overall, "plotting."), 
                 size = 6, color = "red", family = "Inter") +
        theme_void()
      plotly::ggplotly(p)
    } else {
      p <- ggplot(plot_data, aes(x = week_num, y = .data[[y_col_name]],
                                 text = paste("Week:", week_num, "<br>",
                                              selected_metric_overall, ":", round(.data[[y_col_name]], 2)))) +
        geom_point(color = metric_colors[[selected_metric_overall]], size = 1, alpha = 0.9) +
        labs(
          title = paste("Overall Malaria", selected_metric_overall, "Trend in Niger, 2024"),
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
            bgcolor = "transparent",
            color = "#000000",
            activecolor = metric_colors[[selected_metric_overall]]
          )
        ) %>%
        config(
          displaylogo = FALSE,
          modeBarButtonsToAdd = c("downloadImage", "zoom2d", "pan2d", "select2d", "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"),
          toImageButtonOptions = list(format = "png", filename = paste0("overall_", selected_metric_overall, "_trend"))
        )
    }
  })
  
  # --- Reactive Data for District Details Tab ---
  filtered_district_data <- reactive({
    req(input$selected_district)
    
    health_data_for_shiny %>%
      filter(district == input$selected_district) %>%
      mutate(
        week_num = as.numeric(week),
        Cases = as.numeric(Cases),
        Deaths = as.numeric(Deaths),
        Attack_Rate = as.numeric(Attack_Rate),
        Fatality_Rate = as.numeric(Fatality_Rate)
      ) %>%
      filter(!is.na(week_num)) %>%
      arrange(week_num)
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
  }, caption = "Summary for Selected District", caption.placement = "top")
  
  # Plot title for District Details tab
  output$district_plots_title <- renderText({
    paste("Malaria Trends in", input$selected_district, "- 2024")
  })
  
  # Enhanced Render Plotly Function for District Details
  render_district_plotly <- function(metric_name, y_label, color) {
    renderPlotly({
      plot_data <- filtered_district_data()
      
      if (nrow(plot_data) < 2 || all(is.na(plot_data[[metric_name]]))) {
        p <- ggplot() +
          annotate("text", x = 0.5, y = 0.5, 
                   label = paste("No valid data for", metric_name, "in", input$selected_district), 
                   size = 5, color = "red", family = "Inter") +
          theme_void()
        plotly::ggplotly(p)
      } else {
        p <- ggplot(plot_data, aes(x = week_num, y = .data[[metric_name]],
                                   text = paste("Week:", week_num, "<br>",
                                                metric_name, ":", round(.data[[metric_name]], 2)))) +
          geom_point(color = color, size = 1, alpha = 0.9) +
          labs(
            title = paste("Malaria", metric_name, "in", input$selected_district, "- 2024"),
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
            toImageButtonOptions = list(format = "png", filename = paste0(input$selected_district, "_", metric_name, "_trend"))
          )
      }
    })
  }
  
  # Render Plots for District Details Tab
  output$cases_plot <- render_district_plotly("Cases", "Number of Cases", metric_colors$Cases)
  output$deaths_plot <- render_district_plotly("Deaths", "Number of Deaths", metric_colors$Deaths)
  output$attack_rate_plot <- render_district_plotly("Attack_Rate", "Attack Rate (per 100,000)", metric_colors$Attack_Rate)
  output$fatality_rate_plot <- render_district_plotly("Fatality_Rate", "Fatality Rate (%)", metric_colors$Fatality_Rate)
  
  # --- Data Table Tab ---
  output$full_data_table <- DT::renderDataTable({
    DT::datatable(health_data_for_shiny,
                  options = list(pageLength = 10, scrollX = TRUE,
                                 dom = 'lfrtip',
                                 language = list(search = "Search Data:")),
                  rownames = FALSE,
                  caption = htmltools::tags$caption(
                    style = 'caption-side: top; text-align: center; color: black; font-size: 18px;',
                    'Complete Malaria Surveillance Data'
                  )
    )
  })
}

# Run the Shiny app
shinyApp(ui = ui, server = server)
