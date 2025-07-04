# Malaria Surveillance Data Dashboard (Niger 2024)

![Malaria Dashboard Preview](https://via.placeholder.com/800x400.png?text=Malaria+Surveillance+Dashboard)  
*Interactive dashboard for visualizing malaria trends in Niger, 2024*

Welcome to the **Malaria Surveillance Data Dashboard**, a Shiny application designed to analyze and visualize suspected malaria cases in Niger for 2024, using data from the `MDO 2024 S43.xls` Excel file (specifically the `Palu Susp` sheet). This dashboard transforms raw, complex epidemiological data into intuitive, interactive visualizations, enabling public health professionals to monitor malaria trends across districts with ease.

**Date:** July 4, 2025  
**Location:** Dakar, Senegal  
**Developed for:** Médecins Sans Frontières (MSF)

## 🌟 Why This Dashboard?

Malaria remains a critical public health challenge in Niger. This dashboard empowers stakeholders by:
- Providing **interactive visualizations** of suspected malaria cases, deaths, attack rates, and fatality rates.
- Offering **district-level insights** through a user-friendly interface.
- Enabling **data exploration** with searchable, sortable tables.
- Delivering **clean, processed data** from complex Excel files, ready for analysis and reporting.

Built with R and Shiny, the dashboard combines robust data processing with modern, responsive design to support evidence-based decision-making.

## 🚀 Features

- **Overview Tab**: Summarizes total cases, deaths, attack rates, and fatality rates across Niger, with a point-based plot for Cases or Deaths, emphasizing discrete weekly trends.
- **District Details Tab**: Displays point-based plots for Cases, Deaths, Attack Rate, and Fatality Rate for a selected district, highlighting weekly data points.
- **Data Table Tab**: Offers an interactive table for exploring the full cleaned dataset, with search, sort, and pagination capabilities.
- **About Tab**: Details the data source, methodology, and contact information.
- **Visualizations**: Uses `ggplot2` with `plotly` for interactive plots (zoom, pan, PNG download). Points are styled with a professional color palette (Cases: Blue, Deaths: Red, Attack Rate: Green, Fatality Rate: Purple) and a modern `bslib` theme.
- **Data Processing**: Handles multi-level Excel headers, cleans data, and reshapes it into a tidy format using `tidyverse`.

## 📋 Prerequisites

Before setting up the dashboard, ensure you have:
- **R**: Version 4.0 or higher ([download from CRAN](https://cran.r-project.org/)).
- **RStudio**: Recommended for running the app ([download from RStudio](https://rstudio.com/)).
- **Excel File**: `MDO 2024 S43.xls` with the `Palu Susp` sheet containing malaria surveillance data.
- **Internet Connection**: Required for installing R packages and loading Google Fonts for the UI.

## 🛠 Setup Instructions

Follow these steps to set up and run the dashboard:

### 1. Clone the Repository
Clone the repository to your local machine:
```bash
git clone https://github.com/your-username/your-repository-name.git
cd your-repository-name
```
Replace `your-username` and `your-repository-name` with your actual GitHub details.

### 2. Install R Packages
Install the required R packages in R or RStudio:
```R
install.packages(c("shiny", "tidyverse", "readxl", "readr", "janitor", "DT", "plotly", "bslib", "scales"))
```
This installs:
- `shiny`: Web application framework.
- `tidyverse`: Data manipulation (`dplyr`, `tidyr`, etc.).
- `readxl`: Excel file reading.
- `readr`: Numeric parsing.
- `janitor`: Column name cleaning.
- `DT`: Interactive data tables.
- `plotly`: Interactive plots.
- `bslib`: Modern Bootstrap theming.
- `scales`: Axis formatting.

### 3. Place the Excel File
Place `MDO 2024 S43.xls` in the repository’s root directory (same as `app.R`). If the file is elsewhere, update the `file_path` in `app.R`:
```R
file_path <- "path/to/your/MDO 2024 S43.xls"
```

### 4. Run the Dashboard
Open `app.R` in RStudio and click "Run App", or run from the R console:
```R
library(shiny)
runApp()
```
The dashboard will open in your default web browser.

### 5. Explore the Dashboard
- **Overview**: View aggregated statistics and select "Cases" or "Deaths" for a point-based plot of national trends.
- **District Details**: Choose a district to see point-based plots for key metrics.
- **Data Table**: Search, sort, and paginate the cleaned dataset.
- **About**: Learn about the methodology and data source.

## 🧠 Technical Methodology

The dashboard processes complex malaria surveillance data through a robust pipeline:

### Data Loading
- **Dynamic Header Extraction**: Reads the first five rows of `MDO 2024 S43.xls` (`Palu Susp` sheet) to construct column names. Fixed columns (`Region`, `District`, `Isocode`, `Année`, `Population`) are extracted from row 3, while weekly columns (`Semaine X_cas`, `Semaine X_deces`, etc.) are built from rows 3 and 5, accommodating varying numbers of weeks.
- **Raw Data Import**: Loads data from row 6 onward, reading all columns as text to avoid type conversion errors. Column names are aligned dynamically, with warnings for mismatches.

### Data Cleaning and Reshaping
- **Filtering**: Removes summary rows (e.g., "Pays:", "Total") to retain district-level data.
- **Name Standardization**: Uses `janitor::clean_names()` to normalize column names (e.g., "Taux d'attaque" to "taux_d_attaque").
- **Type Conversion**: Converts `population` and weekly metrics to numeric using `readr::parse_number()` for robust handling of Excel formatting.
- **Reshaping**: Transforms data from wide to long format with `pivot_longer()` (creating `Week` and `Metric` columns), then pivots to a wider format with `pivot_wider()` for one row per district per week, with columns for `Cases`, `Deaths`, `Attack_Rate`, and `Fatality_Rate`.

### Visualization
- **Plots**: Uses `ggplot2` with `geom_point(size = 3, alpha = 0.9)` to display discrete weekly data points, without connecting lines or maximum value annotations. Plots are interactive via `plotly`, supporting zoom, pan, and PNG download.
- **Styling**: Applies a modern theme (`bslib`, "Inter" font, `#F8F9FA` background, `#FFFFFF` panel) with a color palette (Cases: `#1F77B4`, Deaths: `#D62728`, Attack Rate: `#2CA02C`, Fatality Rate: `#9467BD`).
- **Dynamic Axes**: X-axis shows weeks with breaks every 5 weeks; y-axis scales to 110% of the maximum value with comma-separated labels.

### Output
- **Dashboard**: A Shiny app with four tabs for overview, district details, data table, and about information.
- **Data Export**: The cleaned dataset is available for exploration within the app. (Optional CSV export can be added; see below.)

## 🛠 Troubleshooting

- **Excel File Not Found**:
  - Ensure `MDO 2024 S43.xls` is in the correct directory.
  - Add this check to `app.R`:
    ```R
    if (!file.exists(file_path)) {
      stop("Excel file not found at: ", file_path)
    }
    ```
- **Package Installation Issues**:
  - Verify internet connectivity.
  - Install packages individually: `install.packages("package-name")`.
- **Plot Issues**:
  - If plots are empty, check that the `Palu Susp` sheet contains valid data.
  - Ensure `week` values are numeric (handled by `as.numeric(week)` in the app).
- **Syntax Errors**:
  - Confirm `app.R` has no stray text or missing commas.
  - Ensure the file ends with a newline (add an empty line).

## 📂 Project Structure

- `app.R`: Main Shiny application script.
- `MDO 2024 S43.xls`: Data file (not included; provide your own).
- `README.md`: This guide.

## 📞 Contact

For questions or feedback, contact Aboubacar Hema at [aboubacarhema94@gmail.com](mailto:aboubacarhema94@gmail.com).

## 🔧 Potential Enhancements

Want to extend the dashboard? Here are some ideas:
- **CSV Export**: Add a download button for the cleaned dataset:
  ```R
  # UI, under Data Table Tab:
  downloadButton("downloadData", "Download Full Data"),
  # Server:
  output$downloadData <- downloadHandler(
    filename = function() { "malaria_data_niger_2024.csv" },
    content = function(file) { write.csv(health_data_for_shiny, file, row.names = FALSE) }
  )
  ```
- **Smoothed Trends**: Add a smoothed line alongside points:
  ```R
  geom_smooth(method = "loess", color = color, linetype = "dashed", size = 0.8, alpha = 0.2)
  ```
- **Bar Charts**: Integrate a `chart.js` bar chart for district comparisons (e.g., total cases by district).

## 📜 License

This project is licensed under the MIT License. See the `LICENSE` file for details (if applicable).

---

*Contribute to this project by submitting pull requests or opening issues on GitHub. Let’s make malaria surveillance more accessible and actionable!*
