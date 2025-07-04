# MDO---Niger
# Report on Malaria Data Processing and Visualization (Niger, MDO 2024 S43)

**Date:** July 4, 2025
**Location:** Dakar, Senegal
**Data Source:** `MDO 2024 S43.xls` (Sheet: "Palu Susp")

---

### 1. Introduction

This report outlines the R programming workflow used to process raw malaria surveillance data from the "MDO 2024 S43.xls" Excel file, specifically focusing on the "Palu Susp" (Suspected Malaria Cases) sheet. The objective of this script is to clean and reshape the weekly reported data, generate time-series plots for key malaria indicators for each district, and compile these plots along with the raw data into a structured Excel workbook.

### 2. Data Loading and Initial Inspection (Part 1)

The first phase of the script focuses on robustly loading the data from the specified Excel sheet. A common challenge with this type of data is that header information spans multiple rows, and the number of weekly columns can vary.

#### 2.1. Dynamic Header Extraction

To address the multi-row header, the script first reads the initial five rows of the Excel sheet without assigning column names. This allows for programmatic extraction of the relevant header information:
* **Fixed Columns:** The first five columns (`Region`, `District`, `Isocode`, `Année`, `Population`) are identified from the third row of the header.
* **Weekly Data Columns:** The script then iteratively constructs names for the weekly epidemiological data. It looks for "Semaine X" (Week X) headers in the third row and corresponding sub-headers (`cas`, `décès`, `Taux d'attaque`, `Létalité`) in the fifth row. For each "Semaine X", four specific column names are generated (e.g., `Semaine 1_cas`, `Semaine 1_deces`, `Semaine 1_taux_d_attaque`, `Semaine 1_letalite`). This dynamic approach ensures that all available weekly data columns are correctly named, regardless of the number of weeks present in the file.

#### 2.2. Raw Data Import

After constructing the full list of column names, the main data block is imported starting from the 6th row (skipping the 5 header rows). All columns are initially read as 'text' to prevent premature and potentially incorrect type conversions by `readxl`.
A crucial step involves comparing the number of columns read from the Excel file (`ncol(health_data_raw)`) with the number of programmatically `all_column_names_constructed`. This check ensures that no data columns are missed or misaligned. If a mismatch occurs, the script either trims the constructed names or pads them with generic names (`Unknown_X`) to align with the actual data dimensions, with warnings issued for such adjustments. Finally, the constructed names are assigned to the raw data frame.

### 3. Data Cleaning and Reshaping (Part 2)

Once the raw data is loaded with correct column names, the next phase involves cleaning and transforming it into a "tidy" format suitable for analysis and visualization.

#### 3.1. Filtering and Renaming

* **Row Filtering:** The data is filtered to remove extraneous summary rows (e.g., "Pays:", "Region", or rows containing "Total") which are often present in raw reports but not needed for granular analysis. This ensures that only valid district-level data is retained.
* **Column Name Cleaning:** The `janitor::clean_names()` function is applied to standardize column names (e.g., "Taux d'attaque" becomes "taux_d_attaque"). This makes column access consistent and robust to special characters or spaces.

#### 3.2. Type Conversion

Key columns are then converted to their appropriate numeric types using `readr::parse_number()`. This function is particularly robust as it can handle non-numeric characters (like spaces or commas) often found in Excel numeric fields, extracting the underlying number. This is applied to `population` and all the generated weekly metric columns (`semaine_X_cas`, `semaine_X_deces`, etc.).

#### 3.3. Wide-to-Long Reshaping

The data is currently in a "wide" format, with each week having multiple columns (e.g., `Semaine_1_cas`, `Semaine_1_deces`, `Semaine_2_cas`, etc.). For time-series analysis and plotting, a "long" format is more suitable.
* The `pivot_longer()` function transforms the weekly columns into two new columns: `Week_Number_Raw` (containing the week number extracted from the original column name) and `Metric_Raw` (containing the type of metric, e.g., "cas", "deces"). The corresponding numeric value is placed in a `Value` column.
* The `Week_Number_Raw` is converted to an integer `Week`, and `Metric_Raw` is renamed to more descriptive English terms (`Cases`, `Deaths`, `Attack_Rate`, `Fatality_Rate`).
* Finally, `pivot_wider()` is used to convert the data back to a slightly wider format, but this time with one row per district per week, and separate columns for `Cases`, `Deaths`, `Attack_Rate`, and `Fatality_Rate`. This is the final tidy format, making it ready for direct plotting.
* Additional filtering ensures no `NA` weeks or "REGION" entries remain, and the `week` column is formatted as "Week X" strings for display, while `population` is rounded.

The cleaned data is then saved to a CSV file named `health_data_cleaned_final_Palu_Susp.csv`.

### 4. Graphics Generation and Excel Export (Part 3)

The final part of the script generates individualized time-series plots for each district and exports them, along with the corresponding raw data, into a single structured Excel workbook.

#### 4.1. Setup for Plotting and Excel

* A list of unique districts is extracted from the cleaned data.
* A new `openxlsx` workbook object (`wb`) is created.
* Standard plot dimensions (`plot_width_in`, `plot_height_in`) are defined in inches for consistent output quality (using a DPI of 300).

#### 4.2. Iterative Plotting and Export

The script iterates through each unique `district_name`:
1.  **District Data Filtering:** Data for the current district is filtered, and a numeric `week_num` is re-extracted from the "Week X" string for proper time-series plotting. Numeric conversions are re-checked for the metric columns to ensure they are ready for plotting. Rows with insufficient or invalid data (e.g., less than 2 data points for line plots, or `NA` values in critical metrics) are skipped with a warning.
2.  **Plot Generation:** Four `ggplot2` plots are created for each district:
    * Malaria Cases over time
    * Malaria Deaths over time
    * Malaria Attack Rate over time
    * Malaria Fatality Rate over time
    Each plot is customized with appropriate titles, axis labels, colors, and themes for clarity.
3.  **Excel Sheet Creation:** A new sheet is added to the Excel workbook for each district. The district name is cleaned (`sheet_safe_name`) to ensure it's a valid Excel sheet name (max 31 characters, no invalid characters).
4.  **Plot Insertion:** **Crucially, for compatibility with older `openxlsx` versions, each plot (`p_cases`, `p_deaths`, `p_attack_rate`, `p_fatality_rate`) is explicitly `print()`ed immediately before calling `openxlsx::insertPlot()`.** This makes the plot active on R's graphics device, allowing `insertPlot()` to capture and embed it into the Excel sheet at specified coordinates (`startRow`, `startCol`) and dimensions. The plots are arranged in a 2x2 grid on each sheet.
5.  **Data Insertion (Optional):** Below the plots on each district's sheet, the filtered raw data for that district is also written, providing a tabular reference alongside the visualizations.

#### 4.3. Workbook Saving

After processing all districts, the entire workbook is saved as `Malaria_District_Reports_Palu_Susp_with_Graphics.xlsx`.

### 5. Conclusion

This R script provides a robust and automated pipeline for handling complex, multi-header Excel data for malaria surveillance. It effectively extracts relevant information, cleans and reshapes it into a tidy format, and then generates detailed, district-specific time-series visualizations. The final output is an organized Excel workbook, making the insights readily accessible for public health reporting and decision-making. The adaptations for `openxlsx` demonstrate flexibility in handling different package versions while maintaining core functionality.
