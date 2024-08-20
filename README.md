# Excel Data Processing and Visualization Script

This Python script automates the process of cleaning, analyzing, and visualizing data from an Excel workbook. It uses the `openpyxl`, `pandas`, `matplotlib`, and `seaborn` libraries for data manipulation and visualization.

## Prerequisites

Before running the script, ensure you have the following Python libraries installed:

- `openpyxl`
- `pandas`
- `matplotlib`
- `seaborn`

You can install them using pip:

```bash
pip install openpyxl pandas matplotlib seaborn
```

## Script Overview

### 1. Processing the Workbook

The script first processes an Excel workbook to clean the data and correct the prices:

- **Load Workbook**: The script loads an Excel workbook (`filename`) and processes the sheet named 'Sheet1'.
- **Data Cleaning**: It reads the data into a pandas DataFrame, cleans the price column by converting it to numeric values, and applies a 10% discount to each price.
- **Save Cleaned Data**: The cleaned data is then saved back to the Excel sheet and also exported to a new Excel file for further analysis.
- **Add Bar Chart**: A bar chart of the corrected prices is generated and added to the Excel sheet using the `openpyxl` library.

### 2. Visualizing the Data

The script then visualizes the cleaned data using `matplotlib` and `seaborn`:

- **Histogram**: Displays the distribution of corrected prices.
- **Scatter Plot**: Plots corrected prices against row numbers to visualize trends.
- **Bar Plot**: Shows a bar plot of corrected prices.

## How to Use

1. **Run the Script**: 
   Execute the script and enter the filename of the Excel workbook when prompted.

   ```bash
   python script.py
   ```

2. **Provide the Filename**: 
   Enter the filename of the Excel file you want to process (e.g., `data.xlsx`).

3. **View the Outputs**: 
   The script will process the workbook, generate charts, and save the cleaned data in a new Excel file with a prefix `cleaned_`.

4. **Examine the Visualizations**: 
   After processing, the script will display visualizations of the cleaned data using matplotlib.

## Example

Assume you have an Excel file named `sales_data.xlsx`. The script will:

- Clean the price data in `sales_data.xlsx`.
- Save the cleaned data to `cleaned_sales_data.xlsx`.
- Generate and display visualizations such as histograms, scatter plots, and bar plots.

## Output Files

- **Original Workbook**: The original Excel file will be updated with cleaned prices and a bar chart.
- **Cleaned Workbook**: A new Excel file (`cleaned_<filename>.xlsx`) containing the cleaned data.

## Visualization Examples

The script will display the following plots:

1. **Histogram**: Visualizes the distribution of corrected prices.
2. **Scatter Plot**: Shows corrected prices across different rows.
3. **Bar Plot**: Illustrates the corrected prices using bars.

---
