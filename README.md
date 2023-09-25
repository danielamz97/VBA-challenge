## VBA Code Explanation

This repository includes VBA (Visual Basic for Applications) code designed to analyze stock data stored in Excel worksheets. Below is an overview of the structure and functionality of the code:

### `change_worksheet` Subroutine

The `change_worksheet` subroutine is responsible for iterating through all the worksheets in the Excel file, calling the `stocks` subroutine for each worksheet. This allows the analysis to be performed on multiple years' worth of stock data.

### `stocks` Subroutine

The `stocks` subroutine performs the stock data analysis for a single worksheet. Here's a breakdown of its functionality:

- **Data Setup**: The subroutine sets up the worksheet by adding headers for the analysis results such as "Ticker," "Yearly Change," "Percent Change," and "Total Stock Volume."

- **Loop Through Data**: It then loops through the rows of data in the worksheet, calculating the yearly change, percentage change, and total stock volume for each stock.

- **Conditional Formatting**: The code includes conditional formatting to visually highlight positive and negative changes in the "Yearly Change" column, using green for positive changes and red for negative changes.

- **Finding Extremes**: After analyzing all the stocks, the code identifies the stocks with the "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume" by looping through the data again and tracking these values.

- **Displaying Results**: The identified extreme values and corresponding stock tickers are displayed in the worksheet.

This VBA code is designed to work on an Excel file containing multiple worksheets, each representing stock data for a different year. By running the `change_worksheet` subroutine, you can perform the analysis across all years' data at once.

## How to Run the Code

To use this VBA code, follow these steps:

1. Download the Excel file containing the stock data.
2. Open the Excel file.
3. Press `ALT` + `F11` to open the VBA editor.
4. Copy and paste both the `change_worksheet` and `stocks` subroutines into VBA modules.
5. Save your changes.
6. Run the `change_worksheet` subroutine to analyze all years' stock data.

Please note that this code assumes your Excel file structure matches the layout and data organization described in the code.
