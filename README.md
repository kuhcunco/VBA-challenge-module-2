# Multiple Year Stock Data VBA Script

## Overview

The **Multiple Year Stock Data VBA Script** is designed to analyze and process stock data for multiple years within an Excel workbook. This script calculates yearly performance metrics for stock tickers, making it easier to identify trends over time.

## Features

- **Multi-year stock analysis**: Process and analyze stock data for multiple years.
- **Yearly performance calculations**: Calculate the yearly opening price, closing price, yearly change, and percent change for each stock ticker.
- **Volume summary**: Compute the total volume of trades for each stock ticker within each year.

## Requirements

- **Excel**: The script is designed to run in Microsoft Excel.
- **VBA (Visual Basic for Applications)**: The script should be executed using the VBA editor within Excel.

## How to Use

1. **Open Excel**: Make sure your stock data is in an Excel sheet. The columns should include at least the following:
   - `Ticker`: The stock ticker symbol.
   - `Date`: The date of the stock entry.
   - `Open`: The stock’s opening price.
   - `Close`: The stock’s closing price.
   - `Volume`: The number of shares traded.

2. **Open the VBA Editor**:
   - Press `Alt + F11` to open the VBA editor in Excel.

3. **Insert the VBA Script**:
   - In the VBA editor, insert a new module (`Insert > Module`) and paste the `Multiple_year_stock_data` VBA script into the module.

4. **Run the Script**:
   - Press `F5` to run the script, or use the “Run” button within the VBA editor.

5. **Output**:
   - The script will generate a summary table, displaying:
     - Ticker
     - Yearly Change (closing price - opening price)
     - Percent Change (percentage change of the yearly stock price)
     - Total Volume for each stock ticker over the selected year.

## License

This project is open-source and free to use.
