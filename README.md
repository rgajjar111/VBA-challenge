# VBA-challenge
**Here's an explanation of the key components and logic in the code LoopThroughStocksForMultipleSheets():**

- The macro loops through a predefined array of worksheet names, representing different years (e.g., "2018", "2019", "2020").

- It initializes variables and collections to store and manipulate data.

**For each year's worksheet:**

- It identifies the last row of data in the worksheet.
- Loops through each row to identify unique stock ticker symbols.
- Calculates the yearly change, percentage change, and total volume traded for each unique ticker symbol.
- Outputs the results to the specified output columns.

**The macro applies conditional formatting to highlight specific data points:**
- Yearly change is color-coded, with red indicating a decrease and green indicating an increase.
- Percentage change is color-coded similarly.
- The colors are applied based on percentiles: red for lower values, green for positive values, and white for values around the median.

**Here's an explanation of the key components and logic in the code SummaryForMultipleSheets():**

- The macro loops through a predefined array of worksheet names, representing different years (e.g., "2018", "2019", "2020").

- It initializes variables and collections to store and manipulate data, including the greatest increase, greatest decrease, and greatest volume.

**For each year's worksheet:**

- It identifies the last row of data in the worksheet.
- Loops through each row to identify unique stock ticker symbols.
- Calculates the yearly change, percentage change, and total trading volume for each unique ticker symbol.
- Updates the greatest increase, greatest decrease, and greatest volume values and corresponding ticker symbols.
- The macro outputs the results in the same worksheet for each year, displaying the stock symbols and metrics in columns.
