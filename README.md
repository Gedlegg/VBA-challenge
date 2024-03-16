# VBA-challenge

# Explanation of Script Creation: (My VBA Script Has two subroutines Stock_Data and Reset)

# Retrieval of Data:

    - The script used a loop to iterate through each row of stock data.

    - From each row, it will extract and store the following values:

            - Ticker symbol: The script will read the ticker symbol from the corresponding column in each row.

            - Volume of stock: It will retrieve the volume of stock traded for that particular ticker symbol.

            - Open price: The script will extract the opening price of the stock for each row.

            - Close price: Similarly, it will obtain the closing price of the stock for each row.

# Column Creation:

    - After extracting the data, the script will create the following columns: "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume" to store this information.

# Conditional Formatting:

    - Conditional formatting applied to get the yearly change and percent change columns to visually represent the data.

    - For the yearly change column, the script will apply formatting to highlight positive changes in green and negative changes in red.

    - Similarly, the percent change column will have appropriate conditional formatting to represent positive and negative changes.

# Calculated Values:

    - The script will calculate three important values: greatest % increase, greatest % decrease, and greatest total volume.

# Looping Across Worksheet:

    - The VBA script designed to run on all sheets in the workbook.
