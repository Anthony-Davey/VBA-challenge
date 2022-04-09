# VBA-challenge

The VBA-challenge involved testing my VBA proficiency using a large dataset of stock ticker price history over 3 years. The macro created will output 4 columns of summary aggregated data in the following format:

Ticker  |  Yearly Change (with conditional formatting  |  Percent Change (with number formatting to 0.00%)  | Total Stock Volume

The macro is designed using a for loop to iterate through all rows starting from row 2 (after the header) to a variable which determines the last row of the code to allow for seamless use of the macro regardless of the number of rows. An if statement tests for whether the ticker in the subsequent column is not equal to the prior cell's ticker. If true, then the macro will assign the last ticker cell's close price to the variable closeP. A variable set to 2 is used to determine the open price for each ticker and set each equal to openP. The yearly change column is then calculated using the formula:  year_change = closeP - openP
Percent change is determined with the following formula: percent_change = year_change / openP
The challenge stated that the percent change be displayed as a percentage with 2 decimal places, which was satisfied using the command: .NumberFormat = "0.00%"
The commands used to create conditional formatting within the yearly change column was: .Interior.ColorIndex = <color #>  'color numbers for green and red were fouund to be 4 and 3, respectively.
The total volume column was included in the else statement which would be True when the ticker was not changing from 1 cell to the next. This would allow me to sum the column using: total_volume = total_volume + Cells(i, 7).Value       'Cells(i, 7).Value target the individual cells within each iteration of the for loop within the total volume column
