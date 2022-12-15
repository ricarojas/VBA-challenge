## Module 2 Challenge by Erica Rojas Wood

### Calculation Subs
The script is set up with a main Sub that can call all the required Sub Modules for calculating
- Total Stock Volume per Ticker (calc_total_volume)
- Yearly Change per Ticker (calc_percentage_change)
- Percentage Change per Ticker (calc_yearly_change)

The main sub is called calc_range_of_unique_ticker()

### get_unique_ticker_values()
This is a standalone Sub which given a list of Ticker Symbols in Column A, it will generate a new list of unique Ticker Symbols. This helps the other Sub Modules run, as they use this unique list in the loops.


### setup_data_formatting()
Formats the data in the Percentage Change column to a percentage.
Adds in conditional formatting for Yearly Change & Percentage Change.

### Recommended way to run
You can run the main Sub as mentioned in the Calculation Subs section, however it is best to only run each of the calcs individually via the main Sub.
To do this, simply remove the Call to 2 of the 3 Subs.


### Running against a Sheet that isn't 2018
The sheet is hardcoded to 2018, but that is simple enougn to change.

Fine **With Sheets("2018")** and replace 2018 with the sheet name you desire.
