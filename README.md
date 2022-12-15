## Module 2 Challenge by Erica Rojas Wood

### Calculation Subs
The script is set up with a main Sub that can call all the required Sub Modules for calculating
- Total Stock Volume per Ticker
- Yearly Change per Ticker
- Percentage Change per Ticker

The main sub is called calc_range_of_unique_ticker()

### get_unique_ticker_values()
This is a standalone Sub which given a list of Ticker Symbols in Column A, it will generate a new list of unique Ticker Symbols. This helps the other Sub Modules run, as they use this unique list in the loops.


### setup_data_formatting()
Formats the data in the Percentage Change column to a percentage.
