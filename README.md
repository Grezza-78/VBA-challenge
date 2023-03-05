# VBA-challenge
Bootcamp Module 2 Challenge - VBA

This assignment has asked to me create a Macro that will run through the "Multiple Year Stock Data" spreadsheet, that contains the multiple stock data points for the 2018, 2019, 2020 worksheets. 

Once the Macro has been run it will output the following information for each of the of worksheets contained in the spreadsheet:

    * The ticker symbol for each stock item

    * Yearly change from the opening price at the beginning a  given year to the closing price at the end of that year for that ticker - apply conditional formatting to identify positive or negative returns for each ticker

    * The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year for that ticker.

    * The total stock volume of the stock traded for that ticker for that year.

In addition the Marco will also identify and list the ticker for each year contained in the spreadsheet with:
    * The Greatest % increase for that year 
    * The Greatest % decrease for that year
    * The Greatest total volume traded for that year

The Macro has been designed by analysing the key challenges in producing the output (as outlined above), and then scripting individual Subs to solve for those challenges. 

Each different Sub has been described within the VBA scripting file:
    
    * The Main Sub that creates the report on any of the years data out(spreadsheet) is called "StockReport" which calls all the other subs that deals with addressing the various challenges analysed
    
    * ReportsTable - generates the headers and space where the data output will be printed
    
    * TickerIdentifier - identifies the all the individual tickers in each of year and prints it
    
    * Yearly_Percent_Change - calculates the yearly and percentage change for each ticker and prints it
    
    * YearlyChangeCheck - calculates and prints the positive or negative price change for each ticker, and then applies conditional formatting accordingly with postivie change being green and negative change being red
    
    * Total_Stock_Vol - calculates the total years stock volume for each ticker and prints it
    
    * GreatestIncrease - identifies the biggest ticker percentage increase for the year and prints it
    
    * GreatestDecrease - identifies the biggest ticker percentage decrease for the year and prints it
    
    * GreatestVolume - identifies the biggest ticker volume traded for the year and prints it

The Parent Sub "SRonallWorksheets" is the sub that initiates the "StockReport" across all the worksheets contained in the date file... and this is the Sub what should be used to initiate the Macro for analysing the "Multiple Year Stock Data" spreadsheet. 

