# VBA-Yearly-Stocks

	The goal of this challenge was to extract the first day of the year and last day of the year for each ticker(stock) to calculate the growth and percent change.
The stock volume from each day was also added up to get the total stock per year for each ticker.  This process was then completed for 3 seperate years.
	There was a bonus to determine which ticker had the highest growth, the lowest growth, and the largest stock volume for each year.

	This was completed by having 5 functions which are included in the 5 script files.

1)  Main_RunAllSheets
	A For loop to go through all worksheets in the workbook and then call the Ticker_Collection and Greatest_Checker functions for each worksheet.

2)  Ticker_Collection
	A Do While loop to go down the rows until there was no longer a value in column A.
	Two variables(i and j) are saved.  First to keep track of the ticker row for the Do While loop.
	The second was to keep track of what row we were creating new values in columns i through l.
	Ticker_Year would be called in each instance of the Do While loop and would collect both i and j variables.

3)  Ticker_Year
	A Do While loop that ended when the ticker name changed.
	Stock volume was added for each row during the loop.
	Beginning of year value was saved and End of year value was gathered at the last instance of the loop.
	The ticker name, growth, percent change and stock volume was entered into columns i through l and then the j variable increased.

4)  Greatest_Checker
	Two For loops.  The first to determine the highest growth, lowest growth, and the largest stock volume.
	The second loop went back through to get the ticker name for each of the variables determined in the first loop.

5)  Clear_Values
	This was just to clear the columns where all the ticker information was pasted before each new run attempt.
