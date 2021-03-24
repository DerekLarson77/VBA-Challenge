Attribute VB_Name = "Ticker_Collection"
Sub Check_Tickers()

'Calls a subroutine to clear all data added from macros, so there can be a clean slate.
Call Clear_Creation

'Creates the 4 headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Counters:  i for rows in the raw data, j for rows in our data creation in columns I-L.
Dim i As Long
Dim j As Integer

i = 2
j = 2

'A Do While loop to go through column A until there is no more data on the sheet.
Do While Cells(i, 1).Value <> ""

Call Ticker_Year(i, j)

i = i + 1
j = j + 1
Loop




End Sub

Function Ticker_Year(i As Long, j As Integer)

Dim Ticker_Name As String
Dim Year_start As Double
Dim Year_end As Double
Dim Stock_Vol As Variant

'Variables we want hard set before the loop begins.
''The variable i means the row is predetermined before this function is run.
''''The name of the ticker is the first value in Column A.
''''The opening price at the beginning of the year is the first value in Column C.
''''The stock volume needs to be reset to 0, so it can be added each time the loop is run later.
Ticker_Name = Cells(i, 1).Value
Year_start = Cells(i, 3).Value
Stock_Vol = 0

'A Do While loop to run while the tracker is still the same in column A.
Do While Ticker_Name = Cells(i, 1).Value

'Each row we continue to add the volume in column G.
Stock_Vol = Stock_Vol + Cells(i, 7).Value

'Advancing our counter by 1, so the loop can progress to the next row.
i = i + 1
Loop

''  The loop ended because a new ticker name appeared.
''  In order to get the closing on the last day of the year we must go back up one row to get our Year end value.
Year_end = Cells(i - 1, 6).Value

'Used for checking what the code calculates the variables at this point in time.  Uncomment to run and the values will display in Immediate window.
'Debug.Print Ticker_Name
'Debug.Print Year_start
'Debug.Print Year_end
'Debug.Print Stock_Vol

Cells(j, 9).Value = Ticker_Name
Cells(j, 10).Value = Year_end - Year_start
Cells(j, 11).Value = (Year_end - Year_start) / Year_start
    Cells(j, 11).NumberFormat = "0.00%"
Cells(j, 12).Value = Stock_Vol

Stock_Vol = 0

End Function


