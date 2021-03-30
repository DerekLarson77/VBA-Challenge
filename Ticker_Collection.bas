Sub Check_Tickers(ws As Worksheet)

'Calls a subroutine to clear all data added from macros, so there can be a clean slate.
Call Clear_Creation(ws)

'Creates the 4 headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Counters:  i for rows in the raw data, j for rows in our data creation in columns I-L.
Dim i As Long
Dim j As Integer

i = 2
j = 2

'A Do While loop to go through column A until there is no more data on the sheet.
Do While ws.Cells(i, 1).Value <> ""

Call Ticker_Year(ws, i, j)

i = i + 1
j = j + 1
Loop


End Sub