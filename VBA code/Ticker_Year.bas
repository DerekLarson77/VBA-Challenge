Function Ticker_Year(ws As Worksheet, i As Long, j As Integer)

Dim Ticker_Name As String
Dim Year_start As Double
Dim Year_end As Double
Dim Stock_Vol As Variant
Dim Percent_Change As Variant

'Variables we want hard set before the loop begins.
''The variable i means the row is predetermined before this function is run.
''''The name of the ticker is the first value in Column A.
''''The opening price at the beginning of the year is the first value in Column C.
''''The stock volume needs to be reset to 0, so it can be added each time the loop is run later.
Ticker_Name = ws.Cells(i, 1).Value
Year_start = ws.Cells(i, 3).Value
Stock_Vol = 0


'A Do While loop to run while the tracker is still the same in column A.
Do While Ticker_Name = ws.Cells(i, 1).Value

'Each row we continue to add the volume in column G.
Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value

'Advancing our counter by 1, so the loop can progress to the next row.
i = i + 1
Loop

''  The loop ended because a new ticker name appeared.
''  In order to get the closing on the last day of the year we must go back up one row to get our Year end value.
Year_end = ws.Cells(i - 1, 6).Value

'Used for checking what the code calculates the variables at this point in time.  Uncomment to run and the values will display in Immediate window.
'Debug.Print Ticker_Name
'Debug.Print Year_start
'Debug.Print Year_end
'Debug.Print Stock_Vol

ws.Cells(j, 9).Value = Ticker_Name
ws.Cells(j, 10).Value = Year_end - Year_start

If Year_start = 0 Then
    ws.Cells(j, 11).Value = Year_end
Else
    Percent_Change = (Year_end - Year_start) / Year_start
    ws.Cells(j, 11).Value = Percent_Change
    
    'Positive Percent Changes are filled with Green.
    'Negative Percent Changes are filled with Red.
    'Zero Percent Changes are left with fill White.
    If Percent_Change > 0 Then
         ws.Cells(j, 11).Interior.ColorIndex = 4
    ElseIf Percent_Change < 0 Then
         ws.Cells(j, 11).Interior.ColorIndex = 3
    End If
    
End If

ws.Cells(j, 11).NumberFormat = "0.00%"
ws.Cells(j, 12).Value = Stock_Vol

Stock_Vol = 0

End Function
