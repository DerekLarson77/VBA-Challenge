Attribute VB_Name = "Greatest_Checker"
Sub Greatest_Stock(ws As Worksheet)

'Adding Header names.
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

Dim lastRow As Integer
Dim i As Integer
Dim Ticker As String
Dim Percent_Increase As Double
Dim Percent_Decrease As Double
Dim Stock_Vol As Variant

'Get last row number from table created in Range("I:L")
lastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row + 1

'Loop to get greatest percent increase/decrease and stock volume
For i = 2 To lastRow

    'Change percent increase if value is higher than previous variable
    If ws.Cells(i, 11).Value > Percent_Increase Then
        Percent_Increase = ws.Cells(i, 11).Value
    End If
    
    'Change percent decrease if value is lower than previous variable
    If ws.Cells(i, 11).Value < Percent_Decrease Then
        Percent_Decrease = ws.Cells(i, 11).Value
    End If
    
    'Change stock volume if value is higher than previous variable
    If ws.Cells(i, 12).Value > Stock_Vol Then
        Stock_Vol = ws.Cells(i, 12).Value
    End If

Next i

'loop to get Ticker name for our 3 variables gathered (percent increase, percent decrease, stock volume)
For i = 2 To lastRow

'If the percent increase is found then get ticker name and enter values in new table Range("O:Q")
    If ws.Cells(i, 11).Value = Percent_Increase Then
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = Percent_Increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
    End If
    
'If the percent decrease is found then get ticker name and enter values in new table Range("O:Q")
    If ws.Cells(i, 11).Value = Percent_Decrease Then
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 17).Value = Percent_Decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
    End If
    
'If the stock volume is found then get ticker name and enter values in new table Range("O:Q")
    If ws.Cells(i, 12).Value = Stock_Vol Then
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 17).Value = Stock_Vol
    End If

Next i

'Auto fits the O column width to the largest character sized cell
ws.Columns("O:O").AutoFit

End Sub
