Attribute VB_Name = "Main_RunAllSheets"
Sub AllSheets()

Application.ScreenUpdating = False

Dim ws As Worksheet

For Each ws In Worksheets

    Call Check_Tickers(ws)
    Call Greatest_Stock(ws)

Next ws

Application.ScreenUpdating = True
End Sub
