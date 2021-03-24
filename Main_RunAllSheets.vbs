Attribute VB_Name = "Main_RunAllSheets"
Sub AllSheets()

Application.ScreenUpdating = False

Sheets("2014").Activate
Call Check_Tickers

Sheets("2015").Activate
Call Check_Tickers

Sheets("2016").Activate
Call Check_Tickers

Application.ScreenUpdating = True
End Sub
