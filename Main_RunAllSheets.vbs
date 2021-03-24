Attribute VB_Name = "Main_RunAllSheets"
Sub AllSheets()

Sheets("2014").Activate
Call Check_Tickers

Sheets("2015").Activate
Call Check_Tickers

Sheets("2016").Activate
Call Check_Tickers

End Sub
