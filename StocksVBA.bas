Attribute VB_Name = "Module1"
Sub StockLoop()


    ' Setting Worksheet Variables
    Dim columnheaders As Variant
    Dim StockWs As Worksheet
    Dim StockWb As Workbook
    
    Set StockWb = ActiveWorkbook
    
    ' Setting Header of Column Information
    columnheaders = Array("Ticker", "Date", "Open", "High", "Low", "Close", "Volume", " ", "Ticker", "Yearly Change", "Percent Change/Yr", "Total Stock Volume", " ", " ", " ", "Ticker", "Value")


End Sub
