Attribute VB_Name = "Module1"
Sub StockLoop()


    ' Setting Worksheet Variables
    Dim columnheaders As Variant
    Dim StockWs As Worksheet
    Dim StockWb As Workbook
    
    Set StockWb = ActiveWorkbook
    
    ' Setting Header of Column Information
    columnheaders = Array("Ticker", "Date", "Open", "High", "Low", "Close", "Volume", " ", "Ticker", "Yearly Change", "Percent Change/Yr", "Total Stock Volume", " ", " ", " ", "Ticker", "Value")
    
     ' Header Array Loop
    For Each StockWs In StockWb.Sheets
        With StockWs
        .Rows(1).Value = ""
        For i = LBound(columnheaders) To UBound(columnheaders)
        .Cells(1, 1 + i).Value = columnheaders(i)
        
        'Setting Row 1 Font and Centering
        Next i
        .Rows(1).Font.Bold = True
        .Rows(1).HorizontalAlignment = xlCenter
        End With
    
    Next StockWs


End Sub
