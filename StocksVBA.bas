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
    
    ' Loop through all worksheets
        For Each StockWs In Worksheets
        
        ' Setting variables for calculations
            Dim Company_Ticker As String
            Company_Ticker = " "
            Dim Total_Volume As Double
            Total_Volume = 0
            Dim Open_Price As Double
            Open_Price = 0
            Dim Close_Price As Double
            Close_Price = 0
            Dim Yearly_Price_Change As Double
            Yearly_Price_Change = 0
            Dim Yearly_Price_Percent As Double
            Yearly_Price_Percent = 0
            Dim Max_Company_Ticker As String
            Max_Company_Ticker = " "
            Dim Min_Company_Ticker As String
            Min_Company_Ticker = " "
            Dim Max_Percent As Double
            Max_Percent = 0
            Dim Min_Percent As Double
            Min_Percent = 0
            Dim Max_Volume_Company_Ticker As String
            Max_Volume_Company_Ticker = " "
            Dim Max_Volume As Double
            Max_Volume = 0
            
            Dim Chart_Row As Long
            Chart_Row = 2
            
            Dim Lastrow As Long
            
            Lastrow = StockWs.Cells(Rows.Count, 1).End(xlUp).Row
            
            Open_Price = StockWs.Cells(2, 3).Value
            
            For i = 2 To Lastrow
            
            If StockWs.Cells(i + 1, 1).Value <> StockWs.Cells(i, 1).Value Then
                
                ' Ticker Name Starting Point
                Company_Ticker = StockWs.Cells(i, 1).Value
                
                Close_Price = StockWs.Cells(i, 6).Value
                Yearly_Price_Change = Close_Price - Open_Price
                
                ' Zero Value Condition Set
                If Open_Price <> 0 Then
                Yearly_Price_Percent = (Yearly_Price_Change / Open_Price) * 100
                    
                End If
                
                ' Add Total Volume
                Total_Volume = Total_Volume + StockWs.Cells(i, 7).Value
                
                ' Ticker Name -> Summary Table Column I
                StockWs.Range("I" & Chart_Row).Value = Company_Ticker
                
                ' Yearly Price Change -> Column J
                StockWs.Range("J" & Chart_Row).Value = Yearly_Price_Change
                
                If (Yearly_Price_Change > 0) Then
                    StockWs.Range("J" & Chart_Row).Interior.ColorIndex = 4
                    
                ElseIf (Yearly_Price_Change <= 0) Then
                
                    StockWs.Range("J" & Chart_Row).Interior.ColorIndex = 3
                End If
                
                ' Yearly Price Change Converted to Percent -> Summary Table Column K
                StockWs.Range("K" & Chart_Row).Value = (CStr(Yearly_Price_Percent) & "%")
                
                ' Total Stock Volume -> Summary Table Column L
                StockWs.Range("L" & Chart_Row).Value = Total_Volume
                
                ' Add 1 to Summary Table Row Count
                Chart_Row = Chart_Row + 1
                
                ' Get the Next Beginning Price
                Open_Price = StockWs.Cells(i + 1, 3).Value
                
                If (Yearly_Price_Percent > Max_Percent) Then
                    Max_Percent = Yearly_Price_Percent
                    Max_Company_Ticker = Company_Ticker
                    
                ElseIf (Yearly_Price_Percent < Min_Percent) Then
                    Min_Percent = Yearly_Price_Percent
                    Min_Company_Ticker = Company_Ticker
                    
                End If
                
                If (Total_Volume > Max_Volume) Then
                    Max_Volume = Total_Volume
                    Max_Volume_Company_Ticker = Company_Ticker
                    
                End If
                
                Yearly_Price_Percent = 0
                Total_Volume = 0
                
            Else
            
                Total_Volume = Total_Volume + StockWs.Cells(i, 7).Value
                
            End If
        
        Next i
                
        ' Setting results to be filled
                StockWs.Range("Q2").Value = (CStr(Max_Percent) & "%")
                StockWs.Range("Q3").Value = (CStr(Min_Percent) & "%")
                StockWs.Range("P2").Value = Max_Company_Ticker
                StockWs.Range("P3").Value = Min_Company_Ticker
                StockWs.Range("P4").Value = Max_Volume_Company_Ticker
                StockWs.Range("Q4").Value = Max_Volume
                StockWs.Range("O2").Value = "Max % Increase"
                StockWs.Range("O3").Value = "Max % Decrease"
                StockWs.Range("O4").Value = "Max Total Volume"
    
    Next StockWs


End Sub
