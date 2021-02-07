# vba-Assignment
Sub VBAHomework()
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Cells(1, 11).Value = "Ticker"
        ws.Cells(1, 12).Value = "Yearly Change"
        ws.Cells(1, 13).Value = "Percentage Change"
        ws.Cells(1, 14).Value = "Total Stock Volume"
    
    Dim ticker As String
    Dim yearstart As Double
        yearstart = 0
    Dim yearend As Double
        yearend = 0
    Dim YearlyChange As Double
        YearlyChange = 0
    Dim PercentageChange As Double
        PercentageChange = 0
    Dim TotalStockVolume As Double
        TotalStockVolume = 0
    Dim Table As Integer
        Table = 2
        
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
    If yearstart = 0 Then
    
        yearstart = ws.Cells(i, 3).Value
        
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    
    End If
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ticker = ws.Cells(i, 1).Value
            ws.Range("K" & Table).Value = ticker
            
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            ws.Range("N" & Table).Value = TotalStockVolume
        
        yearstart = ws.Cells(i, 3).Value
        yearend = ws.Cells(i, 6).Value
        
        YearlyChange = yearend - yearstart
            ws.Range("L" & Table).Value = YearlyChange
        
        Table = Table + 1
        TotalStockVolume = 0
        
    Else
    
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
    End If
    
    If YearlyChange > 0 Then
    
        ws.Cells(i, 12).Interior.ColorIndex = 4
    
    Else
    
        ws.Cells(i, 12).Interior.ColorIndex = 3
    
    End If
    
    If yearstart = 0 Then
    
        PercentageChange = 0
        
    Else
    
         PercentageChange = (YearlyChange / yearstart) * 100
            ws.Range("M" & Table).Value = PercentageChange
            
    End If
    
    yearstart = 0
    
    Next i
    
    ws.Columns("K:N").AutoFit
    
    Next ws
    
End Sub
