Attribute VB_Name = "Module1"
Sub StockAnalysis():

    'declare variable
    Dim ticker As String
    Dim StockVolume As Double
    Dim TotalStockVolume As Double
    Dim ws As Worksheet
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    
    For Each ws In Worksheets

    OpeningPrice = ws.Cells(2, 3).Value
    TotalStockVolume = 0
    j = 2
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'For loop
    For i = 2 To RowCount

    StockVolume = ws.Cells(i, 7).Value
    TotalStockVolume = TotalStockVolume + StockVolume

    'Define variables
    If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
    ticker = ws.Cells(i, 1).Value
    ClosingPrice = ws.Cells(i, 6).Value
    
    YearlyChange = ClosingPrice - OpeningPrice
    PercentageChange = YearlyChange / OpeningPrice
 
    'Highlight positive or negative yearly change.
     If YearlyChange >= 0 Then
     ws.Cells(j, 10).Interior.ColorIndex = 4
     Else
     ws.Cells(j, 10).Interior.ColorIndex = 3
     End If

    'Write Variable onto worksheet
    ws.Cells(j, 9).Value = ticker
    ws.Cells(j, 10).Value = YearlyChange
    ws.Cells(j, 11).Value = PercentageChange
    ws.Cells(j, 11).NumberFormat = "0.00%"
    ws.Cells(j, 12) = TotalStockVolume
    
    TotalStockVolume = 0
    
    j = j + 1
        
    OpeningPrice = ws.Cells(i + 1, 3).Value
    
    End If
    Next i
    Next ws


End Sub



