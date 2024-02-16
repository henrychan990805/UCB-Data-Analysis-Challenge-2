Attribute VB_Name = "Module1"
Sub StockAnnualAnalysis()
Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Columns("I").ColumnWidth = 10
        ws.Columns("J").ColumnWidth = 14
        ws.Columns("K").ColumnWidth = 15
        ws.Columns("L").ColumnWidth = 16
        ws.Columns("O").ColumnWidth = 20
        ws.Columns("P").ColumnWidth = 12
        ws.Columns("Q").ColumnWidth = 12
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).NumberFormat = "0.00E+00"
        Dim Start As Long
        Dim Last As Long
        Start = 2
        Last = ws.Cells(Rows.Count, 1).End(xlUp).Row
        PrintRow = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Value"
        Dim YearlyChange As Double
        YearlyChange = 0
        Dim PercentChange As Double
        PercentChange = 0
        Dim TotalStockValue As Double
        TotalStockValue = 0
        Dim GreatestTotal As Double
        GreatestTotal = 0
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        For i = Start To Last
        TotalStockValue = TotalStockValue + ws.Cells(i, 7).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(PrintRow, 9).Value = ws.Cells(i, 1).Value
            YearlyChange = ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value
            ws.Cells(PrintRow, 10) = YearlyChange
            If YearlyChange > 0 Then
                ws.Cells(PrintRow, 10).Interior.ColorIndex = 4
            ElseIf YearlyChange < 0 Then
                ws.Cells(PrintRow, 10).Interior.ColorIndex = 3
            End If
            PercentChange = YearlyChange / ws.Cells(Start, 3).Value
            ws.Cells(PrintRow, 11) = PercentChange
            If PercentChange > 0 Then
                ws.Cells(PrintRow, 11).Interior.ColorIndex = 4
            ElseIf PercentChange < 0 Then
                ws.Cells(PrintRow, 11).Interior.ColorIndex = 3
            End If
            ws.Cells(PrintRow, 12) = TotalStockValue
            If GreatestTotal < TotalStockValue Then
                GreatestTotal = TotalStockValue
                ws.Cells(4, 17).Value = GreatestTotal
                ws.Cells(4, 16).Value = ws.Cells(PrintRow, 9).Value
            End If
            If GreatestIncrease < PercentChange Then
                GreatestIncrease = PercentChange
                ws.Cells(2, 17).Value = GreatestIncrease
                ws.Cells(2, 16).Value = ws.Cells(PrintRow, 9).Value
            End If
            If GreatestDecrease > PercentChange Then
                GreatestDecrease = PercentChange
                ws.Cells(3, 17).Value = GreatestDecrease
                ws.Cells(3, 16).Value = ws.Cells(PrintRow, 9).Value
            End If
            Start = i + 1
            PrintRow = PrintRow + 1
            TotalStockValue = 0
        End If
        Next i
    Next
End Sub
