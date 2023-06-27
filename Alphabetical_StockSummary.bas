Attribute VB_Name = "Module1"
Sub StockSummary()
For Each ws In Worksheet(2)

    Dim SummaryTableRow As Integer
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim Change As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    OpenValue = ws.Range("C2").Value
    TotalStockVolume = 0
    SummaryTableRow = 2
    For i = 2 To LastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'Calculate summary table values
            CloseValue = ws.Range("F" & i).Value
            Change = CloseValue - OpenValue
            PercentChange = Change / OpenValue
            'Assign variables to summary table
            ws.Range("I" & SummaryTableRow) = ws.Cells(i, 1).Value
            ws.Range("J" & SummaryTableRow) = Change
            If Change > 0 Then
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            ElseIf Change < 0 Then
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            End If
            ws.Range("K" & SummaryTableRow) = FormatPercent(PercentChange)
            ws.Range("L" & SummaryTableRow) = TotalStockVolume
            'Reset variables
            TotalStockVolume = 0
            OpenValue = ws.Range("C" & i + 1).Value
            SummaryTableRow = SummaryTableRow + 1
        Else
            TotalStockVolume = TotalStockVolume + ws.Range("G" & i).Value
        End If
            
    Next i
    
    Dim LargestIncrease As Double
    Dim LargestDecrease As Double
    Dim LargestTotal As Double
    Dim IncTicker As String
    Dim DecTicker As String
    Dim TotalTicker As String
    
    LargestIncrease = 0
    LargestDecrease = 0
    LargestTotal = 0
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To LastRow
        'Determine if each value is greater than previous
        If ws.Range("K" & i).Value > LargestIncrease Then
            LargestIncrease = ws.Range("K" & i).Value
            IncTicker = ws.Range("I" & i).Value
        ElseIf ws.Range("K" & i).Value < LargestDecrease Then
            LargestDecrease = ws.Range("K" & i).Value
            DecTicker = ws.Range("I" & i).Value
        End If
        If ws.Range("L" & i).Value > LargestTotal Then
            LargestTotal = ws.Range("L" & i).Value
            TotalTicker = ws.Range("I" & i).Value
        End If
    Next i
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("P2").Value = IncTicker
    ws.Range("P3").Value = DecTicker
    ws.Range("P4").Value = TotalTicker
    ws.Range("Q2").Value = FormatPercent(LargestIncrease)
    ws.Range("Q3").Value = FormatPercent(LargestDecrease)
    ws.Range("Q4").Value = LargestTotal
Next ws
End Sub
