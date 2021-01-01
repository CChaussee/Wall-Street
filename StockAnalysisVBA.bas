Sub StockData():
'Error Control
    On Error Resume Next
'To loop through all worksheets
Dim ws As Worksheet
For Each ws In Worksheets
'Make code run faster
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

'Creating new Columns with Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
'Defining Varaialbes
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryTableRow As Long
    Dim LastRowAll As Long
    Dim LastRowVolumes As Long
    Dim YearlyChange As Double
    Dim LastAmount As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim LastRowValue As Long
    Dim GreatestTotalVolume As Double
'Variable values
    TotalVolume = 0
    SummaryTableRow = 2
    LastAmount = 2
    GreatestIncrease = 0
    GreatestDecrease = 0
    LastRowAll = ws.Cells(Rows.Count, 1).End(xlUp).Row
    GreatestTotalVolume = 0
'Loop
    For i = 2 To LastRowAll
'Ticker
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & SummaryTableRow).Value = Ticker
            ws.Range("L" & SummaryTableRow).Value = TotalVolume
            TotalVolume = 0
'Opening Price
            OpenPrice = ws.Range("C" & LastAmount)
'Closing Price
            ClosePrice = ws.Range("F" & i)
'Yearly Change
            YearlyChange = ClosePrice - OpenPrice
            ws.Range("J" & SummaryTableRow).Value = YearlyChange
'Coloring of Cells
            If ws.Range("J" & SummaryTableRow).Value >= 0 Then
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            End If
'Percent Change
            If OpenPrice = 0 Then
                PercentChange = 0
                Else
                YearlyOpen = ws.Range("C" & LastAmount)
                PercentChange = YearlyChange / OpenPrice
            End If
'Formatting Percentage
            ws.Range("K" & SummaryTableRow).Value = PercentChange
            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"

'Add 1 to rows
            SummaryTableRow = SummaryTableRow + 1
            LastAmount = i + 1
                
        End If
        Next i
'Finding Greatest Changes Per Worksheet
        LastRowVolumes = ws.Cells(Rows.Count, 11).End(xlUp).Row
        For i = 2 To LastRowVolumes
'Greatest Increase
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
            End If
'Greatest Decrease
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
            End If
'Greatest Total Volume
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
            End If
         Next i
 'Formatting
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    Next ws

End Sub