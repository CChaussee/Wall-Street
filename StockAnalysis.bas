Sub StockData():
'Error Control
    On Error Resume Next
'To loop through all worksheets
Dim ws As Worksheet
For Each ws In Worksheets
'Make code run faster(Thank you Terra)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
'Creating new Columns with Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
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
            ws.Range("I").Value = Ticker
            ws.Range("L").Value = TotalVolume
            TotalVolume = 0
'Opening Price (Stack Overflow helped explain how to hold values for math later)
            OpenPrice = ws.Range("C" & LastAmount)
'Closing Price
            ClosePrice = ws.Range("F" & i)
'Yearly Change
            YearlyChange = ClosePrice - OpenPrice
            ws.Range("J").Value = YearlyChange
'Coloring of Cells (Numbers are from googling how to assign color of cells)
            If ws.Range("J").Value >= 0 Then
            ws.Range("J").Interior.ColorIndex = 4
                Else
                ws.Range("J").Interior.ColorIndex = 3
            End If
'Percent Change
            If OpenPrice = 0 Then
                PercentChange = 0
                Else
                YearlyOpen = ws.Range("C")
                PercentChange = YearlyChange / OpenPrice
            End If
'Formatting Percentage
            ws.Range("K").NumberFormat = "0.00%"
'Add 1 to rows and move to next i
            SummaryTableRow = SummaryTableRow + 1
            LastAmount = i + 1
        End If
        Next i
'Finding Greatest Changes Per Worksheet
        LastRowVolumes = ws.Cells(Rows.Count, 11).End(xlUp).Row
        For i = 2 To LastRowVolumes
'Greatest Increase
            If ws.Range("K").Value > ws.Cells(2, 17).Value Then
                ws.Cells(2, 17).Value = ws.Range("K").Value
                ws.Cells(2, 16).Value = ws.Range("I").Value
            End If
'Greatest Decrease
            If ws.Range("K").Value < ws.Cells(3, 17).Value Then
                ws.Cells(3, 17).Value = ws.Range("K").Value
                ws.Cells(3, 16).Value = ws.Range("I").Value
            End If
'Greatest Total Volume
            If ws.Range("L").Value > ws.Cells(4, 17).Value Then
                ws.Cells(4, 17).Value = ws.Range("L").Value
                ws.Cells(4, 16).Value = ws.Range("I").Value
            End If
         Next i
 'Formatting
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 18).NumberFormat = "0.00%"

    Next ws

End Sub