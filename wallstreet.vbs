Attribute VB_Name = "WallStreet"

Sub StockStats()
  Dim tickerName As String
  Dim tickerOpen As Double
  Dim tickerClose As Double
  Dim volume As Double
  Dim percentageChange As Double
  Dim yearlyChange As Double
  Dim greatestTickerIncrease As String
  Dim greatestIncrease As Double
  Dim greatestTickerDecrease As String
  Dim greatestDecrease As Double
  Dim greatestVolume As Double
  Dim greatestTickerVolume As String
  Dim outputRow As Integer

'Loop to traverse all spreadsheet in Workbook
  For Each ws In Worksheets
    'Initialize variables start of each sheet
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    tickerName = ws.Range("A2").Value
    tickerOpen = ws.Range("C2").Value
    volume = 0
    outputRow = 2
    'Clear Summary Contents
    ws.Range("J:Q").ClearContents
    ws.Range("J:Q").Interior.Pattern = xlNone
    
    'Summary Column Headers
    ws.Range("J1:Q1").Font.FontStyle = "Bold"
    ws.Range("J1:M1").NumberFormat = "@"
    ws.Range("L:L").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("J1:K1").Font.FontStyle = "Bold"
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    
    'Worksheet logic
    For i = 2 To lastRow
      'Ticker Name change
      If ws.Cells(i + 1, 1).Value <> tickerName Then
        tickerClose = ws.Cells(i, 6).Value
        ws.Cells(outputRow, 10).Value = tickerName
        ws.Cells(outputRow, 11).Value = tickerClose - tickerOpen
        If tickerOpen = 0 Then
           ws.Cells(outputRow, 12).Value = 0
        Else
           percentageChange = (tickerClose - tickerOpen) / tickerOpen
           ws.Cells(outputRow, 12).Value = percentageChange
           If percentageChange > 0 Then
             ws.Cells(outputRow, 11).Interior.ColorIndex = 4
           ElseIf percentageChange < 0 Then
             ws.Cells(outputRow, 11).Interior.ColorIndex = 3
             ws.Cells(outputRow, 11).Font.FontStyle = "Bold"
           End If
        End If
        ws.Cells(outputRow, 13).Value = volume + ws.Cells(i, 7)
        
        'Intialize greatest increase/decrease/volume variable.
        'Value sets to first ticker process summary values
        If outputRow = 2 Then
          greatestIncrease = percentageChange
          greatestDecrease = percentageChange
          greatestTickerIncrease = tickerName
          greatestTickerDecrease = tickerName
          greatestVolume = volume + ws.Cells(i, 7)
          greatestTickerVolume = tickerName
        'Each ticker iteration summary metrics compared to variables
        ElseIf percentageChange > greatestIncrease Then
          greatestIncrease = percentageChange
          greatestTickerIncrease = tickerName
        ElseIf percentageChange < greatestDecrease Then
          greatestDecrease = percentageChange
          greatestTickerDecrease = tickerName
        End If
        
       'Check to see if better ticker name volume greater than last.
       If outputRow > 2 And volume + ws.Cells(i, 7) > greatestVolume Then
         greatestVolume = volume + ws.Cells(i, 7)
         greatestTickerVolume = tickerName
       End If
        tickerName = ws.Cells(i + 1, 1).Value
        tickerOpen = ws.Cells(i + 1, 3).Value
        outputRow = outputRow + 1
        volume = 0
    'No ticker name change
     Else
        volume = volume + ws.Cells(i, 7)
     End If
    Next i
    
   'After traversing workseheet set summay metrics
   'Set up bonus cells with greatest percentage increase/decrease and volume values.
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("P2").Value = greatestTickerIncrease
    ws.Range("Q2").Value = greatestIncrease
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("P3").Value = greatestTickerDecrease
    ws.Range("Q3").Value = greatestDecrease
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P4").Value = greatestTickerVolume
    ws.Range("Q4").Value = greatestVolume
  Next ws
  
End Sub
