Sub StockVolume()

  ' Loop through each worksheet
  For Each ws In Worksheets
    Dim TickerSymbol As String
    Dim TickerRow As LongLong
    TickerRow = 2

    Dim VolumeSum As LongLong

    Dim LastRow As Long

    'save last row number
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'add Ticker and total stock volume headers on all sheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"

      ' Loop through each row
    Dim i As Long

    For i = 2 To LastRow
      'Compare adjacent cells to determine if they hold the same or different value. If different value, then a new row is created in the new 2 columns.
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        TickerSymbol = ws.Cells(i, 1).Value
        VolumeSum = VolumeSum + ws.Cells(i, 7).Value

        ws.Cells(TickerRow, 9).Value = TickerSymbol

        ws.Cells(TickerRow, 10).Value = VolumeSum
        TickerRow = TickerRow + 1

        VolumeSum = 0

      Else
        'When the values are the same, keep adding to the VolumeSum counter
        VolumeSum = VolumeSum + ws.Cells(i, 7).Value
      End If

    Next i

  Next ws

End Sub


Sub ClearContents2()

' clear contents of ticker and volume columns in all sheets for testing purposes
    For Each ws In Worksheets
        ws.Range("I:J").ClearContents
    Next ws
End Sub
