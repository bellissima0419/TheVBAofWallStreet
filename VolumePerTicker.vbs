Sub StockVolume()

  ' Loop through each worksheet
  For Each ws In Worksheets
    Dim TickerSymbol As String
    Dim TickerRow As LongLong
    TickerRow = 2

    Dim VolumeSum As LongLong

    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"

      ' Loop through each row
    Dim i As Long

    For i = 2 To LastRow

      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        TickerSymbol = ws.Cells(i, 1).Value
        VolumeSum = VolumeSum + ws.Cells(i, 7).Value

        ws.Cells(TickerRow, 9).Value = TickerSymbol

        ws.Cells(TickerRow, 10).Value = VolumeSum
        TickerRow = TickerRow + 1

        VolumeSum = 0

      Else
        VolumeSum = VolumeSum + ws.Cells(i, 7).Value
      End If

    Next i

  Next ws

End Sub


Sub ClearContents2()

' clear contents columns ticker and volume
    For Each ws In Worksheets
        ws.Range("I:J").ClearContents
    Next ws
End Sub
