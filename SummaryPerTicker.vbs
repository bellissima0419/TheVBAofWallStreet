Sub StockVolume()

  ' Create a variable to hold the Volume sum
  ' Dim TickerSymbol As String
  ' Dim volColumn As Integer
  ' volColumn = 7
  ' Dim TickerColum As Integer
  ' Dim TickerRow As LongLong
  ' TickerRow = 2
  ' TickerColum = 8
  ' Dim TotalColumn As Integer
  ' TotalColumn = 9

  ' counts the number of rows
  ' MsgBox (LastRow)

  ' Loop through each worksheet
  For Each ws In Worksheets
    Dim TickerSymbol As String
    Dim TickerRow As LongLong
    TickerRow = 2

    Dim VolumeSum As LongLong
    ' VolumeSum = 0

    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range("H1").Value = "Ticker"
    ws.Range("I1").Value = "Total Stock Volume"

      ' Loop through each row
    Dim i As Long

    For i = 2 To LastRow

      ' Initially set the VolumeSum to be 0 for each row
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        TickerSymbol = ws.Cells(i, 1).Value
        VolumeSum = VolumeSum + ws.Cells(i, 7).Value

        ws.Cells(TickerRow, 8).Value = TickerSymbol

        ws.Cells(TickerRow, 9).Value = VolumeSum
        TickerRow = TickerRow + 1

        VolumeSum = 0

      Else
        VolumeSum = VolumeSum + ws.Cells(i, 7).Value
      End If

    Next i

      ' ws.Cells(TickerRow, 8).Value = TickerSymbol
      ' ws.Cells(TickerRow, 9).Value = VolumeSum
      ' TickerRow = TickerRow + 1
      ' VolumeSum = 0

  Next ws

End Sub


Sub DeleteHeadersTickerVolume()
    For Each ws In Worksheets
        ws.Range("H1:I1").ClearContents
        ws.Range("HH:II").ClearContents
    Next ws
End Sub


Sub ClearContents2()

' ClearContents2 Macro
' clear contents columns ticker and volume
    For Each ws In Worksheets
        ws.Range("H:I").ClearContents
    Next ws
End Sub
