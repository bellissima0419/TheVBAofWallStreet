Sub StockVolume()

  ' Create a variable to hold the Volume sum
  Dim TickerSymbol As String
  ' Dim volColumn As Integer
  ' volColumn = 7
  ' Dim TickerColum As Integer
  Dim TickerRow As Long

  TickerRow = 2
  ' TickerColum = 8
  ' Dim TotalColumn As Integer
  ' TotalColumn = 9

  Dim lastrow As Long
  ' counts the number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  MsgBox (lastrow)

  Dim VolumeSum As LongLong
  VolumeSum = 0

  Dim i As Long

  ' Loop through each row
  For i = 2 To lastrow

    ' Initially set the VolumeSum to be 0 for each row
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      TickerSymbol = Cells(i, 1).Value
      VolumeSum = VolumeSum + Cells(i, 7).Value

      Cells(TickerRow, 8).Value = TickerSymbol

      Cells(TickerRow, 9).Value = VolumeSum
      TickerRow = TickerRow + 1

      VolumeSum = 0

    Else
      VolumeSum = VolumeSum + Cells(i, 7).Value
    End If

  Next i


End Sub
Sub ClearContents()

' Clear Contents ticker and volume columns

    Range("H2:I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("H2").Select
End Sub
