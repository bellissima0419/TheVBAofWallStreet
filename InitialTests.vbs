Sub StockVolume()

  ' Create a variable to hold the Volume sum
  Dim VolumeSum As long
  Dim TickerSymbol as String
  Dim volColumn as Integer
  volColumn = 7
  Dim TickerColum as Integer
  Dim TickerRow as Long
  TickerRow = 2
  TickerColum = 8
  Dim TotalColumn as Integer
  TotalColumn = 9


  ' counts the number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  msgbox(lastrow)


  VolumeSum = 0


  ' Loop through each row
  For i = 2 To 300

    ' Initially set the VolumeSum to be 0 for each row
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      TickerSymbol = Cells(i, 1).Value
      VolumeSum = VolumeSum + cells(i, volColumn).value

      Cells(TickerRow,TickerColum).Value = TickerSymbol

      Cells(TickerRow, TotalColumn).Value = VolumeSum
      TickerRow = TickerRow + 1

      VolumeSum = 0

    Else
    VolumeSum = VolumeSum + Cells(i, volColumn).Value
    end If


  Next i

End Sub


' Sub StockVolume()

'   ' Create a variable to hold the Volume sum
'   Dim VolumeSum As Long
'   Dim Ticker As String
'   Dim lastrow As Long
'   ' counts the number of rows
'   lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'   MsgBox (lastrow)
'   VolumeSum = 0

'   ' Loop through each row
'   For i = 2 To lastrow

'     ' Initially set the VolumeSum to be 0 for each row
'     If (Cells(i, 1).Value = "A") Then
'       VolumeSum = VolumeSum + Cells(i, 7).Value
'     End If

'   Next i

'   Range("H2").Value = "A"
'   Range("I2").Value = VolumeSum



' End Sub
' ••••ˇˇˇˇ


' Sub StockVolume()

'   ' Create a variable to hold the Volume sum
'   Dim VolumeSum As Long
'   Dim Ticker As String
'   Dim VolumeColumn As Integer
'   VolumeColumn = 7
'   Dim TickerColum As Integer
'   Dim TickerRow As Integer
'   TickerRow = 2
'   TickerColum = 8
'   Dim Column As Integer
'   Column = 1

'   ' counts the number of rows
'   lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'   MsgBox (lastrow)
'   VolumeSum = 0

'   ' Loop through each row
'   For i = 2 To lastrow

'     ' Initially set the VolumeSum to be 0 for each row
'     If Cells(i + 1, Column).Value = Cells(i, Column).Value Then
'       VolumeSum = VolumeSum + VolumeColumn
'     End If
'     Cells(TickerRow, TickerColum).Value = Cells(TickerRow, Column)

'     Cells(TickerRow, VolumeColumn).Value = VolumeSum

'     ' If (Cells(i, 1).value = "A") Then
'     '   VolumeSum = VolumeSum + Cells(i,7).Value
'     ' end If

'   Next i

' End Sub
