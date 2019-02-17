Sub MediumHomework()
Dim currentticker As String
Dim currentvolume As Double
Dim I As Double
Dim stockcount As Double
Dim LastRow As Double
Dim openingvalue As Double
Dim endvalue As Double
Dim changevalue As Double
Dim changeperc As Double
stockcount = 1
currentvolume = 0
openingvalue = Range("C2").Value
endvalue = 0
changevalue = 0
changeperc = 0


With ActiveSheet
    LastRow = Cells(.Rows.Count, "A").End(xlUp).Row 'Find the last used row in a column
End With
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Volume"
currentticker = Cells(2, 1).Value

'MsgBox (currentticker) 'Trying to measure whether or not the commands were properly assigning values
For I = 2 To LastRow
    If Cells(I + 1, 1).Value <> currentticker Then
        Cells(stockcount + 1, 9).Value = currentticker
        Cells(stockcount + 1, 12).Value = currentvolume
        stockcount = stockcount + 1
        currentticker = Cells(I + 1, 1).Value
        currentvolume = 0
        endvalue = Cells(I, 6).Value
        changevalue = endvalue - openingvalue
        Cells(stockcount, 10).Value = changevalue
        If openingvalue = 0 Then
            changeperc = 0
            Else
            changeperc = endvalue / openingvalue - 1
            End If
        Cells(stockcount, 11).Value = changeperc
        changeperc = 0
        changevalue = 0
        openingvalue = Cells(I + 1, 3).Value
        Else
        currentticker = Cells(I, 1).Value
        currentvolume = currentvolume + Cells(I, 7).Value
    End If
Next I

End Sub




