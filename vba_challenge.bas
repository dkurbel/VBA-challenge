Attribute VB_Name = "Module1"
Sub vba_challenge()

Application.ScreenUpdating = False

' create headers for Ticker, Yearly Change, Percent Change, and Total Stock Volume
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' declare variables for Ticker, Yearly Change, Percent Change, and Total Stock Volume
Dim ticker As String
Dim totalvol As Double
Dim openprice As Single
Dim closingprice As Single

' declare variables for [next empty row]
Dim new_row As Long
new_row = CLng(2)

' create for loop to go through rows
For i = 2 To 22771

    ' check if this is the first instance of Ticker value
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then

        ' set the new ticker symbol
        ticker = CStr(Cells(i, 1).Value)
        
        ' insert the new Ticker symbol into a new row
        Cells(new_row, 9).Value = ticker

        ' take in the first open price
        openprice = Cells(i, 3).Value

        ' add the first day volume to [totalvol]
        totalvol = Cells(i, 7).Value

    ' check if this is the last instance of Ticker value
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        ' take in the closing price
        closingprice = Cells(i, 6).Value

        ' insert the yearchange and perchange values
        Cells(new_row, 10).Value = closingprice - openprice
        Cells(new_row, 10).NumberFormat = "$0.00"
            If Cells(new_row, 10).Value <= 0 Then
                Cells(new_row, 10).Interior.ColorIndex = 3
            Else
                Cells(new_row, 10).Interior.ColorIndex = 4
            End If
        Cells(new_row, 11).Value = (closingprice - openprice) / openprice
        Cells(new_row, 11).NumberFormat = "0.00%"

        ' add the day volume to [totalvol]
        totalvol = totalvol + Cells(i, 7).Value

        ' insert the total volume
        Cells(new_row, 12).Value = totalvol

        ' set the next [new_row] value
        new_row = new_row + 1
     
    ' if it's neither the first nor last Ticker instance, just add the volume to the total
    Else
        totalvol = totalvol + Cells(i, 7).Value
    
    End If
    
Next i

Application.ScreenUpdating = True

End Sub


