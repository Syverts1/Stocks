Attribute VB_Name = "Module1"
Sub sumvol():

Dim tickrow As Integer
Dim totalvol As Double
Dim lastrow As Long
Dim openprice As Double
Dim closeprice As Double
Dim pricechange As Double
Dim lastrow2 As Long

tickrow = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"

totalvol = 0
'initialize the sum ticker value, and the openprice (so we don't have to compare a double to a string)
Range("I2") = Cells(tickrow, 1).Value
openprice = Range("C2").Value

   
   For i = 2 To lastrow
        
        'see if the tick row value is different than the one below it
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'set the close price to whats in the F column
            closeprice = Cells(i, 6).Value
            'calculate the change in prices by subtracting the close price from the open price
            pricechange = closeprice - openprice
            Cells(tickrow, 10) = pricechange
            
            'set cell background to green if price change is positive, set red if negative
            If pricechange >= 0 Then
                Range("J" & tickrow).Interior.Color = RGB(155, 187, 89)
            
            Else
                Range("J" & tickrow).Interior.Color = RGB(192, 80, 77)
            
            End If
            
            'calculate the percent change
            If openprice = 0 Then
            Cells(tickrow, 11) = 0
            Else
            Cells(tickrow, 11) = pricechange / openprice
            End If
            Cells(tickrow, 11).Style = "Percent"
            'set the open price value of the next ticker
            openprice = Cells(i + 1, 3).Value
            'place the total volume variable into the sum ticker volume cell
            Cells(tickrow, 12).Value = totalvol + Cells(i, 7)
            'move down a tickrow
            tickrow = tickrow + 1
            
            'initialize the total volume variable with what is in the <vol> column
            totalvol = 0

            'set the ticker value to what is in column 1
            Cells(tickrow, 9).Value = Cells(i + 1, 1).Value


        Else
            'add what is in the tickers volume to the running total volume
            If openprice = 0 Then
            openprice = Cells(i + 1, 3).Value
            End If
            totalvol = totalvol + Cells(i, 7).Value
        
            
        End If
    Next i
    
lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row

maxincr = WorksheetFunction.Max(Range("K2:K" & lastrow2))
Range("Q2") = maxincr
maxdecr = WorksheetFunction.Min(Range("K2:K" & lastrow2))
Range("Q3") = maxdecr
Range("Q2:Q3").Style = "Percent"
maxvol = WorksheetFunction.Max(Range("L2:L" & lastrow2))
Range("Q4") = maxvol

For i = 2 To lastrow2
    
    If Cells(i, 11) = maxincr Then
        Range("P2") = Cells(i, 9)
    
    End If
    
    If Cells(i, 11) = maxdecr Then
        Range("P3") = Cells(i, 9)
    
    End If
    
    If Cells(i, 12) = maxvol Then
        Range("P4") = Cells(i, 9)

    End If
Next i

End Sub

Sub reset()
Dim lastrow As Integer

lastrow = Cells(Rows.Count, 12).End(xlUp).Row

Range("I1: Q" & lastrow) = ""
Range("I1: Q" & lastrow).Interior.ColorIndex = 0

End Sub

Sub runall()
    Dim current As Worksheet
    For Each current In Application.Worksheets
        current.Activate
        Call sumvol
    Next
End Sub

Sub resetall()
    Dim current As Worksheet
    For Each current In Application.Worksheets
        current.Activate
        Call reset
    Next
End Sub


