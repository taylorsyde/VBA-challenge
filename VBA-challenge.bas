Attribute VB_Name = "Module1"
Sub Tickernames():

' this prints out the headers for the summary table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"


'sumcount keeps track of the place you are in the summary table
Dim sumcount As Integer
sumcount = 2
    
'this loop looks a cell and the previous cell, if they are different, it prints to the summary table
    For i = 2 To 58239
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            Cells(sumcount, 9).Value = Cells(i, 1).Value
            sumcount = sumcount + 1
        End If
    Next i
End Sub


