Attribute VB_Name = "Module1"
Sub Tickernames():

' this prints out the headers for the summary table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"


Dim sumcount As Integer
Dim opening, closing As Double
Dim total As Variant
Dim cur_column, next_column As String


'sumcount keeps track of the place you are in the summary table
sumcount = 2
voltotal = 0

opening = Cells(2, 3).Value
closing = 0


'counts the number of rows
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
'this loop looks a cell and the next cell, if they are different, it prints to the summary table
    For i = 2 To lastrow
    
    'keeps variable for current and next column
    cur_column_value = Cells(i, 1).Value
    next_column_value = Cells(i + 1, 1).Value
    
            If cur_column_value <> next_column_value Then
            'print the name of the current stock
            Cells(sumcount, 9).Value = cur_column_value
            
            'add the final vol and print
            voltotal = voltotal + Cells(i, 7).Value
            Cells(sumcount, 12).Value = voltotal
            
            'grab the closing value, calculate difference and print to summary
            closing = Cells(i, 6).Value
            Cells(sumcount, 10).Value = closing - opening
            
            'reset the opening
            opening = Cells(i + 1, 3).Value
            
            'move to the next row in the summary table
            sumcount = sumcount + 1
            
            'reset the voltotal
            voltotal = 0
            
            Else: voltotal = voltotal + Cells(i, 7).Value
        End If
        
        
    Next i
End Sub

