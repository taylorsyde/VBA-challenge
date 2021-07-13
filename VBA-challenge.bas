Attribute VB_Name = "Module1"
Sub Summary():
    
    'these varibables are to sort thru all worksheets
    Dim xSh As Worksheet
    Application.ScreenUpdating = False

    'this begins the outer loop of going thru the worksheets
    For Each xSh In Worksheets
    xSh.Select

    ' this prints out the headers for the summary table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"


    Dim sumcount As Integer
    Dim opening, closing, ydiff As Double
    Dim voltotal As Variant
    Dim cur_column, next_column As String
    
    Dim per, maxper, minper As Double
    Dim maxvol As Variant
    Dim maxname, minname, maxvolname As String
    

    'sumcount keeps track of the place you are in the summary table
    sumcount = 2
    voltotal = 0

    opening = Cells(2, 3).Value
    closing = 0
    ydiff = 0
    per = 0
    
    maxper = 0
    minper = 0
    maxvol = 0


    'counts the number of rows
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'this loop looks at the current cell and the next cell
    'if they are different, it prints to the summary table
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
            
                'grab the closing value, calculate difference
                'and print to summary table
                closing = Cells(i, 6).Value
                ydiff = closing - opening
                Cells(sumcount, 10).Value = ydiff
                
                    'formats the colors of the cells
                    If ydiff > 0 Then
                    Cells(sumcount, 10).Interior.ColorIndex = 4
                
                    Else: Cells(sumcount, 10).Interior.ColorIndex = 3
                    End If
            
                'calculate precent, print to summary table and format cell to %
                per = ydiff / opening
                Cells(sumcount, 11).Value = per
                Cells(sumcount, 11).NumberFormat = "0.00%"
            
                'BONUS max yearly change
                If per > maxper Then
                maxper = per
                maxname = cur_column_value
                End If
                
                If per < minper Then
                minper = per
                minname = cur_column_value
                End If
                
                If voltotal > maxvol Then
                maxvol = voltotal
                maxvolname = cur_column_value
                End If
            
                'reset the opening
                If Cells(i + 1, 3).Value = 0 Then
                opening = 1
                Else: opening = Cells(i + 1, 3).Value
                End If
            
                'move to the next row in the summary table
                sumcount = sumcount + 1
            
                'reset the voltotal
                voltotal = 0
            
                Else: voltotal = voltotal + Cells(i, 7).Value
            End If
                
        Next i
        
        'BONUS
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        Cells(1, 16).Value = "Ticker"
        Cells(2, 16).Value = maxname
        Cells(3, 16).Value = minname
        Cells(4, 16).Value = maxvolname
        
        Cells(1, 17).Value = "Value"
        Range("Q2:Q3").NumberFormat = "0.00%"
        Cells(2, 17).Value = maxper
        Cells(3, 17).Value = minper
        Cells(4, 17).Value = maxvol
        
    Next
    Application.ScreenUpdating = True
    
End Sub

