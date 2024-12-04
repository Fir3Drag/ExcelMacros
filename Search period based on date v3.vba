Sub SearchPeriod()
    Dim yearStartDate As Date
    Dim numOfPeriods As Integer
    
    Dim periodStartDate(1 To 39) As Date 'CHANGE RANGE HERE
    Dim periodEndDate(1 To 39) As Date 'CHANGE RANGE HERE
    
    Dim columnBCell As Range
    Dim incrementValue As Integer
    Dim yearEnd As Integer
    Dim periodNum As Integer
    Dim yearStartEnd As Integer
    
    'Calculating the range of dates
    yearStartDate = #3/5/2023#    'american format #mm/dd/yyyy#
    yearStartEnd = 24
    numOfPeriods = 39 'CHANGE RANGE HERE
    
    'Declaring variables
    periodNum = 1
    incrementValue = 1
    Set columnBCell = Range("B1")
    lastRowNum = ActiveSheet.Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    For i = 1 To numOfPeriods
        periodStartDate(i) = yearStartDate + ((i - 1) * 28)
        periodEndDate(i) = yearStartDate + ((i - 1) * 28) + 27
        
        Debug.Print "Start: " & periodStartDate(i)
        Debug.Print "End: " & periodEndDate(i)
    Next i
    
    For Each columnACell In Range("A1:A" & lastRowNum).Cells
        yearEnd = yearStartEnd
    
        If IsDate(columnACell) Then
            For i = 1 To numOfPeriods
                If columnACell >= periodStartDate(i) And columnACell <= periodEndDate(i) Then
                    If periodNum > 9 Then
                        columnBCell.Value = "P" & periodNum & "-" & yearEnd
                    Else
                        columnBCell.Value = "P0" & periodNum & "-" & yearEnd
                    End If
                End If
                    
                periodNum = periodNum + 1
                
                If periodNum > 13 Then
                    periodNum = periodNum - 13
                    yearEnd = yearEnd + 1
                    
                End If
            Next i
        End If
                
        'incrementing the ranges
        incrementValue = incrementValue + 1
        Set columnBCell = Range("B" & incrementValue)
    Next columnACell
End Sub
