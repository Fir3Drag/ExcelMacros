Sub SearchPeriod()
    Dim columnBCell As Range
    Dim incrementValue As Integer
    Dim yearEnd As Integer

    Dim yearStartDate As Date
    Dim PeriodStartDateList(13) As Date
    Dim PeriodEndDateList(13) As Date
    
    'Calculating the range of dates
    yearStartDate = #3/3/2024#    'american format #mm/dd/yy#
    yearEnd = 25
    
    'Declaring variables
    incrementValue = 1
    Set columnBCell = Range("B1")
    lastRowNum = ActiveSheet.Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    For index = 1 To 13
        PeriodStartDateList(index) = yearStartDate + ((index - 1) * 28)
        Debug.Print PeriodStartDateList(index)
        Next index
        
    For index = 1 To 13
        PeriodEndDateList(index) = yearStartDate + (index * 27)
        Debug.Print PeriodEndDateList(index)
        Next index
    
    For Each columnACell In Range("A1:A" & lastRowNum).Cells
        If IsDate(columnACell) Then
            For index = 1 To 13
                If columnACell >= PeriodStartDateList(index) And columnACell <= PeriodEndDateList(index) Then
                    If index > 9 Then
                        columnBCell.Value = "P" & index & "-" & yearEnd
                    Else
                        columnBCell.Value = "P0" & index & "-" & yearEnd
                    End If
                End If
                Next index
            End If
                
        'incrementing the ranges
        incrementValue = incrementValue + 1
        Set columnBCell = Range("B" & incrementValue)
        Next columnACell
            
End Sub
