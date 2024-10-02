Sub Receipt()
    Dim columnBCell As Range
    Dim columnCCell As Range
    Dim columnDCell As Range
    Dim newTableRowNumber As Integer

    Set columnBCell = Range("B1")
    Set columnCCell = Range("C1")
    Set columnDCell = Range("D1")
    newTableRowNumber = 2
    
    lastRowNum = ActiveSheet.Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    For Each columnACell In Range("A2:A" & lastRowNum).Cells
        'Checks if the current cell starts with £ or -£
        
        If InStr(1, columnACell.Value, "£") = 1 Or InStr(1, columnACell.Value, "£") = 2 Then
            'set the values
            columnBCell.Value = columnACell.Offset(-2, 0).Value 'Due to the blank line each row you need to go up 2 lines
            columnCCell.Value = columnACell.Value
            
            If InStr(1, columnCCell.Value, "A") Then
                columnDCell.Value = "A"
                'strip the A
                columnCCell.Value = Replace(columnCCell.Value, "A", "")
            End If
            
            If InStr(1, columnCCell.Value, "C") Then
                columnDCell.Value = "C"
                'strip the A
                columnCCell.Value = Replace(columnCCell.Value, "C", "")
            End If
            
            'strip the £
            columnCCell.Value = Replace(columnCCell.Value, "£", "")
            columnCCell.Value = Replace(columnCCell.Value, Chr(160), "")
            columnCCell.Value = Trim(columnCCell.Value)
            
            'incrementing the ranges
            Set columnBCell = Range("B" & newTableRowNumber)
            Set columnCCell = Range("C" & newTableRowNumber)
            Set columnDCell = Range("D" & newTableRowNumber)
            newTableRowNumber = newTableRowNumber + 1
        End If
    Next columnACell
End Sub