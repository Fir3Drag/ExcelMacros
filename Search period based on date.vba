Sub SearchPeriod()
    Dim yearStartDate As Date
    
    Dim p1StartDate As Date
    Dim p1EndDate As Date
    Dim p2StartDate As Date
    Dim p2EndDate As Date
    Dim p3StartDate As Date
    Dim p3EndDate As Date
    Dim p4StartDate As Date
    Dim p4EndDate As Date
    Dim p5StartDate As Date
    Dim p5EndDate As Date
    Dim p6StartDate As Date
    Dim p6EndDate As Date
    Dim p7StartDate As Date
    Dim p7EndDate As Date
    Dim p8StartDate As Date
    Dim p8EndDate As Date
    Dim p9StartDate As Date
    Dim p9EndDate As Date
    Dim p10StartDate As Date
    Dim p10EndDate As Date
    Dim p11StartDate As Date
    Dim p11EndDate As Date
    Dim p12StartDate As Date
    Dim p12EndDate As Date
    Dim p13StartDate As Date
    Dim p13EndDate As Date
    
    Dim columnBCell As Range
    Dim incrementValue As Integer
    Dim yearEnd As Integer
    
    'Calculating the range of dates
    yearStartDate = #3/3/2024#
    yearEnd = 25
    
    p1StartDate = yearStartDate
    p1EndDate = yearStartDate + 27
    p2StartDate = p1EndDate + 1
    p2EndDate = p1EndDate + 27
    p3StartDate = p2EndDate + 1
    p3EndDate = p2EndDate + 27
    p4StartDate = p3EndDate + 1
    p4EndDate = p3EndDate + 27
    p5StartDate = p4EndDate + 1
    p5EndDate = p4EndDate + 27
    p6StartDate = p5EndDate + 1
    p6EndDate = p5EndDate + 27
    p7StartDate = p6EndDate + 1
    p7EndDate = p6EndDate + 27
    p8StartDate = p7EndDate + 1
    p8EndDate = p7EndDate + 27
    p9StartDate = p8EndDate + 1
    p9EndDate = p8EndDate + 27
    p10StartDate = p9EndDate + 1
    p10EndDate = p9EndDate + 27
    p11StartDate = p10EndDate + 1
    p11EndDate = p10EndDate + 27
    p12StartDate = p11EndDate + 1
    p12EndDate = p11EndDate + 27
    p13StartDate = p12EndDate + 1
    p13EndDate = p12EndDate + 27
    
    'Declaring variables
    incrementValue = 1
    Set columnBCell = Range("B1")
    lastRowNum = ActiveSheet.Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    For Each columnACell In Range("A1:A" & lastRowNum).Cells
        If IsDate(columnACell) Then
            If columnACell >= p1StartDate And columnACell <= p1EndDate Then
            columnBCell.Value = "P01-" & yearEnd
            End If
            
            If columnACell >= p2StartDate And columnACell <= p2EndDate Then
                columnBCell.Value = "P02-" & yearEnd
            End If
            
            If columnACell >= p3StartDate And columnACell <= p3EndDate Then
                columnBCell.Value = "P03-" & yearEnd
            End If
            
            If columnACell >= p4StartDate And columnACell <= p4EndDate Then
                columnBCell.Value = "P04-" & yearEnd
            End If
                
            If columnACell >= p5StartDate And columnACell <= p5EndDate Then
                columnBCell.Value = "P05-" & yearEnd
            End If
                
            If columnACell >= p6StartDate And columnACell <= p6EndDate Then
                columnBCell.Value = "P06-" & yearEnd
            End If
                
            If columnACell >= p7StartDate And columnACell <= p7EndDate Then
                columnBCell.Value = "P07-" & yearEnd
            End If
                
            If columnACell >= p8StartDate And columnACell <= p8EndDate Then
                columnBCell.Value = "P08-" & yearEnd
            End If
                
            If columnACell >= p9StartDate And columnACell <= p9EndDate Then
                columnBCell.Value = "P09-" & yearEnd
            End If
                
            If columnACell >= p10StartDate And columnACell <= p10EndDate Then
                columnBCell.Value = "P10-" & yearEnd
            End If
                
            If columnACell >= p11StartDate And columnACell <= p11EndDate Then
                columnBCell.Value = "P11-" & yearEnd
            End If
                
            If columnACell >= p12StartDate And columnACell <= p12EndDate Then
                columnBCell.Value = "P12-" & yearEnd
            End If
                
            If columnACell >= p13StartDate And columnACell <= p13EndDate Then
                columnBCell.Value = "P13-" & yearEnd
            End If
        End If
        
        'incrementing the ranges
        incrementValue = incrementValue + 1
        Set columnBCell = Range("B" & incrementValue)
        Next columnACell
            
End Sub
