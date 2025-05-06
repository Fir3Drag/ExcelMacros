Sub GetLastWeekdayOfMonth()
    Dim inputDate As Date
    Dim lastDay As Date
    Dim currentDay As Date
    Dim isWeekend As Boolean
    Dim firstRun As Boolean
    
    ' Modifable Variables
    currentNum = 1  ' START DATE, set this as the start month number (between 1 - 12)
    currentYear = 2025  ' START DATE, set this as the start year
    numOfResults = 24  ' AMOUNT OF MONTHS
    holidayDates = Array(#1/1/2025#, #4/18/2025#, #4/21/2025#, #5/5/2025#, #5/26/2025#, #8/25/2025#, #12/25/2025#, #12/26/2025#, #1/1/2026#, #4/3/2026#, #4/6/2026#, #5/25/2026#, #8/31/2026#, #12/25/2026#, #12/28/2026#, #1/1/2027#, #3/26/2027#, #3/29/2027#, #5/3/2027#, #5/31/2027#, #8/30/2027#, #12/27/2027#, #12/28/2027#)  ' Add holidays days to this, American format
    
    Dim j
    For j = 1 To numOfResults
        If currentNum > 12 Then  ' resets the currentNum and year updates
            currentYear = currentYear + 1
            currentNum = 1
        End If
        
        ' Variables
        inputDate = DateSerial(currentYear, currentNum, 1)
        lastDay = WorksheetFunction.EoMonth(inputDate, 0)  ' Find the last day of the month
        currentDay = lastDay  ' Start checking from the last day of the month
        firstLastDay = True ' This is set to false after the first correct result so it runs twice to give the 2nd last day
        
        Do
            Dim i
            isPublicHoliday = False
            
            For i = LBound(holidayDates) To UBound(holidayDates)  ' checks if the day is in the array of holiday dates
                If holidayDates(i) = currentDay Then
                    isPublicHoliday = True
                End If
            Next i
        
            isWeekend = (Weekday(currentDay, vbMonday) > 5) ' Check if it's a weekend, Returns True if Saturday or Sunday
            
            If Not isWeekend And Not isPublicHoliday Then
                If Not firstLastDay Then
                        Debug.Print (currentDay)
                        Exit Do
                    End If
                    
                firstLastDay = False
            End If
            
            currentDay = currentDay - 1
            
        Loop Until currentDay < inputDate
     
    currentNum = currentNum + 1
    Next j
End Sub