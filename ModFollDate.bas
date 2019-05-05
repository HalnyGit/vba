Attribute VB_Name = "Module1"
Option Explicit

Function Hny_ModFollDate(dInitDate As Date, sInterval As String, iNumInterval As Integer, rHolidays As Range) As Date

'Return the date that is (iNumInterval * sInterval) after dInitDate
'dInitDate: initial calculation date
'sInterval: argument can have these settings:
'    yyyy   Year
'    q      Quarter
'    m      Month
'    y      Day of year
'    d      Day
'    w      Weekday
'    ww     Week
'    h      Hour
'    n      Minute
'    s      Second
'iNumInterval: number of intervals
'rHolidays: argument is range that contains list of holidays, if you have no list of holidays provide empty cell for that argument

Dim dEndDate As Date
Dim dNextDate As Date
Dim dTempDate As Date
Dim dMonth As Integer
Dim iWeekDay As Integer
Dim iEoM As Integer: iEoM = 0

sInterval = LCase(sInterval)
dEndDate = DateAdd(sInterval, iNumInterval, dInitDate)
dMonth = Month(dEndDate)

'Checking if dInitDate is last date of the month
dNextDate = dInitDate + 1
    If Month(dNextDate) > Month(dInitDate) And (sInterval = "m" Or sInterval = "yyyy") Then iEoM = 1
        
 Select Case iEoM
    Case 1
        dTempDate = DateAdd(sInterval, iNumInterval, dInitDate)
        dEndDate = WorksheetFunction.EoMonth(dTempDate, 0)
        Do
            dEndDate = WorksheetFunction.WorkDay(dEndDate + 1, -1, rHolidays)
            iWeekDay = WorksheetFunction.Weekday(dEndDate, 2)
        Loop Until iWeekDay <> 6 And iWeekDay <> 7
    Case Else
        dEndDate = WorksheetFunction.WorkDay(dEndDate - 1, 1, rHolidays)
        iWeekDay = WorksheetFunction.Weekday(dEndDate, 2)

        'Adjusting for Following
        Do
            dEndDate = WorksheetFunction.WorkDay(dEndDate - 1, 1, rHolidays)
            iWeekDay = WorksheetFunction.Weekday(dEndDate, 2)
        Loop Until iWeekDay <> 6 And iWeekDay <> 7

        'Adjusting for Modified
        If sInterval = "m" Or sInterval = "yyyy" Then
            If dMonth <> Month(dEndDate) Then
                Do
                    dEndDate = WorksheetFunction.WorkDay(dEndDate, -1, rHolidays)
                    iWeekDay = WorksheetFunction.Weekday(dEndDate, 2)
                Loop Until iWeekDay <> 6 And iWeekDay <> 7
            End If
        End If
End Select

Hny_ModFollDate = dEndDate

End Function



