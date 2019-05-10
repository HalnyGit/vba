Attribute VB_Name = "ModFollDate"
Function Hny_CalcModFollDate(dInitDate As Date, sInterval As String, iNumInterval As Integer, _
                        Optional rHolCcy1 As Variant, Optional rHolCcy2 As Variant) As Date

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
'rHolCcy1: optional range that contains list of holidays for currency 1
'rHolCcy2: optional range that contains list of holidays for currency 2

Dim dEndDate As Date
Dim dTempDate As Date
Dim iHolidays() As Variant
Dim iMonth As Integer
Dim iWeekDay As Integer
Dim iEoM As Integer: iEoM = 0
Dim rTempRange As Range
Dim c As Range
Dim i As Integer

If IsMissing(rHolCcy1) And IsMissing(rHolCcy2) Then
    ReDim iHolidays(0)
ElseIf IsMissing(rHolCcy2) Then
    ReDim iHolidays(rHolCcy1.Rows.Count)
    i = 0
    For Each c In rHolCcy1
        iHolidays(i) = c.Value * 1
        i = i + 1
    Next
Else
    ReDim iHolidays(rHolCcy1.Rows.Count + rHolCcy2.Rows.Count)
    Set rTempRange = Union(rHolCcy1, rHolCcy2)
    i = 0
    For Each c In rTempRange
        iHolidays(i) = c.Value * 1
        i = i + 1
    Next
End If

sInterval = LCase(sInterval)
dEndDate = DateAdd(sInterval, iNumInterval, dInitDate)
iMonth = Month(dEndDate)

'Checking if dInitDate is last date of the month
dTempDate = dInitDate + 1
    If Month(dTempDate) > Month(dInitDate) And (sInterval = "m" Or sInterval = "yyyy") Then iEoM = 1
        
 Select Case iEoM
    Case 1
        dTempDate = DateAdd(sInterval, iNumInterval, dInitDate)
        dEndDate = WorksheetFunction.EoMonth(dTempDate, 0)
        Do
            dEndDate = WorksheetFunction.WorkDay(dEndDate + 1, -1, iHolidays)
            iWeekDay = WorksheetFunction.Weekday(dEndDate, 2)
        Loop Until iWeekDay <> 6 And iWeekDay <> 7
    Case Else
        dEndDate = WorksheetFunction.WorkDay(dEndDate - 1, 1, iHolidays)
        iWeekDay = WorksheetFunction.Weekday(dEndDate, 2)

        'Adjusting for Following
        Do
            dEndDate = WorksheetFunction.WorkDay(dEndDate - 1, 1, iHolidays)
            iWeekDay = WorksheetFunction.Weekday(dEndDate, 2)
        Loop Until iWeekDay <> 6 And iWeekDay <> 7

        'Adjusting for Modified
        If sInterval = "m" Or sInterval = "yyyy" Then
            If iMonth <> Month(dEndDate) Then
                Do
                    dEndDate = WorksheetFunction.WorkDay(dEndDate, -1, iHolidays)
                    iWeekDay = WorksheetFunction.Weekday(dEndDate, 2)
                Loop Until iWeekDay <> 6 And iWeekDay <> 7
            End If
        End If
End Select

Hny_CalcModFollDate = dEndDate

End Function


