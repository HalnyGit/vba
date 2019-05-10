Attribute VB_Name = "RangesToArray"
Option Explicit

Function RangesToArray(dInitDate As Date, Optional a1 As Variant, Optional a2 As Variant) As Date

Dim rNewRange As Range
Dim aArray() As Variant
Dim dEndDate As Date
Dim c As Range
Dim i As Integer

If IsMissing(a1) And IsMissing(a2) Then
    ReDim aArray(1)
    aArray(0) = 0
ElseIf IsMissing(a2) Then
    ReDim aArray(a1.Rows.Count)
    i = 0
    For Each c In a1
        aArray(i) = c.Value * 1
        i = i + 1
    Next
Else
    Set rNewRange = Union(a1, a2)
    ReDim aArray(a1.Rows.Count + a2.Rows.Count)
    i = 0
    For Each c In rNewRange
        aArray(i) = c.Value * 1
        i = i + 1
    Next
End If

RangesToArray = Application.WorksheetFunction.WorkDay(dInitDate, 1, aArray)

End Function

