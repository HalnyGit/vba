Attribute VB_Name = "ImmDatefinder"
Function Hny_ImmDateFinder(vInitDate As Variant)
'Return exact IMM date (which is third Wednesday of a month) for given month or contract
'vInitDate: date, or contract_name/year, eg. 01/06/2019, 06/2019, h/2020

Dim i As Integer
Dim dTempDate As Date
Dim iDay As Integer
Dim iCode As String
Dim iYear As Integer
Dim arrCodes() As Variant
Dim colCodes As New Collection

ReDim arrCodes(0 To 11) As Variant
arrCodes = Array("f", "g", "h", "j", "k", "m", "n", "q", "u", "v", "x", "z")

For i = 0 To 11
    colCodes.Add Item:=(i + 1), Key:=arrCodes(i)
Next i

If IsDate(vInitDate) Then
    dTempDate = DateSerial(Year(vInitDate), Month(vInitDate), 15)
    Do
    iDay = Weekday(dTempDate, vbMonday)
    If iDay <> 3 Then dTempDate = dTempDate + 1
    Loop Until iDay = 3
    Hny_ImmDateFinder = dTempDate
Else
    iCode = LCase(Left(vInitDate, 1))
    iYear = CInt(Right(vInitDate, 4))
    dTempDate = DateSerial(iYear, colCodes(iCode), 15)
    Do
    iDay = Weekday(dTempDate, vbMonday)
    If iDay <> 3 Then dTempDate = dTempDate + 1
    Loop Until iDay = 3

    Hny_ImmDateFinder = dTempDate

End If

 

End Function
