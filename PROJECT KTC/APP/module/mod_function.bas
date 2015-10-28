Attribute VB_Name = "mod_function"
Public Function getEndDay(ByVal srcmonth As Integer, Optional srcthn As Integer) As String
    Dim h1 As String
'    h1 = Format(srcDate, "mm")
    h1 = srcmonth
    On Error GoTo err
    Select Case h1
        Case Is = 1: getEndDay = 31
        Case Is = 2: getEndDay = Day(h1 & "/29/" & srcthn)
        Case Is = 3: getEndDay = 31
        Case Is = 4: getEndDay = 30
        Case Is = 5: getEndDay = 31
        Case Is = 6: getEndDay = 30
        Case Is = 7: getEndDay = 31
        Case Is = 8: getEndDay = 31
        Case Is = 9: getEndDay = 30
        Case Is = 10: getEndDay = 31
        Case Is = 11: getEndDay = 30
        Case Is = 12: getEndDay = 31
    End Select
    h1 = ""
    Exit Function
err:
        If err.Number = 13 Then getEndDay = 28: h1 = "" 'Day if encounter not a left-year
End Function

Public Function roundDown(dblValue As Double) As Double
On Error GoTo PROC_ERR
Dim myDec As Long

myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
If myDec > 0 Then
     roundDown = CDbl(Left(CStr(dblValue), myDec))
Else
     roundDown = dblValue
End If

PROC_EXIT:
     Exit Function
PROC_ERR:
     MsgBox err.Description, vbInformation, "Round Down"
End Function

Public Function roundUp(dblValue As Double) As Double
On Error GoTo PROC_ERR
Dim myDec As Long

myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
If myDec > 0 Then
     roundUp = CDbl(Left(CStr(dblValue), myDec)) + 1
Else
     roundUp = dblValue
End If

PROC_EXIT:
     Exit Function
PROC_ERR:
     MsgBox err.Description, vbInformation, "Round Up"
End Function

