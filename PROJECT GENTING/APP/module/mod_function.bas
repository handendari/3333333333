Attribute VB_Name = "mod_function"
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
 
Private Const MF_BYPOSITION = &H400&
 
Public Sub RemoveButtonX(frm As Form)
    Dim hSysMenu As Long
    hSysMenu = GetSystemMenu(frm.hWnd, 0)
    Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
    Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub

Public Function getEndDay(ByVal srcmonth As Integer, Optional srcthn As Integer) As String
    Dim h1 As String
'    h1 = Format(srcDate, "mm")
    h1 = srcmonth
    On Error GoTo Err
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
Err:
        If Err.Number = 13 Then getEndDay = 28: h1 = "" 'Day if encounter not a left-year
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
     MsgBox Err.Description, vbInformation, "Round Down"
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
     MsgBox Err.Description, vbInformation, "Round Up"
End Function

Public Function Terbilang(x As Double) As String
Dim tampung As Double
Dim teks As String
Dim bagian As String
Dim i As Integer
Dim tanda As Boolean
Dim letak(5)
    
    letak(1) = "Ribu "
    letak(2) = "Juta "
    letak(3) = "Milyar "
    letak(4) = "Trilyun "

    If (x = 0) Then
        Terbilang = "Nol"
        Exit Function
    End If
    
    If (x < 2000) Then
        tanda = True
    End If
    
    teks = ""
    
    If (x >= 1E+15) Then
        Terbilang = "Nilai terlalu besar"
    Exit Function
    
    End If
    
    For i = 4 To 1 Step -1
        tampung = Int(x / (10 ^ (3 * i)))
        If (tampung > 0) Then
            bagian = ratusan(tampung, tanda)
            If bagian = "Se" Then
                teks = teks & bagian & LCase(letak(i))
            Else
                teks = teks & bagian & letak(i)
            End If
        End If
    
        x = x - tampung * (10 ^ (3 * i))
    Next
    
    teks = teks & ratusan(x, False)
    Terbilang = teks
End Function

Function ratusan(ByVal y As Double, ByVal flag As Boolean) As String
Dim tmp As Double
Dim bilang As String
Dim bag As String
Dim j As Integer
Dim angka(9)

    angka(1) = "Se"
    angka(2) = "Dua "
    angka(3) = "Tiga "
    angka(4) = "Empat "
    angka(5) = "Lima "
    angka(6) = "Enam "
    angka(7) = "Tujuh "
    angka(8) = "Delapan "
    angka(9) = "Sembilan "
    
    Dim posisi(2)
        posisi(1) = "Puluh "
        posisi(2) = "Ratus "
    
    bilang = ""
    
    For j = 2 To 1 Step -1
        tmp = Int(y / (10 ^ j))
        If (tmp > 0) Then
            bag = angka(tmp)
            If (j = 1 And tmp = 1) Then
                y = y - tmp * 10 ^ j
                If (y >= 1) Then
                    posisi(j) = "Belas "
                Else
                    angka(y) = "Se"
                End If
                
                If bag = "Se" Then
                    bilang = bilang & angka(y) & LCase(posisi(j))
                Else
                    bilang = bilang & angka(y) & posisi(j)
                End If
                
                ratusan = bilang
                Exit Function
            Else
                If bag = "Se" Then
                    bilang = bilang & bag & LCase(posisi(j))
                Else
                    bilang = bilang & bag & posisi(j)
                End If
            End If
        End If
        y = y - tmp * 10 ^ j
    Next
    
    If (flag = False) Then
        angka(1) = "Satu "
    End If
    
    bilang = bilang & angka(y)
    ratusan = bilang
End Function
