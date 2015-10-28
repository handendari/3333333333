Attribute VB_Name = "mod_function"
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
 
Private Const MF_BYPOSITION = &H400&

Private sDefInitFileName As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
 
Public Sub RemoveButtonX(frm As Form)
    Dim hSysMenu As Long
    hSysMenu = GetSystemMenu(frm.hwnd, 0)
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

Private Function SpellDigit(strNumeric As Integer)
Dim cRet As String
On Error GoTo Err

cRet = ""
    
    Select Case strNumeric
        Case 0:     cRet = " zero"
        Case 1:     cRet = " one"
        Case 2:     cRet = " two"
        Case 3:     cRet = " three"
    
        Case 4:     cRet = " four"
        Case 5:     cRet = " five"
        Case 6:     cRet = " six"
        Case 7:     cRet = " seven"
    
        Case 8:     cRet = " eight"
        Case 9:     cRet = " nine"
        Case 10:    cRet = " ten"
        Case 11:    cRet = " eleven"
    
        Case 12:    cRet = " twelve"
        Case 13:    cRet = " thirteen"
        Case 14:    cRet = " fourteen"
        Case 15:    cRet = " fifteen"
        Case 16:    cRet = " sixteen"
    
        Case 17:    cRet = " seventeen"
        Case 18:    cRet = " eighteen"
        Case 19:    cRet = " ninetieen"
        Case 20:    cRet = " twenty"
    
        Case 30:    cRet = " thirty"
        Case 40:    cRet = " fourthy"
        Case 50:    cRet = " fifty"
        Case 60:    cRet = " sixty"
    
        Case 70:    cRet = " seventy"
        Case 80:    cRet = " eighty"
        Case 90:    cRet = " ninety"
        Case 100:   cRet = " one hundred"
        Case 200:   cRet = " two hundred"
    
        Case 300:   cRet = " three hundred"
        Case 400:   cRet = " four hundred"
        Case 500:   cRet = " five hundred"
        Case 600:   cRet = " six hundred"
    
        Case 700:   cRet = " seven hundred"
        Case 800:   cRet = " eight hundred"
        Case 900:   cRet = " nine hundred"
    End Select
    
    SpellDigit = cRet
    Exit Function
    
Err:
SpellDigit = "Max 9 Digit"
End Function

Private Function SpellUnit(strNumeric As Integer)
Dim cRet As String
Dim n100 As Integer

Dim n10 As Integer
Dim n1 As Integer

On Error GoTo Err
cRet = ""

    n100 = Int(strNumeric / 100) * 100
    n10 = Int((strNumeric - n100) / 10) * 10
    n1 = (strNumeric - n100 - n10)
    If n100 > 0 Then
       cRet = SpellDigit(n100)
    End If
    
    If n10 > 0 Then
       If n10 = 10 Then
    
          cRet = cRet & SpellDigit(n10 + n1)
       Else
    
          cRet = cRet & SpellDigit(n10)
       End If
    End If
    
    If n1 > 0 And n10 <> 10 Then
       cRet = cRet & SpellDigit(n1)
    
    End If
    SpellUnit = cRet
    Exit Function

Err:
SpellUnit = "Max 9 Digit"
End Function

Public Function TerbilangInggris(strNumeric As String) As String

Dim cRet As String
Dim n1000000 As Long
Dim n1000 As Long
Dim n1 As Integer
Dim n0 As Integer

On Error GoTo Err
    Dim strValid As String, huruf As String * 1
    Dim i As Integer
    'Periksa setiap karakter masukan

    strValid = "1234567890.,"
    For i% = 1 To Len(strNumeric)
      huruf = Chr(Asc(Mid(strNumeric, i%, 1)))
      If InStr(strValid, huruf) = 0 Then

        MsgBox "Must Be Number Format", _
               vbCritical, "Character Not Valid"
        Exit Function
      End If
    Next i%
   
    If strNumeric = "" Then Exit Function
    If Len(Trim(strNumeric)) > 12 Then GoTo Err

    cRet = ""
    n1000000 = Int(strNumeric / 1000000) * 1000000
    n1000 = Int((strNumeric - n1000000) / 1000) * 1000
    n1 = Int(strNumeric - n1000000 - n1000)

    n0 = (strNumeric - n1000000 - n1000 - n1) * 100
    If n1000000 > 0 Then
       cRet = SpellUnit(n1000000 / 1000000) & " million"
    End If
    If n1000 > 0 Then

       cRet = cRet & SpellUnit(n1000 / 1000) & " thousand"
    End If
    If n1 > 0 Then
       cRet = cRet & SpellUnit(n1)

    End If
    If n0 > 0 Then
       cRet = cRet & " and cents" & SpellUnit(n0)
    End If

    TerbilangInggris = cRet & " "
    Exit Function

Err:
TerbilangInggris = "Max 9 Digit"
End Function

Function ratusan(ByVal Y As Double, ByVal Flag As Boolean) As String
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
        tmp = Int(Y / (10 ^ j))
        If (tmp > 0) Then
            bag = angka(tmp)
            If (j = 1 And tmp = 1) Then
                Y = Y - tmp * 10 ^ j
                If (Y >= 1) Then
                    posisi(j) = "Belas "
                Else
                    angka(Y) = "Se"
                End If
                
                If bag = "Se" Then
                    bilang = bilang & angka(Y) & LCase(posisi(j))
                Else
                    bilang = bilang & angka(Y) & posisi(j)
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
        Y = Y - tmp * 10 ^ j
    Next
    
    If (Flag = False) Then
        angka(1) = "Satu "
    End If
    
    bilang = bilang & angka(Y)
    ratusan = bilang
End Function

Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
    sFilePath As String, bSmtpSSL As Boolean) As String
      
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg      As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.from = sFrom
    lobj_cdomsg.Subject = sSubject
    lobj_cdomsg.TextBody = sBody
    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)
    End If
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = Err.Description
End Function

Function EncryptINI(Strg$, passWord$)
   Dim b$, s$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, P$
   j = 1
   For i = 1 To Len(passWord$)
     P$ = P$ & Asc(Mid$(passWord$, i, 1))
   Next
    
   For i = 1 To Len(Strg$)
     A1 = Asc(Mid$(P$, j, 1))
     j = j + 1: If j > Len(P$) Then j = 1
     A2 = Asc(Mid$(Strg$, i, 1))
     A3 = A1 Xor A2
     b$ = Hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b$
     s$ = s$ + b$
   Next
   EncryptINI = s$
End Function

Function DecryptINI(Strg$, passWord$)
   Dim b$, s$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, P$
   j = 1
   For i = 1 To Len(passWord$)
     P$ = P$ & Asc(Mid$(passWord$, i, 1))
   Next
   
   For i = 1 To Len(Strg$) Step 2
     A1 = Asc(Mid$(P$, j, 1))
     j = j + 1: If j > Len(P$) Then j = 1
     b$ = Mid$(Strg$, i, 2)
     A3 = Val("&H" + b$)
     A2 = A1 Xor A3
     s$ = s$ + Chr$(A2)
   Next
   DecryptINI = s$
End Function

Public Function getPeriode(bulan As String, dtp1 As DTPicker, dtp2 As DTPicker)
    If rsPeriode.State Then rsPeriode.Open
    SQL = "SELECT * FROM m_pref_periode"
    rsPeriode.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsPeriode.RecordCount > 0 Then
        vFlagPeriode = rsPeriode!flag_periode
        vDayStart = rsPeriode!day_start
        vDayEnd = rsPeriode!day_end
    Else
        vFlagPeriode = 0
        vDayStart = "01"
        vDayEnd = getEndDay(month(Now), year(Now))
    End If
    rsPeriode.Close
    
    If vFlagPeriode = 0 Then
        dtp1.Value = "01" & Format(bulan, "-MM-yyyy")
        dtp2.Value = getEndDay(month(bulan), year(bulan)) & Format(bulan, "-MM-yyyy")
    ElseIf vFlagPeriode = 1 Then
        If vDayStart < 10 Then
            dtp1.Value = "0" & vDayStart & Format(bulan, "-MM-yyyy")
        Else
            dtp1.Value = vDayStart & Format(DateAdd("m", -1, bulan), "-MM-yyyy")
        End If
        
        If vDayEnd < 10 Then
            dtp2.Value = "0" & vDayEnd & Format(bulan, "-MM-yyyy")
        Else
            dtp2.Value = vDayEnd & Format(bulan, "-MM-yyyy")
        End If
    End If
End Function

Public Function getPeriodeVar(bulan As String, tgl1 As String, tgl2 As String)
    If rsPeriode.State Then rsPeriode.Open
    SQL = "SELECT * FROM m_pref_periode"
    rsPeriode.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsPeriode.RecordCount > 0 Then
        vFlagPeriode = rsPeriode!flag_periode
        vDayStart = rsPeriode!day_start
        vDayEnd = rsPeriode!day_end
    Else
        vFlagPeriode = 0
        vDayStart = "01"
        vDayEnd = getEndDay(month(Now), year(Now))
    End If
    rsPeriode.Close
    
    If vFlagPeriode = 0 Then
        tgl1 = "01" & Format(bulan, "-MM-yyyy")
        tgl2 = getEndDay(month(bulan), year(bulan)) & Format(bulan, "-MM-yyyy")
    ElseIf vFlagPeriode = 1 Then
        If vDayStart < 10 Then
            tgl1 = "0" & vDayStart & Format(bulan, "-MM-yyyy")
        Else
            tgl1 = vDayStart & Format(DateAdd("m", -1, bulan), "-MM-yyyy")
        End If
        
        If vDayEnd < 10 Then
            tgl2 = "0" & vDayEnd & Format(bulan, "-MM-yyyy")
        Else
            tgl2 = vDayEnd & Format(bulan, "-MM-yyyy")
        End If
    End If
End Function

Public Sub check_late(str_employee_code As String, start_time As String, end_time As String)
Dim v_tgl_awal, v_tgl_akhir As Date
Dim v_start_time, v_time_in As String
Dim vLate As Double
Dim vAttDate As String
Dim vFlagLate As Integer
Dim vLateConvert As Double
Dim rsLate As New ADODB.Recordset
Dim vLateRound As Double
Dim vFlagTransport As Integer
Dim vFlagTraining As Integer
    
    vLateConvert = 0
    v_tgl_awal = Format(start_time, "yyyy-MM-dd")
    v_tgl_akhir = Format(end_time, "yyyy-MM-dd")
    
    v_tgl_awal = DateValue(v_tgl_awal)
    v_tgl_akhir = DateValue(v_tgl_akhir)
    
    While v_tgl_awal < v_tgl_akhir
        If rscari.State Then rscari.Close
        SQL = "SELECT att_date,start_time, end_time, time_in, time_out, b.flag_inc_late, a.flag_training " & _
              "FROM h_attendance a JOIN m_employee b ON a.employee_code = b.employee_code " & _
              "WHERE a.employee_code = '" & str_employee_code & "' " & _
                "AND DATE(att_date) = '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' " & _
                "AND status = 'P' " & _
                "AND DATE(att_date) NOT IN (SELECT DATE(pl_date) FROM t_private_leave WHERE employee_code = a.employee_code) " & _
              "ORDER BY IFNULL(time_in,Now()) ASC LIMIT 1"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            vFlagLate = IIf(IsNull(rscari!flag_inc_late), 0, rscari!flag_inc_late)
            vFlagTraining = IIf(IsNull(rscari!flag_training), 0, rscari!flag_training)
            v_start_time = Format(v_tgl_awal, "yyyy-MM-dd") & " " & Format(rscari!start_time, "hh:mm:ss")
            v_time_in = Format(rscari!time_in, "yyyy-MM-dd hh:mm:ss")
            
            If vFlagTraining = 0 Then
                vAttDate = Format(rscari!att_date, "yyyy-MM-dd hh:mm:ss")
                vLate = DateDiff("n", v_start_time, IIf(v_time_in = "", v_start_time, v_time_in))
                vLateRound = Int(vLate / 60)
                If vLateRound < 0 Then vLateRound = 0
                
                vLate = ((vLate / 60) - vLateRound) * 60
                If vLate < 0 Then vLate = 0
                            
                SQL = "SELECT convert_value FROM m_pref_lateconvert " & _
                      "WHERE from_value <= '" & vLate & "' " & _
                        "AND to_value >= '" & vLate & "'"
                rsLate.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rsLate.RecordCount > 0 Then
                    vLateConvert = rsLate!convert_value
                    vLateConvert = vLateConvert / 60
                    
                    If vFlagLate = 1 Then
                        vLateConvert = 0
                    Else
                        vLateConvert = vLateRound + vLateConvert
                    End If
                    
                    SQL = "UPDATE h_attendance SET late_value = " & IIf(vLateConvert = 0, "NULL", vLateConvert) & " " & _
                          "WHERE employee_code = '" & str_employee_code & "' " & _
                            "AND att_date = '" & vAttDate & "'"
                    CnG.Execute SQL
                End If
                rsLate.Close
            Else
                SQL = "UPDATE h_attendance SET late_value = NULL " & _
                      "WHERE employee_code = '" & str_employee_code & "' " & _
                        "AND DATE(att_date) = '" & Format(v_tgl_awal, "yyyy-MM-dd") & "'"
                CnG.Execute SQL
            End If
        Else
            SQL = "UPDATE h_attendance SET late_value = NULL " & _
                  "WHERE employee_code = '" & str_employee_code & "' " & _
                    "AND DATE(att_date) = '" & Format(v_tgl_awal, "yyyy-MM-dd") & "'"
            CnG.Execute SQL
        End If
        rscari.Close
        v_tgl_awal = v_tgl_awal + 1
    Wend
End Sub

Public Sub check_pl(str_employee_code As String, start_time As String, end_time As String)
Dim rsPL As New ADODB.Recordset
Dim vEmployeeCode As String
Dim vTimeIn As Double, vTimeOut As Double
Dim vStatus As String
Dim v_tgl_awal, v_tgl_akhir As Date
Dim v_start_time, v_end_time As String
Dim vFlagShift As Double
Dim vSeq As Integer
Dim vPLDate As Date
        
    v_tgl_awal = Format(start_time, "yyyy-MM-dd")
    v_tgl_akhir = Format(end_time, "yyyy-MM-dd")
    
    v_tgl_awal = DateValue(v_tgl_awal)
    v_tgl_akhir = DateValue(v_tgl_akhir)
    
    If rscari.State Then rscari.Close
    SQL = "SELECT IFNULL(flag_shiftable,0) flag_shiftable FROM m_employee WHERE employee_code = '" & str_employee_code & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vFlagShift = rscari.Fields(0).Value
    End If
    rscari.Close
    
    If vFlagShift = 0 Then
        SQL = "SELECT  b.dt,IFNULL(c.employee_code,0) employee_code,i.group_code,i.shift_code,h.start_time,h.end_time," & _
                    "a.company_code, TIME_TO_SEC(IFNULL(c.time_in,0)) time_in," & _
                    "TIME_TO_SEC(IFNULL(c.time_out,0)) time_out, IFNULL(c.status,'') status " & _
              "FROM m_employee a JOIN m_days b ON 1 = 1 " & _
                    "LEFT JOIN h_attendance c ON DATE(b.dt) = DATE(c.att_date) AND a.employee_code = c.employee_code " & _
                    "LEFT JOIN m_title d ON a.title_code = d.title_code " & _
                    "LEFT JOIN m_division e ON a.division_code = e.division_code AND a.department_code = e.department_code AND a.company_code = e.company_code " & _
                    "LEFT JOIN m_absent_status f ON c.status = f.absent_code " & _
                    "LEFT JOIN m_department g ON a.department_code = g.department_code AND a.company_code = g.company_code " & _
                    "LEFT JOIN m_shift h ON c.shift_code = h.shift_code " & _
                    "LEFT JOIN td_shift2 i ON c.employee_code = i.employee_code " & _
              "WHERE (DATE(b.dt) BETWEEN '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' AND '" & Format(v_tgl_akhir, "yyyy-MM-dd") & "') " & _
                "AND a.employee_code = '" & str_employee_code & "' " & _
                "AND DAYNAME(b.dt) <> 'Saturday' AND DAYNAME(b.dt) <> 'Sunday' " & _
                "AND DATE(b.dt) NOT IN (SELECT DATE(holiday_date) FROM t_holiday) " & _
              "ORDER BY b.dt ASC"
        
'        SQL = "SELECT CASE WHEN IFNULL(c.att_date,0) = 0 THEN b.dt ELSE c.att_date END dt, IFNULL(c.employee_code,0) employee_code, d.group_code, d.shift_code, e.start_time, e.end_time, a.company_code, TIME_TO_SEC(IFNULL(c.time_in,0)) time_in, TIME_TO_SEC(IFNULL(c.time_out,0)) time_out, IFNULL(c.status,'') status " & _
'              "FROM m_employee a JOIN m_days b ON 1 = 1 " & _
'                "LEFT JOIN h_attendance c ON DATE(b.dt) = DATE(c.att_date) AND a.employee_code = c.employee_code " & _
'                "JOIN td_shift2 d ON a.employee_code = d.employee_code " & _
'                "JOIN m_shift e ON d.group_code = e.group_code AND d.shift_code = e.shift_code " & _
'              "WHERE (DATE(b.dt) BETWEEN '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' AND '" & Format(v_tgl_akhir, "yyyy-MM-dd") & "') " & _
'                "AND a.employee_code = '" & str_employee_code & "' " & _
'                "AND DAYNAME(b.dt) <> 'Saturday' AND DAYNAME(b.dt) <> 'Sunday' " & _
'                "AND DATE(b.dt) NOT IN (SELECT DATE(holiday_date) FROM t_holiday) " & _
'              "ORDER BY b.dt ASC"
    Else
        SQL = "SELECT  b.dt,IFNULL(c.employee_code,0) employee_code," & _
                    "(SELECT x.group_code FROM td_emp_group x JOIN m_shift_new z ON x.emp_group_code = z.group_code WHERE employee_code = a.employee_code AND DATE(z.shift_date) = DATE(b.dt) AND DATE(x.start_date) <= DATE(b.dt) ORDER BY x.start_date DESC LIMIT 1) group_code," & _
                    "(SELECT z.shift_code FROM td_emp_group x JOIN m_shift_new z ON x.emp_group_code = z.group_code WHERE employee_code = a.employee_code AND DATE(z.shift_date) = DATE(b.dt) AND DATE(x.start_date) <= DATE(b.dt) ORDER BY x.start_date DESC LIMIT 1) shift_code," & _
                    "h.start_time,h.end_time," & _
                    "a.company_code, TIME_TO_SEC(IFNULL(c.time_in,0)) time_in," & _
                    "TIME_TO_SEC(IFNULL(c.time_out,0)) time_out, IFNULL(c.status,'') status " & _
              "FROM m_employee a JOIN m_days b ON 1 = 1 " & _
                    "LEFT JOIN h_attendance c ON DATE(b.dt) = DATE(c.att_date) AND a.employee_code = c.employee_code " & _
                    "LEFT JOIN m_title d ON a.title_code = d.title_code " & _
                    "LEFT JOIN m_division e ON a.division_code = e.division_code AND a.department_code = e.department_code AND a.company_code = e.company_code " & _
                    "LEFT JOIN m_absent_status f ON c.status = f.absent_code " & _
                    "LEFT JOIN m_department g ON a.department_code = g.department_code AND a.company_code = g.company_code " & _
                    "LEFT JOIN m_shift h ON c.shift_code = h.shift_code " & _
              "WHERE (DATE(b.dt) BETWEEN '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' AND '" & Format(v_tgl_akhir, "yyyy-MM-dd") & "') " & _
                "AND a.employee_code = '" & str_employee_code & "' " & _
                "AND (SELECT z.shift_code FROM td_emp_group x JOIN m_shift_new z ON x.emp_group_code = z.group_code WHERE employee_code = a.employee_code AND DATE(z.shift_date) = DATE(b.dt) AND DATE(x.start_date) <= DATE(b.dt) ORDER BY x.start_date DESC LIMIT 1) <> 'OFF' " & _
              "ORDER BY b.dt ASC"
        
'        SQL = "SELECT CASE WHEN IFNULL(c.att_date,0) = 0 THEN b.dt ELSE c.att_date END dt, IFNULL(c.employee_code,0) employee_code, d.group_code, e.shift_code, f.start_time, f.end_time, a.company_code, TIME_TO_SEC(IFNULL(c.time_in,0)) time_in, TIME_TO_SEC(IFNULL(c.time_out,0)) time_out, IFNULL(c.status,'') status " & _
'              "FROM m_employee a JOIN m_days b ON 1 = 1 " & _
'                "LEFT JOIN h_attendance c ON DATE(b.dt) = DATE(c.att_date) AND a.employee_code = c.employee_code " & _
'                "JOIN td_emp_group d ON a.employee_code = d.employee_code " & _
'                "JOIN m_shift_new e ON d.emp_group_code = e.group_code AND DATE(b.dt) = DATE(e.shift_date) " & _
'                "JOIN m_shift f ON d.group_code = f.group_code AND e.shift_code = f.shift_code " & _
'              "WHERE (DATE(b.dt) BETWEEN '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' AND '" & Format(v_tgl_akhir, "yyyy-MM-dd") & "') " & _
'                "AND a.employee_code = '" & str_employee_code & "' " & _
'                "AND e.shift_code <> 'OFF' " & _
'              "ORDER BY b.dt ASC"
    End If
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
    If rscari.RecordCount > 0 Then
        rscari.MoveFirst
        While Not rscari.EOF
            vEmployeeCode = rscari!employee_code
            vTimeIn = rscari!time_in
            vTimeOut = rscari!time_out
            vStatus = rscari!Status
            
            vPLDate = Format(rscari!Dt, "yyyy-MM-dd")
            vPLDate = DateValue(vPLDate)
            
            If Format(vPLDate, "yyyy-MM-dd") > "2014-01-20" Then
                v_start_time = Format(rscari!Dt, "yyyy-MM-dd") & " " & Format(rscari!start_time, "HH:mm")
                If Format(rscari!start_time, "HH:mm") > Format(rscari!end_time, "HH:mm") Then
                    v_end_time = Format(DateAdd("d", 1, rscari!Dt), "yyyy-MM-dd") & " " & Format(rscari!end_time, "HH:mm")
                Else
                    v_end_time = Format(rscari!Dt, "yyyy-MM-dd") & " " & Format(rscari!end_time, "HH:mm")
                End If
                
                If vEmployeeCode = "0" Then
                    SQL = "INSERT INTO h_attendance (att_date,employee_code,group_code,shift_code,status,flag_present,absent_status," & _
                            "description,entry_date,userinput,shift_number,flag_manual) " & _
                          "VALUES (" & _
                            "'" & Format(rscari!Dt, "yyyy-MM-dd") & "','" & str_employee_code & "','" & rscari!group_code & "'," & _
                            "'" & rscari!shift_code & "','PL',0,9," & _
                            "'PRIVATE LEAVE',now(),'" & LOGIN_NAME & "',1,1)"
                    CnG.Execute SQL
                    
                    If rsPL.State Then rsPL.Close
                    SQL = "SELECT MAX(seq) jmlSeq FROM t_private_leave WHERE employee_code = '" & str_employee_code & "' " & _
                            "AND DATE(pl_date) = '" & Format(rscari!Dt, "yyyy-MM-dd") & "' " & _
                            "AND flag_type = 0"
                    rsPL.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rsPL.RecordCount > 0 Then
                        vSeq = IIf(IsNull(rsPL!jmlSeq), 0, rsPL!jmlSeq) + 1
                    Else
                        vSeq = 1
                    End If
                    rsPL.Close
                    
                    SQL = "INSERT INTO t_private_leave(employee_code,pl_date,time_in,time_out,company_code," & _
                            "description,flag_break,int_break,flag_type,entry_date,entry_user,seq,flag_approval,user_approval) " & _
                          "VALUES ( " & _
                            "'" & str_employee_code & "','" & Format(rscari!Dt, "yyyy-MM-dd HH:mm:ss") & "'," & _
                            "'" & v_start_time & "','" & v_end_time & "'," & _
                            "'" & rscari!COMPANY_CODE & "','AUTOMATICALLY PRIVATE LEAVE',0,0," & _
                            "0,now(),'" & LOGIN_NAME & "'," & vSeq & ",1,'" & LOGIN_NAME & "')"
                    CnG.Execute SQL
                Else
                    If vTimeIn <> 0 Or vTimeOut <> 0 Then
                        If vStatus = "PL" Then
                            SQL = "DELETE FROM t_private_leave WHERE employee_code = '" & str_employee_code & "' " & _
                                    "AND DATE(pl_date) = '" & Format(rscari!Dt, "yyyy-MM-dd") & "'"
                            CnG.Execute SQL
                            
                            SQL = "UPDATE h_attendance SET status = 'P' " & _
                                  "WHERE employee_code = '" & str_employee_code & "' " & _
                                    "AND DATE(att_date) = '" & Format(rscari!Dt, "yyyy-MM-dd HH:mm:ss") & "'"
                            CnG.Execute SQL
                        End If
                    Else
                        If vStatus = "PL" Then
                            If rsPL.State Then rsPL.Close
                            SQL = "SELECT employee_code FROM h_attendance " & _
                                  "WHERE employee_code = '" & str_employee_code & "' " & _
                                        "AND DATE(att_date) = '" & Format(rscari!Dt, "yyyy-MM-dd") & "'"
                            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                            
                            If rs.RecordCount > 1 Then
                                SQL = "DELETE FROM t_private_leave WHERE employee_code = '" & str_employee_code & "' " & _
                                        "AND DATE(pl_date) = '" & Format(rscari!Dt, "yyyy-MM-dd") & "'"
                                CnG.Execute SQL
                                
                                SQL = "DELETE FROM h_attendance " & _
                                      "WHERE employee_code = '" & str_employee_code & "' " & _
                                        "AND att_date = '" & Format(rscari!Dt, "yyyy-MM-dd HH:mm:ss") & "'"
                                CnG.Execute SQL
                            End If
                            rs.Close
                        End If
                    End If
                End If
            End If
            rscari.MoveNext
        Wend
        
        SQL = "DELETE a FROM t_private_leave a JOIN h_attendance b ON a.employee_code = b.employee_code AND DATE(a.pl_date) = DATE(b.att_date)" & _
                  "WHERE a.employee_code = '" & str_employee_code & "' " & _
                    "AND (DATE(a.pl_date) BETWEEN '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' AND '" & Format(v_tgl_akhir, "yyyy-MM-dd") & "') " & _
                    "AND a.description = 'AUTOMATICALLY PRIVATE LEAVE' AND b.status = 'P'"
            CnG.Execute SQL
    End If
    rscari.Close
End Sub

Public Function flagPLAuto() As Double
    str_sql = "SELECT ifnull(flag_pl,0) flag_pl FROM m_pref_gen"
    rs.Open str_sql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        flagPLAuto = rs!flag_pl
    End If
    rs.Close
End Function
