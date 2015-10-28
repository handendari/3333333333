VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsertAtt_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Insert_h_attendance(p_Att_Date As String, p_NIK As String, _
    p_checkIn As String, p_checkOut As String, _
    p_ipAddress As String, p_EnrollNumber As String, _
    p_ket As String, p_shift_code As String, p_status As String, _
    p_jmllembur As Double, p_15 As Double, p_2 As Double, p_3 As Double, _
    p_4 As Double, p_tot_overtime As Double, p_abs_status As Integer, p_EmployeeCode As String, _
    p_meal As Double, p_transport As Double, p_hari_libur As Integer, p_company As String)
    
Dim rs As New ADODB.Recordset
Dim rsnumber As New ADODB.Recordset
Dim strsql As String
Dim StartTime As Date, EndTime As Date
Dim shiftNumber As String, shiftCode As String
Dim checkin_DIFF As Double, checkOUT_DIFF As Double
Dim flagCheckInEarly As Integer, flagCheckOutEarly As Integer
Dim flagCheckInLate As Integer, flagCheckOutLate As Integer

    flagCheckInEarly = 0
    flagCheckOutEarly = 0
    flagCheckInLate = 0
    flagCheckOutLate = 0

    '#################################
    strsql = "SELECT c.shift_number,d.shift_code,e.start_time,e.end_time,c.employee_code " & _
            "FROM  td_shift c JOIN tm_shift d ON c.company_code = d.company_code AND c.shift_number = d.shift_number " & _
            "JOIN m_shift e ON d.shift_code = e.shift_code " & _
            "WHERE c.employee_code = '" & p_EmployeeCode & "'"
    
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        shiftNumber = rs!shift_number
        shiftCode = rs!shift_code
        
        StartTime = DateValue(p_Att_Date) & " " & TimeValue(Format(rs!start_time, "yyyy-MM-dd hh:mm:ss"))
        EndTime = DateValue(p_Att_Date) & " " & TimeValue(Format(rs!end_time, "yyyy-MM-dd hh:mm:ss"))
        checkin_DIFF = DateDiff("h", StartTime, p_checkIn)
        checkOUT_DIFF = DateDiff("h", EndTime, p_checkOut)
        
        If checkin_DIFF > 0 Then
            flagCheckInLate = 1
        ElseIf checkin_DIFF < 0 Then
            flagCheckInEarly = 1
        End If
        If checkOUT_DIFF > 0 Then
            flagCheckOutLate = 1
        ElseIf checkOUT_DIFF < 0 Then
            flagCheckOutEarly = 1
        End If
    Else
        StartTime = p_checkIn
        EndTime = p_checkOut
        shiftNumber = "0"
        shiftCode = "0"
        flagCheckInLate = 0
        flagCheckInEarly = 0
        flagCheckOutLate = 0
        flagCheckOutEarly = 0
    End If
    rs.Close
    '#################################
    
'    strsql = "SElECT att_date FROM h_attendance WHERE employee_code = '" & p_EmployeeCode & "' " & _
'         "AND DATE(att_date) = DATE('" & p_Att_Date & "')"
'
'     rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'     If rs.RecordCount > 0 Then
'         strsql = "UPDATE h_attendance SET time_in = '" & p_Att_Date & "'," & _
'             "time_in_diff = CASE WHEN NOT '" & checkin_DIFF & "' < 0 " & _
'                 "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & p_checkIn & "')) ) " & _
'                 "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & p_checkIn & "'),TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
'             "time_out_diff = CASE WHEN NOT '" & checkOUT_DIFF & "' < 0 " & _
'                 "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & p_checkOut & "')) ) " & _
'                 "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & p_checkOut & "'),TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
'             "flag_early = '" & flagCheckInEarly & "',flag_late = '" & flagCheckInLate & "' " & _
'             "WHERE employee_code = '" & p_EmployeeCode & "' " & _
'             "AND DATE(att_date) = DATE('" & p_Att_Date & "')"
'         CnG.Execute strsql
'
'    Else
'         strsql = "INSERT INTO h_attendance (employee_code, att_date, ip_address, enrollnumber," & _
'             "shift_number, shift_code, start_time," & _
'             "end_time, time_in, time_out," & _
'             "time_in_diff, time_out_diff," & _
'             "flag_io, flag_present, flag_late, flag_early," & _
'             "absent_status, description, entry_date) VALUES " & _
'             "('" & p_EmployeeCode & "', '" & p_Att_Date & "', '" & p_ipAddress & "', '" & p_EnrollNumber & "'," & _
'                 "'" & shiftNumber & "', '" & p_shift_code & "','" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "'," & _
'                 "'" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "', '" & p_checkIn & "', '" & p_checkOut & "'," & _
'             "CASE WHEN NOT '" & checkin_DIFF & "' < 0 " & _
'                 "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & p_checkIn & "')) ) " & _
'                 "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & p_checkIn & "'),TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
'             "CASE WHEN NOT '" & checkOUT_DIFF & "' < 0 " & _
'                 "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & p_checkOut & "')) ) " & _
'                 "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & p_checkOut & "'),TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
'             "0, 1, '" & flagCheckInLate & "','" & flagCheckInEarly & "', 0, '" & p_ket & "',NOW())"
'         CnG.Execute strsql
'    End If
    
    strsql = "DELETE FROM h_attendance WHERE employee_code = '" & p_EmployeeCode & "' " & _
         "AND DATE(att_date) = DATE('" & p_Att_Date & "')"
    CnG.Execute strsql
    
    If p_status <> "" Then
        Select Case UCase(p_status)
            Case "P"
                strsql = "Select start_time, end_time, min_break_in, max_break_out FROM m_shift WHERE shift_code = '" & p_shift_code & "' " & _
                            "AND company_code = '" & p_company & "'"
                rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                Dim v_break_in As String
                Dim v_break_out As String
                Dim v_start_time As String
                Dim v_end_time As String
                
                    v_break_in = Format(p_Att_Date, "yyyy-MM-dd") & " " & Format(rs!min_break_in, "hh:nn:ss")
                    v_break_out = Format(p_Att_Date, "yyyy-MM-dd") & " " & Format(rs!max_break_out, "hh:nn:ss")
                    v_start_time = Format(p_Att_Date, "yyyy-MM-dd") & " " & Format(rs!start_time, "hh:nn:ss")
                    v_end_time = Format(p_Att_Date, "yyyy-MM-dd") & " " & Format(rs!end_time, "hh:nn:ss")
                End If
                rs.Close
                
                strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & p_Att_Date & "' AND employee_code = '" & p_EmployeeCode & "'"
                CnG.Execute strsql
                
                strsql = "INSERT INTO h_attendance (employee_code, att_date, ip_address, enrollnumber," & _
                    "shift_number, shift_code, start_time," & _
                    "end_time, time_in, time_out," & _
                    "time_in_diff, time_out_diff," & _
                    "flag_io, flag_present, flag_late, flag_early," & _
                    "absent_status, description, entry_date, time_in_break,time_out_break," & _
                    "total_ot,15jam,2jam,3jam,4jam,holiday,tot_overtime,flag_meal,flag_transport) VALUES " & _
                    "('" & p_EmployeeCode & "', '" & p_Att_Date & "', '" & p_ipAddress & "', '" & p_EnrollNumber & "'," & _
                        "'" & p_shift_code & "', '" & p_status & "','" & v_start_time & "'," & _
                        "'" & v_end_time & "', '" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "', '" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "'," & _
                    "CASE WHEN NOT '" & checkin_DIFF & "' < 0 " & _
                        "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "')) ) " & _
                        "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
                    "CASE WHEN NOT '" & checkOUT_DIFF & "' < 0 " & _
                        "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "')) ) " & _
                        "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
                    "0, 1, '" & flagCheckInLate & "','" & flagCheckInEarly & "', 0, '" & p_ket & "',NOW(),'" & v_break_in & "','" & v_break_out & "'," & _
                    "" & p_jmllembur & "," & p_15 & "," & p_2 & "," & p_3 & "," & p_4 & "," & p_hari_libur & "," & p_tot_overtime & "," & p_meal & "," & p_transport & ")"
                CnG.Execute strsql
            Case "DT"
                
                strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & p_Att_Date & "' AND employee_code = '" & p_EmployeeCode & "'"
                CnG.Execute strsql
                
                strsql = "INSERT INTO h_attendance (employee_code, att_date, ip_address, enrollnumber," & _
                    "flag_present, flag_duty,description, entry_date, shift_code, shift_number) VALUES " & _
                    "('" & p_EmployeeCode & "', '" & p_Att_Date & "', '" & p_ipAddress & "', '" & p_EnrollNumber & "'," & _
                    "1, 1,'" & p_ket & "',NOW(),'" & p_status & "','" & p_shift_code & "')"
                CnG.Execute strsql
            Case "L"
                strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & p_Att_Date & "' AND employee_code = '" & p_EmployeeCode & "'"
                CnG.Execute strsql
                
                strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_number,shift_code,flag_present,absent_status,description,entry_date,userinput) " _
                    & "VALUES " _
                    & "('" & p_Att_Date & "','" & p_EmployeeCode & "','" & p_shift_code & "','" & p_status & "',0,3,'" & p_ket & "',now(),'" & LOGIN_CODE & "')"
                CnG.Execute strsql
            Case Else
                strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & p_Att_Date & "' AND employee_code = '" & p_EmployeeCode & "'"
                CnG.Execute strsql
                
                strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_number,shift_code,flag_present,absent_status,description,entry_date,userinput) " _
                    & "VALUES " _
                    & "('" & p_Att_Date & "','" & p_EmployeeCode & "','" & p_shift_code & "','" & p_status & "',0,'" & p_abs_status & "','" & p_ket & "',now(),'" & LOGIN_CODE & "')"
                CnG.Execute strsql
            End Select
    End If
          
End Sub


