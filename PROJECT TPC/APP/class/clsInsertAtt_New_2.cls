VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsertAtt_New_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Insert_h_attendance(p_Att_Date As String, p_EmployeeCode As String, p_shift_code As String, p_status As String, _
    p_tot_hour As Double, p_15X As Double, p_2X As Double, p_3X As Double, p_4X As Double, p_tot_overtime As Double, _
    p_ipAddress As String, p_EnrollNumber As String, p_ket As String)
    
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim StartTime As Date, EndTime As Date
Dim shiftNumber As String, shiftCode As String
Dim checkin_DIFF As Double, checkOUT_DIFF As Double
Dim flagCheckInEarly As Integer, flagCheckOutEarly As Integer
Dim flagCheckInLate As Integer, flagCheckOutLate As Integer
Dim abs_number As Integer
Dim rsnumber As New ADODB.Recordset


    flagCheckInEarly = 0
    flagCheckOutEarly = 0
    flagCheckInLate = 0
    flagCheckOutLate = 0

    '#################################
    strSQL = "SELECT c.shift_number,d.shift_code,e.start_time,e.end_time,c.employee_code " & _
            "FROM  td_shift c JOIN tm_shift d ON c.company_code = d.company_code AND c.shift_number = d.shift_number " & _
            "JOIN m_shift e ON d.shift_code = e.shift_code " & _
            "WHERE c.employee_code = '" & p_EmployeeCode & "'"
    
    rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        shiftNumber = rs!shift_number
        shiftCode = rs!shift_code
        
        StartTime = DateValue(p_Att_Date) & " " & TimeValue(Format(rs!start_time, "yyyy-MM-dd hh:mm:ss"))
        EndTime = DateValue(p_Att_Date) & " " & TimeValue(Format(rs!end_time, "yyyy-MM-dd hh:mm:ss"))
        checkin_DIFF = DateDiff("h", StartTime, StartTime)
        checkOUT_DIFF = DateDiff("h", EndTime, EndTime)
        
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
    
'    strSQL = "SElECT att_date FROM h_attendance WHERE employee_code = '" & p_EmployeeCode & "' " & _
'         "AND DATE(att_date) = DATE('" & p_Att_Date & "')"
'
'     rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
'     If rs.RecordCount > 0 Then
'        strSQL = "UPDATE h_overtime SET total_hour = '" & p_tot_hour & "',15jam = '" & p_15X & "',holiday = '" & IIf(p_tot_hour = 0, 1, 0) & "'," & _
'                "2jam = '" & p_2X & "',3jam = '" & p_3X & "',4jam = '" & p_4X & "',tot_overtime = '" & p_tot_overtime & "' " & _
'             "WHERE employee_code = '" & p_EmployeeCode & "' " & _
'             "AND DATE(att_date) = DATE('" & p_Att_Date & "')"
'
'         strSQL = "UPDATE h_attendance SET time_in = '" & p_Att_Date & "'," & _
'             "time_in_diff = CASE WHEN NOT '" & checkin_DIFF & "' < 0 " & _
'                 "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & StartTime & "')) ) " & _
'                 "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & StartTime & "'),TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
'             "time_out_diff = CASE WHEN NOT '" & checkOUT_DIFF & "' < 0 " & _
'                 "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & EndTime & "')) ) " & _
'                 "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & EndTime & "'),TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
'             "flag_early = '" & flagCheckInEarly & "',flag_late = '" & flagCheckInLate & "' " & _
'             "WHERE employee_code = '" & p_EmployeeCode & "' " & _
'             "AND DATE(att_date) = DATE('" & p_Att_Date & "')"
'         CnG.Execute strSQL
'
'    Else

    strSQL = "DELETE FROM h_attendance WHERE employee_code = '" & p_EmployeeCode & "' " & _
         "AND DATE(att_date) = DATE('" & p_Att_Date & "')"
    CnG.Execute strSQL
    
    strSQL = "DELETE FROM h_overtime WHERE employee_code = '" & p_EmployeeCode & "' " & _
         "AND DATE(att_date) = DATE('" & p_Att_Date & "')"
    CnG.Execute strSQL
    
    
    If p_tot_overtime > 0 Then
        strSQL = "INSERT INTO h_overtime (employee_code,att_date,shift_code,total_hour,15jam," & _
                "2jam,3jam,4jam,tot_overtime,holiday,entry_date,entry_user) " & _
            "VALUES " & _
            "('" & p_EmployeeCode & "','" & p_Att_Date & "','" & p_shift_code & "','" & p_tot_hour & "','" & p_15X & "'," & _
            "'" & p_2X & "','" & p_3X & "','" & p_4X & "','" & p_tot_overtime & "','" & IIf(p_tot_hour = 0, 1, 0) & "',now(),'" & LOGIN_CODE & "')"
        CnG.Execute strSQL
    End If
    
    If p_status <> "" Then
        Select Case UCase(p_status)
            Case "P"
                strSQL = "INSERT INTO h_attendance (employee_code, att_date, ip_address, enrollnumber," & _
                    "shift_number, shift_code, start_time," & _
                    "end_time, time_in, time_out," & _
                    "time_in_diff, time_out_diff," & _
                    "flag_io, flag_present, flag_late, flag_early," & _
                    "absent_status, description, entry_date) VALUES " & _
                    "('" & p_EmployeeCode & "', '" & p_Att_Date & "', '" & p_ipAddress & "', '" & p_EnrollNumber & "'," & _
                        "'" & shiftNumber & "', '" & shiftCode & "','" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "'," & _
                        "'" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "', '" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "', '" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "'," & _
                    "CASE WHEN NOT '" & checkin_DIFF & "' < 0 " & _
                        "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "')) ) " & _
                        "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & Format(StartTime, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
                    "CASE WHEN NOT '" & checkOUT_DIFF & "' < 0 " & _
                        "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "')) ) " & _
                        "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & Format(EndTime, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
                    "0, 1, '" & flagCheckInLate & "','" & flagCheckInEarly & "', 0, '" & p_ket & "',NOW())"
            Case "A"
                rsnumber.Open "select ifnull(max(absent_number),0)+1 as absent_number from t_absent", CnG, adOpenForwardOnly, adLockReadOnly
                abs_number = rsnumber!absent_number
                
                strSQL = "INSERT INTO t_absent (absent_number,employee_code,absent_date_from,flag_date_to,absent_status,description,entry_date) " & _
                    "VALUES " & _
                    "('" & abs_number & "','" & p_EmployeeCode & "','" & p_Att_Date & "',0,2,'ALPHA',NOW())"
                CnG.Execute strSQL
                rsnumber.Close
                
            Case "S"
                rsnumber.Open "select ifnull(max(absent_number),0)+1 as absent_number from t_absent", CnG, adOpenForwardOnly, adLockReadOnly
                abs_number = rsnumber!absent_number
                
                strSQL = "INSERT INTO t_absent (absent_number,employee_code,absent_date_from,flag_date_to,absent_status,description,entry_date) " & _
                    "VALUES " & _
                    "('" & abs_number & "','" & p_EmployeeCode & "','" & p_Att_Date & "',0,1,'SICK',NOW())"
                CnG.Execute strSQL
                rsnumber.Close
                
            Case "S'"
                rsnumber.Open "select ifnull(max(absent_number),0)+1 as absent_number from t_absent", CnG, adOpenForwardOnly, adLockReadOnly
                abs_number = rsnumber!absent_number
                
                strSQL = "INSERT INTO t_absent (absent_number,employee_code,absent_date_from,flag_date_to,absent_status,description,entry_date) " & _
                    "VALUES " & _
                    "('" & abs_number & "','" & p_EmployeeCode & "','" & p_Att_Date & "',0,0,'Permision',NOW())"
                CnG.Execute strSQL
                rsnumber.Close
                
            Case "C"
                rsnumber.Open "select ifnull(max(leave_number),0)+1 as leave_number from t_leave", CnG, adOpenForwardOnly, adLockReadOnly
                abs_number = rsnumber!leave_number
                
                strSQL = "INSERT INTO t_leave (leave_number,employee_code,leave_type,flag_doctor_order,leave_date_from,flag_date_to,description,entry_date) " & _
                    "VALUES " & _
                    "('" & abs_number & "','" & p_EmployeeCode & "',0,0,'" & p_Att_Date & "',0,'Cuti',NOW())"
                CnG.Execute strSQL
                rsnumber.Close
            End Select

    End If
'    End If
          
End Sub


