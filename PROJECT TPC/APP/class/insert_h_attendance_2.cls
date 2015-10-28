VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsertAttendance_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Insert_h_attendance(p_Att_Date As String, p_ipAddress As String, p_EnrollNumber As String)
Dim rs As New ADODB.Recordset
Dim strSQL As String

Dim v_company_code As String, v_emp_code As String
Dim v_shift_number As Integer
Dim v_shift_code As String
Dim v_start_time As String, v_end_time As String, vt_start_time As String, vt_end_time As String
Dim v_flag_tolerance As Integer, v_data_created As Integer
Dim v_flag_late As Integer, v_flag_early As Integer
Dim v_str_buf As String
Dim v_int_flag_day_active As Integer
Dim v_flag_io, v_flag_io1 As Integer
Dim v_count_ As Integer
Dim v_device As Integer
Dim v_break_in As String, v_break_out As String, v_dt_buf1 As String, v_dt_buf2 As String
Dim v_break_interval As Integer
Dim v_mode As Integer, v_recs As Integer, v_recs1 As Integer
Dim v_action As Integer, v_action1 As Integer
Dim v_action_mode As Integer, v_action_mode1 As Integer, v_yesterday As Integer
Dim v_att_date As String
Dim v_mindiff As Long, v_mindiff1 As Long, v_mindiff_curr As Long, v_mindiff_curr1 As Long
Dim v_max_date As String, v_max_date1 As String, v_min_date As String, v_min_date1 As String, dt_att_date As String

    v_action = -1
    v_action_mode = -1
    v_action1 = -1
    v_action_mode1 = -1
    v_recs = -1
    v_recs1 = -1
    v_mode = -1
    v_yesterday = -1
    
    '++++++++++++
    strSQL = "SELECT s_device FROM s_device WHERE s_number = 1"
    rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    v_device = IIf(IsNull(rs!s_device), 1, rs!s_device)
    rs.Close
    '+++++++++++
    
    '++++++++++++++++++
    strSQL = "SELECT company_code, employee_code " & _
        "FROM m_enroll_link WHERE ip_address = '" & p_ipAddress & "' AND enrollnumber = '" & p_EnrollNumber & "'"
    rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    
     v_company_code = rs!COMPANY_CODE
     v_emp_code = rs!employee_code
     
     rs.Close
    '+++++++++++++++++++
    
'    CnG.Execute "INSERT INTO h_test (kode,nama) VALUES ('" & v_company_code & "', '" & v_emp_code & "')"
    
    '+++++++++++++++++++
    strSQL = "SELECT MAX(shift_number) shift_number FROM td_shift WHERE employee_code = '" & v_emp_code & "'"
    rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    v_shift_number = IIf(IsNull(rs!shift_number), 1, rs!shift_number)
    rs.Close
    '++++++++++++++++++++
    
    '++++++++++++++++++++
    strSQL = "SELECT shift_code FROM tm_shift WHERE shift_number = '" & v_shift_number & "' AND company_code = " & v_company_code & ""
    rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Dim abc As Integer
    abc = rs.RecordCount
    v_shift_code = rs!shift_code
    rs.Close
    '++++++++++++++++++++
    
    '++++++++++++++++++++
    strSQL = "SELECT MAX(IFNULL(flag_moving,0)) As v_mode FROM m_working_day WHERE shift_code = '" & v_shift_code & "'"
    rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    v_mode = rs!v_mode
    rs.Close
    '++++++++++++++++++++
    
    v_att_date = Format(p_Att_Date, "yyyy-MM-dd")
    dt_att_date = p_Att_Date
       
    '++++++++++++++++++++
    strSQL = "SELECT COUNT(*), MAX(att_date), MIN(att_date)," & _
                "((HOUR( TIMEDIFF(MAX(att_date),MIN(att_date)) )*60) + MINUTE(TIMEDIFF(MAX(att_date),MIN(att_date)) ))," & _
                "Max(flag_io) " & _
            "FROM h_log_attendance WHERE employee_code = '" & v_emp_code & "' AND DATE(att_date) = DATE('" & p_Att_Date & "')"
    rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    v_recs = IIf(IsNull(rs.Fields(0).Value), 0, rs.Fields(0).Value)
    v_max_date = IIf(IsNull(rs.Fields(1).Value), "", rs.Fields(1).Value)
    v_min_date = IIf(IsNull(rs.Fields(2).Value), "", rs.Fields(2).Value)
    v_mindiff = IIf(IsNull(rs.Fields(3).Value), 0, rs.Fields(3).Value)
    v_flag_io = IIf(IsNull(rs.Fields(4).Value), 0, rs.Fields(4).Value)
    rs.Close
    '++++++++++++++++++++
    
    If v_recs > 0 Then 'DATA HARI INI
        v_mindiff_curr = (Hour(DateDiff("h", dt_att_date, v_max_date)) * 60) + Minute(DateDiff("n", dt_att_date, v_max_date))

        If v_flag_io = 0 Then
            v_yesterday = 0
            If v_mindiff_curr > 30 Then
                v_action = 1
                v_action_mode = 2
                p_Flag_IO = 1
            ElseIf v_mindiff_curr <= 30 Then
                v_action = 0
                v_action_mode = 2
                p_Flag_IO = 0
            End If
        ElseIf v_flag_io = 1 Then
            v_yesterday = 1
            If v_mindiff_curr > 30 Then
                v_action = 1
                v_action_mode = 2
                p_Flag_IO = 1
            ElseIf v_mindiff_curr <= 30 Then
                v_action = 1
                v_action_mode = 2
                p_Flag_IO = 1
            End If
        End If
    Else
        '+++++++++++++++++++++++++++++++++
        strSQL = "SELECT COUNT(*), MAX(att_date), MIN(att_date)," & _
                        "((HOUR( TIMEDIFF(MAX(att_date),MIN(att_date)) )*60) + MINUTE( TIMEDIFF(MAX(att_date),MIN(att_date)) ))," & _
                        "Max(flag_io) " & _
                "FROM h_log_attendance WHERE employee_code = '" & v_emp_code & "' AND DATE(att_date) = DATE(ADDDATE('" & p_Att_Date & "',-1))"
        rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        v_recs1 = rs.Fields(0).Value
        v_max_date1 = IIf(IsNull(rs.Fields(1).Value), "", rs.Fields(1).Value)
        v_min_date1 = IIf(IsNull(rs.Fields(2).Value), "", rs.Fields(2).Value)
        v_mindiff1 = IIf(IsNull(rs.Fields(3).Value), 0, rs.Fields(3).Value)
        v_flag_io1 = IIf(IsNull(rs.Fields(4).Value), 0, rs.Fields(4).Value)
        
        rs.Close
        '+++++++++++++++++++++++++++++++
          
        If v_recs1 > 0 Then
            If v_flag_io1 = 0 Then
                v_action1 = 1
                v_yesterday = 1
                v_action_mode1 = 2
                p_Flag_IO = 1
            ElseIf v_flag_io1 = 1 Then
                v_action1 = 0
                v_yesterday = 0
                v_action_mode1 = 1
                p_Flag_IO = 0
            End If
        Else
            v_action1 = 0
            v_yesterday = 0
            v_action_mode1 = 1
            p_Flag_IO = 0
        End If
    End If
    
    If v_mode = 0 Then
       If v_yesterday = 0 Then
        '++++++++++++++++++++++++
            strSQL = "SELECT COUNT(*), MAX(start_time), MAX(end_time), MAX(min_break_in)," & _
                        "MAX(max_break_out), MAX(break_interval_minute) " & _
                    "from m_working_day " & _
                    "WHERE shift_code = '" & v_shift_code & "' AND day_code = WEEKDAY('" & p_Att_Date & "' ) AND flag_active = 1 " & _
                    "GROUP BY shift_code, day_code"
       '+++++++++++++++++++++++++++++
       ElseIf v_yesterday = 1 Then
            strSQL = "SELECT COUNT(*), MAX(start_time), MAX(end_time), MAX(min_break_in)," & _
                        "MAX(max_break_out), MAX(break_interval_minute) " & _
                    "from m_working_day " & _
                    "WHERE shift_code = '" & v_shift_code & "' AND day_code = WEEKDAY(ADDDATE('" & p_add_date & "',-1)) AND flag_active=1 " & _
                    "GROUP BY shift_code, day_code"
       End If
        rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_int_flag_day_active = IIf(IsNull(rs.Fields(0).Value), 0, rs.Fields(0).Value)
            v_start_time = IIf(IsNull(rs.Fields(1).Value), "", rs.Fields(1).Value)
            v_end_time = IIf(IsNull(rs.Fields(2).Value), "", rs.Fields(2).Value)
            v_break_in = IIf(IsNull(rs.Fields(3).Value), "", rs.Fields(3).Value)
            v_break_out = IIf(IsNull(rs.Fields(4).Value), "", rs.Fields(4).Value)
            v_break_interval = IIf(IsNull(rs.Fields(5).Value), 0, rs.Fields(5).Value)
        Else
            v_int_flag_day_active = 0
            v_start_time = ""
            v_end_time = ""
            v_break_in = ""
            v_break_out = ""
            v_break_interval = 0
        End If
        rs.Close
       '+++++++++++++++++++++++++++
       
    'Act
       If v_action_mode = 1 Or v_action_mode1 = 1 Then
            strSQL = "INSERT INTO h_attendance (employee_code, att_date, ip_address, enrollnumber," & _
                "shift_number, shift_code, start_time, end_time, time_in, time_out," & _
                "time_in_diff, time_out_diff, flag_io, flag_present, flag_late, flag_early," & _
                "absent_status, description, entry_date) VALUES " & _
                "('" & v_emp_code & "', '" & p_Att_Date & "', '" & p_ipAddress & "', '" & p_EnrollNumber & "'," & _
                    "'" & v_shift_number & "', '" & v_shift_code & "','" & Format(v_start_time, "yyyy-MM-dd hh:mm:ss") & "'," & _
                    "'" & Format(v_end_time, "yyyy-MM-dd hh:mm:ss") & "', '" & p_Att_Date & "', NULL, NULL, NULL, 0, 1, 0, 0, NULL, '-'," & _
                    "NOW())"
            CnG.Execute strSQL
            
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            strSQL = "SELECT flag_late, flag_early FROM h_attendance WHERE employee_code = '" & v_emp_code & "'"
            rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            flag_late = rs!flag_late
            flag_early = rs!flag_early
            rs.Close
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            
            ref_date = p_Att_Date
            
            v_str_buf = Left(DateDiff("d", DateValue(v_start_time), DateValue(p_Att_Date)), 1)
            
            strSQL = "UPDATE h_attendance SET time_in = '" & p_Att_Date & "'," & _
                "time_in_diff = CASE WHEN NOT '" & v_str_buf & "' = '-' " & _
                    "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(v_start_time, "yyyy-MM-dd hh:mm:ss") & "'),TIME('" & p_Att_Date & "')) ) " & _
                    "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & p_Att_Date & "'),TIME('" & Format(v_start_time, "yyyy-MM-dd hh:mm:ss") & "')) ) END," & _
                "flag_io = IF(flag_io >= 10 ,flag_io,'" & p_Flag_IO & "')," & _
                "flag_late = IF(flag_io >= 10,flag_late, CASE WHEN NOT '" & v_str_buf & "' = '-' THEN 0 ELSE 1 END) " & _
                "WHERE employee_code = '" & v_emp_code & "' " & _
                "AND DATE(att_date) = DATE('" & p_Att_Date & "')"
            CnG.Execute strSQL
            
            ref_date = p_Att_Date
          
       ElseIf v_action_mode = 2 And v_yesterday = 0 Then
          v_str_buf = Left(DateDiff("d", DateValue(p_Att_Date), DateValue(v_end_time)), 1)
            
            strSQL = "UPDATE h_attendance SET time_out = '" & p_Att_Date & "'," & _
                    "time_out_diff = CASE WHEN NOT '" & v_str_buf & "' = '-' " & _
                    "THEN CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(p_Att_Date, "yyyy-MM-dd hh:mm:ss") & "'), TIME('" & Format(v_end_time, "yyyy-MM-dd hh:mm:ss") & "')) ) " & _
                    "ELSE CONCAT(DATE('" & p_Att_Date & "'), ' ', TIMEDIFF(TIME('" & Format(v_end_time, "yyyy-MM-dd hh:mm:ss") & "'), TIME('" & p_Att_Date & "')) ) END," & _
                "flag_io = IF(flag_io >= 10, flag_io, '" & p_Flag_IO & "'), " & _
                "flag_early = IF(flag_io >= 10,flag_early, CASE WHEN NOT '" & v_str_buf & "' = '-' THEN 0 ELSE 1 END) " & _
                "WHERE employee_code = '" & v_emp_code & "' AND DATE(att_date) = DATE('" & p_Att_Date & "')"
            CnG.Execute strSQL
         
         ref_date = p_Att_Date
       End If
    
    ElseIf v_mode = 1 Then
        If v_action_mode1 = 1 Then
            If Hour(p_Att_Date) < 12 Then
                strSQL = "SELECT COUNT(*), MAX(start_time), MAX(end_time), MAX(min_break_in)," & _
                    "MAX(max_break_out), MAX(break_interval_minute) " & _
                    "from m_working_day " & _
                    "WHERE flag_moving = 1 AND moving_number = 1 AND day_code = WEEKDAY('" & p_Att_Date & "') " & _
                    "AND flag_active=1 " & _
                    "GROUP BY shift_code, day_code"
            ElseIf Hour(p_Att_Date) < 18 Then
                strSQL = "SELECT COUNT(*), MAX(start_time), MAX(end_time), MAX(min_break_in)," & _
                    "MAX(max_break_out), MAX(break_interval_minute) " & _
                    "from m_working_day " & _
                    "WHERE flag_moving = 1 AND moving_number = 2 AND day_code = WEEKDAY('" & p_Att_Date & "') AND flag_active=1 " & _
                    "GROUP BY shift_code, day_code"
            Else
                strSQL = "SELECT COUNT(*), MAX(start_time), MAX(end_time), MAX(min_break_in)," & _
                   "MAX(max_break_out), MAX(break_interval_minute) " & _
                   "from m_working_day " & _
                   "WHERE flag_moving = 1 AND moving_number = 3 AND day_code = WEEKDAY('" & p_Att_Date & "') " & _
                   "AND flag_active = 1 " & _
                   "GROUP BY shift_code, day_code"
            End If
            rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            v_int_flag_day_active = rs.Fields(0).Value
            v_start_time = rs.Fields(1).Value
            v_end_time = rs.Fields(2).Value
            v_break_in = rs.Fields(3).Value
            v_break_out = rs.Fields(4).Value
            v_break_interval = rs.Fields(5).Value
            rs.Close
        
        ElseIf v_action_mode = 2 And v_yesterday = 0 Then
            strSQL = "SELECT MAX(start_time), MAX(end_time) " & _
                    "from h_attendance " & _
                    "WHERE employee_code = '" & v_emp_code & "' AND DATE(att_date) = '" & p_Att_Date & "'"
            rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
            v_start_time = rs.Fields(0).Value
            v_end_time = rs.Fields(1).Value
            rs.Close
    
        ElseIf v_action_mode = 2 And v_yesterday = 1 Then
            strSQL = "SELECT MAX(start_time), MAX(end_time) " & _
                    "from h_attendance " & _
                    "WHERE employee_code = '" & v_emp_code & "' AND DATE(att_date) = DATE(ADDDATE('" & p_Att_Date & "',-1))"
            rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
            v_start_time = rs.Fields(0).Value
            v_end_time = rs.Fields(1).Value
            rs.Close
        ElseIf v_action_mode1 = 2 And v_yesterday = 1 Then
            strSQL = "SELECT MAX(start_time), MAX(end_time) " & _
                    "from h_attendance " & _
                    "WHERE employee_code = '" & v_emp_code & "' AND DATE(att_date) = DATE(ADDDATE('" & p_Att_Date & "',-1))"
            rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
            v_start_time = rs.Fields(0).Value
            v_end_time = rs.Fields(1).Value
            rs.Close
        End If
        If v_action_mode1 = 1 Then
            strSQL = "INSERT INTO h_attendance (employee_code, att_date, ip_address, enrollnumber, shift_number, shift_code," & _
                        "start_time, end_time, time_in, time_out,time_in_diff, time_out_diff," & _
                        "flag_io, flag_present, flag_late, flag_early, absent_status, description, entry_date) " & _
                    "VALUES " & _
                    "('" & v_emp_code & "','" & p_Att_Date & "','" & p_ipAddress & "','" & p_EnrollNumber & "','" & v_shift_number & "','" & v_shift_code & "'," & _
                    "'" & Format(v_start_time, "yyyy-MM-dd hh:mm:ss") & "','" & Format(v_end_time, "yyyy-MM-dd hh:mm:ss") & "','" & p_Att_Date & "', NULL, NULL, NULL, 0, 1, 0, 0, NULL, '-',now())"
            CnG.Execute strSQL
              
            v_str_buf = Left(DateDiff("h", Right(v_start_time, 8), Right(p_Att_Date, 8)), 1)
              
            strSQL = "UPDATE h_attendance SET time_in = '" & p_Att_Date & "'," & _
                    "time_in_diff = CASE WHEN NOT '" & v_str_buf & "' = '-' " & _
                        "THEN CONCAT(CAST('" & p_Att_Date & "' AS DATE), ' ', TIMEDIFF(RIGHT('" & v_start_time & "',8) , RIGHT('" & p_Att_Date & "',8)) ) " & _
                        "ELSE CONCAT(CAST(NEW.att_date AS DATE), ' ', TIMEDIFF(DATE(" & p_Att_Date & "), DATE(" & v_start_time & ")) ) END, " & _
                    "flag_io = IF(flag_io >= 10,flag_io, '" & p_Flag_IO & "'), " & _
                    "flag_late = IF(flag_io >= 10,flag_late, CASE WHEN NOT '" & v_str_buf & "' = '-' THEN 0 ELSE 1 END) " & _
                    "WHERE employee_code = '" & v_emp_code & "' AND DATE(att_date) = DATE('" & p_Att_Date & "')"
            CnG.Execute strSQL
            
            ref_date = p_Att_Date
              
        ElseIf v_action_mode = 2 And v_yesterday = 0 Then
            v_str_buf = Left(DateDiff("h", Right(p_Att_Date, 8), Right(v_end_time, 8)), 1)
            
            strSQL = "UPDATE h_attendance SET time_out = '" & p_Att_Date & "', " & _
                    "time_out_diff = CASE WHEN NOT '" & v_str_buf & "' = '-' " & _
                        "THEN CONCAT(CAST('" & p_Att_Date & "' AS DATE), ' ', TIMEDIFF(RIGHT('" & p_Att_Date & "',8), RIGHT(v_end_time,8)) ) " & _
                        "ELSE CONCAT(CAST('" & p_Att_Date & "' AS DATE), ' ', TIMEDIFF( RIGHT(" & v_end_time & ",8), RIGHT('" & p_Att_Date & "',8)) ) END, " & _
                        "flag_io = IF(flag_io >= 10,flag_io, '" & p_Flag_IO & "')," & _
                        "flag_early = IF(flag_io >= 10,flag_early, CASE WHEN NOT '" & v_str_buf & "' = '-' THEN 0 ELSE 1 END) " & _
                    "WHERE employee_code = '" & v_emp_code & "' AND DATE(att_date) = DATE('" & p_Att_Date & "');"
             CnG.Execute strSQL
             
             ref_date = p_Att_Date
              
        ElseIf v_action_mode = 2 And v_yesterday = 1 Then
            v_str_buf = Left(DateDiff("h", Right(p_Att_Date, 8), Right(v_end_time, 8)), 1)
            
            strSQL = "UPDATE h_attendance SET time_out = '" & p_Att_Date & "'," & _
                    "time_out_diff = CASE WHEN NOT '" & v_str_buf & "' = '-' " & _
                        "THEN CONCAT(CAST(ADDDATE('" & p_Att_Date & "',-1) AS DATE), ' ', TIMEDIFF('" & p_Att_Date & "', CONCAT(DATE(ADDDATE('" & p_Att_Date & "',0)),' ',RIGHT(v_end_time,8)) ) ) " & _
                        "ELSE CONCAT(CAST(ADDDATE('" & p_Att_Date & "',-1) AS DATE), ' ', TIMEDIFF( CONCAT(DATE(ADDDATE('" & p_Att_Date & "',0)),' ',RIGHT(v_end_time,8)) , '" & p_Att_Date & "') ) END, " & _
                    "flag_io = IF(flag_io >= 10,flag_io, '" & p_Flag_IO & "'), " & _
                    "flag_early = IF(flag_io>= 10,flag_early, CASE WHEN NOT '" & v_str_buf & "' = '-' THEN 0 ELSE 1 END) " & _
                    "WHERE employee_code = '" & v_emp_code & "' AND DATE(att_date) = DATE(ADDDATE('" & p_Att_Date & "',-1));"
            CnG.Execute strSQL
            
            ref_date = DateAdd("d", p_Att_Date, -1)
              
        ElseIf v_action_mode1 = 2 And v_yesterday = 1 Then
            v_str_buf = Left(DateDiff("h", Right(p_Att_Date, 8), Right(v_end_time, 8)), 1)
            
            strSQL = "UPDATE h_attendance SET time_out = '" & p_Att_Date & "'," & _
                    "time_out_diff = CASE WHEN NOT '" & v_str_buf & "' = '-' " & _
                        "THEN CONCAT(CAST(ADDDATE('" & p_Att_Date & "' ,-1) AS DATE),' ', TIMEDIFF('" & p_Att_Date & "', CONCAT(DATE(ADDDATE('" & p_Att_Date & "',0)),' ',RIGHT(v_end_time,8)) ) ) " & _
                        "ELSE CONCAT(CAST(ADDDATE('" & p_Att_Date & "',-1) AS DATE), ' ', TIMEDIFF( CONCAT(DATE(ADDDATE('" & p_Att_Date & "',0)),' ',RIGHT(v_end_time,8)) , '" & p_Att_Date & "') ) END, " & _
                    "flag_io = IF(flag_io >= 10,flag_io, '" & p_Flag_IO & "')," & _
                    "flag_early = IF(flag_io >= 10,flag_early, CASE WHEN NOT '" & v_str_buf & "' = '-' THEN 0 ELSE 1 END) " & _
                    "WHERE employee_code = '" & v_emp_code & "' AND DATE(att_date) = DATE(ADDDATE('" & p_Att_Date & "',-1));"
             CnG.Execute strSQL
             
             ref_date = DateAdd("d", p_Att_Date, -1)
        End If
    End If
End Sub
