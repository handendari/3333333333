Attribute VB_Name = "mdl_absensi"
Option Explicit

Public strConn As String
Public CnG As New ADODB.Connection

'Public Const pEncryptionPassword As String = "tomcat070509579"   '"HGSDYGDSLWREIUCJD938439402342"
Public Const pEncryptionPassword As String = "ssd12122012"  ' ID Pengenal Software
Public LOGIN_LEVEL, COMPANY_ACCESS As Integer
Public COMPANY_CODE, LOGIN_NAME, LOGIN_PASS, LOGIN_CODE, DEPARTMENT_CODE, EMPLOYEE_NAME As String
Public loginTime As Date
Public DATA_LEVEL As Integer

Public sbox(255) As Variant
Public Key(255) As Variant
Public blnUser_Read, blnUser_Add, blnUser_Edit, blnUser_Delete, blnUser_Posting, blnUser_Printing As Boolean

Private mFile As CFileSys

'===
Public Const headerMSG = "SSD TECHNOLOGY"
Public Const DataSize = 459
Public vMachineNumber As Integer
Public vEMachineNumber As Integer

Public BLN_RUNNING As Boolean
Public int_timer_tick As Integer, BLN_AUTO_LOG As Boolean

Public USER_NAME, USER_PASS, host_name, PORT, DB_NAME, DSN, DSN_CREATED As String
Public DESC, DRIVER_PATH, DRIVER_NAME, HD_SN, DRIVE_ID, RUNNING_COUNT As String
Public FG_IP_ADDRESS, FG_PORT_NUMBER As String
Public IS_LICENSED, MDI_STATE, FIRST_RUN As Boolean
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim vCompany As String
Dim vDept As String
Dim vDiv As String

Public Sub set_data_tdbcombo _
(ByRef Adodc1 As Adodc, ByRef TDBCombo1 As TDBCombo, ByVal str_sql As String)
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find str_sql, 0, adSearchForward, 1
If Not (Adodc1.Recordset.EOF = True Or Adodc1.Recordset.BOF = True) Then
    TDBCombo1.Bookmark = Adodc1.Recordset.AbsolutePosition
Else
    TDBCombo1.Text = ""
End If

End Sub

Public Sub set_data_tdbcombo_new _
(ByRef Adodc1 As Adodc, ByRef TDBCombo1 As TDBCombo, ByVal str_sql As String, str_col As String)
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find str_sql, 0, adSearchForward, 1
If Not (Adodc1.Recordset.EOF = True Or Adodc1.Recordset.BOF = True) Then
    TDBCombo1.Bookmark = Adodc1.Recordset.AbsolutePosition
Else
    TDBCombo1.Text = ""
End If

'set data
If TDBCombo1.ApproxCount > 0 Then
    TDBCombo1.Text = TDBCombo1.Columns(str_col).Value
End If
End Sub

Public Function DropAllComma(ByVal lstri As String) As String
Dim i As Integer
Dim lstrValue As String

lstri = Trim(lstri)
For i = 1 To Len(lstri)
    If Not Mid(lstri, i, 1) = "," Then
        lstrValue = lstrValue & Mid(lstri, i, 1)
    End If
Next i
DropAllComma = Trim(lstrValue)
End Function

Sub Main()
Dim i As Integer
Dim str_file_name As String

i = 0
If App.PrevInstance = True Then
    MsgBox "Program was run!", vbCritical, headerMSG
    End
End If


rerun:
If i > 0 Then
    If remove_mysql_dsn(0) Then
        'DSN_CREATED = "0"
    Else
        MsgBox "Error removing old dsn!", vbInformation, headerMSG
        End
    End If
End If


USER_NAME = USER_PASS = host_name = PORT = DB_NAME = DSN = DSN_CREATED = ""
DESC = DRIVER_PATH = DRIVER_NAME = HD_SN = DRIVE_ID = RUNNING_COUNT = ""
BLN_RUNNING = False
BLN_AUTO_LOG = False


Set mFile = New CFileSys
If Not mFile.readFile(mFile) Then
    MsgBox "It Works...", vbCritical, headerMSG
    End
End If


If Not CInt(DSN_CREATED) = 1 And FIRST_RUN = False Then
    If create_mysql_dsn(1) Then
        DSN_CREATED = "1"
    Else
        MsgBox "Error creating dsn!", vbInformation, headerMSG
        End
    End If
    If Not mFile.WriteFile(mFile) Then
        MsgBox "Error writing file...", vbInformation, headerMSG
        End
    End If

ElseIf Not CInt(DSN_CREATED) = 1 And FIRST_RUN = True Then
    If create_mysql_dsn(0) Then
        DSN_CREATED = "0"
    Else
        MsgBox "Error creating #1 dsn!", vbInformation, headerMSG
        End
    End If
'    If Not mFile.WriteFile(mFile) Then
'        MsgBox "Error writing file...", vbInformation, headerMSG
'        End
'    End If
End If


BLN_RUNNING = False
DRIVE_ID = "0"


If i > 0 Then
'    Dim targetTime2 As Date
'    targetTime2 = DateAdd("s", 5, Now())
'    While targetTime2 >= Now
'        Sleep 10 ' reduce CPU usage
'        'DoEvents ' keep app responsive
'    Wend
    Unload frm_loading_db
End If


If FIRST_RUN = False Then
    i = 1
End If
If Not connect_db(i) Then
    Call connect_db_err
End If


'If FIRST_RUN = True Then
'    str_file_name = App.Path & "\db_attpayroll.sql"
'    frm_loading_db.ProgressBar1.Value = 0
'    frm_loading_db.Show
'    Call MySQLRestore_proc(str_file_name, CnG)
'    'Unload frm_loading_db
'
'    FIRST_RUN = False
'    If Not mFile.WriteFile(mFile) Then
'        MsgBox "Error writing file...", vbInformation, headerMSG
'        End
'    End If
'
'    i = i + 1
'    If CnG.State = 1 Then CnG.Close
'    GoTo rerun
'End If


IS_LICENSED = True

'If check_license() Then
'    IS_LICENSED = True
'    If get_app_unpaid_status Then
'        i = MsgBox("Application Error !" _
'        & vbCrLf & "Please contact us (PT. Solusi Sentral Data - (031) 5616465 ) !", vbOK, headerMSG)
'        End
'    End If
'Else
'    IS_LICENSED = False
'    'If get_attendance_rec_number >= 30 Then
''        i = MsgBox("This trial product has reach data limit." _
''        & vbCrLf & "Please register this product at 'Help' menu!" _
''        & vbCrLf & "Do you want to register it?", vbOKCancel, headerMSG)
'        i = MsgBox("Please register this product!" _
'        & vbCrLf & "Do you want to register it?", vbYesNo, headerMSG)
'        If i = vbYes Then
'            MDI_STATE = False
'            frm_etc_register.Show
'            Exit Sub
'        Else
'            End
'        End If
'    'End If
'End If

frm_etc_login.Timer1.Enabled = True
frm_etc_login.Show
End Sub

'Lakukan cek jika aplikasi sudah di register tetapi belum lunas > 90 hari / 3 bln
Private Function get_app_unpaid_status() As Long
Dim rs1 As New ADODB.Recordset
Dim h As New HDSN

rs1.Open "SELECT date_add(register_date,interval +10 month) as status FROM m_license " _
        & "WHERE license_number = '" & Replace(RC4DeCryptASC(h.GetSerialNumber, pEncryptionPassword), "'", "''") & "'", CnG, adOpenStatic, adLockBatchOptimistic
If rs1.RecordCount > 0 Then
    If Format(Now, "yyyy-mm-dd") >= Format(rs1.Fields("status").Value, "yyyy-mm-dd") Then
        get_app_unpaid_status = True
    Else
        get_app_unpaid_status = False
    End If
End If
End Function

Private Function get_attendance_rec_number() As Long
Dim i As Long
Dim rs1 As New ADODB.Recordset

rs1.Open "select count(*) as rec_ from h_attendance", CnG, adOpenStatic, adLockBatchOptimistic
i = rs1.Fields("rec_").Value
get_attendance_rec_number = i
End Function

Public Sub run_it()
If connect_db(1) Then
    frm_etc_login.Timer1.Enabled = True
    frm_etc_login.Show
'    Form1.Show
Else
    Call connect_db_err
End If
End Sub

Public Sub run_it_bak()
If connect_db(1) Then

Dim rs As New ADODB.Recordset
rs.Open "select sysdate() as dt", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    If Format(rs!Dt, "yyyy-mm-dd") >= "2009-09-01" Then
        MsgBox "It Works...", vbCritical, headerMSG
        End
    End If
Else
    If Format(rs!Dt, "yyyy-mm-dd") >= "2009-09-01" Then
        MsgBox "Database miss-config...", vbCritical, headerMSG
        End
    End If
End If

    frm_etc_login.Timer1.Enabled = True
    frm_etc_login.Show
'    Form1.Show
Else
    Call connect_db_err
End If
End Sub

Public Sub connect_db_err()
MsgBox "Database connection error", vbCritical + vbOKOnly, headerMSG
'MsgBox Err.Description
End
End Sub

Public Function connect_db(ByVal int_db As Integer) As Boolean
On Error GoTo OPEN_DATABASE_CONNECTION_ERROR

'Public USER_NAME, USER_PASS, host_name, PORT, DB_NAME, DSN, DSN_CREATED As String
'Public DESC, DRIVER_PATH, DRIVER_NAME, HD_SN, DRIVE_ID, RUNNING_COUNT As String

CnG.ConnectionTimeout = 120
'strConn = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & IIf(int_db = 0, "dsn_mysql", DSN)
strConn = "DATABASE=" & DB_NAME & ";SERVER=" & host_name & ";port=" & PORT & ";UID=" & USER_NAME & ";PWD=" & USER_PASS & ";provider=MSDASQL.1;DRIVER=" & DRIVER_PATH & ";"
CnG.ConnectionString = strConn
CnG.CursorLocation = adUseClient
CnG.IsolationLevel = adXactSerializable
CnG.Open

If CnG.State Then
    connect_db = True
Else
    connect_db = False
End If
   
Exit Function
OPEN_DATABASE_CONNECTION_ERROR:
'If err.Number <> 0 Then
'   MsgBox err.Description, vbOKOnly, headerMSG
'End If
connect_db = False
End Function

Public Function get_server_datetime() As Date
On Error GoTo ErrHandler
Dim rs As New ADODB.Recordset
Dim str_sql As String

str_sql = "select now() AS dt"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
get_server_datetime = rs!Dt

If rs.State = 1 Then rs.Close

ErrHandler:
If Err.Description <> vbNullString Then
    MsgBox Err.Description, vbInformation, headerMSG
    If rs.State = 1 Then rs.Close
    get_server_datetime = Now
    Exit Function
End If
End Function

Public Sub loadData_userAccess_backup(ByVal str_FormCaption As String)
Dim rs As New ADODB.Recordset
Dim str_sql As String

If LOGIN_LEVEL = 100 Then
    blnUser_Read = True
    blnUser_Add = True
    blnUser_Edit = True
    blnUser_Delete = True
    blnUser_Posting = True
    Exit Sub
End If

blnUser_Read = False
blnUser_Add = False
blnUser_Edit = False
blnUser_Delete = False
blnUser_Posting = False

str_sql = "select F.form_code, F.form_name, F.form_title, " _
    & "F.allow_read, F.allow_add, F.allow_edit, F.allow_delete, F.allow_posting " _
    & "from m_user U, t_user F " _
    & "Where U.user_code = " & LOGIN_CODE & " and user_name = '" & LOGIN_NAME _
    & "' and user_pass = '" & LOGIN_PASS & "'"

rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then

    rs.MoveFirst
    While Not rs.EOF
        If UCase(rs.Fields("form_title").Value) = UCase(str_FormCaption) Then
            
            blnUser_Read = rs.Fields("allow_read").Value
            blnUser_Add = rs.Fields("allow_add").Value
            blnUser_Edit = rs.Fields("allow_edit").Value
            blnUser_Delete = rs.Fields("allow_delete").Value
            blnUser_Posting = rs.Fields("allow_posting").Value
            
            rs.MoveLast
        End If
    
        rs.MoveNext
    Wend

End If
End Sub

Public Sub set_company_mode(ByRef Adodc_company As Adodc, _
                            ByRef TDBCombo_company As TDBCombo, _
                            ByRef txt_company_name As TextBox)
                            
If Not Adodc_company.Recordset.RecordCount > 0 Then Exit Sub

Adodc_company.Recordset.MoveFirst
Adodc_company.Recordset.Find ("company_code='" & COMPANY_CODE & "'")  ', 0, adSearchForward, 1)
If Not (Adodc_company.Recordset.EOF = True Or Adodc_company.Recordset.BOF = True) Then
    TDBCombo_company.Bookmark = Adodc_company.Recordset.AbsolutePosition
    TDBCombo_company.Text = Adodc_company.Recordset.Fields("company_code").Value
    txt_company_name.Text = Adodc_company.Recordset.Fields("company_name").Value
Else
    TDBCombo_company.Text = "": txt_company_name.Text = ""
End If

If LOGIN_LEVEL = 100 Or COMPANY_ACCESS = 1 Then
    TDBCombo_company.Enabled = True
Else
    TDBCombo_company.Enabled = False
End If
End Sub

Public Sub set_company_mode_rs(ByRef rs As Recordset, _
                            ByRef TDBCombo_company As TDBCombo, _
                            ByRef txt_company_name As TextBox)

    If Not rs.RecordCount > 0 Then Exit Sub
    
    rs.MoveFirst
    rs.Find ("company_code='" & COMPANY_CODE & "'")  ', 0, adSearchForward, 1)
    If Not (rs.EOF = True Or rs.BOF = True) Then
        TDBCombo_company.Bookmark = rs.AbsolutePosition
        TDBCombo_company.Text = rs.Fields("company_code").Value
        txt_company_name.Text = rs.Fields("company_name").Value
    Else
        TDBCombo_company.Text = "": txt_company_name.Text = ""
    End If
    
    If LOGIN_LEVEL = 100 Or COMPANY_ACCESS = 1 Then
        TDBCombo_company.Enabled = True
    Else
        TDBCombo_company.Enabled = False
    End If
End Sub

Public Sub set_company_mode1(ByRef rs As Recordset, _
                            ByRef TDBCombo_company As TDBCombo, _
                            ByRef txt_company_name As TextBox)
                            
    If Not rs.RecordCount > 0 Then Exit Sub
    
    rs.MoveFirst
    rs.Find ("company_code='" & COMPANY_CODE & "'")  ', 0, adSearchForward, 1)
    If Not (rs.EOF = True Or rs.BOF = True) Then
        TDBCombo_company.Bookmark = rs.AbsolutePosition
        TDBCombo_company.Text = rs.Fields("company_code").Value
        txt_company_name.Text = rs.Fields("company_name").Value
    Else
        TDBCombo_company.Text = "": txt_company_name.Text = ""
    End If
    
    If LOGIN_LEVEL = 100 Or COMPANY_ACCESS = 1 Then
        TDBCombo_company.Enabled = True
    Else
        TDBCombo_company.Enabled = False
    End If
End Sub

Public Sub set_data_company(ByRef Adodc_company As Adodc, _
                            ByRef TDBCombo_company As TDBCombo, _
                            ByVal str_company_code As String, ByRef txt_company_name As TextBox)
                            
If Not Adodc_company.Recordset.RecordCount > 0 Then Exit Sub

Adodc_company.Recordset.MoveFirst
Adodc_company.Recordset.Find ("company_code='" & str_company_code & "'")  ', 0, adSearchForward, 1)
If Not (Adodc_company.Recordset.EOF = True Or Adodc_company.Recordset.BOF = True) Then
    TDBCombo_company.Bookmark = Adodc_company.Recordset.AbsolutePosition
    TDBCombo_company.Text = Adodc_company.Recordset.Fields("company_code").Value
    txt_company_name.Text = Adodc_company.Recordset.Fields("company_name").Value
Else
    TDBCombo_company.Text = "": txt_company_name.Text = ""
End If

If LOGIN_LEVEL = 100 Or COMPANY_ACCESS = 1 Then
    TDBCombo_company.Enabled = True
Else
    TDBCombo_company.Enabled = False
End If
End Sub

Public Function check_validate_tdbcombo(ByVal TDBCombo1 As TDBCombo) As Boolean
check_validate_tdbcombo = True
If Trim(TDBCombo1.Text) = "" Or Not TDBCombo1.Bookmark > 0 Or IsNull(TDBCombo1.Bookmark) = True Then
    check_validate_tdbcombo = False
End If
End Function

Public Sub get_setting_device()
Dim rs As New ADODB.Recordset

rs.Open "select * from m_device limit 0,1", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount = 1 Then
    FG_IP_ADDRESS = rs.Fields("ip_address").Value
    FG_PORT_NUMBER = rs.Fields("port_number").Value
Else
    FG_IP_ADDRESS = "10.155.9.75"
    FG_PORT_NUMBER = "4370"
End If
End Sub

Public Sub get_setting_auto_log()
Dim rs As New ADODB.Recordset

rs.Open "select * from s_download_mode where s_number=1", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount = 1 Then
    BLN_AUTO_LOG = IIf(rs.Fields("s_auto_download").Value = 1, True, False)
Else
    BLN_AUTO_LOG = False
End If

End Sub


Public Function get_inc_number _
(ByVal str_table As String, ByVal str_field As String, ByVal str_where As String) As Long
Dim rs As New ADODB.Recordset
Dim lng_no As Long

rs.Open "select max(" & str_field & ") as curr_no from " & str_table & " " _
& str_where, CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    If IsNull(rs.Fields("curr_no").Value) = True Then
        lng_no = 1
    Else
        lng_no = rs.Fields("curr_no").Value + 1
    End If
End If

get_inc_number = lng_no
End Function


Public Function check_validate_tdbgrid(ByRef grid1 As TDBGrid) As Boolean
If grid1.ApproxCount > 0 And grid1.Bookmark > 0 Then
    check_validate_tdbgrid = True
Else
    check_validate_tdbgrid = False
End If
End Function


'------------------

Private Sub clear_filter(ByRef Col7 As TrueOleDBGrid70.Column, _
ByRef TDBGrid7 As TDBGrid, ByRef Adodc7 As Adodc)
For Each Col7 In TDBGrid7.Columns
    Col7.FilterText = ""
Next Col7
Adodc7.Recordset.Filter = adFilterNone
End Sub


Private Function getFilter _
(ByRef Col7 As TrueOleDBGrid70.Column, ByRef Cols7 As TrueOleDBGrid70.Columns) As String
Dim tmp As String
Dim n As Integer

For Each Col7 In Cols7
    If Trim(Col7.FilterText) <> "" Then
        n = n + 1
        If n > 1 Then
            tmp = tmp & " AND "
        End If
        
        tmp = tmp & Col7.DataField & " LIKE '" & Col7.FilterText & "*'"
    End If
Next Col7
getFilter = tmp
End Function

Public Sub tdbgrid_filter(ByRef Cols7 As TrueOleDBGrid70.Columns, _
ByRef Col7 As TrueOleDBGrid70.Column, ByRef TDBGrid7 As TDBGrid, ByRef Adodc7 As Adodc)
On Error GoTo ErrHandler

Dim i As Integer

Set Cols7 = TDBGrid7.Columns
i = TDBGrid7.Col
TDBGrid7.HoldFields

Adodc7.Recordset.Filter = getFilter(Col7, Cols7)
TDBGrid7.Col = i
TDBGrid7.EditActive = True

TDBGrid7.SelStart = Len(TDBGrid7.Columns(i).FilterText)
If TDBGrid7.ApproxCount < 1 Then
    Call clear_filter(Col7, TDBGrid7, Adodc7)
    TDBGrid7.Col = i
End If

Exit Sub
ErrHandler:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter(Col7, TDBGrid7, Adodc7)
End Sub


Public Sub load_data_user_access(ByRef frm As Form)
Dim rs As New ADODB.Recordset
Dim str_sql As String

If LOGIN_LEVEL = 100 Then
    blnUser_Read = True
    blnUser_Add = True
    blnUser_Edit = True
    blnUser_Delete = True
    blnUser_Posting = True
    blnUser_Printing = True
    Exit Sub
End If

blnUser_Read = False
blnUser_Add = False
blnUser_Edit = False
blnUser_Delete = False
blnUser_Posting = False
blnUser_Printing = False

str_sql = "select F.sub_menu_code, F.form_name, F.form_title, " _
& "F.allow_read, F.allow_add, F.allow_edit, F.allow_delete, F.allow_post, F.allow_print " _
& "from m_user U left join m_employee b on U.employee_code = b.employee_code " _
& "join t_user F on U.user_code = F.level_code " _
& "Where U.user_code = '" & LOGIN_CODE & "' and user_name = '" & LOGIN_NAME _
& "' and user_pass = '" & LOGIN_PASS & "' and upper(F.form_name)='" & UCase(frm.name) & "'"

rs.Open str_sql, CnG, adOpenStatic, adLockBatchOptimistic
If rs.RecordCount > 0 Then

    'If UCase(rs.Fields("form_title").value) = UCase(str_FormCaption) Then
        
        blnUser_Read = rs.Fields("allow_read").Value
        blnUser_Add = rs.Fields("allow_add").Value
        blnUser_Edit = rs.Fields("allow_edit").Value
        blnUser_Delete = rs.Fields("allow_delete").Value
        blnUser_Posting = rs.Fields("allow_post").Value
        blnUser_Printing = rs.Fields("allow_print").Value
        
        'rs.MoveLast
    'End If

End If
End Sub
