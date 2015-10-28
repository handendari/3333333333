Attribute VB_Name = "mdl_absensi"
Option Explicit

Public strConn As String
Public CnG As New ADODB.Connection

Public rs As New ADODB.Recordset
Public rscari As New ADODB.Recordset

Public Const pEncryptionPassword As String = "ssd12122012"  '"HGSDYGDSLWREIUCJD938439402342"
Public LOGIN_CODE, LOGIN_LEVEL, COMPANY_ACCESS As Integer
Public COMPANY_CODE, LOGIN_NAME, LOGIN_PASS As String
Public sbox(255) As Variant
Public Key(255) As Variant
Public blnUser_Read, blnUser_Add, blnUser_Edit, blnUser_Delete, blnUser_Posting, blnUser_Printing As Boolean

Private mFile As CFileSys

'===
Public Const headerMSG = "e-FREIGHT TECHNOLOGY"
Public Const DATASIZE = 459
Public vMachineNumber As Integer
Public vEMachineNumber As Integer

Public BLN_RUNNING As Boolean
Public int_timer_tick As Integer, BLN_AUTO_LOG As Boolean

Public ServerDB As String, UserDB As String
Public nmDB As String, passDB As String, PortDB As String

Public DESC, DRIVER_PATH, DRIVER_NAME, HD_SN, DRIVE_ID, RUNNING_COUNT As String
Public FG_IP_ADDRESS, FG_PORT_NUMBER As String

Public IS_LICENSED, MDI_STATE, FIRST_RUN As Boolean

Public SQL As String
Dim vCompany As String


Dim konek As New clsConnect
Public vLoadMode As Integer

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
    getServerVariable
    vLoadMode = 0
    
    If ServerDB = "" Then
        frm_etc_db_config.Show
    Else
        konek.Koneksi
    End If
End Sub

Public Sub connect_db_err()
MsgBox "Database connection error", vbCritical + vbOKOnly, headerMSG
End
End Sub

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

Public Sub set_data_company(ByRef rs As Recordset, _
                            ByRef TDBCombo_company As TDBCombo, _
                            ByVal str_company_code As String, _
                            ByRef txt_company_name As TextBox, _
                            ByVal int_mode As Integer)
                            
If Not rs.RecordCount > 0 Then Exit Sub

rs.MoveFirst
rs.Find ("company_code='" & str_company_code & "'")  ', 0, adSearchForward, 1)
If Not (rs.EOF = True Or rs.BOF = True) Then
    TDBCombo_company.Bookmark = rs.AbsolutePosition
    TDBCombo_company.Text = rs.Fields("company_code").Value
    txt_company_name.Text = rs.Fields("company_name").Value
Else
    TDBCombo_company.Text = "": txt_company_name.Text = ""
End If

If int_mode = 0 Then
    TDBCombo_company.Enabled = False
Else
    TDBCombo_company.Enabled = True
End If

End Sub

Public Sub set_company_mode(ByRef Adodc_company As Adodc, _
                            ByRef TDBCombo_company As TDBCombo, _
                            ByRef txt_company_name As TextBox)
                            
    If Not Adodc_company.Recordset.RecordCount > 0 Then Exit Sub
    
    If LOGIN_LEVEL = 100 Or COMPANY_ACCESS = 1 Then
        TDBCombo_company.Enabled = True
        
        vCompany = IIf(IsNull(COMPANY_CODE), "", COMPANY_CODE)
    Else
        TDBCombo_company.Enabled = False
        
        vCompany = IIf(IsNull(COMPANY_CODE), "", COMPANY_CODE)
    End If

    Adodc_company.Recordset.MoveFirst
    Adodc_company.Recordset.Find ("company_code='" & vCompany & "'")  ', 0, adSearchForward, 1)
    If Not (Adodc_company.Recordset.EOF = True Or Adodc_company.Recordset.BOF = True) Then
        TDBCombo_company.Bookmark = Adodc_company.Recordset.AbsolutePosition
        TDBCombo_company.Text = Adodc_company.Recordset.Fields("company_code").Value
        txt_company_name.Text = Adodc_company.Recordset.Fields("company_name").Value
    Else
        TDBCombo_company.Text = "": txt_company_name.Text = ""
    End If
    
'    If LOGIN_LEVEL = 100 Then
'        TDBCombo_company.Enabled = True
'    Else
'        TDBCombo_company.Enabled = False
'    End If
End Sub

Public Sub set_company_mode_rs(ByRef rs As Recordset, _
                            ByRef TDBCombo_company As TDBCombo, _
                            ByRef txt_company_name As TextBox)

    If Not rs.RecordCount > 0 Then Exit Sub
    
    If LOGIN_LEVEL = 100 Or COMPANY_ACCESS = 1 Then
        TDBCombo_company.Enabled = True
                
        vCompany = IIf(IsNull(COMPANY_CODE), "", COMPANY_CODE)
    Else
        TDBCombo_company.Enabled = False
        
        vCompany = IIf(IsNull(COMPANY_CODE), "", COMPANY_CODE)
    End If
    
    rs.MoveFirst
    rs.Find ("company_code = '" & vCompany & "'")  ', 0, adSearchForward, 1)
    If Not (rs.EOF = True Or rs.BOF = True) Then
        TDBCombo_company.Bookmark = rs.AbsolutePosition
        TDBCombo_company.Text = rs.Fields("company_code").Value
        txt_company_name.Text = rs.Fields("company_name").Value
    Else
        TDBCombo_company.Text = "": txt_company_name.Text = ""
    End If
End Sub

Public Sub set_data_tdbcombo_recordset _
(ByRef rs As Recordset, ByRef TDBCombo1 As TDBCombo, ByVal str_sql As String)
rs.MoveFirst
rs.Find str_sql, 0, adSearchForward, 1
If Not (rs.EOF = True Or rs.BOF = True) Then
    TDBCombo1.Bookmark = rs.AbsolutePosition
Else
    TDBCombo1.Text = ""
End If

End Sub

Public Function cek_validate_tdbcombo(ByVal TDBCombo1 As TDBCombo) As Boolean
cek_validate_tdbcombo = True
If Trim(TDBCombo1.Text) = "" Or Not TDBCombo1.Bookmark > 0 Or IsNull(TDBCombo1.Bookmark) = True Then
    cek_validate_tdbcombo = False
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

If rs.State Then rs.Close
str_sql = "select F.sub_menu_code, F.form_name, F.form_title, " _
& "F.allow_read, F.allow_add, F.allow_edit, F.allow_delete, F.allow_post, F.allow_print " _
& "from m_user U join m_employee b on U.employee_code = b.employee_code " _
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

Public Function check_validate_tdbcombo(ByVal TDBCombo1 As TDBCombo) As Boolean
check_validate_tdbcombo = True
'If Trim(TDBCombo1.Text) = "" Or Not TDBCombo1.Bookmark > 0 Or IsNull(TDBCombo1.Bookmark) = True Then
If Trim(TDBCombo1.Text) = "" Then
    check_validate_tdbcombo = False
End If
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

Public Sub getServerVariable()
On Error Resume Next
    If GetInitEntry("CONFIG", "A") <> "" Then
        ServerDB = DecryptINI(GetInitEntry("CONFIG", "A"), pEncryptionPassword)
        UserDB = DecryptINI(GetInitEntry("CONFIG", "B"), pEncryptionPassword)
        passDB = DecryptINI(GetInitEntry("CONFIG", "C"), pEncryptionPassword)
        nmDB = DecryptINI(GetInitEntry("CONFIG", "D"), pEncryptionPassword)
        PortDB = DecryptINI(GetInitEntry("CONFIG", "E"), pEncryptionPassword)
    End If
End Sub
