Attribute VB_Name = "mdl_etc"
'--

Public Const GWL_STYLE As Long = (-16&)
Public Const GWL_EXSTYLE As Long = (-20&)
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_MAXIMIZEBOX As Long = &H10000

Public Const WS_EX_TOOLWINDOW As Long = &H80&

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd&, ByVal nIndex&) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&) As Long
    

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd&, ByVal hWndInsertAfter&, ByVal x&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal wFlags&) As Long


Private Const QUERY_SUCCESS As Long = 0
Private Const ODBC_ADD_DSN = 1
Private Const ODBC_REMOVE_DSN = 3
Private Const ODBC_ADD_SYS_DSN = 4
Private Const vbAPINull = &O0
Private Const MYSQLOptionValue = 131242

#If Win32 Then
    Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
#Else
    Private Declare Function SQLConfigDataSource Lib "ODBCINST.DLL" (ByVal hwndParent As Integer, ByVal fRequest As Integer, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Integer
#End If

Private Declare Function EbExecuteLine Lib "vba6.dll" _
        (ByVal pStringToExec As Long, ByVal Foo1 As Long, _
        ByVal Foo2 As Long, ByVal fCheckOnly As Long) As Long


Function FExecuteCode(stCode As String, _
            Optional fCheckOnly As Boolean) As Boolean

    FExecuteCode = EbExecuteLine(StrPtr(stCode), 0&, 0&, Abs(fCheckOnly)) = 0
    
End Function


Public Function create_mysql_dsn(ByVal int_db As Integer) As Boolean
On Error GoTo ErrorHandler

Dim svlServerName, svlPortNumber, svlDSNName, svlUserName, svlPassword, svlDatabaseName, svlDSNDriver As String
Dim bvlIsUserDSN As Boolean

svlServerName = host_name
svlPortNumber = PORT
svlDSNName = IIf(int_db = 0, "dsn_mysql", DSN)
svlUserName = USER_NAME
svlPassword = USER_PASS
svlDatabaseName = IIf(int_db = 0, "mysql", DB_NAME)
svlDSNDriver = DRIVER_NAME
bvlIsUserDSN = True

create_mysql_dsn = False
#If Win32 Then
  Dim intRet As Long
#Else
    Dim intRet As Integer
#End If
    
strAttributes = _
    "SERVER=" & svlServerName & Chr$(0) & _
    "DSN=" & svlDSNName & Chr$(0) & _
    "DESCRIPTION=Created Automatically on " & Format(Date, "DD-MMM-YYYY") & Chr$(0) & _
    "DATABASE=" & svlDatabaseName & Chr$(0) & _
    "UID=" & svlUserName & Chr$(0) & _
    "PWD=" & svlPassword & Chr$(0) & _
    "PORT=" & svlPortNumber & Chr$(0) & _
    "OPTION=" & MYSQLOptionValue & Chr$(0)
    
If bvlIsUserDSN Then
    intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, svlDSNDriver, strAttributes)
Else
    intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, svlDSNDriver, strAttributes)
End If

If intRet Then
    create_mysql_dsn = True
    'err.Raise vbObjectError + 513, , "DSN " & svlDSNName & " Sucessfully Created..."
Else
    'err.Raise vbObjectError + 514, , "Failed to create DSN " & svlDSNName
    create_mysql_dsn = False
End If
    
    
Exit Function
ErrorHandler:
create_mysql_dsn = False
MsgBox ("Error :" & err.Number & " - " & err.Description)
End Function

Public Function remove_mysql_dsn(ByVal int_db As Integer) As Boolean
On Error GoTo ErrorHandler

Dim svlServerName, svlPortNumber, svlDSNName, svlUserName, svlPassword, svlDatabaseName, svlDSNDriver As String
Dim bvlIsUserDSN As Boolean

svlServerName = host_name
svlPortNumber = PORT
svlDSNName = IIf(int_db = 0, "dsn_mysql", DSN)
svlUserName = USER_NAME
svlPassword = USER_PASS
svlDatabaseName = IIf(int_db = 0, "mysql", DB_NAME)
svlDSNDriver = DRIVER_NAME
bvlIsUserDSN = True

remove_mysql_dsn = False

#If Win32 Then
    Dim intRet As Long
#Else
    Dim intRet As Integer
#End If

Dim strAttributes As String

strAttributes = _
    "SERVER=" & svlServerName & Chr$(0) & _
    "DSN=" & svlDSNName & Chr$(0) & _
    "DESCRIPTION=Created Automatically on " & Format(Date, "DD-MMM-YYYY") & Chr$(0) & _
    "DATABASE=" & svlDatabaseName & Chr$(0) & _
    "UID=" & svlUserName & Chr$(0) & _
    "PWD=" & svlPassword & Chr$(0) & _
    "PORT=" & svlPortNumber & Chr$(0) & _
    "OPTION=" & MYSQLOptionValue & Chr$(0)
'See driver documentation for a complete list of attributes.
strAttributes = "DSN=" & svlDSNName & Chr$(0)
'To show dialog, use Form1.Hwnd instead of vbAPINull.
intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, svlDSNDriver, strAttributes)

If intRet Then
    remove_mysql_dsn = True
    'err.Raise vbObjectError + 513, , "DSN " & svlDSNName & " Sucessfully Created..."
Else
    'err.Raise vbObjectError + 514, , "Failed to create DSN " & svlDSNName
    remove_mysql_dsn = False
End If
    
    
Exit Function
ErrorHandler:
remove_mysql_dsn = False
MsgBox ("Error :" & err.Number & " - " & err.Description)
End Function






Public Sub get_entry_user_access(ByVal str_FormCaption As String)
If blnTop = True Then
    blnUser_Tambah = True
    blnUser_Ubah = True
    blnUser_Hapus = True
    Exit Sub
End If
rsUser.MoveFirst

While Not rsUser.EOF
    If UCase(rsUser.Fields("Nama_Form").Value) = UCase(str_FormCaption) Then
        If rsUser.Fields("Baca").Value = "Y" Then
            blnUser_Baca = True
        ElseIf rsUser.Fields("Baca").Value = "T" Then
            blnUser_Baca = False
        End If
        If rsUser.Fields("Tambah").Value = "Y" Then
            blnUser_Tambah = True
        ElseIf rsUser.Fields("Tambah").Value = "T" Then
            blnUser_Tambah = False
        End If
        If rsUser.Fields("Ubah").Value = "Y" Then
            blnUser_Ubah = True
        ElseIf rsUser.Fields("Ubah").Value = "T" Then
            blnUser_Ubah = False
        End If
        If rsUser.Fields("Hapus").Value = "Y" Then
            blnUser_Hapus = True
        ElseIf rsUser.Fields("Hapus").Value = "T" Then
            blnUser_Hapus = False
        End If
        rsUser.MoveLast
    End If
    
    rsUser.MoveNext
Wend
End Sub



Public Sub set_sizeable_form(ByRef frm As Form)
Dim rs1 As New ADODB.Recordset

rs1.Open "select * from m_sub_menu where form_title='" & frm.Caption & "'", CnG, adOpenStatic, adLockReadOnly
If rs1.RecordCount = 1 Then
    If rs1.Fields("flag_sizeable").Value <> 1 Then
        Exit Sub
    End If
Else
    Exit Sub
End If

Call SetWindowLong(frm.hWnd, GWL_STYLE, GetWindowLong(frm.hWnd, GWL_STYLE) Xor _
                              (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
Call SetWindowLong(frm.hWnd, GWL_EXSTYLE, GetWindowLong(frm.hWnd, GWL_EXSTYLE) Xor WS_EX_TOOLWINDOW)
Call SetWindowPos(frm.hWnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

Public Sub set_sizeable_form_by_hwnd(ByVal form_name As String)
Dim rs1 As New ADODB.Recordset
Dim str1 As String


rs1.Open "select * from m_sub_menu where form_name='" & form_name & "'", CnG, adOpenStatic, adLockReadOnly
If rs1.RecordCount = 1 Then
    If rs1.Fields("flag_sizeable").Value <> 1 Then
        Exit Sub
    End If
Else
    Exit Sub
End If

str1 = "mdi_absensi.txt_hwnd.Text = str(" & form_name & ".hwnd)"
FExecuteCode (str1)

Call SetWindowLong(Val(mdi_absensi.txt_hwnd.Text), GWL_STYLE, _
    GetWindowLong(Val(mdi_absensi.txt_hwnd.Text), GWL_STYLE) Xor _
    (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
Call SetWindowLong(Val(mdi_absensi.txt_hwnd.Text), GWL_EXSTYLE, _
    GetWindowLong(Val(mdi_absensi.txt_hwnd.Text), GWL_EXSTYLE) Xor WS_EX_TOOLWINDOW)
Call SetWindowPos(Val(mdi_absensi.txt_hwnd.Text), 0&, 0&, 0&, 0&, 0&, _
    SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

