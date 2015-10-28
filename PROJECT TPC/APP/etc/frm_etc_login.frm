VERSION 5.00
Begin VB.Form frm_etc_login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " USER LOGIN"
   ClientHeight    =   3390
   ClientLeft      =   4515
   ClientTop       =   2850
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_etc_login.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   870
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   390
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   840
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   2640
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   5415
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   0
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblNama 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   1005
      End
   End
   Begin prj_tpc.vbButton cmdOK 
      Height          =   450
      Left            =   3120
      TabIndex        =   10
      Top             =   2640
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
      BTYPE           =   14
      TX              =   "&Login"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_etc_login.frx":058A
      PICN            =   "frm_etc_login.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdCancel 
      Height          =   450
      Left            =   4470
      TabIndex        =   9
      Top             =   2640
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
      BTYPE           =   14
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_etc_login.frx":1638
      PICN            =   "frm_etc_login.frx":1654
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label LblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SSD ATTENDANCE && PAYROLL 3.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5940
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frm_etc_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LWA_BOTH = 3
Const LWA_ALPHA = 2
Const LWA_COLORKEY = 1
Const GWL_EXSTYLE = -20
Const WS_EX_LAYERED = &H80000

Private Declare Function GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
        (ByVal hwnd As Long, ByVal Color As Long, ByVal x As Byte, _
        ByVal Alpha As Long) As Boolean
        
Dim int_idx As Integer
Dim intOKTime As Integer

Private Sub cmdCancel_Click()
    timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub cmdOK_Click()
Dim rs As New ADODB.Recordset
Dim str_sql As String
Dim v_company As String

str_sql = "select a.user_code,a.user_name,a.full_name,a.user_pass,a.company_code," & _
                "a.flag_company_access,b.level_code, a.user_level,now() tglserver,a.department_code,a.division_code,a.flag_user,b.employee_name " & _
        "from m_user a left join m_employee b on a.employee_code = b.employee_code " & _
        "where user_name = '" & Replace(Trim(txtNama), "'", "''") & "' " & _
        "and user_pass = '" & EncryptINI(txtPassword.Text, pEncryptionPassword) & "'"
        
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.RecordCount > 0 Then
    loginTime = rs!tglServer
    LOGIN_CODE = rs!user_code
    LOGIN_NAME = rs!USER_NAME
    LOGIN_FULLNAME = IIf(IsNull(rs!full_name), "", rs!full_name)
    LOGIN_PASS = rs!USER_PASS
    LOGIN_LEVEL = rs!user_level
    DEPARTMENT_CODE = IIf(IsNull(rs!DEPARTMENT_CODE), "", rs!DEPARTMENT_CODE)
    DIVISION_CODE = IIf(IsNull(rs!DIVISION_CODE), "", rs!DIVISION_CODE)
    EMPLOYEE_NAME = IIf(IsNull(rs!EMPLOYEE_NAME), "", rs!EMPLOYEE_NAME)
    
'    If LOGIN_LEVEL = 100 Then
'        DATA_LEVEL = 1
'    Else
'        DATA_LEVEL = IIf(IsNull(rs!level_code), 10000, rs!level_code)
'    End If
    
    COMPANY_CODE = rs!COMPANY_CODE
    COMPANY_ACCESS = IIf(IsNull(rs!flag_user) = True, 0, rs!flag_user)
    
    If rs.RecordCount = 1 Then
        If Not Trim(rs!USER_NAME) = Trim(txtNama) Or _
            Not Trim(rs!USER_PASS) = EncryptINI(Trim(txtPassword.Text), pEncryptionPassword) Then
            MsgBox "Invalid account", vbCritical, headerMSG
            intOKTime = intOKTime + 1
            If intOKTime >= 3 Then End
            Exit Sub
        End If
    End If
rs.Close

    '++++++++++INSERT UNTUK LOG AKTIFITAS USER++++++++++++
    Dim clsFn As New clsFunction
    clsFn.InsertLog ("Login Program")
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++
Else
    MsgBox "Invalid account", vbCritical, headerMSG
    intOKTime = intOKTime + 1
    If intOKTime >= 3 Then End
    Exit Sub
End If
    
If BLN_RUNNING = True Then
    Unload Me
    
    Call close_all_forms
    Call mdi_absensi.re_create_menu
    
    If TypeDB <> 0 Then MsgBox "Softaware is Running on Local Database..", vbExclamation, headerMSG
Else
    Call mdi_absensi.re_create_menu
        
    Unload Me
    'CnG.Execute "call spg_leave_periode2 ('" & Format(Now, "yyyy-MM-dd") & "')"
    CnG.Execute "DELETE from temp_list"
    mdi_absensi.Show
    
    If TypeDB <> 0 Then MsgBox "Softaware is Running on Local Database..", vbExclamation, headerMSG
End If

End Sub

Private Sub close_all_forms()
    Dim frm As Form
    'First we want to loop through all the
    'Forms and close them (We close the current Form last)
    For Each frm In Forms
       'Make sure we arent looking at the current Form
       If frm.hwnd <> mdi_absensi.hwnd Then
           'Unload this Form
           frm.Hide
           Set frm = Nothing
       End If
   'Now get the next Form
   Next frm

   'Now unload the current Form
   'Unload Me
End Sub

Private Sub Command1_Click()
make_transparent Me.hwnd, 100
End Sub

Private Sub Command2_Click()
make_transparent Me.hwnd, 255
End Sub

Public Sub Form_Load()
    int_idx = 1
'    Call make_transparent(Me.hwnd, int_idx)
    
    Timer2.Enabled = False
    timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
int_idx = 1
timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub TxtNama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK_Click
End Sub

Sub make_transparent(hwnd As Long, idx As Integer)
On Error Resume Next

Dim lg_i As Long
lg_i = GetWindowLong(hwnd, GWL_EXSTYLE)

SetWindowLong hwnd, GWL_EXSTYLE, lg_i Or WS_EX_LAYERED
SetLayeredWindowAttributes hwnd, RGB(255, 255, 0), idx, LWA_ALPHA
Exit Sub
End Sub

Private Sub timer1_Timer()
On Error Resume Next
int_idx = int_idx + 0
If int_idx > 255 Then
    int_idx = 255
    timer1.Enabled = False
End If
'Call make_transparent(Me.hwnd, int_idx)
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
int_idx = int_idx - 5
If int_idx < 0 Then
    int_idx = 0
    Timer2.Enabled = False
    
    If BLN_RUNNING = True Then
        int_idx = 1
        Unload Me
    Else
        If CnG.State = 1 Then CnG.Close
        End
    End If
End If
'Call make_transparent(Me.hwnd, int_idx)
End Sub
