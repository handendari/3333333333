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
      Left            =   2040
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   9
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_etc_login.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Picture         =   "frm_etc_login.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
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
      TabIndex        =   4
      Top             =   840
      Width           =   5415
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
      Begin VB.Label lblNama 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   1170
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "e-FREIGHT ATTENDANCE SYSTEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   5460
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
      TabIndex        =   8
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
        "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
        (ByVal hWnd As Long, ByVal Color As Long, ByVal X As Byte, _
        ByVal Alpha As Long) As Boolean
        
Dim int_idx As Integer
Dim intOKTime As Integer



Private Sub cmdCancel_Click()
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub cmdOK_Click()
Dim rs As New ADODB.Recordset
Dim str_sql As String

str_sql = "select * from m_user where user_name = '" & Trim(txtNama) _
& "' " _
& "and user_pass = '" & Replace(EncryptINI(txtPassword.Text, pEncryptionPassword), "'", "") & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.RecordCount > 0 Then
    LOGIN_CODE = rs!user_code
    LOGIN_NAME = rs!USER_NAME
    LOGIN_PASS = rs!USER_PASS
    LOGIN_LEVEL = rs!user_level
    COMPANY_CODE = rs!COMPANY_CODE
    COMPANY_ACCESS = IIf(IsNull(rs!flag_user), 0, rs!flag_user)
    
    If rs.RecordCount = 1 Then
        If Not Trim(rs!USER_NAME) = Trim(txtNama) Or _
            Not Trim(rs!USER_PASS) = EncryptINI(Replace(Trim(txtPassword.Text), Chr(39), Chr(96)), pEncryptionPassword) Then
            MsgBox "Invalid account", vbCritical, headerMSG
            intOKTime = intOKTime + 1
            If intOKTime >= 3 Then End
            Exit Sub
        End If
    End If
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
Else
    Call mdi_absensi.re_create_menu
        
    Unload Me
    mdi_absensi.Show
End If

End Sub

Private Sub close_all_forms()
    Dim frm As Form
    'First we want to loop through all the
    'Forms and close them (We close the current Form last)
    For Each frm In Forms
       'Make sure we arent looking at the current Form
       If frm.hWnd <> mdi_absensi.hWnd Then
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
make_transparent Me.hWnd, 100
End Sub

Private Sub Command2_Click()
make_transparent Me.hWnd, 255
End Sub

Public Sub Form_Load()
'int_idx = 1: Call make_transparent(Me.hWnd, int_idx)

'Timer2.Enabled = False
'Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'int_idx = 1
'Timer1.Enabled = False
'Timer2.Enabled = True
End Sub

Private Sub TxtNama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Sub make_transparent(hWnd As Long, idx As Integer)
On Error Resume Next

Dim lg_i As Long
lg_i = GetWindowLong(hWnd, GWL_EXSTYLE)

SetWindowLong hWnd, GWL_EXSTYLE, lg_i Or WS_EX_LAYERED
SetLayeredWindowAttributes hWnd, RGB(255, 255, 0), idx, LWA_ALPHA
Exit Sub
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
int_idx = int_idx + 5
If int_idx > 255 Then
    int_idx = 255
    Timer1.Enabled = False
End If
Call make_transparent(Me.hWnd, int_idx)
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
Call make_transparent(Me.hWnd, int_idx)
End Sub
