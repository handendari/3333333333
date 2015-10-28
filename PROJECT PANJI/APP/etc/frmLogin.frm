VERSION 5.00
Begin VB.Form frm_login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   USER LOGIN"
   ClientHeight    =   3240
   ClientLeft      =   4515
   ClientTop       =   2850
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   120
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   600
      Top             =   2640
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmLogin.frx":058A
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
      Left            =   3240
      Picture         =   "frmLogin.frx":0B14
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
      Left            =   120
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
         Text            =   "tomcat"
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
         Text            =   "123"
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
      BackStyle       =   0  'Transparent
      Caption         =   "PT. PANCA WANA INDONESIA"
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
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   5415
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
      Width           =   6015
   End
End
Attribute VB_Name = "frm_login"
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
        (ByVal hWnd As Long, ByVal color As Long, ByVal X As Byte, _
        ByVal alpha As Long) As Boolean
Dim TransparanDonk As Integer
Dim intOKTime As Integer
Dim Batal As Boolean

Private Sub CmdCancel_Click()
Batal = True
If cn.State = 1 Then cn.Close
End
End Sub

Private Sub cmdOK_Click()
Dim rs1 As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim lstri, strNamaUser As String

Set cmd1.ActiveConnection = cn
cmd1.CommandText = "select * from m_user where nama_user = '" _
& Trim(txtNama) & "' and pass = '" & RC4DeCryptASC(Replace(Trim(txtPassword.Text), Chr(39), Chr(96)), pEncryptionPassword) & "'"
rs1.CursorLocation = adUseClient
rs1.Open cmd1, , adOpenStatic, adLockBatchOptimistic
lstri = "Nama atau Password anda belum te-registry dalam program ini"

If rs1.RecordCount > 0 Then
    LOGIN_ID = Trim(rs1!id_user)
    LOGIN_NAME = Trim(rs1!nama_user)
    LOGIN_PWD = Trim(rs1!Pass)
    If rs1.RecordCount = 1 Then
        If Not Trim(rs1!nama_user) = Trim(txtNama) Or _
            Not Trim(rs1!Pass) = RC4DeCryptASC(Replace(Trim(txtPassword.Text), Chr(39), Chr(96)), pEncryptionPassword) Then
            MsgBox lstri, vbCritical, "Invalid account"
            intOKTime = intOKTime + 1
            If intOKTime >= 3 Then End
            Exit Sub
        End If
    End If
Else
    MsgBox lstri, vbCritical, "Invalid account"
    intOKTime = intOKTime + 1
    If intOKTime >= 3 Then End
    Exit Sub
End If

'' --- IDENTIFYING USER ---
'strNamaUser = rs1.Fields("nama_user").Value
'
'If rs1.Fields("user_level").Value = 100 Then 'Admin
'    blnTop = True
'    Set cmdUser.ActiveConnection = CnG
'    If rsUser.State = 1 Then rsUser.Close
'
'    cmdUser.CommandText = "select nama_form from m_form order by nama_form"
'    rsUser.CursorLocation = adUseClient
'    rsUser.Open cmdUser, , adOpenStatic, adLockBatchOptimistic
'
'    ' --- Enable all shortCut menu ---
'    If rsUser.RecordCount > 0 Then
'        rsUser.MoveFirst
'        While Not rsUser.EOF
'            Call EnableMenu(rsUser.Fields("nama_form").Value, True)
'            rsUser.MoveNext
'        Wend
'    End If
'
'    Unload Me
'    Main_Citra_Yasindo.Show
'    GoTo PassIdent
'
'ElseIf rs1.Fields("user_level").Value = 1 Then 'User
'    blnTop = False
'    Set rsUser = Nothing
'    Set cmdUser.ActiveConnection = CnG
'
'    cmdUser.CommandText = "select M.nama_user, M.Password, M.User_Level, D.nama_form, " _
'    & "D.Baca, D.Tambah, D.Ubah, D.Hapus, D.posting " _
'    & "from m_User M, t_User D " _
'    & "Where M.nama_user = D.nama_user " _
'    & "and M.nama_user = '" _
'    & rs1.Fields("nama_user").Value & "'"
'    rsUser.CursorLocation = adUseClient
'    rsUser.Open cmdUser, , adOpenStatic, adLockBatchOptimistic
'
'    ' --- Disable shortCut menu having Baca = 'T' ---
'    Dim i As Integer, blni As Boolean
'    rsUser.MoveFirst
'    While Not rsUser.EOF
'        If rsUser.Fields("Baca").Value = "No" Then
'            blni = False
'        ElseIf rsUser.Fields("Baca").Value = "Yes" Then
'            blni = True
'        End If
'        Call EnableMenu(rsUser.Fields("nama_form").Value, blni)
'        rsUser.MoveNext
'    Wend
'
'    '--------------------------
'    'Disable Form while Form not in rsUser
'    Dim rsForm As New ADODB.Recordset
'    SQL = "select * from m_form where kode_form " & _
'          "not in (select kode_form from t_User Where nama_user = '" _
'          & rs1.Fields("nama_user").Value & "')"
'
'    Call ExecQuery(rsForm, SQL, 3)
'    If rsForm.RecordCount > 0 Then
'        While Not rsForm.EOF
'            Call EnableMenu(rsForm.Fields("nama_form").Value, False)
'            rsForm.MoveNext
'        Wend
'    End If
'    '--------------------------
'    blnActivate = True
    
    mdi.Show
    Unload Me
'End If
'Exit Sub
  
'PassIdent:
'blnTop = True
'blnActivate = True
End Sub

Private Sub Form_Activate()
txtNama.SetFocus
End Sub

Public Sub Form_Load()
Timer2.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Cancel = 1
Timer1.Enabled = False
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
If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Sub TransparanBro(hWndBro As Long, TransBro As Integer)
On Error Resume Next

Dim OKBro As Long
OKBro = GetWindowLong(hWndBro, GWL_EXSTYLE)

SetWindowLong hWndBro, GWL_EXSTYLE, OKBro Or WS_EX_LAYERED
SetLayeredWindowAttributes hWndBro, RGB(255, 255, 0), TransBro, LWA_ALPHA
Exit Sub
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
TransparanDonk = TransparanDonk + 5
If TransparanDonk > 255 Then TransparanDonk = 255: Timer1.Enabled = False
TransparanBro Me.hWnd, TransparanDonk
Me.Show
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
TransparanDonk = TransparanDonk - 5
If TransparanDonk < 0 Then TransparanDonk = 0
Timer2.Enabled = False
If Batal = True Then End
TransparanBro Me.hWnd, TransparanDonk
End Sub
