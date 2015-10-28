VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_loading 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm_loading.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3680
      _ExtentX        =   6482
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Max             =   12
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Loading... Please wait!"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1605
   End
End
Attribute VB_Name = "frm_loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str_license As String


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
str_license = str_license & Chr(KeyCode)
End Sub

Private Sub Form_Load()
str_license = ""
ProgressBar1.Value = 0
Timer1.Enabled = True
End Sub

'Private Sub Timer1_Timer()
'
'If ProgressBar1.Value = 12 Then
'    Timer1.Enabled = False
'
'    Dim rs As New ADODB.Recordset
'
'
'    If get_license(str_license) Then
'        If CnG.State Then
'            Call set_license
'        Else
''            Call connect_db_err
'        End If
'
'        Unload Me
'        frm_etc_login.Timer1.Enabled = True
'        frm_etc_login.Show
'    Else
'        If check_license() Then
'            Unload Me
'            frm_etc_login.Timer1.Enabled = True
'            frm_etc_login.Show
'        Else
'            Call no_license
'        End If
'    End If
'Else
'    ProgressBar1.Value = ProgressBar1.Value + 1
'End If

'If ProgressBar1.Value = 12 Then
'    Timer1.Enabled = False
'
'    If get_license(str_license) Then
'        Call set_license
'        Unload Me
'        Call run_it
'    Else
'        Dim afile As New CFileSys
'
'        RUNNING_COUNT = Int(Trim(CStr(CLng(RC4DeCryptASC(RUNNING_COUNT, pEncryptionPassword)) + 1)))
'        RUNNING_COUNT = RC4DeCryptASC(RUNNING_COUNT, pEncryptionPassword)
'        If Not afile.WriteFile(afile) Then
'            MsgBox "Error writing file...", vbInformation, headerMSG
'        End If
'        End
'    End If
'Else
'    ProgressBar1.Value = ProgressBar1.Value + 1
'End If
'End Sub








