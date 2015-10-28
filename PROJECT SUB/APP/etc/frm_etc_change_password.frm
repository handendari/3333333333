VERSION 5.00
Begin VB.Form frm_etc_change_password 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   Icon            =   "frm_etc_change_password.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPerubahanPasswd 
      Caption         =   "Change Password"
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.TextBox txt_konfirm_pass 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2100
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1440
         Width           =   3405
      End
      Begin VB.TextBox txt_new_pass 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2100
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1080
         Width           =   3405
      End
      Begin VB.TextBox txt_old_pass 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2100
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   3405
      End
      Begin prj_absensi.vbButton cmdGantiPasswd 
         Height          =   450
         Left            =   2850
         TabIndex        =   7
         Top             =   2190
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   794
         BTYPE           =   14
         TX              =   "&Change"
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
         MICON           =   "frm_etc_change_password.frx":058A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_absensi.vbButton cmdBatal 
         Height          =   450
         Left            =   4200
         TabIndex        =   8
         Top             =   2190
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
         MICON           =   "frm_etc_change_password.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Confirm Password"
         Height          =   285
         Left            =   330
         TabIndex        =   3
         Top             =   1500
         Width           =   1665
      End
      Begin VB.Label Label2 
         Caption         =   "New Password"
         Height          =   285
         Left            =   330
         TabIndex        =   2
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Old Password"
         Height          =   285
         Left            =   330
         TabIndex        =   1
         Top             =   780
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_etc_change_password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdGantiPasswd_Click()
Dim strsql As String
Dim rs As New ADODB.Recordset
Dim v_password As String

'+++++++++++++++ Cek Password Lama +++++++++++++++++++++++
If rs.State Then rs.Close
strsql = "SELECT user_pass FROM m_user WHERE user_name = '" & LOGIN_NAME & "'"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
    v_password = DecryptINI(Trim(rs!USER_PASS), pEncryptionPassword)
End If

If txt_old_pass <> v_password Then
    MsgBox "Invalid old password!" & Chr(13) & _
        "Please try again..", vbExclamation, headerMSG
    txt_old_pass.Text = ""
    txt_old_pass.SetFocus
    Exit Sub
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'+++++++++++++++ Cek Password Baru +++++++++++++++++++++++
If txt_new_pass <> txt_konfirm_pass Then
    MsgBox "Invalid new password!" & Chr(13) & _
        "Please try again..", vbExclamation, headerMSG
    txt_konfirm_pass.Text = ""
    txt_konfirm_pass.SetFocus
    Exit Sub
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++

strsql = "UPDATE m_user SET user_pass = '" & EncryptINI(Trim(txt_new_pass.Text), pEncryptionPassword) & "' " _
        & "WHERE user_name = '" & LOGIN_NAME & "'"
CnG.Execute strsql

MsgBox "Change password successfully!", vbInformation, headerMSG

txt_konfirm_pass.Text = ""
txt_new_pass.Text = ""
txt_old_pass.Text = ""

Unload Me
End Sub

Private Sub txt_konfirm_pass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdGantiPasswd_Click
End Sub
