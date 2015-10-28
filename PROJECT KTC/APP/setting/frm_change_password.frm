VERSION 5.00
Begin VB.Form frm_change_password 
   Caption         =   "Change Password"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form2"
   ScaleHeight     =   3570
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frPerubahanPasswd 
      Caption         =   "Form Perubahan Password"
      Height          =   3225
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   3930
         TabIndex        =   8
         Top             =   2460
         Width           =   1575
      End
      Begin VB.CommandButton cmdGantiPasswd 
         Caption         =   "&Ganti Password"
         Default         =   -1  'True
         Height          =   375
         Left            =   2340
         TabIndex        =   7
         Top             =   2460
         Width           =   1605
      End
      Begin VB.TextBox txt_konfirm_pass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2100
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1590
         Width           =   3405
      End
      Begin VB.TextBox txt_new_pass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2100
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1230
         Width           =   3405
      End
      Begin VB.TextBox txt_old_pass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2100
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   870
         Width           =   3405
      End
      Begin VB.Label Label3 
         Caption         =   "Konfirmasi Password"
         Height          =   285
         Left            =   330
         TabIndex        =   3
         Top             =   1650
         Width           =   1665
      End
      Begin VB.Label Label2 
         Caption         =   "Password Baru"
         Height          =   285
         Left            =   330
         TabIndex        =   2
         Top             =   1290
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Password Lama"
         Height          =   285
         Left            =   330
         TabIndex        =   1
         Top             =   930
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_change_password"
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
strsql = "SELECT user_pass FROM m_user WHERE user_name = '" & LOGIN_NAME & "'"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
    v_password = RC4DeCryptASC(Replace(Trim(rs!USER_PASS), Chr(39), Chr(96)), pEncryptionPassword)
End If

If txt_old_pass <> v_password Then
    MsgBox "Password Lama Tidak Sesuai!" & Chr(13) & _
        "Silahkan Coba Lagi..", vbExclamation, headerMSG
    txt_old_pass.Text = ""
    txt_old_pass.SetFocus
    Exit Sub
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'+++++++++++++++ Cek Password Baru +++++++++++++++++++++++
If txt_new_pass <> txt_konfirm_pass Then
    MsgBox "Password Baru Tidak Sesuai!" & Chr(13) & _
        "Silahkan Coba Lagi..", vbExclamation, headerMSG
    txt_konfirm_pass.Text = ""
    txt_konfirm_pass.SetFocus
    Exit Sub
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++

strsql = "UPDATE m_user SET user_pass = '" & RC4DeCryptASC(Replace(Trim(txt_new_pass.Text), Chr(39), Chr(96)), pEncryptionPassword) & "' " _
        & "WHERE user_name = '" & LOGIN_NAME & "'"
CnG.Execute strsql

MsgBox "Ubah Password Sukses!", vbInformation, headerMSG

txt_konfirm_pass.Text = ""
txt_new_pass.Text = ""
txt_old_pass.Text = ""

Unload Me
End Sub
