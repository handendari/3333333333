VERSION 5.00
Begin VB.Form frm_etc_smtp_server 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail (SMTP Server)"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   Icon            =   "frmSMTPServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "General"
      Height          =   1125
      Left            =   2340
      TabIndex        =   17
      Top             =   2910
      Width           =   4575
      Begin VB.TextBox txtSenderName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1770
         TabIndex        =   5
         Top             =   270
         Width           =   2505
      End
      Begin VB.TextBox txtSenderEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1770
         TabIndex        =   6
         Top             =   600
         Width           =   2505
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "SENDER NAME*"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   330
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "SENDER EMAIL*"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   630
         Width           =   1260
      End
   End
   Begin prj_tpc.vbButton cmdSave 
      Height          =   450
      Left            =   4260
      TabIndex        =   8
      Top             =   4185
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
      BTYPE           =   14
      TX              =   "&Save"
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
      MICON           =   "frmSMTPServer.frx":1082
      PICN            =   "frmSMTPServer.frx":109E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "SMTP Server Configuration"
      Height          =   2805
      Left            =   2340
      TabIndex        =   10
      Top             =   90
      Width           =   4575
      Begin VB.CheckBox chkReqAuth 
         Caption         =   "Req Authentication"
         Height          =   195
         Left            =   2340
         TabIndex        =   20
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkSSL 
         Caption         =   "No"
         Height          =   195
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   1005
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   690
         Width           =   945
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "="
         TabIndex        =   4
         Top             =   1680
         Width           =   3345
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   1350
         Width           =   3345
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   3345
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Port*"
         Height          =   255
         Left            =   330
         TabIndex        =   16
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Req SSL*"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   1050
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password*"
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   1710
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Username*"
         Height          =   255
         Left            =   30
         TabIndex        =   13
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter The SMTP Server Configuration Correctly"
         Height          =   405
         Left            =   150
         TabIndex        =   12
         Top             =   2220
         Width           =   4275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server*"
         Height          =   255
         Left            =   330
         TabIndex        =   11
         Top             =   390
         Width           =   645
      End
   End
   Begin prj_tpc.vbButton cmdCancel 
      Height          =   450
      Left            =   5610
      TabIndex        =   9
      Top             =   4185
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
      MICON           =   "frmSMTPServer.frx":2130
      PICN            =   "frmSMTPServer.frx":214C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdTest 
      Height          =   450
      Left            =   2340
      TabIndex        =   7
      Top             =   4170
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
      BTYPE           =   14
      TX              =   "&Test"
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
      MICON           =   "frmSMTPServer.frx":31DE
      PICN            =   "frmSMTPServer.frx":31FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   4770
      Left            =   0
      Picture         =   "frmSMTPServer.frx":428C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2265
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   2280
      X2              =   2280
      Y1              =   -30
      Y2              =   3390
   End
   Begin VB.Image Image1 
      Height          =   3780
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2265
   End
End
Attribute VB_Name = "frm_etc_smtp_server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim objControl      As CONTROL

Private Sub chkReqAuth_Click()
    If chkReqAuth.Value = 1 Then
        txtUsername.Enabled = True
        txtPassword.Enabled = True
    Else
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtUsername.Enabled = False
        txtPassword.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    SQL = "SELECT * FROM s_mail WHERE s_number = 1"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        txtServer.Text = rs!smtp_server
        txtPort.Text = rs!no_port
        chkSSL.Value = rs!req_ssl
        txtUsername.Text = rs!Username
        txtPassword.Text = DecryptINI(rs!Pwd, pEncryptionPassword)
        txtSenderName.Text = rs!sender_name
        txtSenderEmail.Text = rs!sender_mail
        chkReqAuth.Value = IIf(IsNull(rs!req_auth), 0, rs!req_auth)
    Else
        For Each objControl In Me.Controls
            If TypeOf objControl Is TextBox Then
                objControl.Text = ""
            End If
        Next
    End If
    rs.Close
    
    Call load_data_user_access(Me)
    cmdSave.Enabled = blnUser_Add
End Sub

Private Sub chkSSL_Click()
    If chkSSL.Value = 0 Then chkSSL.Caption = "No" Else chkSSL.Caption = "Yes"
End Sub

Private Sub cmdSave_Click()
'On Error GoTo Err:
    For Each objControl In Me.Controls
        If TypeOf objControl Is TextBox Then
            If Trim$(objControl.Text) = vbNullString Then
                If chkReqAuth.Value = 0 Then
                    If LCase(objControl.Text) = "txtusername" Or LCase(objControl.Text) = "txtpassword" Then
                        MsgBox "All Columns Are Required...", vbExclamation, headerMSG
                        Exit Sub
                    End If
                Else
                    MsgBox "All Columns Are Required...", vbExclamation, headerMSG
                    objControl.SetFocus
                    Exit Sub
                End If
            End If
        End If
    Next
    
    SQL = "SELECT * FROM s_mail"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        SQL = "UPDATE s_mail SET sender_mail = '" & Trim$(txtSenderEmail.Text) & "', " & _
                "sender_name = '" & Trim$(txtSenderName.Text) & "', " & _
                "username = '" & Trim$(txtUsername.Text) & "', " & _
                "pwd = '" & EncryptINI(Trim$(txtPassword.Text), pEncryptionPassword) & "', " & _
                "smtp_server = '" & Trim$(txtServer.Text) & "', " & _
                "no_port = " & Trim$(txtPort.Text) & ", " & _
                "req_ssl = " & chkSSL.Value & ", " & _
                "req_auth = " & chkReqAuth.Value & " " & _
              "WHERE s_number = 1"
        CnG.Execute SQL
    Else
        SQL = "INSERT INTO s_mail(s_number,sender_mail,sender_name,username," & _
                "pwd,smtp_server,no_port,req_ssl,req_auth) " & _
                "VALUES( " & _
                "1,'" & Trim$(txtSenderEmail.Text) & "','" & Trim$(txtSenderName.Text) & "'," & _
                "'" & Trim$(txtUsername.Text) & "','" & EncryptINI(Trim$(txtPassword.Text), pEncryptionPassword) & "'," & _
                "'" & Trim$(txtServer.Text) & "'," & Trim$(txtPort.Text) & "," & _
                "" & chkSSL.Value & "," & chkReqAuth.Value & ")"
        CnG.Execute SQL
    End If
    rs.Close
    
    MsgBox "Database Configuration Successfully...", vbInformation, headerMSG
    Exit Sub
    
Err:
MsgBox Err.Description, vbExclamation, headerMSG

End Sub

Private Sub cmdTest_Click()
On Error GoTo Err
        
    Dim lobj_cdomsg      As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = Trim$(txtServer.Text)
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = CInt(Trim$(txtPort.Text))
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = CBool(chkSSL.Value)
    
    If chkReqAuth.Value = 1 Then
        lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
        lobj_cdomsg.Configuration.Fields(cdoSendUserName) = Trim$(txtUsername.Text)
        lobj_cdomsg.Configuration.Fields(cdoSendPassword) = Trim$(txtPassword.Text)
    End If
    
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    
    If chkReqAuth.Value = 1 Then
        lobj_cdomsg.To = Trim$(txtUsername.Text)
    Else
        lobj_cdomsg.To = Trim$(txtSenderEmail.Text)
    End If
    
    lobj_cdomsg.from = Trim$(txtSenderName.Text) & "<" & Trim$(txtSenderEmail.Text) & ">"
    lobj_cdomsg.Subject = "Test Mail (SMTP Server) Configuration"
    lobj_cdomsg.TextBody = "Mail (SMTP Server) Test Successfully!"
    lobj_cdomsg.Send
    
    Set lobj_cdomsg = Nothing
    
    MsgBox "Test SMTP Server Configuration Successfully...", vbInformation, headerMSG
    Exit Sub
          
Err:
MsgBox Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_etc_smtp_server = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Private Sub Text3_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then cmdConnect_Click
'End Sub


