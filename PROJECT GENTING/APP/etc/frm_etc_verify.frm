VERSION 5.00
Begin VB.Form frm_etc_verify 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VERIFYING FINGER..."
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7995
   Icon            =   "frm_etc_verify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_close 
      Caption         =   "&Close"
      Height          =   645
      Left            =   6600
      Picture         =   "frm_etc_verify.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.PictureBox picSample 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1890
         Index           =   5
         Left            =   480
         ScaleHeight     =   1890
         ScaleWidth      =   1335
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1920
         Width           =   3525
      End
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ID / EMP. CODE"
         Height          =   195
         Left            =   1920
         TabIndex        =   11
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FINGER"
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   1500
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         Height          =   195
         Left            =   1920
         TabIndex        =   9
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "STATUS"
         Height          =   195
         Left            =   1920
         TabIndex        =   8
         Top             =   1980
         Width           =   645
      End
   End
End
Attribute VB_Name = "frm_etc_verify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim WithEvents Verification As UareUSDK.clsFPVerification
'Dim FPDatabase As New UareUSDK.clsFPDatabase
'Dim Device As New UareUSDK.clsFPDevice
'
'
'
'Private Sub connect_finger_uareu()
'If FPDatabase.ActiveConnection(strConn) = sc_Success Then
'   'MsgBox "Database Connection Success", 0, "Keterangan"
'Else
'   MsgBox "Database Connection Error!", vbInformation, headerMSG
'   Unload Me
'End If
'End Sub
'
'Private Sub cmd_close_Click()
'Verification.FPVerificationFilter ("")
'Unload Me
'End Sub
'
'Private Sub Form_Load()
'Call connect_finger_uareu
'Call verifikasi
'End Sub
'
'Sub verifikasi()
'Set Verification = New UareUSDK.clsFPVerification
'Verification.PictureSamplePath = App.Path & "\FPTemp.BMP"
'Verification.PictureSampleHeight = picSample(5).Height
'Verification.PictureSampleWidth = picSample(5).Width
'Verification.FPVerification
'End Sub
'
'Private Sub Verification_FPVerificationID(ID As String, FingerNr As UareUSDK.FingerNumber)
'text1.Text = ID
'Text3.Text = FingerNr
'
'If Text3.Text = 0 Then Text4.Text = "Jari Kelingking Kiri"
'If Text3.Text = 1 Then Text4.Text = "Jari Manis Kiri"
'If Text3.Text = 2 Then Text4.Text = "Jari Tengah Kiri"
'If Text3.Text = 3 Then Text4.Text = "Jari Telunjuk Kiri"
'If Text3.Text = 4 Then Text4.Text = "Ibu Jari Kiri"
'If Text3.Text = 5 Then Text4.Text = "Ibu Jari Kanan"
'If Text3.Text = 6 Then Text4.Text = "Jari Telunjuk Kanan "
'If Text3.Text = 7 Then Text4.Text = "Jari Tengah Kanan"
'If Text3.Text = 8 Then Text4.Text = "Jari Manis Kanan"
'If Text3.Text = 9 Then Text4.Text = "Jari Kelingking Kanan"
'If Text3.Text >= 10 Then Text4.Text = ""
'
'Call set_data(True)
'End Sub
'
'Private Sub set_data(ByVal bln1 As Boolean)
'Dim rs1 As New ADODB.Recordset
'
'If bln1 Then
'    rs1.Open "select * from m_employee where employee_code = '" & Trim(text1.Text) & "'", CnG, adOpenStatic, adLockReadOnly
'
'    If rs1.RecordCount <= 0 Then
'       Text4.Text = ""
'       Text3.Text = ""
'       Text2.Text = ""
'       Exit Sub
'    Else
'       Text2.Text = rs1.Fields("employee_name").Value
'    End If
'
'Else
'    Text4.Text = ""
'    Text3.Text = ""
'    Text2.Text = ""
'    text1.Text = ""
'End If
'
'End Sub
'
'Private Sub Verification_FPVerificationImage()
'  picSample(5) = LoadPicture(App.Path & "\FPTemp.BMP")
'End Sub
'
'Private Sub Verification_FPVerificationStatus(Status As VerificationStatus)
'  Select Case Status
'
'    Case v_MultiplelMatch
'      Text5.Text = "Multiple Match"
'      Call set_data(False)
'
'    Case v_OK
'      Text5.Text = "Sidik Jari Diterima"
'      Text5.BackColor = vbGreen
'
'    Case v_NotFound
'      Text5.Text = "Sidik Jari Tidak Ditemukan"
'      Text5.BackColor = vbRed
'      Call set_data(False)
'
'    Case v_WrongDeviceSN
'      Text5.Text = "Wrong Device Serial Number"
'      Text5.BackColor = vbRed
'      Call set_data(False)
'
'    Case v_VerFailed
'      Text5.Text = "Verification False!"
'      Text5.BackColor = vbRed
'      Call set_data(False)
'
'    Case v_NoDevice
'      Text5.Text = "Device not exits"
'      Text5.BackColor = vbRed
'      Call set_data(False)
'
'  End Select
'
'  Verification.FPVerification
'
'End Sub
'
'
Private Sub cmd_close_Click()
Unload Me
End Sub
