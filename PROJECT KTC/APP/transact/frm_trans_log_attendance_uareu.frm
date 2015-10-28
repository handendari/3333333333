VERSION 5.00
Begin VB.Form frm_trans_log_attendance_uareu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LOG ATTENDANCE (BASE SENSOR)"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11010
   Icon            =   "frm_trans_log_attendance_uareu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   645
      Left            =   9120
      Picture         =   "frm_trans_log_attendance_uareu.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6960
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   960
      TabIndex        =   15
      Top             =   2400
      Width           =   9135
      Begin VB.TextBox txt_finger_name 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txt_finger_number 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txt_employee_code 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txt_status 
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
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Width           =   4005
      End
      Begin VB.PictureBox picSample 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1890
         Index           =   5
         Left            =   840
         ScaleHeight     =   1860
         ScaleWidth      =   1305
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "STATUS"
         Height          =   195
         Left            =   2880
         TabIndex        =   20
         Top             =   1980
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         Height          =   195
         Left            =   2880
         TabIndex        =   19
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FINGER"
         Height          =   195
         Left            =   2880
         TabIndex        =   18
         Top             =   1500
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ID / EMP. CODE"
         Height          =   195
         Left            =   2880
         TabIndex        =   17
         Top             =   540
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TYPE"
      Height          =   975
      Left            =   960
      TabIndex        =   13
      Top             =   1440
      Width           =   9135
      Begin VB.OptionButton opt_break_in 
         Caption         =   "BREAK IN"
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton opt_break_out 
         Caption         =   "BREAK OUT"
         Height          =   255
         Left            =   5040
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton opt_out 
         Caption         =   "OUT"
         Height          =   255
         Left            =   7320
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton opt_in 
         Caption         =   "IN"
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   120
      Top             =   1560
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   960
      TabIndex        =   0
      Top             =   5280
      Width           =   9135
   End
   Begin VB.Label lbl_time 
      Alignment       =   2  'Center
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   10575
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   3210
      TabIndex        =   12
      Top             =   1560
      Width           =   2685
   End
   Begin VB.Label lbl_date 
      Alignment       =   2  'Center
      Caption         =   "date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   10575
   End
End
Attribute VB_Name = "frm_trans_log_attendance_uareu"
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
'
'
'Private Function download_data_log() As Boolean
''On Error GoTo capErr
'Dim rs1 As New ADODB.Recordset
'Dim flag_io As Integer
'
'rs1.Open "select * from h_log_attendance where employee_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic
'If opt_in Then
'    flag_io = 0
'ElseIf opt_out Then
'    flag_io = 1
'ElseIf opt_break_in Then
'    flag_io = 2
'ElseIf opt_break_out Then
'    flag_io = 3
'End If
'
'CnG.BeginTrans
'With rs1
'    .AddNew
'
'    '.Fields("att_number").Value = 1 'must have tried by "no auto increment"
'    .Fields("att_date").Value = Now
'    '.Fields("ip_address").Value = FG_IP_ADDRESS
'    '.Fields("enrollnumber").Value = dwEnrollNumber
'    .Fields("employee_code").Value = txt_employee_code
'    .Fields("verifymode").Value = 1
'    .Fields("flag_io").Value = flag_io
'    .Fields("entry_date").Value = Now
'
'    .Update
'End With
'CnG.CommitTrans
'
'download_data_log = True
'Exit Function
'
'capErr:
'MsgBox "Error downloading data...", vbCritical, headerMSG
'MousePointer = vbDefault
'download_data_log = False
'End Function
'
'
'Private Sub CmdExit_Click()
'Unload Me
'End Sub
'
'Private Sub Timer1_Timer()
'lbl_time = Format(Now, "hh:nn:ss")
'End Sub
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
'Unload Me
'End Sub
'
'Private Function day_name(ByVal i As Integer) As String
'Select Case i
'
'Case 0
'    day_name = "Sunday"
'Case 1
'    day_name = "Monday"
'Case 2
'    day_name = "Tuesday"
'Case 3
'    day_name = "Wednesday"
'Case 4
'    day_name = "Thursday"
'Case 5
'    day_name = "Friday"
'Case 6
'    day_name = "Saturday"
'
'End Select
'End Function
'
'Private Sub Form_Load()
'lbl_date = day_name(Weekday(Now)) & ", " & Format(Now, "dd-mm-yyyy")
'
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
'txt_employee_code.Text = ID
'txt_finger_number.Text = FingerNr
'
'If txt_finger_number.Text = 0 Then txt_finger_name.Text = "Jari Kelingking Kiri"
'If txt_finger_number.Text = 1 Then txt_finger_name.Text = "Jari Manis Kiri"
'If txt_finger_number.Text = 2 Then txt_finger_name.Text = "Jari Tengah Kiri"
'If txt_finger_number.Text = 3 Then txt_finger_name.Text = "Jari Telunjuk Kiri"
'If txt_finger_number.Text = 4 Then txt_finger_name.Text = "Ibu Jari Kiri"
'If txt_finger_number.Text = 5 Then txt_finger_name.Text = "Ibu Jari Kanan"
'If txt_finger_number.Text = 6 Then txt_finger_name.Text = "Jari Telunjuk Kanan "
'If txt_finger_number.Text = 7 Then txt_finger_name.Text = "Jari Tengah Kanan"
'If txt_finger_number.Text = 8 Then txt_finger_name.Text = "Jari Manis Kanan"
'If txt_finger_number.Text = 9 Then txt_finger_name.Text = "Jari Kelingking Kanan"
'If txt_finger_number.Text >= 10 Then txt_finger_name.Text = ""
'
'Call set_data(True)
'If download_data_log = False Then
'    MsgBox "Error Recording Data..." & vbCrLf & "Please Check it!", vbInformation, headerMSG
'    Exit Sub
'End If
'
'End Sub
'
'Private Sub set_data(ByVal bln1 As Boolean)
'Dim rs1 As New ADODB.Recordset
'Dim str_type As String
'
'If bln1 Then
'    rs1.Open "select * from m_employee where employee_code = '" & Trim(txt_employee_code.Text) & "'", CnG, adOpenStatic, adLockReadOnly
'
'    If rs1.RecordCount <= 0 Then
'        txt_finger_name.Text = ""
'        txt_finger_number.Text = ""
'        txt_employee_name.Text = ""
'        Exit Sub
'    Else
'        txt_employee_name.Text = rs1.Fields("employee_name").Value
'        If opt_in Then
'            str_type = "(IN)"
'        ElseIf opt_out Then
'            str_type = "(OUT)"
'        ElseIf opt_break_in Then
'            str_type = "(BREAK IN)"
'        ElseIf opt_break_out Then
'            str_type = "(BREAK OUT)"
'        End If
'
'            List1.AddItem vbTab & rs1.Fields("employee_code").Value & vbTab & rs1("employee_name") _
'                & vbTab & vbTab & str_type & vbTab & Format(Now, "yyyy-mm-dd hh:nn:ss")
'        'End If
'    End If
'
'Else
'    txt_finger_name.Text = ""
'    txt_finger_number.Text = ""
'    txt_employee_name.Text = ""
'    txt_employee_code.Text = ""
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
'      txt_status.Text = "Multiple Match"
'      Call set_data(False)
'
'    Case v_OK
'      txt_status.Text = "Sidik Jari Diterima"
'      txt_status.BackColor = vbGreen
'
'    Case v_NotFound
'      txt_status.Text = "Sidik Jari Tidak Ditemukan"
'      txt_status.BackColor = vbRed
'      Call set_data(False)
'
'    Case v_WrongDeviceSN
'      txt_status.Text = "Wrong Device Serial Number"
'      txt_status.BackColor = vbRed
'      Call set_data(False)
'
'    Case v_VerFailed
'      txt_status.Text = "Verification False!"
'      txt_status.BackColor = vbRed
'      Call set_data(False)
'
'    Case v_NoDevice
'      txt_status.Text = "Device not exits"
'      txt_status.BackColor = vbRed
'      Call set_data(False)
'
'  End Select
'
'  Verification.FPVerification
'
'End Sub
'
'
'
