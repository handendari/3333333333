VERSION 5.00
Begin VB.Form frm_etc_finger_reg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FINGER REGISTRATION"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7995
   Icon            =   "frm_etc_finger_reg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_verify 
      Caption         =   "&Verify"
      Height          =   645
      Left            =   2160
      Picture         =   "frm_etc_finger_reg.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmd_close 
      Caption         =   "&Close"
      Height          =   645
      Left            =   6600
      Picture         =   "frm_etc_finger_reg.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmd_set_finger 
      Caption         =   "&Get"
      Height          =   645
      Left            =   5520
      Picture         =   "frm_etc_finger_reg.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3885
      Index           =   1
      Left            =   270
      TabIndex        =   0
      Top             =   158
      Width           =   7395
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frm_etc_finger_reg.frx":10AA
         Left            =   1250
         List            =   "frm_etc_finger_reg.frx":10CC
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.PictureBox picSample 
         Height          =   1635
         Index           =   3
         Left            =   5580
         ScaleHeight     =   1575
         ScaleWidth      =   1215
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   915
         Width           =   1275
      End
      Begin VB.PictureBox picSample 
         Height          =   1635
         Index           =   2
         Left            =   4140
         ScaleHeight     =   1575
         ScaleWidth      =   1215
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   900
         Width           =   1275
      End
      Begin VB.PictureBox picSample 
         Height          =   1635
         Index           =   1
         Left            =   2700
         ScaleHeight     =   1575
         ScaleWidth      =   1215
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   915
         Width           =   1275
      End
      Begin VB.PictureBox picSample 
         Height          =   1635
         Index           =   0
         Left            =   1260
         ScaleHeight     =   1575
         ScaleWidth      =   1215
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   900
         Width           =   1275
      End
      Begin VB.Shape dot 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   6060
         Shape           =   3  'Circle
         Top             =   2655
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape dot 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   4680
         Shape           =   3  'Circle
         Top             =   2655
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape dot 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   3180
         Shape           =   3  'Circle
         Top             =   2655
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape dot 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   1740
         Shape           =   3  'Circle
         Top             =   2655
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "JARI"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   420
         Width           =   345
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "..."
         Height          =   255
         Left            =   900
         TabIndex        =   6
         Top             =   3360
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frm_etc_finger_reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim WithEvents registration As UareUSDK.clsFPRegistration
'Dim WithEvents FPImage As UareUSDK.clsFPImage
'Dim Finger As String
'Dim FPDatabase As New UareUSDK.clsFPDatabase
'Dim Device As New UareUSDK.clsFPDevice
'Dim bln1 As Boolean
'
'
'
'Private Sub Form_Activate()
'Call connect_finger_uareu
'End Sub
'
'Public Sub register()
'bln1 = True
'Call connect_finger_uareu
'
'Set registration = New UareUSDK.clsFPRegistration
'Set FPImage = New UareUSDK.clsFPImage
'
'registration.PictureSamplePath = App.Path & "\FPTemp.BMP"
'registration.PictureSampleHeight = picSample(0).Height
'registration.PictureSampleWidth = picSample(0).Width
'End Sub
'
'Public Sub waiting_finger()
'Dim str_emp_code As String
'
'Call register
'
'str_emp_code = frm_mst_enroll_uareu.Adodc2.Recordset.Fields("employee_code").Value
'registration.FPRegistration str_emp_code, GetFingerNumber(Combo2.Text)
'Label10.Caption = IIf(bln1, "Silahkan Letakkan jari anda pada Sensor!", "...")
'End Sub
'
'
'Private Sub cmd_close_Click()
'Unload Me
'End Sub
'
'Private Sub cmd_set_finger_Click()
'Call waiting_finger
'End Sub
'
'Private Sub cmd_verify_Click()
'frm_etc_verify.Show 1
'End Sub
'
'
'Private Sub Registration_FPRegistrationImage(CurentSample As Integer)
'  picSample(CurentSample) = LoadPicture(App.Path & "\FPTemp.BMP")
'  dot(CurentSample).Visible = True
'End Sub
'
'Private Sub Registration_FPRegistrationStatus(Status As RegistrationStatus)
'Select Case Status
'    Case r_OK
'      MsgBox "Registration Success!", vbInformation, headerMSG
'      Label10.Caption = "..."
'
'      picSample(0) = Nothing
'      picSample(1) = Nothing
'      picSample(2) = Nothing
'      picSample(3) = Nothing
'
'      dot(0).Visible = False
'      dot(1).Visible = False
'      dot(2).Visible = False
'      dot(3).Visible = False
'      Call frm_mst_enroll_uareu.load_data_enroll
'
'    Case r_FpIdAlreadyExist
'      MsgBox "ID dan finger number already exist", vbInformation, headerMSG
'      bln1 = False
'
'    Case r_NoDevice
'      MsgBox "Device not exits", vbInformation, headerMSG
'      Label10.Caption = "..."
'
'    Case r_WrongDeviceSN
'      MsgBox "Wrong Device Serial Number", vbInformation, headerMSG
'      Label10.Caption = "..."
'
'    Case r_RegFailed
'      MsgBox "Registration Error!", vbInformation, headerMSG
'      Label10.Caption = "..."
'
'    Case r_WrongFingerNr
'      MsgBox "Finger number must between 0 to 9", vbInformation, headerMSG
'      Label10.Caption = "..."
'
'End Select
'End Sub
'
'Private Function GetFingerNumber(Finger As String) As FingerNumber
'  Dim j As Byte
'  Select Case Finger
'    Case "Jari Kelingking Kiri"
'      j = 0
'    Case "Jari Manis Kiri"
'      j = 1
'    Case "Jari Tengah Kiri"
'      j = 2
'    Case "Jari Telunjuk Kiri"
'      j = 3
'    Case "Ibu Jari Kiri"
'      j = 4
'    Case "Ibu Jari Kanan"
'      j = 5
'    Case "Jari Telunjuk Kanan"
'      j = 6
'    Case "Jari Tengah Kanan"
'      j = 7
'    Case "Jari Manis Kanan"
'      j = 8
'    Case "Jari Kelingking Kanan"
'      j = 9
'    Case "(None)"
'      j = 10
'    Case Else
'      j = 10
'  End Select
'
'  GetFingerNumber = j
'End Function
'
'Private Sub connect_finger_uareu()
'If FPDatabase.ActiveConnection(strConn) = sc_Success Then
'   'MsgBox "Database Connection Success", 0, "Keterangan"
'Else
'   MsgBox "Database Connection Error!", vbInformation, headerMSG
'   Unload Me
'End If
'End Sub
